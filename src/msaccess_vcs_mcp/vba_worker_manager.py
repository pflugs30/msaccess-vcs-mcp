"""Thread-based management for isolated VBA execution with timeout."""

from __future__ import annotations

import os
import threading
import time
from typing import Any, Optional

try:
    import pythoncom
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

from .addin_integration import VCSAddinIntegration
from .com_recovery import (
    classify_com_error,
    get_recovery_manager,
    is_recoverable_pattern,
)
from .usage_logging import (
    log_com_recovery_event,
    log_diagnostic_event,
    log_vba_worker_event,
)


DEFAULT_RUN_VBA_TIMEOUT_SEC = 45.0
DEFAULT_RECOVERY_PROBE_TIMEOUT_SEC = 10.0


def _read_timeout_env(name: str, default: float) -> float:
    try:
        value = float(os.environ.get(name, str(default)))
    except ValueError:
        return default
    return value if value > 0 else default


def get_run_vba_timeout(timeout_seconds: float | None = None) -> float:
    """Resolve the effective timeout for one `vcs_run_vba` call."""
    if timeout_seconds is not None and timeout_seconds > 0:
        return float(timeout_seconds)
    return _read_timeout_env("ACCESS_VCS_RUN_VBA_TIMEOUT_SEC", DEFAULT_RUN_VBA_TIMEOUT_SEC)


def get_recovery_probe_timeout() -> float:
    """Resolve the timeout for short recovery health probes."""
    return _read_timeout_env(
        "ACCESS_VCS_RECOVERY_PROBE_TIMEOUT_SEC",
        DEFAULT_RECOVERY_PROBE_TIMEOUT_SEC,
    )


class VBAWorkerManager:
    """Run VBA in a daemon thread with a hard timeout.

    Modelled on ``_probe_with_timeout`` in ``addin_integration.py`` and
    ``_run_dao_with_timeout`` in db-inspector-mcp.  The worker thread
    creates its own COM apartment via ``pythoncom.CoInitialize()`` and
    re-acquires the Access instance through the Running Object Table so
    that ``thread.join(timeout)`` can fire even if the COM call blocks.
    """

    _active_worker: Optional[threading.Thread] = None

    def __init__(self) -> None:
        self.recovery = get_recovery_manager()

    def run_vba(
        self,
        database_path: str,
        code: str,
        addin_path: str | None = None,
        timeout_seconds: float | None = None,
    ) -> dict[str, Any]:
        """Run VBA in a timeout-controlled worker thread with recovery."""
        timeout = get_run_vba_timeout(timeout_seconds)
        preflight = self._probe_if_needed(database_path, addin_path)
        if preflight is not None:
            return preflight

        result = self._run_worker(
            operation="run_vba",
            database_path=database_path,
            addin_path=addin_path,
            code=code,
            timeout_seconds=timeout,
        )
        if result.get("success"):
            self.recovery.mark_healthy(database_path)
            return result

        return self._handle_run_failure(
            database_path=database_path,
            addin_path=addin_path,
            code=code,
            timeout_seconds=timeout,
            result=result,
        )

    def probe(self, database_path: str, addin_path: str | None = None) -> dict[str, Any]:
        """Run a short Access/add-in health probe in a worker thread."""
        return self._run_worker(
            operation="probe",
            database_path=database_path,
            addin_path=addin_path,
            code=None,
            timeout_seconds=get_recovery_probe_timeout(),
        )

    # ------------------------------------------------------------------
    # Recovery helpers (unchanged from subprocess version)
    # ------------------------------------------------------------------

    def _probe_if_needed(
        self,
        database_path: str,
        addin_path: str | None,
    ) -> dict[str, Any] | None:
        if not self.recovery.should_probe(database_path):
            return None

        state = self.recovery.mark_probing(database_path)
        log_com_recovery_event(
            "com_recovery_probe_start",
            database_path=database_path,
            status=state.status,
        )
        log_diagnostic_event(
            "com_recovery_probe_start",
            database=str(database_path),
            status=state.status,
        )

        probe_result = self.probe(database_path, addin_path)
        if probe_result.get("success"):
            recovered = self.recovery.mark_healthy(database_path, recovered=True)
            log_com_recovery_event(
                "com_recovery_probe_result",
                database_path=database_path,
                status=recovered.status,
                success=True,
            )
            log_diagnostic_event(
                "com_recovery_probe_result",
                database=str(database_path),
                status=recovered.status,
                success=True,
            )
            return None

        state = self.recovery.mark_failure(
            database_path,
            probe_result.get("error", "Access recovery probe failed"),
            probe_result.get("error_pattern"),
            timed_out=probe_result.get("timed_out", False),
        )
        log_com_recovery_event(
            "com_recovery_probe_result",
            database_path=database_path,
            status=state.status,
            success=False,
            error=probe_result.get("error"),
            error_pattern=state.error_pattern,
        )
        log_diagnostic_event(
            "com_recovery_probe_result",
            database=str(database_path),
            status=state.status,
            success=False,
            error=probe_result.get("error"),
            error_pattern=state.error_pattern,
        )
        return self._recoverable_error(
            "Access is still not responding after a recovery probe.",
            state.error_pattern or "access_unresponsive",
            hint=(
                "Resume execution in the VBE, dismiss any Access modal dialog, "
                "or close and reopen the database, then retry."
            ),
            recovery_status=state.status,
        )

    def _handle_run_failure(
        self,
        database_path: str,
        addin_path: str | None,
        code: str,
        timeout_seconds: float,
        result: dict[str, Any],
    ) -> dict[str, Any]:
        error = result.get("error", "VBA worker failed")
        pattern = result.get("error_pattern") or classify_com_error(error)
        timed_out = result.get("timed_out", False)
        phase = result.get("phase")
        state = self.recovery.mark_failure(
            database_path,
            error,
            pattern,
            timed_out=timed_out,
        )

        if timed_out:
            return self._recoverable_error(
                error,
                "timeout",
                hint=(
                    "The MCP server is still connected, but Access may still "
                    "be running the submitted VBA. Wait for Access to respond, "
                    "resume the VBE if it is in break mode, or dismiss any modal dialog."
                ),
                recovery_status=state.status,
                timed_out=True,
            )

        if (
            is_recoverable_pattern(pattern)
            and phase in {"connect", "load_addin"}
        ):
            probe_result = self.probe(database_path, addin_path)
            if probe_result.get("success"):
                self.recovery.mark_healthy(database_path, recovered=True)
                retry_result = self._run_worker(
                    operation="run_vba",
                    database_path=database_path,
                    addin_path=addin_path,
                    code=code,
                    timeout_seconds=timeout_seconds,
                    retry=True,
                )
                if retry_result.get("success"):
                    self.recovery.mark_healthy(database_path)
                    return retry_result

                retry_error = retry_result.get("error", "VBA worker retry failed")
                retry_pattern = retry_result.get("error_pattern") or classify_com_error(retry_error)
                retry_state = self.recovery.mark_failure(
                    database_path,
                    retry_error,
                    retry_pattern,
                    timed_out=retry_result.get("timed_out", False),
                )
                return self._recoverable_error(
                    retry_error,
                    retry_pattern,
                    hint=(
                        "Access responded to the recovery probe, but the retry failed. "
                        "Confirm Access is responsive before retrying."
                    ),
                    recovery_status=retry_state.status,
                    timed_out=retry_result.get("timed_out", False),
                    phase=retry_result.get("phase"),
                    recoverable=is_recoverable_pattern(retry_pattern),
                )

        return self._recoverable_error(
            error,
            pattern,
            hint=(
                "The failed call was not retried automatically because the VBA "
                "may have started. Retry after confirming Access is responsive."
            ),
            recovery_status=state.status,
            phase=phase,
            recoverable=is_recoverable_pattern(pattern),
        )

    # ------------------------------------------------------------------
    # Thread-based worker (replaces subprocess)
    # ------------------------------------------------------------------

    def _run_worker(
        self,
        operation: str,
        database_path: str,
        addin_path: str | None,
        code: str | None,
        timeout_seconds: float,
        retry: bool = False,
    ) -> dict[str, Any]:
        cls = type(self)
        if cls._active_worker is not None and cls._active_worker.is_alive():
            return {
                "success": False,
                "error": (
                    "A previous VBA worker thread is still running. "
                    "Access may be in VBA break mode or blocked on a modal dialog."
                ),
                "error_pattern": "access_unresponsive",
                "recoverable": True,
                "phase": "guard",
            }

        start = time.perf_counter()
        log_vba_worker_event(
            "vba_worker_start",
            database_path=database_path,
            operation=operation,
            retry=retry,
        )
        log_diagnostic_event(
            "vba_worker_start",
            database=str(database_path),
            operation=operation,
            timeout_seconds=timeout_seconds,
            retry=retry,
        )

        result_box: dict[str, Any] = {}

        def worker() -> None:
            phase = "start"
            try:
                if not COM_AVAILABLE:
                    raise ImportError("pywin32 is required for COM automation")
                pythoncom.CoInitialize()
                try:
                    phase = "connect"
                    worker_app = VCSAddinIntegration._find_access_in_rot(database_path)
                    if worker_app is None:
                        raise RuntimeError(
                            f"Cannot find Access instance for {database_path} "
                            f"from worker thread. The Access application "
                            f"may have been closed."
                        )

                    phase = "load_addin"
                    addin = VCSAddinIntegration(addin_path)
                    addin.load_addin(worker_app, db_path=database_path)

                    if operation == "probe":
                        result_box["result"] = {
                            "success": True,
                            "operation": operation,
                            "phase": phase,
                            "result": "ok",
                        }
                        return

                    phase = "run_vba"
                    vba_result = addin.call_sync("RunVBA", code)
                    result_box["result"] = {
                        "success": True,
                        "operation": operation,
                        "phase": phase,
                        "result": vba_result,
                    }
                finally:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
            except Exception as exc:
                result_box["result"] = {
                    "success": False,
                    "operation": operation,
                    "phase": phase,
                    "error": str(exc),
                    "error_type": type(exc).__name__,
                    "error_pattern": classify_com_error(exc),
                }

        thread = threading.Thread(target=worker, daemon=True, name="vcs-vba-worker")
        cls._active_worker = thread
        thread.start()
        thread.join(timeout=timeout_seconds)

        duration_ms = round((time.perf_counter() - start) * 1000, 2)

        if thread.is_alive():
            error = f"VBA worker timed out after {timeout_seconds} seconds"
            log_vba_worker_event(
                "vba_worker_timeout",
                database_path=database_path,
                operation=operation,
                duration_ms=duration_ms,
                success=False,
                timed_out=True,
                error=error,
                error_pattern="timeout",
                retry=retry,
            )
            log_diagnostic_event(
                "vba_worker_timeout",
                database=str(database_path),
                operation=operation,
                duration_ms=duration_ms,
                error=error,
                retry=retry,
            )
            return {
                "success": False,
                "error": error,
                "error_pattern": "timeout",
                "recoverable": True,
                "timed_out": True,
                "duration_ms": duration_ms,
            }

        cls._active_worker = None

        response = result_box.get("result", {
            "success": False,
            "error": "VBA worker thread completed without producing a result",
            "error_pattern": "unknown",
        })
        response["duration_ms"] = duration_ms
        if not response.get("success"):
            response.setdefault("recoverable", is_recoverable_pattern(response.get("error_pattern")))

        log_vba_worker_event(
            "vba_worker_result",
            database_path=database_path,
            operation=operation,
            duration_ms=duration_ms,
            success=response.get("success"),
            timed_out=False,
            error=response.get("error"),
            error_pattern=response.get("error_pattern"),
            phase=response.get("phase"),
            retry=retry,
        )
        log_diagnostic_event(
            "vba_worker_result",
            database=str(database_path),
            operation=operation,
            duration_ms=duration_ms,
            success=response.get("success"),
            phase=response.get("phase"),
            error=response.get("error"),
            error_pattern=response.get("error_pattern"),
            retry=retry,
        )
        return response

    @staticmethod
    def _recoverable_error(
        error: str,
        error_pattern: str,
        hint: str,
        recovery_status: str,
        timed_out: bool = False,
        phase: str | None = None,
        recoverable: bool = True,
    ) -> dict[str, Any]:
        result: dict[str, Any] = {
            "success": False,
            "error": error,
            "error_pattern": error_pattern,
            "recoverable": recoverable,
            "recovery_status": recovery_status,
            "hint": hint,
        }
        if timed_out:
            result["timed_out"] = True
        if phase:
            result["phase"] = phase
        return result


_worker_manager = VBAWorkerManager()


def run_vba_resilient(
    database_path: str,
    code: str,
    addin_path: str | None = None,
    timeout_seconds: float | None = None,
) -> dict[str, Any]:
    """Run VBA through the thread-isolated worker manager."""
    return _worker_manager.run_vba(
        database_path=database_path,
        code=code,
        addin_path=addin_path,
        timeout_seconds=timeout_seconds,
    )
