"""Parent-side management for isolated VBA worker processes."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

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
    """Launch one short-lived Python worker process per risky VBA call."""

    def __init__(self, worker_module: str = "msaccess_vcs_mcp.vba_worker") -> None:
        self.worker_module = worker_module
        self.recovery = get_recovery_manager()

    def run_vba(
        self,
        database_path: str,
        code: str,
        addin_path: str | None = None,
        timeout_seconds: float | None = None,
    ) -> dict[str, Any]:
        """Run VBA in an isolated worker and apply recovery policy."""
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
        """Run a short Access/add-in health probe in an isolated worker."""
        return self._run_worker(
            operation="probe",
            database_path=database_path,
            addin_path=addin_path,
            code=None,
            timeout_seconds=get_recovery_probe_timeout(),
        )

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
            and phase in {"connect", "load_addin", "validate_database", "load_config"}
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

    def _run_worker(
        self,
        operation: str,
        database_path: str,
        addin_path: str | None,
        code: str | None,
        timeout_seconds: float,
        retry: bool = False,
    ) -> dict[str, Any]:
        request: dict[str, Any] = {
            "operation": operation,
            "database_path": database_path,
            "addin_path": addin_path,
        }
        if code is not None:
            request["code"] = code

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

        with tempfile.TemporaryDirectory(prefix="vcs-vba-worker-") as temp_dir:
            request_path = Path(temp_dir) / "request.json"
            response_path = Path(temp_dir) / "response.json"
            request_path.write_text(json.dumps(request), encoding="utf-8")

            command = [
                sys.executable,
                "-m",
                self.worker_module,
                str(request_path),
                str(response_path),
            ]
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            try:
                _stdout, stderr = process.communicate(timeout=timeout_seconds)
            except subprocess.TimeoutExpired:
                process.kill()
                _stdout, stderr = process.communicate()
                duration_ms = round((time.perf_counter() - start) * 1000, 2)
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
                    "stderr": stderr.strip() if stderr else None,
                }

            duration_ms = round((time.perf_counter() - start) * 1000, 2)
            response = self._read_worker_response(
                response_path=response_path,
                exit_code=process.returncode,
                stderr=stderr,
            )
            response["duration_ms"] = duration_ms
            response["recoverable"] = is_recoverable_pattern(response.get("error_pattern"))

            log_vba_worker_event(
                "vba_worker_result",
                database_path=database_path,
                operation=operation,
                duration_ms=duration_ms,
                success=response.get("success"),
                timed_out=response.get("timed_out", False),
                error=response.get("error"),
                error_pattern=response.get("error_pattern"),
                phase=response.get("phase"),
                exit_code=process.returncode,
                retry=retry,
            )
            log_diagnostic_event(
                "vba_worker_result",
                database=str(database_path),
                operation=operation,
                duration_ms=duration_ms,
                success=response.get("success"),
                phase=response.get("phase"),
                exit_code=process.returncode,
                error=response.get("error"),
                error_pattern=response.get("error_pattern"),
                retry=retry,
            )
            return response

    @staticmethod
    def _read_worker_response(
        response_path: Path,
        exit_code: int | None,
        stderr: str | None,
    ) -> dict[str, Any]:
        if response_path.exists():
            try:
                response = json.loads(response_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError as exc:
                return {
                    "success": False,
                    "error": f"VBA worker returned invalid JSON: {exc}",
                    "error_pattern": "serialization_error",
                    "exit_code": exit_code,
                    "stderr": stderr.strip() if stderr else None,
                }
            response["exit_code"] = exit_code
            if stderr:
                response["stderr"] = stderr.strip()
            return response

        error = "VBA worker exited without writing a response"
        if stderr:
            error = f"{error}: {stderr.strip()}"
        return {
            "success": False,
            "error": error,
            "error_pattern": "worker_crash",
            "exit_code": exit_code,
            "stderr": stderr.strip() if stderr else None,
        }

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
    """Run VBA through the process-isolated worker manager."""
    return _worker_manager.run_vba(
        database_path=database_path,
        code=code,
        addin_path=addin_path,
        timeout_seconds=timeout_seconds,
    )
