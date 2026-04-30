"""COM disconnection classification and in-process recovery state."""

from __future__ import annotations

import os
import threading
import time
from dataclasses import dataclass
from typing import Any


RECOVERABLE_PATTERNS = {
    "timeout",
    "rpc_unavailable",
    "call_rejected",
    "object_disconnected",
    "access_not_found",
    "connection_closed",
    "access_unresponsive",
    "com_error",
}


def normalize_database_key(database_path: str) -> str:
    """Return a stable key for per-database recovery state."""
    return os.path.normcase(os.path.abspath(database_path))


def classify_com_error(error: str | BaseException | None) -> str:
    """Classify common Access/COM transport failures for recovery decisions."""
    if error is None:
        return "unknown"

    text = str(error)
    lowered = text.lower()

    if "timed out" in lowered or "timeout" in lowered:
        return "timeout"
    if "-2147023174" in lowered or "rpc server is unavailable" in lowered:
        return "rpc_unavailable"
    if "-2147418111" in lowered or "call was rejected by callee" in lowered:
        return "call_rejected"
    if (
        "-2147417848" in lowered
        or "object invoked has disconnected" in lowered
        or "object has disconnected" in lowered
        or "disconnected from its clients" in lowered
    ):
        return "object_disconnected"
    if "connection closed" in lowered or "not connected" in lowered:
        return "connection_closed"
    if "cannot find access instance" in lowered or "access application may have been closed" in lowered:
        return "access_not_found"
    if (
        "previous vcs add-in probe is still pending" in lowered
        or "no response from access" in lowered
        or "vba break mode" in lowered
        or "modal dialog" in lowered
    ):
        return "access_unresponsive"
    if "com_error" in lowered or "pywintypes.com_error" in lowered:
        return "com_error"
    if "addin" in lowered or "add-in" in lowered:
        return "addin_error"
    if "not found" in lowered:
        return "file_not_found"

    return "unknown"


def is_recoverable_pattern(pattern: str | None) -> bool:
    """Return whether a failure pattern is worth probing/rebinding after."""
    return pattern in RECOVERABLE_PATTERNS


@dataclass
class RecoveryState:
    """Last known COM health for a single database."""

    status: str = "healthy"
    failure_count: int = 0
    last_error: str | None = None
    error_pattern: str | None = None
    updated_at: float = 0.0

    def as_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "failure_count": self.failure_count,
            "last_error": self.last_error,
            "error_pattern": self.error_pattern,
            "updated_at": self.updated_at,
        }


class COMRecoveryManager:
    """Track broken Access COM connections and when to run recovery probes."""

    def __init__(self) -> None:
        self._states: dict[str, RecoveryState] = {}
        self._lock = threading.Lock()

    def get_state(self, database_path: str) -> RecoveryState:
        key = normalize_database_key(database_path)
        with self._lock:
            state = self._states.get(key)
            if state is None:
                state = RecoveryState(updated_at=time.time())
                self._states[key] = state
            return RecoveryState(**state.as_dict())

    def should_probe(self, database_path: str) -> bool:
        state = self.get_state(database_path)
        return state.status in {"timed_out", "disconnected", "probing", "unresponsive"}

    def mark_probing(self, database_path: str) -> RecoveryState:
        return self._update(database_path, status="probing")

    def mark_healthy(self, database_path: str, recovered: bool = False) -> RecoveryState:
        return self._update(
            database_path,
            status="recovered" if recovered else "healthy",
            failure_count=0,
            last_error=None,
            error_pattern=None,
        )

    def mark_failure(
        self,
        database_path: str,
        error: str,
        error_pattern: str | None = None,
        timed_out: bool = False,
    ) -> RecoveryState:
        pattern = error_pattern or classify_com_error(error)
        if timed_out or pattern == "timeout":
            status = "timed_out"
        elif is_recoverable_pattern(pattern):
            status = "disconnected"
        else:
            status = "unresponsive"
        current = self.get_state(database_path)
        return self._update(
            database_path,
            status=status,
            failure_count=current.failure_count + 1,
            last_error=error,
            error_pattern=pattern,
        )

    def reset(self) -> None:
        """Clear all recovery state. Intended for tests."""
        with self._lock:
            self._states.clear()

    def _update(
        self,
        database_path: str,
        status: str,
        failure_count: int | None = None,
        last_error: str | None = None,
        error_pattern: str | None = None,
    ) -> RecoveryState:
        key = normalize_database_key(database_path)
        with self._lock:
            current = self._states.get(key, RecoveryState())
            state = RecoveryState(
                status=status,
                failure_count=current.failure_count if failure_count is None else failure_count,
                last_error=last_error,
                error_pattern=error_pattern,
                updated_at=time.time(),
            )
            self._states[key] = state
            return RecoveryState(**state.as_dict())


_recovery_manager = COMRecoveryManager()


def get_recovery_manager() -> COMRecoveryManager:
    """Return the process-wide COM recovery manager."""
    return _recovery_manager
