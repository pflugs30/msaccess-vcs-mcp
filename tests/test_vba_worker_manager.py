"""Tests for thread-based VBA worker management and COM recovery."""

from __future__ import annotations

import asyncio
import threading
import time
from unittest.mock import Mock, patch

import pytest

import msaccess_vcs_mcp.tools as tools_module
from msaccess_vcs_mcp.com_recovery import classify_com_error, get_recovery_manager
from msaccess_vcs_mcp.vba_worker_manager import VBAWorkerManager


@pytest.fixture(autouse=True)
def _reset_recovery_and_worker():
    get_recovery_manager().reset()
    VBAWorkerManager._active_worker = None
    yield
    get_recovery_manager().reset()
    VBAWorkerManager._active_worker = None


def _patch_worker_thread(monkeypatch, results: list[dict]):
    """Replace the worker's thread body with a synchronous result feeder.

    Each call to ``_run_worker`` pops the next result dict from *results*
    and stuffs it into the ``result_box`` that the real thread body would
    write.  The thread itself is a no-op so the test never touches COM.
    """
    call_log: list[dict] = []
    real_run_worker = VBAWorkerManager._run_worker

    def fake_run_worker(self, *, operation, database_path, addin_path,
                        code, timeout_seconds, retry=False):
        call_log.append({
            "operation": operation,
            "database_path": database_path,
            "code": code,
            "retry": retry,
        })
        if results and results[0].get("timeout"):
            result = results.pop(0)
            return {
                "success": False,
                "error": f"VBA worker timed out after {timeout_seconds} seconds",
                "error_pattern": "timeout",
                "recoverable": True,
                "timed_out": True,
                "duration_ms": 100.0,
            }
        if results:
            response = dict(results.pop(0))
            response.setdefault("duration_ms", 10.0)
            return response
        return {"success": False, "error": "No more fake results", "error_pattern": "unknown"}

    monkeypatch.setattr(VBAWorkerManager, "_run_worker", fake_run_worker)
    return call_log


# ------------------------------------------------------------------
# COM error classification (unchanged)
# ------------------------------------------------------------------

def test_classifies_recoverable_com_errors():
    assert classify_com_error("The RPC server is unavailable (-2147023174)") == "rpc_unavailable"
    assert classify_com_error("Call was rejected by callee (-2147418111)") == "call_rejected"
    assert classify_com_error("Object invoked has disconnected") == "object_disconnected"
    assert classify_com_error("MCP error -32000: Connection closed") == "connection_closed"
    assert classify_com_error("Operation timed out after 45 seconds") == "timeout"


# ------------------------------------------------------------------
# Happy path
# ------------------------------------------------------------------

def test_run_vba_success(monkeypatch):
    results = [
        {
            "success": True,
            "operation": "run_vba",
            "phase": "run_vba",
            "result": '{"success": true, "result": 42}',
        }
    ]
    calls = _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code="MCP_TempFunction = 42",
        addin_path="C:\\addin.accda",
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert result["result"] == '{"success": true, "result": 42}'
    assert calls[0]["operation"] == "run_vba"
    assert calls[0]["code"] == "MCP_TempFunction = 42"


# ------------------------------------------------------------------
# Timeout
# ------------------------------------------------------------------

def test_run_vba_timeout_returns_recoverable(monkeypatch):
    results = [{"timeout": True}]
    _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code="Do While True: Loop",
        timeout_seconds=0.1,
    )

    assert result["success"] is False
    assert result["timed_out"] is True
    assert result["recoverable"] is True
    assert result["error_pattern"] == "timeout"
    assert get_recovery_manager().get_state("C:\\db.accdb").status == "timed_out"


# ------------------------------------------------------------------
# Recovery probe after prior timeout
# ------------------------------------------------------------------

def test_next_call_probes_after_prior_timeout(monkeypatch):
    get_recovery_manager().mark_failure(
        "C:\\db.accdb",
        "VBA worker timed out after 45 seconds",
        "timeout",
        timed_out=True,
    )
    results = [
        {
            "success": True,
            "operation": "probe",
            "phase": "load_addin",
            "result": "ok",
        },
        {
            "success": True,
            "operation": "run_vba",
            "phase": "run_vba",
            "result": "done",
        },
    ]
    calls = _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code='MCP_TempFunction = "done"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert [c["operation"] for c in calls] == ["probe", "run_vba"]
    assert get_recovery_manager().get_state("C:\\db.accdb").status == "healthy"


# ------------------------------------------------------------------
# Pre-dispatch disconnect -> probe + retry
# ------------------------------------------------------------------

def test_predispatch_disconnect_probes_and_retries_once(monkeypatch):
    results = [
        {
            "success": False,
            "operation": "run_vba",
            "phase": "connect",
            "error": "The RPC server is unavailable",
            "error_pattern": "rpc_unavailable",
        },
        {
            "success": True,
            "operation": "probe",
            "phase": "load_addin",
            "result": "ok",
        },
        {
            "success": True,
            "operation": "run_vba",
            "phase": "run_vba",
            "result": "retried",
        },
    ]
    calls = _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code='MCP_TempFunction = "retried"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert result["result"] == "retried"
    assert [c["operation"] for c in calls] == ["run_vba", "probe", "run_vba"]


# ------------------------------------------------------------------
# run_vba phase failure is NOT retried
# ------------------------------------------------------------------

def test_run_phase_failure_is_not_retried(monkeypatch):
    results = [
        {
            "success": False,
            "operation": "run_vba",
            "phase": "run_vba",
            "error": "Call was rejected by callee",
            "error_pattern": "call_rejected",
        }
    ]
    calls = _patch_worker_thread(monkeypatch, results)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code="CurrentDb.Execute \"UPDATE T SET X = 1\"",
        timeout_seconds=1,
    )

    assert result["success"] is False
    assert result["recoverable"] is True
    assert result["phase"] == "run_vba"
    assert len(calls) == 1


# ------------------------------------------------------------------
# Active-worker single-flight guard
# ------------------------------------------------------------------

def test_active_worker_guard_returns_error():
    stop = threading.Event()
    holding = threading.Thread(target=stop.wait, daemon=True)
    holding.start()
    try:
        VBAWorkerManager._active_worker = holding
        manager = VBAWorkerManager()
        result = manager._run_worker(
            operation="run_vba",
            database_path="C:\\db.accdb",
            addin_path=None,
            code="x = 1",
            timeout_seconds=1,
        )
        assert result["success"] is False
        assert "still running" in result["error"]
        assert result["error_pattern"] == "access_unresponsive"
    finally:
        stop.set()
        holding.join(timeout=1)
        VBAWorkerManager._active_worker = None


# ------------------------------------------------------------------
# tools.py integration (result semantics preserved)
# ------------------------------------------------------------------

def test_vcs_run_vba_preserves_json_result_semantics(tmp_path, monkeypatch):
    db_path = tmp_path / "test.accdb"
    db_path.write_text("", encoding="utf-8")

    async def fake_ensure(ctx):
        return None

    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setattr(tools_module, "_ensure_env_loaded", fake_ensure)
    monkeypatch.setattr(tools_module.mcp, "get_context", lambda: None)
    monkeypatch.setattr(tools_module, "load_config", lambda: {})
    monkeypatch.setattr(
        tools_module,
        "get_config",
        lambda: {"ACCESS_VCS_ADDIN_PATH": "C:\\addin.accda"},
    )
    monkeypatch.setattr(
        tools_module,
        "run_vba_resilient",
        lambda **kwargs: {
            "success": True,
            "result": '{"success": true, "result": 7}',
        },
    )

    result = asyncio.run(
        tools_module.vcs_run_vba(
            database_path=str(db_path),
            code="MCP_TempFunction = 7",
            timeout_seconds=12,
        )
    )

    assert result == {"success": True, "result": 7}
