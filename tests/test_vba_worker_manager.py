"""Tests for isolated VBA worker process management and COM recovery."""

from __future__ import annotations

import json
import subprocess
import asyncio
from pathlib import Path

import pytest

import msaccess_vcs_mcp.tools as tools_module
from msaccess_vcs_mcp.com_recovery import classify_com_error, get_recovery_manager
from msaccess_vcs_mcp.vba_worker_manager import VBAWorkerManager


@pytest.fixture(autouse=True)
def _reset_recovery_state():
    get_recovery_manager().reset()
    yield
    get_recovery_manager().reset()


class FakeProcess:
    """Small subprocess.Popen stand-in that reads/writes worker temp files."""

    def __init__(self, command: list[str], action: dict):
        self.command = command
        self.action = action
        self.returncode = action.get("returncode", 0)
        self.killed = False
        self.request_path = Path(command[-2])
        self.response_path = Path(command[-1])
        self.request = json.loads(self.request_path.read_text(encoding="utf-8"))

        if not action.get("timeout") and action.get("write_response", True):
            self.response_path.write_text(
                json.dumps(action["response"]),
                encoding="utf-8",
            )

    def communicate(self, timeout=None):
        if self.action.get("timeout") and not self.killed:
            raise subprocess.TimeoutExpired(self.command, timeout)
        return "", self.action.get("stderr", "")

    def kill(self):
        self.killed = True
        self.returncode = -9


def install_fake_popen(monkeypatch, actions: list[dict]) -> list[FakeProcess]:
    calls: list[FakeProcess] = []

    def fake_popen(command, stdout=None, stderr=None, text=None):
        process = FakeProcess(command, actions.pop(0))
        calls.append(process)
        return process

    monkeypatch.setattr(subprocess, "Popen", fake_popen)
    return calls


def test_classifies_recoverable_com_errors():
    assert classify_com_error("The RPC server is unavailable (-2147023174)") == "rpc_unavailable"
    assert classify_com_error("Call was rejected by callee (-2147418111)") == "call_rejected"
    assert classify_com_error("Object invoked has disconnected") == "object_disconnected"
    assert classify_com_error("MCP error -32000: Connection closed") == "connection_closed"
    assert classify_com_error("Operation timed out after 45 seconds") == "timeout"


def test_run_vba_success_uses_worker(monkeypatch):
    actions = [
        {
            "response": {
                "success": True,
                "operation": "run_vba",
                "phase": "run_vba",
                "result": '{"success": true, "result": 42}',
            }
        }
    ]
    calls = install_fake_popen(monkeypatch, actions)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code="MCP_TempFunction = 42",
        addin_path="C:\\addin.accda",
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert result["result"] == '{"success": true, "result": 42}'
    assert calls[0].request["operation"] == "run_vba"
    assert calls[0].request["code"] == "MCP_TempFunction = 42"


def test_run_vba_timeout_kills_only_worker(monkeypatch):
    actions = [{"timeout": True, "stderr": "Access did not respond"}]
    calls = install_fake_popen(monkeypatch, actions)

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
    assert calls[0].killed is True
    assert get_recovery_manager().get_state("C:\\db.accdb").status == "timed_out"


def test_next_call_probes_after_prior_timeout(monkeypatch):
    get_recovery_manager().mark_failure(
        "C:\\db.accdb",
        "VBA worker timed out after 45 seconds",
        "timeout",
        timed_out=True,
    )
    actions = [
        {
            "response": {
                "success": True,
                "operation": "probe",
                "phase": "load_addin",
                "result": "ok",
            }
        },
        {
            "response": {
                "success": True,
                "operation": "run_vba",
                "phase": "run_vba",
                "result": "done",
            }
        },
    ]
    calls = install_fake_popen(monkeypatch, actions)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code='MCP_TempFunction = "done"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert [call.request["operation"] for call in calls] == ["probe", "run_vba"]
    assert get_recovery_manager().get_state("C:\\db.accdb").status == "healthy"


def test_predispatch_disconnect_probes_and_retries_once(monkeypatch):
    actions = [
        {
            "returncode": 1,
            "response": {
                "success": False,
                "operation": "run_vba",
                "phase": "connect",
                "error": "The RPC server is unavailable",
                "error_pattern": "rpc_unavailable",
            },
        },
        {
            "response": {
                "success": True,
                "operation": "probe",
                "phase": "load_addin",
                "result": "ok",
            }
        },
        {
            "response": {
                "success": True,
                "operation": "run_vba",
                "phase": "run_vba",
                "result": "retried",
            }
        },
    ]
    calls = install_fake_popen(monkeypatch, actions)

    manager = VBAWorkerManager()
    result = manager.run_vba(
        database_path="C:\\db.accdb",
        code='MCP_TempFunction = "retried"',
        timeout_seconds=1,
    )

    assert result["success"] is True
    assert result["result"] == "retried"
    assert [call.request["operation"] for call in calls] == ["run_vba", "probe", "run_vba"]


def test_run_phase_failure_is_not_retried(monkeypatch):
    actions = [
        {
            "returncode": 1,
            "response": {
                "success": False,
                "operation": "run_vba",
                "phase": "run_vba",
                "error": "Call was rejected by callee",
                "error_pattern": "call_rejected",
            },
        }
    ]
    calls = install_fake_popen(monkeypatch, actions)

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
