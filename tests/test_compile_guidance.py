"""Tests for compile-failure agent_guidance in vcs_compile_vba and vcs_check_vba_compiled."""

from contextlib import contextmanager
from unittest.mock import MagicMock, patch

import pytest


def _unwrap_sync(tool_fn):
    sync_fn = tool_fn
    while hasattr(sync_fn, "__wrapped__"):
        sync_fn = sync_fn.__wrapped__
    return sync_fn


@contextmanager
def _patch_compile_infra(tmp_path, *, is_compiled=None, compile_result=None):
    mock_conn = MagicMock()
    mock_conn.__enter__ = MagicMock(return_value=mock_conn)
    mock_conn.__exit__ = MagicMock(return_value=False)
    mock_conn.connect.return_value = (MagicMock(), MagicMock())

    mock_addin = MagicMock()
    if is_compiled is not None:
        mock_addin.call_sync.return_value = is_compiled
    if compile_result is not None:
        mock_addin.call_sync.return_value = compile_result

    db_path = tmp_path / "test.accdb"
    db_path.touch()

    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=db_path),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
    ):
        yield str(db_path), mock_addin


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    diag_dir = tmp_path_factory.mktemp("diag")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))


class TestCompileGuidance:
    def test_compile_failure_includes_agent_guidance(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_compile_vba

        with _patch_compile_infra(tmp_path, compile_result=False) as (db_path, _):
            result = _unwrap_sync(vcs_compile_vba)(db_path)

        assert result["success"] is False
        assert "agent_guidance" in result
        assert "VBE" in result["agent_guidance"]
        assert "Do not guess" in result["agent_guidance"]

    def test_compile_success_omits_agent_guidance(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_compile_vba

        with _patch_compile_infra(tmp_path, compile_result=True) as (db_path, _):
            result = _unwrap_sync(vcs_compile_vba)(db_path)

        assert result["success"] is True
        assert "agent_guidance" not in result

    def test_compile_exception_includes_agent_guidance(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_compile_vba

        with _patch_compile_infra(tmp_path, compile_result=True) as (db_path, mock_addin):
            mock_addin.load_addin.side_effect = RuntimeError("COM failed")
            result = _unwrap_sync(vcs_compile_vba)(db_path)

        assert result["success"] is False
        assert result["error"] == "COM failed"
        assert "agent_guidance" in result

    def test_not_compiled_check_includes_agent_guidance(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_check_vba_compiled

        with _patch_compile_infra(tmp_path, is_compiled=False) as (db_path, _):
            result = _unwrap_sync(vcs_check_vba_compiled)(db_path)

        assert result["success"] is True
        assert result["compiled"] is False
        assert "agent_guidance" in result

    def test_compiled_check_omits_agent_guidance(self, tmp_path):
        from msaccess_vcs_mcp.tools import vcs_check_vba_compiled

        with _patch_compile_infra(tmp_path, is_compiled=True) as (db_path, _):
            result = _unwrap_sync(vcs_check_vba_compiled)(db_path)

        assert result["success"] is True
        assert result["compiled"] is True
        assert "agent_guidance" not in result
