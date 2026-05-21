"""Tests for vcs_run_tests tool."""

import json
import os
from unittest.mock import Mock, MagicMock, patch, call

import pytest


SAMPLE_RESULTS_ALL_PASS = {
    "runAt": "2026-05-21 14:30:00",
    "databasePath": "C:\\db.accdb",
    "addinVersion": "4.1.0",
    "durationMs": 350,
    "summary": {
        "subs": 5,
        "assertions": 12,
        "passed": 12,
        "failed": 0,
        "errored": 0,
        "empty": 0,
    },
    "tests": {
        "modTestFoo.TestOne": {
            "status": "PASSED",
            "durationMs": 10,
            "tags": ["unit"],
            "assertions": [{"seq": 1, "passed": True, "context": "works"}],
        },
    },
}

SAMPLE_RESULTS_WITH_FAILURE = {
    "runAt": "2026-05-21 14:31:00",
    "databasePath": "C:\\db.accdb",
    "addinVersion": "4.1.0",
    "durationMs": 200,
    "summary": {
        "subs": 3,
        "assertions": 6,
        "passed": 4,
        "failed": 2,
        "errored": 0,
        "empty": 0,
    },
    "tests": {
        "modTestFoo.TestOne": {
            "status": "PASSED",
            "durationMs": 10,
            "tags": [],
            "assertions": [{"seq": 1, "passed": True}],
        },
        "modTestFoo.TestTwo": {
            "status": "FAILED",
            "durationMs": 15,
            "tags": [],
            "assertions": [{"seq": 1, "passed": False, "context": "expected 1"}],
        },
    },
}

SAMPLE_RESULTS_WITH_ERROR = {
    "runAt": "2026-05-21 14:32:00",
    "databasePath": "C:\\db.accdb",
    "addinVersion": "4.1.0",
    "durationMs": 100,
    "summary": {
        "subs": 2,
        "assertions": 1,
        "passed": 1,
        "failed": 0,
        "errored": 1,
        "empty": 0,
    },
    "tests": {
        "modTestFoo.TestOk": {
            "status": "PASSED",
            "durationMs": 5,
            "tags": [],
            "assertions": [{"seq": 1, "passed": True}],
        },
        "modTestFoo.TestBoom": {
            "status": "ERRORED",
            "durationMs": 3,
            "tags": [],
            "errorMessage": "Error 91: Object variable not set",
            "assertions": [],
        },
    },
}

SAMPLE_RESULTS_EMPTY = {
    "runAt": "2026-05-21 14:33:00",
    "databasePath": "C:\\db.accdb",
    "addinVersion": "4.1.0",
    "durationMs": 50,
    "summary": {
        "subs": 0,
        "assertions": 0,
        "passed": 0,
        "failed": 0,
        "errored": 0,
        "empty": 0,
    },
    "tests": {},
}


def _build_mocks(tmp_path, call_sync_return=None):
    """Build the mock objects needed for vcs_run_tests tests.

    Returns (mock_app, mock_addin, addin_path) where call_sync is
    pre-configured with the given return value.
    """
    addin_file = tmp_path / "Version Control.accda"
    addin_file.touch()

    mock_app = Mock()
    mock_db = Mock()
    mock_db.Name = str(tmp_path / "test.accdb")

    mock_conn = MagicMock()
    mock_conn.connect.return_value = (mock_app, mock_db)
    mock_conn.__enter__ = Mock(return_value=mock_conn)
    mock_conn.__exit__ = Mock(return_value=False)

    mock_addin = Mock()
    mock_addin.addin_path = str(addin_file)
    mock_addin.load_addin.return_value = True
    mock_addin.call_sync.return_value = call_sync_return

    return mock_app, mock_conn, mock_addin


@pytest.fixture(autouse=True)
def _patch_tool_infra(monkeypatch, tmp_path_factory):
    """Disable the @vcs_tool decorator's config/logging layers for unit tests."""
    diag_dir = tmp_path_factory.mktemp("diag")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(diag_dir))


def _call_run_tests(tmp_path, call_sync_return, filter_value=None):
    """Call vcs_run_tests with fully mocked COM layer.

    Patches AccessConnection, VCSAddinIntegration, validate_database_path,
    and get_config so the tool body runs against mock objects.
    """
    mock_app, mock_conn, mock_addin = _build_mocks(
        tmp_path, call_sync_return=call_sync_return
    )

    db_path = str(tmp_path / "test.accdb")
    (tmp_path / "test.accdb").touch()

    with (
        patch(
            "msaccess_vcs_mcp.tools.AccessConnection",
            return_value=mock_conn,
        ),
        patch(
            "msaccess_vcs_mcp.tools.VCSAddinIntegration",
            return_value=mock_addin,
        ),
        patch(
            "msaccess_vcs_mcp.tools.validate_database_path",
            return_value=tmp_path / "test.accdb",
        ),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
    ):
        from msaccess_vcs_mcp.tools import vcs_run_tests

        # vcs_run_tests is wrapped by @vcs_tool into an async coroutine;
        # unwrap to the sync body via __wrapped__
        sync_fn = vcs_run_tests
        while hasattr(sync_fn, "__wrapped__"):
            sync_fn = sync_fn.__wrapped__

        result = sync_fn(db_path, filter=filter_value)

    return result, mock_app, mock_addin


class TestRunTestsSuccess:
    """Tests for successful test runs."""

    def test_all_tests_pass(self, tmp_path):
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
        )

        assert result["success"] is True
        assert result["summary"]["subs"] == 5
        assert result["summary"]["failed"] == 0

    def test_no_filter_clears_default_test_filter(self, tmp_path):
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
            filter_value=None,
        )

        mock_addin.call_sync.assert_any_call("SetOption", "DefaultTestFilter", "")
        mock_addin.call_sync.assert_any_call("RunFilteredTests")

    def test_filter_sets_default_test_filter(self, tmp_path):
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
            filter_value="modTestFoo,-slow",
        )

        mock_addin.call_sync.assert_any_call(
            "SetOption", "DefaultTestFilter", "modTestFoo,-slow"
        )
        mock_addin.call_sync.assert_any_call("RunFilteredTests")

    def test_silent_mode_set(self, tmp_path):
        """SetInteractionMode(1) is called via app.Run before tests."""
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
        )

        addin_path = mock_addin.addin_path
        addin_lib = os.path.splitext(os.path.abspath(addin_path))[0]
        mock_app.Run.assert_called_once_with(
            f"{addin_lib}.SetInteractionMode", 1
        )


class TestRunTestsFailure:
    """Tests for test runs that report failures or errors."""

    def test_failed_assertions(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_WITH_FAILURE),
        )

        assert result["success"] is False
        assert result["summary"]["failed"] == 2

    def test_errored_tests(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_WITH_ERROR),
        )

        assert result["success"] is False
        assert result["summary"]["errored"] == 1

    def test_no_tests_found(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_EMPTY),
        )

        assert result["success"] is False
        assert result["summary"]["subs"] == 0


class TestRunTestsEdgeCases:
    """Tests for edge cases and error handling."""

    def test_empty_result(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return="",
        )

        assert result["success"] is False
        assert "no results" in result["error"].lower()

    def test_none_result(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=None,
        )

        assert result["success"] is False
        assert "no results" in result["error"].lower()

    def test_invalid_json_result(self, tmp_path):
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return="not valid json {{{",
        )

        assert result["success"] is False
        assert "parse" in result["error"].lower()

    def test_com_exception(self, tmp_path):
        """COM error during call_sync is caught and returned."""
        mock_app, mock_conn, mock_addin = _build_mocks(tmp_path)
        mock_addin.call_sync.side_effect = Exception("COM disconnected")

        db_path = str(tmp_path / "test.accdb")
        (tmp_path / "test.accdb").touch()

        with (
            patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=mock_conn),
            patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=mock_addin),
            patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=tmp_path / "test.accdb"),
            patch("msaccess_vcs_mcp.tools.get_config", return_value={"ACCESS_VCS_ADDIN_PATH": "test"}),
        ):
            from msaccess_vcs_mcp.tools import vcs_run_tests

            sync_fn = vcs_run_tests
            while hasattr(sync_fn, "__wrapped__"):
                sync_fn = sync_fn.__wrapped__

            result = sync_fn(db_path)

        assert result["success"] is False
        assert "COM disconnected" in result["error"]

    def test_dict_result_passthrough(self, tmp_path):
        """When add-in returns a dict directly (not JSON string), it's handled."""
        result, _, _ = _call_run_tests(
            tmp_path,
            call_sync_return=SAMPLE_RESULTS_ALL_PASS,
        )

        assert result["success"] is True
        assert result["summary"]["subs"] == 5


class TestRunTestsCallOrder:
    """Verify the COM call sequence."""

    def test_call_order(self, tmp_path):
        """SetInteractionMode (app.Run) -> SetOption -> RunFilteredTests (call_sync)."""
        result, mock_app, mock_addin = _call_run_tests(
            tmp_path,
            call_sync_return=json.dumps(SAMPLE_RESULTS_ALL_PASS),
            filter_value="SQL,-slow",
        )

        # SetInteractionMode goes through app.Run, not call_sync
        addin_lib = os.path.splitext(os.path.abspath(mock_addin.addin_path))[0]
        mock_app.Run.assert_called_once_with(
            f"{addin_lib}.SetInteractionMode", 1
        )

        # SetOption + RunFilteredTests go through call_sync
        calls = mock_addin.call_sync.call_args_list
        assert len(calls) == 2
        assert calls[0] == call("SetOption", "DefaultTestFilter", "SQL,-slow")
        assert calls[1] == call("RunFilteredTests")

        mock_addin.load_addin.assert_called_once()
