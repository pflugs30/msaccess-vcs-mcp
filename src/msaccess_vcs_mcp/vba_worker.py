"""Child process entry point for isolated Access VBA execution."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from .access_com.connection import AccessConnection
from .addin_integration import VCSAddinIntegration
from .com_recovery import classify_com_error
from .config import get_config
from .security import validate_database_path


def execute_worker_request(request: dict[str, Any]) -> dict[str, Any]:
    """Execute one VBA worker request and return a JSON-serializable result."""
    operation = request.get("operation")
    database_path = request.get("database_path")
    addin_path = request.get("addin_path")
    code = request.get("code", "")
    context = {"phase": "start"}

    try:
        if not database_path:
            raise ValueError("database_path is required")
        if operation not in {"run_vba", "probe"}:
            raise ValueError(f"Unsupported worker operation: {operation}")

        context["phase"] = "validate_database"
        db_path = validate_database_path(str(database_path))

        if not addin_path:
            context["phase"] = "load_config"
            addin_path = get_config().get("ACCESS_VCS_ADDIN_PATH")

        context["phase"] = "connect"
        with AccessConnection(str(db_path)) as conn:
            app, _db = conn.connect()

            context["phase"] = "load_addin"
            addin = VCSAddinIntegration(addin_path)
            addin.load_addin(app, db_path=str(db_path))

            if operation == "probe":
                return {
                    "success": True,
                    "operation": operation,
                    "phase": context["phase"],
                    "result": "ok",
                }

            context["phase"] = "run_vba"
            result = addin.call_sync("RunVBA", code)
            return {
                "success": True,
                "operation": operation,
                "phase": context["phase"],
                "result": result,
            }
    except Exception as exc:
        return {
            "success": False,
            "operation": operation,
            "phase": context["phase"],
            "error": str(exc),
            "error_type": type(exc).__name__,
            "error_pattern": classify_com_error(exc),
        }


def _write_response(output_path: str, response: dict[str, Any]) -> None:
    Path(output_path).write_text(
        json.dumps(response, default=str),
        encoding="utf-8",
    )


def main(argv: list[str] | None = None) -> int:
    """Run a worker request from a request JSON file into a response JSON file."""
    args = list(sys.argv[1:] if argv is None else argv)
    if len(args) != 2:
        print("Usage: python -m msaccess_vcs_mcp.vba_worker <request.json> <response.json>", file=sys.stderr)
        return 2

    request_path, output_path = args
    try:
        request = json.loads(Path(request_path).read_text(encoding="utf-8"))
    except Exception as exc:
        _write_response(
            output_path,
            {
                "success": False,
                "operation": "unknown",
                "phase": "read_request",
                "error": str(exc),
                "error_type": type(exc).__name__,
                "error_pattern": "worker_request_error",
            },
        )
        return 2

    response = execute_worker_request(request)
    _write_response(output_path, response)
    return 0 if response.get("success") else 1


if __name__ == "__main__":
    raise SystemExit(main())
