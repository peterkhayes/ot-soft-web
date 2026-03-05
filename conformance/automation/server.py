"""
HTTP control server for running conformance tests on Windows.

Start once on the Windows machine (in an interactive desktop session),
then send requests from Mac to trigger test runs.

Usage (on Windows):
    python conformance\\automation\\server.py --otsoft-path "C:\\path\\to\\OTSoft.exe"

Endpoints:
    GET  /status          — check if server is running
    POST /run             — run tests (body: JSON with optional "filter", "no_cleanup", "verbose")
    GET  /results         — get results from the last run
"""

import argparse
import json
import logging
import os
import subprocess
import sys
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

logger = logging.getLogger(__name__)

# Server state
_state = {
    "running": False,
    "last_results": None,
    "last_output": None,
}
_lock = threading.Lock()


class Handler(BaseHTTPRequestHandler):
    otsoft_path = ""
    repo_root = ""

    def do_GET(self):
        if self.path == "/status":
            self._json_response({"status": "ok", "running": _state["running"]})
        elif self.path == "/results":
            self._json_response({
                "running": _state["running"],
                "results": _state["last_results"],
                "output": _state["last_output"],
            })
        else:
            self._json_response({"error": "not found"}, status=404)

    def do_POST(self):
        if self.path != "/run":
            self._json_response({"error": "not found"}, status=404)
            return

        with _lock:
            if _state["running"]:
                self._json_response({"error": "a run is already in progress"}, status=409)
                return
            _state["running"] = True

        # Parse request body
        content_length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(content_length)) if content_length > 0 else {}

        filter_pattern = body.get("filter")
        no_cleanup = body.get("no_cleanup", False)
        verbose = body.get("verbose", False)

        self._json_response({"status": "started", "filter": filter_pattern})

        # Run in background thread so the HTTP response returns immediately
        threading.Thread(
            target=self._run_tests,
            args=(filter_pattern, no_cleanup, verbose),
            daemon=True,
        ).start()

    def _run_tests(self, filter_pattern, no_cleanup, verbose):
        try:
            cmd = [
                sys.executable,
                os.path.join(Handler.repo_root, "conformance", "automation", "run_tests.py"),
                "--otsoft-path", Handler.otsoft_path,
                "--repo-root", Handler.repo_root,
            ]
            if filter_pattern:
                cmd.extend(["--filter", filter_pattern])
            if no_cleanup:
                cmd.append("--no-cleanup")
            if verbose:
                cmd.append("-v")

            logger.info("Running: %s", " ".join(cmd))
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=600,
            )
            output = result.stdout + result.stderr
            logger.info("Run completed (exit code %d)", result.returncode)

            with _lock:
                _state["last_output"] = output
                _state["last_results"] = {
                    "exit_code": result.returncode,
                    "success": result.returncode == 0,
                }
        except Exception as e:
            logger.error("Run failed: %s", e)
            with _lock:
                _state["last_output"] = str(e)
                _state["last_results"] = {"exit_code": -1, "success": False}
        finally:
            with _lock:
                _state["running"] = False

    def _json_response(self, data, status=200):
        body = json.dumps(data, indent=2).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        logger.info(format, *args)


def find_repo_root():
    """Walk up from this script to find the repo root."""
    path = os.path.dirname(os.path.abspath(__file__))
    while path != os.path.dirname(path):
        if os.path.isdir(os.path.join(path, ".git")):
            return path
        path = os.path.dirname(path)
    return os.getcwd()


def main():
    parser = argparse.ArgumentParser(description="Conformance test control server.")
    parser.add_argument("--otsoft-path", required=True, help="Path to OTSoft.exe")
    parser.add_argument("--port", type=int, default=8377, help="Port to listen on (default: 8377)")
    parser.add_argument("--repo-root", default=None, help="Repo root (auto-detected if omitted)")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    Handler.otsoft_path = args.otsoft_path
    Handler.repo_root = args.repo_root or find_repo_root()

    if not os.path.isfile(args.otsoft_path):
        logger.error("OTSoft.exe not found: %s", args.otsoft_path)
        sys.exit(1)

    server = HTTPServer(("0.0.0.0", args.port), Handler)
    logger.info("Control server listening on port %d", args.port)
    logger.info("OTSoft path: %s", Handler.otsoft_path)
    logger.info("Repo root: %s", Handler.repo_root)
    logger.info("Ready for requests from Mac.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        logger.info("Shutting down.")
        server.shutdown()


if __name__ == "__main__":
    main()
