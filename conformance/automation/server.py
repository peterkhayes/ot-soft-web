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
    POST /reload          — git pull and restart the server process
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
        if self.path == "/reload":
            self._handle_reload()
            return

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

            # Auto-commit and push golden files after successful runs
            if result.returncode == 0:
                push_output = self._commit_and_push_golden()
                output += "\n" + push_output

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

    def _commit_and_push_golden(self) -> str:
        """Commit any new/changed golden files and push to remote."""
        cwd = Handler.repo_root
        lines = []

        # Stage golden files
        result = subprocess.run(
            ["git", "add", "conformance/golden/"],
            cwd=cwd, capture_output=True, text=True,
        )

        # Check if there's anything to commit
        result = subprocess.run(
            ["git", "diff", "--cached", "--quiet"],
            cwd=cwd, capture_output=True, text=True,
        )
        if result.returncode == 0:
            lines.append("No new golden files to commit.")
            return "\n".join(lines)

        # Commit
        result = subprocess.run(
            ["git", "commit", "-m", "Collect golden files from VB6 OTSoft"],
            cwd=cwd, capture_output=True, text=True,
        )
        lines.append(f"git commit: {result.stdout.strip()}")
        logger.info("git commit: %s", result.stdout.strip())

        # Push
        result = subprocess.run(
            ["git", "push"],
            cwd=cwd, capture_output=True, text=True,
        )
        if result.returncode == 0:
            lines.append("git push: success")
            logger.info("Pushed golden files to remote.")
        else:
            lines.append(f"git push failed: {result.stderr.strip()}")
            logger.error("git push failed: %s", result.stderr.strip())

        return "\n".join(lines)

    def _handle_reload(self):
        """Git pull and restart the server process."""
        with _lock:
            if _state["running"]:
                self._json_response(
                    {"error": "cannot reload while a run is in progress"}, status=409
                )
                return

        # Git pull first
        logger.info("Pulling latest changes...")
        result = subprocess.run(
            ["git", "pull"],
            cwd=Handler.repo_root,
            capture_output=True,
            text=True,
        )
        pull_output = result.stdout.strip()
        logger.info("git pull: %s", pull_output)

        if result.returncode != 0:
            self._json_response({
                "error": "git pull failed",
                "output": result.stderr,
            }, status=500)
            return

        self._json_response({"status": "restarting", "pull": pull_output})

        # Spawn a new server process, then exit this one.
        # Use the absolute path to this script to avoid escaping issues on Windows.
        def restart():
            import time
            time.sleep(0.5)
            script = os.path.abspath(__file__)
            cmd = [sys.executable, script] + sys.argv[1:]
            logger.info("Restarting: %s", cmd)
            subprocess.Popen(cmd, cwd=Handler.repo_root)
            os._exit(0)

        threading.Thread(target=restart, daemon=True).start()

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
