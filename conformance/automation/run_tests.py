#!/usr/bin/env python3
"""
CLI entry point for collecting VB6 OTSoft golden files.

Usage:
    python run_tests.py --otsoft-path "C:\\path\\to\\OTSoft.exe"
    python run_tests.py --otsoft-path "C:\\path\\to\\OTSoft.exe" --filter "TinyIllustrativeFile_rcd.*"
    python run_tests.py --otsoft-path "C:\\path\\to\\OTSoft.exe" --no-cleanup
"""

import argparse
import logging
import os
import subprocess
import sys


def find_repo_root() -> str:
    """Detect the git repository root."""
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--show-toplevel"],
            capture_output=True,
            text=True,
            check=True,
        )
        return result.stdout.strip()
    except (subprocess.CalledProcessError, FileNotFoundError):
        # Fall back to walking up from this script's location
        path = os.path.dirname(os.path.abspath(__file__))
        while path != os.path.dirname(path):
            if os.path.isdir(os.path.join(path, ".git")):
                return path
            path = os.path.dirname(path)
        return os.getcwd()


def main():
    parser = argparse.ArgumentParser(
        description="Collect VB6 OTSoft golden files via pywinauto automation."
    )
    parser.add_argument(
        "--otsoft-path",
        required=True,
        help='Path to OTSoft.exe (e.g. "C:\\OTSoft\\OTSoft.exe")',
    )
    parser.add_argument(
        "--filter",
        default=None,
        help="Regex pattern to filter test case IDs",
    )
    parser.add_argument(
        "--repo-root",
        default=None,
        help="Path to the ot-soft repository root (auto-detected if omitted)",
    )
    parser.add_argument(
        "--no-cleanup",
        action="store_true",
        help="Keep VB6 output folders after collecting golden files",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable debug logging",
    )
    args = parser.parse_args()

    # Configure logging
    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    repo_root = args.repo_root or find_repo_root()
    logging.info("Repository root: %s", repo_root)

    if not os.path.isfile(args.otsoft_path):
        logging.error("OTSoft.exe not found: %s", args.otsoft_path)
        sys.exit(1)

    # Import here so --help works without pywinauto installed
    from otsoft_driver import OTSoftDriver
    from manifest_runner import run_all

    driver = OTSoftDriver(args.otsoft_path)

    try:
        driver.launch()
        results = run_all(
            driver=driver,
            repo_root=repo_root,
            filter_pattern=args.filter,
            no_cleanup=args.no_cleanup,
        )
    except KeyboardInterrupt:
        logging.info("Interrupted by user")
        results = {"passed": [], "failed": [], "errors": ["interrupted"], "details": {}}
    finally:
        driver.close()

    # Print summary
    print("\n" + "=" * 60)
    print("RESULTS SUMMARY")
    print("=" * 60)
    print(f"  Passed:  {len(results['passed'])}")
    print(f"  Failed:  {len(results['failed'])}")
    print(f"  Errors:  {len(results['errors'])}")

    details = results.get("details", {})

    if results["failed"]:
        print("\nFailed cases:")
        for case_id in results["failed"]:
            detail = details.get(case_id)
            suffix = f" — {detail}" if detail else ""
            print(f"  - {case_id}{suffix}")

    if results["errors"]:
        print("\nError cases:")
        for case_id in results["errors"]:
            detail = details.get(case_id)
            suffix = f" — {detail}" if detail else ""
            print(f"  - {case_id}{suffix}")

    total = len(results["passed"]) + len(results["failed"]) + len(results["errors"])
    if total > 0:
        pct = len(results["passed"]) / total * 100
        print(f"\nPass rate: {pct:.0f}% ({len(results['passed'])}/{total})")

    sys.exit(0 if not results["failed"] and not results["errors"] else 1)


if __name__ == "__main__":
    main()
