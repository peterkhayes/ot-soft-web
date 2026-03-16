"""
Reads conformance/manifest.json and orchestrates test case execution.

Groups cases by input file to minimize file-open operations,
then drives OTSoftDriver for each case.
"""

import json
import logging
import os
import re
import shutil
from collections import defaultdict

from otsoft_driver import OTSoftDriver, DialogDismissedError, StepTimeoutError
from file_collector import collect_output, get_apriori_dest, cleanup_output_dir

logger = logging.getLogger(__name__)


def load_manifest(repo_root: str) -> list[dict]:
    """Load and return test cases from manifest.json."""
    manifest_path = os.path.join(repo_root, "conformance", "manifest.json")
    with open(manifest_path) as f:
        data = json.load(f)
    return data["cases"]


def filter_cases(cases: list[dict], pattern: str | None) -> list[dict]:
    """Filter test cases by regex pattern on case ID."""
    if not pattern:
        return cases
    regex = re.compile(pattern)
    filtered = [c for c in cases if regex.search(c["id"])]
    logger.info("Filtered %d/%d cases matching '%s'", len(filtered), len(cases), pattern)
    return filtered


def group_by_input(cases: list[dict]) -> dict[str, list[dict]]:
    """Group test cases by input_file to minimize file-open operations."""
    groups = defaultdict(list)
    for case in cases:
        groups[case["input_file"]].append(case)
    return dict(groups)


def run_case(driver: OTSoftDriver, case: dict, repo_root: str) -> bool:
    """
    Execute a single test case and collect the output.

    Returns True if the golden file was successfully collected.
    """
    case_id = case["id"]
    algorithm = case["algorithm"]
    params = case.get("params", {})
    output_format = case.get("format", "text")
    apriori_file = case.get("apriori_file")
    input_file = case["input_file"]

    logger.info("=" * 60)
    logger.info("Running case: %s", case_id)
    logger.info("  algorithm=%s, format=%s", algorithm, output_format)
    logger.info("  params=%s", params)

    # No need to restore_defaults() — we relaunch OTSoft between cases,
    # so the app starts fresh. We explicitly set all needed options below.

    # Handle a priori rankings
    if apriori_file:
        abs_apriori = os.path.join(repo_root, apriori_file)
        apriori_dest = get_apriori_dest(input_file, repo_root)
        os.makedirs(os.path.dirname(apriori_dest), exist_ok=True)
        shutil.copy2(abs_apriori, apriori_dest)
        logger.info("Copied a priori file: %s -> %s", abs_apriori, apriori_dest)
        driver.enable_apriori_rankings()
    else:
        driver.disable_apriori_rankings()

    # Configure and run based on algorithm type
    if algorithm in ("rcd", "bcd", "lfcd"):
        driver.select_classical_ot()
        driver.set_algorithm_variant(algorithm, params)
        driver.configure_ranking_options(params)
        driver.run_rank()

    elif algorithm == "maxent":
        driver.run_maxent(params)

    elif algorithm == "factorial_typology":
        driver.select_classical_ot()
        driver.configure_factorial_typology(params)
        driver.run_factorial_typology()

    else:
        logger.error("Unknown algorithm: %s", algorithm)
        return False

    # Collect the output
    success = collect_output(
        input_file=input_file,
        golden_file=case["golden_file"],
        output_format=output_format,
        repo_root=repo_root,
    )

    if success:
        logger.info("SUCCESS: %s -> %s", case_id, case["golden_file"])
    else:
        logger.error("FAILED to collect output for: %s", case_id)

    return success


def run_all(
    driver: OTSoftDriver,
    repo_root: str,
    filter_pattern: str | None = None,
    no_cleanup: bool = False,
) -> dict:
    """
    Run all (or filtered) test cases from the manifest.

    Returns a dict with 'passed', 'failed', and 'errors' lists.
    """
    cases = load_manifest(repo_root)
    cases = filter_cases(cases, filter_pattern)

    if not cases:
        logger.warning("No test cases to run")
        return {"passed": [], "failed": [], "errors": []}

    groups = group_by_input(cases)
    results = {"passed": [], "failed": [], "errors": [], "details": {}}

    for input_file, group_cases in groups.items():
        logger.info("\n" + "=" * 70)
        logger.info("Input file: %s (%d cases)", input_file, len(group_cases))
        logger.info("=" * 70)

        # Start MsgBox dismisser early so it catches dialogs during file open
        driver.start_msgbox_dismisser()
        try:
            # Open the input file (relaunches OTSoft with this file)
            abs_input = os.path.join(repo_root, input_file)
            try:
                driver.open_file(abs_input)
            except Exception as e:
                logger.error("Failed to open file %s: %s", input_file, e)
                for case in group_cases:
                    results["errors"].append(case["id"])
                    results["details"][case["id"]] = f"File open failed: {e}"
                continue
            for i, case in enumerate(group_cases):
                try:
                    success = run_case(driver, case, repo_root)
                    if success:
                        results["passed"].append(case["id"])
                    else:
                        results["failed"].append(case["id"])
                except DialogDismissedError as e:
                    logger.error(
                        "Dialog dismissed during %s: %s",
                        case["id"], e.messages,
                    )
                    results["errors"].append(case["id"])
                    results["details"][case["id"]] = (
                        f"VB6 dialog: {'; '.join(e.messages)}"
                    )
                except (TimeoutError, StepTimeoutError) as e:
                    logger.error("Timeout running %s: %s", case["id"], e)
                    results["errors"].append(case["id"])
                    results["details"][case["id"]] = str(e)
                except Exception as e:
                    logger.error("Error running %s: %s", case["id"], e, exc_info=True)
                    results["errors"].append(case["id"])
                    results["details"][case["id"]] = str(e)

                # Reopen the file between cases to reset VB6 state,
                # but skip after the last case in the group
                if i < len(group_cases) - 1:
                    try:
                        driver.open_file(abs_input)
                    except Exception:
                        pass
        finally:
            driver.stop_msgbox_dismisser()

        # Cleanup VB6 output directories
        if not no_cleanup:
            cleanup_output_dir(input_file, repo_root)

    return results
