"""
Collects VB6 OTSoft output files and copies them to golden file paths.

VB6 writes output to: <input_dir>/FilesFor<filename>/
Text output:  <filename>DraftOutput.txt
HTML output:  ResultsFor<filename>.htm
"""

import logging
import os
import shutil

logger = logging.getLogger(__name__)


def _actual_base_name(input_file: str) -> str:
    """Get the actual filename (without extension) that VB6 uses for output."""
    return os.path.splitext(os.path.basename(input_file))[0]


def _output_dir(input_file: str, repo_root: str) -> str:
    """Derive the VB6 output directory path for a given input file."""
    abs_input = os.path.join(repo_root, input_file)
    input_dir = os.path.dirname(abs_input)
    name = _actual_base_name(input_file)
    return os.path.join(input_dir, f"FilesFor{name}")


def get_apriori_dest(input_file: str, repo_root: str) -> str:
    """
    Get the path where VB6 expects the a priori rankings file.

    VB6 looks for: <input_dir>/FilesFor<name>/<name>apriori.txt
    """
    name = _actual_base_name(input_file)
    return os.path.join(_output_dir(input_file, repo_root), f"{name}apriori.txt")


def collect_output(
    input_file: str,
    golden_file: str,
    output_format: str,
    repo_root: str,
) -> bool:
    """
    Copy VB6 output to the golden file path.

    Returns True if the file was successfully collected, False otherwise.
    """
    name = _actual_base_name(input_file)
    out_dir = _output_dir(input_file, repo_root)

    if output_format == "html":
        source = os.path.join(out_dir, f"ResultsFor{name}.htm")
    else:
        source = os.path.join(out_dir, f"{name}DraftOutput.txt")

    dest = os.path.join(repo_root, golden_file)
    os.makedirs(os.path.dirname(dest), exist_ok=True)

    try:
        shutil.copy2(source, dest)
    except FileNotFoundError:
        logger.warning("Output file not found: %s", source)
        return False

    size = os.path.getsize(dest)
    if size == 0:
        logger.error("Collected file is empty: %s", dest)
        return False

    logger.info("Collected: %s (%d bytes) -> %s", source, size, dest)
    return True


def cleanup_output_dir(input_file: str, repo_root: str):
    """Remove the VB6 output directory for a given input file."""
    shutil.rmtree(_output_dir(input_file, repo_root), ignore_errors=True)
