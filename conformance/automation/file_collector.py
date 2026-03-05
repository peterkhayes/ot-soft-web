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


def _output_dir(input_file: str, display_name: str, repo_root: str) -> str:
    """Derive the VB6 output directory path for a given input file."""
    abs_input = os.path.join(repo_root, input_file)
    input_dir = os.path.dirname(abs_input)
    base_name = os.path.splitext(display_name)[0]
    return os.path.join(input_dir, f"FilesFor{base_name}")


def _base_name(display_name: str) -> str:
    """Extract the base name (without extension) from a VB6 display name."""
    return os.path.splitext(display_name)[0]


def get_apriori_dest(input_file: str, display_name: str, repo_root: str) -> str:
    """
    Get the path where VB6 expects the a priori rankings file.

    VB6 looks for: <input_dir>/FilesFor<name>/<name>apriori.txt
    """
    name = _base_name(display_name)
    return os.path.join(_output_dir(input_file, display_name, repo_root), f"{name}apriori.txt")


def collect_output(
    input_file: str,
    display_name: str,
    golden_file: str,
    output_format: str,
    repo_root: str,
) -> bool:
    """
    Copy VB6 output to the golden file path.

    Returns True if the file was successfully collected, False otherwise.
    """
    name = _base_name(display_name)
    out_dir = _output_dir(input_file, display_name, repo_root)

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


def cleanup_output_dir(input_file: str, display_name: str, repo_root: str):
    """Remove the VB6 output directory for a given input file."""
    shutil.rmtree(
        _output_dir(input_file, display_name, repo_root),
        ignore_errors=True,
    )
