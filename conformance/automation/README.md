# OTSoft Conformance Test Automation

Automates VB6 OTSoft via [pywinauto](https://pywinauto.readthedocs.io/) to collect golden output files for conformance testing.

## Prerequisites

- **Windows machine** with OTSoft 2.3+ (VB6) installed
- Python 3.11+
- The ot-soft repository cloned

## Setup

```powershell
cd conformance\automation
pip install -r requirements.txt
```

## Usage

Run all test cases:

```powershell
python run_tests.py --otsoft-path "C:\path\to\OTSoft.exe"
```

Run a subset (regex on test case IDs):

```powershell
python run_tests.py --otsoft-path "C:\path\to\OTSoft.exe" --filter "TinyIllustrativeFile_rcd.*"
```

Run a single case:

```powershell
python run_tests.py --otsoft-path "C:\path\to\OTSoft.exe" --filter "TinyIllustrativeFile_rcd_defaults$"
```

Keep VB6 output directories after collection:

```powershell
python run_tests.py --otsoft-path "C:\path\to\OTSoft.exe" --no-cleanup
```

Verbose logging:

```powershell
python run_tests.py --otsoft-path "C:\path\to\OTSoft.exe" -v
```

## Workflow

1. **On Mac**: edit code, `git push`
2. **SSH to Windows**: `cd ot-soft && git pull`
3. **Run automation**: `python conformance\automation\run_tests.py --otsoft-path "C:\path\to\OTSoft.exe"`
4. **Commit results**: `git add conformance\golden\ && git commit -m "Collect golden files" && git push`
5. **On Mac**: `git pull`

## Architecture

| File | Purpose |
|------|---------|
| `run_tests.py` | CLI entry point — parses args, launches driver, prints summary |
| `otsoft_driver.py` | `OTSoftDriver` class — all pywinauto interactions with VB6 OTSoft |
| `manifest_runner.py` | Reads `manifest.json`, groups cases by input file, orchestrates execution |
| `file_collector.py` | Locates VB6 output files and copies them to `conformance/golden/` |

## How it works

- **File loading**: OTSoft has no file dialog. The script relaunches OTSoft with the input file as a command-line argument (VB6 reads `Command()` on startup).
- **Algorithm selection**: Clicks radio buttons and toggles menu checkmarks via pywinauto's `win32` backend, matching exact VB6 control names.
- **MsgBox handling**: A background thread polls for VB6 MsgBox dialogs and auto-clicks OK.
- **Output collection**: After each run, copies `<name>DraftOutput.txt` or `ResultsFor<name>.htm` from VB6's output directory to the golden file path.
- **State reset**: Restores default settings and relaunches between test cases for isolation.

## Troubleshooting

- **"Output file not found"**: The algorithm may not have completed. Try `--no-cleanup` to inspect the VB6 output directory, or increase timeouts in `otsoft_driver.py`.
- **Permissions errors**: Run the terminal as Administrator if OTSoft requires elevated access.
- **Wrong control names**: VB6 control names may vary between OTSoft versions. Use Spy++ or pywinauto's `print_control_identifiers()` to inspect the running application.
- **MaxEnt sigma parameter**: For non-default sigma values, you may need to manually edit `ModelParameters.txt` in the VB6 output directory before running. Check the logs for warnings.
