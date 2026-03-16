"""
pywinauto driver for VB6 OTSoft application.

Wraps all GUI interactions: launching, file loading, configuring options,
running algorithms, and collecting output.
"""

import logging
import os
import time
import threading
from typing import Callable

from pywinauto import Application

logger = logging.getLogger(__name__)

# How long to wait for the app to become idle after launching
LAUNCH_TIMEOUT = 10
# How long to wait for a single GUI step (click, menu select, etc.)
STEP_TIMEOUT = 15
# How long to wait for algorithm completion (most finish in seconds)
ALGORITHM_TIMEOUT = 60
# Longer timeout for expensive algorithms (factorial typology)
ALGORITHM_TIMEOUT_LONG = 300
# Polling interval when waiting for output files
POLL_INTERVAL = 0.5


class DialogDismissedError(Exception):
    """Raised when a VB6 dialog was dismissed during an operation.

    Attributes:
        messages: List of dismissed dialog message strings.
    """

    def __init__(self, messages: list[str]):
        self.messages = messages
        summary = "; ".join(messages)
        super().__init__(f"VB6 dialog(s) dismissed during operation: {summary}")


class StepTimeoutError(TimeoutError):
    """Raised when a GUI step (click, menu select, etc.) hangs."""
    pass


def _run_with_timeout(fn, *, label: str, timeout: float = STEP_TIMEOUT):
    """
    Run a blocking function in a thread with a timeout.

    Raises StepTimeoutError if the function doesn't return in time.
    This is needed because pywinauto calls are synchronous and can hang
    indefinitely when a modal dialog blocks the target window.
    """
    result = [None]
    error = [None]

    def wrapper():
        try:
            result[0] = fn()
        except Exception as e:
            error[0] = e

    thread = threading.Thread(target=wrapper, daemon=True)
    thread.start()
    thread.join(timeout=timeout)

    if thread.is_alive():
        raise StepTimeoutError(
            f"{label} did not complete within {timeout}s — "
            f"likely blocked by a modal dialog"
        )
    if error[0] is not None:
        raise error[0]
    return result[0]


# Menu paths used in multiple places
APRIORI_MENU_PATH = (
    "&A Priori Rankings",
    "Rank constraints constrained by a priori rankings",
)


_DISMISS_BUTTON_TITLES = {"ok", "end", "abort", "close", "yes"}


_DISMISS_BUTTON_CLASSES = ("Button", "ThunderRT6CommandButton")


def _dismiss_msgboxes(app) -> list[str]:
    """
    Find and dismiss any visible MsgBox dialogs belonging to the app.

    Matches dismiss buttons case-insensitively ("OK", "Ok", "End", etc.)
    and handles VB6 runtime-error dialogs ("End" button) in addition to
    ordinary MsgBox ("OK") dialogs.  Searches both standard Windows
    Button controls and VB6 ThunderRT6CommandButton controls.

    Returns a list of dismissed message strings (title + body text).
    """
    dismissed = []
    try:
        for win in app.windows():
            try:
                if not win.is_visible():
                    continue
                # Skip the main OTSoft window (it has buttons too)
                if win.rectangle().width() > 400 and win.rectangle().height() > 400:
                    continue
                # Find the first button whose title matches a dismiss action,
                # searching both standard and VB6 button classes.
                dismiss_btn = None
                for cls in _DISMISS_BUTTON_CLASSES:
                    for btn in win.children(class_name=cls):
                        if btn.window_text().strip().lower() in _DISMISS_BUTTON_TITLES:
                            dismiss_btn = btn
                            break
                    if dismiss_btn is not None:
                        break
                if dismiss_btn is None:
                    continue
                title = win.window_text()
                btn_label = dismiss_btn.window_text().strip()
                body_parts = [
                    t.strip()
                    for t in win.texts()
                    if t.strip() and t.strip() != title and t.strip() != btn_label
                ]
                body = " ".join(body_parts) if body_parts else "(no body text)"
                message = f"{title}: {body}"
                logger.warning("MsgBox dismissed — %s", message)
                dismissed.append(message)
                dismiss_btn.click()
                time.sleep(0.3)
            except Exception:
                pass
    except Exception:
        pass
    return dismissed


class MsgBoxDismisser:
    """Background thread that auto-dismisses known VB6 message boxes."""

    def __init__(self, get_app):
        """
        Args:
            get_app: Callable returning the current pywinauto Application.
                     Called on every poll so that a fresh app reference is
                     used after OTSoft is relaunched between test cases.
        """
        self._get_app = get_app
        self._stop = threading.Event()
        self._thread = None
        self._dismissed = []

    def start(self):
        # Guard against double-start: stop existing thread first
        if self._thread and self._thread.is_alive():
            self.stop()
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=2)

    @property
    def dismissed(self):
        return list(self._dismissed)

    def _run(self):
        while not self._stop.is_set():
            self._dismissed.extend(_dismiss_msgboxes(self._get_app()))
            self._stop.wait(0.3)


class OTSoftDriver:
    """Drives VB6 OTSoft via pywinauto."""

    def __init__(self, otsoft_path: str):
        self.otsoft_path = otsoft_path
        self.app = None
        self.main_win = None
        self._msgbox_dismisser = None

    def launch(self):
        """Launch OTSoft.exe and connect to the main window."""
        logger.info("Launching OTSoft: %s", self.otsoft_path)
        self._start_app(self.otsoft_path)

    def _start_app(self, cmd: str):
        """Start OTSoft and wait for the main window to be ready."""
        self.app = Application(backend="win32").start(cmd)
        # VB6 OTSoft may show multiple windows on startup (e.g. splash/about).
        # Wait a moment for them to appear, then dismiss any dialogs and find
        # the main window (the one with a menu bar).
        time.sleep(2)
        _dismiss_msgboxes(self.app)
        self._find_main_window()

    def _find_main_window(self):
        """Find the main OTSoft window among potentially multiple windows."""
        # app.windows() returns DialogWrapper objects; we need to identify
        # the right one, then use app.window() with its handle for a proper
        # WindowSpecification.
        dialogs = self.app.windows(title_re="OTSoft.*")
        logger.info("Found %d OTSoft windows", len(dialogs))
        for d in dialogs:
            logger.info("  Window: %r (%dx%d)", d.window_text(),
                        d.rectangle().width(), d.rectangle().height())

        # Find the main window — the one with actual size (ignore 0x0 helper windows)
        main_dialog = max(
            dialogs,
            key=lambda w: w.rectangle().width() * w.rectangle().height(),
        )

        # Use app.window() with the handle to get a proper WindowSpecification
        self.main_win = self.app.window(handle=main_dialog.handle)
        self.main_win.wait("ready", timeout=LAUNCH_TIMEOUT)
        logger.info("OTSoft main window ready: %s", self.main_win.window_text())

    def close(self):
        """Close OTSoft."""
        self.stop_msgbox_dismisser()
        if self.main_win:
            try:
                self.main_win.close()
            except Exception:
                pass
        # Ensure the process is terminated
        if self.app:
            try:
                self.app.kill()
            except Exception:
                pass

    def start_msgbox_dismisser(self):
        """Start background thread to auto-dismiss MsgBox dialogs."""
        if self._msgbox_dismisser is None:
            self._msgbox_dismisser = MsgBoxDismisser(lambda: self.app)
        self._msgbox_dismisser.start()

    def stop_msgbox_dismisser(self):
        """Stop the MsgBox dismisser thread."""
        if self._msgbox_dismisser:
            self._msgbox_dismisser.stop()

    # ── File loading ──────────────────────────────────────────────────

    def open_file(self, file_path: str):
        """
        Load an input file into OTSoft by relaunching with a command-line argument.

        VB6 OTSoft has no file dialog — it reads Command() on startup
        for the file path. Deletes the settings file first to ensure defaults.
        """
        abs_path = os.path.abspath(file_path)
        logger.info("Opening file: %s", abs_path)

        # Delete VB6's saved settings file so the app starts with defaults.
        # It lives next to the OTSoft.exe.
        settings_file = os.path.join(
            os.path.dirname(self.otsoft_path), "OTSoftRememberUserChoices.txt"
        )
        try:
            os.remove(settings_file)
            logger.info("Deleted settings file: %s", settings_file)
        except FileNotFoundError:
            pass

        if self.main_win:
            try:
                self.main_win.close()
                time.sleep(1)
            except Exception:
                pass

        # Kill the old process unconditionally to avoid leaving a stale OTSoft
        # running alongside the new one (e.g. if a modal GLA window blocked
        # the main-window close above).
        if self.app:
            try:
                self.app.kill()
                time.sleep(0.3)
            except Exception:
                pass

        self._start_app(f'"{self.otsoft_path}" "{abs_path}"')

        # Dismiss any startup MsgBoxes
        time.sleep(1)
        _dismiss_msgboxes(self.app)

        # Verify file was loaded by checking button caption
        try:
            rank_btn = self.main_win.child_window(
                class_name="ThunderRT6CommandButton", found_index=0
            )
            logger.info("Rank button caption: %s", rank_btn.window_text())
        except Exception as e:
            logger.warning("Could not verify file load: %s", e)

    # ── Framework selection ───────────────────────────────────────────

    def select_classical_ot(self):
        """Select the 'Classical OT' radio button."""
        _run_with_timeout(
            lambda: self.main_win.child_window(title="Classical OT").click(),
            label="select Classical OT",
        )
        logger.info("Selected: Classical OT")

    def select_maximum_entropy(self):
        """Select the 'Maximum Entropy' radio button."""
        _run_with_timeout(
            lambda: self.main_win.child_window(title="Maximum Entropy").click(),
            label="select Maximum Entropy",
        )
        logger.info("Selected: Maximum Entropy")

    # ── Menu helpers ──────────────────────────────────────────────────

    def _click_menu(self, *path):
        """Click a sequence of menu items by their titles."""
        menu_path = "->".join(path)
        _run_with_timeout(
            lambda: self.main_win.menu_select(menu_path),
            label=f"menu select {path[-1]}",
        )

    def _set_menu_checked(self, desired: bool, *path):
        """
        Set a checkable menu item to the desired state, toggling only if needed.

        Args:
            desired: True to check, False to uncheck.
            *path: Menu path segments (e.g. "&Options", "Use BCD").
        """
        menu = self.main_win.menu()
        item = menu.get_menu_path("->".join(path))
        if item[-1].is_checked() != desired:
            self._click_menu(*path)
            state = "enabled" if desired else "disabled"
            logger.info("%s: %s", state.capitalize(), path[-1])

    # ── Algorithm variant (Options menu) ──────────────────────────────

    def set_algorithm_variant(self, algorithm: str, params: dict):
        """
        Configure Options menu for the correct Classical OT variant.

        Ensures BCD and LFCD toggles are set correctly, and sets
        BCD specificity if needed.
        """
        want_bcd = algorithm == "bcd"
        want_lfcd = algorithm == "lfcd"

        self._set_menu_checked(
            want_bcd, "&Options", "Use Biased Constraint Demotion"
        )
        self._set_menu_checked(
            want_lfcd, "&Options", "Use the Low Faithfulness version of RCD"
        )

        if want_bcd:
            try:
                self._set_menu_checked(
                    params.get("specific", False),
                    "&Options",
                    "BCD favors specific Faithfulness constraints",
                )
            except Exception as e:
                logger.warning("Could not set BCD specificity: %s", e)

    # ── Ranking argumentation checkboxes ──────────────────────────────

    def configure_ranking_options(self, params: dict):
        """Set the FRed / MIB / details / mini-tableaux checkboxes."""
        self._set_checkbox("Include ranking arguments", params.get("include_fred", True))
        self._set_checkbox("Use Most Informative Basis", params.get("use_mib", False))
        self._set_checkbox("Show details of argumentation", params.get("show_details", False))
        self._set_checkbox(
            "Include illustrative minitableaux",
            params.get("include_mini_tableaux", True),
        )

    def _set_checkbox(self, title: str, desired: bool):
        """Set a checkbox to the desired state."""
        try:
            cb = self.main_win.child_window(title=title)
            current = cb.get_check_state() == 1
            if current != desired:
                _run_with_timeout(
                    cb.click,
                    label=f"checkbox '{title}'",
                )
                state_str = "checked" if desired else "unchecked"
                logger.info("Set '%s' to %s", title, state_str)
        except StepTimeoutError:
            raise
        except Exception as e:
            logger.warning("Could not set checkbox '%s': %s", title, e)

    # ── A priori rankings ─────────────────────────────────────────────

    def enable_apriori_rankings(self):
        """Enable the a priori rankings menu toggle (file must already be in place)."""
        self._set_menu_checked(True, *APRIORI_MENU_PATH)
        # VB6 may show an error dialog asynchronously; give it a moment to appear.
        time.sleep(0.5)
        _dismiss_msgboxes(self.app)

    def disable_apriori_rankings(self):
        """Disable the a priori rankings menu toggle if currently enabled."""
        try:
            self._set_menu_checked(False, *APRIORI_MENU_PATH)
        except Exception:
            pass

    # ── Factorial Typology options ────────────────────────────────────

    def configure_factorial_typology(self, params: dict):
        """Set Factorial Typology menu options."""
        try:
            self._set_menu_checked(
                params.get("include_full_listing", False),
                "&Factorial Typology",
                "Include &rankings in results",
            )
        except Exception as e:
            logger.warning("Could not set FT options: %s", e)

    # ── Running algorithms ────────────────────────────────────────────

    def run_rank(self):
        """Click the Rank button and wait for completion."""
        rank_btn = self.main_win.child_window(
            class_name="ThunderRT6CommandButton", title_re=".*Rank.*"
        )
        logger.info("Clicking Rank button")
        _run_with_timeout(rank_btn.click, label="click Rank")
        self._poll_until(
            lambda: rank_btn.is_enabled() and self.main_win.is_enabled(),
            label="algorithm",
        )

    def run_factorial_typology(self):
        """Click the Factorial Typology button and wait for completion."""
        ft_btn = self.main_win.child_window(
            class_name="ThunderRT6CommandButton", title_re=".*[Ff]actorial.*"
        )
        logger.info("Clicking Factorial Typology button")
        _run_with_timeout(ft_btn.click, label="click Factorial Typology")
        self._poll_until(
            lambda: ft_btn.is_enabled() and self.main_win.is_enabled(),
            label="factorial typology",
            timeout=ALGORITHM_TIMEOUT_LONG,
        )

    def run_maxent(self, params: dict):
        """
        Run batch MaxEnt: select MaxEnt framework, click Rank to open GLA form,
        navigate to batch MaxEnt, configure parameters, and run.
        """
        # If a previous MaxEnt run left a GLA or MaxEnt window open, close it
        # before proceeding so we start from the main window.
        for title_re in (r".*[Mm]aximum [Ee]ntropy.*", r".*Gradual Learning Algorithm.*"):
            try:
                stale = self.app.window(title_re=title_re)
                if stale.exists(timeout=0):
                    stale.close()
                    time.sleep(0.5)
            except Exception:
                pass

        self.select_maximum_entropy()
        time.sleep(0.5)

        # Clicking the Rank button with MaxEnt selected opens the GLA form.
        # VB6's optMaximumEntropy_Click event may not fire via pywinauto, so the
        # button caption may still say "Rank <file>" rather than "Compute weights
        # for <file>" — match both possibilities.
        rank_btn = self.main_win.child_window(
            class_name="ThunderRT6CommandButton", title_re=r"(?i).*(rank|compute weights).*"
        )
        logger.info("Clicking Rank button: %r", rank_btn.window_text())
        _run_with_timeout(rank_btn.click, label="click Rank (MaxEnt)")
        time.sleep(2)

        # Find the GLA form. In OTSoft 2.7 the title is
        # "OTSoft 2.7 - GLA-MaxEnt - <file>" (not "Gradual Learning Algorithm").
        gla_win = self.app.window(title_re=".*GLA.*")
        _run_with_timeout(
            lambda: gla_win.wait("ready", timeout=30),
            label="wait for GLA window",
            timeout=30,
        )

        # Configure Gaussian prior BEFORE opening batch MaxEnt
        use_prior = params.get("use_prior", False)
        if use_prior:
            try:
                _run_with_timeout(
                    lambda: gla_win.menu_select("&MaxEnt->Run MaxEnt with Gaussian prior"),
                    label="enable Gaussian prior",
                )
                logger.info("Enabled Gaussian prior in GLA menu")
            except StepTimeoutError:
                raise
            except Exception as e:
                logger.warning("Could not set Gaussian prior: %s", e)

            sigma = params.get("sigma_squared", 1.0)
            if sigma != 1.0:
                # VB6 reads sigma from ModelParameters.txt — may need manual setup
                logger.warning(
                    "sigma_squared=%s requested — may need manual ModelParameters.txt edit",
                    sigma,
                )

        # Navigate: MaxEnt menu → "Run the batch version of MaxEnt"
        _run_with_timeout(
            lambda: gla_win.menu_select("&MaxEnt->Run the batch version of MaxEnt"),
            label="open batch MaxEnt",
        )
        time.sleep(1)

        # Find the MyMaxEnt form. In OTSoft 2.7 the title is
        # "OTSoft 2.7 - MaxEnt - <file>" (not "Maximum entropy").
        # Use " - MaxEnt - " to distinguish it from the GLA window
        # ("OTSoft 2.7 - GLA-MaxEnt - <file>").
        maxent_win = self.app.window(title_re=r".* - MaxEnt - .*")
        _run_with_timeout(
            lambda: maxent_win.wait("ready", timeout=30),
            label="wait for MaxEnt window",
            timeout=30,
        )

        # Set parameters
        iterations = params.get("iterations", 5)
        weight_min = params.get("weight_min", 0.0)
        weight_max = params.get("weight_max", 50.0)

        self._set_textbox(maxent_win, "txtPrecision", str(iterations), default="5")
        self._set_textbox(maxent_win, "txtWeightMinimum", str(weight_min), default="0")
        self._set_textbox(maxent_win, "txtWeightMaximum", str(weight_max), default="50")

        logger.info(
            "MaxEnt params: iterations=%s, weight_min=%s, weight_max=%s",
            iterations, weight_min, weight_max,
        )

        # Click Run and wait
        run_btn = maxent_win.child_window(title="Run maxent")
        _run_with_timeout(run_btn.click, label="click Run maxent")
        logger.info("Clicked Run maxent")

        self._poll_until(lambda: run_btn.is_enabled(), label="MaxEnt")

        # Return to main window
        try:
            exit_btn = maxent_win.child_window(title_re=".*[Ee]xit.*main.*")
            exit_btn.click()
            time.sleep(0.5)
        except Exception:
            pass

    def _set_textbox(self, window, control_name: str, value: str, *, default: str):
        """
        Set a VB6 text box value, first trying by control title, then by
        matching the expected default value.
        """
        try:
            textbox = window.child_window(
                class_name="ThunderRT6TextBox", title=control_name
            )
            textbox.set_text(value)
        except Exception:
            try:
                for tb in window.children(class_name="ThunderRT6TextBox"):
                    if tb.window_text().strip() == default:
                        tb.set_text(value)
                        logger.info("Set %s = %s (matched by default value)", control_name, value)
                        return
                logger.warning("Could not find textbox %s", control_name)
            except Exception as e:
                logger.warning("Could not set textbox %s: %s", control_name, e)

    # ── Wait helpers ──────────────────────────────────────────────────

    def _poll_until(
        self,
        condition: Callable[[], bool],
        *,
        label: str,
        timeout: float = ALGORITHM_TIMEOUT,
    ):
        """
        Poll until condition() returns True, with timeout.

        Dismisses any MsgBox dialogs upon completion.

        Raises:
            DialogDismissedError: If a VB6 dialog was dismissed during polling.
            TimeoutError: If the condition is not met within the timeout.
        """
        logger.info("Waiting for %s completion (timeout=%ss)...", label, timeout)
        start = time.time()
        time.sleep(1)  # Give the operation a moment to start

        # Snapshot the dismisser state so we only detect NEW dialogs
        dismissed_before = len(self._msgbox_dismisser.dismissed) if self._msgbox_dismisser else 0

        while time.time() - start < timeout:
            # Check for dialogs dismissed since we started polling
            if self._msgbox_dismisser:
                new_dismissed = self._msgbox_dismisser.dismissed[dismissed_before:]
                if new_dismissed:
                    logger.error(
                        "%s: dialog(s) dismissed during operation: %s",
                        label, new_dismissed,
                    )
                    raise DialogDismissedError(new_dismissed)

            try:
                if condition():
                    logger.info("%s completed in %.1fs", label.capitalize(), time.time() - start)
                    _dismiss_msgboxes(self.app)
                    return
            except Exception:
                pass
            time.sleep(POLL_INTERVAL)

        elapsed = time.time() - start
        raise TimeoutError(f"{label} timed out after {elapsed:.0f}s")

    # ── Reset between test cases ──────────────────────────────────────

    def restore_defaults(self):
        """Reset OTSoft to default settings via Options > Restore default settings."""
        try:
            self._click_menu("&Options", "Restore &default settings")
            logger.info("Restored default settings")
            time.sleep(0.5)
            _dismiss_msgboxes(self.app)
        except Exception as e:
            logger.warning("Could not restore defaults: %s", e)
