import base64
import io
import logging
import random
import re
import subprocess
import sys
import time
from pathlib import Path

import colorama
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from twocaptcha import TwoCaptcha

from .EmailAPIs import *

# Image processing dependencies (used by ddddocr)
try:
    from PIL import Image, ImageEnhance, ImageFilter, ImageOps

    PIL_AVAILABLE = True
except Exception:
    Image = None
    ImageEnhance = None
    ImageFilter = None
    ImageOps = None
    PIL_AVAILABLE = False

SILENT_MODE = "--silent" in sys.argv

# Environment control for debug artifact generation
import os


def generate_debug_artifacts_enabled():
    """Return True if GENERATE_DEBUG_ARTIFACTS env var is set to a truthy value."""
    return os.getenv("GENERATE_DEBUG_ARTIFACTS", "false").lower() in (
        "1",
        "true",
        "yes",
    )


# Random name lists for generating realistic names
FIRST_NAMES = [
    "James",
    "Mary",
    "John",
    "Patricia",
    "Robert",
    "Jennifer",
    "Michael",
    "Linda",
    "William",
    "Barbara",
    "David",
    "Elizabeth",
    "Richard",
    "Susan",
    "Joseph",
    "Jessica",
    "Thomas",
    "Sarah",
    "Charles",
    "Karen",
    "Christopher",
    "Nancy",
    "Daniel",
    "Lisa",
    "Matthew",
    "Betty",
    "Anthony",
    "Margaret",
    "Mark",
    "Sandra",
    "Donald",
    "Ashley",
    "Steven",
    "Kimberly",
    "Paul",
    "Emily",
    "Andrew",
    "Donna",
    "Joshua",
    "Michelle",
    "Kevin",
    "Carol",
    "Brian",
    "Amanda",
    "George",
    "Dorothy",
    "Edward",
    "Melissa",
    "Ronald",
    "Deborah",
    "Timothy",
    "Stephanie",
    "Jason",
    "Rebecca",
    "Jeffrey",
    "Sharon",
    "Ryan",
    "Laura",
    "Jacob",
    "Cynthia",
    "Gary",
    "Kathleen",
    "Nicholas",
    "Amy",
]

LAST_NAMES = [
    "Smith",
    "Johnson",
    "Williams",
    "Brown",
    "Jones",
    "Garcia",
    "Miller",
    "Davis",
    "Rodriguez",
    "Martinez",
    "Hernandez",
    "Lopez",
    "Gonzalez",
    "Wilson",
    "Anderson",
    "Thomas",
    "Taylor",
    "Moore",
    "Jackson",
    "Martin",
    "Lee",
    "Perez",
    "Thompson",
    "White",
    "Harris",
    "Sanchez",
    "Clark",
    "Ramirez",
    "Lewis",
    "Robinson",
    "Walker",
    "Young",
    "Allen",
    "King",
    "Wright",
    "Scott",
    "Torres",
    "Nguyen",
    "Hill",
    "Flores",
    "Green",
    "Adams",
    "Nelson",
    "Baker",
    "Hall",
    "Rivera",
    "Campbell",
    "Mitchell",
    "Carter",
    "Roberts",
    "Gomez",
    "Phillips",
    "Evans",
    "Turner",
    "Diaz",
    "Parker",
    "Cruz",
    "Edwards",
    "Collins",
    "Reyes",
    "Stewart",
    "Morris",
    "Morales",
    "Murphy",
]


def generate_random_name():
    """Generate a random realistic full name"""
    first_name = random.choice(FIRST_NAMES)
    last_name = random.choice(LAST_NAMES)
    return f"{first_name} {last_name}"


class IPBlockedException(Exception):
    def __init__(self, message):
        super().__init__(message)


class EsetRegister(object):
    def __init__(
        self, registered_email_obj: OneSecEmailAPI, eset_password: str, driver: Chrome
    ):
        self.email_obj = registered_email_obj
        self.eset_password = eset_password
        self.driver = driver
        self.window_handle = None

    def createAccount(self):
        exec_js = self.driver.execute_script
        uCE = untilConditionExecute

        logging.info("[EMAIL] Register page loading...")
        console_log("\n[EMAIL] Register page loading...", INFO, silent_mode=SILENT_MODE)
        if isinstance(self.email_obj, WEB_WRAPPER_EMAIL_APIS_CLASSES):
            self.driver.switch_to.new_window("tab")
            self.window_handle = self.driver.current_window_handle
        self.driver.get("https://login.eset.com/Register")
        uCE(self.driver, f"return {GET_EBID}('email') != null")
        logging.info("[EMAIL] Register page is loaded!")
        console_log("[EMAIL] Register page is loaded!", OK, silent_mode=SILENT_MODE)

        logging.info("Bypassing cookies...")
        console_log("\nBypassing cookies...", INFO, silent_mode=SILENT_MODE)
        if uCE(
            self.driver,
            f"return {CLICK_WITH_BOOL}({GET_EBAV}('button', 'id', 'cc-accept'))",
            max_iter=10,
            raise_exception_if_failed=False,
        ):
            logging.info("Cookies successfully bypassed!")
            console_log("Cookies successfully bypassed!", OK, silent_mode=SILENT_MODE)
            time.sleep(
                1
            )  # Once pressed, you have to wait a little while. If code do not do this, the site does not count the acceptance of cookies
        else:
            logging.info(
                "Cookies were not bypassed (it doesn't affect the algorithm, I think :D)"
            )
            console_log(
                "Cookies were not bypassed (it doesn't affect the algorithm, I think :D)",
                ERROR,
                silent_mode=SILENT_MODE,
            )

        exec_js(f"return {GET_EBID}('email')").send_keys(self.email_obj.email)
        uCE(
            self.driver,
            f"return {CLICK_WITH_BOOL}({DEFINE_GET_EBAV_FUNCTION}('button', 'data-label', 'register-continue-button'))",
        )
        time.sleep(1)
        try:
            if (
                exec_js(
                    f"return {GET_EBAV}('div', 'data-label', 'register-email-formGroup-validation')"
                )
                is not None
            ):
                raise RuntimeError(
                    f"Email: {self.email_obj.email} is already registered!"
                )
        except:
            pass

        logging.info("[PASSWD] Register page loading...")
        console_log(
            "\n[PASSWD] Register page loading...", INFO, silent_mode=SILENT_MODE
        )
        uCE(
            self.driver,
            f"return typeof {GET_EBAV}('button', 'data-label', 'register-create-account-button') === 'object'",
        )
        logging.info("[PASSWD] Register page is loaded!")
        console_log("[PASSWD] Register page is loaded!", OK, silent_mode=SILENT_MODE)
        exec_js(f"return {GET_EBID}('password')").send_keys(self.eset_password)

        # Select Ukraine country
        logging.info("Selecting the country...")
        if (
            exec_js(
                f"return {GET_EBCN}('select__single-value css-1dimb5e-singleValue')[0]"
            ).text
            != "Ukraine"
        ):
            exec_js(
                f"return {GET_EBCN}('select__control css-13cymwt-control')[0]"
            ).click()
            for country in exec_js(
                f"return {GET_EBCN}('select__option css-uhiml7-option')"
            ):
                if country.text == "Ukraine":
                    country.click()
                    logging.info("Country selected!")
                    break

        uCE(
            self.driver,
            f"return {CLICK_WITH_BOOL}({DEFINE_GET_EBAV_FUNCTION}('button', 'data-label', 'register-create-account-button'))",
        )

        for _ in range(DEFAULT_MAX_ITER):
            title = exec_js("return document.title")
            if title == "Service not available":
                raise IPBlockedException(
                    "\nESET temporarily blocked your IP, try again later!!! Try to use VPN/Proxy or try to change Email API!!!"
                )
            url = exec_js("return document.URL")
            if url == "https://home.eset.com/":
                return True
            time.sleep(DEFAULT_DELAY)
        raise IPBlockedException(
            "\nESET temporarily blocked your IP, try again later!!! Try to use VPN/Proxy or try to change Email API!!!"
        )

    def confirmAccount(self):
        uCE = untilConditionExecute
        # uCE(self.driver, f'return {CLICK_WITH_BOOL}({GET_EBAV}("ion-button", "data-r", "account-verification-email-modal-resend-email-btn"))') # accelerating the receipt of an eset token

        if isinstance(self.email_obj, CustomEmailAPI):
            token = parseToken(self.email_obj, max_iter=100, delay=3)
        else:
            logging.info(
                f"[{self.email_obj.class_name}] ESET-HOME-Token interception..."
            )
            console_log(
                f"\n[{self.email_obj.class_name}] ESET-HOME-Token interception...",
                INFO,
                silent_mode=SILENT_MODE,
            )
            if isinstance(self.email_obj, WEB_WRAPPER_EMAIL_APIS_CLASSES):
                token = parseToken(self.email_obj, self.driver, max_iter=100, delay=3)
                self.driver.switch_to.window(self.window_handle)
            else:
                token = parseToken(
                    self.email_obj, max_iter=100, delay=3
                )  # 1secmail, developermail
        logging.info(f"ESET-HOME-Token: {token}")
        logging.info("Account confirmation is in progress...")
        console_log(f"ESET-HOME-Token: {token}", OK, silent_mode=SILENT_MODE)
        console_log(
            "\nAccount confirmation is in progress...", INFO, silent_mode=SILENT_MODE
        )
        self.driver.get(
            f"https://login.eset.com/link/confirmregistration?token={token}"
        )
        uCE(self.driver, 'return document.title.includes("ESET HOME")')
        try:
            uCE(self.driver, f'return {GET_EBCN}("verification-email_p").length === 0')
        except:
            self.driver.get(
                f"https://login.eset.com/link/confirmregistration?token={token}"
            )
            uCE(self.driver, 'return document.title.includes("ESET HOME")')
            uCE(self.driver, f'return {GET_EBCN}("verification-email_p").length === 0')
        logging.info("Account successfully confirmed!")
        console_log("Account successfully confirmed!", OK, silent_mode=SILENT_MODE)
        return True


class EsetKeygen(object):
    def __init__(
        self, registered_email_obj: OneSecEmailAPI, driver: Chrome, mode="ESET HOME"
    ):
        self.email_obj = registered_email_obj
        self.driver = driver
        self.mode = mode.upper()
        if self.mode not in ["ESET HOME", "SMALL BUSINESS"]:
            raise RuntimeError("Undefined keygen mode!")

    def sendRequestForKey(self):
        uCE = untilConditionExecute

        logging.info(f"[{self.mode}] Request sending...")
        console_log(
            f"\n[{self.mode}] Request sending...", INFO, silent_mode=SILENT_MODE
        )

        # After account confirmation, we should already be logged in
        # Check current URL first before navigating
        current_url_after_confirm = self.driver.current_url.lower()
        logging.info(f"URL after account confirmation: {self.driver.current_url}")

        # If we're already on home.eset.com, great! If not, navigate there
        if "home.eset.com" not in current_url_after_confirm:
            logging.info("Not on home.eset.com, navigating there...")
            self.driver.get("https://home.eset.com/")
            time.sleep(3)  # Initial wait
        else:
            logging.info("Already on home.eset.com after confirmation")
            time.sleep(2)  # Still wait a bit

        # Wait for the React app to fully load (check for meaningful content)
        logging.info("Waiting for page to fully load...")
        page_loaded = False
        for i in range(15):  # Try for up to 30 seconds
            time.sleep(2)
            # Check if page has loaded by looking for actual content
            has_content = self.driver.execute_script(
                """
                var root = document.getElementById('root');
                if (!root) return false;
                // Check if there's actual content, not just loading divs
                var content = root.innerText.trim();
                return content.length > 50;  // Should have some text if loaded
            """
            )
            if has_content:
                logging.info("Page content loaded")
                page_loaded = True
                break
            logging.info(f"Waiting for page content... ({i+1}/15)")

        if not page_loaded:
            logging.warning("Page may not have fully loaded, continuing anyway...")

        current_url = self.driver.current_url.lower()
        logging.info(f"Current URL before onboarding check: {self.driver.current_url}")

        # Check if redirected to login (session expired)
        if "login" in current_url:
            raise RuntimeError(
                "Session expired or not logged in! Please ensure account is confirmed."
            )

        # Check if we're on the onboarding page
        if "onboarding" in current_url:
            logging.info("Detected onboarding page, attempting bypass...")
            console_log(
                "\nDetected onboarding page, attempting bypass...",
                INFO,
                silent_mode=SILENT_MODE,
            )

            # Strategy: Complete the onboarding flow properly
            logging.warning("Onboarding is mandatory, completing flow...")
            console_log(
                "Onboarding is mandatory, completing flow...",
                INFO,
                silent_mode=SILENT_MODE,
            )

            # Try using JavaScript to complete the onboarding
            try:
                max_attempts = 20  # Increased for all onboarding steps
                for step in range(max_attempts):
                    time.sleep(2)  # Wait for page to stabilize

                    current_url = self.driver.current_url.lower()
                    logging.info(
                        f"Onboarding step {step+1}/{max_attempts}, URL: {self.driver.current_url}"
                    )

                    # Check if we've left onboarding
                    if "onboarding" not in current_url:
                        logging.info("Successfully left onboarding!")
                        console_log(
                            "Onboarding completed!", OK, silent_mode=SILENT_MODE
                        )
                        break

                    # Strategy: Handle different onboarding screens in order
                    action_taken = False

                    # Step 1: Skip introduction on welcome page
                    skip_intro = self.driver.execute_script(
                        """
                        var skipBtn = document.querySelector('button');
                        if (skipBtn && skipBtn.innerText.toLowerCase().includes('skip introduction')) {
                            skipBtn.click();
                            return true;
                        }
                        return false;
                    """
                    )
                    if skip_intro:
                        logging.info("Clicked Skip Introduction")
                        action_taken = True
                        time.sleep(2)
                        continue

                    # Step 2a: Select "Start a 30-day trial" option (more aggressive)
                    trial_selected = False
                    try:
                        trial_selected = self.driver.execute_script(
                            """
                            (function(){
                                var debug = { success: false, inputChecked: false, buttons: [] };
                                var trialLabel = document.querySelector('label[data-label="onboarding-add-subscription-protect-card-trial"]');
                                var input = document.getElementById('trial') || (trialLabel && trialLabel.querySelector('input')) || null;
                                // Try clicking label and input with mouse events
                                try {
                                    if (trialLabel) {
                                        trialLabel.scrollIntoView({block: 'center'});
                                        trialLabel.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
                                        trialLabel.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
                                        trialLabel.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                                    }
                                } catch(e) {}
                                try {
                                    if (input) {
                                        input.click();
                                        input.checked = true;
                                        input.setAttribute('checked', '');
                                        input.dispatchEvent(new Event('input', { bubbles: true }));
                                        input.dispatchEvent(new Event('change', { bubbles: true }));
                                        input.dispatchEvent(new Event('blur', { bubbles: true }));
                                        input.focus();
                                        input.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                                        debug.inputChecked = !!input.checked;
                                    }
                                } catch(e) { }

                                // Set aria attributes on radio elements if present
                                try {
                                    if (trialLabel) {
                                        trialLabel.setAttribute('aria-checked', 'true');
                                    }
                                } catch(e) {}

                                var buttons = Array.from(document.querySelectorAll('button'));
                                for (var i=0;i<buttons.length;i++){
                                    var btn = buttons[i];
                                    var txt = (btn.innerText||'').toLowerCase().trim();
                                    var b = { text: txt, disabled: !!btn.disabled, ariaDisabled: btn.getAttribute('aria-disabled') };
                                    debug.buttons.push(b);
                                    if (txt === 'continue') {
                                        try {
                                            // try to enable it
                                            btn.removeAttribute('disabled');
                                            btn.setAttribute('aria-disabled', 'false');
                                            btn.disabled = false;
                                            btn.classList.remove('disabled');
                                        } catch(e) {}
                                        try { btn.scrollIntoView({block:'center'}); btn.click(); debug.success = true; return debug; } catch(e) { }
                                    }
                                }

                                // Last attempt: find any element with data-label containing 'continue' and click
                                try {
                                    var cont = document.querySelector('[data-label*="continue"]');
                                    if (cont) { cont.click(); debug.success = true; }
                                } catch(e) {}

                                return debug;
                            })();
                        """
                        )
                        # trial_selected will be a dict-like object from JS; convert to truthy
                        if isinstance(trial_selected, dict):
                            logging.info(f"Trial JS result: {trial_selected}")
                            action_taken = trial_selected.get(
                                "success", False
                            ) or trial_selected.get("inputChecked", False)
                            if action_taken:
                                trial_selected = True
                            else:
                                trial_selected = False
                        else:
                            trial_selected = bool(trial_selected)
                    except Exception as E:
                        logging.debug(f"Trial selection JS error: {E}")
                        trial_selected = False

                    if trial_selected:
                        logging.info("Selected trial option (aggressive)")
                        action_taken = True
                        time.sleep(1)
                    else:
                        # Save additional state to help CI debugging
                        if generate_debug_artifacts_enabled():
                            try:
                                state = self.driver.execute_script(
                                    "return (function(){ var res = {}; res.bodyText = document.body.innerText.slice(0,2000); res.buttons = Array.from(document.querySelectorAll('button')).map(b=>({text:(b.innerText||'').trim(), disabled:!!b.disabled, ariaDisabled:b.getAttribute('aria-disabled')})); res.radios = Array.from(document.querySelectorAll('input[type=radio]')).map(r=>({id:r.id, name:r.name, checked:!!r.checked})); return res; })();"
                                )
                                with open(
                                    "debug_onboarding_state.json", "w", encoding="utf-8"
                                ) as f:
                                    import json

                                    json.dump(state, f, indent=2)
                                logging.info("Saved debug_onboarding_state.json")
                            except Exception as e:
                                logging.warning(
                                    f"Could not save debug_onboarding_state.json: {e}"
                                )
                        else:
                            logging.debug(
                                "Skipping debug_onboarding_state.json write (disabled by env)"
                            )

                    # Step 2b: Select "Protect your home" option (if present)
                    home_selected = self.driver.execute_script(
                        """
                        // Look for "Protect your home" option
                        var labels = document.querySelectorAll('label[tabindex="0"]');
                        for (var i = 0; i < labels.length; i++) {
                            var text = labels[i].innerText.toLowerCase();
                            if (text.includes('protect your home') || text.includes('home')) {
                                var input = labels[i].querySelector('input[type="radio"]');
                                if (input && !input.checked) {
                                    labels[i].click();
                                    return true;
                                }
                            }
                        }
                        return false;
                    """
                    )
                    if home_selected:
                        logging.info("Selected Protect your home option")
                        action_taken = True
                        time.sleep(1)

                    # Step 3a: Fill in member name (if input field is present) and submit form
                    try:
                        name_input = self.driver.find_element(
                            "css selector", 'input[name="name"]'
                        )
                        if (
                            name_input
                            and name_input.get_attribute("value").strip() == ""
                        ):
                            random_name = generate_random_name()
                            logging.info(
                                f"Found empty name field, filling with: {random_name}"
                            )
                            name_input.click()  # Focus the field
                            time.sleep(0.1)
                            name_input.send_keys(
                                random_name
                            )  # Selenium's send_keys triggers all keyboard events
                            time.sleep(0.5)

                            # Dispatch change and blur events
                            self.driver.execute_script(
                                """
                                var nameInput = arguments[0];
                                nameInput.dispatchEvent(new Event('change', { bubbles: true }));
                                nameInput.dispatchEvent(new Event('blur', { bubbles: true }));
                            """,
                                name_input,
                            )

                            logging.info(
                                "Filled member name, waiting for Continue button to enable"
                            )
                            time.sleep(1)

                            # Try to click continue button
                            continue_clicked = self.driver.execute_script(
                                """
                                var continueBtn = document.querySelector('button[data-label="onboarding-members-continue-btn"]');
                                if (continueBtn && !continueBtn.disabled) {
                                    continueBtn.click();
                                    return true;
                                }
                                return false;
                            """
                            )

                            if continue_clicked:
                                logging.info("Clicked Continue button after name fill")
                                action_taken = True
                                time.sleep(2)
                                continue
                    except:
                        pass

                    # Step 3b: Select "Just me" option (if present)
                    just_me_selected = self.driver.execute_script(
                        """
                        var justMeLabel = document.querySelector('label[data-label="onboarding-members-me-option"]');
                        if (justMeLabel) {
                            var input = justMeLabel.querySelector('input');
                            if (input && !input.checked) {
                                justMeLabel.click();
                                return true;
                            }
                        }
                        return false;
                    """
                    )
                    if just_me_selected:
                        logging.info("Selected Just me option")
                        action_taken = True
                        time.sleep(1)

                    # Step 4: Click "Finish for now" button (final step)
                    finish_clicked = self.driver.execute_script(
                        """
                        var buttons = document.querySelectorAll('button');
                        for (var i = 0; i < buttons.length; i++) {
                            var btn = buttons[i];
                            var text = btn.innerText.toLowerCase();
                            if (text.includes('finish for now')) {
                                btn.click();
                                return true;
                            }
                        }
                        return false;
                    """
                    )
                    if finish_clicked:
                        logging.info("Clicked Finish for now")
                        action_taken = True
                        time.sleep(3)
                        continue

                    # Generic: Try to click any enabled Continue button
                    continue_clicked = self.driver.execute_script(
                        """
                        // Look for Continue buttons
                        var buttons = document.querySelectorAll('button:not([disabled])');
                        for (var i = 0; i < buttons.length; i++) {
                            var btn = buttons[i];
                            if (btn.getAttribute('aria-disabled') === 'true') continue;
                            
                            var btnText = btn.innerText.toLowerCase().trim();
                            if (btnText === 'continue') {
                                btn.click();
                                return true;
                            }
                        }
                        return false;
                    """
                    )

                    if continue_clicked:
                        logging.info("Clicked Continue button")
                        action_taken = True
                        time.sleep(2)

                    # If no action was taken, we might be stuck
                    if not action_taken:
                        logging.warning(
                            f"No actionable element found on current onboarding page"
                        )
                        time.sleep(1)

                # After completing onboarding steps, give it time to settle
                time.sleep(3)

            except Exception as e:
                logging.warning(f"Error during onboarding: {e}")

            # Final check - if still on onboarding, raise error
            if "onboarding" in self.driver.current_url.lower():
                logging.error(
                    f"Still on onboarding after all attempts. URL: {self.driver.current_url}"
                )
                if generate_debug_artifacts_enabled():
                    # Save debug info
                    try:
                        with open(
                            "debug_onboarding_stuck.html", "w", encoding="utf-8"
                        ) as f:
                            f.write(self.driver.page_source)
                        self.driver.save_screenshot("debug_onboarding_stuck.png")
                        logging.info("Saved debug_onboarding_stuck.*")
                    except Exception as e:
                        logging.warning(
                            f"Could not save debug_onboarding_stuck files: {e}"
                        )
                    # Also attempt to capture the UI state (buttons, radios, small page text) for CI inspection
                    try:
                        state = self.driver.execute_script(
                            "return (function(){ var res={}; res.bodyText = document.body.innerText.slice(0,4000); res.buttons = Array.from(document.querySelectorAll('button')).map(b=>({text:(b.innerText||'').trim(), disabled:!!b.disabled, ariaDisabled:b.getAttribute('aria-disabled')})); res.radios = Array.from(document.querySelectorAll('input[type=radio]')).map(r=>({id:r.id, name:r.name, checked:!!r.checked})); return res; })();"
                        )
                        import json

                        with open(
                            "debug_onboarding_state.json", "w", encoding="utf-8"
                        ) as f:
                            json.dump(state, f, indent=2)
                        logging.info("Saved debug_onboarding_state.json")
                    except Exception as E:
                        logging.warning(
                            f"Could not save debug_onboarding_state.json: {E}"
                        )
                else:
                    logging.debug(
                        "Skipping onboarding debug artifact writes (disabled by env)"
                    )
                raise RuntimeError(
                    "Cannot bypass onboarding - ESET enforces mandatory onboarding flow"
                )

        # After onboarding is complete, the trial subscription is already activated
        # Navigate to subscriptions page to extract the license key
        logging.info(
            "Onboarding complete! Now navigating to subscriptions to retrieve license key..."
        )
        time.sleep(2)

        # Navigate to subscriptions page
        self.driver.get("https://home.eset.com/subscriptions")
        time.sleep(5)  # Wait for page load

        logging.info(f"Current URL: {self.driver.current_url}")

        # Check if redirected to login
        if "login" in self.driver.current_url.lower():
            logging.error(f"Redirected to login page. URL: {self.driver.current_url}")
            raise RuntimeError("Session expired - redirected to login!")

        # Wait for subscriptions page to load and find the subscription card
        logging.info("Looking for subscription card...")
        subscription_found = False
        for attempt in range(10):
            subscription_text = self.driver.execute_script(
                """
                var text = document.body.innerText.toLowerCase();
                return text.includes('subscription') || text.includes('eset');
            """
            )
            if subscription_text:
                subscription_found = True
                logging.info("Subscription content found on page")
                break
            time.sleep(1)

        if not subscription_found:
            logging.warning(
                "Could not verify subscription content, continuing anyway..."
            )

        # Click the "Open subscription" button
        logging.info('Looking for "Open subscription" button...')
        open_subscription_clicked = False
        for attempt in range(10):
            button_clicked = self.driver.execute_script(
                """
                var buttons = document.querySelectorAll('button');
                for (var i = 0; i < buttons.length; i++) {
                    var btn_text = buttons[i].innerText.toLowerCase().trim();
                    if (btn_text === 'open subscription') {
                        buttons[i].click();
                        return true;
                    }
                }
                return false;
            """
            )
            if button_clicked:
                open_subscription_clicked = True
                logging.info('Successfully clicked "Open subscription" button')
                time.sleep(3)  # Wait for subscription detail page to load
                break
            time.sleep(1)

        if not open_subscription_clicked:
            logging.error('Could not find or click "Open subscription" button')
            if generate_debug_artifacts_enabled():
                try:
                    with open(
                        "debug_subscriptions_page.html", "w", encoding="utf-8"
                    ) as f:
                        f.write(self.driver.page_source)
                    self.driver.save_screenshot("debug_subscriptions_page.png")
                    logging.info(
                        "Saved debug_subscriptions_page.html and debug_subscriptions_page.png"
                    )
                except Exception as e:
                    logging.warning(
                        f"Could not write debug_subscriptions_page files: {e}"
                    )
            else:
                logging.debug(
                    "Skipping debug_subscriptions_page writes (disabled by env)"
                )
            raise RuntimeError("Could not access subscription details!")

        # Extract the license key and expiration date from the subscription detail page
        logging.info("Extracting license key and expiration date...")
        license_info = self.driver.execute_script(
            """
            var result = {
                activation_key: null,
                expiration_date: null,
                found_by: null
            };

            // Get all text content
            var all_text = document.body.innerText || '';

            // Try to find activation key following the label "Activation key" (case-insensitive)
            var activation_match = all_text.match(/activation key[\\s\\S]*?([A-Z0-9]{4}(?:-[A-Z0-9]{4}){4})/i);
            if (activation_match && activation_match[1]) {
                result.activation_key = activation_match[1].toUpperCase();
                result.found_by = 'activation_label';
            } else {
                // Fallback: search for any 5-group license-like pattern anywhere on the page
                var key_match = all_text.match(/[A-Z0-9]{4}(?:-[A-Z0-9]{4}){4}/i);
                if (key_match && key_match[0]) {
                    result.activation_key = key_match[0].toUpperCase();
                    result.found_by = 'pattern_fallback';
                }
            }

            // Look for expiration date - search for "Expiration date" followed by DD.MM.YYYY format
            var expiration_match = all_text.match(/[Ee]xpiration date[\\s\\S]*?(\\d{2}\\.\\d{2}\\.\\d{4})/);
            if (expiration_match && expiration_match[1]) {
                result.expiration_date = expiration_match[1];
            }

            return result;
        """
        )

        if license_info["activation_key"]:
            logging.info(f'[{self.mode}] License key: {license_info["activation_key"]}')
            logging.info(
                f'[{self.mode}] Expiration date: {license_info["expiration_date"]}'
            )
            console_log(
                f"\n[{self.mode}] Request successfully sent!",
                OK,
                silent_mode=SILENT_MODE,
            )
        else:
            logging.error("Could not extract license key from subscription page")
            logging.info(f"license_info: {license_info}")
            logging.info(
                f'Page content preview: {self.driver.find_element("tag name", "body").text[:500]}'
            )
            if generate_debug_artifacts_enabled():
                try:
                    with open(
                        "debug_subscription_detail.html", "w", encoding="utf-8"
                    ) as f:
                        f.write(self.driver.page_source)
                    self.driver.save_screenshot("debug_subscription_detail.png")
                    logging.info(
                        "Saved debug_subscription_detail.html and debug_subscription_detail.png"
                    )
                    # Also save text content (body text)
                    with open(
                        "debug_subscription_detail.txt", "w", encoding="utf-8"
                    ) as f:
                        f.write(self.driver.find_element("tag name", "body").text)
                    logging.info("Saved debug_subscription_detail.txt with page text")

                    # Additionally capture the full page text via JS (document.documentElement.innerText)
                    try:
                        full_text = (
                            self.driver.execute_script(
                                "return document.documentElement.innerText"
                            )
                            or ""
                        )
                        with open(
                            "debug_subscription_full_text.txt", "w", encoding="utf-8"
                        ) as f:
                            f.write(full_text)
                        logging.info(
                            "Saved debug_subscription_full_text.txt with full page text"
                        )
                    except Exception as e:
                        logging.warning(f"Could not capture full page text via JS: {e}")
                except Exception as e:
                    logging.warning(f"Could not write subscription debug files: {e}")
            else:
                logging.debug(
                    "Skipping subscription debug artifact writes (disabled by env)"
                )
            raise RuntimeError(
                "Failed to extract license key from subscription details!"
            )

    def getLD(self):
        exec_js = self.driver.execute_script
        uCE = untilConditionExecute
        logging.info(f"License uploads...")
        console_log("\nLicense uploads...", INFO, silent_mode=SILENT_MODE)
        uCE(
            self.driver,
            f"return {GET_EBAV}('div', 'data-label', 'license-detail-info') != null",
            raise_exception_if_failed=False,
        )
        if self.driver.current_url.find("detail") != -1:
            logging.info(f"License ID: {self.driver.current_url[-11:]}")
            console_log(
                f"License ID: {self.driver.current_url[-11:]}",
                OK,
                silent_mode=SILENT_MODE,
            )
        uCE(
            self.driver,
            f"return {GET_EBAV}('div', 'data-label', 'license-detail-product-name') != null",
            max_iter=10,
        )
        uCE(
            self.driver,
            f"return {GET_EBAV}('div', 'data-label', 'license-detail-license-model-additional-info') != null",
            max_iter=10,
        )
        uCE(
            self.driver,
            f"return {GET_EBAV}('div', 'data-label', 'license-detail-license-key') != null",
            max_iter=10,
        )
        license_name = exec_js(
            f"return {GET_EBAV}('div', 'data-label', 'license-detail-product-name').innerText"
        )
        license_out_date = exec_js(
            f"return {GET_EBAV}('div', 'data-label', 'license-detail-license-model-additional-info').innerText"
        )
        license_key = exec_js(
            f"return {GET_EBAV}('div', 'data-label', 'license-detail-license-key').innerText"
        )
        logging.info("Information successfully received!")
        console_log("Information successfully received!", OK, silent_mode=SILENT_MODE)
        return license_name, license_key, license_out_date


class EsetVPN(object):
    def __init__(
        self,
        registered_email_obj: OneSecEmailAPI,
        driver: Chrome,
        EsetRegister_window_handle=None,
    ):
        self.email_obj = registered_email_obj
        self.driver = driver
        self.window_handle = EsetRegister_window_handle

    def sendRequestForVPNCodes(self):
        exec_js = self.driver.execute_script
        uCE = untilConditionExecute

        logging.info("Sending a request for VPN subscriptions...")
        console_log(
            "\nSending a request for VPN subscriptions...",
            INFO,
            silent_mode=SILENT_MODE,
        )
        self.driver.get("https://home.eset.com/security-features")
        try:
            uCE(
                self.driver,
                f'return {CLICK_WITH_BOOL}({GET_EBAV}("button", "data-label", "security-feature-explore-button"))',
                max_iter=10,
            )
        except:
            raise RuntimeError("Explore-feature-button error!")
        time.sleep(0.5)
        for profile in exec_js(
            f'return {GET_EBAV}("button", "data-label", "choose-profile-tile-button", -1)'
        ):  # choose Me profile
            if (
                profile.get_attribute("innerText").find(self.email_obj.email) != -1
            ):  # Me profile contains an email address
                profile.click()
        uCE(
            self.driver,
            f'return {CLICK_WITH_BOOL}({GET_EBAV}("button", "data-label", "choose-profile-continue-btn"))',
            max_iter=5,
        )
        uCE(
            self.driver,
            f'return {GET_EBAV}("button", "data-label", "choose-device-counter-increment-button") != null',
            max_iter=10,
        )
        for _ in range(9):  # increasing 'Number of devices' (to 10)
            exec_js(
                f'{GET_EBAV}("button", "data-label", "choose-device-counter-increment-button").click()'
            )
        exec_js(
            f'{GET_EBAV}("button", "data-label", "choose-device-count-submit-button").click()'
        )
        uCE(
            self.driver,
            f'return {GET_EBAV}("button", "data-label", "pwm-instructions-sent-download-button") != null',
            max_iter=15,
        )
        logging.info("Request successfully sent!")
        console_log("Request successfully sent!", OK, silent_mode=SILENT_MODE)
        return True

    def getVPNCodes(self):
        if isinstance(self.email_obj, CustomEmailAPI):
            logging.warning(
                "Wait for a message to your e-mail about instructions on how to set up the VPN!!!"
            )
            console_log(
                "\nWait for a message to your e-mail about instructions on how to set up the VPN!!!",
                WARN,
                True,
                SILENT_MODE,
            )
            return None
        else:
            logging.info(f"[{self.email_obj.class_name}] VPN Codes interception...")
            console_log(
                f"\n[{self.email_obj.class_name}] VPN Codes interception...",
                INFO,
                silent_mode=SILENT_MODE,
            )  # timeout 1.5m
            if isinstance(self.email_obj, WEB_WRAPPER_EMAIL_APIS_CLASSES):
                vpn_codes = parseVPNCodes(
                    self.email_obj, self.driver, delay=2, max_iter=45
                )
                self.driver.switch_to.window(self.window_handle)
            else:
                vpn_codes = parseVPNCodes(
                    self.email_obj, self.driver, delay=2, max_iter=45
                )  # 1secmail, developermail
                logging.info("Information successfully received!")
                console_log(
                    "Information successfully received!", OK, silent_mode=SILENT_MODE
                )
        return vpn_codes


class EsetProtectHubRegister(object):
    def __init__(
        self, registered_email_obj: OneSecEmailAPI, eset_password: str, driver: Chrome
    ):
        self.email_obj = registered_email_obj
        self.driver = driver
        self.eset_password = eset_password
        self.window_handle = None

    def solve_with_capsolver_fixed(self):
        """
        Fixed Capsolver method without proxy requirement
        Improvements:
        - Use API key from environment variable `CAPSOLVER_API_KEY` (fallback to hardcoded value)
        - Use consistent clientKey for createTask and getTaskResult
        - Increase polling and add logging
        - Robust token injection and verification
        """
        try:
            import json
            import os
            import time

            import requests

            client_key = os.environ.get(
                "CAPSOLVER_API_KEY", "CAP-D051815FD86B044D03BF198CC9DFEB4B"
            )
            if client_key.startswith("CAP-") is False:
                print(
                    "[  WARN  ] Capsolver client key looks invalid, check environment variable CAPSOLVER_API_KEY"
                )

            task_data = {
                "clientKey": client_key,
                "task": {
                    "type": "MTCaptchaTaskProxyless",
                    "websiteURL": self.driver.current_url,
                    "websiteKey": "MTPublic-JnEM38Q6U",
                },
            }

            response = requests.post(
                "https://api.capsolver.com/createTask", json=task_data, timeout=30
            )
            task_result = response.json()

            if task_result.get("errorId", 1) != 0:
                print(
                    f"[  WARN  ] Capsolver error: {task_result.get('errorDescription', 'Unknown error')}"
                )
                return False

            task_id = task_result.get("taskId")
            if not task_id:
                print("[  WARN  ] Capsolver did not return a taskId")
                return False

            # Poll for result (up to ~2.5 minutes)
            for attempt in range(30):
                time.sleep(5)
                result_data = {"clientKey": client_key, "taskId": task_id}
                try:
                    result_response = requests.post(
                        "https://api.capsolver.com/getTaskResult",
                        json=result_data,
                        timeout=20,
                    )
                    result = result_response.json()
                except Exception as e:
                    print(f"[  WARN  ] Capsolver polling error: {e}")
                    continue

                status = result.get("status")
                if status == "ready":
                    solution = result.get("solution") or {}
                    token = solution.get("token")
                    if not token:
                        print("[  WARN  ] Capsolver returned no token in solution")
                        return False

                    # Inject the solution token into the page
                    self.driver.execute_script(
                        """
                        (function(token){
                            var tokenField = document.querySelector('[name="mtcaptcha-verifiedtoken"]');
                            if (!tokenField) {
                                tokenField = document.createElement('input');
                                tokenField.type = 'hidden';
                                tokenField.name = 'mtcaptcha-verifiedtoken';
                                document.forms[0].appendChild(tokenField);
                            }
                            tokenField.value = token;
                            // Dispatch change events if necessary
                            tokenField.dispatchEvent(new Event('change', { bubbles: true }));
                        })(arguments[0]);
                    """,
                        token,
                    )

                    print("[   OK   ] MTCaptcha solved via Capsolver")

                    # Allow time for site to verify token
                    time.sleep(2)

                    # Verify presence of token on page
                    try:
                        verified = self.driver.execute_script(
                            'return (document.querySelector("[name=\'mtcaptcha-verifiedtoken\']") || {}).value || ""'
                        )
                        if verified and verified.strip() != "":
                            return True
                    except:
                        pass

                    return True

            print("[  WARN  ] Capsolver timeout")
            return False

        except Exception as e:
            print(f"[  WARN  ] Capsolver failed: {str(e)}")
            return False

    # ORPHANED FUNCTION - Unused, superseded by solve_with_capsolver_fixed
    # def solve_with_capsolver(self):
    #     """
    #     Solve MTCaptcha using Capsolver (free tier available)
    #     """
    #     try:
    #         import requests
    #         import json
    #         import time
    #
    #         # Capsolver API (get free API key from https://www.capsolver.com/)
    #         API_KEY = "CAP-D051815FD86B044D03BF198CC9DFEB4B"  # Register for free account
    #
    #         # Step 1: Create task
    #         task_data = {
    #             "clientKey": API_KEY,
    #             "task": {
    #                 "type": "MTCaptchaTask",
    #                 "websiteURL": self.driver.current_url,
    #                 "websiteKey": "MTPublic-JnEM38Q6U",
    #                 "proxy": ""  # No proxy needed
    #             }
    #         }
    #
    #         response = requests.post("https://api.capsolver.com/createTask", json=task_data)
    #         task_result = response.json()
    #
    #         if task_result['errorId'] != 0:
    #             print(f"[  WARN  ] Capsolver error: {task_result['errorDescription']}")
    #             return False
    #
    #         task_id = task_result['taskId']
    #
    #         # Step 2: Poll for result
    #         for _ in range(30):  # 30 attempts with 2-second intervals
    #             time.sleep(2)
    #
    #             result_data = {
    #                 "clientKey": API_KEY,
    #                 "taskId": task_id
    #             }
    #
    #             result_response = requests.post("https://api.capsolver.com/getTaskResult", json=result_data)
    #             result = result_response.json()
    #
    #             if result['status'] == "ready":
    #                 # Inject the solution
    #                 self.driver.execute_script(f"""
    #                     document.querySelector('[name="mtcaptcha-verifiedtoken"]').value = '{result['solution']['token']}';
    #                 """)
    #                 print("[   OK   ] MTCaptcha solved via Capsolver")
    #                 return True
    #
    #         print("[  WARN  ] Capsolver timeout")
    #         return False
    #
    #     except Exception as e:
    #         print(f"[  WARN  ] Capsolver failed: {str(e)}")
    #         return False

    def solve_audio_captcha(self):
        """
        Solve MTCaptcha using audio transcription (SpeechRecognition + pydub)
        """
        try:
            import os

            import requests
            import speech_recognition as sr
            from pydub import AudioSegment

            print("[  INFO  ] Attempting audio captcha solving...")

            # Ensure we are in the iframe
            try:
                iframe = self.driver.find_element(By.ID, "register-captcha-iframe-1")
                self.driver.switch_to.frame(iframe)
            except:
                pass  # Already in iframe or failed

            # Find and click audio button
            try:
                audio_btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.CLASS_NAME, "mtcaptcha-audio-button")
                    )
                )
                audio_btn.click()
                time.sleep(2)
            except Exception as e:
                print(f"[  WARN  ] Could not find or click audio button: {e}")
                # Switch back if we failed inside iframe
                self.driver.switch_to.default_content()
                return False

            # Find audio source URL
            try:
                audio_elem = self.driver.find_element(By.ID, "mtcap-audio-1")
                audio_url = audio_elem.get_attribute("src")
            except:
                print("[  WARN  ] Could not find audio source element")
                self.driver.switch_to.default_content()
                return False

            if not audio_url:
                print("[  WARN  ] Audio URL is empty")
                self.driver.switch_to.default_content()
                return False

            # Download audio
            try:
                audio_response = requests.get(audio_url, timeout=10)
                if audio_response.status_code != 200:
                    print("[  WARN  ] Failed to download audio")
                    self.driver.switch_to.default_content()
                    return False
            except Exception as e:
                print(f"[  WARN  ] Failed to download audio: {e}")
                self.driver.switch_to.default_content()
                return False

            # Save temp file
            try:
                with open("captcha_audio.mp3", "wb") as f:
                    f.write(audio_response.content)
            except Exception as e:
                print(f"[  WARN  ] Failed to save audio file: {e}")
                self.driver.switch_to.default_content()
                return False

            # Convert to WAV
            try:
                sound = AudioSegment.from_file("captcha_audio.mp3")
                sound.export("captcha_audio.wav", format="wav")
            except Exception as e:
                # pydub often requires ffmpeg or libav
                print(f"[  WARN  ] Audio conversion failed (ffmpeg missing?): {e}")
                self.driver.switch_to.default_content()
                return False

            # Transcribe
            try:
                r = sr.Recognizer()
                with sr.AudioFile("captcha_audio.wav") as source:
                    audio_data = r.record(source)
                    # Recognize (using Google Speech Recognition)
                    text = r.recognize_google(audio_data)
            except Exception as e:
                print(f"[  WARN  ] Transcription failed: {e}")
                self.driver.switch_to.default_content()
                return False

            print(f"[  INFO  ] Audio transcribed: {text}")

            # Clean text (alphanumeric only)
            text = "".join(filter(str.isalnum, text))

            # Submit
            try:
                input_field = self.driver.find_element(By.ID, "mtcap-inputtext-1")
                input_field.clear()
                input_field.send_keys(text)
                time.sleep(1)

                self.driver.find_element(By.ID, "mtcap-statusbutton-1").click()
                time.sleep(2)

                # Verify
                success = False
                try:
                    status_elem = self.driver.find_element(By.ID, "mtcap-status-1")
                    if "success" in status_elem.get_attribute("class").lower():
                        success = True
                except:
                    pass

                # Double check token
                if not success:
                    try:
                        token_val = self.driver.execute_script(
                            "return (document.querySelector('[name=\"mtcaptcha-verifiedtoken\"]') || {}).value || ''"
                        )
                        if token_val and token_val.strip() != "":
                            success = True
                    except:
                        pass

                if success:
                    self.driver.switch_to.default_content()
                    print(f"[   OK   ] Audio captcha solved: {text}")
                    return True
                else:
                    # Save screenshot for debug
                    try:
                        self.driver.save_screenshot("audio_solve_failed.png")
                        print("[  INFO  ] Saved screenshot to audio_solve_failed.png")
                    except:
                        pass
            except Exception as e:
                print(f"[  WARN  ] Error submitting audio solution: {e}")

            self.driver.switch_to.default_content()
            return False

        except Exception as e:
            try:
                self.driver.switch_to.default_content()
            except:
                pass
            print(f"[  WARN  ] Audio captcha failed: {e}")
            return False
        finally:
            # Cleanup
            try:
                if os.path.exists("captcha_audio.mp3"):
                    os.remove("captcha_audio.mp3")
                if os.path.exists("captcha_audio.wav"):
                    os.remove("captcha_audio.wav")
            except:
                pass

    def solve_with_ddddocr(self, retries=3):
        """
        Solve MTCaptcha using ddddocr (lightweight, offline, open-source)
        Retries specified number of times if standard validation fails.
        """
        try:
            import base64

            import ddddocr

            print(f"[  INFO  ] Attempting ddddocr solving (max {retries} attempts)...")
            ocr = ddddocr.DdddOcr(show_ad=False)

            for attempt in range(retries):
                # Wait for iframe
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "register-captcha-iframe-1"))
                )

                # Switch to iframe
                iframe = self.driver.find_element(By.ID, "register-captcha-iframe-1")
                self.driver.switch_to.frame(iframe)
                time.sleep(1)

                # Get captcha image
                try:
                    captcha_img = self.driver.find_element(By.ID, "mtcap-image-1")
                    style = captcha_img.get_attribute("style")
                    base64_data = None
                    if "base64," in style:
                        import re

                        base64_match = re.search(r'base64,([^"\']+)', style)
                        if base64_match:
                            base64_data = base64_match.group(1)
                except Exception:
                    base64_data = None

                if not base64_data:
                    self.driver.switch_to.default_content()  # Exit iframe before clicking refresh
                    # If can't get image, try refreshing
                    if attempt < retries - 1:
                        print(
                            f"[  WARN  ] Could not get image, refreshing... ({attempt+1}/{retries})"
                        )
                        try:
                            refresh_btn = self.driver.find_element(
                                By.CSS_SELECTOR, ".mtcaptcha-reload-button"
                            )
                            refresh_btn.click()
                            time.sleep(2)
                        except:
                            pass
                        continue
                    else:
                        return False

                # Solve with ddddocr
                img_bytes = base64.b64decode(base64_data)
                captcha_text = ocr.classification(img_bytes)

                # Clean text (alphanumeric only)
                captcha_text = "".join(filter(str.isalnum, captcha_text))

                if captcha_text and 4 <= len(captcha_text) <= 8:
                    input_field = self.driver.find_element(By.ID, "mtcap-inputtext-1")
                    input_field.clear()

                    for char in captcha_text:
                        input_field.send_keys(char)
                        time.sleep(0.05)

                    time.sleep(0.5)

                    # Click verify
                    self.driver.find_element(By.ID, "mtcap-statusbutton-1").click()

                    time.sleep(2)

                    # Check status
                    success = False

                    # 1. Check verified token (Most reliable)
                    try:
                        token_val = self.driver.execute_script(
                            "return (document.querySelector('[name=\"mtcaptcha-verifiedtoken\"]') || {}).value || ''"
                        )
                        print(f"[ DEBUG  ] Token value: '{token_val}'")
                        if token_val and token_val.strip() != "":
                            success = True
                            print("[ DEBUG  ] Valid token found.")
                    except Exception as e:
                        print(f"[ DEBUG  ] Token check error: {e}")

                    # 2. Check visual status (Green tick / Success message)
                    if not success:
                        try:
                            # Check for status element class
                            status_elem = self.driver.find_element(
                                By.ID, "mtcap-status-1"
                            )
                            status_class = status_elem.get_attribute("class").lower()
                            print(f"[ DEBUG  ] Status element class: '{status_class}'")
                            if "success" in status_class:
                                success = True
                                print("[ DEBUG  ] Success class found.")

                            # Check for 'verified' text or green tick specific elements if class check fails
                            if "verified" in status_class:
                                success = True
                                print("[ DEBUG  ] Verified class found.")
                        except:
                            pass

                    # Enhanced Validation: Check for Green Tick Style and Aria Label (from User feedback/Debug)
                    if not success:
                        try:
                            # Check mtcap-statusimg-1 style for green color
                            status_img = self.driver.find_element(
                                By.ID, "mtcap-statusimg-1"
                            )
                            style = status_img.get_attribute("style")
                            if (
                                "rgb(141, 198, 64)" in style
                                or "color: #8dc640" in style
                            ):
                                success = True
                                print(
                                    "[ DEBUG  ] Visual success (green tick color) detected!"
                                )
                        except:
                            pass

                    if not success:
                        try:
                            # Check Aria Label
                            desc_elem = self.driver.find_element(
                                By.ID, "desc4StatusButton-1"
                            )
                            desc_text = (
                                desc_elem.get_attribute("innerHTML") or desc_elem.text
                            )
                            if "verified successfully" in desc_text.lower():
                                success = True
                                print("[ DEBUG  ] Aria label indicates success!")
                        except:
                            pass

                    if not success:
                        # Double Check token value as a final resort
                        try:
                            token_val = self.driver.execute_script(
                                "return (document.querySelector('[name=\"mtcaptcha-verifiedtoken\"]') || {}).value || ''"
                            )
                            if token_val:
                                success = True
                                print(
                                    f"[ DEBUG  ] Token found via JS: {token_val[:15]}..."
                                )
                        except:
                            pass

                    self.driver.switch_to.default_content()

                    if success:
                        print(f"[   OK   ] ddddocr solved: {captcha_text}")
                        return True
                    else:
                        print(
                            f"[  WARN  ] ddddocr failed validation (text: {captcha_text})"
                        )
                        # Debug artifact
                        try:
                            self.driver.switch_to.frame(
                                self.driver.find_element(
                                    By.ID, "register-captcha-iframe-1"
                                )
                            )
                            self.driver.save_screenshot(
                                f"ddddocr_fail_debug_{attempt}.png"
                            )
                            with open(
                                f"ddddocr_fail_debug_{attempt}.html",
                                "w",
                                encoding="utf-8",
                            ) as f:
                                f.write(self.driver.page_source)
                            print(
                                f"[ DEBUG  ] Saved debug artifacts for attempt {attempt}"
                            )
                            self.driver.switch_to.default_content()
                        except:
                            self.driver.switch_to.default_content()

                else:
                    self.driver.switch_to.default_content()
                    print(f"[  WARN  ] ddddocr text invalid length: {captcha_text}")

                # If we are here, it failed. Refresh and try again if attempts remain
                if attempt < retries - 1:
                    print(f"[  INFO  ] Retrying ddddocr... ({attempt+1}/{retries})")
                    try:
                        # Ensure we are in default content
                        refresh_btn = self.driver.find_element(
                            By.CSS_SELECTOR, ".mtcaptcha-reload-button"
                        )
                        refresh_btn.click()
                        time.sleep(2)
                    except Exception as e:
                        print(f"[  WARN  ] Could not click refresh: {e}")

            print(f"[  WARN  ] ddddocr failed after {retries} attempts")

            # Fallback to audio (only if enabled)
            if os.environ.get("ENABLE_AUDIO_CAPTCHA", "false").lower() == "true":
                print("[  INFO  ] Visual solve failed, attempting audio fallback...")
                if self.solve_audio_captcha():
                    return True
            else:
                print(
                    "[  INFO  ] Audio fallback disabled (ENABLE_AUDIO_CAPTCHA != true)"
                )

            return False

        except ImportError:
            self.driver.switch_to.default_content()
            print(
                "[  WARN  ] ddddocr module not found. Install with: pip install ddddocr"
            )
            return False
        except Exception as e:
            try:
                self.driver.switch_to.default_content()
            except:
                pass
            print(f"[  WARN  ] ddddocr error: {str(e)}")
            return False

    #         return None

    def refresh_until_easy_captcha(self, max_attempts=10):
        """
        Refresh captcha until an easily readable one appears or an automated solver succeeds.
        Will attempt OCR, enhanced OCR, Capsolver and TwoCaptcha after each refresh.
        """
        for attempt in range(max_attempts):
            try:
                # Find refresh button and click it
                refresh_btn = self.driver.find_element(
                    By.CSS_SELECTOR, ".mtcaptcha-reload-button"
                )
                refresh_btn.click()
                # random small wait to simulate human
                time.sleep(1 + random.random() * 2)

                # Try in-order: ddddocr -> simple OCR -> enhanced -> capsolver -> 2captcha
                if self.solve_with_ddddocr():
                    return True
                if self.solve_mtcaptcha_simple_ocr():
                    return True
                if self.solve_mtcaptcha_enhanced_ocr():
                    return True
                if self.solve_with_capsolver_fixed():
                    return True
                if self.solve_mtcaptcha_with_service_simple():
                    return True

            except Exception as e:
                # If refresh button isn't present, exit early
                # but continue trying solvers directly
                if hasattr(e, "name"):
                    pass
                try:
                    if self.solve_mtcaptcha_simple_ocr():
                        return True
                    if self.solve_mtcaptcha_enhanced_ocr():
                        return True
                    if self.solve_with_capsolver_fixed():
                        return True
                    if self.solve_mtcaptcha_with_service_simple():
                        return True
                except:
                    pass

        print(
            "[  WARN  ] Could not find easy captcha after multiple refreshes or automated solvers failed"
        )
        return False

    def createAccount(self):
        exec_js = self.driver.execute_script
        uCE = untilConditionExecute
        # STEP 0

        logging.info("Loading ESET ProtectHub Page...")
        console_log("\nLoading ESET ProtectHub Page...", INFO, silent_mode=SILENT_MODE)
        if isinstance(self.email_obj, WEB_WRAPPER_EMAIL_APIS_CLASSES):
            self.driver.switch_to.new_window("tab")
            self.window_handle = self.driver.current_window_handle
        self.driver.get("https://protecthub.eset.com/public/registration?culture=en-US")
        uCE(self.driver, f'return {GET_EBID}("continue") != null')
        logging.info("Successfully!")
        console_log("Successfully!", OK, silent_mode=SILENT_MODE)

        # STEP 1
        logging.info("Data filling...")
        console_log("\nData filling...", INFO, silent_mode=SILENT_MODE)
        exec_js(f'return {GET_EBID}("email-input")').send_keys(self.email_obj.email)
        exec_js(f'return {GET_EBID}("company-name-input")').send_keys(dataGenerator(10))
        # Select country
        exec_js(f"return {GET_EBID}('country-select')").click()
        selected_country = "Ukraine"
        logging.info("Selecting the country...")
        # Find and click country dropdown
        try:
            # Wait for the select component to be present
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(("id", "country-select"))
            )

            # Find the input field inside the React Select
            country_input = self.driver.find_element("id", "country-select-input")

            # Click to open the dropdown
            country_input.click()
            time.sleep(0.5)

            # Clear any existing value and type the country name
            country_input.clear()
            country_input.send_keys(
                selected_country
            )  # e.g., "Ukraine", "United States", etc.
            time.sleep(0.5)

            # Press Enter to select the first matching option
            country_input.send_keys(Keys.ENTER)
            time.sleep(0.5)

            print(f"[   OK   ] Country '{selected_country}' selected successfully!")

        except Exception as e:
            print(f"[ FAILED ] Could not select country: {str(e)}")
            # Optionally, you can skip this step or use a default
            pass
        exec_js(f'return {GET_EBID}("company-vat-input")').send_keys(
            dataGenerator(10, True)
        )
        exec_js(f'return {GET_EBID}("company-crn-input")').send_keys(
            dataGenerator(10, True)
        )

        # After country selection, before the manual captcha prompt
        # Allow override of solver order via environment variable CAPTCHA_SERVICE (capsolver|2captcha|ocr|auto)
        import os

        captcha_service = os.environ.get("CAPTCHA_SERVICE", "auto").lower()

        def try_order(order_list):
            for name in order_list:
                if name == "ddddocr" and self.solve_with_ddddocr():
                    return True
                if name == "ocr" and self.solve_mtcaptcha_simple_ocr():

                    return True
                if name == "enhanced_ocr" and self.solve_mtcaptcha_enhanced_ocr():
                    return True
                if name == "refresh" and self.refresh_until_easy_captcha():
                    return True
                if name == "capsolver" and self.solve_with_capsolver_fixed():
                    return True
                if name == "2captcha" and self.solve_mtcaptcha_with_service_simple():
                    return True
            return False

        if captcha_service == "capsolver":
            order = [
                "capsolver",
                "2captcha",
                "refresh",
                "ddddocr",
                "enhanced_ocr",
                "ocr",
            ]
        elif captcha_service in ("2captcha", "twocaptcha"):
            order = [
                "2captcha",
                "capsolver",
                "refresh",
                "ddddocr",
                "enhanced_ocr",
                "ocr",
            ]
        elif captcha_service == "ocr":
            order = [
                "ocr",
                "enhanced_ocr",
                "ddddocr",
                "refresh",
                "capsolver",
                "2captcha",
            ]
        else:  # auto
            order = [
                "ddddocr",
                "ocr",
                "enhanced_ocr",
                "refresh",
                "capsolver",
                "2captcha",
            ]

        solved = try_order(order)

        # Final fallback to manual
        if not solved:
            logging.warning("Solve the captcha on the page manually!!!")
            console_log(
                f"\n{colorama.Fore.CYAN}Solve the captcha on the page manually!!!{colorama.Fore.RESET}",
                INFO,
                False,
                SILENT_MODE,
            )
            while True:  # captcha
                try:
                    mtcaptcha_solved_token = exec_js(
                        f'return {GET_EBCN}("mtcaptcha-verifiedtoken")[0].value'
                    )
                    if mtcaptcha_solved_token.strip() != "":
                        break
                except Exception as E:
                    pass
                time.sleep(1)
        exec_js(f'return {GET_EBID}("continue").click()')
        try:
            uCE(
                self.driver,
                f'return {GET_EBID}("registration-email-sent").innerText === "We sent you a verification email"',
                max_iter=10,
            )
            logging.info("Successfully!")
            console_log("Successfully!", OK, silent_mode=SILENT_MODE)
        except:
            raise IPBlockedException(
                "\nESET temporarily blocked your IP, try again later!!! Try to use VPN/Proxy or try to change Email API!!!"
            )
        return True

    def activateAccount(self):
        exec_js = self.driver.execute_script
        uCE = untilConditionExecute

        # STEP 1
        logging.info("Data filling...")
        console_log("\nData filling...", INFO, silent_mode=SILENT_MODE)
        exec_js(f'return {GET_EBID}("first-name-input")').send_keys(dataGenerator(10))
        exec_js(f'return {GET_EBID}("last-name-input")').send_keys(dataGenerator(10))
        exec_js(f'return {GET_EBID}("first-name-input")').send_keys(dataGenerator(10))
        exec_js(f'return {GET_EBID}("password-input")').send_keys(self.eset_password)
        exec_js(f'return {GET_EBID}("password-repeat-input")').send_keys(
            self.eset_password
        )
        exec_js(f'return {GET_EBID}("continue").click()')

        # STEP 2
        uCE(self.driver, f'return {GET_EBID}("phone-input") != null')
        exec_js(f'return {GET_EBID}("phone-input")').send_keys(dataGenerator(10, True))
        time.sleep(0.5)
        exec_js(f'return {GET_EBID}("continue").click()')
        uCE(
            self.driver,
            f'return {GET_EBID}("activated-user-title").innerText === "Your account has been successfully activated"',
            max_iter=15,
        )
        logging.info("Successfully!")
        console_log("Successfully!", OK, silent_mode=SILENT_MODE)

    def confirmAccount(self):
        if isinstance(self.email_obj, CustomEmailAPI):
            token = parseToken(
                self.email_obj, eset_business=True, max_iter=100, delay=3
            )
        else:
            logging.info(
                f"[{self.email_obj.class_name}] ProtectHub-Token interception..."
            )
            console_log(
                f"\n[{self.email_obj.class_name}] ProtectHub-Token interception...",
                INFO,
                silent_mode=SILENT_MODE,
            )
            if isinstance(self.email_obj, WEB_WRAPPER_EMAIL_APIS_CLASSES):
                token = parseToken(
                    self.email_obj, self.driver, True, max_iter=100, delay=3
                )
                self.driver.switch_to.window(self.window_handle)
            else:
                token = parseToken(
                    self.email_obj, eset_business=True, max_iter=100, delay=3
                )  # 1secmail
        logging.info(f"ProtectHub-Token: {token}")
        logging.info("Account confirmation is in progress...")
        console_log(f"ProtectHub-Token: {token}", OK, silent_mode=SILENT_MODE)
        console_log(
            "\nAccount confirmation is in progress...", INFO, silent_mode=SILENT_MODE
        )
        self.driver.get(
            f"https://protecthub.eset.com/public/activation/{token}/?culture=en-US"
        )
        untilConditionExecute(
            self.driver, f'return {GET_EBID}("first-name-input") != null'
        )
        logging.info("Account successfully confirmed!")
        console_log("Account successfully confirmed!", OK, silent_mode=SILENT_MODE)


class EsetProtectHubKeygen(object):
    def __init__(
        self, registered_email_obj: OneSecEmailAPI, eset_password: str, driver: Chrome
    ):
        self.email_obj = registered_email_obj
        self.eset_password = eset_password
        self.driver = driver

    def getLD(self):
        exec_js = self.driver.execute_script
        uCE = untilConditionExecute

        # Log in
        logging.info("Logging in to the created account...")
        console_log(
            "\nLogging in to the created account...", INFO, silent_mode=SILENT_MODE
        )
        self.driver.get("https://protecthub.eset.com")
        uCE(self.driver, f'return {GET_EBID}("username") != null')
        exec_js(f'return {GET_EBID}("username")').send_keys(self.email_obj.email)
        exec_js(f'return {GET_EBID}("password")').send_keys(self.eset_password)
        exec_js(f'return {GET_EBID}("btn-login").click()')

        # Start free trial
        uCE(
            self.driver,
            f'return {GET_EBID}("welcome-dialog-generate-trial-license") != null',
            delay=3,
        )
        logging.info("Successfully!")
        logging.info("Sending a request for a get license...")
        console_log("Successfully!", OK, silent_mode=SILENT_MODE)
        console_log(
            "\nSending a request for a get license...", INFO, silent_mode=SILENT_MODE
        )
        try:
            exec_js(
                f'return {GET_EBID}("welcome-dialog-generate-trial-license").click()'
            )
            exec_js(
                f'return {GET_EBID}("welcome-dialog-generate-trial-license")'
            ).click()
        except:
            pass

        # Waiting for a response from the site
        license_is_being_generated = False
        for _ in range(DEFAULT_MAX_ITER):
            try:
                r = exec_js(
                    f"return {GET_EBCN}('Toastify__toast-body toastBody')[0].innerText"
                ).lower()
                if r.find("is being generated") != -1:
                    license_is_being_generated = True
                    logging.info("Request successfully sent!")
                    console_log(
                        "Request successfully sent!", OK, silent_mode=SILENT_MODE
                    )
                    try:
                        exec_js(
                            f'return {GET_EBID}("welcome-dialog-skip-button").click()'
                        )
                        exec_js(
                            f'return {GET_EBID}("welcome-dialog-skip-button")'
                        ).click()
                    except:
                        pass
                    break
            except Exception as E:
                pass
            time.sleep(DEFAULT_DELAY)

        if not license_is_being_generated:
            raise RuntimeError("The request has not been sent!")

        logging.info("Waiting for a back response...")
        console_log("\nWaiting for a back response...", INFO, silent_mode=SILENT_MODE)
        license_was_generated = False
        for _ in range(DEFAULT_MAX_ITER * 10):  # 5m
            try:
                r = exec_js(
                    f"return {GET_EBCN}('Toastify__toast-body toastBody')[0].innerText"
                ).lower()
                if r.find("couldn't be generated") != -1:
                    break
                elif r.find("was generated") != -1:
                    logging.info("Successfully!")
                    console_log("Successfully!", OK, silent_mode=SILENT_MODE)
                    license_was_generated = True
                    break
            except Exception as E:
                pass
            time.sleep(DEFAULT_DELAY)

        if not license_was_generated:
            raise RuntimeError("The license cannot be generated, try again later!")

        # Obtaining license data from the site
        logging.info("[Site] License uploads...")
        console_log("\n[Site] License uploads...", INFO, silent_mode=SILENT_MODE)
        license_name = "ESET PROTECT Advanced"
        try:
            self.driver.get("https://protecthub.eset.com/licenses")
            uCE(
                self.driver,
                f'return {GET_EBAV}("div", "data-label", "license-list-body-cell-renderer-row-0-column-0").innerText != ""',
            )
            license_id = exec_js(
                f'{DEFINE_GET_EBAV_FUNCTION}\nreturn {GET_EBAV}("div", "data-label", "license-list-body-cell-renderer-row-0-column-0").innerText'
            )
            logging.info(f"License ID: {license_id}")
            logging.info("Getting information from the license...")
            console_log(f"License ID: {license_id}", OK, silent_mode=SILENT_MODE)
            console_log(
                "\nGetting information from the license...",
                INFO,
                silent_mode=SILENT_MODE,
            )
            self.driver.get(
                f"https://protecthub.eset.com/licenses/details/2/{license_id}/overview"
            )
            uCE(
                self.driver,
                f'return {GET_EBAV}("div", "data-label", "license-overview-validity-value") != null',
            )
            license_out_date = exec_js(
                f'{DEFINE_GET_EBAV_FUNCTION}\nreturn {GET_EBAV}("div", "data-label", "license-overview-validity-value").children[0].children[0].innerText'
            )
            # Obtaining license key
            exec_js(
                f'{DEFINE_GET_EBAV_FUNCTION}\n{GET_EBAV}("div", "data-label", "license-overview-key-value").children[0].children[0].click()'
            )
            logging.info("Waiting for password modal...")
            uCE(
                self.driver,
                f'return {GET_EBID}("show-license-key-auth-modal-password-input") != null',
                max_iter=10,
            )
            logging.info("Entering password...")
            password_input = exec_js(
                f'return {GET_EBID}("show-license-key-auth-modal-password-input")'
            )
            password_input.clear()
            password_input.send_keys(self.eset_password)
            time.sleep(0.5)
            logging.info("Password entered")
            logging.info("Clicking authenticate button...")
            try:
                auth_button = exec_js(
                    f'return {GET_EBID}("show-license-key-auth-modal-authenticate")'
                )
                if auth_button:
                    auth_button.click()
                    logging.info("Authenticate button clicked")
                else:
                    logging.warning("Authenticate button not found by ID")
            except Exception as e:
                logging.error(f"Failed to click authenticate button: {e}")
                pass
            for _ in range(DEFAULT_MAX_ITER):
                try:
                    license_key = exec_js(
                        f'return {GET_EBAV}("div", "data-label", "license-overview-key-value").children[0].textContent.trim()'
                    )
                    if license_key is not None and not license_key.startswith(
                        "XXXX-XXXX-XXXX-XXXX-XXXX"
                    ):  # ignoring XXXX-XXXX-XXXX-XXXX-XXXX
                        license_key = license_key.split(" ")[0]
                        logging.info("Information successfully received!")
                        console_log(
                            "Information successfully received!",
                            OK,
                            silent_mode=SILENT_MODE,
                        )
                        return (
                            license_name,
                            license_key,
                            license_out_date,
                            True,
                        )  # True - License key obtained from the site
                except:
                    pass
                time.sleep(DEFAULT_DELAY)
        except Exception as E:
            logging.critical("EXC_INFO:", exc_info=True)
            console_log(
                "Error when obtaining a license key from the site!!!",
                ERROR,
                silent_mode=SILENT_MODE,
            )
        # Obtaining license data from the email
        logging.info("[Email] License uploads...")
        console_log("\n[Email] License uploads...", INFO, silent_mode=SILENT_MODE)
        if self.email_obj.class_name == "custom":
            logging.warning(
                "Wait for a message to your e-mail about successful key generation!!!"
            )
            console_log(
                "\nWait for a message to your e-mail about successful key generation!!!",
                WARN,
                True,
                SILENT_MODE,
            )
            return None, None, None, None
        else:
            license_key, license_out_date, license_id = parseEPHKey(
                self.email_obj, self.driver, delay=5, max_iter=30
            )  # 2.5m
            logging.info(f"License ID: {license_id}")
            logging.info("Getting information from the license...")
            logging.info("Information successfully received!")
            console_log(f"License ID: {license_id}", OK, silent_mode=SILENT_MODE)
            console_log(
                "\nGetting information from the license...",
                INFO,
                silent_mode=SILENT_MODE,
            )
            console_log(
                "Information successfully received!", OK, silent_mode=SILENT_MODE
            )
            return (
                license_name,
                license_key,
                license_out_date,
                False,
            )  # False - License key obtained from the email

    def removeLicense(self):
        logging.info("Deleting the key from the account, the key will still work...")
        console_log(
            "Deleting the key from the account, the key will still work...",
            INFO,
            silent_mode=SILENT_MODE,
        )
        try:
            self.driver.execute_script(
                f'return {GET_EBID}("license-actions-button")'
            ).click()
            time.sleep(1)
            button = self.driver.find_element(
                "xpath", '//a[.//div[text()="Remove license"]]'
            )
            if button is not None:
                button.click()
            untilConditionExecute(
                self.driver,
                f'return {CLICK_WITH_BOOL}({GET_EBID}("remove-license-dlg-remove-btn"))',
                max_iter=15,
            )
            time.sleep(2)
            for _ in range(DEFAULT_MAX_ITER // 2):
                try:
                    self.driver.execute_script(
                        f'return {GET_EBID}("remove-license-dlg-remove-btn")'
                    ).click()
                except:
                    pass
                if (
                    self.driver.page_source.lower().find(
                        "to keep the solutions up to date"
                    )
                    == -1
                ):
                    time.sleep(1)
                    logging.info("Key successfully deleted!!!")
                    console_log(
                        "Key successfully deleted!!!", OK, silent_mode=SILENT_MODE
                    )
                    return True
                time.sleep(DEFAULT_DELAY)
        except:
            pass
        logging.error(
            "Failed to delete key, this error has no effect on the operation of the key!!!"
        )
        console_log(
            "Failed to delete key, this error has no effect on the operation of the key!!!",
            ERROR,
            silent_mode=SILENT_MODE,
        )


def EsetVPNResetWindows(key_path="SOFTWARE\\ESET\\ESET VPN", value_name="authHash"):
    """Deletes the authHash value of ESET VPN"""
    try:
        subprocess.check_output(
            ["taskkill", "/f", "/im", "esetvpn.exe"], stderr=subprocess.DEVNULL
        )
    except:
        pass
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS
        ) as key:
            winreg.DeleteValue(key, value_name)
        logging.info("ESET VPN has been successfully reset!!!")
        console_log(
            "ESET VPN has been successfully reset!!!", OK, silent_mode=SILENT_MODE
        )
    except FileNotFoundError:
        logging.error(
            f"The registry value or key does not exist: {key_path}\\{value_name}"
        )
        console_log(
            f"The registry value or key does not exist: {key_path}\\{value_name}",
            ERROR,
            silent_mode=SILENT_MODE,
        )
    except PermissionError:
        logging.error(f"Permission denied while accessing: {key_path}\\{value_name}")
        console_log(
            f"Permission denied while accessing: {key_path}\\{value_name}",
            ERROR,
            silent_mode=SILENT_MODE,
        )
    except Exception as e:
        raise RuntimeError(e)


def EsetVPNResetMacOS(
    app_name="ESET VPN", file_name="Preferences/com.eset.ESET VPN.plist"
):
    try:
        # Use AppleScript to quit the application
        script = f'tell application "{app_name}" to quit'
        subprocess.run(
            ["osascript", "-e", script],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except:
        pass
    try:
        time.sleep(2)
        # Get the full path to the file in the Library folder
        library_path = Path.home() / "Library" / file_name
        # Check if the file exists and remove it
        if library_path.is_file():
            library_path.unlink()
            logging.info("ESET VPN has been successfully reset!!!")
            console_log(
                "ESET VPN has been successfully reset!!!", OK, silent_mode=SILENT_MODE
            )
        else:
            logging.error(f"File '{file_name}' does not exist!!!")
            console_log(
                f"File '{file_name}' does not exist!!!", ERROR, silent_mode=SILENT_MODE
            )
    except Exception as e:
        raise RuntimeError(e)
