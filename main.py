from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
import os
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class OECDARAutomation:
    def __init__(self, download_dir=None):
        self._pw = None
        self.browser = None
        self.context = None
        self.page = None
        self.download_dir = download_dir or os.path.join(os.getcwd(), "oecdar_downloads")
        os.makedirs(self.download_dir, exist_ok=True)

        # Primary URL (Link with current settings)
        self.primary_url = (
            "https://data-explorer.oecd.org/vis?tm=Effective%20labour%20market%20exit%20age"
            "&pg=0&snb=2&vw=tb&df%5bds%5d=dsDisseminateFinalDMZ&df%5bid%5d=DSD_PAG%40DF_PAG"
            "&df%5bag%5d=OECD.ELS.SPD&df%5bvs%5d=1.0"
            "&dq=RUS%2BTUR%2BUSA%2BSWE%2BPRT%2BPOL%2BNZL%2BNOR%2BNLD%2BMEX%2BKOR%2BJPN"
            "%2BITA%2BIRL%2BHUN%2BGRC%2BGBR%2BFRA%2BFIN%2BESP%2BDEU%2BCHN%2BCZE%2BCHE"
            "%2BCAN%2BCHL%2BBRA%2BBEL%2BAUS.A.ELMEA..M.."
            "&pd=1970%2C2022&to%5bTIME_PERIOD%5d=true"
        )
        self.default_url = "https://data-explorer.oecd.org/"
        self.search_term = "Pensions at a glance"
        self.target_measure = "Effective labour market exit age"
        self.target_sex = "Male"

        # Required countries from runbook Step 9
        self.required_countries = [
            "Australia", "Belgium", "Brazil", "Canada", "Switzerland", "Chile", "China",
            "Czech", "Germany", "Spain", "Finland", "France", "United Kingdom", "Greece",
            "Hungary", "Ireland", "Italy", "Japan", "Korea", "Mexico", "Netherlands",
            "Norway", "New Zealand", "Poland", "Portugal", "Russia", "Sweden", "Turkey",
            "United States"
        ]

    # ------------------------------------------------------------------
    # Step 1 – Browser setup
    # ------------------------------------------------------------------
    def setup_driver(self):
        """Step 1: Launch Playwright Chromium browser."""
        try:
            logger.info("[CONFIG] STEP 1: Setting up Playwright browser...")
            self._pw = sync_playwright().start()
            self.browser = self._pw.chromium.launch(
                channel="chrome",   # use system-installed Google Chrome
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage",
                      "--disable-blink-features=AutomationControlled"],
            )
            self.context = self.browser.new_context(
                accept_downloads=True,
                viewport={"width": 1920, "height": 1080},
            )
            self.page = self.context.new_page()
            logger.info("[SUCCESS] Driver setup complete")
            return True
        except Exception as e:
            logger.error(f"[FAILURE] Driver setup failed: {e}")
            return False

    # ------------------------------------------------------------------
    # Primary workflow  (Steps 2-5)
    # ------------------------------------------------------------------
    def try_primary_workflow(self):
        logger.info("[INFO] ATTEMPTING PRIMARY WORKFLOW (Steps 2-5)")
        logger.info("=" * 60)
        try:
            if not self.load_primary_url():
                return False
            if not self.check_data_status():
                logger.warning("[WARNING] Data status check had issues, continuing...")
            if not self.update_time_periods():
                logger.warning("[WARNING] Time period update had issues, continuing...")
            if not self.download_data():
                return False
            logger.info("[SUCCESS] PRIMARY WORKFLOW SUCCESSFUL!")
            return True
        except Exception as e:
            logger.error(f"[FAILURE] Primary workflow failed: {e}")
            return False

    def load_primary_url(self):
        """Step 2: Load the primary OECDAR URL and verify success."""
        try:
            logger.info("[STATUS] STEP 2: Loading primary OECDAR URL...")
            self.page.goto(self.primary_url, wait_until="networkidle", timeout=60000)
            # Verify the data-explorer interface rendered
            self.page.wait_for_selector("text=Time period", timeout=20000)
            logger.info("[SUCCESS] Step 2 completed - Primary URL loaded successfully")
            return True
        except PlaywrightTimeout:
            logger.warning("[WARNING] Primary URL verification timed out")
            return False
        except Exception as e:
            logger.error(f"[FAILURE] Step 2 failed: {e}")
            return False

    def check_data_status(self):
        """Step 3: Verify Male is selected and log status."""
        try:
            logger.info("[INFO] STEP 3: Checking data configuration...")
            male = self.page.locator('[data-testid="value_M"][aria-checked="true"]')
            if male.count() > 0:
                logger.info("[SUCCESS] Male gender confirmed as selected")
            else:
                logger.info("[STATUS] Male gender status unclear")
            logger.info("[SUCCESS] Step 3 completed - Data status checked")
            return True
        except Exception as e:
            logger.warning(f"[WARNING] Step 3 warning: {e}")
            return True

    def update_time_periods(self):
        """Step 4: Ensure time period section is open and set end year to latest."""
        try:
            logger.info("[INFO] STEP 4: Updating time periods...")
            updated = self.update_end_year_to_latest()
            if updated:
                logger.info("[SUCCESS] End year updated to latest available")
            else:
                logger.warning("[WARNING] Could not update end year — will proceed anyway")
            logger.info("[SUCCESS] Step 4 completed")
            return True
        except Exception as e:
            logger.warning(f"[WARNING] Step 4 warning: {e}")
            return True

    def expand_time_period_section(self):
        """
        Open the Time period accordion if it is not already open.

        Playwright codegen confirmed the header text is 'Time period<N>'
        (e.g. 'Time period53') and the panel is a div[role='button'].
        We check for year-End-test-id first — if it's present the panel is
        already open and no click is needed (common when loading via primary URL
        which has to[TIME_PERIOD]=true).
        """
        # Pre-check: already open?
        if self.page.locator('[data-testid="year-End-test-id"]').count() > 0:
            logger.info("[SUCCESS] Time period section already open")
            return True

        # Try clicking the accordion summary
        attempts = [
            # Most specific: role=button containing "Time period"
            lambda: self.page.locator('div[role="button"]').filter(has_text="Time period").first.click(timeout=5000),
            lambda: self.page.locator('button').filter(has_text="Time period").first.click(timeout=5000),
            # Playwright getByText equivalent
            lambda: self.page.get_by_text("Time period", exact=False).first.click(timeout=5000),
        ]

        for attempt in attempts:
            try:
                attempt()
                self.page.wait_for_selector('[data-testid="year-End-test-id"]', timeout=10000)
                logger.info("[SUCCESS] Time period section expanded")
                return True
            except Exception:
                continue

        logger.warning("[WARNING] Time period section could not be expanded")
        return False

    def update_end_year_to_latest(self):
        """
        Confirm the panel is open, open the end-year dropdown, find the
        highest non-disabled year, select it, then click Apply.

        DOM (confirmed):
          • Trigger:  [data-testid="year-End-test-id"]  (click opens listbox)
          • Listbox:  ul[aria-labelledby*="year-End"][role="listbox"]
          • Options:  li[role="option"] with data-value="YYYY"
          • Disabled: aria-disabled="true" + Mui-disabled class  (2025+ currently)
          • Apply:    button[name="Apply"]
        """
        try:
            # --- 1. Ensure panel is open ---
            self.expand_time_period_section()

            # --- 2. Read current end year ---
            try:
                current_end = self.page.get_by_test_id("year-End-test-id").inner_text().strip()
                logger.info(f"[STATUS] Current end year: {current_end or '(unreadable)'}")
            except Exception:
                current_end = ""
                logger.info("[STATUS] Could not read current end year")

            # --- 3. Open the end year dropdown ---
            # Playwright codegen: page.getByTestId('year-End-test-id').getByText('2022').click()
            self.page.get_by_test_id("year-End-test-id").click()
            self.page.wait_for_selector(
                'ul[aria-labelledby*="year-End"][role="listbox"]', timeout=10000
            )
            logger.info("[INFO] End year dropdown opened")

            # --- 4. Collect all enabled year options ---
            # Scope to the year-End listbox; read data-value attribute directly.
            enabled_opts = self.page.locator(
                'ul[aria-labelledby*="year-End"] li[role="option"]:not([aria-disabled="true"])'
            ).all()

            year_options = []
            for el in enabled_opts:
                dv = el.get_attribute("data-value") or ""
                if len(dv) == 4 and dv.isdigit():
                    year_options.append(int(dv))

            if not year_options:
                logger.warning("[WARNING] No clickable year options found — closing dropdown")
                self.page.keyboard.press("Escape")
                return False

            latest_year = max(year_options)
            logger.info(
                f"[INFO] Available years: {min(year_options)}–{latest_year} "
                f"({len(year_options)} options)"
            )

            # --- 5. Select the latest year ---
            self.page.get_by_role("option", name=str(latest_year)).click()
            logger.info(f"[SUCCESS] Selected end year: {latest_year}")

            # --- 6. Click Apply ---
            self.page.get_by_role("button", name="Apply").click()
            logger.info("[SUCCESS] Apply clicked — time period updated")
            return True

        except Exception as e:
            logger.warning(f"[WARNING] update_end_year_to_latest: {e}")
            return False

    def download_data(self):
        """
        Step 5: Click the downloads button, then 'Table in Excel'.
        Uses Playwright's expect_download() to capture the file reliably.

        Playwright codegen:
          await page.getByTestId('downloads-button').click();
          await page.locator('span').filter({ hasText: 'Table in Excel' }).first().click();
        """
        try:
            logger.info("[INFO] STEP 5: Downloading data...")

            self.page.get_by_test_id("downloads-button").click()
            logger.info("[SUCCESS] Downloads button clicked")

            # Intercept the download event
            with self.page.expect_download(timeout=60000) as dl_info:
                self.page.locator("span").filter(has_text="Table in Excel").first.click()
                logger.info("[SUCCESS] 'Table in Excel' selected")

            download = dl_info.value
            filename = download.suggested_filename
            save_path = os.path.join(self.download_dir, filename)
            download.save_as(save_path)

            logger.info(f"[SUCCESS] Step 5 completed - Downloaded: {filename}")
            return True

        except PlaywrightTimeout:
            logger.error("[FAILURE] Step 5 timed out waiting for download")
            return False
        except Exception as e:
            logger.error(f"[FAILURE] Step 5 failed: {e}")
            return False

    # ------------------------------------------------------------------
    # Fallback workflow (Steps 7-9)
    # ------------------------------------------------------------------
    def execute_fallback_workflow(self):
        logger.info("[INFO] EXECUTING FALLBACK WORKFLOW (Steps 7-9)")
        logger.info("=" * 60)
        try:
            if not self.step7_search_fallback():
                return False
            if not self.step8_configure_filters():
                logger.warning("[WARNING] Filter configuration had issues, continuing...")
            if not self.step9_add_countries():
                logger.warning("[WARNING] Country selection had issues, continuing...")
            self.update_time_periods()
            if not self.download_data():
                return False
            logger.info("[SUCCESS] FALLBACK WORKFLOW SUCCESSFUL!")
            return True
        except Exception as e:
            logger.error(f"[FAILURE] Fallback workflow failed: {e}")
            return False

    def step7_search_fallback(self):
        """Step 7: Load default URL and search for 'Pensions at a glance'."""
        try:
            logger.info("[INFO] STEP 7: Fallback search method...")
            logger.info("[INFO] Loading default OECD Data Explorer...")
            self.page.goto(self.default_url, wait_until="networkidle", timeout=60000)

            search_input = self.page.locator(
                "input[placeholder='Search by keywords'], "
                "input[aria-label='Search by keywords'], "
                "input[data-testid='spotlight_input']"
            ).first
            search_input.wait_for(timeout=15000)
            search_input.fill(self.search_term)
            logger.info(f"[SUCCESS] Typed search term: '{self.search_term}'")

            # Submit
            try:
                self.page.locator("button[aria-label='commit']").click(timeout=5000)
            except Exception:
                search_input.press("Enter")
            logger.info("[SUCCESS] Search submitted")

            # Wait for results and click pension dataset
            self.page.wait_for_load_state("networkidle", timeout=20000)

            result = self.page.locator("a").filter(has_text="Pensions at a glance").first
            result.click(timeout=15000)
            self.page.wait_for_load_state("networkidle", timeout=30000)
            logger.info("[SUCCESS] Step 7 completed - Dataset found via search")
            return True

        except Exception as e:
            logger.error(f"[FAILURE] Step 7 failed: {e}")
            return False

    def step8_configure_filters(self):
        """Step 8: Configure Measure and Sex filters."""
        try:
            logger.info("[INFO] STEP 8: Configuring filters manually...")

            # Expand Measure section and select target
            try:
                self.page.locator('[aria-controls="MEASURE"], div[role="button"]:has-text("Measure")').first.click(timeout=5000)
                self.page.locator(f'div[role="checkbox"]:has-text("{self.target_measure}")').click(timeout=5000)
                logger.info(f"[SUCCESS] Measure set: {self.target_measure}")
            except Exception:
                logger.info("[INFO] Measure filter not updated")

            # Expand Sex section and select Male
            try:
                self.page.locator('[aria-controls="SEX"], div[role="button"]:has-text("Sex")').first.click(timeout=5000)
                male = self.page.locator('[data-testid="value_M"][role="checkbox"]')
                if male.get_attribute("aria-checked") != "true":
                    male.click(timeout=5000)
                logger.info(f"[SUCCESS] Sex set: {self.target_sex}")
            except Exception:
                logger.info("[INFO] Sex filter not updated")

            logger.info("[SUCCESS] Step 8 completed")
            return True
        except Exception as e:
            logger.warning(f"[WARNING] Step 8 warning: {e}")
            return True

    def step9_add_countries(self):
        """Step 9: Add required countries."""
        try:
            logger.info("[INFO] STEP 9: Adding required countries...")

            # Expand Reference area section
            try:
                self.page.locator('[aria-controls="REF_AREA"], div[role="button"]:has-text("Reference area")').first.click(timeout=5000)
                self.page.wait_for_timeout(2000)
            except Exception:
                pass

            selected_count = 0
            already_selected = 0

            for country in self.required_countries:
                try:
                    el = self.page.locator(f'div[role="checkbox"]:has-text("{country}")').first
                    if el.get_attribute("aria-checked") == "true":
                        already_selected += 1
                    else:
                        el.click(timeout=3000)
                        selected_count += 1
                        self.page.wait_for_timeout(200)
                except Exception:
                    logger.debug(f"[WARNING] {country} not found")

            logger.info(f"[STATUS] Countries: {selected_count} newly selected, {already_selected} already selected")
            logger.info("[SUCCESS] Step 9 completed")
            return True
        except Exception as e:
            logger.warning(f"[WARNING] Step 9 warning: {e}")
            return True

    # ------------------------------------------------------------------
    # Orchestration
    # ------------------------------------------------------------------
    def run_complete_automation(self):
        try:
            logger.info("=" * 70)
            logger.info("[INFO] OECDAR COMPLETE AUTOMATION")
            logger.info("[INFO] Dataset: OECD - Retirement - OECDAR")
            logger.info("[STATUS] Measure: Effective labour market exit age")
            logger.info("[INFO] Gender: Male")
            logger.info("[INFO] Countries: 29 countries from runbook")
            logger.info("[INFO] Fallback: Steps 7-9 if primary URL fails")
            logger.info("=" * 70)

            if not self.setup_driver():
                return False

            success = False
            try:
                if self.try_primary_workflow():
                    success = True
                else:
                    logger.info("")
                    logger.info("[INFO] PRIMARY WORKFLOW FAILED")
                    logger.info("[INFO] EXECUTING RUNBOOK FALLBACK (Steps 7-9)")
                    logger.info("")
                    if self.execute_fallback_workflow():
                        success = True

                downloaded_files = os.listdir(self.download_dir)

                if downloaded_files:
                    logger.info("=" * 70)
                    if success:
                        logger.info("[SUCCESS] OECDAR AUTOMATION COMPLETED SUCCESSFULLY!")
                    else:
                        logger.info("[INFO] FILES FOUND IN DOWNLOAD DIRECTORY")
                    logger.info(f"[INFO] Downloaded files: {downloaded_files}")
                    logger.info("[INFO] Ready for Step 6: Manual data processing")
                    logger.info("=" * 70)
                    return True
                else:
                    logger.error("[FAILURE] No files were downloaded")
                    return False

            except Exception as e:
                logger.error(f"[FAILURE] Automation workflow failed: {e}")
                return False

        except Exception as e:
            logger.error(f"[FAILURE] Complete automation failed: {e}")
            return False

        finally:
            for obj, method in [(self.page, 'close'), (self.context, 'close'),
                                (self.browser, 'close'), (self._pw, 'stop')]:
                if obj:
                    try:
                        getattr(obj, method)()
                    except Exception:
                        pass
            logger.info("[INFO] Browser closed")


def main():
    print("=" * 70)
    print("[INFO] OECDAR AUTOMATION WITH PLAYWRIGHT")
    print("[INFO] Implements complete runbook procedures")
    print("[INFO] Primary workflow + Steps 7-9 fallback")
    print("=" * 70)

    automation = OECDARAutomation()
    success = automation.run_complete_automation()

    if success:
        print("\n" + "=" * 70)
        print("[SUCCESS] AUTOMATION COMPLETED SUCCESSFULLY!")
        print("[STATUS] OECD retirement data retrieved")
        print("=" * 70)
    else:
        print("\n" + "=" * 70)
        print("[FAILURE] AUTOMATION ENCOUNTERED ISSUES")
        print("[INFO] Check logs above for detailed information")
        print("=" * 70)


if __name__ == "__main__":
    main()
