# python -m playwright codegen https://ica.ice.com/ICA/Main
# command to track the actions to be done over the site
import time, csv, re
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
import pyperclip  # pip install pyperclip

# ---------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------
EXCEL_FOLDER = "D:\downloads\scenarios"  # folder containing all your Excel files
OUTPUT_CSV = "margin_results.csv"  # results saved here
SESSION_FILE = "ice_session.json"  # your saved session from login_once.py
APP_URL = "https://ica.ice.com/ICA/Main"
# ---------------------------------------------------------------------
from pathlib import Path


def get_latest_excel(folder="scenarios"):
    folder = Path(folder)
    return max(folder.glob("*.xlsx"), key=lambda f: f.stat().st_mtime)


# Then in your upload step:


def run_margin_calc(file_path: str):
    """Runs the ICE margin calculator for one Excel and returns copied result text."""
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False, slow_mo=150
        )  # headless=True for silent
        context = browser.new_context(storage_state=SESSION_FILE)
        page = context.new_page()

        print(f"\nOpening ICA main page for {Path(file_path).name} ...")
        page.goto(APP_URL)
        page.wait_for_load_state("networkidle")
        checkbox_locator = (
            page.get_by_role(
                "gridcell",
                name="Press Space to toggle row selection (unchecked) All Portfolios (1)",
            )
            .get_by_label("Press Space to toggle row")
            .first
        )

        # Check if it exists
        if checkbox_locator.count() > 0:
            print("✅ Found 'All Portfolios' checkbox — selecting it...")
            checkbox_locator.check()
            page.get_by_role("button", name="Actions").first.click()
            page.get_by_role("button", name="Delete").click()
            page.get_by_text("Delete", exact=True).click()
            page.get_by_role(
                "columnheader",
                name="Press Space to toggle all rows selection (unchecked) Calculation ID",
            ).get_by_label("Press Space to toggle all").first.check()
            page.get_by_role("button", name="Actions").nth(1).click()
            page.get_by_role("button", name="Delete").nth(1).click()
            page.get_by_role("button", name="OK").click()
        else:
            print("⚠️ 'All Portfolios' checkbox not found — skipping selection.")
        # 1️⃣ Navigate to Tools → Upload Trades
        page.get_by_role("menuitem", name="Tools").click()
        page.get_by_role("menuitem", name="Upload Trades").click()

        # 2️⃣ Upload the Excel file
        print("Uploading:", file_path)
        page.get_by_role(
            "button", name=re.compile("Select file", re.I)
        ).set_input_files(file_path)
        page.get_by_role("button", name="Upload").click()

        # Portfolio options
        # page.get_by_text("Add these trades to the").click()
        # page.get_by_text("Clear the portfolio before").click()

        # Confirm and wait for upload success overlay
        page.wait_for_selector("button:has-text('OK')", timeout=60000)
        page.get_by_role("button", name="OK").click()
        # page.wait_for_selector(".ice-overlay.progress-dialog button:has-text('OK')", timeout=120000)
        # page.locator(".ice-overlay.progress-dialog button:has-text('OK')").click()
        # page.wait_for_selector(".ice-overlay.progress-dialog", state="detached", timeout=120000)
        print("✅ Upload completed and overlay dismissed.")

        # 4️⃣ Select all accounts and run calculation
        page.locator(
            "input[aria-label*='Press Space to toggle row selection']"
        ).first.check()
        page.get_by_role("button", name="Run Analytics").click()
        page.get_by_role("tabpanel").filter(has_text="Run").get_by_role("button").nth(
            1
        ).click()
        time.sleep(5)
        print("✅ Processing overlay gone — continuing.")
        # page.get_by_role("button", name="Actions").first.click()
        # page.get_by_role("button", name="Export to Excel").first.click()
        # with page.expect_download() as download_info:
        #     page.get_by_role("button", name="Export", exact=True).click()
        # download = download_info.value
        # export_dir = Path("D:/downloads/exports")
        # export_dir.mkdir(parents=True, exist_ok=True)

        # # Get the suggested name and force '.xlsx' extension
        # suggested_name = Path(
        #     download.suggested_filename
        # ).stem  # remove old extension if any
        # excel_name = f"{suggested_name}.xlsx"

        # save_path = export_dir / excel_name
        # download.save_as(str(save_path))

        # print(f"✅ Excel file downloaded to: {save_path}")
        # # 5️⃣ Wait for results grid
        # print("Waiting for results...")
        # page.wait_for_selector("#cell-1280", timeout=90_000)

        # 6️⃣ Copy the result
        # print("Copying margin result...")
        # page.locator("#cell-1280").click(button="right")
        # page.get_by_text("Copy").first.click()
        # time.sleep(1)

        # copied_text = pyperclip.paste().strip()
        # print(f"✅ {Path(file_path).name} → {copied_text}")

        context.close()
        browser.close()
        return 


def run_all_files():
    folder = Path(EXCEL_FOLDER)
    files = sorted(folder.glob("*.xlsx"))
    results = []

    for f in files:
        print(f"\n=== Processing {f.name} ===")
        try:
            run_margin_calc(str(f))
        except Exception as e:
            result = f"Error: {e}"
            print("❌", result)
        results.append({"file": f.name, "result": result})

    # Save results to CSV
    with open(OUTPUT_CSV, "w", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=["file", "result"])
        writer.writeheader()
        writer.writerows(results)

    print(f"\n✅ All results saved to {OUTPUT_CSV}")


if __name__ == "__main__":
    run_all_files()
