"""
Main margin calculator module with Excel read/write functionality.
This replaces the old run_margin.py with single-file processing.
"""

import time
import re
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
import pyperclip
from openpyxl import load_workbook

# ---------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------
SESSION_FILE = "ice_session.json"
APP_URL = "https://ica.ice.com/ICA/Main"
EXCEL_FILE = "positions_template.xlsx"  # Your working Excel file
RESULT_CELL_ID = "#cell-1468"  # The cell ID where margin result appears
# ---------------------------------------------------------------------


def read_excel_file(excel_path):
    """Read the Excel file and return data info."""
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active

    # Count data rows (excluding header)
    data_rows = sum(
        1
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row)
        if any(cell.value is not None for cell in row)
    )

    print(f"📊 Loaded Excel: {Path(excel_path).name}")
    print(f"   Found {data_rows} position(s)")

    wb.close()
    return data_rows


def write_margin_to_excel(excel_path, margin_result):
    """Write the calculated margin back to Excel."""
    try:
        wb = load_workbook(excel_path)
        ws = wb.active

        # Find the "Calculated Margin" column
        margin_col = None
        for col in range(1, ws.max_column + 1):
            if ws.cell(1, col).value and "Margin" in str(ws.cell(1, col).value):
                margin_col = col
                break

        if not margin_col:
            print("⚠️ 'Calculated Margin' column not found. Adding to column G.")
            margin_col = 7
            ws.cell(1, margin_col, "Calculated Margin")

        # Write margin result to row 2 (first data row)
        ws.cell(2, margin_col, margin_result)

        wb.save(excel_path)
        wb.close()

        print(f"✅ Margin result written to Excel: {margin_result}")
        return True

    except Exception as e:
        print(f"❌ Error writing to Excel: {e}")
        return False


def run_margin_calc(excel_path):
    """
    Main function to run ICE margin calculator.
    1. Uploads the Excel file to ICE
    2. Runs the calculation
    3. Copies the result
    4. Writes result back to Excel
    """
    excel_path = Path(excel_path).resolve()

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    # Verify session file exists
    if not Path(SESSION_FILE).exists():
        raise FileNotFoundError(
            f"Session file '{SESSION_FILE}' not found. Please run 'login_once.py' first."
        )

    print(f"\n{'='*60}")
    print(f"Starting Margin Calculation")
    print(f"{'='*60}")

    # Read Excel to show info
    read_excel_file(excel_path)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=150)
        context = browser.new_context(storage_state=SESSION_FILE)
        page = context.new_page()

        try:
            print(f"\n🌐 Opening ICE ICA application...")
            page.goto(APP_URL, timeout=60000)
            page.wait_for_load_state("networkidle")

            # Check and clear existing portfolios
            checkbox_locator = (
                page.get_by_role(
                    "gridcell",
                    name="Press Space to toggle row selection (unchecked) All Portfolios (1)",
                )
                .get_by_label("Press Space to toggle row")
                .first
            )

            if checkbox_locator.count() > 0:
                print("🗑️  Clearing existing portfolios...")
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
                time.sleep(2)
            else:
                print("✓ No existing portfolios to clear")

            # Navigate to Tools → Upload Trades
            print("\n📤 Uploading positions file...")
            page.get_by_role("menuitem", name="Tools").click()
            page.get_by_role("menuitem", name="Upload Trades").click()

            # Upload the Excel file
            page.get_by_role(
                "button", name=re.compile("Select file", re.I)
            ).set_input_files(str(excel_path))
            page.get_by_role("button", name="Upload").click()

            # Wait for upload confirmation
            page.wait_for_selector("button:has-text('OK')", timeout=60000)
            page.get_by_role("button", name="OK").click()
            print("✅ Upload completed")

            # Select all accounts and run calculation
            print("\n🧮 Running margin calculation...")
            page.locator(
                "input[aria-label*='Press Space to toggle row selection']"
            ).first.check()
            page.get_by_role("button", name="Run Analytics").click()
            page.get_by_role("tabpanel").filter(has_text="Run").get_by_role(
                "button"
            ).nth(1).click()

            # Wait for calculation to complete (fixed time)
            print("⏳ Waiting for calculation to complete (5 seconds)...")
            time.sleep(5)  # Results appear within 5 seconds
            print("✅ Calculation completed")

            # Get the margin result directly from the cell
            print("\n📋 Extracting margin result...")

            # Method 1: Try to get text directly from the cell
            try:
                result_cell = page.locator(RESULT_CELL_ID)
                result_cell.wait_for(timeout=30000, state="visible")  # 30 second timeout
                copied_text = result_cell.inner_text().strip()
                print(f"✅ Margin extracted: {copied_text}")
            except Exception as e:
                print(f"⚠️ Direct extraction failed: {e}")
                # Method 2: Fallback to clipboard method
                print("Trying clipboard method...")
                page.locator(RESULT_CELL_ID).click(button="right", timeout=60000)
                time.sleep(0.5)
                page.get_by_text("Copy").first.click(timeout=5000)
                time.sleep(1)
                copied_text = pyperclip.paste().strip()
                print(f"✅ Margin copied via clipboard: {copied_text}")

            # Result will be shown in the modal
            print(f"\n{'='*60}")
            print(f"✅ SUCCESS! Margin: {copied_text}")
            print(f"{'='*60}\n")

            return copied_text

        except Exception as e:
            print(f"\n❌ Error during calculation: {e}")
            raise


if __name__ == "__main__":
    # Test run
    try:
        result = run_margin_calc(EXCEL_FILE)
        print(f"Final result: {result}")
    except Exception as e:
        print(f"Failed: {e}")
