# ICE Margin Calculator - Automated Tool

Automated margin calculation tool for ICE (Intercontinental Exchange) using Playwright and Excel integration.

## Features

- Simple GUI interface with button to trigger calculations
- Read positions directly from Excel file
- Automated browser automation using Playwright
- Write calculated margin results back to Excel
- Session persistence (login once, use multiple times)

---

## Setup Instructions

### 1. Install Dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. First-Time Login (One-Time Setup)

Run the login script to save your ICE session:

```bash
python login_once.py
```

**Steps:**
1. Browser will open automatically
2. Enter your ICE email and password
3. Complete 2FA authentication
4. Wait until you're redirected to the ICA dashboard
5. Press ENTER in the terminal
6. Session is saved to `ice_session.json`

This only needs to be done once. Future runs will reuse the saved session.

---

## Usage

### Method 1: Create Excel Template (First Time)

Generate the Excel template with proper columns:

```bash
python create_template.py
```

This creates `positions_template.xlsx` with:
- Position columns (Account, Symbol, Quantity, Price, Side, Product Type)
- Calculated Margin column (auto-filled)
- Instructions sheet

### Method 2: Edit Your Positions

Open `positions_template.xlsx` in Excel:
1. Fill in your position data
2. Save the file
3. Keep Excel open or close it (both work)

### Method 3: Run the GUI Application

Start the margin calculator GUI:

```bash
python gui_app.py
```

**GUI Workflow:**
1. Default file is `positions_template.xlsx` (or browse to select another)
2. Click **"Calculate Margin"** button
3. Browser opens automatically and runs calculation
4. Margin result is copied and written back to Excel
5. Success message appears
6. Check your Excel file for the updated margin

---

## File Structure

```
margin calculation/
├── login_once.py              # One-time login script
├── margin_calculator.py       # Core calculation logic
├── gui_app.py                 # GUI application (main entry point)
├── create_template.py         # Excel template generator
├── requirements.txt           # Python dependencies
├── ice_session.json           # Saved session (created after login)
├── positions_template.xlsx    # Your working Excel file
└── README.md                  # This file
```

---

## Excel File Format

### Required Columns (Adjust to match ICE upload format):

| Account | Symbol | Quantity | Price | Side | Product Type | Calculated Margin |
|---------|--------|----------|-------|------|--------------|-------------------|
| ACC001  | ES     | 10       | 4500  | BUY  | Future       | *Auto-filled*     |
| ACC001  | NQ     | 5        | 15000 | SELL | Future       | *Auto-filled*     |

**Note:** The exact columns must match the ICE margin calculator upload format. Adjust the template if needed.

---

## How It Works

1. **Read Excel:** Script reads position data from Excel
2. **Upload to ICE:** Playwright opens browser and uploads Excel to ICE ICA
3. **Calculate:** Runs margin analytics on ICE platform
4. **Copy Result:** Copies the margin value from the results grid
5. **Write Back:** Updates the "Calculated Margin" column in Excel
6. **Done:** Excel file is updated with the result

---

## Configuration

Edit these variables in `margin_calculator.py` if needed:

```python
SESSION_FILE = "ice_session.json"          # Session file location
APP_URL = "https://ica.ice.com/ICA/Main"   # ICE ICA URL
EXCEL_FILE = "positions_template.xlsx"     # Default Excel file
RESULT_CELL_ID = "#cell-1280"              # Cell ID where margin appears
```

---

## Troubleshooting

### Session Expired
If you get authentication errors:
```bash
python login_once.py
```
Re-run the login script to refresh your session.

### Excel File Not Found
- Make sure `positions_template.xlsx` exists in the same folder
- Or use the "Browse" button in GUI to select your file

### Calculation Fails
- Check your Excel file format matches ICE upload requirements
- Verify you have valid position data
- Check browser automation didn't encounter unexpected elements

### Result Not Writing to Excel
- Make sure Excel file is not locked by another process
- Verify "Calculated Margin" column exists (row 1)

---

## Notes

- The browser runs in **non-headless mode** (visible) so you can see what's happening
- Session typically lasts 24-48 hours before re-login is needed
- You can edit Excel while the GUI is open
- Multiple positions can be in the same Excel file
- The script clears existing portfolios before uploading new data

---

## Old Files (Can be deleted)

- `run_margin.py` - Old script that processed multiple files (no longer needed)

---

## Future Enhancements

Possible improvements:
- Add batch processing for multiple Excel files
- Export detailed reports
- Schedule automatic calculations
- Email notifications
- Support for different ICE calculators

---

## Support

For issues:
1. Check the status log in the GUI
2. Verify your Excel format matches ICE requirements
3. Re-run `login_once.py` if session expired
4. Check that Playwright and browser drivers are installed correctly

---

## License

This is a personal automation tool. Use at your own risk.
