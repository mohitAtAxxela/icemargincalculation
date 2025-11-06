# Live Excel Integration Guide

## What's New? No More Saving Required!

The margin calculator now reads data directly from your **open Excel file**, including **unsaved changes**!

## How It Works

### **Method 1: Excel File is OPEN (Recommended)**
- Keep your Excel file open
- Make changes to positions
- **NO NEED TO SAVE**
- Click "Calculate Margin" button
- Script reads your unsaved changes directly from Excel
- Result is written back to the open Excel file
- You'll see the result appear instantly in Excel

### **Method 2: Excel File is CLOSED (Fallback)**
- If the file is closed, it reads from the saved file
- Works like before (requires file to be saved)

## Installation

First, install the new dependency:

```bash
pip install pywin32
```

Or install all requirements:

```bash
pip install -r requirements.txt
```

## Usage

1. **Open Excel file** (positions_template.xlsx)
2. **Edit your positions** (no need to save!)
3. **Run the GUI**:
   ```bash
   python gui_app.py
   ```
4. **Click "Calculate Margin"**
5. **See the result appear in Excel** (in the "Calculated Margin" column)

## Benefits

✅ **No more clicking Save** - Edit and calculate immediately
✅ **Faster workflow** - Reduce manual steps
✅ **Live updates** - See results in real-time
✅ **Works both ways** - Open or closed Excel files

## Technical Details

The script uses **Windows COM API** (via `pywin32`) to:
- Detect if Excel is running
- Find your open workbook
- Read data directly from Excel's memory
- Write results back to the open Excel

This means:
- Reads **unsaved changes** from open Excel
- Writes results **directly to open Excel** (without closing it)
- Falls back to file reading if Excel is closed

## Important Notes

⚠️ **Windows Only**: COM API only works on Windows
⚠️ **Excel Must Be Running**: For live reading, Excel app must be open
⚠️ **Manual Save**: After calculation, you still need to manually save your Excel file to keep the changes

## File Structure

```
margin calculation/
├── excel_live_reader.py      # NEW: Live Excel COM API reader
├── margin_calculator.py       # Updated to use live reading
├── gui_app.py                 # GUI (no changes needed)
└── positions_template.xlsx    # Your working Excel file
```

## Testing

To test the live reader independently:

```bash
python excel_live_reader.py
```

This will show whether it can detect and read from your open Excel file.

## Troubleshooting

### "Reading from SAVED Excel file" instead of "OPEN Excel"
- Make sure Excel is actually open with the file loaded
- Check that the file path matches exactly
- Close and reopen Excel if needed

### COM API Errors
- Make sure `pywin32` is installed: `pip install pywin32`
- Try running: `python -c "import win32com.client; print('OK')"`

### File Not Found
- Make sure the Excel file exists
- Check the file path in the GUI

---

## Comparison: Before vs After

### Before:
1. Edit Excel
2. **Click Save** ⬅️ Extra step!
3. Run calculation
4. Open Excel to see result

### After:
1. Edit Excel (keep it open)
2. Run calculation
3. Result appears instantly in Excel ✨

---

Enjoy your streamlined workflow! 🚀
