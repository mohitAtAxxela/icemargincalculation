"""
GUI Application for ICE Margin Calculator
Simple interface with file selection and calculate button.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
from margin_calculator import run_margin_calc

class MarginCalculatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ICE Margin Calculator")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        # Default Excel file
        self.excel_path = Path("positions_template.xlsx").resolve()
        self.is_calculating = False

        # Close browser when window is closed
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.setup_ui()

    def on_closing(self):
        """Handle window close event."""
        self.root.destroy()

    def setup_ui(self):
        """Setup the user interface."""
        # Title
        title_frame = tk.Frame(self.root, bg="#366092", height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        title_label = tk.Label(
            title_frame,
            text="ICE Margin Calculator",
            font=("Arial", 18, "bold"),
            bg="#366092",
            fg="white"
        )
        title_label.pack(pady=15)

        # Main content frame
        content_frame = tk.Frame(self.root, padx=20, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # File selection section
        file_frame = tk.LabelFrame(
            content_frame,
            text="Excel File Selection",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=10
        )
        file_frame.pack(fill=tk.X, pady=(0, 20))

        # File path display
        self.file_label = tk.Label(
            file_frame,
            text=str(self.excel_path),
            font=("Arial", 9),
            anchor="w",
            bg="#f0f0f0",
            relief=tk.SUNKEN,
            padx=5,
            pady=5
        )
        self.file_label.pack(fill=tk.X, pady=(0, 10))

        # Browse button
        browse_btn = tk.Button(
            file_frame,
            text="📁 Browse Excel File",
            command=self.browse_file,
            font=("Arial", 10),
            bg="#4CAF50",
            fg="white",
            cursor="hand2",
            relief=tk.RAISED,
            padx=10,
            pady=5
        )
        browse_btn.pack()

        # Calculate button section
        calc_frame = tk.Frame(content_frame)
        calc_frame.pack(fill=tk.X, pady=(0, 20))

        self.calc_button = tk.Button(
            calc_frame,
            text="🧮 Calculate Margin",
            command=self.calculate_margin,
            font=("Arial", 14, "bold"),
            bg="#2196F3",
            fg="white",
            cursor="hand2",
            relief=tk.RAISED,
            padx=20,
            pady=15,
            width=25
        )
        self.calc_button.pack()

        # Status section
        status_frame = tk.LabelFrame(
            content_frame,
            text="Status",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=10
        )
        status_frame.pack(fill=tk.BOTH, expand=True)

        # Status text area
        self.status_text = tk.Text(
            status_frame,
            height=8,
            font=("Consolas", 9),
            bg="#f9f9f9",
            relief=tk.SUNKEN,
            state=tk.DISABLED
        )
        self.status_text.pack(fill=tk.BOTH, expand=True)

        # Progress bar
        self.progress = ttk.Progressbar(
            content_frame,
            mode='indeterminate',
            length=400
        )

        # Initial status
        self.update_status("Ready. Select Excel file and click 'Calculate Margin'.")

    def browse_file(self):
        """Open file dialog to select Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialdir=Path.cwd()
        )

        if filename:
            self.excel_path = Path(filename)
            self.file_label.config(text=str(self.excel_path))
            self.update_status(f"Selected: {self.excel_path.name}")

    def update_status(self, message):
        """Update the status text area."""
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)

    def calculate_margin(self):
        """Start margin calculation in a separate thread."""
        if self.is_calculating:
            messagebox.showwarning("Already Running", "Calculation is already in progress!")
            return

        if not self.excel_path.exists():
            messagebox.showerror(
                "File Not Found",
                f"Excel file not found:\n{self.excel_path}\n\nPlease select a valid file."
            )
            return

        # # Confirm before starting
        # response = messagebox.askyesno(
        #     "Confirm Calculation",
        #     # f"Calculate margin for:\n{self.excel_path.name}\n\nThis will open a browser window."
        # )

        # if not response:
        #     return

        # Clear status
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state=tk.DISABLED)

        # Disable button and show progress
        self.is_calculating = True
        self.calc_button.config(state=tk.DISABLED, bg="#cccccc")
        self.progress.pack(pady=10)
        self.progress.start(10)

        # Run calculation in separate thread
        calc_thread = threading.Thread(target=self.run_calculation, daemon=True)
        calc_thread.start()

    def run_calculation(self):
        """Execute the margin calculation (runs in separate thread)."""
        try:
            self.update_status("="*50)
            self.update_status("Starting margin calculation...")
            self.update_status("="*50)

            # Run the calculation
            result = run_margin_calc(str(self.excel_path))

            # Success
            self.update_status("")
            self.update_status("="*50)
            self.update_status(f"✅ SUCCESS!")
            self.update_status(f"Calculated Margin: {result}")
            self.update_status(f"Result saved to: {self.excel_path.name}")
            self.update_status("="*50)

            # Show success dialog
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Success",
                    f"Margin calculated successfully!\n\nResult: {result}\n\nCheck your Excel file."
                )
            )

        except FileNotFoundError as e:
            error_msg = str(e)
            self.update_status(f"\n❌ ERROR: {error_msg}")
            self.root.after(
                0,
                lambda msg=error_msg: messagebox.showerror("File Error", msg)
            )

        except Exception as e:
            error_msg = str(e)
            self.update_status(f"\n❌ ERROR: {error_msg}")
            self.root.after(
                0,
                lambda msg=error_msg: messagebox.showerror(
                    "Calculation Error",
                    f"An error occurred:\n\n{msg}\n\nCheck the status log for details."
                )
            )

        finally:
            # Re-enable button
            self.root.after(0, self.finish_calculation)

    def finish_calculation(self):
        """Reset UI after calculation completes."""
        self.is_calculating = False
        self.calc_button.config(state=tk.NORMAL, bg="#2196F3")
        self.progress.stop()
        self.progress.pack_forget()


def main():
    """Main entry point for the GUI application."""
    root = tk.Tk()
    app = MarginCalculatorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
