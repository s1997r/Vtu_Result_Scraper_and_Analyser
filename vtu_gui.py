'''
import threading
import queue
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Toplevel, Scrollbar, Text
from captcha_handler import CaptchaHandler
from vtu_marks_scraper import generate_usn_list, fetch_vtu_result_with_retry, get_driver
from student_data import parse_student_result
import pandas as pd

# Control flag for stopping threads
stop_flag = threading.Event()
DEFAULT_URL = "https://results.vtu.ac.in/DJcbcs25/index.php"


def get_missing_usns(expected_usns, output_file):
    """Get list of USNs not present in output file"""
    if not os.path.exists(output_file):
        return expected_usns
    try:
        df = pd.read_excel(output_file)
        existing_usns = df['University Seat Number'].astype(str).str.upper().unique()
        return [usn for usn in expected_usns if usn.upper() not in existing_usns]
    except Exception:
        return expected_usns


def run_scraper(usn_list, output_path, log_queue, progress_queue, append=False, base_url=DEFAULT_URL, headless=True):
    """Main scraping function to run in thread"""
    try:
        total = len(usn_list)
        results = []
        handler = CaptchaHandler()
        count = 0

        driver = get_driver(headless=headless)
        missing_usns = []

        for usn in usn_list:
            if stop_flag.is_set():
                log_queue.put("Process manually stopped by user.\n")
                break

            count += 1
            progress = int((count / total) * 100)
            progress_queue.put(progress)
            log_queue.put(f"[{count}/{total}] Fetching: {usn}\n")

            html = fetch_vtu_result_with_retry(driver, usn, handler, base_url=base_url)
            if html:
                tmp_file = f"_tmp_{usn}.html"
                with open(tmp_file, "w", encoding="utf-8") as f:
                    f.write(html)
                df = parse_student_result(tmp_file)
                results.append(df)
                os.remove(tmp_file)
            else:
                missing_usns.append(usn)
                log_queue.put(f"  -> Failed to fetch: {usn}\n")

        driver.quit()

        if results:
            df_all = pd.concat(results, ignore_index=True)
            df_all.fillna("NA", inplace=True)

            if append and os.path.exists(output_path):
                old_df = pd.read_excel(output_path)
                df_all = pd.concat([old_df, df_all], ignore_index=True).drop_duplicates()

            df_all.to_excel(output_path, index=False)
            log_queue.put(f"Results saved to: {output_path}\n")

        if missing_usns:
            log_queue.put("Missing USNs:\n" + ", ".join(missing_usns) + "\n")

        if not results:
            log_queue.put("No results to save.\n")

    except Exception as e:
        log_queue.put(f"Error: {str(e)}\n")
    finally:
        progress_queue.put(100)
        log_queue.put("Finished.\n")
        stop_flag.clear()

import sys

def resource_path(relative_path):
    """Get absolute path to resource, works for PyInstaller and normal runs"""
    try:
        base_path = sys._MEIPASS  # Set by PyInstaller at runtime
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class VTUGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VTU Marks Scraper GUI")
        self.geometry("500x650")  # Increased height to accommodate URL field

        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception as e:
            print(f"Icon load failed: {e}")


        # Configure styles
        self.style = ttk.Style(self)
        self.style.theme_use('clam')

        # Configure button styles
        self.style.configure('TButton',
                             font=('Helvetica', 10),
                             padding=6,
                             relief=tk.RAISED,
                             borderwidth=2)

        self.style.configure('Primary.TButton',
                             foreground='white',
                             background='#007bff',
                             font=('Helvetica', 10, 'bold'))

        self.style.configure('Secondary.TButton',
                             foreground='white',
                             background='#6c757d')

        self.style.configure('Stop.TButton',
                             foreground='white',
                             background='#dc3545')

        self.style.map('Primary.TButton',
                       background=[('active', '#0056b3'), ('disabled', '#cccccc')])

        self.style.map('Secondary.TButton',
                       background=[('active', '#5a6268'), ('disabled', '#cccccc')])

        self.style.map('Stop.TButton',
                       background=[('active', '#c82333'), ('disabled', '#cccccc')])

        # Main container frame
        main_frame = ttk.Frame(self, padding=(15, 10))
        main_frame.pack(fill=tk.BOTH, expand=True)

        # URL frame
        url_frame = ttk.LabelFrame(main_frame, text="Results URL", padding=(10, 5))
        url_frame.pack(fill=tk.X, pady=(0, 10))

        self.url_var = tk.StringVar(value=DEFAULT_URL)
        ttk.Entry(url_frame, textvariable=self.url_var, width=50).pack(fill=tk.X, padx=5, pady=5)

        # Input frame
        input_frame = ttk.LabelFrame(main_frame, text="USN Range Configuration", padding=(10, 5))
        input_frame.pack(fill=tk.X, pady=(0, 10))

        # USN Base row
        base_row = ttk.Frame(input_frame)
        base_row.pack(fill=tk.X, pady=5)
        ttk.Label(base_row, text="USN Base:").pack(side=tk.LEFT, padx=(0, 5))
        self.base_var = tk.StringVar(value="1CR24BA")
        ttk.Entry(base_row, textvariable=self.base_var, width=20).pack(side=tk.LEFT, padx=5)

        # Start/End row
        range_row = ttk.Frame(input_frame)
        range_row.pack(fill=tk.X, pady=5)
        ttk.Label(range_row, text="Start:").pack(side=tk.LEFT, padx=(0, 5))
        self.start_var = tk.StringVar(value="1")
        ttk.Entry(range_row, textvariable=self.start_var, width=5).pack(side=tk.LEFT, padx=5)

        ttk.Label(range_row, text="End:").pack(side=tk.LEFT, padx=(5, 5))
        self.end_var = tk.StringVar(value="10")
        ttk.Entry(range_row, textvariable=self.end_var, width=5).pack(side=tk.LEFT, padx=5)

        # Output file row
        output_row = ttk.Frame(input_frame)
        output_row.pack(fill=tk.X, pady=5)
        ttk.Label(output_row, text="Output File:").pack(side=tk.LEFT, padx=(0, 5))
        self.output_var = tk.StringVar(value="results.xlsx")
        ttk.Entry(output_row, textvariable=self.output_var, width=30).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(output_row, text="Browse", command=self.browse_file, style='Secondary.TButton').pack(side=tk.LEFT,
                                                                                                        padx=5)

        # Options frame
        options_frame = ttk.Frame(input_frame)
        options_frame.pack(fill=tk.X, pady=(5, 0))

        # Append to existing file checkbox
        self.append_var = tk.BooleanVar()
        self.append_check = ttk.Checkbutton(options_frame,
                                            text="Append to existing file",
                                            variable=self.append_var)
        self.append_check.pack(side=tk.LEFT, padx=(0, 10))

        # Show browser checkbox
        self.headless_var = tk.BooleanVar(value=True)  # Default to headless (browser hidden)
        self.headless_check = ttk.Checkbutton(options_frame,
                                              text="Show browser window",
                                              variable=self.headless_var)
        self.headless_check.pack(side=tk.LEFT)

        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 5))

        self.start_btn = ttk.Button(button_frame,
                                    text="Start New Scrape",
                                    command=self.start_scraping,
                                    style='Primary.TButton')
        self.start_btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.retry_btn = ttk.Button(button_frame,
                                    text="Check & Retry Missing",
                                    command=self.retry_missing_usns,
                                    style='Secondary.TButton')
        self.retry_btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.stop_btn = ttk.Button(button_frame,
                                   text="Stop",
                                   command=self.stop_scraping,
                                   style='Stop.TButton')
        self.stop_btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=700, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(10, 5))

        # Log output
        log_frame = ttk.LabelFrame(main_frame, text="Log Output", padding=(5, 5))
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        self.log_text = Text(log_frame,
                             height=25,
                             bg="#1e1e1e",
                             fg="#ffffff",
                             font=('Consolas', 10),
                             wrap=tk.WORD,
                             padx=5,
                             pady=5)

        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.process_queue_id = None
        self.after(100, self._process_queue)

    def _process_queue(self):
        """Internal method to process queue messages with error handling"""
        try:
            self.process_queue()
        except tk.TclError as e:
            if "invalid command name" not in str(e):
                raise
        finally:
            if not self._is_destroyed():
                self.process_queue_id = self.after(100, self._process_queue)

    def _is_destroyed(self):
        """Check if window is being destroyed"""
        try:
            return not self.winfo_exists()
        except tk.TclError:
            return True

    def toggle_append(self):
        pass

    def browse_file(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=self.output_var.get())
        if path:
            self.output_var.set(path)

    def start_scraping(self):
        base = self.base_var.get().strip()
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()
        output = self.output_var.get().strip()
        url = self.url_var.get().strip()

        if not (base and start.isdigit() and end.isdigit()):
            messagebox.showerror("Error", "Please enter valid USN base and numeric start/end")
            return

        if not url:
            messagebox.showerror("Error", "Please enter a valid results URL")
            return

        self.start_btn.config(state=tk.DISABLED)
        self.retry_btn.config(state=tk.DISABLED)
        stop_flag.clear()

        usn_list = generate_usn_list(base=base, start=int(start), end=int(end))
        threading.Thread(
            target=run_scraper,
            args=(usn_list, output, self.log_queue, self.progress_queue,
                  self.append_var.get(), url, not self.headless_var.get()),  # Invert the value for headless
            daemon=True
        ).start()

    def stop_scraping(self):
        stop_flag.set()
        self.start_btn.config(state=tk.NORMAL)
        self.retry_btn.config(state=tk.NORMAL)

    def retry_missing_usns(self):
        """Handle the retry missing USNs operation with full GUI integration"""
        base = self.base_var.get().strip()
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()
        output = self.output_var.get().strip()
        url = self.url_var.get().strip()

        # Input validation
        if not (base and start.isdigit() and end.isdigit()):
            messagebox.showerror("Error", "Please enter valid USN base and numeric start/end")
            return

        if not url:
            messagebox.showerror("Error", "Please enter a valid results URL")
            return

        if not output:
            messagebox.showerror("Error", "Please specify an output file")
            return

        # Generate full USN list and find missing ones
        try:
            usn_list = generate_usn_list(base=base, start=int(start), end=int(end))
            missing = get_missing_usns(usn_list, output)

            if not missing:
                messagebox.showinfo("Info", "No missing USNs found. All data already fetched.")
                return
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate USN list: {str(e)}")
            return

        # Create popup window for USN confirmation/editing
        popup = Toplevel(self)
        popup.title("Missing USNs - Edit & Confirm")
        popup.geometry("600x400")
        popup.resizable(True, True)

        try:
            popup.iconbitmap(resource_path("icon.ico"))
        except Exception as icon_error:
            print(f"Popup icon load failed: {icon_error}")

        # Center the popup
        popup.update_idletasks()
        width = popup.winfo_width()
        height = popup.winfo_height()
        x = (popup.winfo_screenwidth() // 2) - (width // 2)
        y = (popup.winfo_screenheight() // 2) - (height // 2)
        popup.geometry(f'+{x}+{y}')

        # Content frame
        content_frame = ttk.Frame(popup, padding=15)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Information label
        info_text = (f"Found {len(missing)} missing USNs from total {len(usn_list)}.\n"
                     "Edit the list below if needed (one USN per line or comma-separated):")
        ttk.Label(content_frame, text=info_text).pack(anchor=tk.W, pady=(0, 10))

        # Text box with scrollbar
        text_frame = ttk.Frame(content_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text_box = Text(text_frame,
                        wrap=tk.WORD,
                        yscrollcommand=scrollbar.set,
                        height=12,
                        padx=5,
                        pady=5,
                        font=('Consolas', 10))
        text_box.insert(tk.END, "\n".join(missing))
        text_box.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_box.yview)

        # Options frame
        options_frame = ttk.Frame(content_frame)
        options_frame.pack(fill=tk.X, pady=(15, 5))

        # Show browser checkbox
        show_browser_var = tk.BooleanVar(value=not self.headless_var.get())
        show_browser_check = ttk.Checkbutton(options_frame,
                                             text="Show browser window during scraping",
                                             variable=show_browser_var)
        show_browser_check.pack(anchor=tk.W)

        # Button frame
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill=tk.X, pady=(5, 0))

        def confirm_edit():
            """Handle the confirmation of USNs to retry"""
            edited_text = text_box.get("1.0", tk.END).strip()
            popup.destroy()

            # Parse USNs from text box
            usn_list = []
            for line in edited_text.splitlines():
                line = line.strip()
                if line:
                    # Handle both comma-separated and line-separated USNs
                    usn_list.extend([u.strip() for u in line.split(",") if u.strip()])

            if not usn_list:
                messagebox.showinfo("Info", "No USNs entered to fetch")
                return

            # Start scraping with the selected USNs
            self.start_btn.config(state=tk.DISABLED)
            self.retry_btn.config(state=tk.DISABLED)
            stop_flag.clear()

            threading.Thread(
                target=run_scraper,
                args=(usn_list, output, self.log_queue, self.progress_queue,
                      True, url, not show_browser_var.get()),  # Invert for headless
                daemon=True
            ).start()

        # Action buttons
        ttk.Button(button_frame,
                   text="Fetch Missing USNs",
                   command=confirm_edit,
                   style='Primary.TButton').pack(side=tk.LEFT, padx=5, expand=True)

        ttk.Button(button_frame,
                   text="Cancel",
                   command=popup.destroy,
                   style='Secondary.TButton').pack(side=tk.LEFT, padx=5, expand=True)

        # Set focus to text box
        text_box.focus_set()

    def process_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_text.insert(tk.END, msg)
            self.log_text.see(tk.END)

        while not self.progress_queue.empty():
            val = self.progress_queue.get_nowait()
            self.progress['value'] = val

            # Re-enable buttons when progress completes
            if val == 100:
                self.start_btn.config(state=tk.NORMAL)
                self.retry_btn.config(state=tk.NORMAL)

    def destroy(self):
        """Override destroy to clean up scheduled callbacks"""
        if self.process_queue_id:
            self.after_cancel(self.process_queue_id)
        super().destroy()

if __name__ == "__main__":
    app = VTUGUI()
    app.mainloop()
'''

import threading
import queue
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Toplevel, Scrollbar, Text
from captcha_handler import CaptchaHandler
from vtu_marks_scraper import generate_usn_list, fetch_vtu_result_with_retry, get_driver
from student_data import parse_student_result
import pandas as pd
from Analyzer import analyze_results

# Control flag for stopping threads
stop_flag = threading.Event()
DEFAULT_URL = "https://results.vtu.ac.in/DJcbcs25/index.php"


def get_missing_usns(expected_usns, output_file):
    """Get list of USNs not present in output file"""
    if not os.path.exists(output_file):
        return expected_usns
    try:
        df = pd.read_excel(output_file)
        existing_usns = df['University Seat Number'].astype(str).str.upper().unique()
        return [usn for usn in expected_usns if usn.upper() not in existing_usns]
    except Exception:
        return expected_usns


def run_scraper(usn_list, output_path, log_queue, progress_queue, append=False, base_url=DEFAULT_URL, headless=True):
    """Main scraping function to run in thread"""
    try:
        total = len(usn_list)
        results = []
        handler = CaptchaHandler()
        count = 0

        driver = get_driver(headless=headless)
        missing_usns = []

        for usn in usn_list:
            if stop_flag.is_set():
                log_queue.put("Process manually stopped by user.\n")
                break

            count += 1
            progress = int((count / total) * 100)
            progress_queue.put(progress)
            log_queue.put(f"[{count}/{total}] Fetching: {usn}\n")

            html = fetch_vtu_result_with_retry(driver, usn, handler, base_url=base_url)
            if html:
                tmp_file = f"_tmp_{usn}.html"
                with open(tmp_file, "w", encoding="utf-8") as f:
                    f.write(html)
                df = parse_student_result(tmp_file)
                results.append(df)
                os.remove(tmp_file)
            else:
                missing_usns.append(usn)
                log_queue.put(f"  -> Failed to fetch: {usn}\n")

        driver.quit()

        if results:
            df_all = pd.concat(results, ignore_index=True)
            df_all.fillna("NA", inplace=True)

            if append and os.path.exists(output_path):
                old_df = pd.read_excel(output_path)
                df_all = pd.concat([old_df, df_all], ignore_index=True).drop_duplicates()

            df_all.to_excel(output_path, index=False)
            log_queue.put(f"Results saved to: {output_path}\n")

        if missing_usns:
            log_queue.put("Missing USNs:\n" + ", ".join(missing_usns) + "\n")

        if not results:
            log_queue.put("No results to save.\n")

    except Exception as e:
        log_queue.put(f"Error: {str(e)}\n")
    finally:
        progress_queue.put(100)
        log_queue.put("Finished.\n")
        stop_flag.clear()

import sys

def resource_path(relative_path):
    """Get absolute path to resource, works for PyInstaller and normal runs"""
    try:
        base_path = sys._MEIPASS  # Set by PyInstaller at runtime
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class VTUGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VTU Marks Scraper GUI")
        self.geometry("600x650")

        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception as e:
            print(f"Icon load failed: {e}")

        # Configure styles
        self.style = ttk.Style(self)
        self.style.theme_use('clam')

        # Configure button styles
        self.style.configure('TButton',
                             font=('Helvetica', 10),
                             padding=6,
                             relief=tk.RAISED,
                             borderwidth=2)

        self.style.configure('Primary.TButton',
                             foreground='white',
                             background='#007bff',
                             font=('Helvetica', 10, 'bold'))

        self.style.configure('Secondary.TButton',
                             foreground='white',
                             background='#6c757d')

        self.style.configure('Stop.TButton',
                             foreground='white',
                             background='#dc3545')

        self.style.map('Primary.TButton',
                       background=[('active', '#0056b3'), ('disabled', '#cccccc')])

        self.style.map('Secondary.TButton',
                       background=[('active', '#5a6268'), ('disabled', '#cccccc')])

        self.style.map('Stop.TButton',
                       background=[('active', '#c82333'), ('disabled', '#cccccc')])

        # Main container frame
        main_frame = ttk.Frame(self, padding=(15, 10))
        main_frame.pack(fill=tk.BOTH, expand=True)

        # URL frame
        url_frame = ttk.LabelFrame(main_frame, text="Results URL", padding=(10, 5))
        url_frame.pack(fill=tk.X, pady=(0, 10))

        self.url_var = tk.StringVar(value=DEFAULT_URL)
        ttk.Entry(url_frame, textvariable=self.url_var, width=50).pack(fill=tk.X, padx=5, pady=5)

        # Input frame
        input_frame = ttk.LabelFrame(main_frame, text="USN Range Configuration", padding=(10, 5))
        input_frame.pack(fill=tk.X, pady=(0, 10))

        # USN Base row
        base_row = ttk.Frame(input_frame)
        base_row.pack(fill=tk.X, pady=5)
        ttk.Label(base_row, text="USN Base:").pack(side=tk.LEFT, padx=(0, 5))
        self.base_var = tk.StringVar(value="1CR24BA")
        ttk.Entry(base_row, textvariable=self.base_var, width=20).pack(side=tk.LEFT, padx=5)

        # Start/End row
        range_row = ttk.Frame(input_frame)
        range_row.pack(fill=tk.X, pady=5)
        ttk.Label(range_row, text="Start:").pack(side=tk.LEFT, padx=(0, 5))
        self.start_var = tk.StringVar(value="1")
        ttk.Entry(range_row, textvariable=self.start_var, width=5).pack(side=tk.LEFT, padx=5)

        ttk.Label(range_row, text="End:").pack(side=tk.LEFT, padx=(5, 5))
        self.end_var = tk.StringVar(value="10")
        ttk.Entry(range_row, textvariable=self.end_var, width=5).pack(side=tk.LEFT, padx=5)

        # Output file row
        output_row = ttk.Frame(input_frame)
        output_row.pack(fill=tk.X, pady=5)
        ttk.Label(output_row, text="Output File:").pack(side=tk.LEFT, padx=(0, 5))
        self.output_var = tk.StringVar(value="results.xlsx")
        ttk.Entry(output_row, textvariable=self.output_var, width=30).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(output_row, text="Browse", command=self.browse_file, style='Secondary.TButton').pack(side=tk.LEFT,
                                                                                                        padx=5)

        # Options frame
        options_frame = ttk.Frame(input_frame)
        options_frame.pack(fill=tk.X, pady=(5, 0))

        # Append to existing file checkbox
        self.append_var = tk.BooleanVar()
        self.append_check = ttk.Checkbutton(options_frame,
                                            text="Append to existing file",
                                            variable=self.append_var)
        self.append_check.pack(side=tk.LEFT, padx=(0, 10))

        # Show browser checkbox
        self.headless_var = tk.BooleanVar(value=True)
        self.headless_check = ttk.Checkbutton(options_frame,
                                              text="Show browser window",
                                              variable=self.headless_var)
        self.headless_check.pack(side=tk.LEFT)

        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 5))

        self.start_btn = ttk.Button(button_frame,
                                    text="Start New Scrape",
                                    command=self.start_scraping,
                                    style='Primary.TButton')
        self.start_btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.retry_btn = ttk.Button(button_frame,
                                    text="Check & Retry Missing",
                                    command=self.retry_missing_usns,
                                    style='Secondary.TButton')
        self.retry_btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.analyze_btn = ttk.Button(button_frame,
                                      text="Generate Report",
                                      command=self.analyze_data,
                                      style='Primary.TButton')
        self.analyze_btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.stop_btn = ttk.Button(button_frame,
                                   text="Stop",
                                   command=self.stop_scraping,
                                   style='Stop.TButton')
        self.stop_btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=700, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(10, 5))

        # Log output
        log_frame = ttk.LabelFrame(main_frame, text="Log Output", padding=(5, 5))
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        self.log_text = Text(log_frame,
                             height=25,
                             bg="#1e1e1e",
                             fg="#ffffff",
                             font=('Consolas', 10),
                             wrap=tk.WORD,
                             padx=5,
                             pady=5)

        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.process_queue_id = None
        self.after(100, self._process_queue)

    def _process_queue(self):
        """Internal method to process queue messages with error handling"""
        try:
            self.process_queue()
        except tk.TclError as e:
            if "invalid command name" not in str(e):
                raise
        finally:
            if not self._is_destroyed():
                self.process_queue_id = self.after(100, self._process_queue)

    def _is_destroyed(self):
        """Check if window is being destroyed"""
        try:
            return not self.winfo_exists()
        except tk.TclError:
            return True

    def browse_file(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=self.output_var.get())
        if path:
            self.output_var.set(path)

    def start_scraping(self):
        base = self.base_var.get().strip()
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()
        output = self.output_var.get().strip()
        url = self.url_var.get().strip()

        if not (base and start.isdigit() and end.isdigit()):
            messagebox.showerror("Error", "Please enter valid USN base and numeric start/end")
            return

        if not url:
            messagebox.showerror("Error", "Please enter a valid results URL")
            return

        self.start_btn.config(state=tk.DISABLED)
        self.retry_btn.config(state=tk.DISABLED)
        self.analyze_btn.config(state=tk.DISABLED)
        stop_flag.clear()

        usn_list = generate_usn_list(base=base, start=int(start), end=int(end))
        threading.Thread(
            target=run_scraper,
            args=(usn_list, output, self.log_queue, self.progress_queue,
                  self.append_var.get(), url, not self.headless_var.get()),
            daemon=True
        ).start()

    def stop_scraping(self):
        stop_flag.set()
        self.start_btn.config(state=tk.NORMAL)
        self.retry_btn.config(state=tk.NORMAL)
        self.analyze_btn.config(state=tk.NORMAL)

    def retry_missing_usns(self):
        """Handle the retry missing USNs operation with full GUI integration"""
        base = self.base_var.get().strip()
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()
        output = self.output_var.get().strip()
        url = self.url_var.get().strip()

        if not (base and start.isdigit() and end.isdigit()):
            messagebox.showerror("Error", "Please enter valid USN base and numeric start/end")
            return

        if not url:
            messagebox.showerror("Error", "Please enter a valid results URL")
            return

        if not output:
            messagebox.showerror("Error", "Please specify an output file")
            return

        try:
            usn_list = generate_usn_list(base=base, start=int(start), end=int(end))
            missing = get_missing_usns(usn_list, output)

            if not missing:
                messagebox.showinfo("Info", "No missing USNs found. All data already fetched.")
                return
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate USN list: {str(e)}")
            return


        popup = Toplevel(self)
        popup.title("Missing USNs - Edit & Confirm")
        popup.geometry("600x400")
        popup.resizable(True, True)
        try:
            popup.iconbitmap(resource_path("icon.ico"))
        except Exception as icon_error:
            print(f"Popup icon load failed: {icon_error}")


        # Center the popup
        popup.update_idletasks()
        width = popup.winfo_width()
        height = popup.winfo_height()
        x = (popup.winfo_screenwidth() // 2) - (width // 2)
        y = (popup.winfo_screenheight() // 2) - (height // 2)
        popup.geometry(f'+{x}+{y}')

        # Content frame
        content_frame = ttk.Frame(popup, padding=15)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Information label
        info_text = (f"Found {len(missing)} missing USNs from total {len(usn_list)}.\n"
                     "Edit the list below if needed (one USN per line or comma-separated):")
        ttk.Label(content_frame, text=info_text).pack(anchor=tk.W, pady=(0, 10))

        # Text box with scrollbar
        text_frame = ttk.Frame(content_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text_box = Text(text_frame,
                        wrap=tk.WORD,
                        yscrollcommand=scrollbar.set,
                        height=12,
                        padx=5,
                        pady=5,
                        font=('Consolas', 10))
        text_box.insert(tk.END, "\n".join(missing))
        text_box.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_box.yview)

        # Options frame
        options_frame = ttk.Frame(content_frame)
        options_frame.pack(fill=tk.X, pady=(15, 5))

        # Show browser checkbox
        show_browser_var = tk.BooleanVar(value=not self.headless_var.get())
        show_browser_check = ttk.Checkbutton(options_frame,
                                             text="Show browser window during scraping",
                                             variable=show_browser_var)
        show_browser_check.pack(anchor=tk.W)

        # Button frame
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(fill=tk.X, pady=(5, 0))

        def confirm_edit():
            """Handle the confirmation of USNs to retry"""
            edited_text = text_box.get("1.0", tk.END).strip()
            popup.destroy()

            # Parse USNs from text box
            usn_list = []
            for line in edited_text.splitlines():
                line = line.strip()
                if line:
                    usn_list.extend([u.strip() for u in line.split(",") if u.strip()])

            if not usn_list:
                messagebox.showinfo("Info", "No USNs entered to fetch")
                return

            # Start scraping with the selected USNs
            self.start_btn.config(state=tk.DISABLED)
            self.retry_btn.config(state=tk.DISABLED)
            self.analyze_btn.config(state=tk.DISABLED)
            stop_flag.clear()

            threading.Thread(
                target=run_scraper,
                args=(usn_list, output, self.log_queue, self.progress_queue,
                      True, url, not show_browser_var.get()),
                daemon=True
            ).start()

        # Action buttons
        ttk.Button(button_frame,
                   text="Fetch Missing USNs",
                   command=confirm_edit,
                   style='Primary.TButton').pack(side=tk.LEFT, padx=5, expand=True)

        ttk.Button(button_frame,
                   text="Cancel",
                   command=popup.destroy,
                   style='Secondary.TButton').pack(side=tk.LEFT, padx=5, expand=True)

        # Set focus to text box
        text_box.focus_set()

    def analyze_data(self):
        """Handle analysis of the collected results"""
        output_file = self.output_var.get().strip()

        if not output_file:
            messagebox.showerror("Error", "Please specify an output file first")
            return

        if not os.path.exists(output_file):
            messagebox.showerror("Error", f"Output file not found: {output_file}")
            return

        try:
            # Disable button during analysis
            self.analyze_btn.config(state=tk.DISABLED)
            self.log_queue.put("\nStarting analysis of results...\n")

            # Run analysis in a separate thread
            threading.Thread(
                target=self._run_analysis,
                args=(output_file,),
                daemon=True
            ).start()

        except Exception as e:
            messagebox.showerror("Analysis Error", f"Failed to start analysis: {str(e)}")
            self.analyze_btn.config(state=tk.NORMAL)

    def _run_analysis(self, excel_file):
        """Thread-safe analysis runner with proper error handling"""
        try:
            output_file = analyze_results(excel_file)
            self.log_queue.put(f"\nAnalysis report saved as: {output_file}\n")
            self.after(100, lambda f=output_file: messagebox.showinfo(
                "Analysis Complete",
                f"Report generated successfully!\nSaved as: {f}"
            ))
        except Exception as e:
            error_msg = f"\nAnalysis failed: {str(e)}\n"
            self.log_queue.put(error_msg)
            self.after(100, lambda msg=error_msg: messagebox.showerror(
                "Analysis Error",
                f"Failed to generate report:\n{msg}"
            ))
        finally:
            self.after(100, lambda: self.analyze_btn.config(state=tk.NORMAL))


    def process_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_text.insert(tk.END, msg)
            self.log_text.see(tk.END)

        while not self.progress_queue.empty():
            val = self.progress_queue.get_nowait()
            self.progress['value'] = val

            # Re-enable buttons when progress completes
            if val == 100:
                self.start_btn.config(state=tk.NORMAL)
                self.retry_btn.config(state=tk.NORMAL)
                self.analyze_btn.config(state=tk.NORMAL)

    def destroy(self):
        """Override destroy to clean up scheduled callbacks"""
        if self.process_queue_id:
            self.after_cancel(self.process_queue_id)
        super().destroy()


if __name__ == "__main__":
    app = VTUGUI()
    app.mainloop()