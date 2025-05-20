import customtkinter as ctk
import re
import threading
from scrape import fetch, dump, initReader
import scrape
from tkinter import messagebox as mb

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("VTU Result Tool")
        self.geometry("480x480")
        self.fname = ""
        self.cancel_event = None
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.loading_bar = ctk.CTkProgressBar(self, width=200, mode="indeterminate")
        self.loading_bar.pack(pady=20)
        self.loading_bar.start()

        self.loading_label = ctk.CTkLabel(self, text="Loading EasyOCR, please wait...", font=ctk.CTkFont(size=14))
        self.loading_label.pack()

        threading.Thread(target=self.init_easyocr, daemon=True).start()
        self.check_easyocr_ready()

    def init_easyocr(self):
        initReader()

    def check_easyocr_ready(self):
        if scrape.reader is None:
            self.after(100, self.check_easyocr_ready)
        else:
            self.loading_label.destroy()
            self.loading_bar.destroy()
            self.init_tabs()

    def init_tabs(self):
        self.tabs = ctk.CTkTabview(self, width=460, height=420)
        self.tabs.pack(pady=10, padx=10, expand=True, fill="both")

        self.result_tab = self.tabs.add("Fetch & Dump Results")
        self.view_tab = self.tabs.add("View Results")
        self.stats_tab = self.tabs.add("Statistics")
        self.recheck_tab = self.tabs.add("Recheck")

        self.init_result_tab()
        # Future implementation for other tabs

    def init_result_tab(self):
        tab = self.result_tab
        tab.grid_columnconfigure((0, 1), weight=1)

        self.entry_url = ctk.CTkEntry(tab, placeholder_text="Enter VTU Result URL")
        self.entry_url.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="ew")

        self.entry_rstart = ctk.CTkEntry(tab, placeholder_text="Starting USN")
        self.entry_rstart.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        self.entry_rend = ctk.CTkEntry(tab, placeholder_text="Ending USN")
        self.entry_rend.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        self.entry_drstart = ctk.CTkEntry(tab, placeholder_text="Diploma Starting USN")
        self.entry_drstart.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

        self.entry_drend = ctk.CTkEntry(tab, placeholder_text="Diploma Ending USN")
        self.entry_drend.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        self.entry_filename = ctk.CTkEntry(tab, placeholder_text="Filename to save")
        self.entry_filename.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        self.show_browser_var = ctk.BooleanVar(value=True)
        self.checkbox_show_browser = ctk.CTkCheckBox(tab, text="Show Browser", variable=self.show_browser_var)
        self.checkbox_show_browser.grid(row=4, column=0, padx=10, pady=5, sticky="w")

        self.button_grab = ctk.CTkButton(tab, text="Grab Results", command=self.grab_results)
        self.button_grab.grid(row=4, column=1, padx=10, pady=5, sticky="e")

    def extract_segments(self, text):
        pattern = re.compile(r'^(\d)([A-Z]{2})(\d{2})([A-Z]{2})(\d{3})$', re.IGNORECASE)
        match = pattern.match(text)
        if match:
            try:
                return (
                    match.group(1), match.group(2), int(match.group(3)),
                    match.group(4), int(match.group(5))
                )
            except:
                return None
        return None

    def grab_results(self):
        starting_usn = self.entry_rstart.get()
        ending_usn = self.entry_rend.get()
        dip_starting_usn = self.entry_drstart.get()
        dip_ending_usn = self.entry_drend.get()
        url = self.entry_url.get()
        filename = self.entry_filename.get()

        url_pattern = r"^https:\/\/results\.vtu\.ac\.in\/.*"
        if not re.match(url_pattern, url):
            mb.showerror("Invalid URL", "Please enter a valid VTU result URL.")
            return

        start_seg = self.extract_segments(starting_usn)
        end_seg = self.extract_segments(ending_usn)
        dip_start_seg = self.extract_segments(dip_starting_usn)
        dip_end_seg = self.extract_segments(dip_ending_usn)

        if not start_seg or not end_seg:
            mb.showerror("Invalid USN", "Please enter valid starting and ending USNs.")
            return

        if dip_start_seg and dip_end_seg:
            l = [(end_seg[-1], starting_usn[:7], start_seg[-1]),
                 (dip_end_seg[-1], dip_starting_usn[:7], dip_start_seg[-1])]
        else:
            l = [(end_seg[-1], starting_usn[:7], start_seg[-1])]

        self.fname = filename if filename else f"{starting_usn}-{end_seg[-1]}"
        self.cancel_event = threading.Event()

        self.progress_popup()
        threading.Thread(target=self.fetch_and_dump, args=(l, url, self.cancel_event, self.show_browser_var), daemon=True).start()

    def progress_popup(self):
        self.progress_win = ctk.CTkToplevel(self)
        self.progress_win.title("Scraping in progress")
        self.progress_win.geometry("300x100")
        self.progress_label = ctk.CTkLabel(self.progress_win, text="Scraping in progress...")
        self.progress_label.pack(pady=10)
        self.cancel_button = ctk.CTkButton(self.progress_win, text="Cancel", command=self.cancel_scraping)
        self.cancel_button.pack(pady=5)

    def cancel_scraping(self):
        if self.cancel_event:
            self.cancel_event.set()
            mb.showinfo("Cancelled", "Scraping cancelled. Partial data will be saved.")
            self.progress_win.destroy()

    def fetch_and_dump(self, l, url, cancel, show):
        students = fetch(l, url, cancel, show)
        dump(self.fname, students)
        self.progress_win.destroy()
        mb.showinfo("Done", f"Results saved to {self.fname}.")

    def on_closing(self):
        if self.cancel_event and not self.cancel_event.is_set():
            self.cancel_event.set()
        self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()
