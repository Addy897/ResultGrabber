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
        self.title("Result Grabber")
        self.geometry("420x380")
        self.fname = ""
        self.cancel_event = None
        self.protocol("WM_DELETE_WINDOW", self.on_closing)



        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.loading_bar = ctk.CTkProgressBar(self, width=200)
        self.loading_bar.grid(row=0, column=0, columnspan=2, pady=20, sticky="n")
        self.loading_bar.configure(mode="indeterminate")
        self.loading_bar.start()

        self.loading_label = ctk.CTkLabel(self, text="Loading EasyOCR, please wait...", font=ctk.CTkFont(size=14))
        self.loading_label.grid(row=1, column=0, columnspan=2, pady=20, sticky="n")


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
            self.init_widgets()

    def init_widgets(self):

        self.label_url = ctk.CTkLabel(self, text="URL", font=ctk.CTkFont(size=16, family="Cascadia Mono", weight="normal"))
        self.label_url.grid(row=0, column=0, padx=20, pady=(10, 20), sticky="w")
        self.entry_url = ctk.CTkEntry(self, placeholder_text="Enter URL of VTU Site", width=200)
        self.entry_url.grid(row=0, column=1, padx=10, pady=(20, 20), sticky="nsew")


        self.label_range = ctk.CTkLabel(self, text="Starting USN", font=ctk.CTkFont(size=16, family="Cascadia Mono", weight="normal"))
        self.label_range.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="w")
        self.entry_rstart = ctk.CTkEntry(self, placeholder_text="Ex. 1BI23IC001")
        self.entry_rstart.grid(row=1, column=1, padx=10, pady=(0, 20), sticky="nsew")


        self.label_range1 = ctk.CTkLabel(self, text="Ending USN", font=ctk.CTkFont(size=16, family="Cascadia Mono", weight="normal"))
        self.label_range1.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="w")
        self.entry_rend = ctk.CTkEntry(self, placeholder_text="Ex. 1BI23IC063")
        self.entry_rend.grid(row=2, column=1, padx=10, pady=(0, 20), sticky="nsew")


        self.label_range2 = ctk.CTkLabel(self, text="Diploma Starting", font=ctk.CTkFont(size=16, family="Cascadia Mono", weight="normal"))
        self.label_range2.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="w")
        self.entry_drstart = ctk.CTkEntry(self, placeholder_text="Ex. 1BI23IC001")
        self.entry_drstart.grid(row=3, column=1, padx=10, pady=(0, 20), sticky="nsew")


        self.label_range3 = ctk.CTkLabel(self, text="Diploma Ending", font=ctk.CTkFont(size=16, family="Cascadia Mono", weight="normal"))
        self.label_range3.grid(row=4, column=0, padx=20, pady=(0, 20), sticky="w")
        self.entry_drend = ctk.CTkEntry(self, placeholder_text="Ex. 1BI23IC063")
        self.entry_drend.grid(row=4, column=1, padx=10, pady=(0, 20), sticky="nsew")


        self.label_filename = ctk.CTkLabel(self, text="Filename", font=ctk.CTkFont(size=16, family="Cascadia Mono", weight="normal"))
        self.label_filename.grid(row=5, column=0, padx=20, pady=(0, 20), sticky="w")
        self.entry_filename = ctk.CTkEntry(self, placeholder_text="Filename of the saved data")
        self.entry_filename.grid(row=5, column=1, padx=10, pady=(0, 20), sticky="nsew")
        self.show_browser_var = ctk.BooleanVar(value=True)
        self.checkbox_show_browser = ctk.CTkCheckBox(
            master=self, text="Show Browser", variable=self.show_browser_var
        )
        self.checkbox_show_browser.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="w")
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.button_grab = ctk.CTkButton(master=self, text="Grab", fg_color="transparent",
                                         border_width=2, text_color=("gray10", "#DCE4EE"), command=self.grab_results)
        self.button_grab.grid(row=6, column=1, padx=20, pady=(0, 20), sticky="nsew")
    def hide_form_widgets(self):
        self.label_url.grid_remove()
        self.entry_url.grid_remove()
        self.entry_rstart.grid_remove()
        self.entry_rend.grid_remove()
        self.entry_drstart.grid_remove()
        self.entry_drend.grid_remove()
        self.label_filename.grid_remove()
        self.entry_filename.grid_remove()
        self.button_grab.grid_remove()
        self.label_range.grid_remove()
        self.label_range1.grid_remove()
        self.label_range2.grid_remove()
        self.label_range3.grid_remove()
        self.checkbox_show_browser.grid_remove()
    def show_form_widgets(self):
        self.label_url.grid()
        self.entry_url.grid()
        self.entry_rstart.grid()
        self.entry_rend.grid()
        self.entry_drstart.grid()
        self.entry_drend.grid()
        self.label_filename.grid()
        self.entry_filename.grid()
        self.button_grab.grid()
        self.label_range.grid()
        self.label_range1.grid()
        self.label_range2.grid()
        self.label_range3.grid()
        self.checkbox_show_browser.grid()
    def extract_segments(self, text):
        pattern = re.compile(r'^(\d)([A-Z]{2})(\d{2})([A-Z]{2})(\d{3})$',re.IGNORECASE)
        match = pattern.match(text)
        if match:
            try:
                region = match.group(1)
                region_char = match.group(2)
                year = int(match.group(3))
                branch = match.group(4)
                usn_no = int(match.group(5))
                return (region, region_char, year, branch, usn_no)
            except Exception:
                return None
        else:
            return None

    def grab_results(self):
        starting_usn = self.entry_rstart.get()
        ending_usn = self.entry_rend.get()
        dip_starting_usn = self.entry_drstart.get()
        dip_ending_usn = self.entry_drend.get()
        url = self.entry_url.get()
        filename = self.entry_filename.get()
        url_pattern = r"^https:\/\/results\.vtu\.ac\.in\/.*"
        match = re.match(url_pattern, url)
        start_seg = self.extract_segments(text=starting_usn)
        end_seg = self.extract_segments(text=ending_usn)
        dip_start_seg = self.extract_segments(text=dip_starting_usn)
        dip_end_seg = self.extract_segments(text=dip_ending_usn)


        if not match:
            mb.showerror(title="Invalid Field", message="Invalid URL")
            self.entry_url.focus()
            return
        elif not start_seg:
            mb.showerror(title="Empty Field", message="Enter Starting USN")
            self.entry_rstart.focus()
            return
        elif not end_seg:
            mb.showerror(title="Empty Field", message="Enter Ending USN")
            self.entry_rend.focus()
            return

        if dip_start_seg and dip_end_seg:
            l = [(end_seg[-1], starting_usn[:7], start_seg[-1]),
                 (dip_end_seg[-1], dip_starting_usn[:7], dip_start_seg[-1])]
        else:
            l = [(end_seg[-1], starting_usn[:7], start_seg[-1])]
        if not filename:
            self.fname = f"{starting_usn}-{end_seg[-1]}.xlsx"
        else:
            self.fname = filename
        self.cancel_event = threading.Event()

        self.hide_form_widgets()
        self.scraping_label = ctk.CTkLabel(
            self, text="Scraping in progress...", font=ctk.CTkFont(size=14)
        )
        self.scraping_label.grid(row=0, column=0, columnspan=2, pady=20)
        self.cancel_button = ctk.CTkButton(
            self, text="Cancel", command=self.cancel_scraping
        )
        self.cancel_button.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")
        threading.Thread(target=self.fetch_and_dump, args=(l, url,self.cancel_event,self.show_browser_var), daemon=True).start()
    def cancel_scraping(self):
            if self.cancel_event:
                self.cancel_event.set()
                mb.showinfo("Cancelled", "Scraping has been cancelled. Dumping Fetched Students")
                self.scraping_label.configure(text="Scraping cancelled.")
                self.scraping_label.destroy()
                self.cancel_button.destroy()
                self.show_form_widgets()
    def on_closing(self):
        if self.cancel_event and not self.cancel_event.is_set():
            self.cancel_event.set()
        self.destroy()
    def fetch_and_dump(self, l, url,cancel,show):
        students = fetch(l, url,cancel,show)
        if cancel.is_set():
            dump(self.fname, students,False)
            return
        dump(self.fname, students)
        self.scraping_label.destroy()
        self.cancel_button.destroy()
        self.show_form_widgets()
        mb.showinfo(title="Success", message=f"Saved Results in {self.fname} ")



if __name__ == "__main__":
    app = App()
    app.mainloop()
