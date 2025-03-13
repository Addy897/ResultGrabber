
import customtkinter as ctk
import re
from scrape import fetch,dump
from tkinter import messagebox as mb
from threading import Thread
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Result Grabber")
        self.geometry(f"{420}x{380}")
        self.fname=""
        self.label_url=ctk.CTkLabel(self, text="URL", font=ctk.CTkFont(size=16,family="Cascadia Mono", weight="normal"))
        self.label_url.grid(row=0, column=0, padx=20, pady=(10, 20),sticky="w")
        self.entry_url = ctk.CTkEntry(self, placeholder_text="Enter URL of VTU Site",width=200)
        self.entry_url.grid(row=0, column=1, columnspan=1, padx=(10, 0), pady=(20, 20), sticky="nsew")


        self.label_range=ctk.CTkLabel(self, text="Starting USN", font=ctk.CTkFont(size=16,family="Cascadia Mono", weight="normal"))
        self.label_range.grid(row=1, column=0, padx=20, pady=(0, 20),sticky="w")
        self.entry_rstart = ctk.CTkEntry(self, placeholder_text="Ex. 1BI23IC001")
        self.entry_rstart.grid(row=1, column=1, columnspan=1, padx=(10, 0), pady=(0, 20), sticky="nsew")

        self.label_range1=ctk.CTkLabel(self, text="Ending USN", font=ctk.CTkFont(size=16,family="Cascadia Mono", weight="normal"))
        self.label_range1.grid(row=2, column=0, padx=20, pady=(0, 20),sticky="w")
        self.entry_rend = ctk.CTkEntry(self, placeholder_text="Ex. 1BI23IC063")
        self.entry_rend.grid(row=2, column=1, columnspan=1, padx=(10, 0), pady=(0, 20), sticky="nsew")

        self.label_range2=ctk.CTkLabel(self, text="Diploma Starting", font=ctk.CTkFont(size=16,family="Cascadia Mono", weight="normal"))
        self.label_range2.grid(row=3, column=0, padx=20, pady=(0, 20),sticky="w")
        self.entry_drstart = ctk.CTkEntry(self, placeholder_text="Ex. 1BI23IC001")
        self.entry_drstart.grid(row=3, column=1, columnspan=1, padx=(10, 0), pady=(0, 20), sticky="nsew")

        self.label_range3=ctk.CTkLabel(self, text="Diploma Ending", font=ctk.CTkFont(size=16,family="Cascadia Mono", weight="normal"))
        self.label_range3.grid(row=4, column=0, padx=20, pady=(0, 20),sticky="w")
        self.entry_drend = ctk.CTkEntry(self, placeholder_text="Ex. 1BI23IC063")
        self.entry_drend.grid(row=4, column=1, columnspan=1, padx=(10, 0), pady=(0, 20), sticky="nsew")

        self.label_filename=ctk.CTkLabel(self, text="Filename", font=ctk.CTkFont(size=16,family="Cascadia Mono", weight="normal"))
        self.label_filename.grid(row=5, column=0, padx=20, pady=(0, 20),sticky="w")
        self.entry_filename = ctk.CTkEntry(self, placeholder_text="Filename of the saved data")
        self.entry_filename.grid(row=5, column=1, columnspan=1, padx=(10, 0), pady=(0, 20), sticky="nsew")

        self.button_grab = ctk.CTkButton(master=self, text="Grab",fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"),command=self.grab_results)
        self.button_grab.grid(row=6, column=0, padx=(20, 0), pady=(0, 20), sticky="nsew")
    def extract_segments(self,text):
        pattern = r'^(\d)([A-Z]{2})(\d{2})([A-Z]{2})(\d{3})$'
        match = re.match(pattern, text)
        if match:
            try:
                region = match.group(1)
                region_char = match.group(2)
                year = int(match.group(3))
                branch = match.group(4)
                usn_no = int(match.group(5))

                return (region, region_char, year, branch, usn_no)
            except:
                return None

        else:
            return None
    def grab_results(self):
        starting_usn=self.entry_rstart.get()
        ending_usn=self.entry_rend.get()
        dip_starting_usn=self.entry_drstart.get()
        dip_ending_usn=self.entry_drend.get()
        url=self.entry_url.get()
        filename=self.entry_filename.get()
        url_pattern="^https:\/\/results\.vtu\.ac\.in\/.*"
        match=re.match(url_pattern,url)
        start_seg=self.extract_segments(text=starting_usn)
        end_seg=self.extract_segments(text=ending_usn)
        dip_start_seg=self.extract_segments(dip_starting_usn)
        dip_end_seg=self.extract_segments(dip_ending_usn)
        if(not filename):
            self.fname=f"{starting_usn}-{end_seg[-1]}.xlsx"
        else:
            self.fname=filename
        if(not match):
             mb.showerror(title="Invalid Field",message="Invlid URL")
             self.entry_url.focus()

             return
        elif(not start_seg):
               mb.showerror(title="Empty Field",message="Enter Starting USN")
               self.entry_rstart.focus()
               return
        elif(not end_seg):
             mb.showerror(title="Empty Field",message="Enter Ending USN")
             self.entry_rend.focus()
             return
        if(dip_start_seg and dip_end_seg):
            l=[(end_seg[-1],starting_usn[:7],start_seg[-1]),(dip_end_seg[-1],dip_starting_usn[:7],dip_start_seg[-1])]
        else:
            l=[(end_seg[-1],starting_usn[:7],start_seg[-1])]
        t=Thread(target=self.fetch_and_dump,args=(l,url))
        t.daemon=False
        t.start()
    def fetch_and_dump(self,l,url):
        students=fetch(l,url)
        dump(self.fname,students)








if __name__ == "__main__":
    app = App()
    app.mainloop()
