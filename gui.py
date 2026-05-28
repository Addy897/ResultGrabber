

from io import FileIO
import customtkinter as ctk
import re
import threading


from db.models import (
    Student,
    Subject,
    Result,
)
from db.setup import Session, init_db


from scrape_new import fetch, dump, initReader,intc
import scrape_new as scrape

from tkinter import messagebox as mb

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("VTU Result Grabber")
        self.geometry("600x650")

        self.cancel_event = None
        self.fname = ""


        self.loading_label = ctk.CTkLabel(
            self, text="Loading EasyOCR, please wait...", font=ctk.CTkFont(size=14)
        )
        self.loading_label.pack(pady=20)
        self.loading_bar = ctk.CTkProgressBar(self, orientation="horizontal", width=300)
        self.loading_bar.set(0.0)
        self.loading_bar.pack(pady=10)

        threading.Thread(target=self.init_easyocr, daemon=True).start()
        self.check_easyocr_ready()


        self.protocol("WM_DELETE_WINDOW", self.on_closing)

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

        self.tabs = ctk.CTkTabview(self, width=580, height=600)
        self.tabs.pack(pady=10, padx=10, expand=True, fill="both")



        self.fetch_tab = self.tabs.add("Fetch & Dump")
        self.view_tab = self.tabs.add("View Results")
        self.overall_stats_tab = self.tabs.add("Overall Stats")
        self.sem_stats_tab = self.tabs.add("Semester Stats")
        self.subj_stats_tab = self.tabs.add("Subject Stats")
        self.stu_grades_tab = self.tabs.add("Student Grades")


        self.init_fetch_tab()
        self.init_view_tab()
        self.init_overall_stats_tab()
        self.init_semester_stats_tab()
        self.init_subject_stats_tab()
        self.init_student_grades_tab()




    def init_fetch_tab(self):
        tab = self.fetch_tab

        url_frame = ctk.CTkFrame(tab)
        url_frame.grid(row=0, column=0, columnspan=3, padx=20, pady=(20, 10), sticky="ew")
        url_frame.grid_columnconfigure(1, weight=1)

        label_url = ctk.CTkLabel(url_frame, text="VTU Result URL:", anchor="w")
        label_url.grid(row=0, column=0, padx=(5, 10), pady=5, sticky="w")

        self.entry_url = ctk.CTkEntry(url_frame, placeholder_text="https://results.vtu.ac.in/...", width=400)
        self.entry_url.grid(row=0, column=1, padx=(0, 5), pady=5, sticky="ew")

        ug_frame = ctk.CTkFrame(tab)
        ug_frame.grid(row=1, column=0, columnspan=3, padx=20, pady=(5, 10), sticky="ew")
        ug_frame.grid_columnconfigure((1, 2), weight=1)

        label_range = ctk.CTkLabel(ug_frame, text="UG USN Range:", anchor="w")
        label_range.grid(row=0, column=0, padx=(5, 10), pady=5, sticky="w")
        self.entry_rstart = ctk.CTkEntry(ug_frame, placeholder_text="Start USN (e.g. 1AB12CS001)")
        self.entry_rstart.grid(row=0, column=1, padx=(0, 5), pady=5, sticky="ew")
        self.entry_rend = ctk.CTkEntry(ug_frame, placeholder_text="End USN (e.g. 1AB12CS100)")
        self.entry_rend.grid(row=0, column=2, padx=(5, 5), pady=5, sticky="ew")

        dip_frame = ctk.CTkFrame(tab)
        dip_frame.grid(row=2, column=0, columnspan=3, padx=20, pady=(5, 10), sticky="ew")
        dip_frame.grid_columnconfigure((1, 2), weight=1)

        label_dip = ctk.CTkLabel(dip_frame, text="Diploma USN Range (optional):", anchor="w")
        label_dip.grid(row=0, column=0, padx=(5, 10), pady=5, sticky="w")
        self.entry_drstart = ctk.CTkEntry(dip_frame, placeholder_text="Dip Start USN")
        self.entry_drstart.grid(row=0, column=1, padx=(0, 5), pady=5, sticky="ew")
        self.entry_drend = ctk.CTkEntry(dip_frame, placeholder_text="Dip End USN")
        self.entry_drend.grid(row=0, column=2, padx=(5, 5), pady=5, sticky="ew")
        
        out_frame = ctk.CTkFrame(tab)
        out_frame.grid(row=3, column=0, columnspan=3, padx=20, pady=(5, 10), sticky="ew")
        out_frame.grid_columnconfigure(1, weight=1)

        label_fname = ctk.CTkLabel(out_frame, text="Output Filename:", anchor="w")
        label_fname.grid(row=0, column=0, padx=(5, 10), pady=5, sticky="w")
        self.entry_filename = ctk.CTkEntry(out_frame, placeholder_text="results")
        self.entry_filename.grid(row=0, column=1, padx=(0, 5), pady=5, sticky="ew")
        
        self.entry_sem = ctk.CTkEntry(out_frame, placeholder_text="Semester")
        self.entry_sem.grid(row=1, column=1, padx=(5, 5), pady=5, sticky="ew")


        self.var_headless = ctk.BooleanVar(value=True)
        chk_headless = ctk.CTkCheckBox(
            out_frame,
            text="Run Headless",
            variable=self.var_headless,
            onvalue=False,
            offvalue=True
        )
        chk_headless.grid(row=0, column=2, padx=(5, 5), pady=5)
        self.use_proxy = ctk.BooleanVar(value=False)
        chk_proxy = ctk.CTkCheckBox(
            out_frame,
            text="Proxy",
            variable=self.use_proxy,
            onvalue=True,
            offvalue=False
        )
        chk_proxy.grid(row=1, column=2, padx=(5, 5), pady=5)

        btn_frame = ctk.CTkFrame(tab)
        btn_frame.grid(row=4, column=0, columnspan=3, padx=20, pady=(15, 20), sticky="ew")
        btn_frame.grid_columnconfigure(0, weight=1)

        btn_fetch = ctk.CTkButton(
            btn_frame,
            text="Fetch & Dump Results",
            command=self.start_fetch_thread,
            width=200,
            corner_radius=8,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        btn_fetch.grid(row=0, column=0, pady=10)

        tab.grid_columnconfigure((0, 1, 2), weight=1)

    def start_fetch_thread(self):
        starting_usn = self.entry_rstart.get().strip().upper()
        ending_usn = self.entry_rend.get().strip().upper()
        dip_starting_usn = self.entry_drstart.get().strip().upper()
        dip_ending_usn = self.entry_drend.get().strip().upper()
        url = self.entry_url.get().strip()
        filename = self.entry_filename.get().strip()
        sem = self.entry_sem.get()
        url_pattern = r"^https:\/\/results\.vtu\.ac\.in\/.*"
        if not re.match(url_pattern, url):
            mb.showerror("Invalid URL", "Please enter a valid VTU result URL.")
            return

        if not starting_usn or not ending_usn or not filename:
            mb.showerror("Missing Data", "Enter UG USN range and output filename.")
            return

        start_seg = self.extract_segments(starting_usn)
        end_seg = self.extract_segments(ending_usn)
        if not start_seg or not end_seg:
            mb.showerror("Invalid USN", "Please enter valid starting and ending USNs.")
            return

        if dip_starting_usn and dip_ending_usn:
            dip_start_seg = self.extract_segments(dip_starting_usn)
            dip_end_seg = self.extract_segments(dip_ending_usn)
            if not (dip_start_seg and dip_end_seg):
                mb.showerror("Invalid USN", "Please enter valid Diploma USN range.")
                return
            l = [
                (end_seg[-1], starting_usn[:7], start_seg[-1]),
                (dip_end_seg[-1], dip_starting_usn[:7], dip_start_seg[-1]),
            ]
        else:
            l = [(end_seg[-1], starting_usn[:7], start_seg[-1])]

        show = self.var_headless
        self.cancel_event = threading.Event()

        self.progress_win = ctk.CTkToplevel(self)
        self.progress_win.geometry("300x100")
        self.progress_win.title("Fetching…")
        tk_label = ctk.CTkLabel(self.progress_win, text="Fetching VTU results…")
        tk_label.pack(pady=10)
        self.progress_bar = ctk.CTkProgressBar(
            self.progress_win, orientation="horizontal", width=250
        )
        self.progress_bar.set(0.0)
        self.progress_bar.pack(pady=5)
        btn_search_grade = ctk.CTkButton(
            self.progress_win, text="Cancel", command=self.cancel
        )
        btn_search_grade.pack(padx=(5,10), pady=(10,5))
        thread = threading.Thread(
            target=self.fetch_and_dump,
            args=(l, url, self.cancel_event, show,self.use_proxy,sem),
            daemon=True,
        )
        thread.start()

    def fetch_and_dump(self, l, url, cancel, show,proxy,sem):
        try:
            students = fetch(l, url, cancel, show.get(),proxy.get(),sem)
        except Exception as e:
            mb.showerror("Error during fetch", str(e))
            return

        session = Session()
        try:
            for stu in students:
                usn = stu["University Seat Number"]
                name = stu.get("Student Name", "")
                sem = stu.get("Semester")

                db_student = session.query(Student).get(usn)
                if not db_student:
                    db_student = Student(USN=usn, StudentName=name)
                    session.add(db_student)

                for code, marks in stu["marks"].items():
                    existing_subj = session.query(Subject).get(code)
                    if not existing_subj:
                        session.add(
                            Subject(
                                SubjectCode=code,
                                SubjectName=code,
                                Semester=sem if sem is not None else 0,
                            )
                        )

                    db_res = session.query(Result).get((usn, code))
                    if(marks["SEE"]=='NC' and marks["CIE"]=='NC'):
                        continue
                    
                    if not db_res:
                        db_res = Result(USN=usn, SubjectCode=code)
                    SEE=intc(marks["SEE"])
                    CIE=intc(marks["CIE"])
                    if(type(SEE)==str and SEE.startswith("NE")):
                        SEE=0
                    if(type(CIE)==str and CIE.startswith("NE")):
                        CIE=0
                    db_res.SEE = SEE
                    db_res.CIE = CIE
                    db_res.Total = CIE+SEE
                    db_res.Result = marks["Result"]
                    session.add(db_res)

            session.commit()
        except Exception as e:
            session.rollback()
            mb.showerror("Database Error", str(e))
        finally:
            session.close()
        filename=self.entry_filename.get().strip()
        if(not filename):
            filename="results.xlsx"
        try:
            dump(filename, students)
        except Exception as e:
            mb.showerror("Error writing Excel", str(e))
        finally:
            if hasattr(self, "progress_win") and self.progress_win.winfo_exists():
                self.progress_win.destroy()
            mb.showinfo("Done", f"Results saved to {filename}.")

    def _update_progress(self, fraction: float):
        if hasattr(self, "progress_bar"):
            self.progress_bar.set(fraction)




    def init_view_tab(self):
        tab = self.view_tab

        label_usn = ctk.CTkLabel(tab, text="Enter Student USN:")
        label_usn.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        self.entry_lookup_usn = ctk.CTkEntry(tab, placeholder_text="e.g. 1AB12CS001")
        self.entry_lookup_usn.grid(row=0, column=1, padx=10, pady=(10, 5), sticky="ew")

        btn_lookup = ctk.CTkButton(tab, text="Lookup", command=self.lookup_student_results)
        btn_lookup.grid(row=0, column=2, padx=10, pady=(10, 5))

        self.text_results = ctk.CTkTextbox(tab, width=550, height=300)
        self.text_results.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(1, weight=1)

    def lookup_student_results(self):
        usn = self.entry_lookup_usn.get().strip().upper()
        if not usn:
            mb.showerror("Input Error", "Please enter a USN to look up.")
            return

        session = Session()
        try:
            rows = (
                session.query(Result, Subject.SubjectName, Subject.Semester)
                .join(Subject, Result.SubjectCode == Subject.SubjectCode)
                .filter(Result.USN == usn)
                .all()
            )

            if not rows:
                self.text_results.delete("0.0", "end")
                self.text_results.insert("0.0", f"No results found for USN {usn}.")
                return

            output_lines = [f"Results for USN: {usn}", "-" * 60]
            for (res, subj_name, semester) in rows:
                line = (
                    f"Subj Code: {res.SubjectCode} | Name: {subj_name} | "
                    f"Sem: {semester} | CIE: {res.CIE} | SEE: {res.SEE} | "
                    f"Total: {res.Total} | Result: {res.Result}"
                )
                output_lines.append(line)

            self.text_results.delete("0.0", "end")
            self.text_results.insert("0.0", "\n".join(output_lines))

        except Exception as e:
            mb.showerror("Database Error", str(e))
        finally:
            session.close()




    def init_overall_stats_tab(self):
        tab = self.overall_stats_tab


        label_sem_filter = ctk.CTkLabel(tab, text="Filter Semester (optional):")
        label_sem_filter.grid(row=0, column=0, padx=(10,5), pady=(10,5), sticky="w")
        self.entry_filter_sem = ctk.CTkEntry(tab, placeholder_text="e.g. 3")
        self.entry_filter_sem.grid(row=0, column=1, padx=(0,5), pady=(10,5), sticky="ew")


        label_batch_filter = ctk.CTkLabel(tab, text="Filter Batch (optional):")
        label_batch_filter.grid(row=0, column=2, padx=(10,5), pady=(10,5), sticky="w")
        self.entry_filter_batch = ctk.CTkEntry(tab, placeholder_text="e.g. 23")
        self.entry_filter_batch.grid(row=0, column=3, padx=(0,10), pady=(10,5), sticky="ew")


        btn_search_overall = ctk.CTkButton(
            tab, text="Search", command=self.show_overall_statistics
        )
        btn_search_overall.grid(row=1, column=0, padx=10, pady=(5,5), sticky="w")

        btn_refresh_overall = ctk.CTkButton(
            tab, text="Refresh All", command=self.show_overall_statistics
        )
        btn_refresh_overall.grid(row=1, column=1, padx=10, pady=(5,5), sticky="w")


        self.text_overall = ctk.CTkTextbox(tab, width=550, height=300)
        self.text_overall.grid(row=2, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")


        tab.grid_columnconfigure(1, weight=1)
        tab.grid_columnconfigure(3, weight=1)
        tab.grid_rowconfigure(2, weight=1)


        self.show_overall_statistics()


    def show_overall_statistics(self):


        sem_text = self.entry_filter_sem.get().strip()
        batch_text = self.entry_filter_batch.get().strip()


        sem_filter = int(sem_text) if sem_text.isdigit() else None
        batch_filter = batch_text if batch_text else None

        session = Session()
        try:

            rows = (
                session.query(Result.USN, Result.Total, Result.Result, Subject.Semester)
                       .join(Subject, Result.SubjectCode == Subject.SubjectCode)
                       .all()
            )

            stats_dict = {}
            for usn, total, result_char, semester in rows:

                if sem_filter is not None and semester != sem_filter:
                    continue


                if len(usn) >= 5:
                    batch = usn[3:5]
                else:
                    batch = "NA"


                if batch_filter is not None and batch != batch_filter:
                    continue

                key = (semester, batch)
                if key not in stats_dict:
                    stats_dict[key] = {}
                if usn not in stats_dict[key]:
                    stats_dict[key][usn] = []
                stats_dict[key][usn].append((total or 0, result_char))


            if sem_filter is not None or batch_filter is not None:
                header = [f"Filtered Overall Stats", "-" * 70]
                if sem_filter is not None:
                    header.append(f"Semester: {sem_filter}")
                if batch_filter is not None:
                    header.append(f"Batch: {batch_filter}")
                header.append("-" * 70)
                lines = header[:]
            else:
                lines = ["Overall Semester & Batch Statistics", "-" * 70]


            for (semester, batch) in sorted(stats_dict.keys(), key=lambda x: (x[0] or 0, x[1])):
                student_map = stats_dict[(semester, batch)]
                total_students = len(student_map)
                passed_cnt = 0
                failed_cnt = 0
                sum_of_totals = 0
                for usn, records in student_map.items():
                    if all(rc[1] == "P" for rc in records):
                        passed_cnt += 1
                    else:
                        failed_cnt += 1
                line = (
                    f"Sem: {semester} | Batch: {batch} | Total Stud: {total_students} | "
                    f"Pass: {passed_cnt} | Fail: {failed_cnt}"
                )
                lines.append(line)


            self.text_overall.delete("0.0", "end")
            if not stats_dict:
                if sem_filter is not None or batch_filter is not None:
                    self.text_overall.insert("0.0", "No matching data found for given filters.")
                else:
                    self.text_overall.insert("0.0", "No result data in the database.")
            else:
                self.text_overall.insert("0.0", "\n".join(lines))

        except Exception as e:
            mb.showerror("Database Error", str(e))
        finally:
            session.close()








    def init_semester_stats_tab(self):
        tab = self.sem_stats_tab

        btn_refresh = ctk.CTkButton(
            tab, text="Refresh Semester Stats", command=self.show_semester_statistics
        )
        btn_refresh.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

        self.text_sem_stats = ctk.CTkTextbox(tab, width=550, height=300)
        self.text_sem_stats.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        self.show_semester_statistics()

    def show_semester_statistics(self):
        session = Session()
        try:
            rows = (
                session.query(Result.USN, Result.Total, Result.Result, Subject.Semester)
                .join(Subject, Result.SubjectCode == Subject.SubjectCode)
                .all()
            )
            sem_dict = {}
            for usn, total, result_char, semester in rows:
                if semester not in sem_dict:
                    sem_dict[semester] = {}
                if usn not in sem_dict[semester]:
                    sem_dict[semester][usn] = []
                sem_dict[semester][usn].append((total or 0, result_char))

            lines = ["Semester-Level Statistics", "-" * 60]
            for semester in sorted(sem_dict.keys(), key=lambda x: x or 0):
                student_map = sem_dict[semester]
                total_students = len(student_map)
                passed_cnt = 0
                failed_cnt = 0
                for usn, records in student_map.items():
                    if all(rc[1] == "P" for rc in records):
                        passed_cnt += 1
                    else:
                        failed_cnt += 1
                line = (
                    f"Semester: {semester} | Total Stud: {total_students} | "
                    f"Pass: {passed_cnt} | Fail: {failed_cnt}"
                )
                lines.append(line)

            self.text_sem_stats.delete("0.0", "end")
            if not sem_dict:
                self.text_sem_stats.insert("0.0", "No result data available.")
            else:
                self.text_sem_stats.insert("0.0", "\n".join(lines))

        except Exception as e:
            mb.showerror("Database Error", str(e))
        finally:
            session.close()








    def init_subject_stats_tab(self):
        tab = self.subj_stats_tab


        label_subj_filter = ctk.CTkLabel(tab, text="Filter Subject Code (optional):")
        label_subj_filter.grid(row=0, column=0, padx=(10,5), pady=(10,5), sticky="w")
        self.entry_filter_subj = ctk.CTkEntry(tab, placeholder_text="e.g. BCS301")
        self.entry_filter_subj.grid(row=0, column=1, padx=(0,5), pady=(10,5), sticky="ew")


        btn_search_subj = ctk.CTkButton(
            tab, text="Search", command=self.show_subject_statistics
        )
        btn_search_subj.grid(row=1, column=0, padx=10, pady=(5,5), sticky="w")

        btn_refresh_subj = ctk.CTkButton(
            tab, text="Refresh All", command=self.show_subject_statistics
        )
        btn_refresh_subj.grid(row=1, column=1, padx=10, pady=(5,5), sticky="w")


        self.text_subj_stats = ctk.CTkTextbox(tab, width=550, height=300)
        self.text_subj_stats.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(2, weight=1)


        self.show_subject_statistics()

    def show_subject_statistics(self):
        subj_filter = self.entry_filter_subj.get().strip().upper()

        session = Session()
        try:

            rows = session.query(Result.SubjectCode, Result.Total, Result.Result).all()

            subj_dict = {}
            for code, total, result_char in rows:

                if subj_filter and code != subj_filter:
                    continue

                if code not in subj_dict:
                    subj_dict[code] = {"count": 0, "passed": 0, "sum_tot": 0}
                if result_char !='NC':
                    subj_dict[code]["count"] += 1
                if result_char == "P":
                    subj_dict[code]["passed"] += 1
                subj_dict[code]["sum_tot"] += (total or 0)


            if subj_filter:
                header = [f"Stats for Subject: {subj_filter}", "-" * 60]
                lines = header[:]
            else:
                lines = ["Subject-Level Statistics", "-" * 60]

            for code in sorted(subj_dict.keys()):
                data = subj_dict[code]
                cnt = data["count"]
                passed = data["passed"]
                failed = cnt - passed
                pass_pct = (passed / cnt * 100) if cnt else 0.0
                line = (
                    f"Subject: {code} | Total Stud: {cnt} | Pass: {passed} | "
                    f"Fail: {failed} | Pass%: {pass_pct:.2f}%"
                )
                lines.append(line)

            self.text_subj_stats.delete("0.0", "end")
            if not subj_dict:
                if subj_filter:
                    self.text_subj_stats.insert("0.0", f"No data found for subject {subj_filter}.")
                else:
                    self.text_subj_stats.insert("0.0", "No subject data available.")
            else:
                self.text_subj_stats.insert("0.0", "\n".join(lines))

        except Exception as e:
            mb.showerror("Database Error", str(e))
        finally:
            session.close()







    def init_student_grades_tab(self):
            tab = self.stu_grades_tab


            label_grade_usn = ctk.CTkLabel(tab, text="Filter by USN (optional):")
            label_grade_usn.grid(row=0, column=0, padx=(10,5), pady=(10,5), sticky="w")
            self.entry_grade_usn = ctk.CTkEntry(tab, placeholder_text="e.g. 1AB12CS001")
            self.entry_grade_usn.grid(row=0, column=1, padx=(0,5), pady=(10,5), sticky="ew")
            btn_search_grade = ctk.CTkButton(
                tab, text="Search", command=self.show_student_grades
            )
            btn_search_grade.grid(row=0, column=2, padx=(5,10), pady=(10,5), sticky="w")


            btn_refresh = ctk.CTkButton(
                tab, text="Refresh All Grades", command=self.show_student_grades
            )
            btn_refresh.grid(row=1, column=0, columnspan=3, padx=10, pady=(5,10), sticky="w")


            self.text_stu_grades = ctk.CTkTextbox(tab, width=550, height=300)
            self.text_stu_grades.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

            tab.grid_columnconfigure(1, weight=1)
            tab.grid_rowconfigure(2, weight=1)


            self.show_student_grades()


    def show_student_grades(self):
            usn_filter = self.entry_grade_usn.get().strip().upper()

            session = Session()
            try:

                rows = (
                    session.query(Result.USN, Result.Total, Subject.Semester)
                           .join(Subject, Result.SubjectCode == Subject.SubjectCode)
                           .all()
                )




                data = {}
                for usn, total, semester in rows:

                    if usn_filter and usn != usn_filter:
                        continue

                    key = (usn, semester)
                    if key not in data:
                        data[key] = {"sum_tot": 0, "count_subj": 0}
                    data[key]["sum_tot"] += (total or 0)
                    data[key]["count_subj"] += 1


                if usn_filter:
                    header = [f"Grades for USN: {usn_filter}", "-" * 60]
                else:
                    header = ["All Students’ Semester‐Wise Grades", "-" * 60]

                lines = header[:]

                for (usn, semester) in sorted(data.keys(), key=lambda x: (x[0], x[1] or 0)):
                    entry = data[(usn, semester)]
                    total_marks = entry["sum_tot"]
                    num_subj = entry["count_subj"]

                    percent = (total_marks / (num_subj * 100) * 100) if num_subj else 0.0

                    if percent >= 70:
                        grade = "FCD"
                    elif percent >= 60:
                        grade = "FC"
                    elif percent >= 50:
                        grade = "SC"
                    else:
                        grade = "F"


                    line = (
                        f"USN: {usn} | Sem: {semester} | Subjects: {num_subj} | "
                        f"Total: {total_marks} | %: {percent:.2f}% | Grade: {grade}"
                    )
                    lines.append(line)


                self.text_stu_grades.delete("0.0", "end")
                if not data:
                    if usn_filter:
                        self.text_stu_grades.insert("0.0", f"No data found for USN {usn_filter}.")
                    else:
                        self.text_stu_grades.insert("0.0", "No student result data available.")
                else:
                    self.text_stu_grades.insert("0.0", "\n".join(lines))

            except Exception as e:
                mb.showerror("Database Error", str(e))
            finally:
                session.close()




    def extract_segments(self, text):
        pattern = re.compile(
            r"^(\d)([A-Z]{2})(\d{2})([A-Z]{2})(\d{3})$", re.IGNORECASE
        )
        match = pattern.match(text)
        if match:
            try:
                return (
                    match.group(1),
                    match.group(2),
                    int(match.group(3)),
                    match.group(4),
                    int(match.group(5)),
                )
            except:
                return None
        return None



    def cancel(self):
        if self.cancel_event and not self.cancel_event.is_set():
            self.cancel_event.set()
        self.progress_win.destroy()
        
    def on_closing(self):
        if self.cancel_event and not self.cancel_event.is_set():
            self.cancel_event.set()
        self.destroy()


if __name__ == "__main__":
    init_db()
    app = App()
    app.mainloop()
