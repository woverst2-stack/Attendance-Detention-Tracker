"""
Attendance Detention Tracker -- Desktop App
============================================

A windowed front end for attendance.py. All the actual logic (parsing the
report, counting occurrences, updating the workbook) lives in attendance.py
and is reused here unchanged -- this file only adds the window, buttons, and
text boxes.
"""

import datetime as dt
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import attendance as core


class AttendanceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Detention Tracker")
        self.geometry("760x680")
        self.minsize(600, 500)

        self._build_widgets()

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------
    def _build_widgets(self):
        pad = {"padx": 8, "pady": 4}

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        ttk.Label(top, text="Workbook:").grid(row=0, column=0, sticky="w")
        self.workbook_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.workbook_var).grid(
            row=0, column=1, sticky="ew", padx=4
        )
        ttk.Button(top, text="Browse...", command=self._browse_workbook).grid(
            row=0, column=2
        )
        ttk.Button(top, text="New Template...", command=self._new_template).grid(
            row=0, column=3, padx=(4, 0)
        )
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Date (YYYY-MM-DD, blank = today):").grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )
        self.date_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.date_var, width=16).grid(
            row=1, column=1, sticky="w", padx=4, pady=(6, 0)
        )

        ttk.Label(self, text="Paste the attendance report below:").pack(
            anchor="w", **pad
        )
        self.report_box = scrolledtext.ScrolledText(self, height=16, font=("Consolas", 10))
        self.report_box.pack(fill="both", expand=True, padx=8)

        options = ttk.Frame(self)
        options.pack(fill="x", **pad)
        self.dry_run_var = tk.BooleanVar()
        self.force_var = tk.BooleanVar()
        ttk.Checkbutton(options, text="Dry run (show changes, don't save)",
                         variable=self.dry_run_var).pack(side="left")
        ttk.Checkbutton(options, text="Force (allow re-recording this date)",
                         variable=self.force_var).pack(side="left", padx=(16, 0))
        ttk.Button(options, text="Run", command=self._run).pack(side="right")

        weekly = ttk.LabelFrame(self, text="Weekly Detention Report")
        weekly.pack(fill="x", padx=8, pady=(0, 4))

        self._pending_report_through = None
        self.weekly_status_var = tk.StringVar(
            value="Click 'Generate Weekly Report' to see who's earned a detention "
                  "since the last time this was sent."
        )
        ttk.Label(weekly, textvariable=self.weekly_status_var, wraplength=700,
                  justify="left").pack(anchor="w", padx=6, pady=(4, 2))

        weekly_buttons = ttk.Frame(weekly)
        weekly_buttons.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Button(weekly_buttons, text="Generate Weekly Report",
                   command=self._generate_weekly_report).pack(side="left")
        self.mark_sent_button = ttk.Button(
            weekly_buttons, text="Mark as Sent", command=self._mark_report_sent,
            state="disabled",
        )
        self.mark_sent_button.pack(side="left", padx=(8, 0))

        ttk.Label(self, text="Output:").pack(anchor="w", **pad)
        self.output_box = scrolledtext.ScrolledText(
            self, height=14, font=("Consolas", 10), state="disabled", background="#f5f5f5"
        )
        self.output_box.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _browse_workbook(self):
        path = filedialog.askopenfilename(
            title="Choose the Excel workbook",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.workbook_var.set(path)

    def _new_template(self):
        path = filedialog.asksaveasfilename(
            title="Save new workbook as",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return

        try:
            workbook = core.create_template_workbook()
            workbook.save(path)
        except PermissionError:
            messagebox.showerror(
                "Could not save",
                f"'{path}' is open in another program (probably Excel).\n"
                "Close it there and try again.",
            )
            return
        except Exception as exc:
            messagebox.showerror("Could not create workbook", str(exc))
            return

        self.workbook_var.set(path)
        messagebox.showinfo("Template created", f"New workbook created:\n{path}")

    def _generate_weekly_report(self):
        workbook_path = self.workbook_var.get().strip().strip('"')
        if not workbook_path:
            messagebox.showerror("Missing workbook", "Choose an Excel workbook first.")
            return

        try:
            workbook = core.load_workbook(workbook_path)
        except FileNotFoundError:
            messagebox.showerror("File not found", f"Could not find:\n{workbook_path}")
            return
        except Exception as exc:
            messagebox.showerror("Could not open workbook", str(exc))
            return

        tracked_codes = core.read_settings(workbook)
        log_entries = core.read_log(workbook)
        last_sent = core.read_last_report_date(workbook)
        through_date = dt.date.today()

        all_detentions = core.calculate_detentions(log_entries, tracked_codes)
        new_detentions = core.detentions_since(all_detentions, last_sent, through_date)

        self._clear_log()
        if last_sent:
            self._log(f"Detentions earned after {last_sent}, through {through_date}:")
        else:
            self._log(f"No weekly report has been sent before -- showing everything through {through_date}:")
        self._log()

        if new_detentions:
            for d in sorted(new_detentions, key=lambda d: (d["date"], d["name"])):
                self._log(f"  {d['date']}   {d['name']:<28} grade {d['grade']:<3} "
                           f"code {d['code']}   detention #{d['number']}")
        else:
            self._log("  Nobody has earned a new detention since the last report.")

        self._pending_report_through = through_date
        self.mark_sent_button.configure(state="normal")
        self.weekly_status_var.set(
            f"Showing {len(new_detentions)} detention(s) as of {through_date}. "
            f"Copy this into the email, then click 'Mark as Sent'."
        )

    def _mark_report_sent(self):
        if self._pending_report_through is None:
            return

        workbook_path = self.workbook_var.get().strip().strip('"')
        through_date = self._pending_report_through
        try:
            workbook = core.load_workbook(workbook_path)
            tracked_codes = core.read_settings(workbook)
            log_entries = core.read_log(workbook)
            last_sent = core.read_last_report_date(workbook)
            all_detentions = core.calculate_detentions(log_entries, tracked_codes)
            new_detentions = core.detentions_since(all_detentions, last_sent, through_date)

            core.record_weekly_report(workbook, dt.date.today(), last_sent, through_date, new_detentions)
            core.write_last_report_date(workbook, through_date)
            workbook.save(workbook_path)
        except PermissionError:
            messagebox.showerror(
                "Could not save",
                f"'{workbook_path}' is open in another program (probably Excel).\n"
                "Close it there and try again.",
            )
            return
        except Exception as exc:
            messagebox.showerror("Could not save", str(exc))
            return

        self._log(f"\nMarked as sent through {through_date} and saved to the "
                   f"'Weekly Reports' sheet. Next week's report will start from here.")
        self.weekly_status_var.set(f"Last report sent through {through_date}.")
        self.mark_sent_button.configure(state="disabled")
        self._pending_report_through = None

    def _log(self, message=""):
        self.output_box.configure(state="normal")
        self.output_box.insert("end", message + "\n")
        self.output_box.configure(state="disabled")
        self.output_box.see("end")
        self.update_idletasks()

    def _clear_log(self):
        self.output_box.configure(state="normal")
        self.output_box.delete("1.0", "end")
        self.output_box.configure(state="disabled")

    # ------------------------------------------------------------------
    # The actual work -- mirrors attendance.py's main(), but reports to
    # the on-screen log instead of the terminal.
    # ------------------------------------------------------------------
    def _run(self):
        self._clear_log()

        workbook_path = self.workbook_var.get().strip().strip('"')
        if not workbook_path:
            messagebox.showerror("Missing workbook", "Choose an Excel workbook first.")
            return

        date_text = self.date_var.get().strip()
        if date_text:
            try:
                report_date = dt.datetime.strptime(date_text, "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Bad date", "Date must be in YYYY-MM-DD format.")
                return
        else:
            report_date = dt.date.today()

        report_text = self.report_box.get("1.0", "end")
        if not report_text.strip():
            messagebox.showerror("Nothing pasted", "Paste the attendance report first.")
            return

        try:
            workbook = core.load_workbook(workbook_path)
        except FileNotFoundError:
            messagebox.showerror("File not found", f"Could not find:\n{workbook_path}")
            return
        except Exception as exc:
            messagebox.showerror("Could not open workbook", str(exc))
            return

        tracked_codes = core.read_settings(workbook)
        existing_log = core.read_log(workbook)
        last_sent = core.read_last_report_date(workbook)

        self._log(f"Report date:   {report_date}")
        self._log("Codes counted: " + ", ".join(
            f"{c} (every {n})" for c, n in tracked_codes.items()
        ))
        self._log()

        occurrences, unparsed = core.parse_pasted_report(report_text)
        if unparsed:
            self._log(f"!! {len(unparsed)} line(s) could not be read and were skipped:")
            for line in unparsed[:10]:
                self._log(f"     {line}")
            self._log()

        tracked = core.keep_only_tracked(occurrences, tracked_codes)
        if core.COUNT_MODE == "per_day":
            tracked = core.collapse_to_one_per_day(tracked)

        self._log(f"Lines read:            {len(occurrences)} coded entries")
        self._log(f"Counting toward rule:  {len(tracked)}")

        if not tracked:
            self._log("\nNothing to record. No changes made.")
            return

        for code in tracked_codes:
            duplicates = core.already_logged(existing_log, report_date, code)
            if duplicates and not self.force_var.get():
                self._log(
                    f"\n!! STOPPING: {duplicates} '{code}' entries are already recorded "
                    f"for {report_date}."
                )
                self._log("   Recording again would issue detentions that were not earned.")
                self._log("   Tick 'Force' if you are certain you want to add these anyway.")
                return

        known = {e["name"] for e in existing_log}
        incoming = {o.name for o in tracked}
        brand_new = {n for n in incoming if core.name_key(n) not in {core.name_key(k) for k in known}}
        for name, looks_like in core.find_near_matches(brand_new, known):
            self._log(f"!! '{name}' is very close to the existing '{looks_like}'. Check for a typo.")

        if brand_new:
            self._log("\nNew students added to the roster:")
            for name in sorted(brand_new):
                self._log(f"    {name}")

        core.append_to_log(workbook, tracked, report_date)
        updated_log = existing_log + [
            {"date": report_date, "name": o.name, "grade": o.grade,
             "code": o.code, "period": o.period, "raw_code": o.raw_code}
            for o in tracked
        ]
        core.rebuild_derived_sheets(workbook, updated_log, tracked_codes)

        all_detentions = core.calculate_detentions(updated_log, tracked_codes)
        new_detentions = [d for d in all_detentions if d["date"] == report_date]

        self._log("\n" + "=" * 58)
        if new_detentions:
            self._log(f"DETENTIONS EARNED ON {report_date}")
            self._log("=" * 58)
            for d in new_detentions:
                self._log(f"  {d['name']:<28} grade {d['grade']:<3} "
                           f"code {d['code']}   detention #{d['number']}")
        else:
            self._log("No detentions earned today.")
        self._log("=" * 58)

        since_last_report = core.detentions_since(all_detentions, last_sent, report_date)
        self._log()
        if last_sent:
            self._log(f"Since last weekly report (after {last_sent}, through {report_date}):")
        else:
            self._log(f"Since last weekly report (none sent yet, through {report_date}):")
        if since_last_report:
            for d in sorted(since_last_report, key=lambda d: (d["date"], d["name"])):
                self._log(f"  {d['date']}   {d['name']:<28} grade {d['grade']:<3} "
                           f"code {d['code']}   detention #{d['number']}")
        else:
            self._log("  None yet.")

        days = core.days_since_last_report(updated_log, last_sent, report_date)
        if days is not None and days > 7:
            if last_sent:
                self._log(f"\n!! It has been {days} days since the weekly report was last sent "
                           f"(last sent through {last_sent}). Consider sending it soon.")
            else:
                self._log(f"\n!! It has been {days} days since attendance started being logged, "
                           f"and no weekly report has ever been sent. Consider sending one soon.")

        if self.dry_run_var.get():
            self._log("\nDry run -- the workbook was NOT saved.")
            return

        try:
            workbook.save(workbook_path)
        except PermissionError:
            self._log(
                f"\n!! COULD NOT SAVE: '{workbook_path}' is open in another program "
                f"(probably Excel)."
            )
            self._log("   Close the file there and click Run again -- nothing has been lost.")
            messagebox.showerror(
                "Workbook is open",
                "Close the workbook in Excel, then click Run again.\n"
                "Nothing has been lost.",
            )
            return

        self._log(f"\nSaved to {workbook_path}")


def main():
    app = AttendanceApp()
    app.mainloop()


if __name__ == "__main__":
    main()
