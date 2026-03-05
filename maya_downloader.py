# -*- coding: utf-8 -*-
import sys
import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
import requests

# ── Hebrew strings ────────────────────────────────────────────────────────────
T = {
    "title":           "הורדת דוחות שנתיים – מאיה",
    "company_label":   "הזן שם חברה:",
    "year_label":      "בחר שנת דוח שנתי (2015-2025):",
    "search_btn":      "חפש",
    "no_company":      "לא נמצאה חברה מתאימה.\nנסה לכתוב שם אחר.",
    "choose_company":  "נמצאו מספר חברות מתאימות.\nבחר את החברה המתאימה:",
    "no_report":       "לא נמצא דוח שנתי לשנה זו.",
    "choose_report":   "נמצאו מספר דוחות לשנה זו.\nבחר את הדוח הרצוי:",
    "confirm_btn":     "אישור",
    "cancel_btn":      "ביטול",
    "success_title":   "✔  הדוח הורד בהצלחה",
    "success_body":    "חברה: {company}\nשנה: {year}\nתאריך פרסום: {date}\n\nהקובץ נשמר כאן:\n{path}",
    "error_title":     "שגיאה",
    "pdf_fail":        "לא הצלחתי להוריד את קובץ ה-PDF.\nייתכן שהדוח אינו זמין.",
    "conn_error":      "שגיאת חיבור לאינטרנט:\n{err}",
    "save_path_q":     "הכונן E:\\ לא נמצא.\nהזן נתיב שמירה אחר (לדוגמה: C:\\MayaReports):",
    "bad_path":        "הנתיב שהוזן אינו תקין.\nאנא הזן נתיב כגון C:\\MayaReports",
    "status_ready":    "מוכן לחיפוש",
    "status_search":   "מחפש חברה...",
    "status_reports":  "מחפש דוחות שנתיים...",
    "status_pdf":      "מוריד קובץ PDF...",
    "status_save":     "שומר קובץ...",
}

HEADERS = {
    "Accept":           "application/json, text/plain, */*",
    "Accept-Language":  "he-IL,he;q=0.9,en-US;q=0.8,en;q=0.7",
    "Origin":           "https://maya.tase.co.il",
    "Referer":          "https://maya.tase.co.il/",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


# ── API helpers ───────────────────────────────────────────────────────────────

def search_companies(query: str) -> list:
    url = (
        "https://mayaapi.tase.co.il/api/api/company/companysearch"
        f"?query={requests.utils.quote(query)}&lang=1"
    )
    r = SESSION.get(url, timeout=15)
    r.raise_for_status()
    data = r.json()
    if isinstance(data, list):
        return data
    for key in ("CompanyList", "companies", "Companies", "Results", "results"):
        if key in data:
            return data[key]
    return []


def fetch_annual_reports(company_id, year: int) -> list:
    url = (
        "https://mayaapi.tase.co.il/api/report/companyrep"
        f"?companyId={company_id}&reportType=0"
        f"&pageNumber=1&pageSize=100&lang=1"
    )
    r = SESSION.get(url, timeout=15)
    r.raise_for_status()
    data = r.json()

    raw = []
    if isinstance(data, list):
        raw = data
    else:
        for key in ("Reports", "reports", "ReportList", "Items", "items", "Data"):
            if key in data and isinstance(data[key], list):
                raw = data[key]
                break

    annual = []
    for rep in raw:
        title = _field(rep, "ReportName", "reportName", "Title", "title") or ""
        pub   = _field(rep, "PubDate", "pubDate", "PublishDate", "publishDate") or ""
        if "תקופתי" not in title:
            continue
        if str(year) not in str(pub):
            continue
        annual.append(rep)
    return annual


def _field(d: dict, *keys):
    for k in keys:
        v = d.get(k)
        if v is not None:
            return v
    return None


def download_pdf(report_id) -> bytes | None:
    rid = int(report_id)
    start = (rid // 1000) * 1000 + 1
    end   = start + 999
    for suffix in ("00", "01", "02"):
        url = f"https://mayafiles.tase.co.il/rpdf/{start}-{end}/P{rid}-{suffix}.pdf"
        try:
            r = SESSION.get(url, timeout=30, stream=True)
            if r.status_code == 200:
                data = r.content
                if data[:4] == b"%PDF":
                    return data
        except Exception:
            continue
    return None


def sanitize(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()


def fmt_date(raw) -> str:
    s = str(raw or "")[:10]
    s = s.replace("-", ".").replace("/", ".")
    parts = s.split(".")
    if len(parts) == 3 and len(parts[0]) == 4:
        return f"{parts[2]}.{parts[1]}.{parts[0]}"
    return s


def get_save_root() -> str | None:
    if os.path.isdir("E:\\"):
        return "E:\\MayaReports"
    return None


# ── GUI ───────────────────────────────────────────────────────────────────────

BG        = "#f0f4f8"
ACCENT    = "#1a5276"
ACCENT_LT = "#d6eaf8"
FG_DARK   = "#154360"
FONT_MAIN = ("Segoe UI", 11)
FONT_BOLD = ("Segoe UI", 11, "bold")
FONT_BIG  = ("Segoe UI", 13, "bold")
FONT_HDR  = ("Segoe UI", 14, "bold")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(T["title"])
        self.resizable(False, False)
        self.configure(bg=BG)
        self._build()
        self._center()

    def _center(self):
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build(self):
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True)

        # Header
        hdr = tk.Frame(outer, bg=ACCENT, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text=T["title"], bg=ACCENT, fg="white",
                 font=FONT_HDR).pack()

        # Body
        body = tk.Frame(outer, bg=BG, padx=28, pady=20)
        body.pack(fill="both", expand=True)

        # Company row
        tk.Label(body, text=T["company_label"], bg=BG,
                 font=FONT_BOLD, anchor="e").pack(fill="x")
        self.company_var = tk.StringVar()
        ent = tk.Entry(body, textvariable=self.company_var,
                       font=("Segoe UI", 12), justify="right",
                       relief="solid", bd=1)
        ent.pack(fill="x", ipady=4, pady=(2, 12))
        ent.bind("<Return>", lambda _: self._run_search())
        ent.focus()

        # Year row
        tk.Label(body, text=T["year_label"], bg=BG,
                 font=FONT_BOLD, anchor="e").pack(fill="x")
        self.year_var = tk.StringVar(value="2024")
        years = [str(y) for y in range(2025, 2014, -1)]
        cb = ttk.Combobox(body, textvariable=self.year_var,
                          values=years, state="readonly",
                          font=FONT_MAIN, justify="right", width=12)
        cb.pack(anchor="e", pady=(2, 18))

        # Search button
        self.btn = tk.Button(
            body, text=T["search_btn"],
            bg=ACCENT, fg="white", font=FONT_BIG,
            relief="flat", cursor="hand2",
            padx=28, pady=7,
            command=self._run_search,
            activebackground="#154360", activeforeground="white",
        )
        self.btn.pack(anchor="center")

        # Status bar
        self.status_var = tk.StringVar(value=T["status_ready"])
        tk.Label(outer, textvariable=self.status_var,
                 bg=ACCENT_LT, fg=FG_DARK, font=("Segoe UI", 9),
                 anchor="center", relief="sunken", pady=4
                 ).pack(fill="x", side="bottom")

    def _status(self, msg):
        self.status_var.set(msg)
        self.update()

    def _lock(self, locked: bool):
        self.btn.config(state="disabled" if locked else "normal")

    # ── search flow (runs in thread to keep UI alive) ──────────────────────

    def _run_search(self):
        query = self.company_var.get().strip()
        year  = self.year_var.get().strip()
        if not query:
            messagebox.showwarning(T["error_title"], "אנא הזן שם חברה.")
            return
        self._lock(True)
        threading.Thread(target=self._search_thread,
                         args=(query, int(year)), daemon=True).start()

    def _search_thread(self, query, year):
        try:
            self._status(T["status_search"])
            try:
                companies = search_companies(query)
            except Exception as e:
                self._err(T["conn_error"].format(err=e))
                return

            if not companies:
                self._info(T["error_title"], T["no_company"])
                return

            company = companies[0] if len(companies) == 1 \
                else self._pick(T["choose_company"], companies,
                                lambda c: _field(c, "CompanyName", "companyName",
                                                 "Name", "name") or "?")
            if company is None:
                return

            cid   = _field(company, "CompanyId", "companyId", "Id", "id")
            cname = _field(company, "CompanyName", "companyName",
                           "Name", "name") or query

            self._status(T["status_reports"])
            try:
                reports = fetch_annual_reports(cid, year)
            except Exception as e:
                self._err(T["conn_error"].format(err=e))
                return

            if not reports:
                self._info(T["error_title"], T["no_report"])
                return

            def rep_label(r):
                t = _field(r, "ReportName", "reportName", "Title", "title") or "דוח"
                d = fmt_date(_field(r, "PubDate", "pubDate",
                                    "PublishDate", "publishDate"))
                return f"{t} – {d}" if d else t

            report = reports[0] if len(reports) == 1 \
                else self._pick(T["choose_report"], reports, rep_label)
            if report is None:
                return

            rid      = _field(report, "ReportId", "reportId", "Id", "id")
            pub_date = fmt_date(_field(report, "PubDate", "pubDate",
                                       "PublishDate", "publishDate"))

            self._status(T["status_pdf"])
            pdf = download_pdf(rid)
            if not pdf:
                self._err(T["pdf_fail"])
                return

            self._status(T["status_save"])
            save_root = get_save_root()
            if save_root is None:
                save_root = self._ask_path()
                if not save_root:
                    return

            safe = sanitize(cname)
            folder   = os.path.join(save_root, safe)
            os.makedirs(folder, exist_ok=True)
            filepath = os.path.join(folder, f"{safe} - {year}.pdf")
            with open(filepath, "wb") as f:
                f.write(pdf)

            self._info(
                T["success_title"],
                T["success_body"].format(
                    company=cname, year=year,
                    date=pub_date, path=filepath,
                )
            )
            self._status(T["status_ready"])

        finally:
            self.after(0, lambda: self._lock(False))
            self._status(T["status_ready"])

    # ── dialog helpers (must run on main thread) ───────────────────────────

    def _info(self, title, msg):
        self.after(0, lambda: messagebox.showinfo(title, msg, parent=self))

    def _err(self, msg):
        self.after(0, lambda: messagebox.showerror(T["error_title"], msg, parent=self))
        self._status(T["status_ready"])

    def _pick(self, prompt, items, label_fn) -> object | None:
        result = [None]
        done   = threading.Event()

        def _show():
            dlg = _PickDialog(self, prompt, items, label_fn)
            result[0] = dlg.result
            done.set()

        self.after(0, _show)
        done.wait()
        return result[0]

    def _ask_path(self) -> str | None:
        result = [None]
        done   = threading.Event()

        def _show():
            while True:
                path = simpledialog.askstring(
                    T["error_title"], T["save_path_q"],
                    parent=self, initialvalue="C:\\MayaReports"
                )
                if path is None:
                    break
                path = path.strip()
                if re.match(r"^[A-Za-z]:\\", path):
                    result[0] = path
                    break
                messagebox.showerror(T["error_title"], T["bad_path"], parent=self)
            done.set()

        self.after(0, _show)
        done.wait()
        return result[0]


class _PickDialog(tk.Toplevel):
    def __init__(self, parent, prompt, items, label_fn):
        super().__init__(parent)
        self.result = None
        self.title(T["title"])
        self.configure(bg=BG)
        self.grab_set()
        self.resizable(False, False)

        tk.Label(self, text=prompt, bg=BG, font=FONT_BOLD,
                 wraplength=420, justify="right", pady=6).pack(
            fill="x", padx=16, pady=(14, 2))

        frm = tk.Frame(self, bg=BG)
        frm.pack(fill="both", expand=True, padx=16, pady=4)

        sb = tk.Scrollbar(frm, orient="vertical")
        self._lb = tk.Listbox(
            frm, font=FONT_MAIN, justify="right",
            yscrollcommand=sb.set,
            selectbackground=ACCENT, selectforeground="white",
            height=min(len(items), 10), width=54,
            activestyle="none",
        )
        sb.config(command=self._lb.yview)
        sb.pack(side="left", fill="y")
        self._lb.pack(side="right", fill="both", expand=True)

        for idx, item in enumerate(items, 1):
            self._lb.insert("end", f"{idx}. {label_fn(item)}")
        self._lb.selection_set(0)
        self._lb.bind("<Double-Button-1>", lambda _: self._confirm())

        btn_row = tk.Frame(self, bg=BG)
        btn_row.pack(pady=10)
        tk.Button(btn_row, text=T["confirm_btn"],
                  bg=ACCENT, fg="white", font=FONT_BOLD,
                  relief="flat", cursor="hand2",
                  padx=18, pady=5,
                  command=self._confirm).pack(side="right", padx=6)
        tk.Button(btn_row, text=T["cancel_btn"],
                  bg="#aab7b8", fg="white", font=FONT_MAIN,
                  relief="flat", cursor="hand2",
                  padx=18, pady=5,
                  command=self.destroy).pack(side="right", padx=6)

        self._items = items
        self._center(parent)
        self.wait_window()

    def _center(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width()  // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{pw - w//2}+{ph - h//2}")

    def _confirm(self):
        sel = self._lb.curselection()
        if sel:
            self.result = self._items[sel[0]]
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
