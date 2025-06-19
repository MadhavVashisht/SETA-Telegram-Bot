import os
import sys
import json
import threading
import time
import requests
import schedule
import webbrowser
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from subprocess import run, PIPE
from PIL import Image, ImageChops
from datetime import datetime

# ---------------- Resource Path Fix ----------------
def resource_path(filename):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

CONFIG_FILE = resource_path('seta_config.json')
SCHEDULES_FILE = resource_path('separate_schedules.json')
RECIPIENTS_FILE = resource_path('recipients.json')

# ---------------- Core Utilities ----------------

def convert_excel_to_pdf(excel_path, pdf_path):
    cmd = ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(pdf_path), excel_path]
    result = run(cmd, stdout=PIPE, stderr=PIPE)
    return os.path.exists(pdf_path), result.stderr.decode() if result.returncode != 0 else ""

def crop_image_whitespace(image_path):
    img = Image.open(image_path).convert("RGB")
    bg = Image.new("RGB", img.size, img.getpixel((0, 0)))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    if bbox:
        img.crop(bbox).save(image_path)
        return True
    return False

def convert_pdf_to_png(pdf_path, png_path):
    base = os.path.splitext(png_path)[0]
    cmd = ['pdftoppm', '-png', '-singlefile', pdf_path, base]
    result = run(cmd, stdout=PIPE, stderr=PIPE)
    crop_image_whitespace(png_path)
    return os.path.exists(png_path), result.stderr.decode() if result.returncode != 0 else ""

def send_image(bot_token, chat_id, image_path):
    url = f'https://api.telegram.org/bot{bot_token}/sendPhoto'
    with open(image_path, 'rb') as img:
        files = {'photo': img}
        data = {'chat_id': chat_id}
        response = requests.post(url, files=files, data=data)
    return response.ok, response.text

# ---------------- GUI Application ----------------

class SETASchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SETA Scheduler - Dashboard")
        self.root.geometry("850x600")
        self.style = ttk.Style("flatly")

        self.recipients = self.load_recipients()
        self.schedules = self.load_schedules()
        self.bot_token = ""

        self.build_ui()
        self.run_scheduler()

    def load_schedules(self):
        return json.load(open(SCHEDULES_FILE)) if os.path.exists(SCHEDULES_FILE) else []

    def save_schedules(self):
        json.dump(self.schedules, open(SCHEDULES_FILE, 'w'), indent=2)

    def load_recipients(self):
        if os.path.exists(RECIPIENTS_FILE):
            with open(RECIPIENTS_FILE, 'r') as f:
                data = json.load(f)
                return data['users'] if isinstance(data, dict) else data
        return []

    def save_recipients(self):
        json.dump({'users': self.recipients}, open(RECIPIENTS_FILE, 'w'), indent=2)

    def build_ui(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill='both', expand=True)

        ttk.Label(frm, text="SETA Scheduler", font=("Helvetica", 20, "bold")).pack(pady=10)

        token_frame = ttk.Frame(frm)
        token_frame.pack(fill='x', pady=5)
        ttk.Label(token_frame, text="Bot Token:").pack(side='left')
        self.token_var = ttk.StringVar()
        ttk.Entry(token_frame, textvariable=self.token_var, width=50).pack(side='left', padx=5)

        btns = ttk.Frame(frm)
        btns.pack(pady=10)
        ttk.Button(btns, text="Add Schedule", command=self.add_file_schedule, bootstyle=SUCCESS).pack(side='left', padx=5)
        ttk.Button(btns, text="Manage Recipients", command=self.manage_recipients, bootstyle=INFO).pack(side='left', padx=5)
        ttk.Button(btns, text="View Schedules", command=self.view_schedule, bootstyle=PRIMARY).pack(side='left', padx=5)

        self.log_text = tk.Text(frm, height=18, bg="white", fg="black")
        self.log_text.pack(fill='both', expand=True)

        ttk.Button(frm, text="Designed by Madhav", command=lambda: webbrowser.open("https://madhavvashisht.unaux.com/"), bootstyle=LINK).pack(pady=5)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert('end', message + "\n")
        self.log_text.config(state='disabled')
        self.log_text.see('end')

    def manage_recipients(self):
        win = tk.Toplevel(self.root)
        win.title("Manage Recipients")
        lb = tk.Listbox(win)
        lb.pack(fill='both', expand=True)

        for u in self.recipients:
            lb.insert('end', f"{u['name']} ({u['chat_id']})")

        def add():
            name = simpledialog.askstring("Name", "Recipient Name:", parent=win)
            chat = simpledialog.askstring("Chat ID", "Telegram Chat ID:", parent=win)
            if name and chat:
                self.recipients.append({"name": name, "chat_id": chat})
                self.save_recipients()
                lb.insert('end', f"{name} ({chat})")
                self.log(f"Added recipient: {name}")

        def remove():
            sel = lb.curselection()
            if sel:
                idx = sel[0]
                removed = self.recipients.pop(idx)
                self.save_recipients()
                lb.delete(idx)
                self.log(f"Removed recipient: {removed['name']}")

        ttk.Button(win, text="Add", command=add, bootstyle=SUCCESS).pack(fill='x')
        ttk.Button(win, text="Remove", command=remove, bootstyle=DANGER).pack(fill='x')

    def add_file_schedule(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
        if not files:
            return

        win = tk.Toplevel(self.root)
        win.title("Schedule Settings")

        ttk.Label(win, text="Choose Recipients:").pack()
        rec_vars = []
        for r in self.recipients:
            v = tk.BooleanVar()
            ttk.Checkbutton(win, text=f"{r['name']} ({r['chat_id']})", variable=v).pack(anchor='w')
            rec_vars.append((r, v))

        ttk.Label(win, text="Frequency:").pack()
        freq_var = tk.StringVar(value="daily")
        ttk.Combobox(win, textvariable=freq_var, values=["daily", "hourly", "once"], state='readonly').pack()

        time_var = tk.StringVar(value="13:00")
        ttk.Label(win, text="Time (for daily/once):").pack()
        ttk.Entry(win, textvariable=time_var).pack()

        def save():
            selected = [r for r, v in rec_vars if v.get()]
            if not selected:
                messagebox.showerror("Error", "Select at least one recipient")
                return

            for file in files:
                for r in selected:
                    sched = {
                        "file": file,
                        "chat_id": r['chat_id'],
                        "name": r['name'],
                        "frequency": freq_var.get(),
                        "time": time_var.get() if freq_var.get() != "hourly" else None,
                        "sent": False
                    }
                    self.schedules.append(sched)
                    self.setup_schedule(sched)
                    self.log(f"Scheduled {os.path.basename(file)} to {r['name']} ({sched['frequency']})")

            self.save_schedules()
            win.destroy()

        ttk.Button(win, text="Save", command=save, bootstyle=PRIMARY).pack(pady=5)

    def setup_schedule(self, job):
        def send():
            if job['frequency'] == 'hourly' and not (9 <= datetime.now().hour <= 18):
                return
            self.send_file(job)

        if job['frequency'] == 'hourly':
            schedule.every().hour.at(":00").do(send)
        elif job['frequency'] == 'daily':
            schedule.every().day.at(job['time']).do(send)
        elif job['frequency'] == 'once' and not job.get('sent'):
            schedule.every().day.at(job['time']).do(send)

    def send_file(self, job):
        file = job['file']
        pdf = file.replace('.xlsx', '.pdf')
        png = file.replace('.xlsx', '.png')
        convert_excel_to_pdf(file, pdf)
        convert_pdf_to_png(pdf, png)
        self.bot_token = self.token_var.get().strip()
        ok, resp = send_image(self.bot_token, job['chat_id'], png)
        if ok:
            self.log(f"✅ Sent report to {job['name']}")
            if job['frequency'] == 'once':
                job['sent'] = True
        else:
            self.log(f"❌ Failed to send to {job['name']}: {resp}")
        for f in [pdf, png]:
            if os.path.exists(f):
                os.remove(f)
        self.save_schedules()

    def run_scheduler(self):
        for job in self.schedules:
            self.setup_schedule(job)
        threading.Thread(target=self.scheduler_loop, daemon=True).start()

    def scheduler_loop(self):
        while True:
            schedule.run_pending()
            time.sleep(1)

    def view_schedule(self):
        win = tk.Toplevel(self.root)
        win.title("Scheduled Tasks")
        lb = tk.Listbox(win)
        lb.pack(fill='both', expand=True)
        for s in self.schedules:
            freq = s['frequency']
            t = s['time'] if freq != 'hourly' else "Every Hour (9–6)"
            status = "✔️" if s.get("sent") else "⏳"
            lb.insert('end', f"{status} {os.path.basename(s['file'])} → {s['name']} ({t})")
        def remove():
            sel = lb.curselection()
            if sel:
                idx = sel[0]
                removed = self.schedules.pop(idx)
                self.save_schedules()
                lb.delete(idx)
                self.log(f"🗑 Removed {removed['name']} → {os.path.basename(removed['file'])}")
        ttk.Button(win, text="Remove Selected", command=remove, bootstyle=DANGER).pack()

if __name__ == "__main__":
    root = ttk.Window(themename="flatly")
    app = SETASchedulerApp(root)
    root.mainloop()