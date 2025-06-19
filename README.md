# 📊 SETA – Scheduled Excel to Telegram Automation

**SETA** is a Python-based cross-platform desktop application that lets you:

- Automatically schedule Excel files to be sent as **images** via Telegram.
- Manage multiple recipients and file schedules.
- Convert Excel → PDF → PNG with cropped whitespace.
- Create daily, hourly, or one-time schedules.

---

## 🔧 Features

- Modern GUI using `ttkbootstrap`
- Excel → PDF → PNG conversion
- Image cropping for clean visuals
- Telegram Bot integration
- Persistent scheduling and recipient saving
- Compatible with `.exe` (Windows) and macOS App builds

---

## 🧰 Tech Stack

- Python 3.8–3.11 (not yet fully stable on 3.13)
- `Tkinter` and `ttkbootstrap` for GUI
- `schedule` for job scheduling
- `requests`, `Pillow` for Telegram & image handling
- `PyInstaller` for packaging

---

## 🚀 Installation

### 1. Clone or Download this repo

```bash
git clone https://github.com/yourname/SETA.git
cd SETA
