# 📅 Conv-Schedule

> Schedule converter developed with Python and Flask.  
> Web application for transforming Excel timetables into calendar files (.ics).

![Python](https://img.shields.io/badge/Python-3.10-blue?logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-000000?logo=flask&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?logo=pandas&logoColor=white)
![Openpyxl](https://img.shields.io/badge/Openpyxl-success?logo=python&logoColor=white)
![ICS](https://img.shields.io/badge/ICS-Calendar-orange)
![GitHub](https://img.shields.io/badge/GitHub-Repository-black?logo=github)
![Vercel](https://img.shields.io/badge/Vercel-000000?logo=vercel&logoColor=white)
![Render](https://img.shields.io/badge/Render-46a2f1?logo=render&logoColor=white)

[![Live Demo](https://img.shields.io/badge/Live-Demo-success?style=for-the-badge)](https://conv-schedule.vercel.app)

---

## 🌐 Live Demo

👉 https://conv-schedule.vercel.app

---

## 📖 About

**Conv-Schedule** is a lightweight web application designed to convert academic or work schedules from Excel (`.xlsx`) into calendar files (`.ics`).  
These files can be imported into Google Calendar, Outlook, or any other calendar service.

---

## ✨ Features

- ✅ Upload Excel timetable files  
- ✅ Automatic conversion to `.ics` format  
- ✅ Support for multiple Excel parsers (`pandas`, `openpyxl`, `xlrd`)  
- ✅ Export ready‑to‑use calendar files  
- ✅ Simple and responsive Flask interface  
- ✅ Easy deployment on Vercel or Render  

---

## 📷 Screenshots

### Upload & Convert
![Upload](screenshots/upload.png)

### Generated Calendar File
![Result](screenshots/result.png)

---

## 🏗 Architecture

```text
┌─────────────────────┐
│ Flask (Python)      │
│ Backend + Frontend  │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ Excel Parsers       │
│ pandas / openpyxl   │
└──────────┬──────────┘
           │
           ▼
   ICS Generator (ics library)

🛠 Tech Stack
Python

Flask

pandas / openpyxl / xlrd

ics (calendar generation)

Vercel / Render (deployment)
