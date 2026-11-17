# Email KPI Dashboard

A desktop app that turns an exported email log into a visual dashboard of customer-service KPIs: how many emails were received, how many were answered, and how fast.

## Features
- Load any `.csv` or `.xlsx` email log through a file picker, with column validation and clear error messages
- KPI cards: total emails, responded, pending, and response rate
- Interactive data table and charts (matplotlib embedded in the app)
- Toast notifications that grade the response rate (excellent / OK / low)
- One-click export of the processed report to Excel
- Modern dark UI with sidebar, tabs and loading indicator

## Tech stack
Python · CustomTkinter · pandas · matplotlib

## Run it
```bash
pip install -r requirements.txt
python main.py
```
The input file needs the columns `estado`, `fecha recibido` and `fecha respondido`.
