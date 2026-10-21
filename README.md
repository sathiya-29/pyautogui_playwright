# Browser & Desktop Automation (Playwright + PyAutoGUI)

Two Python bots that automate repetitive work: one drives a **web app** with Playwright, the other drives **desktop apps** with PyAutoGUI. Both produce run reports, so every run can be checked afterwards.

## 1. Excel-driven web messaging bot (`playwright_assign.py`)

Reads contacts and message templates from `contacts.xlsx`, personalises each message (`Hello {name}, ...`) and sends it through WhatsApp Web, logging the result of every contact.

- Validates numbers and skips contacts that aren't registered
- Random 2–5 s delays between actions, so it behaves like a person
- Screenshot per contact, plus the last 3 messages read back to confirm delivery
- Writes `whatsapp_report.json` and `whatsapp_report.xlsx` after **every** contact, so nothing is lost if a run stops halfway

```
contacts.xlsx  →  Playwright (Chromium)  →  WhatsApp Web  →  report.xlsx + report.json + screenshots/
```

> **Use responsibly:** this is a learning project for small, personal, opt-in messaging. WhatsApp's terms prohibit bulk or automated messaging, and accounts that do it can be banned. For business messaging, use the official WhatsApp Business API.

## 2. Daily status report bot (`daily_report_bot.py`), macOS

Opens Chrome, collects today's weather, opens Numbers, builds a report sheet (date, weather, comment), saves it as Excel and takes a screenshot as proof. See the sample output in [`reports/`](reports/).

## Run

```bash
pip install playwright pyautogui openpyxl
playwright install chromium

python playwright_assign.py     # needs contacts.xlsx with columns Name, Phone, Message
python daily_report_bot.py      # macOS; grant Accessibility permission to your terminal
```

## What I learned

- Playwright: robust selectors, waits and timeouts, and error handling per item
- PyAutoGUI: keyboard and mouse control, clipboard reading, and checking that each step really worked
- Reporting: saving progress incrementally makes long automations safe to stop and restart

## Tech stack

Python · Playwright (async) · PyAutoGUI · openpyxl
