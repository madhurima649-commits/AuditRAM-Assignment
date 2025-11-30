
# AuditRAM Coding Assignment 2

This project is a simple Python automation agent which:
1. Logs in to Outlook
2. Composes an email
3. Sends it to scittest@auditram.com

---

## Requirements
Install the required packages using:

```
pip install selenium webdriver-manager
```

You must have Google Chrome installed.

---

## How to Run
Open terminal and run:

```
python email_agent.py --email "your_email" --password "your_password" --message "This is my AuditRAM assignment email."
```

Example:

```
python email_agent.py --email "testuser@outlook.com" --password "Password123" --message "Hello, this is my assignment email."
```

This will:
- Open Chrome
- Log in to Outlook
- Compose an email
- Send automatically
- Print "Email sent successfully!"

---

## Repository Structure

```
email_agent.py
README.md
```
