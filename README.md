# 🎂 Birthday Email Agent  
Automatically sends personalized birthday emails using Outlook.

---

## 📌 Overview
The **Birthday Email Agent** is a Python automation tool that reads employee data from an Excel file and sends personalized birthday greetings each day. It uses **Microsoft Outlook (COM automation)** to compose and deliver beautifully formatted HTML birthday emails.

This agent is designed for internal use within organizations that use Outlook as their default mail application.

---

## ✨ Features

- ✔ Automatically checks birthdays once per day  
- ✔ Sends fully custom HTML birthday email  
- ✔ Includes images, circle overlays, emojis, and personalized text  
- ✔ Supports VML for Outlook Desktop  
- ✔ Provides fallback HTML for Gmail, Outlook Mobile, and Apple Mail  
- ✔ Reads contacts from an Excel file  
- ✔ Works with Windows Task Scheduler for daily automation  

---

## 📁 Project Structure

    birthday-agent/
    │
    ├── birthday_agent.py # Main script
    ├── template.html # HTML template for email body
    ├── birthdays.xlsx # Excel file containing user data
    ├── requirements.txt # Gives the details about required python packages
    ├── assets/
    │ ├── birthday_bg.jpg # Background image (converted to base64)
    │ └── confetti.png # Overlays and icons
    │
    └── README.md # Documentation


---

## 📊 Excel Format

Your `birthdays.xlsx` must contain:

Email Address	Full Name	Birthdate	Joining Date 	CC
ab@xyz.com	AB	1-Jan	30-Mar-25	bc@xyz.com;cd@xyz.com
<img width="2369" height="49" alt="image" src="https://github.com/user-attachments/assets/16cb175d-4212-458a-94a5-9f234849c009" />


Dates must be in **YYYY-MM-DD** format.

---

## 🛠 Requirements

- Windows 10/11  
- Python 3.9+  
- Microsoft Outlook installed  
- Libraries:
  ```bash
  pip install pandas pywin32 python-dateutil

## 📧 How It Works

  1. Script loads the Excel file
  2. Compares today's date with each birthday
  3. Sends personalized HTML email using Outlook
  4. Logs last execution (to avoid duplicate sends)
  5. Exits
