\# Huawei TX Intranet — Beta Version
\## 🎯 What is this?

An internal platform for monitoring Huawei TX links (RSL Degraded + Congestion), facilitating coordination between the UOP, DTV, and BO TX teams, as well as senior management.


\*\*Protected Beta Version\*\* — Valid until \*\*2026-12-31\*\*.
\---
\## 📋 Requirements



\- \*\*Windows 10 or later\*\* (64-bit)
\- \*\*No Python installation required\*\* — everything is built-in
\- \*\*At least 1 GB of RAM\*\*
\- \*\*A modern web browser\*\* (Chrome or Edge)
\- \*\*Local network connection\*\* (if you wish to share it with colleagues)
\---



\## 🚀 Getting Started — First-time Setup (5 minutes)
\### 1. Unzip the files
Place the folder in a permanent location, for example:C: \\HuaweiTX\\


⚠️ \*\*Do not place it on a USB drive or your desktop\*\* — because the database is saved alongside the file.
\### 2. Run the application
\*\*Double-click\*\* on: HuaweiTX\_Trial\_Server.exe
A message will appear containing:
\- \*\*Setup Token\*\* (`Mobilis2026`)
\- \*\*Trial Until\*\* (expiration date)
\- \*\*URL\*\* (`http://localhost:8503`)
Click \*\*OK\*\*.

\### 3. Open your browser
If it does not open automatically, go to:http://localhost:8503
\### 4. Create an Admin Account
On the \*\*First-Run Protected Setup\*\* screen:
| Field | Value |
|---|---|
| \*\*Setup Token\*\* | `Mobilis2026` (from the message) |
| \*\*Admin username\*\* | `admin` (or any name) |
| \*\*Full name\*\* | `BO TX Admin` |
| \*\*Password\*\* | 12+ characters, including uppercase, lowercase, a number, and a symbol |
| \*\*Confirm\*\* | The same password |
Click \*\*Create Admin \& Activate\*\*.


\### 5. Log In
Use:
\- \*\*Username:\*\* `admin`
\- \*\*Password:\*\* The one you created
\---
\## 📂 Folder Structure
HuaweiTX\_Trial\_Server
├── HuaweiTX\_Trial\_Server.exe ← Run this
├── README\_TRIAL.md ← this file
├── Copie of Situation HUAWEI... ← (Optional) Raw data
├── huawei\_tx\_database.db ← Automatically created (database)
├── backups\\ ← Automatic backups
├── logs\\ ← logs
└── \_internal\\ ← ❌ Do not touch

---

## 🌐 Sharing with Colleagues

1. Find your device's address

In PowerShell:
```powershell
ipconfig
```
Find the IPv4 address (e.g., 192.168.1.42).

2. Open the port in the firewall (one-time action)
PowerShell as administrator:

New-NetFirewallRule -DisplayName "Huawei TX" `
  -Direction Inbound -Protocol TCP `
  -LocalPort 8503 -Action Allow `
  -Profile Private,Domain

3. Send the link to the colleague
http://192.168.1.42:8503

4. Create an account for him
Under Admin & Audit → Users:
Role: UOP, DTV, or BO TX
Provide him with the username and password.

💾 Backup
Automatic
Before every Excel import
Before every restore
Manual
Under Admin & Audit → Backup & Restore:
Create backup now → Generates an immediate backup
Download DB snapshot → Downloads the file
Recommendation
Copy the `backups\` folder weekly to a secure location (OneDrive, USB drive, file server).

⏰ Expiration Date
This is a trial version, valid until 2026-12-31.
After this date:
The application will cease to function.
A "Trial Expired" message will appear.
It cannot be bypassed without a new version.
To obtain the final version, please contact the system developer.

🔒 Security — Important Notes
✅ System Features
Encrypted passwords (PBKDF2 600k)
Limited sessions (30 minutes of inactivity)
Account lockout after 5 failed attempts
Audit log for every operation

Atomic backups

⚠️ Your Responsibilities
Change the `admin` password after your first login (use the "Change my password" button in the sidebar).

Delete `token.txt` if it exists (it is no longer needed).
Do not share the admin password with anyone.
Do not expose port 8503 to the public internet.

🆘 Troubleshooting
Issue: The browser does not open automatically.
Solution: Open it manually at `http://localhost:8503`.

Issue: A black window appears and then disappears.
Solution: Run `HuaweiTX_Trial_Server.exe` via PowerShell to view the error:

powershell
cd C:\HuaweiTX\HuaweiTX_Trial_Server
.\HuaweiTX_Trial_Server.exe
Issue: A colleague cannot connect.
Check:
Is the `.exe` file running on your machine? Is the colleague on the same network?
Does the firewall allow port 8503?
From the colleague's device, try running: `Test-NetConnection <YOUR-IP> -Port 8503`
Issue: Windows Defender displays a warning.
Solution: Click "More info" → "Run anyway." This occurs because the file is not digitally signed (which is normal for trial versions).
Issue: "Trial Expired" message.
Solution: The trial period has ended. Please contact the developer.
Issue: "Invalid token" upon first launch.
Solution: Ensure you enter `Mobilis2026` exactly as shown (it is case-sensitive).

📞 Support
Developer: Ahmed Chaouli
Email: [chaouliaaa@gmail.com]
Phone: [+213661680039]

📝 Version History
Version	Date	Changes
12.5 Trial	2026-01	Initial trial version
Last updated: 2026-01



