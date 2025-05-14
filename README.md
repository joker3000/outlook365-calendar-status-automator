# Outlook Calendar Status Automator

> **Note:** This tool is mainly intended for environments **where Microsoft Teams is not established** as the company-wide tool and ONLY Outlook 365 and 3CX is used instead.

This Python desktop application allows an administrator to:

- Select a user from the Microsoft 365 tenant
- View and update all future **busy** calendar entries (where the user is the organizer)
- Set them to **free**
- Export changes as CSV
- Backup original calendar entries
- Restore selected entries (including subject, location, body)
- Send optional notification emails

---

## 🔧 Requirements

- Python 3.8 or later
- Microsoft 365 tenant admin account
- Azure AD App Registration (see below)
- Access to Microsoft Graph API

---

## ⚙️ Azure Setup

1. **Register an App:**
   - Go to [https://portal.azure.com](https://portal.azure.com)
   - Navigate to **Azure Active Directory > App registrations > New registration**
   - Name: `OutlookFreeSetter`
   - Supported account types: **Single tenant**
   - Redirect URI: leave empty
   - Click **Register**

2. **Create a Client Secret:**
   - Under **Certificates & secrets**, create a new secret
   - Copy the value **immediately** (you won't see it again)

3. **Add API Permissions:**
   - Go to **API permissions > Add a permission > Microsoft Graph**
   - Select **Application permissions**
     - `Calendars.ReadWrite`
     - `User.Read.All`
     - `Mail.Send` *(if you want email notification)*
   - Click **Add permissions**
   - Click **Grant admin consent**

4. **Note down:**
   - `Tenant ID`
   - `Client ID`
   - `Client Secret`

---

## 📁 Project Setup

1. **Clone or download this repository**

2. **Create a `.env` file in the root folder:**

   ```env
   TENANT_ID=your-tenant-id
   CLIENT_ID=your-client-id
   CLIENT_SECRET=your-client-secret
   ADMIN_MAIL=admin@example.com

3. **Install Python dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

4. **Launch the application:**

   ```bash
   python main.py
   ```

---

## 🖥️ Features

* User dropdown from the tenant
* Date range selector (default: now to +1 year)
* Only modifies appointments with `showAs == "busy"`
* CSV export and calendar backup before changes
* Restore from backup (with field selection)
* Multilingual interface (English & German)
* Email preview (and optional sending via Microsoft Graph)

---

## 📦 Packaging (optional)

To build a standalone `.exe` file (Windows):

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed main.py
```

---

## 🔒 Security

* Do **not** commit `.env` to version control.
* Use application secrets carefully and rotate periodically.

---

## 🧭 Project Purpose

This tool was built for **office employees using a 3CX phone system** integrated with Microsoft 365, who:

* Have **automatic calendar integration enabled** in 3CX (via Microsoft Graph)
* Appear **“busy”** in 3CX during any Outlook appointment – even personal or self-created ones
* Want to remain **available for phone calls** despite having personal calendar events

By converting self-created busy events to **“free”**, users can:

* Keep their calendars intact
* Avoid call redirection or DND status in 3CX
* Maintain full control over availability for telephony

---

## ⚠️ Disclaimer

* This software is provided **“as is”** without warranty of any kind, express or implied.
* The authors and contributors are **not liable** for any damages or data loss resulting from its use.
* You are responsible for reviewing any calendar changes before applying them.
* Always back up calendars or test in a sandbox environment before deploying to production.

---

## 🌐 Community & Support

For discussion, questions, and feedback, visit the official thread in the 3CX community:

👉 [3CX Forum: Outlook Calendar Status Automator](https://www.3cx.com/community/threads/outlook-calendar-status-automator.133628/)

---
