# Outlook Calendar Status Automator

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

```
