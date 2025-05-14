print("Starte Outlook-Automatisierung...")

import os
import requests
import json
import datetime
import csv
import base64
import mimetypes
import logging
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from dotenv import load_dotenv
from tabulate import tabulate

# ==== Logging konfigurieren ====
logging.basicConfig(filename='kalender_updater.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# ==== Umgebungsvariablen laden ====
load_dotenv()
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
ADMIN_MAIL = os.getenv("ADMIN_MAIL")

print("TENANT_ID:", TENANT_ID)
print("CLIENT_ID:", CLIENT_ID)
print("CLIENT_SECRET:", "[ausgeblendet]" if CLIENT_SECRET else "NICHT GESETZT")
print("ADMIN_MAIL:", ADMIN_MAIL)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = "https://graph.microsoft.com/.default"
GRAPH_URL = "https://graph.microsoft.com/v1.0"

# ==== Sprachkonfiguration ====
LANGUAGES = {
    "de": {
        "title": "Outlook Status Automatisierer",
        "choose_user": "Wähle einen Benutzer:",
        "update_btn": "Termine aktualisieren",
        "exit_btn": "Beenden",
        "select_folder": "Wähle Speicherort für CSV-Datei",
        "processing": "Verarbeite Benutzer:",
        "no_selection": "Bitte wähle einen gültigen Benutzer.",
        "abort": "Speicherort nicht ausgewählt. Vorgang abgebrochen.",
        "updated": "Termine geändert.",
        "csv_saved": "CSV gespeichert unter:",
        "email_sent": "E-Mail vom Benutzer aus an Admin & Benutzer gesendet.",
        "no_changes": "Keine Änderungen erforderlich.",
        "error": "Fehler:",
        "email_preview": "E-Mail Vorschau:",
        "backup_saved": "Backup gespeichert unter:",
        "from_label": "Von (YYYY-MM-DD):",
        "to_label": "Bis (YYYY-MM-DD):",
        "restore_btn": "Backup wiederherstellen",
        "restore_done": "Wiederherstellung abgeschlossen.",
        "restore_preview": "Wiederherstellungs-Vorschau:"
    },
    "en": {
        "title": "Outlook Status Automator",
        "choose_user": "Select a user:",
        "update_btn": "Update Appointments",
        "exit_btn": "Exit",
        "select_folder": "Choose folder to save CSV file",
        "processing": "Processing user:",
        "no_selection": "Please select a valid user.",
        "abort": "No folder selected. Aborting.",
        "updated": "Appointments updated.",
        "csv_saved": "CSV saved to:",
        "email_sent": "Email sent from user to Admin & User.",
        "no_changes": "No changes necessary.",
        "error": "Error:",
        "email_preview": "Email Preview:",
        "backup_saved": "Backup saved to:",
        "from_label": "From (YYYY-MM-DD):",
        "to_label": "To (YYYY-MM-DD):",
        "restore_btn": "Restore from Backup",
        "restore_done": "Restore completed.",
        "restore_preview": "Restore Preview:"
    }
}

def get_access_token():
    token_url = f"{AUTHORITY}/oauth2/v2.0/token"
    data = {
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': SCOPE
    }
    r = requests.post(token_url, data=data)
    r.raise_for_status()
    return r.json()['access_token']
def list_users(token):
    users = []
    url = f"{GRAPH_URL}/users?$select=id,displayName,mail"
    headers = {'Authorization': f'Bearer {token}'}
    while url:
        resp = requests.get(url, headers=headers)
        data = resp.json()
        users.extend(data['value'])
        url = data.get('@odata.nextLink')
    return users

def get_calendar_events(token, user_id, start_time, end_time):
    url = f"{GRAPH_URL}/users/{user_id}/calendarView?startDateTime={start_time}&endDateTime={end_time}&$top=1000"
    headers = {
        'Authorization': f'Bearer {token}',
        'Prefer': 'outlook.timezone=\"UTC\"'
    }
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()['value']

def backup_calendar(events, path, user_email):
    filename = os.path.join(path, f"backup_{user_email.replace('@','_')}_{datetime.datetime.utcnow().strftime('%Y%m%dT%H%M%S')}.csv")
    with open(filename, 'w', newline='', encoding='utf-8') as f:
        keys = ["id", "subject", "start", "end", "showAs", "organizer", "isOrganizer"]
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        for e in events:
            writer.writerow({
                "id": e.get("id", ""),
                "subject": e.get("subject", ""),
                "start": e.get("start", {}).get("dateTime", ""),
                "end": e.get("end", {}).get("dateTime", ""),
                "showAs": e.get("showAs", ""),
                "organizer": e.get("organizer", {}).get("emailAddress", {}).get("address", ""),
                "isOrganizer": e.get("isOrganizer", False)
            })
    return filename

def restore_from_backup(token, user_id, preview_widget):
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")], title="Select Backup File")
    if not file_path:
        return "No file selected."

    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    changes = []
    with open(file_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            eid = row["id"]
            patch_url = f"{GRAPH_URL}/users/{user_id}/events/{eid}"
            patch_body = {"showAs": row["showAs"]}
            r = requests.patch(patch_url, headers=headers, data=json.dumps(patch_body))
            log_msg = f"Restore: {row['subject']} ({row['start']}) -> {row['showAs']}"
            logging.info(log_msg)
            changes.append(log_msg)

    preview_widget.delete(1.0, tk.END)
    preview_widget.insert(tk.END, "\n".join(changes[:20]))  # Begrenze auf erste 20 Zeilen zur Übersicht
    return file_path
class OutlookApp:
    def __init__(self, root):
        self.root = root
        self.language = tk.StringVar(value="de")
        self.text = LANGUAGES[self.language.get()]
        self.token = get_access_token()
        print("Token erfolgreich empfangen.")
        self.users = list_users(self.token)
        self.user_map = {f"{u['displayName']} ({u['mail']})": u for u in self.users}
        self.root.title(self.text["title"])
        self.create_widgets()

    def create_widgets(self):
        # Sprachwahl
        lang_frame = tk.Frame(self.root)
        tk.Label(lang_frame, text="Sprache / Language:").pack(side=tk.LEFT)
        tk.Radiobutton(lang_frame, text="Deutsch", variable=self.language, value="de", command=lambda: self.switch_language("de")).pack(side=tk.LEFT)
        tk.Radiobutton(lang_frame, text="English", variable=self.language, value="en", command=lambda: self.switch_language("en")).pack(side=tk.LEFT)
        lang_frame.pack(pady=5)

        # Benutzerwahl
        self.label = tk.Label(self.root, text=self.text["choose_user"], font=("Arial", 12, "bold"))
        self.label.pack(pady=5)

        self.user_var = tk.StringVar()
        self.user_menu = ttk.Combobox(self.root, textvariable=self.user_var, width=60, font=("Arial", 10))
        self.user_menu['values'] = list(self.user_map.keys())
        self.user_menu.pack(pady=5)

        # Datumsbereich
        range_frame = tk.Frame(self.root)
        self.from_label = tk.Label(range_frame, text=self.text["from_label"])
        self.from_label.pack(side=tk.LEFT)
        self.from_entry = tk.Entry(range_frame, width=12)
        self.from_entry.insert(0, datetime.date.today().isoformat())
        self.from_entry.pack(side=tk.LEFT, padx=5)
        self.to_label = tk.Label(range_frame, text=self.text["to_label"])
        self.to_label.pack(side=tk.LEFT)
        self.to_entry = tk.Entry(range_frame, width=12)
        self.to_entry.insert(0, (datetime.date.today() + datetime.timedelta(days=365)).isoformat())
        self.to_entry.pack(side=tk.LEFT, padx=5)
        range_frame.pack(pady=5)

        # Buttons
        self.button = tk.Button(self.root, text=self.text["update_btn"], command=self.process, bg="#4CAF50", fg="white", font=("Arial", 11, "bold"))
        self.button.pack(pady=10)
        self.quit_button = tk.Button(self.root, text=self.text["exit_btn"], command=self.root.quit, bg="#f44336", fg="white", font=("Arial", 11))
        self.quit_button.pack(pady=5)

        self.restore_button = tk.Button(self.root, text=self.text["restore_btn"], command=self.restore_backup, bg="#2196F3", fg="white", font=("Arial", 11))
        self.restore_button.pack(pady=5)

        self.restore_fields = {
            "subject": tk.BooleanVar(value=True),
            "location": tk.BooleanVar(value=True),
            "body": tk.BooleanVar(value=True)
        }
        self.restore_options_frame = tk.Frame(self.root)
        for field in ["subject", "location", "body"]:
            tk.Checkbutton(self.restore_options_frame, text=field.capitalize(), variable=self.restore_fields[field]).pack(side=tk.LEFT, padx=5)
        self.restore_options_frame.pack(pady=2)

        self.restore_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, width=100, height=10)
        self.restore_listbox.pack(pady=5)

        self.restore_label = tk.Label(self.root, text=self.text["restore_preview"], font=("Arial", 10, "bold"))
        self.restore_label.pack()
        self.restore_preview = tk.Text(self.root, height=8, width=100, bg="#e8f0ff", font=("Courier", 9))
        self.restore_preview.pack(pady=5)

        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(pady=5)
        self.output = tk.Text(self.root, height=8, width=100, bg="#f4f4f4", font=("Courier", 10))
        self.output.pack(pady=10)

        self.email_label = tk.Label(self.root, text=self.text["email_preview"], font=("Arial", 10, "bold"))
        self.email_label.pack()
        self.email_preview = tk.Text(self.root, height=6, width=100, bg="#fffbe6", font=("Courier", 9))
        self.email_preview.pack(pady=5)

        tk.Button(self.root, text="Auswahl wiederherstellen", command=self.apply_restore, bg="#8BC34A", fg="black").pack(pady=5)

    def log(self, msg):
        timestamped = f"{datetime.datetime.now().isoformat()} - {msg}"
        logging.info(msg)
        self.output.insert(tk.END, timestamped + "\n")
        self.output.see(tk.END)

    def switch_language(self, lang):
        self.language.set(lang)
        self.text = LANGUAGES[lang]
        self.root.title(self.text["title"])
        self.label.config(text=self.text["choose_user"])
        self.button.config(text=self.text["update_btn"])
        self.quit_button.config(text=self.text["exit_btn"])
        self.email_label.config(text=self.text["email_preview"])
        self.from_label.config(text=self.text["from_label"])
        self.to_label.config(text=self.text["to_label"])
        self.restore_button.config(text=self.text["restore_btn"])
        self.restore_label.config(text=self.text["restore_preview"])

    # ... process() und apply_restore() folgen im Originaltext
    def process(self):
        choice = self.user_var.get()
        if not choice or choice not in self.user_map:
            messagebox.showerror(self.text["error"], self.text["no_selection"])
            return

        user = self.user_map[choice]
        self.log(f"{self.text['processing']} {user['displayName']} - {user['mail']}")

        save_path = filedialog.askdirectory(title=self.text["select_folder"])
        if not save_path:
            self.log(self.text["abort"])
            return

        try:
            start_str = self.from_entry.get()
            end_str = self.to_entry.get()
            start_time = datetime.datetime.fromisoformat(start_str).isoformat() + "Z"
            end_time = datetime.datetime.fromisoformat(end_str).isoformat() + "Z"

            events = get_calendar_events(self.token, user['id'], start_time, end_time)
            backup_file = backup_calendar(events, save_path, user['mail'])
            self.log(f"{self.text['backup_saved']} {backup_file}")

            own_events = [e for e in events if e.get('isOrganizer', False) and e.get('showAs') == 'busy']
            self.progress['maximum'] = len(own_events)
            updated = []

            for i, e in enumerate(own_events):
                self.progress['value'] = i + 1
                self.root.update_idletasks()
                patch_url = f"{GRAPH_URL}/users/{user['id']}/events/{e['id']}"
                r = requests.patch(
                    patch_url,
                    headers={'Authorization': f'Bearer {self.token}', 'Content-Type': 'application/json'},
                    data=json.dumps({'showAs': 'free'})
                )
                updated.append({
                    "Subject": e.get('subject'),
                    "Start": e.get('start', {}).get('dateTime'),
                    "End": e.get('end', {}).get('dateTime'),
                    "OriginalStatus": e.get('showAs'),
                    "UpdatedTo": "free"
                })

            if updated:
                csv_path = os.path.join(save_path, f"updated_{user['mail'].replace('@', '_')}.csv")
                with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=updated[0].keys())
                    writer.writeheader()
                    writer.writerows(updated)
                self.log(f"{len(updated)} {self.text['updated']}")
                self.log(f"{self.text['csv_saved']} {csv_path}")

                subject = f"Kalender-Update: {user['displayName']}"
                body = f"""Die folgenden Termine wurden erfolgreich auf 'frei' gesetzt.

Benutzer: {user['mail']}
Anzahl geänderter Termine: {len(updated)}
Datum: {datetime.datetime.utcnow().isoformat()} UTC
"""
                self.email_preview.delete(1.0, tk.END)
                self.email_preview.insert(tk.END, f"Subject: {subject}\n\n{body}")
                # Optional: send_mail_with_attachment(...) hier aufrufen

            else:
                self.log(self.text["no_changes"])
        except Exception as e:
            self.log(f"{self.text['error']} {str(e)}")

    def apply_restore(self):
        choice = self.user_var.get()
        if not choice or choice not in self.user_map:
            messagebox.showerror(self.text["error"], self.text["no_selection"])
            return

        user = self.user_map[choice]
        headers = {'Authorization': f'Bearer {self.token}', 'Content-Type': 'application/json'}
        selected_indices = self.restore_listbox.curselection()

        if not selected_indices:
            messagebox.showinfo("Info", "Keine Einträge ausgewählt.")
            return

        changes = []
        for idx in selected_indices:
            row = self.restore_data[idx]
            patch = {"showAs": row["showAs"]}
            if self.restore_fields["subject"].get():
                patch["subject"] = row.get("subject", "")
            if self.restore_fields["location"].get():
                patch["location"] = row.get("location", "")
            if self.restore_fields["body"].get():
                patch["body"] = {"contentType": "text", "content": row.get("body", "")}
            eid = row["id"]
            url = f"{GRAPH_URL}/users/{user['id']}/events/{eid}"
            r = requests.patch(url, headers=headers, data=json.dumps(patch))
            msg = f"Restore: {row['subject']} ({row['start']}) updated"
            logging.info(msg)
            changes.append(msg)

        self.restore_preview.delete(1.0, tk.END)
        self.restore_preview.insert(tk.END, "\n".join(changes))
        self.log(f"{self.text['restore_done']} Anzahl: {len(changes)}")

    def restore_backup(self):
        choice = self.user_var.get()
        if not choice or choice not in self.user_map:
            messagebox.showerror(self.text["error"], self.text["no_selection"])
            return

        user = self.user_map[choice]
        self.log(f"{self.text['processing']} {user['displayName']} - {user['mail']}")

        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")], title="Backup-Datei wählen")
        if not file_path:
            return

        self.restore_listbox.delete(0, tk.END)
        with open(file_path, newline='', encoding='utf-8') as f:
            reader = list(csv.DictReader(f))
            self.restore_data = reader
            for row in reader:
                entry = f"{row['start']} | {row['subject']} -> {row['showAs']}"
                self.restore_listbox.insert(tk.END, entry)
        
if __name__ == "__main__":
    print("Initialisiere GUI...")
    root = tk.Tk()
    app = OutlookApp(root)
    root.mainloop()
