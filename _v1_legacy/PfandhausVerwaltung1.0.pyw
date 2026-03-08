import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
from datetime import datetime
import subprocess
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import black
# NEU: Imports für Paragraph und Styles
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import logging
from ttkthemes import ThemedTk
import hashlib
# import os # Bereits importiert
import sys
import uuid
# import subprocess # Bereits importiert
# from tkinter import messagebox # Bereits importiert
# NEU: shutil für Dateiverschiebung
import shutil
# NEU: Imports für Barcode-Generierung
try:
    import barcode
    from barcode.writer import ImageWriter # Benötigt Pillow: pip install Pillow
    from reportlab.lib.utils import ImageReader # Zum Einbetten des Barcodes in PDF
    import io # Um Barcode im Speicher zu erzeugen
    BARCODE_LIB_AVAILABLE = True
except ImportError:
    messagebox.showwarning("Bibliothek fehlt", "Die Bibliothek 'python-barcode' wurde nicht gefunden.\nBitte installieren Sie sie (pip install python-barcode Pillow), um Barcodes zu generieren.\nDie Barcode-Funktionalität ist deaktiviert.")
    BARCODE_LIB_AVAILABLE = False



# Logging konfigurieren
logging.basicConfig(filename='pfandhaus_app.log', level=logging.INFO, # Geändert auf INFO für mehr Details
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Verbindung zur Datenbank herstellen und Tabellen anlegen
def connect_db_static(db_path):
    conn = None
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        # Tabelle kunden (GEÄNDERT: Spalte zifferncode hinzugefügt)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS kunden (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                vorname TEXT NOT NULL,
                geburtsdatum TEXT NOT NULL,
                strasse TEXT,
                hausnummer TEXT,
                plz TEXT,
                ort TEXT,
                telefon TEXT,
                zifferncode INTEGER UNIQUE
            )
        ''')
        # Tabelle pfandscheine – mit neuem Feld 'artikel_beschreibung' (unverändert)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pfandscheine (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kunde_id INTEGER,
                abschlusstag TEXT,
                verfalltag TEXT,
                darlehen REAL,
                monatl_zinsen REAL,
                monatl_kosten REAL,
                versicherungssumme TEXT,
                vertragsnummer TEXT,
                artikel_beschreibung TEXT,
                FOREIGN KEY (kunde_id) REFERENCES kunden(id)
            )
        ''')
        # Tabelle pfandschein_historie – mit neuem Feld 'artikel_beschreibung' (unverändert)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pfandschein_historie (
                historie_id INTEGER PRIMARY KEY AUTOINCREMENT,
                pfandschein_id INTEGER,
                abschlusstag TEXT,
                verfalltag TEXT,
                darlehen REAL,
                monatl_zinsen REAL,
                monatl_kosten REAL,
                versicherungssumme TEXT,
                vertragsnummer TEXT,
                artikel_beschreibung TEXT,
                aenderungsdatum TEXT,
                FOREIGN KEY (pfandschein_id) REFERENCES pfandscheine(id)
            )
        ''')
        # Tabelle kunden_dokumente (unverändert)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS kunden_dokumente (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kunde_id INTEGER,
                dokument_pfad TEXT NOT NULL,
                dateiname TEXT NOT NULL,
                FOREIGN KEY (kunde_id) REFERENCES kunden(id)
            )
        ''')
        conn.commit()
        # NEU: Index für zifferncode erstellen (optional, verbessert Performance bei vielen Kunden)
        try:
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_zifferncode ON kunden(zifferncode)")
            conn.commit()
        except sqlite3.Error as e:
            logging.warning(f"Konnte Index für zifferncode nicht erstellen (möglicherweise existiert er schon): {e}")

        return conn
    except sqlite3.Error as e:
        logging.error(f"Fehler beim Verbinden oder Erstellen der Datenbank: {e}")
        if conn:
            conn.rollback()
        return None

class PfandhausApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pfandhaus Verwaltung")
        self.root.geometry("600x650")

        # NEU: Pfade und Einstellungen aus config laden oder Standard setzen
        self.db_path = "pfandhaus.db" # Standardwert
        self.pdf_background_path = "" # Standardwert
        self.current_theme = "breeze" # Standardwert
        self.last_zins_einheit = "%" # Standardwert
        self.document_base_path = os.path.join(os.path.expanduser("~"), "PfandhausDokumente") # Standardvorschlag

        # Überprüfen, ob die config.txt existiert und Werte laden/fehlende hinzufügen
        config_needs_update = False
        config_dict = {}

        if os.path.exists("config.txt"):
            try:
                with open("config.txt", "r") as f:
                    for line in f:
                        if "=" in line:
                            key, value = line.strip().split("=", 1)
                            config_dict[key] = value
            except Exception as e:
                logging.error(f"Fehler beim Lesen von config.txt beim Start: {e}")
                messagebox.showwarning("Konfigurationsfehler", f"Fehler beim Lesen von config.txt:\n{e}\nVerwende Standardeinstellungen.")
                # Bei Lesefehler mit leeren config_dict und Standardwerten weitermachen

        # Werte aus config_dict übernehmen oder Standard setzen, wenn fehlend
        if 'db_path' in config_dict:
            self.db_path = config_dict['db_path']
        else:
            config_dict['db_path'] = self.db_path
            config_needs_update = True

        if 'pdf_background_path' in config_dict:
            self.pdf_background_path = config_dict['pdf_background_path']
        else:
            config_dict['pdf_background_path'] = self.pdf_background_path
            config_needs_update = True

        if 'theme' in config_dict:
            self.current_theme = config_dict['theme']
        else:
            config_dict['theme'] = self.current_theme
            config_needs_update = True

        if 'last_zins_einheit' in config_dict:
            self.last_zins_einheit = config_dict['last_zins_einheit']
        else:
            config_dict['last_zins_einheit'] = self.last_zins_einheit
            config_needs_update = True

        if 'document_base_path' in config_dict:
            self.document_base_path = config_dict['document_base_path']
        else:
            config_dict['document_base_path'] = self.document_base_path
            config_needs_update = True


        # Wenn etwas fehlte oder die Datei neu erstellt werden muss, config.txt neu schreiben
        if config_needs_update or not os.path.exists("config.txt"):
            try:
                with open("config.txt", "w") as f:
                    for key, value in config_dict.items():
                        f.write(f"{key}={value}\n")
                logging.info("config.txt aktualisiert oder neu erstellt mit fehlenden Einträgen.")
            except Exception as e:
                 logging.error(f"Fehler beim Schreiben/Aktualisieren von config.txt: {e}")
                 messagebox.showwarning("Konfigurationsfehler", f"config.txt konnte nicht aktualisiert/geschrieben werden:\n{e}\nBitte prüfen Sie die Dateirechte.")


        self.available_themes = ThemedTk().get_themes()
        # Sicherstellen, dass das gespeicherte Theme existiert, sonst Fallback
        if self.current_theme not in self.available_themes:
             logging.warning(f"Gespeichertes Theme '{self.current_theme}' nicht verfügbar. Verwende Standardtheme 'clam'.")
             self.current_theme = "clam"

        self.root.set_theme(self.current_theme)


        self.selected_customer_for_edit = None # Variable für die ID des ausgewählten Kunden zum Bearbeiten
        self.hands_free_zifferncode_search = tk.BooleanVar(value=True) # NEU: Variable für Hands-free Modus, default aktiviert


        self.conn = self.connect_db()
        if self.conn is None:
            messagebox.showerror("Fehler", "Die Datenbank konnte nicht geladen werden. Überprüfen Sie die Logdatei für Details.")
            self.root.destroy()
            return
        self.create_widgets()
        self.create_menu()
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<ButtonRelease-1>", self.on_customer_select) # Einzelklick für Auswahl zum Bearbeiten
        self.selected_customer_id = None
        logging.info(f"Pfandhaus-Anwendung gestartet mit Theme: {self.current_theme}, letzte Zinseinheit: {self.last_zins_einheit}, Dokumentenpfad: {self.document_base_path}")

    # --- Methoden bis save_customer (get_last_zins_einheit, save_last_zins_einheit, create_menu, etc.) ---
    # --- sind unverändert zur vorherigen Version und werden hier zur Kürze weggelassen ---
    # --- Fügen Sie hier die Methoden von get_last_zins_einheit bis show_customer_documents ein ---
    def get_last_zins_einheit(self):
        try:
            with open("config.txt", "r") as f:
                for line in f:
                    if line.startswith("last_zins_einheit="):
                        return line.split("=", 1)[1].strip()
            return "%" # Standardwert, falls nicht gefunden
        except FileNotFoundError:
            #logging.warning("config.txt nicht gefunden in get_last_zins_einheit. Verwende Standard-Zinseinheit '%'.") # Log bereits in __init__
            return "%"
        except Exception as e:
            logging.error(f"Fehler beim Lesen der letzten Zinseinheit aus config.txt: {e}")
            return "%"

    def save_last_zins_einheit(self, einheit):
        try:
            with open("config.txt", "r+") as f:
                lines = f.readlines()
                f.seek(0)
                found = False
                for line in lines:
                    if line.startswith("last_zins_einheit="):
                        f.write(f"last_zins_einheit={einheit}\n")
                        found = True
                    else:
                        f.write(line)
                if not found:
                    f.write(f"last_zins_einheit={einheit}\n")
                f.truncate()
            self.last_zins_einheit = einheit
            logging.info(f"Letzte Zinseinheit '{einheit}' in config.txt gespeichert.")
        except Exception as e:
            logging.error(f"Fehler beim Speichern der letzten Zinseinheit in config.txt: {e}")
            messagebox.showerror("Fehler", f"Fehler beim Speichern der letzten Zinseinheit: {e}")

    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Datenbankpfad ändern...", command=self.open_change_db_path_window)
        file_menu.add_command(label="PDF Hintergrundbild ändern...", command=self.open_change_pdf_background_path_window)
        file_menu.add_command(label="Basisordner Dokumente ändern...", command=self.open_change_document_base_path_window) # NEU: Eintrag für Dokumentenpfad
        file_menu.add_command(label="Theme auswählen...", command=self.open_theme_selection_window) # Neuer Eintrag
        file_menu.add_separator()
        file_menu.add_command(label="Beenden", command=self.root.quit)
        menubar.add_cascade(label="Einstellungen", menu=file_menu)

        # NEU: Option für Hands-free Zifferncode-Suche im Menü
        einstellungen_menu = tk.Menu(menubar, tearoff=0)
        einstellungen_menu.add_checkbutton(label="Hands-free Zifferncode-Suche", variable=self.hands_free_zifferncode_search)
        menubar.add_cascade(label="Hands-free", menu=einstellungen_menu)


        # Neuer "Info" Menüpunkt
        info_menu = tk.Menu(menubar, tearoff=0)
        info_menu.add_command(label="Über Pfandhaus Verwaltung", command=self.open_about_window)
        menubar.add_cascade(label="Info", menu=info_menu)

        self.root.config(menu=menubar)

    def open_about_window(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("Über Pfandhaus Verwaltung")
        about_window.resizable(False, False) # Fenstergröße festlegen
        about_window.transient(self.root)
        about_window.grab_set()

        version_info = "Version: 1.2" # NEU: Versionsnummer angepasst
        manufacturer_info = "Hersteller: innobytix-IT Manuel Person"
        support_contact = "Kontakt: manuel.person@outlook.de"

        ttk.Label(about_window, text="Pfandhaus Verwaltung", font=("Arial", 12, "bold")).pack(padx=10, pady=10)
        ttk.Label(about_window, text=version_info).pack(padx=10, pady=5)
        ttk.Label(about_window, text=manufacturer_info).pack(padx=10, pady=5)
        ttk.Label(about_window, text=support_contact).pack(padx=10, pady=5)
        ttk.Label(about_window, text="© 2025 Alle Rechte vorbehalten.").pack(padx=10, pady=10)

        close_button = ttk.Button(about_window, text="Schließen", command=about_window.destroy)
        close_button.pack(pady=10)

    def open_theme_selection_window(self):
        theme_window = tk.Toplevel(self.root)
        theme_window.title("Theme auswählen")
        theme_window.transient(self.root)
        theme_window.grab_set()


        ttk.Label(theme_window, text="Wählen Sie ein Theme:").pack(padx=10, pady=10)

        self.theme_var = tk.StringVar(theme_window)
        self.theme_var.set(self.current_theme) # Aktuelles Theme auswählen

        theme_combobox = ttk.Combobox(theme_window, textvariable=self.theme_var, values=self.available_themes, state="readonly")
        theme_combobox.pack(padx=10, pady=5)

        def apply_selected_theme():
            selected_theme = self.theme_var.get()
            self.root.set_theme(selected_theme)
            self.current_theme = selected_theme
            self.save_theme(selected_theme)
            messagebox.showinfo("Info", f"Theme auf '{selected_theme}' geändert. Die Einstellung wird für zukünftige Starts gespeichert.")
            theme_window.destroy()

        apply_button = ttk.Button(theme_window, text="Theme anwenden und speichern", command=apply_selected_theme)
        apply_button.pack(pady=10)

        cancel_button = ttk.Button(theme_window, text="Abbrechen", command=theme_window.destroy)
        cancel_button.pack(pady=5)

    def save_theme(self, theme_name):
        try:
            with open("config.txt", "r+") as f:
                lines = f.readlines()
                f.seek(0)
                found = False
                for line in lines:
                    if line.startswith("theme="):
                        f.write(f"theme={theme_name}\n")
                        found = True
                    else:
                        f.write(line)
                if not found:
                    f.write(f"theme={theme_name}\n")
                f.truncate()
            logging.info(f"Theme '{theme_name}' in config.txt gespeichert.")
        except Exception as e:
            logging.error(f"Fehler beim Speichern des Themes in config.txt: {e}")
            messagebox.showerror("Fehler", f"Fehler beim Speichern des Themes: {e}")

    def get_saved_theme(self):
        try:
            with open("config.txt", "r") as f:
                for line in f:
                    if line.startswith("theme="):
                        return line.split("=", 1)[1].strip()
            return "clam" # Standardtheme, falls keins gefunden wird (Fallback falls __init__ fehlschlägt)
        except FileNotFoundError:
            #logging.warning("config.txt nicht gefunden in get_saved_theme. Verwende Standardtheme 'clam'.") # Log bereits in __init__
            return "clam"
        except Exception as e:
            logging.error(f"Fehler beim Lesen des Themes aus config.txt: {e}")
            return "clam"

    def open_change_pdf_background_path_window(self):
        change_path_window = tk.Toplevel(self.root)
        change_path_window.title("PDF Hintergrundbild ändern")
        change_path_window.transient(self.root) # Fenster über Hauptfenster halten
        change_path_window.grab_set() # Modales Fenster

        ttk.Label(change_path_window, text="Neuer Pfad zum Hintergrundbild:").pack(padx=10, pady=10)
        self.entry_new_pdf_background_path = ttk.Entry(change_path_window, width=40)
        self.entry_new_pdf_background_path.pack(padx=10, pady=5)
        self.entry_new_pdf_background_path.insert(0, self.get_pdf_background_path())

        btn_browse = ttk.Button(change_path_window, text="Durchsuchen...", command=self.browse_pdf_background_path)
        btn_browse.pack(pady=5)

        btn_save = ttk.Button(change_path_window, text="Speichern", command=self.save_new_pdf_background_path)
        btn_save.pack(pady=10)

        btn_cancel = ttk.Button(change_path_window, text="Abbrechen", command=change_path_window.destroy)
        btn_cancel.pack(pady=5)

    def browse_pdf_background_path(self):
        file_path = filedialog.askopenfilename(
            parent=self.entry_new_pdf_background_path.winfo_toplevel(), # Übergeordnetes Fenster explizit angeben
            title="PDF Hintergrundbild auswählen",
            filetypes=[("Bilddateien", "*.png;*.jpg;*.jpeg;*.gif"), ("Alle Dateien", "*.*")]
        )
        if file_path:
            self.entry_new_pdf_background_path.delete(0, tk.END)
            self.entry_new_pdf_background_path.insert(0, file_path) # Pfad sofort ins Feld schreiben

    def save_new_pdf_background_path(self):
        new_path = self.entry_new_pdf_background_path.get().strip()
        if new_path:
            try:
                with open("config.txt", "r+") as f:
                    lines = f.readlines()
                    found = False
                    f.seek(0)
                    for line in lines:
                        if line.startswith("pdf_background_path="):
                            f.write(f"pdf_background_path={new_path}\n")
                            found = True
                        else:
                            f.write(line)
                    if not found:
                        f.write(f"pdf_background_path={new_path}\n")
                    f.truncate()
                messagebox.showinfo("Erfolg", f"PDF Hintergrundbildpfad gespeichert: {new_path}")
                self.pdf_background_path = new_path
                logging.info(f"PDF Hintergrundbildpfad gespeichert: {new_path}")
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler beim Speichern des PDF Hintergrundbildpfads: {e}")
                logging.error(f"Fehler beim Speichern des PDF Hintergrundbildpfads: {e}")
        else:
            messagebox.showwarning("Warnung", "Bitte geben Sie einen gültigen Pfad zum Hintergrundbild ein.")

    def get_pdf_background_path(self):
        try:
            with open("config.txt", "r") as f:
                for line in f:
                    if line.startswith("pdf_background_path="):
                        return line.split("=", 1)[1].strip()
        except FileNotFoundError:
            #logging.warning("config.txt nicht gefunden in get_pdf_background_path. PDF Hintergrundbildpfad ist leer.") # Log bereits in __init__
            return ""
        except Exception as e:
            logging.error(f"Fehler beim Lesen des PDF Hintergrundbildpfads aus config.txt: {e}")
            return ""

    def get_db_path(self):
        try:
            with open("config.txt", "r") as f:
                for line in f:
                    if line.startswith("db_path="):
                        return line.split("=", 1)[1].strip()
        except FileNotFoundError:
            #logging.warning("config.txt nicht gefunden in get_db_path. Verwende Standard-Datenbankpfad.") # Log bereits in __init__
            return "pfandhaus.db"
        except Exception as e:
            logging.error(f"Fehler beim Lesen des Datenbankpfads aus config.txt: {e}")
            return "pfandhaus.db"

    def connect_db(self):
        db_path = self.get_db_path()
        conn = connect_db_static(db_path)
        if conn is None:
            logging.error(f"Verbindung zur Datenbank unter '{db_path}' konnte nicht hergestellt werden.")
        else:
            logging.info(f"Erfolgreich mit der Datenbank unter '{db_path}' verbunden.")
        return conn

    # --- Neue Methoden zum Verwalten des Kundendokumente-Pfades ---
    def get_document_base_path(self):
        """Holt den Basisordnerpfad für Kundendokumente aus der config.txt."""
        try:
            with open("config.txt", "r") as f:
                for line in f:
                    if line.startswith("document_base_path="):
                        return line.split("=", 1)[1].strip()
        except FileNotFoundError:
            logging.warning("config.txt nicht gefunden in get_document_base_path. Verwende Standard-Dokumentenpfad im Benutzerverzeichnis.")
            default_path = os.path.join(os.path.expanduser("~"), "PfandhausDokumente")
            # Versuch, den Standardpfad direkt in config.txt zu schreiben, falls Datei fehlt
            try:
                 with open("config.txt", "a") as f: # 'a' fügt hinzu, 'w' würde überschreiben
                     f.write(f"document_base_path={default_path}\n")
                 logging.info(f"Standard-Dokumentenpfad in neu erstellte config.txt geschrieben: {default_path}")
            except Exception as e:
                 logging.error(f"Fehler beim Schreiben des Standard-Dokumentenpfads in config.txt: {e}")

            return default_path # Rückgabe des Standardpfads
        except Exception as e:
            logging.error(f"Fehler beim Lesen des Dokumentenpfads aus config.txt: {e}")
            return "" # Leeren String bei Fehler zurückgeben

    def open_change_document_base_path_window(self):
        """Öffnet ein Fenster zum Ändern des Basisordnerpfades für Kundendokumente."""
        change_path_window = tk.Toplevel(self.root)
        change_path_window.title("Basisordner für Kundendokumente ändern")
        change_path_window.transient(self.root) # Fenster über Hauptfenster halten
        change_path_window.grab_set() # Modales Fenster

        ttk.Label(change_path_window, text="Neuer Basisordner für Kundendokumente:\n(Hierin wird ein Unterordner 'Kundendokumente' und dann die Code-Ordner angelegt)").pack(padx=10, pady=10)
        self.entry_new_document_base_path = ttk.Entry(change_path_window, width=50)
        self.entry_new_document_base_path.pack(padx=10, pady=5)
        self.entry_new_document_base_path.insert(0, self.get_document_base_path())

        btn_browse = ttk.Button(change_path_window, text="Ordner auswählen...", command=self.browse_document_base_path)
        btn_browse.pack(pady=5)

        btn_save = ttk.Button(change_path_window, text="Speichern", command=self.save_new_document_base_path)
        btn_save.pack(pady=10)

        btn_cancel = ttk.Button(change_path_window, text="Abbrechen", command=change_path_window.destroy)
        btn_cancel.pack(pady=5)

    def browse_document_base_path(self):
        """Öffnet einen Dialog zur Auswahl eines Ordners für Kundendokumente."""
        folder_path = filedialog.askdirectory(
            parent=self.entry_new_document_base_path.winfo_toplevel(),
            title="Basisordner für Kundendokumente auswählen",
            initialdir=os.path.dirname(self.get_document_base_path() or os.path.expanduser("~")) # Start im aktuellen oder Benutzerverzeichnis
        )
        if folder_path:
            self.entry_new_document_base_path.delete(0, tk.END)
            self.entry_new_document_base_path.insert(0, folder_path)

    def save_new_document_base_path(self):
        """Speichert den neuen Basisordnerpfad für Kundendokumente in der config.txt."""
        new_path = self.entry_new_document_base_path.get().strip()
        if not new_path:
            messagebox.showwarning("Warnung", "Bitte geben Sie einen gültigen Pfad für den Basisordner ein.")
            return

        # Optional: Prüfen, ob der Pfad beschreibbar ist
        if not os.path.isdir(new_path) or not os.access(new_path, os.W_OK):
             messagebox.showerror("Fehler", f"Der ausgewählte Pfad ist kein gültiger Ordner oder nicht beschreibbar:\n{new_path}")
             logging.error(f"Ausgewählter Dokumenten-Basisordner '{new_path}' ist ungültig oder nicht beschreibbar.")
             return

        try:
            with open("config.txt", "r+") as f:
                lines = f.readlines()
                f.seek(0)
                found = False
                for line in lines:
                    if line.startswith("document_base_path="):
                        f.write(f"document_base_path={new_path}\n")
                        found = True
                    else:
                        f.write(line)
                if not found:
                    f.write(f"document_base_path={new_path}\n")
                f.truncate()
            self.document_base_path = new_path
            messagebox.showinfo("Erfolg", f"Basisordner für Kundendokumente gespeichert: {self.document_base_path}\nEin Unterordner 'Kundendokumente' wird darin für die Ablage der Code-Ordner verwendet.")
            logging.info(f"Basisordner für Kundendokumente gespeichert: {self.document_base_path}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern des Basisordners für Kundendokumente: {e}")
            logging.error(f"Fehler beim Speichern des Basisordners für Kundendokumente: {e}")

    # --- Neue Methode zum Erstellen des Kundenordners ---
    def create_customer_document_folder(self, customer_id, zifferncode):
        """Erstellt den spezifischen Dokumentenordner für einen Kunden basierend auf dem Zifferncode."""
        if not self.document_base_path:
            logging.warning(f"Dokumenten-Basisordner ist nicht gesetzt. Konnte keinen Ordner für Kunde ID {customer_id} (Code {zifferncode}) erstellen.")
            #messagebox.showwarning("Warnung", "Der Basisordner für Kundendokumente ist nicht konfiguriert. Ordner für diesen Kunden wurde nicht erstellt.") # Zu aufdringlich? Nur loggen
            return None # Wichtig: None zurückgeben, wenn der Basisordner fehlt

        # Hauptverzeichnis "Kundendokumente" innerhalb des Basisordners
        main_docs_dir = os.path.join(self.document_base_path, "Kundendokumente")
        # Spezifischer Kundenordner (Zifferncode) innerhalb von "Kundendokumente"
        customer_dir = os.path.join(main_docs_dir, str(zifferncode))

        try:
            # Beide Ordner erstellen, falls sie nicht existieren. exist_ok=True verhindert Fehler, wenn sie schon da sind.
            os.makedirs(customer_dir, exist_ok=True)
            logging.info(f"Kundenordner für ID {customer_id} (Code {zifferncode}) erstellt oder existiert bereits: {customer_dir}")
            return customer_dir # Gib den Pfad zum erstellten Ordner zurück
        except OSError as e:
            logging.error(f"Fehler beim Erstellen des Kundenordners für ID {customer_id} (Code {zifferncode}) unter '{customer_dir}': {e}")
            messagebox.showerror("Ordner Fehler", f"Konnte den Ordner für Kunde {zifferncode} nicht erstellen:\n{customer_dir}\n\nFehler: {e}\nBitte prüfen Sie die Berechtigungen.")
            return None # Wichtig: None zurückgeben, wenn ein Fehler auftritt


    def create_widgets(self):
        ttk.Label(self.root, text="Pfandhaus", font=("Arial", 20, "bold")).pack(pady=10)

        # Suchleiste
        search_frame = ttk.Frame(self.root)
        search_frame.pack(pady=5)

        # Allgemeines Suchfeld
        ttk.Label(search_frame, text="Suche:").pack(side=tk.LEFT)
        self.entry_search = ttk.Entry(search_frame)
        self.entry_search.pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="Suchen", command=self.search_customers).pack(side=tk.LEFT)

        # Suchfeld für Zifferncode
        ttk.Label(search_frame, text="Zifferncode:").pack(side=tk.LEFT, padx=10) # Abstand zum vorherigen Suchfeld
        self.entry_zifferncode_search = ttk.Entry(search_frame, width=10) # Breite für Zifferncode angepasst
        self.entry_zifferncode_search.pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="Code Suchen", command=self.search_customers_by_zifferncode).pack(side=tk.LEFT)

        ttk.Button(search_frame, text="Alle anzeigen", command=self.load_customers).pack(side=tk.LEFT, padx=5)


        # Eingabefelder für Kunden
        self.frame_kunde_eingabe = ttk.Frame(self.root) # Frame für bessere Organisation
        self.frame_kunde_eingabe.pack(pady=10)

        ttk.Label(self.frame_kunde_eingabe, text="Name:").grid(row=0, column=0, sticky=tk.E)
        self.entry_name = ttk.Entry(self.frame_kunde_eingabe)
        self.entry_name.grid(row=0, column=1)

        ttk.Label(self.frame_kunde_eingabe, text="Vorname:").grid(row=1, column=0, sticky=tk.E)
        self.entry_vorname = ttk.Entry(self.frame_kunde_eingabe)
        self.entry_vorname.grid(row=1, column=1)

        ttk.Label(self.frame_kunde_eingabe, text="Geburtsdatum (TT.MM.JJJJ):").grid(row=2, column=0, sticky=tk.E)
        self.entry_geburtsdatum = ttk.Entry(self.frame_kunde_eingabe)
        self.entry_geburtsdatum.grid(row=2, column=1)

        ttk.Label(self.frame_kunde_eingabe, text="Straße:").grid(row=3, column=0, sticky=tk.E)
        self.entry_strasse = ttk.Entry(self.frame_kunde_eingabe)
        self.entry_strasse.grid(row=3, column=1)

        ttk.Label(self.frame_kunde_eingabe, text="Hausnummer:").grid(row=4, column=0, sticky=tk.E)
        self.entry_hausnummer = ttk.Entry(self.frame_kunde_eingabe)
        self.entry_hausnummer.grid(row=4, column=1)

        ttk.Label(self.frame_kunde_eingabe, text="Postleitzahl:").grid(row=5, column=0, sticky=tk.E)
        self.entry_plz = ttk.Entry(self.frame_kunde_eingabe)
        self.entry_plz.grid(row=5, column=1)

        ttk.Label(self.frame_kunde_eingabe, text="Ort:").grid(row=6, column=0, sticky=tk.E)
        self.entry_ort = ttk.Entry(self.frame_kunde_eingabe)
        self.entry_ort.grid(row=6, column=1)

        ttk.Label(self.frame_kunde_eingabe, text="Telefonnummer:").grid(row=7, column=0, sticky=tk.E)
        self.entry_telefon = ttk.Entry(self.frame_kunde_eingabe)
        self.entry_telefon.grid(row=7, column=1)


        # Steuerungsknöpfe
        self.btn_save_kunde = ttk.Button(self.root, text="Kunde speichern", command=self.save_customer, width=20)
        self.btn_save_kunde.pack(pady=10)
        ttk.Button(self.root, text="Felder leeren", command=self.clear_fields, width=20).pack(pady=5)
        ttk.Button(self.root, text="Kunde löschen", command=self.delete_customer, width=20).pack(pady=5)


        # Treeview für Kundenliste (GEÄNDERT: Spalte Zifferncode hinzugefügt)
        self.tree = ttk.Treeview(
            self.root,
            columns=("ID", "Zifferncode", "Name", "Vorname", "Geburtsdatum", "Straße", "Hausnummer", "PLZ", "Ort", "Telefon"),
            show="headings"
        )
        for col in ("ID", "Zifferncode", "Name", "Vorname", "Geburtsdatum", "Straße", "Hausnummer", "PLZ", "Ort", "Telefon"):
            self.tree.heading(col, text=col)
            # NEU: Spaltenbreiten anpassen
            if col == "ID":
                self.tree.column(col, width=40, stretch=tk.NO)
            elif col == "Zifferncode":
                self.tree.column(col, width=80, stretch=tk.NO)
            elif col == "Geburtsdatum":
                 self.tree.column(col, width=100, stretch=tk.NO)
            elif col in ["Straße", "Ort"]:
                 self.tree.column(col, width=150)
            elif col in ["Hausnummer", "PLZ"]:
                 self.tree.column(col, width=80, stretch=tk.NO)
            else:
                self.tree.column(col, width=120)

        self.tree.pack(pady=10, fill=tk.BOTH, expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # Doppelklick: Kunde auswählen und Pfandschein verwalten
        self.tree.bind("<Double-1>", self.on_customer_double_click)

        # NEU: Event Bindings für Enter-Taste in Suchfeldern
        self.entry_zifferncode_search.bind("<Return>", self.zifferncode_search_enter_pressed) # Enter für Zifferncode-Suche
        self.entry_search.bind("<Return>", self.fulltext_search_enter_pressed) # Enter für Volltext-Suche


        self.load_customers()

    def on_customer_select(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            customer_id = self.tree.item(selected_item[0])["values"][0]
            self.load_customer_data_for_edit(customer_id)
            self.selected_customer_for_edit = customer_id
            self.btn_save_kunde.config(text="Änderungen speichern") # Button-Text ändern
        else:
            self.selected_customer_for_edit = None
            self.btn_save_kunde.config(text="Kunde speichern") # Button-Text zurücksetzen

    def load_customer_data_for_edit(self, customer_id):
        cursor = self.conn.cursor()
        try:
            # GEÄNDERT: zifferncode wird nicht zum Bearbeiten geladen (ist fix)
            cursor.execute("SELECT name, vorname, geburtsdatum, strasse, hausnummer, plz, ort, telefon FROM kunden WHERE id=?", (customer_id,))
            customer_data = cursor.fetchone()
            if customer_data:
                self.clear_fields() # Felder zuerst leeren
                self.entry_name.insert(0, customer_data[0])
                self.entry_vorname.insert(0, customer_data[1])
                self.entry_geburtsdatum.insert(0, customer_data[2])
                self.entry_strasse.insert(0, customer_data[3])
                self.entry_hausnummer.insert(0, customer_data[4])
                self.entry_plz.insert(0, customer_data[5])
                self.entry_ort.insert(0, customer_data[6])
                self.entry_telefon.insert(0, customer_data[7])
                logging.info(f"Kundendaten für ID {customer_id} zum Bearbeiten geladen.")
                # Markiere den Kunden in der Treeview
                for item in self.tree.get_children():
                    if self.tree.item(item)['values'][0] == str(customer_id):
                        self.tree.selection_set(item)
                        self.tree.focus(item) # Optional: Fokus auf das Item setzen
                        self.tree.see(item) # Optional: Sicherstellen, dass Item sichtbar ist
                        break


        except sqlite3.Error as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Kundendaten zum Bearbeiten: {e}")
            logging.error(f"Fehler beim Laden der Kundendaten für ID {customer_id} zum Bearbeiten: {e}")

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.selected_customer_id = self.tree.item(item)["values"][0]
            self.context_menu = tk.Menu(self.root, tearoff=0)
            self.context_menu.add_command(label="Dokument zuordnen", command=self.assign_document_to_customer)
            self.context_menu.add_command(label="Dokumente anzeigen", command=self.show_customer_documents)
            self.context_menu.post(event.x_root, event.y_root)
        else:
            self.selected_customer_id = None

    def assign_document_to_customer(self):
        if self.selected_customer_id is None:
            messagebox.showwarning("Warnung", "Bitte wählen Sie zuerst einen Kunden per Rechtsklick aus.")
            return

        if not self.document_base_path:
             messagebox.showwarning("Konfigurationsfehler", "Der Basisordner für Kundendokumente ist nicht konfiguriert. Bitte gehen Sie zu Einstellungen -> Basisordner Dokumente ändern.")
             logging.warning("Dokumentenzuordnung abgebrochen: Basisordner nicht konfiguriert.")
             return

        # 1. Kundeninformationen (insb. Zifferncode) abrufen
        cursor = self.conn.cursor()
        try:
            cursor.execute("SELECT zifferncode, name, vorname FROM kunden WHERE id=?", (self.selected_customer_id,))
            kunde_data = cursor.fetchone()
            if not kunde_data:
                messagebox.showerror("Fehler", f"Kunde mit ID {self.selected_customer_id} nicht in der Datenbank gefunden.")
                logging.error(f"Kunde mit ID {self.selected_customer_id} nicht gefunden beim Dokumentenzuordnungsversuch.")
                return
            zifferncode, kunde_name, kunde_vorname = kunde_data

            if zifferncode is None:
                 messagebox.showwarning("Warnung", f"Kunde '{kunde_vorname} {kunde_name}' hat keinen Zifferncode.\nEin Ordner kann erst nach Zuweisung eines Zifferncodes erstellt werden.\nBitte Kunde bearbeiten und speichern, um einen Code zu erhalten.")
                 logging.warning(f"Dokumentenzuordnung abgebrochen: Kunde ID {self.selected_customer_id} hat keinen Zifferncode.")
                 return

            # 2. Zielordner ermitteln und ggf. erstellen
            customer_docs_dir = self.create_customer_document_folder(self.selected_customer_id, zifferncode) # Rufe HELPER-Methode auf
            if not customer_docs_dir:
                # Fehlermeldung und Log passieren bereits in create_customer_document_folder
                return # Abbruch, wenn Ordner nicht erstellt werden konnte

        except sqlite3.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Abrufen des Zifferncodes für Kunde ID {self.selected_customer_id}:\n{e}")
            logging.error(f"Fehler beim Abrufen des Zifferncodes für Kunde ID {self.selected_customer_id}: {e}")
            return
        except Exception as e:
             messagebox.showerror("Systemfehler", f"Ein unerwarteter Fehler ist aufgetreten:\n{e}")
             logging.exception(f"Unerwarteter Fehler beim Start der Dokumentenzuordnung für Kunde ID {self.selected_customer_id}:")
             return


        # 3. Datei zum Verschieben auswählen
        source_file_path = filedialog.askopenfilename(
            title=f"Dokument für Kunde {zifferncode} auswählen",
            initialdir=os.path.expanduser("~") # Start im Benutzerverzeichnis
        )

        if not source_file_path:
            logging.info("Dokumentenauswahl vom Benutzer abgebrochen.")
            return # Benutzer hat den Dialog abgebrochen

        original_file_name = os.path.basename(source_file_path)
        destination_file_name = original_file_name
        destination_path = os.path.join(customer_docs_dir, destination_file_name)

        # 4. Prüfen, ob Datei im Zielordner bereits existiert und ggf. umbenennen
        counter = 1
        name, ext = os.path.splitext(original_file_name)
        while os.path.exists(destination_path):
            destination_file_name = f"{name}_{counter}{ext}"
            destination_path = os.path.join(customer_docs_dir, destination_file_name)
            counter += 1
            if counter > 100: # Schutz vor Endlosschleife
                 messagebox.showerror("Fehler", f"Konnte eindeutigen Dateinamen für '{original_file_name}' nicht finden. Zu viele existierende Kopien.")
                 logging.error(f"Konnte eindeutigen Dateinamen für '{original_file_name}' in '{customer_docs_dir}' nicht finden (counter > 100).")
                 return # Abbruch bei zu vielen Versuchen

        # 5. Datei verschieben und Pfad in Datenbank speichern
        try:
            # Verschieben der Datei
            shutil.move(source_file_path, destination_path)
            logging.info(f"Dokument verschoben von '{source_file_path}' nach '{destination_path}'.")

            # Pfad in der Datenbank speichern
            cursor.execute("INSERT INTO kunden_dokumente (kunde_id, dokument_pfad, dateiname) VALUES (?, ?, ?)",
                           (self.selected_customer_id, destination_path, destination_file_name))
            self.conn.commit()
            messagebox.showinfo("Erfolg", f"Dokument '{original_file_name}' erfolgreich verschoben und als '{destination_file_name}' zu Kunde {zifferncode} zugeordnet.")
            logging.info(f"Dokument '{destination_file_name}' erfolgreich in DB für Kunde ID {self.selected_customer_id} (Code {zifferncode}) gespeichert.")

        except shutil.Error as e:
            messagebox.showerror("Verschieben Fehler", f"Fehler beim Verschieben der Datei:\n{e}\n\nBitte prüfen Sie, ob die Datei '{source_file_path}' geöffnet ist oder Sie die notwendigen Berechtigungen haben.")
            logging.error(f"Fehler beim Verschieben von '{source_file_path}' nach '{destination_path}': {e}")
        except sqlite3.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Speichern des Dokumentenpfads in der Datenbank:\n{e}\n\nDie Datei wurde möglicherweise verschoben, aber der Pfad konnte nicht gespeichert werden.")
            logging.error(f"Fehler beim Speichern des Dokumentenpfads '{destination_path}' für Kunde ID {self.selected_customer_id}: {e}")
            # Loggen, dass die Datei evtl. schon verschoben wurde, aber der DB-Eintrag fehlte
            logging.warning(f"Möglicherweise verschobene Datei unter '{destination_path}', aber DB-Eintrag fehlgeschlagen.")
        except Exception as e:
            messagebox.showerror("Allgemeiner Fehler", f"Ein unerwarteter Fehler ist aufgetreten:\n{e}")
            logging.exception(f"Unerwarteter Fehler in assign_document_to_customer für Kunde ID {self.selected_customer_id}:")


    def show_customer_documents(self):
        if self.selected_customer_id is not None:
            documents_window = tk.Toplevel(self.root)
            documents_window.title(f"Zugeordnete Dokumente für Kunde ID: {self.selected_customer_id}")
            documents_window.geometry("400x300")
            documents_window.transient(self.root)
            documents_window.grab_set()


            cursor = self.conn.cursor()
            try:
                cursor.execute("SELECT id, dateiname, dokument_pfad FROM kunden_dokumente WHERE kunde_id=?", (self.selected_customer_id,))
                documents = cursor.fetchall()

                if documents:
                    listbox_docs = tk.Listbox(documents_window)
                    listbox_docs.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

                    doc_map = {}
                    for doc_id, file_name, file_path in documents:
                        listbox_docs.insert(tk.END, file_name)
                        doc_map[listbox_docs.size() - 1] = (file_path, file_name)

                    def open_selected_document():
                        selected_index = listbox_docs.curselection()
                        if selected_index:
                            file_path, file_name = doc_map[selected_index[0]]
                            try:
                                if os.name == "nt":
                                    os.startfile(file_path)
                                elif os.name == "posix":
                                    # Prüfen, ob der Ordner existiert, bevor xdg-open versucht wird (verhindert hässliche Fehlermeldung)
                                    if os.path.exists(file_path):
                                        subprocess.run(["xdg-open", file_path])
                                    else:
                                         messagebox.showerror("Fehler", f"Datei nicht gefunden: {file_path}")
                                         logging.error(f"Datei nicht gefunden: {file_path}")
                                         # Optional: Eintrag aus DB entfernen, wenn Datei physisch nicht mehr da ist? Komplex.
                                         # Für jetzt nur Fehlermeldung.
                                else:
                                    messagebox.showinfo("Info", f"Datei öffnen: {file_path}")
                                logging.info(f"Dokument '{file_name}' geöffnet: {file_path}")
                            except FileNotFoundError: # Kann auch von xdg-open kommen, wenn es nicht gefunden wird
                                messagebox.showerror("Fehler", f"Programm zum Öffnen der Datei nicht gefunden oder Datei existiert nicht: {file_path}")
                                logging.error(f"Programm zum Öffnen der Datei oder Datei nicht gefunden: {file_path}")
                            except Exception as e:
                                messagebox.showerror("Fehler", f"Fehler beim Öffnen der Datei: {e}")
                                logging.error(f"Fehler beim Öffnen der Datei '{file_name}': {e}")
                        else:
                            messagebox.showwarning("Warnung", "Bitte wählen Sie ein Dokument zum Öffnen aus.")

                    btn_open_doc = ttk.Button(documents_window, text="Ausgewähltes Dokument öffnen", command=open_selected_document)
                    btn_open_doc.pack(pady=5)
                else:
                    ttk.Label(documents_window, text="Keine Dokumente für diesen Kunden gefunden.").pack(padx=10, pady=10)
            except sqlite3.Error as e:
                messagebox.showerror("Fehler", f"Fehler beim Abrufen der Dokumente: {e}")
                logging.error(f"Fehler beim Abrufen der Dokumente: {e}")
        else:
            messagebox.showwarning("Warnung", "Bitte wählen Sie zuerst einen Kunden per Rechtsklick aus.")


    # GEÄNDERT: Logik zum Zuweisen des Zifferncodes für neue Kunden hinzugefügt
    def save_customer(self):
        name = self.entry_name.get().strip()
        vorname = self.entry_vorname.get().strip()
        geburtsdatum_str = self.entry_geburtsdatum.get().strip()
        strasse = self.entry_strasse.get().strip()
        hausnummer = self.entry_hausnummer.get().strip()
        plz = self.entry_plz.get().strip()
        ort = self.entry_ort.get().strip()
        telefon = self.entry_telefon.get().strip()

        if not (name and vorname and geburtsdatum_str):
            messagebox.showwarning("Fehler", "Name, Vorname und Geburtsdatum dürfen nicht leer sein!")
            return

        try:
            datetime.strptime(geburtsdatum_str, "%d.%m.%Y")
        except ValueError:
            messagebox.showwarning("Fehler", "Bitte ein gültiges Geburtsdatum im Format TT.MM.JJJJ eingeben.")
            return

        cursor = self.conn.cursor()
        try:
            if self.selected_customer_for_edit:
                # --- START DER NEUEN LOGIK (Warnmeldung vor Änderung) ---
                confirm_update = messagebox.askyesno(
                    "Warnung",
                    "Achtung! Sie sind gerade dabei, die Daten eines Bestehenden Kunden zu ändern! Möchten Sie dies wirklich tun?",
                    default=messagebox.NO, # Standardmäßig 'Nein' auswählen
                    icon=messagebox.WARNING # Ein Warnsymbol anzeigen
                )

                if not confirm_update: # Wenn der Benutzer 'Nein' klickt
                    messagebox.showinfo("Abgebrochen", "Änderungen wurden nicht gespeichert.") # Optional: Info, dass abgebrochen wurde
                    logging.info(f"Änderungen an Kunde ID {self.selected_customer_for_edit} vom Benutzer abgebrochen.")
                    return # Die Methode hier beenden, ohne zu speichern

                # --- ENDE DER NEUEN LOGIK ---

                # Aktualisiere bestehenden Kunden (Zifferncode wird NICHT geändert)
                # Dieser Teil wird nur ausgeführt, wenn der Benutzer auf 'Ja' geklickt hat
                cursor.execute(
                    "UPDATE kunden SET name=?, vorname=?, geburtsdatum=?, strasse=?, hausnummer=?, plz=?, ort=?, telefon=? WHERE id=?",
                    (name, vorname, geburtsdatum_str, strasse, hausnummer, plz, ort, telefon, self.selected_customer_for_edit)
                )
                self.conn.commit()
                messagebox.showinfo("Erfolg", "Kundendaten erfolgreich geändert!")
                logging.info(f"Kundendaten für ID {self.selected_customer_for_edit} erfolgreich geändert.")
                # BEI ÄNDERUNG: KEINEN NEUEN ORDNER ERSTELLEN, der Ordner existiert bereits mit dem Zifferncode
                self.selected_customer_for_edit = None # Auswahl aufheben
                self.btn_save_kunde.config(text="Kunde speichern") # Button-Text zurücksetzen
            else:
                # Füge neuen Kunden hinzu
                # NEU: Nächsten Zifferncode ermitteln
                cursor.execute("SELECT MAX(zifferncode) FROM kunden")
                max_code_result = cursor.fetchone()
                next_zifferncode = 101 # Standard für den ersten Kunden
                if max_code_result and max_code_result[0] is not None:
                    # Sicherstellen, dass der nächste Code >= 101 ist
                    next_zifferncode = max(101, max_code_result[0] + 1)

                # NEU: zifferncode in INSERT-Statement aufnehmen
                cursor.execute(
                    "INSERT INTO kunden (name, vorname, geburtsdatum, strasse, hausnummer, plz, ort, telefon, zifferncode) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (name, vorname, geburtsdatum_str, strasse, hausnummer, plz, ort, telefon, next_zifferncode)
                )
                self.conn.commit()

                # NEU: Kundenordner nach erfolgreichem Speichern erstellen
                # Erst die ID des neu eingefügten Kunden holen
                new_customer_id = cursor.lastrowid
                if new_customer_id:
                     # Sicherstellen, dass wir den Zifferncode des *gerade* eingefügten Kunden haben
                     self.create_customer_document_folder(new_customer_id, next_zifferncode) # Rufe neue Methode auf
                else:
                     logging.error("Konnte ID des neu eingefügten Kunden nicht abrufen, Ordner wurde nicht erstellt.")
                     messagebox.showwarning("Ordner Fehler", "Konnte den Kundenordner nicht erstellen (ID nicht gefunden).")

                messagebox.showinfo("Erfolg", f"Kunde erfolgreich gespeichert! Zugewiesener Zifferncode: {next_zifferncode}")
                logging.info(f"Kunde '{vorname} {name}' erfolgreich gespeichert mit Zifferncode {next_zifferncode}.")

            self.load_customers()
            self.clear_fields()
        except sqlite3.IntegrityError as e:
             # Fängt den Fall ab, dass der Zifferncode (durch einen seltenen Fehler) schon existiert
             messagebox.showerror("Fehler", f"Fehler beim Speichern: Zifferncode konnte nicht eindeutig zugewiesen werden. Bitte erneut versuchen.\n({e})")
             logging.error(f"IntegrityError beim Speichern des Kunden (möglicherweise Zifferncode-Konflikt): {e}")
             self.conn.rollback() # Änderungen rückgängig machen
        except sqlite3.Error as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern der Kundendaten: {e}")
            logging.error(f"Fehler beim Speichern der Kundendaten: {e}")
            self.conn.rollback()

    def delete_customer(self):
        selected_item = self.tree.selection()
        if selected_item:
            customer_id = self.tree.item(selected_item[0])["values"][0]
            customer_name = f"{self.tree.item(selected_item[0])['values'][3]} {self.tree.item(selected_item[0])['values'][2]}" # Vorname Nachname für die Warnung

            # Warnmeldung anzeigen mit Ja/Nein Optionen
            confirm_delete = messagebox.askyesno(
                "Warnung",
                f"Möchten Sie den Kunden '{customer_name}' wirklich löschen? Gelöschte Daten können nicht wiederhergestellt werden!\n\nKlicken Sie auf 'Ja' zum Löschen oder auf 'Nein' zum Abbrechen.",
                default=messagebox.NO,
                icon=messagebox.WARNING
            )

            if confirm_delete is True:  # Nur bei "Ja" (True) löschen
                cursor = self.conn.cursor()
                try:
                    # 1. Zugeordnete Dokumentpfade abrufen
                    cursor.execute("SELECT dokument_pfad FROM kunden_dokumente WHERE kunde_id=?", (customer_id,))
                    document_paths = [row[0] for row in cursor.fetchall()]

                    # 2. Versuchen, die Dokumentdateien zu löschen
                    # NEU: Auch den Kundenordner löschen, wenn er leer ist oder alle Dateien gelöscht wurden
                    try:
                         cursor.execute("SELECT zifferncode FROM kunden WHERE id=?", (customer_id,))
                         zifferncode_data = cursor.fetchone()
                         customer_folder_path = None
                         if zifferncode_data and zifferncode_data[0] is not None and self.document_base_path:
                             main_docs_dir = os.path.join(self.document_base_path, "Kundendokumente")
                             customer_folder_path = os.path.join(main_docs_dir, str(zifferncode_data[0]))

                         # Lösche die einzelnen Dateien
                         for path in document_paths:
                             try:
                                 if os.path.exists(path):
                                     os.remove(path)
                                     logging.info(f"Dokument gelöscht: {path}")
                                 else:
                                      logging.warning(f"Dokument nicht gefunden (konnte nicht gelöscht werden): {path}") # Warnt, falls Datei fehlt
                             except PermissionError:
                                 messagebox.showerror("Fehler", f"Keine Berechtigung zum Löschen der Datei: {path}")
                                 logging.error(f"Keine Berechtigung zum Löschen der Datei: {path}")
                                 # Hier NICHT return, da die Datenbankeinträge noch gelöscht werden müssen!
                                 # Der Benutzer muss die Datei manuell löschen.
                                 pass # Fahre mit dem Löschen der anderen Dateien und DB-Einträge fort
                             except Exception as e:
                                  messagebox.showerror("Fehler", f"Fehler beim Löschen der Datei {os.path.basename(path)}: {e}")
                                  logging.error(f"Fehler beim Löschen der Datei '{path}': {e}")
                                  pass # Fahre fort


                         # Versuch, den Kundenordner zu löschen, wenn er existiert und leer ist
                         if customer_folder_path and os.path.exists(customer_folder_path):
                            try:
                                # Prüfen, ob der Ordner leer ist, indem man listdir() versucht
                                if not os.listdir(customer_folder_path):
                                    os.rmdir(customer_folder_path) # Nur leere Ordner löschen
                                    logging.info(f"Leerer Kundenordner gelöscht: {customer_folder_path}")
                                else:
                                     logging.warning(f"Kundenordner '{customer_folder_path}' ist nicht leer und wird nicht gelöscht.")
                            except OSError as e:
                                 # Dieser Fehler tritt auf, wenn der Ordner nicht leer ist (trotzdem listdir() bestanden)
                                 # oder wenn Berechtigungen fehlen.
                                 logging.error(f"Fehler beim Löschen des Kundenordners '{customer_folder_path}': {e}")
                                 # messagebox.showwarning("Ordner Fehler", f"Konnte Kundenordner '{zifferncode_data[0]}' nicht löschen:\n{e}") # Zu aufdringlich? Nur loggen
                            except Exception as e:
                                logging.error(f"Unerwarteter Fehler beim Löschen des Kundenordners '{customer_folder_path}': {e}")


                    except sqlite3.Error as e:
                         logging.error(f"Fehler beim Abrufen des Zifferncodes für Kundenordner-Löschung (ID {customer_id}): {e}")
                         # Weiter machen, da die Dokumente selbst und DB-Einträge wichtiger sind
                    except Exception as e:
                         logging.exception(f"Unerwarteter Fehler während Dokument-/Ordnerlöschung für Kunde ID {customer_id}:")


                    # 3. Einträge aus den Datenbanktabellen löschen
                    cursor.execute("DELETE FROM kunden_dokumente WHERE kunde_id=?", (customer_id,))
                    # GEÄNDERT: Historie auch löschen, bevor Pfandscheine gelöscht werden
                    cursor.execute("DELETE FROM pfandschein_historie WHERE pfandschein_id IN (SELECT id FROM pfandscheine WHERE kunde_id=?)", (customer_id,))
                    cursor.execute("DELETE FROM pfandscheine WHERE kunde_id=?", (customer_id,))
                    cursor.execute("DELETE FROM kunden WHERE id=?", (customer_id,))

                    self.conn.commit()
                    messagebox.showinfo("Erfolg", f"Kunde '{customer_name}' und zugeordnete Daten gelöscht!")
                    self.load_customers()
                    logging.info(f"Kunde mit ID {customer_id} ('{customer_name}') und zugeordnete Daten gelöscht.")
                except sqlite3.Error as e:
                    messagebox.showerror("Fehler", f"Fehler beim Löschen des Kunden: {e}")
                    logging.error(f"Fehler beim Löschen des Kunden mit ID {customer_id}: {e}")
                    self.conn.rollback() # Änderungen rückgängig machen bei DB-Fehler
            else:
                # Benutzer hat "Nein" (Abbrechen) geklickt
                messagebox.showinfo("Abgebrochen", "Kunde wurde nicht gelöscht.")
                logging.info(f"Löschen von Kunde mit ID {customer_id} vom Benutzer abgebrochen.")
        else:
            messagebox.showwarning("Fehler", "Kein Kunde ausgewählt!")

    def clear_fields(self):
        self.entry_name.delete(0, tk.END)
        self.entry_vorname.delete(0, tk.END)
        self.entry_geburtsdatum.delete(0, tk.END)
        self.entry_strasse.delete(0, tk.END)
        self.entry_hausnummer.delete(0, tk.END)
        self.entry_plz.delete(0, tk.END)
        self.entry_ort.delete(0, tk.END)
        self.entry_telefon.delete(0, tk.END)
        self.selected_customer_for_edit = None # Auswahl aufheben
        self.btn_save_kunde.config(text="Kunde speichern") # Button-Text zurücksetzen

    # GEÄNDERT: zifferncode in SELECT und WHERE Klausel aufgenommen
    def load_customers(self, query=None, zifferncode_query=None): # NEU: zifferncode_query Parameter
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.selected_customer_for_edit = None # Auswahl aufheben beim Neuladen
        self.btn_save_kunde.config(text="Kunde speichern") # Button-Text zurücksetzen

        cursor = self.conn.cursor()
        try:
            if zifferncode_query: # Priorität für Zifferncode-Suche
                cursor.execute("""
                    SELECT id, zifferncode, name, vorname, geburtsdatum, strasse, hausnummer, plz, ort, telefon
                    FROM kunden WHERE zifferncode = ?
                """, (zifferncode_query,))
            elif query: # Fallback zur allgemeinen Suche
                like_query = f"%{query}%"
                # Suche auch nach Zifferncode
                cursor.execute("""
                    SELECT k.id, k.zifferncode, k.name, k.vorname, k.geburtsdatum, k.strasse, k.hausnummer, k.plz, k.ort, k.telefon
                    FROM kunden k
                    WHERE k.name LIKE ? OR k.vorname LIKE ? OR k.geburtsdatum LIKE ? OR k.strasse LIKE ? OR k.hausnummer LIKE ? OR k.plz LIKE ? OR k.ort LIKE ? OR k.telefon LIKE ? OR CAST(k.zifferncode AS TEXT) LIKE ?
                """, (like_query,)*8 + (like_query,)) # 9 Platzhalter für 9 LIKE-Bedingungen
            else: # Alle laden
                cursor.execute("""
                    SELECT id, zifferncode, name, vorname, geburtsdatum, strasse, hausnummer, plz, ort, telefon
                    FROM kunden ORDER BY id ASC
                """) # Sortiert nach ID

            for row in cursor.fetchall():
                # Stelle sicher, dass zifferncode als String oder leer angezeigt wird
                display_row = list(row)
                if display_row[1] is None: # Index 1 ist zifferncode
                    display_row[1] = ""
                self.tree.insert("", "end", values=display_row)
            logging.info(f"Kundenliste {'(gefiltert nach Zifferncode)' if zifferncode_query else ('(gefiltert)' if query else '')} geladen.")
        except sqlite3.Error as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Kunden: {e}")
            logging.error(f"Fehler beim Laden der Kunden: {e}")

    def search_customers(self):
        search_term = self.entry_search.get().strip()
        self.load_customers(query=search_term)

    def search_customers_by_zifferncode(self): # NEU: Funktion für Zifferncode-Suche
        zifferncode_search_term = self.entry_zifferncode_search.get().strip()
        if zifferncode_search_term:
            try:
                zifferncode = int(zifferncode_search_term) # Versuche in Integer zu konvertieren
                self.load_customers(zifferncode_query=zifferncode) # Suche nur nach Zifferncode
                if self.hands_free_zifferncode_search.get(): # NEU: Hands-free Modus aktiv?
                    items = self.tree.get_children()
                    if len(items) == 1: # Nur wenn genau ein Ergebnis
                        selected_item_id = items[0]
                        customer_id = self.tree.item(selected_item_id)["values"][0]
                        self.load_customer_data_for_edit(customer_id) # Daten laden und Kunde auswählen
                        logging.info(f"Hands-free Zifferncode-Suche: Kunde ID {customer_id} automatisch geladen.")

            except ValueError:
                messagebox.showwarning("Warnung", "Bitte geben Sie eine gültige Zahl für den Zifferncode ein.")
        else:
            self.load_customers() # Lade alle Kunden, wenn das Feld leer ist

    # NEU: Funktion für Enter-Taste in Zifferncode-Suche
    def zifferncode_search_enter_pressed(self, event):
        if self.hands_free_zifferncode_search.get(): # Nur wenn Hands-free aktiviert ist
            self.search_customers_by_zifferncode()
        else: # Wenn Hands-free aus, normale Suche per Button-Click simulieren
            self.search_customers_by_zifferncode() # Verhält sich wie der Button-Click

    # NEU: Funktion für Enter-Taste in Volltext-Suche
    def fulltext_search_enter_pressed(self, event):
        self.search_customers() # Simuliert den Button-Click für Volltextsuche

    def on_customer_double_click(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            customer_id = self.tree.item(selected_item[0])["values"][0]
            self.open_pfandschein_window(customer_id)

    def open_change_db_path_window(self):
        change_path_window = tk.Toplevel(self.root)
        change_path_window.title("Datenbankpfad ändern")
        change_path_window.transient(self.root)
        change_path_window.grab_set()

        ttk.Label(change_path_window, text="Neuer Datenbankpfad:").pack(padx=10, pady=10)
        self.entry_new_db_path = ttk.Entry(change_path_window, width=40)
        self.entry_new_db_path.pack(padx=10, pady=5)
        self.entry_new_db_path.insert(0, self.get_db_path())

        # Button zum Auswählen des Pfads über einen Dateidialog
        btn_browse = ttk.Button(change_path_window, text="Durchsuchen...", command=self.browse_db_path_db)
        btn_browse.pack(pady=5)

        btn_save = ttk.Button(change_path_window, text="Speichern und neu verbinden", command=self.save_new_db_path)
        btn_save.pack(pady=10)

        btn_cancel = ttk.Button(change_path_window, text="Abbrechen", command=change_path_window.destroy)
        btn_cancel.pack(pady=5)

    def browse_db_path_db(self):
        new_path = filedialog.asksaveasfilename(
            parent=self.entry_new_db_path.winfo_toplevel(), # Übergeordnetes Fenster explizit angeben
            initialdir=os.path.dirname(self.get_db_path()),
            defaultextension=".db",
            filetypes=[("SQLite Datenbanken", "*.db"), ("Alle Dateien", "*.*")]
        )
        if new_path:
            self.entry_new_db_path.delete(0, tk.END)
            self.entry_new_db_path.insert(0, new_path) # Pfad sofort ins Feld schreiben

    def save_new_db_path(self):
        new_path = self.entry_new_db_path.get().strip()
        if new_path:
            try:
                with open("config.txt", "r+") as f:
                    lines = f.readlines()
                    found = False
                    f.seek(0)
                    for line in lines:
                        if line.startswith("db_path="):
                            f.write(f"db_path={new_path}\n")
                            found = True
                        else:
                            f.write(line)
                    if not found:
                        f.write(f"db_path={new_path}\n")
                    f.truncate()
                self.db_path = new_path
                self.conn.close()
                self.conn = self.connect_db()
                if self.conn:
                    self.load_customers()
                    messagebox.showinfo("Erfolg", f"Datenbankpfad geändert und gespeichert: {self.db_path}")
                    logging.info(f"Datenbankpfad geändert und gespeichert: {self.db_path}")
                else:
                    messagebox.showerror("Fehler", "Die Verbindung zur neuen Datenbank konnte nicht hergestellt werden. Überprüfen Sie die Logdatei.")
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler beim Speichern des Datenbankpfads: {e}")
                logging.error(f"Fehler beim Speichern des Datenbankpfads: {e}")
        else:
            messagebox.showwarning("Warnung", "Bitte geben Sie einen gültigen Datenbankpfad ein.")


    # --- open_pfandschein_window ---
    # GEÄNDERT: Ruft show_pdf und show_pfandschein_history mit zifferncode auf
    def open_pfandschein_window(self, customer_id):
        # Neues Toplevel-Fenster für die Pfandschein-Verwaltung
        top = tk.Toplevel(self.root)
        top.title("Pfandschein Verwaltung für Kunde ID: " + str(customer_id))
        top.geometry("1200x550")  # Breite angepasst, Höhe unverändert
        top.transient(self.root) # Fenster über Hauptfenster halten
        top.grab_set() # Modales Fenster

        # Linker Bereich: Listbox mit allen Pfandscheinen dieses Kunden (unverändert)
        left_frame = ttk.Frame(top)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        ttk.Label(left_frame, text="Bestehende Pfandscheine:").pack()

        pf_listbox = tk.Listbox(left_frame, width=40)
        pf_listbox.pack(side=tk.LEFT, fill=tk.Y)
        lb_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=pf_listbox.yview)
        pf_listbox.configure(yscrollcommand=lb_scroll.set)
        lb_scroll.pack(side=tk.LEFT, fill=tk.Y)

        # Abfrage aller Pfandscheine des Kunden (sortiert nach ID) - (unverändert)
        cursor = self.conn.cursor()
        pfandscheine = []
        pf_map = {}
        try:
            cursor.execute("SELECT id, kunde_id, abschlusstag, verfalltag, darlehen, monatl_zinsen, monatl_kosten, versicherungssumme, vertragsnummer, artikel_beschreibung FROM pfandscheine WHERE kunde_id=? ORDER BY id ASC", (customer_id,))
            pfandscheine = cursor.fetchall()
            for i, pf in enumerate(pfandscheine):
                summary = f"ID: {pf[0]}, Abschlusstag: {pf[2]}"
                pf_listbox.insert(tk.END, summary)
                pf_map[i] = pf
            logging.info(f"Pfandscheine für Kunde ID {customer_id} geladen.")
        except sqlite3.Error as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Pfandscheine: {e}")
            logging.error(f"Fehler beim Laden der Pfandscheine für Kunde ID {customer_id}: {e}")

        # Mittlerer Bereich: Detailfelder für Pfandschein-Daten (unverändert)
        middle_frame = ttk.Frame(top)
        middle_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        detail_labels = [
            "Abschlusstag (TT.MM.JJJJ):",
            "Verfalltag (TT.MM.JJJJ):",
            "Darlehen:",
            "Monatl. Zinsen:",
            "Monatl. Kosten:",
            "Versicherungssumme:",
            "Nr. des früheren Vertrages:",
            "Artikelbeschreibung:",
        ]
        detail_entries = {}
        grid_row_offset = 0
        for i, lab in enumerate(detail_labels):
            ttk.Label(middle_frame, text=lab).grid(row=i, column=0, sticky=tk.E, padx=5, pady=5)
            ent = ttk.Entry(middle_frame, width=40)
            ent.grid(row=i, column=1, padx=5, pady=5)
            detail_entries[lab] = ent
            grid_row_offset = i

        ttk.Label(middle_frame, text="Zinseinheit:").grid(row=3, column=2, sticky=tk.W, padx=5, pady=5)
        self.zins_einheit = ttk.Combobox(middle_frame, values=["%", "€"], width=5)
        self.zins_einheit.grid(row=3, column=3, padx=5, pady=5)
        self.zins_einheit.set("%")

        if self.last_zins_einheit and self.last_zins_einheit != "%":
             self.zins_einheit.set(self.last_zins_einheit)
        elif self.last_zins_einheit == "%":
            self.zins_einheit.set("%")
        elif self.last_zins_einheit == "€":
             self.zins_einheit.set("€")
        else:
             self.zins_einheit.set("%")

        def save_current_zins_einheit(event):
            selected_einheit = self.zins_einheit.get()
            self.save_last_zins_einheit(selected_einheit)

        self.zins_einheit.bind("<<ComboboxSelected>>", save_current_zins_einheit)

        selected_pf = [None]

        def populate_details(pf):
            detail_entries["Abschlusstag (TT.MM.JJJJ):"].delete(0, tk.END)
            detail_entries["Abschlusstag (TT.MM.JJJJ):"].insert(0, pf[2] if pf[2] else "")
            detail_entries["Verfalltag (TT.MM.JJJJ):"].delete(0, tk.END)
            detail_entries["Verfalltag (TT.MM.JJJJ):"].insert(0, pf[3] if pf[3] else "")
            detail_entries["Darlehen:"].delete(0, tk.END)
            detail_entries["Darlehen:"].insert(0, str(pf[4]) if pf[4] is not None else "")
            detail_entries["Monatl. Zinsen:"].delete(0, tk.END)
            detail_entries["Monatl. Zinsen:"].insert(0, str(pf[5]) if pf[5] is not None else "")
            detail_entries["Monatl. Kosten:"].delete(0, tk.END)
            detail_entries["Monatl. Kosten:"].insert(0, str(pf[6]) if pf[6] is not None else "")
            detail_entries["Versicherungssumme:"].delete(0, tk.END)
            detail_entries["Versicherungssumme:"].insert(0, pf[7] if pf[7] is not None else "")
            detail_entries["Nr. des früheren Vertrages:"].delete(0, tk.END)
            detail_entries["Nr. des früheren Vertrages:"].insert(0, pf[8] if pf[8] is not None else "")
            detail_entries["Artikelbeschreibung:"].delete(0, tk.END)
            detail_entries["Artikelbeschreibung:"].insert(0, pf[9] if pf[9] is not None else "")

        def on_pf_select(event):
            selection = pf_listbox.curselection()
            if selection:
                index = int(selection[0])
                pf = pf_map.get(index)
                if pf:
                    populate_details(pf)
                    selected_pf[0] = pf

        pf_listbox.bind("<<ListboxSelect>>", on_pf_select)

        def create_new_pf():
            abschlusstag = detail_entries["Abschlusstag (TT.MM.JJJJ):"].get().strip()
            verfalltag = detail_entries["Verfalltag (TT.MM.JJJJ):"].get().strip()
            darlehen_str = detail_entries["Darlehen:"].get().strip()
            monatl_zinsen_str = detail_entries["Monatl. Zinsen:"].get().strip()  # Werte aus den Detailfeldern holen
            monatl_kosten_str = detail_entries["Monatl. Kosten:"].get().strip() # Werte aus den Detailfeldern holen
            versicherungssumme = detail_entries["Versicherungssumme:"].get().strip()
            vertragsnummer = detail_entries["Nr. des früheren Vertrages:"].get().strip()
            artikel_beschreibung = detail_entries["Artikelbeschreibung:"].get().strip()
            zinseinheit = self.zins_einheit.get()

            # --- NEU: Überprüfung, ob Zinsen und Kosten befüllt sind ---
            if not monatl_zinsen_str or not monatl_kosten_str:
                messagebox.showwarning("Warnung", "Bitte berechnen Sie zuerst die monatlichen Zinsen und Kosten mit dem Kostenrechner und übernehmen Sie diese, oder geben Sie die Werte manuell ein, bevor Sie den Pfandschein anlegen.")
                return  # Abbruch, wenn Felder leer sind

            try:
                datetime.strptime(abschlusstag, "%d.%m.%Y")
                datetime.strptime(verfalltag, "%d.%m.%Y")
            except ValueError:
                messagebox.showwarning("Fehler", "Bitte gültige Datumsangaben für Abschlusstag und Verfalltag im Format TT.MM.JJJJ eingeben.")
                return
            try:
                darlehen = float(darlehen_str)
                monatl_zinsen = float(monatl_zinsen_str)
                monatl_kosten = float(monatl_kosten_str)
            except ValueError:
                messagebox.showwarning("Fehler", "Darlehen, Monatl. Zinsen und Monatl. Kosten müssen Zahlen sein.")
                return
            try:
                cursor.execute("""
                    INSERT INTO pfandscheine (kunde_id, abschlusstag, verfalltag, darlehen, monatl_zinsen, monatl_kosten, versicherungssumme, vertragsnummer, artikel_beschreibung)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (customer_id, abschlusstag, verfalltag, darlehen, monatl_zinsen, monatl_kosten, versicherungssumme, vertragsnummer, artikel_beschreibung))
                self.conn.commit()
                messagebox.showinfo("Erfolg", "Neuer Pfandschein angelegt!")
                logging.info(f"Neuer Pfandschein für Kunde ID {customer_id} angelegt (Zinseinheit: {zinseinheit}, Artikel: {artikel_beschreibung[:30]}...).")
                refresh_pf_list()
                # --- NEU: Felder im Pfandschein-Formular leeren nach dem Speichern ---
                populate_details(['', '', '', '', '', '', '', '', '', '']) # Leere Felder anzeigen
            except sqlite3.Error as e:
                messagebox.showerror("Fehler", f"Fehler beim Anlegen des neuen Pfandscheins: {e}")
                logging.error(f"Fehler beim Anlegen des neuen Pfandscheins für Kunde ID {customer_id}: {e}")


        def update_pf():
            if not selected_pf[0]:
                messagebox.showwarning("Fehler", "Bitte wähle einen Pfandschein aus!")
                return
            abschlusstag = detail_entries["Abschlusstag (TT.MM.JJJJ):"].get().strip()
            verfalltag = detail_entries["Verfalltag (TT.MM.JJJJ):"].get().strip()
            darlehen_str = detail_entries["Darlehen:"].get().strip()
            monatl_zinsen_str = detail_entries["Monatl. Zinsen:"].get().strip()
            monatl_kosten_str = detail_entries["Monatl. Kosten:"].get().strip()
            versicherungssumme = detail_entries["Versicherungssumme:"].get().strip()
            vertragsnummer = detail_entries["Nr. des früheren Vertrages:"].get().strip()
            artikel_beschreibung = detail_entries["Artikelbeschreibung:"].get().strip()
            zinseinheit = self.zins_einheit.get()

            try:
                datetime.strptime(abschlusstag, "%d.%m.%Y")
                datetime.strptime(verfalltag, "%d.%m.%Y")
            except ValueError:
                messagebox.showwarning("Fehler", "Bitte gültige Datumsangaben für Abschlusstag und Verfalltag im Format TT.MM.JJJJ eingeben.")
                return
            try:
                darlehen = float(darlehen_str)
                monatl_zinsen = float(monatl_zinsen_str)
                monatl_kosten = float(monatl_kosten_str)
            except ValueError:
                messagebox.showwarning("Fehler", "Darlehen, Monatl. Zinsen und Monatl. Kosten müssen Zahlen sein.")
                return

            try:
                cursor = self.conn.cursor()
                cursor.execute("SELECT id, kunde_id, abschlusstag, verfalltag, darlehen, monatl_zinsen, monatl_kosten, versicherungssumme, vertragsnummer, artikel_beschreibung FROM pfandscheine WHERE id=?", (selected_pf[0][0],))
                vorherige_version = cursor.fetchone()
                if vorherige_version:
                    cursor.execute("""
                        INSERT INTO pfandschein_historie (pfandschein_id, abschlusstag, verfalltag, darlehen, monatl_zinsen, monatl_kosten, versicherungssumme, vertragsnummer, artikel_beschreibung, aenderungsdatum)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (vorherige_version[0], vorherige_version[2], vorherige_version[3], vorherige_version[4], vorherige_version[5], vorherige_version[6], vorherige_version[7], vorherige_version[8], vorherige_version[9], datetime.now().strftime("%d.%m.%Y %H:%M:%S")))
                    self.conn.commit()
                    logging.info(f"Pfandschein ID {selected_pf[0][0]} in die Historie verschoben (Zinseinheit: {zinseinheit}).")

                cursor.execute("""
                    UPDATE pfandscheine
                    SET abschlusstag=?, verfalltag=?, darlehen=?, monatl_zinsen=?, monatl_kosten=?, versicherungssumme=?, vertragsnummer=?, artikel_beschreibung=?
                    WHERE id=?
                """, (abschlusstag, verfalltag, darlehen, monatl_zinsen, monatl_kosten, versicherungssumme, vertragsnummer, artikel_beschreibung, selected_pf[0][0]))
                self.conn.commit()
                messagebox.showinfo("Erfolg", "Pfandschein verlängert!")
                logging.info(f"Pfandschein ID {selected_pf[0][0]} verlängert/geändert (Zinseinheit: {zinseinheit}, Artikel: {artikel_beschreibung[:30]}...).")
                refresh_pf_list()
            except sqlite3.Error as e:
                messagebox.showerror("Fehler", f"Fehler beim Verlängern/Ändern des Pfandscheins: {e}")
                logging.error(f"Fehler beim Verlängern/Ändern des Pfandscheins ID {selected_pf[0][0]}: {e}")

        def refresh_pf_list():
            pf_listbox.delete(0, tk.END)
            try:
                cursor.execute("SELECT id, kunde_id, abschlusstag, verfalltag, darlehen, monatl_zinsen, monatl_kosten, versicherungssumme, vertragsnummer, artikel_beschreibung FROM pfandscheine WHERE kunde_id=? ORDER BY id ASC", (customer_id,))
                new_pfandscheine = cursor.fetchall()
                pf_map.clear()
                for i, pf in enumerate(new_pfandscheine):
                    summary = f"ID: {pf[0]}, Abschlusstag: {pf[2]}"
                    pf_listbox.insert(tk.END, summary)
                    pf_map[i] = pf
                # NEU: Zum Ende der Liste scrollen, um neuen Eintrag sichtbar zu machen
                pf_listbox.see(tk.END)
            except sqlite3.Error as e:
                messagebox.showerror("Fehler", f"Fehler beim Aktualisieren der Pfandscheinliste: {e}") # Auch hier Fehlermeldung zeigen
                logging.error(f"Fehler beim Aktualisieren der Pfandscheinliste für Kunde ID {customer_id}: {e}")


        # GEÄNDERT: Ruft generate_pdf mit zifferncode auf
        def show_pdf():
            if not selected_pf[0]:
                messagebox.showwarning("Fehler", "Bitte wähle einen Pfandschein aus!")
                return
            try:
                # Pfandschein-Daten holen
                cursor.execute("SELECT id, kunde_id, abschlusstag, verfalltag, darlehen, monatl_zinsen, monatl_kosten, versicherungssumme, vertragsnummer, artikel_beschreibung FROM pfandscheine WHERE id=? LIMIT 1", (selected_pf[0][0],))
                pf = cursor.fetchone()
                if not pf:
                    messagebox.showwarning("Fehler", "Pfandschein nicht gefunden!")
                    return

                # Kunden-Daten holen (inkl. zifferncode)
                kunde_cursor = self.conn.cursor()
                kunde_cursor.execute("SELECT name, vorname, zifferncode FROM kunden WHERE id=?", (pf[1],)) # pf[1] ist kunde_id
                kunde = kunde_cursor.fetchone()

                if kunde:
                    pf_data = {
                        "Pfandschein-ID": pf[0],
                        "Kunden-Name": f"{kunde[1]} {kunde[0]}", # Vorname Nachname
                        "Zifferncode": kunde[2], # NEU
                        "Abschlusstag": pf[2],
                        "Verfalltag": pf[3],
                        "Darlehen": f"{pf[4]:.2f} €",
                        "Monatl. Zinsen": f"{pf[5]:.2f} {self.zins_einheit.get()}",
                        "Monatl. Kosten": f"{pf[6]:.2f} €",
                        "Versicherungssumme": pf[7],
                        "Vertragsnummer": pf[8],
                        "Artikelbeschreibung": pf[9]
                    }
                    self.generate_pdf_with_background(pf_data, self.pdf_background_path)
                    logging.info(f"PDF für Pfandschein ID {pf[0]} erstellt (Zifferncode: {kunde[2]}, Zinseinheit: {self.zins_einheit.get()}).")
                else:
                    messagebox.showerror("Fehler", "Kunde für diesen Pfandschein nicht gefunden!")
                    logging.error(f"Kunde für Pfandschein ID {pf[0]} nicht gefunden.")
            except sqlite3.Error as e:
                messagebox.showerror("Fehler", f"Fehler beim Abrufen der Pfandscheindaten für PDF: {e}")
                logging.error(f"Fehler beim Abrufen der Pfandscheindaten für PDF: {e}")
            # NEU: Allgemeiner Fehler-Catch für PDF-Generierung selbst
            except Exception as e:
                messagebox.showerror("PDF Fehler", f"Ein unerwarteter Fehler ist beim Erstellen des PDFs aufgetreten:\n{type(e).__name__}: {e}\n\nDetails finden Sie in der Log-Datei.")
                logging.exception(f"Unerwarteter Fehler in show_pdf für Pfandschein ID {selected_pf[0][0] if selected_pf[0] else 'N/A'}:")


        # GEÄNDERT: Ruft print_pdf_from_history mit zifferncode auf
        def show_pfandschein_history():
            if not selected_pf[0]:
                messagebox.showwarning("Fehler", "Bitte wählen Sie zuerst einen Pfandschein aus.")
                return

            pfandschein_id = selected_pf[0][0]
            history_window = tk.Toplevel(top)
            history_window.title(f"Historie für Pfandschein ID: {pfandschein_id}")
            history_window.geometry("1800x600")
            history_window.transient(top) # Fenster über Pfandschein-Fenster halten
            history_window.grab_set() # Modales Fenster

            columns = ("Aenderungsdatum", "Abschlusstag", "Verfalltag", "Darlehen", "Zinsen", "Kosten", "Versicherung", "Vertragsnummer", "Artikel")
            self.tree_history = ttk.Treeview(history_window, columns=columns, show="headings") # Self Attribut für Context Menu
            self.tree_history.heading("Aenderungsdatum", text="Änderungsdatum")
            self.tree_history.heading("Abschlusstag", text="Abschlusstag")
            self.tree_history.heading("Verfalltag", text="Verfalltag")
            self.tree_history.heading("Darlehen", text="Darlehen")
            self.tree_history.heading("Zinsen", text="Zinsen")
            self.tree_history.heading("Kosten", text="Kosten")
            self.tree_history.heading("Versicherung", text="Versicherung")
            self.tree_history.heading("Vertragsnummer", text="Vertragsnummer")
            self.tree_history.heading("Artikel", text="Artikel")
            self.tree_history.pack(fill=tk.BOTH, expand=True)

            self.tree_history.column("Aenderungsdatum", width=140)
            self.tree_history.column("Artikel", width=200)

            try:
                cursor.execute("SELECT aenderungsdatum, abschlusstag, verfalltag, darlehen, monatl_zinsen, monatl_kosten, versicherungssumme, vertragsnummer, artikel_beschreibung FROM pfandschein_historie WHERE pfandschein_id=? ORDER BY aenderungsdatum DESC", (pfandschein_id,))
                for row in cursor.fetchall():
                    self.tree_history.insert("", "end", values=row)
                logging.info(f"Historie für Pfandschein ID {pfandschein_id} angezeigt.")
            except sqlite3.Error as e:
                messagebox.showerror("Fehler", f"Fehler beim Laden der Pfandscheinhistorie: {e}")
                logging.error(f"Fehler beim Laden der Pfandscheinhistorie für Pfandschein ID {pfandschein_id}: {e}")

            # Context Menu für History Treeview BINDEN
            self.tree_history.bind("<Button-3>", lambda event: show_history_context_menu(event)) # Korrigierter Aufruf

            close_button = ttk.Button(history_window, text="Schließen", command=history_window.destroy)
            close_button.pack(pady=10)

        # GEÄNDERT: Holt zifferncode für historisches PDF

        # GEÄNDERT: Holt zifferncode für historisches PDF
        def print_pdf_from_history():
            selected_item = self.tree_history.selection()
            if selected_item:
                history_data = self.tree_history.item(selected_item[0])["values"]
                if history_data and len(history_data) == 9:
                    aenderungsdatum, abschlusstag, verfalltag, darlehen, zinsen, kosten, versicherung, vertragsnummer, artikel_beschreibung = history_data

                    try:
                        darlehen_float = float(darlehen) if darlehen else 0.0 # Umwandlung in float, falls nicht None
                        zinsen_float = float(zinsen) if zinsen else 0.0       # Umwandlung in float, falls nicht None
                        kosten_float = float(kosten) if kosten else 0.0       # Umwandlung in float, falls nicht None
                    except ValueError:
                        messagebox.showerror("Fehler", "Fehler beim Umwandeln von Darlehen, Zinsen oder Kosten in Zahlen für historisches PDF.")
                        logging.error("ValueError beim Umwandeln von Darlehen, Zinsen oder Kosten in Zahlen für historisches PDF.")
                        return # Abbruch, wenn Umwandlung fehlschlägt

                    original_pfandschein_id = selected_pf[0][0]

                    try:
                        # Kunden-ID und Zifferncode vom Original-Pfandschein holen
                        cursor = self.conn.cursor()
                        cursor.execute("SELECT p.kunde_id, k.name, k.vorname, k.zifferncode FROM pfandscheine p JOIN kunden k ON p.kunde_id = k.id WHERE p.id=?", (original_pfandschein_id,))
                        result = cursor.fetchone()
                        if result:
                            kunde_id, kunde_name, kunde_vorname, kunde_zifferncode = result
                            pf_data = {
                                "Pfandschein-ID (Historisch)": f"{original_pfandschein_id} (Stand: {aenderungsdatum})",
                                "Kunden-Name": f"{kunde_vorname} {kunde_name}",
                                "Zifferncode": kunde_zifferncode, # NEU
                                "Abschlusstag": abschlusstag,
                                "Verfalltag": verfalltag,
                                "Darlehen": f"{darlehen_float:.2f} €", # Nutze float-Variable
                                "Monatl. Zinsen": f"{zinsen_float:.2f} {self.last_zins_einheit}", # Nutze float-Variable
                                "Monatl. Kosten": f"{kosten_float:.2f} €", # Nutze float-Variable
                                "Versicherungssumme": versicherung,
                                "Vertragsnummer": vertragsnummer,
                                "Artikelbeschreibung": artikel_beschreibung,
                                "Änderungsdatum": aenderungsdatum
                            }
                            output_filename = f"historischer_pfandschein_{original_pfandschein_id}_{aenderungsdatum.replace(' ', '_').replace(':', '-')}.pdf"
                            self.generate_pdf_with_background(pf_data, self.pdf_background_path, output_path=output_filename)
                            logging.info(f"Historisches PDF für Pfandschein ID {original_pfandschein_id} (Änderungsdatum: {aenderungsdatum}, Zifferncode: {kunde_zifferncode}) erstellt.")
                        else:
                            messagebox.showerror("Fehler", "Ursprünglicher Pfandschein oder zugehöriger Kunde nicht gefunden, um Kundendaten zu laden!")
                            logging.error(f"Ursprünglicher Pfandschein mit ID {original_pfandschein_id} oder Kunde nicht gefunden.")
                    except sqlite3.Error as e:
                        messagebox.showerror("Fehler", f"Fehler beim Abrufen der Informationen für historischen Pfandschein: {e}")
                        logging.error(f"Fehler beim Abrufen der Informationen für historischen Pfandschein: {e}")
                    except Exception as e:
                         messagebox.showerror("Fehler", f"Allgemeiner Fehler beim Erstellen des historischen PDFs: {e}")
                         logging.error(f"Allgemeiner Fehler beim Erstellen des historischen PDFs: {e}")
                else:
                     messagebox.showwarning("Warnung", "Ausgewählte Historiendaten sind unvollständig.")
            else:
                messagebox.showwarning("Warnung", "Bitte wählen Sie einen Eintrag aus der Historie aus.")





        # GEÄNDERT: Context Menu Definition
        def show_history_context_menu(event):
             item = self.tree_history.identify_row(event.y)
             if item:
                 self.tree_history.selection_set(item)
                 # Erstelle das Menü als Attribut des Fensters, um Referenzprobleme zu vermeiden
                 history_context_menu = tk.Menu(top, tearoff=0) # Gehört zum Toplevel 'top'
                 history_context_menu.add_command(label="Diesen historischen Pfandschein drucken (PDF)", command=print_pdf_from_history)
                 history_context_menu.post(event.x_root, event.y_root)

         # --- NEU: Variablen für Zwei-Wege-Datenbindung Darlehen ---
        self.darlehen_summe_calc_var = tk.StringVar()
        self.darlehen_pfandschein_var = tk.StringVar()

        # --- Rechter Bereich für Kostenrechner ---
        calculator_frame = ttk.Frame(top)
        calculator_frame.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.Y)
        ttk.Label(calculator_frame, text="Kostenrechner", font=("Arial", 10, "bold")).pack(pady=5)

        ttk.Label(calculator_frame, text="Darlehenssumme:").pack(anchor=tk.W)
        self.entry_darlehen_summe_calc = ttk.Entry(calculator_frame, textvariable=self.darlehen_summe_calc_var) # Variable hinzugefügt
        self.entry_darlehen_summe_calc.pack(pady=2, anchor=tk.W)

        ttk.Label(calculator_frame, text="Zinsen in %:").pack(anchor=tk.W)
        self.entry_zinsen_prozent_calc = ttk.Entry(calculator_frame)
        self.entry_zinsen_prozent_calc.pack(pady=2, anchor=tk.W)

        ttk.Label(calculator_frame, text="Kosten in %:").pack(anchor=tk.W)
        self.entry_kosten_prozent_calc = ttk.Entry(calculator_frame)
        self.entry_kosten_prozent_calc.pack(pady=2, anchor=tk.W)

        self.label_calc_zinsen_ergebnis = ttk.Label(calculator_frame, text="Monatl. Zinsen: -")
        self.label_calc_zinsen_ergebnis.pack(pady=5, anchor=tk.W)
        self.label_calc_kosten_ergebnis = ttk.Label(calculator_frame, text="Monatl. Kosten: -")
        self.label_calc_kosten_ergebnis.pack(pady=2, anchor=tk.W)

        def calculate_costs_interests():
            try:
                darlehen_summe = float(self.entry_darlehen_summe_calc.get())
                zinsen_prozent = float(self.entry_zinsen_prozent_calc.get())
                kosten_prozent = float(self.entry_kosten_prozent_calc.get())

                monatl_zinsen_calc = (darlehen_summe * zinsen_prozent) / 100
                monatl_kosten_calc = (darlehen_summe * kosten_prozent) / 100

                self.label_calc_zinsen_ergebnis.config(text=f"Monatl. Zinsen: {monatl_zinsen_calc:.2f} €")
                self.label_calc_kosten_ergebnis.config(text=f"Monatl. Kosten: {monatl_kosten_calc:.2f} €")

            except ValueError:
                messagebox.showerror("Fehler", "Bitte gültige Zahlen in den Rechner eingeben.")
                self.label_calc_zinsen_ergebnis.config(text="Monatl. Zinsen: -")
                self.label_calc_kosten_ergebnis.config(text="Monatl. Kosten: -")

        def transfer_calculated_values():
            zinsen_text = self.label_calc_zinsen_ergebnis.cget("text").split(": ")[1].replace(" €", "")
            kosten_text = self.label_calc_kosten_ergebnis.cget("text").split(": ")[1].replace(" €", "")

            if zinsen_text != "-" and kosten_text != "-":
                detail_entries["Monatl. Zinsen:"].delete(0, tk.END)
                detail_entries["Monatl. Zinsen:"].insert(0, zinsen_text)
                detail_entries["Monatl. Kosten:"].delete(0, tk.END)
                detail_entries["Monatl. Kosten:"].insert(0, kosten_text)
                # messagebox.showinfo("Erfolg", "Berechnete Werte in die Pfandscheindetails übernommen.") # Bestätigungsmeldung # Entfernt, da keine Interaktion gewünscht
            else:
                messagebox.showwarning("Warnung", "Bitte zuerst Werte berechnen.")

        ttk.Button(calculator_frame, text="Berechnen", command=calculate_costs_interests).pack(pady=10)
        ttk.Button(calculator_frame, text="Übernehmen", command=transfer_calculated_values).pack(pady=5)

        # --- Detailfelder für Pfandschein-Daten (mittlerer Bereich) ---
        detail_labels = [
            "Abschlusstag (TT.MM.JJJJ):",
            "Verfalltag (TT.MM.JJJJ):",
            "Darlehen:",
            "Monatl. Zinsen:",
            "Monatl. Kosten:",
            "Versicherungssumme:",
            "Nr. des früheren Vertrages:",
            "Artikelbeschreibung:",
        ]
        detail_entries = {}
        grid_row_offset = 0
        for i, lab in enumerate(detail_labels):
            ttk.Label(middle_frame, text=lab).grid(row=i, column=0, sticky=tk.E, padx=5, pady=5)
            if lab == "Darlehen:": # Hier Variable für Zwei-Wege-Bindung verwenden
                ent = ttk.Entry(middle_frame, width=40, textvariable=self.darlehen_pfandschein_var) # Variable hinzugefügt
            else:
                ent = ttk.Entry(middle_frame, width=40)
            ent.grid(row=i, column=1, padx=5, pady=5)
            detail_entries[lab] = ent
            grid_row_offset = i

        # --- Funktionen für Zwei-Wege-Datenbindung ---
        def update_darlehen_calc(*args):
            """Aktualisiert Darlehenssumme im Rechner, wenn Darlehen im Pfandschein geändert wird."""
            darlehen_pf_value = self.darlehen_pfandschein_var.get()
            if darlehen_pf_value != self.darlehen_summe_calc_var.get(): # Verhindert Endlosschleife
                self.darlehen_summe_calc_var.set(darlehen_pf_value)

        def update_darlehen_pfandschein(*args):
            """Aktualisiert Darlehen im Pfandschein, wenn Darlehenssumme im Rechner geändert wird."""
            darlehen_calc_value = self.darlehen_summe_calc_var.get()
            if darlehen_calc_value != self.darlehen_pfandschein_var.get(): # Verhindert Endlosschleife
                detail_entries["Darlehen:"].delete(0, tk.END) # Feld leeren, bevor neuer Wert gesetzt wird
                detail_entries["Darlehen:"].insert(0, darlehen_calc_value)


        # Beobachte Änderungen in den Variablen
        self.darlehen_pfandschein_var.trace_add('write', update_darlehen_calc)
        self.darlehen_summe_calc_var.trace_add('write', update_darlehen_pfandschein)


        # Buttons im mittleren Bereich (unverändert)
        button_start_row = grid_row_offset + 1
        btn_create_pf = ttk.Button(middle_frame, text="Neuen Pfandschein anlegen", command=create_new_pf)
        btn_create_pf.grid(row=button_start_row, column=0, columnspan=2, pady=10, sticky=tk.EW) # sticky=tk.EW hinzugefügt
        btn_update_pf = ttk.Button(middle_frame, text="Pfandschein verlängern/ändern", command=update_pf)
        btn_update_pf.grid(row=button_start_row + 1, column=0, columnspan=2, pady=5, sticky=tk.EW) # sticky=tk.EW hinzugefügt
        btn_show_pdf = ttk.Button(middle_frame, text="Pfandschein drucken (PDF)", command=show_pdf)
        btn_show_pdf.grid(row=button_start_row + 2, column=0, columnspan=2, pady=5, sticky=tk.EW) # sticky=tk.EW hinzugefügt
        btn_history = ttk.Button(middle_frame, text="Pfandschein Historie anzeigen", command=show_pfandschein_history)
        btn_history.grid(row=button_start_row + 3, column=0, columnspan=2, pady=5, sticky=tk.EW) # sticky=tk.EW hinzugefügt

        refresh_pf_list()

    # --- generate_pdf_with_background ---
    # GEÄNDERT: Barcode wird generiert und über der Tabelle platziert
    # GEÄNDERT: Verbesserte Fehlerbehandlung für Barcode
    def generate_pdf_with_background(self, pf_data, background_image_path, output_path="pfandschein.pdf"):
        a4_width, a4_height = A4
        custom_height = a4_height / 3  # Ein Drittel der A4-Höhe
        custom_pagesize = (a4_width, custom_height)
        c = canvas.Canvas(output_path, pagesize=custom_pagesize)
        width, height = custom_pagesize

        # --- Breitenverteilung (unverändert) ---
        tear_off_width = 55 * mm
        main_content_width = width - tear_off_width
        line_x = main_content_width + 1*mm
        tear_off_x_start = line_x + 10*mm

        # --- Hintergrundbild (unverändert) ---
        try:
            if background_image_path and os.path.exists(background_image_path):
                c.drawImage(background_image_path, 0, 0, width=width, height=height, preserveAspectRatio=True, anchor='c')
                logging.info(f"Hintergrundbild geladen von: {background_image_path}")
            else:
                 logging.warning(f"Hintergrundbild nicht gefunden oder Pfad leer: {background_image_path}. Zeichne Fallback-Rahmen.")
                 c.rect(0, 0, width, height) # Fallback: Rahmen zeichnen
        except Exception as e:
            logging.error(f"Fehler beim Laden des Hintergrundbildes von '{background_image_path}': {e}")
            c.rect(0, 0, width, height) # Fallback: Rahmen zeichnen

        # --- Allgemeine Einstellungen (unverändert) ---
        left_margin = 5 * mm
        text_color = black
        font_name = "Helvetica"
        font_size_table_header = 6.5
        font_size_table_data = 7
        font_size_info = 8
        font_size_summary = 7
        line_height_info = 4 * mm
        line_height_summary = 3.8 * mm

        # --- Style für Paragraph im Tabellenkopf (unverändert) ---
        styles = getSampleStyleSheet()
        header_style = styles['Normal']
        header_style.fontName = font_name
        header_style.fontSize = font_size_table_header
        header_style.leading = font_size_table_header + 1
        header_style.alignment = 1

        # --- Hauptinhalt (links) ---

        # NEU: Barcode generieren und platzieren (mit verbesserter Fehlerbehandlung)
        barcode_buffer = None
        if BARCODE_LIB_AVAILABLE and 'Zifferncode' in pf_data and pf_data['Zifferncode']:
            zifferncode_str = None # Vorinitialisieren für den Fehlerfall
            try:
                zifferncode_str = str(pf_data['Zifferncode'])
                Code128 = barcode.get_barcode_class('code128')
                # Optionen: Höhe, Schriftgröße unter Barcode, Abstand Text, Ruhezone
                options = {'module_height': 12.0, 'font_size': 2.0, 'text_distance': 1.0, 'quiet_zone': 2.0, 'write_text': True}
                my_barcode = Code128(zifferncode_str, writer=ImageWriter()) # ImageWriter benötigt Pillow

                # Barcode in einen BytesIO-Puffer schreiben (im Speicher)
                barcode_buffer = io.BytesIO()
                # *** WICHTIG: Fehlerquelle hier möglich (z.B. Pillow fehlt) ***
                my_barcode.write(barcode_buffer, options=options)
                barcode_buffer.seek(0) # Zurück zum Anfang des Puffers

                # Barcode zeichnen
                barcode_img_width = 450 * mm # Breite des Barcodes auf dem PDF
                barcode_img_height = 32 * mm # Höhe des Barcodes auf dem PDF
                barcode_y = height - 13 * mm # Y-Position (oberer Rand des PDFs - Abstand)
                barcode_x = left_margin

                # ImageReader verwenden, um das Bild aus dem Puffer zu lesen
                # *** WICHTIG: Fehlerquelle hier möglich ***
                c.drawImage(ImageReader(barcode_buffer), barcode_x, barcode_y,
                            width=barcode_img_width, height=barcode_img_height,
                            preserveAspectRatio=True, anchor='sw') # 'sw' = South-West (linke untere Ecke)

                logging.info(f"Barcode für Zifferncode {zifferncode_str} erfolgreich generiert und platziert.")

            except Exception as e:
                # *** NEU: Direkte Fehlermeldung anzeigen ***
                error_message = f"Fehler bei der Barcode-Generierung für Code '{zifferncode_str if zifferncode_str else 'N/A'}':\n\n{type(e).__name__}: {e}\n\nÜberprüfen Sie die Log-Datei ('pfandhaus_app.log') für Details.\nStellen Sie sicher, dass 'Pillow' installiert ist (pip install Pillow)."
                # ... (restliche Fehlerbehandlung und Logging)
                logging.error(error_message) # Logge die detaillierte Nachricht
                messagebox.showerror("Barcode Fehler", error_message)
                # Optional: Hier die PDF-Generierung abbrechen, wenn der Barcode kritisch ist
                # return # <-- Wenn du willst, dass bei Barcode-Fehler gar kein PDF kommt

                # Versuch, Fehlermeldung im PDF anzuzeigen (kann fehlschlagen, wenn Canvas kaputt ist)
                try:
                    c.setFillColorRGB(1, 0, 0) # Rot
                    c.setFont("Helvetica", 8)
                    c.drawString(left_margin, height - 30 * mm, f"FEHLER Barcode: {type(e).__name__}") # Kürzerer Text im PDF
                    c.setFillColor(text_color) # Farbe zurücksetzen
                except Exception as draw_error:
                    logging.error(f"Konnte Barcode-Fehlermeldung nicht ins PDF zeichnen: {draw_error}")

            finally:
                if barcode_buffer:
                    barcode_buffer.close() # Puffer schließen
        elif not BARCODE_LIB_AVAILABLE:
             logging.warning("Barcode-Bibliothek nicht verfügbar, Barcode wird übersprungen.")
        # Optional: Loggen, wenn Zifferncode fehlt
        elif 'Zifferncode' not in pf_data or pf_data['Zifferncode'] is None or pf_data['Zifferncode'] == "":
           logging.info("Kein Zifferncode für diesen Kunden vorhanden, Barcode wird übersprungen.")


        # --- Tabellendaten (unverändert) ---
        table_headers = [
            "Abschlusstag", "Verfalltag", "Darlehen", "Monatl. Zinsen", "Monatl. Kosten", "Versicherungs- summe", "Nr. des früheren Vertrages"
        ]
        table_content_row = [
            pf_data.get("Abschlusstag", ""),
            pf_data.get("Verfalltag", ""),
            pf_data.get("Darlehen", ""),
            pf_data.get("Monatl. Zinsen", ""),
            pf_data.get("Monatl. Kosten", ""),
            pf_data.get("Versicherungssumme", ""),
            pf_data.get("Vertragsnummer", "")
        ]

        # --- Tabellen-Layout (unverändert) ---
        # GEÄNDERT: Startposition der Tabelle etwas nach unten verschoben, um Platz für Barcode zu schaffen
        table_y_start = height - 40 * mm # War vorher: height - 45 * mm

        # Spaltenbreiten (unverändert)
        col_widths = [
            20.24 * mm, 20.24 * mm, 20.24 * mm, 18.48 * mm, 23.52 * mm, 23.52 * mm, 18.48 * mm
        ]

        # Berechnung der dynamischen Höhe der Header-Zeile (unverändert)
        header_paragraphs = []
        max_header_height = 0
        for i, header_text in enumerate(table_headers):
            p = Paragraph(header_text.replace(" ", "<br/>"), header_style) # Ersetze Leerzeichen durch Zeilenumbruch für Umbruch
            p_w, p_h = p.wrapOn(c, col_widths[i] - 2*mm, 1000) # Breite - Puffer, Höhe groß
            header_paragraphs.append(p)
            if p_h > max_header_height:
                max_header_height = p_h
        header_row_height = max(max_header_height + 2*mm, 8*mm) # Mindestens 8mm, sonst dynamisch + Puffer

        # Höhe der Datenzeile (fix, unverändert)
        data_row_height = 6 * mm

        line_color = black

        # --- Tabelle zeichnen (unverändert) ---
        c.setStrokeColor(line_color)
        c.setFillColor(text_color)

        # --- Kopfzeile zeichnen (mit Umbruch, unverändert) ---
        current_x = left_margin # Korrektur: Start bei left_margin
        current_y = table_y_start
        for i, p in enumerate(header_paragraphs):
            col_width = col_widths[i]
            c.rect(current_x, current_y - header_row_height, col_width, header_row_height)
            p_draw_y = (current_y - header_row_height) + (header_row_height - p.height) / 2
            p.drawOn(c, current_x + 1*mm, p_draw_y)
            current_x += col_width

        # --- Datenzeile zeichnen (manuell, unverändert) ---
        current_x = left_margin # Korrektur: Start bei left_margin
        current_y = table_y_start - header_row_height # Y-Position unter dem Header
        c.setFont(font_name, font_size_table_data)
        for i, cell_text in enumerate(table_content_row):
            cell_text_str = str(cell_text)
            col_width = col_widths[i]
            max_cell_width = col_width - 2*mm
            # Einfachere Kürzung, falls Text zu lang ist
            while c.stringWidth(cell_text_str, font_name, font_size_table_data) > max_cell_width and len(cell_text_str) > 5:
                 cell_text_str = cell_text_str[:int(len(cell_text_str)*0.8)] + "..." # Kürzen

            c.rect(current_x, current_y - data_row_height, col_width, data_row_height)
            text_y_offset = 1.5 * mm # Kleiner Offset für bessere Zentrierung
            # Zentrieren des Textes in der Zelle (optional, einfacher ist linksbündig mit Rand)
            # text_width = c.stringWidth(cell_text_str, font_name, font_size_table_data)
            # text_x_pos = current_x + (col_width - text_width) / 2
            text_x_pos = current_x + 1 * mm # Linksbündig mit kleinem Rand
            c.drawString(text_x_pos, current_y - data_row_height + text_y_offset, cell_text_str)
            current_x += col_width


        # --- Zusätzliche Informationen unterhalb der Tabelle (unverändert) ---
        info_y = table_y_start - header_row_height - data_row_height - line_height_info
        c.setFillColor(text_color)
        c.setFont(font_name, font_size_info)

        pf_id_key = "Pfandschein-ID (Historisch)" if "Pfandschein-ID (Historisch)" in pf_data else "Pfandschein-ID"
        if pf_id_key in pf_data:
            c.setFont("Helvetica-Bold", 10)
            if info_y > 5 * mm:
                c.drawString(left_margin, info_y, f"Pfandschein Nr: {pf_data[pf_id_key]}")
                info_y -= line_height_info
            c.setFont(font_name, font_size_info)

        if "Kunden-Name" in pf_data:
            if info_y > 5 * mm:
                c.drawString(left_margin, info_y, f"Kunde: {pf_data['Kunden-Name']}")
                # info_y -= line_height_info # Verschieben, damit Zifferncode daneben passt

        # NEU: Zifferncode unter Kunde anzeigen (optional)
        if "Zifferncode" in pf_data and pf_data['Zifferncode'] is not None and pf_data['Zifferncode'] != "":
             if info_y > 5 * mm:
                # Zeichne es neben den Kundennamen, etwas eingerückt
                kunde_name_width = c.stringWidth(f"Kunde: {pf_data['Kunden-Name']}", font_name, font_size_info)
                c.drawString(left_margin + kunde_name_width + 2*mm, info_y, f"(Code: {pf_data['Zifferncode']})")
                info_y -= line_height_info # Jetzt Zeile nach unten

        if "Artikelbeschreibung" in pf_data and pf_data["Artikelbeschreibung"]:
            max_width_artikel = main_content_width - left_margin - 5*mm
            artikel_text = pf_data['Artikelbeschreibung']
            lines = []
            current_line = ""
            for word in artikel_text.split():
                test_line = f"{current_line} {word}".strip()
                if c.stringWidth(test_line, font_name, font_size_info) <= max_width_artikel:
                    current_line = test_line
                else:
                    if current_line:
                        lines.append(current_line)
                    # Prüfen, ob das einzelne Wort schon zu lang ist
                    if c.stringWidth(word, font_name, font_size_info) > max_width_artikel:
                         # Wort aufteilen oder kürzen (einfache Kürzung hier)
                         while c.stringWidth(word, font_name, font_size_info) > max_width_artikel:
                             word = word[:-1]
                         lines.append(word + "...")
                         current_line = "" # Kein neues Wort starten
                    else:
                        current_line = word
            if current_line:
                lines.append(current_line)

            artikel_label = "Artikel: "
            first_line = True
            # Berechne Einzug basierend auf Breite des Labels, falls vorhanden
            if first_line:
                 try:
                     label_width = c.stringWidth(artikel_label, font_name, font_size_info)
                     indent_spaces = int(label_width / c.stringWidth(' ', font_name, font_size_info))
                     indent = ' ' * indent_spaces
                 except ZeroDivisionError: # Falls font_size_info 0 ist (sehr unwahrscheinlich, aber sicherheitshalber)
                     indent = ' ' * 8 # Fallback-Einzug
            else:
                 indent = '' # Nicht die erste Zeile

            for line in lines:
                 if info_y > 5 * mm:
                     # Verwende Einzug für zweite und folgende Zeilen
                     prefix = artikel_label if first_line else indent
                     # Zusätzliche Prüfung, ob die erste Zeile den Einzug braucht (falls label leer ist)
                     effective_prefix = prefix if prefix or first_line else ""
                     c.drawString(left_margin, info_y, f"{effective_prefix}{line}")
                     info_y -= line_height_info
                     first_line = False # Nach der ersten Zeile ist es nicht mehr die erste

        if "Änderungsdatum" in pf_data:
             if info_y > 5 * mm:
                c.drawString(left_margin, info_y, f"Letzte Änderung: {pf_data['Änderungsdatum']}")
                info_y -= line_height_info


        # --- Abriss-Bereich (rechts) ---
        c.setFont(font_name, font_size_summary)
        # GEÄNDERT: Startposition etwas nach unten verschoben wegen Barcode oben
        summary_y = height - 30 * mm # War vorher: height - 25 * mm

        summary_max_width = tear_off_width - (10*mm + 2*mm) # Verfügbare Breite im Abriss

        def draw_summary_line(label, value):
            nonlocal summary_y
            if summary_y < 5 * mm:
                return
            text = f"{label}: {value}"
            # Text kürzen, falls zu lang für den Abrissbereich
            while c.stringWidth(text, font_name, font_size_summary) > summary_max_width and len(text) > 10:
                 text = text[:-4] + "..."
            # Benutze die neue Startposition tear_off_x_start
            c.drawString(tear_off_x_start, summary_y, text)
            summary_y -= line_height_summary

        # Informationen für den Abrissbereich (Inhalt unverändert)
        draw_summary_line("Pfandschein Nr", pf_data.get(pf_id_key, ''))
        draw_summary_line("Kunde", pf_data.get('Kunden-Name', ''))
        # NEU: Zifferncode im Abrissbereich
        if 'Zifferncode' in pf_data and pf_data['Zifferncode'] is not None and pf_data['Zifferncode'] != "":
            draw_summary_line("Code", pf_data['Zifferncode'])
        # Artikelbeschreibung im Abrissbereich kürzen
        artikel_summary = pf_data.get('Artikelbeschreibung', '')
        if len(artikel_summary) > 25: # Beispielhafte Längenbegrenzung
            artikel_summary = artikel_summary[:22] + "..."
        draw_summary_line("Artikel", artikel_summary)
        draw_summary_line("Darlehen", pf_data.get('Darlehen', ''))
        draw_summary_line("Zins", pf_data.get('Monatl. Zinsen', ''))
        draw_summary_line("Kosten", pf_data.get('Monatl. Kosten', ''))
        draw_summary_line("Abschluss", pf_data.get('Abschlusstag', ''))
        draw_summary_line("Verfall", pf_data.get('Verfalltag', ''))
        if "Änderungsdatum" in pf_data:
            draw_summary_line("Hist.Datum", pf_data.get('Änderungsdatum', ''))


        # --- Trennlinie (Position unverändert, verwendet line_x) ---
        c.setStrokeColor(black)
        c.setDash(1, 2) # Gepunktete oder gestrichelte Linie
        c.line(line_x, height - 5*mm, line_x, 5*mm) # Von oben nach unten
        c.setDash() # Linienstil zurücksetzen


        # --- Speichern und Öffnen (Code unverändert) ---
        try:
            c.save()
            logging.info(f"PDF im angepassten Format erstellt: {output_path}")
            if os.name == "nt":
                os.startfile(output_path)
            elif os.name == "posix":
                subprocess.run(["xdg-open", output_path])
            else:
                messagebox.showinfo("Info", f"PDF erstellt: {output_path}")
        except PermissionError:
             logging.error(f"Fehler beim Speichern der PDF: Keine Berechtigung für '{output_path}'. Ist die Datei geöffnet?")
             messagebox.showerror("Fehler", f"PDF konnte nicht gespeichert werden.\nIst die Datei '{os.path.basename(output_path)}' eventuell noch geöffnet?")
        except Exception as e:
            logging.error(f"Fehler beim Speichern oder Öffnen der PDF-Datei: {e}")
            messagebox.showerror("Fehler", f"Fehler beim Speichern/Öffnen der PDF:\n{e}")


    # --- Alle Methoden nach generate_pdf_with_background bleiben unverändert ---
    # (insbesondere __main__ Block)

if __name__ == "__main__":
    root = ThemedTk() # Theme wird beim Start aus der Konfiguration geladen
    # NEU: Prüfen ob Barcode Lib verfügbar ist, bevor App gestartet wird
    # (Die Prüfung ist schon oben im Import-Block, hier nur zur Sicherheit)
    if not BARCODE_LIB_AVAILABLE:
        # Die Warnung wurde schon angezeigt. Ggf. hier nochmals oder App beenden?
        # Fürs Erste lassen wir die App laufen, aber ohne Barcode-Funktion.
        logging.warning("Barcode-Funktionalität ist aufgrund fehlender Bibliotheken deaktiviert.")
        pass # App trotzdem starten
    app = PfandhausApp(root)
    # root.attributes("-fullscreen", True) # Startet im Vollbildmodus (optional)
    root.state('zoomed') # Maximiert starten unter Windows/Mac (besser als fullscreen)
    root.mainloop()
