import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
import datetime
import sys
import csv
import subprocess
import platform
import tempfile
import json
from pathlib import Path

# --- FUNZIONI DI UTILITÀ ---
def format_currency(value):
    if value is None: value = 0.0
    return f"€ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def safe_float_convert(value_str):
    """Converte stringa in float per calcoli, gestendo formattazione italiana"""
    if not value_str: return 0.0
    try:
        clean = str(value_str).replace('€', '').strip()
        if '.' in clean and ',' in clean:
            clean = clean.replace('.', '').replace(',', '.')
        elif ',' in clean:
            clean = clean.replace(',', '.')
        return float(clean)
    except ValueError:
        return 0.0

# --- GESTIONE DATABASE ---
class Database:
    def __init__(self, db_name="np_zero.db"):
        self.db_name = db_name
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.create_tables()
        self.populate_defaults()

    def create_tables(self):
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS unita_misura (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codice TEXT UNIQUE,
                nome TEXT,
                descrizione TEXT
            )
        """)
        
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS progetti (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codice TEXT,
                titolo TEXT,
                cup TEXT,
                committente TEXT
            )
        """)

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS nuovi_prezzi (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                progetto_id INTEGER,
                codice TEXT, 
                descrizione TEXT,
                unita_misura TEXT, 
                perc_spese_generali REAL DEFAULT 17.0,
                perc_sicurezza REAL DEFAULT 5.0,
                perc_utili REAL DEFAULT 10.0,
                prezzo_finale REAL DEFAULT 0.0,
                FOREIGN KEY(progetto_id) REFERENCES progetti(id) ON DELETE CASCADE
            )
        """)

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS voci_costo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                np_id INTEGER,
                ordine INTEGER,
                categoria TEXT, 
                descrizione TEXT,
                um TEXT,
                quantita REAL,
                prezzo_unitario REAL,
                FOREIGN KEY(np_id) REFERENCES nuovi_prezzi(id) ON DELETE CASCADE
            )
        """)
        self.conn.commit()

    def populate_defaults(self):
        self.cursor.execute("SELECT count(*) FROM unita_misura")
        if self.cursor.fetchone()[0] == 0:
            defaults = [
                    ("cad.", "Cadauno", "Sanitari, infissi..."),
                    ("corpo", "A corpo", "Opere indivisibili"),
                    ("h", "Ora", "Manodopera..."),
                    ("kg", "Chilogrammo", "Acciaio..."),
                    ("m", "Metro lineare", "Tubazioni, cavi..."),
                    ("mc", "Metro cubo", "Getti in cls, scavi..."),
                    ("mq", "Metro quadrato", "Pavimenti, intonaci..."),
            ]
            self.cursor.executemany("INSERT INTO unita_misura (codice, nome, descrizione) VALUES (?, ?, ?)", defaults)
            self.conn.commit()

    def select_specific(self, table, columns, where_clause="", params=()):
        cols_str = ", ".join(columns)
        query = f"SELECT id, {cols_str} FROM {table} {where_clause}"
        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def fetch_all(self, table, where_clause="", params=()):
        query = f"SELECT * FROM {table} {where_clause}"
        self.cursor.execute(query, params)
        return self.cursor.fetchall()
    
    def fetch_one(self, table, where_clause="", params=()):
        query = f"SELECT * FROM {table} {where_clause}"
        self.cursor.execute(query, params)
        return self.cursor.fetchone()

    def get_max_order(self, np_id):
        query = "SELECT MAX(ordine) FROM voci_costo WHERE np_id=?"
        self.cursor.execute(query, (np_id,))
        val = self.cursor.fetchone()[0]
        return val if val is not None else 0

    def insert(self, table, data):
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        values = list(data.values())
        query = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
        self.cursor.execute(query, values)
        self.conn.commit()
        return self.cursor.lastrowid 

    def delete(self, table, record_id):
        self.cursor.execute(f"DELETE FROM {table} WHERE id=?", (record_id,))
        self.conn.commit()

    def update(self, table, record_id, data):
        set_clause = ', '.join([f"{k}=?" for k in data.keys()])
        values = list(data.values()) + [record_id]
        query = f"UPDATE {table} SET {set_clause} WHERE id=?"
        self.cursor.execute(query, values)
        self.conn.commit()

# --- FINESTRA DI IMPORTAZIONE ---
class ImportDialog(tk.Toplevel):
    def __init__(self, parent, db, current_project_id, on_import_callback):
        super().__init__(parent)
        self.title("Importa NP da altro progetto")
        self.geometry("600x400")
        self.db = db
        self.current_project_id = current_project_id
        self.on_import_callback = on_import_callback
        
        ttk.Label(self, text="1. Seleziona Progetto di Origine:", font=("Arial", 10, "bold")).pack(pady=5)
        
        self.combo_proj = ttk.Combobox(self, state="readonly", width=50)
        self.combo_proj.pack(pady=5)
        self.combo_proj.bind("<<ComboboxSelected>>", self.load_nps)
        
        ttk.Label(self, text="2. Seleziona NP da Copiare:", font=("Arial", 10, "bold")).pack(pady=5)
        
        cols = ("codice", "desc")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=10)
        self.tree.heading("codice", text="Codice")
        self.tree.heading("desc", text="Descrizione")
        self.tree.column("codice", width=100)
        self.tree.column("desc", width=400)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        
        btn_import = ttk.Button(self, text="IMPORTA SELEZIONATO", command=self.do_import)
        btn_import.pack(pady=10, fill="x", padx=20)
        
        self.load_projects()

    def load_projects(self):
        projs = self.db.fetch_all("progetti")
        self.proj_map = {}
        values = []
        for p in projs:
            if p[0] != self.current_project_id:
                label = f"{p[1]} - {p[2]}"
                values.append(label)
                self.proj_map[label] = p[0]
        self.combo_proj['values'] = values

    def load_nps(self, event):
        for row in self.tree.get_children(): self.tree.delete(row)
        selected_label = self.combo_proj.get()
        if not selected_label: return
        
        pid = self.proj_map[selected_label]
        nps = self.db.fetch_all("nuovi_prezzi", "WHERE progetto_id=?", (pid,))
        for np in nps:
            self.tree.insert("", "end", iid=np[0], values=(np[2], np[3]))

    def do_import(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Attenzione", "Seleziona un NP dalla lista")
            return
        
        source_np_id = sel[0]
        self.on_import_callback(source_np_id)
        self.destroy()

# --- PANNELLO CRUD ---
class CrudPanel(ttk.Frame):
    def __init__(self, parent, db, table_name, fields, callbacks=None):
        super().__init__(parent)
        self.db = db
        self.table_name = table_name
        self.fields = fields 
        self.callbacks = callbacks or {}
        self.entries = {}
        self.buttons = {}

        self.frame_top = ttk.LabelFrame(self, text="Dati Inserimento", padding=10)
        self.frame_top.pack(side="top", fill="x", padx=10, pady=5)
        
        if self.table_name == "nuovi_prezzi":
            self.setup_custom_np_ui()
        elif self.table_name == "progetti":
            self.setup_custom_progetti_ui()
        else:
            self.setup_generic_ui()

        self.frame_mid = ttk.Frame(self)
        self.frame_mid.pack(side="top", fill="both", expand=True, padx=10, pady=5)

        cols = [f[0] for f in self.fields]
        self.tree = ttk.Treeview(self.frame_mid, columns=cols, show="headings")
        
        for field, label, width in self.fields:
            align = "w"
            if field == "prezzo_finale": align = "e"
            self.tree.heading(field, text=label, 
                              command=lambda c=field: self.sort_column(c, False))
            self.tree.column(field, width=width, anchor=align)
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        sb = ttk.Scrollbar(self.frame_mid, orient="vertical", command=self.tree.yview)
        sb.pack(side="left", fill="y")
        self.tree.configure(yscroll=sb.set)

        self.frame_btns = ttk.Frame(self.frame_mid)
        self.frame_btns.pack(side="right", fill="y", padx=5)

        self.buttons['add'] = ttk.Button(self.frame_btns, text="Nuovo", command=self.add_record)
        self.buttons['add'].pack(fill="x", pady=2)
        self.buttons['upd'] = ttk.Button(self.frame_btns, text="Modifica", command=self.update_record)
        self.buttons['upd'].pack(fill="x", pady=2)
        self.buttons['del'] = ttk.Button(self.frame_btns, text="Cancella", command=self.delete_record)
        self.buttons['del'].pack(fill="x", pady=2)
        self.buttons['dup'] = ttk.Button(self.frame_btns, text="Duplica", command=self.duplicate_record)
        self.buttons['dup'].pack(fill="x", pady=2)
        self.buttons['imp'] = ttk.Button(self.frame_btns, text="Importa da...", command=lambda: print("Implementare Wrapper")) 
        self.buttons['imp'].pack(fill="x", pady=2)
        
        ttk.Separator(self.frame_btns, orient="horizontal").pack(fill="x", pady=5)
        self.buttons['clr'] = ttk.Button(self.frame_btns, text="Pulisci", command=self.clear_fields)
        self.buttons['clr'].pack(fill="x", pady=10)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.on_double_click)
        
        self.refresh_data()

    def setup_generic_ui(self):
        for i, (field_name, label_text, width) in enumerate(self.fields):
            if field_name == "prezzo_finale": continue
            container = ttk.Frame(self.frame_top)
            container.pack(side="left", padx=5, fill="y")
            ttk.Label(container, text=label_text, font=("Arial", 9, "bold")).pack(side="top", anchor="w")
            entry = ttk.Entry(container, width=width // 8) 
            entry.pack(side="top", fill="x")
            self.entries[field_name] = entry

    def setup_custom_progetti_ui(self):
        row1 = ttk.Frame(self.frame_top)
        row1.pack(side="top", fill="x", pady=2)
        def add_field(parent, key, label, width_factor):
            f = ttk.Frame(parent)
            f.pack(side="left", padx=(0, 15), fill="x", expand=False)
            ttk.Label(f, text=label, font=("Arial", 9, "bold")).pack(anchor="w")
            e = ttk.Entry(f, width=width_factor)
            e.pack(fill="x")
            self.entries[key] = e
        add_field(row1, 'codice', 'Codice', 15)
        add_field(row1, 'cup', 'CUP', 15)
        f_comm = ttk.Frame(row1)
        f_comm.pack(side="left", padx=0, fill="x", expand=True)
        ttk.Label(f_comm, text="Committente", font=("Arial", 9, "bold")).pack(anchor="w")
        self.entries['committente'] = ttk.Entry(f_comm)
        self.entries['committente'].pack(fill="x")
        row2 = ttk.Frame(self.frame_top)
        row2.pack(side="top", fill="x", pady=5)
        ttk.Label(row2, text="Titolo Intervento", font=("Arial", 9, "bold")).pack(anchor="w")
        self.entries['titolo'] = ttk.Entry(row2)
        self.entries['titolo'].pack(fill="x")

    def setup_custom_np_ui(self):
        row1 = ttk.Frame(self.frame_top)
        row1.pack(side="top", fill="x", pady=2)
        f_cod = ttk.Frame(row1)
        f_cod.pack(side="left", padx=(0, 20))
        ttk.Label(f_cod, text="Codice NP", font=("Arial", 9, "bold")).pack(anchor="w")
        self.entries['codice'] = ttk.Entry(f_cod, width=15)
        self.entries['codice'].pack(fill="x")
        f_um = ttk.Frame(row1)
        f_um.pack(side="left")
        ttk.Label(f_um, text="Unità di Misura", font=("Arial", 9, "bold")).pack(anchor="w")
        self.entries['unita_misura'] = ttk.Combobox(f_um, width=15, postcommand=self.populate_um_combo)
        self.entries['unita_misura'].pack(fill="x")
        row2 = ttk.Frame(self.frame_top)
        row2.pack(side="top", fill="x", pady=5)
        ttk.Label(row2, text="Descrizione NP", font=("Arial", 9, "bold")).pack(anchor="w")
        txt_frame = ttk.Frame(row2)
        txt_frame.pack(fill="x", expand=True)
        self.entries['descrizione'] = tk.Text(txt_frame, height=4, font=("Arial", 9))
        self.entries['descrizione'].pack(side="left", fill="x", expand=True)
        sb = ttk.Scrollbar(txt_frame, orient="vertical", command=self.entries['descrizione'].yview)
        sb.pack(side="right", fill="y")
        self.entries['descrizione'].config(yscrollcommand=sb.set)

    def populate_um_combo(self):
        if 'unita_misura' in self.entries and isinstance(self.entries['unita_misura'], ttk.Combobox):
            ums = self.db.fetch_all("unita_misura")
            values = [u[1] for u in ums]
            self.entries['unita_misura']['values'] = values

    def sort_column(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        def sort_key(t):
            val = t[0]
            if not val: return ""
            try:
                clean = str(val).replace('€', '').strip()
                if '.' in clean and ',' in clean: clean = clean.replace('.', '').replace(',', '.')
                elif ',' in clean: clean = clean.replace(',', '.')
                return float(clean)
            except ValueError:
                return str(val).lower()
        l.sort(key=sort_key, reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)
        self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))

    def get_data_from_ui(self):
        data = {}
        for field, _, _ in self.fields:
            if field in self.entries:
                widget = self.entries[field]
                if isinstance(widget, tk.Text):
                    val = widget.get("1.0", "end-1c").strip()
                else:
                    val = widget.get().strip()
                data[field] = val
        return data

    def clear_fields(self):
        for widget in self.entries.values():
            if isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
            else:
                widget.delete(0, tk.END)

    def refresh_data(self, condition="", params=()):
        for row in self.tree.get_children(): self.tree.delete(row)
        target_columns = [f[0] for f in self.fields]
        rows = self.db.select_specific(self.table_name, target_columns, condition, params)
        for r in rows:
            vals = list(r[1:])
            for idx, f in enumerate(self.fields):
                if f[0] == "prezzo_finale" and idx < len(vals):
                     vals[idx] = format_currency(vals[idx] or 0.0)
            self.tree.insert("", "end", iid=r[0], values=vals)

    def add_record(self):
        try:
            data = self.get_data_from_ui()
            if not any(data.values()): return
            self.db.insert(self.table_name, data)
            self.refresh_data() 
            self.clear_fields()
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def update_record(self):
        selected = self.tree.selection()
        if not selected: return
        try:
            data = self.get_data_from_ui()
            self.db.update(self.table_name, selected[0], data)
            self.refresh_data() 
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def delete_record(self):
        selected = self.tree.selection()
        if not selected: return
        if messagebox.askyesno("Conferma", "Cancellare record?"):
            self.db.delete(self.table_name, selected[0])
            self.refresh_data() 
            self.clear_fields()

    def duplicate_record(self):
        selected = self.tree.selection()
        if not selected: return
        data = self.get_data_from_ui()
        if 'codice' in data: data['codice'] += "_cp"
        if 'titolo' in data: data['titolo'] += " (Copia)"
        try:
            self.db.insert(self.table_name, data)
            self.refresh_data() 
        except Exception as e:
            messagebox.showerror("Errore", str(e))

    def on_select(self, event):
        selected = self.tree.selection()
        if not selected: return
        item = self.tree.item(selected[0])
        values = item['values']
        self.clear_fields()
        col_idx = 0
        for field, _, _ in self.fields:
            if field in self.entries and col_idx < len(values):
                widget = self.entries[field]
                val = values[col_idx]
                if isinstance(widget, tk.Text):
                    widget.insert("1.0", str(val))
                elif isinstance(widget, ttk.Combobox):
                     widget.set(str(val))
                else:
                    widget.insert(0, str(val))
            col_idx += 1
        if 'on_select' in self.callbacks:
            self.callbacks['on_select'](selected[0])
            
    def on_double_click(self, event):
        selected = self.tree.selection()
        if not selected: return
        if 'on_double_click' in self.callbacks:
            self.callbacks['on_double_click'](selected[0])

# --- SCHEDA STAMPE (HTML) ---
class PrintPanel(ttk.Frame):
    def __init__(self, parent, db):
        super().__init__(parent)
        self.db = db
        self.current_np_id = None
        self.print_proj_map = {}

        self.lbl_title = ttk.Label(self, text="Stampa Singola NP", font=("Arial", 14, "bold"))
        self.lbl_title.pack(pady=(20, 5))
        self.info_lbl = ttk.Label(self, text="Seleziona un NP nella Scheda 2 per abilitare la stampa singola.", font=("Arial", 10))
        self.info_lbl.pack(pady=5)
        self.btn_print = ttk.Button(self, text="GENERA STAMPA NP SELEZIONATO", command=self.generate_html, width=40)
        self.btn_print.pack(pady=10, ipady=5)
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=50, pady=20)
        self.lbl_batch = ttk.Label(self, text="Stampa Massiva per Intervento", font=("Arial", 14, "bold"))
        self.lbl_batch.pack(pady=5)
        ttk.Label(self, text="Seleziona un progetto per stampare tutti i suoi NP:").pack(pady=5)
        self.combo_stampa_progetti = ttk.Combobox(self, state="readonly", width=60, postcommand=self.load_projects_for_print)
        self.combo_stampa_progetti.pack(pady=5)
        self.btn_print_all = ttk.Button(self, text="GENERA TUTTE LE STAMPE DEL PROGETTO", command=self.generate_batch_html, width=40)
        self.btn_print_all.pack(pady=10, ipady=5)
        self.lbl_status = ttk.Label(self, text="", foreground="green", font=("Arial", 9, "italic"))
        self.lbl_status.pack(pady=20)

    def set_current_np(self, np_id):
        self.current_np_id = np_id
        if np_id:
            np = self.db.fetch_one("nuovi_prezzi", "WHERE id=?", (np_id,))
            self.info_lbl.config(text=f"NP Selezionato: {np[2]} - {np[3]}")
        else:
            self.info_lbl.config(text="Nessun NP selezionato nella Scheda 2.")

    def load_projects_for_print(self):
        projs = self.db.fetch_all("progetti")
        values = []
        self.print_proj_map = {}
        for p in projs:
            label = f"{p[1]} - {p[2]}"
            values.append(label)
            self.print_proj_map[label] = p[0]
        self.combo_stampa_progetti['values'] = values

    def generate_batch_html(self):
        selected_label = self.combo_stampa_progetti.get()
        if not selected_label:
            messagebox.showwarning("Attenzione", "Seleziona un progetto dal menu.")
            return
        proj_id = self.print_proj_map[selected_label]
        nps = self.db.fetch_all("nuovi_prezzi", "WHERE progetto_id=?", (proj_id,))
        if not nps:
            messagebox.showinfo("Info", "Nessun NP associato a questo progetto.")
            return
        if messagebox.askyesno("Conferma", f"Verranno generati {len(nps)} file HTML. Continuare?"):
            original_np_id = self.current_np_id
            count = 0
            for np in nps:
                self.current_np_id = np[0]
                self.generate_html()
                count += 1
            self.current_np_id = original_np_id
            self.lbl_status.config(text=f"Operazione completata: {count} stampe generate.")
            messagebox.showinfo("Successo", f"Generati {count} file nella cartella NP_STAMPE.")

    def generate_html(self):
        if not hasattr(self, 'current_np_id') or not self.current_np_id:
            messagebox.showerror("Errore", "Nessun NP selezionato.")
            return
        np = self.db.fetch_one("nuovi_prezzi", "WHERE id=?", (self.current_np_id,))
        items = self.db.fetch_all("voci_costo", "WHERE np_id=? ORDER BY categoria, ordine", (self.current_np_id,))
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{timestamp}_NP_{np[2]}.html"
        if getattr(sys, 'frozen', False):
            current_path = Path(sys.executable).parent
        else:
            current_path = Path(__file__).parent.resolve()
        target_dir_path = current_path.parent / "NP_STAMPE"
        target_dir_path.mkdir(parents=True, exist_ok=True)
        full_path = str(target_dir_path / filename)

        css = """
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 10px; color: #333; max-width: 900px; margin: auto; }
        h1 { color: #0055aa; border-bottom: 2px solid #0055aa; padding-bottom: 10px; font-size: 24px; }
        h2 { background-color: #f2f2f2; padding: 5px; margin-top: 30px; border-left: 5px solid #0055aa; font-size: 18px; }
        .header-box { background: #eef; padding: 10px; border-radius: 5px; margin-bottom: 20px; border: 1px solid #ccd; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 14px; }
        th, td { border: 1px solid #ddd; padding: 3px; text-align: left; }
        th { background-color: #0055aa; color: white; text-transform: uppercase; font-size: 12px; }
        .num { text-align: right; }
        .footer { font-size: 0.8em; color: #777; margin-top: 40px; text-align: center; border-top: 1px solid #ddd; padding-top: 10px; }
        .small-note { font-size: 0.85em; color: #666; margin-top: 2px; display: block; }
        """
        html = f"""
        <!DOCTYPE html>
        <html lang="it">
        <head><meta charset="UTF-8"><title>Analisi Prezzo {np[2]}</title><style>{css}</style></head>
        <body>
            <h1>ANALISI NUOVO PREZZO: {np[2]}</h1>
            <div class="header-box">
                <table style="width: 100%; border: none;">
                    <tr><td colspan="2" style="border: none; padding: 2px 0;"><strong>Descrizione:</strong> {np[3]}</td></tr>
                    <tr style="font-size: 1.1em;">
                        <td style="border: none; width: 50%;"><strong>Unità di Misura: {np[4]}</strong></td>
                        <td style="border: none; text-align: right; width: 50%;"><strong>Prezzo d'applicazione: {format_currency(np[8])}</strong></td>
                    </tr>
                </table>
            </div>
            <h2>Dettaglio Voci Costitutive</h2>
            <table><thead><tr><th width="5%">N.</th><th>Descrizione</th><th width="5%">U.M.</th><th width="10%">Quantità</th><th width="15%">P. Unitario</th><th width="15%">Totale</th></tr></thead><tbody>
        """
        current_cat, tot_a, tot_manodopera = "", 0.0, 0.0
        for item in items:
            cat, desc, um, q, pu = item[3], item[4], item[5], item[6], item[7]
            tot = q * pu
            tot_a += tot
            if "Manodopera" in cat: tot_manodopera += tot
            if cat != current_cat:
                html += f"""<tr><td colspan="6" style="background-color:#e0e0e0; font-weight:bold; color:#333; padding-left: 10px;">{cat.upper()}</td></tr>"""
                current_cat = cat
            html += f"""<tr><td></td><td>{desc}</td><td>{um}</td><td class="num">{q:.3f}</td><td class="num">€ {pu:.3f}</td><td class="num">€ {tot:.2f}</td></tr>"""

        sg_val = tot_a * (np[5]/100)
        sic_val = sg_val * (np[6]/100)
        utili_val = (tot_a + sg_val) * (np[7]/100)
        perc_man = (tot_manodopera / np[8] * 100) if np[8] > 0 else 0
        perc_sic = (sic_val / np[8] * 100) if np[8] > 0 else 0

        html += f"""
                </tbody></table>
            <h2>Riepilogo economico</h2>
            <table style="width: 100%; border: 2px solid #ddd;">
                <tr><td>(A) Sommano</td><td class="num">{format_currency(tot_a)}</td></tr>
                <tr><td>(B) Spese Generali = {np[5]}% di A <span class="small-note"> - compresi oneri di sicurezza per euro {format_currency(sic_val)} = {np[6]}% di B</span></td><td class="num">{format_currency(sg_val)}</td></tr>
                <tr><td>(C) Utile d'Impresa = {np[7]}% di (A + B)</td><td class="num">{format_currency(utili_val)}</td></tr>
                <tr style="background-color: #dbeeff;"><td style="font-weight: bold;">(D) TOTALE (A + B + C)</td><td class="num" style="font-weight: bold;">{format_currency(np[8])}</td></tr>
            </table>
            <h2>Riepilogo incidenze</h2>
            <table style="width: 100%; border: 1px solid #ddd;">
                <tr><td>Incidenza Manodopera</td><td class="num">{perc_man:.2f}%</td><td class="num">{format_currency(tot_manodopera)}</td></tr>
                <tr><td>Incidenza Sicurezza</td><td class="num">{perc_sic:.2f}%</td><td class="num">{format_currency(sic_val)}</td></tr>
            </table>
            <div class="footer">Generato il {datetime.datetime.now().strftime("%d/%m/%Y")}</div>
        </body></html>
        """
        with open(full_path, "w", encoding="utf-8") as f: f.write(html)
        self.lbl_status.config(text=f"Stampa salvata in:\n{full_path}")
        try:
            if os.name == 'nt': os.startfile(full_path)
            else: os.system(f"open '{full_path}'")
        except: pass

# --- SCHEDA CONVERTITORE PDF (INTEGRATA) ---
class ConverterPanel(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        # Directory di default: NP_STAMPE
        if getattr(sys, 'frozen', False):
            base_path = Path(sys.executable).parent
        else:
            base_path = Path(__file__).parent.resolve()
        
        self.work_dir = str(base_path.parent / "NP_STAMPE")
        if not os.path.exists(self.work_dir):
            os.makedirs(self.work_dir, exist_ok=True)

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Selezione cartella
        dir_frame = ttk.LabelFrame(main_frame, text="Cartella di lavoro HTML", padding="5")
        dir_frame.pack(fill="x", pady=5)
        self.dir_label = ttk.Label(dir_frame, text=self.work_dir, relief="sunken", anchor="w")
        self.dir_label.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(dir_frame, text="Cambia...", command=self.select_directory).pack(side="right")

        # Lista file
        files_frame = ttk.LabelFrame(main_frame, text="File HTML disponibili per conversione", padding="5")
        files_frame.pack(fill="both", expand=True, pady=5)
        scrollbar = ttk.Scrollbar(files_frame)
        scrollbar.pack(side="right", fill="y")
        self.files_listbox = tk.Listbox(files_frame, selectmode="multiple", yscrollcommand=scrollbar.set, height=10)
        self.files_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.files_listbox.yview)

        # Pulsanti
        btn_f = ttk.Frame(main_frame)
        btn_f.pack(pady=5)
        ttk.Button(btn_f, text="Seleziona tutti", command=lambda: self.files_listbox.selection_set(0, "end")).pack(side="left", padx=2)
        ttk.Button(btn_f, text="Deseleziona", command=lambda: self.files_listbox.selection_clear(0, "end")).pack(side="left", padx=2)
        ttk.Button(btn_f, text="Aggiorna lista", command=self.refresh_file_list).pack(side="left", padx=2)

        # Opzioni
        opt_f = ttk.LabelFrame(main_frame, text="Opzioni PDF", padding="5")
        opt_f.pack(fill="x", pady=5)
        self.mode = tk.StringVar(value="single")
        ttk.Radiobutton(opt_f, text="PDF separati (A4)", variable=self.mode, value="single").pack(anchor="w")
        ttk.Radiobutton(opt_f, text="Unisci in unico PDF", variable=self.mode, value="merged").pack(anchor="w")
        
        merge_n = ttk.Frame(opt_f)
        merge_n.pack(fill="x")
        ttk.Label(merge_n, text="Nome file unito:").pack(side="left", padx=5)
        self.merged_name = ttk.Entry(merge_n)
        self.merged_name.insert(0, "Progetto_Completo.pdf")
        self.merged_name.pack(side="left", fill="x", expand=True)

        self.btn_conv = ttk.Button(main_frame, text="AVVIA CONVERSIONE PDF", command=self.convert_to_pdf)
        self.btn_conv.pack(pady=10, ipady=5)

        self.status = ttk.Label(main_frame, text="Pronto", relief="sunken", anchor="w")
        self.status.pack(fill="x")
        self.prog = ttk.Progressbar(main_frame, mode='determinate')
        self.prog.pack(fill="x", pady=5)
        
        self.refresh_file_list()

    def select_directory(self):
        d = filedialog.askdirectory(initialdir=self.work_dir)
        if d:
            self.work_dir = d
            self.dir_label.config(text=d)
            self.refresh_file_list()

    def refresh_file_list(self):
        self.files_listbox.delete(0, "end")
        if os.path.exists(self.work_dir):
            files = sorted([f for f in os.listdir(self.work_dir) if f.lower().endswith('.html')])
            for f in files: self.files_listbox.insert("end", f)

    def get_chrome_path(self):
        sys_p = platform.system()
        if sys_p == "Darwin":
            ps = ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"]
        elif sys_p == "Windows":
            ps = ["C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe", "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", os.path.expanduser("~\\AppData\\Local\\Microsoft\\Edge\\Application\\msedge.exe")]
        else:
            ps = ["google-chrome", "microsoft-edge", "chromium"]
        for p in ps:
            if os.path.exists(p) if os.path.isabs(p) else True: return p
        return None

    def convert_to_pdf(self):
        sel = self.files_listbox.curselection()
        if not sel:
            messagebox.showwarning("Attenzione", "Seleziona almeno un file HTML")
            return
        
        c_path = self.get_chrome_path()
        if not c_path:
            messagebox.showerror("Errore", "Chrome o Edge non trovati. Installa un browser basato su Chromium.")
            return

        files = [self.files_listbox.get(i) for i in sel]
        mode = self.mode.get()
        self.prog['maximum'] = len(files)
        temp_pdfs = []

        try:
            for i, f in enumerate(files):
                h_p = os.path.abspath(os.path.join(self.work_dir, f))
                if mode == "single":
                    pdf_p = h_p.replace('.html', '.pdf')
                else:
                    t_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
                    pdf_p = t_pdf.name
                    t_pdf.close()
                    temp_pdfs.append(pdf_p)

                self.status.config(text=f"Conversione {i+1}/{len(files)}...")
                self.update_idletasks()
                
                f_url = f"file:///{h_p.replace(os.sep, '/')}" if platform.system() == "Windows" else f"file://{h_p}"
                cmd = [c_path, "--headless", "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf=" + pdf_p, f_url]
                subprocess.run(cmd, check=True, timeout=30)
                self.prog['value'] = i + 1

            if mode == "merged":
                from pypdf import PdfWriter, PdfReader
                out_n = self.merged_name.get()
                if not out_n.endswith('.pdf'): out_n += ".pdf"
                out_p = os.path.join(self.work_dir, out_n)
                writer = PdfWriter()
                for p in temp_pdfs:
                    r = PdfReader(p)
                    for page in r.pages: writer.add_page(page)
                with open(out_p, "wb") as o: writer.write(o)
                for p in temp_pdfs: os.unlink(p)
                messagebox.showinfo("Successo", f"PDF Unito creato: {out_n}")
            else:
                messagebox.showinfo("Successo", "Conversione completata.")
        except Exception as e:
            messagebox.showerror("Errore", str(e))
        finally:
            self.status.config(text="Pronto")
            self.prog['value'] = 0

# --- APP PRINCIPALE ---
class NPApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestione Nuovi Prezzi (NP)")
        self.geometry("1100x850")
        self.db = Database()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

        # TAB 1: PROGETTI
        self.tab_progetti = CrudPanel(
            self.notebook, self.db, "progetti", 
            [("codice", "Codice", 100), ("titolo", "Titolo Intervento", 300), ("cup", "CUP", 100), ("committente", "Committente", 200)],
            callbacks={'on_select': self.on_project_select, 'on_double_click': self.goto_tab_np}
        )
        self.notebook.add(self.tab_progetti, text="1. Progetti")

        # TAB 2: ELENCO NP
        self.frame_tab2 = ttk.Frame(self.notebook) 
        self.notebook.add(self.frame_tab2, text="2. Elenco NP")
        self.lbl_project_title = ttk.Label(self.frame_tab2, text="Nessun Progetto Selezionato", font=("Arial", 12, "bold"), foreground="#444", padding=10)
        self.lbl_project_title.pack(side="top", fill="x")
        self.current_project_id = None
        self.tab_np = CrudPanel(
            self.frame_tab2, self.db, "nuovi_prezzi",
            [("codice", "Codice NP", 100), ("descrizione", "Descrizione NP", 400), ("unita_misura", "Unità di Misura", 100), ("prezzo_finale", "Prezzo Finale", 120)],
            callbacks={'on_select': self.on_np_select, 'on_double_click': self.goto_tab_details}
        )
        self.tab_np.pack(fill="both", expand=True)
        
        # Override bottoni Tab 2
        def add_np_wrapper():
            if not self.current_project_id:
                messagebox.showwarning("Attenzione", "Seleziona prima un Progetto")
                return
            data = self.tab_np.get_data_from_ui()
            data['progetto_id'] = self.current_project_id
            self.db.insert("nuovi_prezzi", data)
            self.tab_np.refresh_data("WHERE progetto_id=?", (self.current_project_id,))
        
        self.tab_np.buttons['add'].config(command=add_np_wrapper)
        self.tab_np.buttons['dup'].config(command=lambda: self._copy_np_logic(False))
        self.tab_np.buttons['imp'].config(command=lambda: ImportDialog(self, self.db, self.current_project_id, self._import_np_callback))

        # TAB 3: MODIFICA NP
        self.frame_det_np = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_det_np, text="3. Modifica NP")
        self.setup_details_tab()

        # TAB 4: UM
        self.tab_admin = CrudPanel(self.notebook, self.db, "unita_misura", [("codice", "Codice", 80), ("nome", "Unità", 150), ("descrizione", "Utilizzo", 400)])
        self.notebook.add(self.tab_admin, text="4. Unità di Misura")

        # TAB 5: STAMPE
        self.tab_print = PrintPanel(self.notebook, self.db)
        self.notebook.add(self.tab_print, text="5. Stampe HTML")

        # TAB 6: CONVERTITORE PDF
        self.tab_conv = ConverterPanel(self.notebook)
        self.notebook.add(self.tab_conv, text="6. Convertitore PDF")

    def _import_np_callback(self, source_np_id):
        self._copy_np_logic(True, source_np_id)

    def _copy_np_logic(self, is_import=False, source_id_override=None):
        if not self.current_project_id: return
        source_id = source_id_override if is_import else (self.tab_np.tree.selection()[0] if self.tab_np.tree.selection() else None)
        if not source_id: return
        src = self.db.fetch_one("nuovi_prezzi", "WHERE id=?", (source_id,))
        new_id = self.db.insert("nuovi_prezzi", {
            'progetto_id': self.current_project_id, 'codice': src[2] + "_cp", 'descrizione': src[3] + " (Copia)",
            'unita_misura': src[4], 'perc_spese_generali': src[5], 'perc_sicurezza': src[6], 'perc_utili': src[7], 'prezzo_finale': src[8]
        })
        items = self.db.fetch_all("voci_costo", "WHERE np_id=?", (source_id,))
        for it in items:
            self.db.insert("voci_costo", {'np_id': new_id, 'ordine': it[2], 'categoria': it[3], 'descrizione': it[4], 'um': it[5], 'quantita': it[6], 'prezzo_unitario': it[7]})
        self.tab_np.refresh_data("WHERE progetto_id=?", (self.current_project_id,))

    def on_tab_change(self, event):
        idx = self.notebook.index("current")
        if idx == 1 and self.current_project_id:
            self.tab_np.refresh_data("WHERE progetto_id=?", (self.current_project_id,))
        elif idx == 4:
            self.tab_print.set_current_np(getattr(self, 'current_np_id', None))
        elif idx == 5:
            self.tab_conv.refresh_file_list()

    def on_project_select(self, project_id):
        self.current_project_id = project_id
        p = self.db.fetch_one("progetti", "WHERE id=?", (project_id,))
        if p: self.lbl_project_title.config(text=f"Progetto: {p[2]}")
        self.tab_np.refresh_data("WHERE progetto_id=?", (project_id,))

    def goto_tab_np(self, project_id):
        self.on_project_select(project_id)
        self.notebook.select(1)

    def on_np_select(self, np_id):
        self.current_np_id = np_id
        self.refresh_details_tree()

    def goto_tab_details(self, np_id):
        self.on_np_select(np_id)
        self.notebook.select(2)

    def setup_details_tab(self):
        self.lbl_np_details = ttk.Label(self.frame_det_np, text="Nessun NP Selezionato", font=("Arial", 12, "bold"), foreground="#0055aa", padding=10)
        self.lbl_np_details.pack(fill="x")
        f_in = ttk.LabelFrame(self.frame_det_np, text="Nuova Voce", padding=10)
        f_in.pack(fill="x", padx=10)
        
        self.entry_order = ttk.Entry(f_in, width=5)
        self.combo_cat = ttk.Combobox(f_in, values=["Manodopera", "Mezzi e Noli", "Materiali"], width=15)
        self.txt_desc = tk.Text(f_in, height=2, width=40)
        self.combo_um = ttk.Combobox(f_in, width=10)
        self.entry_q = ttk.Entry(f_in, width=10)
        self.entry_pu = ttk.Entry(f_in, width=10)

        ttk.Label(f_in, text="Ord").grid(row=0, column=0)
        self.entry_order.grid(row=1, column=0)
        ttk.Label(f_in, text="Cat").grid(row=0, column=1)
        self.combo_cat.grid(row=1, column=1)
        ttk.Label(f_in, text="Descrizione").grid(row=0, column=2)
        self.txt_desc.grid(row=1, column=2, padx=5)
        ttk.Label(f_in, text="UM").grid(row=0, column=3)
        self.combo_um.grid(row=1, column=3)
        ttk.Label(f_in, text="Q.tà").grid(row=0, column=4)
        self.entry_q.grid(row=1, column=4)
        ttk.Label(f_in, text="P.U.").grid(row=0, column=5)
        self.entry_pu.grid(row=1, column=5)

        btn_f = ttk.Frame(self.frame_det_np)
        btn_f.pack(pady=5)
        ttk.Button(btn_f, text="Aggiungi", command=self.add_cost_item).pack(side="left", padx=5)
        ttk.Button(btn_f, text="Elimina", command=self.delete_cost_item).pack(side="left", padx=5)

        cols = ("ord", "cat", "desc", "um", "q", "pu", "tot")
        self.tree_det = ttk.Treeview(self.frame_det_np, columns=cols, show="headings", height=10)
        for c in cols: self.tree_det.heading(c, text=c.upper()); self.tree_det.column(c, width=80)
        self.tree_det.pack(fill="both", expand=True, padx=10)

        self.frame_summary = ttk.LabelFrame(self.frame_det_np, text="Riepilogo", padding=10)
        self.frame_summary.pack(fill="x", padx=10, pady=10)
        self.lbl_total_final = ttk.Label(self.frame_summary, text="TOTALE: € 0,00", font=("Arial", 12, "bold"))
        self.lbl_total_final.pack(side="right")
        
        # Campi percentuali
        perc_f = ttk.Frame(self.frame_summary)
        perc_f.pack(side="left")
        self.entry_perc_spese = ttk.Entry(perc_f, width=5)
        self.entry_perc_spese.insert(0, "17")
        self.entry_perc_utili = ttk.Entry(perc_f, width=5)
        self.entry_perc_utili.insert(0, "10")
        ttk.Label(perc_f, text="SG %:").pack(side="left")
        self.entry_perc_spese.pack(side="left", padx=5)
        ttk.Label(perc_f, text="Utili %:").pack(side="left")
        self.entry_perc_utili.pack(side="left", padx=5)
        ttk.Button(perc_f, text="Ricalcola", command=self.recalculate_totals).pack(side="left", padx=10)

    def refresh_details_tree(self):
        for row in self.tree_det.get_children(): self.tree_det.delete(row)
        if not getattr(self, 'current_np_id', None): return
        items = self.db.fetch_all("voci_costo", "WHERE np_id=? ORDER BY ordine", (self.current_np_id,))
        self.total_a = 0.0
        for it in items:
            t = round(it[6] * it[7], 2)
            self.total_a += t
            self.tree_det.insert("", "end", iid=it[0], values=(it[2], it[3], it[4], it[5], it[6], it[7], t))
        self.recalculate_totals()

    def recalculate_totals(self):
        if not getattr(self, 'current_np_id', None): return
        s = safe_float_convert(self.entry_perc_spese.get())
        u = safe_float_convert(self.entry_perc_utili.get())
        val_b = self.total_a * (s/100)
        val_c = (self.total_a + val_b) * (u/100)
        tot = self.total_a + val_b + val_c
        self.lbl_total_final.config(text=f"TOTALE: {format_currency(tot)}")
        self.db.update("nuovi_prezzi", self.current_np_id, {'perc_spese_generali': s, 'perc_utili': u, 'prezzo_finale': tot})

    def add_cost_item(self):
        if not getattr(self, 'current_np_id', None): return
        d = {
            'np_id': self.current_np_id, 'ordine': self.entry_order.get(), 'categoria': self.combo_cat.get(),
            'descrizione': self.txt_desc.get("1.0", "end-1c"), 'um': self.combo_um.get(),
            'quantita': safe_float_convert(self.entry_q.get()), 'prezzo_unitario': safe_float_convert(self.entry_pu.get())
        }
        self.db.insert("voci_costo", d)
        self.refresh_details_tree()

    def delete_cost_item(self):
        sel = self.tree_det.selection()
        if sel:
            self.db.delete("voci_costo", sel[0])
            self.refresh_details_tree()

# --- CONTROLLO DIPENDENZE ALL'AVVIO ---
def check_pypdf_dependency():
    """Verifica se pypdf è installato per le funzioni di unione PDF"""
    try:
        import pypdf
    except ImportError:
        # Mostra un avviso non bloccante
        messagebox.showinfo(
            "Dipendenza Mancante", 
            "Attenzione: La libreria 'pypdf' non è installata.\n\n"
            "Il programma funzionerà correttamente per la generazione dei singoli PDF, "
            "ma la funzione 'Unisci in unico PDF' nella Tab 6 non sarà disponibile.\n\n"
            "Per abilitarla sul tuo computer, esegui: pip install pypdf"
        )

# --- AVVIO APPLICAZIONE ---
if __name__ == "__main__":
    app = NPApp()
    
    # Esegue il controllo dopo 100ms dall'avvio per assicurarsi 
    # che la finestra principale sia già visibile
    app.after(100, check_pypdf_dependency)
    
    app.mainloop()