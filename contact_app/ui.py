from __future__ import annotations

"""
Interfaz gráfica (Tkinter/ttk) de la agenda de contactos.

Incluye:
- CRUD
- búsqueda exacta y fuzzy (RapidFuzz)
- importación/exportación CSV
- deduplicación con fusión
- mejoras de UX (menú, tooltips, modo oscuro, etc.)
"""

import csv
import sqlite3
import tkinter as tk
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from typing import Callable, Optional

from contact_app import db
from contact_app.duplicates import RapidFuzzNotAvailable, find_duplicate_groups, is_available, merge_contacts
from contact_app.models import Contact
from contact_app.validators import validate_email, validate_name, validate_phone


class ContactApp:
    def __init__(self, root: tk.Tk, conn: sqlite3.Connection, on_close: Optional[Callable[[], None]] = None) -> None:
        self.root = root
        self.conn = conn
        self.on_close = on_close
        self.selected_contact_id: Optional[int] = None
        self._placeholders: dict[ttk.Entry, str] = {}
        self.dark_mode_var = tk.BooleanVar(value=False)
        self._palette: dict[str, str] = {}

        self.search_var = tk.StringVar()
        self.rf_status_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.phone_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self._sort_col: Optional[str] = None
        self._sort_reverse = False

        self._build_style()
        self._build_layout()
        self._wire_events()
        self.refresh_contacts()

    def refresh_contacts(self) -> None:
        search = self._get_search_text()
        if search and is_available():
            contacts = self._fuzzy_search(search)
        else:
            contacts = db.list_contacts(self.conn, search=search if search else None)
        self._render_contacts(contacts)
        self._update_status(displayed=len(contacts), search=search)

    def _fuzzy_search(self, query: str) -> list[Contact]:
        try:
            from rapidfuzz import fuzz
        except Exception:
            return db.list_contacts(self.conn, search=query)

        q = query.strip().lower()
        if not q:
            return db.list_contacts(self.conn)

        contacts = db.list_contacts(self.conn)
        scored: list[tuple[int, Contact]] = []
        for c in contacts:
            name_score = int(fuzz.WRatio(q, c.name.lower()))
            phone_score = 100 if q in (c.phone or "") else 0
            email_score = int(fuzz.partial_ratio(q, c.email.lower())) if c.email else 0
            score = max(name_score, phone_score, email_score)
            if score >= 60:
                scored.append((score, c))

        scored.sort(key=lambda t: (-t[0], t[1].name.lower(), t[1].id or 0))
        return [c for _, c in scored]

    def _update_status(self, displayed: int, search: str) -> None:
        try:
            total = db.count_contacts(self.conn)
        except sqlite3.Error:
            total = displayed
        mode = "fuzzy" if (search.strip() and is_available()) else "exacta"
        search_part = f' | Buscar="{search.strip()}" ({mode})' if search.strip() else ""
        self.status_var.set(f"Mostrando {displayed} de {total}{search_part}")

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        self._apply_theme("dark" if self.dark_mode_var.get() else "light")

    def _apply_theme(self, mode: str) -> None:
        light = {
            "bg": "#F5F7FB",
            "text": "#111827",
            "muted": "#4B5563",
            "card_bg": "#FFFFFF",
            "border": "#E5E7EB",
            "entry_bg": "#FFFFFF",
            "entry_fg": "#111827",
            "placeholder": "#9CA3AF",
            "accent": "#2563EB",
            "accent_active": "#1D4ED8",
            "accent_pressed": "#1E40AF",
            "danger": "#DC2626",
            "danger_active": "#B91C1C",
            "danger_pressed": "#991B1B",
            "secondary": "#E5E7EB",
            "secondary_active": "#D1D5DB",
            "secondary_pressed": "#CBD5E1",
            "tree_bg": "#FFFFFF",
            "tree_odd": "#F3F4F6",
            "tree_head_bg": "#E5E7EB",
            "selected_bg": "#BFDBFE",
            "badge_ok_bg": "#DCFCE7",
            "badge_ok_fg": "#166534",
            "badge_warn_bg": "#FEF3C7",
            "badge_warn_fg": "#92400E",
            "tip_bg": "#111827",
            "tip_fg": "#F9FAFB",
        }
        dark = {
            "bg": "#0B1220",
            "text": "#E5E7EB",
            "muted": "#9CA3AF",
            "card_bg": "#111827",
            "border": "#1F2937",
            "entry_bg": "#0F172A",
            "entry_fg": "#E5E7EB",
            "placeholder": "#64748B",
            "accent": "#3B82F6",
            "accent_active": "#2563EB",
            "accent_pressed": "#1D4ED8",
            "danger": "#EF4444",
            "danger_active": "#DC2626",
            "danger_pressed": "#B91C1C",
            "secondary": "#1F2937",
            "secondary_active": "#334155",
            "secondary_pressed": "#0F172A",
            "tree_bg": "#0F172A",
            "tree_odd": "#111827",
            "tree_head_bg": "#1F2937",
            "selected_bg": "#1D4ED8",
            "badge_ok_bg": "#064E3B",
            "badge_ok_fg": "#D1FAE5",
            "badge_warn_bg": "#78350F",
            "badge_warn_fg": "#FFFBEB",
            "tip_bg": "#0B1220",
            "tip_fg": "#E5E7EB",
        }

        palette = dark if mode == "dark" else light
        self._palette = palette
        self.root.configure(bg=palette["bg"])
        style = ttk.Style(self.root)

        style.configure(".", font=("Segoe UI", 10))
        style.configure("TFrame", background=palette["bg"])
        style.configure("TLabel", background=palette["bg"], foreground=palette["text"])
        style.configure("Card.TFrame", background=palette["card_bg"])
        style.configure("Card.TLabel", background=palette["card_bg"], foreground=palette["text"])
        style.configure("CardTitle.TLabel", background=palette["card_bg"], foreground=palette["text"], font=("Segoe UI", 14, "bold"))
        style.configure("CardSubtitle.TLabel", background=palette["card_bg"], foreground=palette["muted"])
        style.configure("BadgeOk.TLabel", background=palette["badge_ok_bg"], foreground=palette["badge_ok_fg"], padding=(8, 3))
        style.configure("BadgeWarn.TLabel", background=palette["badge_warn_bg"], foreground=palette["badge_warn_fg"], padding=(8, 3))

        style.configure(
            "TEntry",
            padding=6,
            fieldbackground=palette["entry_bg"],
            foreground=palette["entry_fg"],
        )

        style.configure(
            "TButton",
            padding=(12, 7),
        )

        style.configure(
            "Accent.TButton",
            background=palette["accent"],
            foreground="#FFFFFF",
            padding=(12, 7),
        )
        style.map(
            "Accent.TButton",
            background=[("active", palette["accent_active"]), ("pressed", palette["accent_pressed"])],
            foreground=[("disabled", palette["muted"])],
        )

        style.configure(
            "Danger.TButton",
            background=palette["danger"],
            foreground="#FFFFFF",
            padding=(12, 7),
        )
        style.map(
            "Danger.TButton",
            background=[("active", palette["danger_active"]), ("pressed", palette["danger_pressed"])],
            foreground=[("disabled", palette["muted"])],
        )

        style.configure(
            "Secondary.TButton",
            background=palette["secondary"],
            foreground=palette["text"],
            padding=(12, 7),
        )
        style.map(
            "Secondary.TButton",
            background=[("active", palette["secondary_active"]), ("pressed", palette["secondary_pressed"])],
        )

        style.configure(
            "Treeview",
            background=palette["tree_bg"],
            fieldbackground=palette["tree_bg"],
            foreground=palette["text"],
            rowheight=26,
            borderwidth=0,
            relief="flat",
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background=palette["tree_head_bg"],
            foreground=palette["text"],
            padding=(8, 6),
        )
        style.map(
            "Treeview",
            background=[("selected", palette["selected_bg"])],
            foreground=[("selected", "#FFFFFF")],
        )

        ToolTip.BG = palette["tip_bg"]
        ToolTip.FG = palette["tip_fg"]

        if hasattr(self, "tree"):
            self.tree.tag_configure("even", background=palette["tree_bg"])
            self.tree.tag_configure("odd", background=palette["tree_odd"])
        if hasattr(self, "rf_badge_label"):
            badge_style = "BadgeOk.TLabel" if is_available() else "BadgeWarn.TLabel"
            self.rf_badge_label.configure(style=badge_style)

    def _toggle_theme(self) -> None:
        self._apply_theme("dark" if self.dark_mode_var.get() else "light")
        self._apply_placeholders()

    def _set_entry_color(self, entry: ttk.Entry, color: str) -> None:
        try:
            entry.configure(foreground=color)
        except tk.TclError:
            return

    def _set_placeholder(self, entry: ttk.Entry, var: tk.StringVar, text: str) -> None:
        already_bound = entry in self._placeholders
        self._placeholders[entry] = text

        if already_bound:
            if not var.get().strip():
                var.set(text)
                self._set_entry_color(entry, "#9CA3AF")
            return

        def on_focus_in(_evt: object) -> None:
            if var.get() == text:
                var.set("")
                self._set_entry_color(entry, "#111827")

        def on_focus_out(_evt: object) -> None:
            if not var.get().strip():
                var.set(text)
                self._set_entry_color(entry, "#9CA3AF")

        entry.bind("<FocusIn>", on_focus_in, add=True)
        entry.bind("<FocusOut>", on_focus_out, add=True)
        if not var.get().strip():
            var.set(text)
            self._set_entry_color(entry, "#9CA3AF")

    def _value_from_entry(self, entry: ttk.Entry, var: tk.StringVar) -> str:
        text = var.get()
        placeholder = self._placeholders.get(entry)
        if placeholder and text == placeholder:
            return ""
        return text

    def _get_search_text(self) -> str:
        return self._value_from_entry(self.search_entry, self.search_var).strip() if hasattr(self, "search_entry") else ""

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Importar CSV...", command=self.on_import_csv, accelerator="Ctrl+O")
        file_menu.add_command(label="Exportar CSV...", command=self.on_export_csv, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Salir", command=self._request_close)
        menubar.add_cascade(label="Archivo", menu=file_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Detectar duplicados", command=self.on_detect_duplicates)
        menubar.add_cascade(label="Herramientas", menu=tools_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_checkbutton(label="Modo oscuro", variable=self.dark_mode_var, command=self._toggle_theme, accelerator="Ctrl+D")
        menubar.add_cascade(label="Ver", menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Acerca de", command=self.on_about)
        menubar.add_cascade(label="Ayuda", menu=help_menu)

        self.root.config(menu=menubar)

    def _request_close(self) -> None:
        if self.on_close is not None:
            self.on_close()
            return
        self.root.destroy()

    def on_about(self) -> None:
        messagebox.showinfo(
            "Acerca de",
            "Agenda inteligente de contactos\n\n"
            "Funciones:\n"
            "- CRUD (añadir/actualizar/eliminar)\n"
            "- Búsqueda en tiempo real (exacta o fuzzy con RapidFuzz)\n"
            "- Detección y fusión de duplicados (RapidFuzz)\n"
            "- Importar / Exportar CSV\n\n"
            "Atajos:\n"
            "- Ctrl+O: importar\n"
            "- Ctrl+S: exportar\n"
            "- Esc: limpiar formulario",
            parent=self.root,
        )

    def _build_layout(self) -> None:
        self.root.title("Agenda Inteligente de Contactos")
        self.root.minsize(820, 520)
        self._build_menu()

        header_wrap = ttk.Frame(self.root, padding=10)
        header_wrap.pack(fill=tk.X)

        header = ttk.Frame(header_wrap, style="Card.TFrame", padding=12)
        header.pack(fill=tk.X)
        header.configure(relief="solid", borderwidth=1)

        left = ttk.Frame(header, style="Card.TFrame")
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(left, text="Agenda inteligente", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            left,
            text="CRUD, búsqueda instantánea, import/export CSV y deduplicación con RapidFuzz",
            style="CardSubtitle.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        right = ttk.Frame(header, style="Card.TFrame")
        right.pack(side=tk.RIGHT, anchor="e")
        self.rf_status_var.set("RapidFuzz: OK" if is_available() else "RapidFuzz: no instalado")
        badge_style = "BadgeOk.TLabel" if is_available() else "BadgeWarn.TLabel"
        self.rf_badge_label = ttk.Label(right, textvariable=self.rf_status_var, style=badge_style)
        self.rf_badge_label.pack(anchor="e")

        search_row = ttk.Frame(header_wrap, padding=(2, 10, 2, 0))
        search_row.pack(fill=tk.X)
        ttk.Label(search_row, text="Buscar:").pack(side=tk.LEFT)
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_var, width=46)
        self.search_entry.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(search_row, text="Limpiar búsqueda", command=lambda: self._clear_search(), style="Secondary.TButton").pack(
            side=tk.LEFT, padx=(10, 0)
        )

        mid = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        mid.pack(fill=tk.X)

        form = ttk.Frame(mid)
        form.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(form, text="Nombre y Apellido:").grid(row=0, column=0, sticky="w")
        self.name_entry = ttk.Entry(form, textvariable=self.name_var, width=40)
        self.name_entry.grid(row=0, column=1, sticky="we", padx=(8, 12), pady=4)

        ttk.Label(form, text="Teléfono:").grid(row=1, column=0, sticky="w")
        self.phone_entry = ttk.Entry(form, textvariable=self.phone_var, width=40)
        self.phone_entry.grid(row=1, column=1, sticky="we", padx=(8, 12), pady=4)

        ttk.Label(form, text="Email (opcional):").grid(row=2, column=0, sticky="w")
        self.email_entry = ttk.Entry(form, textvariable=self.email_var, width=40)
        self.email_entry.grid(row=2, column=1, sticky="we", padx=(8, 12), pady=4)

        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(mid)
        actions.pack(side=tk.RIGHT, fill=tk.Y)

        self.add_btn = ttk.Button(actions, text="Añadir", command=self.on_add, style="Accent.TButton")
        self.add_btn.pack(fill=tk.X, pady=2)

        self.update_btn = ttk.Button(actions, text="Actualizar", command=self.on_update, style="Accent.TButton")
        self.update_btn.pack(fill=tk.X, pady=2)

        self.delete_btn = ttk.Button(actions, text="Eliminar", command=self.on_delete, style="Danger.TButton")
        self.delete_btn.pack(fill=tk.X, pady=2)

        self.clear_btn = ttk.Button(actions, text="Limpiar", command=self.on_clear, style="Secondary.TButton")
        self.clear_btn.pack(fill=tk.X, pady=(10, 2))

        self.dup_btn = ttk.Button(actions, text="Detectar duplicados", command=self.on_detect_duplicates, style="Secondary.TButton")
        self.dup_btn.pack(fill=tk.X, pady=(10, 2))

        bottom = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        bottom.pack(fill=tk.BOTH, expand=True)

        cols = ("id", "name", "phone", "email", "created_at")
        self.tree = ttk.Treeview(bottom, columns=cols, show="headings")
        self.tree.heading("id", text="ID", command=lambda: self._on_sort("id"))
        self.tree.heading("name", text="Nombre", command=lambda: self._on_sort("name"))
        self.tree.heading("phone", text="Teléfono", command=lambda: self._on_sort("phone"))
        self.tree.heading("email", text="Email", command=lambda: self._on_sort("email"))
        self.tree.heading("created_at", text="Creado", command=lambda: self._on_sort("created_at"))

        self.tree.column("id", width=60, anchor=tk.CENTER, stretch=False)
        self.tree.column("name", width=240, anchor=tk.W)
        self.tree.column("phone", width=140, anchor=tk.W)
        self.tree.column("email", width=220, anchor=tk.W)
        self.tree.column("created_at", width=120, anchor=tk.W, stretch=False)

        scrollbar = ttk.Scrollbar(bottom, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.tag_configure("odd", background="#F3F4F6")
        self._build_tree_menu()

        if not is_available():
            self.dup_btn.state(["disabled"])

        status = ttk.Frame(self.root, padding=(10, 2))
        status.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Label(status, textvariable=self.status_var).pack(side=tk.LEFT)
        self._update_status(displayed=0, search="")
        self._add_tooltips()

    def _wire_events(self) -> None:
        self.search_var.trace_add("write", lambda *_: self.refresh_contacts())
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.root.bind("<Escape>", lambda *_: self.on_clear())
        self.root.bind("<Control-n>", lambda *_: self.on_clear())
        self.root.bind("<Control-f>", lambda *_: self._focus_search())
        self.root.bind("<Control-d>", lambda *_: self._toggle_theme_key())
        self.root.bind("<Control-s>", lambda *_: self.on_export_csv())
        self.root.bind("<Control-o>", lambda *_: self.on_import_csv())
        self.root.bind("<Delete>", lambda *_: self._delete_if_tree_focused())
        self.root.bind("<Return>", lambda *_: self._submit_form())
        self._apply_placeholders()

    def _render_contacts(self, contacts: list[Contact]) -> None:
        self.tree.delete(*self.tree.get_children())
        for idx, c in enumerate(contacts):
            created = c.created_at.isoformat(timespec="seconds") if c.created_at else ""
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert("", tk.END, values=(c.id, c.name, c.phone, c.email, created), tags=(tag,))

    def on_tree_select(self, _evt: Optional[object] = None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel[0], "values")
        if not values:
            return
        self.selected_contact_id = int(values[0])
        self._clear_placeholders_for_form()
        self.name_var.set(values[1])
        self.phone_var.set(values[2])
        self.email_var.set(values[3])

    def on_add(self) -> None:
        try:
            contact = self._read_form()
        except ValueError as e:
            messagebox.showerror("Validación", str(e), parent=self.root)
            return
        try:
            new_id = db.add_contact(self.conn, contact)
        except sqlite3.Error as e:
            messagebox.showerror("Base de datos", f"No se pudo añadir el contacto.\n\n{e}", parent=self.root)
            return
        self.on_clear()
        self.refresh_contacts()
        self._select_contact(new_id)

    def on_update(self) -> None:
        if self.selected_contact_id is None:
            messagebox.showwarning("Actualizar", "Selecciona un contacto para actualizar.", parent=self.root)
            return
        try:
            contact = self._read_form()
        except ValueError as e:
            messagebox.showerror("Validación", str(e), parent=self.root)
            return
        try:
            db.update_contact(self.conn, self.selected_contact_id, contact)
        except sqlite3.Error as e:
            messagebox.showerror("Base de datos", f"No se pudo actualizar.\n\n{e}", parent=self.root)
            return
        contact_id = self.selected_contact_id
        self.refresh_contacts()
        self._select_contact(contact_id)

    def on_delete(self) -> None:
        if self.selected_contact_id is None:
            messagebox.showwarning("Eliminar", "Selecciona un contacto para eliminar.", parent=self.root)
            return
        if not messagebox.askyesno(
            "Eliminar",
            "¿Seguro que quieres eliminar el contacto seleccionado?",
            parent=self.root,
        ):
            return
        try:
            db.delete_contact(self.conn, self.selected_contact_id)
        except sqlite3.Error as e:
            messagebox.showerror("Base de datos", f"No se pudo eliminar.\n\n{e}", parent=self.root)
            return
        self.on_clear()
        self.refresh_contacts()

    def on_clear(self) -> None:
        self.selected_contact_id = None
        self.name_var.set("")
        self.phone_var.set("")
        self.email_var.set("")
        self.tree.selection_remove(self.tree.selection())
        self.name_entry.focus_set()
        self._apply_placeholders()

    def on_detect_duplicates(self) -> None:
        if not is_available():
            messagebox.showinfo(
                "Duplicados",
                "La función de duplicados requiere instalar RapidFuzz:\n\npip install rapidfuzz",
                parent=self.root,
            )
            return
        try:
            contacts = db.list_contacts(self.conn)
            groups = find_duplicate_groups(contacts, threshold=88)
        except RapidFuzzNotAvailable:
            messagebox.showinfo(
                "Duplicados",
                "La función de duplicados requiere instalar RapidFuzz:\n\npip install rapidfuzz",
                parent=self.root,
            )
            return
        except sqlite3.Error as e:
            messagebox.showerror("Base de datos", f"No se pudo leer la lista de contactos.\n\n{e}", parent=self.root)
            return

        if not groups:
            messagebox.showinfo("Duplicados", "No se han detectado duplicados.", parent=self.root)
            return

        DuplicateWindow(self.root, self.conn, groups, on_merged=self.refresh_contacts)

    def on_export_csv(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Exportar contactos a CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not path:
            return
        try:
            contacts = db.list_contacts(self.conn)
        except sqlite3.Error as e:
            messagebox.showerror("Base de datos", f"No se pudo leer la lista de contactos.\n\n{e}", parent=self.root)
            return

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=["name", "phone", "email"])
                writer.writeheader()
                for c in contacts:
                    writer.writerow({"name": c.name, "phone": c.phone, "email": c.email})
        except OSError as e:
            messagebox.showerror("Exportar", f"No se pudo escribir el archivo.\n\n{e}", parent=self.root)
            return

        messagebox.showinfo("Exportar", f"Exportados {len(contacts)} contactos a:\n{path}", parent=self.root)

    def on_import_csv(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Importar contactos desde CSV",
            filetypes=[("CSV", "*.csv")],
        )
        if not path:
            return
        try:
            with open(path, "r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
        except OSError as e:
            messagebox.showerror("Importar", f"No se pudo leer el archivo.\n\n{e}", parent=self.root)
            return

        added = 0
        errors: list[str] = []
        for idx, row in enumerate(rows, start=2):
            name = _first_key(row, ["name", "nombre", "Nombre"])
            phone = _first_key(row, ["phone", "telefono", "teléfono", "Telefono", "Teléfono"])
            email = _first_key(row, ["email", "correo", "mail", "Email", "Correo"])
            try:
                contact = Contact(
                    id=None,
                    name=validate_name(name),
                    phone=validate_phone(phone),
                    email=validate_email(email),
                )
                db.add_contact(self.conn, contact)
                added += 1
            except (ValueError, sqlite3.Error) as e:
                errors.append(f"Línea {idx}: {e}")

        self.refresh_contacts()

        if not errors:
            messagebox.showinfo("Importar", f"Importados {added} contactos.", parent=self.root)
            return

        shown = "\n".join(errors[:8])
        more = f"\n... y {len(errors) - 8} más" if len(errors) > 8 else ""
        messagebox.showwarning(
            "Importar",
            f"Importados {added} contactos.\nErrores: {len(errors)}\n\n{shown}{more}",
            parent=self.root,
        )

    def _read_form(self) -> Contact:
        name_raw = self._value_from_entry(self.name_entry, self.name_var)
        phone_raw = self._value_from_entry(self.phone_entry, self.phone_var)
        email_raw = self._value_from_entry(self.email_entry, self.email_var)
        name = validate_name(name_raw)
        phone = validate_phone(phone_raw)
        email = validate_email(email_raw)
        return Contact(id=None, name=name, phone=phone, email=email)

    def _select_contact(self, contact_id: int) -> None:
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            if values and int(values[0]) == contact_id:
                self.tree.selection_set(item)
                self.tree.see(item)
                self.tree.focus(item)
                self.on_tree_select()
                return

    def _on_sort(self, col: str) -> None:
        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col = col
            self._sort_reverse = False
        self._sort_tree(col, reverse=self._sort_reverse)

    def _sort_tree(self, col: str, reverse: bool) -> None:
        cols = list(self.tree["columns"])
        if col not in cols:
            return
        idx = cols.index(col)
        items = self.tree.get_children("")
        values = []
        for item in items:
            row = self.tree.item(item, "values")
            values.append((item, row[idx] if idx < len(row) else ""))

        def key(val: str):
            if col == "id":
                try:
                    return int(val)
                except Exception:
                    return 0
            if col == "created_at":
                try:
                    return datetime.fromisoformat(val)
                except Exception:
                    return datetime.min
            return str(val).lower()

        values.sort(key=lambda t: key(t[1]), reverse=reverse)
        for i, (item, _) in enumerate(values):
            self.tree.move(item, "", i)

    def _focus_search(self) -> None:
        self.search_entry.focus_set()
        try:
            self.search_entry.selection_range(0, tk.END)
        except tk.TclError:
            return

    def _clear_search(self) -> None:
        self.search_var.set("")
        self._apply_placeholders()
        self.refresh_contacts()

    def _toggle_theme_key(self) -> None:
        self.dark_mode_var.set(not self.dark_mode_var.get())
        self._toggle_theme()

    def _delete_if_tree_focused(self) -> None:
        focus = self.root.focus_get()
        if focus == self.tree:
            self.on_delete()

    def _submit_form(self) -> None:
        focus = self.root.focus_get()
        if focus not in {self.name_entry, self.phone_entry, self.email_entry}:
            return
        if self.selected_contact_id is None:
            self.on_add()
        else:
            self.on_update()

    def _apply_placeholders(self) -> None:
        if hasattr(self, "search_entry"):
            self._set_placeholder(self.search_entry, self.search_var, "Nombre, teléfono o email...")
        if hasattr(self, "name_entry"):
            self._set_placeholder(self.name_entry, self.name_var, "Ej: Ana García")
        if hasattr(self, "phone_entry"):
            self._set_placeholder(self.phone_entry, self.phone_var, "Ej: 600123456")
        if hasattr(self, "email_entry"):
            self._set_placeholder(self.email_entry, self.email_var, "Ej: ana@correo.com")

    def _clear_placeholders_for_form(self) -> None:
        for entry, var in [(self.name_entry, self.name_var), (self.phone_entry, self.phone_var), (self.email_entry, self.email_var)]:
            placeholder = self._placeholders.get(entry)
            if placeholder and var.get() == placeholder:
                var.set("")
                self._set_entry_color(entry, "#111827")

    def _build_tree_menu(self) -> None:
        self._tree_menu = tk.Menu(self.root, tearoff=0)
        self._tree_menu.add_command(label="Editar", command=self._menu_edit)
        self._tree_menu.add_command(label="Eliminar", command=self.on_delete)
        self._tree_menu.add_separator()
        self._tree_menu.add_command(label="Copiar teléfono", command=self._menu_copy_phone)
        self._tree_menu.add_command(label="Copiar email", command=self._menu_copy_email)
        self.tree.bind("<Button-3>", self._on_tree_right_click, add=True)

    def _on_tree_right_click(self, evt: tk.Event) -> None:
        item = self.tree.identify_row(evt.y)
        if item:
            self.tree.selection_set(item)
            self.tree.focus(item)
            self.on_tree_select()
            try:
                self._tree_menu.tk_popup(evt.x_root, evt.y_root)
            finally:
                self._tree_menu.grab_release()

    def _menu_edit(self) -> None:
        self.name_entry.focus_set()

    def _menu_copy_phone(self) -> None:
        phone = self.phone_var.get()
        self._copy_to_clipboard(phone)

    def _menu_copy_email(self) -> None:
        email = self.email_var.get()
        self._copy_to_clipboard(email)

    def _copy_to_clipboard(self, text: str) -> None:
        if not text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    def _add_tooltips(self) -> None:
        ToolTip(self.add_btn, "Añade el contacto (Enter)")
        ToolTip(self.update_btn, "Actualiza el contacto seleccionado (Enter)")
        ToolTip(self.delete_btn, "Elimina el contacto seleccionado (Supr)")
        ToolTip(self.clear_btn, "Limpia el formulario (Esc)")
        ToolTip(self.dup_btn, "Detecta y fusiona contactos duplicados")
        ToolTip(self.tree, "Click derecho para más opciones")


def _first_key(row: dict, keys: list[str]) -> str:
    for k in keys:
        if k in row and row[k] is not None:
            return str(row[k])
    return ""


class ToolTip:
    BG = "#111827"
    FG = "#F9FAFB"

    def __init__(self, widget: tk.Widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self._tip: Optional[tk.Toplevel] = None
        self._after_id: Optional[str] = None
        widget.bind("<Enter>", self._schedule, add=True)
        widget.bind("<Leave>", self._hide, add=True)
        widget.bind("<ButtonPress>", self._hide, add=True)

    def _schedule(self, _evt: object) -> None:
        self._cancel()
        self._after_id = self.widget.after(500, self._show)

    def _cancel(self) -> None:
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except tk.TclError:
                pass
            self._after_id = None

    def _show(self) -> None:
        if self._tip is not None:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except tk.TclError:
            return
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        frame = tk.Frame(self._tip, bg=self.BG, bd=0)
        frame.pack()
        label = tk.Label(frame, text=self.text, bg=self.BG, fg=self.FG, font=("Segoe UI", 9))
        label.pack(padx=10, pady=6)

    def _hide(self, _evt: Optional[object] = None) -> None:
        self._cancel()
        if self._tip is not None:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None


class DuplicateWindow:
    def __init__(
        self,
        parent: tk.Tk,
        conn: sqlite3.Connection,
        groups: list[dict],
        on_merged: Callable[[], None],
    ) -> None:
        self.conn = conn
        self.groups = groups
        self.on_merged = on_merged
        self.contacts_by_id: dict[int, Contact] = {
            int(c.id): c for c in db.list_contacts(conn) if c.id is not None
        }

        self.win = tk.Toplevel(parent)
        self.win.title("Duplicados detectados")
        self.win.minsize(820, 480)

        container = ttk.Frame(self.win, padding=10)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            container,
            text="Selecciona un grupo y pulsa “Fusionar” para unificarlo en un único contacto.",
        ).pack(anchor="w")

        upper = ttk.Frame(container)
        upper.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.group_tree = ttk.Treeview(upper, columns=("group", "score", "ids"), show="headings", height=6)
        self.group_tree.heading("group", text="Grupo")
        self.group_tree.heading("score", text="Score")
        self.group_tree.heading("ids", text="IDs")
        self.group_tree.column("group", width=80, anchor=tk.CENTER, stretch=False)
        self.group_tree.column("score", width=80, anchor=tk.CENTER, stretch=False)
        self.group_tree.column("ids", width=580, anchor=tk.W)

        group_scroll = ttk.Scrollbar(upper, orient=tk.VERTICAL, command=self.group_tree.yview)
        self.group_tree.configure(yscrollcommand=group_scroll.set)
        self.group_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        group_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        lower = ttk.Frame(container)
        lower.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.detail_tree = ttk.Treeview(lower, columns=("id", "name", "phone", "email"), show="headings")
        self.detail_tree.heading("id", text="ID")
        self.detail_tree.heading("name", text="Nombre")
        self.detail_tree.heading("phone", text="Teléfono")
        self.detail_tree.heading("email", text="Email")
        self.detail_tree.column("id", width=60, anchor=tk.CENTER, stretch=False)
        self.detail_tree.column("name", width=260, anchor=tk.W)
        self.detail_tree.column("phone", width=140, anchor=tk.W)
        self.detail_tree.column("email", width=260, anchor=tk.W)

        detail_scroll = ttk.Scrollbar(lower, orient=tk.VERTICAL, command=self.detail_tree.yview)
        self.detail_tree.configure(yscrollcommand=detail_scroll.set)
        self.detail_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        detail_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        actions = ttk.Frame(container)
        actions.pack(fill=tk.X, pady=(10, 0))
        self.merge_btn = ttk.Button(actions, text="Fusionar grupo seleccionado", command=self.on_merge)
        self.merge_btn.pack(side=tk.RIGHT)

        self._render_groups()
        self.group_tree.bind("<<TreeviewSelect>>", self.on_group_select)
        self.win.grab_set()

    def _render_groups(self) -> None:
        self.group_tree.delete(*self.group_tree.get_children())
        for idx, g in enumerate(self.groups, start=1):
            ids = g["ids"]
            self.group_tree.insert("", tk.END, values=(idx, g["score"], ", ".join(map(str, ids))))

        first = self.group_tree.get_children()
        if first:
            self.group_tree.selection_set(first[0])
            self.on_group_select()

    def on_group_select(self, _evt: Optional[object] = None) -> None:
        sel = self.group_tree.selection()
        if not sel:
            return
        values = self.group_tree.item(sel[0], "values")
        if not values:
            return
        group_index = int(values[0]) - 1
        ids = self.groups[group_index]["ids"]
        self._render_details(ids)

    def _render_details(self, ids: list[int]) -> None:
        self.detail_tree.delete(*self.detail_tree.get_children())
        for cid in ids:
            c = self.contacts_by_id.get(cid)
            if c is None:
                continue
            self.detail_tree.insert("", tk.END, values=(c.id, c.name, c.phone, c.email))

    def on_merge(self) -> None:
        sel = self.group_tree.selection()
        if not sel:
            return
        values = self.group_tree.item(sel[0], "values")
        if not values:
            return
        group_index = int(values[0]) - 1
        ids = self.groups[group_index]["ids"]

        contacts = []
        for cid in ids:
            c = db.get_contact(self.conn, cid)
            if c is not None:
                contacts.append(c)
        if len(contacts) < 2:
            messagebox.showinfo("Fusionar", "El grupo ya no está disponible.", parent=self.win)
            return

        if not messagebox.askyesno(
            "Fusionar",
            f"Se fusionarán {len(contacts)} contactos en uno solo.\n\n"
            f"Se mantendrá el ID {contacts[0].id} y se eliminarán los demás.\n\n¿Continuar?",
            parent=self.win,
        ):
            return

        merged = merge_contacts(contacts)
        try:
            db.update_contact(self.conn, int(merged.id), merged)
            for c in contacts[1:]:
                db.delete_contact(self.conn, int(c.id))
        except sqlite3.Error as e:
            messagebox.showerror("Base de datos", f"No se pudo fusionar.\n\n{e}", parent=self.win)
            return

        self.contacts_by_id = {int(c.id): c for c in db.list_contacts(self.conn) if c.id is not None}
        self.groups.pop(group_index)
        if not self.groups:
            self.on_merged()
            self.win.destroy()
            return
        self._render_groups()
        self.on_merged()
