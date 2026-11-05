from __future__ import annotations

import json
import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from llamamientos import AnalysisResult, LlamamientosError, analyze_files, import_month


APP_TITLE = "Importación mensual de llamamientos"


class LlamamientosApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1080x760")
        self.minsize(900, 650)

        self.source_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.summary_var = tk.StringVar(value="Selecciona los dos archivos y pulsa Analizar.")
        self.status_var = tk.StringVar(value="Preparado")
        self.analysis: AnalysisResult | None = None
        self._busy = False

        self._configure_style()
        self._build_ui()
        self._load_config()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#17324D")
        style.configure("Section.TLabel", font=("Segoe UI", 10, "bold"), foreground="#17324D")
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(18, 8))
        style.configure("TButton", padding=(10, 6))
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=22)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(5, weight=1)

        ttk.Label(outer, text=APP_TITLE, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            outer,
            text=(
                "Carga de una vez todos los días del mes cerrado. El origen debe haberse abierto, "
                "recalculado y guardado después de aplicar las correcciones. Antes de importar, "
                "cierra todas las ventanas de Microsoft Excel."
            ),
            wraplength=1000,
        ).grid(row=1, column=0, sticky="ew", pady=(5, 18))

        files = ttk.LabelFrame(outer, text="Archivos", padding=14)
        files.grid(row=2, column=0, sticky="ew")
        files.columnconfigure(1, weight=1)

        ttk.Label(files, text="Excel mensual de origen", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 12), pady=5
        )
        ttk.Entry(files, textvariable=self.source_var).grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Button(files, text="Examinar…", command=self._choose_source).grid(
            row=0, column=2, padx=(10, 0), pady=5
        )

        ttk.Label(files, text="Herramienta de destino", style="Section.TLabel").grid(
            row=1, column=0, sticky="w", padx=(0, 12), pady=5
        )
        ttk.Entry(files, textvariable=self.target_var).grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Button(files, text="Examinar…", command=self._choose_target).grid(
            row=1, column=2, padx=(10, 0), pady=5
        )

        actions = ttk.Frame(outer)
        actions.grid(row=3, column=0, sticky="ew", pady=14)
        self.analyze_button = ttk.Button(
            actions, text="Analizar archivos", style="Primary.TButton", command=self._analyze
        )
        self.analyze_button.pack(side="left")
        self.import_button = ttk.Button(
            actions,
            text="Importar y reemplazar el mes",
            style="Primary.TButton",
            command=self._import,
            state="disabled",
        )
        self.import_button.pack(side="left", padx=(10, 0))
        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=180)
        self.progress.pack(side="right")

        summary = ttk.LabelFrame(outer, text="Resultado del análisis", padding=12)
        summary.grid(row=4, column=0, sticky="ew", pady=(0, 12))
        summary.columnconfigure(0, weight=1)
        ttk.Label(summary, textvariable=self.summary_var, wraplength=1000).grid(
            row=0, column=0, sticky="ew"
        )

        details = ttk.Panedwindow(outer, orient="horizontal")
        details.grid(row=5, column=0, sticky="nsew")

        days_frame = ttk.LabelFrame(details, text="Registros por día", padding=8)
        notes_frame = ttk.LabelFrame(details, text="Avisos y errores", padding=8)
        details.add(days_frame, weight=3)
        details.add(notes_frame, weight=2)

        days_frame.rowconfigure(0, weight=1)
        days_frame.columnconfigure(0, weight=1)
        self.days_tree = ttk.Treeview(
            days_frame,
            columns=("date", "sheet", "records", "state"),
            show="headings",
            selectmode="browse",
        )
        self.days_tree.heading("date", text="Fecha")
        self.days_tree.heading("sheet", text="Hoja")
        self.days_tree.heading("records", text="Registros")
        self.days_tree.heading("state", text="Estado hoja")
        self.days_tree.column("date", width=100, anchor="center")
        self.days_tree.column("sheet", width=95, anchor="center")
        self.days_tree.column("records", width=90, anchor="e")
        self.days_tree.column("state", width=100, anchor="center")
        scrollbar = ttk.Scrollbar(days_frame, orient="vertical", command=self.days_tree.yview)
        self.days_tree.configure(yscrollcommand=scrollbar.set)
        self.days_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        notes_frame.rowconfigure(0, weight=1)
        notes_frame.columnconfigure(0, weight=1)
        self.notes = tk.Text(
            notes_frame,
            wrap="word",
            font=("Segoe UI", 9),
            relief="flat",
            padx=8,
            pady=8,
            state="disabled",
            background="#FAFAFA",
        )
        notes_scroll = ttk.Scrollbar(notes_frame, orient="vertical", command=self.notes.yview)
        self.notes.configure(yscrollcommand=notes_scroll.set)
        self.notes.grid(row=0, column=0, sticky="nsew")
        notes_scroll.grid(row=0, column=1, sticky="ns")

        ttk.Separator(outer).grid(row=6, column=0, sticky="ew", pady=(14, 8))
        ttk.Label(outer, textvariable=self.status_var).grid(row=7, column=0, sticky="w")

    @property
    def _config_path(self) -> Path:
        base = Path(os.environ.get("APPDATA", Path.home()))
        return base / "LlamamientosMensuales" / "config.json"

    def _load_config(self) -> None:
        try:
            data = json.loads(self._config_path.read_text(encoding="utf-8"))
            self.source_var.set(data.get("source", ""))
            self.target_var.set(data.get("target", ""))
        except (OSError, ValueError, TypeError):
            pass

    def _save_config(self) -> None:
        try:
            self._config_path.parent.mkdir(parents=True, exist_ok=True)
            self._config_path.write_text(
                json.dumps(
                    {"source": self.source_var.get().strip(), "target": self.target_var.get().strip()},
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
        except OSError:
            pass

    def _initial_dir(self, value: str) -> str | None:
        path = Path(value.strip())
        if path.is_file():
            return str(path.parent)
        if path.is_dir():
            return str(path)
        return None

    def _choose_source(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Seleccionar Excel mensual de llamamientos",
            initialdir=self._initial_dir(self.source_var.get()),
            filetypes=(("Excel", "*.xlsm *.xlsx"), ("Todos los archivos", "*.*")),
        )
        if chosen:
            self.source_var.set(chosen)
            self._invalidate_analysis()
            self._save_config()

    def _choose_target(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Seleccionar Herramienta Llamamientos",
            initialdir=self._initial_dir(self.target_var.get()),
            filetypes=(("Excel", "*.xlsx"), ("Todos los archivos", "*.*")),
        )
        if chosen:
            self.target_var.set(chosen)
            self._invalidate_analysis()
            self._save_config()

    def _invalidate_analysis(self) -> None:
        self.analysis = None
        self.import_button.configure(state="disabled")
        self.summary_var.set("Los archivos han cambiado. Pulsa Analizar.")

    def _set_busy(self, busy: bool, status: str) -> None:
        self._busy = busy
        self.status_var.set(status)
        self.analyze_button.configure(state="disabled" if busy else "normal")
        if busy:
            self.import_button.configure(state="disabled")
            self.progress.start(12)
        else:
            self.progress.stop()

    def _run_background(self, work, success, status: str) -> None:
        if self._busy:
            return
        self._set_busy(True, status)

        def runner() -> None:
            try:
                result = work()
            except Exception as exc:
                self.after(0, lambda error=exc: self._background_error(error))
            else:
                self.after(0, lambda: success(result))

        threading.Thread(target=runner, daemon=True).start()

    def _background_error(self, exc: Exception) -> None:
        self._set_busy(False, "Se ha producido un error")
        message = str(exc) if isinstance(exc, LlamamientosError) else f"Error inesperado:\n{exc}"
        self._write_notes([message])
        messagebox.showerror(APP_TITLE, message, parent=self)

    def _analyze(self) -> None:
        source = self.source_var.get().strip()
        target = self.target_var.get().strip()
        self._save_config()
        self._run_background(
            lambda: analyze_files(source, target), self._analysis_complete, "Analizando los libros…"
        )

    def _analysis_complete(self, analysis: AnalysisResult) -> None:
        self.analysis = analysis
        self._set_busy(False, "Análisis completado")
        for item in self.days_tree.get_children():
            self.days_tree.delete(item)
        for day in analysis.days:
            self.days_tree.insert(
                "",
                "end",
                values=(
                    day.work_date.strftime("%d/%m/%Y"),
                    day.sheet_name,
                    day.record_count,
                    "Oculta" if day.sheet_state != "visible" else "Visible",
                ),
            )

        self.summary_var.set(
            f"{analysis.period_label}: {_number(len(analysis.records))} registros válidos. "
            f"La hoja {analysis.target_sheet} contiene actualmente "
            f"{_number(analysis.existing_record_count)} registros y su tabla tiene capacidad para "
            f"{_number(analysis.target_capacity)}."
        )
        notes = []
        notes.extend(f"ERROR: {item}" for item in analysis.errors)
        notes.extend(f"AVISO: {item}" for item in analysis.warnings)
        if not notes:
            notes = ["Sin errores ni avisos. La importación está lista."]
        self._write_notes(notes)
        self.import_button.configure(state="normal" if analysis.can_import else "disabled")

    def _write_notes(self, lines: list[str]) -> None:
        self.notes.configure(state="normal")
        self.notes.delete("1.0", "end")
        self.notes.insert("1.0", "\n\n".join(lines))
        self.notes.configure(state="disabled")

    def _import(self) -> None:
        if not self.analysis or not self.analysis.can_import:
            return
        analysis = self.analysis
        prompt = (
            "Se modificará y guardará directamente el archivo original seleccionado.\n\n"
            f"Se reemplazarán todos los datos A:E de {analysis.target_sheet} por "
            f"{_number(len(analysis.records))} registros.\n\n"
            "Justo antes se creará una copia de seguridad fechada. Las fórmulas F:K se conservarán "
            "y extenderán.\n\n¿Continuar?"
        )
        if not messagebox.askyesno(APP_TITLE, prompt, icon="warning", parent=self):
            return

        self._run_background(
            lambda: import_month(self.source_var.get().strip(), self.target_var.get().strip()),
            self._import_complete,
            "Importando el mes y validando el resultado…",
        )

    def _import_complete(self, result) -> None:
        self._set_busy(False, "Importación completada")
        self.analysis = None
        self.import_button.configure(state="disabled")
        message = (
            f"Importación completada.\n\n"
            f"Hoja: {result.target_sheet}\n"
            f"Registros: {_number(result.imported_records)}\n\n"
            f"Copia de seguridad:\n{result.backup_path}"
        )
        self._write_notes([message])
        self.summary_var.set(
            "El mes se ha reemplazado y validado. Pulsa Analizar para comprobar de nuevo los archivos."
        )
        messagebox.showinfo(APP_TITLE, message, parent=self)


def _number(value: int) -> str:
    return f"{value:,}".replace(",", ".")


if __name__ == "__main__":
    LlamamientosApp().mainloop()
