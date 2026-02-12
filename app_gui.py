import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
import subprocess

from engine.reference_cache import build_reference_cache, cache_exists
from engine.runner import run_exclusion_check
from engine.config_loader import ClientConfig


PROJECT_ROOT = Path(__file__).resolve().parent
CLIENTS_DIR = PROJECT_ROOT / "clients"


class ExclusionAppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ExclusionApp")
        self.root.geometry("950x650")

        self.client_var = tk.StringVar()
        self.month_var = tk.StringVar(value=datetime.now().strftime("%Y-%m"))

        self.oig_path = tk.StringVar()
        self.sam_path = tk.StringVar()
        self.staff_path = tk.StringVar()
        self.board_path = tk.StringVar()
        self.vendor_path = tk.StringVar()

        self._build_ui()
        self._load_clients()

    # ---------------- UI ----------------

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True)

        # Run Setup
        setup = ttk.LabelFrame(main, text="Run Setup")
        setup.pack(fill="x", **pad)

        ttk.Label(setup, text="Client:").grid(row=0, column=0, sticky="w", **pad)
        self.client_combo = ttk.Combobox(setup, textvariable=self.client_var, state="readonly", width=40)
        self.client_combo.grid(row=0, column=1, sticky="w", **pad)
        self.client_combo.bind("<<ComboboxSelected>>", lambda e: self._update_button_states())

        ttk.Label(setup, text="Month (YYYY-MM):").grid(row=0, column=2, sticky="w", **pad)
        self.month_entry = ttk.Entry(setup, textvariable=self.month_var, width=12)
        self.month_entry.grid(row=0, column=3, sticky="w", **pad)
        self.month_entry.bind("<KeyRelease>", lambda e: self._update_button_states())

        # Input Files
        files = ttk.LabelFrame(main, text="Input Files")
        files.pack(fill="x", **pad)

        self._file_row(files, 0, "OIG CSV:", self.oig_path, [("CSV files", "*.csv")])
        self._file_row(files, 1, "SAM CSV:", self.sam_path, [("CSV files", "*.csv")])
        self._file_row(files, 2, "Staff CSV:", self.staff_path, [("CSV files", "*.csv")])
        self._file_row(files, 3, "Board CSV:", self.board_path, [("CSV files", "*.csv")])
        self._file_row(files, 4, "Vendors XLSX:", self.vendor_path, [("Excel files", "*.xlsx")])

        # Buttons
        actions = ttk.Frame(main)
        actions.pack(fill="x", **pad)

        self.btn_build = ttk.Button(actions, text="Build Reference Cache (Monthly)", command=self.on_build)
        self.btn_build.pack(side="left", padx=10)

        self.btn_run = ttk.Button(actions, text="Run Exclusion Check", command=self.on_run)
        self.btn_run.pack(side="left", padx=10)

        self.btn_open = ttk.Button(actions, text="Open Runs Folder", command=self.open_runs)
        self.btn_open.pack(side="left", padx=10)

        # Log
        log_frame = ttk.LabelFrame(main, text="Status Log")
        log_frame.pack(fill="both", expand=True, **pad)

        self.log_box = tk.Text(log_frame, height=15)
        self.log_box.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_box.configure(state="disabled")


    def _file_row(self, parent, row, label, var, filetypes):
        pad = {"padx": 8, "pady": 6}

        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", **pad)

        entry = ttk.Entry(parent, textvariable=var, width=70, state="readonly")
        entry.grid(row=row, column=1, sticky="w", **pad)

        btn = ttk.Button(parent, text="Select", command=lambda: self._select_file(var, filetypes))
        btn.grid(row=row, column=2, sticky="w", **pad)

    def _select_file(self, var, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)
        self._update_button_states()

    def _log(self, message):
        self.log_box.configure(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{ts}] {message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # ---------------- Client Load ----------------

    def _load_clients(self):
        self.client_map = {}
        names = []

        for yaml_file in CLIENTS_DIR.glob("*.yaml"):
            try:
                cfg = ClientConfig(yaml_file)
                name = cfg.client_name
                self.client_map[name] = str(yaml_file)
                names.append(name)
            except Exception:
                pass

        self.client_combo["values"] = names
        if names:
            self.client_combo.current(0)

        self._update_button_states()

    # ---------------- Validation ----------------

    def _month_valid(self):
        try:
            datetime.strptime(self.month_var.get(), "%Y-%m")
            return True
        except Exception:
            return False

    def _file_exists(self, path):
        return path and Path(path).exists()

    def _update_button_states(self):
        month_ok = self._month_valid()
        client_ok = self.client_var.get() in self.client_map

        oig_ok = self._file_exists(self.oig_path.get())
        sam_ok = self._file_exists(self.sam_path.get())

        cache_ready = month_ok and cache_exists(self.month_var.get())

        self.btn_build.config(state="normal" if (client_ok and month_ok and oig_ok and sam_ok and not cache_ready) else "disabled")

        any_sources = any([
            self._file_exists(self.staff_path.get()),
            self._file_exists(self.board_path.get()),
            self._file_exists(self.vendor_path.get())
        ])

        self.btn_run.config(state="normal" if (client_ok and month_ok and cache_ready and any_sources) else "disabled")

    # ---------------- Actions ----------------

    def on_build(self):
        month = self.month_var.get()
        self._log("Building reference cache...")
        build_reference_cache(month, self.oig_path.get(), self.sam_path.get())
        self._log("Reference cache complete.")
        self._update_button_states()

    def on_run(self):
        month = self.month_var.get()
        client_yaml = self.client_map[self.client_var.get()]

        self._log("Running exclusion check...")

        result = run_exclusion_check(
            client_yaml,
            month,
            staff_path=self.staff_path.get() or None,
            board_path=self.board_path.get() or None,
            vendor_path=self.vendor_path.get() or None,
            oig_path=self.oig_path.get(),
            sam_path=self.sam_path.get()
        )

        self._log("Run complete.")
        self._log(f"Outputs saved in: {result['run_directory']}")

    def open_runs(self):
        runs_dir = Path.home() / "ExclusionAppData" / "runs"
        runs_dir.mkdir(parents=True, exist_ok=True)
        subprocess.run(["open", str(runs_dir)])


if __name__ == "__main__":
    root = tk.Tk()
    app = ExclusionAppGUI(root)
    root.mainloop()