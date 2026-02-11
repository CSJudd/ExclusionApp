import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime

from engine.reference_cache import build_reference_cache, cache_exists
from engine.runner import run_exclusion_check
from engine.config_loader import ClientConfig


PROJECT_ROOT = Path(__file__).resolve().parent
CLIENTS_DIR = PROJECT_ROOT / "clients"


class ExclusionAppGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("ExclusionApp")
        self.root.geometry("900x620")

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
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self.root)
        frm.pack(fill="both", expand=True)

        # Top controls
        top = ttk.LabelFrame(frm, text="Run Setup")
        top.pack(fill="x", **pad)

        ttk.Label(top, text="Client:").grid(row=0, column=0, sticky="w", **pad)
        self.client_combo = ttk.Combobox(top, textvariable=self.client_var, state="readonly", width=45)
        self.client_combo.grid(row=0, column=1, sticky="w", **pad)
        self.client_combo.bind("<<ComboboxSelected>>", lambda e: self._update_button_states())

        ttk.Label(top, text="Month (YYYY-MM):").grid(row=0, column=2, sticky="w", **pad)
        self.month_entry = ttk.Entry(top, textvariable=self.month_var, width=15)
        self.month_entry.grid(row=0, column=3, sticky="w", **pad)
        self.month_entry.bind("<KeyRelease>", lambda e: self._update_button_states())

        # File selectors
        files = ttk.LabelFrame(frm, text="Input Files")
        files.pack(fill="x", **pad)

        self._file_row(files, 0, "OIG CSV:", self.oig_path, filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        self._file_row(files, 1, "SAM CSV:", self.sam_path, filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        self._file_row(files, 2, "Staff CSV:", self.staff_path, filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        self._file_row(files, 3, "Board CSV:", self.board_path, filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        self._file_row(files, 4, "Vendors XLSX:", self.vendor_path, filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

        # Action buttons
        actions = ttk.Frame(frm)
        actions.pack(fill="x", **pad)

        self.btn_build_cache = ttk.Button(actions, text="Build Reference Cache (Monthly)", command=self.on_build_cache)
        self.btn_build_cache.pack(side="left", padx=10)

        self.btn_run = ttk.Button(actions, text="Run Exclusion Check", command=self.on_run)
        self.btn_run.pack(side="left", padx=10)

        self.btn_open_runs = ttk.Button(actions, text="Open Runs Folder", command=self.on_open_runs_folder)
        self.btn_open_runs.pack(side="left", padx=10)

        # Log window
        logbox = ttk.LabelFrame(frm, text="Status Log")
        logbox.pack(fill="both", expand=True, **pad)

        self.log_text = tk.Text(logbox, height=18, wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_text.configure(state="disabled")

        self._update_button_states()

    def _file_row(self, parent, row, label, var, filetypes):
        pad = {"padx": 10, "pady": 6}

        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", **pad)
        entry = ttk.Entry(parent, textvariable=var, width=80)
        entry.grid(row=row, column=1, sticky="w", **pad)

        btn = ttk.Button(parent, text="Select", command=lambda: self._pick_file(var, filetypes))
        btn.grid(row=row, column=2, sticky="w", **pad)

    def _pick_file(self, var, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)
            self._update_button_states()

    def _log(self, msg: str):
        self.log_text.configure(state="normal")
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {msg}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # ---------------- Clients ----------------
    def _load_clients(self):
        if not CLIENTS_DIR.exists():
            messagebox.showerror("Missing clients folder", f"Could not find: {CLIENTS_DIR}")
            return

        yamls = sorted(CLIENTS_DIR.glob("*.yaml"))
        client_names = []
        self.client_map = {}

        for y in yamls:
            try:
                cfg = ClientConfig(y)
                name = cfg.client_name or y.stem
                client_names.append(name)
                self.client_map[name] = str(y)
            except Exception as e:
                self._log(f"Failed to load client config {y.name}: {e}")

        self.client_combo["values"] = client_names
        if client_names:
            self.client_combo.current(0)
            self.client_var.set(client_names[0])

        self._update_button_states()

    # ---------------- Validation ----------------
    def _month_ok(self):
        val = self.month_var.get().strip()
        try:
            datetime.strptime(val, "%Y-%m")
            return True
        except Exception:
            return False

    def _client_ok(self):
        return bool(self.client_var.get().strip()) and self.client_var.get().strip() in self.client_map

    def _update_button_states(self):
        # Build cache requires: client + valid month + oig + sam
        can_build = (
            self._client_ok()
            and self._month_ok()
            and self.oig_path.get().strip()
            and self.sam_path.get().strip()
        )

        # Run requires: client + valid month + cache exists + staff/board/vendor selected (at least one)
        month = self.month_var.get().strip()
        any_sources = any([
            self.staff_path.get().strip(),
            self.board_path.get().strip(),
            self.vendor_path.get().strip()
        ])

        can_run = (
            self._client_ok()
            and self._month_ok()
            and cache_exists(month)
            and any_sources
        )

        # If cache already exists, disable build cache (prevents accidental rebuild)
        if self._month_ok() and cache_exists(month):
            self.btn_build_cache.configure(state="disabled")
        else:
            self.btn_build_cache.configure(state=("normal" if can_build else "disabled"))

        self.btn_run.configure(state=("normal" if can_run else "disabled"))

    # ---------------- Actions ----------------
    def on_build_cache(self):
        if not self._client_ok():
            messagebox.showerror("Missing client", "Select a client first.")
            return
        if not self._month_ok():
            messagebox.showerror("Invalid month", "Month must be YYYY-MM.")
            return

        month = self.month_var.get().strip()
        oig = self.oig_path.get().strip()
        sam = self.sam_path.get().strip()

        if not oig or not sam:
            messagebox.showerror("Missing files", "Select both OIG CSV and SAM CSV.")
            return

        if cache_exists(month):
            messagebox.showinfo("Cache exists", f"Reference cache already exists for {month}.")
            self._update_button_states()
            return

        try:
            self._log(f"Building reference cache for {month}...")
            self._log(f"OIG: {oig}")
            self._log(f"SAM: {sam}")
            build_reference_cache(month, oig, sam)
            self._log(f"Reference cache built for {month}.")
        except Exception as e:
            self._log(f"ERROR building cache: {e}")
            messagebox.showerror("Cache build failed", str(e))

        self._update_button_states()

    def on_run(self):
        if not self._client_ok():
            messagebox.showerror("Missing client", "Select a client first.")
            return
        if not self._month_ok():
            messagebox.showerror("Invalid month", "Month must be YYYY-MM.")
            return

        month = self.month_var.get().strip()
        if not cache_exists(month):
            messagebox.showerror("Missing cache", f"Reference cache not found for {month}. Build it first.")
            return

        client_name = self.client_var.get().strip()
        client_yaml = self.client_map[client_name]

        staff = self.staff_path.get().strip() or None
        board = self.board_path.get().strip() or None
        vendor = self.vendor_path.get().strip() or None

        if not any([staff, board, vendor]):
            messagebox.showerror("No source files", "Select at least one of Staff, Board, or Vendors.")
            return

        try:
            self._log(f"Running exclusion check: client={client_name}, month={month}")
            out = run_exclusion_check(
                client_yaml,
                month,
                staff_path=staff,
                board_path=board,
                vendor_path=vendor,
                oig_path=self.oig_path.get().strip() or None,
                sam_path=self.sam_path.get().strip() or None
            )
            self._log("Run complete.")
            self._log(f"Run directory: {out.get('run_directory')}")
            self._log(f"Audit file: {out.get('audit_file')}")
            if out.get("staff_pdf"):
                self._log(f"Staff PDF: {out.get('staff_pdf')}")
            if out.get("board_pdf"):
                self._log(f"Board PDF: {out.get('board_pdf')}")
            if out.get("vendor_pdf"):
                self._log(f"Vendor PDF: {out.get('vendor_pdf')}")

            messagebox.showinfo("Success", f"Run complete.\n\nOutputs saved in:\n{out.get('run_directory')}")
        except Exception as e:
            self._log(f"ERROR running exclusion check: {e}")
            messagebox.showerror("Run failed", str(e))

        self._update_button_states()

    def on_open_runs_folder(self):
        # Opens the base runs folder in Finder (macOS)
        runs_dir = Path.home() / "ExclusionAppData" / "runs"
        runs_dir.mkdir(parents=True, exist_ok=True)
        try:
            import subprocess
            subprocess.run(["open", str(runs_dir)], check=False)
        except Exception as e:
            messagebox.showerror("Could not open folder", str(e))


def main():
    root = tk.Tk()
    # Use ttk theme
    style = ttk.Style()
    try:
        style.theme_use("aqua")
    except Exception:
        pass

    app = ExclusionAppGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()