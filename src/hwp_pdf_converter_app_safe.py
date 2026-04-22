import os
import csv
import shutil
import threading
import queue
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "HWP/HWPX → PDF Converter (Safe)"
DEFAULT_OPEN_OPTION = "forceopen:true;versionwarning:false;"
TEMP_WORKDIR = Path(r"C:\temp\hwp_convert_safe")


def ensure_pywin32():
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
        return True, ""
    except Exception as e:
        return False, str(e)


def is_hwp_running():
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Hwp.exe"],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return "Hwp.exe" in result.stdout
    except Exception:
        return False


def kill_hwp():
    try:
        subprocess.run(
            ["taskkill", "/IM", "Hwp.exe", "/F"],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return True
    except Exception:
        return False


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("920x690")
        self.root.minsize(820, 560)

        self.folder_var = tk.StringVar()
        self.overwrite_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=True)
        self.use_safe_copy_var = tk.BooleanVar(value=True)

        self.log_queue = queue.Queue()
        self.worker = None
        self.stop_requested = False
        self.is_running = False

        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="x")

        ttk.Label(top, text="Root folder").grid(row=0, column=0, sticky="w")
        top.columnconfigure(0, weight=1)

        path_row = ttk.Frame(top)
        path_row.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        path_row.columnconfigure(0, weight=1)

        ttk.Entry(path_row, textvariable=self.folder_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(path_row, text="Browse...", command=self.browse_folder).grid(row=0, column=1, sticky="e")

        opts = ttk.LabelFrame(self.root, text="Options", padding=12)
        opts.pack(fill="x", padx=12, pady=(0, 12))

        ttk.Checkbutton(opts, text="Include subfolders", variable=self.recursive_var).grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Checkbutton(opts, text="Overwrite existing PDF", variable=self.overwrite_var).grid(row=0, column=1, sticky="w", padx=(0, 16))
        ttk.Checkbutton(
            opts,
            text="Use safe local temp conversion (recommended for Google Drive / network drives)",
            variable=self.use_safe_copy_var
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))

        actions = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        actions.pack(fill="x")

        self.start_btn = ttk.Button(actions, text="Start conversion", command=self.start_conversion)
        self.start_btn.pack(side="left")

        self.stop_btn = ttk.Button(actions, text="Stop", command=self.request_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))

        ttk.Button(actions, text="Open selected folder", command=self.open_selected_folder).pack(side="left", padx=(8, 0))

        progress_frame = ttk.Frame(self.root, padding=(12, 0, 12, 0))
        progress_frame.pack(fill="x")

        self.progress_label_var = tk.StringVar(value="Ready")
        ttk.Label(progress_frame, textvariable=self.progress_label_var).pack(anchor="w")

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(6, 0))

        log_frame = ttk.Frame(self.root, padding=12)
        log_frame.pack(fill="both", expand=True)

        ttk.Label(log_frame, text="Log").pack(anchor="w")
        self.log_text = tk.Text(log_frame, wrap="word")
        self.log_text.pack(side="left", fill="both", expand=True)

        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        note_frame = ttk.LabelFrame(self.root, text="Notes", padding=12)
        note_frame.pack(fill="x", padx=12, pady=(0, 12))
        notes = (
            "• Close Hancom HWP before starting for best stability.\n"
            "• Safe temp mode copies each file to a short local path before conversion.\n"
            "• A CSV log will be written to the selected root folder."
        )
        ttk.Label(note_frame, text=notes, justify="left").pack(anchor="w")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)

    def open_selected_folder(self):
        folder = self.folder_var.get().strip()
        if folder and os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning(APP_TITLE, "Select a valid folder first.")

    def append_log(self, text: str):
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")

    def request_stop(self):
        self.stop_requested = True
        self.append_log("Stop requested. Current file will finish first.")

    def _poll_log_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()

                if kind == "log":
                    self.append_log(payload)

                elif kind == "progress":
                    current, total, label = payload
                    self.progress["maximum"] = max(total, 1)
                    self.progress["value"] = current
                    self.progress_label_var.set(label)

                elif kind == "done":
                    success, failed, skipped, log_csv = payload
                    self.is_running = False
                    self.start_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    self.progress_label_var.set(f"Done. Success: {success}, Failed: {failed}, Skipped: {skipped}")
                    messagebox.showinfo(
                        APP_TITLE,
                        f"Conversion finished.\n\nSuccess: {success}\nFailed: {failed}\nSkipped: {skipped}\n\nLog file:\n{log_csv}"
                    )

                elif kind == "error":
                    self.is_running = False
                    self.start_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    self.progress_label_var.set("Error")
                    messagebox.showerror(APP_TITLE, payload)

        except queue.Empty:
            pass

        self.root.after(150, self._poll_log_queue)

    def start_conversion(self):
        if self.is_running:
            messagebox.showwarning(APP_TITLE, "A conversion job is already running.")
            return

        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror(APP_TITLE, "Select a valid root folder.")
            return

        ok, detail = ensure_pywin32()
        if not ok:
            messagebox.showerror(
                APP_TITLE,
                "pywin32 is not available.\n\nInstall it with:\n"
                "python -m pip install pywin32\n\nDetails:\n" + detail
            )
            return

        if is_hwp_running():
            answer = messagebox.askyesnocancel(
                APP_TITLE,
                "Hancom HWP is already running.\n\n"
                "Yes: force close HWP and continue\n"
                "No: continue anyway\n"
                "Cancel: stop"
            )
            if answer is None:
                return
            if answer is True:
                killed = kill_hwp()
                if not killed:
                    messagebox.showerror(APP_TITLE, "Could not close HWP automatically.")
                    return

        self.stop_requested = False
        self.is_running = True
        self.log_text.delete("1.0", "end")
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.progress["value"] = 0
        self.progress_label_var.set("Scanning files...")

        self.worker = threading.Thread(
            target=self._run_conversion,
            args=(
                folder,
                self.recursive_var.get(),
                self.overwrite_var.get(),
                self.use_safe_copy_var.get(),
            ),
            daemon=True,
        )
        self.worker.start()

    @staticmethod
    def collect_files(folder: str, recursive: bool):
        root = Path(folder)
        iterator = root.rglob("*") if recursive else root.glob("*")
        files = []
        for p in iterator:
            if p.is_file() and p.suffix.lower() in (".hwp", ".hwpx"):
                files.append(p)
        return files

    def _run_conversion(self, folder: str, recursive: bool, overwrite: bool, use_safe_copy: bool):
        try:
            files = self.collect_files(folder, recursive)
            total = len(files)
            if total == 0:
                self.log_queue.put(("error", "No .hwp or .hwpx files found."))
                return

            log_csv = str(Path(folder) / "hwp_to_pdf_log_safe.csv")
            success = 0
            failed = 0
            skipped = 0

            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            hwp = None

            try:
                TEMP_WORKDIR.mkdir(parents=True, exist_ok=True)

                hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                try:
                    hwp.XHwpWindows.Item(0).Visible = False
                except Exception:
                    pass

                self.log_queue.put(("log", f"Found {total} file(s)."))
                self.log_queue.put(("log", f"CSV log: {log_csv}"))
                self.log_queue.put(("log", f"Safe temp mode: {'ON' if use_safe_copy else 'OFF'}"))

                with open(log_csv, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(["status", "source", "pdf", "message"])

                    for idx, src_path in enumerate(files, start=1):
                        if self.stop_requested:
                            writer.writerow(["STOPPED", "", "", "User requested stop"])
                            self.log_queue.put(("log", "Stopped by user."))
                            break

                        src_path = Path(src_path)
                        pdf_path = src_path.with_suffix(".pdf")
                        fmt = src_path.suffix.replace(".", "").upper()

                        self.log_queue.put(("progress", (idx - 1, total, f"[{idx}/{total}] {src_path.name}")))
                        self.log_queue.put(("log", f"Processing: {src_path}"))

                        temp_input = None
                        temp_pdf = None

                        try:
                            if pdf_path.exists() and not overwrite:
                                skipped += 1
                                msg = "Skipped because PDF already exists"
                                writer.writerow(["SKIPPED", str(src_path), str(pdf_path), msg])
                                self.log_queue.put(("log", f"SKIPPED -> {pdf_path}"))
                                self.log_queue.put(("progress", (idx, total, f"[{idx}/{total}] Skipped")))
                                continue

                            if use_safe_copy:
                                temp_input = TEMP_WORKDIR / f"{idx:05d}_{src_path.name}"
                                temp_pdf = TEMP_WORKDIR / f"{idx:05d}_{src_path.stem}.pdf"

                                if temp_input.exists():
                                    temp_input.unlink()
                                if temp_pdf.exists():
                                    temp_pdf.unlink()

                                shutil.copy2(src_path, temp_input)
                                open_target = temp_input
                                save_target = temp_pdf
                            else:
                                open_target = src_path
                                save_target = pdf_path

                            if pdf_path.exists() and overwrite:
                                pdf_path.unlink()

                            opened = hwp.Open(str(open_target), fmt, DEFAULT_OPEN_OPTION)
                            if opened is False:
                                raise RuntimeError("Open failed")

                            saved = hwp.SaveAs(str(save_target), "PDF", "")
                            if saved is False:
                                raise RuntimeError("SaveAs PDF failed")

                            try:
                                hwp.Clear(1)
                            except Exception:
                                pass

                            if use_safe_copy:
                                if not temp_pdf.exists():
                                    raise RuntimeError("Temporary PDF was not created")
                                shutil.move(str(temp_pdf), str(pdf_path))

                            success += 1
                            writer.writerow(["OK", str(src_path), str(pdf_path), ""])
                            self.log_queue.put(("log", f"OK -> {pdf_path}"))

                        except Exception as e:
                            failed += 1
                            writer.writerow(["FAILED", str(src_path), str(pdf_path), str(e)])
                            self.log_queue.put(("log", f"FAILED -> {src_path} | {e}"))
                            try:
                                hwp.Clear(1)
                            except Exception:
                                pass

                        finally:
                            for tmp in (temp_input, temp_pdf):
                                try:
                                    if tmp and Path(tmp).exists():
                                        Path(tmp).unlink()
                                except Exception:
                                    pass

                        self.log_queue.put(("progress", (idx, total, f"[{idx}/{total}] Done")))

            finally:
                try:
                    if hwp is not None:
                        hwp.Quit()
                except Exception:
                    pass
                pythoncom.CoUninitialize()

            self.log_queue.put(("done", (success, failed, skipped, log_csv)))

        except Exception as e:
            self.log_queue.put(("error", f"Unexpected error:\n{e}"))


def main():
    root = tk.Tk()
    style = ttk.Style(root)
    try:
        style.theme_use("vista")
    except Exception:
        pass
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
