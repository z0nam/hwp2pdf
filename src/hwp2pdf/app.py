import csv
import os
import queue
import shutil
import subprocess
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from hwp2pdf.version import __version__

APP_NAME = "HWP/HWPX -> PDF/DOCX Converter"
APP_TITLE = f"{APP_NAME} v{__version__}"
DEFAULT_OPEN_OPTION = "forceopen:true;versionwarning:false;"
TEMP_WORKDIR = Path(r"C:\temp\hwp_convert_safe")
BASE_EXTENSIONS = (".hwp", ".hwpx")
OUTPUT_FORMATS = {
    "PDF": ".pdf",
    "DOCX": ".docx",
}
SAVE_FORMAT_ALIASES = {
    "PDF": ("PDF",),
    "DOCX": ("OOXML", "DOCX", "MSWORD"),
}
HWP_SECURITY_MODULE = ("FilePathCheckDLL", "FilePathCheckerModule")
MESSAGE_BOX_AUTO_CONFIRM = 0x10


def ensure_pywin32():
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401

        return True, ""
    except Exception as e:
        return False, str(e)


def get_hwp_processes():
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Hwp.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode != 0:
            return []

        processes = []
        for row in csv.reader(result.stdout.splitlines()):
            if len(row) >= 2 and row[0].lower() == "hwp.exe":
                processes.append({"name": row[0], "pid": row[1]})
        return processes
    except Exception:
        return []


def is_hwp_running():
    return bool(get_hwp_processes())


def kill_hwp():
    try:
        result = subprocess.run(
            ["taskkill", "/IM", "Hwp.exe", "/F"],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return result.returncode == 0 and not is_hwp_running()
    except Exception:
        return False


def enabled_extensions():
    return BASE_EXTENSIONS


def output_extension(output_format: str):
    return OUTPUT_FORMATS[output_format]


def save_document_as(hwp, save_target: Path, output_format: str):
    errors = []
    for save_format in SAVE_FORMAT_ALIASES[output_format]:
        previous_message_box_mode = None
        try:
            if output_format == "DOCX":
                _, previous_message_box_mode, _ = enable_auto_confirm_message_boxes(hwp)
            saved = hwp.SaveAs(str(save_target), save_format, "")
            if saved is not False and save_target.exists():
                return save_format
            errors.append(f"SaveAs {save_format} returned {saved}")
        except Exception as e:
            errors.append(f"SaveAs {save_format}: {e}")
        finally:
            if output_format == "DOCX":
                restore_message_box_mode(hwp, previous_message_box_mode)

    for save_format in SAVE_FORMAT_ALIASES[output_format]:
        previous_message_box_mode = None
        try:
            pset = hwp.HParameterSet.HFileOpenSave
            hwp.HAction.GetDefault("FileSaveAs_S", pset.HSet)

            # pywin32 can expose this property with different casing depending on generated wrappers.
            for attr in ("filename", "FileName"):
                try:
                    setattr(pset, attr, str(save_target))
                except Exception:
                    pass

            pset.Format = save_format
            try:
                pset.Attributes = 0
            except Exception:
                pass

            if output_format == "DOCX":
                _, previous_message_box_mode, _ = enable_auto_confirm_message_boxes(hwp)
            executed = hwp.HAction.Execute("FileSaveAs_S", pset.HSet)
            if executed is not False and save_target.exists():
                return save_format
            errors.append(f"FileSaveAs_S {save_format} returned {executed}")
        except Exception as e:
            errors.append(f"FileSaveAs_S {save_format}: {e}")
        finally:
            if output_format == "DOCX":
                restore_message_box_mode(hwp, previous_message_box_mode)

    raise RuntimeError(f"SaveAs {output_format} failed. Tried: {'; '.join(errors)}")


def force_one_page_view(hwp):
    ps = hwp.HParameterSet.HViewProperties
    hwp.HAction.GetDefault("ViewZoom", ps.HSet)
    ps.ZoomCntX = 1
    ps.ZoomCntY = 1
    ps.PageDir = 0
    executed = hwp.HAction.Execute("ViewZoom", ps.HSet)
    if executed is False:
        raise RuntimeError("ViewZoom one-page setting failed")


def enable_auto_confirm_message_boxes(hwp):
    previous_mode = None
    try:
        previous_mode = hwp.GetMessageBoxMode()
    except Exception:
        pass

    try:
        mode = MESSAGE_BOX_AUTO_CONFIRM
        if isinstance(previous_mode, int):
            mode = previous_mode | MESSAGE_BOX_AUTO_CONFIRM
        hwp.SetMessageBoxMode(mode)
        return True, previous_mode, ""
    except Exception as e:
        return False, previous_mode, str(e)


def restore_message_box_mode(hwp, previous_mode):
    if previous_mode is None:
        return
    try:
        hwp.SetMessageBoxMode(previous_mode)
    except Exception:
        pass


def register_hwp_security_module(hwp):
    try:
        module_name, module_class = HWP_SECURITY_MODULE
        return bool(hwp.RegisterModule(module_name, module_class)), ""
    except Exception as e:
        return False, str(e)


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("920x710")
        self.root.minsize(820, 580)

        self.folder_var = tk.StringVar()
        self.overwrite_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=True)
        self.use_safe_copy_var = tk.BooleanVar(value=True)
        self.force_one_page_var = tk.BooleanVar(value=True)
        self.output_pdf_var = tk.BooleanVar(value=True)
        self.output_docx_var = tk.BooleanVar(value=False)

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
        ttk.Button(path_row, text="Browse folder...", command=self.browse_folder).grid(
            row=0, column=1, sticky="e", padx=(0, 8)
        )
        ttk.Button(path_row, text="Pick file...", command=self.pick_file_folder).grid(row=0, column=2, sticky="e")

        opts = ttk.LabelFrame(self.root, text="Options", padding=12)
        opts.pack(fill="x", padx=12, pady=(0, 12))

        ttk.Checkbutton(opts, text="Include subfolders", variable=self.recursive_var).grid(
            row=0, column=0, sticky="w", padx=(0, 16)
        )
        ttk.Checkbutton(opts, text="Overwrite existing output", variable=self.overwrite_var).grid(
            row=0, column=1, sticky="w", padx=(0, 16)
        )

        output_frame = ttk.Frame(opts)
        output_frame.grid(row=0, column=2, sticky="w")
        ttk.Label(output_frame, text="Output").pack(side="left", padx=(0, 8))
        ttk.Checkbutton(output_frame, text="PDF", variable=self.output_pdf_var).pack(side="left")
        ttk.Checkbutton(output_frame, text="DOCX", variable=self.output_docx_var).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(
            opts,
            text="Use safe local temp conversion (recommended for Google Drive / network drives)",
            variable=self.use_safe_copy_var,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Checkbutton(
            opts,
            text="Force one-page view before export",
            variable=self.force_one_page_var,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

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
            "- Close Hancom HWP before starting for best stability.\n"
            "- Safe temp mode copies each file to a short local path before conversion.\n"
            "- One-page view is forced before export by default to avoid two-page PDF output.\n"
            "- DOCX output uses Hancom Office export, so layout fidelity depends on Hancom's DOCX support.\n"
            "- A CSV log will be written to the selected root folder."
        )
        ttk.Label(note_frame, text=notes, justify="left").pack(anchor="w")

    def browse_folder(self):
        initial_dir = self.folder_var.get().strip()
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = str(Path.home())

        folder = filedialog.askdirectory(parent=self.root, title="Select target folder", initialdir=initial_dir)
        if folder:
            self.folder_var.set(folder)

    def pick_file_folder(self):
        initial_dir = self.folder_var.get().strip()
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = str(Path.home())

        selected_file = filedialog.askopenfilename(
            parent=self.root,
            title="Select any file in the target folder",
            initialdir=initial_dir,
            filetypes=[
                ("Supported documents", "*.hwp *.hwpx *.docx *.pdf"),
                ("HWP/HWPX", "*.hwp *.hwpx"),
                ("PDF/DOCX", "*.pdf *.docx"),
                ("All files", "*.*"),
            ],
        )
        if selected_file:
            self.folder_var.set(str(Path(selected_file).parent))

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
                    success, failed, skipped, log_csv, all_success = payload
                    self.is_running = False
                    self.start_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    if all_success:
                        self.progress_label_var.set("Success")
                        messagebox.showinfo(APP_TITLE, "Conversion succeeded.")
                    else:
                        self.progress_label_var.set(f"Done. Success: {success}, Failed: {failed}, Skipped: {skipped}")
                        messagebox.showinfo(
                            APP_TITLE,
                            f"Conversion finished.\n\n"
                            f"Success: {success}\nFailed: {failed}\nSkipped: {skipped}\n\n"
                            f"Log file:\n{log_csv}",
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

        output_formats = self.selected_output_formats()
        if not output_formats:
            messagebox.showerror(APP_TITLE, "Select at least one output format: PDF or DOCX.")
            return

        ok, detail = ensure_pywin32()
        if not ok:
            messagebox.showerror(
                APP_TITLE,
                "pywin32 is not available.\n\nInstall it with:\n"
                "python -m pip install pywin32\n\nDetails:\n" + detail,
            )
            return

        hwp_processes = get_hwp_processes()
        if hwp_processes:
            process_detail = ", ".join(f"PID {process['pid']}" for process in hwp_processes)
            answer = messagebox.askyesnocancel(
                APP_TITLE,
                "Hancom HWP process is already running in the background.\n\n"
                f"Detected: {process_detail}\n\n"
                "Yes: force close HWP and continue\n"
                "No: continue anyway\n"
                "Cancel: stop",
            )
            if answer is None:
                return
            if answer is True:
                killed = kill_hwp()
                if not killed:
                    messagebox.showerror(
                        APP_TITLE,
                        "Could not close HWP automatically.\n\n"
                        "Close Hwp.exe from Task Manager, then start conversion again.",
                    )
                    return

        self.stop_requested = False
        self.is_running = True
        self.log_text.delete("1.0", "end")
        self.append_log("Starting conversion...")
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
                self.force_one_page_var.get(),
                output_formats,
            ),
            daemon=True,
        )
        self.worker.start()

    def selected_output_formats(self):
        formats = []
        if self.output_pdf_var.get():
            formats.append("PDF")
        if self.output_docx_var.get():
            formats.append("DOCX")
        return tuple(formats)

    @staticmethod
    def collect_files(folder: str, recursive: bool):
        root = Path(folder)
        iterator = root.rglob("*") if recursive else root.glob("*")
        allowed_extensions = enabled_extensions()
        files = []
        for path in iterator:
            if path.is_file() and path.suffix.lower() in allowed_extensions:
                files.append(path)
        return files

    def _run_conversion(
        self, folder: str, recursive: bool, overwrite: bool, use_safe_copy: bool, force_one_page: bool, output_formats
    ):
        try:
            self.log_queue.put(("log", "Scanning files..."))
            files = self.collect_files(folder, recursive)
            total_files = len(files)
            total_jobs = total_files * len(output_formats)
            extension_label = ", ".join(ext.upper() for ext in enabled_extensions())
            if total_files == 0:
                self.log_queue.put(("error", f"No {extension_label} files found."))
                return

            log_csv = str(Path(folder) / "hwp2pdf_log.csv")
            success = 0
            failed = 0
            skipped = 0
            stopped = False

            import pythoncom
            import win32com.client

            self.log_queue.put(("log", "Initializing Hancom COM automation..."))
            pythoncom.CoInitialize()
            hwp = None

            try:
                TEMP_WORKDIR.mkdir(parents=True, exist_ok=True)

                self.log_queue.put(("log", "Starting HWPFrame.HwpObject..."))
                hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                self.log_queue.put(("log", "HWPFrame.HwpObject started."))
                try:
                    hwp.XHwpWindows.Item(0).Visible = False
                except Exception:
                    pass

                self.log_queue.put(("log", "Registering HWP file access security module..."))
                security_ok, security_detail = register_hwp_security_module(hwp)

                self.log_queue.put(("log", f"Found {total_files} file(s): {extension_label}"))
                self.log_queue.put(("log", f"CSV log: {log_csv}"))
                self.log_queue.put(("log", f"Safe temp mode: {'ON' if use_safe_copy else 'OFF'}"))
                self.log_queue.put(("log", f"Force one-page view: {'ON' if force_one_page else 'OFF'}"))
                self.log_queue.put(("log", f"Output formats: {', '.join(output_formats)}"))
                self.log_queue.put(("log", "Auto-confirm DOCX compatibility dialogs: ON"))
                security_msg = "ON" if security_ok else f"OFF ({security_detail or 'module unavailable'})"
                self.log_queue.put(("log", f"HWP file access security module: {security_msg}"))

                with open(log_csv, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(["status", "source", "output", "message"])

                    job_index = 0
                    for file_index, src_path in enumerate(files, start=1):
                        src_path = Path(src_path)
                        fmt = src_path.suffix.replace(".", "").upper()

                        self.log_queue.put(("log", f"Processing: {src_path}"))

                        for output_format in output_formats:
                            if self.stop_requested:
                                writer.writerow(["STOPPED", "", "", "User requested stop"])
                                self.log_queue.put(("log", "Stopped by user."))
                                stopped = True
                                break

                            job_index += 1
                            output_ext = output_extension(output_format)
                            output_path = src_path.with_suffix(output_ext)
                            self.log_queue.put(
                                (
                                    "progress",
                                    (
                                        job_index - 1,
                                        total_jobs,
                                        f"[{job_index}/{total_jobs}] {src_path.name} -> {output_format}",
                                    ),
                                )
                            )

                            temp_input = None
                            temp_output = None

                            try:
                                if output_path.exists() and not overwrite:
                                    skipped += 1
                                    msg = f"Skipped because {output_format} already exists"
                                    writer.writerow(["SKIPPED", str(src_path), str(output_path), msg])
                                    self.log_queue.put(("log", f"SKIPPED {output_format} -> {output_path}"))
                                    self.log_queue.put(
                                        ("progress", (job_index, total_jobs, f"[{job_index}/{total_jobs}] Skipped"))
                                    )
                                    continue

                                if use_safe_copy:
                                    temp_input = TEMP_WORKDIR / f"{file_index:05d}_{output_format}_{src_path.name}"
                                    temp_output = TEMP_WORKDIR / f"{file_index:05d}_{output_format}_{src_path.stem}{output_ext}"

                                    if temp_input.exists():
                                        temp_input.unlink()
                                    if temp_output.exists():
                                        temp_output.unlink()

                                    shutil.copy2(src_path, temp_input)
                                    open_target = temp_input
                                    save_target = temp_output
                                else:
                                    open_target = src_path
                                    save_target = output_path

                                if output_path.exists() and overwrite:
                                    output_path.unlink()

                                opened = hwp.Open(str(open_target), "", DEFAULT_OPEN_OPTION)
                                if opened is False:
                                    raise RuntimeError(f"Open failed for {fmt}")

                                if force_one_page:
                                    force_one_page_view(hwp)

                                actual_save_format = save_document_as(hwp, save_target, output_format)

                                try:
                                    hwp.Clear(1)
                                except Exception:
                                    pass

                                if use_safe_copy:
                                    if not temp_output.exists():
                                        raise RuntimeError(f"Temporary {output_format} was not created")
                                    shutil.move(str(temp_output), str(output_path))

                                success += 1
                                writer.writerow(["OK", str(src_path), str(output_path), ""])
                                self.log_queue.put(
                                    ("log", f"OK {output_format} ({actual_save_format}) -> {output_path}")
                                )

                            except Exception as e:
                                failed += 1
                                writer.writerow(["FAILED", str(src_path), str(output_path), str(e)])
                                self.log_queue.put(("log", f"FAILED {output_format} -> {src_path} | {e}"))
                                try:
                                    hwp.Clear(1)
                                except Exception:
                                    pass

                            finally:
                                for tmp in (temp_input, temp_output):
                                    try:
                                        if tmp and Path(tmp).exists():
                                            Path(tmp).unlink()
                                    except Exception:
                                        pass

                            self.log_queue.put(("progress", (job_index, total_jobs, f"[{job_index}/{total_jobs}] Done")))

                        if self.stop_requested:
                            stopped = True
                            break

            finally:
                try:
                    if hwp is not None:
                        hwp.Quit()
                except Exception:
                    pass
                pythoncom.CoUninitialize()

            all_success = success == total_jobs and failed == 0 and skipped == 0 and not stopped
            if all_success:
                try:
                    Path(log_csv).unlink()
                except FileNotFoundError:
                    pass
                except Exception as e:
                    all_success = False
                    self.log_queue.put(("log", f"Could not remove success log: {e}"))

            self.log_queue.put(("done", (success, failed, skipped, log_csv, all_success)))

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
