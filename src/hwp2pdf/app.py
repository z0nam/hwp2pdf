import os
import queue
import subprocess
import sys
import textwrap
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

from hwp2pdf import config
from hwp2pdf.backends import BackendUnavailable, create_backend
from hwp2pdf.backends.windows_com import (
    ensure_pywin32,
    get_hwp_processes,
    is_hwp_running,
    kill_hwp,
)
from hwp2pdf.constants import APP_NAME, APP_TITLE, enabled_extensions
from hwp2pdf.i18n import LANGUAGE_CODES, LANGUAGE_LABELS, translate
from hwp2pdf.jobs import collect_files, run_batch
from hwp2pdf.paths import IS_WINDOWS, reveal_in_file_manager
from hwp2pdf.updater import (
    GITHUB_RELEASES_PAGE_URL,
    UPDATE_DOWNLOAD_DIR,
    fetch_latest_release,
    is_installed_build,
    is_setup_asset_url,
    latest_release_download_url,
    latest_release_version,
    load_update_state,
    parse_version,
    save_update_state,
    should_check_updates,
)
from hwp2pdf.version import __version__

# Compatibility surface. Before the backend split every one of these names was
# defined in this module, and outside callers import them from here --
# ``cli.py``, ``src/hwp_pdf_converter_app_safe.py`` and
# ``scripts/check_windows.ps1`` among them. Re-export the whole previous public
# surface so a split that was meant to be internal cannot break them.
from hwp2pdf.backends.windows_com import (  # noqa: F401
    HancomDialogWatcher,
    blocked_conversion_reason,
    build_save_failure_message,
    configure_pdf_print,
    detect_hwp_arch,
    enable_auto_confirm_message_boxes,
    ensure_hwp_security_module_registered,
    force_one_page_view,
    hwp_process_id,
    is_nup_print_method,
    read_hwp_file_flags,
    register_hwp_security_module,
    reset_pdf_print_method,
    restore_message_box_mode,
    save_document_as,
    save_pdf_with_print_to_pdf,
    set_hwp_parameter,
)
from hwp2pdf.constants import (  # noqa: F401
    BASE_EXTENSIONS,
    DEFAULT_OPEN_OPTION,
    HANCOM_BLOCKING_DIALOG_MESSAGES,
    HANCOM_DIALOG_CONFIRM_BUTTONS,
    HWP_FILE_SIGNATURE,
    HWP_FILEHEADER_STREAM,
    HWP_FLAG_COMPRESSED,
    HWP_FLAG_DISTRIBUTION_DOCUMENT,
    HWP_FLAG_PASSWORD_PROTECTED,
    HWP_SECURITY_DLL_NAME,
    HWP_SECURITY_INSTALL_DIR,
    HWP_SECURITY_MODULE,
    HWP_SECURITY_REG_KEY,
    HWP_SECURITY_REG_VALUE,
    MESSAGE_BOX_AUTO_CONFIRM,
    OUTPUT_FORMATS,
    SAVE_FORMAT_ALIASES,
    TEMP_WORKDIR,
    output_extension,
)
from hwp2pdf.i18n import PRINT_METHOD_LABELS, TEXT, print_method_label  # noqa: F401
from hwp2pdf.updater import (  # noqa: F401
    GITHUB_RELEASES_API_URL,
    UPDATE_CHECK_INTERVAL_SECONDS,
    UPDATE_STATE_PATH,
)

#: Names this module has always exposed. ``tests/test_app_surface.py`` asserts
#: that every one of them still resolves. Written as a literal so linters can
#: see that the re-export imports above are deliberate.
__all__ = [
    "APP_NAME", "APP_TITLE", "BASE_EXTENSIONS", "DEFAULT_OPEN_OPTION",
    "GITHUB_RELEASES_API_URL", "GITHUB_RELEASES_PAGE_URL",
    "HANCOM_BLOCKING_DIALOG_MESSAGES", "HANCOM_DIALOG_CONFIRM_BUTTONS",
    "HWP_FILEHEADER_STREAM", "HWP_FILE_SIGNATURE", "HWP_FLAG_COMPRESSED",
    "HWP_FLAG_DISTRIBUTION_DOCUMENT", "HWP_FLAG_PASSWORD_PROTECTED",
    "HWP_SECURITY_DLL_NAME", "HWP_SECURITY_INSTALL_DIR", "HWP_SECURITY_MODULE",
    "HWP_SECURITY_REG_KEY", "HWP_SECURITY_REG_VALUE", "HancomDialogWatcher",
    "LANGUAGE_CODES", "LANGUAGE_LABELS", "MESSAGE_BOX_AUTO_CONFIRM",
    "ModernGradientButton", "OUTPUT_FORMATS", "PRINT_METHOD_LABELS",
    "SAVE_FORMAT_ALIASES", "START_BUTTON_PALETTE", "STOP_BUTTON_PALETTE",
    "TEMP_WORKDIR", "TEXT", "UPDATE_CHECK_INTERVAL_SECONDS",
    "UPDATE_DOWNLOAD_DIR", "UPDATE_STATE_PATH", "ConverterApp",
    "blend_hex_color", "blocked_conversion_reason", "build_save_failure_message",
    "configure_pdf_print", "detect_hwp_arch", "enable_auto_confirm_message_boxes",
    "enabled_extensions", "ensure_hwp_security_module_registered",
    "ensure_pywin32", "fetch_latest_release", "force_one_page_view",
    "get_hwp_processes", "hwp_process_id", "is_hwp_running", "is_installed_build",
    "is_nup_print_method", "is_setup_asset_url", "kill_hwp",
    "latest_release_download_url", "latest_release_version", "load_update_state",
    "main", "output_extension", "parse_version", "print_method_label",
    "read_hwp_file_flags", "register_hwp_security_module",
    "reset_pdf_print_method", "restore_message_box_mode", "save_document_as",
    "save_pdf_with_print_to_pdf", "save_update_state", "set_hwp_parameter",
    "should_check_updates", "translate",
]

LEGACY_EXPORTS = tuple(__all__)


def blend_hex_color(start: str, end: str, ratio: float):
    ratio = max(0.0, min(1.0, ratio))
    start_rgb = tuple(int(start[index : index + 2], 16) for index in (1, 3, 5))
    end_rgb = tuple(int(end[index : index + 2], 16) for index in (1, 3, 5))
    blended = tuple(
        round(start_channel + (end_channel - start_channel) * ratio)
        for start_channel, end_channel in zip(start_rgb, end_rgb)
    )
    return "#{:02x}{:02x}{:02x}".format(*blended)


class ModernGradientButton(tk.Button):
    """Rounded image-backed button with subtle gradients and native semantics."""

    def __init__(
        self,
        master,
        *,
        command,
        palette,
        icon,
        text="",
        state="normal",
        disabled_foreground="#5f6662",
        width=142,
        height=38,
    ):
        self._button_width = width
        self._button_height = height
        self._palette = palette
        self._icon = icon
        self._hovered = False
        self._pressed = False
        self._logical_state = state
        self._command = command
        self._disabled_foreground = disabled_foreground
        self._images = {}

        style = ttk.Style(master)
        surface_color = style.lookup("TFrame", "background") or "#f0f0f0"
        if not surface_color.startswith("#"):
            red, green, blue = master.winfo_rgb(surface_color)
            surface_color = "#{:02x}{:02x}{:02x}".format(
                red // 256, green // 256, blue // 256
            )
        self._button_font = tkfont.nametofont("TkDefaultFont", root=master).copy()
        self._button_font.configure(weight="bold")
        for name, colors in palette.items():
            self._images[name] = self._create_gradient_image(
                master,
                colors[0],
                colors[1],
                colors[2],
                surface_color,
            )

        super().__init__(
            master,
            command=self._invoke_command,
            text=text,
            state="normal",
            image=self._images["normal"],
            compound="center",
            font=self._button_font,
            foreground="#ffffff",
            activeforeground="#ffffff",
            disabledforeground=disabled_foreground,
            background=surface_color,
            activebackground=surface_color,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=surface_color,
            highlightcolor=palette["normal"][2],
            padx=0,
            pady=0,
            relief="flat",
            takefocus=state != "disabled",
            cursor="hand2" if state != "disabled" else "arrow",
        )

        self.bind("<Enter>", self._on_enter, add="+")
        self.bind("<Leave>", self._on_leave, add="+")
        self.bind("<ButtonPress-1>", self._on_press, add="+")
        self.bind("<ButtonRelease-1>", self._on_release, add="+")
        self.bind("<FocusIn>", lambda _event: self._sync_visual(), add="+")
        self.bind("<FocusOut>", lambda _event: self._sync_visual(), add="+")
        self._sync_visual()

    def _create_gradient_image(self, master, top, bottom, border, surface):
        image = tk.PhotoImage(
            master=master,
            width=self._button_width,
            height=self._button_height,
        )
        radius = 9
        for y in range(self._button_height):
            ratio = y / max(self._button_height - 1, 1)
            fill = blend_hex_color(top, bottom, ratio)
            if y < self._button_height * 0.35:
                highlight = 0.045 * (1 - y / (self._button_height * 0.35))
                fill = blend_hex_color(fill, "#ffffff", highlight)

            row = []
            for x in range(self._button_width):
                if not self._inside_rounded_rect(
                    x, y, self._button_width, self._button_height, radius
                ):
                    row.append(surface)
                elif not self._inside_rounded_rect(
                    x - 1,
                    y - 1,
                    self._button_width - 2,
                    self._button_height - 2,
                    radius - 1,
                ):
                    row.append(border)
                else:
                    row.append(fill)
            image.put("{" + " ".join(row) + "}", to=(0, y))
        self._draw_debossed_icon(image, top, bottom)
        return image

    def _draw_debossed_icon(self, image, top, bottom):
        pixels = self._icon_pixels(self._icon)
        icon_x = 11 if self._button_width < 100 else 17
        icon_y = self._button_height // 2
        shadow = blend_hex_color(bottom, "#000000", 0.28)
        highlight = blend_hex_color(top, "#ffffff", 0.30)

        for x, y in pixels:
            image.put(highlight, to=(icon_x + x + 1, icon_y + y + 1))
        for x, y in pixels:
            image.put(shadow, to=(icon_x + x, icon_y + y))

    @staticmethod
    def _icon_pixels(icon):
        if icon == "play":
            pixels = []
            for y in range(-6, 7):
                max_x = round(8 * (1 - abs(y) / 6))
                pixels.extend((x, y) for x in range(max_x + 1))
            return pixels
        if icon == "stop":
            return [(x, y) for y in range(-5, 5) for x in range(10)]
        return []

    @staticmethod
    def _inside_rounded_rect(x, y, width, height, radius):
        if x < 0 or y < 0 or x >= width or y >= height:
            return False
        if radius <= x < width - radius or radius <= y < height - radius:
            return True

        center_x = radius - 0.5 if x < radius else width - radius - 0.5
        center_y = radius - 0.5 if y < radius else height - radius - 0.5
        return (x - center_x) ** 2 + (y - center_y) ** 2 <= radius**2

    def _on_enter(self, _event):
        self._hovered = True
        self._sync_visual()

    def _on_leave(self, _event):
        self._hovered = False
        self._pressed = False
        self._sync_visual()

    def _on_press(self, _event):
        if self._logical_state == "disabled":
            return "break"
        self._pressed = True
        self._sync_visual()

    def _on_release(self, event):
        if self._logical_state == "disabled":
            return "break"
        self._pressed = False
        self._hovered = 0 <= event.x < self.winfo_width() and 0 <= event.y < self.winfo_height()
        self._sync_visual()

    def _sync_visual(self):
        if not self._images:
            return

        if self._logical_state == "disabled":
            image_name = "disabled"
            cursor = "arrow"
            foreground = self._disabled_foreground
        elif self._pressed:
            image_name = "pressed"
            cursor = "hand2"
            foreground = "#ffffff"
        elif self._hovered:
            image_name = "hover"
            cursor = "hand2"
            foreground = "#ffffff"
        else:
            image_name = "normal"
            cursor = "hand2"
            foreground = "#ffffff"
        super().configure(
            image=self._images[image_name],
            cursor=cursor,
            foreground=foreground,
            activeforeground=foreground,
            takefocus=self._logical_state != "disabled",
        )

    def _invoke_command(self):
        if self._logical_state != "disabled" and self._command is not None:
            return self._command()
        return None

    def invoke(self):
        if self._logical_state == "disabled":
            return ""
        return super().invoke()

    def cget(self, key):
        if key == "state":
            return self._logical_state
        return super().cget(key)

    def configure(self, cnf=None, **kwargs):
        if isinstance(cnf, dict) and "state" in cnf:
            cnf = dict(cnf)
            self._logical_state = cnf.pop("state")
        if "state" in kwargs:
            self._logical_state = kwargs.pop("state")
        result = super().configure(cnf, **kwargs)
        if hasattr(self, "_images"):
            self._sync_visual()
        return result

    config = configure


START_BUTTON_PALETTE = {
    "normal": ("#2caf78", "#17895b", "#126f49"),
    "hover": ("#38bb83", "#1b9864", "#137650"),
    "pressed": ("#168050", "#0f6e43", "#0b5b37"),
    "disabled": ("#a4bdb1", "#8fa79b", "#82978d"),
}

STOP_BUTTON_PALETTE = {
    "normal": ("#e76066", "#c83f47", "#a93239"),
    "hover": ("#ee6c71", "#d14951", "#b23840"),
    "pressed": ("#bd3941", "#a92e35", "#8f252b"),
    "disabled": ("#c7aaac", "#b18f92", "#9e7e81"),
}


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("920x710")
        self.root.minsize(820, 580)

        self.settings = config.load()
        saved = self.settings["options"]
        server = config.server_settings(self.settings)

        self.folder_var = tk.StringVar(value=self.settings.get("last_target", ""))
        self.overwrite_var = tk.BooleanVar(value=saved["overwrite"])
        self.recursive_var = tk.BooleanVar(value=saved["recursive"])
        self.folder_recursive_preference = saved["recursive"]
        self._updating_recursive_for_target = False
        self.use_safe_copy_var = tk.BooleanVar(value=saved["safe_temp"])
        self.force_one_page_var = tk.BooleanVar(value=saved["force_one_page"])
        self.output_pdf_var = tk.BooleanVar(value="PDF" in saved["formats"])
        self.output_docx_var = tk.BooleanVar(value="DOCX" in saved["formats"])
        self.language_var = tk.StringVar(
            value=LANGUAGE_LABELS.get(self.settings.get("language"), LANGUAGE_LABELS["ko"])
        )

        # Conversion happens locally on Windows and on a server everywhere else.
        self.server_url_var = tk.StringVar(value=server.get("url", ""))
        self.server_token_var = tk.StringVar(value=server.get("token", ""))
        self.server_transport_var = tk.StringVar(value=server.get("transport", config.TRANSPORT_AUTO))
        self.transport_label_var = tk.StringVar()
        self.use_remote_var = tk.BooleanVar(value=not IS_WINDOWS or bool(server.get("url")))
        self.server_status_var = tk.StringVar()
        self.server_test_running = False
        self._save_settings_job = None

        self.log_queue = queue.Queue()
        self.worker = None
        self.stop_requested = False
        self.is_running = False
        self.recursive_check = None
        self.file_count_var = tk.StringVar()
        self.update_status_var = tk.StringVar()
        self.upgrade_btn = None
        self.latest_release_url = GITHUB_RELEASES_PAGE_URL
        self.latest_download_url = GITHUB_RELEASES_PAGE_URL
        self.update_check_running = False
        self.ui = {}

        self._build_ui()
        try:
            TkinterDnD.require(self.root)
            drop_widget = (
                self.root
                if hasattr(self.root, "drop_target_register")
                else self.ui["path_entry"]
            )
            drop_widget.drop_target_register(DND_FILES)
            drop_widget.dnd_bind("<<Drop>>", self._on_drop_event)
            self.drop_target_enabled = True
        except Exception:
            self.drop_target_enabled = False
            self.ui["drop_hint_label"].grid_remove()
        self.folder_var.trace_add("write", self._on_target_path_changed)
        self.recursive_var.trace_add("write", self._on_recursive_option_changed)
        for var in (
            self.folder_var, self.overwrite_var, self.recursive_var,
            self.use_safe_copy_var, self.force_one_page_var,
            self.output_pdf_var, self.output_docx_var, self.language_var,
            self.server_url_var, self.server_token_var, self.server_transport_var,
            self.use_remote_var,
        ):
            var.trace_add("write", self._schedule_save_settings)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._apply_cached_update_state()
        self._poll_log_queue()
        self.root.after(1000, self.check_for_updates_if_due)

    def lang(self):
        return LANGUAGE_CODES.get(self.language_var.get(), "ko")

    def tr(self, key: str, **kwargs):
        return translate(self.lang(), key, **kwargs)

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="x")

        header_row = ttk.Frame(top)
        header_row.grid(row=0, column=0, sticky="ew")
        header_row.columnconfigure(0, weight=1)
        self.ui["target_label"] = ttk.Label(header_row)
        self.ui["target_label"].grid(row=0, column=0, sticky="w")
        language_frame = ttk.Frame(header_row)
        language_frame.grid(row=0, column=1, sticky="e")
        self.ui["language_label"] = ttk.Label(language_frame)
        self.ui["language_label"].pack(side="left", padx=(0, 6))
        language_combo = ttk.Combobox(
            language_frame,
            textvariable=self.language_var,
            values=[LANGUAGE_LABELS["ko"], LANGUAGE_LABELS["en"]],
            width=10,
            state="readonly",
        )
        language_combo.pack(side="left")
        language_combo.bind("<<ComboboxSelected>>", self._on_language_changed)
        top.columnconfigure(0, weight=1)

        path_row = ttk.Frame(top)
        path_row.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        path_row.columnconfigure(0, weight=1)

        self.ui["path_entry"] = ttk.Entry(path_row, textvariable=self.folder_var)
        self.ui["path_entry"].grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.ui["browse_btn"] = ttk.Button(path_row, command=self.browse_folder)
        self.ui["browse_btn"].grid(row=0, column=1, sticky="e", padx=(0, 8))
        self.ui["pick_btn"] = ttk.Button(path_row, command=self.pick_file_folder)
        self.ui["pick_btn"].grid(row=0, column=2, sticky="e")
        self.ui["file_count_label"] = ttk.Label(top, textvariable=self.file_count_var)
        self.ui["file_count_label"].grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.ui["drop_hint_label"] = ttk.Label(top)
        self.ui["drop_hint_label"].grid(row=3, column=0, sticky="w", pady=(4, 0))

        update_row = ttk.Frame(top)
        update_row.grid(row=4, column=0, sticky="w", pady=(4, 0))
        self.ui["update_status_label"] = ttk.Label(update_row, textvariable=self.update_status_var)
        self.ui["update_status_label"].pack(side="left")
        self.auto_update_btn = ttk.Button(update_row, command=self.start_auto_update)
        self.auto_update_btn.pack(side="left", padx=(8, 0))
        self.auto_update_btn.pack_forget()
        self.upgrade_btn = ttk.Button(update_row, command=self.open_latest_release)
        self.upgrade_btn.pack(side="left", padx=(8, 0))
        self.upgrade_btn.pack_forget()

        opts = ttk.LabelFrame(self.root, padding=12)
        self.ui["opts"] = opts
        opts.pack(fill="x", padx=12, pady=(0, 12))

        self.recursive_check = ttk.Checkbutton(opts, variable=self.recursive_var)
        self.recursive_check.grid(row=0, column=0, sticky="w", padx=(0, 16))
        self.ui["overwrite_check"] = ttk.Checkbutton(opts, variable=self.overwrite_var)
        self.ui["overwrite_check"].grid(
            row=0, column=1, sticky="w", padx=(0, 16)
        )

        output_frame = ttk.Frame(opts)
        output_frame.grid(row=0, column=2, sticky="w")
        self.ui["output_label"] = ttk.Label(output_frame)
        self.ui["output_label"].pack(side="left", padx=(0, 8))
        ttk.Checkbutton(output_frame, text="PDF", variable=self.output_pdf_var).pack(side="left")
        ttk.Checkbutton(output_frame, text="DOCX", variable=self.output_docx_var).pack(side="left", padx=(8, 0))
        self.ui["safe_temp_check"] = ttk.Checkbutton(
            opts,
            variable=self.use_safe_copy_var,
        )
        self.ui["safe_temp_check"].grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.ui["force_one_page_check"] = ttk.Checkbutton(
            opts,
            variable=self.force_one_page_var,
        )
        self.ui["force_one_page_check"].grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        server_frame = ttk.LabelFrame(self.root, padding=12)
        self.ui["server_frame"] = server_frame
        server_frame.columnconfigure(1, weight=1)

        self.ui["server_remote_check"] = ttk.Checkbutton(
            server_frame, variable=self.use_remote_var, command=self._apply_backend_mode
        )
        if IS_WINDOWS:
            # Elsewhere there is no local engine, so the choice is not offered.
            self.ui["server_remote_check"].grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        self.ui["server_address_label"] = ttk.Label(server_frame)
        self.ui["server_address_label"].grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.ui["server_address_entry"] = ttk.Entry(server_frame, textvariable=self.server_url_var)
        self.ui["server_address_entry"].grid(row=1, column=1, sticky="ew")
        self.ui["server_test_btn"] = ttk.Button(server_frame, command=self.test_server_connection)
        self.ui["server_test_btn"].grid(row=1, column=2, sticky="e", padx=(8, 0))

        self.ui["server_token_label"] = ttk.Label(server_frame)
        self.ui["server_token_label"].grid(row=2, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
        self.ui["server_token_entry"] = ttk.Entry(
            server_frame, textvariable=self.server_token_var, show="\u2022"
        )
        self.ui["server_token_entry"].grid(row=2, column=1, sticky="ew", pady=(6, 0))

        self.ui["server_transport_label"] = ttk.Label(server_frame)
        self.ui["server_transport_label"].grid(row=3, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
        self.ui["server_transport_combo"] = ttk.Combobox(
            server_frame, state="readonly", width=14, textvariable=self.transport_label_var
        )
        self.ui["server_transport_combo"].grid(row=3, column=1, sticky="w", pady=(6, 0))
        self.ui["server_transport_combo"].bind("<<ComboboxSelected>>", self._on_transport_changed)

        self.ui["server_status_label"] = ttk.Label(
            server_frame, textvariable=self.server_status_var, wraplength=760, justify="left"
        )
        self.ui["server_status_label"].grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        actions = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        self.ui["actions_frame"] = actions
        actions.pack(fill="x")

        self.start_btn = ModernGradientButton(
            actions,
            command=self.start_conversion,
            palette=START_BUTTON_PALETTE,
            icon="play",
        )
        self.start_btn.pack(side="left")

        self.stop_btn = ModernGradientButton(
            actions,
            command=self.request_stop,
            palette=STOP_BUTTON_PALETTE,
            icon="stop",
            state="disabled",
            disabled_foreground="#72575a",
            width=80,
        )
        self.stop_btn.pack(side="left", padx=(8, 0))

        self.ui["open_btn"] = ttk.Button(actions, command=self.open_selected_folder)
        self.ui["open_btn"].pack(side="left", padx=(8, 0))

        progress_frame = ttk.Frame(self.root, padding=(12, 0, 12, 0))
        progress_frame.pack(fill="x")

        self.progress_label_var = tk.StringVar()
        ttk.Label(progress_frame, textvariable=self.progress_label_var).pack(anchor="w")

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(6, 0))

        log_frame = ttk.Frame(self.root, padding=12)
        log_frame.pack(fill="both", expand=True)

        self.ui["log_label"] = ttk.Label(log_frame)
        self.ui["log_label"].pack(anchor="w")
        self.log_text = tk.Text(log_frame, wrap="word")
        self.log_text.tag_configure("error", foreground="#b00020")
        self.log_text.tag_configure("warning", foreground="#8a5a00")
        self.log_text.pack(side="left", fill="both", expand=True)

        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        note_frame = ttk.LabelFrame(self.root, padding=12)
        self.ui["note_frame"] = note_frame
        note_frame.pack(fill="x", padx=12, pady=(0, 12))
        self.ui["notes_label"] = ttk.Label(note_frame, justify="left")
        self.ui["notes_label"].pack(anchor="w")
        self._apply_language()

    def _on_language_changed(self, _event=None):
        self._apply_language()

    def _apply_language(self):
        self.root.title(APP_TITLE)
        self.ui["target_label"].configure(text=self.tr("target_label"))
        self.ui["drop_hint_label"].configure(text=self.tr("drop_hint"))
        self.ui["language_label"].configure(text=self.tr("language"))
        self.ui["browse_btn"].configure(text=self.tr("browse_folder"))
        self.ui["pick_btn"].configure(text=self.tr("pick_file"))
        self.ui["opts"].configure(text=self.tr("options"))
        self.recursive_check.configure(text=self.tr("include_subfolders"))
        self.ui["overwrite_check"].configure(text=self.tr("overwrite"))
        self.ui["output_label"].configure(text=self.tr("output"))
        self.ui["safe_temp_check"].configure(text=self.tr("safe_temp"))
        self.ui["force_one_page_check"].configure(text=self.tr("force_one_page"))
        self.start_btn.configure(text=self.tr("start"))
        self.stop_btn.configure(text=self.tr("stop"))
        self.ui["open_btn"].configure(text=self.tr("open_selected"))
        self.upgrade_btn.configure(text=self.tr("upgrade"))
        self.auto_update_btn.configure(text=self.tr("auto_update"))
        self.ui["log_label"].configure(text=self.tr("log"))
        self.ui["note_frame"].configure(text=self.tr("notes_title"))
        self.ui["notes_label"].configure(
            text=self.tr("notes_remote" if self.use_remote_var.get() else "notes")
        )
        self.ui["server_frame"].configure(text=self.tr("server_section"))
        self.ui["server_remote_check"].configure(text=self.tr("server_use_remote"))
        self.ui["server_address_label"].configure(text=self.tr("server_address"))
        self.ui["server_token_label"].configure(text=self.tr("server_token"))
        self.ui["server_transport_label"].configure(text=self.tr("server_transport_label"))
        self.ui["server_test_btn"].configure(text=self.tr("server_test"))
        self._apply_transport_labels()
        self._apply_backend_mode()
        if not self.is_running:
            self.progress_label_var.set(self.tr("ready"))
        self._apply_cached_update_state()
        self._update_file_count_estimate()

    def browse_folder(self):
        initial_dir = self.folder_var.get().strip()
        if initial_dir and os.path.isfile(initial_dir):
            initial_dir = str(Path(initial_dir).parent)
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = str(Path.home())

        folder = filedialog.askdirectory(
            parent=self.root, title=self.tr("select_folder_title"), initialdir=initial_dir
        )
        if folder:
            self.folder_var.set(folder)

    def pick_file_folder(self):
        initial_dir = self.folder_var.get().strip()
        if initial_dir and os.path.isfile(initial_dir):
            initial_dir = str(Path(initial_dir).parent)
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = str(Path.home())

        selected_file = filedialog.askopenfilename(
            parent=self.root,
            title=self.tr("select_file_title"),
            initialdir=initial_dir,
            filetypes=[
                ("HWP/HWPX", "*.hwp *.hwpx"),
                (self.tr("all_files"), "*.*"),
            ],
        )
        if selected_file:
            self.folder_var.set(selected_file)

    def _handle_dropped_paths(self, paths):
        valid_targets = []
        for raw_path in paths:
            path = Path(raw_path)
            if path.is_dir() or (
                path.is_file() and path.suffix.lower() in enabled_extensions()
            ):
                valid_targets.append(path)

        if not valid_targets:
            messagebox.showwarning(APP_TITLE, self.tr("invalid_drop"))
            return

        self.folder_var.set(str(valid_targets[0]))
        if len(paths) > 1:
            messagebox.showinfo(APP_TITLE, self.tr("multiple_drop"))

    def _on_drop_event(self, event):
        paths = self.root.tk.splitlist(event.data)
        self._handle_dropped_paths(paths)
        return event.action

    def _on_target_path_changed(self, *_args):
        if self.recursive_check is None:
            return

        target = self.folder_var.get().strip()
        if target and os.path.isfile(target):
            self.folder_recursive_preference = self.recursive_var.get()
            self._set_recursive_for_target(False)
            self.recursive_check.state(["disabled"])
        else:
            self.recursive_check.state(["!disabled"])
            if target and os.path.isdir(target):
                self._set_recursive_for_target(self.folder_recursive_preference)
        self._update_file_count_estimate()

    def _on_recursive_option_changed(self, *_args):
        if self._updating_recursive_for_target:
            return

        target = self.folder_var.get().strip()
        if target and os.path.isdir(target):
            self.folder_recursive_preference = self.recursive_var.get()
        self._update_file_count_estimate()

    def _set_recursive_for_target(self, value: bool):
        if self.recursive_var.get() == value:
            return

        self._updating_recursive_for_target = True
        try:
            self.recursive_var.set(value)
        finally:
            self._updating_recursive_for_target = False

    def _update_file_count_estimate(self):
        if not hasattr(self, "file_count_var"):
            return

        target = self.folder_var.get().strip()
        if not target:
            self.file_count_var.set("")
            return

        try:
            if os.path.isfile(target):
                count = 1 if Path(target).suffix.lower() in enabled_extensions() else 0
            elif os.path.isdir(target):
                count = len(self.collect_files(target, self.recursive_var.get()))
            else:
                self.file_count_var.set("")
                return
        except Exception:
            self.file_count_var.set(self.tr("file_count_unavailable"))
            return

        self.file_count_var.set(self.tr("file_count_estimate", count=count))

    # -- backend settings ------------------------------------------------
    def _transport_label_map(self):
        return {
            config.TRANSPORT_AUTO: self.tr("transport_auto"),
            config.TRANSPORT_UPLOAD: self.tr("transport_upload"),
            config.TRANSPORT_SHARE: self.tr("transport_share"),
        }

    def _apply_transport_labels(self):
        labels = self._transport_label_map()
        self.ui["server_transport_combo"].configure(values=list(labels.values()))
        self.transport_label_var.set(labels.get(self.server_transport_var.get(), labels[config.TRANSPORT_AUTO]))

    def _on_transport_changed(self, _event=None):
        chosen = self.transport_label_var.get()
        for code, label in self._transport_label_map().items():
            if label == chosen:
                self.server_transport_var.set(code)
                return

    def use_remote_backend(self) -> bool:
        return bool(self.use_remote_var.get()) or not IS_WINDOWS

    def _apply_backend_mode(self):
        """Show the server panel only when conversion happens remotely."""
        remote = self.use_remote_backend()
        frame = self.ui["server_frame"]
        # winfo_ismapped() is false while the window is withdrawn, so ask the
        # geometry manager whether the frame is in the layout at all.
        packed = frame.winfo_manager() == "pack"
        if remote and not packed:
            frame.pack(fill="x", padx=12, pady=(0, 12), before=self.ui["actions_frame"])
        elif not remote and packed:
            frame.pack_forget()
        state = "normal" if remote else "disabled"
        for key in ("server_address_entry", "server_token_entry", "server_test_btn"):
            self.ui[key].configure(state=state)
        self.ui["server_transport_combo"].configure(state="readonly" if remote else "disabled")
        if "notes_label" in self.ui:
            self.ui["notes_label"].configure(
                text=self.tr("notes_remote" if remote else "notes")
            )

    def backend_settings(self):
        if not self.use_remote_backend():
            return {"url": "", "token": "", "transport": config.TRANSPORT_AUTO, "shares": []}
        return {
            "url": self.server_url_var.get().strip(),
            "token": self.server_token_var.get().strip(),
            "transport": self.server_transport_var.get(),
            "shares": self.settings["server"].get("shares", []),
        }

    # -- settings persistence --------------------------------------------
    def _schedule_save_settings(self, *_args):
        if self._save_settings_job is not None:
            try:
                self.root.after_cancel(self._save_settings_job)
            except Exception:
                pass
        self._save_settings_job = self.root.after(500, self._save_settings)

    def _save_settings(self):
        self._save_settings_job = None
        formats = []
        if self.output_pdf_var.get():
            formats.append("PDF")
        if self.output_docx_var.get():
            formats.append("DOCX")

        self.settings["language"] = self.lang()
        self.settings["last_target"] = self.folder_var.get().strip()
        self.settings["options"] = {
            "recursive": bool(self.recursive_var.get()),
            "overwrite": bool(self.overwrite_var.get()),
            "safe_temp": bool(self.use_safe_copy_var.get()),
            "force_one_page": bool(self.force_one_page_var.get()),
            "formats": formats or ["PDF"],
        }
        self.settings["server"].update({
            "url": self.server_url_var.get().strip(),
            "token": self.server_token_var.get().strip(),
            "transport": self.server_transport_var.get(),
        })
        config.save(self.settings)

    def _on_close(self):
        self._save_settings()
        self.root.destroy()

    # -- connection test --------------------------------------------------
    def test_server_connection(self):
        if self.server_test_running:
            return
        url = self.server_url_var.get().strip()
        if not url:
            self.server_status_var.set(self.tr("server_test_failed", detail=self.tr("server_not_configured")))
            return
        self.server_test_running = True
        self.server_status_var.set(self.tr("server_test_running"))
        threading.Thread(
            target=self._server_test_worker,
            args=(self.backend_settings(), self.lang()),
            daemon=True,
        ).start()

    def _server_test_worker(self, server, lang):
        from hwp2pdf.backends.remote_http import RemoteHttpBackend

        try:
            backend = RemoteHttpBackend(server)
            backend.preflight(lang)
            self.log_queue.put(("server_test", (True, backend.capabilities_payload)))
        except Exception as e:
            self.log_queue.put(("server_test", (False, str(e))))

    def open_selected_folder(self):
        target = self.folder_var.get().strip()
        if target and os.path.isfile(target):
            reveal_in_file_manager(str(Path(target).parent))
        elif target and os.path.isdir(target):
            reveal_in_file_manager(target)
        else:
            messagebox.showwarning(APP_TITLE, self.tr("invalid_open_target"))

    def append_log(self, text: str, level: str = "info"):
        tag = level if level in {"error", "warning"} else None
        if tag:
            self.log_text.insert("end", text + "\n", tag)
        else:
            self.log_text.insert("end", text + "\n")
        self.log_text.see("end")

    def request_stop(self):
        self.stop_requested = True
        self.append_log(self.tr("stop_requested"))

    def _poll_log_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()

                if kind == "log":
                    if isinstance(payload, tuple):
                        text, level = payload
                        self.append_log(text, level)
                    else:
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
                        self.progress_label_var.set(self.tr("success_status"))
                        messagebox.showinfo(APP_TITLE, self.tr("success_message"))
                    else:
                        self.progress_label_var.set(
                            self.tr("done_status", success=success, failed=failed, skipped=skipped)
                        )
                        messagebox.showinfo(
                            APP_TITLE,
                            self.tr("done_message", success=success, failed=failed, skipped=skipped, log_csv=log_csv),
                        )

                elif kind == "server_test":
                    ok, detail = payload
                    self.server_test_running = False
                    if ok:
                        self.server_status_var.set(self.tr(
                            "server_test_ok",
                            version=detail.get("version", "?"),
                            hwp=self.tr(
                                "server_hwp_ok" if detail.get("hwp_installed") else "server_hwp_missing"
                            ),
                            queue=detail.get("queue_depth", 0),
                        ))
                    else:
                        self.server_status_var.set(self.tr("server_test_failed", detail=detail))

                elif kind == "error":
                    self.is_running = False
                    self.start_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    self.progress_label_var.set(self.tr("error_status"))
                    self.append_log(self.tr("error_log", message=payload), "error")
                    messagebox.showerror(APP_TITLE, payload)

                elif kind == "update_dl_progress":
                    self.update_status_var.set(self.tr("auto_update_downloading", pct=payload))

                elif kind == "update_dl_error":
                    self.auto_update_btn.state(["!disabled"])
                    self.upgrade_btn.state(["!disabled"])
                    self._apply_cached_update_state()
                    messagebox.showerror(APP_TITLE, self.tr("auto_update_failed", error=payload))

                elif kind == "update_relaunch":
                    self.update_status_var.set(self.tr("auto_update_installing"))
                    self.root.after(1500, self._exit_for_update)

                elif kind == "update_done":
                    if len(payload) == 5:
                        status, latest, release_url, download_url, error_message = payload
                    else:
                        status, latest, release_url, error_message = payload
                        download_url = release_url
                    self.update_check_running = False
                    state = {
                        "checked_at": time.time(),
                        "status": status,
                        "latest": latest,
                        "release_url": release_url,
                        "download_url": download_url,
                        "error": error_message,
                    }
                    save_update_state(state)
                    self._apply_update_state(state)

        except queue.Empty:
            pass

        self.root.after(150, self._poll_log_queue)

    def open_latest_release(self):
        webbrowser.open(self.latest_download_url or self.latest_release_url or GITHUB_RELEASES_PAGE_URL)

    def _show_upgrade_button(self, visible: bool):
        if visible:
            if not self.upgrade_btn.winfo_manager():
                self.upgrade_btn.pack(side="left", padx=(8, 0))
        else:
            self.upgrade_btn.pack_forget()

    def _show_auto_update_button(self, visible: bool):
        if visible:
            if not self.auto_update_btn.winfo_manager():
                self.auto_update_btn.pack(side="left", padx=(8, 0))
        else:
            self.auto_update_btn.pack_forget()

    def _apply_cached_update_state(self):
        state = load_update_state()
        if state:
            self._apply_update_state(state)
        else:
            self.update_status_var.set(self.tr("update_status_current", current=__version__))
            self._show_upgrade_button(False)

    def _apply_update_state(self, state: dict):
        status = state.get("status")
        latest = state.get("latest") or ""
        release_url = state.get("release_url") or GITHUB_RELEASES_PAGE_URL
        download_url = state.get("download_url") or release_url
        self.latest_release_url = release_url
        self.latest_download_url = download_url

        if status == "newer" and latest and parse_version(latest) > parse_version(__version__):
            self.update_status_var.set(self.tr("update_status_available", current=__version__, latest=latest))
            self._show_upgrade_button(True)
            self._show_auto_update_button(
                is_installed_build() and is_setup_asset_url(self.latest_download_url)
            )
        elif status == "no_release":
            self.update_status_var.set(self.tr("update_status_no_release", current=__version__))
            self._show_upgrade_button(False)
            self._show_auto_update_button(False)
        elif status == "error":
            self.update_status_var.set(self.tr("update_status_failed", current=__version__))
            self._show_upgrade_button(False)
            self._show_auto_update_button(False)
        else:
            self.update_status_var.set(self.tr("update_status_current", current=__version__))
            self._show_upgrade_button(False)
            self._show_auto_update_button(False)

    def start_auto_update(self):
        if self.is_running:
            messagebox.showwarning(APP_TITLE, self.tr("auto_update_busy"))
            return
        if not is_setup_asset_url(self.latest_download_url):
            self.open_latest_release()
            return
        if not is_installed_build():
            if messagebox.askyesno(APP_TITLE, self.tr("auto_update_portable")):
                self.open_latest_release()
            return
        state = load_update_state()
        latest = state.get("latest") or ""
        if not messagebox.askyesno(APP_TITLE, self.tr("auto_update_confirm", latest=latest)):
            return
        self.auto_update_btn.state(["disabled"])
        self.upgrade_btn.state(["disabled"])
        threading.Thread(
            target=self._auto_update_worker, args=(self.latest_download_url,), daemon=True
        ).start()

    def _auto_update_worker(self, url):
        try:
            UPDATE_DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
            dest = UPDATE_DOWNLOAD_DIR / url.rsplit("/", 1)[-1]
            try:
                if dest.exists():
                    dest.unlink()
            except OSError:
                pass
            req = urllib.request.Request(
                url, headers={"User-Agent": f"hwp2pdf/{__version__}", "Accept": "application/octet-stream"}
            )
            with urllib.request.urlopen(req, timeout=60) as resp:
                total = int(resp.headers.get("Content-Length") or 0)
                written = 0
                last_pct = -5
                with dest.open("wb") as f:
                    while True:
                        chunk = resp.read(131072)
                        if not chunk:
                            break
                        f.write(chunk)
                        written += len(chunk)
                        pct = int(100 * written / total) if total else 0
                        if pct - last_pct >= 2:
                            self.log_queue.put(("update_dl_progress", pct))
                            last_pct = pct
            self._launch_installer_and_signal_exit(dest)
        except Exception as e:
            self.log_queue.put(("update_dl_error", str(e)))

    def _launch_installer_and_signal_exit(self, setup_path: Path):
        our_exe = sys.executable
        parent_pid = os.getpid()
        ps_path = UPDATE_DOWNLOAD_DIR / "hwp2pdf-update.ps1"
        ready_path = UPDATE_DOWNLOAD_DIR / "hwp2pdf-update.ready"
        helper_log = UPDATE_DOWNLOAD_DIR / "hwp2pdf-update.log"
        install_log = UPDATE_DOWNLOAD_DIR / "hwp2pdf-install.log"
        try:
            ready_path.unlink(missing_ok=True)
        except OSError:
            pass

        script = textwrap.dedent(f"""
            $ErrorActionPreference = 'Stop'
            $helperLog = {self._ps_quote(helper_log)}
            function Write-UpdateLog([string]$Message) {{
                Add-Content -LiteralPath $helperLog -Value "$(Get-Date -Format o) $Message" -Encoding UTF8
            }}

            try {{
                Set-Content -LiteralPath {self._ps_quote(ready_path)} -Value 'ready' -Encoding ASCII
                Write-UpdateLog 'Helper started; waiting for the app to exit.'
                for ($i = 0; $i -lt 120; $i++) {{
                    if (-not (Get-Process -Id {parent_pid} -ErrorAction SilentlyContinue)) {{
                        break
                    }}
                    Start-Sleep -Milliseconds 250
                }}
                if (Get-Process -Id {parent_pid} -ErrorAction SilentlyContinue) {{
                    throw 'The app did not exit within 30 seconds.'
                }}
                Start-Sleep -Seconds 1
                Write-UpdateLog 'Launching installer.'
                $p = Start-Process -FilePath {self._ps_quote(setup_path)} `
                    -ArgumentList '/SP-','/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART','/CLOSEAPPLICATIONS','/HWP2PDFAUTOUPDATE=1',{self._ps_quote(f'/LOG="{install_log}"')} `
                    -Verb RunAs -PassThru
                $p.WaitForExit()
                Write-UpdateLog "Installer exit code: $($p.ExitCode)"
                if ($p.ExitCode -ne 0) {{
                    throw "Installer failed with exit code $($p.ExitCode)."
                }}
                Remove-Item -LiteralPath {self._ps_quote(ready_path)} -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath {self._ps_quote(ps_path)} -Force -ErrorAction SilentlyContinue
            }}
            catch {{
                Write-UpdateLog "Update failed: $($_.Exception.Message)"
                Remove-Item -LiteralPath {self._ps_quote(ready_path)} -Force -ErrorAction SilentlyContinue
                if (Test-Path -LiteralPath {self._ps_quote(our_exe)}) {{
                    Start-Process -FilePath {self._ps_quote(our_exe)}
                }}
            }}
        """).strip()
        ps_path.write_text(script, encoding="utf-8-sig")

        CREATE_NEW_PROCESS_GROUP = 0x00000200
        CREATE_NO_WINDOW = 0x08000000
        powershell_path = (
            Path(os.environ.get("SystemRoot") or r"C:\Windows")
            / "System32"
            / "WindowsPowerShell"
            / "v1.0"
            / "powershell.exe"
        )
        helper = subprocess.Popen(
            [
                str(powershell_path),
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(ps_path),
            ],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW,
            close_fds=True,
        )

        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            if ready_path.exists():
                self.log_queue.put(("update_relaunch", None))
                return
            if helper.poll() is not None:
                raise RuntimeError(f"Updater helper exited early. See {helper_log}")
            time.sleep(0.05)
        helper.terminate()
        raise RuntimeError(f"Updater helper did not start. See {helper_log}")

    @staticmethod
    def _ps_quote(value) -> str:
        return "'" + str(value).replace("'", "''") + "'"

    def _exit_for_update(self):
        try:
            self.root.destroy()
        except Exception:
            pass
        raise SystemExit(0)

    def check_for_updates_if_due(self):
        state = load_update_state()
        if self.update_check_running or not should_check_updates(state):
            return

        self.update_check_running = True
        self.update_status_var.set(self.tr("update_status_checking"))
        self._show_upgrade_button(False)
        threading.Thread(target=self._check_for_updates_worker, daemon=True).start()

    def _check_for_updates_worker(self):
        try:
            release = fetch_latest_release()
            latest = latest_release_version(release)
            release_url = release.get("html_url") or GITHUB_RELEASES_PAGE_URL
            download_url = latest_release_download_url(release) or release_url
            if latest and parse_version(latest) > parse_version(__version__):
                self.log_queue.put(("update_done", ("newer", latest, release_url, download_url, "")))
            else:
                self.log_queue.put(("update_done", ("current", latest or __version__, release_url, download_url, "")))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                self.log_queue.put(("update_done", ("no_release", "", "", "", "")))
            else:
                self.log_queue.put(("update_done", ("error", "", "", "", str(e))))
        except Exception as e:
            self.log_queue.put(("update_done", ("error", "", "", "", str(e))))

    def _confirm_local_engine_ready(self) -> bool:
        ok, detail = ensure_pywin32()
        if not ok:
            messagebox.showerror(APP_TITLE, self.tr("pywin32_missing", detail=detail))
            return False

        hwp_processes = get_hwp_processes()
        if hwp_processes:
            process_detail = ", ".join(f"PID {process['pid']}" for process in hwp_processes)
            answer = messagebox.askyesnocancel(
                APP_TITLE,
                self.tr("hwp_running_prompt", process_detail=process_detail),
            )
            if answer is None:
                return False
            if answer is True:
                hwp_processes = get_hwp_processes()
                if not hwp_processes:
                    self.append_log(self.tr("hwp_closed_already"))
                elif not kill_hwp():
                    hwp_processes = get_hwp_processes()
                    if not hwp_processes:
                        self.append_log(self.tr("hwp_closed_already"))
                    else:
                        messagebox.showerror(
                            APP_TITLE,
                            self.tr("hwp_kill_failed"),
                        )
                        return False
                else:
                    hwp_processes = get_hwp_processes()
                    if hwp_processes:
                        process_detail = ", ".join(f"PID {process['pid']}" for process in hwp_processes)
                        messagebox.showerror(
                            APP_TITLE,
                            self.tr("hwp_kill_failed") + f"\n\n{process_detail}",
                        )
                        return False

        return True

    def start_conversion(self):
        if self.is_running:
            messagebox.showwarning(APP_TITLE, self.tr("already_running"))
            return

        target = self.folder_var.get().strip()
        if not target or not (os.path.isdir(target) or os.path.isfile(target)):
            messagebox.showerror(APP_TITLE, self.tr("invalid_target"))
            return

        if os.path.isfile(target) and Path(target).suffix.lower() not in enabled_extensions():
            messagebox.showerror(APP_TITLE, self.tr("invalid_file"))
            return

        output_formats = self.selected_output_formats()
        if not output_formats:
            messagebox.showerror(APP_TITLE, self.tr("select_output"))
            return

        try:
            backend = create_backend(self.backend_settings(), self.lang())
        except BackendUnavailable as e:
            messagebox.showerror(APP_TITLE, str(e))
            return

        # Only the local COM engine cares about pywin32 and stray Hwp.exe; a
        # remote server reports its own problems through the conversion log.
        if backend.capabilities.manages_hwp_process and not self._confirm_local_engine_ready():
            return

        self.stop_requested = False
        self.is_running = True
        self.log_text.delete("1.0", "end")
        lang = self.lang()
        self.append_log(translate(lang, "starting_conversion"))
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.progress["value"] = 0
        self.progress_label_var.set(translate(lang, "scanning"))

        self.worker = threading.Thread(
            target=self._run_conversion,
            args=(
                target,
                self.recursive_var.get(),
                self.overwrite_var.get(),
                self.use_safe_copy_var.get(),
                self.force_one_page_var.get(),
                output_formats,
                lang,
                self.backend_settings(),
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

    collect_files = staticmethod(collect_files)

    def _run_conversion(
        self,
        target: str,
        recursive: bool,
        overwrite: bool,
        use_safe_copy: bool,
        force_one_page: bool,
        output_formats,
        lang: str,
        backend_settings=None,
    ):
        # Tk variables may only be read on the main thread, so the caller
        # resolves the backend settings before starting this worker.
        if backend_settings is None:
            backend_settings = getattr(self, "backend_settings", None)
            if callable(backend_settings):
                backend_settings = None
        try:
            backend = create_backend(backend_settings, lang)
        except BackendUnavailable as e:
            self.log_queue.put(("error", str(e)))
            return

        run_batch(
            self.log_queue,
            backend,
            target=target,
            recursive=recursive,
            overwrite=overwrite,
            use_safe_copy=use_safe_copy,
            force_one_page=force_one_page,
            output_formats=output_formats,
            lang=lang,
            is_stopped=lambda: self.stop_requested,
            file_collector=self.collect_files,
        )


def main():
    try:
        root = TkinterDnD.Tk()
    except RuntimeError:
        root = tk.Tk()
    style = ttk.Style(root)
    available = set(style.theme_names())
    for theme in ("aqua", "vista", "clam"):
        if theme in available:
            try:
                style.theme_use(theme)
                break
            except Exception:
                continue
    ConverterApp(root)
    root.mainloop()
