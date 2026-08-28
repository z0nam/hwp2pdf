"""Format, extension and Hancom automation constants.

Importable without tkinter or pywin32 so the conversion server and the remote
client can share them.
"""

from hwp2pdf import paths

from hwp2pdf.version import __version__

APP_NAME = "HWP/HWPX -> PDF/DOCX Converter"


DEFAULT_OPEN_OPTION = "forceopen:true;versionwarning:false;"


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


HWP_SECURITY_REG_KEY = r"Software\HNC\HwpAutomation\Modules"


HWP_SECURITY_REG_VALUE = "FilePathCheckerModule"


HWP_SECURITY_DLL_NAME = "FilePathCheckerModule.dll"


MESSAGE_BOX_AUTO_CONFIRM = 0x10


HANCOM_BLOCKING_DIALOG_MESSAGES = (
    "복합 파일을 현재 구현하기에 너무 큽니다.",
    "알 수 없는 형식의 파일입니다.",
)


HANCOM_DIALOG_CONFIRM_BUTTONS = ("확인", "OK", "예", "Yes", "계속", "Continue", "닫기", "Close")


HWP_FILEHEADER_STREAM = "FileHeader"


HWP_FILE_SIGNATURE = b"HWP Document File"


HWP_FLAG_COMPRESSED = 1 << 0


HWP_FLAG_PASSWORD_PROTECTED = 1 << 1


HWP_FLAG_DISTRIBUTION_DOCUMENT = 1 << 2


def enabled_extensions():
    return BASE_EXTENSIONS


def output_extension(output_format: str):
    return OUTPUT_FORMATS[output_format]


APP_TITLE = f"{APP_NAME} v{__version__}"

# Both were ``%LOCALAPPDATA%``/``C:\temp`` literals before the macOS port; on
# Windows ``paths`` resolves to exactly the same locations.
TEMP_WORKDIR = paths.temp_workdir()
HWP_SECURITY_INSTALL_DIR = paths.security_install_dir()
