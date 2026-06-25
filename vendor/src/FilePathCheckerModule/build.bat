@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem Build FilePathCheckerModule.dll for both x86 and x64 using MSVC Build Tools.
rem Requires: Visual Studio 2022 Build Tools (or full VS) with C++ workload installed.
rem Output: ..\..\x86\FilePathCheckerModule.dll, ..\..\x64\FilePathCheckerModule.dll

set "SRC_DIR=%~dp0"
set "ROOT_DIR=%SRC_DIR%..\..\"
set "X86_OUT=%ROOT_DIR%x86"
set "X64_OUT=%ROOT_DIR%x64"
set "TMP_DIR=%SRC_DIR%build"

set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" set "VSWHERE=%ProgramFiles%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" (
    echo [ERROR] vswhere.exe not found. Install Visual Studio 2022 Build Tools first.
    exit /b 1
)
for /f "usebackq tokens=*" %%I in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do (
    set "VS_INSTALL=%%I"
)
if not defined VS_INSTALL (
    echo [ERROR] Visual Studio with C++ tools not found.
    exit /b 1
)
set "VCVARSALL=%VS_INSTALL%\VC\Auxiliary\Build\vcvarsall.bat"
if not exist "%VCVARSALL%" (
    echo [ERROR] vcvarsall.bat not found at "%VCVARSALL%".
    exit /b 1
)

if not exist "%X86_OUT%" mkdir "%X86_OUT%"
if not exist "%X64_OUT%" mkdir "%X64_OUT%"
if exist "%TMP_DIR%" rmdir /S /Q "%TMP_DIR%"
mkdir "%TMP_DIR%"

call :build x86 "%X86_OUT%\FilePathCheckerModule.dll" || exit /b 1
call :build x64 "%X64_OUT%\FilePathCheckerModule.dll" || exit /b 1

rmdir /S /Q "%TMP_DIR%"
echo.
echo [OK] x86 ^-^> "%X86_OUT%\FilePathCheckerModule.dll"
echo [OK] x64 ^-^> "%X64_OUT%\FilePathCheckerModule.dll"
exit /b 0

:build
set "ARCH=%~1"
set "OUT=%~2"
set "WORK=%TMP_DIR%\%ARCH%"
mkdir "%WORK%"
echo.
echo === Building %ARCH% ===
call "%VCVARSALL%" %ARCH% >nul || (
    echo [ERROR] vcvarsall %ARCH% failed.
    exit /b 1
)
pushd "%WORK%"
cl /nologo /LD /O2 /MT /W4 /EHsc /DUNICODE /D_UNICODE ^
    "%SRC_DIR%FilePathCheckerModule.cpp" ^
    /Fe:FilePathCheckerModule.dll ^
    /link /DEF:"%SRC_DIR%FilePathCheckerModule.def" kernel32.lib user32.lib
set "RC=%ERRORLEVEL%"
popd
if not "%RC%"=="0" (
    echo [ERROR] cl %ARCH% failed with %RC%.
    exit /b %RC%
)
copy /Y "%WORK%\FilePathCheckerModule.dll" "%OUT%" >nul
exit /b 0
