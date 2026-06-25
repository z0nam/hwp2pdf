// FilePathCheckerModule.cpp
//
// Minimal stand-alone Hancom HWP file-access security module.
//
// Hancom's HWP automation calls RegisterModule("FilePathCheckDLL", "<reg value name>")
// which loads the DLL whose path is stored in
//   HKCU\Software\HNC\HwpAutomation\Modules\<reg value name>
// and invokes the DLL's exported entry point IsAccessiblePath() to decide whether HWP
// may open / save a given file from automation. Without a registered handler HWP shows
// the "<...> 파일에 접근하려는 시도가 있습니다. 접근을 허용하시겠습니까?" dialog
// and a headless conversion hangs forever.
//
// This module is intentionally permissive: hwp2pdf only converts files the user already
// chose in its own UI / CLI, so silently approving every access request is what we want.
// No prompts, no logging, no I/O — just return TRUE.
//
// SPDX-License-Identifier: MIT
// Copyright (c) 2026 Namun Cho

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

extern "C" {

BOOL APIENTRY DllMain(HINSTANCE hInstance, DWORD reason, LPVOID reserved) {
    (void)hInstance;
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) {
        DisableThreadLibraryCalls(hInstance);
    }
    return TRUE;
}

// HWP's RegisterModule contract: returning TRUE means "this file path is OK to access".
// HWND parent, LONG callerId, file path, optional caller / site info.
__declspec(dllexport) BOOL __stdcall IsAccessiblePath(
        HWND hWnd, LONG nID, LPCTSTR szFilePath, LPCTSTR szSiteInfo) {
    (void)hWnd;
    (void)nID;
    (void)szFilePath;
    (void)szSiteInfo;
    return TRUE;
}

}  // extern "C"
