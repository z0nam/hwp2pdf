# vendor/

Bundled binaries shipped with hwp2pdf.

## FilePathCheckerModule.dll

Tiny stand-alone DLL that satisfies Hancom HWP's automation file-access security
contract. When registered, HWP automation stops popping the
"…파일에 접근하려는 시도가 있습니다. 접근을 허용하시겠습니까?" dialog and headless
conversions can finish without human input.

| Path                                | Architecture |
| ----------------------------------- | ------------ |
| `x86/FilePathCheckerModule.dll`     | 32-bit HWP   |
| `x64/FilePathCheckerModule.dll`     | 64-bit HWP   |

Source and license: `src/FilePathCheckerModule/` — MIT, written from scratch for
hwp2pdf. Only the Hancom-defined export signature (`IsAccessiblePath`) is reused,
which is an ABI contract and not copyrightable. No Hancom-supplied code is
included or redistributed.

### Build

Requires Visual Studio 2022 Build Tools with the **C++ x86/x64 build tools**
component. Then:

```cmd
cd vendor\src\FilePathCheckerModule
build.bat
```

This drops freshly built `FilePathCheckerModule.dll` into `vendor\x86\` and
`vendor\x64\`. Commit the resulting binaries — `app.py` and the installer expect
them to be present.
