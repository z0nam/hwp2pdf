# Security Policy

## Supported Versions

Security fixes are handled for the latest published GitHub Release.

## Reporting A Vulnerability

Please do not open a public issue for a vulnerability involving arbitrary code execution, unsafe file handling, credential exposure, or malicious document behavior.

Report security concerns privately to:

```text
namun.cho@gmail.com
```

Please include:

- hwp2pdf version
- Windows version
- Hancom Office version
- Steps to reproduce
- Whether the issue affects GUI, CLI, installer, or all of them

## Scope

hwp2pdf automates Hancom Office through COM and opens user-selected HWP/HWPX files. Treat untrusted documents carefully, and only run documents from sources you trust.

Release binaries may show Windows SmartScreen warnings until they are code signed.
