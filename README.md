# psquickaccess

A small PowerShell script that pins folders to **Quick Access** in Windows File Explorer.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1 or PowerShell 7+

## Usage

Pass folder paths as arguments:

```powershell
.\pinqa.ps1 "C:\Projects" "D:\Media"
```

Or run it with no arguments to enter paths interactively (one per line; type `done` or leave the line blank to finish):

```powershell
.\pinqa.ps1
```

If script execution is blocked on your machine, run it with a one-off bypass:

```powershell
powershell -ExecutionPolicy Bypass -File .\pinqa.ps1 "C:\Projects"
```

## How it works

The script uses the `Shell.Application` COM object and invokes the `pintohome` verb on each folder. `pintohome` is the canonical (language-independent) name of File Explorer's "Pin to Quick access" context-menu action, so the script works on any Windows display language.

## Notes

- Quick Access only pins **folders**. File paths are skipped with a warning — pin the file's containing folder instead.
- Paths that don't exist are reported and skipped; the remaining paths are still processed.
- To unpin a folder, right-click it under Quick Access in File Explorer and choose **Unpin from Quick access**.
