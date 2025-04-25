# New-BootableIsoFile

Create bootable ISO files from a source directory using PowerShell.  
This module provides a more controlled and transparent method for generating ISO files, supporting both BIOS and UEFI boot options.

[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/New-BootableIsoFile.svg?style=flat-square)](https://www.powershellgallery.com/packages/New-BootableIsoFile)

---

## 📦 Installation

You can install this module directly from the PowerShell Gallery:

```powershell
Install-Module -Name New-BootableIsoFile -Scope CurrentUser
```

---

## 🚀 Usage

```powershell
# Minimal
New-BootableIsoFile .\Win11_24H2_English_x64

# Full control
New-BootableIsoFile -SourceDir "C:\Win11_24H2_English_x64" -IsoPath "D:\ISO\Win11_24H2_English_x64.iso" -IsoLabel "ESD-ISO"
```

### Parameters:
- `-SourceDir` (**Required**)  
  The path to the directory containing files you want to include in the ISO.

- `-IsoPath` (Optional)  
  The full output path for the ISO file. Defaults to `<SourceDir>.iso` in the same directory.

- `-IsoLabel` (Optional)  
  The label for the ISO volume. Defaults to the ISO filename.

---

## ⚠️ Notes

- The script relies on COM objects like `IMAPI2FS.MsftFileSystemImage` and works on **Windows only**.
- Relative paths are automatically resolved to full paths.

---