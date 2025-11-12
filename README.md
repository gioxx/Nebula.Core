# Nebula.Core

Work in progress.

> [!WARNING]  
> This module is not yet available on PowerShell Gallery.

![PowerShell Gallery](https://img.shields.io/powershellgallery/v/Nebula.Core?label=PowerShell%20Gallery)
![Downloads](https://img.shields.io/powershellgallery/dt/Nebula.Core?color=blue)

---

## 📦 Installation

Install from PowerShell Gallery:

```powershell
Install-Module -Name Nebula.Core -Scope CurrentUser
```

---

## 🧽 How to clean up old module versions (optional)

When updating from previous versions, old files (such as unused `.psm1`, `.yml`, or `LICENSE` files) are not automatically deleted.  
If you want a completely clean setup, you can remove all previous versions manually:

```powershell
# Remove all installed versions of the module
Uninstall-Module -Name Nebula.Core -AllVersions -Force

# Reinstall the latest clean version
Install-Module -Name Nebula.Core -Scope CurrentUser -Force
```

ℹ️ This is entirely optional — PowerShell always uses the most recent version installed.

---

## 📄 License

All scripts in this repository are licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## 🔧 Development

This module is part of the [Nebula](https://github.com/gioxx?tab=repositories&q=Nebula) PowerShell tools family.

Feel free to fork, improve and submit pull requests.

---

## 📬 Feedback and Contributions

Feedback, suggestions, and pull requests are welcome!  
Feel free to [open an issue](https://github.com/gioxx/Nebula.Core/issues) or contribute directly.