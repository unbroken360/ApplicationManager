# ApplicationManager

**ApplicationManager** is a powerful PowerShell-based automation tool designed to streamline the creation and management of software applications in **Microsoft ConfigMgr (SCCM)** and **Microsoft Intune**.  
It supports packaging with **PSAppDeployToolkit (PSADT)** as well as traditional `cmd` or `ps1` based setups.

---

## 🔧 Key Features

### ✅ Microsoft ConfigMgr (SCCM)

- Create ConfigMgr Applications automatically
- Generate Device/User Collections
- Create and manage Deployments
- Fully customizable naming and folder structure
- Supports PSADT, CMD, or PS1-based setups

### ✅ Microsoft Intune

- Create Intune Win32 Applications packages
- Automatically assign to Azure AD groups
- Handle Required/Available/Uninstall assignments
- Manage Intune groups if missing
- Full PSADT integration

### ✅ WinGet Integration

- Search for WinGet packages
- Automatically download and wrap installers using PSAppDeployToolkit
- Build and deploy ready-to-go apps

---

## 🚀 Quick Start

```powershell
.\ApplicationManager.ps1
```
