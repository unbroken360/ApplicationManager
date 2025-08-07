# 📦 Changelog

All notable changes to **ApplicationManager** will be documented in this file.

---

## [0.11.0] - 2025-08-07

### ✨ Added

- Detection script template system: uses `Files\DetectionScriptWingetTemplate.ps1` with `<APPID>` and `<APPVERSION>` placeholders.
- Version-aware detection: compares installed version using `[Version]` and `-ge` operator.
- Automatic generation of detection script files per app (`SupportFiles\Detect-<AppID>.ps1`).
- Support for `_install.bat` and `_uninstall.bat`: automatically used as install/uninstall commands if present in the app folder.
- Smart `ServiceUI` wrapping: automatically applied **only if** install/uninstall is **not** a `.bat` file, prioritizing `ServiceUI_x64.exe`.

### 🛠 Fixed

- Resolved `Rename-Item` exceptions when `.intunewin` file already exists.
- Fixed issue where ServiceUI command line was broken due to multiple matches (e.g., both x64 and x86).

### 🧼 Improved

- PSADT template copying logic: now safely skips existing items and handles subfolders cleanly.
- Improved detection command generation and script validation.
- Clearer log messages for detection, copy, rename, and error handling.

## [0.10.0] - 2025-03-28

### ✨ Added

- Full compatibility with PowerShell App Deployment Toolkit (PSADT) v4.x
- Winget PowerShell Module integration to support all Winget apps
- Automatic MSI detection method when MSI is found
- Extracted core app creation logic into reusable function: `Create-ApplicationObjects`
- Parameter-based behavior control (`$CreateInConfigMgr`, `$CreateInIntune`, etc.)
- Reusable logic used across `ButtonCreateClick` and `ButtonCreateWinGet`

### 🧼 Improved

- Removed UI dependencies inside logic functions
- Improved maintainability and testability of the code

---

## [0.9.6]

- Updated drive connect behavior: connect to SiteCode drive only if ConfigMgr app creation is selected

## [0.9.5]

- Changed Azure authentication to use Application Auth  
  → [Reference](https://learn.microsoft.com/en-us/samples/microsoftgraph/powershell-intune-samples/important/)
- Enhanced Azure authentication security

## [0.9.4]

- Fixed: Uninstall Program not set
- Support for opening log files from UNC paths

## [0.9.3]

- Fixed: Some variables misinterpreted as `System.String` in global state
- Fixed: Uninstall deployment issue

## [0.9.2]

- Fixed: `RunInstallAs32Bit` interpreted as `System.String`

## [0.9.1]

- Rebuilt Azure authentication
- Performed rebranding

## [0.8]

- Integrated Winget App creation with PSADT
- Modified Intune app creation function (based on updated Cmdlets)
- Added SourcePatch to Intune description
- Reworked ProgressBar behavior (added to config save action)
- Split app creation functions for better future reuse
- Updated to use `Connect-MSIntuneGraph -TenantID $TenantName` additionally to `-TenantName`

## [0.7]

- Bugfix: `Connect-MSIntuneGraph` now uses both `-TenantID` and `-TenantName`

## [0.6]

- Added `cscript` to command line for `.vbs` scripts
- Bugfix for handling subfolders
- Added `ServiceUI` to command line if present and `Allow User to interact` is enabled

## [0.5]

- Applied AAD group prefix to Pilot Group
- Added error handler for `TextBoxIntuneOutputFolder`

## [0.4]

- Install AAD Modules automatically
- Set `SchUseStrongCrypto` registry key if NuGet is not registered
- Add AAD group prefix

## [0.3]

- Added `IntuneWinAppUtil.exe` to `Files` folder and as a parameter
- Bugfix: `CleanUpIntuneOutputFolder`
- Added ToolTips for UI

## [0.2]

- Changed Uninstall Collection naming
- Added browse button for application selection

---
