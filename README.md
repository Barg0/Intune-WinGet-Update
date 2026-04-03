# 🪟 Intune WinGet Update

Scheduled **Microsoft Intune** remediations that use **Windows Package Manager (WinGet)** to find outdated apps and upgrade them.

**Two PowerShell scripts:** detection decides *if* something needs doing; remediation *does* the upgrades.

---

## 📑 Table of contents

1. [What you get](#-what-you-get)  
2. [Win32 apps via Intune-WinGet](#-win32-apps-via-intune-winget)  
3. [WinGet in SYSTEM context (platform script)](#-winget-in-system-context-platform-script)  
4. [Prerequisites](#-prerequisites)  
5. [List modes (Blacklist / Whitelist)](#-list-modes-blacklist--whitelist)  
6. [Deploy in Intune (step by step)](#-deploy-in-intune-step-by-step)  
7. [Log files & Collect diagnostics](#-log-files--collect-diagnostics)  
8. [Configuration reference](#-configuration-reference)  
9. [How remediation fallbacks work](#-how-remediation-fallbacks-work)  
10. [Exit codes & summary lines](#-exit-codes--summary-lines)  
11. [Log examples](#-log-examples)  
12. [Troubleshooting](#-troubleshooting)  
13. [References & credits](#-references--credits)

---

## ✨ What you get

- 🔍 **Detection** — Queries WinGet for upgrades, applies your list rules, exits **non-compliant** when updates remain.  
- 🔧 **Remediation** — Refreshes sources, walks each pending app through an **upgrade ladder**, optional **locale** retry, optional **`install` fallback**, optional **`--uninstall-previous`**, then logs a **one-line portal summary**.  
- 🎯 **Blacklist or Whitelist** — Wildcards on package IDs (`Mozilla.Firefox*`, …).  
- ⏳ **Busy installer handling** — Waits and retries when WinGet reports *another installation in progress*.  
- 📝 **Logs** — `%ProgramData%\IntuneLogs\Scripts\WinGet-Update\` → `detection.log` / `remediation.log`

---

## 📦 Win32 apps via Intune-WinGet

If you want to **deploy WinGet applications as Intune Win32 apps** (`.intunewin` packages with **install / uninstall / detection** scripts, use this project instead of (or alongside) remediations:

👉 **[github.com/Barg0/Intune-WinGet](https://github.com/Barg0/Intune-WinGet)** — CSV-driven pipeline: **`package.ps1`** builds packages from **`apps.csv`**; **`deploy.ps1`** publishes them to Intune via **Microsoft Graph**. Install templates there share the same kind of **scope / busy-install** retry patterns as this repo.

**This repository** (**Intune-WinGet-Update**) only covers **scheduled detection + remediation** to **upgrade** apps that are already installed and visible to WinGet.

---

## ⚙️ WinGet in SYSTEM context (platform script)

🤖 **Detection and remediation here run as SYSTEM.**

On some devices, WinGet can **fail in SYSTEM context** because required **UWP dependencies** — in particular **`Microsoft.VCLibs`** and **`Microsoft.UI.Xaml`** — are not available to the **SYSTEM** account the way they are to an interactive user. Until those paths are registered for SYSTEM, `winget.exe` may not work reliably for platform scripts or Win32 deployments.

1. Deploy **Winget-SystemContext** as an **Intune platform script** (Settings Catalog → **Scripts**, or your tenant’s equivalent) **before** or **alongside** rolling out WinGet-heavy policies.  
2. The script **registers the UWP dependency paths** so WinGet can run correctly **as SYSTEM**.  
3. **Without it**, this remediation package or Intune-WinGet **install** scripts may fail on PCs where those dependencies have **never** been made visible to SYSTEM.

**Where to get the script:** [Barg0/Intune-Platform-Scripts — **Winget-SystemContext**](https://github.com/Barg0/Intune-Platform-Scripts/tree/main/Winget-SystemContext)

> 💡 **Scope:** This applies to **SYSTEM**-executed WinGet.

---

## ✅ Prerequisites

| Requirement | Notes |
|-------------|--------|
| 🖥️ Windows 10 / 11 | WinGet via **App Installer** (`Microsoft.DesktopAppInstaller`). |
| ☁️ Intune + Management Extension | Scripts are designed for **SYSTEM** (64-bit PowerShell). |
| 🔗 Matching config | **`$listMode`** and list arrays must be **identical** in **both** scripts. |
| 🔧 WinGet healthy as SYSTEM | See [**WinGet in SYSTEM context**](#-winget-in-system-context-platform-script) — deploy **Winget-SystemContext**. |

---

## 🎚️ List modes (Blacklist / Whitelist)

| Mode | Variable | Behavior |
|------|----------|----------|
| 🚫 **Blacklist** | `$listMode = 'Blacklist'` | Upgrade **every** WinGet app that has an update **except** IDs matching `$blacklistApps`. |
| ✅ **Whitelist** | `$listMode = 'Whitelist'` | Upgrade **only** IDs matching `$whitelistApps`. |

**Wildcards** use PowerShell `-like` rules: `*` and `?` work (e.g. `Microsoft.DotNet*`).

**Package IDs** come from `winget list` / `winget search` (**Id** column).

---

## ☁️ Deploy in Intune (step by step)

### 1️⃣ Prepare the files

1. Download this repository.  
2. Open **`detection.ps1`** and **`remediation.ps1`** in an editor.  
3. Set **[Configuration reference](#-configuration-reference)** values.  
4. **Critical:** Copy the **same** `$listMode`, `$blacklistApps`, and `$whitelistApps` into **both** files.

### 2️⃣ Upload the scripts

1. Sign in to the [Microsoft Intune admin center](https://aka.ms/intune).  
2. Go to **Devices** → **Remediations** (menu names can vary slightly by tenant).  
3. **Create script package**.  
4. Name it clearly (e.g. `WinGet Update - Weekly`).  
5. **Detection script:** upload **`detection.ps1`**.  
6. **Remediation script:** upload **`remediation.ps1`**.  
7. Recommended settings:  
   - **Run using logged-on credentials:** **No** (run as **SYSTEM**).  
   - **Run in 64-bit PowerShell:** **Yes**.  
   - **Enforce script signature check:** **No** (unless you sign the scripts).  
8. **Assign** to device groups.  
9. **Schedule** (e.g. daily / weekly).

Intune runs **detection** first. Exit **1** = **non-compliant** → **remediation** runs. Exit **0** = compliant → remediation is skipped for that evaluation.

### 3️⃣ Verify

- **Intune:** package → **Device status** / run output.  
- **Device:** log folder and **Collect diagnostics** setup → [Log files & Collect diagnostics](#-log-files--collect-diagnostics).  
- **Console:** last line should be the **portal summary** (see below).

---

## 📂 Log files & Collect diagnostics

Scripts write UTF-8 logs here (default folder uses **`$scriptName = 'WinGet-Update'`** in both scripts):

`C:\ProgramData\IntuneLogs\Scripts\WinGet-Update\`

| File | Written by |
|------|------------|
| 📄 **`detection.log`** | Detection script |
| 📄 **`remediation.log`** | Remediation script |

If you change **`$scriptName`**, the folder becomes `...\Scripts\<YourName>\` instead—keep the **same** name in **both** scripts.

📲 **Collect diagnostics in Intune** does not always include arbitrary folders like `ProgramData\IntuneLogs\...` in the diagnostic package. To pull these logs remotely, deploy this **platform script** on devices:

👉 **[Diagnostics - Custom Log File Directory](https://github.com/Barg0/Intune-Platform-Scripts/tree/main/Diagnostics%20-%20Custom%20Log%20File%20Directory)**

That script registers your custom log path with the Intune Management Extension so **Collect diagnostics** can bundle those `.log` files.

**Path (default):**

```text
C:\ProgramData\IntuneLogs\Scripts\WinGet-Update\
├── detection.log
└── remediation.log
```

---

## 🎛️ Configuration reference

### 🔗 In **both** `detection.ps1` and `remediation.ps1` (must match)

| Setting | Purpose |
|---------|---------|
| `$listMode` | `'Blacklist'` or `'Whitelist'`. |
| `$blacklistApps` | String array of IDs/patterns to **exclude** in Blacklist mode. |
| `$whitelistApps` | String array of IDs/patterns to **include** in Whitelist mode. |
| `$log`, `$logDebug`, `$logGet`, `$logRun`, `$enableLogFile` | Control console/file logging verbosity. |
| `$scriptName` | Log folder name under `...\Scripts\<name>\`. Keep **same** in both scripts if you change it. |

### 🔧 Only in **`remediation.ps1`**

| Setting | Default (typical) | Purpose |
|---------|-------------------|---------|
| `$wingetLocaleWorkaround` | `'en-US'` | After a failed **first** locale pass ending in **“No packages found”** (`0x8A150014`), repeats the **full upgrade ladder** with `--locale`. Use `''` to disable. |
| `$wingetAllowInstallFallback` | `$false` | If **`$true`**, after upgrades still fail with **5212** and **`AvailableVersion`** is known, tries **`winget install --id … --version … --force`** across machine → default → user scopes. |
| `$wingetAllowUninstallPrevious` | `$false` | If **`$true`**, last resort: **`winget upgrade … --uninstall-previous`** (high impact if reinstall fails). |

---

## 🔁 How remediation fallbacks work

### 1. ⬆️ Upgrade ladder (every attempt wrapped for “installer busy”)

For each **locale pass** (normal, then optional `--locale` workaround):

1. **`winget upgrade`** with **`--scope machine`** (`--source winget`, `-e`, `--force`, agreements, silent).  
2. If WinGet indicates **wrong scope / no package / no applicable** (and not a hard stop), retry **without** `--scope`.  
3. If still needed, retry with **`--scope user`**.  

Success is accepted only if WinGet’s category is **Success** *and* the output does not claim a bogus “no applicable upgrade” (guards false **exit 0**).

### 2. 🩹 Per-invocation repairs (inside the ladder)

- **Hash mismatch** → `winget source update --name winget`, then **same** scope command again.  
- **Transient download errors** → sleep **`$wingetDownloadRetryWaitSeconds`**, then **one** repeat of the same scope command.

### 3. 🗂️ Source repair (once per app, if needed)

If WinGet reports **corrupt / missing source** (mapped as **RetrySourceRepair**), the script runs **`winget source reset --force`** and **`winget source update`**, then runs the **whole machine → default → user ladder again** (same locale pass).

### 4. 🌐 Locale pass (optional)

If **`$wingetLocaleWorkaround`** is non-empty and the **first** pass ends with **5212**, the **entire ladder** (including source repair logic) runs again with **`--locale`**.

### 5. 📥 Install fallback (optional, **`$wingetAllowInstallFallback`**)

If still **5212** and the detection list supplied **`AvailableVersion`**, the script can run **`winget install --id … --version … --force`** in **machine → default → user** order. This uses the **catalog** instead of relying on upgrade’s installed-app matching (helps with known WinGet edge cases).

### 6. ♻️ Uninstall-previous (optional, **`$wingetAllowUninstallPrevious`**)

If enabled, runs **`winget upgrade`** with **`--uninstall-previous`** (and version) across scopes. **Uninstalls the old build before installing the new one** — test thoroughly before enabling.

### 🔍 Detection script behavior (no “fallbacks” like remediation)

- Three **`winget upgrade --source winget`** list passes: **no scope**, **`--scope user`**, **`--scope machine`**, merged by **App ID**.  
- Rows with **Unknown** version are skipped.  
- **No** install / uninstall-previous / override logic.

---

## 🚦 Exit codes & summary lines

### 🔎 Detection

| Exit | Meaning |
|------|---------|
| **0** | Compliant: no pending updates after filters, **or** WinGet unavailable (intentionally **no** remediation trigger). |
| **1** | Non-compliant: at least one filtered update pending. |

**Last console line (examples):**

- `Available: (none)`  
- `Available: Google.Chrome, 7zip.7zip`  
- `Available: (none) | Winget unavailable`  
- `Available: (unknown) | see detection.log` (unhandled error)

### 🛠️ Remediation

| Exit | Meaning |
|------|---------|
| **0** | No **hard** failures (some apps may be **deferred**). |
| **1** | At least one app **failed** after all enabled steps. |

**Last console line (examples):**

- `Updated: 7zip.7zip, Google.Chrome | Failed: (none)`  
- `Updated: 7zip.7zip | Failed: Vendor.App | Deferred: BigInstaller`  
- `Updated: (none) | Failed: (none) | Winget unavailable`

---

## 📋 Log examples

Timestamps and versions are illustrative.

### ✅ Detection — compliant (no updates)

```text
2026-03-29 09:00:00 [ Start   ] ==================== Start ====================
2026-03-29 09:00:00 [ Info    ] Host DESKTOP01 | SYSTEM | WinGet-Update
2026-03-29 09:00:02 [ Get     ] Upgrades: 0
2026-03-29 09:00:02 [ Success ] No upgrades
2026-03-29 09:00:02 [ Info    ] Runtime 00:00:02.10
2026-03-29 09:00:02 [ Info    ] Exit 0
2026-03-29 09:00:02 [ End     ] ==================== End ====================
Available: (none)
```

### ⚠️ Detection — non-compliant

```text
2026-03-29 09:05:00 [ Get     ] Upgrades: 4
2026-03-29 09:05:01 [ Get     ] Filtered: 2
2026-03-29 09:05:01 [ Info    ] Available:
2026-03-29 09:05:01 [ Info    ]   Google.Chrome 131.0.0 -> 132.0.0
2026-03-29 09:05:01 [ Info    ]   7zip.7zip 24.08 -> 24.09 [machine]
2026-03-29 09:05:01 [ Success ] Detect: 2 non-compliant
2026-03-29 09:05:01 [ Info    ] Runtime 00:00:03.05
2026-03-29 09:05:01 [ Info    ] Exit 1
2026-03-29 09:05:01 [ End     ] ==================== End ====================
Available: Google.Chrome, 7zip.7zip
```

### ✅ Remediation — success after scope retry

```text
2026-03-29 09:10:00 [ Run     ] WinGet source update
2026-03-29 09:10:02 [ Run     ] [1/2] Google.Chrome 131.0.0 -> 132.0.0
2026-03-29 09:10:05 [ Info    ] Retry: no --scope
2026-03-29 09:10:40 [ Success ] Google.Chrome (default scope)
2026-03-29 09:10:43 [ Run     ] [2/2] 7zip.7zip 24.08 -> 24.09
2026-03-29 09:10:55 [ Success ] 7zip.7zip (machine)
2026-03-29 09:10:55 [ Success ] OK count: 2
2026-03-29 09:10:55 [ Success ] Done (no hard failures)
2026-03-29 09:10:55 [ Info    ] Runtime 00:00:55.20
2026-03-29 09:10:55 [ Info    ] Exit 0
2026-03-29 09:10:55 [ End     ] ==================== End ====================
Updated: Google.Chrome, 7zip.7zip | Failed: (none)
```

### ⏸️ Remediation — defer (installer busy)

```text
2026-03-29 10:00:00 [ Run     ] [1/1] Contoso.App 1.0 -> 1.1
2026-03-29 10:00:05 [ Info    ] Install busy; wait 120s (1/15)
...
2026-03-29 10:25:00 [ Info    ] Defer Contoso.App: install busy (max waits)
2026-03-29 10:25:00 [ Success ] OK count: 0
2026-03-29 10:25:00 [ Info    ] Deferred (1): Contoso.App
2026-03-29 10:25:00 [ Info    ] Exit 0
2026-03-29 10:25:00 [ End     ] ==================== End ====================
Updated: (none) | Failed: (none) | Deferred: Contoso.App
```

*(Enable **`$logDebug = $true`** to see full `winget` command lines in the log.)*

---

## 🆘 Troubleshooting

| Symptom | Check |
|---------|--------|
| 🔕 Detection never triggers | Same lists/mode in **both** scripts; correct **package IDs**. |
| ❌ Remediation exit **1** | `remediation.log` → **Fail** lines; compare codes with [WinGet return codes](https://github.com/microsoft/winget-cli/blob/master/doc/windows/package-manager/winget/returnCodes.md). |
| 🐛 One package always fails | Blacklist it temporarily; test **`install` fallback** / **`uninstall-previous`** only with care. |
| 📭 WinGet missing | Deploy / repair **App Installer**; remediation exits **1** if WinGet is required and missing. |
| 🤖 WinGet fails only as **SYSTEM** (works for users) | Deploy **[Winget-SystemContext](https://github.com/Barg0/Intune-Platform-Scripts/tree/main/Winget-SystemContext)**; see [WinGet in SYSTEM context](#-winget-in-system-context-platform-script). |
| 📦 **Collect diagnostics** missing these logs | Deploy **[Diagnostics - Custom Log File Directory](https://github.com/Barg0/Intune-Platform-Scripts/tree/main/Diagnostics%20-%20Custom%20Log%20File%20Directory)**; see [Log files & Collect diagnostics](#-log-files--collect-diagnostics). |

---

## 📚 References & credits

| Link | Notes |
|------|------|
| [**Intune Vita Doctrina** (YouTube)](https://www.youtube.com/@IntuneVitaDoctrina) | Introduced me to the concept of Intune working with WinGet |
| [Winget-AutoUpdate](https://github.com/Romanitho/Winget-AutoUpdate) | Early inspiration for WinGet automation ideas |
| [WinGet](https://learn.microsoft.com/windows/package-manager/winget/) | Product documentation |
| [winget upgrade](https://learn.microsoft.com/windows/package-manager/winget/upgrade) | Upgrade command |
| [WinGet return codes](https://github.com/microsoft/winget-cli/blob/master/doc/windows/package-manager/winget/returnCodes.md) | HRESULT reference |
| [Intune remediations](https://learn.microsoft.com/mem/intune/fundamentals/powershell-scripts-remediation) | Policy type |
