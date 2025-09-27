# Windows Installer Cache Recovery Automation (with FixMissingMSI)

A set of PowerShell scripts to automate the recovery of missing files in the **Windows Installer cache (`C:\Windows\Installer`)**, using the **[FixMissingMSI utility](https://github.com/suyouquan/SQLSetupTools/releases/tag/V2.2.1)** at scale.  

This automation is designed for scenarios where `C:\Windows\Installer` was "cleaned" to save disk space using scripts that compare files against only the Patches registry and then delete everything else (e.g., a script that enumerates `HKLM:\...\Installer\UserData\S-1-5-18\Patches\*\LocalPackage`, compares to `C:\Windows\Installer\*.msp,*.msi`, and removes the "unregistered" remainder). Because that approach ignores the Products registration (for MSIs), it can delete legitimate MSI cache files. Once those MSI files are gone, many applications fail to update or uninstall without manual sourcing (e.g., Azure Arc Agent, SQL Server, Exchange Server, Microsoft Edge, and others).

By orchestrating FixMissingMSI through **centralized deployment tools** (MECM, Ansible, Azure Arc, etc.), these scripts allow you to:  
- Run FixMissingMSI **non-interactively** on all servers.  
- Collect **CSV reports** of missing MSI/MSP files.  
- Build a **shared cache** populated by hosts that already possess the needed files, driven by the merged report.
- Restore installer functionality at scale without manual intervention.  

---

## What This Adds Beyond FixMissingMSI

FixMissingMSI is a GUI-only troubleshooting tool. On its own it:
- Cannot be run non-interactively.
- Cannot scale across many servers.
- Cannot share repair data between machines.

This automation adds:
- Non-interactive execution of FixMissingMSI via .NET Reflection.
- Centralized reporting in `.CSV` format of unresolved installer files.
- Shared cache support driven by a unified report: machines that already have the required MSI/MSP upload them for peers.
- Integration with centralized deployment tools (MECM, Ansible, Azure Arc, etc.) to scale across entire server fleets.

---

## Workflow

| Script | Run Context | Purpose |
|--------|-------------|---------|
| `Step0-Initialize-FileShare.ps1` | Interactive | Prepares the file share, downloads FixMissingMSI, and stages the app and folder layout. |
| `Step1-Invoke-FixMissingMSI.ps1` | Centralized Deployment (MECM/Ansible/Azure Arc/etc.) | Runs FixMissingMSI non-interactively on each server. Attempts to restore missing MSI/MSP files from the local cache (and from the shared cache if it exists). Produces a per-host `.CSV` report. |
| `Step2-Merge-MissingMSIReports.ps1` | Interactive | Merges all per-host `.CSV` reports from Step1 into a unified list of missing files. |
| `Step3-Populate-MsiCache.ps1` | Centralized Deployment (MECM/Ansible/Azure Arc/etc.) | Populates the shared cache only with MSI/MSP files flagged as missing in the merged report. Hosts upload copies they already have locally; this step does not indiscriminately mirror local caches. |
| `Step1-Invoke-FixMissingMSI.ps1` (re-run) | Centralized Deployment (MECM/Ansible/Azure Arc/etc.) | Run again after Step3. This time servers can source their missing files directly from the populated shared cache. |

> Cache population is demand-driven: a server uploads a file to the shared cache only when another server’s report marks it as missing.

---

## How the Non-Interactive Execution Works

FixMissingMSI was designed as a WinForms GUI with no CLI support.
This automation bypasses the UI by loading the EXE via .NET Reflection and invoking internal methods directly:

1. Load the assembly (`FixMissingMSI.exe`).
2. Create the UI form object (`Form1`) to initialize internal state (UI never displayed).
3. Access internal data structures (`myData`, `CacheFileStatus`).
4. Call internal methods (`ScanSetupMedia`, `ScanProducts`, `AddMsiMspPackageFromLastUsedSource`).
5. Use `UpdateFixCommand` to generate copy commands for missing/mismatched files.
6. Collect results from the `rows` collection for further processing.

> Note: This relies on internal implementation details. If FixMissingMSI changes in future versions, reflection bindings may need to be updated.

---

## Credits

- FixMissingMSI is authored and maintained by [suyouquan (Simon Su @ Microsoft)](https://github.com/suyouquan/SQLSetupTools/releases/tag/V2.2.1).
- This automation orchestrates the tool in a non-interactive, scalable way and adds a shared cache mechanism.

---

## Disclaimer

These scripts are provided for internal automation purposes.
Review, test, and validate them in a non-production environment before deployment.
