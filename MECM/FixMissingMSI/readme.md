# FixMissingMSI Automation

This folder contains a set of PowerShell scripts and supporting files to automate rebuilding the **MSI/MSP cache** on servers using **MECM**  and the **FixMissingMSI** utility via PowerShell.  

The workflow is broken into **four steps**. Some are run interactively, while others are deployed via MECM or other centralized management.

---

## Contents

| Script | Execution Context | Purpose |
|--------|------------------|---------|
| `Step0-Initialize-FileShare.ps1` | **Interactive** | Prepares the file share and downloads the [FixMissingMSI application by suyouquan (Simon Su @Microsoft)](https://github.com/suyouquan/SQLSetupTools/releases/tag/V2.2.1), then copies it to the file share for use in later steps. |
| `Step1-Invoke-FixMissingMSI.ps1` | **MECM Deployed** | Runs FixMissingMSI **non-interactively** on each server. It attempts to resolve missing MSI/MSP files from the **local cache first** and then from the **shared cache** if available. On the very first run, the shared cache will not yet exist, so only local sources can be used. Each server generates a `.CSV` report listing unresolved files. |
| `Step2-Merge-MissingMSIReports.ps1` | **Interactive** | Merges the `.CSV` reports generated in Step 1 into a consolidated list for reference by Step3. |
| `Step3-Populate-MsiCache.ps1` | **MECM Deployed** | Uses both the **local cache** and the now-populated **shared cache** to repopulate the MSI/MSP cache across all targeted servers. |

---

## Overview

1. **Prepare the environment**  
   Run **Step0-Initialize-FileShare.ps1** interactively to set up the network file share and download + stage the FixMissingMSI application.

2. **Collect missing MSI/MSP data**  
   Deploy **Step1-Invoke-FixMissingMSI.ps1** via MECM to all relevant servers. Each server runs FixMissingMSI **in non-interactive mode**, attempts to repair missing installer files from its **local cache**, and then checks the **shared cache** if it exists.  
   > On the first run, the shared cache won’t exist yet, so only local cache recovery will be possible. A per-server `.CSV` report is produced listing unresolved files.

3. **Merge reports**  
   Run **Step2-Merge-MissingMSIReports.ps1** interactively to consolidate `.CSV` files into one master report of unresolved files.

4. **Populate local caches**  
   Deploy **Step3-Populate-MsiCache.ps1** via MECM to populate the shared MSI/MSP cache.

5. **Restore missing files from shared cache**  
   Deploy **Step1-Invoke-FixMissingMSI.ps1** again via MECM to restore missing files from the shared MSI/MSP cache.


---

## Technical Details

The **FixMissingMSI** utility was designed as a WinForms GUI application and does not provide a native way to run in non-interactive or command-line mode.  

This automation works around that limitation by leveraging **.NET Reflection** to load and interact with the application’s internal types and methods directly, bypassing the UI.  

In practice, the Step1-Invoke-FixMissingMSI.ps1 does the following:

1. **Load the FixMissingMSI assembly** (`FixMissingMSI.exe`) into the current PowerShell session.  
2. **Instantiate the UI form** (`Form1`) only to initialize handles; the UI itself is never shown, but this step is required or backend methods will throw null reference errors.  
3. **Access internal data structures** such as `myData` and `CacheFileStatus` to configure scan parameters (setup source, filters, etc.).  
4. **Invoke private and public methods** like `ScanSetupMedia`, `ScanProducts`, and `AddMsiMspPackageFromLastUsedSource` to replicate what the GUI would normally trigger when clicking buttons.  
5. **Call `UpdateFixCommand`** to populate "FixCommand" which are copy cmds for each missing/mismatched installer it found a source for.  
6. **Extract results from the `rows` collection**, filtering to only those entries with a `Missing` or `Mismatched` status.  

This technique effectively runs FixMissingMSI in non-interactive mode, enabling it to be orchestrated by MECM, and the additional scripts enable building a shared cache in the environment enabling you to source missing files from another host without manual efforts..   

>  Note: Because this is essentially calling into an app’s internal implementation details, it’s not an officially supported API surface. If the FixMissingMSI codebase changes in future releases, reflection bindings may need to be updated.

---

##  Credits

**FixMissingMSI** is authored and maintained by **[suyouquan](https://github.com/suyouquan/SQLSetupTools/releases/tag/V2.2.1)**.  
This automation simply downloads and orchestrates the tool with the ability to create a shared cache, allowing you to resolve missing MSI/MSP files across your servers.  

---

##  Disclaimer

These scripts are provided for internal automation purposes. Review, test, and validate them in a non-production environment before wide deployment.  
