# Unjunk v52

A native Windows cleanup dashboard built with PowerShell and WPF. Scan C:, inspect the results, and choose which files to remove.

For future development, see [AGENTS.md](AGENTS.md) for the project map, established behavior and validation notes. Historical files are indexed in [archived scripts](archived%20scripts/README.md); the current launchers, build assets, modules and tests remain in the project root and `tests/`.

## Download

**[Download Unjunk v52 for Windows](https://github.com/tay0thman/unjunk/releases/download/v52/Unjunk-v52-windows.zip)**

Extract the entire ZIP into a folder, then double-click **Unjunk.exe**. The package includes the executable and all required scripts; keep those files together. No build step is needed for the release download.

## Launch from source

GitHub source checkouts do not include the compiled executable. Double-click **Launch Unjunk.cmd**, or run:

```powershell
.\unjunk.ps1 -Gui
```

No arguments also opens the GUI. Requires Windows PowerShell 5.1 (included with Windows); no modules or downloads are needed. The small native launcher hosts Windows PowerShell without a console window. Keep it beside the scripts. The launcher uses an execution-policy override for this process only; it does not change your system policy.

To build the launcher from `Unjunk.Launcher.cs`, run `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build.ps1`. The build uses Windows' existing .NET compiler and does not download dependencies. The generated executable is ignored by Git. To rebuild, close Unjunk and rename the existing executable first.

1. Click **Scan C:**. Scanning never deletes files.
2. Filter by category or name. Select a row to see the path, cleanup effects, age rule, and largest files.
3. Check a **category checkbox** to select its eligible targets, or check **All categories** to select eligible targets across categories. Click the category name to filter the table. Category selection is independent of search filters; previously selected categories remain selected when you browse another. A dash indicates partial selection. Uncheck individual targets to fine-tune the selection, or uncheck a category to clear it. Category selection can include Review items in Standard mode, so inspect the confirmation carefully. Empty/informational rows, policy-blocked targets, access-error targets, and running-app targets are excluded unless Close apps is enabled for caches or reboot mode is selected. In reboot mode, category selection excludes backups, plug-ins, installers, and diagnostics.
4. Alternatively, **Select recommended** selects cache/temp targets only, excluding review items and targets with access errors; running caches are eligible with Close apps enabled. It applies to all results, including filtered-out rows; the total and confirmation include all selected rows.
5. **Clean selected** displays the exact selected targets and asks for confirmation. Deletion is permanent, not sent to the Recycle Bin. Caches rebuild as needed.

The GUI stays responsive during scanning and cleaning. **Cancel** stops after the current file; files already deleted cannot be restored. Every completed operation saves a JSON report under `%LOCALAPPDATA%\Unjunk\Reports`. **Export report** saves a copy wherever you choose. Reports include local file paths; review them before sharing.

## Features

### One-click deep clean and mouse-controlled timeline

Click **One-click deep clean** at the top of the main window, or double-click **One-click full clean.cmd** beside Unjunk.exe. The full preset starts after the normal Windows administrator prompt, with no scan/selection/confirmation steps. It runs in its own process and continues if you close the main window. The button deliberately overrides Unjunk's local CrowdStrike Safe limits for this run, leaving your saved profile and Falcon settings unchanged.

The preset selects every category in the current Windows cleanup catalog automatically, plus every verified machine-wide app leftover. This includes application caches, C:\Autodesk extracted installers, user/system temp, application diagnostics, LiveKernelReports, MEMORY.DMP/minidumps, all inventoried Event Viewer channels, privacy traces, shadow copies, DISM WinSxS cleanup and the full registered Windows Disk Cleanup profile. Disk Cleanup can include Recycle Bin and Downloads where those handlers are installed. Saved path exclusions apply to scanned-file work; native Windows tools use their own eligibility rules.

App pruning now means **machine-wide leftovers of uninstalled apps**, not uninstalling apps. The old bloatware list and keep-newest app-version uninstaller are retired. A successful all-user Appx inventory must prove that no account has the app installed. Registered apps (including older versions still registered), shared frameworks, infrastructure, ChatGPT, Codex and Windows Terminal are preserved. An unavailable, empty or ambiguous inventory preserves app data.

The leftover list includes residual package data in local profiles on the system drive, plus unused provisioned copies removed through Windows' supported provisioning API. Scanned residual files are deleted individually with identity, path, exclusion and reparse-point checks; new, changed, excluded and locked files survive. Registration is rechecked before each leftover file and immediately before deprovisioning. An app reinstalled for any account stops deletion of its remaining data. No installed app is uninstalled, WindowsApps ownership is unchanged, and no raw WindowsApps or Installer tree is deleted. Empty folder shells remain. This does not guess which arbitrary Program Files/ProgramData folders belong to an uninstalled Win32 app.

The one-click sequence is **application caches (close mapped apps first when enabled) → machine-wide app leftovers → C:\Autodesk extracted installers → user/system temp → logs and crash reports → Event Viewer (including security/audit and security-product channels) → privacy cleanup → Explorer Home recent files (pins and other apps' Jump Lists kept) → shadow copies → foreground DISM → foreground Windows Disk Cleanup**. The manual panel offers the same categories, plus separate individual/select-all controls for app leftovers. Installed apps are never offered for uninstallation.
The preset uses the validated file catalog and normal Windows maintenance tools. It does not change security-product settings, disable logging, take ownership of WindowsApps or delete raw servicing/Installer directories. **Close apps before cache cleanup** is enabled by default and saved with the main options; it applies to one-click too. A normal close is requested first, then mapped apps still running after five seconds are terminated. Save work before cleaning. Selected cache targets can close apps; C:\Autodesk cleanup also stops Autodesk Access and its UI without changing the service startup setting. Windows, servicing, security processes, installers, other sessions and Unjunk itself are protected. If an app cannot be stopped or restarts, its remaining cache is skipped. Locked, changed and excluded files remain protected.

On completion, **Cleanup timeline** opens with the completed run selected. Drag the slider or anywhere across the bar chart to select a cleanup date; nothing plays or animates automatically. Each date has two bars: green for counted items removed, and blue for total free space after cleanup on the selected drive. The two series use separate labeled scales. Blue bars are normalized to the drive's recorded range, with padding and an explicitly labeled zoomed minimum/maximum, so small changes are readable. The scale stays fixed while scrubbing and recalculates when switching drives. Tooltips and the final-space card retain actual byte measurements. A missing measurement is marked as not measured instead of zero; a drive with unchanged free space shows level bars. Each run also shows measured net space freed, final available space per fixed drive, and each stage's dates, duration and result. Select a stage for its detailed output. Counts include confirmed file deletions, unused provisioned copies, privacy records, cleared event channels and snapshots, while native DISM/Disk Cleanup counts are marked **Not reported**. Disk-space changes are measured before/after and may include other activity; negative changes are preserved. Partial and failed runs remain visible.

Reopen **Cleanup timeline** without cleaning to browse the last 120 runs. Complete history is saved in `%LOCALAPPDATA%\Unjunk\History`, detailed reports in `Reports`. Older v41 reports lack before/after measurements and are not converted into invented history. The main window reads the detached worker's status during native maintenance. Windows operations show elapsed time and process ID, and wait for the command itself instead of lingering descendant hosts. No timeout kills Windows servicing. An already-running Disk Cleanup blocks a second one-click run before it deletes files.

### Windows cleanup and app leftovers

Open **Windows cleanup & apps** in the main window. **Enable Windows cleanup** switches to Standard cleaning and, when needed, uses the existing administrator launch. Then open the panel and click **Scan Windows**. Scanning is read-only.

Cleanup checkboxes and **Clean all** are available immediately, even before a scan. Your cleanup choices are kept when a scan completes or is refreshed. Only execution requires an administrator, Standard mode and a completed scan; selection is temporarily disabled while an operation is running. **Run selected** opens a review dialog before cleanup begins.

The panel offers 14 independently selectable cleanup categories, including privacy, Explorer Home recent files, C:\Autodesk and all Event Viewer channels. **Clean all** selects the complete category list, including shadow copies and full Windows Disk Cleanup. App leftovers have separate **Select all leftovers** and **Clear leftovers** controls. One-click automatically includes both every category and every eligible leftover. Main-window file cleanup also has **Select all**.

**Run selected** presents a scrollable confirmation with every effective operation, exact package names/scopes, Disk Cleanup categories and shadow-copy IDs. No operation runs merely by scanning or toggling a checkbox. Stop finishes the current Windows operation before skipping the remaining jobs. The panel cannot close while its worker is active. If a different Unjunk instance or the detached one-click process is cleaning, Scan and Run are disabled with a clear busy message until it finishes. An existing run is never force-terminated to unlock the UI. Each run saves a JSON report and native output logs; failures, partial file cleanup and restart requirements are reported separately. A new scan is required after execution.

- **App leftovers** uses an all-user registration check, scans residual package data across local profiles on the system drive, and offers unused machine provisioning. Installed apps are excluded and the engine rejects old installed-app selections. See the one-click policy above.
- **User/system temp** uses the configured 1-365 day cutoff, exclusions and existing scanned-file validation. **LiveKernelReports** and **memory dumps** in this panel include all ages; their explicit selection removes crash evidence. The main file catalog retains its 30-day diagnostic cutoff and now also lists MEMORY.DMP. Running-app caches and locked/changed files are skipped.
- **Old installations** and **Windows Update** use selected installed Windows Disk Cleanup handlers. Removing previous installations removes rollback files. WinSxS uses `DISM /Online /Cleanup-Image /StartComponentCleanup /NoRestart`, without ResetBase. The existing Analyze WinSxS button remains read-only.
- **Full Disk Cleanup** includes every registered VolumeCaches handler, including DownloadsFolder, Recycle Bin, old installations and diagnostics when registered. The exact list appears in the preview. Windows determines eligibility at execution time; main-window file exclusions/age settings do not govern native handlers. `/sagerun` processes all drives. It runs unattended with a hidden native process and no category/setup dialog; Windows or third-party handler UI behavior is outside Unjunk's control. Full cleanup absorbs the separate old-installation/update selections. App leftovers and shadow copies are handled by separate operations.
- **Shadow copies** removes only the scanned, client-accessible snapshot IDs on their recorded local volumes, using `vssadmin delete shadows /Shadow=... /Quiet`. Volume identity is revalidated before each deletion; `/For` is not combined with `/Shadow`, which current Windows rejects. This includes dependent System Restore points and Previous Versions. Copies created after the scan are preserved. Backup-provider snapshots outside the client-accessible context are not deleted. In the selectable panel these operations are blocked by CrowdStrike Safe mode; the dedicated one-click preset uses its explicit per-run override.
- Disk Cleanup uses an unused temporary profile in the 9000-9999 range. Existing profiles are preserved, and only this run's StateFlags values are removed afterward, including on failure. A named mutex prevents concurrent Unjunk Windows maintenance runs. No Windows operation is force-terminated and no automatic restart is requested.

Windows integration references: [Disk Cleanup profiles](https://learn.microsoft.com/en-us/troubleshoot/windows-server/backup-and-storage/automating-disk-cleanup-tool), [WinSxS cleanup](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/clean-up-the-winsxs-folder), [all-user app inventory](https://learn.microsoft.com/en-us/powershell/module/appx/get-appxpackage), [provisioned package removal](https://learn.microsoft.com/en-us/powershell/module/dism/remove-appxprovisionedpackage), [shadow-copy deletion](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/vssadmin-delete-shadows).

### Privacy cleanup, including hidden history, without removing Jump Lists

**One-click deep clean automatically includes privacy cleanup.** The manual category is **Privacy: IDM, recents and hidden history**. Preview and results identify each group separately:

| Group | Cleanup |
| --- | --- |
| Recent files | Direct Recent `*.lnk` shortcuts, RecentDocs, RunMRU, TypedPaths, WordWheelQuery, OpenSavePidlMRU and LastVisitedPidlMRU values. |
| Recent programs | UserAssist `Count`, Search RecentApps/RecentItems, MUI application-name/company records, taskbar FeatureUsage, and Compatibility Assistant Store/Persisted values. |
| IDM | Completed/saved download-list records, `UrlHistory.txt` and `foldresHistory.txt`, including a configured IDM data directory. Close apps closes `IDMan` before rescanning and clearing history. Downloaded files, unfinished transfers, registration and settings remain. |
| Prefetch | Direct `%SystemRoot%\Prefetch\*.pf` files at any age, with administrator access. Windows recreates these; initial launches may be slower. |
| Folder browsing | Numbered ShellBag/BagMRU records in current-user Shell and ShellNoRoam stores, including the UsrClass-backed Local Settings branch. This resets saved per-folder views; global view defaults remain. |
| Background/compatibility | Current-user BAM/DAM executable-use records and the machine AppCompatCache history value, using normal registry permissions. Service bookkeeping and compatibility database configuration remain. |
| Application MRUs | Known MPC-HC/MPC-BE, Paint and WordPad recent-file/media-history registry locations. |
| PCA logs | The three named Compatibility Assistant logs: PcaAppLaunchDic.txt, PcaGeneralDb0.txt and PcaGeneralDb1.txt. Locked files are reported. |
| Local activity history | ActivitiesCache.db history/operation/package/asset rows under ConnectedDevicesPlatform. SQLite transactions preserve settings and encryption metadata; a write failure rolls back. Journals are checkpointed/truncated where available; a busy journal is reported. |

The general privacy stage never enters or changes `Recent\AutomaticDestinations` or `Recent\CustomDestinations`. Jump Lists therefore retain their own history at that stage. File exclusions, path/link boundaries and changed-record checks apply. Current-account history and system Prefetch/PCA/compatibility traces are covered; other user accounts are not edited.

**Explorer Home recent files (keep pins)** is a separate, explicitly authorized exception included in one-click. Emptying the Recent shortcuts folder and RecentDocs registry does not empty Explorer Home/Quick Access: Explorer maintains its own list. This option clears usage for each scanned, unpinned file through Windows' [IApplicationDestinations::RemoveDestination](https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-iapplicationdestinations-removedestination), scoped to `Microsoft.Windows.Explorer_RecentFiles`. Files can appear normally when reopened. It keeps all folders, pinned favorites, original files and other apps' Jump Lists. Explorer's own history store can change; no Jump List database is directly deleted. No global SHAddToRecentDocs reset is used.

The Explorer action rechecks pins, exclusions and the presence of a recognized history-removal command at execution, using that command only as an eligibility hint. No context-menu command is invoked: **Remove from list** can suppress future appearances as well as the current entry. It refreshes Home between batches because Explorer reveals older entries after the visible list is removed. Currently eligible unpinned entries, including ones that appear during cleanup, are cleared; newly pinned items remain protected. Unsupported/localized commands are reported as skipped. Each removal is confirmed with a fresh Explorer inventory, and only confirmed removals appear in the timeline as privacy records. Results also show the final eligible-entry count. A two-minute/100-batch bound and no repeated attempts on failed identities prevent endlessly recreated history from hanging cleanup. The manual panel offers this option independently of general privacy cleanup. Its worker runs in STA for Windows Shell compatibility.

Archives such as ZIP files can be reported as folders by the Windows Shell. Unjunk checks filesystem directory attributes before treating a Shell container as a protected folder. Unpinned archive history is cleared through the same usage API; archive contents remain intact and reopening them can add a new recent entry. Real directories (including names ending in `.zip`), virtual folders, pinned archives and exclusions remain protected. Unreadable container metadata without a confirmed pin is reported as an inventory issue. See Microsoft's [Shell folder and stream attributes](https://learn.microsoft.com/en-us/windows/win32/shell/sfgao).

The result shows confirmed removals and skipped/inaccessible stores for every group, without copying erased document names, URLs, usage values or registry hashes into reports. One inaccessible group no longer hides all other privacy work. Large registry inventories report progress and can be cancelled. Privacy runs before shadow copies, DISM and Disk Cleanup, and its breakdown is saved in the mouse-controlled timeline.

Coverage is the listed local stores, not a guarantee that every trace is erased. Live Amcache/SRUM databases, filesystem journals, unrecognized application stores and cloud-synced histories are not wiped. Windows and running applications can recreate local records. No tracking policies or security-product settings are disabled. Event Viewer clearing remains its separate existing one-click stage.

References: [IDM URL-history location](https://www.internetdownloadmanager.com/register/new_faq/functions16.html), [IDM settings and download-list storage](https://www.internetdownloadmanager.com/register/new_faq/functions17.html), [IDM completed-download behavior](https://www.internetdownloadmanager.com/support/main.html), and [NIST Windows registry history artifacts](https://www.nist.gov/document/windows-registry-forensic-tool-test-assertions-and-test-plan-draft-2-version-10). IDM completion markers were additionally verified against the local installed schema; unknown schemas are left unchanged.
### File cleanup

- **CrowdStrike Safe cleaning** is on by default and persists between launches. This is a conservative Unjunk profile, not a CrowdStrike certification or guarantee against alerts. It permits only cataloged user-profile caches and temporary files at least seven days old. System targets, diagnostics, downloaded installers, administrator relaunch, and Windows maintenance tools are blocked while it is enabled. The restriction is also enforced in the cleanup engine, not just the GUI. Switching profiles invalidates the previous scan and selection.
- Neither profile changes Falcon, antivirus settings, exclusions, services, policies, or telemetry. Event logs are preserved. Restore points are removed only by the explicitly selected shadow-copy operation in Windows cleanup (also included in Clean all). A local cleaning profile cannot guarantee how an organization's security policy will classify a script or executable.
- Size previews and a separate selected-size total; no automatic selection on launch or after scanning.
- All detected Edge, Chrome, and Brave Default/Profile-number caches.
- Claude, Vagon, Gemini, Oculus, Calibrite, Discord, Slack, VS Code, and Vectorworks Install Manager cache leaves.
- **Codex caches** appear under **Desktop apps** (search `Codex`) and are included in **Application caches** and the one-click preset. Handles both `%APPDATA%\Codex` and the Windows app's `LocalCache\Roaming\Codex` storage: browser HTTP/code/GPU caches, Dawn caches, and shader caches, including the in-app browser partition. Only existing, explicitly named cache folders are offered. Close Codex first or enable the close-apps option; its `ChatGPT.exe`, `Codex.exe` and computer-use helper processes must exit before immediate cleanup. Chats, workspaces, settings, cookies/sign-in, browser storage, extensions, runtimes and the entire `.codex` home are preserved. Linked folders are skipped, including relocated `.codex\cache` storage. CrowdStrike Safe keeps the same seven-day minimum; one-click uses its existing full-clean override.
- Targeted Adobe, Revit, Bluebeam, Rhino, V-Ray, Teams, Epic, Steam, NVIDIA, DirectX, pip, and NuGet caches.
- Separate review items for downloaded Vectorworks Cloud, Calibrite, Vagon, and Upscayl installers, plus old driver/setup extraction folders.
- **Rhino 9 settings backup ZIPs**, under Creative apps, matches only `*.zip` files directly in `%APPDATA%\McNeel\Rhinoceros\9.0`. No subfolders are scanned for this target. It preserves active settings, plug-ins, UI files, and non-ZIP files. Because these are potential restore copies, it is an unchecked Review item available in Standard mode only; close Rhino and review the filenames before selecting. No age limit is applied. It is not eligible for reboot deletion. If no ZIPs exist, enable **Show empty** to see the target.
- **Rhino plug-ins** inventories `%APPDATA%\McNeel\Rhinoceros\packages\9.0` and offers one Review target per verifiably older version. It preserves the version selected by `manifest.txt`, the highest installed version, and the highest stable version; equivalent build-metadata versions and unknown layouts are kept. Numeric 3/4-part versions and prerelease identifiers are compared by version precedence, not text or dates. Missing/ambiguous pointers, invalid manifests, and linked folders are preserved. The inventory row explains preserved packages even when there is nothing to clean. Standard mode and explicit selection are required; reboot deletion is blocked. Rhino/Yak must be closed. The active pointer and eligibility are rechecked before cleanup, with an additional active-pointer/process check between files. Only scanned files and then empty directories under the selected old-version root are removed; parent package folders and active pointers are preserved. This removes rollback copies, so keep any older plug-in version you still need. [McNeel's explanation of the active-version pointer](https://discourse.mcneel.com/t/docked-webview-panels-throw-plugin-id-already-in-use-on-yak-update/216878/2).
- Diagnostic reports and logs shown separately, with a 30-day minimum age and no recommended selection.
- Temporary files default to seven days old. Set 1-365 days and path exclusions in **Scan options**. Settings persist in `%LOCALAPPDATA%\Unjunk\settings.json`.
- **Close apps before cache cleanup** is visible beside the main selection buttons and in the Windows panel. It is enabled by default; turning it off preserves running-app caches. The main setting also controls one-click. CLI users can opt in with `-CloseApps`. Preview/WhatIf, empty targets and delete-at-reboot never close apps. Only catalog-mapped processes in the current interactive session are candidates; protected/system/security processes and installers are excluded.
- Standard-user scanning and cleanup; inaccessible paths show partial/access errors. **Open as administrator** requests UAC and starts a fresh window. Use the same Windows account to scan its profile.
- **Analyze WinSxS** runs read-only DISM in a visible progress console when elevated. Component cleanup uses the same visible console. Disk Cleanup opens its normal progress dialog. Unjunk waits for completion, captures DISM output/exit code and then advances to the next stage. No servicing process is force-terminated. Analysis and cleanup share the maintenance lock.
- **Storage settings** opens Windows' own cleanup controls for update files, thumbnails, and Recycle Bin review.

## Cleanup guarantees and limits

Only files in the completed scan and selected catalog targets are considered. The engine validates target boundaries, rechecks exclusions and reparse points, and skips files whose size, creation time, or modification time changed. Scans expire after 30 minutes. New files are not swept up by a recursive directory deletion. In normal cleanup, locked/access-denied files are skipped, never automatically queued. Empty directory trees are left intact except within verified old Rhino plug-in versions; those are pruned with nonrecursive, empty-only deletions.

## Delete at next reboot

1. Turn off **CrowdStrike Safe cleaning**, then choose **Open as administrator**.
2. Scan, review the file previews, and select cache or temporary-file targets.
3. Check **Delete selected at next reboot**, then click **Queue for reboot** and review the confirmation.

This mode queues all eligible selected files (including locked files or files in running-app caches); it does not delete anything immediately and does not automatically restart Windows. Installers and diagnostic targets are not allowed. Queued counts and estimated sizes are separate from deleted counts in the result and JSON report. Open/export the report to inspect the complete queued-path list.

The option is off at every launch and is disabled by CrowdStrike Safe cleaning. The engine also enforces both this restriction and administrator access. Preview/WhatIf never adds pending deletions. A failed Windows API call is reported as skipped, not queued.

Windows uses the documented [MoveFileExW restart-deletion mechanism](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-movefileexw). It queues paths, not immutable file identities. The app validates the scanned file immediately before queuing, but cannot revalidate at startup: a changed or replacement file at the same path can be deleted. Queue success means the request was registered, not proof of eventual deletion or reclaimed space. Restart when ready and scan afterward to verify. Cancelling a running queue operation stops further additions but does not undo entries already queued; neither does unchecking the option. Unjunk does not edit or clear other applications' pending operations and has no queue-undo feature.

Sizes are logical file lengths, not guaranteed disk-space recovery: compression, hard links, and background activity can affect the actual change. Access errors and skipped links mean a scan may be incomplete. There is no broad custom-path delete feature.

The file-cleanup catalog preserves Windows Installer, Package Cache, WinSxS, servicing files, WindowsApps permissions, event logs, browser databases and cookies, cloud/collaboration data, recovery files, and application settings. The separate Windows cleanup panel provides explicit native maintenance, app-leftover cleanup and shadow-copy removal as described above. It never directly deletes component-store or WindowsApps trees, resets services, kills Explorer or erases downloaded models.

## Console use

```powershell
# Preview only; prints stable target IDs and saves a JSON report
.\unjunk.ps1 -DryRun
.\unjunk.ps1 -Console -TempAgeDays 14 -ExcludePath 'C:\Users\you\AppData\Local\Temp\Keep'

# Opt out of the default CrowdStrike Safe profile for a broader preview
.\unjunk.ps1 -DryRun -StandardCleaning

# Explicit target selection; use IDs printed on your own machine
.\unjunk.ps1 -Console -Clean -TargetId 'TARGET_ID_FROM_PREVIEW'

# Validate the selection without deleting
.\unjunk.ps1 -Console -Clean -TargetId 'TARGET_ID_FROM_PREVIEW' -WhatIf

# Queue selected cache/temp files for restart (administrator, Standard profile)
.\unjunk.ps1 -Console -Clean -StandardCleaning -QueueOnReboot -TargetId 'TARGET_ID_FROM_PREVIEW'
# -DeleteAfterReboot is an alias for -QueueOnReboot; add -WhatIf to preview without queuing
```

`-Clean` always performs a fresh scan. `-Force` is accepted only with `-Clean` and explicit target IDs, and only bypasses the selected-target confirmation. `-Force` alone no longer performs broad/deep maintenance. `-DryRun` takes precedence over cleanup. `-ReportPath` sets the JSON output path, including for preview/WhatIf reports. Console mode uses explicit arguments; saved GUI preferences apply to the GUI. Console mode defaults to CrowdStrike Safe cleaning; add `-StandardCleaning` to opt out for that invocation.

## Migration from v39

The previous script combined cache cleanup with app removal, system changes, forced process termination, and recovery/history deletion. The v40 entry point replaces that sequence with selected-file cleanup. In particular, the previous Rhino MSI deletion and WindowsApps ownership/deletion routines are not carried forward. Full app data such as Edge IndexedDB/OPFS, Revit collaboration data, Bluebeam recovery/Studio sessions, Rhino autosaves, Honeybee simulations, and IDM/qBittorrent download data are not junk targets.

Before this migration, the working copy (including uncommitted changes) was saved byte-for-byte to `%LOCALAPPDATA%\Unjunk\Backups\unjunk-before-gui-20261030-160753.ps1.txt`. It is retained for reference, not invoked by the new app. The legacy `windowsApps_cleanup.ps1` now opens the protected GUI instead of piping selected installed packages into removal. Historical scripts, registry data and shortcuts are preserved unchanged in [archived scripts](archived%20scripts/README.md); use **Launch Unjunk.cmd** for the current interface.


### Autodesk installer files and Event Viewer

**C:\Autodesk extracted installers** includes scanned WI/Access download and extraction files of all ages in Windows cleanup and one-click. The main Standard scan also includes all ages as an explicit Review target. With Close apps enabled, Unjunk stops the Autodesk Access updater service and closes its UI before deletion. Active setup processes still block deletion and are never terminated. Empty extraction folders are removed afterward using nonrecursive deletion; links and license-manager folders are preserved. Network License Manager folders and saved exclusions are preserved. Future repairs may need the installers downloaded again. [Autodesk guidance](https://www.autodesk.com/support/technical/article/caas/sfdcarticles/sfdcarticles/C-Autodesk-contain-to-many-files.html).

**Event Viewer logs (including security/audit)** inventories registered channels and clears the scanned channels through Windows `wevtutil cl`, including Security/audit and security-product channels. Logging remains enabled; no channel configuration, security-product setting or file permission is changed. Empty channels are skipped. Newly registered channels are preserved; denied/unsupported clears are reported as partial results. Windows may immediately record new events, including the clear itself. The timeline counts successfully cleared channels, not an invented number of individual events. Reports contain channel names and outcomes, not copies of cleared event messages. [Windows event-log command reference](https://learn.microsoft.com/windows-server/administration/windows-commands/wevtutil).

## Validation

Run filesystem tests from a copy of the project on the system drive without junction/reparse-point ancestors. Some fixtures are created under `tests/`; the environment must allow disposable junction creation and native file moves. Sandboxes or linked workspaces can otherwise trigger the same path protections that the cleaner intentionally enforces.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Core.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\RhinoPackages.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Selection.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Windows.Tests.ps1
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Gui.Tests.ps1
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Gui.Callbacks.Tests.ps1
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\OneClick.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\AppLeftovers.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Privacy.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Privacy.Extended.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\ActivityHistory.Tests.ps1
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\ExplorerHistory.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\SystemScan.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\CodexCache.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\CloseApps.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Autodesk.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\AutodeskService.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Events.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Foreground.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\WindowsBusy.Tests.ps1
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\History.Gui.Tests.ps1
```

The file regression checks use uniquely named disposable fixtures under `tests/` or the current user's Temp folder. They exercise age/exclusion filters, preview behavior, changed and locked files, target boundaries, junctions, protected paths, cancellation, stale scans, active-process detection, selection validation, and report serialization. Windows tests mock native processes, registry state, app providers and shadow-copy providers, including failure recovery and WhatIf. GUI tests construct and lay out controls offscreen. These tests do not clean real application data, remove installed apps or delete real snapshots.
