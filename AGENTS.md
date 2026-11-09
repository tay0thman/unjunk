# Unjunk contributor handoff

Updated 2026-11-08. The current application is v52, a native Windows PowerShell/WPF cleanup dashboard. Read `README.md` for the full behavior and commands; this file records the development map and decisions that future changes must preserve. Update this handoff when those decisions change.

## Project map

| Files | Responsibility |
| --- | --- |
| `unjunk.ps1`, `Unjunk.Gui.ps1` | Main entry point and dashboard. |
| `Unjunk.Core.psm1` | File catalog, scanning, exclusions, identity checks, cleanup and reports. |
| `Unjunk.Windows.psm1`, `Unjunk.Windows.Gui.ps1` | Selectable Windows maintenance operations and their UI. |
| `Unjunk.OneClick.ps1`, `Unjunk.OneClick.psm1` | Detached full-clean worker, sequencing, measurements and history persistence. |
| `Unjunk.History.Gui.ps1` | Mouse-controlled cleanup timeline and per-drive graphs. |
| `Unjunk.AppLeftovers.psm1` | Machine-wide leftovers of apps no longer installed. |
| `Unjunk.Privacy.psm1`, `Unjunk.ActivityHistory.psm1` | Privacy traces, recent activity and application histories. |
| `Unjunk.ExplorerHistory.psm1`, `Unjunk.ExplorerRecentNative.cs` | Targeted Explorer Home recent-file cleanup. |
| `Unjunk.Events.psm1`, `Unjunk.Dism.ps1` | Event-log cleanup and visible DISM progress. |
| `Unjunk.Launcher.cs`, `Build.ps1`, `unjunk.ico` | Native launcher source and build assets. The icon is still required. |
| `tests/` | Fixture-based regression checks, including offscreen WPF checks. |
| `archived scripts/` | Historical reference files; not application dependencies or test inputs. See its README. |

Keep `Unjunk.exe`, `Launch Unjunk.cmd`, `One-click full clean.cmd` and the runtime modules together in the project root. `windowsApps_cleanup.ps1` is a current compatibility entry point that opens the protected GUI; it is not the old app-removal implementation. The executable loads the sibling scripts, so script-only changes do not require rebuilding it.

## Established product behavior

- One-click includes every current Windows cleanup category and eligible machine-wide app leftovers. Preserve the sequence: application caches, app leftovers, Autodesk installers, temp files, logs/dumps, Event Viewer, privacy, Explorer recents, shadow copies, DISM, then Windows Disk Cleanup.
- The one-click preset overrides only Unjunk's local CrowdStrike Safe limits for that run. It does not alter Falcon settings, exclusions, security-product configuration or the saved profile.
- App pruning means leftovers, not uninstalling installed applications. Revalidate successful all-user registration data before removal. Preserve registered apps, shared frameworks and protected infrastructure, including ChatGPT, Codex and Windows Terminal. Ambiguous inventories preserve data.
- Close apps is enabled by default. Use existing mapped-process/session protections and the normal-close-then-terminate behavior. Autodesk cleanup may stop its Access updater service/UI, but must preserve license-manager files and active installers.
- The user requested privacy cleanup including IDM history, recent files/programs, hidden activity and Prefetch. Jump Lists remain protected except for the explicit Explorer recent-file exception below.
- Event Viewer includes scanned security/audit and security-product channels through supported Windows APIs, without disabling logging or changing permissions/settings. Keep partial/error reporting honest.
- Preserve scanned-file boundaries, exclusions, identity checks, reparse-point protection and locked/changed-file handling. Do not replace them with recursive deletion of live application trees or raw WindowsApps/WinSxS/Installer directories.
- Keep DISM and Disk Cleanup progress visible and sequential. Preserve maintenance locking, cancellation boundaries and exit-code reporting; do not force-terminate servicing to release a busy lock.

## Explorer recents: preserve the v51 repair

The approved exception clears unpinned files from Explorer Home while keeping folders, pins, favorites and other apps' Jump Lists. Newly opened files must appear again after cleaning.

Use the existing `IApplicationDestinations.RemoveDestination` bridge scoped to `Microsoft.Windows.Explorer_RecentFiles`. Do not restore the Shell context-menu `Remove from list`/`Remove from Quick access` invocation: it previously hid files persistently, including reopened PDFs. Do not delete whole destination stores or rewrite their binary contents as routine cleanup. ZIP files can report `IsFolder`; retain the physical file-attribute check that distinguishes archives from directories.

The user verified the repair by double-clicking a real PDF in Explorer and confirming it reappeared after targeted cleanup. Synthetic `SHAddToRecentDocs` calls alone did not reliably reproduce this real application workflow.

## Timeline: preserve the v52 behavior

- Mouse/slider selection only; no animation, playback timer or automatic cycling.
- Each run has green item-count and blue total-free-space bars for the selected drive, using separate labeled scales.
- Blue heights use the recorded minimum/maximum with padding and an explicit zoomed-range label. Use all displayed runs for that drive so scrubbing does not change the scale. Switching drives recalculates it.
- Keep actual free-space values in data, tooltips and the selected-run card. Missing measurements differ from measured zero; constant values remain level. Preserve signed net disk changes and unknown native item counts.
- Windows PowerShell overload binding matters: use double-valued arguments for byte-range/fraction `Math.Min`/`Math.Max` calculations to avoid Int32 overflow or rounding normalized fractions to integers.

## Development and validation

Use Windows PowerShell 5.1 and STA for WPF/native-shell tests. Run filesystem fixture suites from a copy on the system drive with no junction/reparse-point ancestors; some fixtures are created under `tests/`. The test environment must allow creating disposable junctions and native file moves. Keep production path protections intact when a sandbox or linked workspace prevents these checks. From that copy's repository root, for example:

```powershell
$ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
& $ps51 -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\History.Gui.Tests.ps1
& $ps51 -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\ExplorerHistory.Tests.ps1
& $ps51 -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\OneClick.Tests.ps1
& $ps51 -NoProfile -ExecutionPolicy Bypass -File .\tests\Windows.Tests.ps1
```

Run checks relevant to the change; the README lists the remaining focused suites. For timeline layout changes, `History.Gui.Tests.ps1 -RenderDirectory <directory>` renders an offscreen PNG for inspection. Before publishing v52, all 21 fixture suites passed, including 28 timeline checks; the timeline was also visually checked against saved history at a narrow width. Filesystem suites initially blocked by the sandbox/linked workspace passed when rerun from a normal directory with the required Windows filesystem capabilities.

Routine verification uses disposable fixtures, mocked providers and offscreen views. Do not run the full cleaner, clear real history/logs or remove real snapshots just to test a code change. Tests that touch files should stay within verified disposable roots.

History, reports and settings live under `%LOCALAPPDATA%\Unjunk`. A sandbox may substitute a different profile path: verify the effective account/path before diagnosing live user state. Keep personal paths and report contents out of source fixtures.

Preserve unrelated working-copy changes. Use a baseline and hashes when staging an installation outside the writable workspace; verify copied files. Do not reset the repository or change Git configuration to make a tool command succeed. Root launch/build files and tests are active; archived scripts remain reference material and should not be loaded, executed or packaged as runtime dependencies.
