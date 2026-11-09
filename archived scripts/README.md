# Archived scripts

Moved here from the project root on 2026-11-08. These six files are historical references and are not used by the active Unjunk v52 application. Their original filenames and contents are preserved.

| File | Historical purpose |
| --- | --- |
| `text.txt` | Earlier monolithic PowerShell cleanup script saved as text. |
| `Clear_Event_Viewer_Logs.bat` | Standalone Event Viewer cleanup, superseded by the current selectable Event Viewer operation. |
| `___________Clear_Event_Viewer_Logs.lnk` | Shortcut to the former root-level batch file. |
| `UnJunk_.lnk` | Old PowerShell 7 shortcut targeting a separate Google Drive project path. |
| `Remove_Gallery_from_navigation_pane_in_File_Explorer_for_current_user.reg` | Standalone File Explorer Gallery navigation tweak. |
| `f01b4d95cf55d32a.automaticDestinations-ms` | Historical Explorer destination-store file, not current application data or a source dependency. |

Use the launchers in the project root for the current application. Archived shortcuts retain their historical targets; they are not maintained launchers. These scripts and registry data are retained for reading, not for automatic execution, import, tests or deployment. The destination-store file may contain historical personal paths; do not treat it as a generic test fixture or copy it into a live profile.

`windowsApps_cleanup.ps1` remains in the root because its current implementation is a compatibility launcher for the protected GUI. The active launcher icon, source, modules and regression tests also remain in place.
