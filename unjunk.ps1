cls
$ErrorActionPreference = "SilentlyContinue"
$FreespaceBefore = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
echo "
                                                   __       
                                                  |  \      
 __    __  _______         __  __    __  _______  | XX   __ 
|  \  |  \|       \       |  \|  \  |  \|       \ | XX  /  \
| XX  | XX| XXXXXXX\       \XX| XX  | XX| XXXXXXX\| XX_/  XX
| XX  | XX| XX  | XX      |  \| XX  | XX| XX  | XX| XX   XX 
| XX__/ XX| XX  | XX      | XX| XX__/ XX| XX  | XX| XXXXXX\ 
 \XX    XX| XX  | XX      | XX \XX    XX| XX  | XX| XX  \XX\
  \XXXXXX  \XX   \XX __   | XX  \XXXXXX  \XX   \XX \XX   \XX
                    |  \__/ XX                              
                     \XX    XX                              
                      \XXXXXX

                      By Tay Othman      
                      "



#...........................................................................................................................................................................................................................................................
#....CCCCCCC....LLLL.......EEEEEEEEEEE....AAAAA.....NNNN...NNNN..UUUU...UUUU..PPPPPPPPP...........AAAAA.....UUUU...UUUU..TTTTTTTTTTO..OOOOOOO....DDDDDDDDDD...EEEEEEEEEEE...SSSSSSS...SKKK....KKKKK..... FFFFFFFFFFFIII.ILLL.......EEEEEEEEEEE...SSSSSSS....
#...CCCCCCCCC...LLLL.......EEEEEEEEEEE....AAAAA.....NNNNN..NNNN..UUUU...UUUU..PPPPPPPPPP..........AAAAA.....UUUU...UUUU..TTTTTTTTTTO.OOOOOOOOOO..DDDDDDDDDDD..EEEEEEEEEEE..SSSSSSSSS..SKKK...KKKKK...... FFFFFFFFFFFIII.ILLL.......EEEEEEEEEEE..SSSSSSSSS...
#..CCCCCCCCCCC..LLLL.......EEEEEEEEEEE...AAAAAA.....NNNNN..NNNN..UUUU...UUUU..PPPPPPPPPPP........AAAAAA.....UUUU...UUUU..TTTTTTTTTTOOOOOOOOOOOOO.DDDDDDDDDDDD.EEEEEEEEEEE.ESSSSSSSSSS.SKKK..KKKKK....... FFFFFFFFFFFIII.ILLL.......EEEEEEEEEEE.ESSSSSSSSSS..
#..CCCC...CCCCC.LLLL.......EEEE..........AAAAAAA....NNNNNN.NNNN..UUUU...UUUU..PPPP...PPPP........AAAAAAA....UUUU...UUUU.....TTTT...OOOOO...OOOOO.DDDD....DDDD.EEEE........ESSS...SSSS.SKKK.KKKKK........ FFF.......FIII.ILLL.......EEEE........ESSS...SSSS..
#.CCCC.....CCC..LLLL.......EEEE.........AAAAAAAA....NNNNNN.NNNN..UUUU...UUUU..PPPP...PPPP.......AAAAAAAA....UUUU...UUUU.....TTTT...OOOO.....OOOOODDDD....DDDDDEEEE........ESSSS.......SKKKKKKKK......... FFF.......FIII.ILLL.......EEEE........ESSSS........
#.CCCC..........LLLL.......EEEEEEEEEE...AAAAAAAA....NNNNNNNNNNN..UUUU...UUUU..PPPPPPPPPPP.......AAAAAAAA....UUUU...UUUU.....TTTT...OOOO......OOOODDDD.....DDDDEEEEEEEEEEE..SSSSSSS....SKKKKKKK.......... FFFFFFFFF.FIII.ILLL.......EEEEEEEEEEE..SSSSSSS.....
#.CCCC..........LLLL.......EEEEEEEEEE...AAAA.AAAA...NNNNNNNNNNN..UUUU...UUUU..PPPPPPPPPP........AAAA.AAAA...UUUU...UUUU.....TTTT...OOOO......OOOODDDD.....DDDDEEEEEEEEEEE...SSSSSSSS..SKKKKKKKK......... FFFFFFFFF.FIII.ILLL.......EEEEEEEEEEE...SSSSSSSS...
#.CCCC..........LLLL.......EEEEEEEEEE..AAAAAAAAAA...NNNNNNNNNNN..UUUU...UUUU..PPPPPPPPP........AAAAAAAAAA...UUUU...UUUU.....TTTT...OOOO......OOOODDDD.....DDDDEEEEEEEEEEE.....SSSSSSS.SKKKKKKKKK........ FFFFFFFFF.FIII.ILLL.......EEEEEEEEEEE.....SSSSSSS..
#.CCCC.....CCC..LLLL.......EEEE........AAAAAAAAAAA..NNNNNNNNNNN..UUUU...UUUU..PPPP.............AAAAAAAAAAA..UUUU...UUUU.....TTTT...OOOO.....OOOOODDDD....DDDDDEEEE...............SSSSSSKKKK.KKKK........ FFF.......FIII.ILLL.......EEEE...............SSSS..
#..CCCC...CCCCC.LLLL.......EEEE........AAAAAAAAAAA..NNNN.NNNNNN..UUUU...UUUU..PPPP.............AAAAAAAAAAA..UUUU...UUUU.....TTTT...OOOOOO..OOOOO.DDDD....DDDD.EEEE........ESSS...SSSSSSKKK..KKKKK....... FFF.......FIII.ILLL.......EEEE........ESSS...SSSS..
#..CCCCCCCCCCC..LLLLLLLLLL.EEEEEEEEEEEAAAA....AAAA..NNNN..NNNNN..UUUUUUUUUUU..PPPP............ AAA....AAAA..UUUUUUUUUUU.....TTTT....OOOOOOOOOOOO.DDDDDDDDDDDD.EEEEEEEEEEEEESSSSSSSSSS.SKKK...KKKKK...... FFF.......FIII.ILLLLLLLLLLEEEEEEEEEEEEESSSSSSSSSS..
#...CCCCCCCCCC..LLLLLLLLLL.EEEEEEEEEEEAAAA.....AAAA.NNNN..NNNNN...UUUUUUUUU...PPPP............ AAA.....AAAA..UUUUUUUUU......TTTT.....OOOOOOOOOO..DDDDDDDDDDD..EEEEEEEEEEEE.SSSSSSSSSS.SKKK....KKKK...... FFF.......FIII.ILLLLLLLLLLEEEEEEEEEEEE.SSSSSSSSSS..
#....CCCCCCC....LLLLLLLLLL.EEEEEEEEEEEAAAA.....AAAA.NNNN...NNNN....UUUUUUU....PPPP........... AAA.....AAAA...UUUUUUU.......TTTT......OOOOOOO....DDDDDDDDDD...EEEEEEEEEEEE..SSSSSSS...SKKK....KKKKK..... FFF.......FIII.ILLLLLLLLLLEEEEEEEEEEEE..SSSSSSS....
#...........................................................................................................................................................................................................................................................

write-output ("Disk C:\ is currently at [{0:N2}" -f ($FreeSpaceBefore.freespace/1GB) + "] Gb available.")
Remove-Item "$env:SYSTEMROOT\Temp\*" -Recurse -Force

Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2020\CollaborationCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2020\Journals\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2021\CollaborationCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2021\Journals\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2023\CollaborationCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2023\Journals\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2022\CollaborationCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2022\Journals\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2024\CollaborationCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2024\Journals\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2025\CollaborationCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2025\Journals\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Autodesk\Revit\PacCache\*" -Recurse -Force
Remove-Item "C:\Autodesk\*" -Recurse -Force
Remove-Item "$env:AppData\Autodesk\ADPSDK\JSON" -Recurse -Force
Remove-Item "C:\ProgramData\RevitInterProcess\*" -Recurse -Force
freespace1 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared1 = $FreeSpaceBefore.freespace - $freespace1.freespace
write-output ("Cleared [{0:N2}" -f ($cleared1/1GB) + "] Gb of space.")
echo "............Done Cleaning Autodesk Files"
# ***************************************************************************************************************************************************Delete ACCDocs Cache
Remove-Item "$env:HOMEPATH\ACCDOCS\*" -Recurse -Force

freespace2 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared2 = $freespace1.freespace - $freespace2.freespace
write-output ("Cleared [{0:N2}" -f ($cleared2/1GB) + "] Gb of space.")
echo "............Done ACCDOCS Files"
# ***************************************************************************************************************************************************Delete Chaos Vantage Cache
Remove-Item "$env:LOCALAPPDATA\Chaos Group\Vantage\cache\QtWebEngine\Default\Cache\*" -Recurse -Force

freespace3 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared3 = $freespace2.freespace - $freespace3.freespace
write-output ("Cleared [{0:N2}" -f ($cleared3/1GB) + "] Gb of space.")
echo "............Done Cleaning Chaos Vantage Files"
# ***************************************************************************************************************************************************Delete Vray Logs
# Delete All *.log files located here "C:\Users\tayO\AppData\Roaming\Chaos Group\V-Ray for Rhinoceros\vrayneui"
Remove-Item "$env:APPDATA\Chaos Group\V-Ray for Rhinoceros\vrayneui\*.log" -Recurse -Force
freespace4 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared4 = $freespace3.freespace - $freespace4.freespace
write-output ("Cleared [{0:N2}" -f ($cleared4/1GB) + "] Gb of space.")
echo "............Done Cleaning Vray Logs"

# ***************************************************************************************************************************************************Delete Logitech Ghub Cache
Remove-Item "C:\ProgramData\LGHUB\cache\*" -Recurse -Force
Remove-Item "C:\ProgramData\LGHUB\depots\*" -Recurse -Force
freespace5 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared5 = $freespace4.freespace - $freespace5.freespace
write-output ("Cleared [{0:N2}" -f ($cleared5/1GB) + "] Gb of space.")
echo "............Done Cleaning Logitech Files"
# ***************************************************************************************************************************************************Delete Adobe Sensei Cache
# Remove files located on "C:\Users\tayO\AppData\Roaming\Adobe\Creative Cloud Libraries\*"
Remove-Item "$env:APPDATA\Adobe\Creative Cloud Libraries\*" -Recurse -Force
freespace6 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared6 = $freespace5.freespace - $freespace6.freespace
write-output ("Cleared [{0:N2}" -f ($cleared6/1GB) + "] Gb of space.")
echo "............Done Cleaning Adobe Creative Cloud Files"

# ***************************************************************************************************************************************************Remove old uninstalled programs enties
Remove-Item "$env:LOCALAPPDATA\Downloaded Installations" -Recurse -Force
# ***************************************************************************************************************************************************Delete Bluebeam Sessions Cache
Remove-Item "$env:LOCALAPPDATA\Revu\data\Sessions\studio.bluebeam.com\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\CrashDumps\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Bluebeam\Revu\21\WebCache" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Bluebeam\Revu\21\Recovery\*" -Recurse -Force
freespace7 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared7 = $freespace6.freespace - $freespace7.freespace
write-output ("Cleared [{0:N2}" -f ($cleared7/1GB) + "] Gb of space.")
echo "............Done Cleaning Bluebeam Files"

# ===================================================================================================================================================Delete Adobe Cache
Remove-Item "$env:LOCALAPPDATA\Adobe\Lightroom\Caches" -Recurse -Force
Remove-Item "$env:APPDATA\Adobe\Logs\Adobe Illustrator\25.0\Adobe Illustrator\ACPLLogs\*" -fORCE
Remove-Item "$env:AppData\com.adobe.dunamis\*" -Recurse -Force
Remove-Item "C:\adobeTemp" -Recurse -Force
del C:\adobeTemp /f /s /q
freespace8 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared8 = $freespace7.freespace - $freespace8.freespace
write-output ("Cleared [{0:N2}" -f ($cleared8/1GB) + "] Gb of space.")
echo "............Done Cleaning Adobe Temp Files"
# ***************************************************************************************************************************************************Delete Mcneel Update
Remove-Item "C:\ProgramData\McNeel\McNeelUpdate\DownloadCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\McNeel\Rhinoceros\6.0\AutoSave\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\McNeel\Rhinoceros\7.0\AutoSave\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\McNeel\Rhinoceros\8.0\AutoSave\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\McNeel\McNeelUpdate\DownloadCache\*" -Recurse -Force
freepace9 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared9 = $freespace8.freespace - $freespace9.freespace
write-output ("Cleared [{0:N2}" -f ($cleared9/1GB) + "] Gb of space.")
echo "............Done Cleaning Mcneel Files"

# ***************************************************************************************************************************************************Delete Honeybee Simulations
Remove-Item "$env:USERPROFILE\simulation\*" -Recurse -Force

freepace10 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared10 = $freespace9.freespace - $freespace10.freespace
write-output ("Cleared [{0:N2}" -f ($cleared10/1GB) + "] Gb of space.")
echo "............Done Cleaning Honeybee Simulation Files"
# ***************************************************************************************************************************************************Delete Zoom Error Logs
Remove-Item "$env:APPDATA\Zoom\logs\*" -Recurse -Force
Remove-Item "$env:APPDATA\Zoom\data\WebviewCacheX64\bzn01rg8r86wd78tyy1dhw\EBWebView\*" -Recurse -Force
freepace11 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared11 = $freespace10.freespace - $freespace11.freespace
write-output ("Cleared [{0:N2}" -f ($cleared11/1GB) + "] Gb of space.")
echo "............Done Cleaning Zoom Cache Files"

# ***************************************************************************************************************************************************Delete Remote Desktop Cache
Remove-Item "$env:LOCALAPPDATA\Microsoft\Terminal Server Client\Cache\*" -Recurse -Force
freepace12 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared12 = $freespace11.freespace - $freespace12.freespace
write-output ("Cleared [{0:N2}" -f ($cleared12/1GB) + "] Gb of space.")
echo "............Done Cleaning Remote Desktop Cache Files"

# ***************************************************************************************************************************************************Delete WDF Cache
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\WDF\*" -Recurse -Force
Remove-Item "C:\ProgramData\Microsoft\Windows\WDF\*" -Recurse -Force
freepace13 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared13 = $freespace12.freespace - $freespace13.freespace
write-output ("Cleared [{0:N2}" -f ($cleared13/1GB) + "] Gb of space.")
echo "............Done Cleaning WDF Cache Files"

# ***************************************************************************************************************************************************Delete Windows Search Index
# Cleaning up microsoft maps cache
Remove-Item "C:\ProgramData\Microsoft\MapData\*" -Recurse -Force
freepace14 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared14 = $freespace13.freespace - $freespace14.freespace
write-output ("Cleared [{0:N2}" -f ($cleared14/1GB) + "] Gb of space.")
echo "............Done Cleaning Microsoft Maps Cache Files"
# ***************************************************************************************************************************************************Delete MS Teams Cache
# kill all teams processes
taskkill /f /im "ms-teams.exe"
Remove-Item "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\WebStorage\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\Cache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\Service Worker\CacheStorage\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\Service Worker\ScriptCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs\*" -Recurse -Force
Remove-Item "$env:AppData\Microsoft\Teams\Service Worker\CacheStorage\*" -Recurse -Force
Remove-Item "$env:AppData\Microsoft\Teams\Cache\*" -Recurse -Force
freepace15 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared15 = $freespace14.freespace - $freespace15.freespace
write-output ("Cleared [{0:N2}" -f ($cleared15/1GB) + "] Gb of space.")
echo "............Done Cleaning MS Teams Cache Files"
# ***************************************************************************************************************************************************Delete Google Earth
Remove-Item "$env:USERPROFILE\AppData\LocalLow\Google\GoogleEarth\Cache\*" -Recurse -Force
freepace16 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared16 = $freespace15.freespace - $freespace16.freespace
write-output ("Cleared [{0:N2}" -f ($cleared16/1GB) + "] Gb of space.")
echo "............Done Cleaning Google Earth Cache"

# ***************************************************************************************************************************************************Delete Google Drive Cache
#....Google\DriveFS\109972538989880061485\photos_cache_temp
Remove-Item "$env:LOCALAPPDATA\Google\DriveFS\109972538989880061485\photos_cache_temp" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Google\DriveFS\logs\*" -Recurse -Force
freepace17 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared17 = $freespace16.freespace - $freespace17.freespace
write-output ("Cleared [{0:N2}" -f ($cleared17/1GB) + "] Gb of space.")
echo "............Done Cleaning Google Drive Cache"
# ***************************************************************************************************************************************************VSCode Cache
Remove-Item "$env:APPDATA\Code\Cache\*"  -Recurse -Force
Remove-Item "$env:APPDATA\Code\CachedData\*" -Recurse -Force
freepace18 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared18 = $freespace17.freespace - $freespace18.freespace
write-output ("Cleared [{0:N2}" -f ($cleared18/1GB) + "] Gb of space.")
echo "............Done Removing VSCode Cache"
# ***************************************************************************************************************************************************Internet Download Manager Data
taskkill /f /im "idman.exe"
Remove-Item HKCU:\Software\DownloadManager\1* -Recurse
Remove-Item HKCU:\Software\DownloadManager\2* -Recurse
Remove-Item HKCU:\Software\DownloadManager\3* -Recurse
Remove-Item HKCU:\Software\DownloadManager\4* -Recurse
Remove-Item HKCU:\Software\DownloadManager\6* -Recurse
Remove-Item HKCU:\Software\DownloadManager\7* -Recurse
Remove-Item HKCU:\Software\DownloadManager\8* -Recurse
Remove-Item HKCU:\Software\DownloadManager\9* -Recurse
Remove-Item "C:\Users\tayO\AppData\Roaming\IDM\foldresHistory.txt" -Force
freepace19 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared19 = $freespace18.freespace - $freespace19.freespace
write-output ("Cleared [{0:N2}" -f ($cleared19/1GB) + "] Gb of space.")
echo "............Done Cleaning IDMAN Hisotry & Cache Files"
# ***************************************************************************************************************************************************Teamviewer Cache
Remove-Item "$env:LOCALAPPDATA\TeamViewer\EdgeBrowserControl\Temporary\*" -Recurse -Force
freepace20 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared20 = $freespace19.freespace - $freespace20.freespace
write-output ("Cleared [{0:N2}" -f ($cleared20/1GB) + "] Gb of space.")
echo "............Done Cleaning Teamviewer Cache Files"

# ***************************************************************************************************************************************************Delete Nvidia DX Cache
Remove-Item "$env:LOCALAPPDATA\NVIDIA\DXCache\*" -Recurse -Force
Remove-Item "$env:USERPROFILE\AppData\LocalLow\NVIDIA\PerDriverVersion\DXCache\*" -Recurse -Force
Remove-Item "$env:USERPROFILE\AppData\Roaming\NVIDIA\ComputeCache\*" -Recurse -Force
freepace21 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared21 = $freespace20.freespace - $freespace21.freespace
write-output ("Cleared [{0:N2}" -f ($cleared21/1GB) + "] Gb of space.")
echo "............Done Cleaning Nvidia DX and compute Cache Files"
# ***************************************************************************************************************************************************Microsoft Edge
Remove-Item "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\Cache_Data\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache\js\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache\wasm\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Service Worker\CacheStorage\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Service Worker\ScriptCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Edge\User Data\BrowserMetrics-spare.pma" -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Edge\User Data\\BrowserMetrics\*" -Force -Recurse
Remove-Item "$env:LOCALAPPDATA\pip\*" -Force -Recurse
freepace22 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared22 = $freespace21.freespace - $freespace22.freespace
write-output ("Cleared [{0:N2}" -f ($cleared22/1GB) + "] Gb of space.")
echo "............Done Cleaning Microsoft Edge Cache Files"

# ***************************************************************************************************************************************************Delete Windows Web Cache
# kill windows host task service
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\WebCache\*" -Recurse -Force
freepace23 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared23 = $freespace22.freespace - $freespace23.freespace
write-output ("Cleared [{0:N2}" -f ($cleared23/1GB) + "] Gb of space.")
echo "............Done Cleaning Windows Web Cache Files"
# ******************************************************************************************************************************************************Delete FFMPEG
Remove-Item "$env:LOCALAPPDATA\Eibolsoft\*" -Recurse -Force
remove-item "$env:APPDATA\FFbatch\saved*" -force
Remove-Item "$env:LOCALAPPDATA\qBittorrent\*" -Recurse -Force

# ***************************************************************************************************************************************************Windows Defender Logs
Set-MpPreference -ScanPurgeItemsAfterDelay 1
Remove-Item "C:\ProgramData\Microsoft\Windows Defender\Scans\History\Service" -Recurse -Force

# ***************************************************************************************************************************************************Notification Cache
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Notifications\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\ActionCenterCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\PPBCompat*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\*" -Recurse -Force
freepace24 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared24 = $freespace23.freespace - $freespace24.freespace
write-output ("Cleared [{0:N2}" -f ($cleared24/1GB) + "] Gb of space.")

echo "............Done Cleaning Windows Hisotry & Cache Files"

# ***************************************************************************************************************************************************Previous Windows Installation
# takeown /F C:\Windows.old* /R /A /D Y
# cacls C:\Windows.old*.* /T /grant administrators:F
# Remove-Item "C:\Windows.old\*" -Recurse -Force
# rmdir /S /Q C:\Windows.old

# ***************************************************************************************************************************************************Windows temp
Remove-Item "$env:SYSTEMROOT\Temp\*" -Recurse -Force
Remove-Item "$env:SYSTEMROOT\Prefetch\*" -Recurse -Force
Remove-Item "$env:TEMP\*" -Recurse -Force
Remove-Item "$env:APPDATA\Temp\*" -Recurse -Force
Remove-Item "$env:USERPROFILE\AppData\LocalLow\Temp\*" -Recurse -Force
Remove-Item "C:\Windows\*\*.log" -Recurse -Force
Remove-Item "$env:SystemDrive\File*.chk" -Force
Remove-Item "C:\MATS" -Force -Recurse
Remove-Item "$env:LOCALAPPDATA\Temp\*" -Recurse -Force
Remove-Item "C:\Windows\System32\sru\*" -Recurse -Force
Remove-Item "C:\WINDOWS\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Logs\*" -Recurse -Force

$logpath = "C:\" 
echo "............Removing tmp,old and log files"
Get-ChildItem $logpath -recurse *.tmp -force | Remove-Item -force -Recurse
Get-ChildItem $logpath -recurse *.old -force | Remove-Item -force -Recurse
freepace25 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared25 = $freespace24.freespace - $freespace25.freespace
write-output ("Cleared [{0:N2}" -f ($cleared25/1GB) + "] Gb of space.")
echo "............Done Removing tmp,old and log files"

# ****************************************************************************************************************************************************Crash Dumps
Remove-Item "$env:USERPROFILE\AppData\Local\CrashDumps\*" -Recurse -Force

# ****************************************************************************************************************************************************Cinebench Cache
Remove-Item "$env:APPDATA\MAXON\*" -Recurse -Force
Remove-Item "$env:APPDATA\MAXON\Maxon\Cinebench2024_FD9B9CBE\Redshift\Cache\*" -Recurse -Force
freepace26 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared26 = $freespace25.freespace - $freespace26.freespace
write-output ("Cleared [{0:N2}" -f ($cleared26/1GB) + "] Gb of space.")
echo "............Done Cleaning Cinebench Cache Files"

# Log Files
$logpath = "C:\Windows\"
Get-ChildItem $logpath -recurse *.log -force | Remove-Item -Force -Recurse
Get-ChildItem $logpath -recurse *.dmp -force | Remove-Item -Force -Recurse
freepace27 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared27 = $freespace26.freespace - $freespace27.freespace
write-output ("Cleared [{0:N2}" -f ($cleared27/1GB) + "] Gb of space.")
echo "............Done Cleaning System Dump Files"
# ****************************************************************************************************************************************************Delete Elgato Camera Hub Logs
Remove-Item "$env:APPDATA\Elgato\CameraHub\logs\*" -Recurse -Force
Remove-Item "$env:APPDATA\Elgato\CameraHub\SW\*" -Recurse -Force
Remove-Item "$env:APPDATA\Elgato\CameraHub\Tmp\*" -Recurse -Force
freepace28 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared28 = $freespace27.freespace - $freespace28.freespace
write-output ("Cleared [{0:N2}" -f ($cleared28/1GB) + "] Gb of space.")
echo "............Done Cleaning Elgato Camera Hub Logs"
# ******************************************************************************************************************************************************Game Cache
taskkill /f /im "upc.exe"
taskkill /f /im "steam.exe"
taskkill /f /im "EpicGamesLauncher.exe"
# Remove-Item "C:\Program Files (x86)\Ubisoft\Ubisoft Game Launcher\cache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Steam\htmlcache\Cache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\EpicGamesLauncher\Saved\webcache_*\*" -Force -Recurse
Remove-Item "$env:LOCALAPPDATA\EpicGamesLauncher\Saved\webcache_*" -Force -Recurse
Remove-Item "$env:LOCALAPPDATA\EpicGamesLauncher\Saved\Crashes\*" -Force -Recurse
Remove-Item "$env:LOCALAPPDATA\UnrealDatasmithExporter\Saved\Crashes\*" -Force -Recurse
freepace29 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared29 = $freespace28.freespace - $freespace29.freespace
write-output ("Cleared [{0:N2}" -f ($cleared29/1GB) + "] Gb of space.")
echo "............Done Cleaning Game Cache Files"

# ******************************************************************************************************************************************************Nvidia Cache
remove-item "$env:HOMEPATH\AppData\Local\Nvidia\DXCache\*" -force -recurse
remove-item "$env:HOMEPATH\AppData\Local\Nvidia\GLCache\*" -force -recurse
freepace30 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared30 = $freespace29.freespace - $freespace30.freespace
write-output ("Cleared [{0:N2}" -f ($cleared30/1GB) + "] Gb of space.")
echo "............Done Cleaning Nvidia Cache Files"

# ****************************************************************************************************************************************************InstallShield Leftover
Remove-Item "C:\Program Files (x86)\InstallShield Installation Information\*" -Recurse -Force
# Remove-Item "C:\ProgramData\Package Cache\*" -Recurse -Force
freepace31 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared31 = $freespace30.freespace - $freespace31.freespace
write-output ("Cleared [{0:N2}" -f ($cleared31/1GB) + "] Gb of space.")
echo "............Done Cleaning Installshield Cache Files"

# ****************************************************************************************************************************************************Google Update Cache
Remove-Item "C:\Program Files (x86)\Google\Update\Download\*" -Recurse -Force
freepace32 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared32 = $freespace31.freespace - $freespace32.freespace
write-output ("Cleared [{0:N2}" -f ($cleared32/1GB) + "] Gb of space.")
echo "............Done Cleaning Google Update Cache Files"

# ****************************************************************************************************************************************************Teamviewer Cache
Remove-Item "$env:LOCALAPPDATA\TeamViewer\*" -Recurse -Force
freepace33 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared33 = $freespace32.freespace - $freespace33.freespace
write-output ("Cleared [{0:N2}" -f ($cleared33/1GB) + "] Gb of space.")
echo "............Done Cleaning Teamviewer Cache Files"

# ***************************************************************************************************************************************************Windows Memory Dump
Remove-Item "$env:SYSTEMROOT\memory.dmp" -Recurse -Force
freepace34 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared34 = $freespace33.freespace - $freespace34.freespace
write-output ("Cleared [{0:N2}" -f ($cleared34/1GB) + "] Gb of space.")
echo "............Done Cleaning Windows Memory Dump Files"
# ***************************************************************************************************************************************************VSS Shadow Copies
##vssadmin delete shadows /all /quiet
# echo "............Done Cleaning VSS Shadow Copies"

# ***************************************************************************************************************************************************Windows Update
net stop wuauserv
Remove-Item "C:\WINDOWS\SoftwareDistribution\Download" -Recurse -Force
#Windows.old
#Rollup Fix Backup
Remove-Item "C:\Windows\servicing\LCU\*.*" -Recurse -Force
freepace35 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared35 = $freespace34.freespace - $freespace35.freespace
write-output ("Cleared [{0:N2}" -f ($cleared35/1GB) + "] Gb of space.")
echo "............Done Cleaning Windows Update Temporary Files"
# ***************************************************************************************************************************************************Temporary Internet Files
remove-item "$env:HOMEPATH\AppData\Local\Microsoft\Windows\Temporary Internet Files\*" -force -recurse
# Google Chrome
remove-item "$env:HOMEPATH\AppData\Local\Google\Chrome\User Data\Profile 1\Cache\Cache_Data\*" -force -recurse
remove-item "$env:HOMEPATH\AppData\Local\Google\Chrome\User Data\Profile 1\Code Cache*" -force -recurse
remove-item "$env:HOMEPATH\AppData\Local\Google\Chrome\User Data\Profile 1\Cache\Cache_Data\*" -force -recurse
remove-item "$env:HOMEPATH\AppData\Local\Google\Chrome\User Data\Profile 2\Code Cache*" -force -recurse

# Microsoft Edge
remove-item "$env:HOMEPATH\AppData\Local\Microsoft\Edge\User Data\Default\Cache\Cache_Data\*" -force -recurse
remove-item "$env:HOMEPATH\AppData\Local\Microsoft\Edge\User Data\Profile 1\Cache\Cache_Data\*" -force -recurse
freepace36 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared36 = $freespace35.freespace - $freespace36.freespace
write-output ("Cleared [{0:N2}" -f ($cleared36/1GB) + "] Gb of space.")
echo "............Done Cleaning Temporary Browser Files"


# IE History Cleaned
RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 1
freepace37 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared37 = $freespace36.freespace - $freespace37.freespace
write-output ("Cleared [{0:N2}" -f ($cleared37/1GB) + "] Gb of space.")

echo "............Done Cleaning IE History Tracks"



# ***************************************************************************************************************************************************Windows Logs
remove-item "C:\Windows\Logs\NetSetup\*" -Force -Recurse
remove-item "C:\Windows\Logs\SIH\*" -Force -Recurse
remove-item "C:\Windows\Logs\WindowsUpdate\*" -Force -Recurse
remove-item "C:\ProgramData\USOShared\Logs\User\*" -Force -Recurse
#....Windows Error Report
Remove-Item "$env:LOCALAPPDATA\ElevatedDiagnostics\*" -Recurse -Force
freepace38 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared38 = $freespace37.freespace - $freespace38.freespace
write-output ("Cleared [{0:N2}" -f ($cleared38/1GB) + "] Gb of space.")
echo "............Done Cleaning Windows Logs"


# /      \|  \     |        \/      \|  \  |  \      |  \  _  |  |      |  \  |  |       \ /      \|  \  _  |  \/      \       |        |  \  |  |       \|  \     /      \|       \|        |       \ 
# |  $$$$$$| $$     | $$$$$$$|  $$$$$$| $$\ | $$      | $$ / \ | $$\$$$$$| $$\ | $| $$$$$$$|  $$$$$$| $$ / \ | $|  $$$$$$\      | $$$$$$$| $$  | $| $$$$$$$| $$    |  $$$$$$| $$$$$$$| $$$$$$$| $$$$$$$\
# | $$   \$| $$     | $$__   | $$__| $| $$$\| $$      | $$/  $\| $$ | $$ | $$$\| $| $$  | $| $$  | $| $$/  $\| $| $$___\$$      | $$__    \$$\/  $| $$__/ $| $$    | $$  | $| $$__| $| $$__   | $$__| $$
# | $$     | $$     | $$  \  | $$    $| $$$$\ $$      | $$  $$$\ $$ | $$ | $$$$\ $| $$  | $| $$  | $| $$  $$$\ $$\$$    \       | $$  \    >$$  $$| $$    $| $$    | $$  | $| $$    $| $$  \  | $$    $$
# | $$   __| $$     | $$$$$  | $$$$$$$| $$\$$ $$      | $$ $$\$$\$$ | $$ | $$\$$ $| $$  | $| $$  | $| $$ $$\$$\$$_\$$$$$$\      | $$$$$   /  $$$$\| $$$$$$$| $$    | $$  | $| $$$$$$$| $$$$$  | $$$$$$$\
# | $$__/  | $$_____| $$_____| $$  | $| $$ \$$$$      | $$$$  \$$$$_| $$_| $$ \$$$| $$__/ $| $$__/ $| $$$$  \$$$|  \__| $$      | $$_____|  $$ \$$| $$     | $$____| $$__/ $| $$  | $| $$_____| $$  | $$
#  \$$    $| $$     | $$     | $$  | $| $$  \$$$      | $$$    \$$|   $$ | $$  \$$| $$    $$\$$    $| $$$    \$$$\$$    $$      | $$     | $$  | $| $$     | $$     \$$    $| $$  | $| $$     | $$  | $$
#  \$$$$$$ \$$$$$$$$\$$$$$$$$\$$   \$$\$$   \$$       \$$      \$$\$$$$$$\$$   \$$\$$$$$$$  \$$$$$$ \$$      \$$ \$$$$$$        \$$$$$$$$\$$   \$$\$$      \$$$$$$$$\$$$$$$ \$$   \$$\$$$$$$$$\$$   \$$
                                                                                                                                                                                                      









# ***************************************************************************************************************************************************Windows Explorer

Write-Host "`nClearing . . . " -NoNewline

Get-ChildItem $env:APPDATA\Microsoft\Windows\Recent\* -File -Force -Exclude desktop.ini | 
Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem $env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\* -File -Force -Exclude desktop.ini, f01b4d95cf55d32a.automaticDestinations-ms| 
Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem $env:APPDATA\Microsoft\Windows\Recent\CustomDestinations\* -File -Force -Exclude desktop.ini | 
Remove-Item -Force -ErrorAction SilentlyContinue

# Clear unpinned folders from Quick Access, using the Verbs() method
# $UnpinnedQAFolders = (0,0)
# While ($UnpinnedQAFolders) {
#    $UnpinnedQAFolders = (((New-Object -ComObject Shell.Application).Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items() | 
#    where IsFolder -eq $true).Verbs() | where Name -match "Remove from Quick access")
#    If ($UnpinnedQAFolders) { $UnpinnedQAFolders.DoIt() }
# }

Write-Host "Done!`n"
Stop-Process -Name explorer -Force

Remove-Variable UnpinnedQAFolders

# ***************************************************************************************************************************************************Start Win 11 Privacy Protector
# Start "C:\Program Files\Yamicsoft\Windows 11 Manager\PrivacyProtector.exe" -ArgumentList auto


#Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\*" -Exclude "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\*" -Force

Remove-Item "C:\Program Files\Microsoft OneDrive\*" -Recurse -Force
Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\*.lnk" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\History" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\WebCache\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\*" -Recurse -Force
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Caches\*" -Recurse -Force
Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU" -Recurse -Force
Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\WordWheelQuery" -Recurse -Force
Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs" -Recurse -Force
Remove-Item "HKCU:\SOFTWARE\MPC-HC\MPC-HC\Recent File List\*" -Recurse -Force
Remove-Item "HKCU:\Software\MPC-HC\MPC-HC\MediaHistory\*" -Recurse -Force
Remove-Item "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store" -Recurse -Force
# MEDIA PLAYER CLASSIC
Remove-Item "HKCU:\SOFTWARE\MPC-HC\MPC-HC\Recent File List\*" -Recurse -Force
Remove-Item "HKCU:\SOFTWARE\MPC-HC\MPC-HC\Recent File List" -Force
Remove-Item "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths" -Recurse -Force
Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FeatureUsage" -Recurse -Force
Remove-Item "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache\*" -Force -Recurse
Remove-Item "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache" -Force -Exclude "(Default)", "Default"
Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\*" -Recurse -Force
Remove-Item "HKCU:\SYSTEM\ControlSet001\Control\Session Manager\AppCompatCache" -Recurse -Force
Remove-Item "HKCU:\SYSTEM\CurrentControlSet\Control\Session Manager\AppCompatCache" -Recurse -Force
Remove-Item "HKLM:\SYSTEM\ControlSet001\Control\Session Manager\AppCompatCache" -Recurse -Force
Remove-Item "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\AppCompatCache" -Recurse -Force
Remove-Item "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU\*" -Recurse -Force

start explorer.exe
freepace39 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared39 = $freespace38.freespace - $freespace39.freespace
write-output ("Cleared [{0:N2}" -f ($cleared39/1GB) + "] Gb of space.")
echo "............Done Cleaning Windows Explorer Hisotry Files"


#******************************************************************************************************************************************************Remove Trackers
Remove-Item "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache" -Recurse -Force
Remove-Item "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Persisted" -Recurse -Force
Remove-Item "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store" -Recurse -Force


# ***************************************************************************************************************************************************Event Viewer Logs
/F “tokens=*” %1 in ('wevtutil.exe el') DO wevtutil.exe cl “%1”

# ***************************************************************************************************************************************************Delete Temporary Folders
Remove-Item "$env:WINDIR\Temp" -Recurse -Force
Remove-Item "$env:WINDIR\Prefetch" -Recurse -Force
Remove-Item "$env:TEMP" -Recurse -Force
Remove-Item "$env:APPDATA\Temp" -Recurse -Force
Remove-Item "$env:HOMEPATH\AppData\LocalLow\Temp" -Recurse -Force

# ***************************************************************************************************************************************************Delete Driver Installation
Remove-Item "$env:SYSTEMDRIVE\AMD\*" -Recurse -Force
Remove-Item "$env:SYSTEMDRIVE\NVIDIA\*" -Recurse -Force
Remove-Item "$env:SYSTEMDRIVE\INTEL\*" -Recurse -Force

Remove-Item "$env:SYSTEMDRIVE\AMD" -Recurse -Force
Remove-Item "$env:SYSTEMDRIVE\NVIDIA" -Recurse -Force
Remove-Item "$env:SYSTEMDRIVE\INTEL" -Recurse -Force

# ***************************************************************************************************************************************************Create Temporary Folders
New-Item "$env:WINDIR\Temp" -ItemType Directory
New-Item "$env:WINDIR\Prefetch" -ItemType Directory
New-Item "$env:TEMP" -ItemType Directory
New-Item "$env:APPDATA\Temp" -ItemType Directory
New-Item "$env:HOMEPATH\AppData\LocalLow\Temp" -ItemType Directory

Remove-Item "$env:LOCALAPPDATA\Packages\5319275A.WhatsAppDesktop_cv1g1gvanyjgm\LocalCache\Roaming\WhatsApp\Service Worker\CacheStorage" -Recurse -Force
freepace40 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
cleared40 = $freespace39.freespace - $freespace40.freespace
write-output ("Cleared [{0:N2}" -f ($cleared40/1GB) + "] Gb of space.")
echo "............Done Cleaning Windows Temporary Files"

# ***************************************************************************************************************************************************Kill Onedrive
 # TASKKILL /f /im OneDrive.exe
 # Remove-item "HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
 # Remove-Item "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"

# ****************************************************************************************************************************************************Cleanup component Store in a new CMD window
Dism.exe /Online /Cleanup-Image /AnalyzeComponentStore
Dism /Online /Cleanup-Image /restoreHealth
Dism.exe /online /Cleanup-Image /StartComponentCleanup /ResetBase

freepace41 = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select -Object Freespace)
cleared41 = $freespace40.freespace - $freespace41.freespace
write-output ("Cleared [{0:N2}" -f ($cleared41/1GB) + "] Gb of space.")
echo "............Done Cleaning Component Store"

# ***************************************************************************************************************************************************Clean Old Drivers
# Use this PowerShell script to find and remove old and unused device drivers from the Windows Driver Store
# Explanation: http://woshub.com/how-to-remove-unused-drivers-from-driver-store/

$dismOut = dism /online /get-drivers
$Lines = $dismOut | select -Skip 10
$Operation = "theName"
$Drivers = @()
foreach ( $Line in $Lines ) {
    $tmp = $Line
    $txt = $($tmp.Split( ':' ))[1]
    switch ($Operation) {
        'theName' { $Name = $txt
                     $Operation = 'theFileName'
                     break
                   }
        'theFileName' { $FileName = $txt.Trim()
                         $Operation = 'theEntr'
                         break
                       }
        'theEntr' { $Entr = $txt.Trim()
                     $Operation = 'theClassName'
                     break
                   }
        'theClassName' { $ClassName = $txt.Trim()
                          $Operation = 'theVendor'
                          break
                        }
        'theVendor' { $Vendor = $txt.Trim()
                       $Operation = 'theDate'
                       break
                     }
        'theDate' { # we'll change the default date format for easy sorting
                     $tmp = $txt.split( '.' )
                     $txt = "$($tmp[2]).$($tmp[1]).$($tmp[0].Trim())"
                     $Date = $txt
                     $Operation = 'theVersion'
                     break
                   }
        'theVersion' { $Version = $txt.Trim()
                        $Operation = 'theNull'
                        $params = [ordered]@{ 'FileName' = $FileName
                                              'Vendor' = $Vendor
                                              'Date' = $Date
                                              'Name' = $Name
                                              'ClassName' = $ClassName
                                              'Version' = $Version
                                              'Entr' = $Entr
                                            }
                        $obj = New-Object -TypeName PSObject -Property $params
                        $Drivers += $obj
                        break
                      }
         'theNull' { $Operation = 'theName'
                      break
                     }
    }
}
$last = ''
$NotUnique = @()
foreach ( $Dr in $($Drivers | sort Filename) ) {
    if ($Dr.FileName -eq $last  ) {  $NotUnique += $Dr  }
    $last = $Dr.FileName
}
$NotUnique | sort FileName | ft
# search for duplicate drivers 
$list = $NotUnique | select -ExpandProperty FileName -Unique
$ToDel = @()
foreach ( $Dr in $list ) {
    Write-Host "duplicate driver found" -ForegroundColor Yellow
    $sel = $Drivers | where { $_.FileName -eq $Dr } | sort date -Descending | select -Skip 1
    $sel | ft
    $ToDel += $sel
}
Write-Host "List of driver version  to remove" -ForegroundColor Red
$ToDel | ft
# Removing old driver versions
# Uncomment the Invoke-Expression to automatically remove old versions of device drivers 
foreach ( $item in $ToDel ) {
    $Name = $($item.Name).Trim()
    Write-Host "deleting $Name" -ForegroundColor Yellow
   # Write-Host "pnputil.exe /remove-device  $Name" -ForegroundColor Yellow
    Invoke-Expression -Command "pnputil.exe /remove-device $Name"
}
echo "............Done Cleaning Old Drivers"

#.........................................................................................................................................
#.DDDDDDDDD...DIIII..SSSSSSS....SKKK...KKKKK........CCCCCCC....CLLL.......EEEEEEEEEEE....AAAAA.....ANNN...NNNN..NUUU...UUUU..UPPPPPPPP....
#.DDDDDDDDDD..DIIII.SSSSSSSSS...SKKK..KKKKK........CCCCCCCCC...CLLL.......EEEEEEEEEEE....AAAAA.....ANNNN..NNNN..NUUU...UUUU..UPPPPPPPPP...
#.DDDDDDDDDDD.DIIII.SSSSSSSSSS..SKKK.KKKKK........ CCCCCCCCCC..CLLL.......EEEEEEEEEEE...AAAAAA.....ANNNN..NNNN..NUUU...UUUU..UPPPPPPPPPP..
#.DDDD...DDDD.DIIIIISSSS..SSSS..SKKKKKKKK......... CCC...CCCCC.CLLL.......EEEE..........AAAAAAA....ANNNNN.NNNN..NUUU...UUUU..UPPP...PPPP..
#.DDDD....DDDDDIIIIISSSS........SKKKKKKK......... CC.....CCC..CLLL.......EEEE.........AAAAAAAA....ANNNNN.NNNN..NUUU...UUUU..UPPP...PPPP..
#.DDDD....DDDDDIIII.SSSSSSS.....SKKKKKKK......... CC..........CLLL.......EEEEEEEEEE...AAAAAAAA....ANNNNNNNNNN..NUUU...UUUU..UPPPPPPPPPP..
#.DDDD....DDDDDIIII..SSSSSSSSS..SKKKKKKK......... CC..........CLLL.......EEEEEEEEEE...AAAA.AAAA...ANNNNNNNNNN..NUUU...UUUU..UPPPPPPPPP...
#.DDDD....DDDDDIIII....SSSSSSS..SKKKKKKKK........ CC..........CLLL.......EEEEEEEEEE..EAAAAAAAAA...ANNNNNNNNNN..NUUU...UUUU..UPPPPPPPP....
#.DDDD....DDDDDIIII.......SSSSS.SKKK.KKKKK....... CC.....CCC..CLLL.......EEEE........EAAAAAAAAAA..ANNNNNNNNNN..NUUU...UUUU..UPPP.........
#.DDDD...DDDDDDIIIIISSS....SSSS.SKKK..KKKK........ CCC...CCCCC.CLLL.......EEEE........EAAAAAAAAAA..ANNN.NNNNNN..NUUU...UUUU..UPPP.........
#.DDDDDDDDDDD.DIIIIISSSSSSSSSSS.SKKK..KKKKK....... CCCCCCCCCC..CLLLLLLLLL.EEEEEEEEEEEEEAA....AAAA..ANNN..NNNNN..NUUUUUUUUUU..UPPP.........
#.DDDDDDDDDD..DIIII.SSSSSSSSSS..SKKK...KKKKK.......CCCCCCCCCC..CLLLLLLLLL.EEEEEEEEEEEEEAA.....AAAA.ANNN..NNNNN...UUUUUUUUU...UPPP.........
#.DDDDDDDDD...DIIII..SSSSSSSS...SKKK...KKKKK........CCCCCCC....CLLLLLLLLL.EEEEEEEEEEEEEAA.....AAAA.ANNN...NNNN....UUUUUUU....UPPP.........
#.........................................................................................................................................

    # Create Cleanmgr profile:
    write-output "Starting Disk Cleanup utility..."
    $ErrorActionPreference = "SilentlyContinue"
    $CleanMgrKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
    if (-not (get-itemproperty -path "$CleanMgrKey\Temporary Files" -name StateFlags1221))
    {
        set-itemproperty -path "$CleanMgrKey\Active Setup Temp Folders" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\BranchCache" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Downloaded Program Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Delivery Optimization Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Internet Cache Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Memory Dump Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Old ChkDsk Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Previous Installations" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Recycle Bin" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Service Pack Cleanup" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Setup Log Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\System error memory dump files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\System error minidump files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Temporary Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Temporary Setup Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Thumbnail Cache" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Update Cleanup" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Upgrade Discarded Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\User file versions" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Windows Defender" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Windows Error Reporting Archive Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Windows Error Reporting Queue Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Windows Error Reporting System Archive Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Windows Error Reporting System Queue Files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Windows ESD installation files" -name StateFlags1221 -type DWORD -Value 2
        set-itemproperty -path "$CleanMgrKey\Windows Upgrade Log Files" -name StateFlags1221 -type DWORD -Value 2
    }
    # run it:
    write-output "Starting Cleanmgr with full set of checkmarks (might take a while)..."
    $Process = (Start-Process -FilePath "$env:systemroot\system32\cleanmgr.exe" -ArgumentList "/sagerun:1221" -Wait -PassThru)
    write-output "Process ended with exitcode [$($Process.ExitCode)]."         

    write-output "Calculating disk usage on C:\..."
    $FreespaceAfter = (Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object Freespace)
    write-output ("Disk C:\ now has [{0:N2}" -f ($FreeSpaceAfter.freespace/1GB) + "] Gb available.")
    write-output ("Disk C:\ was [{0:N2}" -f ($FreeSpaceBefore.freespace/1GB) + "] Gb available.")
$ErrorActionPreference = "Continue"
write-output ("[{0:N2}" -f (($FreespaceAfter.Freespace-$FreespaceBefore.FreeSpace)/1GB) + "] Gb has been liberated on C:\.")    
    
        Write-Host "Press any key to End ..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")