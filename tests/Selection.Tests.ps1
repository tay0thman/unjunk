#requires -Version 5.1
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.Core.psm1') -Force
$passed=0
function Assert([bool]$Condition,[string]$Message) { if(-not $Condition){throw "FAIL: $Message"}; $script:passed++; Write-Host "PASS: $Message" }
function Row([string]$Name,[string]$Category,[string]$Risk='Cache') {
    [pscustomobject]@{Name=$Name;Category=$Category;Risk=$Risk;Selected=$false;CanSelect=$true;PolicyBlocked=$false;FileCount=2;Issues=@();RunningApps=@()}
}
$rows=@((Row 'Edge' 'Browsers'),(Row 'Chrome' 'Browsers'),(Row 'Claude' 'Desktop apps'),(Row 'Old plugin' 'Rhino plug-ins' 'Review'),(Row 'Inventory' 'Rhino plug-ins' 'Review'),(Row 'Blocked' 'Browsers'),(Row 'Running' 'Browsers'),(Row 'Access error' 'Browsers'))
$rows[4].CanSelect=$false; $rows[4].FileCount=0; $rows[5].PolicyBlocked=$true; $rows[6].RunningApps=@('browser'); $rows[7].Issues=@('Access denied')
Set-UnjunkCategorySelection -Rows $rows -Category Browsers
Assert ($rows[0].Selected -and $rows[1].Selected -and -not $rows[2].Selected) 'Category selection changes only the requested category'
Assert (-not $rows[5].Selected -and -not $rows[6].Selected -and -not $rows[7].Selected) 'Category selection skips blocked, running and access-error targets'
Assert ((Get-UnjunkCategoryState $rows Browsers).Checked -eq $true) 'Fully selected eligible category is checked'
$rows[0].Selected=$false
Assert ($null -eq (Get-UnjunkCategoryState $rows Browsers).Checked) 'Individual target deselection produces a partial category state'
Set-UnjunkCategorySelection -Rows $rows -Category 'Desktop apps'
Assert ($rows[1].Selected -and $rows[2].Selected) 'Picking another category preserves earlier selections'
Set-UnjunkCategorySelection -Rows $rows -Category Browsers -Selected $false
Assert (-not $rows[0].Selected -and -not $rows[1].Selected -and $rows[2].Selected) 'Clearing one category preserves other categories'
Set-UnjunkCategorySelection -Rows $rows -Category 'Rhino plug-ins'
Assert ($rows[3].Selected -and -not $rows[4].Selected) 'Explicit plug-in category choice selects review targets but never its inventory row'
Set-UnjunkCategorySelection -Rows $rows -Category 'All categories' -Selected $false
Assert (@($rows | Where-Object Selected).Count -eq 0) 'All-categories deselection clears all rows'
Set-UnjunkCategorySelection -Rows $rows -QueueOnReboot
Assert ($rows[6].Selected -and -not $rows[3].Selected) 'Reboot category selection allows running caches and excludes plug-in review targets'
Assert (-not (Get-UnjunkCategoryState @() Browsers).CanToggle) 'Empty categories cannot be toggled'
Assert (-not (Get-UnjunkCategoryState @($rows[4]) 'Rhino plug-ins').CanToggle) 'Inventory-only category cannot be toggled'
Write-Host "All $passed selection checks passed."
