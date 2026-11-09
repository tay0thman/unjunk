#requires -Version 5.1
# Real SQLite transactions on disposable databases; never reads or writes real activity stores.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '..\Unjunk.ActivityHistory.psm1') -Force
$module=Get-Module Unjunk.ActivityHistory
& $module {Initialize-UnjunkActivitySqlite}
$fixture=Join-Path $PSScriptRoot ('ActivityFixture-'+[Guid]::NewGuid().ToString('N'))
$root=Join-Path $fixture 'ConnectedDevicesPlatform';$profile=Join-Path $root 'FixtureUser'
$null=New-Item -ItemType Directory -Path $profile -Force
$database=Join-Path $profile 'ActivitiesCache.db'
& $module {param($root) $script:activityFixtureRoot=$root;function script:Get-UnjunkActivityRoot {$script:activityFixtureRoot}} $root
$passed=0
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw "FAIL: $Message"};$script:passed++;Write-Host "PASS: $Message"}
$flags=[Reflection.BindingFlags]'NonPublic,Static'
$open=[Unjunk.ActivitySqlite].GetMethod('Open',$flags);$execute=[Unjunk.ActivitySqlite].GetMethod('Execute',$flags)
$scalarMethod=[Unjunk.ActivitySqlite].GetMethod('Scalar',$flags);$close=[Unjunk.ActivitySqlite].GetMethod('sqlite3_close',$flags)
function Sql([string]$Path,[string]$Text,[switch]$Scalar){
    $db=$open.Invoke($null,@($Path,$true))
    try{if($Scalar){$scalarMethod.Invoke($null,@($db,$Text))}else{$null=$execute.Invoke($null,@($db,$Text))}}finally{$null=$close.Invoke($null,@($db))}
}
try {
    [IO.File]::WriteAllBytes($database,[byte[]]@())
    foreach($statement in @('CREATE TABLE Activity(Id TEXT PRIMARY KEY,Payload TEXT)',
        'CREATE TABLE ActivityOperation(Id TEXT,Payload TEXT)',
        'CREATE TABLE Activity_PackageId(ActivityId TEXT,PackageName TEXT)',
        'CREATE TABLE Asset(Id TEXT,AssetPayload TEXT)',
        'CREATE TABLE AppSettings(AppId TEXT,SettingsPropertyBag TEXT)',
        'CREATE TABLE DataEncryptionKeys(KeyVersion INT,KeyValue TEXT)',
        "INSERT INTO Activity VALUES('a','private-document')", "INSERT INTO ActivityOperation VALUES('a','private-app')",
        "INSERT INTO Activity_PackageId VALUES('a','private-package')", "INSERT INTO Asset VALUES('a','private-asset')",
        "INSERT INTO AppSettings VALUES('a','keep-settings')", "INSERT INTO DataEncryptionKeys VALUES(1,'keep-key')")){Sql $database $statement}
    $inventory=Get-UnjunkActivityInventory
    Assert ($inventory.Databases.Count -eq 1 -and $inventory.Databases[0].Count -eq 4 -and -not $inventory.Issues.Count) 'Activity inventory counts all known history tables'
    Assert (($inventory|ConvertTo-Json -Depth 5) -notmatch 'private-|keep-key') 'Inventory never reads history payloads or encryption keys'
    Assert ((Get-UnjunkActivityInventory -Exclusions @($root)).Databases.Count -eq 0) 'Excluded activity roots are not opened'
    Assert ((Get-UnjunkActivityInventory -Exclusions @(($database+'-wal'))).Databases.Count -eq 0) 'An excluded database journal protects the whole database'
    $rejected=$false;try{Clear-UnjunkActivityDatabase -Path $database -Exclusions @($root)}catch{$rejected=$true}
    Assert $rejected 'Execution revalidates activity exclusions'
    Sql $database "CREATE TRIGGER BlockDelete BEFORE DELETE ON Activity BEGIN SELECT RAISE(ABORT,'fixture failure'); END"
    $rejected=$false;try{Clear-UnjunkActivityDatabase -Path $database}catch{$rejected=$true}
    Assert ($rejected -and [Unjunk.ActivitySqlite]::Count($database) -eq 4) 'A failed history delete rolls back all table changes'
    Sql $database 'DROP TRIGGER BlockDelete'
    $lock=$open.Invoke($null,@([string]$database,$true));$null=$execute.Invoke($null,@($lock,'BEGIN IMMEDIATE'))
    try{$rejected=$false;$watch=[Diagnostics.Stopwatch]::StartNew();try{Clear-UnjunkActivityDatabase -Path $database}catch{$rejected=$true};$watch.Stop()
        Assert ($rejected -and $watch.Elapsed.TotalSeconds -lt 5) 'A locked activity database fails promptly without hanging cleanup'
    }finally{$null=$execute.Invoke($null,@($lock,'ROLLBACK'));$null=$close.Invoke($null,@($lock))}
    $result=Clear-UnjunkActivityDatabase -Path $database
    Assert ($result.RemovedRecords -eq 4 -and [Unjunk.ActivitySqlite]::Count($database) -eq 0 -and $result.JournalCleared) 'History rows and journal are cleared with accurate counts'
    Assert ((Sql $database 'SELECT COUNT(*) FROM AppSettings' -Scalar) -eq 1 -and (Sql $database 'SELECT COUNT(*) FROM DataEncryptionKeys' -Scalar) -eq 1) 'Settings and encryption metadata survive activity cleanup'
    Sql $database 'PRAGMA journal_mode=WAL'
    Sql $database "INSERT INTO Activity VALUES('b','private-wal-record')"
    $result=Clear-UnjunkActivityDatabase -Path $database
    Assert ($result.RemovedRecords -eq 1 -and $result.JournalCleared -and (-not (Test-Path -LiteralPath ($database+'-wal')) -or (Get-Item -LiteralPath ($database+'-wal')).Length -eq 0)) 'WAL history is checkpointed and truncated after commit'
    $other=Join-Path $fixture 'ActivitiesCache.db';Copy-Item -LiteralPath $database -Destination $other
    $rejected=$false;try{Clear-UnjunkActivityDatabase -Path $other}catch{$rejected=$true}
    Assert $rejected 'A database outside the exact activity-store structure cannot be cleaned'
    Sql $database 'DROP TABLE ActivityOperation'
    $inventory=Get-UnjunkActivityInventory
    Assert ($inventory.Issues.Count -eq 1 -and $inventory.Databases.Count -eq 0) 'Unknown schemas are reported rather than treated as empty history'
} finally {
    $path=[IO.Path]::GetFullPath($fixture);$parent=[IO.Path]::GetFullPath($PSScriptRoot)+'\'
    if(-not $path.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $path -Leaf) -notlike 'ActivityFixture-*'){throw 'Unsafe activity fixture path'}
    Remove-Item -LiteralPath $path -Recurse -Force
}
Write-Host "All $passed activity history checks passed."
