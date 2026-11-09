#requires -Version 5.1
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot 'Unjunk.Core.psm1')

function Initialize-UnjunkActivitySqlite {
    if('Unjunk.ActivitySqlite' -as [type]){return}
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
[assembly: DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
namespace Unjunk {
    public sealed class ActivityClearResult {public long RemovedRecords;public bool JournalCleared;}
    public static class ActivitySqlite {
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern int sqlite3_open_v2(byte[] path,out IntPtr db,int flags,IntPtr vfs);
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern int sqlite3_close(IntPtr db);
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern int sqlite3_busy_timeout(IntPtr db,int ms);
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern int sqlite3_prepare_v2(IntPtr db,byte[] sql,int length,out IntPtr statement,IntPtr tail);
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern int sqlite3_step(IntPtr statement);
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern int sqlite3_finalize(IntPtr statement);
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern long sqlite3_column_int64(IntPtr statement,int column);
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern IntPtr sqlite3_column_text(IntPtr statement,int column);
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern int sqlite3_column_bytes(IntPtr statement,int column);
        [DllImport("winsqlite3.dll",CallingConvention=CallingConvention.Cdecl)] static extern int sqlite3_changes(IntPtr db);
        static byte[] Utf8(string text){return Encoding.UTF8.GetBytes(text+"\0");}
        static void Check(int result){if(result!=0)throw new IOException("Windows activity database unavailable (SQLite code "+result+").");}
        static IntPtr Open(string path,bool write){
            if(!File.Exists(path))throw new IOException("Activity database no longer exists.");
            IntPtr db;int code=sqlite3_open_v2(Utf8(path),out db,(write?2:1)|0x10000,IntPtr.Zero);
            if(code!=0){if(db!=IntPtr.Zero)sqlite3_close(db);Check(code);}
            sqlite3_busy_timeout(db,1500);return db;
        }
        static IntPtr Prepare(IntPtr db,string sql){IntPtr statement;Check(sqlite3_prepare_v2(db,Utf8(sql),-1,out statement,IntPtr.Zero));return statement;}
        static long Scalar(IntPtr db,string sql){
            IntPtr s=Prepare(db,sql);
            try{int code=sqlite3_step(s);if(code!=100){if(code==101)throw new IOException("Activity schema is incomplete.");Check(code);}return sqlite3_column_int64(s,0);}
            finally{sqlite3_finalize(s);}
        }
        static void Execute(IntPtr db,string sql){
            IntPtr s=Prepare(db,sql);
            try{int code;do{code=sqlite3_step(s);}while(code==100);if(code!=101)Check(code);}
            finally{sqlite3_finalize(s);}
        }
        static string[] Tables(IntPtr db){
            List<string> names=new List<string>();IntPtr s=Prepare(db,"SELECT name FROM sqlite_master WHERE type='table' ORDER BY name");
            try{int code;while((code=sqlite3_step(s))==100){int length=sqlite3_column_bytes(s,0);byte[] bytes=new byte[length];Marshal.Copy(sqlite3_column_text(s,0),bytes,0,length);names.Add(Encoding.UTF8.GetString(bytes));}if(code!=101)Check(code);}
            finally{sqlite3_finalize(s);}return names.ToArray();
        }
        public static string[] InspectTables(string path){IntPtr db=Open(path,false);try{return Tables(db);}finally{sqlite3_close(db);}}
        public static string[] InspectSchema(string path){
            IntPtr db=Open(path,false);var result=new List<string>();
            try{IntPtr s=Prepare(db,"SELECT sql FROM sqlite_master WHERE type IN ('table','trigger') ORDER BY name");
                try{int code;while((code=sqlite3_step(s))==100){int n=sqlite3_column_bytes(s,0);byte[] b=new byte[n];if(n>0)Marshal.Copy(sqlite3_column_text(s,0),b,0,n);result.Add(Encoding.UTF8.GetString(b));}if(code!=101)Check(code);}finally{sqlite3_finalize(s);}
                return result.ToArray();
            }finally{sqlite3_close(db);}
        }
        static string[] HistoryTables(IntPtr db){
            var present=new HashSet<string>(Tables(db),StringComparer.Ordinal);
            if(!present.Contains("Activity") || !present.Contains("ActivityOperation"))throw new IOException("Unrecognized Windows activity-history schema; left unchanged.");
            // Order children before Activity. Settings, metadata and unknown tables are untouched.
            var tables=new List<string>();
            foreach(string table in new[]{"ActivityOperation","Activity_PackageId","Asset","Activity"})if(present.Contains(table))tables.Add(table);
            return tables.ToArray();
        }
        public static long Count(string path){
            IntPtr db=Open(path,false);try{long count=0;foreach(string table in HistoryTables(db))count+=Scalar(db,"SELECT COUNT(*) FROM \""+table+"\"");return count;}finally{sqlite3_close(db);}
        }
        public static ActivityClearResult Clear(string path){
            IntPtr db=Open(path,true);bool transaction=false;
            try{
                Execute(db,"PRAGMA foreign_keys=ON");Execute(db,"PRAGMA secure_delete=ON");
                Execute(db,"BEGIN IMMEDIATE");transaction=true;long removed=0;
                foreach(string table in HistoryTables(db)){Execute(db,"DELETE FROM \""+table+"\"");removed+=sqlite3_changes(db);}
                Execute(db,"COMMIT");transaction=false;
                var result=new ActivityClearResult{RemovedRecords=removed,JournalCleared=false};
                try{result.JournalCleared=Scalar(db,"PRAGMA wal_checkpoint(TRUNCATE)")==0;}catch{}
                return result;
            }finally{
                if(transaction){try{Execute(db,"ROLLBACK");}catch{}}
                sqlite3_close(db);
            }
        }
    }
}
'@ -ErrorAction Stop
}

function Get-UnjunkActivityRoot {Join-Path $env:LOCALAPPDATA 'ConnectedDevicesPlatform'}

function Test-UnjunkActivityPath([string]$Path,[string[]]$Exclusions=@()) {
    $core=Get-Module Unjunk.Core
    & $core {param($path,$root,$exclusions)
        $root=Assert-UnjunkRoot $root;$path=Get-UnjunkFullPath $path
        if([IO.Path]::GetFileName($path) -ne 'ActivitiesCache.db' -or -not (Test-UnjunkWithin $path $root) -or
            [IO.Path]::GetDirectoryName([IO.Path]::GetDirectoryName($path)) -ne $root){return $false}
        foreach($candidate in @($path,($path+'-wal'),($path+'-shm'),($path+'-journal'))){
            if(Test-UnjunkExcluded $candidate $exclusions){return $false}
            if((Test-Path -LiteralPath $candidate) -and -not (Test-UnjunkNoLinks $candidate)){return $false}
        }
        return (Test-Path -LiteralPath $path -PathType Leaf)
    } $Path (Get-UnjunkActivityRoot) $Exclusions
}

function Get-UnjunkActivityInventory {
    param([string[]]$Exclusions=@())
    $paths=New-Object 'System.Collections.Generic.List[object]';$issues=@();$root=Get-UnjunkActivityRoot
    try {
        if(Test-Path -LiteralPath $root){
            $core=Get-Module Unjunk.Core
            if(-not (& $core {param($root) Test-UnjunkNoLinks $root} $root)){throw 'Linked activity store.'}
            foreach($folder in @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction Stop)){
                $path=Join-Path $folder.FullName 'ActivitiesCache.db'
                if(-not (Test-Path -LiteralPath $path -PathType Leaf)){continue}
                if(-not (Test-UnjunkActivityPath $path $Exclusions)){continue}
                try{Initialize-UnjunkActivitySqlite;$count=[Unjunk.ActivitySqlite]::Count($path);$paths.Add([pscustomobject]@{Path=$path;Count=$count})}
                catch{$issues+=@('An activity database was inaccessible or had an unsupported schema.')}
            }
        }
    } catch {$issues+=@('Windows activity-history inventory was unavailable.')}
    [pscustomobject]@{Databases=$paths.ToArray();Issues=@($issues | Sort-Object -Unique)}
}

function Clear-UnjunkActivityDatabase {
    param([string]$Path,[string[]]$Exclusions=@())
    if(-not (Test-UnjunkActivityPath $Path $Exclusions)){throw 'Activity database changed, was excluded or is outside its allowed location.'}
    Initialize-UnjunkActivitySqlite
    [Unjunk.ActivitySqlite]::Clear($Path)
}

Export-ModuleMember -Function Get-UnjunkActivityInventory,Clear-UnjunkActivityDatabase
