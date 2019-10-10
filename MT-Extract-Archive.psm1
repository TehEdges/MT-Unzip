Function Get-RunspaceData {
    [cmdletbinding()]
    param(
        [switch]$Wait
    )
    Do {
        $more = $false         
        Foreach($runspace in $runspaces) {
            If ($runspace.Runspace.isCompleted) {
                $runspace.powershell.EndInvoke($runspace.Runspace)
                $runspace.powershell.dispose()
                $runspace.Runspace = $null
                $runspace.powershell = $null                 
            } ElseIf ($runspace.Runspace -ne $null) {
                $more = $true
            }
        }
        If ($more -AND $PSBoundParameters['Wait']) {
            Start-Sleep -Milliseconds 100
        }   
        #Clean out unused runspace jobs
        $temphash = $runspaces.clone()
        $temphash | Where {
            $_.runspace -eq $Null
        } | ForEach {
            Write-Verbose ("Removing {0}" -f $_.computer)
            $Runspaces.remove($_)
        }  
        #[console]::Title = ("Remaining Runspace Jobs: {0}" -f ((@($runspaces | Where {$_.Runspace -ne $Null}).Count)))
        UpdateProgress             
    } while ($more -AND $PSBoundParameters['Wait'])
}

Function GetEntries ($ZipFile)
{  
    $FS = New-Object System.IO.FileStream ($ZipFile,[System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    $ZipArchive = New-Object System.IO.Compression.ZipArchive -ArgumentList ($fs, [System.IO.Compression.ZipArchiveMode]::Read, $true)

    $Ret = $ZipArchive.Entries | Select FullName

    $fs.Close()
    $ZipArchive.Dispose()
    Return $Ret
}

Function UpdateProgress()
{
    $PercentComplete = ($Hash.CompletedEntries / $hash.TotalEntries) * 100
    Write-Progress -id 1 -Activity ("Unziping: " + $Hash.ZipFile) -PercentComplete $PercentComplete
}

Function Extract($ZipLocation, $Destination){
$ScriptBlock = {
    param($hash,
    $EntryStart,
    $EntryEnd)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO

    Function ReadAllBytes ($Reader)
    {
        $MS = New-Object System.IO.MemoryStream
        $Buffer = [byte[]]::CreateInstance([System.Byte],1024)
        $Reader.CopyTo($MS)
        Return $MS.ToArray()
    }

    Function GetSubEntries ($ZipFile, $EntryStart, $EntryEnd)
    {
        
        $FS = New-Object System.IO.FileStream ($ZipFile,[System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $ZipArchive = New-Object System.IO.Compression.ZipArchive -ArgumentList ($fs, [System.IO.Compression.ZipArchiveMode]::Read, $true)
        
        Return $($ZipArchive.Entries[$EntryStart..$EntryEnd] | Select FullName -ExpandProperty FullName)
    }

    $Entries = $(GetSubEntries -ZipFile $hash.ZipFile -EntryStart $EntryStart -EntryEnd $EntryEnd)

    $hash.add("Test: $EntryStart - $EntryEnd", $Entries) 
    
    Try
    {
        $fs = New-Object System.IO.FileStream -ArgumentList ($hash.ZipFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $Archive = New-Object System.IO.Compression.ZipArchive -ArgumentList ($FS, [System.IO.Compression.ZipArchiveMode]::Read, $true)
        
        Foreach($entry in $Entries)
        {
            [System.IO.Compression.ZipArchiveEntry]$ze = $Archive.GetEntry($entry)
            $es = $ze.Open()
            $MS =New-Object System.IO.MemoryStream
            $es.CopyTo($MS)
            $data = $Ms.ToArray()
            $dst = ($hash.destination + "\" + $ze.FullName.Replace("/", "\"))
            if($ze.Length -gt 0)
            {
                [System.io.file]::WriteAllBytes($dst, $data)
                $hash.CompletedEntries += 1
            }
        }
        $Archive.Dispose()
        $fs.Close()
        $fs.Dispose()
    }
    catch
    {
        $hash["Test: $EntryStart - $EntryEnd"] = $_.InvocationInfo.Line + ":" + $_.Exception.Message
    }
    #>
  
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Script:runspaces = New-Object System.Collections.ArrayList
$script:hash = [hashtable]::Synchronized(@{})
$hash.add("ZipFile", $ZipLocation)
$hash.add("Destination", $Destination)
$hash.add("TotalEntries", $(GetEntries $Hash.ZipFile).Count)
$hash.add("CompletedEntries", 0)
$sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$runspacepool = [runspacefactory]::CreateRunspacePool(1, 10, $sessionstate, $Host)
$runspacepool.Open()

$Directories = @()
Foreach($E in [System.IO.Compression.ZipFile]::OpenRead($hash.ZipFile).Entries)
{
    $Name = (Split-Path -Path $E.FullName -Parent)
    if($Name -ne "")
    {
        $Directories += $Name
    }
}

$Directories = $Directories | Select -Unique | Sort-Object

Foreach($Dir in $Directories)
{
    if($Dir -contains "\")
    {
        #SubFolder Create it
        $Name = $($Dir -split "\"| Select -Last 1)
        $Path = $Dir.SubString(0, $Dir.LastIndexOf("\"))
        New-Item -Path $Path -Name $Name -ItemType Directory | Out-Null
    }
    else
    {
        #Root Folder create it
        New-Item -Path $hash.Destination -Name $Dir -ItemType Directory | Out-Null
    }
}

$MaxThreads = (Get-CimInstance -ClassName win32_processor | Select NumberOfLogicalProcessors -ExpandProperty NumberOfLogicalProcessors)
$ThreadItemStart = 0
$TotalEntries = $(GetEntries -ZipFile $hash.ZipFile).count
$MaxThreadItems = [math]::Round($TotalEntries / $MaxThreads)

For ($i = 0; $I -lt $TotalEntries; $i = $I + $MaxThreadItems + 1) {
    $ThreadStart = $i
    $ThreadEnd = if(($ThreadStart + $MaxThreadItems) -le $TotalEntries){$ThreadStart + $MaxThreadItems}else{$TotalEntries}

    #Create the powershell instance and supply the scriptblock with the other parameters 
    $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($hash).AddArgument($ThreadStart).AddArgument($ThreadEnd)
           
    #Add the runspace into the powershell instance
    $powershell.RunspacePool = $runspacepool
           
    #Create a temporary collection for each runspace
    $temp = "" | Select-Object PowerShell,Runspace
    $temp.PowerShell = $powershell
           
    #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
    $temp.Runspace = $powershell.BeginInvoke()
    Write-Verbose ("Adding {0} collection" -f $temp.Computer)
    $runspaces.Add($temp) | Out-Null               
}
    Get-RunspaceData -Wait
}
