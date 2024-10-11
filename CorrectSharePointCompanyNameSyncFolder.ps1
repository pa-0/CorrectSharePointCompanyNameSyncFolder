#This script will update all a user's sync'd folders to be within the same, latest company folder.
#There is no way I've found to force the personal OneDrive folder name to update.
#Does not require admin rights.  Intended for manual execution.
#All old folders will still exist, in case you need to revert for some reason.  The appropriate registry keys will be exported to their corresponding folders.

#Must be in chronological order
$companyNameHistory = @"
Information Technology Strategies LLC
Information Technology Strategies, Inc
Information Technology Strategies
"@.Split("`n") | foreach {$_.trim()}

$companyName = "Information Technology Strategies"
$workDir = $env:USERPROFILE

<#
$DestFldrName = "Information Technology Strategies, Inc"
if (Test-Path "$workdir\$companyName") {
    $DestFldrName = $companyName
    }
#>

#Ensure we use the latest company folder that already exists on the machine
foreach ($cname in $companyNameHistory) {
    if (Test-Path "$workdir\$cname") {
        $DestFldrName = $cname
        }
    else {break}
    }

Write-Host "`n`n====================================================================================================="
Write-Host "Your SharePoint folders will be sync'd under `"" -NoNewline -ForegroundColor Green
Write-Host $DestFldrName -NoNewline -ForegroundColor Magenta
Write-Host "`"" -ForegroundColor Green
Write-Host "=====================================================================================================`n`n"

#Get all folders with the company name, excluding the one we want to use
[array]$folders = Get-ChildItem -Path $workDir -Filter "$companyName*" | where {$_.name -NE $DestFldrName}
[array]$folderNames = $folders | select -ExpandProperty name | sort -Descending

#Recursive function to ignore online only files
function copyLocalFiles {
    param (
        $sourceFolder,
        $destFolder,
        $recursePath = $null
        )
    <#Testing
    $sourceFolder = $f.FullName
    $destFolder = "$workDir\$DestFldrName"
    $recursePath = $null
    #>

    #Input validation
    if ($sourceFolder -eq $null) {
        Write-Error "ERROR: Source Folder not provided"
        return
        }
    if ($destFolder -eq $null) {
        Write-Error "ERROR: destination Folder not provided"
        return
        }
    
    #Ensure good starting state
    $sourceFolder = $sourceFolder.trimEnd("\")
    $destFolder = $destFolder.trimEnd("\")
    
    #Get all items from source, no recurse
    $copyItems = Get-ChildItem $($sourceFolder + "\" + $recursePath).TrimEnd("\") -Attributes !o

    :copyLoop foreach ($item in $copyItems) {
        #Skip the reg files that may have been backed up from previous runs
        if ($recursePath -eq $path -and ($item.name -eq "CLSIDx32.reg" -or $item.name -eq "CLSIDx64.reg" -or $item.name -eq "NamespaceOneDrive.reg")) {
            continue copyLoop
            }

        :itemType switch -Regex ($item.Mode) {
            "d" {
        #Write-Host "Making folder '$destFolder$recursePath\$($item.name)', and recursing"
                New-Item "$destFolder$recursePath" -Name $item.name -ItemType Directory -Force -EA SilentlyContinue | Out-Null
                copyLocalFiles $sourceFolder $destFolder $($recursePath + "\" + $item.name)
                break itemType
                }
            "a" {
        #Write-Host "Copying file $($item.fullname)"
                Copy-Item $item.fullname -Destination "$destFolder$recursePath" -EA SilentlyContinue
                break itemType
                }
            default {
                Write-Error "ERROR: Unhandled File Type '$($item.attributes)' for file '$($item.fullname)'"
                pause
                }
            }
        }
    }

#Get UserID
#Now unnecessary
Remove-Variable currUserID -EA SilentlyContinue
while ($true) {
    $IDFromReg = (Get-ChildItem -Path "REGISTRY::HKEY_CURRENT_USER\Software\Policies\Microsoft" -EA SilentlyContinue | where {$_.name -Match "[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{8}"} | select -ExpandProperty PSChildName)
    if ($IDFromReg) {$currUserID = $IDFromReg.split("_")[0]}
    if ($currUserID) {break}
    $currUserID = Get-ItemPropertyValue -Path "REGISTRY::HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\Business1" -name cid
    if ($currUserID) {break}
    $currUserID = Get-ItemPropertyValue -Path "REGISTRY::HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\Business1" -name OneAuthAccountId
    if ($currUserID) {break}
    $currUserID = Get-ItemPropertyValue -Path "REGISTRY::HKEY_CURRENT_USER\Software\Microsoft\Edge\Profiles\Default" -Name OID
    if ($currUserID) {break}
    $currUserID = (Get-ItemPropertyValue -Path "REGISTRY::HKEY_CURRENT_USER\Software\Microsoft\IdentityCRL\TokenBroker\DefaultAccount" -name accountID).split(":")[1]
    break
    }

#Stop Onedrive
#C:\Users\ChrisSteele\AppData\Local\Microsoft\OneDrive
$oneDriveExePath = "C:\Program Files\Microsoft OneDrive\OneDrive.exe"
if (!(Test-Path $oneDriveExePath -EA SilentlyContinue)) {
    $oneDriveExePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" 
    }
&$oneDriveExePath /shutdown

#Loop until process exit
while (Get-Process -Name OneDrive -EA SilentlyContinue) {
    Write-Host "Checking in 3 seconds for OneDrive's closure"
    Start-Sleep -Seconds 3
    }
Write-Host "OneDrive Stopped" -ForegroundColor Green

#Get INI file.  Is always named a GUID, but that GUID is not guaranteed to be the user's Azure UserID
#    %LocalAPPDATA%\Microsoft\OneDrive\settings\Business1\<currUserID>.ini
#    C:\Users\ChrisSteele\AppData\Local\Microsoft\OneDrive\settings
#Make sure its for ITS by looking in each Business folder, and check the ini contents for our company name
Remove-Variable iniPath,iniContents,success -EA SilentlyContinue
$bidnessFolders = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\OneDrive\settings" -Filter "Business*" 
foreach ($b in $bidnessFolders) {
    $iniPath = "$($b.fullname)\$($currUserID).ini"
    #If the user's guid isn't the ini's name, use the last edited file of all <guid>.ini files
    if (!(Test-Path $iniPath)) {
        $iniPath = Get-ChildItem $b.FullName -Filter "*.ini" | where name -match '^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{8}' | sort LastWriteTime -Descending | select -First 1 -ExpandProperty fullname
        }
    #grab contents of ini file so we can look for our company name
    $iniContents = Get-Content $iniPath -Encoding Unicode
    if ($iniContents -match $companyName) {
        Write-Host "Found INI file for company '$companyName'"  -ForegroundColor Green
        $success = $true;break}
    }

#if not found, don't continue.
if (!$success) {
    Write-Host -ForegroundColor Red "FATAL ERROR: Cannot find appropriate INI file for '$companyName'."
    Read-Host -Prompt "Please contact chrissteele@it-strat.com for assistance.  Press enter to exit this script." | Out-Null
    exit
    }

#Grab the personal folder path from the ini file
$OneDrivePath = $iniContents[0].Split('"')[9]

#Replace ITS names for desired name
#replace personal library location
##DO NOT DO.  Anything using a direct path will break, and will disconnect from the QuickStart links in explorer  
#Can be commented in later, but can't be undone once run
#$iniContents[0] = $iniContents[0].Replace($OneDrivePath,"$($env:USERPROFILE + "\" + "OneDrive - $DestFldrName")")
#replace sync'd folder locations
for ($i = 1; $i -lt $iniContents.Count; $i++) {
    foreach ($fName in $folderNames) {
        $iniContents[$i] = $iniContents[$i].replace("\$fName\","\$DestFldrName\")
        }
    }

#Make backup of ini
$timestamp = get-date -Format "yyyyMMdd_hhmmss"
Rename-Item $iniPath -NewName "$($iniPath | Split-Path -Leaf)_$timestamp`.bak"
#Output back to file
Out-File -InputObject $iniContents -FilePath $iniPath -Encoding unicode
Write-Host "Updated INI file" -ForegroundColor Green

#Make OneDrive folder with new company name, if it doesn't already exist
#You dont get to control the displayname, just the folder it sits in
#$destOneDrive = "OneDrive - $DestFldrName"
$destOneDrive = $OneDrivePath | Split-Path -Leaf
<#better yet, leave it where it lies.  Moving the folder breaks the QuickStart links for documents and whatnot.
if (!(Test-Path "$workDir\$destOneDrive")) {
    Write-Host "Creating new OneDrive personal folder" -ForegroundColor Green
    New-Item -Path $workDir -Name $destOneDrive -ItemType Directory -EA SilentlyContinue | Out-Null

    #Copy old OneDrive folder into new
    #Get-ChildItem -Path $OneDrivePath | Copy-Item -Destination "$workDir\$destOneDrive" -Recurse -Force
    copyLocalFiles $OneDrivePath "$workDir\$destOneDrive"
    Write-Host "Files copied to new OneDrive personal folder" -ForegroundColor Green
    }
    #>

#Make desired Sync folder name
New-Item -Path $workdir -Name $DestFldrName -ItemType Directory -EA SilentlyContinue | Out-Null
#Copy all other company folders into new one
foreach ($f in $folders) {
    Write-Host "Migrating '$($f.name)' into '$workDir\$DestFldrName'"
    #Get-ChildItem -Path $f.FullName -Attributes !o | Copy-Item -Destination "$workDir\$DestFldrName" -Recurse -Force
    #robocopy $f.FullName "$workDir\$DestFldrName" /s /z
    copyLocalFiles $f.FullName "$workDir\$DestFldrName"
    }
Write-Host "SharePoint sync'd files copied to $DestFldrName" -ForegroundColor Green

#Cleanup registry.  This part removes the company folders from the explorer side pane
#region registry
#CLSID x64
Remove-Variable subKeysCLSID64,subKeysCLSID32,subKeysNamespace -EA SilentlyContinue
$subKeysCLSID64 = Get-ChildItem "REGISTRY::HKEY_CURRENT_USER\Software\Classes\CLSID" | where {$_.Property.Contains("SortOrderIndex")}
foreach ($key in $subKeysCLSID64) {
    Remove-Variable value,folderpath -EA SilentlyContinue
    $value = (Get-ItemProperty "REGISTRY::$($key.name)").'(default)'
    if ($value -eq $DestFldrName -or $value -eq $destOneDrive) {
        Write-Host "Found key for desired folder $value; skipping" -ForegroundColor Cyan
        continue
        }
    #$folderPath = $folders | where name -eq $value | select -ExpandProperty fullname
    #if ($folderPath -eq $null) {write-host "Line 163";exit}
    $folderPath = $env:USERPROFILE + "\" + $value
    if (!(Test-Path $folderPath)) {write-host "Line 166 - '$folderPath' not found";pause}
    Write-Host "exporting $($key.name) to $("$($folderPath)\CLSIDx64.reg")"
    Remove-Item $("$($folderPath)\CLSIDx64.reg") -Force -Confirm:$false -EA SilentlyContinue
    reg export $($key.name) $("$($folderPath)\CLSIDx64.reg")
    Remove-Item "REGISTRY::$($key.name)" -Recurse -Force -Confirm:$false
    }

#CLSID x32
$subKeysCLSID32 = Get-ChildItem "REGISTRY::HKEY_CURRENT_USER\Software\Classes\WOW6432Node\CLSID" | where {$_.Property.Contains("SortOrderIndex")}
foreach ($key in $subKeysCLSID32) {
    Remove-Variable value,folderpath -EA SilentlyContinue
    $value = (Get-ItemProperty "REGISTRY::$($key.name)").'(default)'
    if ($value -eq $DestFldrName -or $value -eq $destOneDrive) {
        Write-Host "Found key for desired folder $value; skipping" -ForegroundColor Cyan
        continue
        }
    #$folderPath = $folders | where name -eq $value | select -ExpandProperty fullname
    #if ($folderPath -eq $null) {write-host "Line 181";exit}
    $folderPath = $env:USERPROFILE + "\" + $value
    if (!(Test-Path $folderPath)) {write-host "Line 185 - '$folderPath' not found";pause}
    Write-Host "exporting $($key.name) to $("$($folderPath)\CLSIDx32.reg")"
    Remove-Item $("$($folderPath)\CLSIDx32.reg") -Force -Confirm:$false -EA SilentlyContinue
    reg export $($key.name) $("$($folderPath)\CLSIDx32.reg")
    Remove-Item "REGISTRY::$($key.name)" -Recurse -Force -Confirm:$false
    }

#Delete namespace regkey for old company folders
#    Computer\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{<guid>}:(Default) -eq "company name"
$subKeysNamespace = Get-ChildItem "REGISTRY::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace"
foreach ($key in $subKeysNamespace) {
    Remove-Variable value,folderpath -EA SilentlyContinue
    $value = (Get-ItemProperty "REGISTRY::$($key.name)").'(default)'
    if ($value -eq $DestFldrName -or $value -eq $destOneDrive) {
        Write-Host "Found key for desired folder $value; skipping" -ForegroundColor Cyan
        continue
        }
    #$folderPath = $folders | where name -eq $value | select -ExpandProperty fullname
    $folderPath = $env:USERPROFILE + "\" + $value
    if (!(Test-Path $folderPath)) {write-host "Line 213 - '$folderPath' not found";pause}
    Write-Host "exporting $($key.name) to $("$($folderPath)\NamespaceOneDrive.reg")"
    Remove-Item $("$($folderPath)\NamespaceOneDrive.reg") -Force -Confirm:$false -EA SilentlyContinue
    reg export $($key.name) $("$($folderPath)\NamespaceOneDrive.reg")
    Remove-Item "REGISTRY::$($key.name)" -Recurse -Force -Confirm:$false
    }
Write-Host "Registry keys exported, then deleted" -ForegroundColor Green
#endregion

<#Not necessary
#Map Registry
$key = Get-ItemProperty "REGISTRY::HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\$($b.name)\ScopeIdToMountPointPathCache"
#For each sync'd folder already listed, update to new path
#redundant, since we're setting it later as well
$keyProps = $key | get-member -MemberType NoteProperty | where name -NotMatch '^PS[PC]'
foreach ($property in $keyProps) {
    
    }

#add new entries for all in ini file
$iniContents | select -Skip 1 | where {$_ -match "\\$DestFldrName\\"}
#>

#Start OneDrive
&$oneDriveExePath /start
Write-Host "`n`nOneDrive is now resyncing.  Please allow it time to update." -ForegroundColor Green

Write-Host "`n`n====================================================================================================="
Write-Host "Your SharePoint folders will be sync'd under `"" -NoNewline -ForegroundColor Green
Write-Host $DestFldrName -NoNewline -ForegroundColor Magenta
Write-Host "`"" -ForegroundColor Green
Write-Host "=====================================================================================================`n`n"

Write-Host "You can check its progress by clicking the OneDrive cloud icon in your system tray."
Write-Host "If you see pop-up errors from OneDrive stating `"[...] is not your original [...] folder`", this is expected, and just click `"Try Again`"."
Write-Host "(You can use the Space or Enter keys to close the pop-ups.  It wont even take a minute, I promise.)"
Read-Host "You may press enter at any time to finish this script" | Out-Null