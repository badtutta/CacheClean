<#


               Name:    CacheClean.ps1                                                               
                                                                                                   
          Author(s):    Navy Flank Speed MCS                                                                                                                                                 
                                                                                                       
            Version:    1.0.4                                                                        
                                                                                                  
            Created:    01.18.2023                                                                 
           Modified:    12.21.2023                                                                  
                                                                                                  
                                                                                                  
                                                                                                  
          Info:                                                                                   
                                                                                                  
          This script can determine the size of the user's profile and be used to perform a disk cleanup 
          to free up space via CleanMgr.exe, as well as, backup of Microsoft Edge favorites to an html file.
		  It will all perform the cleaning of the following caches...
		  
		    -  PowerBI
			-  Teams
			-  Outlook
			-  OneDrive

          

          
          Parameters:
    
          -ProfileSize = Determines disk space used by the user's profile
		  -Clean = Cleans with default options
		  -Options = Opens GUI and allows selection of more options to clean
		  -MSTeams = Cleans Microsoft Teams cache
		  -OneDrive = Cleans Microsoft OneDrive cache
		  -PowerBI = Cleans Microsoft PowerBI cache
		  -Outlook = Cleans Microsoft Outlook cache
		  -Edge = Backs up Edge Favorites to location of user's choosing
		  
		  
          

          Version Log:

          .1 - Initial script
		  .2 - Added relaunch of apps after cache cleaning
		  .3 - Added Edge Favorites backup ability
		  .4 - Added support for new Teams version
                        

#>
 

Param (
    [Parameter(Mandatory=$false)] [switch]$ProfileSize = $false,
	[Parameter(Mandatory=$false)] [switch]$Clean = $false,
	[Parameter(Mandatory=$false)] [switch]$Options = $false,
	[Parameter(Mandatory=$false)] [switch]$MSTeams = $false,
	[Parameter(Mandatory=$false)] [switch]$OneDrive = $false,
	[Parameter(Mandatory=$false)] [switch]$PowerBI = $false,
	[Parameter(Mandatory=$false)] [switch]$Outlook = $false,
	[Parameter(Mandatory=$false)] [switch]$Edge = $false,
	[Parameter(Mandatory=$false)] [switch]$Default = $false	
)



###############################
# Determine profile sizes in MB
###############################
If ($ProfileSize -eq $true){
	$profile = $Env:USERNAME
	$largeprofile = Get-ChildItem C:\Users\$profile -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Sum length | Select -ExpandProperty Sum
	$largeprofile = [math]::Round(($largeprofile/1MB),2)
	If ($largeprofile -lt 20){
		Continue}
	$object = New-Object -TypeName PSObject
	$object | Add-Member -MemberType NoteProperty -Name Name -Value $profile
	$object | Add-Member -MemberType NoteProperty -Name "Size(MB)" -Value $largeprofile
	($object | fl | Out-String).Trim();Write-Output "`n"
}


#######################################
# Clean user profile using CleanMgr.exe
#######################################
If ($Clean -eq $true){
	# Clean disk using CleanMgr with custom selected settings via GUI
	If ($Options -eq $true){
		$arg1 = "/D" 
		$arg2 = $Env:SystemDrive
		
		& "CleanMgr.exe" $arg1 $arg2 | Out-Null 
	}Else{
		# Clean disk using CleanMgr with default settings
		$arg1 = "/VERYLOWDISK"
	
		& "CleanMgr.exe" $arg1 | Out-Null 
	}
}	


#######################################
# Clean Microsoft Teams cache
#######################################
If ($MSTeams -eq $true){
	$teamsVersion = Get-Process ms-teams -ErrorAction SilentlyContinue
	If ($teamsVersion.ProcessName -eq "ms-teams"){
		$processName = "ms-teams"
	}else{
		$processName = "teams"
	}

	#Write-Host "Stopping Microsoft Teams in order to clear cache."
	try{
		Get-Process -ProcessName $processName -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue -Force
		Start-Sleep -Seconds 5
	#	Write-Host "Microsoft Teams has been successfully stopped."
	}
	catch{
		echo $_
	}
	# The cache is now being cleared.
	#Write-Host "Clearing Microsoft Teams cache."
	try{
		If ($processName -eq "teams"){
			Remove-Item -Path $env:APPDATA\"Microsoft\Teams" -ErrorAction SilentlyContinue -Recurse -Force
		}elseif($processName -eq "ms-teams"){
			Remove-Item -Path $env:APPDATA\"Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams" -ErrorAction SilentlyContinue -Recurse -Force
		}
	}
	catch{
		echo $_
	}
	 
	#write-host "The Microsoft Teams cache has been successfully cleared."

	Start-Sleep 5
	If ($processName -eq "teams"){
		& ${env:PROGRAMFILES(x86)}\"Microsoft\Teams\current\Teams.exe"
	}elseif($processName -eq "ms-teams"){
		& ${env:PROGRAMFILES}\"WindowsApps\MicrosoftTeams_23306.3309.2530.1346_x64__8wekyb3d8bbwe\msteams.exe"
	}
}
	

#######################################
# Clean OneDrive cache
#######################################
If ($OneDrive -eq $true){
	# Launching OneDrive with /reset
	Get-Process "OneDrive" -ErrorAction SilentlyContinue | kill -Force -ErrorAction SilentlyContinue | Out-Null
	Remove-Item -Path "HKCU:\SOFTWARE\Microsoft\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
	Remove-Item -Path "env:LOCALAPPDATA\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
	& $env:LOCALAPPDATA\"Microsoft\OneDrive\OneDrive.exe" /reset | Out-Null
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\OneDrive\Settings\" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
	Start-Sleep 5
	& $env:PROGRAMFILES\"Microsoft OneDrive\OneDrive.exe" 
}


#######################################
# Clean Outlook cache
#######################################
If ($Outlook -eq $true){
	#Write-Host "Stopping Microsoft Outlook in order to clear cache."
	try{
		Get-Process -ProcessName Outlook | Stop-Process -ErrorAction SilentlyContinue -Force
		Start-Sleep -Seconds 5
	#	Write-Host "Microsoft Outlook has been successfully stopped."
	}
	catch{
		echo $_
	}
	# The cache is now being cleared.
	#Write-Host "Clearing Microsoft Outlook cache."
	try{
		Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Outlook" -ErrorAction SilentlyContinue -Recurse -Force
		Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Windows\INetCache\Content.Outlook\*" -ErrorAction SilentlyContinue -Recurse -Force
	}
	catch{
		echo $_
	}
	 
	#write-host "The Microsoft Outlook cache has been successfully cleared."
	Start-Sleep 5
	& $env:PROGRAMFILES\"Microsoft Office\root\Office16\Outlook.exe" 
}



#######################################
# Clean PowerBI cache
#######################################
If ($PowerBI -eq $true){
	# Stop Power BI Desktop if it is running
	Stop-Process -ProcessName "PBIDesktop" -Force -ErrorAction SilentlyContinue

	# Delete the contents of the cache folders
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\AnalysisServicesWorkspaces"\* -ErrorAction SilentlyContinue -Recurse -Force
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\Cache"\* -ErrorAction SilentlyContinue -Recurse -Force
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\ExtensionCache"\* -ErrorAction SilentlyContinue -Recurse -Force
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\FoldedArtifactsCache"\* -ErrorAction SilentlyContinue -Recurse -Force
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\LuciaCache"\* -ErrorAction SilentlyContinue -Recurse -Force

	Start-Sleep 5
	Start shell:appsFolder\"PowerBIDesktop_7zvfj5dg98bjr!PBIDESKTOP"

}



#######################################
# Backup Edge Favorites
#######################################
If ($Edge -eq $true){
	#Variables
	If ($Default -eq $true){
		$HTML_File_Dir = "$env:OneDrive"
	}Else{
		Add-Type -AssemblyName 'System.Windows.Forms'
		$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
		$foldername.rootfolder = "MyComputer"

		If ($foldername.ShowDialog() -eq "OK"){
			$HTML_File_Dir = $foldername.SelectedPath
		}
	}
	#Path to bookmarks JSON file
	$JSON_File_Path = "$($env:localappdata)\Microsoft\Edge\User Data\Default\Bookmarks"

	#Path to exported HTML file
	$ExportedTime = Get-Date -Format "MM-dd-yyyy"
	$HTML_File_Path = "$($HTML_File_Dir)\EdgeFavorites-Bookmarks.backup_$($ExportedTime).html"


### Definitions
$EdgeStable="Edge"
$EdgeBeta="Edge Beta"
$EdgeDev="Edge Dev"
$ExportedTime = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'

### Choose the Edge Release ($EdgeStable, $EdgeBeta, $EdgeDev) you like to Backup:
$EdgeRelease=$EdgeStable

### Path to Edge Bookmarks Source-File
$JSON_File_Path = "$($env:localappdata)\Microsoft\$($EdgeRelease)\User Data\Default\Bookmarks"


### Filename of HTML-Export (Backup-Filename), choose with YYYY-MM-DD_HH-MM-SS Date-Suffix or fixed Filename
#$HTML_File_Path = "$($HTML_File_Dir)\EdgeChromium-Bookmarks.backup.html"
$HTML_File_Path = "$($HTML_File_Dir)\EdgeFavorites-Bookmarks.backup_$($ExportedTime).html"

## Reference-Timestamp needed to convert Timestamps of JSON (Milliseconds / Ticks since LDAP / NT epoch 01.01.1601 00:00:00 UTC) to Unix-Timestamp (Epoch)
$Date_LDAP_NT_EPOCH = Get-Date -Year 1601 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0

if (!(Test-Path -Path $JSON_File_Path -PathType Leaf)) {
    throw "Source-File Path $JSON_File_Path does not exist!" 
}
if (!(Test-Path -Path $HTML_File_Dir -PathType Container)) { 
    throw "Destination-Directory Path $HTML_File_Dir does not exist!" 
}

# ---- HTML Header ----
$BookmarksHTML_Header = @'
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<!-- This is an automatically generated file.
     It will be read and overwritten.
     DO NOT EDIT! -->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>Bookmarks</TITLE>
<H1>Bookmarks</H1>
<DL><p>
'@

$BookmarksHTML_Header | Out-File -FilePath $HTML_File_Path -Force -Encoding utf8

# ---- Enumerate Bookmarks Folders ----
Function Get-BookmarkFolder {
    [cmdletbinding()] 
    Param( 
        [Parameter(Position = 0, ValueFromPipeline = $True)]
        $Node 
    )
    function ConvertTo-UnixTimeStamp {
        param(
            [Parameter(Position = 0, ValueFromPipeline = $True)]
            $TimeStamp 
        )
        $date = [Decimal] $TimeStamp
        if ($date -gt 0) { 
            # Timestamp Conversion: JSON-File uses Timestamp-Format "Ticks-Offset since LDAP/NT-Epoch" (reference Timestamp, Epoch since 1601 see above), HTML-File uses Unix-Timestamp (Epoch, since 1970)																																																   
            $date = $Date_LDAP_NT_EPOCH.AddTicks($date * 10) # Convert the JSON-Timestamp to a valid PowerShell date
            # $DateAdded # Show Timestamp in Human-Readable-Format (Debugging-purposes only)																					
            $date = $date | Get-Date -UFormat %s # Convert to Unix-Timestamp
            $unixTimeStamp = [int][double]::Parse($date) - 1 # Cut off the Milliseconds
            return $unixTimeStamp
        }
    }   
    if ($node.name -like "Favorites Bar") {
        $DateAdded = [Decimal] $node.date_added | ConvertTo-UnixTimeStamp
        $DateModified = [Decimal] $node.date_modified | ConvertTo-UnixTimeStamp
        "        <DT><H3 FOLDED ADD_DATE=`"$($DateAdded)`" LAST_MODIFIED=`"$($DateModified)`" PERSONAL_TOOLBAR_FOLDER=`"true`">$($node.name )</H3>" | Out-File -FilePath $HTML_File_Path -Append -Force -Encoding utf8
        "        <DL><p>" | Out-File -FilePath $HTML_File_Path -Append -Force -Encoding utf8
    }
    foreach ($child in $node.children) {
        $DateAdded = [Decimal] $child.date_added | ConvertTo-UnixTimeStamp    
        $DateModified = [Decimal] $child.date_modified | ConvertTo-UnixTimeStamp
        if ($child.type -eq 'folder') {
            "        <DT><H3 ADD_DATE=`"$($DateAdded)`" LAST_MODIFIED=`"$($DateModified)`">$($child.name)</H3>" | Out-File -FilePath $HTML_File_Path -Append -Force -Encoding utf8
            "        <DL><p>" | Out-File -FilePath $HTML_File_Path -Append -Force -Encoding utf8
            Get-BookmarkFolder $child # Recursive call in case of Folders / SubFolders
            "        </DL><p>" | Out-File -FilePath $HTML_File_Path -Append -Force -Encoding utf8
        }
        else {
            # Type not Folder => URL
            "        <DT><A HREF=`"$($child.url)`" ADD_DATE=`"$($DateAdded)`">$($child.name)</A>" | Out-File -FilePath $HTML_File_Path -Append -Encoding utf8
        }
    }
    if ($node.name -like "Favorites Bar") {
        "        </DL><p>" | Out-File -FilePath $HTML_File_Path -Append -Force -Encoding utf8
    }
}

# ---- Convert the JSON Contens (recursive) ----
$data = Get-content $JSON_File_Path -Encoding UTF8 | out-string | ConvertFrom-Json
$sections = $data.roots.PSObject.Properties | Select-Object -ExpandProperty name
ForEach ($entry in $sections) { 
    $data.roots.$entry | Get-BookmarkFolder
}

# ---- HTML Footer ----
'</DL>' | Out-File -FilePath $HTML_File_Path -Append -Force -Encoding utf8


}






