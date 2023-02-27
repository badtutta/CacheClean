<#


               Name:    CacheClean.ps1                                                               
                                                                                                   
          Author(s):    Navy Flank Speed MCS                                                                                                                                                 
                                                                                                       
            Version:    1.0.x                                                                        
                                                                                                  
            Created:    01.18.2023                                                                 
           Modified:    01.18.2023                                                                  
                                                                                                  
                                                                                                  
                                                                                                  
          Info:                                                                                   
                                                                                                  
          This script can determine the size of the user's profile and be used to perform a disk cleanup 
          to free up space via CleanMgr.exe, as well as, cleaning of the following cache...
		  
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
		  
		  
          

          Version Log:

          .1 - Initial script 
                        

#>
 
[CmdLetBinding()]
Param (
    [Parameter(Mandatory=$false)] [switch]$ProfileSize = $false,
	[Parameter(Mandatory=$false)] [switch]$Clean = $false,
	[Parameter(Mandatory=$false)] [switch]$Options = $false,
	[Parameter(Mandatory=$false)] [switch]$MSTeams = $fals,
	[Parameter(Mandatory=$false)] [switch]$OneDrive = $false,
	[Parameter(Mandatory=$false)] [switch]$PowerBI = $false,
	[Parameter(Mandatory=$false)] [switch]$Outlook = $false
	
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
	Write-Host "Stopping Microsoft Teams in order to clear cache."
	try{
		Get-Process -ProcessName Teams | Stop-Process -Force
		Start-Sleep -Seconds 5
		Write-Host "Microsoft Teams has been successfully stopped."
	}
	catch{
		echo $_
	}
	# The cache is now being cleared.
	Write-Host "Clearing Microsoft Teams cache."
	try{
		Remove-Item -Path $env:APPDATA\"Microsoft\teams\blob_storage" -Recurse -Force
		Remove-Item -Path $env:APPDATA\"Microsoft\teams\databases" -Recurse -Force
		Remove-Item -Path $env:APPDATA\"Microsoft\teams\Cache" -Recurse -Force
		Remove-Item -Path $env:APPDATA\"Microsoft\teams\GPUcache" -Recurse -Force
		Remove-Item -Path $env:APPDATA\"Microsoft\teams\IndexedDB" -Recurse -Force
		Remove-Item -Path $env:APPDATA\"Microsoft\teams\Local Storage" -Recurse -Force
		Remove-Item -Path $env:APPDATA\"Microsoft\teams\tmp" -Recurse -Force 
	}
	catch{
		echo $_
	}
	 
	write-host "The Microsoft Teams cache has been successfully cleared."
}
	

#######################################
# Clean OneDrive cache
#######################################
If ($OneDrive -eq $true){
	# Launching OneDrive with /reset
	$arg1 = "/reset"
	& $env:LOCALAPPDATA\"Microsoft\OneDrive\OneDrive.exe" $arg1 | Out-Null
}


#######################################
# Clean Outlook cache
#######################################
If ($Outlook -eq $true){
	Write-Host "Stopping Microsoft Outlook in order to clear cache."
	try{
		Get-Process -ProcessName Outlook | Stop-Process -Force
		Start-Sleep -Seconds 5
		Write-Host "Microsoft Outlook has been successfully stopped."
	}
	catch{
		echo $_
	}
	# The cache is now being cleared.
	Write-Host "Clearing Microsoft Outlook cache."
	try{
		Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Windows\INetCache\Content.Outlook\*" -Recurse -Force
	}
	catch{
		echo $_
	}
	 
	write-host "The Microsoft Outlook cache has been successfully cleared."
}



#######################################
# Clean PowerBI cache
#######################################
If ($PowerBI -eq $true){
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\cache" -Recurse -Force
	

}