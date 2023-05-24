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
 

Param (
    [Parameter(Mandatory=$false)] [switch]$ProfileSize = $false,
	[Parameter(Mandatory=$false)] [switch]$Clean = $false,
	[Parameter(Mandatory=$false)] [switch]$Options = $false,
	[Parameter(Mandatory=$false)] [switch]$MSTeams = $fals,
	[Parameter(Mandatory=$false)] [switch]$OneDrive = $false,
	[Parameter(Mandatory=$false)] [switch]$PowerBI = $false,
	[Parameter(Mandatory=$false)] [switch]$Outlook = $false
	
)


# .Net methods for hiding/showing the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}
Hide-Console

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
	Get-Process "OneDrive" | kill -Force -ErrorAction SilentlyContinue | Out-Null
	Remove-Item -Path "HKCU:\SOFTWARE\Microsoft\OneDrive" -Recurse -Force | Out-Null
	Remove-Item -Path "$env:LOCALAPPDATA\OneDrive" -Recurse -Force | Out-Null
	& $env:ProgramFiles"\Microsoft OneDrive\OneDrive.exe" /reset | Out-Null
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\OneDrive\Settings\" -Recurse -Force | Out-Null
	Start-Sleep 5
	& $env:ProgramFiles\"Microsoft OneDrive\OneDrive.exe"
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
	# Stop Power BI Desktop if it is running
	Stop-Process -ProcessName "PBIDesktop" -Force -ErrorAction SilentlyContinue

	# Delete the contents of the cache folders
	"$env:LOCALAPPDATA\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces"
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\AnalysisServicesWorkspaces"\* -Recurse -Force
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\Cache"\* -Recurse -Force
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\ExtensionCache"\* -Recurse -Force
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\FoldedArtifactsCache"\* -Recurse -Force
	Remove-Item -Path $env:LOCALAPPDATA\"Microsoft\Power BI Desktop\LuciaCache"\* -Recurse -Force

}
# SIG # Begin signature block
# MIIF0wYJKoZIhvcNAQcCoIIFxDCCBcACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUmakCKIOEWF+QDixzGXtL7/Gy
# dhagggNMMIIDSDCCAjCgAwIBAgIQWgJUYzsrk7dAXVlfryMOoDANBgkqhkiG9w0B
# AQsFADA8MTowOAYDVQQDDDFOYXV0aWx1cyBWaXJ0dWFsIERlc2t0b3AgQ29kZSBT
# aWduaW5nIENlcnRpZmljYXRlMB4XDTIzMDUyMjIwMTQzMloXDTI0MDUyMjIwMzQz
# MlowPDE6MDgGA1UEAwwxTmF1dGlsdXMgVmlydHVhbCBEZXNrdG9wIENvZGUgU2ln
# bmluZyBDZXJ0aWZpY2F0ZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# APzLCLETl7GZkzTgUNFM4ynsPnyQXRIMimenwWCLdquBjGXWrsyIL2zYnw3Tppx5
# zJLtJM8kYC7vR4ycW4bZBNBE+QweorRzQSI5afJ9Gmr7R0DOWAw/zm/2XhhZGLnb
# MHb3Wnh+Ll4HVdhKLzsXb7D2rEKdUGC5iLeH5LSztYqXzE4yb05kNL9zQNOGuFNC
# R82+Fyz+tEex24JVZSSCH6xVnHMn5x9osiuVqHsegLNo752UVhUFz8ajfAiP6Owt
# hc7M8fCNtRwpI8s4eZBbt0OeeqjYWiWQMOYBrP9LDI9z23PYBFDmkVsJ+NbYVsq2
# KC+zPETo/XGWoX4cnyWbYpECAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBTK5onkx2bgvViKK5enaFSXc7QkEjAN
# BgkqhkiG9w0BAQsFAAOCAQEAK5RAgEWJpqb+QUzfodcAzrfvMN+tq0qGvrE2wfhn
# jh5klpMyPPJDdUukS/R5O0Vm6NrXsvSgwK1rwT+qq1rJwoTXDAy+s7HmOlk4TyQ/
# ShnGm4UG2/j4NIW8ouTxNVwWRLFV/nQSLf2/PvyqI+zvjeFm86ABcdBZY0Yqewmm
# D7ifDVzBX/3vVd5XJeT6IOpOHr6fOHv8bxgikEUgLwviFPwEswP7SGzvTAuqipIU
# UxBf096sM50lzal8MUydmDhEXIjm8SVxeLx89RnWAx1hfjiJaGgm5qwfPkkvIxoo
# wPKkLzXdKc8fZDYSk9oS3XqiF5yKTbA6aajqtwXfZAI08jGCAfEwggHtAgEBMFAw
# PDE6MDgGA1UEAwwxTmF1dGlsdXMgVmlydHVhbCBEZXNrdG9wIENvZGUgU2lnbmlu
# ZyBDZXJ0aWZpY2F0ZQIQWgJUYzsrk7dAXVlfryMOoDAJBgUrDgMCGgUAoHgwGAYK
# KwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIB
# BDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU
# FwZ6rGEdO4nF7+yrqvFue4sRGZ0wDQYJKoZIhvcNAQEBBQAEggEA2OhkKpfJxj4J
# EMNml/7tuDcnxsXwAF7DgrVfxhHReEMM86ulQhniQb+KvNs474H1IEAq/KJfe6zR
# VW+GWlVDr1i9xPGBF+WaV0DUGdiywSVHwm03QMRHKjkG/T4HgvEW27TvtJFjTdUu
# aqWYbsQ1tenb1YnuvAHfTYQpDO1oYWj3eoKE9/Ssc+0e4j53hl++mDxFJQ4fZeQR
# dTn9qy6azy/RU9pYUup/9XciUFgDC09mPTNAuTIMuJ5+RQqOzNq3jXA+5jx7K7lj
# XBwnqak3kOadeyu0WlMcjlB5Uv8QxtlHk7pdWbWNsv7cGRUDPb90kBqQ8wSPc37V
# 4+4dBZs4wA==
# SIG # End signature block
