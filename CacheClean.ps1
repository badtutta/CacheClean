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
	[Parameter(Mandatory=$false)] [switch]$MSTeams = $false,
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
	#Write-Host "Stopping Microsoft Teams in order to clear cache."
	try{
		Get-Process -ProcessName Teams -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue -Force
		Start-Sleep -Seconds 5
	#	Write-Host "Microsoft Teams has been successfully stopped."
	}
	catch{
		echo $_
	}
	# The cache is now being cleared.
	#Write-Host "Clearing Microsoft Teams cache."
	try{
		Remove-Item -Path $env:APPDATA\"Microsoft\Teams" -ErrorAction SilentlyContinue -Recurse -Force
	}
	catch{
		echo $_
	}
	 
	#write-host "The Microsoft Teams cache has been successfully cleared."

	Start-Sleep 5
	& ${env:PROGRAMFILES(x86)}\"Microsoft\Teams\current\Teams.exe"
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

}
# SIG # Begin signature block
# MIIF0wYJKoZIhvcNAQcCoIIFxDCCBcACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUf+/XtUmMMM4eA6BXzQbW2frD
# p6WgggNMMIIDSDCCAjCgAwIBAgIQcXHqnYdQMbxAV5XYyArlADANBgkqhkiG9w0B
# AQsFADA8MTowOAYDVQQDDDFOYXV0aWx1cyBWaXJ0dWFsIERlc2t0b3AgQ29kZSBT
# aWduaW5nIENlcnRpZmljYXRlMB4XDTIzMDUyMjE4MjU1N1oXDTI0MDUyMjE4NDU1
# N1owPDE6MDgGA1UEAwwxTmF1dGlsdXMgVmlydHVhbCBEZXNrdG9wIENvZGUgU2ln
# bmluZyBDZXJ0aWZpY2F0ZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# AKH0I9g2NPsiqgP2Gmi9ifMe9GRSQOTVj5nnp/zwb0AbzJF5qpIOX9IeIC0/gmUz
# dj+WwkGL4lONzVKIPNTKPbBPDzRnBK9wJazJjSYDbRh4Yd3HbMCo7W2e5NMII0li
# WkPDJioVHw8Zh7BdN+0i0yOLb3A6Jr62CP3wHOn/LIYcXK6HA6jS2WUHjGoBr2rW
# lYAdumrrTtL1KbBYYjIjAdkrAgX/XU5dzVBosMkaWQXJTrdoJzAPe9erYItowKQS
# 9iiCSlT6K0bbCpxt9EDg1GpLN7yqH4Ut/Q2ZMxr5rc21r80gqoCX7+gi49flOO5X
# saf16boIv0CtPysQZWXPCZkCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBSXn4hbjGNiJmu5Zveef2+DUTM4VTAN
# BgkqhkiG9w0BAQsFAAOCAQEAUrepMOc58NSqFtzV8o4cpxn9L0jjXpTdibg0zRtj
# av5pIZgGtrH063rWHlstLluuF9OPIfAlVOK5cAIj21pOY4ecqVMaF/QslhZ2ahiL
# N71kT9JpcOPMWewwfPfl6OBh+Dnofj4WdkdIF+Y05g19dMjF2og93w3cnjVFzH1Y
# RSRK7R44EZ3QMISFOHvJSSv9xsf0jLItYNqxDFf9ySFMpOss/ZiZEsCoHhP4kMv0
# EVgIwbTp8JQd0Q/ffohufLS1fiD7HhVBPMlhAHhnNbBFjUtxSFGifW/vqt8oNk2I
# 9CySPYfoD5A0x0beVyDrVlNamgKhp/jk9FQZ5FJSPt00HTGCAfEwggHtAgEBMFAw
# PDE6MDgGA1UEAwwxTmF1dGlsdXMgVmlydHVhbCBEZXNrdG9wIENvZGUgU2lnbmlu
# ZyBDZXJ0aWZpY2F0ZQIQcXHqnYdQMbxAV5XYyArlADAJBgUrDgMCGgUAoHgwGAYK
# KwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIB
# BDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU
# omv0h1ETJ2bDuBbUkEMy0mrJVPcwDQYJKoZIhvcNAQEBBQAEggEAULf/jm9EvYOH
# qa/Tg9XW7m16W0TypJSgLj8mpPfOOgghw/uZRzWBW1MyrIXzzuiYzcF+weZAoKWi
# XfjKw48iCU3Ft0el6sH1QEMEQdinO2rZqZdHl2jzuYHXZIJCDx6WwVgoG8IncA39
# lNIcLt7Rg0oKgkylucqrTJisPpZsWsINftkLquo3kBhIQxO7dl2xuXrl+4Gm6cLm
# BAkG/c0Ifu8CGRyHjY9CcmA7dSErHN6qWeAoPJEpcmlW7ai+vX+LJBqS+IknpEBO
# RDo3iORGUeiiHY7RfsCtkiWkRnwZVg7ZILBy5+pJr9YI9i8Obgw4ZDXXDZSp0RJG
# gV0mnzhc7g==
# SIG # End signature block
