#########################
# WEB service installer
#
# Version 1.0 By Diagg/OSD-Couture.com 
#
# Release Date 18/03/2019
# Latest release: 26/03/2019
#
# Purpose: 
#	UnInstall OSD WEB service and all his dependencies
#
# History:
#
# 18/03/2019 - V1.0 - Initial Release
# 26/03/2019 - V2.0 - Added try/Cacth when shuting down the service
# 
#
#########################

#Requires -Version 4
#Requires -RunAsAdministrator

Clear-Host
##== Debug
$ErrorActionPreference = "stop"
#$ErrorActionPreference = "Continue"

##== Global Variables
$Script:CurrentScriptName = $MyInvocation.MyCommand.Name
$Script:CurrentScriptFullName = $MyInvocation.MyCommand.Path
$Script:CurrentScriptPath = split-path $MyInvocation.MyCommand.Path

##== Init Script
."$CurrentScriptPath\DiaggFunctions.ps1"
$CurrentUser = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).split("\")[1]
$Script:LogFile = Init-Logging

##== UnInstall Service
$NewTools = "C:\ProgramData\MDT-Manager\Tools"
$nssm = "$NewTools\nssm-2.24-101-g897c7ad\win64\nssm.exe"
$serviceName = 'OSD-WEBService'
$powershell = (Get-Command powershell).Source
$scriptPath = "$NewTools\WebService-OSD.ps1"
$arguments = '-ExecutionPolicy Bypass -NoProfile -File "{0}"' -f $scriptPath

Log-ScriptEvent -value "Uninstalling OSD Web service !!"


If ((Get-Service $serviceName -ErrorAction SilentlyContinue).Name -eq $serviceName)
	{

        ##== Closing Web Service
        Log-ScriptEvent -value "Stopping web server !!"
        Try {Invoke-WebRequest -Uri "http://$($env:computername):8530/OSDInfo?command=Exit" -ErrorAction SilentlyContinue|Out-Null}
        Catch {If ($error[0] -like "*404.0*"){Log-ScriptEvent -value "Web server stopped successfully!!"}}

		Log-ScriptEvent -value "Stopping Windows service"
		If ((Get-Service $serviceName).status -ne "Stopped"){Stop-Service $serviceName -Force -ErrorAction SilentlyContinue}
		
        If ((Get-Service $serviceName).status -eq "Stopped")
            {
                Log-ScriptEvent -value "Service $serviceName stopped successfully !!"
                Log-ScriptEvent -value "Uninstalling NSSM service"
                $Cmd = Invoke-Executable -Path $nssm -Arguments "remove $serviceName confirm"
                If ($Cmd -match "service does not exist"){Log-ScriptEvent -value "Service $serviceName already uninstalled !!"}
            }
        Else
            {Log-ScriptEvent -value "[ERROR] Unable to stop service $serviceName, Aborting!!!!" ; EXIT}
	}

##== Schedule task to ping WEB service evry 5 minutes
$TaskName ="Ping OSD WEB Service"
If ([string]::IsNullOrWhiteSpace($(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue).State))
    {
        Log-ScriptEvent -value "Uninstalling Scheduled task"
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$False
    }
Else
    {Log-ScriptEvent -value "Scheduled task already uninstalled !!"}


##== Remove Firewal Rule
If ((Get-NetFirewallRule -DisplayName "OSD-WebService" -ErrorAction SilentlyContinue).DisplayName -eq "OSD-WebService") 
    {
        Log-ScriptEvent -value "Removing Firewall Execption"
        Remove-NetFirewallRule -DisplayName "OSD-WebService" -Confirm:$false
    }
Else
    {Log-ScriptEvent -value "Firewall Execption already uninstalled !!"}


##== Delete all Files
Log-ScriptEvent -value "Removing remaining files !!"
Remove-item -path $NewTools -Recurse -Force -ErrorAction SilentlyContinue


##== The end my friend
Log-ScriptEvent -value "OSD WEB Service Uninstallation finished !!"
  