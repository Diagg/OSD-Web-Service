#########################
# WEB service installer
#
# Version 1.5 By Diagg/OSD-Couture.com 
#
# Release Date 04/02/2019
# Latest release: 26/03/2019
#
# Purpose: 
#	Install OSD WEB service and all his dependencies
#
# History:
#
# 02/02/2019 - V1.0 - Initial Release
# 08/02/2019 - V1.1 - Added "run as a service feature" and credencials
# 10/02/2019 - V1.2 - Added Scheduled task to ping the service
# 08/03/2019 - V1.3 - Added credential validation by Jaap Brasser
# 22/03/2019 - V1.4 - Removed unused features
#                   - Added firewall Rule
# 26/03/2019 - V1.5 - The service now run in "above normal" priority
#
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
[Int]$Port = 8550


##== Functions
function Test-LocalCredential 
    {
        [CmdletBinding()]
        Param
        (
            [string]$UserName,
            [string]$ComputerName = $env:COMPUTERNAME,
            [string]$Password
        )
        if (!($UserName) -or !($Password)) {
            Write-Warning 'Test-LocalCredential: Please specify both user name and password'
        } else {
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
            $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('machine',$ComputerName)
            $DS.ValidateCredentials($UserName, $Password)
        }
    }


function Test-ADCredential 
    {
        [CmdletBinding()]
        Param
        (
            [string]$UserName,
            [string]$Password
        )
        if (!($UserName) -or !($Password)) {
            Write-Warning 'Test-ADCredential: Please specify both user name and password'
        } else {
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
            $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('domain')
            $DS.ValidateCredentials($UserName, $Password)
        }
    }



##== Create folder for log file
$NewLog = "C:\ProgramData\MDT-Manager\$CurrentUser\Logs"
If (-not (test-path $NewLog))
    {
        New-Item -Path $NewLog -ItemType Directory -Force|out-null
        $testPath = TestAndLog-Path $NewLog -Action created
        If ($testPath -eq $false) {Log-ScriptEvent -value "[ERROR] Unable to create $NewLog, Aborting!!!!" ; EXIT}
    }
Else
    {Log-ScriptEvent -value "Path $NewLog already created, nothing to do !!!"}


##== Move log file to C:\ProgramData\MDT-Manager\Logs
$Script:LogFile = Relocate-Logging -path $NewLog
Log-ScriptEvent -value "New location for Logs is $Script:LogFile"


##== Create folder Tools
$NewTools = "C:\ProgramData\MDT-Manager\Tools"
If (-not (test-path $NewTools))
    {
        New-Item -Path $NewTools -ItemType Directory -Force|out-null
        $testPath = TestAndLog-Path $NewTools -Action created
        If ($testPath -eq $false) {Log-ScriptEvent -value "[ERROR] Unable to create $NewTools, Aborting!!!!" ; EXIT}
    }
Else
    {Log-ScriptEvent -value "Path $NewTools already created, nothing to do !!!"}


##== Download NSSM
$output = "$NewTools\nssm-2.24-101-g897c7ad\win64\nssm.exe"
If (-not(Test-path $output))
    {
        Log-ScriptEvent -value "Downloading NSSM to $output"
        $url = "https://nssm.cc/ci/nssm-2.24-101-g897c7ad.zip"
        Start-BitsTransfer -Source $url -Destination $env:temp
		Expand-Archive -Path "$env:temp\nssm-2.24-101-g897c7ad.zip" -DestinationPath $NewTools
        $testPath = TestAndLog-Path $output
        If ($testPath -eq $false) {Log-ScriptEvent -value "[ERROR] Unable to download NSSM to folder $output, Please download and install the tool manually and retry, Aborting!!!!" ; EXIT}
    }
Else
    {Log-ScriptEvent -value "NSSM already downloaded, nothing to do !!!"}


##== Copy Files in C:\ProgramData\MDT-Manager\Tools Folder
Log-ScriptEvent -value "Copying Scripts to $NewTools"
$ManagerFiles = @("WebService-OSD.ps1","DiaggFunctions.ps1","WebService-Functions.ps1","WebService-UnInstaller.ps1")
Foreach ($File in $ManagerFiles)
	{
		If (Test-Path "$Script:CurrentScriptPath\$File"){Copy-Item -Path "$Script:CurrentScriptPath\$File" -Destination $NewTools}
	}


##== Install AD Powershell Module
If ((Get-WmiObject -class Win32_OperatingSystem).caption -notlike '*server*')
	{
		If (-not ((Get-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0").State -eq 'Installed'))
		    {Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"}
	}
Else
	{
		If (-not ((Get-WindowsFeature -Name "RSAT-ad-powershell").InstallState -eq 'Installed'))
		    {
		        Enable-WindowsOptionalFeature -Online -FeatureName RSAT-ADDS-Tools-Feature -ErrorAction SilentlyContinue
                Add-WindowsFeature -Name "RSAT-AD-AdminCenter" -IncludeAllSubFeature
                Add-WindowsFeature -Name "RSAT-AD-PowerShell" –IncludeAllSubFeature
		    } 
	}
Import-Module activedirectory


##== Get Account that will run the service
Log-ScriptEvent -value "Getting Credential for account used to run the service"
Do
    {
        $creds = Get-Credential -Message "Enter Account with enought rights to run the service"
        $User = $creds.GetNetworkCredential().UserName 
        $Pass = $creds.GetNetworkCredential().password
        $Dom = $creds.GetNetworkCredential().Domain
        If ((-not [string]::IsNullOrWhiteSpace($Dom)) -or ($Dom -ne $env:COMPUTERNAME))
            {$TestCred = Test-ADCredential -UserName $User -Password $Pass}
        Else
            {$TestCred = Test-LocalCredential -UserName $User -Password $Pass} 
            
        If ($TestCred -eq $false) {Log-ScriptEvent -value "Unable to validate credential, please enter valid credential" -Severity 2}    
            
              
    }Until($TestCred -eq $true)

Log-ScriptEvent -value "Credential Validated successfully"
If ( -not [string]::IsNullOrWhiteSpace($Dom)) {$User= "$Dom\$User"}


##== Install Service
$nssm = $output
$serviceName = 'OSD-WEBService'
$powershell = (Get-Command powershell).Source
$scriptPath = "$NewTools\WebService-OSD.ps1"
$arguments = '-ExecutionPolicy Bypass -NoProfile -File "{0}"' -f $scriptPath

If ((Get-Service $serviceName -ErrorAction SilentlyContinue).Name -eq $serviceName)
	{
		Log-ScriptEvent -value "Uninstalling previous version of the service"
		If ((Get-Service $serviceName).status -eq "Running"){Stop-Service $serviceName -Force -ErrorAction SilentlyContinue}
		
        If ((Get-Service $serviceName).status -eq "Stopped")
            {
                Log-ScriptEvent -value "Service $serviceName stopped successfully !!"
                $Cmd = Invoke-Executable -Path $nssm -Arguments "remove $serviceName confirm"
                If ($Cmd -match "service does not exist"){Log-ScriptEvent -value "Service $serviceName already uninstalled !!"}
            }
        Else
            {Log-ScriptEvent -value "[ERROR] Unable to stop service $serviceName, Aborting!!!!" ; EXIT}
	}

Log-ScriptEvent -value "Registering WEB service"
& $nssm install $serviceName $powershell $arguments
Log-ScriptEvent -value "Adding service account"
$arguments = "ObjectName"
& $nssm set $serviceName $arguments $User $Pass
$arguments = "AppPriority" ; $Value = "BELOW_NORMAL_PRIORITY_CLASS"
& $nssm set $serviceName $arguments $Value
Log-ScriptEvent -value "Getting WEB service status:"
& $nssm status $serviceName
Log-ScriptEvent -value "Starting WEB service:"
Start-service $serviceName
Log-ScriptEvent -value "Starting New WEB service:"
Get-Service $serviceName

##== Create Firewall Rule for WebService
If ((Get-NetFirewallRule -DisplayName "OSD-WebService" -ErrorAction SilentlyContinue).DisplayName -eq "OSD-WebService") 
    {
        Log-ScriptEvent -value "Uninstalling previous version of the Firewall Rules"
        Remove-NetFirewallRule -DisplayName "OSD-WebService" -Confirm:$false
        If (-not((Get-NetFirewallRule -DisplayName "OSD-WebService" -ErrorAction SilentlyContinue).DisplayName -eq "OSD-WebService")){Log-ScriptEvent -value "Firewall Rules uninstalled successfully !!"} 
    }

Log-ScriptEvent -value "Installing Firewall Rules"
New-NetFirewallRule -Name "OSD-WebService" -DisplayName "OSD-WebService" -Description "OSDC - Quality Deployment since 1884" -Enabled True -Profile Domain,Private -Direction Inbound -Action Allow -LocalPort $Port -Protocol TCP -RemotePort Any|Out-Null

Log-ScriptEvent -value "Installation Finished !!!!"