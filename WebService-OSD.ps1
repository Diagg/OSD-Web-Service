#########################
# OSD WEB Service 
# 
# Initial code by Steve Lee from Microsoft (https://www.powershellgallery.com/packages/HttpListener/1.0.2/Content/HTTPListener.psm1)
# First edit and CSV formating by Sylvain Lesire
# Post method by Peter Hinchley (https://hinchley.net/articles/create-a-web-server-using-powershell/)
# Async .Net callback function by Oisin Grehan (http://www.nivot.org/post/2009/10/09/PowerShell20AsynchronousCallbacksFromNET)
# Async Request from Brandon Olin (https://www.powershellgallery.com/packages/PSHealthZ/1.0.0/Content/Public%5CStart-HealthZListener.ps1)
# Put together + all the rest by Diagg/OSD-Couture.com 
#
#
# Version 0.8.1 By Diagg/OSD-Couture.com 
#
# Release Date 20/12/2018
# Latest relase: 27/03/2019
#
# Purpose: 
#
#
# History:
#
# 07/12/2018 - V0.1    - Initial Release
#            - V0.2    - Added MDT output
#                      - Added Post support
#                      - Service State and request are logged in Event log
#            - V0.3    - Added Ping
#                      - Added multi listener auto registering
# 07/03/2019 - V0.4    - Added Basic Active Directory Support: GETPCINFO, ADDNEWPC, ADDGENERATEDPC, ADDPC2OU, ADDPC2GROUP
# 15/03/2019 - v0.5    - Added Basic SCCM support: ADDPC2COLLECTION
# 20/03/2019 - V0.6    - Refactored to support Asynchronious Web Requests (Boy that was hot!!!) 
# 22/03/2019 - V0.7    - TCP port to use is retived from firewall rule created from the installer
# 26/03/2019 - v0.8    - Heavy code refactoring (4 lines changed) to avoid a monstruous memory leak !!!
# 27/03/2019 - v0.8.1  - Fixed bugs in Get-CMComputername were Mac and IP adresses were not reported as strings
#                      - Changed Add-CMComputerToCollection. Now we wait until the collection has refeshed before moving on 
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
$Port = (Get-NetFirewallRule -DisplayName "OSD-WebService"| Get-NetFirewallPortFilter).LocalPort
$Url = "OSDInfo"

##== Init Script
."$CurrentScriptPath\DiaggFunctions.ps1"
."$CurrentScriptPath\WebService-Functions.ps1"
$CurrentUser = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).split("\")[1]
$Script:LogFile = Init-Logging -oPath "C:\ProgramData\MDT-Manager\$CurrentUser\Logs\WebService-Status.log"


##== Record start Time
$StartServiceTime = Get-Date

##== Check Port
If ([String]::IsNullOrWhiteSpace($Port) -or ($port -eq "Any")){Log-ScriptEvent -value "[ERROR] No specified Port in firewall rules, pleas re-run installer!" -severity 3 ; Exit}
If (Get-NetTCPConnection -State Listen -LocalPort $Port -ErrorAction SilentlyContinue){Log-ScriptEvent -value "[ERROR] Port $Port is already in use, please pick another one!" -severity 3 ; Exit}
                
				
#== Adding Listeners (localhost, FGDN, IP, local DNS suffix)
if ($Url.Length -gt 0 -and -not $Url.EndsWith('/')) {$Url += "/"}

$listener = New-Object System.Net.HttpListener
	        	
$prefix = "http://+:$Port/$Url"
Log-ScriptEvent -value "Registering listener: $prefix"
$listener.Prefixes.Add($prefix)

$prefix = "http://Localhost:$Port/$Url"
Log-ScriptEvent -value "Registering listener: $prefix"
$listener.Prefixes.Add($prefix)
                
$prefix = "http://$($env:COMPUTERNAME):$Port/$Url"
Log-ScriptEvent -value "Registering listener: $prefix"                
$listener.Prefixes.Add($prefix)

[System.Net.Dns]::GetHostByName($env:computerName).addresslist.IPAddressToString|ForEach-Object{
        $prefix = "http://$($_):$Port/$Url" 
        Log-ScriptEvent -value "Registering listener: $prefix"               
        $listener.Prefixes.Add($prefix)
    } 

Get-DnsClient|Select-Object Hostname, ConnectionSpecificSuffix -Unique|ForEach-Object { 
        If(-not [String]::IsNullOrWhiteSpace($_.ConnectionSpecificSuffix))
            {
                $prefix = "http://$($_.Hostname).$($_.ConnectionSpecificSuffix):$Port/$Url"
                Log-ScriptEvent -value "Registering listener: $prefix"                
                $listener.Prefixes.Add($prefix)
            }
    }


##== Create EventLog
$EventLogName = 'WEBService OSD'
$EventLogListener = 'OSD Listener'
New-EventLog -LogName $EventLogName -Source $EventLogListener -ErrorAction SilentlyContinue



##== Start Listenting
Try {$listener.Start()}
Catch 
    {
		$Message = "[ERROR] Unable to start WEB service, aborting!!!"
        Write-EventLog -LogName $EventLogName -Source $EventLogListener -EntryType Error -EventId 99 -Message $Message
        Log-ScriptEvent -value $Message -severity 3
        $ErrorOutput = $_ | Resolve-Error -AsString
        Exit
    }
 
$Message = "OSD WEB Service has started - $prefix - Start time: $StartServiceTime"
Write-EventLog -LogName $EventLogName -Source $EventLogListener -EntryType Information -EventId 1 -Message $Message
Log-ScriptEvent -value $Message
Log-ScriptEvent -value "Listening on port $port..."


##== Create runspace script
$requestListener = {
                    [cmdletbinding()]
                    param($result)

                    [System.Net.HttpListener]$listener = $result.AsyncState
                    $context = $listener.EndGetContext($result)
                    $request = $context.Request
                    $response = $context.Response
                    $statusCode = [System.Net.HttpStatusCode]::OK
                    $buffer = $Null

                    ##== Analysing Request Type
                    $Message = "Received $($request.httpMethod) request on $(get-date) with url $($request.url.OriginalString) :"
                    Write-EventLog -LogName $EventLogName -Source $EventLogListener -EntryType Information -EventId 2 -Message $Message
                    Log-ScriptEvent -value $Message	
                                
                    If ($request.httpMethod  -eq 'POST')
                        {
                            $Data = Read-PostRequest -request $request
							if ($null -eq $Data) {$commandOutput = "[ERROR] Bad request formating, use syntax: command=<string> format=[JSON|TEXT|XML|NONE|CLIXML]" ; $Format = "TEXT"}	
                        } 
                                
                    If	($request.httpMethod  -eq 'GET')	                    
						{
					        if (-not $request.QueryString.HasKeys()) 
								{$commandOutput = "[ERROR] Bad request formating, use syntax:  command=<string> format=[JSON|TEXT|XML|NONE|CLIXML]" ; $Format = "TEXT"} 
							Else
								{$Data = Read-GetRequest -request $request}
						}


                    ##== Setup the Output format
                    If ([String]::IsNullOrWhiteSpace($Format) -and ( -not [String]::IsNullOrWhiteSpace($Data.get_Item("format"))) ) 
                        {$Format = $Data.get_Item("format")}
                    Elseif ([String]::IsNullOrWhiteSpace($Format))
                        {$Format = "MDTXML"}
	                Log-ScriptEvent -value "Output Format = $Format"
								
                    ##== Process requests
                    If (-not [String]::IsNullOrWhiteSpace($Data))
                        {
					        foreach ($kvp in $Data.GetEnumerator()) {Log-ScriptEvent -value $("Recieved " + $kvp.Key + " = "  + $kvp.Value)}			
                    
	                        $command = $Data.get_Item("command")
		                    switch ($command.toupper() ) 
						        {
	                                "EXIT" 
								        {
	                                        $Message = "[EXIT] Received command to exit listener"
                                            Write-EventLog -LogName $EventLogName -Source $EventLogListener -EntryType Information -EventId 100 -Message $Message
                                            Log-ScriptEvent -value $Message
                                            $listener.Stop()
                                            $listener.Close()
                                            Exit
	                                    }

							        "PING"
								        {
                                            $commandOutput = Ping-service
                                            $Message = "The service is now running for $($commandOutput.days) Days, $($commandOutput.Hours) Hours and $($commandOutput.Minutes) Minutes"
                                            If ($Format.ToUpper() -eq "TEXT") {$commandOutput = $Message}
                                            Log-ScriptEvent -value $Message
                                            break
								        }
                                                    
                                    "GETTIME"
                                        {
                                            $commandOutput = Get-Time
                                            Break
                                        }
                                                    
                                    "GETPCINFO"
                                        {
                                            $commandOutput = Get-ComputerName -ComputerName $Data.get_Item("ComputerName") -SearchPath $Data.get_Item("SearchPath")
                                            Break
                                        }    
                                                
                                    "ADDNEWPC"
                                        {
                                            $commandOutput = Add-NewComputer -ComputerName $Data.get_Item("ComputerName") -OU $Data.get_Item("OU")
                                            Break  
                                        }

                                    "ADDGENERATEDPC"
                                        {
                                            $commandOutput = Add-NewGeneratedComputer -Prefix $Data.get_Item("Prefix") -Suffix $Data.get_Item("Suffix") -Digits $Data.get_Item("Digits") -OU $Data.get_Item("OU")
                                            Break  
                                        }                                                    

                                    "ADDPC2OU"
                                        {
                                            $commandOutput = Add-ComputerToOU -ComputerName $Data.get_Item("ComputerName") -OU $Data.get_Item("OU")
                                            Break  
                                        }

                                    "ADDPC2GROUP"
                                        {
                                            $commandOutput = Add-ComputerToGroup -ComputerName $Data.get_Item("ComputerName") -Group $Data.get_Item("Group")
                                            Break  
                                        }

                                    "ADDPC2COLLECTION"
                                        {
                                            $commandOutput = Add-CMComputerToCollection -ComputerName $Data.get_Item("ComputerName") -Collection $Data.get_Item("Collection")
                                            Break  
                                        }
                                                    			
							        Default 
                                        {
                                            $commandOutput = "404 Bro ! This is not the page you're looking for."
                                            $Format = "TEXT"
                                            $statusCode = 404
                                            Break
                                        }
    					        }	
                        }


		            ##== Formating Output
				    if (!$commandOutput) {$commandOutput = " "}
                    $commandOutput = switch ($Format.ToUpper()) 
						{
			                TEXT    { $commandOutput | Out-String ; break } 
			                JSON    { $commandOutput | ConvertTo-JSON; break }
			                CSV     { $result= $null ; ($commandOutput | ConvertTo-csv -delimiter ";" -NoTypeInformation) | ForEach-Object{$result += $_+"###"} ; [string]$result ; break }
			                XML     { $commandOutput | ConvertTo-XML -As String; break }
                            MDTXML  { Dump-ObjectToXml -obj $commandOutput -tag 'MDT' ; break }              
			                CLIXML  { [System.Management.Automation.PSSerializer]::Serialize($commandOutput) ; break }
			                default { "Invalid output format selected, valid choices are TEXT, JSON, CSV, XML, MDTXML and CLIXML"; $statusCode = 501; break }
	                    }


				    ##== Responding
                    Log-ScriptEvent -value "Response:"
                    If ([String]::IsNullOrWhiteSpace($commandOutput)){Log-ScriptEvent -value "No data to send, back !!" } Else {Log-ScriptEvent -value $commandOutput}

				    $response.ContentType = 'text/html'
                    $response.StatusCode = $statusCode
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($commandOutput)

				    $response.ContentLength64 = $buffer.Length
				    $response.OutputStream.Write($buffer,0,$buffer.Length)
				    $response.Close()
                    
            }  

## Launch Async callback for the first time
$context = $listener.BeginGetContext((New-ScriptBlockCallback -Callback $requestListener), $listener)

 
##== Run until you send a GET request to /end
while ($listener.IsListening)
    {
         
        ##== Run new New Async Callback instance once the previous one get filled by a request
        If ($context.IsCompleted -eq $true) {$context = $listener.BeginGetContext((New-ScriptBlockCallback -Callback $requestListener), $listener)}
        
        ##== Heartbeat (every 10 Minutes)
        $oTimeLapse = New-TimeSpan -Start $StartServiceTime -End $(Get-date)
        If ($PreviousTime.Minutes -ne $oTimeLapse.Minutes)
            {
                $PreviousTime = $oTimeLapse
                If ($oTimeLapse.Minutes % 10 -eq 0) 
                    {
                        Log-ScriptEvent -value "$((Get-Date).ToShortDateString()) - $((Get-Date).ToShortTimeString()) - The process is still running since $($oTimeLapse.days) Days - $($oTimeLapse.Hours) Hours - $($oTimeLapse.Minutes) Minutes - $($oTimeLapse.Seconds) Seconds"
                        Log-ScriptEvent -value "Average Memory consumption is: $((Get-Process -Id $pid).WS/1MB) MB" 
                    }
            }    
 
    }
 
##== Terminate the listener
$listener.Close()
Log-ScriptEvent -value 'Terminated...'