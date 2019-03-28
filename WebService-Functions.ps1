############################################################
#
#  WEB Service Functions
#
############################################################


Function ConvertTo-HashTable 
	{
	    Param (
		        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
		        [Object]$InputObject,
		        [string[]]$ExcludeTypeName = @("ListDictionaryInternal","Object[]"),
		        [ValidateRange(1,10)][Int]$MaxDepth = 4
	    	)

        $propNames = $InputObject.psobject.Properties | Select-Object -ExpandProperty Name
        $hash = @{}
        $propNames | ForEach-Object {
	            if ($null -ne $InputObject.$_ ) 
					{
		                if ($InputObject.$_ -is [string] -or (Get-Member -MemberType Properties -InputObject ($InputObject.$_) ).Count -eq 0) 
							{$hash.Add($_,$InputObject.$_)} 
						else 
							{
			                    if ($InputObject.$_.GetType().Name -in $ExcludeTypeName){Write-Verbose "Skipped $_"} 
								elseif ($MaxDepth -gt 1){$hash.Add($_,(ConvertTo-HashTable -InputObject $InputObject.$_ -MaxDepth ($MaxDepth - 1)))}
	                		}
	            	}
	        }
	    $hash
	}


Function Read-PostRequest($request) 
	{
 		# Get post data from the input stream.
		# This function by Peter Hinchley
		# https://hinchley.net/articles/create-a-web-server-using-powershell/
		$length = $request.contentlength64
  		$buffer = new-object "byte[]" $length

  		[void]$request.inputstream.read($buffer, 0, $length)
  		$body = [system.text.encoding]::ascii.getstring($buffer)

  		$data = @{}
  		$body.split('&') | ForEach-Object{$part = $_.split('=') ; $data.add($part[0], $part[1])}
  		return $data
	}


Function Read-GetRequest($request) 
	{
		$data = @{}
		$request.QueryString|ForEach-Object {$data.add($_ , $request.QueryString.get_Item($_))}		
		return $data
	}


Function Get-Time
    {
        $Result = Get-Date|select Date, Day, DayOfWeek, DayOfYear, Hour, Kind, Millisecond, Minute, Month, Second, Ticks, TimeOfDay, Year, DateTime
        Return $Result
    }


Function Ping-service
    {
        $now = Get-Date
        $Elapsed = (Get-Date $now) - (Get-Date $StartServiceTime)
        $Result = New-Object PSObject
        Add-Member -InputObject $Result -MemberType NoteProperty -Name Status -Value "Running"
        Add-Member -InputObject $Result -MemberType NoteProperty -Name Days -Value $Elapsed.days														
        Add-Member -InputObject $Result -MemberType NoteProperty -Name Hours -Value $Elapsed.Hours
        Add-Member -InputObject $Result -MemberType NoteProperty -Name Minutes -Value $Elapsed.Minutes
        Return $Result 
    }


function New-ScriptBlockCallback 
    {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
        param(
            [parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [scriptblock]$Callback
        )

        # Is this type already defined?
        if (-not ( 'CallbackEventBridge' -as [type])) {
            Add-Type @' 
                using System; 
 
                public sealed class CallbackEventBridge { 
                    public event AsyncCallback CallbackComplete = delegate { }; 
 
                    private CallbackEventBridge() {} 
 
                    private void CallbackInternal(IAsyncResult result) { 
                        CallbackComplete(result); 
                    } 
 
                    public AsyncCallback Callback { 
                        get { return new AsyncCallback(CallbackInternal); } 
                    } 
 
                    public static CallbackEventBridge Create() { 
                        return new CallbackEventBridge(); 
                    } 
                } 
'@
        }
        $bridge = [callbackeventbridge]::create()
        Register-ObjectEvent -InputObject $bridge -EventName callbackcomplete -Action $Callback -MessageData $args > $null
        $bridge.Callback
    }




############################################################
#
#  AD Functions
#
############################################################



Function Get-ComputerName 
    {
	    Param (
                [String]$ComputerName,
                [String]$SearchPath
            )

            
        If ([String]::IsNullOrWhiteSpace($ComputerName))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] No value specified for property 'ComputerName'"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 9														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $Null
                Return $ReturnObj 
            }              

        If (-not([String]::IsNullOrWhiteSpace($SearchPath)))
            {$Result = Get-ADComputer -Filter {name -eq $ComputerName} -SearchBase $SearchPath -properties MemberOf}    
        else 
            {$Result = Get-ADComputer -Filter {name -eq $ComputerName} -properties MemberOf}

        If ([String]::IsNullOrWhiteSpace($Result))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] There is no Computer with Name $ComputerName in AD"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 10														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $Null
                Return $ReturnObj 
            }
        else 
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $True
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "Machine $ComputerName Found in AD!!"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 0														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $Result.Name
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerADpath -Value $Result.DistinguishedName
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerSID -Value $Result.SID
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerObjGUID -Value $Result.ObjectGUID
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerMemberOf -Value $Result.MemberOf.Value
                Return $ReturnObj          
            }
    }


Function Get-GroupName
    {
	    Param (
                [String]$GroupName,
                [String]$SearchPath
            )
        
            
        If ([String]::IsNullOrWhiteSpace($GroupName))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] No value specified for Group Name"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 9														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $Null
                Return $ReturnObj 
            } 

        If (-not([String]::IsNullOrWhiteSpace($SearchPath)))
            {$Result = Get-ADGroup -filter 'Name -eq $GroupName' -SearchBase $SearchPath -Properties Members,SID}
        else
            {$Result = Get-ADGroup -filter 'Name -eq $GroupName' -Properties Members,SID}

        If ([String]::IsNullOrWhiteSpace($Result))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] There is no Group with Name $GroupName in AD"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 18														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupName -Value $Null
                Return $ReturnObj 
            }
        else 
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $True
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "Group $GroupName Found in AD!!"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 0														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupName -Value $Result.Name
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupADpath -Value $Result.DistinguishedName
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupSID -Value $Result.SID
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupObjGUID -Value $Result.ObjectGUID
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupMembers -Value $Result.Members.Value
                Return $ReturnObj              
            }    
    }


Function Get-OuName
    {
	    Param (
                [String]$OuName,
                [String]$SearchPath
	    	)

            
        If ([String]::IsNullOrWhiteSpace($OuName))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] No value specified for OU Name"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 9														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $Null
                Return $ReturnObj 
            }

        If (-not([String]::IsNullOrWhiteSpace($SearchPath)))
            {$Result = Get-ADOrganizationalUnit -Filter 'Name -eq $OuName' -SearchBase $SearchPath -ErrorAction SilentlyContinue}
        else
            {$Result = Get-ADOrganizationalUnit -Filter 'Name -eq $OuName' -ErrorAction SilentlyContinue}

        If ([String]::IsNullOrWhiteSpace($Result))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] There is no OU with Name $OuName"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 13														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupName -Value $Null
                Return $ReturnObj 
            }
        else 
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $True
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "Ou $OuName Found !!"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 0														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name OuName -Value $Result.Name
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name OuADpath -Value $Result.DistinguishedName
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name OuObjGUID -Value $Result.ObjectGUID
                Return $ReturnObj              
            }    
    }


Function Add-NewComputer
	{
	    Param (
		        [String]$ComputerName,
		        [string]$OU
	    	)

        ##== Validate Computer name
        If ($ComputerName.Length -gt 15)
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] Computer Name $ComputerName is more than 15 Caraters"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 11														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $Null
                Return $ReturnObj 
            }


        $Result = Get-ComputerName -ComputerName $ComputerName
        If (-not([String]::IsNullOrWhiteSpace($Result.ComputerName)))
            {
                $Result.Success = $False
                $Result.Message = "[ERROR] Computer Name Already exists"
                $Result.ErrorNumber = 12
                Return $Result
            }

        ##== Create Computername
        If ([String]::IsNullOrWhiteSpace($OU))
            {New-ADComputer -Name $ComputerName -SAMAccountName $ComputerName -Enabled $true}
        Else
            {
                ##== Check OU
                $Result = Get-ADOrganizationalUnit -Filter 'Name -like $OU'
                If (-not([String]::IsNullOrWhiteSpace($Result)))
                    {New-ADComputer -Name $ComputerName -SAMAccountName $ComputerName -Path $Result.DistinguishedName -Enabled $true}
                else
                    {
                        $ReturnObj = New-Object PSObject
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] Specified OU does not exists!!!"
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 13														
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $Null
                        Return $ReturnObj 
                    }
            }
      
        ##== Return Result
        $Result = Get-ComputerName -ComputerName $ComputerName
        $Result.Message = "Machine created successfully !!"
        $Result.ErrorNumber = 0
        Return $Result  
    }


Function Add-NewGeneratedComputer
	{ 
	    Param (
		        [String]$Prefix,
                [string]$Suffix,
                [Int]$Digits,
                [string]$OU
            )
            
        If (([String]::IsNullOrWhiteSpace($Prefix)) -and ([String]::IsNullOrWhiteSpace($Suffix)))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] Parameters Prefix and Suffix can not both be empty, One parameter must be submitted at least!!!"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 15														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $Null
                Return $ReturnObj 
            } 

        If ([String]::IsNullOrWhiteSpace($Digits))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] Parameter digits missing !!!"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 16														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $Null
                Return $ReturnObj 
            } 
  
        ##== Create Computername
        If (-not ([String]::IsNullOrWhiteSpace($Prefix))) {$AdFilter = "$Prefix*" ; [Int]$TmpLength = $Prefix.Length}
        If (-not ([String]::IsNullOrWhiteSpace($Suffix))) {$AdFilter = "*$Suffix" ; [Int]$TmpLength = $Suffix.Length}
        
        ##== Change digits into friendly powershell something
        [String]$Digits = "{0:d$Digits}"
        [String]$DigitsAsString = $Digits -f 0
        $NameLength = $TmpLength + $DigitsAsString.Length
        [String]$RegexRule = "^.{$NameLength,$NameLength}$"

        
        [Int]$I = 0
        Do
            {
                $I += 1
                #$NewComputerName = (Get-ADComputer -Filter {name -like $AdFilter}).name|Select-Object -Last 1
                $NewComputerName = (Get-ADComputer -properties Name -filter *| Where-object { $_.Name -like $AdFilter -and $_.Name -match $RegexRule}).name|Select-Object -Last 1
                
                ##== Manage case when this is the very first computer to create
                If ([String]::IsNullOrWhiteSpace($NewComputerName)){$NewComputerName = $($AdFilter.replace("*",$DigitsAsString)) }

                ##== Cleanup name from strings to allow number incrementation
                If (-not ([String]::IsNullOrWhiteSpace($Prefix))) {$NewComputerName = $NewComputerName.replace($Prefix,"")}
                If (-not ([String]::IsNullOrWhiteSpace($Suffix))) {$NewComputerName = $NewComputerName.replace($Suffix,"")}
                
                ##== Increment number, and format with requested Digits
                [int]$NewComputerName = $NewComputerName ; $NewComputerName += $I ; [String]$NewComputerName = $Digits -f $NewComputerName
                
                ##== Rebuild new name
                If (-not ([String]::IsNullOrWhiteSpace($Prefix))) {$NewComputerName = $($Prefix + $NewComputerName) }
                If (-not ([String]::IsNullOrWhiteSpace($Suffix))) {$NewComputerName = $($NewComputerName + $Suffix) }
                    
            }while ($null -ne ((Get-ADComputer -Filter {name -eq  $NewComputerName}).name))	
        
        If ([String]::IsNullOrWhiteSpace($OU))    
            {$Result = Add-NewComputer -ComputerName $NewComputerName}
        else 
            {$Result = Add-NewComputer -ComputerName $NewComputerName -OU $OU}           
        Return $Result

    }


Function Add-ComputerToOU
    {
	    Param (
		        [String]$ComputerName,
		        [string]$OU
            )
            
        ##== Check Computer Name
        $CompResult = Get-ComputerName -ComputerName $ComputerName
        If ($CompResult.Success -eq $False){Return $CompResult}

        ##== Check OU
        $OuResult = Get-OuName -OuName $OU
        If ($OuResult.Success -eq $False){Return $OuResult}

        ##== Check If Computer is not already a member of the OU 
        $Result = Get-ComputerName -ComputerName $ComputerName -SearchPath $OuResult.OuADpath
        If ($Result.Success -eq $True)
            {$Message = "Computer $ComputerName is already a member of the OU $OU"}
        Else
            {
                ##== Move Computer to target OU
                Move-ADObject -Identity $CompResult.ComputerADpath -TargetPath $OuResult.OuADpath -ErrorAction SilentlyContinue

                ##== Check if Move was successful !!    
                $Result = Get-ComputerName -ComputerName $ComputerName -SearchPath $OuResult.OuADpath
                If ($Result.Success -eq $False)
                    {
                        $Result.Message = "[ERROR] Unable to move Computer $ComputerName to OU $OU"
                        $Result.ErrorNumber = 17
                        Return $Result
                    }
                else
                    {$Message = "Computer $ComputerName moved successfully to OU $OU"}     
            }

        $Result.Message = $Message
        $Result.ErrorNumber =  0
        Add-Member -InputObject $Result -MemberType NoteProperty -Name OuName -Value $OuResult.OuName
        Add-Member -InputObject $Result -MemberType NoteProperty -Name OuAdPath -Value $OuResult.OuADpath
        Add-Member -InputObject $Result -MemberType NoteProperty -Name OuGUID -Value $OuResult.OuObjGUID
        Return $Result                                    
    }


Function Add-ComputerToGroup
    {
        Param (
            [String]$ComputerName,
            [string]$Group
        )

        ##== Check Computer Name
        $CompResult = Get-ComputerName -ComputerName $ComputerName
        If ($CompResult.Success -eq $False){Return $CompResult}        

        ##== Check Computer Group
        $GrpResult = Get-GroupName -GroupName $Group
        If ($GrpResult.Success -eq $False){Return $GrpResult} 

        ##== Check If Computer is not already a member of the Group 
        If ($GrpResult.GroupMembers -contains $CompResult.ComputerADpath)
            {$Message = "Computer $ComputerName is already a member of group $Group"}
        else 
            {
                ##== Add computer to Group
                Add-ADGroupMember $Group -Members $($ComputerName + "$")

                ##== Check If Computer has successfully joined the group
                $GrpResult = Get-GroupName -GroupName $Group

                If ($GrpResult.GroupMembers -contains $CompResult.ComputerADpath)
                    {
                        ##== Check again to send back updated information
                        $CompResult = Get-ComputerName -ComputerName $ComputerName
                        
                        $ReturnObj = New-Object PSObject
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $True
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "Computer $ComputerName was successfully added to group $Group"
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 0
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerName -Value $CompResult.ComputerName
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerADpath -Value $CompResult.ComputerADpath
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerSID -Value $CompResult.ComputerSID
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerObjGUID -Value $CompResult.ComputerObjGUID
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerMemberOf -Value $CompResult.ComputerMemberOf
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupName -Value $GrpResult.GroupName
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupADpath -Value $GrpResult.GroupADpath
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupSID -Value $GrpResult.GroupSID
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupObjGUID -Value $GrpResult.GroupObjGUID
                        Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name GroupMembers -Value $GrpResult.GroupMembers               														
                        Return $ReturnObj
                    }
                else 
                    {
                        $GrpResult.Success = $False
                        $GrpResult.Message = "[ERROR] Unable to move Computer $ComputerName to group $group"
                        $GrpResult.ErrorNumber = 19
                        Add-Member -InputObject $GrpResult -MemberType NoteProperty -Name ComputerName -Value $ComputerName
                        Return $GrpResult                         
                    }                
            }
    }



############################################################
#
#  SCCM Functions
#
############################################################


Function Get-CMComputerCollection
    {
       
        Param (
                [String]$Collection
            )        


        ##== INIT
        If ([String]::IsNullOrWhiteSpace($Script:CMDrive)){$Script:CMDrive = Get-CMSite}
        If ([String]::IsNullOrWhiteSpace($Script:CurrentLocation)){$Script:CurrentLocation = Get-Location}


        
        If ([String]::IsNullOrWhiteSpace($Collection))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] No value specified for Collection Name"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 20														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name CollectionName -Value $Null
                Return $ReturnObj 
            }              

        ##==Change Location to SCCM
        Set-Location $Script:CMDrive.SiteDrive
               
        $Result = Get-WmiObject -computername $Script:CMDrive.SiteServer -Namespace $Script:CMDrive.WMInameSpace -Query "select * from SMS_Collection Where SMS_Collection.Name='$Collection'"     

        If ([String]::IsNullOrWhiteSpace($Result))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] There is no Collection with Name $Collection"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 21														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name CollectionName -Value $Null
                Set-Location $Script:CurrentLocation
                Return $ReturnObj 
            }
        else 
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $True
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "Collection $Collection Found !!"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 0														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name CollectionName -Value $Result.Name
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name CollectionID -Value $Result.CollectionID
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name CollectionMember -Value $Result.MemberCount
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name CollectionLastRefreshTime -Value $([datetime]::parseexact(($Result.LastRefreshTime).split('.')[0],"yyyyMMddHHmmss",[System.Globalization.CultureInfo]::InvariantCulture))
                Set-Location $Script:CurrentLocation
                Return $ReturnObj          
            }
    }


Function Get-CMComputerName 
    {
	    Param (
                [String]$ComputerName
            )

        ##== INIT
        If ([String]::IsNullOrWhiteSpace($Script:CMDrive)){$Script:CMDrive = Get-CMSite}
        If ([String]::IsNullOrWhiteSpace($Script:CurrentLocation)){$Script:CurrentLocation = Get-Location}
         
            
        If ([String]::IsNullOrWhiteSpace($ComputerName))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] No value specified for Computer Name"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 22														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMName -Value $Null
                Return $ReturnObj 
            }
            
            
         ##==Change Location to SCCM
        Set-Location $Script:CMDrive.SiteDrive
                                 
        $Result = Get-WmiObject -computername $Script:CMDrive.SiteServer -Namespace $Script:CMDrive.WMInameSpace -Class SMS_R_System -Filter "NetbiosName='$ComputerName'" -ErrorAction SilentlyContinue

        If ([String]::IsNullOrWhiteSpace($Result))
            {
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $False
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "[ERROR] There is no Computer with Name $ComputerName In SCCM"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 23														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMName -Value $Null
                Set-Location $Script:CurrentLocation
                Return $ReturnObj 
            }
        else 
            {
                               
                $CollectionsMemberchip = (Get-WmiObject -computername $Script:CMDrive.SiteServer -Namespace $Script:CMDrive.WMInameSpace -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$ComputerName' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID" -ErrorAction SilentlyContinue).Name 
                $ClientActiveStatus = (Get-WmiObject -computername $Script:CMDrive.SiteServer -Namespace $Script:CMDrive.WMInameSpace -query "SELECT SMS_G_System_CH_ClientSummary.* FROM SMS_G_System_CH_ClientSummary, SMS_R_System WHERE name = '$ComputerName'and SMS_G_System_CH_ClientSummary.ResourceId = SMS_R_System.ResourceId" -ErrorAction SilentlyContinue).ClientActiveStatus
                                
                $ReturnObj = New-Object PSObject
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Success -Value $True
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name Message -Value "Machine $ComputerName Found In SCCM!!"
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ErrorNumber -Value 0														
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMName -Value $Result.Name
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMADpath -Value $Result.DistinguishedName
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMResourceId -Value $Result.ResourceId
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMSID -Value $Result.SID
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMWindowsBuild -Value $Result.Build
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMClientVersion -Value $Result.ClientVersion
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMIPAddresses -Value $($Result.IPAddresses -join ",")
                Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMMACAddresses -Value $($Result.MACAddresses -join ",")
                If ( -not [String]::IsNullOrWhiteSpace($CollectionsMemberchip)){Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMCollectionMemberchip -Value $($CollectionsMemberchip -join ",")}
                If ( -not [String]::IsNullOrWhiteSpace($ClientActiveStatus)){Add-Member -InputObject $ReturnObj -MemberType NoteProperty -Name ComputerCMClientIsActive -Value $([system.convert]::ToBoolean($ClientActiveStatus))}
                Set-Location $Script:CurrentLocation
                Return $ReturnObj          
            }
    }


Function Add-CMComputerToCollection
    {
	    Param (
                [String]$ComputerName,
                [String]$Collection
            )
 

        ##== INIT
        If ([String]::IsNullOrWhiteSpace($Script:CMDrive)){$Script:CMDrive = Get-CMSite}
        If ([String]::IsNullOrWhiteSpace($Script:CurrentLocation)){$Script:CurrentLocation = Get-Location}

                
        ##==Change Location to SCCM
        Set-Location $Script:CMDrive.SiteDrive
             
        ##== Check Computer Name
        $CompResult = Get-CMComputerName -ComputerName $ComputerName
        If ($CompResult.Success -eq $False){Set-Location $Script:CurrentLocation ; Return $CompResult}        

        ##== Check Collection
        $GrpResult = Get-CMComputerCollection -Collection $Collection
        If ($GrpResult.Success -eq $False){Set-Location $Script:CurrentLocation ; Return $GrpResult} 

        ##== Check If Computer is not already a member of the Collection 
        If ($CompResult.ComputerCMCollectionMemberchip -contains $GrpResult.CollectionName)
            {
                $CompResult.Success = $False
                $CompResult.Message = "[ERROR] Computer $ComputerName is already a member of collection $Collection"
                $CompResult.ErrorNumber = 24
                Add-Member -InputObject $CompResult -MemberType NoteProperty -Name CollectionName -Value $GrpResult.CollectionName
                Add-Member -InputObject $CompResult -MemberType NoteProperty -Name CollectionID -Value $GrpResult.CollectionID
                Add-Member -InputObject $CompResult -MemberType NoteProperty -Name CollectionMember -Value $GrpResult.CollectionMember 
                Add-Member -InputObject $CompResult -MemberType NoteProperty -Name CollectionLastRefreshTime -Value $GrpResult.CollectionLastRefreshTime 
                Set-Location $Script:CurrentLocation
                Return $CompResult
            }
        else 
            {
                ##== Add computer to collection
                ### This code section by Kaido & Keith 
                ## More info here https://keithga.wordpress.com/2018/01/25/a-replacement-for-sccm-add-cmdevicecollectiondirectmembershiprule-powershell-cmdlet/
                ## Code taken from https://github.com/keithga/CMPSLib/blob/master/PSCMLib/Collections/Add-CMDeviceToCollection.ps1
                # Note to self: method used here is "AddMembershipRule", not "AddMembershipRule(s)" used in keith code (no array needed !)   
                $CollectionQuery = Get-WmiObject -computername $Script:CMDrive.SiteServer -Namespace $Script:CMDrive.WMInameSpace -Class SMS_Collection -filter "Name='$Collection' and CollectionType='2'"
                $InParams = $CollectionQuery.PSBase.GetMethodParameters('AddMembershipRule')

                $cls = Get-WmiObject -computername $Script:CMDrive.SiteServer -Namespace $Script:CMDrive.WMInameSpace -Class SMS_CollectionRuleDirect -list
                $NewRule = $cls.CreateInstance()
                $NewRule.ResourceClassName = "SMS_R_System"
                $NewRule.ResourceID = $CompResult.ComputerCMResourceId
                $NewRule.Rulename = $CompResult.ComputerCMName
                $Rules = $NewRule.psobject.BaseObject

                $InParams.CollectionRule = $Rules.psobject.BaseOBject
                $PreviousEval = $([datetime]::parseexact(($CollectionQuery.LastMemberChangeTime).split('.')[0],"yyyyMMddHHmmss",[System.Globalization.CultureInfo]::InvariantCulture))
                $CollectionQuery.PSBase.InvokeMethod('AddMembershipRule',$InParams,$null) | Out-Null
                $CollectionQuery.RequestRefresh() | Out-Null

                ##== Check If Computer has successfully joined the collection by updating collection membership
                $NewEval = $([datetime]::parseexact(($((Get-WmiObject -computername $Script:CMDrive.SiteServer -Namespace $Script:CMDrive.WMInameSpace -Class SMS_Collection -filter "Name='$Collection' and CollectionType='2'" -Property LastMemberChangeTime).LastMemberChangeTime)).split('.')[0],"yyyyMMddHHmmss",[System.Globalization.CultureInfo]::InvariantCulture))
                While ($PreviousEval -ge $NewEval)
                    {
                        start-sleep -Seconds 2
                        $NewEval = $([datetime]::parseexact(($((Get-WmiObject -computername $Script:CMDrive.SiteServer -Namespace $Script:CMDrive.WMInameSpace -Class SMS_Collection -filter "Name='$Collection' and CollectionType='2'" -Property LastMemberChangeTime).LastMemberChangeTime)).split('.')[0],"yyyyMMddHHmmss",[System.Globalization.CultureInfo]::InvariantCulture))
                    }
                
                ##== Get updated infos
                $CompResult = Get-CMComputerName -ComputerName $ComputerName

                If ($CompResult.ComputerCMCollectionMemberchip -like "*$($GrpResult.CollectionName)*")
                    {
                        $GrpResult = Get-CMComputerCollection -Collection $Collection
                        $CompResult.Message = "Computer $ComputerName was successfully added to collection $Collection"
                        Add-Member -InputObject $CompResult -MemberType NoteProperty -Name CollectionName -Value $GrpResult.CollectionName
                        Add-Member -InputObject $CompResult -MemberType NoteProperty -Name CollectionID -Value $GrpResult.CollectionID
                        Add-Member -InputObject $CompResult -MemberType NoteProperty -Name CollectionMember -Value $GrpResult.CollectionMember 
                        Add-Member -InputObject $CompResult -MemberType NoteProperty -Name CollectionLastRefreshTime -Value $GrpResult.CollectionLastRefreshTime 
                        Set-Location $Script:CurrentLocation
                        Return $CompResult
                    }
                else 
                    {
                        $CompResult.Success = $False
                        $CompResult.Message = "[ERROR] Unable to move Computer $ComputerName to collection $Collection"
                        $CompResult.ErrorNumber = 25
                        Set-Location $Script:CurrentLocation
                        Return $CompResult                         
                    }                
            }
    }