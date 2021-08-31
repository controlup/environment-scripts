<#
.SYNOPSIS
    Creates the folder structure and adds/removes or moves machines into the structure.

.DESCRIPTION
    Creates the folder structure and adds/removes or moves machines into the structure.

.PARAMETER FolderPath
    The target folder path in ControlUp to save these Objects.

.PARAMETER Preview
    Shows the expected results without committing any changes to the ControlUp environment
    
.PARAMETER Delete
    Enables the script to execute object removals. Use with -Preview to see the proposed changes without committing changes.

.PARAMETER LogFile
    Tells the script to log the output to a text file. Can be used with -Preview to see the proposed changes.

.PARAMETER Site
    Specify a ControlUp Monitor Site by name to assign the objects.

.PARAMETER Brokers
    A list of brokers to contact for Delivery Groups and Computers to sync. Multiple brokers can be specified if seperated by commas.

.PARAMETER maxRecordCount
    The maximum number of machines to retrieve. If not specified and the default is exceeded, the script will figure out the total to use

.PARAMETER includeDeliveryGroup
    Include only these specific Delivery Groups to be added to the ControlUp tree. Wild cards are supported as well, so if you have 
    Delivery Groups named like "Epic North", "Epic South" you can specify "Epic*" and it will capture both. Multiple delivery groups
    can be specified if they are seperated by commas. If you also use the parameter "excludeDeliveryGroups" then the exclude will
    supersede any includes and remove any matching Delivery Groups. Omitting this parameter and the script will capture all detected
    Delivery Groups.

.PARAMETER excludeDeliveryGroup
    Exclude specific delivery groups from being added to the ControlUp tree. Wild cards are supported as well, so if you have 
    Delivery Groups named like "Epic CGY", "Cerner CGY" you can specify "*CGY" and it will exclude any that ends with CGY. Multiple 
    delivery groups can be specified if they are seperated by commas. If you also use the parameter "includeDeliveryGroup" then the exclude will
    supersede any includes and remove any matching Delivery Groups.

.PARAMETER addBrokersToControlUp
    Add brokers to the ControlUp Tree. This optional parameter can be specified if you prefer this script to add broker machines as they
    are detected. If this parameter is omitted then Broker machines will not be moved or added to your ControlUp tree.

.PARAMETER enabledOnly
    Only include Delivery Groups which are enabled

.PARAMETER MatchEUCEnvTree
    Configures the script to match the same structure used by ControlUp for the EUC Environment Tree. If this parameter is omitted the
    Delivery Group is added to the FolderPath. 
    With this parameter enabled, the structure is like so:
    $SiteName -|
               |-Delivery Groups -|
               |                  |-DG1 -|
               |                         |-Machine001
               |-Brokers -|
                          |-Broker001

.PARAMETER batchCreateFolders
    Create folders in batches rather than sequentially

.PARAMETER force
    When the number of new folders to create is large, force the operation to continue otherwise it will abort before commencing

.PARAMETER SmtpServer
    Name/IP address of an SMTP server to send email alerts via. Optionally specify : and a port number if not the default of 25

.PARAMETER emailFrom
    Email address from which to send email alerts from

.PARAMETER emailTo
    Email addresses to which to send email alerts to

.PARAMETER emailUseSSL
    Use SSL for SMTP server communication

.EXAMPLE
    . .\CU_SyncScript.ps1 -Brokers "ddc1.bottheory.local","ctxdc01.bottheory.local" -folderPath "CUSync\Citrix" -includeDeliveryGroup "EpicNorth","EpicSouth","EpicCentral","Cerner*" -excludeDeliveryGroup "CernerNorth" -addBrokersToControlUp -MatchEUCEnvTree
        Contacts the brokers ddc1.bottheory.local and ctxdc01.bottheory.local, it will save the objects to the ControlUp folder 
        "CUSync\Citrix", only include specific Delivery Groups including all Delivery Groups that start wtih Cerner and exclude
        the Delivery Group "CernerNorth", adds the broker machines to ControlUp, have the script match the same structure as
        the ControlUp EUC Environment.

.EXAMPLE
    . .\CU_SyncScript.ps1 -Brokers "ddc1.bottheory.local" -folderPath "CUSync"
        Contacts the brokers ddc1.bottheory.local and adds all Delivery Groups and their machines to ControlUp under the folder "CUSync"

.CONTEXT
    Citrix

.MODIFICATION_HISTORY
    Trentent Tye,         2020-08-06 - Original Code
    Guy Leech,            2020-09-16 - Fix missing siteid hashtable, changed from dot sourcing common module from cwd to same path as this sript
    Guy Leech,            2020-10-09 - Added parameter -enabledOnly to only include Delivery Groups which are enabled
    Guy Leech,            2020-10-13 - Accommodate Build-CUTree returning error count
    Guy Leech,            2020-10-30 - Fixed bug where -Adminaddress not passed to Get-BrokerDesktopGroup
    Guy Leech,            2020-11-02 - Added -batchCreateFolders option to create folders in batches (faster) otherwise creates them one at a time
    Guy Leech,            2021-02-12 - Added -force for when large number of folders to add
    Guy Leech,            2021-07-30 - Added email alerting and writing to logfile
    Guy Leech,            2021-08-20 - Fixed issue where dash in Citrix site name caused 2 different folders to be created
    Guy Leech,            2021-08-25 - Fixed manual merge errors and possible flattening of delivery group include and exclude arrays. Errors if no objects to sync found.
                                       Extra logging for inclusions, exclusions and module import failures
    Guy Leech,            2021-08-20 - Added support for -MaxRecordCount & checking it there are more machines if not specified

.LINK

.COMPONENT

.NOTES
    Requires rights to read Citrix environment.

    Version:        0.1
    Author:         Trentent Tye
    Creation Date:  2020-08-06
    Updated:        2020-08-06
                    Changed ...
    Purpose:        Created for Citrix Sync
#>

[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$true, HelpMessage='Enter a ControlUp subfolder to save your Citrix tree' )]
    [ValidateNotNullOrEmpty()]
    [string] $folderPath,

    [Parameter(Mandatory=$false, HelpMessage='Preview the changes' )]
    [switch] $Preview,

    [Parameter(Mandatory=$false, HelpMessage='Execute removal operations. When combined with preview it will only display the proposed changes' )]
    [switch] $Delete,

    [Parameter(Mandatory=$false, HelpMessage='Enter a path to generate a log file of the proposed changes' )]
    [string] $LogFile,

    [Parameter(Mandatory=$false, HelpMessage='Enter a ControlUp Site' )]
    [ValidateNotNullOrEmpty()]
    [string] $Site,

    [Parameter(Mandatory=$false, HelpMessage='Creates the ControlUp folder structure based on the EUC Environment tree' )]
    [switch] $MatchEUCEnvTree,

    [Parameter(Mandatory=$true,  HelpMessage='A list of Brokers to connect and pull data' )]
    [ValidateNotNullOrEmpty()]
    [array] $Brokers,
    
    [Parameter(Mandatory=$false, HelpMessage='Maximum number of items to request from broker' )]
    [int] $maxRecordCount,

    [Parameter(Mandatory=$false, HelpMessage='A list of Delivery Groups to include.  Works with wildcards' )]
    [array] $includeDeliveryGroup,

    [Parameter(Mandatory=$false, HelpMessage='A list of Delivery Groups to exclude.  Works with wildcards. Exclusions supercede inclusions' )]
    [array] $excludeDeliveryGroup,

    [Parameter(Mandatory=$false, HelpMessage='Adds the Citrix Brokers to the ControlUp tree' )]
    [switch] $addBrokersToControlUp ,

    [Parameter(Mandatory=$false, HelpMessage='Only adds Delivery Groups which are enabled' )]
    [switch] $enabledOnly ,

    [Parameter(Mandatory=$false, HelpMessage='Create folders in batches rather than individually' )]
	[switch] $batchCreateFolders ,

    [Parameter(Mandatory=$false, HelpMessage='Force folder creation if number exceeds safe limit' )]
	[switch] $force ,

    [Parameter(Mandatory=$false, HelpMessage='Smtp server to send alert emails from' )]
	[string] $SmtpServer ,

    [Parameter(Mandatory=$false, HelpMessage='Email address to send alert email from' )]
	[string] $emailFrom ,

    [Parameter(Mandatory=$false, HelpMessage='Email addresses to send alert email to' )]
	[string[]] $emailTo ,

    [Parameter(Mandatory=$false, HelpMessage='Use SSL to send email alert' )]
	[switch] $emailUseSSL
) 

## GRL this way allows script to be run with debug/verbose without changing script
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

##$Global:LogFile = $PSBoundParameters.LogFile

## Script from ControlUp which must reside in the same folder as this script
[string]$buildCuTreeScript = 'Build-CUTree.ps1'

function Make-NameWithSafeCharacters ([string]$string) {
    ###### TODO need to replace the folder path characters that might be illegal
    #list of illegal characters : '/', '\', ':', '*','?','"','<','>','|','{','}'
    $returnString = (($string).Replace("/","-")).Replace("\","-").Replace(":","-").Replace("*","-").Replace("?","-").Replace("`"","-").Replace("<","-").Replace(">","-").Replace("|","-").Replace("{","-").Replace("}","-")
    return $returnString
}
    
#Create ControlUp structure object for synchronizing
class ControlUpObject{
    [string]$Name
    [string]$FolderPath
    [string]$Type
    [string]$Domain
    [string]$Description
    [string]$DNSName
        ControlUpObject ([String]$Name,[string]$folderPath,[string]$type,[string]$domain,[string]$description,[string]$DNSName) {
        $this.Name = $Name
        $this.FolderPath = $folderPath
        $this.Type = $type
        $this.Domain = $domain
        $this.Description = $description
        $this.DNSName = $DNSName
    }
}

# dot sourcing Functions

## GRL Don't assume user has changed location so get the script path instead
[string]$thisScript = & { $myInvocation.ScriptName }
[string]$scriptPath = Split-Path -Path $thisScript -Parent
[string]$buildCuTreeScriptPath = [System.IO.Path]::Combine( $scriptPath , $buildCuTreeScript )
[string]$errorMessage = $null

if( ! ( Test-Path -Path $buildCuTreeScriptPath -PathType Leaf -ErrorAction SilentlyContinue ) )
{
    Throw "Unable to find script `"$buildCuTreeScript`" in `"$scriptPath`""
}

. $buildCuTreeScriptPath

# Add required Citrix cmdlets

## new CVAD have modules so use these in preference to snapins which are there for backward compatibility
if( ! (  Import-Module -Name Citrix.DelegatedAdmin.Commands -ErrorAction SilentlyContinue -PassThru -Verbose:$false) `
    -and ! ( Add-PSSnapin -Name Citrix.Broker.Admin.* -ErrorAction SilentlyContinue -PassThru -Verbose:$false) )
{
    $errorMessage = 'Failed to load Citrix PowerShell cmdlets - is this a Delivery Controller or have Studio or the PowerShell SDK installed ?'
    Write-CULog -Msg $errorMessage -ShowConsole
    Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$thisScript`" on $env:COMPUTERNAME" -body "$errorMessage"
    Throw $errorMessage
}

Write-Host "Brokers: $Brokers"
$DeliveryGroups = New-Object System.Collections.Generic.List[PSObject]
$BrokerMachines = New-Object System.Collections.Generic.List[PSObject]
$CTXSites       = New-Object System.Collections.Generic.List[PSObject]

[hashtable]$brokerParameters = @{ }
if( $enabledOnly )
{
    $brokerParameters.Add( 'Enabled' , $true )
}

if ($brokers.count -eq 1 -and $brokers[0].IndexOf(',') -ge 0) {
    $brokers = @( $brokers -split ',' )
}

Try {
    foreach ($adminAddr in $brokers) {
        $brokerParameters.AdminAddress = $adminAddr
        $CTXSite = Get-BrokerSite -AdminAddress $adminAddr
        $CTXSites.Add($CTXSite)
        Write-Verbose -Message "Querying $adminAddr for Delivery Groups"
        #Get list of Delivery Groups
        foreach ($DeliveryGroup in $(Get-BrokerDesktopGroup @brokerParameters)) {
            if ($DeliveryGroups.Count -eq 0) {
                $DeliveryGroupObject = [PSCustomObject]@{
                        MachineName         = ""
                        DNSName             = ""
                        Name                = $DeliveryGroup.Name
                        Site                = $CTXSite.Name
                        Broker              = $adminAddr
                    }
                    $DeliveryGroups.Add($DeliveryGroupObject)
            } else {
                if (-not($DeliveryGroups.Name.Contains($DeliveryGroup.Name))) {  #ensures we don't add duplicate delivery groups so you can specify multiple brokers incase one goes down.
                    Write-Verbose -Message "Add $($DeliveryGroup.Name)"
                    $DeliveryGroupObject = [PSCustomObject]@{
                        MachineName         = ""
                        DNSName             = ""
                        Name                = $DeliveryGroup.Name
                        Site                = $CTXSite.Name
                        Broker              = $adminAddr
                    }
                    $DeliveryGroups.Add($DeliveryGroupObject)
                }
            }
        }

        Write-Verbose -Message "Querying $adminAddr for Broker Machines"
        foreach ($BrokerMachine in Get-BrokerController -AdminAddress $adminAddr) {
            if ($BrokerMachines.Count -eq 0) {
                $BrokerMachineObject = [PSCustomObject]@{
                        MachineName         = $BrokerMachine.MachineName
                        DNSName             = $BrokerMachine.DNSName
                        Name                = ""
                        Site                = $CTXSite.Name
                        Broker              = $adminAddr
                    }
                    $BrokerMachines.Add($BrokerMachineObject)
            } else {
                if (-not($BrokerMachines.MachineName.Contains($BrokerMachine.MachineName))) {  #ensures we don't add duplicate broker machines so you can specify multiple brokers incase one goes down.
                    Write-Verbose -Message "Add $($BrokerMachine.MachineName)"
                    $BrokerMachineObject = [PSCustomObject]@{
                        MachineName         = $BrokerMachine.MachineName
                        DNSName             = $BrokerMachine.DNSName
                        Name                = ""
                        Site                = $CTXSite.Name
                        Broker              = $adminAddr
                    }
                    $BrokerMachines.Add($BrokerMachineObject)
                }
            }
        }
    }
}
catch
{
    Write-CULog -Msg $_ -ShowConsole -Type E
    Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$thisScript`" on $env:COMPUTERNAME" -body $_
}

Write-Host "Total Number of Delivery Groups : $($DeliveryGroups.Count)"
Write-Host "Total Number of Brokers         : $($BrokerMachines.Count)"
Write-Host "Total Number of Sites           : $($($($CTXSites.Name | Sort-Object -Unique) | Measure-Object).count)"

#Add Included Delivery Groups
if ($PSBoundParameters.ContainsKey("includeDeliveryGroup")) {

    if ($includeDeliveryGroup.count -eq 1 -and $includeDeliveryGroup[0].IndexOf(',') -ge 0) {
        $includeDeliveryGroup = @( $includeDeliveryGroup -split ',' )
    }

    $DeliveryGroupsInclusionList = New-Object System.Collections.Generic.List[PSObject]
    foreach ($DeliveryGroupInclusion in $includeDeliveryGroup) {
        foreach ($DeliveryGroup in $DeliveryGroups) {
            if ($DeliveryGroup.Name -like $DeliveryGroupInclusion) {
                [string]$message = "Including delivery group: $($DeliveryGroup.Name) - because: $DeliveryGroupInclusion"
                Write-CULog -Msg $message -ShowConsole
                Write-Verbose -Message $message
                $DeliveryGroupsInclusionList.Add($DeliveryGroup)
            }
        }
    }
    #Set DeliveryGroups Object to be the filtered one
    $DeliveryGroups = $DeliveryGroupsInclusionList
    Write-Verbose -Message "Total Number of Delivery Groups after filtering for inclusions: $($DeliveryGroups.Count)"
}

#Remove Excluded Delivery Groups
if ($PSBoundParameters.ContainsKey("excludeDeliveryGroup")) {
    if ($excludeDeliveryGroup.count -eq 1 -and $excludeDeliveryGroup[0].IndexOf(',') -ge 0) {
        $excludeDeliveryGroup = @( $excludeDeliveryGroup -split ',' )
    }

    $IndexesToRemove = New-Object System.Collections.Generic.List[PSObject]
    foreach ($DeliveryGroup in $DeliveryGroups) {
        foreach ($exclusion in $excludeDeliveryGroup) {
            if ($DeliveryGroup.Name -like $exclusion) {
                [string]$message = "Excluding delivery group: $($DeliveryGroup.Name)  - because: $exclusion"
                Write-CULog -Msg $message -ShowConsole
                Write-Verbose -Message $message
                $IndexesToRemove.add($DeliveryGroups.IndexOf($DeliveryGroup))
            }
        }
    }
    for ($i = $IndexesToRemove.Count-1; $i -ge 0; $i--) { 
        $DeliveryGroups.RemoveAt($IndexesToRemove[$i])
    }
}
Write-Verbose -Message "Total Number of Delivery Groups after filtering for exclusions: $($DeliveryGroups.Count)"

Write-Host "Adding Delivery Groups to ControlUp Environmental Object"
$ControlUpEnvironmentObject = New-Object System.Collections.Generic.List[PSObject]
foreach ($DeliveryGroup in $DeliveryGroups) {
    if( $newObject = [ControlUpObject]::new($($DeliveryGroup.Name) ,"$($DeliveryGroup.Name)","Folder","","$($DeliveryGroup.site)-DeliveryGroup",""))
    {
        $ControlUpEnvironmentObject.Add( $newObject )
    }
}

[hashtable]$brokerMachineParameters = @{
    ReturnTotalRecordCount = $true
    ErrorVariable = 'recordCount'
    ErrorAction = 'SilentlyContinue '
}

if( $PSBoundParameters[ 'maxrecordcount' ] )
{
    $brokerMachineParameters.Add( 'maxrecordcount' , $maxRecordCount )
}
#Add machines from the delivery group to the environmental object
foreach ($DeliveryGroup in $DeliveryGroups) {
    $recordCount = $null

    $CTXMachines = @( Get-BrokerMachine -DesktopGroupName $DeliveryGroup.Name -AdminAddress $DeliveryGroup.Broker @brokerMachineParameters )
    if( $recordCount -and $recordCount.Count )
    {
        if( $recordCount[0] -match 'Returned (\d+) of (\d+) items' )
        {
            [int]$returned   = $matches[1]
            [int]$totalItems = $matches[2]
            if( $returned -lt $totalItems )
            {
                Write-Warning -Message "Querying $($DeliveryGroup.Broker) again as only got $returned machines out of $totalItems for delivery group $($DeliveryGroup.Name)"
                $brokerMachineParameters[ 'maxrecordcount' ] = $totalItems
                $CTXMachines = @( Get-BrokerMachine -AdminAddress $DeliveryGroup.Broker -DesktopGroupName $DeliveryGroup.Name @brokerMachineParameters )
            }
        }
        else
        {
            Write-Error -Message $recordCount[0]
        }
    }
    else
    {
        Write-Warning -Message "Failed to get total record count from $($DeliveryGroup.Broker), retrieved $($CTXMachines|Measure-Object -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Count) machines"
    }

    ## if failed then array may be $null so no count property
    Write-Verbose -Message "Got $($CTXMachines|Measure-Object -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Count) machines from delivery group $($DeliveryGroup.Name) on $($DeliveryGroup.Broker)"

    foreach ($Machine in $CTXMachines) {
        if ($Machine.MachineName -like "S-1-5*") {
            Write-Host "Detected a machine with a SID for a name. These cannot be added to ControlUp. Skipping: $($DeliveryGroup.Name) - $($Machine.machineName)" -ForegroundColor Yellow
        } else {
            if ([string]::IsNullOrEmpty($machine.DNSName)) {
                $DNSName = $null
            } else {
                $DNSName = $machine.DNSName
            }
            $Domain = $Machine.MachineName.split("\")[0]
            $Name =$Machine.MachineName.split("\")[1]
            if( $newObject = [ControlUpObject]::new( $Name , $DeliveryGroup.Name , "Computer" , $Domain , "$($DeliveryGroup.site)-Machine" , $DNSName ) )
            {
                $ControlUpEnvironmentObject.Add( $newObject )
            }
        }
    }
}

#Add Brokers to ControlUpEnvironmentalObject
if ($addBrokersToControlUp) {
    if (-not($MatchEUCEnvTree)) {  # MatchEUCEnvTree will add environment specific brokers folders
        if( $newObject = [ControlUpObject]::new( 'Brokers' ,'Brokers' , 'Folder' , '' , 'Brokers' , '' ))
        {
            $ControlUpEnvironmentObject.Add( $newObject )
        }
    }
    foreach ($Machine in $BrokerMachines) {
        if ([string]::IsNullOrEmpty($machine.DNSName)) {
            $DNSName = $null
        } else {
            $DNSName = $machine.DNSName
        }
        $Domain = $Machine.MachineName.split("\")[0]
        $Name =$Machine.MachineName.split("\")[1]
        if( $newObject = [ControlUpObject]::new( $Name , 'Brokers' , 'Computer' , $Domain , "$($Machine.site)-BrokerMachine" , $DNSName ))
        {
            $ControlUpEnvironmentObject.Add( $newObject )
        }
    }
}
## TYE
if ($MatchEUCEnvTree) {
    for ($i=0; $i -lt $ControlUpEnvironmentObject.Count; $i++) {
        if ($ControlUpEnvironmentObject[$i].FolderPath -eq "Brokers" -and $ControlUpEnvironmentObject[$i].Type -eq "Computer") {
            $BrokerObj = $BrokerMachines | Where-Object ({$_.MachineName.split("\")[1] -eq $ControlUpEnvironmentObject[$i].Name})
            $ControlUpEnvironmentObject[$i].FolderPath = "$($BrokerObj.Site)\$($ControlUpEnvironmentObject[$i].FolderPath)" #Sets the path to $SiteName\Brokers
        } else {
            ## changed 2021/08/20 GRL as was splitting on dash which meant it broke environments where the Citrix site had a dash in the name
            $ControlUpEnvironmentObject[$i].FolderPath = "$($ControlUpEnvironmentObject[$i].Description -replace '\-(DeliveryGroup|Machine)$')\Delivery Groups\$($ControlUpEnvironmentObject[$i].FolderPath)"
        }
    }
    if ($addBrokersToControlUp) {
        foreach ($CtxSite in $($CTXSites | Sort-Object -Unique)) {
            if( $newObject = [ControlUpObject]::new("Brokers" ,"$($CTXSite.Name)\Brokers","Folder","","Brokers",""))
            {
                $ControlUpEnvironmentObject.Add( $newObject )
            }
        }
    }
}

Write-Debug "$($ControlUpEnvironmentObject | Format-Table | Out-String)"

$BuildCUTreeParams = @{
    CURootFolder = $folderPath
}

if ($Preview) {
    $BuildCUTreeParams.Add("Preview",$true)
}

if ($Delete) {
    $BuildCUTreeParams.Add("Delete",$true)
}

if ($LogFile){
    $BuildCUTreeParams.Add("LogFile",$LogFile)
}

if ($Site){
    $BuildCUTreeParams.Add("SiteName",$Site)
}

if ($batchCreateFolders){
    $BuildCUTreeParams.Add("batchCreateFolders",$true)
}

if ($Force){
    $BuildCUTreeParams.Add("Force",$true)
}

if ($SmtpServer){
    $BuildCUTreeParams.Add("SmtpServer",$SmtpServer)
}

if ($emailFrom){
    $BuildCUTreeParams.Add("emailFrom",$emailFrom)
}

if ($emailTo){
    $BuildCUTreeParams.Add("emailTo",$emailTo)
}

if ($emailUseSSL){
    $BuildCUTreeParams.Add("emailUseSSL",$emailUseSSL)
}

[int]$errorCount = 1

if( $null -eq $ControlUpEnvironmentObject -or $ControlUpEnvironmentObject.Count -eq 0 )
{
    $errorMessage = "No Citrix on-premises objects found to sync with ControlUp from brokers $($Brokers -join ,',')"
    Write-CULog -Msg $errorMessage -ShowConsole -Type E
    Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$thisScript`" on $env:COMPUTERNAME" -body $errorMessage
}
else
{
    $errorCount = Build-CUTree -ExternalTree $ControlUpEnvironmentObject @BuildCUTreeParams
}

Exit $errorCount
