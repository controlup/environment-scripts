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
    Specify a ControlUp Monitor Site to assign the objects.

.PARAMETER Brokers
    A list of brokers to contact for Delivery Groups and Computers to sync. Multiple brokers can be specified if seperated by commas.

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
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        HelpMessage='Enter a ControlUp subfolder to save your Citrix tree'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $folderPath,

    [Parameter(
        Position=1, 
        Mandatory=$false, 
        HelpMessage='Preview the changes'
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $Preview,

    [Parameter(
        Position=2, 
        Mandatory=$false, 
        HelpMessage='Execute removal operations. When combined with preview it will only display the proposed changes'
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $Delete,

    [Parameter(
        Position=3, 
        Mandatory=$false, 
        HelpMessage='Enter a path to generate a log file of the proposed changes'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $LogFile,

    [Parameter(
        Position=4, 
        Mandatory=$false, 
        HelpMessage='Enter a ControlUp Site'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $Site,

    [Parameter(
        Position=5,
        Mandatory=$false, 
        HelpMessage='Creates the ControlUp folder structure based on the EUC Environment tree'
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $MatchEUCEnvTree,

    [Parameter(
        Position=6,
        Mandatory=$true, 
        HelpMessage='A list of Brokers to connect and pull data'
    )]
    [ValidateNotNullOrEmpty()]
    [array] $Brokers,

    [Parameter(
        Position=7,
        Mandatory=$false, 
        HelpMessage='A list of Delivery Groups to include.  Works with wildcards'
    )]
    [ValidateNotNullOrEmpty()]
    [array] $includeDeliveryGroup,

    [Parameter(
        Position=8,
        Mandatory=$false, 
        HelpMessage='A list of Delivery Groups to exclude.  Works with wildcards. Exclusions supercede inclusions'
    )]
    [ValidateNotNullOrEmpty()]
    [array] $excludeDeliveryGroup,

    [Parameter(
        Position=9,
        Mandatory=$false, 
        HelpMessage='Adds the Citrix Brokers to the ControlUp tree'
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $addBrokersToControlUp ,

    [Parameter(
        Mandatory=$false, 
        HelpMessage='Only adds Delivery Groups which are enabled'
    )]
    [switch] $enabledOnly
) 

## GRL this way allows script to be run with debug/verbose without changing script
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

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
[string]$scriptPath = Split-Path -Path (& { $myInvocation.ScriptName }) -Parent
[string]$buildCuTreeScriptPath = [System.IO.Path]::Combine( $scriptPath , $buildCuTreeScript )

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
    Throw 'Failed to load Citrix PowerShell cmdlets - is this a Delivery Controller or have Studio or the PowerShell SDK installed ?'
}

Write-Host "Brokers: $Brokers"
$DeliveryGroups = New-Object System.Collections.Generic.List[PSObject]
$BrokerMachines = New-Object System.Collections.Generic.List[PSObject]
$CTXSites       = New-Object System.Collections.Generic.List[PSObject]

[hashtable]$brokerParameters = @{ 'AdminAddress' = $adminAddr }
if( $enabledOnly )
{
    $brokerParameters.Add( 'Enabled' , $true )
}

foreach ($adminAddr in $brokers) {
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

Write-Host "Total Number of Delivery Groups : $($DeliveryGroups.Count)"
Write-Host "Total Number of Brokers         : $($BrokerMachines.Count)"
Write-Host "Total Number of Sites           : $($($($CTXSites.Name | Sort-Object -Unique) | Measure-Object).count)"

#Add Included Delivery Groups
if ($PSBoundParameters.ContainsKey("includeDeliveryGroup")) {
    $DeliveryGroupsInclusionList = New-Object System.Collections.Generic.List[PSObject]
    foreach ($DeliveryGroupInclusion in $includeDeliveryGroup) {
        foreach ($DeliveryGroup in $DeliveryGroups) {
            if ($DeliveryGroup.Name -like $DeliveryGroupInclusion) {
            Write-Verbose -Message "Including DG $($DeliveryGroup.Name) - reason: $DeliveryGroupInclusion"
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
    $IndexesToRemove = New-Object System.Collections.Generic.List[PSObject]
    foreach ($DeliveryGroup in $DeliveryGroups) {
        foreach ($exclusion in $excludeDeliveryGroup) {
            if ($DeliveryGroup.Name -like $exclusion) {
                Write-Verbose "Excluding: $($DeliveryGroup.Name)  - because: $exclusion"
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
    $ControlUpEnvironmentObject.Add([ControlUpObject]::new($($DeliveryGroup.Name) ,"$($DeliveryGroup.Name)","Folder","","$($DeliveryGroup.site)-DeliveryGroup",""))
}

#Add machines from the delivery group to the environmental object
foreach ($DeliveryGroup in $DeliveryGroups) {
    
    $CTXMachines = Get-BrokerMachine -DesktopGroupName $DeliveryGroup.Name -AdminAddress $DeliveryGroup.Broker
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
            $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new($($Name) ,"$($DeliveryGroup.Name)","Computer","$Domain","$($DeliveryGroup.site)-Machine","$DNSName")) )
        }
    }
}

#Add Brokers to ControlUpEnvironmentalObject
if ($addBrokersToControlUp) {
    if (-not($MatchEUCEnvTree)) {  # MatchEUCEnvTree will add environment specific brokers folders
        $ControlUpEnvironmentObject.Add([ControlUpObject]::new("Brokers" ,"Brokers","Folder","","Brokers",""))
    }
    foreach ($Machine in $BrokerMachines) {
        if ([string]::IsNullOrEmpty($machine.DNSName)) {
            $DNSName = $null
        } else {
            $DNSName = $machine.DNSName
        }
        $Domain = $Machine.MachineName.split("\")[0]
        $Name =$Machine.MachineName.split("\")[1]
        $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new($($Name) ,"Brokers","Computer","$Domain","$($Machine.site)-BrokerMachine","$DNSName")))
    }
}
## TYE
if ($MatchEUCEnvTree) {
    for ($i=0; $i -lt $ControlUpEnvironmentObject.Count; $i++) {
        if ($ControlUpEnvironmentObject[$i].FolderPath -eq "Brokers" -and $ControlUpEnvironmentObject[$i].Type -eq "Computer") {
            $BrokerObj = $BrokerMachines | Where-Object ({$_.MachineName.split("\")[1] -eq $ControlUpEnvironmentObject[$i].Name})
            $ControlUpEnvironmentObject[$i].FolderPath = "$($BrokerObj.Site)\$($ControlUpEnvironmentObject[$i].FolderPath)" #Sets the path to $SiteName\Brokers
        } else {
            $ControlUpEnvironmentObject[$i].FolderPath = "$($ControlUpEnvironmentObject[$i].Description.Split("-")[0])\Delivery Groups\$($ControlUpEnvironmentObject[$i].FolderPath)"
        }
    }
    if ($addBrokersToControlUp) {
        foreach ($CtxSite in $($CTXSites | Sort-Object -Unique)) {
            $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new("Brokers" ,"$($CTXSite.Name)\Brokers","Folder","","Brokers","")))
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
    $BuildCUTreeParams.Add("SiteId",$Site)
}

[int]$errorCount = Build-CUTree -ExternalTree $ControlUpEnvironmentObject @BuildCUTreeParams
