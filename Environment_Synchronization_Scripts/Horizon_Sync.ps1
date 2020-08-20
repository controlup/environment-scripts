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

.PARAMETER ConnectionServer
    FQDN for one of the connection servers.

.PARAMETER LocalHVSiteOnly
    Configures the script to sync only the local Site to ControlUp

.EXAMPLE
    . .\CU_SyncScript.ps1 -ConnectionServer "connectionserver.domain.example" -folderPath "CUSync\Horizon" 
        Contacts the connection server connectionserver.domain.example, discover and connect to all other pods in the cloud pod federation. Tt will save the objects to the ControlUp folder 
        "CUSync\Horizon". I

.EXAMPLE
    \CU_SyncScript.ps1 -ConnectionServer "connectionserver.domain.example" -folderPath "CUSync\Horizon" -localHVsite only
        Contacts the connection server connectionserver.domain.example, discover and connect to all other pods in the local Site in the cloud pod federation. It will save the objects to the ControlUp folder 
        "CUSync\Horizon".

.CONTEXT
    VMware Horizon

.MODIFICATION_HISTORY
    Wouter Kursten,         2020-08-18 - Original Code

.LINK

.COMPONENT

.NOTES
    Requires rights to connect to the Horizon API's and read the Horizon environment.

    Version:        0.1
    Author:         Wouter Kursten
    Creation Date:  2020-08-06
    Updated:        2020-08-06
                    Changed ...
    Purpose:        Created for VMware Horizon Sync
#>

[CmdletBinding()]
Param
(
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        HelpMessage='Enter a ControlUp subfolder to save your Horizon tree'
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
        HelpMessage='FQDN of the connectionserver'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $ConnectionServer,

    [Parameter(
        Position=3, 
        Mandatory=$false, 
        HelpMessage='Execute removal operations. When combined with preview it will only display the proposed changes'
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $Delete,

    [Parameter(
        Position=4, 
        Mandatory=$false, 
        HelpMessage='Enter a path to generate a log file of the proposed changes'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $LogFile,

    [Parameter(
        Position=5, 
        Mandatory=$false, 
        HelpMessage='Synchronise the local site only'
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $LocalHVSiteOnly,

   [Parameter(
        Position=6, 
        Mandatory=$false, 
        HelpMessage='Enter a ControlUp Site'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $Site,

    [Parameter(
        Position=7, 
        Mandatory=$false, 
        HelpMessage='File with a list of exceptions, machine names and/or desktop pools'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $Exceptionsfile


) 

## For debugging uncomment
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'continue'
#$DebugPreference = 'SilentlyContinue'
Set-StrictMode -Version Latest

$Pooldivider="Desktop Pools"
$RDSDivider="RDS Farms"

# dot sourcing Functions
. ".\Build-CUTree.ps1"

function Make-NameWithSafeCharacters ([string]$string) {
    ###### TODO need to replace the folder path characters that might be illegal
    #list of illegal characters : '/', '\', ':', '*','?','"','<','>','|','{','}'
    $returnString = (($string).Replace("/","-")).Replace("\","-").Replace(":","-").Replace("*","-").Replace("?","-").Replace("`"","-").Replace("<","-").Replace(">","-").Replace("|","-").Replace("{","-").Replace("}","-")
    return $returnString
}
    
function Get-CUStoredCredential {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The system the credentials will be used for.")]
        [string]$System
    )
    # Get the stored credential object
    [string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
    Import-Clixml $strCUCredFolder\$($env:USERNAME)_$($System)_Cred.xml
}

function Load-VMWareModules {
    <# Imports VMware PowerCLI modules
    NOTES:
    - The required modules to be loaded are passed as an array.
    - If the PowerCLI versions is below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
    #>

    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware module to be loaded. Can be single or multiple values (as array).")]
        [array]$Components
    )

    # Try Import-Module for each passed component, try Add-PSSnapin if this fails (only if -Prefix was not specified)
    # Import each module, if Import-Module fails try Add-PSSnapin
    foreach ($component in $Components) {
        try {
            $null = Import-Module -Name VMware.$component
        }
        catch {
            try {
                $null = Add-PSSnapin -Name VMware
            }
            catch {
                Write-CULog -Msg 'The required VMWare PowerShell components were not found as modules or snapins. Please make sure VMWare PowerCLI (version 6.5 or higher required) is installed and available for the user running the script.' -ShowConsole -Type E
            }
        }
    }
}

function Load-ControlUPModule {
    # Try Import-Module for each passed component.
    try {
        $pathtomodule = (Get-ChildItem "C:\Program Files\Smart-X\ControlUpMonitor\*ControlUp.PowerShell.User.dll" -Recurse | Sort-Object LastWriteTime -Descending)[0]
        Import-Module $pathtomodule
    }
    catch {
        Write-CULog -Msg 'The required module was not found. Please make sure COntrolUP.CLI Module is installed and available for the user running the script.' -ShowConsole -Type E
        }
}

function Connect-HorizonConnectionServer {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The FQDN of the Horizon View Connection server. IP address may be used.")]
        [string]$HVConnectionServerFQDN,
        [parameter(Mandatory = $true,
            HelpMessage = "The PSCredential object used for authentication.")]
        [PSCredential]$Credential
    )
    # Try to connect to the Connection server
    try {
        Connect-HVServer -Server $HVConnectionServerFQDN -Credential $Credential
    }
    catch {
        Write-CULog -Msg 'There was a problem connecting to the Horizon View Connection server.' -ShowConsole -Type E
    }
}

function Disconnect-HorizonConnectionServer {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    # Try to connect from the connection server
    try {
        Disconnect-HVServer -Server $HVConnectionServer -Confirm:$false
    }
    catch {
        Write-CULog -Msg 'There was a problem disconnecting from the Horizon View Connection server.' -ShowConsole -Type W
    }
}

function Get-HVDesktopPools {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    # Try to get the Desktop pools in this pod
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'DesktopSummaryView'
        # Filter oud rds desktop pools since they don't contain machines
        $defn.Filter = New-Object VMware.Hv.QueryFilterNotEquals -property @{'memberName'='desktopSummaryData.type'; 'value' = "RDS"}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Write-CULog -Msg 'There was a problem retreiving the Horizon View Desktop Pool(s).' -ShowConsole -Type E
    }
}

function Get-HVFarms {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'FarmSummaryView'
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Write-CULog -Msg 'There was a problem retreiving the Horizon View RDS Farm(s).' -ShowConsole -Type E
    }
}

function Get-HVDesktopMachines {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Horizon View Desktop Pool.")]
        [VMware.Hv.DesktopId]$HVPoolID,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'MachineSummaryView'
        # Filter for only the machines within the provided desktop pool
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='base.desktop'; 'value' = $HVPoolID}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Write-CULog -Msg 'There was a problem retreiving the Horizon View machines.' -ShowConsole -Type E
    }
}

function Get-HVRDSMachines {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "ID of the Horizon View RDS Farm.")]
        [VMware.Hv.FarmId]$HVFarmID,
        [parameter(Mandatory = $true,
            HelpMessage = "The Horizon View Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'RDSServerSummaryView'
        # Filter for only the machines within the provided desktop pool
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='base.farm'; 'value' = $HVFarmID}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        return $queryResults
    }
    catch {
        Write-CULog -Msg 'There was a problem retreiving the Horizon View machines.' -ShowConsole -Type E
    }
}

function Write-CULog {
    <#
    .SYNOPSIS
        Write the Logfile
    .DESCRIPTION
        Helper Function to Write Log Messages to Console Output and corresponding Logfile
        use get-help <functionname> -full to see full help
    .EXAMPLE
        Write-CULog -Msg "Warining Text" -Type W
    .EXAMPLE
        Write-CULog -Msg "Text would be shown on Console" -ShowConsole
    .EXAMPLE
        Write-CULog -Msg "Text would be shown on Console in Cyan Color, information status" -ShowConsole -Color Cyan
    .EXAMPLE
        Write-CULog -Msg "Error text, script would be existing automaticaly after this message" -Type E
    .EXAMPLE
        Write-CULog -Msg "External log contenct" -Type L
    .NOTES
        Author: Matthias Schlimm
        Company:  EUCWeb.com
        History:
        dd.mm.yyyy MS: function created
        07.09.2015 MS: add .SYNOPSIS to this function
        29.09.2015 MS: add switch -SubMSg to define PreMsg string on each console line
        21.11.2017 MS: if Error appears, exit script with Exit 1
        08.07.2020 TT: Borrowed Write-BISFLog and modified to meet the purpose for this script
    .LINK
        https://eucweb.com
    #>

    Param(
        [Parameter(Mandatory = $True)][Alias('M')][String]$Msg,
        [Parameter(Mandatory = $False)][Alias('S')][switch]$ShowConsole,
        [Parameter(Mandatory = $False)][Alias('C')][String]$Color = "",
        [Parameter(Mandatory = $False)][Alias('T')][String]$Type = "",
        [Parameter(Mandatory = $False)][Alias('B')][switch]$SubMsg
    )

    $LogType = "INFORMATION..."
    IF ($Type -eq "W" ) { $LogType = "WARNING........."; $Color = "Yellow" }
    IF ($Type -eq "E" ) { $LogType = "ERROR..............."; $Color = "Red" }

    IF (!($SubMsg)) {
        $PreMsg = "+"
    }
    ELSE {
        $PreMsg = "`t>"
    }

    $date = Get-Date -Format G
    if ($LogFile) {
        Write-Output "$date | $LogType | $Msg"  | Out-file $($LogFile) -Append
    }
    

    If (!($ShowConsole)) {
        IF (($Type -eq "W") -or ($Type -eq "E" )) {
            #IF ($VerbosePreference -eq 'SilentlyContinue') {
                Write-Host "$PreMsg $Msg" -ForegroundColor $Color
                $Color = $null
            #}
        }
        ELSE {
            Write-Verbose -Msg "$PreMsg $Msg"
            $Color = $null
        }

    }
    ELSE {
        if ($Color -ne "") {
            #IF ($VerbosePreference -eq 'SilentlyContinue') {
                Write-Host "$PreMsg $Msg" -ForegroundColor $Color
                $Color = $null
            #}
        }
        else {
            Write-Host "$PreMsg $Msg"
        }
    }
}

# Set the credentials location
[string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.HorizonView')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server
[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $ConnectionServer -Credential $CredsHorizon

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

# checks if this connectionserver is member of a cloud pod federation

[VMware.Hv.PodFederationLocalPodStatus]$HVpodstatus=($objHVConnectionServer.ExtensionData.PodFederation.PodFederation_Get()).localpodstatus
if ($HVpodstatus.status -eq "ENABLED"){
    # Retreives all pods
    [array]$HVpods=$objHVConnectionServer.ExtensionData.Pod.Pod_List()
    # retreive the first connection server from each pod
    $HVPodendpoints=@()
    if ($LocalHVSiteOnly){
        $hvlocalpod=$hvpods | where-object {$_.LocalPod -eq $true}
        $hvlocalsite=$objHVConnectionServer.ExtensionData.Site.Site_Get($hvlocalpod.site)
        foreach ($hvpod in $hvlocalsite.pods){$HVPodendpoints+=$objHVConnectionServer.ExtensionData.PodEndpoint.PodEndpoint_list($hvpod) | select-object -first 1}
        }

    else {
            [array]$HVPodendpoints += foreach ($hvpod in $hvpods) {$objHVConnectionServer.ExtensionData.PodEndpoint.PodEndpoint_List($hvpod.id) | select-object -first 1}
    }
    # Convert from url to only the name
    [array]$hvconnectionservers=$HVPodendpoints.serveraddress.replace("https://","").replace(":8472/","")
    # Disconnect from the current connection server
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
}
else {
    # Create list with one entry
    $hvconnectionservers=$hvConnectionServerfqdn
    # Disconnect from the current connection server
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
}
# Get the content of the exception file and put it into an array
if ($Exceptionsfile){
    [array]$exceptionlist=get-content $Exceptionsfile
}
else {
    $exceptionlist=@()
}

$ControlUpEnvironmentObject = New-Object System.Collections.Generic.List[PSObject]

foreach ($hvconnectionserver in $hvconnectionservers){
    if ($HVpodstatus.status -eq "ENABLED"){
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $hvconnectionserver -Credential $CredsHorizon
        # Retreive the name of the pod

        $pods=$objHVConnectionServer.extensionData.pod.Pod_list()
        [string]$podname=$pods | where-object {$_.localpod -eq $True} | select-object -expandproperty Displayname

        Write-CULog -Msg "Processing Pod $podname"

        # Add folder with the podname to the batch
        [string]$targetfolderpath="$podname"
        $ControlUpEnvironmentObject.Add([ControlUpObject]::new("$podname" ,"$podname","Folder","","",""))

    }
    else{
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $hvconnectionserver -Credential $CredsHorizon
        Write-CULog -Msg "Processing a non Cloud Pod Architecture Environment"
        [string]$targetfolderpath=""
    }
    # Get the Horizon View desktop Pools
    [array]$HVPools = Get-HVDesktopPools -HVConnectionServer $objHVConnectionServer
    

    if($NULL -ne $hvpools){
        [array]$HVPools = $HVPools | Where-Object {$exceptionlist -notcontains $_.DesktopSummaryData.Name}
        if($targetfolderpath -eq ""){
            [string]$Poolspath=$Pooldivider
        }
        else{
            [string]$Poolspath=$targetfolderpath+"\"+$Pooldivider
        }
        if($ControlUpEnvironmentObject.name -notcontains $Pooldivider){
            $ControlUpEnvironmentObject.Add([ControlUpObject]::new("$Pooldivider" ,"$Poolspath","Folder","","$($podname)-Pod",""))
        }

        # first the folders for the pools need to be created
        foreach ($hvpool in $hvpools){
            # Create the variable for the batch of machines that will be used to add and remove machines
            $poolname=($hvpool).DesktopSummaryData.Name
            Write-CULog -Msg "Processing Desktop Pool $poolname"
            [string]$poolnamepath=$Poolspath+"\"+$poolname
            $ControlUpEnvironmentObject.Add([ControlUpObject]::new("$poolname" ,"$poolnamepath","Folder","","$($poolname)-Pool",""))

            # Retreive all the desktops in the desktop pool.
            [array]$HVDesktopmachines=@{}
            $HVDesktopmachines = Get-HVDesktopMachines -HVPoolID $HVPool.id -HVConnectionServer $objHVConnectionServer
            if($NULL -ne $HVDesktopmachines) {
            # Filtering out any desktops without a DNS name
            $HVDesktopmachines = $HVDesktopmachines | where-object {$_.base.dnsname -ne $null}
            # Remove machines in the exceptionlist
            [array]$HVDesktopmachines = $HVDesktopmachines | Where-Object {$exceptionlist -notcontains $_.base.dnsname}
                foreach ($HVDesktopmachine in $HVDesktopmachines){
                    $dnsname=$HVDesktopmachine.base.dnsname
                    # Try to convert to lowercase
                    try{
                        $toaddname=$HVDesktopmachine.base.name.ToLower()
                    }
                    catch {
                        Write-CULog -Msg "error converting Machine names to lowercase" -ShowConsole -Type E
                    }
                    # If this is a manual non-managed pool the name equals the dnsname so we extract the name in another way.
                    if ($toaddname -eq $HVDesktopmachine.base.dnsname){
                        $toaddname=$toaddname.split('.')[0]
                    }
                    # Try getting the domain name from the dnsname

                    try{
                        $domainname=$HVDesktopmachine.base.dnsname.replace($toaddname+".","")
                    }
                    catch {
                        Write-CULog -Msg "Error retreiving the domainname from the DNS name" -ShowConsole -Type E
                    }
                    #add-cucomputer -name $toaddname -domain $domainname -folderpath $poolnamepath -Batch $batch
                    $ControlUpEnvironmentObject.Add([ControlUpObject]::new("$toaddname" ,"$poolnamepath","Computer","$domainname","$($poolname)-Pool","$dnsname"))
                }
            }
        }
    }
    write-culog -Msg "Processing RDS Farms"
    [array]$HVFarms = Get-HVfarms -HVConnectionServer $objHVConnectionServer
    
    if ($NULL -ne $hvfarms){
        [array]$HVFarms = $HVFarms | Where-Object {$exceptionlist -notcontains $_.Data.Name}
        if($targetfolderpath -eq ""){
            [string]$Farmspath=$RDSDivider
        }
        else{
            [string]$Farmspath=$targetfolderpath+"\"+$RDSDivider
        }
        if($ControlUpEnvironmentObject.name -notcontains $RDSDivider){
            $ControlUpEnvironmentObject.Add([ControlUpObject]::new("$RDSDivider" ,"$Farmspath","Folder","","$($podname)-Pod",""))
        }

        foreach ($HVFarm in $HVFarms){
            # Create the variable for the batch of machines that will be used to add and remove machines
            $farmname=($hvfarm).Data.Name
            [string]$farmnamepath=$Farmspath+"\"+$farmname
            $ControlUpEnvironmentObject.Add([ControlUpObject]::new("$farmname" ,"$farmnamepath","Folder","","$($farmname)-Pool",""))
            $farmname=($hvfarm).Data.Name
            Write-CULog -Msg "Processing RDS Farm $farmname"
            # Retreive all the desktops in the desktop pool.
            [array]$HVfarmmachines=@{}
            $HVfarmmachines = Get-HVRDSMachines -HVFarmID $HVfarm.id -HVConnectionServer $objHVConnectionServer
            # Filtering out any RDS Machines without a DNS name
            if($HVfarmmachines -ne ""){
                $HVfarmmachines = $HVfarmmachines | where-object {$_.AgentData.dnsname -ne $null}
                # Remove machines in the exceptionlist
                [array]$HVfarmmachines = $HVfarmmachines | Where-Object {$exceptionlist -notcontains $_.AgentData.dnsname}
                # Create List with desktops that need to be removed
                foreach ($HVfarmmachine in $HVfarmmachines){
                    $dnsname=$HVfarmmachine.AgentData.dnsname
                    # Try to convert to lowercase
                    try {
                        $toaddname=$HVfarmmachine.base.name.ToLower()
                    }
                    catch {
                        Write-CULog -Msg "error converting Machine names to lowercase" -ShowConsole -Type E
                    }
                    # If this is a manual non-managed pool the name equals the dnsname so we extract the name in another way.
                    if ($toaddname -eq $HVfarmmachine.AgentData.dnsname){
                        $toaddname=$toaddname.split('.')[0]
                    }
                    # Try getting the domain name from the dnsname
                    try {
                        $domainname=$HVfarmmachine.AgentData.dnsname.replace($toaddname+".","")
                    }
                    catch {
                        Write-CULog -Msg "Error retreiving the domainname from the DNS name" -ShowConsole -Type E
                    }
                    $ControlUpEnvironmentObject.Add([ControlUpObject]::new("$toaddname" ,"$farmnamepath","Computer","$domainname","$($farmname)-Farm","$dnsname"))
                }
            }
        }
    }
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
}
$ControlUpEnvironmentObject | fl

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

Build-CUTree -ExternalTree $ControlUpEnvironmentObject @BuildCUTreeParams
