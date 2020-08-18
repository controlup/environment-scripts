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

.PARAMETER Connection_servers
    A list of Connection Servers to contact for Horizon Pools,Farms and Computers to sync. Multiple Connection Servers can be specified if seperated by commas.

.PARAMETER includeDesktopPools
    Include only these specific Delivery Groups to be added to the ControlUp tree. Wild cards are supported as well, so if you have 
    Delivery Groups named like "Epic North", "Epic South" you can specify "Epic*" and it will capture both. Multiple delivery groups
    can be specified if they are seperated by commas. If you also use the parameter "excludeDeliveryGroups" then the exclude will
    supersede any includes and remove any matching Delivery Groups. Omitting this parameter and the script will capture all detected
    Delivery Groups.

.PARAMETER excludeDesktopPools
    Exclude specific delivery groups from being added to the ControlUp tree. Wild cards are supported as well, so if you have 
    Delivery Groups named like "Epic CGY", "Cerner CGY" you can specify "*CGY" and it will exclude any that ends with CGY. Multiple 
    delivery groups can be specified if they are seperated by commas. If you also use the parameter "includeDeliveryGroup" then the exclude will
    supersede any includes and remove any matching Delivery Groups.

.PARAMETER LocalPodOnly
    Configures the script to sync only the local Pod to ControlUp

.PARAMETER LocalSiteOnly
    Configures the script to sync only the local Site to ControlUp

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
    VMware Horizon

.MODIFICATION_HISTORY
    Wouter Kursten,         2020-08-11 - Original Code

.LINK

.COMPONENT

.NOTES
    Requires rights to read Citrix environment.

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
    [switch] $addBrokersToControlUp
) 

## For debugging uncomment
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'continue'
#$DebugPreference = 'SilentlyContinue'
Set-StrictMode -Version Latest


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
                Write-CULog -Message 'The required VMWare PowerShell components were not found as modules or snapins. Please make sure VMWare PowerCLI (version 6.5 or higher required) is installed and available for the user running the script.' -ShowConsole -Type E
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
        Write-CULog -Message 'The required module was not found. Please make sure COntrolUP.CLI Module is installed and available for the user running the script.' -ShowConsole -Type E
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
        Write-CULog -Message 'There was a problem connecting to the Horizon View Connection server.' -ShowConsole -Type E
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
        Write-CULog -Message 'There was a problem disconnecting from the Horizon View Connection server.' -ShowConsole -Type W
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
        Write-CULog -Message 'There was a problem retreiving the Horizon View Desktop Pool(s).' -ShowConsole -Type E
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
        Write-CULog -Message 'There was a problem retreiving the Horizon View RDS Farm(s).' -ShowConsole -Type E
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
        Write-CULog -Message 'There was a problem retreiving the Horizon View machines.' -ShowConsole -Type E
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
        Write-CULog -Message 'There was a problem retreiving the Horizon View machines.' -ShowConsole -Type E
    }
}

# Set the credentials location
[string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"

# Import the VMware PowerCLI modules
Load-VMwareModules -Components @('VimAutomation.HorizonView')

# Get the stored credentials for running the script
[PSCredential]$CredsHorizon = Get-CUStoredCredential -System 'HorizonView'

# Connect to the Horizon View Connection Server
[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $HVConnectionServerFQDN -Credential $CredsHorizon

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
    if ($localsiteonly -eq $true){
        $hvlocalpod=$hvpods | where-object {$_.LocalPod -eq $true}
        $hvlocalsite=$objHVConnectionServer.ExtensionData.Site.Site_Get($hvlocalpod.site)
        foreach ($hvpod in $hvlocalsite.pods){$HVPodendpoints+=$objHVConnectionServer.ExtensionData.PodEndpoint.PodEndpoint_list($hvpod) | select-object -first 1}
        }

    else {
            [array]$HVPodendpoints = foreach ($hvpod in $hvpods) {$objHVConnectionServer.ExtensionData.PodEndpoint.PodEndpoint_List($hvpod.id) | select-object -first 1}
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
if ($exceptionfile){
    [array]$exceptionlist=get-content $exceptionfile
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

        Write-CULog -message "Processing Pod $podname"

        # Add folder with the podname to the batch
        $ControlUpEnvironmentObject.Add([ControlUpObject]::new("$podname" ,"$podname","Folder","","",""))

        [string]$targetfolderpath=$podname
    }
    else{
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $hvconnectionserver -Credential $CredsHorizon
        Write-CULog -message "Processing a non Cloud Pod Architecture Environment"
        [string]$targetfolderpath=""
    }
        # Get the Horizon View desktop Pools
        [array]$HVPools = Get-HVDesktopPools -HVConnectionServer $objHVConnectionServer
        [array]$HVPools = $HVPools | Where-Object {$exceptionlist -notcontains $_.DesktopSummaryData.Name}
        Write-CULog -message "Processing Dividers"
        # Add folder for the Desktop pool divider to the batch
        add-CUFolder -ParentPath $strpodnamepath -Name $Pooldivider -batch $batch
        $ControlUpEnvironmentObject.Add([ControlUpObject]::new("$Desktop Pools" ,"$podname","Folder","","$($podname)-Pod",""))

        # Add folder for the RDS divider to the batch

        add-CUFolder -ParentPath $strpodnamepath -name $RDSDivider -Batch $batch

        # Apply the batch for the dividers

        [string]$Poolspath=$strpodnamepath+"\"+$Pooldivider

        # first the folders for the pools need to be created
        foreach ($hvpool in $hvpools){
            # Create the variable for the batch of machines that will be used to add and remove machines
            $poolname=($hvpool).DesktopSummaryData.Name
            Write-CULog -message "Processing Desktop Pool $poolname"
                Write-CULog -message "Adding folder for Desktop Pool $poolname"
                # Defines the batch for folders
                # Adds folder for the Desktop pool to the batch
                add-CUFolder -ParentPath $Poolspath -Name $poolname -batch $batch
                # count the amount of pools first


                Write-Host "Applying batch update every $folderbatchsize folders"


        }


        foreach ($hvpool in $hvpools){
            # Create the variable for the batch of machines that will be used to add and remove machines
            $poolname=($hvpool).DesktopSummaryData.Name
            Write-CULog -message "Processing Desktop Pool $poolname"
            [string]$poolnamepath=$Poolspath+"\"+$poolname
            [array]$CUComputers=@{}
            $CUComputers=Get-CUComputers -FolderPath $poolnamepath
            # Retreive all the desktops in the desktop pool.
            [array]$HVDesktopmachines=@{}
            $HVDesktopmachines = Get-HVDesktopMachines -HVPoolID $HVPool.id -HVConnectionServer $objHVConnectionServer
            # Filtering out any desktops without a DNS name
            $HVDesktopmachines = $HVDesktopmachines | where-object {$_.base.dnsname -ne $null}
            # Remove machines in the exceptionlist
            [array]$HVDesktopmachines = $HVDesktopmachines | Where-Object {$exceptionlist -notcontains $_.base.dnsname}
            # Create list with desktops that need to be added
            [array]$toAddmachines = $HVDesktopmachines | Where-Object {($CUComputers).FQDN -notcontains $_.base.dnsname}
            # Create List with desktops that need to be removed

                Write-CULog -message "Adding VDI Machines actions to the batch"

                foreach ($toaddmachine in $toaddmachines){
                    # Try to convert to lowercase
                    try{
                        $toaddname=$toaddmachine.base.name.ToLower()
                    }
                    catch {
                        Write-CULog -message "error converting Machine names to lowercase" -ShowConsole -Type E
                    }
                    # If this is a manual non-managed pool the name equals the dnsname so we extract the name in another way.
                    if ($toaddname -eq $toaddmachine.base.dnsname){
                        $toaddname=$toaddname.split('.')[0]
                    }
                    # Try getting the domain name from the dnsname
                    try{
                        $domainname=$toaddmachine.base.dnsname.replace($toaddname+".","")
                    }
                    catch {
                        Write-CULog -Message "Error retreiving the domainname from the DNS name" -ShowConsole -Type E
                    }
                    add-cucomputer -name $toaddname -domain $domainname -folderpath $poolnamepath -Batch $batch

            }
        }

        [array]$HVFarms = Get-HVfarms -HVConnectionServer $objHVConnectionServer
        [array]$HVFarms = $HVFarms | Where-Object {$exceptionlist -notcontains $_.Data.Name}
        [string]$Farmspath=$strpodnamepath+"\"+$RDSDivider

        # first the folders for the farms need to be created
        foreach ($HVFarm in $HVFarms){
            # Create the variable for the batch of machines that will be used to add and remove machines
            $farmname=($hvfarm).Data.Name

                Write-CULog -message "Adding folder for Desktop Pool $farmname"
                # Defines the batch for folders
                # Adds folder for the Desktop pool to the batch
                add-CUFolder -ParentPath $Farmspath -Name $farmname -batch $batch
                # count the amount of pools first
        }
        foreach ($HVFarm in $HVFarms){
            # Create the variable for the batch of machines that will be used to add and remove machines
            $farmname=($hvfarm).Data.Name
            Write-CULog -message "Processing RDS Farm $farmname"
            [string]$farmnamepath=$Farmspath+"\"+$farmname
            [array]$CUComputers=@{}
            $CUComputers=Get-CUComputers -FolderPath $farmnamepath
            # Retreive all the desktops in the desktop pool.
            [array]$HVfarmmachines=@{}
            $HVfarmmachines = Get-HVRDSMachines -HVFarmID $HVfarm.id -HVConnectionServer $objHVConnectionServer
            # Filtering out any RDS Machines without a DNS name
            $HVfarmmachines = $HVfarmmachines | where-object {$_.AgentData.dnsname -ne $null}
            # Remove machines in the exceptionlist
            [array]$HVfarmmachines = $HVfarmmachines | Where-Object {$exceptionlist -notcontains $_.AgentData.dnsname}
            # Create list with desktops that need to be added
            [array]$toAddmachines = $HVfarmmachines | Where-Object {($CUComputers).FQDN -notcontains $_.AgentData.dnsname}
            # Create List with desktops that need to be removed

            if ($todeletemachines -or $toaddmachines){
                Write-CULog -message "Adding RDS host actions to the batch"
                $batch = New-CUBatchUpdate
                $count=0
                foreach ($toaddmachine in $toaddmachines){
                    # Try to convert to lowercase
                    try {
                        $toaddname=$toaddmachine.base.name.ToLower()
                    }
                    catch {
                        Write-CULog -message "error converting Machine names to lowercase" -ShowConsole -Type E
                    }
                    # If this is a manual non-managed pool the name equals the dnsname so we extract the name in another way.
                    if ($toaddname -eq $toaddmachine.AgentData.dnsname){
                        $toaddname=$toaddname.split('.')[0]
                    }
                    # Try getting the domain name from the dnsname
                    try {
                        $domainname=$toaddmachine.AgentData.dnsname.replace($toaddname+".","")
                    }
                    catch {
                        Write-CULog -Message "Error retreiving the domainname from the DNS name" -ShowConsole -Type E
                    }
                    add-cucomputer -name $toaddname -domain $domainname -folderpath $farmnamepath -Batch $batch



            }
        }
    }

    }
    Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
}

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
