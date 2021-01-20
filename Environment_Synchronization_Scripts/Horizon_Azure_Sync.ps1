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
    Guy Leech,              2020-09-23 - Added more logging for fatal errors to aid troubleshooting when run as ascheduled task
    Guy Leech,              2020-10-13 - Accommodate Build-CUTree returning error count
    Guy Leech,              2020-11-13 - Added getting RDS servers
    Guy Leech,              2020-11-26 - Changed folder structure created in CU
    Guy Leech,              2020-12-16 - Added findpool code
    Guy Leech,              2020-12-24 - Verbose output for findpools result for troubleshooting missing machines
    Guy Leech,              2021-01-06 - Use findpools output to find any pools not already retrieved
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
        Mandatory=$true,
        HelpMessage='Enter a ControlUp subfolder to save your Horizon tree'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $folderPath,

    [Parameter(
        HelpMessage='Preview the changes'
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $Preview,

    [Parameter(
        HelpMessage='Base URL used to logon to VMware Cloud'
    )]
    [ValidateNotNullOrEmpty()]
    [string]$base = 'cloud-us-2' , ## get this from the URL shown after manual logon to VMware cloud

    [Parameter(
        HelpMessage='Execute removal operations. When combined with preview it will only display the proposed changes'
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $Delete,

    [Parameter(
        HelpMessage='Enter a path to generate a log file of the proposed changes'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $LogFile,

    [Parameter(
        HelpMessage='Synchronise the local site only'
    )]
    [ValidateNotNullOrEmpty()]
    [switch] $LocalHVSiteOnly,

   [Parameter(
        HelpMessage='Enter a ControlUp Site'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $Site,

    [Parameter(
        HelpMessage='File with a list of exceptions, machine names and/or desktop pools'
    )]
    [ValidateNotNullOrEmpty()]
    [string] $Exceptionsfile
)

## GRL this way allows script to be run with debug/verbose without changing script
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[string]$Pooldivider="VDI Desktops"
[string]$RDSDivider="RDS Farms"
[string]$buildCuTreeScript = 'Build-CUTree.ps1' ## Script from ControlUp which must reside in the same folder as this script

# dot sourcing Functions
## GRL Don't assume user has changed location so get the script path instead
[string]$scriptPath = Split-Path -Path (& { $myInvocation.ScriptName }) -Parent
[string]$buildCuTreeScriptPath = [System.IO.Path]::Combine( $scriptPath , $buildCuTreeScript )

if( ! ( Test-Path -Path $buildCuTreeScriptPath -PathType Leaf -ErrorAction SilentlyContinue ) )
{
    [string]$errorMessage = "Unable to find script `"$buildCuTreeScript`" in `"$scriptPath`""
    ## Write-CULog -Msg $errorMessage -ShowConsole -Type E ## can't write to CU-Log yet as that function is in the scripot we can't find
    Throw $errorMessage
}

. $buildCuTreeScriptPath

function Get-CUStoredCredential {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The system the credentials will be used for.")]
        [string]$System
    )
    # Get the stored credential object
    [string]$strCUCredFolder = "$([environment]::GetFolderPath('CommonApplicationData'))\ControlUp\ScriptSupport"
    [string]$strCUCredFile = Join-Path -Path "$strCUCredFolder" -ChildPath "$($env:USERNAME)_$($System)_Cred.xml"
    if( ! ( Test-Path -Path $strCUCredFile -ErrorAction SilentlyContinue ) )
    {
        [string]$errorMessage = "Unable to find stored credential file `"$strCUCredFile`" - have you previously run the `"Create Credentials for Horizon Scripts`" script for user $env:username ?"
        Write-CULog -Msg $errorMessage -ShowConsole -Type E
        Throw $errorMessage
    }
    else
    {
        Try
        {
            Import-Clixml -Path $strCUCredFile
        }
        Catch
        {
            [string]$errorMessage = "Error reading stored credentials from `"$strCUCredFile`" : $_"
            Write-CULog -Msg $errorMessage -ShowConsole -Type E
            Throw $errorMessage
        }
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
            Write-Verbose -Message "$PreMsg $Msg"
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
[string]$strCUCredFolder = [System.IO.Path]::Combine( [environment]::GetFolderPath( [Environment+SpecialFolder]::CommonApplicationData ) , 'ControlUp' , 'ScriptSupport' )

# Get the stored credentials for running the script - two sets - myVMware and domain account

[System.Management.Automation.PSCredential]$myVMwareCredential = Get-CUStoredCredential -System 'HorizonCloudmyVMware'

if( ! $myVMwareCredential )
{
    [string]$errorMessage = "Failed to get stored credentials for $env:username for myVMware"
    Write-CULog -Msg $errorMessage -ShowConsole -Type E
    Throw $errorMessage
}

[System.Management.Automation.PSCredential]$domainCredential = Get-CUStoredCredential -System 'HorizonCloudDomain'

if( ! $domainCredential )
{
    [string]$errorMessage = "Failed to get stored credentials for $env:username for domain account"
    Write-CULog -Msg $errorMessage -ShowConsole -Type E
    Throw $errorMessage
}

Write-Verbose -Message "myVMware user is $($myVMwareCredential.Username), domain user is $($domainCredential.Username)"

##https://stackoverflow.com/questions/41897114/unexpected-error-occurred-running-a-simple-unauthorized-rest-query?rq=1
Add-Type -TypeDefinition @'
public class SSLHandler
{
    public static System.Net.Security.RemoteCertificateValidationCallback GetSSLHandler()
    {
        return new System.Net.Security.RemoteCertificateValidationCallback((sender, certificate, chain, policyErrors) => { return true; });
    }
}
'@

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = [SSLHandler]::GetSSLHandler()
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

$sessionVariable = $null

[string]$baseURL = "https://$base.horizon.vmware.com"

[hashtable]$RESTparams = @{
    'ContentType' = 'application/json'
    'Method' = 'Post'
}

#region authentication

## 1. logon with myVmware account

[hashtable]$authParams = @{
        username = $myVMwareCredential.UserName
        password = $myVMwareCredential.GetNetworkCredential().Password.ToCharArray()
}

$RESTparams.Body = ( $authParams | ConvertTo-Json ).ToString()
$RESTparams.Uri = $baseURL + '/api/login/login'

try
{
    $response1 = Invoke-RestMethod @RESTparams -SessionVariable sessionVariable
}
catch
{
    $response1 = $null
    Throw "Failed to logon to $($RESTparams.uri) as $($myVMwareCredential.UserName) : $_"
}

if( ! $response1 -or ! $response1.PSObject.Properties[ 'authSession' ] )
{
    Throw "Failed to get authSession from $($RESTparams.uri) as $($myVMwareCredential.UserName), got $response1"
}

$RESTparams.websession = $sessionVariable

## 2. logon to AD

[string]$domain = ($domainCredential.UserName -split '\\')[0]

if( $null -eq $response1.PSObject.Properties[ 'credentialRequested' ] )
{
    Write-Warning -Message "No credentialRequested returned from initial logon"
}

if( $domain -notin $response1.domainNames )
{
    Write-Warning -Message "Domain name `"$domain`" in AD credential not in list of domains returned from Horizon - $($response1.domainNames -join ',')"
}

$authParams.username = ($domainCredential.UserName -split '\\')[-1]
$authParams.domain   = $domain
$authParams.password = $domainCredential.GetNetworkCredential().Password.ToCharArray()
$authParams.credentialRequested = $response1.credentialRequested
$authParams.authSession = $response1.authSession

$RESTparams.Body = ( $authParams | ConvertTo-Json ).ToString()

try
{
    $authentication = Invoke-RestMethod @RESTparams
}
catch
{
    $authentication = $null
    Throw "Failed AD logon to $($RESTparams.uri) as $($domainCredential.UserName) : $_"
}

if( ! $authentication -or $null -eq $authentication.PSObject.Properties[ 'apiToken' ] )
{
    Throw "No apiToken returned from AD logon"
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

# Get the content of the exception file and put it into an array
if ($Exceptionsfile){
    [array]$exceptionlist= get-content -Path $Exceptionsfile
}
else {
    $exceptionlist=@()
}

## Map VMware terms to CU
[hashtable]$poolTypeToFolder = @{
    'desktop' = $Pooldivider
    'session' = $RDSDivider
}

$ControlUpEnvironmentObject = New-Object System.Collections.Generic.List[PSObject]

$RESTparams.Headers = @{ 'Authorization' = "Bearer $($authentication.apiToken)" ; 'Accept' = 'application/json' ; 'Content-Type' = 'application/json' }

## 6. Get broker type

$RESTparams.Remove( 'Body' )
$RESTparams.Method = 'GET'
$RESTparams.uri = $baseURL + '/api/cloudbrokersync/brokerPreference'

try
{
    $broker = Invoke-RestMethod @RESTparams
}
catch
{
    $broker = $null
    Throw "Failed to get broker via $($RESTparams.uri) : $_"
}

if( $broker )
{
    ## If azureBrokeringType is ‘BROKER_NEXT’, it means customer is using Universal Broker.
    ## If value for azureBrokeringType is “LEGACY_BROKER”, it means customer is using single pod broker.
    Write-Verbose -Message "Broker Type is $($broker.azureBrokeringType)"
}

## 7. Fetch all multi pod VDI assignments for customer

$RESTparams.Uri = $baseURL + '/api/mcw/assignments?type=WORKSPACE'

try
{
    $assignments = Invoke-RestMethod @RESTparams
}
catch
{
    $assignments = $null
    Throw "Failed to get all multi pod VDI assignments via $($RESTparams.uri) : $_"
}

## 8. Get farms
$RESTparams.Uri = $baseURL + '/dt-rest/v100/farm/manager/farms/session'

try
{
    $farms = Invoke-RestMethod @RESTparams
}
catch
{
    $farms = $null
    Write-Warning -Message "Failed to get farms via $($RESTparams.uri) : $_"
}

[hashtable]$parentFolders = @{}
[hashtable]$poolsProcessed = @{}

if( ! $farms -or ! $farms.Count )
{
    Write-Warning -Message "Call to $($RESTparams.uri) returned no farms"
}
else
{
    [string]$farmsURL = $RESTparams.Uri

    ForEach( $farm in $farms )
    {
        $RESTparams.Uri = $farmsURL + ( '/{0}/allservers' -f $farm.id )

        try
        {
            $VMs = Invoke-RestMethod @RESTparams
        }
        catch
        {
            $VMs = $null
            Write-Warning -Message "Failed to get VMs via $($RESTparams.uri) : $_"
        }

        if( $VMs )
        {
            $vms | Select-Object -Property name , guestOs , numcpus , memorySizeMB , ipaddress , activesessions , poolname , pooltype , vmpowerstate | Format-Table -AutoSize | Out-String | Write-Verbose

            if( $null -eq $poolsProcessed[ $farm.name ] )
            {
                $poolsProcessed.Add( $farm.name , $vms.Count )
            }

            ForEach( $VM in $VMs )
            {
                [string]$divider = $poolTypeToFolder[ $VM.pooltype ]
                [string]$baseFolder = Join-Path -Path "Pod $($farm.hydraNodeName)" -ChildPath $divider

                ## requested folder structure is POD\RDS Farms\Farm
                if( ! $parentFolders[ $baseFolder ] )
                {
                    $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( (Split-Path -Path $baseFolder -Leaf) , $baseFolder , 'Folder' , "" , "" ,"" )) )
                    $parentFolders.Add( $baseFolder , $true )
                }

                [string]$farmLabel = "Farm $($farm.name)"
                [string]$nextLevel = Join-Path -Path $baseFolder -ChildPath $farmLabel

                ## requested folder structure is POD\RDS Farms\Farm
                if( ! $parentFolders[ $nextLevel ] )
                {
                    $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( (Split-Path -Path $nextLevel -Leaf) , $nextLevel , 'Folder' , "" , "" ,"" )) )
                    $parentFolders.Add( $nextLevel , $true )
                }

                Write-Verbose -Message "Adding computer $($VM.name) to folder `"$nextLevel`""

                $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( $VM.name , $nextLevel , 'Computer' , ( $VM.dnsname.Split( '@' )[-1] ) ,"" , $VM.dnsname )))
            }

        }
    }
}

if( ! $assignments -or $assignments.status -ne 'SUCCESS' )
{
    Write-Warning -Message "Call to $($RESTparams.uri) failed - $assignments"
}
else
{
    Write-Verbose -Message "Got $($assignments.data.Count) assignments"

    ##$ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( $Pooldivider , $Pooldivider , 'Folder' , "" , "" ,"" )) )

    ForEach( $pool in $assignments.data )
    {
        $RESTparams.Uri = $baseURL + ( '/api/mcw/assignments/{0}?workspace=true&containers=true' -f $pool.id )

        try
        {
            $poolDetails = Invoke-RestMethod @RESTparams
        }
        catch
        {
            $poolDetails = $null
            Write-Warning -Message "Failed to get assignment via $($RESTparams.uri) : $_"
        }

        if( $poolDetails -and $poolDetails.status -eq 'SUCCESS' -and $pooldetails.PSObject.Properties[ 'data' ] )
        {
            Write-Verbose -Message "Got $($poolDetails.data.desktop_containers.Count) desktop containers"

            ForEach( $desktopContainer in $poolDetails.data.desktop_containers )
            {
                Write-Verbose -Message "Got $($desktopContainer.machine_groups.Count) machine groups"
                ForEach( $machineGroup in $desktopContainer.machine_groups )
                {
                    $RESTparams.Uri = $baseURL + ( '/dt-rest/v100/infrastructure/pool/desktop/{0}/vms' -f $machineGroup.pool_or_farm_id )

                    try
                    {
                        $poolVMs = Invoke-RestMethod @RESTparams
                    }
                    catch
                    {
                        $poolVMs = $null
                        Write-Warning -Message "Failed to get machine group via $($RESTparams.uri) : $_"
                    }

                    if( $poolVMs )
                    {
                        Write-Verbose -Message "Got $($poolVMs.Count) VMs in pool $($machineGroup.name)"
                        $poolVMs | Select-Object -property name , guestOS , numcpus , memorySizeMB , ipaddress , sessionAllocationState , poolname , pooltype , vmpowerstate | Format-Table -AutoSize | Out-String | Write-Verbose

                        if( $null -eq $poolsProcessed[ $pool.name ] )
                        {
                            $poolsProcessed.Add( $pool.name , $poolVMs.Count )
                        }

                        ForEach( $poolVm in $poolVMs )
                        {
                            [string]$divider = $poolTypeToFolder[ $poolVM.pooltype ]
                            [string]$baseFolder = Join-Path -Path "Pod $($poolVm.hydraNodeName)" -ChildPath $divider

                            ## requested folder structure is POD\VDI Desktops\Pool
                            if( ! $parentFolders[ $baseFolder ] )
                            {
                                $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( (Split-Path -Path $baseFolder -Leaf) , $baseFolder , 'Folder' , "" , "" ,"" )) )
                                $parentFolders.Add( $baseFolder , $true )
                            }

                            [string]$poolName = $poolVM.poolname
                            if( $poolName -match "^$($pool.name)-[0-9a-f]{8}$" )
                            {
                                $poolName = $pool.Name
                            }
                            [string]$poolLabel = "Pool $poolName"
                            [string]$nextLevel = Join-Path -Path $baseFolder -ChildPath $poolLabel

                            ## requested folder structure is POD\VDI Desktops\Pool
                            if( ! $parentFolders[ $nextLevel ] )
                            {
                                $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( (Split-Path -Path $nextLevel -Leaf) , $nextLevel , 'Folder' , "" , "" ,"" )) )
                                $parentFolders.Add( $nextLevel , $true )
                            }

                            Write-Verbose -Message "Adding computer $($poolVM.name) to folder `"$nextLevel`""

                            $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( $poolVM.name , $nextLevel , 'Computer' , ( $poolVM.dnsname.Split( '@' )[-1] ) ,"" , $poolVM.dnsname )))
                        }
                    }
                }
            }
        }
    }
}

Write-Verbose -Message "Processed $($poolsProcessed.Count) pools"
$poolsProcessed.GetEnumerator() | Format-Table -AutoSize | Out-String | Write-Verbose

## See if there are any pools and thence VMs which we do not already have

$RESTparams.Uri = $baseURL + '/dt-rest/v100/pool/manager/findpools'
$RESTparams.Method = 'POST'
$RESTparams.Body = ( @{ searchName = '*' ; skipBuiltInPools = $true } | ConvertTo-Json ).ToString()

try
{
    $pools = Invoke-RestMethod @RESTparams ## for single pod broker customer, not universal
}
catch
{
    $pools = $null
    Throw "Failed to get pools via $($RESTparams.uri) : $_"
}

$RESTparams.Method = 'GET'
$RESTparams.Remove( 'Body' )

if( $pools )
{
    [int]$totalSize = $pools | Measure-Object -Property actualSize -Sum | Select-Object -ExpandProperty Sum
    Write-Verbose -Message "Got $($pools.Count) pools from /findpools, containing $totalSize machines"
    $pools | Select-Object -Property Name, HydraNodeName , PoolSessionType , PoolSizeType , sessionBased , actualSize , requestedSize , domainName , provisioningState , poolOnline | Format-Table -AutoSize | Out-String | Write-Verbose

    ForEach( $pool in $pools )
    {
        ## check pool name not already processed - could be decorated , e.g. Floating-45016da6 for Floating
        if( ( $null -eq ( $existingPool = $poolsProcessed[ $pool.name ] ) ) -and ( $null -eq ( $existingPool = $poolsProcessed[ ($pool.name -creplace '-[0-9a-f]{8}$' ) ] ) ) )
        {
            $RESTparams.Uri = $baseURL + ( '/dt-rest/v100/infrastructure/pool/desktop/{0}/vms' -f $pool.id )

            try
            {
                if( $poolVMs = Invoke-RestMethod @RESTparams )
                {
                    Write-Verbose -Message "Got $($poolVMs.Count) VMs in /findpools pool $($pool.name)"
                    $poolVMs | Select-Object -property name , guestOS , numcpus , memorySizeMB , ipaddress , sessionAllocationState , poolname , pooltype , vmpowerstate | Format-Table -AutoSize | Out-String | Write-Verbose

                    ForEach( $poolVm in $poolVMs )
                    {
                        [string]$divider = $poolTypeToFolder[ $poolVM.pooltype ]
                        [string]$baseFolder = Join-Path -Path "Pod $($poolVm.hydraNodeName)" -ChildPath $divider

                        ## requested folder structure is POD\VDI Desktops\Pool
                        if( ! $parentFolders[ $baseFolder ] )
                        {
                            $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( (Split-Path -Path $baseFolder -Leaf) , $baseFolder , 'Folder' , "" , "" ,"" )) )
                            $parentFolders.Add( $baseFolder , $true )
                        }

                        [string]$poolName = $poolVM.poolname -creplace '-[0-9a-f]{8}$'
                        [string]$poolLabel = "Pool $poolName"
                        [string]$nextLevel = Join-Path -Path $baseFolder -ChildPath $poolLabel

                        ## requested folder structure is POD\VDI Desktops\Pool
                        if( ! $parentFolders[ $nextLevel ] )
                        {
                            $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( (Split-Path -Path $nextLevel -Leaf) , $nextLevel , 'Folder' , "" , "" ,"" )) )
                            $parentFolders.Add( $nextLevel , $true )
                        }

                        Write-Verbose -Message "Adding computer $($poolVM.name) to folder `"$nextLevel`""

                        $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new( $poolVM.name , $nextLevel , 'Computer' , ( $poolVM.dnsname.Split( '@' )[-1] ) ,"" , $poolVM.dnsname )))
                    }
                }
            }
            catch
            {
                Write-Warning -Message "Failed to get desktops for pool $($pool.Name) id $($pool.id)"
            }
        }
        else
        {
            Write-Verbose -Message "Already got $existingPool VMs in existing pool `"$($pool.Name)`""
        }
    }
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

[int]$errorCount = Build-CUTree -ExternalTree $ControlUpEnvironmentObject @BuildCUTreeParams

Exit $errorCount
