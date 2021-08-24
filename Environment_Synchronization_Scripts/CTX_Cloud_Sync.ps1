#requires -version 3

<#
.SYNOPSIS
    Creates the folder structure and adds/removes or moves Citrix Cloud machines into the ControlUp structure.

.DESCRIPTION

.PARAMETER customerID

    The customer ID of the Citrix Cloud tenant to connect to

.PARAMETER clientID
    The client ID (GUID) of the API client created for use with this customer ID.
    If not specified, previously created credentials will be sought in C:\ProgramData\ControlUp\ScriptSupport for the user running the script

.PARAMETER clientSecret
    The client secret of the API client created for use with this customer ID
    If not specified, previously created credentials will be sought in C:\ProgramData\ControlUp\ScriptSupport for the user running the script

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

.NOTES
    Requires previously created stored credentials. Create a Citrix Cloud API client at https://us.cloud.com/identity/api-access/secure-clients

    MODIFICATION_HISTORY
    Guy Leech,            2021-08-18 - Initial version, copied from on-premises Citrix script
    Guy Leech,            2021-08-20 - Fixed issue where dash in Citrix site name caused 2 different folders to be created

#>

[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$true, HelpMessage='Enter Citrix Cloud customer id' )]
    [ValidateNotNullOrEmpty()]
    [string]$customerID,

    [Parameter(Mandatory=$false, HelpMessage='Enter Citrix Cloud API client id' )]
    [ValidateNotNullOrEmpty()]
    [guid]$clientID,

    [Parameter(Mandatory=$false, HelpMessage='Enter Citrix Cloud API client secret' )]
    [ValidateNotNullOrEmpty()]
    [string]$clientSecret,

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
    
    [Parameter(Mandatory=$false, HelpMessage='Use this domain for computer membership rather than what REST API returns' )]
    [string] $domainOverride,
    
    [Parameter(Mandatory=$false, HelpMessage='Creates the ControlUp folder structure based on the EUC Environment tree' )]
    [switch] $MatchEUCEnvTree,
    
    [string]$baseURI = 'https://api-us.cloud.com' ,

    [Parameter(Mandatory=$false, HelpMessage='A list of Delivery Groups to include.  Works with wildcards' )]
    [array] $includeDeliveryGroup,

    [Parameter(Mandatory=$false, HelpMessage='A list of Delivery Groups to exclude.  Works with wildcards. Exclusions supercede inclusions' )]
    [array] $excludeDeliveryGroup,

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

##$Global:LogFile = $LogFile
$sessionVariable = $null
[string]$errorMessage = $null

## this function should really go in the common Build-CUTree script
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
        $errorMessage = "Unable to find stored credential file `"$strCUCredFile`" - have you previously run the `"Create Credentials for Horizon Scripts`" script for user $env:username ?"
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
            $errorMessage = "Error reading stored credentials from `"$strCUCredFile`" : $_"
            Write-CULog -Msg $errorMessage -ShowConsole -Type E
            Throw $errorMessage
        }
    }
}
## Script from ControlUp which must reside in the same folder as this script
[string]$buildCuTreeScript = 'Build-CUTree.ps1'

function Make-NameWithSafeCharacters ([string]$string) {
    ###### TODO need to replace the folder path characters that might be illegal
    #list of illegal characters : '/', '\', ':', '*','?','"','<','>','|','{','}'
    $returnString = (($string).Replace('/','-')).Replace('\','-').Replace(':','-').Replace('*','-').Replace('?','-').Replace('''','-').Replace('<','-').Replace('>','-').Replace('|','-').Replace('{','-').Replace('}','-')
    return $returnString
}
    
#Create ControlUp structure object for synchronizing
## This should be in Build-CUTree script
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

if( ! $PSBoundParameters[ 'clientSecret' ] -or ! $PSBoundParameters[ 'clientId' ] )
{
    # Get the stored credentials for using the REST API
    [PSCredential]$CredsCitrixCloud = Get-CUStoredCredential -System 'CitrixCloud'

    if( ! $CredsCitrixCloud )
    {
        $errorMessage = "Failed to get stored API client secret for $env:username for Citrix Cloud"
        Write-CULog -Msg $errorMessage -ShowConsole -Type E
        Throw $errorMessage
    }
    if( [string]::IsNullOrEmpty( $clientId ))
    {
        $clientId = $CredsCitrixCloud.UserName
    }
    if( [string]::IsNullOrEmpty( $clientSecret ))
    {
        $clientSecret = $CredsCitrixCloud.GetNetworkCredential().Password
    }
}

# Setup TLS Setup
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12 ## -bor [Net.SecurityProtocolType]::Tls13

## https://stackoverflow.com/questions/41897114/unexpected-error-occurred-running-a-simple-unauthorized-rest-query?rq=1
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

[hashtable]$body = @{
    client_id = $clientID
    client_secret = $clientSecret
    grant_type = 'client_credentials'
}

[hashtable]$authHeaders = @{
    'Content-Type' = 'application/x-www-form-urlencoded'
    Accept = 'application/json'
}

## authenticate to get bearer token
## https://developer.cloud.com/explore-more-apis-and-sdk/cloud-services-platform/citrix-cloud-api-overview/docs/get-started-with-citrix-cloud-apis#bearer_token_tab_oauth_2.0_flow

if( ! ( $bearerToken = (Invoke-RestMethod -Uri "$baseURI/cctrustoauth2/$customerID/tokens/clients" -Method POST -Headers $authHeaders -Body $body -SessionVariable sessionVariable ) ) )
{
    $errorMessage = "Failed to get bearer token for customer id $customerID"
}
elseif( ! $bearerToken.PSObject.Properties[ 'token_type' ] )
{
    $errorMessage = "Bearer token missing from auth response - $bearerToken"
}
elseif( $bearerToken.token_type -ne 'bearer' )
{
    $errorMessage = "Unexpected token type `"$($bearerToke.token_type)`" from auth response"
}
elseif( ! $bearerToken.PSObject.Properties[ 'access_token' ] )
{
    $errorMessage = "Access token missing from auth response - $bearerToken"
}

if( $errorMessage )
{
    Write-CULog -Msg $errorMessage -ShowConsole -Type E
    Throw $errorMessage
}

[hashtable]$requestHeaders = @{
    Authorization = "CwsAuth Bearer=$($bearerToken.access_token)"
}

##$resourceLocations = Invoke-RestMethod -Uri "https://registry.citrixworkspacesapi.net/$customerId/resourcelocations" -Headers $requestHeaders -Method Get -WebSession $sessionVariable

## https://developer.cloud.com/citrixworkspace/virtual-apps-and-desktops/cvad-rest-apis/apis/Me-APIs/Me_GetMe
$requestHeaders += @{
    'Citrix-CustomerId' = $customerID
    Accept = 'application/json'
    'Content-Type' = 'application/json'
}

if( ! ( $me = Invoke-RestMethod -Uri "$baseURI/cvadapis/me" -Headers $requestHeaders -Method Get -WebSession $sessionVariable ) )
{
    $errorMessage = "Failed to get result from 'me' REST call"
    Write-CULog -Msg $errorMessage -ShowConsole -Type E
    Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$thisScript`" on $env:COMPUTERNAME" -body $_
    Throw $errorMessage
}

$DeliveryGroups = New-Object -Typename System.Collections.Generic.List[PSObject]
$BrokerMachines = New-Object -Typename System.Collections.Generic.List[PSObject]
$ControlUpEnvironmentObject = New-Object -Typename System.Collections.Generic.List[PSObject]

## Get delivery groups for each site
## https://developer.cloud.com/citrixworkspace/virtual-apps-and-desktops/cvad-rest-apis/apis/DeliveryGroups-APIs/DeliveryGroups_GetDeliveryGroups
$me.Customers.Where( { $_.Id -eq $customerID } ).ForEach(
{
    $customer = $_
    Write-Verbose -Message "Customer $($customer.id)"
    ForEach( $cloudSite in $customer.Sites.GetEnumerator() )
    {
        Write-Verbose -Message "Site $($cloudSite.Name) id $($cloudSite.Id)"
        if( ! ( [array]$cloudDeliveryGroups = @( Invoke-RestMethod -Uri "$baseURI/cvadapis/$($cloudSite.id)/DeliveryGroups" -Headers $requestHeaders -Method Get -WebSession $sessionVariable | Select-Object -ExpandProperty Items -ErrorAction SilentlyContinue) ) )
        {
            Write-Warning -Message "Failed to get any delivery groups for site $($cloudSite.id)"
        }
        else
        {
            Write-Verbose -Message "Got $($cloudDeliveryGroups.Count) delivery groups for site $($cloudSite.id)"
            ForEach( $deliveryGroup in $cloudDeliveryGroups )
            {
                $DeliveryGroupObject = [PSCustomObject]@{
                    MachineName         = ""
                    DNSName             = ""
                    Name                = $DeliveryGroup.Name
                    Site                = $cloudSite.Name
                    Broker              = $null
                }
                $DeliveryGroups.Add( $DeliveryGroupObject )
            }
        }

        if( ! ( [array]$cloudMachines = @( Invoke-RestMethod -Uri "$baseURI/cvadapis/$($cloudSite.id)/Machines" -Headers $requestHeaders -Method Get -WebSession $sessionVariable | Select-Object -ExpandProperty Items -ErrorAction SilentlyContinue) ) )
        {
            Write-Warning -Message "Failed to get any machines for site $($cloudSite.id)"
        }
        else
        {
            Write-Verbose -Message "Got $($cloudMachines.Count) machines for site $($cloudSite.id)"
            ForEach( $machine in $cloudMachines )
            {
                $BrokerMachineObject = [PSCustomObject]@{
                    MachineName         = ($machine.Name -split '\\')[-1] ## strip domain
                    DNSName             = $machine.DNSName
                    Name                = ""
                    Site                = $cloudSite.Name
                    Broker              = $null
                }
                $BrokerMachines.Add( $BrokerMachineObject )

                if ($machine.Name -like "*S-1-5*") {
                    Write-Warning "Detected a machine with a SID for a name. These cannot be added to ControlUp. Skipping: $($Machine.Name)" 
                } else {
                    if ([string]::IsNullOrEmpty($machine.DNSName)) {
                        $DNSName = $null
                    } else {
                        $DNSName = $machine.DNSName
                    }
                    if( $PSBoundParameters[ 'domainOverride' ] )
                    {
                        $domain = $domainOverride
                    }
                    else
                    {
                        $Domain = $Machine.Name.split("\")[0]
                    }
                    $Name =$Machine.Name.split("\")[-1]
                    try
                    {
                        if( $machine.DeliveryGroup -and ! [string]::IsNullOrEmpty( $machine.DeliveryGroup.Name ) )
                        {
                            $ControlUpEnvironmentObject.Add( [ControlUpObject]::new( $Name , $machine.DeliveryGroup.Name , "Computer" , $Domain , "$($CloudSite.Name)-$DNSName.Machine" , $DNSName ) )
                        }
                        else
                        {
                            Write-Warning -Message "Ignoring machine $name as not in a delivery group, catalog is `"$($machine.MachineCatalog | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue)`""
                        }
                    }
                    catch
                    {
                        Write-Warning -Message $_
                    }
                }
            }
        }
    }
} )

<#
Try
{
    $brokerParameters.AdminAddress = $adminAddr
    $CTXSite = Get-BrokerSite -AdminAddress $adminAddr
    $CTXSites.Add($CTXSite)
    Write-Verbose -Message "Querying Delivery Groups"
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
        }
    }

    Write-Verbose -Message "Querying  Machines"
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
catch
{
    Write-CULog -Msg $_ -ShowConsole -Type E
    Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$thisScript`" on $env:COMPUTERNAME" -body $_
}
#>

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

Write-Verbose -Message "Adding Delivery Groups to ControlUp Environmental Object"

foreach ($DeliveryGroup in $DeliveryGroups) {
        $ControlUpEnvironmentObject.Add( ([ControlUpObject]::new($($DeliveryGroup.Name) ,"$($DeliveryGroup.Name)","Folder","","$($DeliveryGroup.site)-DeliveryGroup","")))
}

#Add machines from the delivery group to the environmental object
<#
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
            if( $newObject = [ControlUpObject]::new( $Name , $DeliveryGroup.Name , "Computer" , $Domain , "$($DeliveryGroup.site)-Machine" , $DNSName ) )
            {
                $ControlUpEnvironmentObject.Add( $newObject )
            }
        }
    }
}
#>

## TYE
if ($MatchEUCEnvTree) {
    for ($i=0; $i -lt $ControlUpEnvironmentObject.Count; $i++) {
        if ($ControlUpEnvironmentObject[$i].FolderPath -eq "Brokers" -and $ControlUpEnvironmentObject[$i].Type -eq "Computer") {
            $BrokerObj = $BrokerMachines | Where-Object ({$_.MachineName.split("\")[1] -eq $ControlUpEnvironmentObject[$i].Name})
            $ControlUpEnvironmentObject[$i].FolderPath = "$($BrokerObj.Site)\$($ControlUpEnvironmentObject[$i].FolderPath)" #Sets the path to $SiteName\Brokers
        } else {
            $ControlUpEnvironmentObject[$i].FolderPath = "$($ControlUpEnvironmentObject[$i].Description -replace '\-(DeliveryGroup|Machine)$')\Delivery Groups\$($ControlUpEnvironmentObject[$i].FolderPath)"
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

<#
if ($Site){
    $BuildCUTreeParams.Add("SiteName",$Site)
}
#>

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

[int]$errorCount = Build-CUTree -ExternalTree $ControlUpEnvironmentObject @BuildCUTreeParams

Exit $errorCount
