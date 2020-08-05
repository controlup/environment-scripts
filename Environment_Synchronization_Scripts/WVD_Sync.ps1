<#
.SYNOPSIS
    Creates the folder structure and adds/removes or moves machines into the structure.
.DESCRIPTION
    Creates the folder structure and adds/removes or moves machines into the structure. Based on the WVD 2020 Spring Release.
.EXAMPLE
    . .\Azure_WVD_SyncScript.ps1 -folderPath VDI_and_SBC
.CONTEXT
    Windows Virtual Desktops
.MODIFICATION_HISTORY
    Esther Barthel, MSc - 06/06/20 - Original code
    Esther Barthel, MSc - 06/06/20 - Changed the script to support WVD 2020 Spring Release (ARM Architecture update)
    Trentent Tye,         07/06/20 - Updated to run on the ControlUp Monitor.

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-clixml?view=powershell-7

.COMPONENT
    Set-AzSPCredentials - The required Azure Service Principal (Subcription level) and tenantID information need to be securely stored in a Credentials File. The Set-AzSPCredentials Script Action will ensure the file is created according to ControlUp standards
    Az.Desktopvirtualization PowerShell Module - The Az.Desktopvirtualization PowerShell Module must be installed on the machine running this Script Action

.NOTES
    Requires Service Principal credentials stored in order for this script to work. In order to use as "LocalSystem" context, you must use PSExec to create
    the credential file. See documentation on ControlUp's website for how to create the Service Principal and how to create the stored credential file
    for the context you want to use.

    Version:        0.1
    Author:         Esther Barthel, MSc
    Creation Date:  2020-06-06
    Updated:        2020-07-06
                    Changed ...
    Purpose:        Script Action, created for ControlUp WVD Monitoring
        
    Copyright (c) cognition IT. All rights reserved.
#>


[CmdletBinding()]
Param
(
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        HelpMessage='Enter a subfolder to save your WVD tree'
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
    [string] $Site
) 


<#
## For debugging uncomment
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'continue'
$DebugPreference = 'SilentlyContinue'
Set-StrictMode -Version Latest
#>
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

#region WVD

# dot sourcing WVD Functions
. ".\WVDFunctions.ps1"
. ".\Build-CUTree.ps1"
        
## Check if the required PowerShell Modules are installed and can be imported
Invoke-CheckInstallAndImportPSModulePrereq -ModuleName "Az.Accounts" #-Verbose
Invoke-CheckInstallAndImportPSModulePrereq -ModuleName "Az.DesktopVirtualization" #-Verbose

$WVDEnvironment = New-Object System.Collections.Generic.List[PSObject]

If (Invoke-NETFrameworkCheck)
{
    If ($azSPCredentials = Get-AzSPStoredCredentials)
    {
        # Sign in to Azure with a Service Principal with Contributor Role at Subscription level
        try
        {
            $azSPSession = Connect-AzAccount -Credential $azSPCredentials.spCreds -Tenant $($azSPCredentials.tenantID).ToString() -ServicePrincipal -WarningAction SilentlyContinue
        }
        catch
        {
            Write-Error ("A [" + $_.Exception.GetType().FullName + "] ERROR occurred. " + $_.Exception.Message)
            Exit
        }

        # Retrieve the Subscription information for the Service Principal (that is logged on)
        $azSubscriptions = Get-AzSubscription
        Write-Verbose "Subscription Name = $($azSubscriptions.Name)"
        #region Generate ControlUpObject that will be used to Add objects into ControlUp.
        # Create the Azure Subscriptions folders
        foreach ($azSubscription in $azSubscriptions)
        {
            # Create the Azure Subscription folder that is linked to the Service Principal
            # The following characters are allowed: letters, numbers, space, dash and underscore. 
            $folderNameSubscription = Make-NameWithSafeCharacters -string $azSubscription.Name
            $ControlUpObject = [ControlUpObject]::new($folderNameSubscription,"$folderNameSubscription","Folder","","WVD Subcription","")
            $WVDEnvironment.Add($ControlUpObject)
            Write-Verbose ("SubscriptionFolderName: $folderNameSubscription")
            #Get all Hostpools
            $hostPoolQueryDuration = Measure-Command {
                $hostPools = Get-AzWvdHostPool -SubscriptionId $($azSubscription.Id.ToString())
            }
            Write-Verbose "Querying Hostpools took: $($hostPoolQueryDuration.TotalMilliseconds)"
            Write-Verbose "Number of Hostpools: $($($hostpools| Measure-Object).count)"
            foreach ($hostPool in $hostPools)
            {
                $folderNameHostPool = Make-NameWithSafeCharacters -string $hostPool.Name
                Write-Verbose ("HostPoolFolderName  : $folderNameHostPool")
                $parentPathHostPool = "$folderNameSubscription"
                $ControlUpObject = [ControlUpObject]::new($folderNameHostPool,"$parentPathHostPool`\$folderNameHostPool","Folder","","WVD Hostpool","")
                $WVDEnvironment.Add($ControlUpObject)
                $sessionHostQueryDuration = Measure-Command {
                    $sessionHosts = Get-AzWvdSessionHost -HostPoolName $hostpool.Name -ResourceGroupName $($hostPool.Id.Split("/")[4])
                }
                Write-Verbose "Querying Session Hosts took: $($sessionHostQueryDuration.TotalMilliseconds)"
                Write-Verbose "Found $($($sessionHosts | Measure-Object).count) session hosts"
                foreach ($sessionHost in $sessionHosts)
                {
                    #Get-SessionHosts
                    $sessionHostName = ($sessionHost.Name.Split("/")[1].Split(".")[0])
                    Write-Verbose "SessionHost Name    :  $sessionHostName"
                    $sessionHostDomain = $sessionHost.Name.substring($sessionHost.Name.IndexOf(".")+1)
                    $folderPathSessionHost = "$folderNameSubscription`\$folderNameHostPool"
                                
 
                    $ControlUpObject = [ControlUpObject]::new($sessionHostName,"$folderPathSessionHost","Computer","$sessionHostDomain","WVD SessionHost","")
                    $WVDEnvironment.Add($ControlUpObject)
                }
            }
        }
        #endregion
    }
}
#Write-Verbose "$($WVDEnvironment | Format-Table | Out-String)"
<#
Returns an object like so:
    Name                   FolderPath                           Type     Domain                      Description     DNSName
    ----                   ----------                           ----     ------                      -----------     -------
    Pay-As-You-Go Dev-Test Pay-As-You-Go Dev-Test               Folder                               WVD Subcription
    GPU                    Pay-As-You-Go Dev-Test\GPU           Folder                               WVD Hostpool
    GPU-WVD-0              Pay-As-You-Go Dev-Test\GPU           Computer AcmeOnAzure.onmicrosoft.com WVD SessionHost
    GPUNV                  Pay-As-You-Go Dev-Test\GPUNV         Folder                               WVD Hostpool
    Spring2020WVD          Pay-As-You-Go Dev-Test\Spring2020WVD Folder                               WVD Hostpool
    wvd-20spr-0            Pay-As-You-Go Dev-Test\Spring2020WVD Computer AcmeOnAzure.onmicrosoft.com WVD SessionHost
    wvd-20spr-1            Pay-As-You-Go Dev-Test\Spring2020WVD Computer AcmeOnAzure.onmicrosoft.com WVD SessionHost
    WVDHP                  Pay-As-You-Go Dev-Test\WVDHP         Folder                               WVD Hostpool
    WVDSH-0                Pay-As-You-Go Dev-Test\WVDHP         Computer AcmeOnAzure.onmicrosoft.com WVD SessionHost
    WVDSH-1                Pay-As-You-Go Dev-Test\WVDHP         Computer AcmeOnAzure.onmicrosoft.com WVD SessionHost
#>

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

Build-CUTree -ExternalTree $WVDEnvironment @BuildCUTreeParams


#endregion WVD



