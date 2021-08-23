#requires -Version 3.0

<#
    .SYNOPSIS
    Prepares the a PSCredential object on the target device for running ControlUp Horizon View scripts

    .DESCRIPTION
    This script creates an encrypted PSCredential object on the target machine in order to allow running of script for Horizon View without having to authenticate manually.

    .EXAMPLE
    This script should be run on any machines that will run Horizon View scripts. In general these are the machines that run the ControlUp Console or ControlUp Monitors. 
    
    .NOTES
    Connecting to a Horizon View Connection server is required for running Horizon View scripts. The server does not allow passthrough (Active Directory) authentication. In order to allow scripts to run without asking for a password each time (such as in Automated Actions) a PSCredential
    object needs to be stored on each target device (ie. each machine that will be used for running Horizon View scripts). This script can create this PSCredential object on the targets.
    PSCREDENTIAL OBJECTS CAN ONLY BE USED BY THE USER THAT CREATED THE OBJECT AND ON THE MACHINE THE OBJECT WAS CREATED.
    - The User that creates the file is required to have a local profile when creating the file. This is a limitation from Powershell
    
    Modification history:   20/08/2019 - Anthonie de Vreede - First version
                            03/06/2020 - Wouter Kursten - Second version
                            10/09/2020 - WOuter Kursten - Third Version
                            12/11/2020 - Guy Leech - added credential type argument for use with Horizon Cloud so one user can have multiple credentials
                            08/12/2020 - Added parameter sets with option for PSCredential object passing (pass as $null to prompt for credentials)
                            07/06/2021 - Merged Azure credentials script

    Changelog ;
        Second Version
            - Added check for local profile
            - changed error message when failing to create the xml file
            - Fixed issue where users without local admin rights and no active session on the target machine couldn't create a credentrials file ($env:USERPROFILE returns c:\users\default)

    .PARAMETER username
    The username for the PSCredential object
    
    .PARAMETER password
    The password for the credential object

    .PARAMETER passwordAgain
    Double check the password
    
    .PARAMETER credentialType
    The type of the credential

    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli
    https://github.com/vmware/PowerCLI-Example-Scripts/tree/master/Modules/VMware.Hv.Helper
#>

[CmdletBinding(DefaultParameterSetName='ClearText')]

Param
(
    [Parameter(Mandatory,HelpMessage='EUC environment to create credential file for')]
    [ValidateSet('HorizonView','Azure','HorizonCloudmyVMware','HorizonCloudDomain')]
    [string]$credentialType ,
    [Parameter(Mandatory,ParameterSetName='ClearText',HelpMessage='username to store in credential file - email or domain format')]
    [string]$userName ,
    [Parameter(Mandatory,ParameterSetName='ClearText',HelpMessage='Password')]
    [string]$password ,
    [Parameter(Mandatory,ParameterSetName='ClearText',HelpMessage='Password repeated')]
    [string]$passwordAgain ,
    [Parameter(Mandatory,ParameterSetName='Credential',HelpMessage='PSCredential object')]
    [System.Management.Automation.PSCredential]$credential ,
    [Parameter(Mandatory,ParameterSetName='Azure',HelpMessage='Service Principal Tenant Id')]
    [string]$tenantId ,
    [Parameter(Mandatory,ParameterSetName='Azure',HelpMessage='Service Principal Application (client) Id')]
    [string]$applicationId ,
    [Parameter(Mandatory,ParameterSetName='Azure',HelpMessage='Service Principal Application (client) secret')]
    [string]$applicationSecret
)

$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

[int]$outputwidth = 400

if( ( $PSWindow = (Get-Host).UI.RawUI ) -and ( $WideDimensions = $PSWindow.BufferSize ) )
{
    $WideDimensions.Width = $outputWidth
    $PSWindow.BufferSize = $WideDimensions
}

Function Out-CUConsole {
    <# This function provides feedback in the console on errors or progress, and aborts if error has occured.
    If only Message is passed this message is displayed
    If Warning is specified the message is displayed in the warning stream (Message must be included)
    If Stop is specified the stop message is displayed in the warning stream and an exception with the Stop message is thrown (Message must be included)
    If an Exception is passed a warning is displayed and the exception is thrown
    If an Exception AND Message is passed the Message message is displayed in the warning stream and the exception is thrown
    #>

    Param (
        [Parameter(Mandatory = $false)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$Warning,
        [Parameter(Mandatory = $false)]
        [switch]$Stop,
        [Parameter(Mandatory = $false)]
        $Exception
    )

    # Throw error, include $Exception details if they exist
    if ($Exception) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        If ($Message) {
            Write-Warning -Message "$Message`n$($Exception.CategoryInfo.Category)`nPlease see the Error tab for the exception details."
            Write-Error "$Message`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)`n" -ErrorAction Stop
        }
        Else {
            Write-Warning "There was an unexpected error: $($Exception.CategoryInfo.Category)`nPlease see the Error tab for details."
            Throw $Exception
        }
    }
    elseif ($Stop) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        Write-Warning -Message "There was an error.`n$Message"
        Throw $Message
    }
    elseif ($Warning) {
        # Write the warning to Warning stream, thats it. It's a warning.
        Write-Warning -Message $Message
    }
    else {
        # Not an exception or a warning, output the message
        Write-Output -InputObject $Message
    }
}

Function New-CUStoredCredential {
    param (
        [parameter(Mandatory = $true,
            HelpMessage = "The username to be stored in the PSCredential object.")]
        [string]$Username,
        [parameter(Mandatory = $true,
            HelpMessage = "The password to be stored in the PSCredential object.")]
        [string]$password ,
        [parameter(Mandatory = $false,
            HelpMessage = "The Azure Service Principal tenant id")]
        [string]$tenantId,
        [parameter(Mandatory = $true,
            HelpMessage = "The system the credentials will be used for.")]
        [string]$System
    )
    # Username and password correct, check if target folder exists and create it if necessary
    
    $strCredTargetFolder = [System.IO.Path]::Combine( [Environment]::GetFolderPath( [Environment+SpecialFolder]::CommonApplicationData ) , 'ControlUp' , 'ScriptSupport' )

    # Create the folder if it does not exist
    If ( ! (Test-Path -Path $strCredTargetFolder -ErrorAction SilentlyContinue)) {
        Write-Output "Folder does not exist"
        try {
            if( ! ( $newFolder = New-Item -Path $strCredTargetFolder -ItemType Directory ) ) {
                Write-Warning -Message "Problem creating folder `"$strCredTargetFolder`""
            }
        }
        catch {
            Out-CUConsole -Message "There was a problem creating the folder used to store the credentials object ($strCredTargetFolder). Please make sure you have permission to write to the parent folder." -Exception $_
        }
    }

    # Create the PSCredential object
    [System.Management.Automation.PSCredential]$Cred = $null

    try {
        [System.Security.SecureString]$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $Cred = New-Object System.Management.Automation.PSCredential ( $UserName , $SecurePassword ) 
    }
    catch {
        Out-CUConsole -Message "There was a problem creating the PSCredential object." -Exception $_
    }

    if( $Cred ) {
        [string]$credsfile = [System.IO.Path]::Combine( $strCredTargetFolder , ( $Env:Username + '_' + $System + '_Cred.xml' ) )
        Write-Verbose -Message "Writing credentials to file `"$credsfile`""

        # Store the PSCredential object or the Azure details
        try {
            $export = $cred

            if( ! [string]::IsNullOrEmpty( $tenantId ) )
            {
                [string]$guidRegex = '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
                if( $tenantId -notmatch $guidRegex )
                {
                    $export = $null
                    Out-CUConsole -Message "Tenant id `"$tenantId`" is not a correctly formed GUID"
                }
                elseif( $Username -notmatch $guidRegex )
                {
                    $export = $null
                    Out-CUConsole -Message "Application id `"$Username`" is not a correctly formed GUID"
                }
                else
                {
                    $export = @{
                        'tenantID' = $tenantID
                        'spCreds' = $cred }
                }
            }
            if( $export )
            {
                Export-Clixml -Path $credsfile -InputObject $export -Force

                Out-CUConsole -Message "Credential object created and stored in `"$credsFile`"" 
            }
        }
        catch {
            Remove-Item -path $credsfile -force
            Out-CUConsole -Message "There was a problem saving the PSCredential object to `"$credsfile`" - this may be a permission issue or there is no local profile." -Exception $_
        }
    }
}

$userprofile = $env:USERPROFILE

if( ! (Get-CimInstance -Classname win32_userprofile | Where-Object localpath -eq $userprofile )){
    Out-CUConsole -message "User $Env:Username has no profile on this system. This is a requirement for creating the credentials file. Please log onto this machine once in order to create your user profile."  -exception "No local profile found" # this is a limitation of Powershell
}

If ( $credentialType -match 'Azure' ) {
    if( $PsCmdlet.ParameterSetName -eq 'Azure' ) {
        New-CUStoredCredential -Username $applicationId -Password $applicationSecret -System $credentialType -TenantId $tenantId
    }
    else {
        Out-CUConsole -Message "Wrong parameters used for $credentialType credential type - use -applicationId, -applicationSecret & -tenantId" -Stop
    }
}
ElseIf( $PsCmdlet.ParameterSetName -eq 'Credential' )
{
    New-CUStoredCredential -Username $credential.userName -Password $credential.GetNetworkCredential().Password -System $credentialType
}
ElseIf (!([string]::IsNullOrWhiteSpace( $userName )) -and !([string]::IsNullOrWhiteSpace( $password )) -and $password -eq $passwordAgain ) {
    New-CUStoredCredential -Username $userName -Password $password -System $credentialType
}
Else {
    If ($password -ne $passwordAgain ) {
        Out-CUConsole -Message "The passwords do not match. Please enter the same password in both password fields." -Stop
    }
}
