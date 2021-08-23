
<#
    .SYNOPSIS
        Provides the function Build-CUTree which creates/updates a ControlUp folder structure with an external source (Active Directory, Citrix, Horizon, WVD)

    .DESCRIPTION
        See the help for Build-CUTree for usage detail once the script has been dot sourced

    .NOTES
    
    MODIFICATION_HISTORY

        @guyrleech 2020-10-13   Build-CUTree returns error count
        @guyrleech 2020-10-26   Bug fixed when deleting - path repeated organization name so never matched items to delete
        @guyrleech 2020-10-30   Bug fixed where computers were not being deleted, fixed bug in dll version detection
        @guyrleech 2020-11-02   Workaround for bug where batch folder creation fails where folder name already exists at top level
        @guyrleech 2020-12-20   Reorganised help comment block to be get-help compatible for script and function
        @guyrleech 2021-02-11   Change SiteId to SiteName and errors if does not exist. Added batch folder warning as can cause issues
        @guyrleech 2021-02-12   Added delay between each folder add when a large number being added
        @guyrleech 2021-07-29   Added more logging to log file. Added email notification
        @guyrleech 2021-08-13   Added checking and more logging for CU Monitor service state
        @guyrleech 2021-08-16   Changed service checking as was causing access denied errors
#>

<#
    .SYNOPSIS
	    Synchronizes ControlUp folder structure with an external source (Active Directory, Citrix, Horizon, WVD)

    .DESCRIPTION
        This function is meant to be dot-sourced into your external source script.  The Build-CUTree function expects to be
        passed an ExternalTree object of a specific format.

        Please see an existing sync script for an exmample of building this object.

        The expectations for this function is the ExternalTree object being passed to it will have had all objects you want sync'ed to be
        included and extra machines/folders excluded already. It assumes the ExternalTree object as a source of truth so no filtering
        or exclusions will occur in this function.

        This function has the following features:
          - Specify a folder in ControlUp to sync the external source
          - Remove all extra objects in the ControlUp folder path but not in the ExternalTree
          - Move any ControlUp machines detected outside the path into the path that matches the ExternalTree
          - Preview actions to be executed
          - Save a preview log
    
        The default executions will be run in a 'batch' fashion of 100 folder operations or 1000 computer operations. Each type of operation
        gets its own batch and will execute in kind.

        In order to use this function, it must be run on a ControlUp Monitor server. Any outside credentials needed to connect to external sources
        will be required to be stored with the user account/profile specified.  Please see the external source script you wish to you for further
        documentation.

    .PARAMETER ExternalTree
        An object that represents the logical formation of your external object.
        The object should in this format:

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

    .PARAMETER CURootFolder
        The folder in the ControlUp console that will be the root for the ExternalTree. Do NOT specify the organization name in this path, just the folder path
        underneath. For instance, in ControlUp if I right-click the folder I want to use as my root, and select "Properties" the "path:" will say something like:
        "a.c.m.e.\vdi_and_sbc\wvd".  In this example, "a.c.m.e." is the organization name so ignore it and just enter the folders underneath, which in this
        example is "vdi_and_sbc\wvd".

    .PARAMETER Delete
        Enables removal of excess objects in the sync directory. If a machine or folder object is found in the ControlUp path but not in the ExternalTree object
        the object will be marked for removal. If you do not use this parameter than only Add or Move operations will occur.

    .PARAMETER Preview
        Generates a preview showing the actions this script wants to take. If more than 25 operations is going to be executed the script will just
        return the number of operations. To see the individual operations use the PreviewOutputPath parameter. See that paramater help for more information.
            + CU Computers Count: 101
            + Organization Name: a.c.m.e.
            + Target Folder Paths:
                    > a.c.m.e.\VDI_and_SBC\WVD\Pay-As-You-Go Dev-Test
                    > a.c.m.e.\VDI_and_SBC\WVD\Pay-As-You-Go Dev-Test\GPU
                    > a.c.m.e.\VDI_and_SBC\WVD\Pay-As-You-Go Dev-Test\GPUNV
                    > a.c.m.e.\VDI_and_SBC\WVD\Pay-As-You-Go Dev-Test\Spring2020WVD
                    > a.c.m.e.\VDI_and_SBC\WVD\Pay-As-You-Go Dev-Test\WVDHP
            + External Computers Total Count: 5
            + Folders to Add     : 2
                    > Folders to Add Batches     : 1
                    > Add-CUFolder -Name GPU -ParentPath "a.c.m.e.\VDI_and_SBC\WVD\Pay-As-You-Go Dev-Test"
                    > Add-CUFolder -Name GPUNV -ParentPath "a.c.m.e.\VDI_and_SBC\WVD\Pay-As-You-Go Dev-Test"
            + Folders to Remove  : 0
                    > Folders to Remove Batches  : 0
            + Computers to Add   : 2
                    > Computers to Add Batches   : 1
                    > Add-CUComputer -Domain AcmeOnAzure.onmicrosoft.com -Name GPU-WVD-0 -FolderPath "a.c.m.e.\VDI_and_SBC\WVD\Pay-As-You-Go Dev-Test\GPU"
                    > Add-CUComputer -Domain AcmeOnAzure.onmicrosoft.com -Name WVDSH-1 -FolderPath "a.c.m.e.\VDI_and_SBC\WVD\Pay-As-You-Go Dev-Test\WVDHP"
            + Computers to Move  : 0
                    > Computers to Move Batches  : 0
            + Computers to Remove: 0
                    > Computers to Remove Batches: 0
            + Build-CUTree took: 0 Seconds.
            + Committing Changes:
                    > Executing Folder Object Adds. Batch 1/1
                    > Execution Time: PREVIEW MODE
                    > Executing Computer Object Adds. Batch 1/1
                    > Execution Time: PREVIEW MODE

    .PARAMETER LogFile
        Specifies that all output will be saved to a log file. Individual operations will also be saved to the log file. The operations are saved in
        such a way that you should be able to copy-paste them into a powershell prompt that has the ControlUp Powershell modules loaded and they should
        be executable.  Use this for testing individual operations to validate it will work as you expect.

    .PARAMETER SiteName
        An optional parameter to specify which site you want the machine object assigned. By default, the site name is "Default". Enter the name of the site
        to assign the object

    .PARAMETER DebugCUMachineEnvironment
        CONTROLUP INTERNAL USE ONLY

    .PARAMETER DebugCUFolderEnvironment
        CONTROLUP INTERNAL USE ONLY

    .PARAMETER batchCreateFolders
        Create ControlUp folders in batches rather than one by one

    .PARAMETER batchCountWarning
    
        When the number of new folders to add exceeds this number, either a warning will be produced and throttling introduced or the operation will be aborted if -force is not specified

    .PARAMETER force
    
        When the number of new folders to add exceeds -batchCountWarning a warning will be produced and throttling introduced otherwise the operation will be aborted

    .PARAMETER folderCreateDelaySeconds
    
        When the number of new folders to add exceeds -batchCountWarning, a delay of this number of seconds will be introduced between each folder creation when -force is specified.
        The delay can also be set using the %CU_delay% environment variable

    .EXAMPLE
        Build-CUTree -ExternalTree $WVDEnvironment -CURootFolder "VDI_and_SBC\WVD" -Preview -Delete -PreviewOutputPath C:\temp\sync.log
            Executes a logged preview of what sync'ing the WVDEnvironment object to the ControlUp folder VDI_and_SBC\WVD with what object removals would look like.

    .EXAMPLE
        Build-CUTree -ExternalTree $WVDEnvironment -CURootFolder "VDI_and_SBC\WVD" -Delete
            Executes a sync of the $WVDEnvironment object to the ControlUp folder "VDI_and_SBC\WVD" with object removal enabled.

    .NOTES
	    Runs on a ControlUp Monitor computer
	    Connects to an external source, retrieves the folder structure to synchronize
	    Adds to ControlUp folder structure all folders and computers from the external source
	    Moves folders and computers which exist in locations that differ from the external source
	    Optionally, removes folders and computers which do not exist in the external source
#>
function Build-CUTree {
    [CmdletBinding()]
    Param
    (

	    [Parameter(Mandatory=$true,HelpMessage='Object to build tree within ControlUp')]
	    [PSObject] $ExternalTree,

	    [Parameter(Mandatory=$false,HelpMessage='ControlUp root folder to sync')]
	    [string] $CURootFolder,

 	    [Parameter(Mandatory=$false, HelpMessage='Delete CU objects which are not in the external source' )]
	    [switch] $Delete,

        [Parameter(Mandatory=$false, HelpMessage='Generate a report of the actions to be executed' )]
        [switch]$Preview,

        [Parameter(Mandatory=$false, HelpMessage='Save a log file' )]
	    [string] $LogFile,

        [Parameter(Mandatory=$false, HelpMessage='ControlUp Site name to assign the machine object to' )]
	    [string] $SiteName,

        [Parameter(Mandatory=$false, HelpMessage='Debug CU Machine Environment Objects' )]
	    [Object] $DebugCUMachineEnvironment,

        [Parameter(Mandatory=$false, HelpMessage='Debug CU Folder Environment Object' )]
	    [switch] $DebugCUFolderEnvironment ,

        [Parameter(Mandatory=$false, HelpMessage='Create folders in batches rather than individually' )]
	    [switch] $batchCreateFolders ,

        [Parameter(Mandatory=$false, HelpMessage='Number of folders to be created that generates warning and requires -force' )]
        [int] $batchCountWarning = 100 ,
        
        [Parameter(Mandatory=$false, HelpMessage='Force creation of folders if -batchCountWarning size exceeded' )]
        [switch] $force ,
        
        [Parameter(Mandatory=$false, HelpMessage='Smtp server to send alert emails from' )]
	    [string] $SmtpServer ,

        [Parameter(Mandatory=$false, HelpMessage='Email address to send alert email from' )]
	    [string] $emailFrom ,

        [Parameter(Mandatory=$false, HelpMessage='Email addresses to send alert email to' )]
	    [string[]] $emailTo ,

        [Parameter(Mandatory=$false, HelpMessage='Use SSL to send email alert' )]
	    [switch] $emailUseSSL ,

        [Parameter(Mandatory=$false, HelpMessage='Delay between each folder creation when count exceeds -batchCountWarning' )]
        [double] $folderCreateDelaySeconds = 0.5
    )

    Begin {

        #This variable sets the maximum computer batch size to apply the changes in ControlUp. It is not recommended making it bigger than 1000
        $maxBatchSize = 1000
        #This variable sets the maximum batch size to apply the changes in ControlUp. It is not recommended making it bigger than 100
        $maxFolderBatchSize = 100
        [int]$errorCount = 0
        [array]$stack = @( Get-PSCallStack )
        [string]$callingScript = $stack.Where( { $_.ScriptName -ne $stack[0].ScriptName } ) | Select-Object -First 1 -ExpandProperty ScriptName
        if( ! $callingScript -and ! ( $callingScript = $stack | Select-Object -First 1 -ExpandProperty ScriptName ) )
        {
            $callingScript = $stack[-1].Position ## if no script name then use this which should give us the full command line used to invoke the script
        }

        function Execute-PublishCUUpdates {
            Param(
	            [Parameter(Mandatory = $True)][Object]$BatchObject,
	            [Parameter(Mandatory = $True)][string]$Message
            )
            [int]$returnCode = 0
            [int]$batchCount = 0
            foreach ($batch in $BatchObject) {
                $batchCount++
                Write-CULog -Msg "$Message. Batch $batchCount/$($BatchObject.count)" -ShowConsole -Color DarkYellow -SubMsg
                if (-not($preview)) {
                    [datetime]$timeBefore = [datetime]::Now
                    $result = Publish-CUUpdates -Batch $batch 
                    [datetime]$timeAfter = [datetime]::Now
                    [array]$results = @( Show-CUBatchResult -Batch $batch )
                    [array]$failures = @( $results.Where( { $_.IsSuccess -eq $false } )) ## -and $_.ErrorDescription -notmatch 'Folder with the same name already exists' } ) )

                    Write-CULog -Msg "Execution Time: $(($timeAfter - $timeBefore).TotalSeconds) seconds" -ShowConsole -Color Green -SubMsg
                    Write-CULog -Msg "Result: $result" -ShowConsole -Color Green -SubMsg
                    Write-CULog -Msg "Failures: $($failures.Count) / $($results.Count)" -ShowConsole -Color Green -SubMsg

                    if( $failures -and $failures.Count -gt 0 ) {
                        $returnCode += $failures.Count
                        ForEach( $failure in $failures ) {
                            Write-CULog -Msg "Action $($failure.ActionName) on `"$($failure.Subject)`" gave error $($failure.ErrorDescription) ($($failure.ErrorCode))" -ShowConsole -Type E
                        }
                    }
                } else {
                    Write-CULog -Msg "Execution Time: PREVIEW MODE" -ShowConsole -Color Green -SubMsg
                }
            }
        }
        
        <#
        ## Paths must be absolute
        function Test-CUFolderPath {
            Param(
                [parameter(Mandatory = $true,
                HelpMessage = "Specifies a path to be tested. The value of the Path parameter is case insensitive and used exactly as it is typed. No characters are interpreted as wildcard characters.")]
                [string]$Path
            )
            ## GRL Previous method relied on checking a cache of folders which did not have newly crated folders in. Apparently there's a risk though that Get-CUFolders can miss recently created folders.
            [string]$trimmed = $path.Trim( '\' )
            Get-CUFolders | Where-Object { ( $_.FolderType -eq 'Folder' -or $_.FolderType -eq 'RootFolder' ) -and $_.Path -eq $trimmed } | . { Process {
                return $true
            }}

            return $false
        }
        #>

        #attempt to setup the log file
        if ($PSBoundParameters.ContainsKey("LogFile")) {
            $Global:LogFile = $PSBoundParameters.LogFile
            Write-Host "Saving Output to: $Global:LogFile"
            if (-not(Test-Path $($PSBoundParameters.LogFile))) {
                Write-CULog -Msg "Creating Log File" #Attempt to create the file
                if (-not(Test-Path $($PSBoundParameters.LogFile))) {
                    Write-Error "Unable to create the report file" -ErrorAction Stop
                }
            } else {
                Write-CULog -Msg "Beginning Synchronization"
            }
            Write-CULog -Msg "Detected the following parameters:"
            foreach($psbp in $PSBoundParameters.GetEnumerator())
            {
                if ($psbp.Key -like "ExternalTree" -or $psbp.Key -like "DebugCUMachineEnvironment") {
                    Write-CULog -Msg $("Parameter={0} Value={1}" -f $psbp.Key,$psbp.Value.count)
                } else {
                    Write-CULog -Msg $("Parameter={0} Value={1}" -f $psbp.Key,$psbp.Value)
                }
            }
        } else {
            $Global:LogFile = $false
        }

        if( ! $PSBoundParameters[ 'folderCreateDelaySeconds' ] -and $env:CU_delay )
        {
            $folderCreateDelaySeconds = $env:CU_delay
        }
    }

    Process {
        $startTime = Get-Date
        [string]$errorMessage = $null

        #region Load ControlUp PS Module
        try
        {
            ## Check CU monitor is installed and at least minimum required version
            [string]$cuMonitor = 'ControlUp Monitor'
            [string]$cuDll = 'ControlUp.PowerShell.User.dll'
            [string]$cuMonitorProcessName = 'CUmonitor'
            [version]$minimumCUmonitorVersion = '8.1.5.600'
            if( ! ( $installKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -Name DisplayName -ErrorAction SilentlyContinue| Where-Object DisplayName -eq $cuMonitor ) )
            {
                Write-CULog -ShowConsole -Type W -Msg "$cuMonitor does not appear to be installed"
            }
            ## when running via scheduled task we do not have sufficient rights to query services
            if( ! ( $cuMonitorProcess = Get-Process -Name $cuMonitorProcessName -ErrorAction SilentlyContinue ) )
            {
                Write-CULog -ShowConsole -Type W -Msg "Unable to find process $cuMonitorProcessName for $cuMonitor service" ## pid $($cuMonitorService.ProcessId)"
            }
            else
            {
                [string]$message =  "$cuMonitor service running as pid $($cuMonitorProcess.Id)"
                ## if not running as admin/elevated then won't be able to get start time
                if( $cuMonitorProcess.StartTime )
                {
                    $message += ", started at $(Get-Date -Date $cuMonitorProcess.StartTime -Format G)"
                }
                Write-CULog -Msg $message
            }

	        # Importing the latest ControlUp PowerShell Module - need to find path for dll which will be where cumonitor is running from. Don't use Get-Process as may not be elevated so would fail to get path to exe and win32_service fails as scheduled task with access denied
            if( ! ( $cuMonitorServicePath = ( Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\cuMonitor' -Name ImagePath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ImagePath ) ) )
            {
                Throw "$cuMonitor service path not found in registry"
            }
            elseif( ! ( $cuMonitorProperties = Get-ItemProperty -Path $cuMonitorServicePath.Trim( '"' ) -ErrorAction SilentlyContinue) )
            {
                Throw  "Unable to find CUmonitor service at $cuMonitorServicePath"
            }
            elseif( $cuMonitorProperties.VersionInfo.FileVersion -lt $minimumCUmonitorVersion )
            {
                Throw "Found version $($cuMonitorProperties.VersionInfo.FileVersion) of cuMonitor.exe but need at least $($minimumCUmonitorVersion.ToString())"
            }
            elseif( ! ( $pathtomodule = Join-Path -Path (Split-Path -Path $cuMonitorServicePath.Trim( '"' ) -Parent) -ChildPath $cuDll ) )
            {
                Throw "Unable to find $cuDll in `"$pathtomodule`""
            }
	        elseif( ! ( Import-Module $pathtomodule -PassThru ) )
            {
                Throw "Failed to import module from `"$pathtomodule`""
            }
            elseif( ! ( Get-Command -Name 'Get-CUFolders' -ErrorAction SilentlyContinue ) )
            {
                Throw "Loaded CU Monitor PowerShell module from `"$pathtomodule`" but unable to find cmdlet Get-CUFolders"
            }
        }
        catch
        {
            $exception = $_
            Write-CULog -Msg $exception -ShowConsole -Type E
            Write-CULog -Msg (Get-PSCallStack|Format-Table)
            Write-CULog -Msg 'The required ControlUp PowerShell module was not found or could not be loaded. Please make sure this is a ControlUp Monitor machine.' -ShowConsole -Type E
            Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$exception"
            $errorCount++
            break
        }
        #endregion

        #region validate SiteName parameter
        [hashtable] $SiteIdParam = @{}
        [string]$SiteIdGUID = $null
        if ($PSBoundParameters.ContainsKey("SiteName")) {
            Write-CULog -Msg "Assigning resources to specific site: $SiteName" -ShowConsole
            
            [array]$cusites = @( Get-CUSites )
            if( ! ( $SiteIdGUID = $cusites | Where-Object { $_.Name -eq $SiteName } | Select-Object -ExpandProperty Id ) -or ( $SiteIdGUID -is [array] -and $SiteIdGUID.Count -gt 1 ) )
            {
                $errorMessage = "No unique ControlUp site `"$SiteName`" found (the $($cusites.Count) sites are: $(($cusites | Select-Object -ExpandProperty Name) -join ' , ' ))"
                Write-CULog -Msg $errorMessage -ShowConsole -Type E
                Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$exception"
                $errorCount++
                break
            }
            else
            {
                Write-CULog -Msg "SiteId GUID: $SiteIdGUID" -ShowConsole -SubMsg
                $SiteIdParam.Add( 'SiteId' , $SiteIdGUID )
            }
        }

        #region Retrieve ControlUp folder structure
        if (-not($DebugCUMachineEnvironment)) {
            try {
                $CUComputers = Get-CUComputers # add a filter on path so only computers within the $rootfolder are used
            } catch {
                $errorMessage = "Unable to get computers from ControlUp: $_" 
                Write-CULog -Msg $errorMessage -ShowConsole -Type E
                Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$errorMessage"
                $errorCount++
                break
            }
        } else {
            Write-Debug "Number of objects in DebugCUMachineEnvironment: $($DebugCUMachineEnvironment.count)"
            if ($($DebugCUMachineEnvironment.count) -eq 2) {
                foreach ($envObjects in $DebugCUMachineEnvironment) {
                    if  ($($envObjects  | Get-Member).TypeName[0] -eq "Create-CrazyCUEnvironment.CUComputerObject") {
                        $CUComputers = $envObjects
                    }
                }
            } else {
                $CUComputers = $DebugCUMachineEnvironment
            }
        }
        
        Write-CULog -Msg  "CU Computers Count: $(if( $CUComputers ) { $CUComputers.count } else { 0 } )" -ShowConsole -Color Cyan
        #create a hashtable out of the CUMachines object as it's much faster to query. This is critical when looking up Machines when ControlUp contains ten's of thousands of machines.
        $CUComputersHashTable = @{}
        foreach ($machine in $CUComputers) {
            foreach ($obj in $machine) {
                $CUComputersHashTable.Add($Obj.Name, $obj)
            }
        }

        if (-not($DebugCUFolderEnvironment)) {
            try {
                $CUFolders   = Get-CUFolders # add a filter on path so only folders within the rootfolder are used
            } catch {
                $errorMessage = "Unable to get folders from ControlUp: $_"
                Write-CULog -Msg $errorMessage  -ShowConsole -Type E
                Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$errorMessage"
                $errorCount++
                break
            }
        } else {
            Write-Debug "Number of folder objects in DebugCUMachineEnvironment: $($DebugCUMachineEnvironment.count)"
            if ($($DebugCUMachineEnvironment.count) -eq 2) {
                foreach ($envObjects in $DebugCUMachineEnvironment) {
                    if  ($($envObjects  | Get-Member).TypeName[0] -eq "Create-CrazyCUEnvironment.CUFolderObject") {
                        $CUFolders = $envObjects
                    }
                }
            } else {
                $CUFolders = Get-CUFolders
            }
        }

        #endregion
        $OrganizationName = ($CUFolders)[0].path
        Write-CULog -Msg "Organization Name: $OrganizationName" -ShowConsole

        [array]$rootFolders = @( Get-CUFolders | Where-Object FolderType -eq 'RootFolder' )

        Write-Verbose -Message "Got $($rootFolders.Count) root folders/organisations: $( ($rootFolders | Select-Object -ExpandProperty Path) -join ' , ' )"

        [string]$pathSoFar = $null
        [bool]$builtPath = $false
        ## strip off leading \ as CU cmdlets don't like it
        [string[]]$CURootFolderElements = @( ($CURootFolder.Trim( '\' ).Split( '\' ) ))
        Write-Verbose -Message "Got $($CURootFolderElements.Count) elements in path `"$CURootFolder`""

        ## see if first folder element is the organisation name and if not then we will prepend it as must have that
        if( $OrganizationName -ne $CURootFolderElements[0] ) {
            Write-CULog -Msg "Organization Name `"$OrganizationName`" not found in path `"$CURootFolder`" so adding" -Verbose
            $CURootFolder = Join-Path -Path $OrganizationName -ChildPath $CURootFolder
        }

        ## Code making folders checks if each element in folder exists and if not makes it so no pointmaking path here

        #region Prepare items for synchronization
        #replace FolderPath in ExternalTree object with the local ControlUp Path:
        foreach ($obj in $externalTree) {
            $obj.FolderPath = (Join-Path -Path $CURootFolder -ChildPath $obj.FolderPath).Trim( '\' ) ## CU doesn't like leading \
        }

        #We also create a hashtable to improve lookup performance for computers in large organizations.
        $ExtTreeHashTable = @{}
        $ExtFolderPaths = New-Object -TypeName System.Collections.Generic.List[psobject]
        foreach ($ExtObj in $externalTree) {
            foreach ($obj in $ExtObj) {
                ## GRL only add computers since that is all we look up and get duplicate error if OU and computer have the same name
                if( $obj.Type -eq 'Computer' ) {
                    $ExtTreeHashTable.Add($Obj.Name, $obj)
                }
                else {
                    $ExtFolderPaths.Add( $obj )
                }
            }
        }

        Write-CULog -Msg "Target Folder Paths:" -ShowConsole
        if ($ExtFolderPaths.count -ge 25) {
            Write-CULog "$($ExtFolderPaths.count) paths detected" -ShowConsole -SubMsg
            Foreach ($ExtFolderPath in $ExtFolderPaths) {
                Write-CULog -Msg "$($ExtFolderPath.FolderPath)" -SubMsg
            }
        } else {
            Foreach ($ExtFolderPath in $ExtFolderPaths) {
                Write-CULog -Msg "$($ExtFolderPath.FolderPath)" -ShowConsole -SubMsg
            }
        }

        $FolderAddBatches   = New-Object System.Collections.Generic.List[PSObject]
        $FoldersToAddBatch  = New-CUBatchUpdate
        $FoldersToAddCount  = 0

        #we'll output the statistics at the end -- also helps with debugging
        $FoldersToAdd          = New-Object System.Collections.Generic.List[PSObject]
        ## There can be problems when folders are added in large numbers so we will see how many new ones are being requested so we can control if necessary
        $FoldersToAddBatchless = New-Object System.Collections.Generic.List[PSObject]
        [hashtable]$newFoldersAdded = @{} ## keep track of what we've issued btch commands to create so we don't duplicate

        foreach ($ExtFolderPath in $ExtFolderPaths.FolderPath) {
            if ( $ExtFolderPath -notin $CUFolders.Path ) {  ##check if folder doesn't already exist
                [string]$pathSoFar = $null
                ## Check each part of the path exists, or will be created, and if not add a task to create it
                ForEach( $pathElement in ($ExtFolderPath.Trim( '\' )).Split( '\' ) ) {
                    [string]$absolutePath = $(if( $pathSoFar ) { Join-Path -Path $pathSoFar -ChildPath $pathElement } else { $pathElement })
                    if( $null -eq $newFoldersAdded[ $absolutePath ] -and $absolutePath -notin $CUFolders.Path  ) ## not already added it to folder creations or already exists
                    {
                        ## there is a bug that causes an error if a folder name being created in a batch already exists at the top level so we workaround it
                        if( $batchCreateFolders )
                        {
                            if ($FoldersToAddCount -ge $maxFolderBatchSize) {  ## we will execute folder batch operations $maxFolderBatchSize at a time
                                Write-Verbose "Generating a new add folder batch"
                                $FolderAddBatches.Add($FoldersToAddBatch)
                                $FoldersToAddCount = 0
                                $FoldersToAddBatch = New-CUBatchUpdate
                            }
                            Add-CUFolder -Name $pathElement -ParentPath $pathSoFar -Batch $FoldersToAddBatch
                        }
                        else ## create folders immediately rather than in batch but first make a list so we can see how many new ones are needed since some may exist already
                        {
                            if( -not $Preview )
                            {
                                $FoldersToAddBatchless.Add( [pscustomobject]@{ PathElement = $pathElement ; PathSoFar = $pathSoFar } )
                            }
                        }
                        $FoldersToAdd.Add("Add-CUFolder -Name `"$pathElement`" -ParentPath `"$pathSoFar`"")
                        $FoldersToAddCount++
                        $newFoldersAdded.Add( $absolutePath , $ExtFolderPath )
                    }
                    $pathSoFar = $absolutePath
                }
            }
        }

        if( $FoldersToAddBatchless -and $FoldersToAddBatchless.Count )
        {
            [int]$folderDelayMilliseconds = 0

            if( $FoldersToAddBatchless.Count -ge $batchCountWarning )
            {
                [string]$logText = "$($FoldersToAddBatchless.Count) folders to add which could cause performance issues"

                if( $force )
                {
                    Write-CULog -Msg $logText -ShowConsole -Type W
                    $folderDelayMilliseconds = $folderCreateDelaySeconds * 1000
                }
                else
                {
                    $errorMessage = "$logText, aborting - use -force to override" 
                    Write-CULog -Msg $errorMessage -ShowConsole -Type E
                    Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$errorMessage"
                    $errorCount++
                    break
                }
            }
            ForEach( $item in $FoldersToAddBatchless )
            {
                Write-Verbose -Message "Creating folder `"$($item.pathElement)`" in `"$($item.pathSoFar)`""
                if( ! ( $folderCreated = Add-CUFolder -Name $item.pathElement -ParentPath $item.pathSoFar ) -or $folderCreated -notmatch "^Folder '$($item.pathElement)' was added successfully$" )
                {
                    Write-CULog -Msg "Failed to create folder `"$($item.pathElement)`" in `"$($item.pathSoFar)`" - $folderCreated" -ShowConsole -Type E
                }
                ## to help avoid central CU service becoming overwhelmed
                if( $folderDelayMilliseconds -gt 0 )
                {
                    Start-Sleep -Milliseconds $folderDelayMilliseconds
                }
            }
        }

        if ($FoldersToAddCount -le $maxFolderBatchSize -and $FoldersToAddCount -ne 0) { $FolderAddBatches.Add($FoldersToAddBatch) }

        # Build computers batch
        $ComputersAddBatches    = New-Object System.Collections.Generic.List[PSObject]
        $ComputersMoveBatches   = New-Object System.Collections.Generic.List[PSObject]
        $ComputersRemoveBatches = New-Object System.Collections.Generic.List[PSObject]
        $ComputersAddBatch      = New-CUBatchUpdate
        $ComputersMoveBatch     = New-CUBatchUpdate
        $ComputersRemoveBatch   = New-CUBatchUpdate
        $ComputersAddCount      = 0
        $ComputersMoveCount     = 0
        $ComputersRemoveCount   = 0

        $ExtComputers = $externalTree.Where{$_.Type -eq "Computer"}
        Write-CULog -Msg  "External Computers Total Count: $($ExtComputers.count)" -ShowConsole -Color Cyan

        #we'll output the statistics at the end -- also helps with debugging
        $MachinesToMove   = New-Object System.Collections.Generic.List[PSObject]
        $MachinesToAdd    = New-Object System.Collections.Generic.List[PSObject]
        $MachinesToRemove = New-Object System.Collections.Generic.List[PSObject]
        
        Write-CULog "Determining Computer Objects to Add or Move" -ShowConsole
        foreach ($ExtComputer in $ExtComputers) {
	        if (($CUComputersHashTable.Contains("$($ExtComputer.Name)"))) {
    	        if ("$($ExtComputer.FolderPath)\" -notlike "$($CUComputersHashTable[$($ExtComputer.name)].Path)\") {
                    if ($ComputersMoveCount -ge $maxBatchSize) {  ## we will execute computer batch operations $maxBatchSize at a time
                        Write-Verbose "Generating a new computer move batch"
                        $ComputersMoveBatches.Add($ComputersMoveBatch)
                        $ComputersMoveCount = 0
                        $ComputersMoveBatch = New-CUBatchUpdate
                    }

        	        Move-CUComputer -Name $ExtComputer.Name -FolderPath "$($ExtComputer.FolderPath)" -Batch $ComputersMoveBatch
                    $MachinesToMove.Add("Move-CUComputer -Name $($ExtComputer.Name) -FolderPath `"$($ExtComputer.FolderPath)`"")
                    $ComputersMoveCount = $ComputersMoveCount+1
    	        }
	        } else {
                if ($ComputersAddCount -ge $maxBatchSize) {  ## we will execute computer batch operations $maxBatchSize at a time
                        Write-Verbose "Generating a new add computer batch"
                        $ComputersAddBatches.Add($ComputersAddBatch)
                        $ComputersAddCount = 0
                        $ComputersAddBatch = New-CUBatchUpdate
                    }
                
    	        try {
                         Add-CUComputer -Domain $ExtComputer.Domain -Name $ExtComputer.Name -FolderPath "$($ExtComputer.FolderPath)" -Batch $ComputersAddBatch @SiteIdParam
                } catch {
                         Write-CULog "Error while attempting to run Add-CUComputer" -ShowConsole -Type E
                         Write-CULog "$($Error[0])"  -ShowConsole -Type E
                }
                if ( ! [string]::IsNullOrEmpty( $SiteIdGUID )) {
                    $MachinesToAdd.Add("Add-CUComputer -Domain $($ExtComputer.Domain) -Name $($ExtComputer.Name) -FolderPath `"$($ExtComputer.FolderPath)`" -SiteId $SiteIdGUID")
                } else {
                    $MachinesToAdd.Add("Add-CUComputer -Domain $($ExtComputer.Domain) -Name $($ExtComputer.Name) -FolderPath `"$($ExtComputer.FolderPath)`"")
                }
                $ComputersAddCount = $ComputersAddCount+1
	        }
        }
        if ($ComputersMoveCount -le $maxBatchSize -and $ComputersMoveCount -ne 0) { $ComputersMoveBatches.Add($ComputersMoveBatch) }
        if ($ComputersAddCount -le $maxBatchSize -and $ComputersAddCount -ne 0)   { $ComputersAddBatches.Add($ComputersAddBatch)   }

        $FoldersToRemoveBatches = New-Object System.Collections.Generic.List[PSObject]
        $FoldersToRemoveBatch   = New-CUBatchUpdate
        $FoldersToRemoveCount   = 0
        #we'll output the statistics at the end -- also helps with debugging
        $FoldersToRemove = New-Object System.Collections.Generic.List[PSObject]
        
        if ($Delete) {
            Write-CULog "Determining Objects to be Removed" -ShowConsole
	        # Build batch for folders which are in ControlUp but not in the external source
<#
            if ($CUFolders.where{ $_.Path -like "$("$CURootFolder")\*" }.count -eq 0) { ## Get CUFolders filtered to targetted sync path
               $CUFolderSyncRoot = $CUFolders.where{$_.Path -like "$("$CURootFolder")"} ## if count is 0 then no subfolders exist
               Write-CULog "Root Target Path : Only Target Folder Exists" -ShowConsole -Verbose
            }
            if ($CUFolders.where{$_.Path -like "$("$CURootFolder")\*"}.count -ge 1) { ## if count is ge 1 then grab all subfolders
                $CUFolderSyncRoot = $CUFolders.where{$_.Path -like "$("$CURootFolder")\*"} 
                Write-CULog "Root Target Path : Subfolders detected" -ShowConsole -Verbose
            }
#>
            [string]$folderRegex = "^$([regex]::Escape( $CURootFolder ))\\.+"
            [array]$CUFolderSyncRoot = @( $CUFolders.Where{ $_.Path -match $folderRegex } )
            if( $CUFolderSyncRoot -and $CUFolderSyncRoot.Count )
            {
                Write-CULog "Root Target Path : $($CUFolderSyncRoot.Count) subfolders detected" -ShowConsole -Verbose
            }
            else
            {
                Write-CULog "Root Target Path : Only Target Folder Exists" -ShowConsole -Verbose
            }
            Write-CULog "Determining Folder Objects to be Removed" -ShowConsole
	        foreach ($CUFolder in $($CUFolderSyncRoot.Path)) {
                $folderRegex = "$([regex]::Escape( $CUFolder ))"
                ## need to test if the whole path matches or it's a sub folder (so "Folder 1" won't match "Folder 12")
                if( $ExtFolderPaths.Where( { $_.FolderPath -match "^$folderRegex$" -or $_.FolderPath -match "^$folderRegex\\" } ).Count -eq 0 -and $CUFolder -ne $CURootFolder ) {
                ## can't use a simple -notin as path may be missing but there may be child paths of it - GRL
    	        ##if (($CUFolder -notin $ExtFolderPaths.FolderPath) -and ($CUFolder -ne $("$CURootFolder"))) { #prevents excluding the root folder
                    if ($Delete) {
                        if ($FoldersToRemoveCount -ge $maxFolderBatchSize) {  ## we will execute computer batch operations $maxBatchSize at a time
                            Write-Verbose "Generating a new remove folder batch"
                            $FoldersToRemoveBatches.Add($FoldersToRemoveBatch)
                            $FoldersToRemoveCount = 0
                            $FoldersToRemoveBatch = New-CUBatchUpdate
                        }
        	            Remove-CUFolder -FolderPath "$CUFolder" -Force -Batch $FoldersToRemoveBatch
                        $FoldersToRemove.Add("Remove-CUFolder -FolderPath `"$CUFolder`" -Force")
                        $FoldersToRemoveCount = $FoldersToRemoveCount+1
                    }
    	        }
	        }

            Write-CULog "Determining Computer Objects to be Removed" -ShowConsole
	        # Build batch for computers which are in ControlUp but not in the external source
            [string]$curootFolderAllLower = $CURootFolder.ToLower()
	        foreach ($CUComputer in $CUComputers.Where{$_.path.startsWith( $curootFolderAllLower ) }) { #hey! StartsWith is case sensitive..  at least we return path in lowercase.
                ##if ($($ExtFolderPaths.FolderPath) -contains $CUComputer.path) {
    	            if (-not $ExtTreeHashTable[ $CUComputer.name ] ) {
                        if ($Delete) {
                            if ($FoldersToRemoveCount -ge $maxFolderBatchSize) {  ## we will execute computer batch operations $maxBatchSize at a time
                                Write-Verbose "Generating a new remove computer batch"
                                $ComputersRemoveBatches.Add($ComputersRemoveBatch)
                                $ComputersRemoveCount = 0
                                $ComputersRemoveBatch = New-CUBatchUpdate
                            }
        	                Remove-CUComputer -Name $($CUComputer.Name) -Force -Batch $ComputersRemoveBatch
                            $MachinesToRemove.Add("Remove-CUComputer -Name $($CUComputer.Name) -Force")
                            $ComputersRemoveCount = $ComputersRemoveCount+1
                        }
                    }
    	        ##}
	        }
        }
        if ($FoldersToRemoveCount -le $maxFolderBatchSize -and $FoldersToRemoveCount -ne 0) { $FoldersToRemoveBatches.Add($FoldersToRemoveBatch)   }
        if ($ComputersRemoveCount -le $maxBatchSize -and $ComputersRemoveCount -ne 0)       { $ComputersRemoveBatches.Add($ComputersRemoveBatch)   }

        #endregion

        Write-CULog -Msg "Folders to Add     : $($FoldersToAdd.Count)" -ShowConsole -Color White 
        Write-CULog -Msg "Folders to Add Batches     : $($FolderAddBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($FoldersToAdd.Count) -ge 25) {
            foreach ($obj in $FoldersToAdd) {
                Write-CULog -Msg "$obj" -SubMsg
            }
        } else {
            foreach ($obj in $FoldersToAdd) {
                Write-CULog -Msg "$obj" -ShowConsole -Color Green -SubMsg
            }
        }

        Write-CULog -Msg "Folders to Remove  : $($FoldersToRemove.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Folders to Remove Batches  : $($FoldersToRemoveBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($FoldersToRemove.Count) -ge 25) {
            foreach ($obj in $FoldersToRemove) {
                Write-CULog -Msg "$obj" -SubMsg
            }
        } else {
            foreach ($obj in $FoldersToRemove) {
                Write-CULog -Msg "$obj" -ShowConsole -Color DarkYellow -SubMsg
            }
        }

        Write-CULog -Msg "Computers to Add   : $($MachinesToAdd.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Computers to Add Batches   : $($ComputersAddBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($MachinesToAdd.Count) -ge 25) {
            foreach ($obj in $MachinesToAdd) {
                Write-CULog -Msg "$obj" -SubMsg
            } 
        } else {
            foreach ($obj in $MachinesToAdd) {
                Write-CULog -Msg "$obj" -ShowConsole -Color Green -SubMsg
            }
        }

        Write-CULog -Msg "Computers to Move  : $($MachinesToMove.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Computers to Move Batches  : $($ComputersMoveBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($MachinesToMove.Count) -ge 25) {
            foreach ($obj in $MachinesToMove) {
                Write-CULog -Msg "$obj" -SubMsg
            }
        } else {
            foreach ($obj in $MachinesToMove) {
                Write-CULog -Msg "$obj" -ShowConsole -Color DarkYellow -SubMsg
            }
        }

        Write-CULog -Msg "Computers to Remove: $($MachinesToRemove.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Computers to Remove Batches: $($ComputersRemoveBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($MachinesToRemove.Count -ge 25)) {
            foreach ($obj in $MachinesToRemove) {
                Write-CULog -Msg "$obj" -SubMsg
            }
        } else {
            foreach ($obj in $MachinesToRemove) {
                Write-CULog -Msg "$obj" -ShowConsole -Color DarkYellow -SubMsg
            }
        }
            
        $endTime = Get-Date

        Write-CULog -Msg "Build-CUTree took: $($(New-TimeSpan -Start $startTime -End $endTime).Seconds) Seconds." -ShowConsole -Color White
        Write-CULog -Msg "Committing Changes:" -ShowConsole -Color DarkYellow
        if ($ComputersRemoveBatches.Count -gt 0) { $errorCount += Execute-PublishCUUpdates -BatchObject $ComputersRemoveBatches -Message "Executing Computer Object Removal" }
        if ($FoldersToRemoveBatches.Count -gt 0) { $errorCount += Execute-PublishCUUpdates -BatchObject $FoldersToRemoveBatches -Message "Executing Folder Object Removal"   }
        if ($FolderAddBatches.Count -gt 0 -and $batchCreateFolders )       { $errorCount += Execute-PublishCUUpdates -BatchObject $FolderAddBatches -Message "Executing Folder Object Adds"            }
        if ($ComputersAddBatches.Count -gt 0)    { $errorCount += Execute-PublishCUUpdates -BatchObject $ComputersAddBatches -Message "Executing Computer Object Adds"       }
        if ($ComputersMoveBatches.Count -gt 0)   { $errorCount += Execute-PublishCUUpdates -BatchObject $ComputersMoveBatches -Message "Executing Computer Object Moves"     }

        Write-CULog -Msg "Returning $errorCount to caller"

        return $errorCount
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
    if ($Global:LogFile) {
        Write-Output "$date | $LogType | $Msg"  | Out-file $($Global:LogFile) -Append
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

Function Send-EmailAlert
{
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory=$false, HelpMessage='Smtp server to send emails from' )]
	    [string] $SmtpServer ,

        [Parameter(Mandatory=$false, HelpMessage='Email address to send email from' )]
	    [string] $from ,

        [Parameter(Mandatory=$false, HelpMessage='Email addresses to send email to' )]
	    [string[]] $to ,

        [Parameter(Mandatory=$false, HelpMessage='Body of email' )]
	    [string] $body ,

        [Parameter(Mandatory=$false, HelpMessage='Subject of email' )]
	    [string] $subject ,

        [Parameter(Mandatory=$false, HelpMessage='Use SSL to send email alert' )]
	    [switch] $useSSL
    )

    if( [string]::IsNullOrEmpty( $SmtpServer ) -or [string]::IsNullOrEmpty( $to ) )
    {
        return $null ## don't check if set at caller end to make code sleaker
    }

    [int]$port = 25
    [string[]]$serverParts = @( $SmtpServer -split ':' )
    if( $serverParts.Count -gt 1 )
    {
        $port = $serverParts[-1]
    }

    if( $to -and $to.Count -eq 1 -and $to[0].IndexOf( ',' ) -ge 0 )
    {
        $to = @( $to -split ',' )
    }

    if( [string]::IsNullOrEmpty( $from ) )
    {
        $from = "$env:COMPUTERNAME@$env:USERDNSDOMAIN"
    }

    Send-MailMessage -SmtpServer $serverParts[0] -Port $port -UseSsl:$useSSL -Body $body -Subject $subject -From $from -to $to
}
