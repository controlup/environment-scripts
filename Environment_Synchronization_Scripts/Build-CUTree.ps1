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

    .PARAMETER ExternalTree <Object>
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

    .PARAMETER CURootFolder <string>
        The folder in the ControlUp console that will be the root for the ExternalTree. Do NOT specify the organization name in this path, just the folder path
        underneath. For instance, in ControlUp if I right-click the folder I want to use as my root, and select "Properties" the "path:" will say something like:
        "a.c.m.e.\vdi_and_sbc\wvd".  In this example, "a.c.m.e." is the organization name so ignore it and just enter the folders underneath, which in this
        example is "vdi_and_sbc\wvd".

    .PARAMETER Delete <Switch>
        Enables removal of excess objects in the sync directory. If a machine or folder object is found in the ControlUp path but not in the ExternalTree object
        the object will be marked for removal. If you do not use this parameter than only Add or Move operations will occur.

    .PARAMETER Preview <switch>
        Generates a preview showing the actions this script wants to take. If more than 25 operations is going to be executed the script will just
        return the number of operations. To see the individual operations use the PreviewOutputPath parameter. See that paramater help for more information.
        The Preview parameter will output do the operation calculation and return the expected operations as in this sample:
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

    .PARAMETER LogFile <path to log file>
        Specifies that all output will be saved to a log file. Individual operations will also be saved to the log file. The operations are saved in
        such a way that you should be able to copy-paste them into a powershell prompt that has the ControlUp Powershell modules loaded and they should
        be executable.  Use this for testing individual operations to validate it will work as you expect.

    .PARAMETER SiteId <string>
        An optional parameter to specify which site you want the machine object assigned. By default, the site ID is "Default". Enter the name of the site
        to assign the object

    .EXAMPLE
        Build-CUTree -ExternalTree $WVDEnvironment -CURootFolder "VDI_and_SBC\WVD" -Preview -Delete -PreviewOutputPath C:\temp\sync.log
            Executes a logged preview of what sync'ing the WVDEnvironment object to the ControlUp folder VDI_and_SBC\WVD with what object removals would look like.

    .EXAMPLE
        Build-CUTree -ExternalTree $WVDEnvironment -CURootFolder "VDI_and_SBC\WVD" -Delete
            Executes a sync of the $WVDEnvironment object to the ControlUp folder "VDI_and_SBC\WVD" with object removal enabled.

    .CONTEXT
    .MODIFICATION_HISTORY
    .LINK
        https://www.controlup.com

    .COMPONENT
    .NOTES
	    Runs on a ControlUp Monitor computer
	    Connects to an external source, retrieves the folder structure to synchronize
	    Adds to ControlUp folder structure all folders and computers from the external source
	    Moves folders and computers which exist in locations that differ from the external source
	    Optionally, removes folders and computers which do not exist in the external source
#>

#region Bind input parameters
function Build-CUTree {
    [CmdletBinding()]
    Param
    (

	    [Parameter(
    	    Position=1,
    	    Mandatory=$true,
    	    HelpMessage='Object to build tree within ControlUp'
	    )]
	    [PSObject] $ExternalTree,

	    [Parameter(
    	    Mandatory=$false,
    	    HelpMessage='ControlUp root folder to sync'
	    )]
	    [string] $CURootFolder,

 	    [Parameter(
    	    Mandatory=$false,
    	    HelpMessage='Delete CU objects which are not in the external source'
	    )]
	    [switch] $Delete,

        [Parameter(
            Mandatory=$false,
            HelpMessage='Generate a report of the actions to be executed'
        )]
        [switch]$Preview,

        [Parameter(
    	    Mandatory=$false,
    	    HelpMessage='Save a log file'
	    )]
	    [string] $LogFile,

        [Parameter(
    	    Mandatory=$false,
    	    HelpMessage='ControlUp Site Id to assign the machine object'
	    )]
	    [string] $SiteId,

        [Parameter(
    	    Mandatory=$false,
    	    HelpMessage='Debug CU Machine Environment Objects'
	    )]
	    [Object] $DebugCUMachineEnvironment,

        [Parameter(
    	    Mandatory=$false,
    	    HelpMessage='Debug CU Folder Environment Object'
	    )]
	    [switch] $DebugCUFolderEnvironment

    )
    #endregion


    Begin {

    #This variable sets the maximum computer batch size to apply the changes in ControlUp. It is not recommended making it bigger than 1000
    $maxBatchSize = 1000
    #This variable sets the maximum batch size to apply the changes in ControlUp. It is not recommended making it bigger than 100
    $maxFolderBatchSize = 100

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

        function Execute-PublishCUUpdates {
            Param(
	            [Parameter(Mandatory = $True)][Object]$BatchObject,
	            [Parameter(Mandatory = $True)][string]$Message
            )
            $batchCount = 1
            foreach ($batch in $BatchObject) {
                Write-CULog -Msg "$Message. Batch $batchCount/$($BatchObject.count)" -ShowConsole -Color DarkYellow -SubMsg
                if (-not($preview)) {
                    $PublishTime = Measure-Command { Publish-CUUpdates $batch }
                    Write-CULog -Msg "Execution Time: $($PublishTime.TotalSeconds)" -ShowConsole -Color Green -SubMsg
                } else {
                    Write-CULog -Msg "Execution Time: PREVIEW MODE" -ShowConsole -Color Green -SubMsg
                }
                $batchCount = $batchCount+1
            }
        }
        
        function Test-CUFolderPath {
            Param(
                [parameter(Mandatory = $true,
                HelpMessage = "Specifies a path to be tested. The value of the Path parameter is case insensitive and used exactly as it is typed. No characters are interpreted as wildcard characters.")]
                [string]$Path
            )
            if ($path.EndsWith("\")) { $Path = $Path.TrimEnd("\") }  # remove last character if it ends in a backslash
            foreach ($folder in $CUFolders) {
                if (($folder.Path) -eq "$Path") {return $true}
            }
            return $false
        }

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
    }

    Process {
        <#
        ## For debugging uncomment
        $ErrorActionPreference = 'Stop'
        $VerbosePreference = 'SilentlyContinue'
        $DebugPreference = 'SilentlyContinue'
        Set-StrictMode -Version Latest
        #>

        $startTime = Get-Date

        #region Load ControlUp PS Module
        try {
            ## Check CU monitor is installed and at least minimum required version
            [string]$cuMonitor = 'ControlUp Monitor'
            [version]$minimumCUmonitorVersion = '8.1.5.600'
            if( ! ( $installKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -Name DisplayName -ErrorAction SilentlyContinue| Where-Object DisplayName -eq $cuMonitor ) )
            {
                Write-Error -Message "$cuMonitor does not appear to be installed"
                break
            }
	        # Importing the latest ControlUp PowerShell Module
	        if( ! ( $pathtomodule = (Get-ChildItem "$($env:ProgramFiles)\Smart-X\ControlUpMonitor\*\ControlUp.PowerShell.User.dll" -Recurse -Force | Sort-Object -Property VersionInfo.FileVersion -Descending) | Select-Object -First 1 ) )
            {
                Write-Error -Message "Unable to find ControlUp.PowerShell.User.dll"
                break
            }
            ## dll version is 1.0.0.0 so we check cuMonitor.exe
            if( $cuMonitorExeProperties = Get-ChildItem -Path (Join-Path -Path $pathtomodule.DirectoryName -ChildPath 'cuMonitor.exe') -Force -ErrorAction SilentlyContinue )
            {
                if( $cuMonitorExeProperties.VersionInfo.FileVersion -lt $minimumCUmonitorVersion )
                {
                    Write-Warning -Message "Found version $($cuMonitorExeProperties.VersionInfo.FileVersion) of cuMonitor.exe but need at least $($minimumCUmonitorVersion.ToString())"
                }
            }
            else
            {
                Write-Warning -Message "Unable to find cuMonitor.exe in folder `"$($pathtomodule.DirectoryName)`""
            }
	        if( ! ( Import-Module $pathtomodule -PassThru ) )
            {
                break
            }
            if( ! ( Get-Command -Name 'Get-CUFolders' -ErrorAction SilentlyContinue ) )
            {
                Write-Error -Message "Loaded CU Monitor PowerShell module but unable to find cmdlet Get-CUFolders"
            }
        }
        catch {
            Write-CULog -Msg 'The required ControlUp PowerShell module was not found or could not be loaded. Please make sure this is a ControlUp Monitor machine.' -ShowConsole -Type E
        }
        #endregion


        #region validate SiteId parameter
        [hashtable] $SiteIdParam = @{}
        if ($PSBoundParameters.ContainsKey("SiteId")) {
            Write-CULog -Msg "Assigning resources to specific site: $SiteId" -ShowConsole
            $Sites = Get-CUSites
            $SiteIdGUID = ($Sites.Where{$_.Name -eq $SiteId}).Id
            Write-CULog -Msg "SiteId GUID: $SiteIdGUID" -ShowConsole -SubMsg
            $SiteIdParam.Add( 'SiteId' , $SiteIdGUID )
        }


        #region Retrieve ControlUp folder structure
        if (-not($DebugCUMachineEnvironment)) {
            try {
                $CUComputers = Get-CUComputers # add a filter on path so only computers within the $rootfolder are used
            } catch {
                Write-Error "Unable to get computers from ControlUp"
                break
            }
        } else {
            Write-Debug "Number of objects in DebugCUMachineEnvironment: $($DebugCUMachineEnvironment.count)"
            if ($($DebugCUMachineEnvironment.count) -eq 2) {
                foreach ($envObjects in $DebugCUMachineEnvironment) {
                    if  ($($envObjects  | gm).TypeName[0] -eq "Create-CrazyCUEnvironment.CUComputerObject") {
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
                Write-Error "Unable to get folders from ControlUp"
                break
            }
        } else {
            Write-Debug "Number of folder objects in DebugCUMachineEnvironment: $($DebugCUMachineEnvironment.count)"
            if ($($DebugCUMachineEnvironment.count) -eq 2) {
                foreach ($envObjects in $DebugCUMachineEnvironment) {
                    if  ($($envObjects  | gm).TypeName[0] -eq "Create-CrazyCUEnvironment.CUFolderObject") {
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

        #testing for folder structure:
        Write-CULog -Msg "Preparing root path: $("$OrganizationName\$CURootFolder")" -ShowConsole
        $builtPath = $false
        if (-not(Test-CUFolderPath -Path $("$OrganizationName\$CURootFolder"))) {
            [string]$folderTree = ""
            foreach ($folder in "$OrganizationName\$CURootFolder".split("\")) {
                Write-CULog -Msg "Checking for folder : $folder" -Verbose
                if ($folder -eq $OrganizationName) {
                    $folderTree += "$folder"
                } else {
                    $folderTree += "\$folder"
                }
                if (Test-CUFolderPath -Path $folderTree) {
                    Write-CULog -Msg "Path Found          : $folderTree" -Verbose
                } else {
                    Write-CULog -Msg "Path NOT found      : $folderTree" -Verbose
                    $LastBackslashPosition = $folderTree.LastIndexOf("\")
	                if ($LastBackslashPosition -gt 0) {
                        Write-CULog "Adding Folder $folder" -ShowConsole -Color Green
                        Write-CULog "Add-CUFolder -Name $folder -ParentPath $($folderTree.Substring(0,$LastBackslashPosition))" -ShowConsole -SubMsg
                        $builtPath = $true
                        Add-CUFolder -Name $folder -ParentPath $folderTree.Substring(0,$LastBackslashPosition)
                        
                    }
                }
            }
        }

        #if we built the path we need to update our variable:
        if ($builtPath) {
            $attempts = 0
            Write-CULog -Msg "Updating CUFolders variable" -Verbose
            do {
                Write-CULog -Msg "Checking for $("$OrganizationName\$CURootFolder")" -Verbose
                sleep 10 ## Need to sleep some amount of time to allow the update on the common config or else the folder changes aren't picked up.
                $CUFolders = Get-CUFolders
                $attempts = $attempts+1
            } While ((-not(Test-CUFolderPath -Path $("$OrganizationName\$CURootFolder"))) -or $attempts -ge 10)
        }

        #region Prepare items for synchronization
        #replace FolderPath in ExternalTree object with the local ControlUp Path:
        foreach ($obj in $externalTree) {
            $objectFolderPath = $obj.FolderPath
            $obj.FolderPath = "$("$OrganizationName\$CURootFolder")\$objectFolderPath"
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
        $FoldersToAdd       = New-Object System.Collections.Generic.List[PSObject]
        
        foreach ($ExtFolderPath in $ExtFolderPaths.FolderPath) {
            if ("$ExtFolderPath" -notin $($CUFolders.Path)) {  ##check if folder doesn't already exist
	            $LastBackslashPosition = $ExtFolderPath.LastIndexOf("\")
	            if ($LastBackslashPosition -gt 0) {
                    if ($FoldersToAddCount -ge $maxFolderBatchSize) {  ## we will execute folder batch operations $maxFolderBatchSize at a time
                        Write-Verbose "Generating a new add folder batch"
                        $FolderAddBatches.Add($FoldersToAddBatch)
                        $FoldersToAddCount = 0
                        $FoldersToAddBatch = New-CUBatchUpdate
                    }
                    Add-CUFolder -Name $ExtFolderPath.Substring($LastBackslashPosition+1) -ParentPath $ExtFolderPath.Substring(0,$LastBackslashPosition) -Batch $FoldersToAddBatch
                    $FoldersToAdd.Add("Add-CUFolder -Name $($ExtFolderPath.Substring($LastBackslashPosition+1)) -ParentPath `"$($ExtFolderPath.Substring(0,$LastBackslashPosition))`"")
                    $FoldersToAddCount = $FoldersToAddCount+1
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
                if ($PSBoundParameters.ContainsKey("SiteId")) {
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
            if ($CUFolders.where{$_.Path -like "$("$OrganizationName\$CURootFolder")\*"}.count -eq 0) { ## Get CUFolders filtered to targetted sync path
               $CUFolderSyncRoot = $CUFolders.where{$_.Path -like "$("$OrganizationName\$CURootFolder")"} ## if count is 0 then no subfolders exist
               Write-CULog "Root Target Path : Only Target Folder Exists" -ShowConsole -Verbose
            }
            if ($CUFolders.where{$_.Path -like "$("$OrganizationName\$CURootFolder")\*"}.count -ge 1) { ## if count is ge 1 then grab all subfolders
                $CUFolderSyncRoot = $CUFolders.where{$_.Path -like "$("$OrganizationName\$CURootFolder")\*"} 
                Write-CULog "Root Target Path : Subfolders detected" -ShowConsole -Verbose
            }
            Write-CULog "Determining Folder Objects to be Removed" -ShowConsole
	        foreach ($CUFolder in $($CUFolderSyncRoot.Path)) {
    	        if (($CUFolder -notin $ExtFolderPaths.FolderPath) -and ($CUFolder -ne $("$OrganizationName\$CURootFolder"))) { #prevents excluding the root folder
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
	        foreach ($CUComputer in $CUComputers.where{$_.path.startsWith(("$OrganizationName\$CURootFolder").ToLower())}) { #hey! StartsWith is case sensitive..  at least we return path in lowercase.
                if ($($ExtFolderPaths.FolderPath) -contains $CUComputer.path) {
    	            if (-not($ExtTreeHashTable.Contains("$($CUComputer.name)"))) {
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
    	        }
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
        if ($ComputersRemoveBatches.Count -gt 0) { Execute-PublishCUUpdates -BatchObject $ComputersRemoveBatches -Message "Executing Computer Object Removal" }
        if ($FoldersToRemoveBatches.Count -gt 0) { Execute-PublishCUUpdates -BatchObject $FoldersToRemoveBatches -Message "Executing Folder Object Removal"   }
        if ($FolderAddBatches.Count -gt 0)       { Execute-PublishCUUpdates -BatchObject $FolderAddBatches -Message "Executing Folder Object Adds"            }
        if ($ComputersAddBatches.Count -gt 0)    { Execute-PublishCUUpdates -BatchObject $ComputersAddBatches -Message "Executing Computer Object Adds"       }
        if ($ComputersMoveBatches.Count -gt 0)   { Execute-PublishCUUpdates -BatchObject $ComputersMoveBatches -Message "Executing Computer Object Moves"     }

    }
    
}
