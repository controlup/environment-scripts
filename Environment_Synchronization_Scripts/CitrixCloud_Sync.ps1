[CmdletBinding()]
Param(
	[Parameter(Mandatory=$false, HelpMessage='Enter a ControlUp subfolder to save your DeliveryGroup tree to')]
	[ValidateNotNullOrEmpty()]
	[string] $folderPath,

	[Parameter(Mandatory=$false, HelpMessage='Enter the Citrix Cloud Client ID')]
	[ValidateNotNullOrEmpty()]
	[string] $clientID,

	[Parameter(Mandatory=$false, HelpMessage='Enter the Citrix Cloud Secret Key')]
	[ValidateNotNullOrEmpty()]
	[string] $secretKey,

	[Parameter(Mandatory=$false, HelpMessage='Enter the Citrix Cloud Customer ID')]
	[ValidateNotNullOrEmpty()]
	[string] $cloudEnvironmentName,

	[Parameter(Mandatory=$false, HelpMessage='Domain to force on the CU object')]
	[ValidateNotNullOrEmpty()]
	[string] $forceDomain,

	[Parameter(Mandatory=$false, HelpMessage='Preview the changes' )]
	[ValidateNotNullOrEmpty()]
	[switch] $Preview,

	[Parameter(Mandatory=$false, HelpMessage='Execute removal operations. When combined with preview it will only display the proposed changes')]
	[ValidateNotNullOrEmpty()]
	[switch] $Delete,

	[Parameter(Mandatory=$false, HelpMessage='Enter a path to generate a log file of the proposed changes')]
	[ValidateNotNullOrEmpty()]
	[string] $LogFile,

	[Parameter(Mandatory=$false, HelpMessage='Enter a ControlUp Site Name' )]
	[ValidateNotNullOrEmpty()]
	[string] $Site ,

	[Parameter(Mandatory=$false, HelpMessage='Create folders in batches rather than individually')]
	[switch] $batchCreateFolders ,

	[Parameter(Mandatory=$false, HelpMessage='Force folder creation if number exceeds safe limit')]
	[switch] $force
) 
# PS Module Import
Get-Item "$((get-childitem 'C:\Program Files\Smart-X\ControlUpMonitor\')[-1].fullName)\*powershell*.dll"|import-module

# Config Items
$cuOrg = (Get-CUFolders |?{$_.FolderType -eq "RootFolder"}).Name
$cuDGFolder =""
$domainOverride = ""
$batchCreateFolders = $true

#citrixCloud ClientID and SecretKey
$id=''
$secret=''
$cloudEnvironment = ''

#Override Pesets with commandline args
if($clientID){$id = $clientID}
if($secretKey){$secret = $secretKey}
if($cloudEnvironmentName){$customerID = $cloudEnvironmentName}
if($folderPath){$cuDGFolder = $folderPath}
if($forceDomain){$domainOverride = $forceDomain}
$Environment = New-Object -TypeName System.Collections.Generic.List[PSObject]

# Setup TLS Setup
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor
[Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

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

# Setup TLS Setup
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor
[Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

$body = @{grant_type = "client_credentials";client_id = $id;client_secret = $secret}

$token = ((Invoke-webrequest "https://api-us.cloud.com/cctrustoauth2/$customerID/tokens/clients" -Method 'POST' -body $body).content|convertfrom-json).access_token

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
$h = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$h.Add("Authorization", "CWSAuth Bearer $token")
$h.Add("Citrix-CustomerId", $customerID)
$wr = Invoke-RestMethod "https://api-us.cloud.com/cvad/manage/Me " -Method 'GET' -Headers $h -Body $body

$cId = $wr.customers.id
$csId = $wr.customers.sites.id
$csName = $wr.customers.sites.name

$h = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$h.Add("Authorization", "CWSAuth Bearer $token")
$h.Add("Citrix-CustomerId", $customerID)
$h.Add("Citrix-InstanceId", $csId)

$mDGList = Invoke-RestMethod "https://api-us.cloud.com/cvad/manage/DeliveryGroups" -Method 'GET' -Headers $h -Body $body
	$mDGList.items.name|%{
			$af++
			$folderName = $_
			$Environment.Add([ControlUpObject]::new($folderName,"$folderName","Folder",$null,$null,$null))
	}
$machines = @()
$body = '{"SearchFilters": [{"Property": "FaultState","Value": "None","Operator": "Equals"}]}'
$m = Invoke-RestMethod "https://api-us.cloud.com/cvad/manage/Machines/`$search?limit=1000&async=false" -Method 'POST' -Headers $h -contentType "application/json" -body $body
$machines += $m.items
if($m.ContinuationToken){
	do{
		if($m.ContinuationToken){
			$m = invoke-restmethod "https://api-us.cloud.com/cvad/manage/Machines/`$search?limit=1000&async=false&continuationToken=$($m.ContinuationToken)"  -Method 'POST' -Headers $h -contentType "application/json" -body $body
			$machines += $m.items
		}
	}until(!$m.ContinuationToken)
}

$cumList = @()
$cmList = (get-cucomputers).name
foreach ($machine in $machines){
	$dns = $machine.dnsname.toLower()
	$dg = $machine.deliverygroup.name
	$name = $machine.name.split('\')[1]
	$domain = if($domainOverride){$domainOverride}else{$machine.name.split('\')[0]}
		$am++
		$Environment.Add(([ControlUpObject]::new( $name , "$dg","Computer", $Domain ,"Citrix Cloud Machine", $dns )))
}

function Build-CUTree {
    [CmdletBinding()]
    Param(
	    [Parameter(Mandatory=$true,HelpMessage='Object to build tree within ControlUp')]
	    [PSObject] $ExternalTree,
	    [Parameter(Mandatory=$false,HelpMessage='ControlUp root folder to sync')]
	    [string] $CURootFolder,
 	    [Parameter(Mandatory=$false, HelpMessage='Delete CU objects which are not in the external source')]
	    [switch] $Delete,
        [Parameter(Mandatory=$false, HelpMessage='Generate a report of the actions to be executed')]
        [switch]$Preview,
        [Parameter(Mandatory=$false, HelpMessage='Save a log file')]
	    [string] $LogFile,
        [Parameter(Mandatory=$false, HelpMessage='ControlUp Site name to assign the machine object to')]
	    [string] $SiteName,
        [Parameter(Mandatory=$false, HelpMessage='Debug CU Machine Environment Objects')]
	    [Object] $DebugCUMachineEnvironment,
        [Parameter(Mandatory=$false, HelpMessage='Debug CU Folder Environment Object')]
	    [switch] $DebugCUFolderEnvironment ,
        [Parameter(Mandatory=$false, HelpMessage='Create folders in batches rather than individually')]
	    [switch] $batchCreateFolders ,
        [Parameter(Mandatory=$false, HelpMessage='Number of folders to be created that generates warning and requires -force')]
        [int] $batchCountWarning = 100 ,
        [Parameter(Mandatory=$false, HelpMessage='Force creation of folders if -batchCountWarning size exceeded')]
        [switch] $force ,
        [Parameter(Mandatory=$false, HelpMessage='Smtp server to send alert emails from')]
	    [string] $SmtpServer ,
        [Parameter(Mandatory=$false, HelpMessage='Email address to send alert email from')]
	    [string] $emailFrom ,
        [Parameter(Mandatory=$false, HelpMessage='Email addresses to send alert email to')]
	    [string[]] $emailTo ,
        [Parameter(Mandatory=$false, HelpMessage='Use SSL to send email alert')]
	    [switch] $emailUseSSL ,
        [Parameter(Mandatory=$false, HelpMessage='Delay between each folder creation when count exceeds -batchCountWarning')]
        [double] $folderCreateDelaySeconds = 0.5
   )

        #This variable sets the maximum computer batch size to apply the changes in ControlUp. It is not recommended making it bigger than 1000
        $maxBatchSize = 1000
        #This variable sets the maximum batch size to apply the changes in ControlUp. It is not recommended making it bigger than 100
        $maxFolderBatchSize = 100
        [int]$errorCount = 0
        [array]$stack = @(Get-PSCallStack)
        [string]$callingScript = $stack.Where({ $_.ScriptName -ne $stack[0].ScriptName }) | Select-Object -First 1 -ExpandProperty ScriptName
        if(!$callingScript -and !($callingScript = $stack | Select-Object -First 1 -ExpandProperty ScriptName)){$callingScript = $stack[-1].Position}

        function Execute-PublishCUUpdates {
            Param(
	            [Parameter(Mandatory = $True)][Object]$BatchObject,
	            [Parameter(Mandatory = $True)][string]$Message
           )
            [int]$returnCode = 0
            [int]$batchCount = 0
            foreach ($batch in $BatchObject){
                $batchCount++
                Write-CULog -Msg "$Message. Batch $batchCount/$($BatchObject.count)" -ShowConsole -Color DarkYellow -SubMsg
                if (-not($preview)){
                    [datetime]$timeBefore = [datetime]::Now
                    $result = Publish-CUUpdates -Batch $batch 
                    [datetime]$timeAfter = [datetime]::Now
                    [array]$results = @(Show-CUBatchResult -Batch $batch)
                    [array]$failures = @($results.Where({$_.IsSuccess -eq $false})) ## -and $_.ErrorDescription -notmatch 'Folder with the same name already exists' }))

                    Write-CULog -Msg "Execution Time: $(($timeAfter - $timeBefore).TotalSeconds) seconds" -ShowConsole -Color Green -SubMsg
                    Write-CULog -Msg "Result: $result" -ShowConsole -Color Green -SubMsg
                    Write-CULog -Msg "Failures: $($failures.Count) / $($results.Count)" -ShowConsole -Color Green -SubMsg

                    if($failures -and $failures.Count -gt 0){
                        $returnCode += $failures.Count
                        foreach($failure in $failures){Write-CULog -Msg "Action $($failure.ActionName) on `"$($failure.Subject)`" gave error $($failure.ErrorDescription) ($($failure.ErrorCode))" -ShowConsole -Type E}
                    }
                }else{Write-CULog -Msg "Execution Time: PREVIEW MODE" -ShowConsole -Color Green -SubMsg}
            }
        }
        

        #attempt to setup the log file
        if ($PSBoundParameters.ContainsKey("LogFile")){
            $Global:LogFile = $PSBoundParameters.LogFile
            Write-Host "Saving Output to: $Global:LogFile"
            if (-not(Test-Path $($PSBoundParameters.LogFile))){
                Write-CULog -Msg "Creating Log File" #Attempt to create the file
                if (-not(Test-Path $($PSBoundParameters.LogFile))){Write-Error "Unable to create the report file" -ErrorAction Stop}
            }else{Write-CULog -Msg "Beginning Synchronization"}
            Write-CULog -Msg "Detected the following parameters:"
            foreach($psbp in $PSBoundParameters.GetEnumerator()){
                if ($psbp.Key -like "ExternalTree" -or $psbp.Key -like "DebugCUMachineEnvironment"){
                    Write-CULog -Msg $("Parameter={0} Value={1}" -f $psbp.Key,$psbp.Value.count)
                }else{Write-CULog -Msg $("Parameter={0} Value={1}" -f $psbp.Key,$psbp.Value)}
            }
        }else{$Global:LogFile = $false}

        if(!$PSBoundParameters['folderCreateDelaySeconds' ] -and $env:CU_delay){
            $folderCreateDelaySeconds = $env:CU_delay
        }

        $startTime = Get-Date
        [string]$errorMessage = $null

        #region Load ControlUp PS Module
        try{
            ## Check CU monitor is installed and at least minimum required version
            [string]$cuMonitor = 'ControlUp Monitor'
            [string]$cuDll = 'ControlUp.PowerShell.User.dll'
            [string]$cuMonitorProcessName = 'CUmonitor'
            [version]$minimumCUmonitorVersion = '8.1.5.600'
            if(!($installKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -Name DisplayName -ErrorAction SilentlyContinue| Where-Object DisplayName -eq $cuMonitor)){
                Write-CULog -ShowConsole -Type W -Msg "$cuMonitor does not appear to be installed"
            }
            ## when running via scheduled task we do not have sufficient rights to query services
            if(!($cuMonitorProcess = Get-Process -Name $cuMonitorProcessName -ErrorAction SilentlyContinue)){
                Write-CULog -ShowConsole -Type W -Msg "Unable to find process $cuMonitorProcessName for $cuMonitor service" ## pid $($cuMonitorService.ProcessId)"
            }else{
                [string]$message =  "$cuMonitor service running as pid $($cuMonitorProcess.Id)"
                ## if not running as admin/elevated then won't be able to get start time
                if($cuMonitorProcess.StartTime){
                    $message += ", started at $(Get-Date -Date $cuMonitorProcess.StartTime -Format G)"
                }
                Write-CULog -Msg $message
            }

	        # Importing the latest ControlUp PowerShell Module - need to find path for dll which will be where cumonitor is running from. Don't use Get-Process as may not be elevated so would fail to get path to exe and win32_service fails as scheduled task with access denied
            if(!($cuMonitorServicePath = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\cuMonitor' -Name ImagePath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ImagePath))){
                Throw "$cuMonitor service path not found in registry"
            }elseif(!($cuMonitorProperties = Get-ItemProperty -Path $cuMonitorServicePath.Trim('"') -ErrorAction SilentlyContinue)){
                Throw  "Unable to find CUmonitor service at $cuMonitorServicePath"
            }elseif($cuMonitorProperties.VersionInfo.FileVersion -lt $minimumCUmonitorVersion){
                Throw "Found version $($cuMonitorProperties.VersionInfo.FileVersion) of cuMonitor.exe but need at least $($minimumCUmonitorVersion.ToString())"
            }elseif(!($pathtomodule = Join-Path -Path (Split-Path -Path $cuMonitorServicePath.Trim('"') -Parent) -ChildPath $cuDll)){
                Throw "Unable to find $cuDll in `"$pathtomodule`""
            }elseif(!(Import-Module $pathtomodule -PassThru)){
                Throw "Failed to import module from `"$pathtomodule`""
            }elseif(!(Get-Command -Name 'Get-CUFolders' -ErrorAction SilentlyContinue)){
                Throw "Loaded CU Monitor PowerShell module from `"$pathtomodule`" but unable to find cmdlet Get-CUFolders"
            }
        }catch{
            $exception = $_
            Write-CULog -Msg $exception -ShowConsole -Type E
            ##Write-CULog -Msg (Get-PSCallStack|Format-Table)
            Write-CULog -Msg 'The required ControlUp PowerShell module was not found or could not be loaded. Please make sure this is a ControlUp Monitor machine.' -ShowConsole -Type E
            Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$exception"
            $errorCount++
            break
        }
        #endregion

        #region validate SiteName parameter
        [hashtable] $SiteIdParam = @{}
        [string]$SiteIdGUID = $null
        if ($PSBoundParameters.ContainsKey("SiteName")){
            Write-CULog -Msg "Assigning resources to specific site: $SiteName" -ShowConsole
            
            [array]$cusites = @(Get-CUSites)
            if(!($SiteIdGUID = $cusites | Where-Object { $_.Name -eq $SiteName } | Select-Object -ExpandProperty Id) -or ($SiteIdGUID -is [array] -and $SiteIdGUID.Count -gt 1)){
                $errorMessage = "No unique ControlUp site `"$SiteName`" found (the $($cusites.Count) sites are: $(($cusites | Select-Object -ExpandProperty Name) -join ' , '))"
                Write-CULog -Msg $errorMessage -ShowConsole -Type E
                Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$exception"
                $errorCount++
                break
            }else{
                Write-CULog -Msg "SiteId GUID: $SiteIdGUID" -ShowConsole -SubMsg
                $SiteIdParam.Add('SiteId' , $SiteIdGUID)
            }
        }

        #region Retrieve ControlUp folder structure
        if (-not($DebugCUMachineEnvironment)){
            try {
                $CUComputers = Get-CUComputers # add a filter on path so only computers within the $rootfolder are used
            }catch{
                $errorMessage = "Unable to get computers from ControlUp: $_" 
                Write-CULog -Msg $errorMessage -ShowConsole -Type E
                Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$errorMessage"
                $errorCount++
                break
            }
        }else{
            Write-Debug "Number of objects in DebugCUMachineEnvironment: $($DebugCUMachineEnvironment.count)"
            if ($($DebugCUMachineEnvironment.count) -eq 2){
                foreach ($envObjects in $DebugCUMachineEnvironment){
                    if ($($envObjects  | Get-Member).TypeName[0] -eq "Create-CrazyCUEnvironment.CUComputerObject"){$CUComputers = $envObjects}
                }
            }else{$CUComputers = $DebugCUMachineEnvironment}
        }
        
        Write-CULog -Msg  "CU Computers Count: $(if($CUComputers){ $CUComputers.count }else{ 0 })" -ShowConsole -Color Cyan
        #create a hashtable out of the CUMachines object as it's much faster to query. This is critical when looking up Machines when ControlUp contains ten's of thousands of machines.
        $CUComputersHashTable = @{}
        foreach ($machine in $CUComputers){
            foreach ($obj in $machine){
                $CUComputersHashTable.Add($Obj.Name, $obj)
            }
        }

        if (-not($DebugCUFolderEnvironment)){
            try {
                $CUFolders   = Get-CUFolders # add a filter on path so only folders within the rootfolder are used
            }catch{
                $errorMessage = "Unable to get folders from ControlUp: $_"
                Write-CULog -Msg $errorMessage  -ShowConsole -Type E
                Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$errorMessage"
                $errorCount++
                break
            }
        }else{
            Write-Debug "Number of folder objects in DebugCUMachineEnvironment: $($DebugCUMachineEnvironment.count)"
            if ($($DebugCUMachineEnvironment.count) -eq 2){
                foreach ($envObjects in $DebugCUMachineEnvironment){
                    if ($($envObjects  | Get-Member).TypeName[0] -eq "Create-CrazyCUEnvironment.CUFolderObject"){$CUFolders = $envObjects}
                }
            }else{$CUFolders = Get-CUFolders}
        }

        #endregion
        $OrganizationName = ($CUFolders)[0].path
        Write-CULog -Msg "Organization Name: $OrganizationName" -ShowConsole
        [array]$rootFolders = @(Get-CUFolders | Where-Object FolderType -eq 'RootFolder')
        Write-Verbose -Message "Got $($rootFolders.Count) root folders/organisations: $(($rootFolders | Select-Object -ExpandProperty Path) -join ' , ')"

        [string]$pathSoFar = $null
        [bool]$builtPath = $false
        ## strip off leading \ as CU cmdlets don't like it
        [string[]]$CURootFolderElements = @(($CURootFolder.Trim('\').Split('\')))
        Write-Verbose -Message "Got $($CURootFolderElements.Count) elements in path `"$CURootFolder`""

        ## see if first folder element is the organisation name and if not then we will prepend it as must have that
        if($OrganizationName -ne $CURootFolderElements[0]){
            Write-CULog -Msg "Organization Name `"$OrganizationName`" not found in path `"$CURootFolder`" so adding" -Verbose
            $CURootFolder = Join-Path -Path $OrganizationName -ChildPath $CURootFolder
        }

        ## Code making folders checks if each element in folder exists and if not makes it so no pointmaking path here

        #region Prepare items for synchronization
        #replace FolderPath in ExternalTree object with the local ControlUp Path:
        foreach ($obj in $externalTree){$obj.FolderPath = (Join-Path -Path $CURootFolder -ChildPath $obj.FolderPath).Trim('\')}

        #We also create a hashtable to improve lookup performance for computers in large organizations.
        $ExtTreeHashTable = @{}
        $ExtFolderPaths = New-Object -TypeName System.Collections.Generic.List[psobject]
        foreach ($ExtObj in $externalTree){
            foreach ($obj in $ExtObj){
                ## GRL only add computers since that is all we look up and get duplicate error if OU and computer have the same name
                if($obj.Type -eq 'Computer'){
                    $ExtTreeHashTable.Add($Obj.Name, $obj)
                }else{
                    $ExtFolderPaths.Add($obj)
                }
            }
        }

        Write-CULog -Msg "Target Folder Paths:" -ShowConsole
        if ($ExtFolderPaths.count -ge 25){
            Write-CULog "$($ExtFolderPaths.count) paths detected" -ShowConsole -SubMsg
            foreach ($ExtFolderPath in $ExtFolderPaths){Write-CULog -Msg "$($ExtFolderPath.FolderPath)" -SubMsg}
        }else{
            foreach ($ExtFolderPath in $ExtFolderPaths){Write-CULog -Msg "$($ExtFolderPath.FolderPath)" -ShowConsole -SubMsg}
        }

        $FolderAddBatches   = New-Object System.Collections.Generic.List[PSObject]
        $FoldersToAddBatch  = New-CUBatchUpdate
        $FoldersToAddCount  = 0

        #we'll output the statistics at the end -- also helps with debugging
        $FoldersToAdd          = New-Object System.Collections.Generic.List[PSObject]
        ## There can be problems when folders are added in large numbers so we will see how many new ones are being requested so we can control if necessary
        $FoldersToAddBatchless = New-Object System.Collections.Generic.List[PSObject]
        [hashtable]$newFoldersAdded = @{} ## keep track of what we've issued btch commands to create so we don't duplicate

        foreach ($ExtFolderPath in $ExtFolderPaths.FolderPath){
            if ($ExtFolderPath -notin $CUFolders.Path){ 
                [string]$pathSoFar = $null
                ## Check each part of the path exists, or will be created, and if not add a task to create it
                foreach($pathElement in ($ExtFolderPath.Trim('\')).Split('\')){
                    [string]$absolutePath = $(if($pathSoFar){ Join-Path -Path $pathSoFar -ChildPath $pathElement }else{ $pathElement })
                    if($null -eq $newFoldersAdded[$absolutePath ] -and $absolutePath -notin $CUFolders.Path ){
                        ## there is a bug that causes an error if a folder name being created in a batch already exists at the top level so we workaround it
                        if($batchCreateFolders){
                            if ($FoldersToAddCount -ge $maxFolderBatchSize){
                                Write-Verbose "Generating a new add folder batch"
                                $FolderAddBatches.Add($FoldersToAddBatch)
                                $FoldersToAddCount = 0
                                $FoldersToAddBatch = New-CUBatchUpdate
                            }
                            Add-CUFolder -Name $pathElement -ParentPath $pathSoFar -Batch $FoldersToAddBatch
                        }else{if(!$Preview){$FoldersToAddBatchless.Add([pscustomobject]@{ PathElement = $pathElement ; PathSoFar = $pathSoFar })}}
						
                        $FoldersToAdd.Add("Add-CUFolder -Name `"$pathElement`" -ParentPath `"$pathSoFar`"")
                        $FoldersToAddCount++
                        $newFoldersAdded.Add($absolutePath , $ExtFolderPath)
                    }
                    $pathSoFar = $absolutePath
                }
            }
        }

        if($FoldersToAddBatchless -and $FoldersToAddBatchless.Count){
            [int]$folderDelayMilliseconds = 0
            if($FoldersToAddBatchless.Count -ge $batchCountWarning){
                [string]$logText = "$($FoldersToAddBatchless.Count) folders to add which could cause performance issues"

                if($force){
                    Write-CULog -Msg $logText -ShowConsole -Type W
                    $folderDelayMilliseconds = $folderCreateDelaySeconds * 1000
                }else{
                    $errorMessage = "$logText, aborting - use -force to override" 
                    Write-CULog -Msg $errorMessage -ShowConsole -Type E
                    Send-EmailAlert -SmtpServer $SmtpServer -from $emailFrom -to $emailTo -useSSL:$emailUseSSL -subject "Fatal error from ControlUp sync script `"$callingScript`" on $env:COMPUTERNAME" -body "$errorMessage"
                    $errorCount++
                    break
                }
            }
            foreach($item in $FoldersToAddBatchless){
                Write-Verbose -Message "Creating folder `"$($item.pathElement)`" in `"$($item.pathSoFar)`""
                if(!($folderCreated = Add-CUFolder -Name $item.pathElement -ParentPath $item.pathSoFar) -or $folderCreated -notmatch "^Folder '$($item.pathElement)' was added successfully$"){
                    Write-CULog -Msg "Failed to create folder `"$($item.pathElement)`" in `"$($item.pathSoFar)`" - $folderCreated" -ShowConsole -Type E
                }
                ## to help avoid central CU service becoming overwhelmed
                if($folderDelayMilliseconds -gt 0){
                    Start-Sleep -Milliseconds $folderDelayMilliseconds
                }
            }
        }

        if ($FoldersToAddCount -le $maxFolderBatchSize -and $FoldersToAddCount -ne 0){$FolderAddBatches.Add($FoldersToAddBatch)}

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
        foreach ($ExtComputer in $ExtComputers){
	        if (($CUComputersHashTable.Contains("$($ExtComputer.Name)"))){
    	        if ("$($ExtComputer.FolderPath)\" -notlike "$($CUComputersHashTable[$($ExtComputer.name)].Path)\"){
                    if ($ComputersMoveCount -ge $maxBatchSize){  ## we will execute computer batch operations $maxBatchSize at a time
                        Write-Verbose "Generating a new computer move batch"
                        $ComputersMoveBatches.Add($ComputersMoveBatch)
                        $ComputersMoveCount = 0
                        $ComputersMoveBatch = New-CUBatchUpdate
                    }

        	        Move-CUComputer -Name $ExtComputer.Name -FolderPath "$($ExtComputer.FolderPath)" -Batch $ComputersMoveBatch
                    $MachinesToMove.Add("Move-CUComputer -Name $($ExtComputer.Name) -FolderPath `"$($ExtComputer.FolderPath)`"")
                    $ComputersMoveCount = $ComputersMoveCount+1
    	        }
	        }else{
                if ($ComputersAddCount -ge $maxBatchSize){
                        Write-Verbose "Generating a new add computer batch"
                        $ComputersAddBatches.Add($ComputersAddBatch)
                        $ComputersAddCount = 0
                        $ComputersAddBatch = New-CUBatchUpdate
                    }
                ##write-host $($ExtComputer.FolderPath)
    	        try{Add-CUComputer -Domain $ExtComputer.Domain -Name $ExtComputer.Name -DNSName $ExtComputer.DNSName -FolderPath "$($ExtComputer.FolderPath)" -Batch $ComputersAddBatch @SiteIdParam}
				catch{
                         Write-CULog "Error while attempting to run Add-CUComputer" -ShowConsole -Type E
                         Write-CULog "$($Error[0])"  -ShowConsole -Type E
                }
                if (![string]::IsNullOrEmpty($SiteIdGUID)){
                    $MachinesToAdd.Add("Add-CUComputer -Domain $($ExtComputer.Domain) -Name $($ExtComputer.Name) -DNSName $($ExtComputer.DNSName) -FolderPath `"$($ExtComputer.FolderPath)`" -SiteId $SiteIdGUID")
                }else{
                    $MachinesToAdd.Add("Add-CUComputer -Domain $($ExtComputer.Domain) -Name $($ExtComputer.Name) -DNSName $($ExtComputer.DNSName) -FolderPath `"$($ExtComputer.FolderPath)`"")
                }
                $ComputersAddCount = $ComputersAddCount+1
	        }
        }
        if ($ComputersMoveCount -le $maxBatchSize -and $ComputersMoveCount -ne 0){ $ComputersMoveBatches.Add($ComputersMoveBatch) }
        if ($ComputersAddCount -le $maxBatchSize -and $ComputersAddCount -ne 0)   { $ComputersAddBatches.Add($ComputersAddBatch)   }

        $FoldersToRemoveBatches = New-Object System.Collections.Generic.List[PSObject]
        $FoldersToRemoveBatch   = New-CUBatchUpdate
        $FoldersToRemoveCount   = 0
        #we'll output the statistics at the end -- also helps with debugging
        $FoldersToRemove = New-Object System.Collections.Generic.List[PSObject]
        
        if ($Delete){
            Write-CULog "Determining Objects to be Removed" -ShowConsole
	        # Build batch for folders which are in ControlUp but not in the external source

            [string]$folderRegex = "^$([regex]::Escape($CURootFolder))\\.+"
            [array]$CUFolderSyncRoot = @($CUFolders.Where{ $_.Path -match $folderRegex })
            if($CUFolderSyncRoot -and $CUFolderSyncRoot.Count){
                Write-CULog "Root Target Path : $($CUFolderSyncRoot.Count) subfolders detected" -ShowConsole -Verbose
            }else{
                Write-CULog "Root Target Path : Only Target Folder Exists" -ShowConsole -Verbose
            }
            Write-CULog "Determining Folder Objects to be Removed" -ShowConsole
	        foreach ($CUFolder in $($CUFolderSyncRoot.Path)){
                $folderRegex = "$([regex]::Escape($CUFolder))"
                ## need to test if the whole path matches or it's a sub folder (so "Folder 1" won't match "Folder 12")
                if($ExtFolderPaths.Where({ $_.FolderPath -match "^$folderRegex$" -or $_.FolderPath -match "^$folderRegex\\" }).Count -eq 0 -and $CUFolder -ne $CURootFolder){
                ## can't use a simple -notin as path may be missing but there may be child paths of it - GRL
    	        ##if (($CUFolder -notin $ExtFolderPaths.FolderPath) -and ($CUFolder -ne $("$CURootFolder"))){ #prevents excluding the root folder
                    if ($Delete -and $CUFolder){
                        if ($FoldersToRemoveCount -ge $maxFolderBatchSize){  ## we will execute computer batch operations $maxBatchSize at a time
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

	        foreach ($CUComputer in $CUComputers.Where{$_.path -like "$CURootFolder*"}){
    	            if (!($ExtTreeHashTable[$CUComputer.name].name)){

                        if ($Delete){
							write-host $cucomputer.path
                            if ($FoldersToRemoveCount -ge $maxFolderBatchSize){
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
        if ($FoldersToRemoveCount -le $maxFolderBatchSize -and $FoldersToRemoveCount -ne 0){ $FoldersToRemoveBatches.Add($FoldersToRemoveBatch)   }
        if ($ComputersRemoveCount -le $maxBatchSize -and $ComputersRemoveCount -ne 0)       { $ComputersRemoveBatches.Add($ComputersRemoveBatch)   }

        #endregion

        Write-CULog -Msg "Folders to Add     : $($FoldersToAdd.Count)" -ShowConsole -Color White 
        Write-CULog -Msg "Folders to Add Batches     : $($FolderAddBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($FoldersToAdd.Count) -ge 25){
            foreach ($obj in $FoldersToAdd){Write-CULog -Msg "$obj" -SubMsg}
        }else{
            foreach ($obj in $FoldersToAdd){Write-CULog -Msg "$obj" -ShowConsole -Color Green -SubMsg}
        }

        Write-CULog -Msg "Folders to Remove  : $($FoldersToRemove.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Folders to Remove Batches  : $($FoldersToRemoveBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($FoldersToRemove.Count) -ge 25){
            foreach ($obj in $FoldersToRemove){Write-CULog -Msg "$obj" -SubMsg}
        }else{
            foreach ($obj in $FoldersToRemove){Write-CULog -Msg "$obj" -ShowConsole -Color DarkYellow -SubMsg}
        }

        Write-CULog -Msg "Computers to Add   : $($MachinesToAdd.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Computers to Add Batches   : $($ComputersAddBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($MachinesToAdd.Count) -ge 25){
            foreach ($obj in $MachinesToAdd){Write-CULog -Msg "$obj" -SubMsg} 
        }else{
            foreach ($obj in $MachinesToAdd){Write-CULog -Msg "$obj" -ShowConsole -Color Green -SubMsg}
        }

        Write-CULog -Msg "Computers to Move  : $($MachinesToMove.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Computers to Move Batches  : $($ComputersMoveBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($MachinesToMove.Count) -ge 25){
            foreach ($obj in $MachinesToMove){Write-CULog -Msg "$obj" -SubMsg}
        }else{
            foreach ($obj in $MachinesToMove){Write-CULog -Msg "$obj" -ShowConsole -Color DarkYellow -SubMsg}
        }

        Write-CULog -Msg "Computers to Remove: $($MachinesToRemove.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Computers to Remove Batches: $($ComputersRemoveBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($MachinesToRemove.Count -ge 25)){
            foreach ($obj in $MachinesToRemove){Write-CULog -Msg "$obj" -SubMsg}
        }else{
            foreach ($obj in $MachinesToRemove){Write-CULog -Msg "$obj" -ShowConsole -Color DarkYellow -SubMsg}
        }
            
        $endTime = Get-Date

        Write-CULog -Msg "Build-CUTree took: $($(New-TimeSpan -Start $startTime -End $endTime).Seconds) Seconds." -ShowConsole -Color White
        Write-CULog -Msg "Committing Changes:" -ShowConsole -Color DarkYellow
        if ($ComputersRemoveBatches.Count -gt 0){ $errorCount += Execute-PublishCUUpdates -BatchObject $ComputersRemoveBatches -Message "Executing Computer Object Removal" }
        if ($FoldersToRemoveBatches.Count -gt 0){ $errorCount += Execute-PublishCUUpdates -BatchObject $FoldersToRemoveBatches -Message "Executing Folder Object Removal"   }
        if ($FolderAddBatches.Count -gt 0 -and $batchCreateFolders)       { $errorCount += Execute-PublishCUUpdates -BatchObject $FolderAddBatches -Message "Executing Folder Object Adds"            }
        if ($ComputersAddBatches.Count -gt 0)    { $errorCount += Execute-PublishCUUpdates -BatchObject $ComputersAddBatches -Message "Executing Computer Object Adds"       }
        if ($ComputersMoveBatches.Count -gt 0)   { $errorCount += Execute-PublishCUUpdates -BatchObject $ComputersMoveBatches -Message "Executing Computer Object Moves"     }
        Write-CULog -Msg "Returning $errorCount to caller"
        return $errorCount
}

function Write-CULog {
    Param(
	    [Parameter(Mandatory = $True)][Alias('M')][String]$Msg,
	    [Parameter(Mandatory = $False)][Alias('S')][switch]$ShowConsole,
	    [Parameter(Mandatory = $False)][Alias('C')][String]$Color = "",
	    [Parameter(Mandatory = $False)][Alias('T')][String]$Type = "",
	    [Parameter(Mandatory = $False)][Alias('B')][switch]$SubMsg
   )
    
    $LogType = "INFORMATION..."
    if ($Type -eq "W"){ $LogType = "WARNING........."; $Color = "Yellow" }
    if ($Type -eq "E"){ $LogType = "ERROR..............."; $Color = "Red" }
    if (!($SubMsg)){$PreMsg = "+"}else{$PreMsg = "`t>"}
    $date = Get-Date -Format G
    if ($Global:LogFile){Write-Output "$date | $LogType | $Msg"  | Out-file $($Global:LogFile) -Append}
            
    if (!($ShowConsole)){
	    if (($Type -eq "W") -or ($Type -eq "E")){Write-Host "$PreMsg $Msg" -ForegroundColor $Color;$Color = $null}
		else{Write-Verbose -Message "$PreMsg $Msg";$Color = $null}
    }else{
	    if ($Color -ne ""){Write-Host "$PreMsg $Msg" -ForegroundColor $Color;$Color = $null}
		else{Write-Host "$PreMsg $Msg"}
    }
}
$BuildCUTreeParams = @{CURootFolder = $cuDGFolder}
if ($Preview) {$BuildCUTreeParams.Add("Preview",$true)}
if ($Delete) {$BuildCUTreeParams.Add("Delete",$true)}
if ($LogFile){$BuildCUTreeParams.Add("LogFile",$LogFile)}
if ($Site){$BuildCUTreeParams.Add("SiteName",$Site)}
if ($Force){$BuildCUTreeParams.Add("Force",$true)}
if ($batchCreateFolders){$BuildCUTreeParams.Add("batchCreateFolders",$true)}
[int]$errorCount = Build-CUTree -ExternalTree $Environment @BuildCUTreeParams
	
