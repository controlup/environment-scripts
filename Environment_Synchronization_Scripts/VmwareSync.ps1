[CmdletBinding()]
Param(
	[Parameter(Mandatory=$false)][string]$FolderPath,
	[Parameter(Mandatory=$false)][string]$server,
	[Parameter(Mandatory=$false)][switch]$Delete,
	[Parameter(Mandatory=$false)][switch]$Preview
)
Get-Item "$((get-childitem 'C:\Program Files\Smart-X\ControlUpMonitor\' -Directory)[-1].fullName)\*powershell*.dll"|import-module
import-module vmware.powercli
$Environment = New-Object -TypeName System.Collections.Generic.List[PSObject]
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
$exportPath = "$($env:programdata)\ControlUp\SyncScripts"
$credFile = "$exportPath\VCenterSync.xml"
if(!(test-path $credFile)){$cred = Get-Credential -Message "Please enter VCenter Credentials"|Export-Clixml -path $credFile}
New-Item -ItemType Directory -Force -Path $exportPath |out-null
$syncFolder = "$FolderPath"
if(!$syncFolder){throw "No sync folder found. `n`nPlease use arguments or the config file"}
if((test-path $credFile)){$cred=$null;$cred = Import-Clixml -path $credFile}else{throw "Error importing Credential File"}
if(!$cred){throw "No Credential Found"}

$exPath = "$exportPath\exclusions.lst"
if(!(test-path $exPath)){$null|Out-File -FilePath $exPath -Force}
	

$global:sf = $syncFolder
$root = (Get-CUFolders)[0].name.toLower()
$rootPath = "$root\$syncFolder"
$vmList = [System.Collections.ArrayList]@()
$folderList = [System.Collections.ArrayList]@()

$global:exclusions = get-content $exPath

Connect-VIServer $server -credential $cred

Function Get-VMFolderPath {
    param([string]$VMFolderId)

    $Folders = [system.collections.arraylist]::new()
    $tracker = Get-Folder -Id $VMFolderId
    $Obj = [pscustomobject][ordered]@{FolderName = $tracker.Name; FolderID = $tracker.Id}
    $null = $Folders.add($Obj)

    while ($tracker) {
       if ($tracker.parent.type) {
        $tracker = (Get-Folder -Id $tracker.parentId)
        $Obj = [pscustomobject][ordered]@{FolderName = $tracker.Name; FolderID = $tracker.Id}
        $null = $Folders.add($Obj)
           }
           else {
        $Obj = [pscustomobject][ordered]@{FolderName = $tracker.parent.name; FolderID = $tracker.parentId}
        $null = $Folders.add($Obj)
            $tracker = $null
       }
    }
    $Folders.Reverse()
    $Folders.FolderName -join "/"
}

$vmList = get-vm


foreach ($vm in $vmlist){

	try{$DNS=$null;$DNS = [System.Net.Dns]::GetHostByName($vm.name).hostname}catch{}
	$folder = $vm|%{get-vmfolderpath $_.folder.id}
	$FolderList.add($folder)|out-null
	
	
	if($DNS){$domain = $DNS.substring($DNS.indexof(".")+1)}
	
	$skip = $false
	foreach ($i in $global:exclusions){
		$name = "$folder\$($vm.name)".toLower()
		if($name -like "*$i*".toLower()){$skip = $true;break}
	}
	
	if($DNS -and $folder -and $($vm.name) -and $domain -and !$skip){
		$Environment.Add([ControlUpObject]::new($vm.name, $folder ,"Computer", $Domain ,"Added by Sync Script", $DNS))
	}

}

$uniqueFolders = ($FolderList|?{$_ -ne $root -and $_ -ne $rootPath})|sort -unique
foreach ($uf in $uniqueFolders){
	$folderName = $uf.split("\")[-1]
	$addFolderTo = $uf.replace("$rootPath\","")
		$skip = $false
	foreach ($i in $global:exclusions){
		$name = "$folderName".toLower()
		if($name -like "*$i*".toLower()){$skip = $true;break}
	}
	if(!$skip){	$Environment.Add([ControlUpObject]::new($FolderName,$addFolderTo,"Folder",$null,$null,$null))}
}


############################
##### Start BuildCUTree ####
############################
function Build-CUTree {
    [CmdletBinding()]
    Param(
	    [Parameter(Mandatory=$true,HelpMessage='Object to build tree within ControlUp')]
	    [PSObject] $ExternalTree,
	    [Parameter(Mandatory=$false,HelpMessage='ControlUp root folder to sync')]
	    [string] $CURootFolder,
	    [Parameter(Mandatory=$false,HelpMessage='ControlUp root ')]
	    [string] $CUSyncFolder,
 	    [Parameter(Mandatory=$false, HelpMessage='Delete CU objects which are not in the external source')]
	    [switch] $Delete,
        [Parameter(Mandatory=$false, HelpMessage='Generate a report of the actions to be executed')]
        [switch]$Preview,
        [Parameter(Mandatory=$false, HelpMessage='Save a log file')]
	    [string] $LogFile,
        [Parameter(Mandatory=$false, HelpMessage='ControlUp Site name to assign the machine object to')]
	    [string] $SiteName,
        [Parameter(Mandatory=$false, HelpMessage='Create folders in batches rather than individually')]
	    [switch] $batchCreateFolders 
	)	
		$batchCreateFolders = $true
        $maxBatchSize = 1000
        $maxFolderBatchSize = 100
        [int]$errorCount = 0
		
        function Execute-PublishCUUpdates {
            Param([Parameter(Mandatory = $True)][Object]$BatchObject,[Parameter(Mandatory = $True)][string]$Message)
            [int]$returnCode = 0
            [int]$batchCount = 0
            foreach ($batch in $BatchObject){
                $batchCount++
                Write-CULog -Msg "$Message. Batch $batchCount/$($BatchObject.count)" -ShowConsole -Color DarkYellow -SubMsg
                if (!($preview)){
                    [datetime]$timeBefore = [datetime]::Now
                    $result = Publish-CUUpdates -Batch $batch 
                    [datetime]$timeAfter = [datetime]::Now
                    [array]$results = @(Show-CUBatchResult -Batch $batch)
                    [array]$failures = @($results.Where({$_.IsSuccess -eq $false}))

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
            if (!(Test-Path $($PSBoundParameters.LogFile))){
                Write-CULog -Msg "Creating Log File" #Attempt to create the file
                if (!(Test-Path $($PSBoundParameters.LogFile))){Write-Error "Unable to create the report file" -ErrorAction Stop}
            }else{Write-CULog -Msg "Beginning Synchronization"}
            Write-CULog -Msg "Detected the following parameters:"
            foreach($psbp in $PSBoundParameters.GetEnumerator()){
                if ($psbp.Key -like "ExternalTree"){
                    Write-CULog -Msg $("Parameter={0} Value={1}" -f $psbp.Key,$psbp.Value.count)
                }else{Write-CULog -Msg $("Parameter={0} Value={1}" -f $psbp.Key,$psbp.Value)}
            }
        }else{$Global:LogFile = $false}


        $startTime = Get-Date
        [string]$errorMessage = $null

        #region Retrieve ControlUp folder structure
            try {$CUComputers = Get-CUComputers}
			catch{
                $errorMessage = "Unable to get computers from ControlUp: $_" 
                Write-CULog -Msg $errorMessage -ShowConsole -Type E
                $errorCount++
                break
            }
			
        Write-CULog -Msg  "CU Computers Count: $(if($CUComputers){ $CUComputers.count }else{ 0 })" -ShowConsole -Color Cyan
        #create a hashtable out of the CUMachines object as it's much faster to query. This is critical when looking up Machines when ControlUp contains ten's of thousands of machines.
        $CUComputersHashTable = @{}
        foreach ($machine in $CUComputers){
			foreach ($obj in $machine){
					$CUComputersHashTable.Add($Obj.Name, $obj)
				}
		}
		
			try {$CUFolders = Get-CUFolders # add a filter on path so only folders within the rootfolder are used
			}catch{
				$errorMessage = "Unable to get folders from ControlUp: $_"
				Write-CULog -Msg $errorMessage  -ShowConsole -Type E
				$errorCount++
				break
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
        if($OrganizationName -ne $CURootFolderElements[0]){$CURootFolder = Join-Path -Path $OrganizationName -ChildPath $CURootFolder}

        ## Code making folders checks if each element in folder exists and if not makes it so no pointmaking path here
        #replace FolderPath in ExternalTree object with the local ControlUp Path:
        foreach ($obj in $externalTree){$obj.FolderPath = (Join-Path -Path $CURootFolder -ChildPath $obj.FolderPath).Trim('\')}

        #We also create a hashtable to improve lookup performance for computers in large organizations.
        $ExtTreeHashTable = @{}
        $ExtFolderPaths = New-Object -TypeName System.Collections.Generic.List[psobject]
        foreach ($ExtObj in $externalTree){foreach ($obj in $ExtObj){if($obj.Type -eq 'Computer'){$ExtTreeHashTable.Add($Obj.Name, $obj)}else{$ExtFolderPaths.Add($obj)}}}
        #foreach ($ExtObj in $externalTree){foreach ($obj in $ExtObj){if($obj.Type -eq 'Computer'){$ExtTreeHashTable.Add($Obj.Name, $obj)}else{if($obj.folderpath -notlike "*$OrganizationName*\$OrganizationName*"){$ExtFolderPaths.Add($obj)}}}}
		#$ExtFolderPaths.folderpath|%{write-host $_};pause
        Write-CULog -Msg "Target Folder Paths:"
        Write-CULog "$($ExtFolderPaths.count) paths detected" -ShowConsole -SubMsg
        foreach ($ExtFolderPath in $ExtFolderPaths){Write-CULog -Msg "$($ExtFolderPath.FolderPath)" -SubMsg}

        $FolderAddBatches   = New-Object System.Collections.Generic.List[PSObject]
        $FoldersToAddBatch  = New-CUBatchUpdate
        $FoldersToAddCount  = 0

        #we'll output the statistics at the end -- also helps with debugging
        $FoldersToAdd          = New-Object System.Collections.Generic.List[PSObject]
        [hashtable]$newFoldersAdded = @{} ## keep track of what we've issued batch commands to create so we don't duplicate
		
        foreach ($ExtFolderPath in $ExtFolderPaths.FolderPath){
            if ($ExtFolderPath -notin $CUFolders.Path){ 
                [string]$pathSoFar = $null
                ## Check each part of the path exists, or will be created, and if not add a task to create it
				#write-host $ExtFolderPath
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
                        }
						
                        $FoldersToAdd.Add("Add-CUFolder -Name `"$pathElement`" -ParentPath `"$pathSoFar`"")
                        $FoldersToAddCount++
                        $newFoldersAdded.Add($absolutePath , $ExtFolderPath)
                    }
                    $pathSoFar = $absolutePath
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
                #write-host $($ExtComputer.FolderPath)
				#write-host "$($extComputer.Name) - $($extComputer.Domain) - $($extComputer.Name) - $($extComputer.DNSName) - $($extComputer.Site)"
    	        try{Add-CUComputer -Domain $ExtComputer.Domain -Name $ExtComputer.Name -DNSName $ExtComputer.DNSName -FolderPath "$($ExtComputer.FolderPath)" -siteId $extComputer.Site -Batch $ComputersAddBatch}
				catch{Write-CULog "Error while attempting to run Add-CUComputer" -ShowConsole -Type E; Write-CULog "$($Error[0])"  -ShowConsole -Type E}
				
                $MachinesToAdd.Add("Add-CUComputer -Domain $($ExtComputer.Domain) -Name $($ExtComputer.Name) -DNSName $($ExtComputer.DNSName) -FolderPath `"$($ExtComputer.FolderPath)`" -SiteId $SiteIdGUID")
                $ComputersAddCount = $ComputersAddCount+1
	        }
        }
        if ($ComputersMoveCount -le $maxBatchSize -and $ComputersMoveCount -ne 0){$ComputersMoveBatches.Add($ComputersMoveBatch)}
        if ($ComputersAddCount -le $maxBatchSize -and $ComputersAddCount -ne 0){$ComputersAddBatches.Add($ComputersAddBatch)}

        $FoldersToRemoveBatches = New-Object System.Collections.Generic.List[PSObject]
        $FoldersToRemoveBatch   = New-CUBatchUpdate
        $FoldersToRemoveCount   = 0
        #we'll output the statistics at the end -- also helps with debugging
        $FoldersToRemove = New-Object System.Collections.Generic.List[PSObject]
        
		
        if ($Delete){
            Write-CULog "Determining Objects to be Removed" -ShowConsole
	        # Build batch for folders which are in ControlUp but not in the external source
			$cuFolderSyncroot = "$CURootFolder$CUSyncFolder"
            [string]$folderRegex = "^$([regex]::Escape($cuFolderSyncroot))\\.+"
            [array]$CUFolderSyncRoot = @($CUFolders.Where{ $_.Path -match $folderRegex })
			
            if($CUFolderSyncRoot -and $CUFolderSyncRoot.Count){Write-CULog "Root Target Path : $($CUFolderSyncRoot.Count) subfolders detected" -ShowConsole -Verbose}
			else{Write-CULog "Root Target Path : Only Target Folder Exists" -ShowConsole -Verbose}
            Write-CULog "Determining Folder Objects to be Removed" -ShowConsole

	        foreach ($CUFolder in $($CUFolderSyncRoot.Path)){
                $folderRegex = "$([regex]::Escape($CUFolder))"
                ## need to test if the whole path matches or it's a sub folder (so "Folder 1" won't match "Folder 12")
                if($ExtFolderPaths.Where({ $_.FolderPath -match "^$folderRegex$" -or $_.FolderPath -match "^$folderRegex\\" }).Count -eq 0 -and $CUFolder -ne $CURootFolder){
                ## can't use a simple -notin as path may be missing but there may be child paths of it - GRL
    	        ##if (($CUFolder -notin $ExtFolderPaths.FolderPath) -and ($CUFolder -ne $("$CURootFolder"))){ #prevents excluding the root folder
						$skip = $false
						foreach ($path in $global:eucDisconnectedMsg){if ($CUFolder -like "$path*"){$skip = $true;break}}
						if ($Delete -and $CUFolder -and !$Skip){
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
			
	        foreach ($CUComputer in $CUComputers.Where{$_.path.toLower() -like "$CURootFolder$CUSyncFolder*".toLower()}){
				
    	            if (!($ExtTreeHashTable[$CUComputer.name].name)){
						$CUComputerPath = $cucomputer.path
						$skip = $false
						foreach ($path in $global:eucDisconnectedMsg){
							if ($CUComputerPath -like "$path*"){$skip = $true;break}
						}
						if (($ExtComputers.Contains("$($CUComputer.name)"))){$skip = $true}

                        if ($Delete -and !$skip){							
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
            foreach ($obj in $FoldersToAdd){Write-CULog -Msg "$obj"} #-ShowConsole -Color Green -SubMsg}
        }

        Write-CULog -Msg "Folders to Remove  : $($FoldersToRemove.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Folders to Remove Batches  : $($FoldersToRemoveBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($FoldersToRemove.Count) -ge 25){
            foreach ($obj in $FoldersToRemove){Write-CULog -Msg "$obj" -SubMsg}
        }else{
            foreach ($obj in $FoldersToRemove){Write-CULog -Msg "$obj"} #-ShowConsole -Color DarkYellow -SubMsg}
        }

        Write-CULog -Msg "Computers to Add   : $($MachinesToAdd.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Computers to Add Batches   : $($ComputersAddBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($MachinesToAdd.Count) -ge 25){
            foreach ($obj in $MachinesToAdd){Write-CULog -Msg "$obj" -SubMsg} 
        }else{
            foreach ($obj in $MachinesToAdd){Write-CULog -Msg "$obj"} #-ShowConsole -Color Green -SubMsg}
        }

        Write-CULog -Msg "Computers to Move  : $($MachinesToMove.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Computers to Move Batches  : $($ComputersMoveBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($MachinesToMove.Count) -ge 25){
            foreach ($obj in $MachinesToMove){Write-CULog -Msg "$obj" -SubMsg}
        }else{
            foreach ($obj in $MachinesToMove){Write-CULog -Msg "$obj"} #-ShowConsole -Color DarkYellow -SubMsg}
        }

        Write-CULog -Msg "Computers to Remove: $($MachinesToRemove.Count)" -ShowConsole -Color White
        Write-CULog -Msg "Computers to Remove Batches: $($ComputersRemoveBatches.Count)" -ShowConsole -Color Gray -SubMsg
        if ($($MachinesToRemove.Count -ge 25)){foreach ($obj in $MachinesToRemove){Write-CULog -Msg "$obj" -SubMsg}}else{foreach ($obj in $MachinesToRemove){Write-CULog -Msg "$obj"}}
            
        $endTime = Get-Date
		$bcutStart = get-date
        Write-CULog -Msg "Build-CUTree took: $($(New-TimeSpan -Start $startTime -End $endTime).Seconds) Seconds." -ShowConsole -Color White
        Write-CULog -Msg "Committing Changes:" -ShowConsole -Color DarkYellow
		if ($FolderAddBatches.Count -gt 0 -and $batchCreateFolders){ $errorCount += Execute-PublishCUUpdates -BatchObject $FolderAddBatches -Message "Executing Folder Object Adds"}
		#write-host "Waiting 30 seconds for folder creation";start-sleep 30
		if ($ComputersAddBatches.Count -gt 0){ $errorCount += Execute-PublishCUUpdates -BatchObject $ComputersAddBatches -Message "Executing Computer Object Adds"}
		if ($ComputersMoveBatches.Count -gt 0){$errorCount += Execute-PublishCUUpdates -BatchObject $ComputersMoveBatches -Message "Executing Computer Object Moves"}
        if ($ComputersRemoveBatches.Count -gt 0){ $errorCount += Execute-PublishCUUpdates -BatchObject $ComputersRemoveBatches -Message "Executing Computer Object Removal"}
		if ($FoldersToRemoveBatches.Count -gt 0){ $errorCount += Execute-PublishCUUpdates -BatchObject $FoldersToRemoveBatches -Message "Executing Folder Object Removal"}
        Write-CULog -Msg "Returning $errorCount to caller"
		$bcutEnd = get-date
		Write-CULog -Msg "Committing Changes took: $($(New-TimeSpan -start $bcutStart -end $bcutEnd).Seconds) Seconds."
		Write-Host -Msg "Committing Changes took: $($(New-TimeSpan -start $bcutStart -end $bcutEnd).totalSeconds) Seconds."
        return $errorCount
}

############################
#####  End BuildCUTree  ####
############################


$BuildCUTreeParams = @{CURootFolder = $syncFolder}
$BuildCUTreeParams.Add("Force",$true)
if ($Preview){$BuildCUTreeParams.Add("Preview",$true)}
if ($Delete){$BuildCUTreeParams.Add("Delete",$true)}
if ($LogFile){$BuildCUTreeParams.Add("LogFile",$LogFile)}
if ($batchCreateFolders){$BuildCUTreeParams.Add("batchCreateFolders",$true)}


[int]$errorCount = Build-CUTree -ExternalTree $Environment @BuildCUTreeParams



