#requires -version 5
[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$true, HelpMessage='Enter a ControlUp subfolder to save your AD tree to')]
    [ValidateNotNullOrEmpty()]
    [string] $folderPath,

    [Parameter(Mandatory=$true, HelpMessage='Enter the Distinguished Name of the OU to sync')]
    [ValidateNotNullOrEmpty()]
    [string] $OU,
    
    [Parameter(Mandatory=$false, HelpMessage='Domain to query if not the default')]
    [ValidateNotNullOrEmpty()]
    [string] $domain,

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

$Exclude = "Xen Virtual,RemotePC"
[Array]$ExcludedWords = if($Exclude){$Exclude.split(",")}else{$null}

## GRL this way allows script to be run with debug/verbose without changing script
$VerbosePreference = $(if( $PSBoundParameters[ 'verbose' ] ) { $VerbosePreference } else { 'SilentlyContinue' })
$DebugPreference = $(if( $PSBoundParameters[ 'debug' ] ) { $DebugPreference } else { 'SilentlyContinue' })
$ErrorActionPreference = $(if( $PSBoundParameters[ 'erroraction' ] ) { $ErrorActionPreference } else { 'Stop' })
$ProgressPreference = 'SilentlyContinue'

## Script from ControlUp which must reside in the same folder as this script
[string]$buildCuTreeScript = 'Build-CUTree.ps1'

function Make-NameWithSafeCharacters ([string]$string) {
    #list of illegal characters : '/', '\', ':', '*','?','"','<','>','|','{','}'
    $returnString = (($string).Replace("/","-")).Replace("\","-").Replace(":","-").Replace("*","-").Replace("?","-").Replace("`"","-").Replace("<","-").Replace(">","-").Replace("|","-").Replace("{","-").Replace("}","-")
    return $returnString
}
## GRL This needs PoSH 5.0 minimum so is it worth it rather than just using PSObjects?
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

#region AD Functions

Function ConvertTo-LDAP {
    
    Param
    (
        [string]$name
    )

    [string]$convertedName = $Null

    if( ! [string]::IsNullOrEmpty( $name ) )
    {
        if( $name -match '^DC=' )
        {
            $convertedName = 'LDAP://{0}' -f $name
        }
        else
        {
            $convertedName = 'LDAP://DC={0}' -f ( $name -replace '\.' , ',DC=' )
        }
    }

    $convertedName
}

Function Get-DomainNameFromDistinguishedName {
  
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Mandatory=$true, 
            HelpMessage='Enter the Distinguished name for an absolute path.  Eg OU=TestOU,DC=bottheory,DC=local',
            ParameterSetName="WithDistinguishedName"        )]
        [ValidateNotNullOrEmpty()]
        [string] $DistinguishedName ,
        
        [Parameter(
            Mandatory=$false, 
            HelpMessage='Enter the domain name to search if not the current.  Eg acme.local'        )]
        [AllowNull()]
        [string] $DomainName
    )

    Process {

        $RootDN = ($DistinguishedName | Select-String "(DC.+)\w+"  -AllMatches).Matches.Value
        
        $Search = [adsisearcher]"(&(distinguishedName=$RootDN))"
        if( $PSBoundParameters[ 'domainName' ] -and ! [string]::IsNullOrEmpty( $domainName ) )
        {
            $search.SearchRoot = ConvertTo-LDAP -name $DomainName
        }
        $search.PropertiesToLoad.AddRange(@('canonicalname'))
        $RootDNProperties = $Search.FindAll()
        if (-not([string]::IsNullOrEmpty($RootDNProperties.properties.canonicalname))) {
            return ($RootDNProperties.properties.canonicalname).Replace("/","")
        } else {
            Write-Error "Unable to return the canonical name for the root domain"
        }
    }
}

Function Get-OU {
  
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Mandatory=$true, 
            HelpMessage='Enter a organizational unit name',
            ParameterSetName="WithName"        )]
        [ValidateNotNullOrEmpty()]
        [string] $Name,

        [Parameter(
            Mandatory=$true, 
            HelpMessage='Enter the Distinguished name for an absolute path.  Eg OU=TestOU,DC=bottheory,DC=local',
            ParameterSetName="WithDistinguishedName"        )]
        [ValidateNotNullOrEmpty()]
        [string] $DistinguishedName ,

        [Parameter(
            Mandatory=$false, 
            HelpMessage='Enter the domain name to search if not the current.  Eg acme.local'        )]
        [AllowNull()]
        [string] $DomainName
    )

    Process {

        [string]$searchString = "(&(objectCategory=organizationalUnit)"
        $OUObjects = $null

        if( $PSBoundParameters[ 'name' ] )
        {
            $searchString = -join $searchString , "(name=$name))"
        }
        elseif( $PSBoundParameters[ 'DistinguishedName' ] )
        {
            $searchString = -join $searchString , "(distinguishedName=$DistinguishedName))"
        }

        if ( $Search = [adsisearcher]$searchString )
        {
            if( $PSBoundParameters[ 'domainName' ] -and ! [string]::IsNullOrEmpty( $domainName ) )
            {
                $search.SearchRoot = ConvertTo-LDAP -name $DomainName
            }
            $OUObjects = @( $Search.FindAll() )
            if ($OUObjects.count -ne 1 )
            {
                Write-Error -Message "$($OUObjects.count) OUs were found for `"$(($search.Filter.Split( '&=()' , [System.StringSplitOptions]::RemoveEmptyEntries ))[-1])`""
            }
        }
        else
        {
            Write-Error -Message "Failed to create ADSI searcher from $searchString"
        }

        $OUObjects ## return
    }
}

Function Get-ObjectsInOU {
  
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='Enter an organizational unit'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $DistinguishedName

    )

    Process {

        $Search = [adsisearcher]::new()
        $Search.SearchRoot.Path = "LDAP://$DistinguishedName"
        $Search.SearchRoot.distinguishedName = "LDAP://$DistinguishedName"
        $Search.PageSize = 100000 #retrieve a maximum of 100,000 objects
        [void]$Search.PropertiesToLoad.Add("name") #filter down the properties to return to speed up the results.
        [void]$Search.PropertiesToLoad.Add("cn")
        [void]$Search.PropertiesToLoad.Add("distinguishedname")
        [void]$Search.PropertiesToLoad.Add("samaccountname")
        [void]$Search.PropertiesToLoad.Add("objectcategory")
        [void]$Search.PropertiesToLoad.Add("dnshostname")

        $OUCollectionObj = New-Object System.Collections.Generic.List[PSObject]
        Write-Verbose "Searching $($DistinguishedName)"
        $SearchStartTime = Get-Date
        #Write-Verbose "$SearchStartTime"
        $Search.FindAll().ForEach( { $OUCollectionObj.Add( $_ ) } )
    
        $SearchEndTime = Get-Date
        #Write-Verbose "$SearchEndTime"
        Write-Verbose "Searching $($DistinguishedName) found $($OUCollectionObj.count) objects in $($($SearchEndTime.Second - $SearchStartTime.Second)) seconds"

        if ($OUCollectionObj.count -gt 1) {
            return $OUCollectionObj
        } else {
            Write-Error "No objects found"
        }
    }
}

Function Convert-LDAPPathToPath {
  
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            HelpMessage='LDAP Path'
        )]
        [ValidateNotNullOrEmpty()]
        [string] $LDAPPath
    )

    Process {
        if ($LDAPPath.StartsWith("LDAP://")) {
        
            $LDAPPath = $LDAPPath.Replace("LDAP://","").replace('OU=','').replace('OU=','').replace('DC=','').Replace('CN=','') #remove URI
            $SplitLDAPPath = $LDAPPath -split '(?<!\\),' #remove commas that are not escaped
            [char[]]$CharsToEscape = ',\#+<>;"=' #all escaped LDAP characters https://social.technet.microsoft.com/wiki/contents/articles/5312.active-directory-characters-to-escape.aspx
            $PathBuilder = New-Object System.Collections.Generic.List[PSObject]
            #$pathBuilderTimer = Measure-Command {
            foreach ($obj in $SplitLDAPPath) {
               # $innerpathBuilderTimer = Measure-Command {
                if ($obj.Contains($CharsToEscape)) {      
                    foreach ($char in $CharsToEscape) {
                        if ($char -eq "\") {
                            $obj = $obj.replace("\$char", $char ) #escape out the escape backslash differently
                        } else {
                            $obj = $obj.replace("`\$char", $char )
                        }
                    }
                }
                
                $pathString = (($obj).Replace("/","-")).Replace("\","-").Replace(":","-").Replace("*","-").Replace("?","-").Replace("`"","-").Replace("<","-").Replace(">","-").Replace("|","-").Replace("{","-").Replace("}","-") #this is faster than calling the function.  Calling the function adds ~16ms to each iteration as opposed to just doing it directly (on a 2012 CPU)
                #$pathString = Make-NameWithSafeCharacters -string $obj  #this will replace most of the ldap caharacters with dashes
                
                #$pathString = $pathString -Replace "$","\" #adds a backslash to the end of each line
                $PathBuilder.Insert(0,"$pathString\") #LDAP returns objects with the leaf on the left side. Insert(0... will reverse the order so the leaf ends up on the right
                #}
                #[Console]::WriteLine("innerPathBuilderTimer took: $($innerpathBuilderTimer.totalMilliseconds)")
                }
            
            #}
            #[Console]::WriteLine("pathBuilderTimer took: $($pathBuilderTimer.totalMilliseconds)")
            
            $LDAPPath = (-join($PathBuilder))
            return $LDAPPath.TrimEnd("\")
        } else {
            Write-Error "Unable to detect proper LDAP:\\ path:  $LDAPPath"
        }
    }
}
#endregion

# dot sourcing Functions
## GRL Don't assume user has changed location so get the script path instead
[string]$scriptPath = Split-Path -Path (& { $myInvocation.ScriptName }) -Parent
[string]$buildCuTreeScriptPath = [System.IO.Path]::Combine( $scriptPath , $buildCuTreeScript )

if( ! ( Test-Path -Path $buildCuTreeScriptPath -PathType Leaf -ErrorAction SilentlyContinue ) )
{
    Throw "Unable to find script `"$buildCuTreeScript`" in `"$scriptPath`""
}

. $buildCuTreeScriptPath       

$ADEnvironment = New-Object -TypeName System.Collections.Generic.List[PSObject]

#region Generate ControlUpObject that will be used to Add objects into ControlUp.
Write-Host "Getting OU Object: $OU" -foregroundColor Yellow
if( ! ( $RootOU = Get-OU -DistinguishedName $OU -DomainName $domain ) )
{
    Throw "Failed to locate OU $OU"
}

if( ! ( $OUObjects = Get-ObjectsInOU -DistinguishedName $OU ) -or ! $OUObjects.Count )
{
    Throw "Found no objects in OU $OU"
}

[string]$RootPath = Convert-LDAPPathToPath -LDAPPath $RootOU.Path
Write-Verbose "RootPath: $RootPath"
[string]$RootFolder = $RootPath.Split("\")[-1]
[System.TimeSpan]$ListOfOUsTimer = Measure-Command { $ListOfOUs =  $OUObjects.Where( { $_.properties.objectcategory -like "*organizational-Unit*" } ) }
[int]$OUCounts = $ListOfOUs.count
Write-Verbose "Found $($OUCounts) OUs in $($ListOfOUsTimer.TotalSeconds) seconds"


#find AD OU's to be made into ControlUp folders:
[System.TimeSpan]$ConvertOU_ToFoldersTimer = Measure-Command {
[int]$OUCount = 0
[Console]::WriteLine("$((get-date).ToLongTimeString()) : Processing $OUCounts OU's...")
    foreach ($OUObject in $ListOfOUs) {
        if (0 -eq $OUCount % 10) {
            [Console]::WriteLine("$((get-date).ToLongTimeString()) : Processed $OUCount / $($OUCounts)...")
        }
        [string]$OUObjectPath = Convert-LDAPPathToPath -LDAPPath $OUObject.Path  # Get the Path in the format we require
        $OUObjectPath = $OUObjectPath.Replace("$RootPath","")            # Remove the RootPath of the LDAP structure
        $OUObjectPath = $RootFolder + $OUObjectPath                      # Re-add the RootFolder (You can probably remove this if you want to add the machines directly to the root of the targetFolder
        $FolderName = $OUObjectPath.Split("\")[-1]                       # Folder name in ControlUp
		
		foreach ($exclusion in $excludedWords){
			if($OUObjectPath.toLower() -notlike "*$($exclusion.toLower())*"){
				$ADEnvironment.Add([ControlUpObject]::new($FolderName ,"$OUObjectPath","Folder","","AD OU",""))
				$OUCount++
			}
		}
    }
}
## GRL Why do we have Write-Host and [Console]::WriteLine ?
[Console]::WriteLine("$((get-date).ToLongTimeString()) : Processed $($OUCounts) / $($OUCounts)...")

Write-Verbose "Converting OUs to ControlUp Folder Import objects took $($ConvertOU_ToFoldersTimer.TotalSeconds) seconds"
$ListOfComputersTimer = Measure-Command { $ListOfComputers =  $OUObjects.Where({$_.properties.objectcategory -like "*computer*"}) }
$CanonicalNameOfDomain = Get-DomainNameFromDistinguishedName -DistinguishedName $OU -DomainName $domain

$ComputerCount = $ListOfComputers.count
Write-Verbose "Found $($ListOfComputers.count) computer objects in $($ListOfComputersTimer.TotalSeconds) seconds"
[Console]::WriteLine("$((get-date).ToLongTimeString()) : Processing $($ListOfComputers.count) computers's...")
$ConvertComputer_ToCUComputersTimer = Measure-Command {
    [int]$compCount = 0
    foreach ($CompObject in $ListOfComputers) {
        if (0 -eq $compCount % 250) {
            [Console]::WriteLine("$((get-date).ToLongTimeString()) : Processed $compCount / $($ComputerCount)...")
        }
        $CompObjectPath = Convert-LDAPPathToPath -LDAPPath $CompObject.Path   # Get the Path in the format we require
        $CompObjectPath = $CompObjectPath.Replace("$RootPath","")             # Remove the RootPath of the LDAP structure
        $CompObjectPath = $RootFolder + $CompObjectPath                       # Re-add the RootFolder (You can probably remove this if you want to add the machines directly to the root of the targetFolder
        $ComputerName = $CompObjectPath.Split("\")[-1]                        # Computer name in ControlUp
        $CompObjectPath = $CompObjectPath.Replace("\$ComputerName","")        # Remove computer name from the folder path
        $CompObjectDNSName = $CompObject.Properties[ 'dnshostname' ]          # DNS Name
		foreach ($exclusion in $excludedWords){
			if($CompObjectPath.toLower() -notlike "*$($exclusion.toLower())*"){
				$ADEnvironment.Add( ([ControlUpObject]::new( $ComputerName , $CompObjectPath ,"Computer", $CanonicalNameOfDomain , "AD Computer", $CompObjectDNSName )))
				$compCount++
			}
		}
    }
}
[Console]::WriteLine("$((get-date).ToLongTimeString()) : Processed $($ComputerCount) / $($ComputerCount)...")
Write-Verbose "Converting OUs to ControlUp Folder Import objects took $($ConvertComputer_ToCUComputersTimer.TotalSeconds) seconds"

#endregion
    

#Write-Verbose "$($ADEnvironment | Format-Table | Out-String)"
<#
Returns an object like so:
    Name                   FolderPath                           Type     Domain                      Description     DNSName
    ----                   ----------                           ----     ------                      -----------     -------
    Citrix                 Citrix                               Folder                               AD OU
    Edmonton               Citrix\Edmonton                      Folder                               AD OU
    North                  Citrix\Edmonton\North                Folder                               AD OU
    WSCTXEDN001            Citrix\Edmonton\North                Computer bottheory.local             AD Computer     WSCTXEDN001.bottheory.local
    South                  Citrix\Edmonton                      Folder                               AD OU
    WSCTXEDS001            Citrix\Edmonton\South                Computer bottheory.local             AD Computer     WSCTXEDS001.bottheory.local
    WSCTXEDS002            Citrix\Edmonton\South                Computer bottheory.local             AD Computer     WSCTXEDS002.bottheory.local
    Calgary                Citrix\Calgary                       Folder                               AD OU
    WSCTXCGY001            Citrix\Calgary                       Computer bottheory.local             AD Computer     WSCTXCGY001.bottheory.local
    WSCTXCGY002            Citrix\Calgary                       Computer bottheory.local             AD Computer     WSCTXCGY002.bottheory.local
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
    $BuildCUTreeParams.Add("SiteName",$Site)
}

if ($Force){
    $BuildCUTreeParams.Add("Force",$true)
}

if ($batchCreateFolders){
    $BuildCUTreeParams.Add("batchCreateFolders",$true)
}

[int]$errorCount = Build-CUTree -ExternalTree $ADEnvironment @BuildCUTreeParams

Exit $errorCount

#endregion WVD
