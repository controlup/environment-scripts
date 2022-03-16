[CmdletBinding()]
Param(
	[Parameter(Mandatory=$false, HelpMessage='Separate AD groups with Comma.' )]
    [array] $Groups,
	[Parameter(Mandatory=$false, HelpMessage='Export File to current DIR for full path.' )]
    [string]$ExportFile,
	[Parameter(Mandatory=$false, HelpMessage='If the user account does not have an email address, this will skip the input prompt.' )]
    [switch]$IgnoreEmptyEmail
)

$global:Users = New-Object -TypeName System.Collections.Generic.List[PSObject]
class UserObject{
    [string]$upn
    [string]$fname
    [string]$lname
    [string]$email
    [string]$samaccountname
    [string]$DNSName
        UserObject ([String]$upn,[string]$fname,[string]$lname,[string]$email,[string]$samaccountname,[string]$DNSName) {
        $this.upn = $upn
        $this.FName = $fname
        $this.LName = $lname
        $this.Email = $email
        $this.SAMAccountName = $samaccountname
        $this.DNSName = $DNSName
    }
}

$global:ignoreEmail = $IgnoreEmptyEmail
Add-Type -AssemblyName System.Windows.Forms
$saveAs = New-Object System.Windows.Forms.SaveFileDialog -Property @{
	InitialDirectory = [Environment]::GetFolderPath('Desktop')
	Filter = 'Comma Separated Values (*.csv)|*.csv'
}

$rootgroups = @()
if(!$groups){$groups = read-host "Enter AD Group Name to Sync"}
if($groups){
foreach ($groupname in $groups){
	$filter = "(&(objectClass=group)(|(Name=$groupname)(CN=$groupname)))"
	$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().name
	$rootEntry = New-Object System.DirectoryServices.DirectoryEntry
	$searcher= [adsisearcher]([adsi]"LDAP://$($rootentry.distinguishedName)")
	$searcher.Filter = $filter
	$searcher.SearchScope = "Subtree"
	$searcher.PageSize = 100000
	$rootGroups += $searcher.FindOne().properties
}
}else{write-host "There was no group listed, please input a group name" -ForegroundColor red -BackgroundColor black}

function recurseGroups{
	Param([string]$DN)
	if (!([adsi]::Exists("LDAP://$DN"))) {write-host "$DN does not exist" -ForegroundColor red -BackgroundColor black;return}
	write-host "Processing $DN" -ForegroundColor green
	$group = [adsi]("LDAP://$DN")
	($group).member | ForEach-Object {
		$groupObject = [adsisearcher]"(&(distinguishedname=$($_)))"  
		$groupObjectProps = $groupObject.FindAll().Properties
		if ($groupObjectProps.objectcategory -like "CN=group*"){ 
			recurseGroups $_
		} else {
			$userenabled = ($groupObjectProps.useraccountcontrol[0] -band 2) -ne 2
			if ($userenabled) {
				$dnc = $groupObjectProps.distinguishedname.replace("DC=",".")
				$mail = $null
				$mail = $groupObjectProps.mail
				$groupObjectProps.mail
				if(!$groupObjectProps.mail){
					if(!$global:ignoreEmail){$mail = read-host "Unable to find email for user '$($groupObjectProps.userprincipalname)' Enter email or leave blank to skip"}
				}
				if($mail){$global:Users.Add([UserObject]::new($groupObjectProps.userprincipalname,$groupObjectProps.givenname,$groupObjectProps.sn,$mail,$groupObjectProps.samaccountname,$($dnc.substring($dnc.indexof(",.")+2).replace(",.","."))))}
				
			}              
		}
	}
}
foreach ($rootGroup in $rootGroups){recurseGroups $($rootGroup.distinguishedname)}

if($exportFile){
	$global:Users|export-csv $exportFile -NoTypeInformation
}else{
	$saveDialog = [System.Windows.Forms.MessageBox]::Show("Would you like to save?." , "Save Confirmation" , 4,32)
}

if($saveDialog){
	$SaveAs.ShowDialog()
	if($saveAs.FileName){$global:Users|export-csv $saveAs.FileName -NoTypeInformation }
}
