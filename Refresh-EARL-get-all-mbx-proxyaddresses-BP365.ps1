


<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.148
	 Created on:   	10/03/2023 14:00
	 Created by:   	Neil Poultney
	 Organization: 	NDP Consultancy Ltd - neil.poultney@ndpconsultancyltd.co.uk
	 Filename:     	get-proxyaddresses-bp365.ps1
	===========================================================================
	.DESCRIPTION
		Exports proxyAddresses across all object types in O365.

		Change Log
		V1.00, 18/03/2020 - Initial full version
		v1.1,  16-02-2022 updated to work with EARL import requirement
		v1.2 testing with v3 commandlets		

	.Known Issues
		none
		
	.DISCLAIMER
		Be aware that all scripts are run at your own risk and while every script has been written with the intention of minimising the potential for unintended consequences
		The author cannot be held responsible for any misuse or script problems.
		
		The software and service is provided by the author "as is" and any express or implied warranties, including, but not limited to,
		the implied warranties of merchantability and fitness for a particular purpose are disclaimed.
		In no event shall the author be liable for any direct, indirect, incidental, special, exemplary, or consequential damages
		(including, but not limited to, procurement of substitute goods or services; loss of use, data, or profits; or business interruption)
		however caused and on any theory of liability, whether in contract, strict liability, or tort (including negligence or otherwise) arising in any way out of the use of this software or service
	#>	


param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeGroupMailboxes,[switch]$IncludeDGs,[switch]$IncludeMailUsers,[switch]$IncludeMailContacts,[switch]$CondensedOutput,[switch]$IncludeSIPAliases,[switch]$IncludeSPOAliases)




# Make Windows negotiate higher TLS version:
[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")

$global:nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"

[System.GC]::Collect()


##workoutwhereweare
$Domainwearein = (Get-WmiObject Win32_ComputerSystem).Name
$whoweare = $ENV:USERNAME
if ($domainwearein -eq "BP1GXEIS801") { $global:Envirionmentchoice = "Dev" }
if ($domainwearein -eq "zneepacp11eme2" -or $domainwearein -eq "zneepacp11emrg") { $global:Envirionmentchoice = "ProdNE" }
if ($domainwearein -eq "zweepacp11emg3" -or $domainwearein -eq "zweepacp11em50") { $global:Envirionmentchoice = "ProdWE" }




if (($Envirionmentchoice -eq "ProdNE") -or ($Envirionmentchoice -eq "ProdWE"))
{
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	$transcriptlog = "H:\EARLTranscripts\ProxyAddresses\proxyAddresses-prod-" + $nowfiledate + ".log"
	
	Start-Transcript -Path $transcriptlog
	
	

	$loglocation = "H:\EARLPSLogs\BulkExports\"
	Set-Variable -Name logfolder -Value $logslocation -Option ReadOnly -Scope Script -Force
	$logfilelocation = $loglocation + "BulkExport-ProxyAddresses-" + $nowfiledate + ".log"
	Set-Variable -Name logfile -Value $logfilelocation -Option ReadOnly -Scope Script -Force
	#$filewatcherlocationout = "Q:\EARL\FileLocation\"
	$filewatcherlocationout = "Q:\EARL\CSVFileLocation\"
	Set-Variable -Name FileWatcherOut -Value $filewatcherlocationout -Option ReadOnly -Scope Script -Force
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Start of Log File"
	add-content $logfile  $now
	add-content $logfile  "Processing in Live environment on $Domainwearein for  $Envirionmentchoice for user $whoweare"
	
}




[Int] $connecttry = 0


Function Connect-EXO
{
	$Error.Clear()
	
	
	$modules = get-module | Select-Object name
	if (!$Modules -like "*ExchangeOnline*")
	{
		
		Import-Module ExchangeOnlineManagement
		
	}
	
	
	Disconnect-ExchangeOnline -Confirm:$false
	Get-PSSession | Remove-PSSession
	
	$PFXPassword = ""
	$cert = ""
	$sessionstate = ""
	$sessionstate1 = ""
	$EXOconnectionavailble1 = ""
	$EXOconnectionavailble2 = ""
	$EXOconnectionavailble3 = ""
	$EXOconnectionavailble4 = ""
	
	
	
	if (($Envirionmentchoice -eq "ProdWE") -or ($Envirionmentchoice -eq "ProdNE"))
	{
		
		if ($failconnect -le "10")
		{
			
			#APP Connection EARL EXO Reports
			#Connect-ExchangeOnline -CertificateThumbprint "f658b65fe915b1204cfeefe399259333f744c315" -AppID "920938ea-809a-4a52-bf9e-0ae65fd12d53" -Organization "bp365.onmicrosoft.com" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
			Connect-ExchangeOnline -CertificateThumbprint "a98251f44faf329cd3d1474f1440aca8356edaa0" -AppID "920938ea-809a-4a52-bf9e-0ae65fd12d53" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Get-Recipient,Get-EXORecipient,Get-DistributionGroup,Get-ConnectionInformation" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErrr
$time = get-date -Format dd-MM-yyyy-HH:mm:ss
			
			
			$sessionstateA = Get-ConnectionInformation | select-Object *
			
			
			
			$EXOconnectionavailableA = $sessionstateA.TokenStatus
			$EXOconnectionavailableB = $sessionstateA.Name
			
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "TokenStatus : $EXOconnectionavailableA | Session Name: $EXOconnectionavailableB .... | $now "
			
			if (($EXOconnectionavailableA -eq "Active") -and ($EXOconnectionavailableB -match "ExchangeOnline*"))
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Connected to Exchange Online EXO V3 APP connection EARL EXO Reports (thumbprint).... | $now "
				$failconnect = 0
			}
			
			if (($EXOconnectionavailableA -ne "Active") -and ($EXOconnectionavailableB -notmatch "ExchangeOnline*"))
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Unable to Connect to  Exchange Online EXO V3 EARL EXO Reports APP(thumbprint).... pausing for 5 minutes | $now"
				Start-Sleep -Seconds 600
				$failconnect = $failconnect + 1
				Connect-EXO
			}
			
		}
		
	}
	
	if ($failconnect -gt "9")
	{
		
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Tried to connect to EXO Powershell 10 times and failed so aborting using V3 EARL EXO Reports APP(thumbprint).... | $now"
	}
	
}

[int]$failconnect = 0


function Disconnect-EXO
{
	
	Disconnect-ExchangeOnline -Confirm:$false
	$now = get-date -Format dd-MM-yyyy-HH:mm:ss
	Add-Content $logfile "Disconnected From Exchange Online Remote Powershell  .... Time: $now"
}




Function Map-Filewatcher
{
	$connectionok = "False"
	$testdrives = Get-PsDrive | select Name
	foreach ($drive in $drives)
	{
		if ($drive -match "Q")
		{
			$connectionok = "True"
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			add-content $logfile  "Connected to EARL Filewatcher Q already | $now"
		}
		
	}
	
	if ($connectionok -eq "False")
	{
		$connectTestResult = Test-NetConnection -ComputerName earlfiles.file.core.windows.net -Port 445
		if ($connectTestResult.TcpTestSucceeded)
		{
			
			if ($Envirionmentchoice -eq "ProdWE")
			{
				cmd.exe /C "cmdkey /add:`"earlfiles.file.core.windows.net`" /user:`"localhost\earlfiles`" /pass:`"sN3o0NyMn5VwSyLFP6EvhIR/siBc8uHm/mfa196up7hZDJnjLr36Op7vWfqGOtayEKKcjEZSicY5pSb6Rx9YoQ==`""
				# Mount the drive
				New-PSDrive -Name Q -PSProvider FileSystem -Scope Script -Root "\\earlfiles.file.core.windows.net\filewatcher\filewatcherwe\"
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Connected to EARL Filewatcher WE Q | $now"
			}
			
			if ($Envirionmentchoice -eq "ProdNE")
			{
				cmd.exe /C "cmdkey /add:`"earlfiles.file.core.windows.net`" /user:`"localhost\earlfiles`" /pass:`"sN3o0NyMn5VwSyLFP6EvhIR/siBc8uHm/mfa196up7hZDJnjLr36Op7vWfqGOtayEKKcjEZSicY5pSb6Rx9YoQ==`""
				# Mount the drive
				New-PSDrive -Name Q -PSProvider FileSystem -Scope Script -Root "\\earlfiles.file.core.windows.net\filewatcher\filewatcherne\"
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Connected to EARL FilewatcherNE Q | $now"
			}
			
		}
		else
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			add-content $logfile  "Cannot Connect to EARL Filewatcher - critical error : Unable to reach the Azure storage account via port 445.\\earlfiles.file.core.windows.net\filewatcher | $now"
			if ($FWConnectionTry -gt 10)
			{
				add-content $logfile  "Cannot Connect to EARL Filewatcher - critical error. logs only created locally  | $now"
				
				
			}
			
			Start-Sleep -Seconds 360
			$FWConnectionTry = $FWConnectionTry + 1
			Map-Filewatcher
		}
		
	}
}

function RemoveFilewatcher
{
	#Map-Filewatcher
	#RemoveFilewatcher
	#Map-Logdrive
	#RemoveLogdrive
	$testdrives = ""
	
	$testdrives = Get-PsDrive | Select-Object *
	
	
	foreach ($drive in $testdrives)
	{
		$drivename = $drive.Name
		if ($drive -match "Q")
		{
			Get-PSDrive Q | Remove-PSDrive
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			add-content $logfile  "disconnected Filewatcher Q | $now"
		}
		
		
	}
	
	#Get-PSDrive Q | Remove-PSDrive
	
}



Add-Type -As System.IO.Compression.FileSystem

function New-ZipFile {
	#.Synopsis
	#  Create a new zip file, optionally appending to an existing zip...
	[CmdletBinding()]
	param(
		# The path of the zip to create
		[Parameter(Position=0, Mandatory=$true)]
		$ZipFilePath,
 
		# Items that we want to add to the ZipFile
		[Parameter(Position=1, Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("PSPath","Item")]
		[string[]]$InputObject = $Pwd,
 
		# Append to an existing zip file, instead of overwriting it
		[Switch]$Append,
 
		# The compression level (defaults to Optimal):
		#   Optimal - The compression operation should be optimally compressed, even if the operation takes a longer time to complete.
		#   Fastest - The compression operation should complete as quickly as possible, even if the resulting file is not optimally compressed.
		#   NoCompression - No compression should be performed on the file.
		[System.IO.Compression.CompressionLevel]$Compression = "Optimal"
	)
	begin {
		# Make sure the folder already exists
		[string]$File = Split-Path $ZipFilePath -Leaf
		[string]$Folder = $(if($Folder = Split-Path $ZipFilePath) { Resolve-Path $Folder } else { $Pwd })
		$ZipFilePath = Join-Path $Folder $File
		# If they don't want to append, make sure the zip file doesn't already exist.
		if(!$Append) {
			if(Test-Path $ZipFilePath) { Remove-Item $ZipFilePath }
		}
		$Archive = [System.IO.Compression.ZipFile]::Open( $ZipFilePath, "Update" )
	}
	process {
		foreach($path in $InputObject) {
			foreach($item in Resolve-Path $path) {
				# Push-Location so we can use Resolve-Path -Relative
				Push-Location (Split-Path $item)
				# This will get the file, or all the files in the folder (recursively)
				foreach($file in Get-ChildItem $item -Recurse -File -Force | % FullName) {
					# Calculate the relative file path
					$relative = (Resolve-Path $file -Relative).TrimStart(".\")
					# Add the file to the zip
					$null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, $file, $relative, $Compression)
				}
				Pop-Location
			}
		}
	}
	end {
		$Archive.Dispose()
		Get-Item $ZipFilePath
	}
}
     
     
function Expand-ZipFile {
	#.Synopsis
	#  Expand a zip file, ensuring it's contents go to a single folder ...
	[CmdletBinding()]
	param(
		# The path of the zip file that needs to be extracted
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=0, Mandatory=$true)]
		[Alias("PSPath")]
		$FilePath,
 
		# The path where we want the output folder to end up
		[Parameter(Position=1)]
		$OutputPath = $Pwd,
 
		# Make sure the resulting folder is always named the same as the archive
		[Switch]$Force
	)
	process {
		$ZipFile = Get-Item $FilePath
		$Archive = [System.IO.Compression.ZipFile]::Open( $ZipFile, "Read" )
 
		# Figure out where we'd prefer to end up
		if(Test-Path $OutputPath) {
			# If they pass a path that exists, we want to create a new folder
			$Destination = Join-Path $OutputPath $ZipFile.BaseName
		} else {
			# Otherwise, since they passed a folder, they must want us to use it
			$Destination = $OutputPath
		}
 
		# The root folder of the first entry ...
		$ArchiveRoot = ($Archive.Entries[0].FullName -Split "/|\\")[0]
 
		Write-Verbose "Desired Destination: $Destination"
		Write-Verbose "Archive Root: $ArchiveRoot"
 
		# If any of the files are not in the same root folder ...
		if($Archive.Entries.FullName | Where-Object { @($_ -Split "/|\\")[0] -ne $ArchiveRoot }) {
			# extract it into a new folder:
			New-Item $Destination -Type Directory -Force
			[System.IO.Compression.ZipFileExtensions]::ExtractToDirectory( $Archive, $Destination )
		} else {
			# otherwise, extract it to the OutputPath
			[System.IO.Compression.ZipFileExtensions]::ExtractToDirectory( $Archive, $OutputPath )
 
			# If there was only a single file in the archive, then we'll just output that file...
			if($Archive.Entries.Count -eq 1) {
				# Except, if they asked for an OutputPath with an extension on it, we'll rename the file to that ...
				if([System.IO.Path]::GetExtension($Destination)) {
					Move-Item (Join-Path $OutputPath $Archive.Entries[0].FullName) $Destination
				} else {
					Get-Item (Join-Path $OutputPath $Archive.Entries[0].FullName)
				}
			} elseif($Force) {
				# Otherwise let's make sure that we move it to where we expect it to go, in case the zip's been renamed
				if($ArchiveRoot -ne $ZipFile.BaseName) {
					Move-Item (join-path $OutputPath $ArchiveRoot) $Destination
					Get-Item $Destination
				}
			} else {
				Get-Item (Join-Path $OutputPath $ArchiveRoot)
			}
		}
 
		$Archive.Dispose()
	}
}

# Add the aliases ZIP and UNZIP
new-alias zip new-zipfile
new-alias unzip expand-zipfile




function Get-EmailAddressesInventory {
<#
.Synopsis
    Lists all aliases for all recipients of the selected type(s).
.DESCRIPTION
    The Get-EmailAddressesInventory cmdlet finds all recipients of the selected type(s) and lists their email and/or non-email aliases.
    Running the cmdlet without parameters will return entries for User mailboxes only. Specifying particular recipient type(s) can be done with the corresponding switch parameter.
    To use condensed output (one line per recipient), use the CondensedOutput switch.
    To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.

.EXAMPLE
    .\O365_aliases_inventory.ps1 -IncludeUserMailboxes

    This command will return a list of email aliases for all user mailboxes.

.EXAMPLE
    Get-EmailAddressesInventory -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "accessrights.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the recipient type and its aliases.
#>

    [CmdletBinding()]
    
    Param
    (
    #Specify whether to include User mailboxes in the result
    [Switch]$IncludeUserMailboxes,    
    #Specify whether to include Shared mailboxes in the result
    [Switch]$IncludeSharedMailboxes,
    #Specify whether to include Room and Equipment mailboxes in the result
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to include Group mailboxes in the result
    [Switch]$IncludeGroupMailboxes,
    #Specify whether to include Distribution Groups, Dynamic Distribution Groups, Room Lists and Mail-enabled Security Groups in the result
    [Switch]$IncludeDGs,
    #Specify whether to include Mail Users and Guest Mail Users in the result
    [switch]$IncludeMailUsers,
    #Specify whether to include Mail Contacts in the result
    [switch]$IncludeMailContacts,
    #Specify whether to return all recipient types in the result
    [Switch]$IncludeAll,
    #Specify whether to write the output in condensed format
    [Switch]$CondensedOutput,
    #Specify whether to include SIP/EUM aliases in the output
    [switch]$IncludeSIPAliases,
    #Specify whether to include SPO aliases in the output
    [switch]$IncludeSPOAliases)

    
    #Initialize the variable used to designate recipient types, based on the input parameters
	$included = @()
	
	
	if ($IncludeUserMailboxes)
	{
		$included += "UserMailbox"
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\UserO365TempProxyaddressTable_" + $filedate + ".csv"
		$outzipfile = "H:\M365Reports\UserMailboxesO365Emailaddresses_" + $filedate + ".zip"
		if (test-path $outzipfile) { rm $outzipfile }
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for User Mailboxes.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force
		
		
	}
	
	
	
    if($IncludeSharedMailboxes) { 
$included += "SharedMailbox"
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\SharedO365TempProxyAddressTable_" + $filedate + ".csv"
		$outzipfile = "H:\M365Reports\SharedMailboxesO365Emailaddresses_" + $filedate + ".zip"
		if (test-path $outzipfile) { rm $outzipfile }
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for shared mailboxes.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force

}
    if($IncludeRoomMailboxes) {
 $included += "RoomMailbox","EquipmentMailbox"
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss

$outfilecsv = "H:\M365Reports\RoomO365TempProxyAddressTable_" +$filedate+ ".csv"
			$outzipfile = "H:\M365Reports\RoomMailboxesO365Emailaddresses_" + $filedate + ".zip"
	if (test-path $outzipfile) { rm $outzipfile }
		
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for room mailboxes.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force
}
    if($IncludeMailUsers) { 
$included += "MailUser"
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\MailUsersO365TempProxyAddressTable_" + $filedate + ".csv"
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for mailusers.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force

}
    if($IncludeMailContacts) { $included += "MailContact"}
	if ($IncludeGroupMailboxes)
	{
		
		$included += "GroupMailbox"
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\GroupsO365TempProxyAddressTable_" + $filedate + ".csv"
		$outzipfile = "H:\M365Reports\O365GroupsEmailaddresses_" + $filedate + ".zip"
		if (test-path $outzipfile) { rm $outzipfile }
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for O365 group mailboxes.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force
		
	}
    if($IncludeDGs) { 
$included += 'DynamicDistributionGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'RoomList'
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\GroupsO365TempProxyAddressTable_" + $filedate + ".csv"
			$outzipfile = "H:\M365Reports\GroupsO365Emailaddresses_" + $filedate + ".zip"
	if (test-path $outzipfile) { rm $outzipfile }
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for Distribution groups.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force

}

if($IncludeAll) {
$Included = @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','TeamMailbox','DiscoveryMailbox','MailUser','MailContact','DynamicDistributionGroup','MailUniversalDistributionGroup','MailUniversalSecurityGroup','RoomList','GuestMailUser','GroupMailbox')
		
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\All_O365TempProxyAddressTable_" + $filedate + ".csv"
	if (test-path $outfile) { rm $outfile }	
		$outzipfile = "H:\M365Reports\All_O365-Emailaddresses_" + $filedate + ".zip"
	if (test-path $outzipfile) { rm $outzipfile }
		Write-Host "Exporting proxyaddresses for All Email Addresses"
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force

}
    
Connect-EXO


	$Checksessionstate = ""
	$EXOconnectionavailableA1 = ""
	$EXOconnectionavailableB1 = ""
	
	$Checksessionstate = Get-ConnectionInformation
	
	$EXOconnectionavailableA1 = $Checksessionstate.TokenStatus
	$EXOconnectionavailableB1 = $Checksessionstate.Name
	
	$Checksessionstate = ""
	if ((($EXOconnectionavailableA1 -ne "Active") -and ($EXOconnectionavailableB1 -notlike "ExchangeOnline*")) -or (!$Checksessionstate))
	{
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "`tNo Exchange online remote access |  Token status : $EXOconnectionavailableA1 | Connection Name : $EXOconnectionavailableB1  | $now "
		
		Connect-EXO
	}
	
	#Get a minimal set of properties for the selected recipients. Make sure to add any additional properties you want included in the report to the list here!
   if($IncludeUserMailboxes) {$MBList = Get-EXORecipient -ResultSize Unlimited -Properties Displayname,PrimarySMTPAddress,WindowsLiveID,EmailAddresses,RecipientTypeDetails -RecipientTypeDetails UserMailbox }

    if($IncludeSharedMailboxes) {$MBList = Get-EXORecipient -ResultSize Unlimited -Properties Displayname,PrimarySMTPAddress,WindowsLiveID,EmailAddresses,RecipientTypeDetails -RecipientTypeDetails SharedMailbox }
	
	if ($IncludeRoomMailboxes) { $MBList = Get-EXORecipient -ResultSize Unlimited -Properties Displayname, PrimarySMTPAddress, WindowsLiveID, EmailAddresses, RecipientTypeDetails -RecipientTypeDetails RoomMailbox }
	
	if ($IncludeGroupMailboxes) { $MBList = Get-EXORecipient -ResultSize Unlimited -Properties Displayname, PrimarySMTPAddress, WindowsLiveID, EmailAddresses, RecipientTypeDetails -RecipientTypeDetails Groupmailbox }
	
	if ($IncludeDGs) { $MBList = Get-EXORecipient -ResultSize Unlimited -Properties Displayname, PrimarySMTPAddress, WindowsLiveID, EmailAddresses, RecipientTypeDetails -RecipientTypeDetails MailUniversalSecurityGroup, MailUniversalDistributionGroup, DynamicDistributionGroup }
	
	if ($IncludeMailUsers) { $MBList = Get-EXORecipient -ResultSize Unlimited -Properties Displayname, PrimarySMTPAddress, WindowsLiveID, EmailAddresses, RecipientTypeDetails -RecipientTypeDetails MailUser }
	
if($IncludeAll) {
write-host "Getting all email addresses available"

$MBList = Get-EXORecipient -ResultSize Unlimited -Properties Displayname,PrimarySMTPAddress,WindowsLiveID,EmailAddresses,RecipientTypeDetails

}
	
	
	
$foundcount = $MBList.count
    #If no recipients are returned from the above cmdlet, stop the script and inform the user
    if (!$MBList) { Write-Error "No recipients of the specified types were found, specify different criteria." -ErrorAction Stop}
#if (($foundcount -lt 75000) -and ($included -eq "UserMailbox")) {
#Write-Host "Not enough user mailbox recipients  were found, trying again."
#$MBList = Get-EXORecipient -ResultSize Unlimited -Properties Displayname,PrimarySMTPAddress,WindowsLiveID,EmailAddresses,RecipientTypeDetails -RecipientTypeDetails UserMailbox 
#} 

    #Once we have the recipient list, cycle over each recipient to prepare the output object
    $arrAliases = @()

$outputcount = $MBList.count

	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	Add-Content $logfile "Found $outputcount for export of proxyaddresses for the types $included . Now creating output :::$time"

[int]$u = 1

$elapsed = 0
[int]$i = 0
$perusersum = 0

    foreach ($MB in $MBList) {

$mbxbname = $MB.DisplayName
    $peruser = [Diagnostics.Stopwatch]::StartNew()
    $formin = ($i / $outputcount) * 100
    $formated = "{0:N2}" -f $formin


#Write-Progress -activity "Processing mailbox: $mbxbname" -status "${u} of ${outputcount}" -percentComplete (($u / $outputcount) * 100)

        #If we want condensed output, one line per recipient
        if ($CondensedOutput) {
            $objAliases = New-Object PSObject
            $i++;Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient" -Value $MB.DisplayName
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Primary SMTP address" -Value $MB.PrimarySMTPAddress
            #we use WindowsLiveID as a workaround to get the UPN, as Get-Recipient does not return the UPN property
            if ($MB.WindowsLiveID) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "UPN" -Value $MB.WindowsLiveID }
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Email Aliases" -Value (($MB.EmailAddresses | ? {$_.Prefix -eq "SMTP" -or $_.Prefix -eq "X500"}).ProxyAddressString -join ";")

            #Handle SIP/SPO aliases and external email address depending on the parameters provided
            if ($IncludeSIPAliases -or $IncludeAll) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "SIP Aliases" -Value (($MB.EmailAddresses | ? {$_.Prefix -eq "SIP" -or $_.Prefix -eq "EUM"}).ProxyAddressString -join ";") }
            if ($IncludeSPOAliases -or $IncludeAll) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "SPO Aliases" -Value (($MB.EmailAddresses | ? {$_.Prefix -eq "SPO"}).ProxyAddressString -join ";") }
           # if ($IncludeMailUsers -or $IncludeMailContacts -or $IncludeAll -or $IncludeUserMailboxes -or $IncludeSharedMailboxes -or $IncludeRoomMailboxes -or $IncludeDGs -or $IncludeGroupMailboxes) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "External email address" -Value $MB.ExternalEmailAddress }

            $arrAliases += $objAliases
        }
        #Otherwise, write each proxy entry on separate line
        else {
            foreach ($entry in $MB.EmailAddresses) {



		if (($entry.startswith("smtp:")) -or ($entry.startswith("SMTP:")))  {
					$objAliases = New-Object PSObject
					
					$outentry = $entry.substring(5)
               # Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "DisplayName" -Value $MB.DisplayName
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "PrimarySMTPAddress" -Value $MB.PrimarySMTPAddress
                #we use WindowsLiveID as a workaround to get the UPN, as Get-Recipient does not return the UPN property
                #if ($MB.WindowsLiveID) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "UPN" -Value $MB.WindowsLiveID }
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $MB.RecipientTypeDetails
                #Handle SIP/SPO aliases depending on the parameters provided
                #if (($entry.Prefix -eq "SIP" -or $entry.Prefix -eq "EUM") -and !($IncludeSIPAliases -or $IncludeAll)) { continue }
                #if ($entry.Prefix -eq "SPO" -and !($IncludeSPOAliases -or $IncludeAll)) { continue }
                #Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Aliases" -Value $entry.ProxyAddressString
					Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "EmailAddresses" -Value $outentry
		
		#Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "ExternalEmailAddress" -Value $MB.ExternalEmailAddress               
                $arrAliases += $objAliases



}
            }

            #Handle External email address for Mail User/Mail Contact objects
           # if (($IncludeMailUsers -or $IncludeMailContacts -or $IncludeAll) -and $MB.ExternalEmailAddress) {
              #  if ($MB.ExternalEmailAddress.AddressString -eq $MB.PrimarySMTPAddress) { continue }
               # $objAliases = New-Object PSObject
               # Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Number" -Value $i
               # Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient" -Value $MB.DisplayName
               # Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Primary SMTP address" -Value $MB.PrimarySMTPAddress
               # Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails
               # Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Aliases" -Value $MB.ExternalEmailAddress
              #  $arrAliases += $objAliases
          #  }
        }
$i++
  $perusersum += $peruser.ElapsedMilliseconds
    $peruseravg = $perusersum / $i
    $stillneeded = $outputcount - $i
    $remain = (($peruseravg * $stillneeded) / 1000)
    write-progress -activity "Processing mailbox: $mbxbname" -status "${i} of ${outputcount} ($formated%) Remaining: $((New-TimeSpan -Seconds $remain).ToString()) Average(ms): $($peruseravg.ToString("#.##"))" -percentcomplete $([Math]::Round(($i / $outputcount * 100))) -id 1
    }
    #Output the result to the console host. Rearrange/sort as needed.
   # $arrAliases | select * -ExcludeProperty Number

$arrAliases | Export-Csv $outfile -nti -encoding UTF8 -delimiter '|'

write-host "Exported csv file to $outfile" 
#new-zipfile $outzipfile $outfile
}

#Invoke the Get-EmailAddressesInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!

[System.GC]::Collect()


$start = get-date -Format dd-MM-yyyy-hh:mm
$StopWatch = New-Object System.Diagnostics.Stopwatch
$StopWatch.Start()

Write-Host "The Script Started at $start"

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "The Script Started at $start"

Get-EmailAddressesInventory @PSBoundParameters
#Get-EmailAddressesInventory @PSBoundParameters -OutVariable global:varAliases | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_EmailAddresses.csv" -NoTypeInformation -Encoding UTF8 -UseCulture

<#

Try
{
	Map-Filewatcher
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	Add-Content $logfile "`tFile $outfile trying to copy to $filewatcherout | $now"
	Copy-item -path $outfile -destination $filewatcherout
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	Add-Content $logfile "`tFile Copied to FileWatcher $outfile to $filewatcherout | $now"
	
	#Map-Filewatcher
	RemoveFilewatcher
}
catch
{
	Start-Sleep -s 20
	try
	{
		Map-Filewatcher
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	Add-Content $logfile "`tFile $outfile trying to copy to $filewatcherout [try2] | $now"
		Copy-item -path $outfile -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "`tFile Copied to FileWatcher $outfile to $filewatcherout | $now"
		#Map-Filewatcher
		RemoveFilewatcher
	}
	catch
	{
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Cannot copy file to FileWatcher $outfile | $now"
	}
}

#>

$finished = get-date -Format dd-MM-yyyy-HH:mm
$StopWatch.Stop()
$ElapsedTime = $StopWatch.Elapsed

$scripthours = $ElapsedTime.Hours
$scriptminutes = $ElapsedTime.Minutes

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "The Script Finished at $finished"
Add-Content $logfile "The script took" $ElapsedTime.Hours "hours," $ElapsedTime.Minutes "minutes, and" $ElapsedTime.Seconds "seconds to run."

#.\EARL-get-all-mbx-proxyaddresses-BP365.ps1 -IncludeMailUsers
#.\EARL-get-all-mbx-proxyaddresses-BP365.ps1 -IncludeUserMailboxes
#.\O365_aliases_inventory1.ps1 -IncludeDGs
#.\O365_aliases_inventory1.ps1 -IncludeSharedMailboxes
#.\O365_aliases_inventory1.ps1 -IncludeGroupMailboxes

<# None, UserMailbox, LinkedMailbox, SharedMailbox, LegacyMailbox, RoomMailbox, EquipmentMailbox, MailContact, MailUser,
MailUniversalDistributionGroup, MailNonUniversalGroup, MailUniversalSecurityGroup, DynamicDistributionGroup, PublicFolder, SystemAttendantMailbox, SystemMailbox,
MailForestContact, User, Contact, UniversalDistributionGroup, UniversalSecurityGroup, NonUniversalGroup, DisabledUser, MicrosoftExchange, ArbitrationMailbox, MailboxPlan,
LinkedUser, RoomList, DiscoveryMailbox, RoleGroup, RemoteUserMailbox, Computer, RemoteRoomMailbox, RemoteEquipmentMailbox, RemoteSharedMailbox, PublicFolderMailbox, TeamMailbox,
RemoteTeamMailbox, MonitoringMailbox, GroupMailbox, LinkedRoomMailbox, AuditLogMailbox, RemoteGroupMailbox, SchedulingMailbox, GuestMailUser, AuxAuditLogMailbox,
SupervisoryReviewPolicyMailbox, ExchangeSecurityGroup, AllUniqueRecipientTypes""
#>
Disconnect-EXO

	Stop-Transcript
	EXIT-PSSession
	Exit
