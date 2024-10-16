


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


param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeGroupMailboxes,[switch]$IncludeGuests,[switch]$IncludeDGs,[switch]$IncludeMailUsers,[switch]$IncludeMailContacts,[switch]$CondensedOutput,[switch]$IncludeSIPAliases,[switch]$IncludeSPOAliases)




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
	$transcriptlog = "H:\EARLTranscripts\ProxyAddresses\TempproxyAddresses-prod-" + $nowfiledate + ".log"
	
	Start-Transcript -Path $transcriptlog
	
	

	$loglocation = "H:\EARLPSLogs\BulkExports\"
	Set-Variable -Name logfolder -Value $logslocation -Option ReadOnly -Scope Script -Force
	$logfilelocation = $loglocation + "BulkExport-TempProxyAddresses-" + $nowfiledate + ".log"
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
			#Connect-ExchangeOnline -CertificateThumbprint "a98251f44faf329cd3d1474f1440aca8356edaa0" -AppID "920938ea-809a-4a52-bf9e-0ae65fd12d53" -Organization "bp365.onmicrosoft.com" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
			Connect-ExchangeOnline -CertificateThumbprint "8f901a3fbc0f30746f1f5309806314aa32841e2b" -AppID "920938ea-809a-4a52-bf9e-0ae65fd12d53" -Organization "bp365.onmicrosoft.com" -ShowProgress  -SkipLoadingCmdletHelp $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
			

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
				add-content $logfile  "Connected to EARL FilewatcherWE Q | $now"
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
	#Specify whether to include Guest Mail Users in the result
	[switch]$IncludeGuests,
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
		$outfilecsv1 = "H:\M365Reports\UserO365TempProxyaddressTable_1_" + $filedate + ".csv"
		$outfilecsv2 = "H:\M365Reports\UserO365TempProxyaddressTable_2_" + $filedate + ".csv"
		$outfilecsv3 = "H:\M365Reports\UserO365TempProxyaddressTable_3_" + $filedate + ".csv"
		$outfilecsv4 = "H:\M365Reports\UserO365TempProxyaddressTable_4_" + $filedate + ".csv"
		$outfilecsv5 = "H:\M365Reports\UserO365TempProxyaddressTable_5_" + $filedate + ".csv"
		$outfilecsv6 = "H:\M365Reports\UserO365TempProxyaddressTable_6_" + $filedate + ".csv"
		$outfilecsv7 = "H:\M365Reports\UserO365TempProxyaddressTable_7_" + $filedate + ".csv"
		$outfilecsv8 = "H:\M365Reports\UserO365TempProxyaddressTable_8_" + $filedate + ".csv"
		$outfilecsv9 = "H:\M365Reports\UserO365TempProxyaddressTable_9_" + $filedate + ".csv"
		$outfilecsv10 = "H:\M365Reports\UserO365TempProxyaddressTable_10_" + $filedate + ".csv"
		$outfilecsv11 = "H:\M365Reports\UserO365TempProxyaddressTable_11_" + $filedate + ".csv"
		$outfilecsv12 = "H:\M365Reports\UserO365TempProxyaddressTable_12_" + $filedate + ".csv"
		$outfilecsv13 = "H:\M365Reports\UserO365TempProxyaddressTable_13_" + $filedate + ".csv"
		$outfilecsv14 = "H:\M365Reports\UserO365TempProxyaddressTable_14_" + $filedate + ".csv"
		$outfilecsv15 = "H:\M365Reports\UserO365TempProxyaddressTable_15_" + $filedate + ".csv"
		$outfilecsv16 = "H:\M365Reports\UserO365TempProxyaddressTable_16_" + $filedate + ".csv"
		$outfilecsv17 = "H:\M365Reports\UserO365TempProxyaddressTable_17_" + $filedate + ".csv"
		$outfilecsv18 = "H:\M365Reports\UserO365TempProxyaddressTable_18_" + $filedate + ".csv"
		$outfilecsv19 = "H:\M365Reports\UserO365TempProxyaddressTable_19_" + $filedate + ".csv"
		$outfilecsv20 = "H:\M365Reports\UserO365TempProxyaddressTable_20_" + $filedate + ".csv"
		$outfilecsv21 = "H:\M365Reports\UserO365TempProxyaddressTable_21_" + $filedate + ".csv"
		$outfilecsv22 = "H:\M365Reports\UserO365TempProxyaddressTable_22_" + $filedate + ".csv"
		$outfilecsv23 = "H:\M365Reports\UserO365TempProxyaddressTable_23_" + $filedate + ".csv"
		$outfilecsv24 = "H:\M365Reports\UserO365TempProxyaddressTable_24_" + $filedate + ".csv"
		$outfilecsv25 = "H:\M365Reports\UserO365TempProxyaddressTable_25_" + $filedate + ".csv"
		$outfilecsv26 = "H:\M365Reports\UserO365TempProxyaddressTable_26_" + $filedate + ".csv"
		$outfilecsv27 = "H:\M365Reports\UserO365TempProxyaddressTable_27_" + $filedate + ".csv"
		$outfilecsv28 = "H:\M365Reports\UserO365TempProxyaddressTable_28_" + $filedate + ".csv"
		$outfilecsv29 = "H:\M365Reports\UserO365TempProxyaddressTable_29_" + $filedate + ".csv"
		$outfilecsv30 = "H:\M365Reports\UserO365TempProxyaddressTable_30_" + $filedate + ".csv"
		$outfilecsv31 = "H:\M365Reports\UserO365TempProxyaddressTable_31_" + $filedate + ".csv"
		$outfilecsv32 = "H:\M365Reports\UserO365TempProxyaddressTable_32_" + $filedate + ".csv"
		$outfilecsv33 = "H:\M365Reports\UserO365TempProxyaddressTable_33_" + $filedate + ".csv"
		$outfilecsv34 = "H:\M365Reports\UserO365TempProxyaddressTable_34_" + $filedate + ".csv"
		$outfilecsv35 = "H:\M365Reports\UserO365TempProxyaddressTable_35_" + $filedate + ".csv"
		$outfilecsv36 = "H:\M365Reports\UserO365TempProxyaddressTable_36_" + $filedate + ".csv"
		$outfilecsv37 = "H:\M365Reports\UserO365TempProxyaddressTable_37_" + $filedate + ".csv"
		$outfilecsv38 = "H:\M365Reports\UserO365TempProxyaddressTable_38_" + $filedate + ".csv"
		$outfilecsv39 = "H:\M365Reports\UserO365TempProxyaddressTable_39_" + $filedate + ".csv"
		$outfilecsv40 = "H:\M365Reports\UserO365TempProxyaddressTable_40_" + $filedate + ".csv"
		$outfilecsv41 = "H:\M365Reports\UserO365TempProxyaddressTable_41_" + $filedate + ".csv"
		$outfilecsv42 = "H:\M365Reports\UserO365TempProxyaddressTable_42_" + $filedate + ".csv"
		$outfilecsv43 = "H:\M365Reports\UserO365TempProxyaddressTable_43_" + $filedate + ".csv"
		$outfilecsv44 = "H:\M365Reports\UserO365TempProxyaddressTable_44_" + $filedate + ".csv"
		$outfilecsv45 = "H:\M365Reports\UserO365TempProxyaddressTable_45_" + $filedate + ".csv"
		$outfilecsv46 = "H:\M365Reports\UserO365TempProxyaddressTable_46_" + $filedate + ".csv"
		$outfilecsv47 = "H:\M365Reports\UserO365TempProxyaddressTable_47_" + $filedate + ".csv"
		$outfilecsv48 = "H:\M365Reports\UserO365TempProxyaddressTable_48_" + $filedate + ".csv"
		$outfilecsv49 = "H:\M365Reports\UserO365TempProxyaddressTable_49_" + $filedate + ".csv"
		$outfilecsv50 = "H:\M365Reports\UserO365TempProxyaddressTable_50_" + $filedate + ".csv"
		
		$outzipfile = "H:\M365Reports\UserMailboxesO365Emailaddresses_" + $filedate + ".zip"
		if (test-path $outzipfile) { rm $outzipfile }
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for User Mailboxes.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile1 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile2 -Value $outfilecsv2 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile3 -Value $outfilecsv3 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile4 -Value $outfilecsv4 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile5 -Value $outfilecsv5 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile6 -Value $outfilecsv6 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile7 -Value $outfilecsv7 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile8 -Value $outfilecsv8 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile9 -Value $outfilecsv9 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile10 -Value $outfilecsv10 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile11 -Value $outfilecsv11 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile12 -Value $outfilecsv12 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile13 -Value $outfilecsv13 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile14 -Value $outfilecsv14 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile15 -Value $outfilecsv15 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile16 -Value $outfilecsv16 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile17 -Value $outfilecsv17 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile18 -Value $outfilecsv18 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile19 -Value $outfilecsv19 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile20 -Value $outfilecsv20 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile21 -Value $outfilecsv21 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile22 -Value $outfilecsv22 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile23 -Value $outfilecsv23 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile24 -Value $outfilecsv24 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile25 -Value $outfilecsv25 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile26 -Value $outfilecsv26 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile27 -Value $outfilecsv27 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile28 -Value $outfilecsv28 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile29 -Value $outfilecsv29 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile30 -Value $outfilecsv30 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile31 -Value $outfilecsv31 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile32 -Value $outfilecsv32 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile33 -Value $outfilecsv33 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile34 -Value $outfilecsv34 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile35 -Value $outfilecsv35 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile36 -Value $outfilecsv36 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile37 -Value $outfilecsv37 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile38 -Value $outfilecsv38 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile39 -Value $outfilecsv39 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile40 -Value $outfilecsv40 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile41 -Value $outfilecsv41 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile42 -Value $outfilecsv42 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile43 -Value $outfilecsv43 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile44 -Value $outfilecsv44 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile45 -Value $outfilecsv45 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile46 -Value $outfilecsv46 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile47 -Value $outfilecsv47 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile48 -Value $outfilecsv48 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile49 -Value $outfilecsv49 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile50 -Value $outfilecsv50 -Option ReadOnly -Scope Script -Force
	}
	
	
	
    if($IncludeSharedMailboxes) { 
$included += "SharedMailbox"
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\SharedO365TempProxyAddressTable_" + $filedate + ".csv"
		$outfilecsv1 = "H:\M365Reports\SharedO365TempProxyAddressTable_1_" + $filedate + ".csv"
		$outfilecsv2 = "H:\M365Reports\SharedO365TempProxyAddressTable_2_" + $filedate + ".csv"
		$outfilecsv3 = "H:\M365Reports\SharedO365TempProxyAddressTable_3_" + $filedate + ".csv"
		$outfilecsv4 = "H:\M365Reports\SharedO365TempProxyAddressTable_4_" + $filedate + ".csv"
		$outfilecsv5 = "H:\M365Reports\SharedO365TempProxyAddressTable_5_" + $filedate + ".csv"
		$outfilecsv6 = "H:\M365Reports\SharedO365TempProxyAddressTable_6_" + $filedate + ".csv"
		$outfilecsv7 = "H:\M365Reports\SharedO365TempProxyAddressTable_7_" + $filedate + ".csv"
		$outfilecsv8 = "H:\M365Reports\SharedO365TempProxyAddressTable_8_" + $filedate + ".csv"
		$outfilecsv9 = "H:\M365Reports\SharedO365TempProxyAddressTable_9_" + $filedate + ".csv"
		$outfilecsv10 = "H:\M365Reports\SharedO365TempProxyAddressTable_10_" + $filedate + ".csv"
		$outfilecsv11 = "H:\M365Reports\SharedO365TempProxyAddressTable_11_" + $filedate + ".csv"
		$outfilecsv12 = "H:\M365Reports\SharedO365TempProxyAddressTable_12_" + $filedate + ".csv"
		$outfilecsv13 = "H:\M365Reports\SharedO365TempProxyAddressTable_13_" + $filedate + ".csv"
		$outfilecsv14 = "H:\M365Reports\SharedO365TempProxyAddressTable_14_" + $filedate + ".csv"
		$outfilecsv15 = "H:\M365Reports\SharedO365TempProxyAddressTable_15_" + $filedate + ".csv"
		
		$outzipfile = "H:\M365Reports\SharedMailboxesO365Emailaddresses_" + $filedate + ".zip"
		if (test-path $outzipfile) { rm $outzipfile }
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for shared mailboxes.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile1 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile2 -Value $outfilecsv2 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile3 -Value $outfilecsv3 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile4 -Value $outfilecsv4 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile5 -Value $outfilecsv5 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile6 -Value $outfilecsv6 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile7 -Value $outfilecsv7 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile8 -Value $outfilecsv8 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile9 -Value $outfilecsv9 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile10 -Value $outfilecsv10 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile11 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile12 -Value $outfilecsv12 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile13 -Value $outfilecsv13 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile14 -Value $outfilecsv14 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile15 -Value $outfilecsv15 -Option ReadOnly -Scope Script -Force
}
    if($IncludeRoomMailboxes) {
 $included += "RoomMailbox","EquipmentMailbox"
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\RoomO365TempProxyAddressTable_" + $filedate + ".csv"
		$outfilecsv1 = "H:\M365Reports\RoomO365TempProxyAddressTable_1_" + $filedate + ".csv"
		$outfilecsv2 = "H:\M365Reports\RoomO365TempProxyAddressTable_2_" + $filedate + ".csv"
		$outfilecsv3 = "H:\M365Reports\RoomO365TempProxyAddressTable_3_" + $filedate + ".csv"
		$outfilecsv4 = "H:\M365Reports\RoomO365TempProxyAddressTable_4_" + $filedate + ".csv"
		$outfilecsv5 = "H:\M365Reports\RoomO365TempProxyAddressTable_5_" + $filedate + ".csv"
		$outfilecsv6 = "H:\M365Reports\RoomO365TempProxyAddressTable_6_" + $filedate + ".csv"
		$outfilecsv7 = "H:\M365Reports\RoomO365TempProxyAddressTable_7_" + $filedate + ".csv"
		$outfilecsv8 = "H:\M365Reports\RoomO365TempProxyAddressTable_8_" + $filedate + ".csv"
		$outfilecsv9 = "H:\M365Reports\RoomO365TempProxyAddressTable_9_" + $filedate + ".csv"
		$outfilecsv10 = "H:\M365Reports\RoomO365TempProxyAddressTable_10_" + $filedate + ".csv"
		$outfilecsv11 = "H:\M365Reports\RoomO365TempProxyAddressTable_11_" + $filedate + ".csv"
		$outfilecsv12 = "H:\M365Reports\RoomO365TempProxyAddressTable_12_" + $filedate + ".csv"
		$outfilecsv13 = "H:\M365Reports\RoomO365TempProxyAddressTable_13_" + $filedate + ".csv"
		$outfilecsv14 = "H:\M365Reports\RoomO365TempProxyAddressTable_14_" + $filedate + ".csv"
		$outfilecsv15 = "H:\M365Reports\RoomO365TempProxyAddressTable_15_" + $filedate + ".csv"
		
			$outzipfile = "H:\M365Reports\RoomMailboxesO365Emailaddresses_" + $filedate + ".zip"
	if (test-path $outzipfile) { rm $outzipfile }
		
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for room mailboxes.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile1 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile2 -Value $outfilecsv2 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile3 -Value $outfilecsv3 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile4 -Value $outfilecsv4 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile5 -Value $outfilecsv5 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile6 -Value $outfilecsv6 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile7 -Value $outfilecsv7 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile8 -Value $outfilecsv8 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile9 -Value $outfilecsv9 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile10 -Value $outfilecsv10 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile11 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile12 -Value $outfilecsv12 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile13 -Value $outfilecsv13 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile14 -Value $outfilecsv14 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile15 -Value $outfilecsv15 -Option ReadOnly -Scope Script -Force
}
    if($IncludeMailUsers) { 
$included += "MailUser"
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\MailUsersO365TempProxyAddressTable_" + $filedate + ".csv"
		$outfilecsv1 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_1_" + $filedate + ".csv"
		$outfilecsv2 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_2_" + $filedate + ".csv"
		$outfilecsv3 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_3_" + $filedate + ".csv"
		$outfilecsv4 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_4_" + $filedate + ".csv"
		$outfilecsv5 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_5_" + $filedate + ".csv"
		$outfilecsv6 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_6_" + $filedate + ".csv"
		$outfilecsv7 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_7_" + $filedate + ".csv"
		$outfilecsv8 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_8_" + $filedate + ".csv"
		$outfilecsv9 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_9_" + $filedate + ".csv"
		$outfilecsv10 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_10_" + $filedate + ".csv"
		$outfilecsv11 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_11_" + $filedate + ".csv"
		$outfilecsv12 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_12_" + $filedate + ".csv"
		$outfilecsv13 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_13_" + $filedate + ".csv"
		$outfilecsv14 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_14_" + $filedate + ".csv"
		$outfilecsv15 = "H:\M365Reports\MailUsersO365TempProxyAddressTable_15_" + $filedate + ".csv"
		
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for mailusers.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile1 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile2 -Value $outfilecsv2 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile3 -Value $outfilecsv3 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile4 -Value $outfilecsv4 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile5 -Value $outfilecsv5 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile6 -Value $outfilecsv6 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile7 -Value $outfilecsv7 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile8 -Value $outfilecsv8 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile9 -Value $outfilecsv9 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile10 -Value $outfilecsv10 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile11 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile12 -Value $outfilecsv12 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile13 -Value $outfilecsv13 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile14 -Value $outfilecsv14 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile15 -Value $outfilecsv15 -Option ReadOnly -Scope Script -Force
	}
	
	
	if ($IncludeGuests)
	{
		$included += "GuestMailUser"
		
		$filedate = get-date -f dd-MM-yyyy-HH-mm-ss
		
		$outfilecsv = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_" + $filedate + ".csv"
		$outfilecsv1 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_1_" + $filedate + ".csv"
		$outfilecsv2 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_2_" + $filedate + ".csv"
		$outfilecsv3 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_3_" + $filedate + ".csv"
		$outfilecsv4 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_4_" + $filedate + ".csv"
		$outfilecsv5 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_5_" + $filedate + ".csv"
		$outfilecsv6 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_6_" + $filedate + ".csv"
		$outfilecsv7 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_7_" + $filedate + ".csv"
		$outfilecsv8 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_8_" + $filedate + ".csv"
		$outfilecsv9 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_9_" + $filedate + ".csv"
		$outfilecsv10 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_10_" + $filedate + ".csv"
		$outfilecsv11 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_11_" + $filedate + ".csv"
		$outfilecsv12 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_12_" + $filedate + ".csv"
		$outfilecsv13 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_13_" + $filedate + ".csv"
		$outfilecsv14 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_14_" + $filedate + ".csv"
		$outfilecsv15 = "H:\M365Reports\GuestUsersO365TempProxyAddressTable_15_" + $filedate + ".csv"
		
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for GuestMailUsers.... | $now "
		Set-Variable -Name outfile -Value $outfilecsv -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile1 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile2 -Value $outfilecsv2 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile3 -Value $outfilecsv3 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile4 -Value $outfilecsv4 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile5 -Value $outfilecsv5 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile6 -Value $outfilecsv6 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile7 -Value $outfilecsv7 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile8 -Value $outfilecsv8 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile9 -Value $outfilecsv9 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile10 -Value $outfilecsv10 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile11 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile12 -Value $outfilecsv12 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile13 -Value $outfilecsv13 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile14 -Value $outfilecsv14 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile15 -Value $outfilecsv15 -Option ReadOnly -Scope Script -Force
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
		$outfilecsv1 = "H:\M365Reports\GroupsO365TempProxyAddressTable_1_" + $filedate + ".csv"
		$outfilecsv2 = "H:\M365Reports\GroupsO365TempProxyAddressTable_2_" + $filedate + ".csv"
		$outfilecsv3 = "H:\M365Reports\GroupsO365TempProxyAddressTable_3_" + $filedate + ".csv"
		$outfilecsv4 = "H:\M365Reports\GroupsO365TempProxyAddressTable_4_" + $filedate + ".csv"
		$outfilecsv5 = "H:\M365Reports\GroupsO365TempProxyAddressTable_5_" + $filedate + ".csv"
		$outfilecsv6 = "H:\M365Reports\GroupsO365TempProxyAddressTable_6_" + $filedate + ".csv"
		$outfilecsv7 = "H:\M365Reports\GroupsO365TempProxyAddressTable_7_" + $filedate + ".csv"
		$outfilecsv8 = "H:\M365Reports\GroupsO365TempProxyAddressTable_8_" + $filedate + ".csv"
		$outfilecsv9 = "H:\M365Reports\GroupsO365TempProxyAddressTable_9_" + $filedate + ".csv"
		$outfilecsv10 = "H:\M365Reports\GroupsO365TempProxyAddressTable_10_" + $filedate + ".csv"
		$outfilecsv11 = "H:\M365Reports\GroupsO365TempProxyAddressTable_11_" + $filedate + ".csv"
		$outfilecsv12 = "H:\M365Reports\GroupsO365TempProxyAddressTable_12_" + $filedate + ".csv"
		$outfilecsv13 = "H:\M365Reports\GroupsO365TempProxyAddressTable_13_" + $filedate + ".csv"
		$outfilecsv14 = "H:\M365Reports\GroupsO365TempProxyAddressTable_14_" + $filedate + ".csv"
		$outfilecsv15 = "H:\M365Reports\GroupsO365TempProxyAddressTable_15_" + $filedate + ".csv"
		$outfilecsv16 = "H:\M365Reports\GroupsO365TempProxyAddressTable_16_" + $filedate + ".csv"
		$outfilecsv17 = "H:\M365Reports\GroupsO365TempProxyAddressTable_17_" + $filedate + ".csv"
		$outfilecsv18 = "H:\M365Reports\GroupsO365TempProxyAddressTable_18_" + $filedate + ".csv"
		$outfilecsv19 = "H:\M365Reports\GroupsO365TempProxyAddressTable_19_" + $filedate + ".csv"
		$outfilecsv20 = "H:\M365Reports\GroupsO365TempProxyAddressTable_20_" + $filedate + ".csv"
		$outfilecsv21 = "H:\M365Reports\GroupsO365TempProxyAddressTable_21_" + $filedate + ".csv"
		$outfilecsv22 = "H:\M365Reports\GroupsO365TempProxyAddressTable_22_" + $filedate + ".csv"
		$outfilecsv23 = "H:\M365Reports\GroupsO365TempProxyAddressTable_23_" + $filedate + ".csv"
		$outfilecsv24 = "H:\M365Reports\GroupsO365TempProxyAddressTable_24_" + $filedate + ".csv"
		$outfilecsv25 = "H:\M365Reports\GroupsO365TempProxyAddressTable_25_" + $filedate + ".csv"
		$outfilecsv26 = "H:\M365Reports\GroupsO365TempProxyAddressTable_26_" + $filedate + ".csv"
		$outfilecsv27 = "H:\M365Reports\GroupsO365TempProxyAddressTable_27_" + $filedate + ".csv"
		$outfilecsv28 = "H:\M365Reports\GroupsO365TempProxyAddressTable_28_" + $filedate + ".csv"
		$outfilecsv29 = "H:\M365Reports\GroupsO365TempProxyAddressTable_29_" + $filedate + ".csv"
		$outfilecsv30 = "H:\M365Reports\GroupsO365TempProxyAddressTable_30_" + $filedate + ".csv"
		$outfilecsv31 = "H:\M365Reports\GroupsO365TempProxyAddressTable_31_" + $filedate + ".csv"
		$outfilecsv32 = "H:\M365Reports\GroupsO365TempProxyAddressTable_32_" + $filedate + ".csv"
		$outfilecsv33 = "H:\M365Reports\GroupsO365TempProxyAddressTable_33_" + $filedate + ".csv"
		$outfilecsv34 = "H:\M365Reports\GroupsO365TempProxyAddressTable_34_" + $filedate + ".csv"
		$outfilecsv35 = "H:\M365Reports\GroupsO365TempProxyAddressTable_35_" + $filedate + ".csv"
		$outfilecsv36 = "H:\M365Reports\GroupsO365TempProxyAddressTable_36_" + $filedate + ".csv"
		$outfilecsv37 = "H:\M365Reports\GroupsO365TempProxyAddressTable_37_" + $filedate + ".csv"
		$outfilecsv38 = "H:\M365Reports\GroupsO365TempProxyAddressTable_38_" + $filedate + ".csv"
		$outfilecsv39 = "H:\M365Reports\GroupsO365TempProxyAddressTable_39_" + $filedate + ".csv"
		$outfilecsv40 = "H:\M365Reports\GroupsO365TempProxyAddressTable_40_" + $filedate + ".csv"
		$outfilecsv41 = "H:\M365Reports\GroupsO365TempProxyAddressTable_41_" + $filedate + ".csv"
		$outfilecsv42 = "H:\M365Reports\GroupsO365TempProxyAddressTable_42_" + $filedate + ".csv"
		$outfilecsv43 = "H:\M365Reports\GroupsO365TempProxyAddressTable_43_" + $filedate + ".csv"
		$outfilecsv44 = "H:\M365Reports\GroupsO365TempProxyAddressTable_44_" + $filedate + ".csv"
		$outfilecsv45 = "H:\M365Reports\GroupsO365TempProxyAddressTable_45_" + $filedate + ".csv"
		$outfilecsv46 = "H:\M365Reports\GroupsO365TempProxyAddressTable_46_" + $filedate + ".csv"
		$outfilecsv47 = "H:\M365Reports\GroupsO365TempProxyAddressTable_47_" + $filedate + ".csv"
		$outfilecsv48 = "H:\M365Reports\GroupsO365TempProxyAddressTable_48_" + $filedate + ".csv"
		$outfilecsv49 = "H:\M365Reports\GroupsO365TempProxyAddressTable_49_" + $filedate + ".csv"
		$outfilecsv50 = "H:\M365Reports\GroupsO365TempProxyAddressTable_50_" + $filedate + ".csv"
		
			$outzipfile = "H:\M365Reports\GroupsO365Emailaddresses_" + $filedate + ".zip"
	if (test-path $outzipfile) { rm $outzipfile }
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Exporting proxyaddresses for Distribution groups.... | $now "
		Set-Variable -Name outfile1 -Value $outfilecsv1 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile2 -Value $outfilecsv2 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile3 -Value $outfilecsv3 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile4 -Value $outfilecsv4 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile5 -Value $outfilecsv5 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile6 -Value $outfilecsv6 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile7 -Value $outfilecsv7 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile8 -Value $outfilecsv8 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile9 -Value $outfilecsv9 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile10 -Value $outfilecsv10 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile11 -Value $outfilecsv11 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile12 -Value $outfilecsv12 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile13 -Value $outfilecsv13 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile14 -Value $outfilecsv14 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile15 -Value $outfilecsv15 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile16 -Value $outfilecsv16 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile17 -Value $outfilecsv17 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile18 -Value $outfilecsv18 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile19 -Value $outfilecsv19 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile20 -Value $outfilecsv20 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile21 -Value $outfilecsv21 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile22 -Value $outfilecsv22 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile23 -Value $outfilecsv23 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile24 -Value $outfilecsv24 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile25 -Value $outfilecsv25 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile26 -Value $outfilecsv26 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile27 -Value $outfilecsv27 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile28 -Value $outfilecsv28 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile29 -Value $outfilecsv29 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile30 -Value $outfilecsv30 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile31 -Value $outfilecsv31 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile32 -Value $outfilecsv32 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile33 -Value $outfilecsv33 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile34 -Value $outfilecsv34 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile35 -Value $outfilecsv35 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile36 -Value $outfilecsv36 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile37 -Value $outfilecsv37 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile38 -Value $outfilecsv38 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile39 -Value $outfilecsv39 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile40 -Value $outfilecsv40 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile41 -Value $outfilecsv41 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile42 -Value $outfilecsv42 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile43 -Value $outfilecsv43 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile44 -Value $outfilecsv44 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile45 -Value $outfilecsv45 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile46 -Value $outfilecsv46 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile47 -Value $outfilecsv47 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile48 -Value $outfilecsv48 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile49 -Value $outfilecsv49 -Option ReadOnly -Scope Script -Force
		Set-Variable -Name outfile50 -Value $outfilecsv50 -Option ReadOnly -Scope Script -Force

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
	
	if ($IncludeGuests) { $MBList = Get-EXORecipient -ResultSize Unlimited -Properties Displayname, PrimarySMTPAddress, WindowsLiveID, EmailAddresses, RecipientTypeDetails -RecipientTypeDetails GuestMailUser }
	
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
	if ($IncludeUserMailboxes) { $outputcount = $outputcount * 4 }
	if ($IncludeSharedMailboxes) { $outputcount = $outputcount * 2 }
	if ($IncludeDGs) { $outputcount = $outputcount * 2 }
	if ($IncludeRoomMailboxes) { $outputcount = $outputcount * 2 }
	if ($IncludeMailUsers) { $outputcount = $outputcount }
	if ($IncludeGuests) { $outputcount = $outputcount }
	
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
					$I++
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
                #$arrAliases += $objAliases
					
					if ($I -le 10000)
					{
						$objAliases | Export-CSV $outfile1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 10000) -and ($I -le 20000))
					{
						$objAliases | Export-CSV $outfile2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 20000) -and ($I -le 30000))
					{
						$objAliases | Export-CSV $outfile3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 30000) -and ($I -le 40000))
					{
						$objAliases | Export-CSV $outfile4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 40000) -and ($I-le 50000))
					{
						$objAliases | Export-CSV $outfile5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 50000) -and ($I -le 60000))
					{
						$objAliases | Export-CSV $outfile6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 60000) -and ($I -le 70000))
					{
						$objAliases  | Export-CSV $outfile7 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 70000) -and ($I -le 80000))
					{
						$objAliases| Export-CSV $outfile8 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 80000) -and ($I -le 90000))
					{
						$objAliases | Export-CSV $outfile9 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 90000) -and ($I -le 100000))
					{
						$objAliases | Export-CSV $outfile10 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 100000) -and ($I -le 110000))
					{
						$objAliases | Export-CSV $outfile11 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 110000) -and ($I -le 120000))
					{
						$objAliases | Export-CSV $outfile12 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 120000) -and ($I -le 130000))
					{
						$objAliases | Export-CSV $outfile13 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 130000) -and ($I -le 140000))
					{
						$objAliases | Export-CSV $outfile14 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 140000) -and ($I -le 150000))
					{
						$objAliases | Export-CSV $outfile15 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					if (($I -ge 150000) -and ($I -le 160000))
					{
						$objAliases | Export-CSV $outfile16 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					if (($I -ge 160000) -and ($I -le 170000))
					{
						$objAliases | Export-CSV $outfile17 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					if (($I -ge 170000) -and ($I -le 180000))
					{
						$objAliases | Export-CSV $outfile18 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 180000) -and ($I -le 190000))
					{
						$objAliases | Export-CSV $outfile19 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 190000) -and ($I -le 200000))
					{
						$objAliases | Export-CSV $outfile20 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 200000) -and ($I -le 210000))
					{
						$objAliases | Export-CSV $outfile21 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
										
					
					if (($I -ge 210000) -and ($I -le 220000))
					{
						$objAliases | Export-CSV $outfile22 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					if (($I -ge 220000) -and ($I -le 230000))
					{
						$objAliases | Export-CSV $outfile23 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 230000) -and ($I -le 240000))
					{
						$objAliases | Export-CSV $outfile24 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 240000) -and ($I -le 250000))
					{
						$objAliases | Export-CSV $outfile25 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 250000) -and ($I -le 260000))
					{
						$objAliases | Export-CSV $outfile26 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 260000) -and ($I -le 270000))
					{
						$objAliases | Export-CSV $outfile27 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 270000) -and ($I -le 280000))
					{
						$objAliases | Export-CSV $outfile28 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 280000) -and ($I -le 290000))
					{
						$objAliases | Export-CSV $outfile29 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 290000) -and ($I -le 300000))
					{
						$objAliases | Export-CSV $outfile30 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 300000) -and ($I -le 310000))
					{
						$objAliases | Export-CSV $outfile31 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 310000) -and ($I -le 320000))
					{
						$objAliases | Export-CSV $outfile32 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 320000) -and ($I -le 330000))
					{
						$objAliases | Export-CSV $outfile33 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 330000) -and ($I -le 340000))
					{
						$objAliases | Export-CSV $outfile34 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 340000) -and ($I -le 350000))
					{
						$objAliases | Export-CSV $outfile35 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 350000) -and ($I -le 360000))
					{
						$objAliases | Export-CSV $outfile36 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 360000) -and ($I -le 370000))
					{
						$objAliases | Export-CSV $outfile37 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 370000) -and ($I -le 380000))
					{
						$objAliases | Export-CSV $outfile38 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 380000) -and ($I -le 390000))
					{
						$objAliases | Export-CSV $outfile39 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 390000) -and ($I -le 400000))
					{
						$objAliases | Export-CSV $outfile40 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 400000) -and ($I -le 410000))
					{
						$objAliases | Export-CSV $outfile41 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 410000) -and ($I -le 420000))
					{
						$objAliases | Export-CSV $outfile42 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 420000) -and ($I -le 430000))
					{
						$objAliases | Export-CSV $outfile43 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 430000) -and ($I -le 440000))
					{
						$objAliases | Export-CSV $outfile44 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($I -ge 440000) -and ($I -le 450000))
					{
						$objAliases | Export-CSV $outfile45 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 450000) -and ($I -le 460000))
					{
						$objAliases | Export-CSV $outfile46 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 460000) -and ($I -le 470000))
					{
						$objAliases | Export-CSV $outfile47 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 470000) -and ($I -le 480000))
					{
						$objAliases | Export-CSV $outfile48 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 480000) -and ($I -le 490000))
					{
						$objAliases | Export-CSV $outfile49 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($I -ge 490000) -and ($I -le 500000))
					{
						$objAliases | Export-CSV $outfile50 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					
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
        }$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Found $outputcount for export of proxyaddresses for the types $included . Now creating output :::$time"

  $perusersum += $peruser.ElapsedMilliseconds
    $peruseravg = $perusersum / $i
    $stillneeded = $outputcount - $i
    $remain = (($peruseravg * $stillneeded) / 1000)
    write-progress -activity "Processing mailbox: $mbxbname" -status "${i} of ${outputcount} ($formated%) Remaining: $((New-TimeSpan -Seconds $remain).ToString()) Average(ms): $($peruseravg.ToString("#.##"))" -percentcomplete $([Math]::Round(($i / $outputcount * 100))) -id 1
    }
    #Output the result to the console host. Rearrange/sort as needed.
   # $arrAliases | select * -ExcludeProperty Number

#$arrAliases | Export-Csv $outfile -nti -encoding UTF8 -delimiter '|'
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	Add-Content $logfile "Found $i proxyaddresses for the types $included . Now copying to CSV files to Filewatcher :::$time"
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


Map-Filewatcher

if (Test-Path $Outfile1)
{
	Try
	{
		
		Copy-item -path $Outfile1 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "ProxyAddresses File Copied to FileWatcher $Outfile1 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile1 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "ProxyAddresses File Copied to FileWatcher $Outfile1 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			#RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile1 | $now"
		}
	}
}


if (Test-Path $Outfile2)
{
	Try
	{
		
		#Map-Filewatcher
		Copy-item -path $Outfile2 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile2 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Fileout -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile2 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile2 | $now"
		}
	}
}


if (Test-Path $Outfile3)
{
	
	
	Try
	{
		
		#Map-Filewatcher
		Copy-item -path $Outfile3 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile3 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile3 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile3 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile3 | $now"
		}
	}
	
}

if (Test-Path $Outfile4)
{
	
	
	Try
	{
		
		#Map-Filewatcher
		Copy-item -path $Outfile4 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile4 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile4 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile4 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile4 | $now"
		}
	}
	
}

if (Test-Path $Outfile5)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile5 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile5 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile5 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile5 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile5 | $now"
		}
	}
	
}

if (Test-Path $Outfile6)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile6 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile6 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile6 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile6 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile6 | $now"
		}
	}
	
}


if (Test-Path $Outfile7)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile7 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile7 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile7 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile7 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile7 | $now"
		}
	}
	
}


if (Test-Path $Outfile8)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile8 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile8 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile8 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile8 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile8 | $now"
		}
	}
	
}

if (Test-Path $Outfile9)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile9 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile9 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile9 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile9 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "remotemailbox copy files to FileWatcher $Outfile9 | $now"
		}
	}
	
}


if (Test-Path $Outfile10)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile10 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile10 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile10 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile10 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile10 | $now"
		}
	}
	
}

if (Test-Path $Outfile11)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile11 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile11 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile11 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile11 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile11 | $now"
		}
	}
	
}


if (Test-Path $Outfile12)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile12 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile12 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile12 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile12 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile12 | $now"
		}
	}
	
}

if (Test-Path $Outfile13)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile13 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile13 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile13 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile13 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile13 | $now"
		}
	}
	
}

if (Test-Path $Outfile14)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile14 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile14 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile14 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile14 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile14 | $now"
		}
	}
	
}


if (Test-Path $Outfile15)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile15 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile15 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile15 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile15 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile15 | $now"
		}
	}
	
}

if (Test-Path $Outfile16)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile16 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile16 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile16 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile16 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile16 | $now"
		}
	}
	
}


if (Test-Path $Outfile17)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile17 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile17 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile17 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile17 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile17 | $now"
		}
	}
	
}


if (Test-Path $Outfile18)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile18 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile18 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile18 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile18 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile18 | $now"
		}
	}
	
}

if (Test-Path $Outfile19)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile19 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile19 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile19 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile19 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "remotemailbox copy files to FileWatcher $Outfile19 | $now"
		}
	}
	
}


if (Test-Path $Outfile20)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile20 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile20 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile20 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile20 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile20 | $now"
		}
	}
	
}

if (Test-Path $Outfile21)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile21 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile21 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile21 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile21 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile11 | $now"
		}
	}
	
}


if (Test-Path $Outfile22)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile22 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile22 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile22 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile22 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile22 | $now"
		}
	}
	
}

if (Test-Path $Outfile23)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile23 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile23 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile23 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile23 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile23 | $now"
		}
	}
	
}

if (Test-Path $Outfile24)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile24 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile24 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile24 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile24 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile24 | $now"
		}
	}
	
}


if (Test-Path $Outfile25)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile25 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile25 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile25 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile25 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile25 | $now"
		}
	}
	
}

if (Test-Path $Outfile26)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile26 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile26 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile26 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile26 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile26 | $now"
		}
	}
	
}


if (Test-Path $Outfile27)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile27 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile27 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile27 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile27 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile27 | $now"
		}
	}
	
}


if (Test-Path $Outfile28)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile28 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile28 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile28 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile28 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile28 | $now"
		}
	}
	
}

if (Test-Path $Outfile29)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile29 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile29 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile29 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile29 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "remotemailbox copy files to FileWatcher $Outfile29 | $now"
		}
	}
	
}


if (Test-Path $Outfile30)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile30 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile30 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile30 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile30 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile30 | $now"
		}
	}
	
}



if (Test-Path $Outfile31)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile31 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile31 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile31 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile31 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile31 | $now"
		}
	}
	
}


if (Test-Path $Outfile32)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile32 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile32 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile32 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile32 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile32 | $now"
		}
	}
	
}

if (Test-Path $Outfile33)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile33 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile33 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile33 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile33 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile33 | $now"
		}
	}
	
}

if (Test-Path $Outfile34)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile34 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile34 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile34 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile34 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile34 | $now"
		}
	}
	
}


if (Test-Path $Outfile35)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile35 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile35 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile35 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile35 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile35 | $now"
		}
	}
	
}

if (Test-Path $Outfile36)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile36 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile36 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile36 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile36 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile36 | $now"
		}
	}
	
}


if (Test-Path $Outfile37)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile37 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile37 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile37 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile37 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile37 | $now"
		}
	}
	
}


if (Test-Path $Outfile38)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile38 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile38 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile38 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile38 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile38 | $now"
		}
	}
	
}

if (Test-Path $Outfile39)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile39 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile39 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile39 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile39 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "remotemailbox copy files to FileWatcher $Outfile39 | $now"
		}
	}
	
}


if (Test-Path $Outfile40)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile40 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile40 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile40 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile40 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile40 | $now"
		}
	}
	
}


if (Test-Path $Outfile41)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile41 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile41 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile41 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile41 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile41 | $now"
		}
	}
	
}


if (Test-Path $Outfile42)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile42 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile42 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile42 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile42 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile42 | $now"
		}
	}
	
}

if (Test-Path $Outfile43)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile43 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile43 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile43 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile43 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile43 | $now"
		}
	}
	
}

if (Test-Path $Outfile44)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile44 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile44 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile44 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile44 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile44 | $now"
		}
	}
	
}


if (Test-Path $Outfile45)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile45 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile45 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile45 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile45 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile45 | $now"
		}
	}
	
}

if (Test-Path $Outfile46)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile46 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile46 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile46 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile46 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile46 | $now"
		}
	}
	
}


if (Test-Path $Outfile47)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile47 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile47 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile47 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile47 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile47 | $now"
		}
	}
	
}


if (Test-Path $Outfile48)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile48 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile48 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile48 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile48 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile48 | $now"
		}
	}
	
}

if (Test-Path $Outfile49)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile49 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile49 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile49 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile49 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "remotemailbox copy files to FileWatcher $Outfile49 | $now"
		}
	}
	
}


if (Test-Path $Outfile50)
{
	
	
	Try
	{
		#Map-Filewatcher
		Copy-item -path $Outfile50 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile50 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 30
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		Start-Sleep -s 30
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile50 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp proxyaddresses File Copied to FileWatcher $Outfile50 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile50 | $now"
		}
	}
	
}


$finished = get-date -Format dd-MM-yyyy-HH:mm
$StopWatch.Stop()
$ElapsedTime = $StopWatch.Elapsed

$scripthours = $ElapsedTime.Hours
$scriptminutes = $ElapsedTime.Minutes

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "The Script Finished at $finished"
Add-Content $logfile "The script took $ElapsedTime [$scripthours hours, $scriptminutes minutes] to run."

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
