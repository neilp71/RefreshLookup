



<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.148
	 Created on:    07/08/2023 14:00
	 Created by:   	Neil Poultney
	 Organization: 	NDP Consultancy Ltd - neil.poultney@ndpconsultancyltd.co.uk
	 Filename:     	Refresh-EARL-get-lookupTable-Export-generic.ps1
	===========================================================================
	.DESCRIPTION
		Exports for EARL MailDb to Temp Lookup Table for reresh of Database.

		Change Log
		V1.00, 20/01/2024 - Initial full version
		

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
		even if advised of the possibility of such damage.
#>


[System.GC]::Collect()

$StopWatch = New-Object System.Diagnostics.Stopwatch
$StopWatch.Start()

# Make Windows negotiate higher TLS version:
[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Import-Module ActiveDirectory
$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")

$global:nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"

[System.GC]::Collect()

#change 
#$lookuptime = (get-date).adddays(-2)
$lookuptime = (get-date).addhours(-3)
Set-Variable -Name lasthour -Value $lookuptime -Option ReadOnly -Scope Script -Force

##workoutwhereweare
$Domainwearein = (Get-WmiObject Win32_ComputerSystem).Name
$whoweare = $ENV:USERNAME

if ($domainwearein -eq "zneepacp11eme2" -or $domainwearein -eq "zneepacp11emrg") { $global:Envirionmentchoice = "ProdNE" }
if ($domainwearein -eq "zweepacp11emg3" -or $domainwearein -eq "zweepacp11em50") { $global:Envirionmentchoice = "ProdWE" }



$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
$transcriptlog = "H:\EARLTranscripts\LookupTbl\Refresh-lookup-export-generic-" + $nowfiledate + ".log"

Start-Transcript -Path $transcriptlog

$now
$ENV:USERNAME
$Domainwearein
$Envirionmentchoice


if ($Envirionmentchoice -eq "ProdWE")
{
	
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$loglocation = "H:\EARLPSLogs\BulkExports\"  # change to usetype RPMBCREATE etc
	Set-Variable -Name logfolder -Value $logslocation -Option ReadOnly -Scope Script -Force
	$logfilelocation = $loglocation + "Refresh-Lookup-Table-generic-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
	Set-Variable -Name logfile -Value $logfilelocation -Option ReadOnly -Scope Script -Force
	
	$EARLNTID = "BP1\task-EARLEXCWE1"
	$EARLNTID2 = "BP1\task-EARLEXCWE2"
	if ($domainwearein -eq "zweepacp11emg3")
	{
		$secureAES = "F:\AppCerts\PSUserCred\secureaes.key"
	}
	
	if ($domainwearein -eq "zweepacp11em50")
	{
		$secureAES = "F:\PSDetails\EASKey\secureaes.key"
	}
	
	$EARLUserPWFile = "F:\PSDetails\PSUserCred\EARLEXCWE1.txt"
	$EARLUserPWFile2 = "F:\PSDetails\PSUserCred\EARLEXCWE2.txt"
	Set-Variable -Name EARLPW -Value $EARLUserPWFile -Option ReadOnly -Scope Script -Force
	Set-Variable -Name EARLPW2 -Value $EARLUserPWFile2 -Option ReadOnly -Scope Script -Force
	Set-Variable -Name EARLPWSecureAES -Value $secureAES -Option ReadOnly -Scope Script -Force
	Set-Variable -Name EARLUser -Value $EARLNTID -Option ReadOnly -Scope Script -Force
	Set-Variable -Name EARLUser2 -Value $EARLNTID2 -Option ReadOnly -Scope Script -Force
	
	#$filewatcherlocationout = "Q:\EARL\FileLocation\"
	$filewatcherlocationout = "Q:\EARL\CSVFileLocation\"
	Set-Variable -Name FileWatcherOut -Value $filewatcherlocationout -Option ReadOnly -Scope Script -Force
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Start of Log File"
	add-content $logfile  $now
	add-content $logfile  "Processing in Live environment for $Envirionmentchoice for user $whoweare"
	
}

if ($Envirionmentchoice -eq "ProdNE")
{
	
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$loglocation = "H:\EARLPSLogs\BulkExports\"  # change to usetype RPMBCREATE etc
	Set-Variable -Name logfolder -Value $logslocation -Option ReadOnly -Scope Script -Force
	$logfilelocation = $loglocation + "Lookup-Table-generic-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
	Set-Variable -Name logfile -Value $logfilelocation -Option ReadOnly -Scope Script -Force
	
	$EARLNTID = "BP1\task-EARLEXCNE1"
	$EARLNTID2 = "BP1\task-EARLEXCNE2"
	if ($domainwearein -eq "zneepacp11eme2")
	{
		$secureAES = "F:\AppCerts\PSUserCred\secureaes.key"
	}
	
	if ($domainwearein -eq "zneepacp11emrg")
	{
		$secureAES = "F:\PSDetails\EASKey\secureaes.key"
	}
	
	$EARLUserPWFile = "F:\PSDetails\PSUserCred\EARLEXCNE1.txt"
	$EARLUserPWFile2 = "F:\PSDetails\PSUserCred\EARLEXCNE2.txt"
	Set-Variable -Name EARLPW -Value $EARLUserPWFile -Option ReadOnly -Scope Script -Force
	Set-Variable -Name EARLPW2 -Value $EARLUserPWFile2 -Option ReadOnly -Scope Script -Force
	Set-Variable -Name EARLPWSecureAES -Value $secureAES -Option ReadOnly -Scope Script -Force
	Set-Variable -Name EARLUser -Value $EARLNTID -Option ReadOnly -Scope Script -Force
	Set-Variable -Name EARLUser2 -Value $EARLNTID2 -Option ReadOnly -Scope Script -Force
	
	$filewatcherlocationout = "Q:\EARL\CSVFileLocation\"
	#$filewatcherlocationout = "Q:\EARL\FileLocation\"
	Set-Variable -Name FileWatcherOut -Value $filewatcherlocationout -Option ReadOnly -Scope Script -Force
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Start of Log File"
	add-content $logfile  $now
	add-content $logfile  "Processing in Live environment for $Envirionmentchoice for user $whoweare"
	
}




#Import-Module ActiveDirectory


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

Function ConnectExchangeonPrem
{
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
	Add-Content $logfile "Exchange OnPremise remote powershell $Envirionmentchoice | $now"
	
	$EXpassword = ""
	$EXCredentials = ""
	
	$Error.Clear()
	Get-PSSession | Remove-PSSession
	#Disconnect-MgGraph
	
	
	if ($Envirionmentchoice -eq "ProdWE")
	{
		
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
		Add-Content $logfile "Attempting to connect to Exchange OnPremise remote powershell Prod || $now"
		#set random figure to set which exchange mailbox server to connect to - fail back to next server
		$randomchoice = Get-Random -Minimum 1 -Maximum 3
		
		if ($randomchoice -eq "1")
		{
			$securepassword = Get-Content $EARLUserPWFile | ConvertTo-SecureString -Key (Get-Content $EARLPWSecureAES)
			$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $EARLUser, $securepassword
			$accchosen = "task-EARLEXCWE1"
			
			
		}
		
		if ($randomchoice -eq "2")
		{
			$securepassword = Get-Content $EARLUserPWFile2 | ConvertTo-SecureString -Key (Get-Content $EARLPWSecureAES)
			$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $EARLUser2, $securepassword
			$accchosen = "task-EARLEXCWE2"
		}
		
		Try
		{
			$randomchoice2 = Get-Random -Minimum 1 -Maximum 3
			if ($randomchoice2 -eq "1")
			{
				$exchangesession = New-PSSession -Credential $Credentials -ConfigurationName Microsoft.Exchange -authentication Basic -ConnectionUri https://zneepacp11em3z.bp1.ad.bp.com/PowerShell -AllowRedirection
			}
			
			if ($randomchoice2 -eq "2")
			{
				$exchangesession = New-PSSession -Credential $Credentials -ConfigurationName Microsoft.Exchange -authentication Basic -ConnectionUri https://zweepacp11emhx.bp1.ad.bp.com/PowerShell -AllowRedirection
			}
			
			
			#$exchangesession = New-PSSession -Credential $Credentials -ConfigurationName Microsoft.Exchange -authentication Basic -ConnectionUri https://mail.bp.com/PowerShell -AllowRedirection
			
			
			Import-PsSession $exchangesession
			$connectiontoonprem = Get-PSSession | select *
			$connectsessionID = $connectiontoonprem.ConfigurationName
			$connectsessionComputer = $connectiontoonprem.ComputerType
			if (($connectsessionID -eq "Microsoft.Exchange") -and ($connectsessionComputer -eq "mail.bp.com"))
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
				Add-Content $logfile "Connected to Exchange OnPremise remote powershell with account choice $randomchoice | $accchosen  | $now"
				$connecttry = "0"
			}
			
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
			$connecttry = $connecttry + 1
			$errorMessage = $_.Exception.Message
			$itemfailed = $_.Exception.item
			Add-Content $logfile "could not connect to Exchange 2019 on Premise Will try again this is the $connecttry try .... $errorMessage and $itemfailed Time: $now"
			Start-Sleep -Seconds 300
			if ($connecttry -ge "9") { Add-Content $logging "could not connect to On Premise Powershell i have tried $connecttry times and will quit .... Time: $now"; quit }
			else
			{
				ConnectExchangeonPrem
			}
			
		}
	}
	
	if ($Envirionmentchoice -eq "ProdNE")
	{
		
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
		Add-Content $logfile "Attempting to connect to Exchange OnPremise remote powershell Prod || $now"
		#set random figure to set which exchange mailbox server to connect to - fail back to next server
		$randomchoice = Get-Random -Minimum 1 -Maximum 3
		
		if ($randomchoice -eq "1")
		{
			$securepassword = Get-Content $EARLUserPWFile | ConvertTo-SecureString -Key (Get-Content $EARLPWSecureAES)
			$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $EARLUser, $securepassword
			$accchosen = "task-EARLEXCNE1"
			
			
		}
		
		if ($randomchoice -eq "2")
		{
			$securepassword = Get-Content $EARLUserPWFile2 | ConvertTo-SecureString -Key (Get-Content $EARLPWSecureAES)
			$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $EARLUser2, $securepassword
			$accchosen = "task-EARLEXCNE2"
		}
		
		Try
		{
			$randomchoice2 = Get-Random -Minimum 1 -Maximum 3
			if ($randomchoice2 -eq "1")
			{
				$exchangesession = New-PSSession -Credential $Credentials -ConfigurationName Microsoft.Exchange -authentication Basic -ConnectionUri https://zneepacp11em3z.bp1.ad.bp.com/PowerShell -AllowRedirection
			}
			
			if ($randomchoice2 -eq "2")
			{
				$exchangesession = New-PSSession -Credential $Credentials -ConfigurationName Microsoft.Exchange -authentication Basic -ConnectionUri https://zweepacp11emhx.bp1.ad.bp.com/PowerShell -AllowRedirection
			}
			
						#$exchangesession = New-PSSession -Credential $Credentials -ConfigurationName Microsoft.Exchange -authentication Basic -ConnectionUri https://mail.bp.com/PowerShell -AllowRedirection
			
		
			Import-PsSession $exchangesession
			$connectiontoonprem = Get-PSSession | select *
			$connectsessionID = $connectiontoonprem.ConfigurationName
			$connectsessionComputer = $connectiontoonprem.ComputerType
			if (($connectsessionID -eq "Microsoft.Exchange") -and ($connectsessionComputer -eq "mail.bp.com"))
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
				Add-Content $logfile "Connected to Exchange OnPremise remote powershell with account choice $randomchoice | $accchosen  | $now"
				$connecttry = "0"
			}
			
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
			$connecttry = $connecttry + 1
			$errorMessage = $_.Exception.Message
			$itemfailed = $_.Exception.item
			Add-Content $logfile "could not connect to Exchange 2019 on Premise Will try again this is the $connecttry try .... $errorMessage and $itemfailed Time: $now"
			Start-Sleep -Seconds 300
			if ($connecttry -ge "9") { Add-Content $logging "could not connect to On Premise Powershell i have tried $connecttry times and will quit .... Time: $now"; quit }
			else
			{
				ConnectExchangeonPrem
			}
			
		}
	}
}


function DisconnectExchangeOnPrem
{
	
	Get-PSSession | Remove-PsSession
	$now = get-date -Format dd-MM-yyyy-HH:mm:ss
	Add-Content $logfile "Disconnected From Exchange 2019 remote Powershell  .... Time: $now"
}




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
	Add-Content $logfile "Disconnected From Exchange Online remote Powershell  .... Time: $now"
}




$now = Get-Date -format dd-M-yyyy-HH-mm

$wheretoProcess = ""
$outputdate = get-date -f yyyy-MM-dd-HH-mm-ss




#ConnectExchangeonPrem
#$Mbx = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox -ResultSize 500 | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
#Mbx += Get-RemoteMailbox -ResultSize 500 | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
#$Mbx += Get-Mailuser -ResultSize 500 | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress





#We have some mailboxes, so we can process them...

$Report1 = [System.Collections.Generic.List[Object]]::new() # Create output file 
$Report2 = [System.Collections.Generic.List[Object]]::new() # Create output file 





function exportlocalmbx
{
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export Generic users for export"
	
	
	$count = ""
	$M = ""
	$reconnect = 0
	
	ConnectExchangeonPrem
	
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv1 = "H:\M365Reports\EARL-Prod-LookupTable-localmbx-Hourly-" + $nowfiledate + ".csv"
	$lasthour = (get-date).adddays(-1)
	#$lasthour = (get-date).addhours(-1)
	#$Mbx1 = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	$Mbx1 = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	$count = $Mbx1.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count local mailbox accounts to process"
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv2 for remote mbx "
	
	$reconnect = 0
	ForEach ($M in $Mbx1)
	{
		$MbxNumber++
		$reconnect++
		$NTID = $null #SamAccountName
		$Display = $null #displayName
		$UsrACCCtrl = $null #useraccountcontrol
		$HideAB = $null #msExchHideFromAddressLists
		$SN = $null #sn
		$firstName = $Null #givenName
		$mail = $null #mail
		$managerDN = $null #manager
		$managerDisp = $null
		$mgrEmail = $null
		$mgrAlias = $null
		$rectypedetail = $Null #msExchRecipientTypeDetails
		$dept = $null #Department
		$BPtext3201 = $null #GPID
		$distName = $null # DN
		$country = $null #co
		$Comp = $null #Company
		$CA9 = $null #Employee
		$country = $null #co
		$recipientdetailsEX = $null #RecipientTypeDetails
		$managerDisp = $null #ManagerDisplayName
		$mgrEmail = $null #ManagerEmail
		$mgrAlias = $null #Manager
		
		$usr = $M.PrimarySmtpAddress
		$Display = $M.DisplayName
		$recipientdetailsEX = $M.RecipientTypeDetails
		
		#write-host "Getting mailbox:: $usr"
		$acc = Get-ADUser -filter 'mail -eq $usr' -properties * | Select-Object *
		#$recpdetails = Get-Recipient -identity $usr -properties *
		#$recpdetails = Get-Recipient -identity $usr | Select-Object *
		
		[int]$Progress = $MbxNumber/$count * 100
		$PercentComplete = [math]::Round($Progress, 3)
		$disp = $acc.DisplayName
		$UPN = $acc.UserPrincipalName
		$MailboxType = $acc.msExchRecipientTypeDetails
		$mail = $acc.mail
		$SN = $acc.sn
		$firstName = $acc.GivenName
		$dept = $acc.Department
		$Comp = $acc.Company
		$country = $acc.co
		$UsrACCCtrl = $acc.useraccountcontrol
		$NTID = $acc.samaccountname
		$distName = $acc.DistinguishedName
		
		$managerDN = $acc.manager
		$BPtext3201 = $acc.'bp-Text32-01'
		$CA9 = $acc.extensionAttribute9
		$HideAB = $acc.msExchHideFromAddressLists
		
		
		if ($null -ne $managerDN)
		{
			$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
			
			$managerDisp = $mgrout.DisplayName
			$mgrEmail = $mgrout.PrimarySMTPAddress
			$mgrAlias = $mgrout.Alias
			
		}
		
		if ($recipientdetailsEX -eq "MailUser" -and ! $MailboxType)
		{
			$MailboxType = "128"
		}
		
		
		if (! $managerDisp)
		{
			
			
			$managerDisp = "NULL"
			
			
		}
		
		if (! $mgrEmail)
		{
			
			$mgrEmail = "NULL"
			
		}
		
		if (! $mgrAlias)
		{
			
			$mgrAlias = "NULL"
			
		}
		
		if (! $Comp)
		{
			$Comp = "NULL"
		}
		
		if (! $SN)
		{
			$SN = "NULL"
		}
		
		if (! $firstName)
		{
			$firstName = "NULL"
		}
		
		if (! $CA9)
		{
			$CA9 = "NULL"
		}
		
		if (! $BPtext3201)
		{
			$BPtext3201 = "NULL"
		}
		
		if (! $dept)
		{
			$dept = "NULL"
		}
		
		if (! $HideAB)
		{
			$HideAB = "False"
		}
		
		if (! $country)
		{
			$country = "NULL"
		}
		
		if (!$UPN)
		{
			$UPN = "NULL"
		}
		#$MbxStatus = $disp + " [" + $MbxNumber + "/" + $count + "]"
		
		
		
		if ($PercentComplete -eq "5.000" -or $PercentComplete -eq "10.000" -or $PercentComplete -eq "15.000" -or $PercentComplete -eq "20.000" -or $PercentComplete -eq "25.000" -or $PercentComplete -eq "30.000" -or $PercentComplete -eq "35.000" -or $PercentComplete -eq "40.000" -or $PercentComplete -eq "45.000" -or $PercentComplete -eq "50.000" -or $PercentComplete -eq "55.000" -or $PercentComplete -eq "60.000" -or $PercentComplete -eq "65.000" -or $PercentComplete -eq "70.00" -or $PercentComplete -eq "75.000" -or $PercentComplete -eq "80.000" -or $PercentComplete -eq "85.000" -or $PercentComplete -eq "90.000" -or $PercentComplete -eq "95.000")
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
		}
		
		
		
		#deal with sendas
		
		
		If (($Disp -notmatch "System.Object*") -and ($mail))
		{
			
			$ReportLine1 = [PSCustomObject][ordered] @{
				
				
				Samaccountname			   = $NTID
				displayName			       = $Disp
				mail					   = $mail
				useraccountcontrol		   = $UsrACCCtrl
				msExchRecipientTypeDetails = $MailboxType
				DN						   = $distName
				RecipientTypeDetails	   = $recipientdetailsEX
				Manager				       = $mgrAlias
				ManagerDisplayName		   = $managerDisp
				Manageremail			   = $mgrEmail
				msExchHideFromAddressLists = $HideAB
				Surname				       = $SN
				givenName				   = $firstname
				Employee				   = $CA9
				co						   = $country
				GPID					   = $BPtext3201
				Department				   = $dept
				Company				       = $comp
				UserPrincipalName          = $UPN
			}
			
			
			
			$ReportLine1 | Export-CSV $exportreportcsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
			
			
		}
	}
	
	
	
	DisconnectExchangeOnPrem
	
	
	
	
	#sortoutput so no blank lines and no duplicates
	$inputCsv1 = Import-Csv $exportreportcsv1 -delimiter "|" | Sort-Object * -Unique
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-localmbx-" + $nowfiledate + "-1.csv"
	$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$Outfile1 = "H:\M365Reports\EARL-LookupTbl-localmbx-" + $nowfiledate + "-2.csv"
	gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Outfile1
	
	
	if (Test-Path $exportreportcsv1)
	{
		Remove-Item $exportreportcsv1
	}
	
	if (Test-Path $finaloutcsv)
	{
		Remove-Item $finaloutcsv
	}
	
	Try
	{
		Map-Filewatcher
		Copy-item -path $Outfile1 -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Output file File Copied to FileWatcher $Outfile1 to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 360
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
			Add-Content $logfile "Output file File Copied to FileWatcher $Outfile1 to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 360
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile1 | $now"
		}
		
	}
	
}

function exportremoteusermbx
{
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export remotemailboxes for export"
	
	
	$count = ""
	$M = ""
	$reconnect = 0
	$MbxNumber = 0
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv1 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-1-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv2 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-2-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv3 = "H:\M365Reports\EARL-Prod-LookupTbl-remotmbx-hourly-3-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv4 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-4-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv5 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-5-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv6 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-6-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv7 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-7-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv8 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-8-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv9 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-9-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv10 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-10-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv11 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-11-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv12 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-12-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv13 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-13-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportusercsv14 = "H:\M365Reports\EARL-Prod-LookupTbl-remotembx-hourly-14-" + $nowfiledate + ".csv"
	
	#get-remotemailbox -resultsize 20 -filter 'RecipientTypedetails -eq "RemoteUserMailbox"'
	#$Mbx2 = Get-RemoteMailbox -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	
	#$Mbx2 = Get-RemoteMailbox -ResultSize unlimited -filter "WhenCreated -gt '$lasthour'" | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	$Mbx2 = Get-RemoteMailbox -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	
	$count = $mbx2.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count remote mailbox accounts to process"
	
	add-content $logfile  "LookupTable Exporting to file $exportreportusercsv1  for remote mbx "
	
	
	
	
	if ($count -ge 1)
	{
		
		$attributecountset1 = "0"
		$attributecountset2 = "0"
		$attributecountset3 = "0"
		$attributecountset4 = "0"
		$attributecountset5 = "0"
		$attributecountset6 = "0"
		$attributecountset7 = "0"
		$attributecountset8 = "0"
		$attributecountset9 = "0"
		$attributecountset10 = "0"
		$attributecountset11 = "0"
		$attributecountset12 = "0"
		$attributecountset13 = "0"
		$attributecountset14 = "0"
		$attributecountset15 = "0"
		$attributecountset16 = "0"
		$attributecountset17 = "0"
		$attributecountset18 = "0"
		$attributecountset19 = "0"
		$attributecountset20 = "0"
		
		
		ForEach ($M in $Mbx2)
		{
			$MbxNumber = $MbxNumber + 1
			$reconnect++
			$acc = ""
			$NTID = $null #SamAccountName
			$Display = $null #displayName
			$UsrACCCtrl = $null #useraccountcontrol
			$HideAB = $null #msExchHideFromAddressLists
			$SN = $null #sn
			$firstName = $Null #givenName
			$mail = $null #mail
			$managerDN = $null #manager
			$managerDisp = $null
			$mgrEmail = $null
			$mgrAlias = $null
			$rectypedetail = $Null #msExchRecipientTypeDetails
			$dept = $null #Department
			$BPtext3201 = $null #GPID
			$distName = $null # DN
			$country = $null #co
			$Comp = $null #Company
			$CA9 = $null #Employee
			$country = $null #co
			$recipientdetailsEX = $null #RecipientTypeDetails
			$managerDisp = $null #ManagerDisplayName
			$mgrEmail = $null #ManagerEmail
			$mgrAlias = $null #Manager
			
			$usr = $M.PrimarySmtpAddress
			$Display = $M.DisplayName
			$recipientdetailsEX = $M.RecipientTypeDetails
			

			
			$getacccount = 0
			try
			{
				$acc = Get-ADUser -filter 'mail -eq $usr' -properties * | Select-Object *
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "unable to get details for $usr - skipping"
			}
			
			$getacccount = $acc.count
			#$recpdetails = Get-Recipient -identity $usr -properties *
			#$recpdetails = Get-Recipient -identity $usr | Select-Object *
			
			if ($getacccount -lt 2)
			{
				
				
				[int]$Progress = $MbxNumber/$count * 100
				$PercentComplete = [math]::Round($Progress, 3)
				$disp = $acc.DisplayName
				$UPN = $acc.UserPrincipalName
				$MailboxType = $acc.msExchRecipientTypeDetails
				$mail = $acc.mail
				$SN = $acc.sn
				$firstName = $acc.GivenName
				$dept = $acc.Department
				$Comp = $acc.Company
				$country = $acc.co
				$UsrACCCtrl = $acc.useraccountcontrol
				$NTID = $acc.samaccountname
				$distName = $acc.DistinguishedName
				
				$managerDN = $acc.manager
				$BPtext3201 = $acc.'bp-Text32-01'
				$CA9 = $acc.extensionAttribute9
				$HideAB = $acc.msExchHideFromAddressLists
				
				
				if ($null -ne $managerDN)
				{
					$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
					
					$managerDisp = $mgrout.DisplayName
					$mgrEmail = $mgrout.PrimarySMTPAddress
					$mgrAlias = $mgrout.Alias
					
				}
				
				if ($recipientdetailsEX -eq "MailUser" -and ! $MailboxType)
				{
					$MailboxType = "128"
				}
				
				
				if (! $managerDisp)
				{
					
					
					$managerDisp = "NULL"
					
					
				}
				
				if (! $mgrEmail)
				{
					
					$mgrEmail = "NULL"
					
				}
				
				if (! $mgrAlias)
				{
					
					$mgrAlias = "NULL"
					
				}
				
				if (! $Comp)
				{
					$Comp = "NULL"
				}
				
				if (! $SN)
				{
					$SN = "NULL"
				}
				
				if (! $firstName)
				{
					$firstName = "NULL"
				}
				
				if (! $CA9)
				{
					$CA9 = "NULL"
				}
				
				if (! $BPtext3201)
				{
					$BPtext3201 = "NULL"
				}
				
				if (! $dept)
				{
					$dept = "NULL"
				}
				
				if (! $HideAB)
				{
					$HideAB = "False"
				}
				
				if (! $country)
				{
					$country = "NULL"
				}
				
				if (! $UPN)
				{
					$UPN = "NULL"
				}
				
				
				
				#$MbxStatus = $disp + " [" + $MbxNumber + "/" + $count + "]"
				
				
				
					if (($attributecountset1 -eq "0") -and ($PercentComplete -eq "5.000"))
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
						$attributecountset1 = "1" 
					}
					
					if (($attributecountset2 -eq "0") -and ($PercentComplete -eq "10.000"))
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
						 $attributecountset2 = "1"
					}
					
					if (($attributecountset3 -eq "0") -and ($PercentComplete -eq "15.000"))
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
						$attributecountset3= "1"
					}
					
					if (($attributecountset4 -eq "0") -and ($PercentComplete -eq "20.000"))
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
						$attributecountset4 = "1"
					}
					
					if (($attributecountset5 -eq "0") -and ($PercentComplete -eq "25.000"))
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
						$attributecountset5 = "1"
					}
					
					if (($attributecountset6 -eq "0") -and ($PercentComplete -eq "30.000"))
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
						$attributecountset6 = "1"
					}
					
					if (($attributecountset7 -eq "0") -and ($PercentComplete -eq "35.000"))
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
						$attributecountset7 = "1"
					}
				
				if (($attributecountset8 -eq "0") -and ($PercentComplete -eq "40.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset8 = "1"
				}
				
				if (($attributecountset9 -eq "0") -and ($PercentComplete -eq "45.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset9 = "1"
				}
				
				if (($attributecountset10 -eq "0") -and ($PercentComplete -eq "50.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset10 = "1"
				}
				
				
				if (($attributecountset11 -eq "0") -and ($PercentComplete -eq "55.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset11 = "1"
				}
				
				if (($attributecountset12 -eq "0") -and ($PercentComplete -eq "60.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset12 = "1"
				}
				
				if (($attributecountset13 -eq "0") -and ($PercentComplete -eq "65.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset13 = "1"
				}
				
				if (($attributecountset14 -eq "0") -and ($PercentComplete -eq "70.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset14 = "1"
				}
				
				if (($attributecountset15 -eq "0") -and ($PercentComplete -eq "75.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset15 = "1"
				}
				
				if (($attributecountset16 -eq "0") -and ($PercentComplete -eq "80.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset16 = "1"
				}
				
				if (($attributecountset17 -eq "0") -and ($PercentComplete -eq "85.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset17 = "1"
				}
				
				if (($attributecountset18 -eq "0") -and ($PercentComplete -eq "90.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset18 = "1"
				}
				
				if (($attributecountset19 -eq "0") -and ($PercentComplete -eq "95.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset19 = "1"
				}
				
				if (($attributecountset20 -eq "0") -and ($PercentComplete -eq "99.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset20 = "1"
				}
				
				
				
				
				
				#deal with sendas
				
				
				If ($Disp)
				{
					
					
					$ReportLine2 = [PSCustomObject][ordered] @{
						
						
						Samaccountname			   = $NTID
						displayName			       = $Disp
						mail					   = $mail
						useraccountcontrol		   = $UsrACCCtrl
						msExchRecipientTypeDetails = $MailboxType
						DN						   = $distName
						RecipientTypeDetails	   = $recipientdetailsEX
						Manager				       = $mgrAlias
						ManagerDisplayName		   = $managerDisp
						Manageremail			   = $mgrEmail
						msExchHideFromAddressLists = $HideAB
						Surname				       = $SN
						givenName				   = $firstname
						Employee				   = $CA9
						co						   = $country
						GPID					   = $BPtext3201
						Department				   = $dept
						Company				       = $comp
						UserPrincipalName		   = $UPN
					}
					
					
					if ($MbxNumber -le 10000)
					{
						$ReportLine2 | Export-CSV $exportreportusercsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 10000) -and ($MbxNumber -le 20000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 20000) -and ($MbxNumber -le 30000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 30000) -and ($MbxNumber -le 40000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 40000) -and ($MbxNumber -le 50000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 50000) -and ($MbxNumber -le 60000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 60000) -and ($MbxNumber -le 70000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 70000) -and ($MbxNumber -le 80000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 80000) -and ($MbxNumber -le 90000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 90000) -and ($MbxNumber -le 100000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 100000) -and ($MbxNumber -le 110000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 110000) -and ($MbxNumber -le 120000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 120000) -and ($MbxNumber -le 130000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if ($MbxNumber -gt 130000)
					{
						$ReportLine2 | Export-CSV $exportreportusercsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
				}
				
			}
		}
		
	
	
	
	DisconnectExchangeOnPrem
		
		if ($count -ge 1)
		{
			
			#sortoutput so no blank lines and no duplicates
			$inputCsv1 = Import-Csv $exportreportusercsv1 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-1-" + $nowfiledate + "-1.csv"
			$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile1 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-1-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Outfile1 -Encoding UTF8
			
			
			$checkfile = Test-Path -Path $exportreportusercsv3
			if ($checkfile -eq "True")
			{
				$inputCsv2 = Import-Csv $exportreportusercsv2 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv2 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-2-" + $nowfiledate + "-1.csv"
				$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile2 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-2-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Outfile2 -Encoding UTF8
			}
			
			$checkfile = Test-Path -Path $exportreportusercsv3
			if ($checkfile -eq "True") { 
			$inputCsv3 = Import-Csv $exportreportusercsv3 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv3 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-3-" + $nowfiledate + "-1.csv"
			$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile3 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-3-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Outfile3 -Encoding UTF8
		}
		
		
		Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv4
			if ($checkfile -eq "True")
			{
				$inputCsv4 = Import-Csv $exportreportusercsv4 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv4 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-4-" + $nowfiledate + "-1.csv"
				$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile4 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-4-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Outfile4 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv5
			if ($checkfile -eq "True")
			{
				$inputCsv5 = Import-Csv $exportreportusercsv5 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv5 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-5-" + $nowfiledate + "-1.csv"
				$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile5 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-5-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Outfile5 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv6
			if ($checkfile -eq "True")
			{
				$inputCsv6 = Import-Csv $exportreportusercsv6 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv6 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-6-" + $nowfiledate + "-1.csv"
				$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile6 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-6-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Outfile6 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv7
			if ($checkfile -eq "True")
			{
				$inputCsv7 = Import-Csv $exportreportusercsv7 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv7 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-7-" + $nowfiledate + "-1.csv"
				$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile7 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-7-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Outfile7 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv8
			if ($checkfile -eq "True")
			{
				$inputCsv8 = Import-Csv $exportreportusercsv8 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv8 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-8-" + $nowfiledate + "-1.csv"
				$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile8 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-8-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Outfile8 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv9
			if ($checkfile -eq "True")
			{
				$inputCsv9 = Import-Csv $exportreportusercsv9 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv9 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-9-" + $nowfiledate + "-1.csv"
				$inputCsv9 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile9 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-9-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv9 | ? { $_.trim() -ne "" } | set-content $Outfile9 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv10
			if ($checkfile -eq "True")
			{
				$inputCsv10 = Import-Csv $exportreportusercsv10 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv10 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-10-" + $nowfiledate + "-1.csv"
				$inputCsv10 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile10 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-10-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv10 | ? { $_.trim() -ne "" } | set-content $Outfile10 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv11
			if ($checkfile -eq "True")
			{
				$inputCsv11 = Import-Csv $exportreportusercsv11 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv11 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-11-" + $nowfiledate + "-1.csv"
				$inputCsv11 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile11 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-11-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv11 | ? { $_.trim() -ne "" } | set-content $Outfile11 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv12
			if ($checkfile -eq "True")
			{
				$inputCsv12 = Import-Csv $exportreportusercsv12 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv12 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-12-" + $nowfiledate + "-1.csv"
				$inputCsv12 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile12 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-12-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv12 | ? { $_.trim() -ne "" } | set-content $Outfile12 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv13
			if ($checkfile -eq "True")
			{
				$inputCsv13 = Import-Csv $exportreportusercsv13 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv13 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-13-" + $nowfiledate + "-1.csv"
				$inputCsv13 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile13 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-13-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv13 | ? { $_.trim() -ne "" } | set-content $Outfile13 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv14
			if ($checkfile -eq "True")
			{
				$inputCsv14 = Import-Csv $exportreportusercsv14 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv14 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-14-" + $nowfiledate + "-1.csv"
				$inputCsv14 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile14 = "H:\M365Reports\EARL-LookupTbl-remotembx-hourly-14-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv14 | ? { $_.trim() -ne "" } | set-content $Outfile14 -Encoding UTF8
				
			}
			#copy to filewatcher	
			
			
			Map-Filewatcher
			
			if (Test-Path $Outfile1)
			{
				Try
				{
					
					Copy-item -path $Outfile1 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile1 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Fileout to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Start-Sleep -Seconds 60
					#Map-Filewatcher
					Copy-item -path $Outfile2 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile2 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile2 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Start-Sleep -Seconds 60
					#Map-Filewatcher
					Copy-item -path $Outfile3 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile3 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile3 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Start-Sleep -Seconds 60
					#Map-Filewatcher
					Copy-item -path $Outfile4 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile4 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile4 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile5 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile5 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile6 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile6 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile7 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile7 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile8 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile8 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile9 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile9 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile10 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile10 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
						RemoveFilewatcher
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile10 | $now"
					}
				}
				
			}
			
			RemoveFilewatcher
			#Map-Filewatcher
			
		}
		
		#cleanup files
		
		if (Test-Path $exportreportusercsv1)
		{
			Remove-Item $exportreportusercsv1
		}
		
		if (Test-Path $finaloutcsv1)
		{
			Remove-Item $finaloutcsv1
		}
		
		if (Test-Path $exportreportusercsv2)
		{
			Remove-Item $exportreportusercsv2
		}
		
		if (Test-Path $finaloutcsv2)
		{
			Remove-Item $finaloutcsv2
		}
		
		if (Test-Path $exportreportusercsv3)
		{
			Remove-Item $exportreportusercsv3
		}
		
		if (Test-Path $finaloutcsv3)
		{
			Remove-Item $finaloutcsv3
		}
		
		if (Test-Path $exportreportusercsv4)
		{
			Remove-Item $exportreportusercsv4
		}
		
		if (Test-Path $finaloutcsv4)
		{
			Remove-Item $finaloutcsv4
		}
		
		if (Test-Path $exportreportusercsv5)
		{
			Remove-Item $exportreportusercsv5
		}
		
		if (Test-Path $finaloutcsv5)
		{
			Remove-Item $finaloutcsv5
		}
		
		if (Test-Path $exportreportusercsv6)
		{
			Remove-Item $exportreportusercsv6
		}
		
		if (Test-Path $finaloutcsv6)
		{
			Remove-Item $finaloutcsv6
		}
		
		if (Test-Path $exportreportusercsv7)
		{
			Remove-Item $exportreportusercsv7
		}
		
		if (Test-Path $finaloutcsv7)
		{
			Remove-Item $finaloutcsv7
		}
		
		#>
		
	}
	
}


function exportremotesharedroommbx
{
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export remotenonusermailboxes for export"
	
	
	$count = ""
	$M = ""
	$reconnect = 0
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv5 = "H:\M365Reports\EARL-Prod-LookupTbl-remotenonusermbx-" + $nowfiledate + ".csv"
	$lasthour = (get-date).adddays(-2)
	#$lasthour = (get-date).addhours(-1)
	$Mbx5 = Get-RemoteMailbox -ResultSize unlimited -filter 'RecipientTypedetails -eq "RemotesharedMailbox"' | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	$Mbx5 += Get-RemoteMailbox -ResultSize unlimited -filter 'RecipientTypedetails -eq "RemoteRoomMailbox"' | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	
	$count = $mbx5.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count remote mailbox accounts to process"
	
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv5 for remote non user mailboxes mbx "
	
	ForEach ($M in $Mbx5)
	{
		$MbxNumber++
		$reconnect++
		$NTID = $null #SamAccountName
		$Display = $null #displayName
		$UsrACCCtrl = $null #useraccountcontrol
		$HideAB = $null #msExchHideFromAddressLists
		$SN = $null #sn
		$firstName = $Null #givenName
		$mail = $null #mail
		$managerDN = $null #manager
		$managerDisp = $null
		$mgrEmail = $null
		$mgrAlias = $null
		$rectypedetail = $Null #msExchRecipientTypeDetails
		$dept = $null #Department
		$BPtext3201 = $null #GPID
		$distName = $null # DN
		$country = $null #co
		$Comp = $null #Company
		$CA9 = $null #Employee
		$country = $null #co
		$recipientdetailsEX = $null #RecipientTypeDetails
		$managerDisp = $null #ManagerDisplayName
		$mgrEmail = $null #ManagerEmail
		$mgrAlias = $null #Manager
		
		$usr = $M.PrimarySmtpAddress
		$Display = $M.DisplayName
		$recipientdetailsEX = $M.RecipientTypeDetails
		
		#write-host "Getting mailbox:: $usr"
		$acc = Get-ADUser -filter 'mail -eq $usr' -properties * | Select-Object *
		#$recpdetails = Get-Recipient -identity $usr -properties *
		#$recpdetails = Get-Recipient -identity $usr | Select-Object *
		
		[int]$Progress = $MbxNumber/$count * 100
		$PercentComplete = [math]::Round($Progress, 3)
		$disp = $acc.DisplayName
		$UPN = $acc.UserPrincipalName
		$MailboxType = $acc.msExchRecipientTypeDetails
		$mail = $acc.mail
		$SN = $acc.sn
		$firstName = $acc.GivenName
		$dept = $acc.Department
		$Comp = $acc.Company
		$country = $acc.co
		$UsrACCCtrl = $acc.useraccountcontrol
		$NTID = $acc.samaccountname
		$distName = $acc.DistinguishedName
		
		$managerDN = $acc.manager
		$BPtext3201 = $acc.'bp-Text32-01'
		$CA9 = $acc.extensionAttribute9
		$HideAB = $acc.msExchHideFromAddressLists
		
		
		if ($null -ne $managerDN)
		{
			$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
			
			$managerDisp = $mgrout.DisplayName
			$mgrEmail = $mgrout.PrimarySMTPAddress
			$mgrAlias = $mgrout.Alias
			
		}
		
		if ($recipientdetailsEX -eq "MailUser" -and ! $MailboxType)
		{
			$MailboxType = "128"
		}
		
		
		if (! $managerDisp)
		{
			
			
			$managerDisp = "NULL"
			
			
		}
		
		if (! $mgrEmail)
		{
			
			$mgrEmail = "NULL"
			
		}
		
		if (! $mgrAlias)
		{
			
			$mgrAlias = "NULL"
			
		}
		
		if (! $Comp)
		{
			$Comp = "NULL"
		}
		
		if (! $SN)
		{
			$SN = "NULL"
		}
		
		if (! $firstName)
		{
			$firstName = "NULL"
		}
		
		if (! $CA9)
		{
			$CA9 = "NULL"
		}
		
		if (! $BPtext3201)
		{
			$BPtext3201 = "NULL"
		}
		
		if (! $dept)
		{
			$dept = "NULL"
		}
		
		if (! $HideAB)
		{
			$HideAB = "False"
		}
		
		if (! $country)
		{
			$country = "NULL"
		}
		
		
		#$MbxStatus = $disp + " [" + $MbxNumber + "/" + $count + "]"
		
		
		
		if ($PercentComplete -eq "5.000" -or $PercentComplete -eq "10.000" -or $PercentComplete -eq "15.000" -or $PercentComplete -eq "20.000" -or $PercentComplete -eq "25.000" -or $PercentComplete -eq "30.000" -or $PercentComplete -eq "35.000" -or $PercentComplete -eq "40.000" -or $PercentComplete -eq "45.000" -or $PercentComplete -eq "50.000" -or $PercentComplete -eq "55.000" -or $PercentComplete -eq "60.000" -or $PercentComplete -eq "65.000" -or $PercentComplete -eq "70.00" -or $PercentComplete -eq "75.000" -or $PercentComplete -eq "80.000" -or $PercentComplete -eq "85.000" -or $PercentComplete -eq "90.000" -or $PercentComplete -eq "95.000")
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
		}
		
		
		
		#deal with sendas
		
		
		If (($Disp -notmatch "System.Object*") -and ($mail))
		{
			
			$ReportLine5 = [PSCustomObject][ordered] @{
				
				
				Samaccountname			   = $NTID
				displayName			       = $Disp
				mail					   = $mail
				useraccountcontrol		   = $UsrACCCtrl
				msExchRecipientTypeDetails = $MailboxType
				DN						   = $distName
				RecipientTypeDetails	   = $recipientdetailsEX
				Manager				       = $mgrAlias
				ManagerDisplayName		   = $managerDisp
				Manageremail			   = $mgrEmail
				msExchHideFromAddressLists = $HideAB
				Surname				       = $SN
				givenName				   = $firstname
				Employee				   = $CA9
				co						   = $country
				GPID					   = $BPtext3201
				Department				   = $dept
				Company				       = $comp
			}
			
			
			
			$ReportLine5 | Export-CSV $exportreportcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
			
			
		}
	}
	
	
	
	
	DisconnectExchangeOnPrem
	
	
	
	#sortoutput so no blank lines and no duplicates
	$inputCsv1 = Import-Csv $exportreportcsv5 -delimiter "|" | Sort-Object * -Unique
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-remotenonusermbx-" + $nowfiledate + "-1.csv"
	$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$Fileout = "H:\M365Reports\EARL-LookupTbl-remotenonusermbx-" + $nowfiledate + "-2.csv"
	gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Fileout
	

	
	<#
Try
{
	Map-Filewatcher
	Copy-item -path $Fileout -destination $filewatcherout
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	Add-Content $logfile "outfile File Copied to FileWatcher $Fileout to $filewatcherout [1st try] | $now"
	Start-Sleep -Seconds 360
	
	RemoveFilewatcher
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
		Add-Content $logfile "outfile File Copied to FileWatcher $Fileout to $filewatcherout [2nd try] | $now"
		Start-Sleep -Seconds 360
		
		RemoveFilewatcher
	}
	catch
	{
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Cannot copy files to FileWatcher $Fileout | $now"
	}
}

#cleanup files
if (Test-Path $finaloutcsv)
{
	Remove-Item $finaloutcsv
}

if (Test-Path $finaloutcsv2)
{
	Remove-Item $finaloutcsv2
}


if (Test-Path $exportreportcsv1)
{
	Remove-Item $exportreportcsv1
}

if (Test-Path $exportreportcsv2)
{
	Remove-Item $exportreportcsv2
}
#>
	
	
	
}



function exportmailuser
{
	
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export mailusers for export"
	
	

	$M = ""
	$reconnect = 0
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv3 = "H:\M365Reports\EARL-Prod-LookupTable-mailusers-" + $nowfiledate + ".csv"
	#$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress

	$Mbx3 = Get-Mailuser -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	$count = $mbx3.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count mailusers accounts to process"
	
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv3 for Mail Users "
	
	
	if ($count -ge 1)
	{
		
		
		$attributecountset1 = "0"
		$attributecountset2 = "0"
		$attributecountset3 = "0"
		$attributecountset4 = "0"
		$attributecountset5 = "0"
		$attributecountset6 = "0"
		$attributecountset7 = "0"
		$attributecountset8 = "0"
		$attributecountset9 = "0"
		$attributecountset10 = "0"
		$attributecountset11 = "0"
		$attributecountset12 = "0"
		$attributecountset13 = "0"
		$attributecountset14 = "0"
		$attributecountset15 = "0"
		$attributecountset16 = "0"
		$attributecountset17 = "0"
		$attributecountset18 = "0"
		$attributecountset19 = "0"
		$attributecountset20 = "0"
		
		
		ForEach ($M in $Mbx3)
		{
			$MbxNumber++
			$reconnect++
			$acc = ""
			$NTID = $null #SamAccountName
			$Display = $null #displayName
			$UsrACCCtrl = $null #useraccountcontrol
			$HideAB = $null #msExchHideFromAddressLists
			$SN = $null #sn
			$firstName = $Null #givenName
			$mail = $null #mail
			$managerDN = $null #manager
			$managerDisp = $null
			$mgrEmail = $null
			$mgrAlias = $null
			$rectypedetail = $Null #msExchRecipientTypeDetails
			$dept = $null #Department
			$BPtext3201 = $null #GPID
			$distName = $null # DN
			$country = $null #co
			$Comp = $null #Company
			$CA9 = $null #Employee
			$country = $null #co
			$recipientdetailsEX = $null #RecipientTypeDetails
			$managerDisp = $null #ManagerDisplayName
			$mgrEmail = $null #ManagerEmail
			$mgrAlias = $null #Manager
			
			$usr = $M.PrimarySmtpAddress
			$Display = $M.DisplayName
			$recipientdetailsEX = $M.RecipientTypeDetails
			
			#write-host "Getting mailbox:: $usr"
			$getacccount = 0
			try
			{
				$acc = Get-ADUser -filter 'mail -eq $usr' -properties * | Select-Object *
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "unable to get details for $usr - skipping"
			}
			
			$getacccount = $acc.count
			#$recpdetails = Get-Recipient -identity $usr -properties *
			#$recpdetails = Get-Recipient -identity $usr | Select-Object *
			
			if ($getacccount -lt 2)
			{
				[int]$Progress = $MbxNumber/$count * 100
				$PercentComplete = [math]::Round($Progress, 3)
				
				[string]$disp = $acc.DisplayName
				$UPN = $acc.UserPrincipalName
				$MailboxType = $acc.msExchRecipientTypeDetails
				$mail = $acc.mail
				$SN = $acc.sn
				$firstName = $acc.GivenName
				$dept = $acc.Department
				$Comp = $acc.Company
				$country = $acc.co
				$UsrACCCtrl = $acc.useraccountcontrol
				$NTID = $acc.samaccountname
				$distName = $acc.DistinguishedName
				
				$managerDN = $acc.manager
				$BPtext3201 = $acc.'bp-Text32-01'
				$CA9 = $acc.extensionAttribute9
				$HideAB = $acc.msExchHideFromAddressLists
				
				
				if ($null -ne $managerDN)
				{
					$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
					
					$managerDisp = $mgrout.DisplayName
					$mgrEmail = $mgrout.PrimarySMTPAddress
					$mgrAlias = $mgrout.Alias
					
				}
				
				if ($recipientdetailsEX -eq "MailUser" -and ! $MailboxType)
				{
					$MailboxType = "128"
				}
				
				
				if (! $managerDisp)
				{
					
					
					$managerDisp = "NULL"
					
					
				}
				
				if (! $mgrEmail)
				{
					
					$mgrEmail = "NULL"
					
				}
				
				if (! $mgrAlias)
				{
					
					$mgrAlias = "NULL"
					
				}
				
				if (! $Comp)
				{
					$Comp = "NULL"
				}
				
				if (! $SN)
				{
					$SN = "NULL"
				}
				
				if (! $firstName)
				{
					$firstName = "NULL"
				}
				
				if (! $CA9)
				{
					$CA9 = "NULL"
				}
				
				if (! $BPtext3201)
				{
					$BPtext3201 = "NULL"
				}
				
				if (! $dept)
				{
					$dept = "NULL"
				}
				
				if (! $HideAB)
				{
					$HideAB = "False"
				}
				
				if (! $country)
				{
					$country = "NULL"
				}
				
				if (! $UPN)
				{
					$UPN = "NULL"
				}
				
				
				#$MbxStatus = $disp + " [" + $MbxNumber + "/" + $count + "]"
				
				
				
				if (($attributecountset1 -eq "0") -and ($PercentComplete -eq "5.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset1 = "1"
				}
				
				if (($attributecountset2 -eq "0") -and ($PercentComplete -eq "10.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset2 = "1"
				}
				
				if (($attributecountset3 -eq "0") -and ($PercentComplete -eq "15.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset3 = "1"
				}
				
				if (($attributecountset4 -eq "0") -and ($PercentComplete -eq "20.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset4 = "1"
				}
				
				if (($attributecountset5 -eq "0") -and ($PercentComplete -eq "25.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset5 = "1"
				}
				
				if (($attributecountset6 -eq "0") -and ($PercentComplete -eq "30.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset6 = "1"
				}
				
				if (($attributecountset7 -eq "0") -and ($PercentComplete -eq "35.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset7 = "1"
				}
				
				if (($attributecountset8 -eq "0") -and ($PercentComplete -eq "40.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset8 = "1"
				}
				
				if (($attributecountset9 -eq "0") -and ($PercentComplete -eq "45.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset9 = "1"
				}
				
				if (($attributecountset10 -eq "0") -and ($PercentComplete -eq "50.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset10 = "1"
				}
				
				
				if (($attributecountset11 -eq "0") -and ($PercentComplete -eq "55.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset11 = "1"
				}
				
				if (($attributecountset12 -eq "0") -and ($PercentComplete -eq "60.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset12 = "1"
				}
				
				if (($attributecountset13 -eq "0") -and ($PercentComplete -eq "65.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset13 = "1"
				}
				
				if (($attributecountset14 -eq "0") -and ($PercentComplete -eq "70.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset14 = "1"
				}
				
				if (($attributecountset15 -eq "0") -and ($PercentComplete -eq "75.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset15 = "1"
				}
				
				if (($attributecountset16 -eq "0") -and ($PercentComplete -eq "80.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset16 = "1"
				}
				
				if (($attributecountset17 -eq "0") -and ($PercentComplete -eq "85.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset17 = "1"
				}
				
				if (($attributecountset18 -eq "0") -and ($PercentComplete -eq "90.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset18 = "1"
				}
				
				if (($attributecountset19 -eq "0") -and ($PercentComplete -eq "95.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset19 = "1"
				}
				
				if (($attributecountset20 -eq "0") -and ($PercentComplete -eq "99.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset20 = "1"
				}
				
				
				
				If ($Disp)
				{
					
					
					
					$ReportLine3 = [PSCustomObject][ordered] @{
						
						
						Samaccountname			   = $NTID
						displayName			       = $Disp
						mail					   = $mail
						useraccountcontrol		   = $UsrACCCtrl
						msExchRecipientTypeDetails = $MailboxType
						DN						   = $distName
						RecipientTypeDetails	   = $recipientdetailsEX
						Manager				       = $mgrAlias
						ManagerDisplayName		   = $managerDisp
						Manageremail			   = $mgrEmail
						msExchHideFromAddressLists = $HideAB
						Surname				       = $SN
						givenName				   = $firstname
						Employee				   = $CA9
						co						   = $country
						GPID					   = $BPtext3201
						Department				   = $dept
						Company				       = $comp
						UserPrincipalName		   = $UPN
					}
					
					
					
					$ReportLine3 | Export-CSV $exportreportcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					
				}
			}
			
		}
		
		
		
		
		
		DisconnectExchangeOnPrem
		
		
		
		#sortoutput so no blank lines and no duplicates
		$inputCsv1 = Import-Csv $exportreportcsv3 -delimiter "|" | Sort-Object * -Unique
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-mailuser-" + $nowfiledate + "-1.csv"
		$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$Fileout = "H:\M365Reports\EARL-LookupTbl-mailuser-" + $nowfiledate + "-2.csv"
		gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Fileout -Encoding UTF8
		
	
		
			Map-Filewatcher
Try
{
	
	Copy-item -path $Fileout -destination $filewatcherout
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "MailUsers File Copied to FileWatcher $Fileout to $filewatcherout [1st try] | $now"

	
	RemoveFilewatcher
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
				Add-Content $logfile "MailUsers File Copied to FileWatcher $Fileout to $filewatcherout [2nd try] | $now"

		RemoveFilewatcher
	}
	catch
	{
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Cannot copy file for mail users to FileWatcher $Fileout | $now"
	}
}

#cleanup files
if (Test-Path $finaloutcsv)
{
	Remove-Item $finaloutcsv
}

if (Test-Path $exportreportcsv3)
{
			Remove-Item $exportreportcsv3
}

#>
	}
	
	
}



function exportcontact
{
	
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export contacts for export"
	
	$Mbx5 = ""
	$exportreportcsv5 = ""
	$M = ""
	$reconnect = 0
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv5 = "H:\M365Reports\EARL-Prod-LookupTable-contacts-" + $nowfiledate + ".csv"
	#$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	#temp
	$lookuptime1 = (get-date).adddays(-366)
	#$lookuptime = (get-date).addhours(-3)
	Set-Variable -Name lasthour1 -Value $lookuptime1 -Option ReadOnly -Scope Script -Force
	
	#$Mbx5 = Get-MailContact -ResultSize unlimited | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	$Mbx5 = Get-MailContact -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	$count = $mbx5.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count mailcontacts accounts to process"
	
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv5 for Contacts "
	
	
	
	
	
	
	if ($count -ge 1)
	{
		
		
		$attributecountset1 = "0"
		$attributecountset2 = "0"
		$attributecountset3 = "0"
		$attributecountset4 = "0"
		$attributecountset5 = "0"
		$attributecountset6 = "0"
		$attributecountset7 = "0"
		$attributecountset8 = "0"
		$attributecountset9 = "0"
		$attributecountset10 = "0"
		$attributecountset11 = "0"
		$attributecountset12 = "0"
		$attributecountset13 = "0"
		$attributecountset14 = "0"
		$attributecountset15 = "0"
		$attributecountset16 = "0"
		$attributecountset17 = "0"
		$attributecountset18 = "0"
		$attributecountset19 = "0"
		$attributecountset20 = "0"
		
		
		ForEach ($M in $Mbx5)
		{
			$MbxNumber++
			$reconnect++
			$acc = ""
			$NTID = "Null"
			$Display = $null #displayName
			$UsrACCCtrl = $null #useraccountcontrol
			$HideAB = $null #msExchHideFromAddressLists
			$SN = $null #sn
			$firstName = $Null #givenName
			$mail = $null #mail
			$managerDN = $null #manager
			$managerDisp = $null
			$mgrEmail = $null
			$mgrAlias = $null
			$rectypedetail = $Null #msExchRecipientTypeDetails
			$dept = $null #Department
			$BPtext3201 = $null #GPID
			$distName = $null # DN
			$country = $null #co
			$Comp = $null #Company
			$CA9 = $null #Employee
			$country = $null #co
			$recipientdetailsEX = $null #RecipientTypeDetails
			$managerDisp = $null #ManagerDisplayName
			$mgrEmail = $null #ManagerEmail
			$mgrAlias = $null #Manager
			
			$usr = $M.PrimarySmtpAddress
			$Display = $M.DisplayName
			$recipientdetailsEX = $M.RecipientTypeDetails
			
			#write-host "Getting mailbox:: $usr"
			$getacccount = 0
			try
			{
				$acc = Get-Adobject -filter 'mail -eq $usr' -properties * | Select-Object *
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "unable to get details for $usr - skipping"
			}
			
			$getacccount = $acc.count
			#$recpdetails = Get-Recipient -identity $usr -properties *
			#$recpdetails = Get-Recipient -identity $usr | Select-Object *
			
			if ($getacccount -lt 2)
			{
				[int]$Progress = $MbxNumber/$count * 100
				$PercentComplete = [math]::Round($Progress, 3)
				
				[string]$disp = $acc.DisplayName
				$UPN = "NULL"
				$MailboxType = $acc.msExchRecipientTypeDetails
				$mail = $acc.mail
				$SN = $acc.sn
				$firstName = $acc.GivenName
				$dept = $acc.Department
				$Comp = $acc.Company
				$country = $acc.co
				$UsrACCCtrl = "0"
				$NTID = $acc.mail
				$distName = $acc.DistinguishedName
				
				$managerDN = $acc.manager
				$BPtext3201 = "NULL"
				$CA9 = "N"
				$HideAB = $acc.msExchHideFromAddressLists
				
				
				if ($null -ne $managerDN)
				{
					$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
					
					$managerDisp = $mgrout.DisplayName
					$mgrEmail = $mgrout.PrimarySMTPAddress
					$mgrAlias = $mgrout.Alias
					
				}
				
				if ($recipientdetailsEX -eq "MailContact" -and ! $MailboxType)
				{
					$MailboxType = "64"
				}
				
				
				if (! $managerDisp)
				{
					
					
					$managerDisp = "NULL"
					
					
				}
				
				if (! $mgrEmail)
				{
					
					$mgrEmail = "NULL"
					
				}
				
				if (! $mgrAlias)
				{
					
					$mgrAlias = "NULL"
					
				}
				
				if (! $Comp)
				{
					$Comp = "NULL"
				}
				
				if (! $SN)
				{
					$SN = "NULL"
				}
				
				if (! $firstName)
				{
					$firstName = "NULL"
				}
				
				if (! $CA9)
				{
					$CA9 = "NULL"
				}
				
				if (! $BPtext3201)
				{
					$BPtext3201 = "NULL"
				}
				
				if (! $dept)
				{
					$dept = "NULL"
				}
				
				if (! $HideAB)
				{
					$HideAB = "False"
				}
				
				if (! $country)
				{
					$country = "NULL"
				}
				
				if (! $UPN)
				{
					$UPN = "NULL"
				}
				
				
				#$MbxStatus = $disp + " [" + $MbxNumber + "/" + $count + "]"
				
				
				
				if (($attributecountset1 -eq "0") -and ($PercentComplete -eq "5.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset1 = "1"
				}
				
				if (($attributecountset2 -eq "0") -and ($PercentComplete -eq "10.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset2 = "1"
				}
				
				if (($attributecountset3 -eq "0") -and ($PercentComplete -eq "15.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Mailcontact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset3 = "1"
				}
				
				if (($attributecountset4 -eq "0") -and ($PercentComplete -eq "20.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset4 = "1"
				}
				
				if (($attributecountset5 -eq "0") -and ($PercentComplete -eq "25.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset5 = "1"
				}
				
				if (($attributecountset6 -eq "0") -and ($PercentComplete -eq "30.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset6 = "1"
				}
				
				if (($attributecountset7 -eq "0") -and ($PercentComplete -eq "35.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Mailcontact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset7 = "1"
				}
				
				if (($attributecountset8 -eq "0") -and ($PercentComplete -eq "40.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Mailcontact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset8 = "1"
				}
				
				if (($attributecountset9 -eq "0") -and ($PercentComplete -eq "45.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset9 = "1"
				}
				
				if (($attributecountset10 -eq "0") -and ($PercentComplete -eq "50.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset10 = "1"
				}
				
				
				if (($attributecountset11 -eq "0") -and ($PercentComplete -eq "55.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset11 = "1"
				}
				
				if (($attributecountset12 -eq "0") -and ($PercentComplete -eq "60.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset12 = "1"
				}
				
				if (($attributecountset13 -eq "0") -and ($PercentComplete -eq "65.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset13 = "1"
				}
				
				if (($attributecountset14 -eq "0") -and ($PercentComplete -eq "70.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset14 = "1"
				}
				
				if (($attributecountset15 -eq "0") -and ($PercentComplete -eq "75.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset15 = "1"
				}
				
				if (($attributecountset16 -eq "0") -and ($PercentComplete -eq "80.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset16 = "1"
				}
				
				if (($attributecountset17 -eq "0") -and ($PercentComplete -eq "85.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset17 = "1"
				}
				
				if (($attributecountset18 -eq "0") -and ($PercentComplete -eq "90.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset18 = "1"
				}
				
				if (($attributecountset19 -eq "0") -and ($PercentComplete -eq "95.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset19 = "1"
				}
				
				if (($attributecountset20 -eq "0") -and ($PercentComplete -eq "99.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset20 = "1"
				}
				
				
				
				If (($Disp) -and ($mail))
				{
					
					
					
					$ReportLine3 = [PSCustomObject][ordered] @{
						
						
						Samaccountname			   = $NTID
						displayName			       = $Disp
						mail					   = $mail
						useraccountcontrol		   = $UsrACCCtrl
						msExchRecipientTypeDetails = $MailboxType
						DN						   = $distName
						RecipientTypeDetails	   = $recipientdetailsEX
						Manager				       = $mgrAlias
						ManagerDisplayName		   = $managerDisp
						Manageremail			   = $mgrEmail
						msExchHideFromAddressLists = $HideAB
						Surname				       = $SN
						givenName				   = $firstname
						Employee				   = $CA9
						co						   = $country
						GPID					   = $BPtext3201
						Department				   = $dept
						Company				       = $comp
						UserPrincipalName		   = $UPN
					}
					
					
					
					$ReportLine3 | Export-CSV $exportreportcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					
				}
			}
			
		}
		
		
		
		
		
		DisconnectExchangeOnPrem
		
		
		
		#sortoutput so no blank lines and no duplicates
		$inputCsv1 = Import-Csv $exportreportcsv5 -delimiter "|" | Sort-Object * -Unique
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-mailcontact-" + $nowfiledate + "-1.csv"
		$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$Fileout = "H:\M365Reports\EARL-LookupTbl-mailcontact-" + $nowfiledate + "-2.csv"
		gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Fileout -Encoding UTF8
		
		
		
		Map-Filewatcher
		Try
		{
			
			Copy-item -path $Fileout -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "MailUsers File Copied to FileWatcher $Fileout to $filewatcherout [1st try] | $now"
			
			
			RemoveFilewatcher
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
				Add-Content $logfile "MailUsers File Copied to FileWatcher $Fileout to $filewatcherout [2nd try] | $now"
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy file for mail users to FileWatcher $Fileout | $now"
			}
		}
		
		#cleanup files
		if (Test-Path $finaloutcsv)
		{
			Remove-Item $finaloutcsv
		}
		
		if (Test-Path $exportreportcsv3)
		{
			Remove-Item $exportreportcsv3
		}
		
		#>
	}
	
	
}


function exportpriv
{
	
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export Privileged accounts for export"
	
	$Mbx5 = ""
	$exportreportcsv5 = ""
	$M = ""
	$reconnect = 0
	$accounts = ""
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv6 = "H:\M365Reports\EARL-Prod-LookupTable-priv-" + $nowfiledate + ".csv"
	#$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	#temp
	$lookuptime1 = (get-date).adddays(-366)
	#$lookuptime = (get-date).addhours(-3)
	Set-Variable -Name lasthour1 -Value $lookuptime1 -Option ReadOnly -Scope Script -Force
	
	#$Mbx5 = Get-MailContact -ResultSize unlimited | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	#$Mbx5 = Get-MailContact -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	$accounts = Get-User -filter "samaccountname -like 'svc*'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like '-svc*'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like 'serv-*'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like '-serv-*'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like '-tsk*'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like 'task*'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like 'tsk*'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like '-task-*'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like '-gbl*'" -resultsize unlimited | Select-Object *
	
	#"samaccountname -like 'svc*' -and WhenChanged -gt '$lookuptime1' "
	
	#get-user -OrganizationalUnit "bp1.ad.bp.com/Client/ORG/GenericAccounts"
	
	$count = $accounts.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count Privileged accounts to process"
	
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv6 for Privileged accounts "
	
	
	
	
	
	
	if ($count -ge 1)
	{
		[int]$skippedpriv = 0
		[int]$addedpriv = 0
		$attributecountset1 = "0"
		$attributecountset2 = "0"
		$attributecountset3 = "0"
		$attributecountset4 = "0"
		$attributecountset5 = "0"
		$attributecountset6 = "0"
		$attributecountset7 = "0"
		$attributecountset8 = "0"
		$attributecountset9 = "0"
		$attributecountset10 = "0"
		$attributecountset11 = "0"
		$attributecountset12 = "0"
		$attributecountset13 = "0"
		$attributecountset14 = "0"
		$attributecountset15 = "0"
		$attributecountset16 = "0"
		$attributecountset17 = "0"
		$attributecountset18 = "0"
		$attributecountset19 = "0"
		$attributecountset20 = "0"
		
		
		ForEach ($account in $accounts)
		{
			$MbxNumber++
			$reconnect++
			$acc1 = ""
			$NTID = ""
			$skipped = "NO"
			$Display = $null #displayName
			$UsrACCCtrl = $null #useraccountcontrol
			$HideAB = $null #msExchHideFromAddressLists
			$SN = $null #sn
			$firstName = $Null #givenName
			$mail = $null #mail
			$managerDN = $null #manager
			$managerDisp = $null
			$mgrEmail = $null
			$mgrAlias = $null
			$rectypedetail = $Null #msExchRecipientTypeDetails
			$dept = $null #Department
			$BPtext3201 = $null #GPID
			$distName = $null # DN
			$country = $null #co
			$Comp = $null #Company
			$CA9 = $null #Employee
			$country = $null #co
			$recipientdetailsEX = $null #RecipientTypeDetails
			$managerDisp = $null #ManagerDisplayName
			$mgrEmail = $null #ManagerEmail
			$mgrAlias = $null #Manager
			
			$usr = $account.samaccountname
			$Disp = $account.DisplayName
			$recipientdetailsEX = $account.RecipientTypeDetails
			
			#write-host "Getting mailbox:: $usr"
			$getacccount = 0
			try
			{
				$acc1 = Get-AdUser -filter 'samaccountname -eq $usr' -properties * | Select-Object *
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "unable to get details for $usr - skipping"
			}
			
			$getacccount = $acc1.count
			#$recpdetails = Get-Recipient -identity $usr -properties *
			#$recpdetails = Get-Recipient -identity $usr | Select-Object *
			
			if ($getacccount -lt 2)
			{
				[int]$Progress = $MbxNumber/$count * 100
				$PercentComplete = [math]::Round($Progress, 3)
				
				[string]$disp = $acc1.DisplayName
				$UPN = $acc1.UserPrincipalName
				$MailboxType = $acc1.msExchRecipientTypeDetails
				$mail = $acc1.mail
				$SN = $acc1.sn
				$firstName = $acc1.GivenName
				$dept = $acc1.Department
				$Comp = $acc1.Company
				$country = $acc1.co
				$UsrACCCtrl = $acc1.useraccountcontrol
				$NTID = $acc1.samaccountname
				$distName = $acc1.DistinguishedName
				
				$managerDN = $acc1.manager
				$BPtext3201 = $acc1.'bp-Text32-01'
				$CA9 = $acc1.extensionAttribute9
				$HideAB = $acc1.msExchHideFromAddressLists
				
				
				if (($null -ne $managerDN) -and ($UsrACCCtrl -ne "514" -or $UsrACCCtrl -ne "6650"))
				{
					$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
					
					$managerDisp = $mgrout.DisplayName
					$mgrEmail = $mgrout.PrimarySMTPAddress
					$mgrAlias = $mgrout.Alias
					
				}
				
				if ($recipientdetailsEX -eq "MailUser" -and ! $MailboxType)
				{
					$MailboxType = "128"
				}
				
				
				if (! $MailboxType)
				{
					$MailboxType = "NULL"
				}
				
				if (! $managerDisp)
				{
					
					
					$managerDisp = "NULL"
					
					
				}
				
				if (! $mgrEmail)
				{
					
					$mgrEmail = "NULL"
					
				}
				
				if (! $mgrAlias)
				{
					
					$mgrAlias = "NULL"
					
				}
				
				if (! $Comp)
				{
					$Comp = "NULL"
				}
				
				if (! $SN)
				{
					$SN = "NULL"
				}
				
				if (! $firstName)
				{
					$firstName = "NULL"
				}
				
				if (! $CA9)
				{
					$CA9 = "NULL"
				}
				
				if (! $BPtext3201)
				{
					$BPtext3201 = "NULL"
				}
				
				if (! $dept)
				{
					$dept = "NULL"
				}
				
				if (! $HideAB)
				{
					$HideAB = "False"
				}
				
				if (! $country)
				{
					$country = "NULL"
				}
				
				if (! $UPN)
				{
					$UPN = "NULL"
				}
				
				if ((! $Disp) -or ($Disp -eq ""))
				{
					$Display = "NULL"
				}
				
				
				if ($mail)
				{
					$skipped = "YES"
					
				}
				
				if (! $mail)
				{
					$mail = "NULL"
					
				}
				
				
				#$MbxStatus = $disp + " [" + $MbxNumber + "/" + $count + "]"
				
			
				
				if (($attributecountset1 -eq "0") -and ($PercentComplete -eq "5.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset1 = "1"
				}
				
				if (($attributecountset2 -eq "0") -and ($PercentComplete -eq "10.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset2 = "1"
				}
				
				if (($attributecountset3 -eq "0") -and ($PercentComplete -eq "15.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset3 = "1"
				}
				
				if (($attributecountset4 -eq "0") -and ($PercentComplete -eq "20.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset4 = "1"
				}
				
				if (($attributecountset5 -eq "0") -and ($PercentComplete -eq "25.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset5 = "1"
				}
				
				if (($attributecountset6 -eq "0") -and ($PercentComplete -eq "30.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset6 = "1"
				}
				
				if (($attributecountset7 -eq "0") -and ($PercentComplete -eq "35.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset7 = "1"
				}
				
				if (($attributecountset8 -eq "0") -and ($PercentComplete -eq "40.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset8 = "1"
				}
				
				if (($attributecountset9 -eq "0") -and ($PercentComplete -eq "45.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset9 = "1"
				}
				
				if (($attributecountset10 -eq "0") -and ($PercentComplete -eq "50.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset10 = "1"
				}
				
				
				if (($attributecountset11 -eq "0") -and ($PercentComplete -eq "55.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset11 = "1"
				}
				
				if (($attributecountset12 -eq "0") -and ($PercentComplete -eq "60.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset12 = "1"
				}
				
				if (($attributecountset13 -eq "0") -and ($PercentComplete -eq "65.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset13 = "1"
				}
				
				if (($attributecountset14 -eq "0") -and ($PercentComplete -eq "70.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset14 = "1"
				}
				
				if (($attributecountset15 -eq "0") -and ($PercentComplete -eq "75.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset15 = "1"
				}
				
				if (($attributecountset16 -eq "0") -and ($PercentComplete -eq "80.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset16 = "1"
				}
				
				if (($attributecountset17 -eq "0") -and ($PercentComplete -eq "85.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset17 = "1"
				}
				
				if (($attributecountset18 -eq "0") -and ($PercentComplete -eq "90.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset18 = "1"
				}
				
				if (($attributecountset19 -eq "0") -and ($PercentComplete -eq "95.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset19 = "1"
				}
				
				if (($attributecountset20 -eq "0") -and ($PercentComplete -eq "99.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Priv User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset20 = "1"
				}
				
				
				
				If ($UsrACCCtrl -eq "514")
				{
					$skipped = "YES"
					$skippedpriv = $skippedpriv + 1
				}
				
				If ($UsrACCCtrl -eq "66050")
				{
					$skipped = "YES"
					$skippedpriv = $skippedpriv + 1
				}
				
				
				If ($NTID -like "*-dou")
					{
					$skipped = "YES"
					
				}
				
				
				
				If ($skipped -eq "NO")
				{
					
					$addedpriv = $addedpriv + 1
					
					$ReportLine4 = [PSCustomObject][ordered] @{
						
						
						Samaccountname			   = $NTID
						displayName			       = $Disp
						mail					   = $mail
						useraccountcontrol		   = $UsrACCCtrl
						msExchRecipientTypeDetails = $MailboxType
						DN						   = $distName
						RecipientTypeDetails	   = $recipientdetailsEX
						Manager				       = $mgrAlias
						ManagerDisplayName		   = $managerDisp
						Manageremail			   = $mgrEmail
						msExchHideFromAddressLists = $HideAB
						Surname				       = $SN
						givenName				   = $firstname
						Employee				   = $CA9
						co						   = $country
						GPID					   = $BPtext3201
						Department				   = $dept
						Company				       = $comp
						UserPrincipalName		   = $UPN
					}
					
					
					
					$ReportLine4 | Export-CSV $exportreportcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					
				}
			}
			
		}
		
		
		
		
		
		DisconnectExchangeOnPrem
		
		
		Add-Content $logfile "Skipped $skippedpriv Priv users as they are disabled with no mail for import"
		Add-Content $logfile "Imported $addedpriv Priv Users as they are enabled and no mail currently in place"
		
		#sortoutput so no blank lines and no duplicates
		$inputCsv1 = Import-Csv $exportreportcsv6 -delimiter "|" | Sort-Object * -Unique
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-privuser-" + $nowfiledate + "-1.csv"
		$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$Fileout = "H:\M365Reports\EARL-LookupTbl-privuser-" + $nowfiledate + "-2.csv"
		gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Fileout -Encoding UTF8
		
		
		
		Map-Filewatcher
		Try
		{
			
			Copy-item -path $Fileout -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Priv Users File Copied to FileWatcher $Fileout to $filewatcherout [1st try] | $now"
			
			
			RemoveFilewatcher
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
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Fileout to $filewatcherout [2nd try] | $now"
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Fileout | $now"
			}
		}
		
		#cleanup files
		if (Test-Path $finaloutcsv)
		{
			Remove-Item $finaloutcsv
		}
		
		if (Test-Path $exportreportcsv6)
		{
			Remove-Item $exportreportcsv6
		}
		
		
	}
	
	
}


function exportgeneric
{
	
	[int]$MbxNumber = 0
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export Generic accounts for export to Temp Table"
	
	$Mbx5 = ""
	$exportreportcsv5 = ""
	$M = ""
	$reconnect = 0
	$accounts = ""
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	#$exportreportcsv7 = "H:\M365Reports\EARL-Prod-LookupTable-generic-" + $nowfiledate + ".csv"
	$exportreportgencsv = "H:\M365Reports\EARL-Prod-LookupTable-generic-All-" + $nowfiledate + ".csv"
	$exportreportgencsv1 = "H:\M365Reports\EARL-Prod-LookupTable-generic-1-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgencsv2 = "H:\M365Reports\EARL-Prod-LookupTable-generic-2-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgencsv3 = "H:\M365Reports\EARL-Prod-LookupTable-generic-3-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgencsv4 = "H:\M365Reports\EARL-Prod-LookupTable-generic-4-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgencsv5 = "H:\M365Reports\EARL-Prod-LookupTable-generic-5-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgencsv6 = "H:\M365Reports\EARL-Prod-LookupTable-generic-6-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgencsv7 = "H:\M365Reports\EARL-Prod-LookupTable-generic-7-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgencsv8 = "H:\M365Reports\EARL-Prod-LookupTable-generic-8-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgencsv9 = "H:\M365Reports\EARL-Prod-LookupTable-generic-9-" + $nowfiledate + ".csv"
	
	#$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	#temp
	$lookuptime1 = (get-date).adddays(-1250)
	#$lookuptime = (get-date).addhours(-3)
	Set-Variable -Name lasthour1 -Value $lookuptime1 -Option ReadOnly -Scope Script -Force
	
	#$accounts = Get-User -filter "samaccountname -like 'svc*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	
	#$accounts = get-user -OrganizationalUnit "bp1.ad.bp.com/Client/ORG/GenericAccounts" -filter "WindowsEmailAddress -ne '*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	$accounts = get-user -OrganizationalUnit "bp1.ad.bp.com/Client/ORG/GenericAccounts" -filter "WindowsEmailAddress -ne '*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	$accounts += get-user -OrganizationalUnit "bp1.ad.bp.com/Client/DWP/rEU/SUN/Users" -filter "WindowsEmailAddress -ne '*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	
	
	$count = $accounts.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count Generic accounts to process"
	
	add-content $logfile  "Refresh LookupTable Exporting to file  $exportreportcsv7 for Generic accounts "
	
	
	
	
	
	
	if ($count -ge 1)
	{
		[int]$skippedpriv = 0
		[int]$addedpriv = 0
		$attributecountset1 = "0"
		$attributecountset2 = "0"
		$attributecountset3 = "0"
		$attributecountset4 = "0"
		$attributecountset5 = "0"
		$attributecountset6 = "0"
		$attributecountset7 = "0"
		$attributecountset8 = "0"
		$attributecountset9 = "0"
		$attributecountset10 = "0"
		$attributecountset11 = "0"
		$attributecountset12 = "0"
		$attributecountset13 = "0"
		$attributecountset14 = "0"
		$attributecountset15 = "0"
		$attributecountset16 = "0"
		$attributecountset17 = "0"
		$attributecountset18 = "0"
		$attributecountset19 = "0"
		$attributecountset20 = "0"
		
		
		ForEach ($account in $accounts)
		{
			$MbxNumber++
			$reconnect++
			$acc1 = ""
			$NTID = ""
			$skipped = "NO"
			$Display = $null #displayName
			$UsrACCCtrl = $null #useraccountcontrol
			$HideAB = $null #msExchHideFromAddressLists
			$SN = $null #sn
			$firstName = $Null #givenName
			$mail = $null #mail
			$managerDN = $null #manager
			$managerDisp = $null
			$mgrEmail = $null
			$mgrAlias = $null
			$rectypedetail = $Null #msExchRecipientTypeDetails
			$dept = $null #Department
			$BPtext3201 = $null #GPID
			$distName = $null # DN
			$country = $null #co
			$Comp = $null #Company
			$CA9 = $null #Employee
			$country = $null #co
			$recipientdetailsEX = $null #RecipientTypeDetails
			$managerDisp = $null #ManagerDisplayName
			$mgrEmail = $null #ManagerEmail
			$mgrAlias = $null #Manager
			$descript = $null #Description
			$descript1 = $null
			$descript2 = $null
			$descript3 = $null
			$descript4 = $null
			$Display1 = $null
			$Display2 = $null
			$Display3 = $null
			$Display4 = $null
			$Display = $null
			$usr = $account.samaccountname
			$Display = $account.DisplayName
			$recipientdetailsEX = $account.RecipientTypeDetails
			
			#write-host "Getting mailbox:: $usr"
			$getacccount = 0
			try
			{
				$acc1 = Get-AdUser -filter 'samaccountname -eq $usr' -properties * | Select-Object *
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "unable to get details for $usr - skipping"
			}
			
			$getacccount = $acc1.count
			#$recpdetails = Get-Recipient -identity $usr -properties *
			#$recpdetails = Get-Recipient -identity $usr | Select-Object *
			
			if ($getacccount -lt 2)
			{
				[int]$Progress = $MbxNumber/$count * 100
				$PercentComplete = [math]::Round($Progress, 3)
				
				[string]$disp = $acc1.DisplayName
				$UPN = $acc1.UserPrincipalName
				$MailboxType = $acc1.msExchRecipientTypeDetails
				$mail = $acc1.mail
				$SN = $acc1.sn
				$firstName = $acc1.GivenName
				$dept = $acc1.Department
				$Comp = $acc1.Company
				$country = $acc1.co
				$UsrACCCtrl = $acc1.useraccountcontrol
				$NTID = $acc1.samaccountname
				$distName = $acc1.DistinguishedName
				$descript = $acc1.Description
				$managerDN = $acc1.manager
				$BPtext3201 = $acc1.'bp-Text32-01'
				$CA9 = $acc1.extensionAttribute9
				$HideAB = $acc1.msExchHideFromAddressLists
				
				
				if (($null -ne $managerDN) -and ($UsrACCCtrl -ne "514" -or $UsrACCCtrl -ne "6650"))
				{
					$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
					
					$managerDisp = $mgrout.DisplayName
					$mgrEmail = $mgrout.PrimarySMTPAddress
					$mgrAlias = $mgrout.Alias
					
				}
				
				if ($recipientdetailsEX -eq "MailUser" -and ! $MailboxType)
				{
					$MailboxType = "128"
				}
				
				
				if (! $MailboxType)
				{
					$MailboxType = "NULL"
				}
				
				
				if ($null -ne $descript)
				{
					[string]$descript1 = $descript -replace "`r`n", ""
					
					if ($descript1 -ne $descript)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed carriage return and new line in Description Field for $usr | $NTID | $now"
						[string]$descript = $descript1
					}
				}
				
				
				if ($null -ne $descript)
				{
					[string]$descript2 = $descript -replace "`n", ""
					
					
					if ($descript2 -ne $descript)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed new line in Description Field for $usr | $NTID | $now"
						[string]$descript = $descript2
					}
				}
				
				
				
				
				
				if ($null -ne $descript)
				{
					[string]$descript3 = $descript -replace "`r", ""
					
					if ($descript3 -ne $descript)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed carriage return in Description Field for $usr | $NTID | $now"
						[string]$descript = $descript3
					}
				}
				
				
				<#
				
				
				if ($descript)
				{
					$descript = $descript -replace '`r*`n*', ''
				}
				
				if ($descript)
				{
					$descript = $descript -replace '\r\n', ''
				}
				
				#>
				
				if ($null -ne $descript)
				{
					[string]$descript4 = $descript -replace '`t', ''
					
					if ($descript4 -ne $descript)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed Tab in Description Field for $usr | $NTID | $now"
						[string]$descript = $descript4
					}
				}
				
				
				
				
				
				if (!$descript)
				{
					$descript = "NULL"
				}
				
				
				if (! $managerDisp)
				{
					
					
					$managerDisp = "NULL"
					
					
				}
				
				if (! $mgrEmail)
				{
					
					$mgrEmail = "NULL"
					
				}
				
				if (! $mgrAlias)
				{
					
					$mgrAlias = "NULL"
					
				}
				
				if (! $Comp)
				{
					$Comp = "NULL"
				}
				
				if (! $SN)
				{
					$SN = "NULL"
				}
				
				if (! $firstName)
				{
					$firstName = "NULL"
				}
				
				if (! $CA9)
				{
					$CA9 = "NULL"
				}
				
				if (! $BPtext3201)
				{
					$BPtext3201 = "NULL"
				}
				
				if (! $dept)
				{
					$dept = "NULL"
				}
				
				if (! $HideAB)
				{
					$HideAB = "False"
				}
				
				if (! $country)
				{
					$country = "NULL"
				}
				
				if (! $UPN)
				{
					$UPN = "NULL"
				}
				
				if ($null -ne $Display)
				{
					[string]$Display1 = $Display -replace "`r`n", ""
					
					if ($Display1 -ne $Display)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed carriage return and new line in Display Field for $usr | $NTID | $now"
						[string]$Display = $Display1
					}
				}
				
				
				if ($null -ne $Display)
				{
					[string]$Display2 = $Display -replace "`n", ""
					
					
					if ($Display2 -ne $Display)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed new line in Display Field for $usr | $NTID | $now"
						[string]$Display = $Display2
					}
				}
				
				
				
				
				
				if ($null -ne $Display)
				{
					[string]$Display3 = $Display -replace "`r", ""
					
					if ($Display3 -ne $Display)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed carriage return in Display Field for $usr | $NTID | $now"
						[string]$Display = $Display3
					}
				}
				
				
				
				if ($null -ne $Display)
				{
					[string]$Display4 = $Display -replace '`t', ''
					
					if ($Display4 -ne $Display)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed Tab in Display Field for $usr | $NTID | $now"
						[string]$Display = $Display4
					}
				}
				
				
				if ((! $Disp) -or ($Disp -eq ""))
				{
					$Display = "NULL"
				}
				
				
				
				if ($mail)
				{
					$skipped = "YES"
					
				}
				
				if (! $mail)
				{
					$mail = "NULL"
					
				}
				
				
				#$MbxStatus = $disp + " [" + $MbxNumber + "/" + $count + "]"
				
				
				
				if (($attributecountset1 -eq "0") -and ($PercentComplete -eq "5.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset1 = "1"
				}
				
				if (($attributecountset2 -eq "0") -and ($PercentComplete -eq "10.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset2 = "1"
				}
				
				if (($attributecountset3 -eq "0") -and ($PercentComplete -eq "15.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset3 = "1"
				}
				
				if (($attributecountset4 -eq "0") -and ($PercentComplete -eq "20.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset4 = "1"
				}
				
				if (($attributecountset5 -eq "0") -and ($PercentComplete -eq "25.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset5 = "1"
				}
				
				if (($attributecountset6 -eq "0") -and ($PercentComplete -eq "30.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset6 = "1"
				}
				
				if (($attributecountset7 -eq "0") -and ($PercentComplete -eq "35.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset7 = "1"
				}
				
				if (($attributecountset8 -eq "0") -and ($PercentComplete -eq "40.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset8 = "1"
				}
				
				if (($attributecountset9 -eq "0") -and ($PercentComplete -eq "45.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset9 = "1"
				}
				
				if (($attributecountset10 -eq "0") -and ($PercentComplete -eq "50.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset10 = "1"
				}
				
				
				if (($attributecountset11 -eq "0") -and ($PercentComplete -eq "55.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset11 = "1"
				}
				
				if (($attributecountset12 -eq "0") -and ($PercentComplete -eq "60.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset12 = "1"
				}
				
				if (($attributecountset13 -eq "0") -and ($PercentComplete -eq "65.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset13 = "1"
				}
				
				if (($attributecountset14 -eq "0") -and ($PercentComplete -eq "70.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset14 = "1"
				}
				
				if (($attributecountset15 -eq "0") -and ($PercentComplete -eq "75.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset15 = "1"
				}
				
				if (($attributecountset16 -eq "0") -and ($PercentComplete -eq "80.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset16 = "1"
				}
				
				if (($attributecountset17 -eq "0") -and ($PercentComplete -eq "85.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset17 = "1"
				}
				
				if (($attributecountset18 -eq "0") -and ($PercentComplete -eq "90.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset18 = "1"
				}
				
				if (($attributecountset19 -eq "0") -and ($PercentComplete -eq "95.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset19 = "1"
				}
				
				if (($attributecountset20 -eq "0") -and ($PercentComplete -eq "99.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Generic User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset20 = "1"
				}
				
				
				
				If ($UsrACCCtrl -eq "514")
				{
					$skipped = "YES"
					$skippedpriv = $skippedpriv + 1
				}
				
				If ($UsrACCCtrl -eq "66050")
				{
					$skipped = "YES"
					$skippedpriv = $skippedpriv + 1
				}
				
				
				If ($NTID -like "*-dou")
				{
					$skipped = "YES"
					
				}
				
				
				
				If (($skipped -eq "NO") -and ($NTID))
				{
					
					$addedpriv = $addedpriv + 1
					
					$ReportLine5 = [PSCustomObject][ordered] @{
						
						
						Samaccountname			   = $NTID
						displayName			       = $Display
						mail					   = $mail
						useraccountcontrol		   = $UsrACCCtrl
						msExchRecipientTypeDetails = $MailboxType
						DN						   = $distName
						RecipientTypeDetails	   = $recipientdetailsEX
						Manager				       = $mgrAlias
						ManagerDisplayName		   = $managerDisp
						Manageremail			   = $mgrEmail
						msExchHideFromAddressLists = $HideAB
						Surname				       = $SN
						givenName				   = $firstname
						Employee				   = $CA9
						co						   = $country
						GPID					   = $BPtext3201
						Department				   = $dept
						Company				       = $comp
						UserPrincipalName		   = $UPN
						Description			       = $descript
					}
					
					if ($MbxNumber -le 5000)
					{
						$ReportLine5 | Export-CSV $exportreportgencsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 5000) -and ($MbxNumber -le 10000))
					{
						$ReportLine5 | Export-CSV $exportreportgencsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 10000) -and ($MbxNumber -le 15000))
					{
						$ReportLine5 | Export-CSV $exportreportgencsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 15000) -and ($MbxNumber -le 20000))
					{
						$ReportLine5 | Export-CSV $exportreportgencsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 20000) -and ($MbxNumber -le 25000))
					{
						$ReportLine5 | Export-CSV $exportreportgencsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
				
					if (($MbxNumber -gt 25000) -and ($MbxNumber -le 30000))
					{
						$ReportLine5 | Export-CSV $exportreportgencsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 30000) -and ($MbxNumber -le 35000))
					{
						$ReportLine5 | Export-CSV $exportreportgencsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if ($MbxNumber -gt 35000) 
					{
						$ReportLine5 | Export-CSV $exportreportgencsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					$ReportLine5 | Export-CSV $exportreportgencsv -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					
				}
			}
			
		}
		
	}
	
	
	
	DisconnectExchangeOnPrem
		
		
		Add-Content $logfile "Skipped $skippedpriv Generic users as they are disabled with no mail for import"
		Add-Content $logfile "Imported $addedpriv Generic Users as they are enabled and no mail currently in place"
		
		#sortoutput so no blank lines and no duplicates
		if ($count -ge 1)
		{
			#sortoutput so no blank lines and no duplicates
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
			$inputCsv1 = Import-Csv $exportreportgencsv1 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-generic-1" + $nowfiledate + "-1.csv"
			$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile1 = "H:\M365Reports\EARL-TempLookupTable-generic-1-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Outfile1 -Encoding UTF8
			
			$checkfile = Test-Path -Path $exportreportgencsv2
			if ($checkfile -eq "True")
			{
				Start-Sleep -s 2
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
				$inputCsv2 = Import-Csv $exportreportgencsv2 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv2 = "H:\M365Reports\EARL-LookupTbl-generic-2-" + $nowfiledate + "-1.csv"
				$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile2 = "H:\M365Reports\EARL-TempLookupTable-generic-2-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Outfile2 -Encoding UTF8
			}
			
			
			$checkfile = Test-Path -Path $exportreportgencsv3
			if ($checkfile -eq "True")
			{
				Start-Sleep -s 2
				$inputCsv3 = Import-Csv $exportreportgencsv3 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv3 = "H:\M365Reports\EARL-LookupTbl-generic-3-" + $nowfiledate + "-1.csv"
				$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile3 = "H:\M365Reports\EARL-TempLookupTable-generic-3-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Outfile3 -Encoding UTF8
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgencsv4
			if ($checkfile -eq "True")
			{
				$inputCsv4 = Import-Csv $exportreportgencsv4 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv4 = "H:\M365Reports\EARL-LookupTbl-generic-4-" + $nowfiledate + "-1.csv"
				$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile4 = "H:\M365Reports\EARL-TempLookupTable-generic-4-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Outfile4 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgencsv5
			if ($checkfile -eq "True")
			{
				$inputCsv5 = Import-Csv $exportreportgencsv5 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv5 = "H:\M365Reports\EARL-LookupTbl-generic-5-" + $nowfiledate + "-1.csv"
				$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile5 = "H:\M365Reports\EARL-TempLookupTable-generic-5-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Outfile5 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgencsv6
			if ($checkfile -eq "True")
			{
				$inputCsv6 = Import-Csv $exportreportgencsv6 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv6 = "H:\M365Reports\EARL-LookupTbl-generic-6-" + $nowfiledate + "-1.csv"
				$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile6 = "H:\M365Reports\EARL-TempLookupTable-generic-6-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Outfile6 -Encoding UTF8
				
			}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportgencsv7
		if ($checkfile -eq "True")
		{
			$inputCsv7 = Import-Csv $exportreportgencsv7 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv7 = "H:\M365Reports\EARL-LookupTbl-generic-7-" + $nowfiledate + "-1.csv"
			$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile7 = "H:\M365Reports\EARL-TempLookupTable-generic-7-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Outfile7 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportgencsv8
		if ($checkfile -eq "True")
		{
			$inputCsv8 = Import-Csv $exportreportgencsv8 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv8 = "H:\M365Reports\EARL-LookupTbl-generic-8-" + $nowfiledate + "-1.csv"
			$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile8 = "H:\M365Reports\EARL-TempLookupTable-generic-8-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Outfile8 -Encoding UTF8
			
		}
		
		
	}
	
	
	#copy to filewatcher	
	#Map-Filewatcher
		
	if (Test-Path $Outfile1)
	{
		Try
		{
			
			Copy-item -path $Outfile1 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile1 to $filewatcherout [1st try] | $now"
			Start-Sleep -Seconds 360
			
			
			#RemoveFilewatcher
			
		
			Map-Filewatcher
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
				Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile1 to $filewatcherout [2nd try] | $now"
				Start-Sleep -Seconds 360
				
				#RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy Generic Users to FileWatcher $Outfile1 | $now"
			}
		}
	}
	
	
	
	
	
	if (Test-Path $Outfile2)
	{
		Try
		{
			Start-Sleep -Seconds 60
			#Map-Filewatcher
			Copy-item -path $Outfile2 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile2 to $filewatcherout [1st try] | $now"
			Start-Sleep -Seconds 360
			
			
			
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
				Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile2 to $filewatcherout [2nd try] | $now"
				Start-Sleep -Seconds 360
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy Generic Users to FileWatcher $Outfile2 | $now"
			}
		}
	}
	
	
	if (Test-Path $Outfile3)
	{
		
		
		Try
		{
			Start-Sleep -Seconds 60
			#Map-Filewatcher
			Copy-item -path $Outfile3 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile3 to $filewatcherout [1st try] | $now"
			Start-Sleep -Seconds 360
			
			
			
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
				Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile3 to $filewatcherout [2nd try] | $now"
				Start-Sleep -Seconds 360
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy Generic Users to FileWatcher $Outfile3 | $now"
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
			Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile4 to $filewatcherout [1st try] | $now"
			Start-Sleep -Seconds 360
			
			
			
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
				Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile4 to $filewatcherout [2nd try] | $now"
				Start-Sleep -Seconds 360
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy Generic Users to FileWatcher $Outfile4 | $now"
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
			Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile5 to $filewatcherout [1st try] | $now"
			Start-Sleep -Seconds 360
			
			
			
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
				Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile5 to $filewatcherout [2nd try] | $now"
				Start-Sleep -Seconds 360
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy Generic Users to FileWatcher $Outfile5 | $now"
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
			Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile6 to $filewatcherout [1st try] | $now"
			Start-Sleep -Seconds 360
			
			
			
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
				Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile6 to $filewatcherout [2nd try] | $now"
				Start-Sleep -Seconds 360
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy Generic Users to FileWatcher $Outfile6 | $now"
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
			Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile7 to $filewatcherout [1st try] | $now"
			Start-Sleep -Seconds 360
			
			
			
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
				Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile7 to $filewatcherout [2nd try] | $now"
				Start-Sleep -Seconds 360
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy Generic Users to FileWatcher $Outfile7 | $now"
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
			Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile8 to $filewatcherout [1st try] | $now"
			Start-Sleep -Seconds 360
			
			
			
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
				Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile8 to $filewatcherout [2nd try] | $now"
				Start-Sleep -Seconds 360
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy Generic Users to FileWatcher $Outfile8 | $now"
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
			Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile9 to $filewatcherout [1st try] | $now"
			Start-Sleep -Seconds 360
			
			
			
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
				Add-Content $logfile "Generic Users File Copied to FileWatcher $Outfile9 to $filewatcherout [2nd try] | $now"
				Start-Sleep -Seconds 360
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy Generic Users to FileWatcher $Outfile9 | $now"
			}
		}
		
	}
			#>
	
	
	if (Test-Path $exportreportgencsv1)
	{
		Remove-Item $exportreportgencsv1
	}
	
	if (Test-Path $finaloutcsv)
	{
		Remove-Item $finaloutcsv
	}
	
	
	if (Test-Path $exportreportgencsv2)
	{
		Remove-Item $exportreportgencsv2
	}
	
	if (Test-Path $finaloutcsv2)
	{
		Remove-Item $finaloutcsv2
	}
	
	if (Test-Path $exportreportgrpcsv3)
	{
		Remove-Item $exportreportgrpcsv3
	}
	
	if (Test-Path $finaloutcsv4)
	{
		Remove-Item $finaloutcsv4
	}
	
	if (Test-Path $exportreportgrpcsv5)
	{
		Remove-Item $exportreportgrpcsv5
	}
	
	if (Test-Path $finaloutcsv5)
	{
		Remove-Item $finaloutcsv5
	}
	if (Test-Path $exportreportgrpcsv6)
	{
		Remove-Item $exportreportgrpcsv6
	}
	
	if (Test-Path $finaloutcsv7)
	{
		Remove-Item $finaloutcsv7
	}
	
	if (Test-Path $exportreportgrpcsv8)
	{
		Remove-Item $exportreportgrpcsv8
	}
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Finished Function to export Generic accounts for export"
}

function exportDL
{
	
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export Groups for export"
	
	
	
	$M = ""
	$reconnect = 0
	#$exportreportcsv4 = "H:\M365Reports\EARL-Prod-LookupTable-groups-" + $nowfiledate + ".csv"
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	
	$exportreportgrpcsv1 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-1-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv2 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-2-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv3 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-3-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv4 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-4-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv5 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-5-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv6 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-6-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv7 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-7-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv8 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-8-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv9 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-9-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv10 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-10-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv11 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-11-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv12 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-12-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv13 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-13-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv14 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-14-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv15 = "H:\M365Reports\EARL-Prod-LookupTable-group-Hourly-15-" + $nowfiledate + ".csv"
	
	
	
	
	$Mbx4 = Get-DistributionGroup -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'"  | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress
	
	$count = $mbx4.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count groups to process"
	
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv4 for groups "
	
	$attributecountset1 = "0"
	$attributecountset2 = "0"
	$attributecountset3 = "0"
	$attributecountset4 = "0"
	$attributecountset5 = "0"
	$attributecountset6 = "0"
	$attributecountset7 = "0"
	$attributecountset8 = "0"
	$attributecountset9 = "0"
	$attributecountset10 = "0"
	$attributecountset11 = "0"
	$attributecountset12 = "0"
	$attributecountset13 = "0"
	$attributecountset14 = "0"
	$attributecountset15 = "0"
	$attributecountset16 = "0"
	$attributecountset17 = "0"
	$attributecountset18 = "0"
	$attributecountset19 = "0"
	$attributecountset20 = "0"
	
	if ($count -ge 1)
	{
		ForEach ($M in $Mbx4)
		{
			$MbxNumber = $MbxNumber + 1
			$reconnect++
			$NTID = $null #SamAccountName
			$Display = $null #displayName
			$UsrACCCtrl = $null #useraccountcontrol
			$HideAB = $null #msExchHideFromAddressLists
			$SN = $null #sn
			$firstName = $Null #givenName
			$mail = $null #mail
			$managerDN = $null #manager
			$managerDisp = $null
			$mgrEmail = $null
			$mgrAlias = $null
			$rectypedetail = $Null #msExchRecipientTypeDetails
			$dept = $null #Department
			$BPtext3201 = $null #GPID
			$distName = $null # DN
			$country = $null #co
			$Comp = $null #Company
			$CA9 = $null #Employee
			$country = $null #co
			$recipientdetailsEX = $null #RecipientTypeDetails
			$managerDisp = $null #ManagerDisplayName
			$mgrEmail = $null #ManagerEmail
			$mgrAlias = $null #Manager
			
			$usr = $M.PrimarySmtpAddress
			$Display = $M.DisplayName
			$recipientdetailsEX = $M.RecipientTypeDetails
			
			
			try
			{
				$acc = Get-ADGroup -filter 'mail -eq $usr' -properties * | Select-Object *
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "unable to get details for $usr - skipping"
			}
			
			$getacccount = $acc.count
			#$recpdetails = Get-Recipient -identity $usr -properties *
			#$recpdetails = Get-Recipient -identity $usr | Select-Object *
			
			if ($getacccount -lt 2)
			{
				
				#$recpdetails = Get-Recipient -identity $usr -properties *
				#$recpdetails = Get-Recipient -identity $usr | Select-Object *
				
				[int]$Progress = $MbxNumber/$count * 100
				$PercentComplete = [math]::Round($Progress, 3)
				$disp = $acc.DisplayName
				$UPN = "NULL"
				$MailboxType = $acc.msExchRecipientTypeDetails
				$mail = $acc.mail
				$SN = $acc.DisplayName
				$firstName = "NULL"
				$dept = "NULL"
				$Comp = "NULL"
				$country = "NULL"
				$UsrACCCtrl = $acc.grouptype
				$NTID = $acc.samaccountname
				$distName = $acc.DistinguishedName
				
				$managerDN = $acc.ManagedBy
				$BPtext3201 = "NULL"
				$CA9 = "NULL"
				$HideAB = $acc.msExchHideFromAddressLists
				
				
				if ($null -ne $managerDN)
				{
					$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
					
					$managerDisp = $mgrout.DisplayName
					$mgrEmail = $mgrout.PrimarySMTPAddress
					$mgrAlias = $mgrout.Alias
					
				}
				
				if ($recipientdetailsEX -eq "MailUser" -and ! $MailboxType)
				{
					$MailboxType = "128"
				}
				
				
				if (! $managerDisp)
				{
					
					
					$managerDisp = "NULL"
					
					
				}
				
				if (! $mgrEmail)
				{
					
					$mgrEmail = "NULL"
					
				}
				
				if (! $mgrAlias)
				{
					
					$mgrAlias = "NULL"
					
				}
				
				if (! $Comp)
				{
					$Comp = "NULL"
				}
				
				if (! $SN)
				{
					$SN = "NULL"
				}
				
				if (! $firstName)
				{
					$firstName = "NULL"
				}
				
				if (! $CA9)
				{
					$CA9 = "NULL"
				}
				
				if (! $BPtext3201)
				{
					$BPtext3201 = "NULL"
				}
				
				if (! $dept)
				{
					$dept = "NULL"
				}
				
				if (! $HideAB)
				{
					$HideAB = "False"
				}
				
				if (! $country)
				{
					$country = "NULL"
				}
				
				if (! $MailboxType)
				{
					$MailboxType = 999
				}
				
				#$MbxStatus = $disp + " [" + $MbxNumber + "/" + $count + "]"
				
				
				
				if (($attributecountset1 -eq "0") -and ($PercentComplete -eq "5.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset1 = "1"
				}
				
				if (($attributecountset2 -eq "0") -and ($PercentComplete -eq "10.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset2 = "1"
				}
				
				if (($attributecountset3 -eq "0") -and ($PercentComplete -eq "15.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset3 = "1"
				}
				
				if (($attributecountset4 -eq "0") -and ($PercentComplete -eq "20.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset4 = "1"
				}
				
				if (($attributecountset5 -eq "0") -and ($PercentComplete -eq "25.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset5 = "1"
				}
				
				if (($attributecountset6 -eq "0") -and ($PercentComplete -eq "30.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset6 = "1"
				}
				
				if (($attributecountset7 -eq "0") -and ($PercentComplete -eq "35.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset7 = "1"
				}
				
				if (($attributecountset8 -eq "0") -and ($PercentComplete -eq "40.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset8 = "1"
				}
				
				if (($attributecountset9 -eq "0") -and ($PercentComplete -eq "45.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset9 = "1"
				}
				
				if (($attributecountset10 -eq "0") -and ($PercentComplete -eq "50.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset10 = "1"
				}
				
				
				if (($attributecountset11 -eq "0") -and ($PercentComplete -eq "55.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset11 = "1"
				}
				
				if (($attributecountset12 -eq "0") -and ($PercentComplete -eq "60.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset12 = "1"
				}
				
				if (($attributecountset13 -eq "0") -and ($PercentComplete -eq "65.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset13 = "1"
				}
				
				if (($attributecountset14 -eq "0") -and ($PercentComplete -eq "70.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset14 = "1"
				}
				
				if (($attributecountset15 -eq "0") -and ($PercentComplete -eq "75.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset15 = "1"
				}
				
				if (($attributecountset16 -eq "0") -and ($PercentComplete -eq "80.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset16 = "1"
				}
				
				if (($attributecountset17 -eq "0") -and ($PercentComplete -eq "85.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset17 = "1"
				}
				
				if (($attributecountset18 -eq "0") -and ($PercentComplete -eq "90.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset18 = "1"
				}
				
				if (($attributecountset19 -eq "0") -and ($PercentComplete -eq "95.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset19 = "1"
				}
				
				if (($attributecountset20 -eq "0") -and ($PercentComplete -eq "99.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset20 = "1"
				}
				
				
				
				
				
				If (($Disp -notmatch "System.Object*") -and ($mail))
				{
					
					$ReportLine4 = [PSCustomObject][ordered] @{
						
						
						Samaccountname			   = $NTID
						displayName			       = $Disp
						mail					   = $mail
						useraccountcontrol		   = $UsrACCCtrl
						msExchRecipientTypeDetails = $MailboxType
						DN						   = $distName
						RecipientTypeDetails	   = $recipientdetailsEX
						Manager				       = $mgrAlias
						ManagerDisplayName		   = $managerDisp
						Manageremail			   = $mgrEmail
						msExchHideFromAddressLists = $HideAB
						Surname				       = $SN
						givenName				   = $firstname
						Employee				   = $CA9
						co						   = $country
						GPID					   = $BPtext3201
						Department				   = $dept
						Company				       = $comp
						UserPrincipalName		   = $UPN
					}
					
					
					
					if ($MbxNumber -le 10000)
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 10000) -and ($MbxNumber -le 20000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 20000) -and ($MbxNumber -le 30000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 30000) -and ($MbxNumber -le 40000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 40000) -and ($MbxNumber -le 50000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 50000) -and ($MbxNumber -le 60000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 60000) -and ($MbxNumber -le 70000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					if (($MbxNumber -gt 70000) -and ($MbxNumber -le 80000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 80000) -and ($MbxNumber -le 90000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 90000) -and ($MbxNumber -le 100000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 100000) -and ($MbxNumber -le 110000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 110000) -and ($MbxNumber -le 120000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 120000) -and ($MbxNumber -le 130000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if ($MbxNumber -gt 130000)
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
				}
			}
			
		}
		
	
	
	
	DisconnectExchangeOnPrem
		
		
		if ($count -ge 1)
		{
			#sortoutput so no blank lines and no duplicates
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
			$inputCsv1 = Import-Csv $exportreportgrpcsv1 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-1" + $nowfiledate + "-1.csv"
			$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile1 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-1" + $nowfiledate + "-2.csv"
			gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Outfile1 -Encoding UTF8
			
			Start-Sleep -s 2
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
			$inputCsv2 = Import-Csv $exportreportgrpcsv2 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv2 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-2-" + $nowfiledate + "-1.csv"
			$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile2 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-2-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Outfile2 -Encoding UTF8
			
			Start-Sleep -s 2
			$inputCsv3 = Import-Csv $exportreportgrpcsv3 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv3 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-3-" + $nowfiledate + "-1.csv"
			$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile3 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-3-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Outfile3 -Encoding UTF8
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv4
			if ($checkfile -eq "True")
			{
				$inputCsv4 = Import-Csv $exportreportgrpcsv4 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv4 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-4-" + $nowfiledate + "-1.csv"
				$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile4 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-4-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Outfile4 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv5
			if ($checkfile -eq "True")
			{
				$inputCsv5 = Import-Csv $exportreportgrpcsv5 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv5 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-5-" + $nowfiledate + "-1.csv"
				$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile5 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-5-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Outfile5 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv6
			if ($checkfile -eq "True")
			{
				$inputCsv6 = Import-Csv $exportreportgrpcsv6 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv6 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-6-" + $nowfiledate + "-1.csv"
				$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile6 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-6-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Outfile6 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv7
			if ($checkfile -eq "True")
			{
				$inputCsv7 = Import-Csv $exportreportgrpcsv7 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv7 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-7-" + $nowfiledate + "-1.csv"
				$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile7 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-7-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Outfile7 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv8
			if ($checkfile -eq "True")
			{
				$inputCsv8 = Import-Csv $exportreportgrpcsv8 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv8 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-8-" + $nowfiledate + "-1.csv"
				$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile8 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-8-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Outfile8 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv9
			if ($checkfile -eq "True")
			{
				$inputCsv9 = Import-Csv $exportreportgrpcsv9 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv9 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-9-" + $nowfiledate + "-1.csv"
				$inputCsv9 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile9 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-9-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv9 | ? { $_.trim() -ne "" } | set-content $Outfile9 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv10
			if ($checkfile -eq "True")
			{
				$inputCsv10 = Import-Csv $exportreportgrpcsv10 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv10 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-10-" + $nowfiledate + "-1.csv"
				$inputCsv10 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile10 = "H:\M365Reports\EARL-LookupTbl-groups10-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv10 | ? { $_.trim() -ne "" } | set-content $Outfile10 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv11
			if ($checkfile -eq "True")
			{
				$inputCsv11 = Import-Csv $exportreportgrpcsv11 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv11 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-11-" + $nowfiledate + "-1.csv"
				$inputCsv11 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile11 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-11-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv11 | ? { $_.trim() -ne "" } | set-content $Outfile11 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv12
			if ($checkfile -eq "True")
			{
				$inputCsv12 = Import-Csv $exportreportgrpcsv12 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv12 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-12-" + $nowfiledate + "-1.csv"
				$inputCsv12 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile12 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-12-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv12 | ? { $_.trim() -ne "" } | set-content $Outfile12 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv13
			if ($checkfile -eq "True")
			{
				$inputCsv13 = Import-Csv $exportreportgrpcsv13 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv13 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-13-" + $nowfiledate + "-1.csv"
				$inputCsv13 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile13 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-13-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv13 | ? { $_.trim() -ne "" } | set-content $Outfile13 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv14
			if ($checkfile -eq "True")
			{
				$inputCsv14 = Import-Csv $exportreportgrpcsv14 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv14 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-14-" + $nowfiledate + "-1.csv"
				$inputCsv14 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile14 = "H:\M365Reports\EARL-LookupTbl-groups-Hourly-14-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv14 | ? { $_.trim() -ne "" } | set-content $Outfile14 -Encoding UTF8
				
			}
			
			#copy to filewatcher	
			Map-Filewatcher
			
			if (Test-Path $Outfile1)
			{
				Try
				{
					
					Copy-item -path $Outfile1 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile1 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile1 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Start-Sleep -Seconds 60
					#Map-Filewatcher
					Copy-item -path $Outfile2 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile2 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile2 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Start-Sleep -Seconds 60
					#Map-Filewatcher
					Copy-item -path $Outfile3 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile3 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile3 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile4 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile4 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile5 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile5 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile6 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile6 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile7 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile7 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile8 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile8 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile9 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile9 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
						RemoveFilewatcher
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile9 | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile10 to $filewatcherout [1st try] | $now"
					Start-Sleep -Seconds 360
					
					
					
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile10 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
						RemoveFilewatcher
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile10 | $now"
					}
				}
				
			}
			
			#cleanup files
			
			if (Test-Path $exportreportgrpcsv1)
			{
				Remove-Item $exportreportgrpcsv1
			}
			
			if (Test-Path $finaloutcsv)
			{
				Remove-Item $finaloutcsv
			}
			
			
			if (Test-Path $exportreportgrpcsv2)
			{
				Remove-Item $exportreportgrpcsv2
			}
			
			if (Test-Path $finaloutcsv2)
			{
				Remove-Item $finaloutcsv2
			}
			
			
			if (Test-Path $exportreportgrpcsv3)
			{
				Remove-Item $exportreportgrpcsv3
			}
			
			if (Test-Path $finaloutcsv3)
			{
				Remove-Item $finaloutcsv3
			}
			
			if (Test-Path $exportreportgrpcsv4)
			{
				Remove-Item $exportreportgrpcsv4
			}
			
			if (Test-Path $finaloutcsv4)
			{
				Remove-Item $finaloutcsv4
			}
			
			
			if (Test-Path $exportreportgrpcsv5)
			{
				Remove-Item $exportreportgrpcsv5
			}
			
			if (Test-Path $finaloutcsv5)
			{
				Remove-Item $finaloutcsv5
			}
			
			
			if (Test-Path $exportreportgrpcsv6)
			{
				Remove-Item $exportreportgrpcsv6
			}
			
			if (Test-Path $finaloutcsv6)
			{
				Remove-Item $finaloutcsv6
			}
			
			if (Test-Path $exportreportgrpcsv7)
			{
				Remove-Item $exportreportgrpcsv7
			}
			
			if (Test-Path $finaloutcsv7)
			{
				Remove-Item $finaloutcsv7
			}
			
			if (Test-Path $exportreportgrpcsv8)
			{
				Remove-Item $exportreportgrpcsv8
			}
			
			if (Test-Path $finaloutcsv8)
			{
				Remove-Item $finaloutcsv8
			}
			
			
			if (Test-Path $exportreportgrpcsv9)
			{
				Remove-Item $exportreportgrpcsv9
			}
			
			if (Test-Path $finaloutcsv9)
			{
				Remove-Item $finaloutcsv9
			}
			
			if (Test-Path $exportreportgrpcsv10)
			{
				Remove-Item $exportreportgrpcsv10
			}
			
			if (Test-Path $finaloutcsv10)
			{
				Remove-Item $finaloutcsv10
			}
			
			if (Test-Path $exportreportgrpcsv11)
			{
				Remove-Item $exportreportgrpcsv11
			}
			
			if (Test-Path $finaloutcsv11)
			{
				Remove-Item $finaloutcsv11
			}
			
			if (Test-Path $exportreportgrpcsv12)
			{
				Remove-Item $exportreportgrpcsv12
			}
			
			if (Test-Path $finaloutcsv12)
			{
				Remove-Item $finaloutcsv12
			}
			
			if (Test-Path $exportreportgrpcsv13)
			{
				Remove-Item $exportreportgrpcsv13
			}
			
			if (Test-Path $finaloutcsv13)
			{
				Remove-Item $finaloutcsv13
			}
			
			if (Test-Path $exportreportgrpcsv14)
			{
				Remove-Item $exportreportgrpcsv14
			}
			
			if (Test-Path $finaloutcsv14)
			{
				Remove-Item $finaloutcsv14
			}
			
		}
		RemoveFilewatcher
	}
}


#exportDL
#exportmailuser
#exportlocalmbx
#exportremoteusermbx
#exportcontact
#exportpriv
exportgeneric

#exportremotesharedroommbx



$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "All finished with exports for LDAP replacement Temp LookupTable Generic Users | $now"
DisconnectExchangeOnPrem
RemoveFilewatcher

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "Closing down script - bye $now"
Stop-Transcript

Disconnect-EXO
DisconnectExchangeOnPrem

Exit-PSSession
Exit



