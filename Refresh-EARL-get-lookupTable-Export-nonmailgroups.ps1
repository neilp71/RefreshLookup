



<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.148
	 Created on:    07/08/2023 14:00
	 Created by:   	Neil Poultney
	 Organization: 	NDP Consultancy Ltd - neil.poultney@ndpconsultancyltd.co.uk
	 Filename:     	Refresh-EARL-get-lookupTable-Export-nonmailgroups.ps1
	===========================================================================
	.DESCRIPTION
		Exports for EARL MailDb to Lookup Table on hourly basis.

		Change Log
		V1.00, 19/01/2024 - Initial full version
		

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
$transcriptlog = "H:\EARLTranscripts\LookupTbl\refresh-lookup-export-nonmailgroup-" + $nowfiledate + ".log"

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
	$logfilelocation = $loglocation + "Refresh-Lookup-Table-nonmailgroup-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
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
	$logfilelocation = $loglocation + "Refresh-Lookup-Table-nonmailgroup-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
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
			
			Import-PsSession $exchangesession -AllowClobber
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
			
			Import-PsSession $exchangesession -AllowClobber
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






function exportNonMailGroups
{
	
	

	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export nonmail Groups for export"
	
	
	
	$M = ""
	$reconnect = 0
	#$exportreportcsv4 = "H:\M365Reports\EARL-Prod-LookupTable-groups-" + $nowfiledate + ".csv"
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	
	$exportreportgrpcsv1 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-1-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv2 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-2-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv3 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-3-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv4 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-4-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv5 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-5-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv6 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-6-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv7 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-7-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv8 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-8-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv9 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-9-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv10 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-10-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv11 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-11-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv12 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-12-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv13 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-13-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv14 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-14-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportgrpcsv15 = "H:\M365Reports\EARL-Prod-TempLookupTable-nonmailgroup-15-" + $nowfiledate + ".csv"
	
	
	$lookuptime1 = (get-date).adddays(-720)
	$lookuptime = (get-date).addhours(-3)
	Set-Variable -Name lasthour -Value $lookuptime1 -Option ReadOnly -Scope Script -Force
	
	#$Mbx5 = Get-MailContact -ResultSize unlimited | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	
	#$Adgroup = Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups,OU=MOR,OU=rEU,OU=Client,DC=bp1,DC=ad,DC=bp,DC=com" | where { $_.whenChanged -gt $lasthour } 
	

	$Adgroup = @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups,OU=MOR,OU=rEU,OU=Client,DC=bp1,DC=ad,DC=bp,DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=MailEnabledGroups,OU=ORG,OU=Client,DC=bp1,DC=ad,DC=bp,DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=SIN, OU=rSE, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=TKP, OU=rSE, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=HOU, OU=rAM, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=CHI, OU=rAM, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=CHP, OU=rAM, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=SUN, OU=rEU, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=DefaultDistributionLists, OU=Unity, OU=ORG, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=NAP, OU=rAM, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=BAK, OU=rEU, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=POS, OU=rAM, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=HTN, OU=rST, OU=DWP, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=LDC, OU=rST, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=LCS, OU=rST, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=HTN, OU=rST, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=IAM, OU=rST, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=IDA, OU=rST, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=LST, OU=rST, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=UEU, OU=rEU, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=RoleGroups, OU=WHI, OU=rAM, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	$Adgroup += @(Get-ADGroup -Filter 'mail -notlike "*"' -Property whenChanged -SearchBase "OU=BOC, OU=rEC, OU=Client, DC=bp1, DC=ad, DC=bp, DC=com" | where { $_.whenChanged -gt $lasthour })
	
	$count = $Adgroup.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count non mail groups to process"
	
	add-content $logfile  "LookupTable Exporting to files for non mail groups "
	
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
	
	ConnectExchangeonPrem
	
	if ($count -ge 1)
	{
		ForEach ($M in $Adgroup)
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
			$usr = $M.DistinguishedName
			$Display = $M.DisplayName
			$recipientdetailsEX = $M.RecipientTypeDetails
			$GrpName = $M.Name
			
			try
			{
				$acc = Get-ADGroup -filter 'DistinguishedName -eq $usr' -properties * | Select-Object *
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
				$MailboxType = "NULL"
				$mail = $acc.mail
				$SN = $acc.DisplayName
				$firstName = "NULL"
				$dept = "NULL"
				$Comp = "NULL"
				$country = "NULL"
				$UsrACCCtrl = $acc.grouptype
				$NTID = $acc.samaccountname
				$distName = $acc.DistinguishedName
				$descript = $acc.Description
				$managerDN = $acc.ManagedBy
				$BPtext3201 = "NULL"
				$CA9 = "NULL"
				$HideAB = $acc.msExchHideFromAddressLists
				
				if (!$mail)
				{
					$mail = "NULL"
				}
				
				if ($null -ne $managerDN)
				{
					$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
					
					$managerDisp = $mgrout.DisplayName
					$mgrEmail = $mgrout.PrimarySMTPAddress
					$mgrAlias = $mgrout.Alias
					
				}
				
				if (!$recipientdetailsEX )
				{
					$recipientdetailsEX = "NonMailEnabledGroup"
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
				
				
				if ((! $Display) -or ($Display -eq ""))
				{
					$Display = "NULL"
				}
				
				
				if ($descript)
				{
					$descript.replace('`r`n', ' ')
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
				
				if (! $MailboxType)
				{
					$MailboxType = 999
				}
				
				#$MbxStatus = $disp + " [" + $MbxNumber + "/" + $count + "]"
				
				
				
				if (($attributecountset1 -eq "0") -and ($PercentComplete -eq "5.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed non mail Group number : $MbxNumber | Percent Complete: $PercentComplete | $now "
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
				
				
				
				
				
				If (($Disp -notmatch "System.Object*") -and ($NTID))
				{
					
					$ReportLine4 = [PSCustomObject][ordered] @{
						
						
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
					
					
					
					if ($MbxNumber -le 3000)
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 3000) -and ($MbxNumber -le 6000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 6000) -and ($MbxNumber -le 9000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 9000) -and ($MbxNumber -le 12000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 12000) -and ($MbxNumber -le 15000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 15000) -and ($MbxNumber -le 18000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 18000) -and ($MbxNumber -le 21000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					if (($MbxNumber -gt 21000) -and ($MbxNumber -le 24000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 24000) -and ($MbxNumber -le 27000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 27000) -and ($MbxNumber -le 30000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 30000) -and ($MbxNumber -le 33000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 33000) -and ($MbxNumber -le 36000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 36000) -and ($MbxNumber -le 39000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if ($MbxNumber -gt 39000)
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
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv1
			if ($checkfile -eq "True")
			{
				
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Processing CSv file $exportreportgrpcsv1 | $now "
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
				$inputCsv1 = Import-Csv $exportreportgrpcsv1 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-1-" + $nowfiledate + "-1.csv"
				$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile1 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-1-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Outfile1 -Encoding UTF8
			}
			
			$checkfile = Test-Path -Path $exportreportgrpcsv2
			if ($checkfile -eq "True")
			{
				Start-Sleep -s 2
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Processing CSV file $exportreportgrpcsv2 | $now "
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
				$inputCsv2 = Import-Csv $exportreportgrpcsv2 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv2 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-2-" + $nowfiledate + "-1.csv"
				$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile2 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-2-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Outfile2 -Encoding UTF8
			}
			
			$checkfile = Test-Path -Path $exportreportgrpcsv3
			if ($checkfile -eq "True")
			{
				Start-Sleep -s 2
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Processing CSV file $exportreportgrpcsv3 | $now "
				$inputCsv3 = Import-Csv $exportreportgrpcsv3 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv3 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-3-" + $nowfiledate + "-1.csv"
				$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile3 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-3-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Outfile3 -Encoding UTF8
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv4
			if ($checkfile -eq "True")
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Processing CSV file $exportreportgrpcsv4 | $now "
				$inputCsv4 = Import-Csv $exportreportgrpcsv4 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv4 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-4-" + $nowfiledate + "-1.csv"
				$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile4 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-4-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Outfile4 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv5
			if ($checkfile -eq "True")
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Processing CSV file $exportreportgrpcsv5 | $now "
				$inputCsv5 = Import-Csv $exportreportgrpcsv5 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv5 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-5-" + $nowfiledate + "-1.csv"
				$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile5 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-5-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Outfile5 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv6
			if ($checkfile -eq "True")
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Processing CSV file $exportreportgrpcsv6 | $now "
				$inputCsv6 = Import-Csv $exportreportgrpcsv6 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv6 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-6-" + $nowfiledate + "-1.csv"
				$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile6 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-6-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Outfile6 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv7
			if ($checkfile -eq "True")
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Processing CSV file $exportreportgrpcsv7 | $now "
				$inputCsv7 = Import-Csv $exportreportgrpcsv7 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv7 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-7-" + $nowfiledate + "-1.csv"
				$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile7 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-7-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Outfile7 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv8
			if ($checkfile -eq "True")
			{
				$inputCsv8 = Import-Csv $exportreportgrpcsv8 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv8 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-8-" + $nowfiledate + "-1.csv"
				$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile8 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-8-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Outfile8 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv9
			if ($checkfile -eq "True")
			{
				$inputCsv9 = Import-Csv $exportreportgrpcsv9 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv9 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-9-" + $nowfiledate + "-1.csv"
				$inputCsv9 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile9 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-9-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv9 | ? { $_.trim() -ne "" } | set-content $Outfile9 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv10
			if ($checkfile -eq "True")
			{
				$inputCsv10 = Import-Csv $exportreportgrpcsv10 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv10 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-10-" + $nowfiledate + "-1.csv"
				$inputCsv10 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile10 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-10-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv10 | ? { $_.trim() -ne "" } | set-content $Outfile10 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv11
			if ($checkfile -eq "True")
			{
				$inputCsv11 = Import-Csv $exportreportgrpcsv11 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv11 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-11-" + $nowfiledate + "-1.csv"
				$inputCsv11 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile11 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-11-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv11 | ? { $_.trim() -ne "" } | set-content $Outfile11 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv12
			if ($checkfile -eq "True")
			{
				$inputCsv12 = Import-Csv $exportreportgrpcsv12 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv12 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-12-" + $nowfiledate + "-1.csv"
				$inputCsv12 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile12 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-12-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv12 | ? { $_.trim() -ne "" } | set-content $Outfile12 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv13
			if ($checkfile -eq "True")
			{
				$inputCsv13 = Import-Csv $exportreportgrpcsv13 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv13 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-13-" + $nowfiledate + "-1.csv"
				$inputCsv13 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile13 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-13-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv13 | ? { $_.trim() -ne "" } | set-content $Outfile13 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv14
			if ($checkfile -eq "True")
			{
				$inputCsv14 = Import-Csv $exportreportgrpcsv14 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv14 = "H:\M365Reports\EARL-LookupTbl-nonmailgroups-14-" + $nowfiledate + "-1.csv"
				$inputCsv14 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile14 = "H:\M365Reports\EARL-TempLookupTable-nonmailgroups-14-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv14 | ? { $_.trim() -ne "" } | set-content $Outfile14 -Encoding UTF8
				
			}
			
			
			Map-Filewatcher
			
			if (Test-Path $Outfile1)
			{
				Try
				{
					
					Copy-item -path $Outfile1 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile1 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile1 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile2 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile2 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile3 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile3 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile4 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile4 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile5 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile6 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile7 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile8 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile9 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile10 to $filewatcherout [2nd try] | $now"
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
					
					Copy-item -path $Outfile11 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile11 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile11 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						#RemoveFilewatcher
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile12 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile12 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile13 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile13 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile14 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile14 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile15 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile15 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile16 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile16 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile16 | $now"
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


exportNonMailGroups




$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "All finished with exports for LDAP replacement Temp LookupTable hourly for nonmailgroups | $now"
DisconnectExchangeOnPrem
RemoveFilewatcher

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "Closing down script - bye $now"
Stop-Transcript

Disconnect-EXO
DisconnectExchangeOnPrem

Exit-PSSession
Exit



