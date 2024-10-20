﻿



<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.148
	 Created on:    07/08/2023 14:00
	 Created by:   	Neil Poultney
	 Organization: 	NDP Consultancy Ltd - neil.poultney@ndpconsultancyltd.co.uk
	 Filename:     	Refresh-EARL-get-lookupTable-Export.ps1
	===========================================================================
	.DESCRIPTION
		Exports for EARL MailDb to Lookup Table to temp db for refresh.

		Change Log
		V1.00, 11/10/2023 - Initial full version
		V1.10, 23/03/2024 - Update to reduce overall file sizes to imporive performance

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
$lookuptime = (get-date).addhours(-36)
Set-Variable -Name lasthour -Value $lookuptime -Option ReadOnly -Scope Script -Force

##workoutwhereweare
$Domainwearein = (Get-WmiObject Win32_ComputerSystem).Name
$whoweare = $ENV:USERNAME

if ($domainwearein -eq "zneepacp11eme2" -or $domainwearein -eq "zneepacp11emrg") { $global:Envirionmentchoice = "ProdNE" }
if ($domainwearein -eq "zweepacp11emg3" -or $domainwearein -eq "zweepacp11em50") { $global:Envirionmentchoice = "ProdWE" }



$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
$transcriptlog = "H:\EARLTranscripts\LookupTbl\refresh-lookup-export-" + $nowfiledate + ".log"

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
	$logfilelocation = $loglocation + "Refresh-Lookup-Table-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
	Set-Variable -Name logfile -Value $logfilelocation -Option ReadOnly -Scope Script -Force
	$counttype = $loglocation + "Refresh-Lookup-Table-count.txt"
	Set-Variable -Name countypeoutput -Value $counttype -Option ReadOnly -Scope Script -Force
	
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
	$filedatetext = get-date -f "yyyy-MM-dd"
	$loglocation = "H:\EARLPSLogs\BulkExports\"  # change to usetype RPMBCREATE etc
	Set-Variable -Name logfolder -Value $logslocation -Option ReadOnly -Scope Script -Force
	$logfilelocation = $loglocation + "Refresh-Lookup-Table-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
	Set-Variable -Name logfile -Value $logfilelocation -Option ReadOnly -Scope Script -Force
	$counttype = $loglocation + "Refresh-Lookup-Table-count.txt"
	Set-Variable -Name countypeoutput -Value $counttype -Option ReadOnly -Scope Script -Force
	
	
	
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
	
	#$filewatcherlocationout = "Q:\EARL\FileLocation\"
	$filewatcherlocationout = "Q:\EARL\CSVFileLocation\"
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
			#new 2023 thumprint : a98251f44faf329cd3d1474f1440aca8356edaa0
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


function Process-CsvFileremoteusermailboxes
{
	param (
		[string]$csvFilePath,
		[string]$outputFolder,
		[int]$index
	)
	
	if (Test-Path -Path $csvFilePath)
	{
		$inputCsv = Import-Csv -Path $csvFilePath -Delimiter "|" | Sort-Object * -Unique
		
		$nowFileDate = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
		$finalOutCsv = "$outputFolder\EARL-TempLookupTbl-remotembx-$index-$nowFileDate-1.csv"
		$outputFile = "$outputFolder\TempLookupUploads\EARL-TempLookupTable-remotembx-$index-$nowFileDate-2.csv"
		New-Variable -Name "outfile$index" -Value $outputFile -Force
		$inputCsv | Sort-Object -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } |
		Export-Csv -Path $finalOutCsv -NoTypeInformation -Delimiter "|" -Encoding UTF8
		
		Start-Sleep -Seconds 5
		
		Get-Content -Path $finalOutCsv | Where-Object { $_.Trim() -ne "" } | Set-Content -Path $outputFile -Encoding UTF8
	}
	
	Try
	{
		
		Copy-item -path $outputFile  -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp Lookup remotemailboxes Users File Copied to FileWatcher $outputFile  to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 20
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp Lookup remotemailboxes Users File Copied to FileWatcher $outputFile  to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			#RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $outputFile  | $now"
		}
	}
	
	if (Test-Path $finaloutcsv)
	{
		Remove-Item $finaloutcsv
	}
	
	if (Test-Path $filePath)
	{
		Remove-Item $filePath
	}
	
	
}


function Process-CsvFile
{
	param (
		[string]$filePath,
		[string]$outputBaseName
	)
	
	if (Test-Path -Path $filePath)
	{
		$nowfiledate = Get-Date -f "yyyy-MM-dd-HH-mm"
		$inputCsv = Import-Csv $filePath -Delimiter "|" | Sort-Object * -Unique
		$finaloutcsv = "H:\M365Reports\${outputBaseName}-${nowfiledate}-1.csv"
		
		$inputCsv | Sort-Object -Property @{ Expression = { $_.Samaccountname }; Ascending = $false } |
		Export-Csv $finaloutcsv -NoTypeInformation -Delimiter "|" -Encoding UTF8
		
		Start-Sleep -Seconds 5
		
		$nowfiledate = Get-Date -f "yyyy-MM-dd-HH-mm-ss"
		$Outfile = "H:\M365Reports\TempLookupUploads\${outputBaseName}-${nowfiledate}-2.csv"
		
		Get-Content $finaloutcsv | Where-Object { $_.Trim() -ne "" } | Set-Content $Outfile -Encoding UTF8
	}
	
	Try
	{
		
		Copy-item -path $Outfile -destination $filewatcherout
		$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		Add-Content $logfile "Temp Refresh File Copied to FileWatcher $Outfile to $filewatcherout [1st try] | $now"
		Start-Sleep -Seconds 20
		
		
		
		#RemoveFilewatcher
		
		
		#Map-Filewatcher
	}
	catch
	{
		
		try
		{
			RemoveFilewatcher
			Start-Sleep -s 15
			Map-Filewatcher
			Copy-item -path $Outfile -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Temp Refresh file Copied to FileWatcher $Outfile to $filewatcherout [2nd try] | $now"
			Start-Sleep -Seconds 30
			
			#RemoveFilewatcher
		}
		catch
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Cannot copy files to FileWatcher $Outfile | $now"
		}
	}
	
	if (Test-Path $finaloutcsv)
	{
		Remove-Item $finaloutcsv
	}
	
	if (Test-Path $filePath)
	{
		Remove-Item $filePath
	}
	
	
}

function Export-ReportLine
{
	param (
		[int]$MBXNumber,
		[object]$reportLine,
		[string[]]$exportReportPaths
	)
	
	$index = [math]::Floor($MBXNumber / 1000)
	if ($index -lt $exportReportPaths.Length)
	{
		$reportPath = $exportReportPaths[$index]
		$reportLine | Export-Csv $reportPath -NoTypeInformation -Delimiter "|" -Encoding UTF8 -Append -Force
	}
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
	add-content $logfile  "Running Function to export localmailboxes for export"
	
	
	$count = ""
	$M = ""
	$reconnect = 0
	
	ConnectExchangeonPrem
	
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv1 = "H:\M365Reports\EARL-Prod-LookupTable-localmbx-Hourly-" + $nowfiledate + ".csv"
	$lasthour = (get-date).adddays(-1)
	#$lasthour = (get-date).addhours(-1)
	$Mbx1 = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	#$Mbx1 = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	$count = $Mbx1.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count local mailbox accounts to process"
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv2 for remote mbx "
	add-content $counttype "LocalMailboxCount : $count"
	
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
		$descript = $null #Description
		$descript1 = $null
		$descript2 = $null
		$descript3 = $null
		$descript4 = $null
		
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
		$descript = $acc.Description
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
		
		if ($Comp)
		{
			
			[string]$Comp = $Comp -replace '|', ''
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
				UserPrincipalName		   = $UPN
				Description			       = $descript
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
	$Outfile1 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-localmbx-" + $nowfiledate + "-2.csv"
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
		Start-Sleep -Seconds 60
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
	
	
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export remotemailboxes for export"
	
	
	$count = ""
	$M = ""
	$reconnect = 0
	$MbxNumber = 0
	
	
	Start-Sleep -Seconds 1
	# Set the base file path
	$basePath = "H:\M365Reports\EARL-Prod-TempLookupTbl-remotembx-"
	$dateFormat = "yyyy-MM-dd-hh-mm-ss"
	
	# Loop through numbers 1 to 30 to create file paths
	for ($i = 1; $i -le 50; $i++)
	{
		$nowfiledate = Get-Date -Format $dateFormat
		# Construct the file path with the current index
		$filePath = "$basePath$i-$nowfiledate.csv"
		# Dynamically create a variable name for each file path
		New-Variable -Name "exportreportusercsv$i" -Value $filePath  -Force
		
		Start-Sleep -Seconds 1
	}
	
	ConnectExchangeonPrem


	$Mbx1 = Get-RemoteMailbox -ResultSize unlimited -Filter "RecipientTypedetails -eq 'RemoteUserMailbox'"  | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	$Mbx2 = Get-RemoteMailbox -ResultSize unlimited -Filter "RecipientTypedetails -eq 'RemoteSharedMailbox'" | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	$Mbx3 = Get-RemoteMailbox -ResultSize unlimited -Filter "RecipientTypedetails -eq 'RemoteRoomMailbox'" | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	$AllMbx = $Mbx1 + $Mbx2 + $Mbx3
	
	
	$usrmbxcount = $Mbx1.count
	$sharedmbxcount = $Mbx2.count
	$roommbxcount = $Mbx3.count
	
	
	$count = $AllMbx.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count remote mailbox accounts to process and upload for lookuptable refresh [Temp]"
	add-content $logfile  "Found: $usrmbxcount remote user mailbox accounts to process and upload for lookuptable refresh [Temp]"
	add-content $logfile  "Found: $sharedmbxcount remote shared mailbox accounts to process and upload for lookuptable refresh [Temp]"
	add-content $logfile  "Found: $roommbxcount remote room mailbox accounts to process and upload for lookuptable refresh [Temp]"
	add-content $logfile  "LookupTable Exporting to files $exportreportusercsv1  for remote mbx to refresh [Temp] "
	
	add-content $counttype "RemoteMailboxCount : $count"
	
	
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
		
		
		ForEach ($M in $AllMbx)
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
			$descript = $null #Description
			$descript1 = $null
			$descript2 = $null
			$descript3 = $null
			$descript4 = $null
			
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
				$descript = $acc.Description
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
					[string]$Display3 = $descript -replace "`r", ""
					
					if ($descript3 -ne $Display)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed carriage return in Description Field for $usr | $NTID | $now"
						[string]$descript = $descript3
					}
				}
				
				
				
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
				
				if ($null -ne $Display)
				{
					[string]$Display5 = $Display -replace '|', ''
					
					if ($Display5 -ne $Display)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Removed Pipe character in Display Field for $usr | $NTID | $now"
						[string]$Display = $Display5
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
				
				if ($Comp)
				{
					
					[string]$comp = $comp -replace '|', ''
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
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset3 -eq "0") -and ($PercentComplete -eq "15.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset3 = "1"
					
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
					
					
				}
				
				if (($attributecountset4 -eq "0") -and ($PercentComplete -eq "20.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset4 = "1"
					
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
					
					
				}
				
				if (($attributecountset5 -eq "0") -and ($PercentComplete -eq "25.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset5 = "1"
					
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
					
				}
				
				if (($attributecountset6 -eq "0") -and ($PercentComplete -eq "30.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset6 = "1"
					
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
					
				}
				
				if (($attributecountset7 -eq "0") -and ($PercentComplete -eq "35.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset7 = "1"
					
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
					
				}
				
				if (($attributecountset8 -eq "0") -and ($PercentComplete -eq "40.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset8 = "1"
					
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset9 -eq "0") -and ($PercentComplete -eq "45.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset9 = "1"
					
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset10 -eq "0") -and ($PercentComplete -eq "50.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset10 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				
				if (($attributecountset11 -eq "0") -and ($PercentComplete -eq "55.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset11 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset12 -eq "0") -and ($PercentComplete -eq "60.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset12 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset13 -eq "0") -and ($PercentComplete -eq "65.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset13 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset14 -eq "0") -and ($PercentComplete -eq "70.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset14 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset15 -eq "0") -and ($PercentComplete -eq "75.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset15 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset16 -eq "0") -and ($PercentComplete -eq "80.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset16 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset17 -eq "0") -and ($PercentComplete -eq "85.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset17 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset18 -eq "0") -and ($PercentComplete -eq "90.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset18 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset19 -eq "0") -and ($PercentComplete -eq "95.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset19 = "1"
					DisconnectExchangeOnPrem
					Start-Sleep -s 10
					
					ConnectExchangeonPrem
				}
				
				if (($attributecountset20 -eq "0") -and ($PercentComplete -eq "99.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed User number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset20 = "1"
				}
				
				
				
				
				
				
				
				
				#deal with sendas
				
				
				If ($Display)
				{
					
					
					$ReportLine2 = [PSCustomObject][ordered] @{
						
						
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
						$ReportLine2 | Export-CSV $exportreportusercsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 3000) -and ($MbxNumber -le 6000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 6000) -and ($MbxNumber -le 9000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 9000) -and ($MbxNumber -le 12000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 12000) -and ($MbxNumber -le 15000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 15000) -and ($MbxNumber -le 18000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 18000) -and ($MbxNumber -le 21000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 21000) -and ($MbxNumber -le 24000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 24000) -and ($MbxNumber -le 27000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 27000) -and ($MbxNumber -le 30000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 30000) -and ($MbxNumber -le 33000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 33000) -and ($MbxNumber -le 36000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 36000) -and ($MbxNumber -le 39000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 39000) -and ($MbxNumber -le 42000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 42000) -and ($MbxNumber -le 45000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv15 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 45000) -and ($MbxNumber -le 48000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv16 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 48000) -and ($MbxNumber -le 51000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv17 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 51000) -and ($MbxNumber -le 54000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv18 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 54000) -and ($MbxNumber -le 57000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv19 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 57000) -and ($MbxNumber -le 60000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv20 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 60000) -and ($MbxNumber -le 63000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv21 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 63000) -and ($MbxNumber -le 66000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv22 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 66000) -and ($MbxNumber -le 69000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv23 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 69000) -and ($MbxNumber -le 72000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv24 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 72000) -and ($MbxNumber -le 75000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv25 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 75000) -and ($MbxNumber -le 78000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv26 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 78000) -and ($MbxNumber -le 81000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv27 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 81000) -and ($MbxNumber -le 84000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv28 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 84000) -and ($MbxNumber -le 87000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv29 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 87000) -and ($MbxNumber -le 90000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv30 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 90000) -and ($MbxNumber -le 93000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv31 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					if (($MbxNumber -gt 93000) -and ($MbxNumber -le 96000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv32 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 96000) -and ($MbxNumber -le 99000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv33 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 99000) -and ($MbxNumber -le 102000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv34 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 102000) -and ($MbxNumber -le 105000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv35 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 105000) -and ($MbxNumber -le 108000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv36 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 108000) -and ($MbxNumber -le 111000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv37 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 111000) -and ($MbxNumber -le 114000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv38 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 114000) -and ($MbxNumber -le 117000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv39 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 117000) -and ($MbxNumber -le 120000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv40 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 120000) -and ($MbxNumber -le 123000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv41 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 123000) -and ($MbxNumber -le 126000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv42 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 126000) -and ($MbxNumber -le 129000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv43 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 129000) -and ($MbxNumber -le 132000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv44 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 132000) -and ($MbxNumber -le 135000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv45 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 135000) -and ($MbxNumber -le 138000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv46 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 138000) -and ($MbxNumber -le 141000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv47 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 141000) -and ($MbxNumber -le 144000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv48 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 144000) -and ($MbxNumber -le 147000))
					{
						$ReportLine2 | Export-CSV $exportreportusercsv49 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if ($MbxNumber -gt 147000)
					{
						$ReportLine2 | Export-CSV $exportreportusercsv50 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					
					
					
					
				}
				
			}
		}
		
		
		
		
		DisconnectExchangeOnPrem
		
		
		
		$exportfiles = @(
			@{ Path = $exportreportusercsv1; BaseName = "EARL-TempLookupTable-remotemailbox-1" },
			@{ Path = $exportreportusercsv2; BaseName = "EARL-TempLookupTable-remotemailbox-2" },
			@{ Path = $exportreportusercsv3; BaseName = "EARL-TempLookupTable-remotemailbox-3" },
			@{ Path = $exportreportusercsv4; BaseName = "EARL-TempLookupTable-remotemailbox-4" },
			@{ Path = $exportreportusercsv5; BaseName = "EARL-TempLookupTable-remotemailbox-5" },
			@{ Path = $exportreportusercsv6; BaseName = "EARL-TempLookupTable-remotemailbox-6" },
			@{ Path = $exportreportusercsv7; BaseName = "EARL-TempLookupTable-remotemailbox-7" },
			@{ Path = $exportreportusercsv8; BaseName = "EARL-TempLookupTable-remotemailbox-8" },
			@{ Path = $exportreportusercsv9; BaseName = "EARL-TempLookupTable-remotemailbox-9" },
			@{ Path = $exportreportusercsv10; BaseName = "EARL-TempLookupTable-remotemailbox-10" },
			@{ Path = $exportreportusercsv11; BaseName = "EARL-TempLookupTable-remotemailbox-11" },
			@{ Path = $exportreportusercsv12; BaseName = "EARL-TempLookupTable-remotemailbox-12" },
			@{ Path = $exportreportusercsv13; BaseName = "EARL-TempLookupTable-remotemailbox-13" },
			@{ Path = $exportreportusercsv14; BaseName = "EARL-TempLookupTable-remotemailbox-14" },
			@{ Path = $exportreportusercsv15; BaseName = "EARL-TempLookupTable-remotemailbox-15" },
			@{ Path = $exportreportusercsv16; BaseName = "EARL-TempLookupTable-remotemailbox-16" },
			@{ Path = $exportreportusercsv17; BaseName = "EARL-TempLookupTable-remotemailbox-17" },
			@{ Path = $exportreportusercsv18; BaseName = "EARL-TempLookupTable-remotemailbox-18" },
			@{ Path = $exportreportusercsv19; BaseName = "EARL-TempLookupTable-remotemailbox-19" },
			@{ Path = $exportreportusercsv20; BaseName = "EARL-TempLookupTable-remotemailbox-20" },
			@{ Path = $exportreportusercsv21; BaseName = "EARL-TempLookupTable-remotemailbox-21" },
			@{ Path = $exportreportusercsv22; BaseName = "EARL-TempLookupTable-remotemailbox-22" },
			@{ Path = $exportreportusercsv23; BaseName = "EARL-TempLookupTable-remotemailbox-23" },
			@{ Path = $exportreportusercsv24; BaseName = "EARL-TempLookupTable-remotemailbox-24" },
			@{ Path = $exportreportusercsv25; BaseName = "EARL-TempLookupTable-remotemailbox-25" },
			@{ Path = $exportreportusercsv26; BaseName = "EARL-TempLookupTable-remotemailbox-26" },
			@{ Path = $exportreportusercsv27; BaseName = "EARL-TempLookupTable-remotemailbox-27" },
			@{ Path = $exportreportusercsv28; BaseName = "EARL-TempLookupTable-remotemailbox-28" },
			@{ Path = $exportreportusercsv29; BaseName = "EARL-TempLookupTable-remotemailbox-29" },
			@{ Path = $exportreportusercsv30; BaseName = "EARL-TempLookupTable-remotemailbox-30" },
			@{ Path = $exportreportusercsv31; BaseName = "EARL-TempLookupTable-remotemailbox-31" },
			@{ Path = $exportreportusercsv32; BaseName = "EARL-TempLookupTable-remotemailbox-32" },
			@{ Path = $exportreportusercsv33; BaseName = "EARL-TempLookupTable-remotemailbox-33" },
			@{ Path = $exportreportusercsv34; BaseName = "EARL-TempLookupTable-remotemailbox-34" },
			@{ Path = $exportreportusercsv35; BaseName = "EARL-TempLookupTable-remotemailbox-35" },
			@{ Path = $exportreportusercsv36; BaseName = "EARL-TempLookupTable-remotemailbox-36" },
			@{ Path = $exportreportusercsv37; BaseName = "EARL-TempLookupTable-remotemailbox-37" },
			@{ Path = $exportreportusercsv38; BaseName = "EARL-TempLookupTable-remotemailbox-38" },
			@{ Path = $exportreportusercsv39; BaseName = "EARL-TempLookupTable-remotemailbox-39" },
			@{ Path = $exportreportusercsv40; BaseName = "EARL-TempLookupTable-remotemailbox-40" },
			@{ Path = $exportreportusercsv41; BaseName = "EARL-TempLookupTable-remotemailbox-41" },
			@{ Path = $exportreportusercsv42; BaseName = "EARL-TempLookupTable-remotemailbox-42" },
			@{ Path = $exportreportusercsv43; BaseName = "EARL-TempLookupTable-remotemailbox-43" },
			@{ Path = $exportreportusercsv44; BaseName = "EARL-TempLookupTable-remotemailbox-44" },
			@{ Path = $exportreportusercsv45; BaseName = "EARL-TempLookupTable-remotemailbox-45" },
			@{ Path = $exportreportusercsv46; BaseName = "EARL-TempLookupTable-remotemailbox-46" },
			@{ Path = $exportreportusercsv47; BaseName = "EARL-TempLookupTable-remotemailbox-47" },
			@{ Path = $exportreportusercsv48; BaseName = "EARL-TempLookupTable-remotemailbox-48" },
			@{ Path = $exportreportusercsv49; BaseName = "EARL-TempLookupTable-remotemailbox-49" },
			@{ Path = $exportreportusercsv50; BaseName = "EARL-TempLookupTable-remotemailbox-50" }
			
		
			
			
			
			
			
		)
		
		Map-Filewatcher
		# Process each file
		foreach ($file in $exportfiles)
		{
			Process-CsvFile -filePath $file.Path -outputBaseName $file.BaseName
			Start-Sleep -Seconds 2
		}
		
		
		RemoveFilewatcher
		
		
		
				
			
			
		
		
	}
	
}


function exportcontact
{
	$count = 0
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export contacts for export"
	
	$Mbx5 = ""
	$exportreportcsv5 = ""
	$M = ""
	$reconnect = 0
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	
	$basePath = "H:\M365Reports\Refresh-EARL-TempLookupTable-contacts-"
	$dateFormat = "yyyy-MM-dd-hh-mm-ss"
	$fileCount = 20
	$delay = 2 # Delay in seconds
	
	for ($i = 1; $i -le $fileCount; $i++)
	{
		Start-Sleep -Seconds $delay
		$nowfiledate = Get-Date -Format $dateFormat
		$filePath = "${basePath}${i}-${nowfiledate}.csv"
		# Output or process the file path
		New-Variable -Name "exportreportcontactcsv$i" -Value $filePath -Force
	}
	
	
	
	$exportreportcontactcsvall = "H:\M365Reports\Refresh-EARL-Prod-LookupTable-contacts-All-" + $nowfiledate + ".csv"
	#$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	#temp
	#$lookuptime1 = (get-date).adddays(-366)
	$lookuptime1 = (get-date).addhours(-3)
	Set-Variable -Name lasthour -Value $lookuptime1 -Option ReadOnly -Scope Script -Force
	
	$Mbx5 = Get-MailContact -ResultSize unlimited | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	#$Mbx5 = Get-MailContact -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	$count = $mbx5.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count mailcontacts accounts to process"
	
	add-content $logfile  "Refresh LookupTable Exporting to file  $exportreportcontactcsv1 | $exportreportcontactcsv2 | $exportreportcontactcsv3  for Contacts "
	
	add-content $counttype "ContactsCount : $count"
	
	
	
	
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
			$descript = $null #Description
			$descript1 = $null
			$descript2 = $null
			$descript3 = $null
			$descript4 = $null
			
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
				$descript = $acc.Description
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
				
				if ($Comp)
				{
					
					[string]$Comp = $Comp -replace '|', ''
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
				
				for ($i = 1; $i -le 20; $i++)
				{
					$attributeVar = "attributecountset$i"
					$percentComplete = "{0:D2}.000" -f ($i * 5)
					$attributeValue = Get-Variable -Name $attributeVar -ValueOnly
					
					if ($attributeValue -eq "0" -and $PercentComplete -eq $percentComplete)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Processed MailContact number : $MbxNumber | Percent Complete: $PercentComplete | $now"
						Set-Variable -Name $attributeVar -Value "1"
					}
				}
				
				
				
				
				
				If (($Display) -and ($mail) -and ($NTID))
				{
					
					
					
					$ReportLine3 = [PSCustomObject][ordered] @{
						
						
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
					
					if ($MbxNumber -le 2000)
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 2000) -and ($MbxNumber -le 4000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 4000) -and ($MbxNumber -le 6000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 6000) -and ($MbxNumber -le 8000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 8000) -and ($MbxNumber -le 10000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 10000) -and ($MbxNumber -le 12000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 12000) -and ($MbxNumber -le 14000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 14000) -and ($MbxNumber -le 16000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 16000) -and ($MbxNumber -le 18000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 18000) -and ($MbxNumber -le 20000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 20000) -and ($MbxNumber -le 22000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 22000) -and ($MbxNumber -le 24000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 24000) -and ($MbxNumber -le 26000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 26000) -and ($MbxNumber -le 28000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 28000) -and ($MbxNumber -le 30000))
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv15 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if ($MbxNumber -gt 30000)
					{
						$ReportLine3 | Export-CSV $exportreportcontactcsv16 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
					
					$ReportLine3 | Export-CSV $exportreportcontactcsvall -NoTypeInformation -Encoding UTF8 -Append -force
					
				}
			}
			
		}
		
		
		
		
		
		DisconnectExchangeOnPrem
		
		
		
		#sortoutput so no blank lines and no duplicates
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv1
		if ($checkfile -eq "True")
		{
			$inputCsv1 = Import-Csv $exportreportcontactcsv1 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv1 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-1-" + $nowfiledate + "-1.csv"
			$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout1 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-1-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv1 | ? { $_.trim() -ne "" } | set-content $Fileout1 -Encoding UTF8
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv2
		if ($checkfile -eq "True")
		{
			
			$inputCsv2 = Import-Csv $exportreportcontactcsv2 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv2 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-2-" + $nowfiledate + "-1.csv"
			$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout2 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-2-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Fileout2 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv3
		if ($checkfile -eq "True")
		{
			#sortoutput so no blank lines and no duplicates
			$inputCsv3 = Import-Csv $exportreportcontactcsv3 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv3 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-3-" + $nowfiledate + "-1.csv"
			$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout3 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-3-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Fileout3 -Encoding UTF8
			
		}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv4
		if ($checkfile -eq "True")
		{
			#sortoutput so no blank lines and no duplicates
			$inputCsv4 = Import-Csv $exportreportcontactcsv4 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv4 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-4-" + $nowfiledate + "-1.csv"
			$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout4 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-4-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Fileout4 -Encoding UTF8
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv5
		if ($checkfile -eq "True")
		{
			
			#sortoutput so no blank lines and no duplicates
			$inputCsv5 = Import-Csv $exportreportcontactcsv5 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv5 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-5-" + $nowfiledate + "-1.csv"
			$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout5 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-5-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Fileout5 -Encoding UTF8
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv6
		if ($checkfile -eq "True")
		{
			
			#sortoutput so no blank lines and no duplicates
			$inputCsv6 = Import-Csv $exportreportcontactcsv6 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv6 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-6-" + $nowfiledate + "-1.csv"
			$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout6 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-6-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Fileout6 -Encoding UTF8
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv7
		if ($checkfile -eq "True")
		{
			
			#sortoutput so no blank lines and no duplicates
			$inputCsv7 = Import-Csv $exportreportcontactcsv7 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv7 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-7-" + $nowfiledate + "-1.csv"
			$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout7 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-7-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Fileout7 -Encoding UTF8
		}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv8
		if ($checkfile -eq "True")
		{
			
			#sortoutput so no blank lines and no duplicates
			$inputCsv8 = Import-Csv $exportreportcontactcsv8 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv8 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-8-" + $nowfiledate + "-1.csv"
			$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout8 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-8-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Fileout8 -Encoding UTF8
		}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv9
		if ($checkfile -eq "True")
		{
			
			#sortoutput so no blank lines and no duplicates
			$inputCsv9 = Import-Csv $exportreportcontactcsv9 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv9 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-9-" + $nowfiledate + "-1.csv"
			$inputCsv9 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout9 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-9-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv9 | ? { $_.trim() -ne "" } | set-content $Fileout9 -Encoding UTF8
		}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv10
		if ($checkfile -eq "True")
		{
			$inputCsv10 = Import-Csv $exportreportcontactcsv10 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv10 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-10-" + $nowfiledate + "-1.csv"
			$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout10 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-10-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv10 | ? { $_.trim() -ne "" } | set-content $Fileout10 -Encoding UTF8
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv11
		if ($checkfile -eq "True")
		{
			$inputCsv11 = Import-Csv $exportreportcontactcsv11 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv11 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-11-" + $nowfiledate + "-1.csv"
			$inputCsv11 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout11 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-11-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv11 | ? { $_.trim() -ne "" } | set-content $Fileout11 -Encoding UTF8
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv12
		if ($checkfile -eq "True")
		{
			$inputCsv12 = Import-Csv $exportreportcontactcsv12 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv12 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-12-" + $nowfiledate + "-1.csv"
			$inputCsv12 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout12 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-12-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv12 | ? { $_.trim() -ne "" } | set-content $Fileout12 -Encoding UTF8
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv13
		if ($checkfile -eq "True")
		{
			$inputCsv13 = Import-Csv $exportreportcontactcsv13 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv13 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-13-" + $nowfiledate + "-1.csv"
			$inputCsv13 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout13 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-13-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv13 | ? { $_.trim() -ne "" } | set-content $Fileout13 -Encoding UTF8
		}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv14
		if ($checkfile -eq "True")
		{
			$inputCsv14 = Import-Csv $exportreportcontactcsv1 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv14 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-14-" + $nowfiledate + "-1.csv"
			$inputCsv14 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout14 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-14-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv14 | ? { $_.trim() -ne "" } | set-content $Fileout14 -Encoding UTF8
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcontactcsv15
		if ($checkfile -eq "True")
		{
			$inputCsv15 = Import-Csv $exportreportcontactcsv15 -delimiter "|" | Sort-Object * -Unique
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv15 = "H:\M365Reports\EARL-TempLookupTbl-mailcontact-15-" + $nowfiledate + "-1.csv"
			$inputCsv15 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv15 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Fileout15 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailcontact-15-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv15 | ? { $_.trim() -ne "" } | set-content $Fileout15 -Encoding UTF8
		}
		
		Map-Filewatcher
		Try
		{
			
			Copy-item -path $Fileout1 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout1 to $filewatcherout [1st try] | $now"
			
			
			
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
				Copy-item -path $Fileout1 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout1 to $filewatcherout [2nd try] | $now"
				
				RemoveFilewatcher
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout1 | $now"
			}
		}
		Start-Sleep -s 30
		if (Test-Path $Fileout2)
		{
			Try
			{
				
				Copy-item -path $Fileout2 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout2 to $filewatcherout [1st try] | $now"
				
				
				
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
					Copy-item -path $Fileout2 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout2 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout2 | $now"
				}
			}
			
		}
		
		Start-Sleep -s 30
		if (Test-Path $Fileout3)
		{
			Try
			{
				
				Copy-item -path $Fileout3 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout3 to $filewatcherout [1st try] | $now"
				
				
				
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
					Copy-item -path $Fileout3 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout3 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout3 | $now"
				}
			}
		}
		
		Start-Sleep -s 30
		if (Test-Path $Fileout4) {
		Try
		{
			
			Copy-item -path $Fileout4 -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout4 to $filewatcherout [1st try] | $now"
			
			
			
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
				Copy-item -path $Fileout4 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout4 to $filewatcherout [2nd try] | $now"
				
				
			}
			catch
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout4 | $now"
			}
		}
	}
		
		Start-Sleep -s 30
		if (Test-Path $Fileout5) {
		
			
			Try
			{
				
				Copy-item -path $Fileout5 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout5 to $filewatcherout [1st try] | $now"
				
				
				
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
					Copy-item -path $Fileout5 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout5 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout5 | $now"
				}
			}
		}
		
		Start-Sleep -s 30
		if (Test-Path $Fileout6)
		{
			Try
			{
				
				Copy-item -path $Fileout6 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout6 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout6 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout6 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout6 | $now"
				}
			}
		}
		
		
		Start-Sleep -s 30
		if (Test-Path $Fileout7)
		{
			Try
			{
				
				Copy-item -path $Fileout7 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout7 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout7 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout7 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout7 | $now"
				}
			}
		}
		
		
		Start-Sleep -s 30
		if (Test-Path $Fileout8)
		{
			Try
			{
				
				Copy-item -path $Fileout8 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout8 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout8 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout8 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout8 | $now"
				}
			}
		}
		
		
		Start-Sleep -s 30
		if (Test-Path $Fileout9)
		{
			Try
			{
				
				Copy-item -path $Fileout9 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout9 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout9 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout9 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout9 | $now"
				}
			}
		}
		
		
		Start-Sleep -s 30
		if (Test-Path $Fileout10)
		{
			Try
			{
				
				Copy-item -path $Fileout10 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout10 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout10 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout10 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout10 | $now"
				}
			}
		}
		
		
		Start-Sleep -s 30
		if (Test-Path $Fileout11)
		{
			Try
			{
				
				Copy-item -path $Fileout11 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout11 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout11 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout11 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout11 | $now"
				}
			}
		}
		
		
		Start-Sleep -s 30
		if (Test-Path $Fileout12)
		{
			Try
			{
				
				Copy-item -path $Fileout12 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout12 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout12 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout12 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout12 | $now"
				}
			}
		}
		
		
		Start-Sleep -s 30
		if (Test-Path $Fileout13)
		{
			Try
			{
				
				Copy-item -path $Fileout13 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout13 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout13 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout13 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout13 | $now"
				}
			}
		}
		
		Start-Sleep -s 30
		if (Test-Path $Fileout14)
		{
			Try
			{
				
				Copy-item -path $Fileout14 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout14 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout14 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout14 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout14 | $now"
				}
			}
		}
		
		Start-Sleep -s 30
		
		if (Test-Path $Fileout15)
		{
			Try
			{
				
				Copy-item -path $Fileout15 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout15 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout15 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout15 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout15 | $now"
				}
			}
		}
		
		Start-Sleep -s 30
				
		if (Test-Path $Fileout16)
		{
			Try
			{
				
				Copy-item -path $Fileout16 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout16 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Fileout13 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Contacts File Copied to FileWatcher $Fileout16 to $filewatcherout [2nd try] | $now"
					
					
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for Contacts to FileWatcher $Fileout16 | $now"
				}
			}
		}
		
		
		#cleanup files
		if (Test-Path $finaloutcsv2)
		{
			Remove-Item $finaloutcsv2
		}
		
		if (Test-Path $exportreportcsv2)
		{
			Remove-Item $exportreportcsv2
		}
		
		if (Test-Path $finaloutcsv1)
		{
			Remove-Item $finaloutcsv1
		}
		
		if (Test-Path $exportreportcsv1)
		{
			Remove-Item $exportreportcsv1
		}
		
		
		if (Test-Path $finaloutcsv3)
		{
			Remove-Item $finaloutcsv3
		}
		
		if (Test-Path $exportreportcsv3)
		{
			Remove-Item $exportreportcsv3
		}
		
		if (Test-Path $finaloutcsv4)
		{
			Remove-Item $finaloutcsv4
		}
		
		if (Test-Path $exportreportcsv4)
		{
			Remove-Item $exportreportcsv4
		}
		
		if (Test-Path $finaloutcsv5)
		{
			Remove-Item $finaloutcsv5
		}
		
		if (Test-Path $exportreportcsv5)
		{
			Remove-Item $exportreportcsv5
		}
		
		if (Test-Path $finaloutcsv6)
		{
			Remove-Item $finaloutcsv6
		}
		
		if (Test-Path $exportreportcsv6)
		{
			Remove-Item $exportreportcsv6
		}
		
		if (Test-Path $finaloutcsv7)
		{
			Remove-Item $finaloutcsv7
		}
		
		if (Test-Path $exportreportcsv7)
		{
			Remove-Item $exportreportcsv7
		}
		
		if (Test-Path $finaloutcsv8)
		{
			Remove-Item $finaloutcsv8
		}
		
		if (Test-Path $exportreportcsv8)
		{
			Remove-Item $exportreportcsv8
		}
		
		if (Test-Path $finaloutcsv9)
		{
			Remove-Item $finaloutcsv9
		}
		
		if (Test-Path $exportreportcsv9)
		{
			Remove-Item $exportreportcsv9
		}
		
		if (Test-Path $finaloutcsv10)
		{
			Remove-Item $finaloutcsv10
		}
		
		if (Test-Path $exportreportcsv10)
		{
			Remove-Item $exportreportcsv10
		}
	}
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Finished export contacts for Refresh LookupTable"
	
	DisconnectExchangeOnPrem
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
	$exportreportcsv3 = "H:\M365Reports\EARL-Prod-TempLookupTable-mailusers-" + $nowfiledate + ".csv"
	#$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	#$Mbx3 = Get-Mailuser -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	$count = $mbx3.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count mailusers accounts to process"
	
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv3 for Mail Users "
	add-content $counttype "MailusersCount : $count"
	
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
			$descript = $null #Description
			$descript1 = $null
			$descript2 = $null
			$descript3 = $null
			$descript4 = $null
			
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
				add-content $logfile  "Unable to get details for $usr - skipping"
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
				$descript = $acc.Description
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
				
				if ($Comp)
				{
					
					[string]$Comp = $Comp -replace '|', ''
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
						Description			       = $descript
						
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
		$finaloutcsv = "H:\M365Reports\EARL-TempLookupTbl-mailuser-" + $nowfiledate + "-1.csv"
		$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$Fileout = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-mailuser-" + $nowfiledate + "-2.csv"
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

		#>
#cleanup files
if (Test-Path $finaloutcsv)
{
	Remove-Item $finaloutcsv
}

if (Test-Path $exportreportcsv3)
{
			Remove-Item $exportreportcsv3
}

	}
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Finished export mailusers for Refresh LookupTable"
	
	DisconnectExchangeOnPrem
}


function exportDL
{
	$count = 0
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export Groups for export to Temp LookupTable"
	
	
	
	$M = ""
	$reconnect = 0
	#$exportreportcsv4 = "H:\M365Reports\EARL-Prod-LookupTable-groups-" + $nowfiledate + ".csv"
	
	$basePath = "H:\M365Reports\EARL-Prod-TempLookupTable-group-"
	$dateFormat = "yyyy-MM-dd-hh-mm-ss"
	$fileCount = 50
	$delay = 2 # Delay in seconds
	
	for ($i = 1; $i -le $fileCount; $i++)
	{
		Start-Sleep -Seconds $delay
		$nowfiledate = Get-Date -Format $dateFormat
		$filePath = "${basePath}${i}-${nowfiledate}.csv"
		# Output or process the file path
		# Dynamically create a variable name for each file path
		New-Variable -Name "exportreportgrpcsv$i" -Value $filePath -Force
		
	}
	
	
	$Mbx4 = Get-DistributionGroup -ResultSize unlimited  | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress
	
	$count = $mbx4.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count groups to process for refresh of table"
	
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv4 for groups for refresh of lookuptable [Temp] "
	
	add-content $counttype "GroupsCount : $count"
	
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
			$descript = $null #Description
			$usr = $M.PrimarySmtpAddress
			$Display = $M.DisplayName
			$recipientdetailsEX = $M.RecipientTypeDetails
			$descript1 = $null
			$descript2 = $null
			$descript3 = $null
			$descript4 = $null
			
			
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
				$descript = $acc.Description
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
				
				
				if ($null -ne  $descript)
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
				
				if ($null -ne  $descript)
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
				
				
				for ($i = 1; $i -le 20; $i++)
				{
					$attributeVar = "attributecountset$i"
					$percentComplete = "{0:D3}.000" -f ($i * 5)
					
					if ($(Get-Variable -Name $attributeVar -ValueOnly) -eq "0" -and $PercentComplete -eq $percentComplete)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now"
						Set-Variable -Name $attributeVar -Value "1"
						DisconnectExchangeOnPrem
						Start-Sleep -Seconds 10
						ConnectExchangeOnPrem
					}
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
					
					if (($MbxNumber -gt 39000) -and ($MbxNumber -le 42000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 42000) -and ($MbxNumber -le 45000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv15 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 45000) -and ($MbxNumber -le 48000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv16 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 48000) -and ($MbxNumber -le 51000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv17 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 51000) -and ($MbxNumber -le 54000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv18 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 54000) -and ($MbxNumber -le 57000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv19 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 57000) -and ($MbxNumber -le 60000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv20 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 60000) -and ($MbxNumber -le 63000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv21 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 63000) -and ($MbxNumber -le 66000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv22 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 66000) -and ($MbxNumber -le 69000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv23 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 69000) -and ($MbxNumber -le 72000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv24 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 72000) -and ($MbxNumber -le 75000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv25 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 75000) -and ($MbxNumber -le 78000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv26 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 78000) -and ($MbxNumber -le 81000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv27 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 81000) -and ($MbxNumber -le 84000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv28 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 84000) -and ($MbxNumber -le 87000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv29 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 87000) -and ($MbxNumber -le 90000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv30 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 90000) -and ($MbxNumber -le 93000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv31 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 93000) -and ($MbxNumber -le 96000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv32 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 96000) -and ($MbxNumber -le 99000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv33 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 99000) -and ($MbxNumber -le 102000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv34 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 102000) -and ($MbxNumber -le 105000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv35 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 105000) -and ($MbxNumber -le 108000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv36 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 108000) -and ($MbxNumber -le 111000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv37 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 111000) -and ($MbxNumber -le 114000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv38 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 114000) -and ($MbxNumber -le 117000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv39 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 117000) -and ($MbxNumber -le 120000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv40 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 120000) -and ($MbxNumber -le 123000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv41 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 123000) -and ($MbxNumber -le 126000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv42 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 126000) -and ($MbxNumber -le 129000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv43 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 129000) -and ($MbxNumber -le 132000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv44 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 132000) -and ($MbxNumber -le 135000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv45 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 135000) -and ($MbxNumber -le 138000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv46 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 138000) -and ($MbxNumber -le 141000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv47 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 141000) -and ($MbxNumber -le 144000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv48 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 144000) -and ($MbxNumber -le 147000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv49 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if ($MbxNumber -gt 147000)
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv50 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
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
			$finaloutcsv = "H:\M365Reports\EARL-TempLookupTbl-groups-1" + $nowfiledate + "-1.csv"
			$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile1 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-1-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Outfile1 -Encoding UTF8
			
			Start-Sleep -s 2
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
			$inputCsv2 = Import-Csv $exportreportgrpcsv2 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv2 = "H:\M365Reports\EARL-TempLookupTbl-groups-2-" + $nowfiledate + "-1.csv"
			$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile2 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-2-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Outfile2 -Encoding UTF8
			
			Start-Sleep -s 2
			$inputCsv3 = Import-Csv $exportreportgrpcsv3 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv3 = "H:\M365Reports\EARL-TempLookupTbl-groups-3-" + $nowfiledate + "-1.csv"
			$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile3 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-3-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Outfile3 -Encoding UTF8
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv4
			if ($checkfile -eq "True")
			{
				$inputCsv4 = Import-Csv $exportreportgrpcsv4 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv4 = "H:\M365Reports\EARL-TempLookupTbl-groups-4-" + $nowfiledate + "-1.csv"
				$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile4 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-4-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Outfile4 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv5
			if ($checkfile -eq "True")
			{
				$inputCsv5 = Import-Csv $exportreportgrpcsv5 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv5 = "H:\M365Reports\EARL-TempLookupTbl-groups-5-" + $nowfiledate + "-1.csv"
				$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile5 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-5-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Outfile5 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv6
			if ($checkfile -eq "True")
			{
				$inputCsv6 = Import-Csv $exportreportgrpcsv6 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv6 = "H:\M365Reports\EARL-TempLookupTbl-groups-6-" + $nowfiledate + "-1.csv"
				$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile6 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-6-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Outfile6 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv7
			if ($checkfile -eq "True")
			{
				$inputCsv7 = Import-Csv $exportreportgrpcsv7 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv7 = "H:\M365Reports\EARL-TempLookupTbl-groups-7-" + $nowfiledate + "-1.csv"
				$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile7 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-7-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Outfile7 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv8
			if ($checkfile -eq "True")
			{
				$inputCsv8 = Import-Csv $exportreportgrpcsv8 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv8 = "H:\M365Reports\EARL-TempLookupTbl-groups-8-" + $nowfiledate + "-1.csv"
				$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile8 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-8-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Outfile8 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv9
			if ($checkfile -eq "True")
			{
				$inputCsv9 = Import-Csv $exportreportgrpcsv9 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv9 = "H:\M365Reports\EARL-TempLookupTbl-groups-9-" + $nowfiledate + "-1.csv"
				$inputCsv9 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile9 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-9-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv9 | ? { $_.trim() -ne "" } | set-content $Outfile9 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv10
			if ($checkfile -eq "True")
			{
				$inputCsv10 = Import-Csv $exportreportgrpcsv10 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv10 = "H:\M365Reports\EARL-TempLookupTbl-groups-10-" + $nowfiledate + "-1.csv"
				$inputCsv10 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile10 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-10-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv10 | ? { $_.trim() -ne "" } | set-content $Outfile10 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv11
			if ($checkfile -eq "True")
			{
				$inputCsv11 = Import-Csv $exportreportgrpcsv11 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv11 = "H:\M365Reports\EARL-TempLookupTbl-groups-11-" + $nowfiledate + "-1.csv"
				$inputCsv11 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile11 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-11-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv11 | ? { $_.trim() -ne "" } | set-content $Outfile11 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv12
			if ($checkfile -eq "True")
			{
				$inputCsv12 = Import-Csv $exportreportgrpcsv12 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv12 = "H:\M365Reports\EARL-TempLookupTbl-groups-12-" + $nowfiledate + "-1.csv"
				$inputCsv12 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile12 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-12-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv12 | ? { $_.trim() -ne "" } | set-content $Outfile12 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv13
			if ($checkfile -eq "True")
			{
				$inputCsv13 = Import-Csv $exportreportgrpcsv13 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv13 = "H:\M365Reports\EARL-TempLookupTbl-groups-13-" + $nowfiledate + "-1.csv"
				$inputCsv13 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile13 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-13-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv13 | ? { $_.trim() -ne "" } | set-content $Outfile13 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv14
			if ($checkfile -eq "True")
			{
				$inputCsv14 = Import-Csv $exportreportgrpcsv14 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv14 = "H:\M365Reports\EARL-TempLookupTbl-groups-14-" + $nowfiledate + "-1.csv"
				$inputCsv14 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile14 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-14-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv14 | ? { $_.trim() -ne "" } | set-content $Outfile14 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv15
			if ($checkfile -eq "True")
			{
				$inputCsv15 = Import-Csv $exportreportgrpcsv15 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv15 = "H:\M365Reports\EARL-TempLookupTbl-groups-15-" + $nowfiledate + "-1.csv"
				$inputCsv15 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv15 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile15 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-15-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv15 | ? { $_.trim() -ne "" } | set-content $Outfile15 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv16
			if ($checkfile -eq "True")
			{
				$inputCsv16 = Import-Csv $exportreportgrpcsv16 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv16 = "H:\M365Reports\EARL-TempLookupTbl-groups-16-" + $nowfiledate + "-1.csv"
				$inputCsv16 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv16 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile16 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-16-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv16 | ? { $_.trim() -ne "" } | set-content $Outfile16 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv17
			if ($checkfile -eq "True")
			{
				$inputCsv17 = Import-Csv $exportreportgrpcsv17 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv17 = "H:\M365Reports\EARL-TempLookupTbl-groups-17-" + $nowfiledate + "-1.csv"
				$inputCsv17 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv17 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile17 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-17-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv17 | ? { $_.trim() -ne "" } | set-content $Outfile17 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv18
			if ($checkfile -eq "True")
			{
				$inputCsv18 = Import-Csv $exportreportgrpcsv18 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv18 = "H:\M365Reports\EARL-TempLookupTbl-groups-18-" + $nowfiledate + "-1.csv"
				$inputCsv18 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv18 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile18 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-18-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv18 | ? { $_.trim() -ne "" } | set-content $Outfile18 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv19
			if ($checkfile -eq "True")
			{
				$inputCsv19 = Import-Csv $exportreportgrpcsv19 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv19 = "H:\M365Reports\EARL-TempLookupTbl-groups-19-" + $nowfiledate + "-1.csv"
				$inputCsv19 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv19 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile19 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-19-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv19 | ? { $_.trim() -ne "" } | set-content $Outfile19 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv20
			if ($checkfile -eq "True")
			{
				$inputCsv20 = Import-Csv $exportreportgrpcsv20 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv20 = "H:\M365Reports\EARL-TempLookupTbl-groups-20-" + $nowfiledate + "-1.csv"
				$inputCsv20 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv20 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile20 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-20-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv20 | ? { $_.trim() -ne "" } | set-content $Outfile20 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv21
			if ($checkfile -eq "True")
			{
				$inputCsv21 = Import-Csv $exportreportgrpcsv21 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv21 = "H:\M365Reports\EARL-TempLookupTbl-groups-21-" + $nowfiledate + "-1.csv"
				$inputCsv21 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv21 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile21 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-21-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv21 | ? { $_.trim() -ne "" } | set-content $Outfile21 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv22
			if ($checkfile -eq "True")
			{
				$inputCsv22 = Import-Csv $exportreportgrpcsv22 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv22 = "H:\M365Reports\EARL-TempLookupTbl-groups-22-" + $nowfiledate + "-1.csv"
				$inputCsv22 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv22 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile22 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-22-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv22 | ? { $_.trim() -ne "" } | set-content $Outfile22 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv23
			if ($checkfile -eq "True")
			{
				$inputCsv23 = Import-Csv $exportreportgrpcsv23 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv23 = "H:\M365Reports\EARL-TempLookupTbl-groups-23-" + $nowfiledate + "-1.csv"
				$inputCsv23 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv23 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile23 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-23-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv23 | ? { $_.trim() -ne "" } | set-content $Outfile23 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv24
			if ($checkfile -eq "True")
			{
				$inputCsv24 = Import-Csv $exportreportgrpcsv24 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv24 = "H:\M365Reports\EARL-TempLookupTbl-groups-24-" + $nowfiledate + "-1.csv"
				$inputCsv24 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv24 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile24 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-24-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv24 | ? { $_.trim() -ne "" } | set-content $Outfile24 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv25
			if ($checkfile -eq "True")
			{
				$inputCsv25 = Import-Csv $exportreportgrpcsv25 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv25 = "H:\M365Reports\EARL-TempLookupTbl-groups-25-" + $nowfiledate + "-1.csv"
				$inputCsv25 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv25 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile25 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-25-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv25 | ? { $_.trim() -ne "" } | set-content $Outfile25 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv26
			if ($checkfile -eq "True")
			{
				$inputCsv26 = Import-Csv $exportreportgrpcsv26 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv26 = "H:\M365Reports\EARL-TempLookupTbl-groups-26-" + $nowfiledate + "-1.csv"
				$inputCsv26 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv26 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile26 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-26-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv26 | ? { $_.trim() -ne "" } | set-content $Outfile26 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv27
			if ($checkfile -eq "True")
			{
				$inputCsv27 = Import-Csv $exportreportgrpcsv27 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv27 = "H:\M365Reports\EARL-TempLookupTbl-groups-27-" + $nowfiledate + "-1.csv"
				$inputCsv27 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv27 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile27 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-27-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv27 | ? { $_.trim() -ne "" } | set-content $Outfile27 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv28
			if ($checkfile -eq "True")
			{
				$inputCsv28 = Import-Csv $exportreportgrpcsv28 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv28 = "H:\M365Reports\EARL-TempLookupTbl-groups-28-" + $nowfiledate + "-1.csv"
				$inputCsv28 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv28 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile28 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-28-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv28 | ? { $_.trim() -ne "" } | set-content $Outfile28 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv29
			if ($checkfile -eq "True")
			{
				$inputCsv29 = Import-Csv $exportreportgrpcsv29 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv29 = "H:\M365Reports\EARL-TempLookupTbl-groups-29-" + $nowfiledate + "-1.csv"
				$inputCsv29 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv29 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile29 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-29-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv29 | ? { $_.trim() -ne "" } | set-content $Outfile29 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv30
			if ($checkfile -eq "True")
			{
				$inputCsv30 = Import-Csv $exportreportgrpcsv30 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv30 = "H:\M365Reports\EARL-TempLookupTbl-groups-30-" + $nowfiledate + "-1.csv"
				$inputCsv30 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv30 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile30 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-30-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv30 | ? { $_.trim() -ne "" } | set-content $Outfile30 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv31
			if ($checkfile -eq "True")
			{
				$inputCsv31 = Import-Csv $exportreportgrpcsv31 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv31 = "H:\M365Reports\EARL-TempLookupTbl-groups-31-" + $nowfiledate + "-1.csv"
				$inputCsv31 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv31 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile31 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-31-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv31 | ? { $_.trim() -ne "" } | set-content $Outfile31 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv32
			if ($checkfile -eq "True")
			{
				$inputCsv32 = Import-Csv $exportreportgrpcsv32 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv32 = "H:\M365Reports\EARL-TempLookupTbl-groups-32-" + $nowfiledate + "-1.csv"
				$inputCsv32 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv32 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile32 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-32-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv32 | ? { $_.trim() -ne "" } | set-content $Outfile32 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv33
			if ($checkfile -eq "True")
			{
				$inputCsv33 = Import-Csv $exportreportgrpcsv33 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv33 = "H:\M365Reports\EARL-TempLookupTbl-groups-33-" + $nowfiledate + "-1.csv"
				$inputCsv33 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv33 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile33 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-33-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv33 | ? { $_.trim() -ne "" } | set-content $Outfile33 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv34
			if ($checkfile -eq "True")
			{
				$inputCsv34 = Import-Csv $exportreportgrpcsv34 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv34 = "H:\M365Reports\EARL-TempLookupTbl-groups-34-" + $nowfiledate + "-1.csv"
				$inputCsv34 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv34 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile34 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-34-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv34 | ? { $_.trim() -ne "" } | set-content $Outfile34 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv35
			if ($checkfile -eq "True")
			{
				$inputCsv35 = Import-Csv $exportreportgrpcsv35 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv35 = "H:\M365Reports\EARL-TempLookupTbl-groups-35-" + $nowfiledate + "-1.csv"
				$inputCsv35 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv35 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile35 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-35-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv35 | ? { $_.trim() -ne "" } | set-content $Outfile35 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv36
			if ($checkfile -eq "True")
			{
				$inputCsv36 = Import-Csv $exportreportgrpcsv36 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv36 = "H:\M365Reports\EARL-TempLookupTbl-groups-36-" + $nowfiledate + "-1.csv"
				$inputCsv36 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv36 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile36 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-36-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv36 | ? { $_.trim() -ne "" } | set-content $Outfile36 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv37
			if ($checkfile -eq "True")
			{
				$inputCsv37 = Import-Csv $exportreportgrpcsv37 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv37 = "H:\M365Reports\EARL-TempLookupTbl-groups-37-" + $nowfiledate + "-1.csv"
				$inputCsv37 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv37 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile37 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-37-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv37 | ? { $_.trim() -ne "" } | set-content $Outfile37 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv38
			if ($checkfile -eq "True")
			{
				$inputCsv38 = Import-Csv $exportreportgrpcsv38 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv38 = "H:\M365Reports\EARL-TempLookupTbl-groups-38-" + $nowfiledate + "-1.csv"
				$inputCsv38 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv38 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile38 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-38-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv38 | ? { $_.trim() -ne "" } | set-content $Outfile38 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv39
			if ($checkfile -eq "True")
			{
				$inputCsv39 = Import-Csv $exportreportgrpcsv39 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv39 = "H:\M365Reports\EARL-TempLookupTbl-groups-39-" + $nowfiledate + "-1.csv"
				$inputCsv39 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv39 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile39 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-39-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv39 | ? { $_.trim() -ne "" } | set-content $Outfile39 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv40
			if ($checkfile -eq "True")
			{
				$inputCsv40 = Import-Csv $exportreportgrpcsv40 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv40 = "H:\M365Reports\EARL-TempLookupTbl-groups-40-" + $nowfiledate + "-1.csv"
				$inputCsv40 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv40 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile40 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-40-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv40 | ? { $_.trim() -ne "" } | set-content $Outfile40 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv41
			if ($checkfile -eq "True")
			{
				$inputCsv41 = Import-Csv $exportreportgrpcsv41 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv41 = "H:\M365Reports\EARL-TempLookupTbl-groups-41-" + $nowfiledate + "-1.csv"
				$inputCsv41 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv41 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile41 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-41-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv41 | ? { $_.trim() -ne "" } | set-content $Outfile41 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv42
			if ($checkfile -eq "True")
			{
				$inputCsv42 = Import-Csv $exportreportgrpcsv42 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv42 = "H:\M365Reports\EARL-TempLookupTbl-groups-42-" + $nowfiledate + "-1.csv"
				$inputCsv42 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv42 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile42 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-42-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv42 | ? { $_.trim() -ne "" } | set-content $Outfile42 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv43
			if ($checkfile -eq "True")
			{
				$inputCsv43 = Import-Csv $exportreportgrpcsv43 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv43 = "H:\M365Reports\EARL-TempLookupTbl-groups-43-" + $nowfiledate + "-1.csv"
				$inputCsv43 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv43 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile43 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-43-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv43 | ? { $_.trim() -ne "" } | set-content $Outfile43 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv44
			if ($checkfile -eq "True")
			{
				$inputCsv44 = Import-Csv $exportreportgrpcsv44 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv44 = "H:\M365Reports\EARL-TempLookupTbl-groups-44-" + $nowfiledate + "-1.csv"
				$inputCsv44 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv44 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile44 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-44-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv44 | ? { $_.trim() -ne "" } | set-content $Outfile44 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv45
			if ($checkfile -eq "True")
			{
				$inputCsv45 = Import-Csv $exportreportgrpcsv45 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv45 = "H:\M365Reports\EARL-TempLookupTbl-groups-45-" + $nowfiledate + "-1.csv"
				$inputCsv45 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv45 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile45 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-45-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv45 | ? { $_.trim() -ne "" } | set-content $Outfile45 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv46
			if ($checkfile -eq "True")
			{
				$inputCsv46 = Import-Csv $exportreportgrpcsv46 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv46 = "H:\M365Reports\EARL-TempLookupTbl-groups-46-" + $nowfiledate + "-1.csv"
				$inputCsv46 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv46 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile46 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-46-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv46 | ? { $_.trim() -ne "" } | set-content $Outfile46 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv47
			if ($checkfile -eq "True")
			{
				$inputCsv47 = Import-Csv $exportreportgrpcsv47 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv47 = "H:\M365Reports\EARL-TempLookupTbl-groups-47-" + $nowfiledate + "-1.csv"
				$inputCsv47 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv47 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile47 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-47-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv47 | ? { $_.trim() -ne "" } | set-content $Outfile47 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv48
			if ($checkfile -eq "True")
			{
				$inputCsv48 = Import-Csv $exportreportgrpcsv48 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv48 = "H:\M365Reports\EARL-TempLookupTbl-groups-48-" + $nowfiledate + "-1.csv"
				$inputCsv48 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv48 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile48 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-48-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv48 | ? { $_.trim() -ne "" } | set-content $Outfile48 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv49
			if ($checkfile -eq "True")
			{
				$inputCsv49 = Import-Csv $exportreportgrpcsv49 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv49 = "H:\M365Reports\EARL-TempLookupTbl-groups-49-" + $nowfiledate + "-1.csv"
				$inputCsv49 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv49 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile49 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-49-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv49 | ? { $_.trim() -ne "" } | set-content $Outfile49 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv50
			if ($checkfile -eq "True")
			{
				$inputCsv50 = Import-Csv $exportreportgrpcsv50 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv50 = "H:\M365Reports\EARL-TempLookupTbl-groups-50-" + $nowfiledate + "-1.csv"
				$inputCsv50 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv50 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile50 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-groups-50-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv50 | ? { $_.trim() -ne "" } | set-content $Outfile50 -Encoding UTF8
				
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
			
			
			if (Test-Path $Outfile17)
			{
				
				
				Try
				{
					#Map-Filewatcher
					Copy-item -path $Outfile17 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile17 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile17 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile18 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile18 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile19 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile19 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile19 | $now"
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
					Add-Content $logfile "Group File Copied to FileWatcher $Outfile20 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile20 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					
					Copy-item -path $Outfile21 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile21 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile21 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						#RemoveFilewatcher
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile21 | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile22 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile22 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile23 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile23 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile24 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile24 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile25 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile25 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile26 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile26 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile27 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile27 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile28 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile28 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile29 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile29 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile29 | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile30 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile30 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					
					Copy-item -path $Outfile31 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile31 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile31 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						#RemoveFilewatcher
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile32 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile32 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile33 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile33 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile34 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile34 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile35 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile35 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile36 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile36 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile37 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile37 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile38 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile38 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile39 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile39 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile39 | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile40 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile40 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					
					Copy-item -path $Outfile41 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile41 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile41 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						#RemoveFilewatcher
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile42 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile42 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile43 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile43 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile44 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile44 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile45 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile45 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile46 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile46 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile47 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile47 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile48 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile48 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile49 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile49 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile49 | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile50 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile50 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile50 | $now"
					}
				}
				
			}
			
			
			#cleanup files
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Clearing up files no longer needed for groups | $now"
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
			
			if (Test-Path $exportreportgrpcsv15)
			{
				Remove-Item $exportreportgrpcsv15
			}
			
			if (Test-Path $finaloutcsv15)
			{
				Remove-Item $finaloutcsv15
			}
			
			if (Test-Path $exportreportgrpcsv16)
			{
				Remove-Item $exportreportgrpcsv16
			}
			
			if (Test-Path $finaloutcsv16)
			{
				Remove-Item $finaloutcsv16
			}
			
			if (Test-Path $exportreportgrpcsv17)
			{
				Remove-Item $exportreportgrpcsv17
			}
			
			if (Test-Path $finaloutcsv17)
			{
				Remove-Item $finaloutcsv17
			}
			
			if (Test-Path $exportreportgrpcsv18)
			{
				Remove-Item $exportreportgrpcsv18
			}
			
			if (Test-Path $finaloutcsv18)
			{
				Remove-Item $finaloutcsv18
			}
			
			if (Test-Path $exportreportgrpcsv19)
			{
				Remove-Item $exportreportgrpcsv19
			}
			
			if (Test-Path $finaloutcsv19)
			{
				Remove-Item $finaloutcsv19
			}
			
			if (Test-Path $exportreportgrpcsv20)
			{
				Remove-Item $exportreportgrpcsv20
			}
			
			if (Test-Path $finaloutcsv20)
			{
				Remove-Item $finaloutcsv20
			}
			
			if (Test-Path $exportreportgrpcsv21)
			{
				Remove-Item $exportreportgrpcsv21
			}
			
			if (Test-Path $finaloutcsv21)
			{
				Remove-Item $finaloutcsv21
			}
			
			if (Test-Path $exportreportgrpcsv22)
			{
				Remove-Item $exportreportgrpcsv22
			}
			
			if (Test-Path $finaloutcsv22)
			{
				Remove-Item $finaloutcsv22
			}
			
			if (Test-Path $exportreportgrpcsv23)
			{
				Remove-Item $exportreportgrpcsv23
			}
			
			if (Test-Path $finaloutcsv23)
			{
				Remove-Item $finaloutcsv23
			}
			
			if (Test-Path $exportreportgrpcsv24)
			{
				Remove-Item $exportreportgrpcsv24
			}
			
			if (Test-Path $finaloutcsv24)
			{
				Remove-Item $finaloutcsv24
			}
			
			if (Test-Path $exportreportgrpcsv25)
			{
				Remove-Item $exportreportgrpcsv25
			}
			
			if (Test-Path $finaloutcsv25)
			{
				Remove-Item $finaloutcsv25
			}
			
			if (Test-Path $exportreportgrpcsv26)
			{
				Remove-Item $exportreportgrpcsv26
			}
			
			if (Test-Path $finaloutcsv26)
			{
				Remove-Item $finaloutcsv26
			}
			
			if (Test-Path $exportreportgrpcsv27)
			{
				Remove-Item $exportreportgrpcsv27
			}
			
			if (Test-Path $finaloutcsv27)
			{
				Remove-Item $finaloutcsv27
			}
			
			if (Test-Path $exportreportgrpcsv28)
			{
				Remove-Item $exportreportgrpcsv28
			}
			
			if (Test-Path $finaloutcsv28)
			{
				Remove-Item $finaloutcsv28
			}
			
			if (Test-Path $exportreportgrpcsv29)
			{
				Remove-Item $exportreportgrpcsv29
			}
			
			if (Test-Path $finaloutcsv29)
			{
				Remove-Item $finaloutcsv29
			}
			
			if (Test-Path $exportreportgrpcsv30)
			{
				Remove-Item $exportreportgrpcsv30
			}
			
			if (Test-Path $finaloutcsv30)
			{
				Remove-Item $finaloutcsv30
			}
			
			if (Test-Path $exportreportgrpcsv31)
			{
				Remove-Item $exportreportgrpcsv31
			}
			
			if (Test-Path $finaloutcsv31)
			{
				Remove-Item $finaloutcsv31
			}
			
			if (Test-Path $exportreportgrpcsv32)
			{
				Remove-Item $exportreportgrpcsv32
			}
			
			if (Test-Path $finaloutcsv32)
			{
				Remove-Item $finaloutcsv32
			}
			
			if (Test-Path $exportreportgrpcsv33)
			{
				Remove-Item $exportreportgrpcsv33
			}
			
			if (Test-Path $finaloutcsv33)
			{
				Remove-Item $finaloutcsv33
			}
			
			if (Test-Path $exportreportgrpcsv34)
			{
				Remove-Item $exportreportgrpcsv34
			}
			
			if (Test-Path $finaloutcsv34)
			{
				Remove-Item $finaloutcsv34
			}
			
			if (Test-Path $exportreportgrpcsv35)
			{
				Remove-Item $exportreportgrpcsv35
			}
			
			if (Test-Path $finaloutcsv35)
			{
				Remove-Item $finaloutcsv35
			}
			
			if (Test-Path $exportreportgrpcsv36)
			{
				Remove-Item $exportreportgrpcsv36
			}
			
			if (Test-Path $finaloutcsv36)
			{
				Remove-Item $finaloutcsv36
			}
			
			if (Test-Path $exportreportgrpcsv37)
			{
				Remove-Item $exportreportgrpcsv37
			}
			
			if (Test-Path $finaloutcsv37)
			{
				Remove-Item $finaloutcsv37
			}
			
			if (Test-Path $exportreportgrpcsv38)
			{
				Remove-Item $exportreportgrpcsv38
			}
			
			if (Test-Path $finaloutcsv38)
			{
				Remove-Item $finaloutcsv38
			}
			
			if (Test-Path $exportreportgrpcsv39)
			{
				Remove-Item $exportreportgrpcsv39
			}
			
			if (Test-Path $finaloutcsv39)
			{
				Remove-Item $finaloutcsv39
			}
			
			if (Test-Path $exportreportgrpcsv40)
			{
				Remove-Item $exportreportgrpcsv40
			}
			
			if (Test-Path $finaloutcsv40)
			{
				Remove-Item $finaloutcsv40
			}
			
			
			
			
		}
		RemoveFilewatcher
	}
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	Add-Content $logfile "Finished with exports for Temp LookupTable for Groups | $now"
	
}


function exportM365Grps
{
	$count = 0
	
	Connect-EXO
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export M365Groups for export to Temp LookupTable"
	
	
	
	$M = ""
	$reconnect = 0
	#$exportreportcsv4 = "H:\M365Reports\EARL-Prod-LookupTable-groups-" + $nowfiledate + ".csv"
	
	$basePath = "H:\M365Reports\EARL-Prod-TempLookupTable-M365group-"
	$dateFormat = "yyyy-MM-dd-hh-mm-ss"
	$fileCount = 50
	$delay = 2 # Delay in seconds
	
	for ($i = 1; $i -le $fileCount; $i++)
	{
		Start-Sleep -Seconds $delay
		$nowfiledate = Get-Date -Format $dateFormat
		$filePath = "${basePath}${i}-${nowfiledate}.csv"
		# Output or process the file path
		# Dynamically create a variable name for each file path
		New-Variable -Name "exportreportgrpcsv$i" -Value $filePath -Force
		
	}
	
	
	$Mbx4 = Get-UnifiedGroup -ResultSize unlimited | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress
	
	$count = $mbx4.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count M365groups to process for refresh of table"
	
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv1 .. for M365groups for refresh of lookuptable [Temp] "
	
	add-content $counttype "GroupsCount : $count"
	
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
			$descript = $null #Description
			$usr = $M.PrimarySmtpAddress
			$Display = $M.DisplayName
			$recipientdetailsEX = $M.RecipientTypeDetails
			$descript1 = $null
			$descript2 = $null
			$descript3 = $null
			$descript4 = $null
			
			
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
				$descript = $acc.Description
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
				
				
				for ($i = 1; $i -le 20; $i++)
				{
					$attributeVar = "attributecountset$i"
					$percentComplete = "{0:D3}.000" -f ($i * 5)
					
					if ($(Get-Variable -Name $attributeVar -ValueOnly) -eq "0" -and $PercentComplete -eq $percentComplete)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Processed Group number : $MbxNumber | Percent Complete: $PercentComplete | $now"
						Set-Variable -Name $attributeVar -Value "1"
						DisconnectExchangeOnPrem
						Start-Sleep -Seconds 10
						ConnectExchangeOnPrem
					}
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
					
					if (($MbxNumber -gt 39000) -and ($MbxNumber -le 42000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 42000) -and ($MbxNumber -le 45000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv15 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 45000) -and ($MbxNumber -le 48000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv16 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 48000) -and ($MbxNumber -le 51000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv17 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 51000) -and ($MbxNumber -le 54000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv18 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 54000) -and ($MbxNumber -le 57000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv19 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 57000) -and ($MbxNumber -le 60000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv20 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 60000) -and ($MbxNumber -le 63000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv21 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 63000) -and ($MbxNumber -le 66000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv22 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 66000) -and ($MbxNumber -le 69000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv23 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 69000) -and ($MbxNumber -le 72000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv24 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 72000) -and ($MbxNumber -le 75000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv25 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 75000) -and ($MbxNumber -le 78000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv26 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 78000) -and ($MbxNumber -le 81000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv27 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 81000) -and ($MbxNumber -le 84000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv28 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 84000) -and ($MbxNumber -le 87000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv29 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 87000) -and ($MbxNumber -le 90000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv30 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 90000) -and ($MbxNumber -le 93000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv31 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 93000) -and ($MbxNumber -le 96000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv32 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 96000) -and ($MbxNumber -le 99000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv33 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 99000) -and ($MbxNumber -le 102000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv34 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 102000) -and ($MbxNumber -le 105000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv35 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 105000) -and ($MbxNumber -le 108000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv36 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 108000) -and ($MbxNumber -le 111000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv37 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 111000) -and ($MbxNumber -le 114000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv38 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 114000) -and ($MbxNumber -le 117000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv39 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 117000) -and ($MbxNumber -le 120000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv40 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 120000) -and ($MbxNumber -le 123000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv41 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 123000) -and ($MbxNumber -le 126000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv42 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 126000) -and ($MbxNumber -le 129000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv43 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 129000) -and ($MbxNumber -le 132000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv44 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 132000) -and ($MbxNumber -le 135000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv45 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 135000) -and ($MbxNumber -le 138000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv46 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 138000) -and ($MbxNumber -le 141000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv47 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 141000) -and ($MbxNumber -le 144000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv48 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -gt 144000) -and ($MbxNumber -le 147000))
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv49 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if ($MbxNumber -gt 147000)
					{
						$ReportLine4 | Export-CSV $exportreportgrpcsv50 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
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
			$finaloutcsv = "H:\M365Reports\EARL-TempLookupTbl-groups-1" + $nowfiledate + "-1.csv"
			$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile1 = "H:\M365Reports\EARL-TempLookupTable-groups-1-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Outfile1 -Encoding UTF8
			
			Start-Sleep -s 2
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
			$inputCsv2 = Import-Csv $exportreportgrpcsv2 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv2 = "H:\M365Reports\EARL-TempLookupTbl-groups-2-" + $nowfiledate + "-1.csv"
			$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile2 = "H:\M365Reports\EARL-TempLookupTable-groups-2-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Outfile2 -Encoding UTF8
			
			Start-Sleep -s 2
			$inputCsv3 = Import-Csv $exportreportgrpcsv3 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv3 = "H:\M365Reports\EARL-TempLookupTbl-groups-3-" + $nowfiledate + "-1.csv"
			$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile3 = "H:\M365Reports\EARL-TempLookupTable-groups-3-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Outfile3 -Encoding UTF8
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv4
			if ($checkfile -eq "True")
			{
				$inputCsv4 = Import-Csv $exportreportgrpcsv4 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv4 = "H:\M365Reports\EARL-TempLookupTbl-groups-4-" + $nowfiledate + "-1.csv"
				$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile4 = "H:\M365Reports\EARL-TempLookupTable-groups-4-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Outfile4 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv5
			if ($checkfile -eq "True")
			{
				$inputCsv5 = Import-Csv $exportreportgrpcsv5 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv5 = "H:\M365Reports\EARL-TempLookupTbl-groups-5-" + $nowfiledate + "-1.csv"
				$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile5 = "H:\M365Reports\EARL-TempLookupTable-groups-5-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Outfile5 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv6
			if ($checkfile -eq "True")
			{
				$inputCsv6 = Import-Csv $exportreportgrpcsv6 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv6 = "H:\M365Reports\EARL-TempLookupTbl-groups-6-" + $nowfiledate + "-1.csv"
				$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile6 = "H:\M365Reports\EARL-TempLookupTable-groups-6-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Outfile6 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv7
			if ($checkfile -eq "True")
			{
				$inputCsv7 = Import-Csv $exportreportgrpcsv7 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv7 = "H:\M365Reports\EARL-TempLookupTbl-groups-7-" + $nowfiledate + "-1.csv"
				$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile7 = "H:\M365Reports\EARL-TempLookupTable-groups-7-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Outfile7 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv8
			if ($checkfile -eq "True")
			{
				$inputCsv8 = Import-Csv $exportreportgrpcsv8 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv8 = "H:\M365Reports\EARL-TempLookupTbl-groups-8-" + $nowfiledate + "-1.csv"
				$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile8 = "H:\M365Reports\EARL-TempLookupTable-groups-8-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Outfile8 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv9
			if ($checkfile -eq "True")
			{
				$inputCsv9 = Import-Csv $exportreportgrpcsv9 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv9 = "H:\M365Reports\EARL-TempLookupTbl-groups-9-" + $nowfiledate + "-1.csv"
				$inputCsv9 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile9 = "H:\M365Reports\EARL-TempLookupTable-groups-9-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv9 | ? { $_.trim() -ne "" } | set-content $Outfile9 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv10
			if ($checkfile -eq "True")
			{
				$inputCsv10 = Import-Csv $exportreportgrpcsv10 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv10 = "H:\M365Reports\EARL-TempLookupTbl-groups-10-" + $nowfiledate + "-1.csv"
				$inputCsv10 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile10 = "H:\M365Reports\EARL-TempLookupTable-groups-10-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv10 | ? { $_.trim() -ne "" } | set-content $Outfile10 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv11
			if ($checkfile -eq "True")
			{
				$inputCsv11 = Import-Csv $exportreportgrpcsv11 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv11 = "H:\M365Reports\EARL-TempLookupTbl-groups-11-" + $nowfiledate + "-1.csv"
				$inputCsv11 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile11 = "H:\M365Reports\EARL-TempLookupTable-groups-11-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv11 | ? { $_.trim() -ne "" } | set-content $Outfile11 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv12
			if ($checkfile -eq "True")
			{
				$inputCsv12 = Import-Csv $exportreportgrpcsv12 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv12 = "H:\M365Reports\EARL-TempLookupTbl-groups-12-" + $nowfiledate + "-1.csv"
				$inputCsv12 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile12 = "H:\M365Reports\EARL-TempLookupTable-groups-12-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv12 | ? { $_.trim() -ne "" } | set-content $Outfile12 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv13
			if ($checkfile -eq "True")
			{
				$inputCsv13 = Import-Csv $exportreportgrpcsv13 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv13 = "H:\M365Reports\EARL-TempLookupTbl-groups-13-" + $nowfiledate + "-1.csv"
				$inputCsv13 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile13 = "H:\M365Reports\EARL-TempLookupTable-groups-13-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv13 | ? { $_.trim() -ne "" } | set-content $Outfile13 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv14
			if ($checkfile -eq "True")
			{
				$inputCsv14 = Import-Csv $exportreportgrpcsv14 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv14 = "H:\M365Reports\EARL-TempLookupTbl-groups-14-" + $nowfiledate + "-1.csv"
				$inputCsv14 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile14 = "H:\M365Reports\EARL-TempLookupTable-groups-14-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv14 | ? { $_.trim() -ne "" } | set-content $Outfile14 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv15
			if ($checkfile -eq "True")
			{
				$inputCsv15 = Import-Csv $exportreportgrpcsv15 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv15 = "H:\M365Reports\EARL-TempLookupTbl-groups-15-" + $nowfiledate + "-1.csv"
				$inputCsv15 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv15 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile15 = "H:\M365Reports\EARL-TempLookupTable-groups-15-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv15 | ? { $_.trim() -ne "" } | set-content $Outfile15 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv16
			if ($checkfile -eq "True")
			{
				$inputCsv16 = Import-Csv $exportreportgrpcsv16 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv16 = "H:\M365Reports\EARL-TempLookupTbl-groups-16-" + $nowfiledate + "-1.csv"
				$inputCsv16 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv16 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile16 = "H:\M365Reports\EARL-TempLookupTable-groups-16-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv16 | ? { $_.trim() -ne "" } | set-content $Outfile16 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv17
			if ($checkfile -eq "True")
			{
				$inputCsv17 = Import-Csv $exportreportgrpcsv17 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv17 = "H:\M365Reports\EARL-TempLookupTbl-groups-17-" + $nowfiledate + "-1.csv"
				$inputCsv17 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv17 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile17 = "H:\M365Reports\EARL-TempLookupTable-groups-17-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv17 | ? { $_.trim() -ne "" } | set-content $Outfile17 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv18
			if ($checkfile -eq "True")
			{
				$inputCsv18 = Import-Csv $exportreportgrpcsv18 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv18 = "H:\M365Reports\EARL-TempLookupTbl-groups-18-" + $nowfiledate + "-1.csv"
				$inputCsv18 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv18 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile18 = "H:\M365Reports\EARL-TempLookupTable-groups-18-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv18 | ? { $_.trim() -ne "" } | set-content $Outfile18 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv19
			if ($checkfile -eq "True")
			{
				$inputCsv19 = Import-Csv $exportreportgrpcsv19 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv19 = "H:\M365Reports\EARL-TempLookupTbl-groups-19-" + $nowfiledate + "-1.csv"
				$inputCsv19 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv19 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile19 = "H:\M365Reports\EARL-TempLookupTable-groups-19-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv19 | ? { $_.trim() -ne "" } | set-content $Outfile19 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv20
			if ($checkfile -eq "True")
			{
				$inputCsv20 = Import-Csv $exportreportgrpcsv20 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv20 = "H:\M365Reports\EARL-TempLookupTbl-groups-20-" + $nowfiledate + "-1.csv"
				$inputCsv20 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv20 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile20 = "H:\M365Reports\EARL-TempLookupTable-groups-20-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv20 | ? { $_.trim() -ne "" } | set-content $Outfile20 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv21
			if ($checkfile -eq "True")
			{
				$inputCsv21 = Import-Csv $exportreportgrpcsv21 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv21 = "H:\M365Reports\EARL-TempLookupTbl-groups-21-" + $nowfiledate + "-1.csv"
				$inputCsv21 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv21 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile21 = "H:\M365Reports\EARL-TempLookupTable-groups-21-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv21 | ? { $_.trim() -ne "" } | set-content $Outfile21 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv22
			if ($checkfile -eq "True")
			{
				$inputCsv22 = Import-Csv $exportreportgrpcsv22 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv22 = "H:\M365Reports\EARL-TempLookupTbl-groups-22-" + $nowfiledate + "-1.csv"
				$inputCsv22 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv22 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile22 = "H:\M365Reports\EARL-TempLookupTable-groups-22-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv22 | ? { $_.trim() -ne "" } | set-content $Outfile22 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv23
			if ($checkfile -eq "True")
			{
				$inputCsv23 = Import-Csv $exportreportgrpcsv23 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv23 = "H:\M365Reports\EARL-TempLookupTbl-groups-23-" + $nowfiledate + "-1.csv"
				$inputCsv23 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv23 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile23 = "H:\M365Reports\EARL-TempLookupTable-groups-23-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv23 | ? { $_.trim() -ne "" } | set-content $Outfile23 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv24
			if ($checkfile -eq "True")
			{
				$inputCsv24 = Import-Csv $exportreportgrpcsv24 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv24 = "H:\M365Reports\EARL-TempLookupTbl-groups-24-" + $nowfiledate + "-1.csv"
				$inputCsv24 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv24 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile24 = "H:\M365Reports\EARL-TempLookupTable-groups-24-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv24 | ? { $_.trim() -ne "" } | set-content $Outfile24 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv25
			if ($checkfile -eq "True")
			{
				$inputCsv25 = Import-Csv $exportreportgrpcsv25 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv25 = "H:\M365Reports\EARL-TempLookupTbl-groups-25-" + $nowfiledate + "-1.csv"
				$inputCsv25 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv25 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile25 = "H:\M365Reports\EARL-TempLookupTable-groups-25-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv25 | ? { $_.trim() -ne "" } | set-content $Outfile25 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv26
			if ($checkfile -eq "True")
			{
				$inputCsv26 = Import-Csv $exportreportgrpcsv26 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv26 = "H:\M365Reports\EARL-TempLookupTbl-groups-26-" + $nowfiledate + "-1.csv"
				$inputCsv26 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv26 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile26 = "H:\M365Reports\EARL-TempLookupTable-groups-26-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv26 | ? { $_.trim() -ne "" } | set-content $Outfile26 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv27
			if ($checkfile -eq "True")
			{
				$inputCsv27 = Import-Csv $exportreportgrpcsv27 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv27 = "H:\M365Reports\EARL-TempLookupTbl-groups-27-" + $nowfiledate + "-1.csv"
				$inputCsv27 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv27 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile27 = "H:\M365Reports\EARL-TempLookupTable-groups-27-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv27 | ? { $_.trim() -ne "" } | set-content $Outfile27 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv28
			if ($checkfile -eq "True")
			{
				$inputCsv28 = Import-Csv $exportreportgrpcsv28 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv28 = "H:\M365Reports\EARL-TempLookupTbl-groups-28-" + $nowfiledate + "-1.csv"
				$inputCsv28 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv28 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile28 = "H:\M365Reports\EARL-TempLookupTable-groups-28-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv28 | ? { $_.trim() -ne "" } | set-content $Outfile28 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv29
			if ($checkfile -eq "True")
			{
				$inputCsv29 = Import-Csv $exportreportgrpcsv29 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv29 = "H:\M365Reports\EARL-TempLookupTbl-groups-29-" + $nowfiledate + "-1.csv"
				$inputCsv29 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv29 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile29 = "H:\M365Reports\EARL-TempLookupTable-groups-29-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv29 | ? { $_.trim() -ne "" } | set-content $Outfile29 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv30
			if ($checkfile -eq "True")
			{
				$inputCsv30 = Import-Csv $exportreportgrpcsv30 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv30 = "H:\M365Reports\EARL-TempLookupTbl-groups-30-" + $nowfiledate + "-1.csv"
				$inputCsv30 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv30 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile30 = "H:\M365Reports\EARL-TempLookupTable-groups-30-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv30 | ? { $_.trim() -ne "" } | set-content $Outfile30 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv31
			if ($checkfile -eq "True")
			{
				$inputCsv31 = Import-Csv $exportreportgrpcsv31 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv31 = "H:\M365Reports\EARL-TempLookupTbl-groups-31-" + $nowfiledate + "-1.csv"
				$inputCsv31 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv31 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile31 = "H:\M365Reports\EARL-TempLookupTable-groups-31-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv31 | ? { $_.trim() -ne "" } | set-content $Outfile31 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv32
			if ($checkfile -eq "True")
			{
				$inputCsv32 = Import-Csv $exportreportgrpcsv32 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv32 = "H:\M365Reports\EARL-TempLookupTbl-groups-32-" + $nowfiledate + "-1.csv"
				$inputCsv32 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv32 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile32 = "H:\M365Reports\EARL-TempLookupTable-groups-32-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv32 | ? { $_.trim() -ne "" } | set-content $Outfile32 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv33
			if ($checkfile -eq "True")
			{
				$inputCsv33 = Import-Csv $exportreportgrpcsv33 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv33 = "H:\M365Reports\EARL-TempLookupTbl-groups-33-" + $nowfiledate + "-1.csv"
				$inputCsv33 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv33 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile33 = "H:\M365Reports\EARL-TempLookupTable-groups-33-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv33 | ? { $_.trim() -ne "" } | set-content $Outfile33 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv34
			if ($checkfile -eq "True")
			{
				$inputCsv34 = Import-Csv $exportreportgrpcsv34 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv34 = "H:\M365Reports\EARL-TempLookupTbl-groups-34-" + $nowfiledate + "-1.csv"
				$inputCsv34 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv34 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile34 = "H:\M365Reports\EARL-TempLookupTable-groups-34-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv34 | ? { $_.trim() -ne "" } | set-content $Outfile34 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv35
			if ($checkfile -eq "True")
			{
				$inputCsv35 = Import-Csv $exportreportgrpcsv35 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv35 = "H:\M365Reports\EARL-TempLookupTbl-groups-35-" + $nowfiledate + "-1.csv"
				$inputCsv35 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv35 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile35 = "H:\M365Reports\EARL-TempLookupTable-groups-35-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv35 | ? { $_.trim() -ne "" } | set-content $Outfile35 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv36
			if ($checkfile -eq "True")
			{
				$inputCsv36 = Import-Csv $exportreportgrpcsv36 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv36 = "H:\M365Reports\EARL-TempLookupTbl-groups-36-" + $nowfiledate + "-1.csv"
				$inputCsv36 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv36 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile36 = "H:\M365Reports\EARL-TempLookupTable-groups-36-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv36 | ? { $_.trim() -ne "" } | set-content $Outfile36 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv37
			if ($checkfile -eq "True")
			{
				$inputCsv37 = Import-Csv $exportreportgrpcsv37 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv37 = "H:\M365Reports\EARL-TempLookupTbl-groups-37-" + $nowfiledate + "-1.csv"
				$inputCsv37 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv37 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile37 = "H:\M365Reports\EARL-TempLookupTable-groups-37-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv37 | ? { $_.trim() -ne "" } | set-content $Outfile37 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv38
			if ($checkfile -eq "True")
			{
				$inputCsv38 = Import-Csv $exportreportgrpcsv38 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv38 = "H:\M365Reports\EARL-TempLookupTbl-groups-38-" + $nowfiledate + "-1.csv"
				$inputCsv38 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv38 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile38 = "H:\M365Reports\EARL-TempLookupTable-groups-38-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv38 | ? { $_.trim() -ne "" } | set-content $Outfile38 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv39
			if ($checkfile -eq "True")
			{
				$inputCsv39 = Import-Csv $exportreportgrpcsv39 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv39 = "H:\M365Reports\EARL-TempLookupTbl-groups-39-" + $nowfiledate + "-1.csv"
				$inputCsv39 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv39 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile39 = "H:\M365Reports\EARL-TempLookupTable-groups-39-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv39 | ? { $_.trim() -ne "" } | set-content $Outfile39 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv40
			if ($checkfile -eq "True")
			{
				$inputCsv40 = Import-Csv $exportreportgrpcsv40 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv40 = "H:\M365Reports\EARL-TempLookupTbl-groups-40-" + $nowfiledate + "-1.csv"
				$inputCsv40 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv40 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile40 = "H:\M365Reports\EARL-TempLookupTable-groups-40-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv40 | ? { $_.trim() -ne "" } | set-content $Outfile40 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv41
			if ($checkfile -eq "True")
			{
				$inputCsv41 = Import-Csv $exportreportgrpcsv41 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv41 = "H:\M365Reports\EARL-TempLookupTbl-groups-41-" + $nowfiledate + "-1.csv"
				$inputCsv41 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv41 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile41 = "H:\M365Reports\EARL-TempLookupTable-groups-41-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv41 | ? { $_.trim() -ne "" } | set-content $Outfile41 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv42
			if ($checkfile -eq "True")
			{
				$inputCsv42 = Import-Csv $exportreportgrpcsv42 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv42 = "H:\M365Reports\EARL-TempLookupTbl-groups-42-" + $nowfiledate + "-1.csv"
				$inputCsv42 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv42 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile42 = "H:\M365Reports\EARL-TempLookupTable-groups-42-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv42 | ? { $_.trim() -ne "" } | set-content $Outfile42 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv43
			if ($checkfile -eq "True")
			{
				$inputCsv43 = Import-Csv $exportreportgrpcsv43 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv43 = "H:\M365Reports\EARL-TempLookupTbl-groups-43-" + $nowfiledate + "-1.csv"
				$inputCsv43 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv43 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile43 = "H:\M365Reports\EARL-TempLookupTable-groups-43-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv43 | ? { $_.trim() -ne "" } | set-content $Outfile43 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv44
			if ($checkfile -eq "True")
			{
				$inputCsv44 = Import-Csv $exportreportgrpcsv44 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv44 = "H:\M365Reports\EARL-TempLookupTbl-groups-44-" + $nowfiledate + "-1.csv"
				$inputCsv44 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv44 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile44 = "H:\M365Reports\EARL-TempLookupTable-groups-44-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv44 | ? { $_.trim() -ne "" } | set-content $Outfile44 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv45
			if ($checkfile -eq "True")
			{
				$inputCsv45 = Import-Csv $exportreportgrpcsv45 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv45 = "H:\M365Reports\EARL-TempLookupTbl-groups-45-" + $nowfiledate + "-1.csv"
				$inputCsv45 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv45 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile45 = "H:\M365Reports\EARL-TempLookupTable-groups-45-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv45 | ? { $_.trim() -ne "" } | set-content $Outfile45 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv46
			if ($checkfile -eq "True")
			{
				$inputCsv46 = Import-Csv $exportreportgrpcsv46 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv46 = "H:\M365Reports\EARL-TempLookupTbl-groups-46-" + $nowfiledate + "-1.csv"
				$inputCsv46 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv46 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile46 = "H:\M365Reports\EARL-TempLookupTable-groups-46-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv46 | ? { $_.trim() -ne "" } | set-content $Outfile46 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv47
			if ($checkfile -eq "True")
			{
				$inputCsv47 = Import-Csv $exportreportgrpcsv47 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv47 = "H:\M365Reports\EARL-TempLookupTbl-groups-47-" + $nowfiledate + "-1.csv"
				$inputCsv47 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv47 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile47 = "H:\M365Reports\EARL-TempLookupTable-groups-47-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv47 | ? { $_.trim() -ne "" } | set-content $Outfile47 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv48
			if ($checkfile -eq "True")
			{
				$inputCsv48 = Import-Csv $exportreportgrpcsv48 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv48 = "H:\M365Reports\EARL-TempLookupTbl-groups-48-" + $nowfiledate + "-1.csv"
				$inputCsv48 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv48 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile48 = "H:\M365Reports\EARL-TempLookupTable-groups-48-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv48 | ? { $_.trim() -ne "" } | set-content $Outfile48 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv49
			if ($checkfile -eq "True")
			{
				$inputCsv49 = Import-Csv $exportreportgrpcsv49 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv49 = "H:\M365Reports\EARL-TempLookupTbl-groups-49-" + $nowfiledate + "-1.csv"
				$inputCsv49 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv49 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile49 = "H:\M365Reports\EARL-TempLookupTable-groups-49-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv49 | ? { $_.trim() -ne "" } | set-content $Outfile49 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportgrpcsv50
			if ($checkfile -eq "True")
			{
				$inputCsv50 = Import-Csv $exportreportgrpcsv50 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv50 = "H:\M365Reports\EARL-TempLookupTbl-groups-50-" + $nowfiledate + "-1.csv"
				$inputCsv50 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv50 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile50 = "H:\M365Reports\EARL-TempLookupTable-groups-50-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv50 | ? { $_.trim() -ne "" } | set-content $Outfile50 -Encoding UTF8
				
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
			
			
			if (Test-Path $Outfile17)
			{
				
				
				Try
				{
					#Map-Filewatcher
					Copy-item -path $Outfile17 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile17 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile17 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile18 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile18 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile19 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile19 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile19 | $now"
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
					Add-Content $logfile "Group File Copied to FileWatcher $Outfile20 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile20 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					
					Copy-item -path $Outfile21 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile21 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile21 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						#RemoveFilewatcher
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile21 | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile22 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile22 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile23 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile23 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile24 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile24 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile25 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile25 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile26 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile26 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile27 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile27 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile28 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile28 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile29 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile29 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile29 | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile30 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile30 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					
					Copy-item -path $Outfile31 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile31 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile31 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						#RemoveFilewatcher
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile32 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile32 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile33 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile33 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile34 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile34 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile35 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile35 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile36 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile36 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile37 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile37 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile38 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile38 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile39 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile39 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile39 | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile40 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile40 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					
					Copy-item -path $Outfile41 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "group File Copied to FileWatcher $Outfile41 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile41 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						#RemoveFilewatcher
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile42 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile42 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile43 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile43 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile44 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile44 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile45 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile45 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile46 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile46 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile47 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile47 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile48 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile48 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile49 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile49 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile49 | $now"
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
					Add-Content $logfile "group File Copied to FileWatcher $Outfile50 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "Groups File Copied to FileWatcher $Outfile50 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 30
						
						
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile50 | $now"
					}
				}
				
			}
			
			
			#cleanup files
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Clearing up files no longer needed for groups | $now"
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
			
			if (Test-Path $exportreportgrpcsv15)
			{
				Remove-Item $exportreportgrpcsv15
			}
			
			if (Test-Path $finaloutcsv15)
			{
				Remove-Item $finaloutcsv15
			}
			
			if (Test-Path $exportreportgrpcsv16)
			{
				Remove-Item $exportreportgrpcsv16
			}
			
			if (Test-Path $finaloutcsv16)
			{
				Remove-Item $finaloutcsv16
			}
			
			if (Test-Path $exportreportgrpcsv17)
			{
				Remove-Item $exportreportgrpcsv17
			}
			
			if (Test-Path $finaloutcsv17)
			{
				Remove-Item $finaloutcsv17
			}
			
			if (Test-Path $exportreportgrpcsv18)
			{
				Remove-Item $exportreportgrpcsv18
			}
			
			if (Test-Path $finaloutcsv18)
			{
				Remove-Item $finaloutcsv18
			}
			
			if (Test-Path $exportreportgrpcsv19)
			{
				Remove-Item $exportreportgrpcsv19
			}
			
			if (Test-Path $finaloutcsv19)
			{
				Remove-Item $finaloutcsv19
			}
			
			if (Test-Path $exportreportgrpcsv20)
			{
				Remove-Item $exportreportgrpcsv20
			}
			
			if (Test-Path $finaloutcsv20)
			{
				Remove-Item $finaloutcsv20
			}
			
			if (Test-Path $exportreportgrpcsv21)
			{
				Remove-Item $exportreportgrpcsv21
			}
			
			if (Test-Path $finaloutcsv21)
			{
				Remove-Item $finaloutcsv21
			}
			
			if (Test-Path $exportreportgrpcsv22)
			{
				Remove-Item $exportreportgrpcsv22
			}
			
			if (Test-Path $finaloutcsv22)
			{
				Remove-Item $finaloutcsv22
			}
			
			if (Test-Path $exportreportgrpcsv23)
			{
				Remove-Item $exportreportgrpcsv23
			}
			
			if (Test-Path $finaloutcsv23)
			{
				Remove-Item $finaloutcsv23
			}
			
			if (Test-Path $exportreportgrpcsv24)
			{
				Remove-Item $exportreportgrpcsv24
			}
			
			if (Test-Path $finaloutcsv24)
			{
				Remove-Item $finaloutcsv24
			}
			
			if (Test-Path $exportreportgrpcsv25)
			{
				Remove-Item $exportreportgrpcsv25
			}
			
			if (Test-Path $finaloutcsv25)
			{
				Remove-Item $finaloutcsv25
			}
			
			if (Test-Path $exportreportgrpcsv26)
			{
				Remove-Item $exportreportgrpcsv26
			}
			
			if (Test-Path $finaloutcsv26)
			{
				Remove-Item $finaloutcsv26
			}
			
			if (Test-Path $exportreportgrpcsv27)
			{
				Remove-Item $exportreportgrpcsv27
			}
			
			if (Test-Path $finaloutcsv27)
			{
				Remove-Item $finaloutcsv27
			}
			
			if (Test-Path $exportreportgrpcsv28)
			{
				Remove-Item $exportreportgrpcsv28
			}
			
			if (Test-Path $finaloutcsv28)
			{
				Remove-Item $finaloutcsv28
			}
			
			if (Test-Path $exportreportgrpcsv29)
			{
				Remove-Item $exportreportgrpcsv29
			}
			
			if (Test-Path $finaloutcsv29)
			{
				Remove-Item $finaloutcsv29
			}
			
			if (Test-Path $exportreportgrpcsv30)
			{
				Remove-Item $exportreportgrpcsv30
			}
			
			if (Test-Path $finaloutcsv30)
			{
				Remove-Item $finaloutcsv30
			}
			
			if (Test-Path $exportreportgrpcsv31)
			{
				Remove-Item $exportreportgrpcsv31
			}
			
			if (Test-Path $finaloutcsv31)
			{
				Remove-Item $finaloutcsv31
			}
			
			if (Test-Path $exportreportgrpcsv32)
			{
				Remove-Item $exportreportgrpcsv32
			}
			
			if (Test-Path $finaloutcsv32)
			{
				Remove-Item $finaloutcsv32
			}
			
			if (Test-Path $exportreportgrpcsv33)
			{
				Remove-Item $exportreportgrpcsv33
			}
			
			if (Test-Path $finaloutcsv33)
			{
				Remove-Item $finaloutcsv33
			}
			
			if (Test-Path $exportreportgrpcsv34)
			{
				Remove-Item $exportreportgrpcsv34
			}
			
			if (Test-Path $finaloutcsv34)
			{
				Remove-Item $finaloutcsv34
			}
			
			if (Test-Path $exportreportgrpcsv35)
			{
				Remove-Item $exportreportgrpcsv35
			}
			
			if (Test-Path $finaloutcsv35)
			{
				Remove-Item $finaloutcsv35
			}
			
			if (Test-Path $exportreportgrpcsv36)
			{
				Remove-Item $exportreportgrpcsv36
			}
			
			if (Test-Path $finaloutcsv36)
			{
				Remove-Item $finaloutcsv36
			}
			
			if (Test-Path $exportreportgrpcsv37)
			{
				Remove-Item $exportreportgrpcsv37
			}
			
			if (Test-Path $finaloutcsv37)
			{
				Remove-Item $finaloutcsv37
			}
			
			if (Test-Path $exportreportgrpcsv38)
			{
				Remove-Item $exportreportgrpcsv38
			}
			
			if (Test-Path $finaloutcsv38)
			{
				Remove-Item $finaloutcsv38
			}
			
			if (Test-Path $exportreportgrpcsv39)
			{
				Remove-Item $exportreportgrpcsv39
			}
			
			if (Test-Path $finaloutcsv39)
			{
				Remove-Item $finaloutcsv39
			}
			
			if (Test-Path $exportreportgrpcsv40)
			{
				Remove-Item $exportreportgrpcsv40
			}
			
			if (Test-Path $finaloutcsv40)
			{
				Remove-Item $finaloutcsv40
			}
			
			
			
			
		}
		RemoveFilewatcher
	}
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	Add-Content $logfile "Finished with exports for Temp LookupTable for Groups | $now"
	
}


exportremoteusermbx
#exportDL
#exportcontact
#exportmailuser
#exportlocalmbx



#exportremotesharedroommbx



$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "All finished with exports for LDAP replacement LookupTable hourly | $now"
DisconnectExchangeOnPrem
RemoveFilewatcher

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "Closing down script - bye $now"
Stop-Transcript

Disconnect-EXO
DisconnectExchangeOnPrem

Exit-PSSession
Exit



