



<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.148
	 Created on:    07/08/2023 14:00
	 Created by:   	Neil Poultney
	 Organization: 	NDP Consultancy Ltd - neil.poultney@ndpconsultancyltd.co.uk
	 Filename:     	EARL-get-lookupTable-Export-guests.ps1
	===========================================================================
	.DESCRIPTION
		Exports for EARL MailDb to Lookup Table on hourly basis.

		Change Log
		V1.00, 11/10/2023 - Initial full version
		V1.01, 21/11/2023 - updated for description field

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
$transcriptlog = "H:\EARLTranscripts\LookupTbl\lookup-export-Guests-Refresh-" + $nowfiledate + ".log"

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
	$logfilelocation = $loglocation + "Refresh-Lookup-Table-Guests-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
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
	$logfilelocation = $loglocation + "Refresh-Lookup-Table-Guests-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
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
		Add-Content $logfile "Attempting to connect to Exchange OnPremise remote powershell Prod | $now"
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
		Add-Content $logfile "EXO Guest Users File Copied to FileWatcher $Outfile to $filewatcherout [1st try] | $now"
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
			Add-Content $logfile "EXO Guests File Copied to FileWatcher $Outfile to $filewatcherout [2nd try] | $now"
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



function exportguests
{
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export guests for export to Refresh LookupTable"
	
	$Mbx5 = ""
	$exportreportcsv5 = ""
	$M = ""
	$reconnect = 0
	
	
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	
	
	# Set the base file path
	$GRPbasePath = "H:\M365Reports\EARL-TempLookupTable-Guestmailuser"
	
	# Loop through numbers 1 to 20 to create file paths
	for ($i = 1; $i -le 30; $i++)
	{
		# Get the current date and time in the specified format
		$nowfiledate = Get-Date -Format "yyyy-MM-dd-hh-mm-ss"
		
		# Construct the file path with the current index
		$GRPfilePath = "$GRPbasePath-$i-$nowfiledate.csv"
		
		# Dynamically create a variable name for each file path
		New-Variable -Name "exportreportguestcsv$i" -Value $GRPfilePath -Force
		
		# Sleep for 1 seconds between file path creations
		Start-Sleep -Seconds 1
	}
	
	
	
	
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsvAll = "H:\M365Reports\EARL-Temp-LookupTable-AllGuests-" + $nowfiledate + ".csv"
	
	Connect-EXO
	
	

	Start-Sleep -s 5

	#$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	#temp
	#$lookuptime1 = (get-date).adddays(-720)
	$lookuptime1 = (get-date).addhours(-4)
	Set-Variable -Name lasthour -Value $lookuptime1 -Option ReadOnly -Scope Script -Force
	
	#$Mbx5 = Get-MailContact -ResultSize unlimited | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	$Mbx5 = Get-Recipient -ResultSize unlimited -filter " RecipientTypeDetails -eq 'GuestMailUser' " | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, ExternalemailAddress, HiddenFromAddressListsEnabled, whenchanged, whencreated, Manager, LastName, FirstName, WindowsLiveID, Name, Alias
	
	$count = $mbx5.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count guest accounts to process"
	
	add-content $logfile  "LookupTable Exporting to file  $exportreportcsv5 for Guest Users "
	
	
	
	
	
	
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
			$liveId = $M.WindowsLiveID
			
			#write-host "Getting mailbox:: $usr"
			$getacccount = 0
			try
			{
				$acc = Get-user $usr | Select-Object *
			}
			catch
			{
				$acc = Get-user $liveID | Select-Object *
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "unable to get details for $usr trying with $liveId - skipping"
			}
			
			$getacccount = $acc.count
			#$recpdetails = Get-Recipient -identity $usr -properties *
			#$recpdetails = Get-Recipient -identity $usr | Select-Object *
			
			if ($getacccount -lt 2)
			{
				[int]$Progress = $MbxNumber/$count * 100
				$PercentComplete = [math]::Round($Progress, 3)
				
				[string]$disp = $M.DisplayName
				$UPN = $acc.UserPrincipalName
				$MailboxType = $M.msExchRecipientTypeDetails
				$mail = $M.PrimarySmtpAddress
				$SN = $acc.LastName
				$firstName = $acc.FirstName
				$dept = "NULL"
				$Comp = $M.Company
				$country = $acc.CountryOrRegion
				$UsrACCCtrl = "0"
				$NTID = $acc.samaccountname
				$distName = $acc.DistinguishedName
				$createtype = $acc.CreationType
				$persona = $acc.UserPersona
				$whencreated = $acc.WhenCreated
				$disacc = $acc.AccountDisabled
				$descript = "GuestMailUser: " + $createtype + " - " + $persona + "- Created: " + $whencreated + " - AccountDisabled: $disacc"
				$managerDN = $M.Manager
				$BPtext3201 = "NULL"
				$CA9 = "N"
				$HideAB = $acc.HiddenFromAddressListsEnabled
				
				
				if ($null -ne $managerDN)
				{
					$mgrout = Get-Recipient $managerDN | Select-Object DisplayName, PrimarySMTPAddress, Alias
					
					$managerDisp = $mgrout.DisplayName
					$mgrEmail = $mgrout.PrimarySMTPAddress
					$mgrAlias = $mgrout.Alias
					
				}
				
				if ($recipientdetailsEX -eq "GuestMailUser" -and ! $MailboxType)
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
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset1 = "1"
				}
				
				if (($attributecountset2 -eq "0") -and ($PercentComplete -eq "10.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset2 = "1"
				}
				
				if (($attributecountset3 -eq "0") -and ($PercentComplete -eq "15.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset3 = "1"
				}
				
				if (($attributecountset4 -eq "0") -and ($PercentComplete -eq "20.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset4 = "1"
				}
				
				if (($attributecountset5 -eq "0") -and ($PercentComplete -eq "25.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset5 = "1"
				}
				
				if (($attributecountset6 -eq "0") -and ($PercentComplete -eq "30.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset6 = "1"
				}
				
				if (($attributecountset7 -eq "0") -and ($PercentComplete -eq "35.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset7 = "1"
				}
				
				if (($attributecountset8 -eq "0") -and ($PercentComplete -eq "40.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset8 = "1"
				}
				
				if (($attributecountset9 -eq "0") -and ($PercentComplete -eq "45.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset9 = "1"
				}
				
				if (($attributecountset10 -eq "0") -and ($PercentComplete -eq "50.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset10 = "1"
				}
				
				
				if (($attributecountset11 -eq "0") -and ($PercentComplete -eq "55.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset11 = "1"
				}
				
				if (($attributecountset12 -eq "0") -and ($PercentComplete -eq "60.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset12 = "1"
				}
				
				if (($attributecountset13 -eq "0") -and ($PercentComplete -eq "65.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset13 = "1"
				}
				
				if (($attributecountset14 -eq "0") -and ($PercentComplete -eq "70.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset14 = "1"
				}
				
				if (($attributecountset15 -eq "0") -and ($PercentComplete -eq "75.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset15 = "1"
				}
				
				if (($attributecountset16 -eq "0") -and ($PercentComplete -eq "80.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset16 = "1"
				}
				
				if (($attributecountset17 -eq "0") -and ($PercentComplete -eq "85.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset17 = "1"
				}
				
				if (($attributecountset18 -eq "0") -and ($PercentComplete -eq "90.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset18 = "1"
				}
				
				if (($attributecountset19 -eq "0") -and ($PercentComplete -eq "95.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset19 = "1"
				}
				
				if (($attributecountset20 -eq "0") -and ($PercentComplete -eq "99.000"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Processed GuestMailUser number : $MbxNumber | Percent Complete: $PercentComplete | $now "
					$attributecountset20 = "1"
				}
				
				
				
				If (($Disp) -and ($mail) )
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
					
					
					
					$exportReportPaths = @(
						$exportreportguestcsv1,
						$exportreportguestcsv2,
						$exportreportguestcsv3,
						$exportreportguestcsv4,
						$exportreportguestcsv5,
						$exportreportguestcsv6,
						$exportreportguestcsv7,
						$exportreportguestcsv8,
						$exportreportguestcsv9,
						$exportreportguestcsv10,
						$exportreportguestcsv11,
						$exportreportguestcsv12,
						$exportreportguestcsv13,
						$exportreportguestcsv14,
						$exportreportguestcsv15,
						$exportreportguestcsv16,
						$exportreportguestcsv17,
						$exportreportguestcsv18,
						$exportreportguestcsv19,
						$exportreportguestcsv20,
						$exportreportguestcsv21,
						$exportreportguestcsv22,
						$exportreportguestcsv23,
						$exportreportguestcsv24,
						$exportreportguestcsv25
						
					)
					
					Export-ReportLine -MBXNumber $MbxNumber -reportLine $ReportLine3 -exportReportPaths $exportReportPaths
					
					<#
					if ($MbxNumber -le 1000)
					{
						$ReportLine3 | Export-CSV $exportreportusercsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 1000) -and ($MbxNumber -le 2000))
					{
						$ReportLine3 | Export-CSV $exportreportusercsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 2000) -and ($MbxNumber -le 3000))
					{
						$ReportLine3 | Export-CSV $exportreportusercsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 3000) -and ($MbxNumber -le 4000))
					{
						$ReportLine3 | Export-CSV $exportreportusercsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -gt 4000) -and ($MbxNumber -le 5000))
					{
						$ReportLine3 | Export-CSV $exportreportusercsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					#>
					
									
					$ReportLine3 | Export-CSV $exportreportcsvAll -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					
				}
			}
			
		}
		
		
		
		
		
		Disconnect-EXO
		
		
		$exportfiles = @(
			@{ Path = $exportreportguestcsv1; BaseName = "EARL-TmpLookupTable-guestmailuser-1" },
			@{ Path = $exportreportguestcsv2; BaseName = "EARL-TmpLookupTable-guestmailuser-2" },
			@{ Path = $exportreportguestcsv3; BaseName = "EARL-TmpLookupTable-guestmailuser-3" },
			@{ Path = $exportreportguestcsv4; BaseName = "EARL-TmpLookupTable-guestmailuser-4" },
			@{ Path = $exportreportguestcsv5; BaseName = "EARL-TmpLookupTable-guestmailuser-5" },
			@{ Path = $exportreportguestcsv6; BaseName = "EARL-TmpLookupTable-guestmailuser-6" },
			@{ Path = $exportreportguestcsv7; BaseName = "EARL-TmpLookupTable-guestmailuser-7" },
			@{ Path = $exportreportguestcsv8; BaseName = "EARL-TmpLookupTable-guestmailuser-8" },
			@{ Path = $exportreportguestcsv9; BaseName = "EARL-TmpLookupTable-guestmailuser-9" },
			@{ Path = $exportreportguestcsv10; BaseName = "EARL-TmpLookupTable-guestmailuser-10" },
			@{ Path = $exportreportguestcsv11; BaseName = "EARL-TmpLookupTable-guestmailuser-11" },
			@{ Path = $exportreportguestcsv12; BaseName = "EARL-TmpLookupTable-guestmailuser-12" },
			@{ Path = $exportreportguestcsv13; BaseName = "EARL-TmpLookupTable-guestmailuser-13" },
			@{ Path = $exportreportguestcsv14; BaseName = "EARL-TmpLookupTable-guestmailuser-14" },
			@{ Path = $exportreportguestcsv15; BaseName = "EARL-TmpLookupTable-guestmailuser-15" },
			@{ Path = $exportreportguestcsv16; BaseName = "EARL-TmpLookupTable-guestmailuser-16" },
			@{ Path = $exportreportguestcsv17; BaseName = "EARL-TmpLookupTable-guestmailuser-17" },
			@{ Path = $exportreportguestcsv18; BaseName = "EARL-TmpLookupTable-guestmailuser-18" },
			@{ Path = $exportreportguestcsv19; BaseName = "EARL-TmpLookupTable-guestmailuser-19" },
			@{ Path = $exportreportguestcsv20; BaseName = "EARL-TmpLookupTable-guestmailuser-20" },
			@{ Path = $exportreportguestcsv21; BaseName = "EARL-TmpLookupTable-guestmailuser-21" },
			@{ Path = $exportreportguestcsv22; BaseName = "EARL-TmpLookupTable-guestmailuser-22" },
			@{ Path = $exportreportguestcsv23; BaseName = "EARL-TmpLookupTable-guestmailuser-23" },
			@{ Path = $exportreportguestcsv24; BaseName = "EARL-TmpLookupTable-guestmailuser-24" },
			@{ Path = $exportreportguestcsv25; BaseName = "EARL-TmpLookupTable-guestmailuser-25" },
			@{ Path = $exportreportguestcsv26; BaseName = "EARL-TmpLookupTable-guestmailuser-26" },
			@{ Path = $exportreportguestcsv27; BaseName = "EARL-TmpLookupTable-guestmailuser-27" },
			@{ Path = $exportreportguestcsv28; BaseName = "EARL-TmpLookupTable-guestmailuser-28" },
			@{ Path = $exportreportguestcsv29; BaseName = "EARL-TmpLookupTable-guestmailuser-29" },
			@{ Path = $exportreportguestcsv30; BaseName = "EARL-TmpLookupTable-guestmailuser-30" }
			
			


			
		)
		
		Map-Filewatcher
		# Process each file
		foreach ($file in $exportfiles)
		{
			Process-CsvFile -filePath $file.Path -outputBaseName $file.BaseName
			Start-Sleep -Seconds 2
		}
		
		
		RemoveFilewatcher
		
		
		<#
		
		if ($count -ge 1)
		{
			
			#sortoutput so no blank lines and no duplicates
			$inputCsv1 = Import-Csv $exportreportusercsv1 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-1-" + $nowfiledate + "-1.csv"
			$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile1 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-1-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv | ? { $_.trim() -ne "" } | set-content $Outfile1 -Encoding UTF8
			
			
			$checkfile = Test-Path -Path $exportreportusercsv3
			if ($checkfile -eq "True")
			{
				$inputCsv2 = Import-Csv $exportreportusercsv2 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv2 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-2-" + $nowfiledate + "-1.csv"
				$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile2 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-2-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Outfile2 -Encoding UTF8
			}
			
			$checkfile = Test-Path -Path $exportreportusercsv3
			if ($checkfile -eq "True")
			{
				$inputCsv3 = Import-Csv $exportreportusercsv3 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv3 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-3-" + $nowfiledate + "-1.csv"
				$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
				Start-Sleep -s 5
				$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
				$Outfile3 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-3-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Outfile3 -Encoding UTF8
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv4
			if ($checkfile -eq "True")
			{
				$inputCsv4 = Import-Csv $exportreportusercsv4 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv4 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-4-" + $nowfiledate + "-1.csv"
				$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile4 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-4-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Outfile4 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv5
			if ($checkfile -eq "True")
			{
				$inputCsv5 = Import-Csv $exportreportusercsv5 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv5 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-5-" + $nowfiledate + "-1.csv"
				$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile5 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-5-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Outfile5 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv6
			if ($checkfile -eq "True")
			{
				$inputCsv6 = Import-Csv $exportreportusercsv6 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv6 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-6-" + $nowfiledate + "-1.csv"
				$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile6 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-6-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Outfile6 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv7
			if ($checkfile -eq "True")
			{
				$inputCsv7 = Import-Csv $exportreportusercsv7 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv7 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-7-" + $nowfiledate + "-1.csv"
				$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile7 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-7-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Outfile7 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv8
			if ($checkfile -eq "True")
			{
				$inputCsv8 = Import-Csv $exportreportusercsv8 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv8 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-8-" + $nowfiledate + "-1.csv"
				$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile8 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-8-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Outfile8 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv9
			if ($checkfile -eq "True")
			{
				$inputCsv9 = Import-Csv $exportreportusercsv9 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv9 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-9-" + $nowfiledate + "-1.csv"
				$inputCsv9 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile9 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-9-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv9 | ? { $_.trim() -ne "" } | set-content $Outfile9 -Encoding UTF8
				
			}
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv10
			if ($checkfile -eq "True")
			{
				$inputCsv10 = Import-Csv $exportreportusercsv10 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv10 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-10-" + $nowfiledate + "-1.csv"
				$inputCsv10 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile10 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-10-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv10 | ? { $_.trim() -ne "" } | set-content $Outfile10 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv11
			if ($checkfile -eq "True")
			{
				$inputCsv11 = Import-Csv $exportreportusercsv11 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv11 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-11-" + $nowfiledate + "-1.csv"
				$inputCsv11 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile11 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-11-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv11 | ? { $_.trim() -ne "" } | set-content $Outfile11 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv12
			if ($checkfile -eq "True")
			{
				$inputCsv12 = Import-Csv $exportreportusercsv12 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv12 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-12-" + $nowfiledate + "-1.csv"
				$inputCsv12 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile12 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-12-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv12 | ? { $_.trim() -ne "" } | set-content $Outfile12 -Encoding UTF8
				
			}
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv13
			if ($checkfile -eq "True")
			{
				$inputCsv13 = Import-Csv $exportreportusercsv13 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv13 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-13-" + $nowfiledate + "-1.csv"
				$inputCsv13 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile13 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-13-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv13 | ? { $_.trim() -ne "" } | set-content $Outfile13 -Encoding UTF8
				
			}
			
			
			Start-Sleep -s 2
			$checkfile = Test-Path -Path $exportreportusercsv14
			if ($checkfile -eq "True")
			{
				$inputCsv14 = Import-Csv $exportreportusercsv14 -delimiter "|" | Sort-Object * -Unique
				$finaloutcsv14 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-14-" + $nowfiledate + "-1.csv"
				$inputCsv14 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8
				$Outfile14 = "H:\M365Reports\EARL-LookupTbl-guestmailuser-hourly-14-" + $nowfiledate + "-2.csv"
				gc $finaloutcsv14 | ? { $_.trim() -ne "" } | set-content $Outfile14 -Encoding UTF8
				
			}
			
			
			
			Map-Filewatcher
			if (Test-Path $Outfile1)
			{
				Try
				{
					
					Copy-item -path $Outfile1 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile1 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $outfile1 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile2 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile2 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile3 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile3 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile4 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile4 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile5 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile5 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile6 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile6 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile7 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile7 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile8 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile8 to $filewatcherout [2nd try] | $now"
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
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile9 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile9 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
						RemoveFilewatcher
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "GuestMailUser copy files to FileWatcher $Outfile9 | $now"
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
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile10 to $filewatcherout [1st try] | $now"
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
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile10 to $filewatcherout [2nd try] | $now"
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
			
			if (Test-Path $Outfile11)
			{
				
				
				Try
				{
					#Map-Filewatcher
					Copy-item -path $Outfile11 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile11 to $filewatcherout [1st try] | $now"
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
						Copy-item -path $Outfile11 -destination $filewatcherout
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "GuestMailUser File Copied to FileWatcher $Outfile11 to $filewatcherout [2nd try] | $now"
						Start-Sleep -Seconds 360
						
						RemoveFilewatcher
					}
					catch
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Cannot copy files to FileWatcher $Outfile11 | $now"
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




exportguests




$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "All finished with exports for LDAP replacement LookupTable All guests | $now"
Disconnect-EXO
RemoveFilewatcher

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "Closing down script - bye $now"
Stop-Transcript


Exit-PSSession
Exit



