



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


function Process-CsvFileremotembx
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
		$finalOutCsv = "$outputFolder\EARL-TempLookupTbl-remotesharedmbx-$index-$nowFileDate-1.csv"
		$outputFile = "$outputFolder\EARL-TempLookupTable-remotesharedmbx-$index-$nowFileDate-2.csv"
		
		New-Variable -Name "outfile$index" -Value $outputFile -Force
		#$Outfile1
		
		
		$inputCsv | Sort-Object -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } |
		Export-Csv -Path $finalOutCsv -NoTypeInformation -Delimiter "|" -Encoding UTF8
		
		Start-Sleep -Seconds 5
		
		Get-Content -Path $finalOutCsv | Where-Object { $_.Trim() -ne "" } | Set-Content -Path $outputFile -Encoding UTF8
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







function exportremotesharedmbx
{
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export remote shared mailboxes for export"
	
	
	$count = ""
	$M = ""
	$reconnect = 0
	$MbxNumber = 0
	
	
	Start-Sleep -Seconds 1
	# Set the base file path
	$basePath = "H:\M365Reports\EARL-Prod-TempLookupTbl-remotesharedmbx-"
	$dateFormat = "yyyy-MM-dd-hh-mm-ss"
	
	# Loop through numbers 1 to 30 to create file paths
	for ($i = 1; $i -le 60; $i++)
	{
		$nowfiledate = Get-Date -Format $dateFormat
		# Construct the file path with the current index
		$filePath = "$basePath$i-$nowfiledate.csv"
		# Dynamically create a variable name for each file path
		New-Variable -Name "exportreportusercsv$i" -Value $filePath  -Force
		
		Start-Sleep -Seconds 1
	}
	
	
	$Mbx2 = Get-RemoteMailbox -ResultSize unlimited -filter "RecipientTypedetails 'RemotesharedMailbox'" | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	
	
	$count = $mbx2.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count remote shared mailbox accounts to process and upload for lookuptable refresh [Temp]"
	
	add-content $logfile  "LookupTable Exporting to file $exportreportusercsv1  for remote mbx to refresh [Temp] "
	
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
		
		
		if ($count -ge 1)
		{
			$outputFolder = "H:\M365Reports"
			
			for ($i = 1; $i -le 50; $i++)
			{
				$csvFilePath = Get-Variable -Name "exportreportusercsv$i" -ValueOnly
				if ($csvFilePath)
				{
					Process-CsvFileremotembx -csvFilePath $csvFilePath -outputFolder $outputFolder -index $i
				}
				
				Start-Sleep -Seconds 2
			}
		}
		
		
				
			
			#copy to filewatcher	
			
			$Outfiles = @(
				$Outfile1,
				$Outfile2,
				$Outfile3,
				$Outfile4,
				$Outfile5,
				$Outfile6,
				$Outfile7,
				$Outfile8,
				$Outfile9,
				$Outfile10,
				$Outfile11,
				$Outfile12,
				$Outfile13,
				$Outfile14,
				$Outfile15,
				$Outfile16,
				$Outfile17,
				$Outfile18,
				$Outfile19,
				$Outfile20,
				$Outfile21,
				$Outfile22,
				$Outfile23,
				$Outfile24,
				$Outfile25,
				$Outfile26,
				$Outfile27,
				$Outfile28,
				$Outfile29,
				$Outfile30,
				$Outfile31,
				$Outfile32,
				$Outfile33,
				$Outfile34,
				$Outfile35,
				$Outfile36,
				$Outfile37,
				$Outfile38,
				$Outfile39,
				$Outfile40,
				$Outfile41,
				$Outfile42,
				$Outfile43,
				$Outfile44,
				$Outfile45,
				$Outfile46,
				$Outfile47,
				$Outfile48,
				$Outfile49,
				$Outfile50
			)
			
			Map-Filewatcher
			
			foreach ($Outfile in $Outfiles)
			{
				if (Test-Path $Outfile)
				{
					Try
					{
						Copy-item -path $Outfile -destination $filewatcherout
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						Add-Content $logfile "Mailboxes File Copied to FileWatcher $Outfile to $filewatcherout [1st try] | $now"
						Start-Sleep -Seconds 30
					}
					catch
					{
						Start-Sleep -Seconds 30
						try
						{
							RemoveFilewatcher
							Start-Sleep -Seconds 15
							Map-Filewatcher
							Copy-item -path $Outfile -destination $filewatcherout
							$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
							Add-Content $logfile "Mailboxes File Copied to FileWatcher $Outfile to $filewatcherout [2nd try] | $now"
							Start-Sleep -Seconds 30
						}
						catch
						{
							$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
							Add-Content $logfile "Cannot copy files to FileWatcher $Outfile | $now"
						}
					}
				}
			}
			
			RemoveFilewatcher
			
			
			
			
			#cleanup files
			
			# Define the number of CSV files
			$csvCount = 50
			
			# Loop through each numbered CSV variable
			for ($i = 1; $i -le $csvCount; $i++)
			{
				$exportVar = "exportreportusercsv$i"
				$finalVar = "finaloutcsv$i"
				
				if (Test-Path (Get-Variable -Name $exportVar).Value)
				{
					Remove-Item (Get-Variable -Name $exportVar).Value
				}
				
				if (Test-Path (Get-Variable -Name $finalVar).Value)
				{
					Remove-Item (Get-Variable -Name $finalVar).Value
				}
			}
			
		
		
	}
	
}








exportremotesharedmbx





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



