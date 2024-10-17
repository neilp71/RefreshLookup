



<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.148
	 Created on:    18/08/2024 14:00
	 Created by:   	Neil Poultney
	 Organization: 	NDP Consultancy Ltd - neil.poultney@ndpconsultancyltd.co.uk
	 Filename:     	EARL-get-GRPlookupTable-Export-Full.ps1
	===========================================================================
	.DESCRIPTION
		Exports for EARL MailDb to GRPLookup Table.

		Change Log
		V1.00, 18/08/2024 14:00 - Initial full version
		

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

##workoutwhereweare deprecated WMI
#$Domainwearein = (Get-WmiObject Win32_ComputerSystem).Name

$Domainwearein = (Get-CimInstance Win32_ComputerSystem).Name
$whoweare = $ENV:USERNAME

if ($domainwearein -eq "BP1GXEIS801") { $global:Envirionmentchoice = "Dev" }
if ($domainwearein -eq "zneepacp11emrg") { $global:Envirionmentchoice = "ProdNE"; $global:ProcEARLServer = "zneepacp11emrg" }
if ($domainwearein -eq "zweepacp11em50") { $global:Envirionmentchoice = "ProdWE"; $global:ProcEARLServer = "zweepacp11em50" }

if ($domainwearein -eq "zneepacp11eme2") { $global:Envirionmentchoice = "ProdNE"; $global:ProcEARLServer = "zneepacp11eme2" }
if ($domainwearein -eq "zneepacp11emfk") { $global:Envirionmentchoice = "ProdNE"; $global:ProcEARLServer = "zneepacp11emfk" }
if ($domainwearein -eq "zweepacp11emg3") { $global:Envirionmentchoice = "ProdWE"; $global:ProcEARLServer = "zweepacp11emg3" }
if ($domainwearein -eq "zweepacp11emce") { $global:Envirionmentchoice = "ProdWE"; $global:ProcEARLServer = "zweepacp11emce" }



$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
$transcriptlog = "H:\EARLTranscripts\GRPLookup\TempGRPlookup-export-OnPremSG-" + $nowfiledate + ".log"

Start-Transcript -Path $transcriptlog

$now
$ENV:USERNAME
$Domainwearein
$Envirionmentchoice


if (($domainwearein -eq "zneepacp11eme2") -or ($domainwearein -eq "zneepacp11emfk"))
{
	
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$loglocation = "H:\EARLPSLogs\GRPLookup\" # change to usetype RPMBCREATE etc
	Set-Variable -Name logfolder -Value $logslocation -Option ReadOnly -Scope Script -Force
	$logfilelocation = $loglocation + "TempGRPLookup-Table-OnPremSG-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
	Set-Variable -Name logfile -Value $logfilelocation -Option ReadOnly -Scope Script -Force
	$EARLNTID = "BP1\task-EARLEXCNE1"
	$EARLNTID2 = "BP1\task-EARLEXCNE2"
	$secureAES = "F:\AppCerts\PSUserCred\secureaes.key"
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



if ($domainwearein -eq "zweepacp11em50")
{
	
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$loglocation = "H:\EARLPSLogs\GRPLookup\"  # change to usetype RPMBCREATE etc
	Set-Variable -Name logfolder -Value $logslocation -Option ReadOnly -Scope Script -Force
	$logfilelocation = $loglocation + "TempGRPLookup-Table-OnPremSG-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
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

if ($domainwearein -eq "zneepacp11emrg")
{
	
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$loglocation = "H:\EARLPSLogs\GRPLookup\"  # change to usetype RPMBCREATE etc
	Set-Variable -Name logfolder -Value $logslocation -Option ReadOnly -Scope Script -Force
	$logfilelocation = $loglocation + "TempGRPLookup-Table-OnPremSG-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
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
	Add-Content $logfile "Exchange OnPremise remote powershell $Envirionmentchoice || $now"
	
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
		$randomchoice = Get-Random -Minimum 1 -Maximum 2
		
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
			$exchangesession = New-PSSession -Credential $Credentials -ConfigurationName Microsoft.Exchange -authentication Basic -ConnectionUri https://mail.bp.com/PowerShell -AllowRedirection
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
		$randomchoice = Get-Random -Minimum 1 -Maximum 2
		
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
			$exchangesession = New-PSSession -Credential $Credentials -ConfigurationName Microsoft.Exchange -authentication Basic -ConnectionUri https://mail.bp.com/PowerShell -AllowRedirection
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
	
	
	
	if (($ProcEARLServer -eq "zneepacp11emrg") -or ($ProcEARLServer -eq "zweepacp11em50"))
	{
		
		if ($failconnect -le "10")
		{
			
			#APP Connection EARL EXO Reports
			#Connect-ExchangeOnline -CertificateThumbprint "a98251f44faf329cd3d1474f1440aca8356edaa0" -AppID "920938ea-809a-4a52-bf9e-0ae65fd12d53" -Organization "bp365.onmicrosoft.com" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
			Connect-ExchangeOnline -CertificateThumbprint "8f901a3fbc0f30746f1f5309806314aa32841e2b" -AppID "920938ea-809a-4a52-bf9e-0ae65fd12d53" -Organization "bp365.onmicrosoft.com" -ShowProgress $false -SkipLoadingCmdletHelp -ShowBanner:$false -EA SilentlyContinue -EV silentErr
			
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
	
	
	if (($ProcEARLServer -eq "zneepacp11eme2") -or ($ProcEARLServer -eq "zneepacp11emfk"))
	{
		$randomchoice = ""
		$Connectchoice = ""
		$Error.Clear()
		#get-Pssession | remove-pssession
		
		
		$modules = get-module | select name
		if (!$Modules -like "*ExchangeOnline*")
		{
			
			Import-Module ExchangeOnlineManagement
		}
		
		
		##check if there is already a broken connection - then re-use if possible
		
		$Checksessionstate = ""
		$EXOconnectionavailableA1 = ""
		$EXOconnectionavailableB1 = ""
		$outcheck = ""
		$Checksessionstate = Get-ConnectionInformation
		
		$EXOconnectionavailableA1 = $Checksessionstate.TokenStatus
		$EXOconnectionavailableB1 = $Checksessionstate.Name
		$EXOconnectionavailableC1 = $Checksessionstate.State
		
		if (($EXOconnectionavailableA1 -eq "Expired") -and ($EXOconnectionavailableB1 -like "ExchangeOnline*") -and ($EXOconnectionavailableC1 -eq "Broken"))
		{
			$outcheck = Get-EXOMailbox -ResultSize 1
			
			if ($outcheck)
			{
				$Checksessionstate = ""
				$EXOconnectionavailableA1 = ""
				$EXOconnectionavailableB1 = ""
				$EXOconnectionavailableC1 = ""
				
				$Checksessionstate = Get-ConnectionInformation
				
				$EXOconnectionavailableA1 = $Checksessionstate.TokenStatus
				$EXOconnectionavailableB1 = $Checksessionstate.Name
				$EXOconnectionavailableC1 = $Checksessionstate.State
				
				if (($EXOconnectionavailableA1 -eq "Active") -and ($EXOconnectionavailableB1 -like "ExchangeOnline*") -and ($EXOconnectionavailableC1 -eq "Connected"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Connection re-established : Token status : $EXOconnectionavailableA1 | Connection Name : $EXOconnectionavailableB1 | Connection state : $EXOconnectionavailableC1 | $now"
				}
			}
			
			
		}
		
		if ((! $Checksessionstate) -or (!$outcheck) -or (($EXOconnectionavailableA1 -ne "Expired") -and ($EXOconnectionavailableB1 -notlike "ExchangeOnline*") -and ($EXOconnectionavailableC1 -ne "Broken")))
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "I Need to establish a new ExchangeOnline connection | $now"
			
			
			$randomchoice = Get-Random -Minimum 1 -Maximum 3
			
			
			if ($randomchoice -eq 1) { $Connectchoice = "NE3App" }
			if ($randomchoice -eq 2) { $Connectchoice = "NE4App" }
			
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
			Add-Content $logfile "Connecting to Exchange Online EXO V3 APP connection $Connectchoice (thumbprint)  | $now "
			Write-Host "Connecting to Exchange Online EXO V3 APP connection $Connectchoice (thumbprint)  | $now "
			
			
			
			
			
			if ($Connectchoice -eq "NE3App")
			{
				#Connect-ExchangeOnline -CertificateThumbprint "c02052ba5de2dfaa3f65108d00f267be316d657c" -AppID "f3dbffb6-788b-4ae5-a236-feec4897c1ac" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
				Connect-ExchangeOnline -CertificateThumbprint "3d03ccf5d36dd7beecac18814b9afba517f2eb59" -AppID "f3dbffb6-788b-4ae5-a236-feec4897c1ac" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
				
				$time = get-date -Format dd-MM-yyyy-HH:mm:ss
				$connecttry = 0
			}
			
			
			if ($Connectchoice -eq "NE4App")
			{
				#Connect-ExchangeOnline -CertificateThumbprint "37c0375f3791d6c9fa67bc70b2775911b1a14ae6" -AppID "2c35c6a7-23ce-4e56-936e-2c75c3b8101e" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
				Connect-ExchangeOnline -CertificateThumbprint "85db39cf7dcb5ec59233c031d5b9cf0c3734339c" -AppID "2c35c6a7-23ce-4e56-936e-2c75c3b8101e" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
				
				$time = get-date -Format dd-MM-yyyy-HH:mm:ss
				$connecttry = 0
			}
			
			
			$EXOconnectionavailableA = ""
			$EXOconnectionavailableB = ""
			$Tokenexpiry = ""
			$certauthen = ""
			$appId = ""
			$sessionstateA = ""
			$EarlApp = ""
			
			#$sessionstateA = Get-PSSession | Select-Object -Property State, Name
			$sessionstateA = Get-ConnectionInformation
			
			#$EXOconnectionavailableA = $sessionstateA.state
			#$sessionstateA = Get-ConnectionInformation
			
			if ($SessionstateA)
			{
				$EXOconnectionavailableA = $sessionstateA.TokenStatus
				$EXOconnectionavailableB = $sessionstateA.Name
				$Tokenexpiry = $sessionstateA.TokenExpiryTimeUTC
				$certauthen = $sessionstateA.CertificateAuthentication
				$appId = $sessionstateA.AppID
				
				
				
				if ($appId -eq "f3dbffb6-788b-4ae5-a236-feec4897c1ac")
				{
					$EarlApp = "NE3App"
				}
				
				if ($appId -eq "2c35c6a7-23ce-4e56-936e-2c75c3b8101e")
				{
					$EarlApp = "NE4App"
				}
				
				
				
			}
			
			
			
			
			if (($EXOconnectionavailableA -eq "Active") -and ($EXOconnectionavailableB -like "ExchangeOnline*"))
			#if (($EXOconnectionavailableA -eq "Opened") -and ($EXOconnectionavailableB -like "ExchangeOnlineInternal*"))
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
				Add-Content $logfile "Connected to Exchange Online EXO V3 APP connection $Connectchoice (thumbprint).... | Name : $EXOconnectionavailableB | Token Expiry : $Tokenexpiry | AppId : [$EarlApp] $appId | CertAuthentication : $certauthen  $now "
				Write-Host "Connected to Exchange Online EXO V3 APP connection $Connectchoice (thumbprint).... | Name : $EXOconnectionavailableB | Token Expiry : $Tokenexpiry | AppId : [$EarlApp] $appId | CertAuthentication : $certauthen  $now "
			}
			
			#if (($EXOconnectionavailableA -ne "Opened") -and ($EXOconnectionavailableB -notlike "ExchangeOnlineInternal*"))
			if ((($EXOconnectionavailableA -ne "Active") -and ($EXOconnectionavailableB -notlike "ExchangeOnline*")) -or (!$sessionStateA))
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
				Add-Content $logfile "Unable to Connect to  Exchange Online EXO V3 $Connectchoice (thumbprint) | AppId : [$EarlApp] $appId | $now "
				Write-Host "Unable to Connect to  Exchange Online EXO V3 $Connectchoice (thumbprint) | AppId : [$EarlApp] $appId| $now "
				
				
				if ($Connectchoice -eq "NE3App")
				{
					#Connect-ExchangeOnline -CertificateThumbprint "e950757aeea06c1c5a611705e02a1d36a98104f7" -AppID "de9a91a6-3b83-4985-9fbf-53b241af54d2" -Organization "bp365.onmicrosoft.com" -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					#Connect-ExchangeOnline -CertificateThumbprint "37c0375f3791d6c9fa67bc70b2775911b1a14ae6" -AppID "2c35c6a7-23ce-4e56-936e-2c75c3b8101e" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					Connect-ExchangeOnline -CertificateThumbprint "85db39cf7dcb5ec59233c031d5b9cf0c3734339c" -AppID "2c35c6a7-23ce-4e56-936e-2c75c3b8101e" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					
					
					
					$time = get-date -Format dd-MM-yyyy-HH:mm:ss
					$connecttry = 0
				}
				
				if ($Connectchoice -eq "NE4App")
				{
					#Connect-ExchangeOnline -CertificateThumbprint "e950757aeea06c1c5a611705e02a1d36a98104f7" -AppID "de9a91a6-3b83-4985-9fbf-53b241af54d2" -Organization "bp365.onmicrosoft.com" -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					#Connect-ExchangeOnline -CertificateThumbprint "c02052ba5de2dfaa3f65108d00f267be316d657c" -AppID "f3dbffb6-788b-4ae5-a236-feec4897c1ac" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					Connect-ExchangeOnline -CertificateThumbprint "3d03ccf5d36dd7beecac18814b9afba517f2eb59" -AppID "f3dbffb6-788b-4ae5-a236-feec4897c1ac" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					
					
					$time = get-date -Format dd-MM-yyyy-HH:mm:ss
					$connecttry = 0
				}
				
				
				$EXOconnectionavailableC = ""
				$EXOconnectionavailableD = ""
				$Tokenexpiry = ""
				$certauthen = ""
				$appId = ""
				$sessionstateB = ""
				
				
				$sessionstateB = Get-ConnectionInformation
				
				if ($SessionstateB)
				{
					$EXOconnectionavailableC = $sessionstateB.TokenStatus
					$EXOconnectionavailbleD = $sessionstateB.Name
					
					$Tokenexpiry = $sessionstateB.TokenExpiryTimeUTC
					$certauthen = $sessionstateB.CertificateAuthentication
					$appId = $sessionstateB.AppID
					
					if ($appId -eq "de9a91a6-3b83-4985-9fbf-53b241af54d2")
					{
						$EarlApp = "NE1App"
					}
					
					if ($appId -eq "e9408c38-a082-4763-9184-bc84c9bb5e63")
					{
						$EarlApp = "NE2App"
					}
					
					if ($appId -eq "f3dbffb6-788b-4ae5-a236-feec4897c1ac")
					{
						$EarlApp = "NE3App"
					}
					
					if ($appId -eq "2c35c6a7-23ce-4e56-936e-2c75c3b8101e")
					{
						$EarlApp = "NE4App"
					}
					
					
					
				}
				if ((($EXOconnectionavailableC -eq "Active") -and ($EXOconnectionavailableD -like "ExchangeOnline*")) -or (!$SessionStateB))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
					Add-Content $logfile "Connected to Exchange Online EXO V3 APP connection $Connectchoice (thumbprint).... Name : $EXOconnectionavailableD | Token Expiry : $Tokenexpiry | AppId : [$EarlApp] $appId | CertAuthentication : $certauthen  $now "
					Write-Host "Connected to Exchange Online EXO V3 APP connection $Connectchoice (thumbprint).... Name : $EXOconnectionavailableD | Token Expiry : $Tokenexpiry | AppId : [$EarlApp] $appId | CertAuthentication : $certauthen  $now "
				}
				
				#if (($EXOconnectionavailableC -ne "Opened") -and ($EXOconnectionavailableD -notlike "ExchangeOnlineInternal*"))
				if (($EXOconnectionavailableC -ne "Active") -and ($EXOconnectionavailableD -notlike "ExchangeOnline*"))
				{
					
					$connecttry = $connecttry + 1
					$time = get-date -Format dd-MM-yyyy-HH:mm:ss
					#write-host "Could not connect to Exchange Online EXO V2"
					
					start-sleep -s 120
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd hh:mm:ss")
					if ($connecttry -eq "10")
					{
						Add-Content $logfile "Could not connect to EXO V3 commands I have tried $connecttry times to both App IDs in NE and will quit | $now "
						Write-Host "Could not connect to EXO V3 commands I have tried $connecttry times to both App IDs in NE and will quit | $now "
						Exit
					}
					else
					{
						Add-Content $logfile "Could not connect to EXO V3 commands I have tried $connecttry times and will try again in 1 minute  | $now "
						Write-Host "Could not connect to EXO V3 commands I have tried $connecttry times and will try again in 1 minute  | $now "
						Connect-EXO
					}
					
				}
				
				
			}
		}
	}
	
	
	if (($ProcEARLServer -eq "zweepacp11emg3") -or ($ProcEARLServer -eq "zweepacp11emce"))
	{
		$randomchoice = ""
		$Connectchoice = ""
		$Error.Clear()
		#get-Pssession | remove-pssession
		
		
		$modules = get-module | select name
		if (!$Modules -like "*ExchangeOnline*")
		{
			
			Import-Module ExchangeOnlineManagement
		}
		
		##check if there is already a broken connection - then re-use if possible
		
		$Checksessionstate = ""
		$EXOconnectionavailableA1 = ""
		$EXOconnectionavailableB1 = ""
		$outcheck = ""
		$Checksessionstate = Get-ConnectionInformation
		
		$EXOconnectionavailableA1 = $Checksessionstate.TokenStatus
		$EXOconnectionavailableB1 = $Checksessionstate.Name
		$EXOconnectionavailableC1 = $Checksessionstate.State
		
		if (($EXOconnectionavailableA1 -eq "Expired") -and ($EXOconnectionavailableB1 -like "ExchangeOnline*") -and ($EXOconnectionavailableC1 -eq "Broken"))
		{
			$outcheck = Get-EXOMailbox -ResultSize 1
			
			if ($outcheck)
			{
				$Checksessionstate = ""
				$EXOconnectionavailableA1 = ""
				$EXOconnectionavailableB1 = ""
				$EXOconnectionavailableC1 = ""
				
				$Checksessionstate = Get-ConnectionInformation
				
				$EXOconnectionavailableA1 = $Checksessionstate.TokenStatus
				$EXOconnectionavailableB1 = $Checksessionstate.Name
				$EXOconnectionavailableC1 = $Checksessionstate.State
				
				if (($EXOconnectionavailableA1 -eq "Active") -and ($EXOconnectionavailableB1 -like "ExchangeOnline*") -and ($EXOconnectionavailableC1 -eq "Connected"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Connection re-established : Token status : $EXOconnectionavailableA1 | Connection Name : $EXOconnectionavailableB1 | Connection state : $EXOconnectionavailableC1 | $now"
				}
			}
			
			
		}
		
		if ((! $Checksessionstate) -or (!$outcheck) -or (($EXOconnectionavailableA1 -ne "Expired") -and ($EXOconnectionavailableB1 -notlike "ExchangeOnline*") -and ($EXOconnectionavailableC1 -ne "Broken")))
		{
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "I Need to establish a new ExchangeOnline connection | $now"
			
			$randomchoice = Get-Random -Minimum 1 -Maximum 3
			
			if ($randomchoice -eq 1) { $Connectchoice = "WE3App" }
			if ($randomchoice -eq 2) { $Connectchoice = "WE4App" }

			
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "Connection choice $Connectchoice | $now"
			Write-Host "Connection choice $Connectchoice | $now"
			

			
			if ($Connectchoice -eq "WE3App")
			{
				#Connect-ExchangeOnline -CertificateThumbprint "1afe35507431e0d5a7c0c0e7ed8537f3a323ca62" -AppID "b54d9fb1-88ff-4f8d-aa2d-4f7db11fb988" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
				Connect-ExchangeOnline -CertificateThumbprint "a146944da02e307172ce6b1e8fecad773ba303d8" -AppID "b54d9fb1-88ff-4f8d-aa2d-4f7db11fb988" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
				
				
				
				$time = get-date -Format dd-MM-yyyy-HH:mm:ss
				$connecttry = 0
			}
			
			if ($Connectchoice -eq "WE4App")
			{
				#Connect-ExchangeOnline -CertificateThumbprint "2b82fe273372045d0f8586e29ac03a4af69ce704" -AppID "ef78df15-91d7-438c-bdb0-efd2055d11fd" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
				Connect-ExchangeOnline -CertificateThumbprint "01db657b6f59a1a384c3803b0a4f53a5036f6840" -AppID "ef78df15-91d7-438c-bdb0-efd2055d11fd" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
				
				
				$time = get-date -Format dd-MM-yyyy-HH:mm:ss
				$connecttry = 0
			}
			
			$EarlApp = ""
			#$sessionstateA = Get-PSSession | Select-Object -Property State, Name
			$sessionstateA = Get-ConnectionInformation
			
			
			$EXOconnectionavailableA = $sessionstateA.TokenStatus
			$EXOconnectionavailableB = $sessionstateA.Name
			$Tokenexpiry = $sessionstateA.TokenExpiryTimeUTC
			$certauthen = $sessionstateA.CertificateAuthentication
			$appId = $sessionstateA.AppID
			
			#$EXOconnectionavailableA = $sessionstateA.state
			#$EXOconnectionavailableB = $sessionstateA.Name
			

			
			if ($appId -eq "b54d9fb1-88ff-4f8d-aa2d-4f7db11fb988")
			{
				$EarlApp = "WE3App"
			}
			
			if ($appId -eq "ef78df15-91d7-438c-bdb0-efd2055d11fd")
			{
				$EarlApp = "WE4App"
			}
			
			if (($EXOconnectionavailableA -eq "Active") -and ($EXOconnectionavailableB -like "ExchangeOnline*"))
			
			#if (($EXOconnectionavailableA -eq "Opened") -and ($EXOconnectionavailableB -like "ExchangeOnlineInternal*"))
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Connected to Exchange Online EXO V3 APP connection with limited commands $Connectchoice (thumbprint).... | Token Expiry : $Tokenexpiry | AppId : [$EarlApp] $appId | CertAuthentication : $certauthen  $now "
				Write-Host "Connected to Exchange Online EXO V3 APP connection with limited commands $Connectchoice (thumbprint).... | Token Expiry : $Tokenexpiry | AppId : [$EarlApp] $appId | CertAuthentication : $certauthen  $now "
			}
			
			#if (($EXOconnectionavailableA -ne "Opened") -and ($EXOconnectionavailableB -notlike "ExchangeOnlineInternal*"))
			if (($EXOconnectionavailableA -ne "Active") -and ($EXOconnectionavailableB -notlike "ExchangeOnline*"))
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Unable to Connect to  Exchange Online EXO V3 $Connectchoice (thumbprint).... Name : $EXOconnectionavailableB | Token Status : $EXOconnectionavailable1 | Token Expiry : $Tokenexpiry | AppId : [$EarlApp] $appId | CertAuthentication : $certauthen  $now"
				

				
				
				if ($Connectchoice -eq "WE3App")
				{
					#APP Connection WE4
					#Connect-ExchangeOnline -CertificateThumbprint "6fe884540f82e07c68da0a87554454eeef0f113e" -AppID "60873704-3696-4378-b19b-64993c08440e" -Organization "bp365.onmicrosoft.com" -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					#Connect-ExchangeOnline -CertificateThumbprint "2b82fe273372045d0f8586e29ac03a4af69ce704" -AppID "ef78df15-91d7-438c-bdb0-efd2055d11fd" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					Connect-ExchangeOnline -CertificateThumbprint "01db657b6f59a1a384c3803b0a4f53a5036f6840" -AppID "ef78df15-91d7-438c-bdb0-efd2055d11fd" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					
					
					$time = get-date -Format dd-MM-yyyy-HH:mm:ss
					$connecttry = 0
					
				}
				
				if ($Connectchoice -eq "WE4App")
				{
					#APP Connection WE3
					#Connect-ExchangeOnline -CertificateThumbprint "6fe884540f82e07c68da0a87554454eeef0f113e" -AppID "60873704-3696-4378-b19b-64993c08440e" -Organization "bp365.onmicrosoft.com" -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					#Connect-ExchangeOnline -CertificateThumbprint "1afe35507431e0d5a7c0c0e7ed8537f3a323ca62" -AppID "b54d9fb1-88ff-4f8d-aa2d-4f7db11fb988" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					Connect-ExchangeOnline -CertificateThumbprint "a146944da02e307172ce6b1e8fecad773ba303d8" -AppID "b54d9fb1-88ff-4f8d-aa2d-4f7db11fb988" -Organization "bp365.onmicrosoft.com" -SkipLoadingCmdletHelp -CommandName "Get-User,Set-User,Get-Mailbox,Get-EXOMailbox,Set-Mailbox,Add-Mailboxpermission,Remove-MailboxPermission,Get-RecipientPermission,Add-RecipientPermission,Remove-RecipientPermission,Get-MailboxPermission,Get-EXOMailboxPermission,Get-Recipient,Get-EXORecipient,Get-CASMailbox,Get-EXOCASMailbox,Set-CASMailbox,Get-DistributionGroup,Get-ConnectionInformation,Get-Mailboxstatistics,Get-EXOMailboxstatistics,Get-MailboxFolderstatistics,Get-EXOMailboxFolderstatistics,Get-MailboxFolderPermission,Get-EXOMailboxFolderPermission" -ShowProgress $false -ShowBanner:$false -EA SilentlyContinue -EV silentErr
					
					
					$time = get-date -Format dd-MM-yyyy-HH:mm:ss
					$connecttry = 0
					
				}
				
				#$sessionstateB = Get-PSSession | Select-Object -Property State, Name
				$EarlApp = ""
				$sessionstateB = Get-ConnectionInformation
				
				#$EXOconnectionavailableC = $sessionstateB.state
				$EXOconnectionavailableD = $sessionstateB.Name
				$EXOconnectionavailableC = $sessionstateB.TokenStatus
				$Tokenexpiry = $sessionstateB.TokenExpiryTimeUTC
				$certauthen = $sessionstateB.CertificateAuthentication
				$appId = $sessionstateB.AppID
				
	
				
				if ($appId -eq "b54d9fb1-88ff-4f8d-aa2d-4f7db11fb988")
				{
					$EarlApp = "WE3App"
				}
				
				if ($appId -eq "ef78df15-91d7-438c-bdb0-efd2055d11fd")
				{
					$EarlApp = "WE4App"
				}
				
				
				if (($EXOconnectionavailableC -eq "Active") -and ($EXOconnectionavailableD -like "ExchangeOnline*"))
				
				
				
				#if (($EXOconnectionavailableC -eq "Opened") -and ($EXOconnectionavailableD -like "ExchangeOnlineInternal*"))
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Connected to Exchange Online EXO V3 APP connection with limited commands $Connectchoice (thumbprint).... Name : $EXOconnectionavailableD | Token Expiry : $Tokenexpiry | AppId : [$EarlApp] $appId | CertAuthentication : $certauthen  $now "
				}
				
				#if (($EXOconnectionavailableC -ne "Opened") -and ($EXOconnectionavailableD -notlike "ExchangeOnlineInternal*"))
				if (($EXOconnectionavailableC -ne "Active") -and ($EXOconnectionavailableD -notlike "ExchangeOnline*"))
				{
					
					$connecttry = $connecttry + 1
					$time = get-date -Format dd-MM-yyyy-HH:mm:ss
					#write-host "Could not connect to Exchange Online EXO V2"
					
					start-sleep -s 120
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					if ($connecttry -eq "10")
					{
						Add-Content $logfile "Could not connect to EXO V3 commands I have tried $connecttry times to both App IDs in WE and will quit .... | $now "
						Exit
					}
					else
					{
						Add-Content $logfile "Could not connect to EXO V3 commands I have tried $connecttry times and will try again in 1 minute  | $now "
						Connect-EXO
					}
					
				}
				
				
				
				
				
			}
		}
	}
	
}




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
		$nowfiledate = Get-Date -f "yyyy-MM-dd-hh-mm"
		$inputCsv = Import-Csv $filePath -Delimiter "|" | Sort-Object * -Unique
		$finaloutcsv = "H:\M365Reports\${outputBaseName}-${nowfiledate}-1.csv"
		
		$inputCsv | Sort-Object -Property @{ Expression = { $_.Samaccountname }; Ascending = $false } |
		Export-Csv $finaloutcsv -NoTypeInformation -Delimiter "|" -Encoding UTF8
		
		Start-Sleep -Seconds 5
		
		$nowfiledate = Get-Date -f "yyyy-MM-dd-HH-mm-ss"
		$Outfile = "H:\M365Reports\${outputBaseName}-${nowfiledate}-2.csv"
		
		Get-Content $finaloutcsv | Where-Object { $_.Trim() -ne "" } | Set-Content $Outfile -Encoding UTF8
		
		Try
		{
			
			Copy-item -path $Outfile -destination $filewatcherout
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			Add-Content $logfile "SG on prem group File Copied to FileWatcher $Outfile to $filewatcherout [1st try] | $now"
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
				Add-Content $logfile "Groups File Copied to FileWatcher $Outfile to $filewatcherout [2nd try] | $now"
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
}

function Export-ReportLine
{
	param (
		[int]$GRPNumber,
		[object]$reportLine,
		[string[]]$exportReportPaths
	)
	
	$index = [math]::Floor($GRPNumber / 3000)
	if ($index -lt $exportReportPaths.Length)
	{
		$reportPath = $exportReportPaths[$index]
		$reportLine | Export-Csv $reportPath -NoTypeInformation -Delimiter "|" -Encoding UTF8 -Append -Force
	}
}




function exportSGOnPrem
{
	
	
	
	
	
	[int]$GRPNo = 0
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export Security Groups for export to Temp GrpLookup"
	
	
	
	$M = ""
	$reconnect = 0
	#$exportreportcsv4 = "H:\M365Reports\EARL-Prod-LookupTable-groups-" + $nowfiledate + ".csv"
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	
	
	# Set the base file path
	$GRPbasePath = "H:\M365Reports\Refresh-EARL-GRP-OnPremSG"
	
	# Loop through numbers 1 to 30 to create file paths
	for ($i = 1; $i -le 80; $i++)
	{
		# Get the current date and time in the specified format
		$nowfiledate = Get-Date -Format "yyyy-MM-dd-hh-mm-ss"
		
		# Construct the file path with the current index
		$GRPfilePath = "$GRPbasePath-$i-$nowfiledate.csv"
		
		# Dynamically create a variable name for each file path
		New-Variable -Name "exportreportgrpcsv$i" -Value $GRPfilePath -Force
		
		# Sleep for 1 seconds between file path creations
		Start-Sleep -Seconds 1
	}
	
	
		

	$grpprogresscount = 0
	
	ConnectExchangeonPrem
	$lasthour = (get-date).addhours(-3)
	
	$grp1 = Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize unlimited | Select DisplayName, RecipientTypeDetails, SamAccountName, PrimarySmtpAddress, HiddenFromAddressListsEnabled, Description, Alias, GroupType, IsDirSynced, ExternalDirectoryObjectId, RequireSenderAuthenticationEnabled, DistinguishedName, ManagedBy
	
	
	$count1 = $grp1.count
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count1 on premise SG groups to process"
	
	add-content $logfile  "Exporting for SG groups to temp GRPLookup Table for on premise "
	
	ForEach ($G in $grp1)
	{
		$grpprogresscount = $grpprogresscount + 1
		$reconnect++
		$grpprimarysmtp = $null
		$grpDisplay = $null
		$grprecipientdetailsEX = $null
		$grphideAB = $null
		$grpdescription = $null
		$grpalias = $null
		$grpEXTdirectID = $null
		$grpsenderauth = $null
		$grpDN = $null
		$grpOwner = $null
		$mgrEmail = $null #ManagerEmail
		$mgrAlias = $null #Manager
		$OwnerDisp = $Null
		$OwnerEmail = $Null
		$OwnerAlias = $Null
		$OwnerAccState = $Null
		$restrictfrom = $null
		$acceptfrom = $null
		$sendtodluser = $Null
		$sendtodlgrp = $Null #
		$restrictedsendtouser = "False"
		$restrictedsendtogrp = "False"
		$grpSamAccountName = $Null
		[int]$restrictionDLcount = 0
		[int]$countofacls = 0
		[int]$restrictionusercount = 0
		
		$grpprimarysmtp = $G.PrimarySmtpAddress
		[string]$grpDisplay = $G.DisplayName
		$grprecipientdetailsEX = $G.RecipientTypeDetails
		$grphideAB = $G.HiddenFromAddressListsEnabled
		[string]$grpdescription = $G.Description
		$grpalias = $G.Alias
		$grpEXTdirectID = $G.ExternalDirectoryObjectId
		$grpsenderauth = $G.RequireSenderAuthenticationEnabled
		[string]$grpDN = $G.DistinguishedName
		[string]$grpOwner = $G.ManagedBy
		$locale = "OnPremise"
		$dirsync = $G.IsDirSynced
		$grptype = $G.GroupType
		[string]$grpSamAccountName = $G.SamAccountName
		
		
		
		[int]$Progress = $grpprogresscount/$count1 * 100
		$PercentComplete = [math]::Round($Progress, 3)
		
		
		
		$percentages = 5 .. 95 + 99
		$attributeSets = 1 .. 20
		
		# Create a hashtable to store the attribute count sets
		$attributeCountSets = @{ }
		
		# Initialize the attribute count set variables to "0"
		foreach ($i in $attributeSets)
		{
			$attributeCountSets["attributecountset$i"] = "0"
		}
		
		foreach ($i in 0 .. ($percentages.Length - 1))
		{
			$expectedPercent = $percentages[$i]
			$roundedPercent = [math]::Round($PercentComplete)
			
			if ($attributeCountSets["attributecountset$($attributeSets[$i])"] -eq "0" -and $roundedPercent -eq $expectedPercent)
			{
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile "Processed Group number : $grpprogresscount | Percent Complete: $PercentComplete | $now"
				$attributeCountSets["attributecountset$($attributeSets[$i])"] = "1"
			}
		}
		
		
		
		
		if (($dirsync -match "False") -and ($grptype -like "Universal*"))
		{
			
			$SendToUserNTID = "NULL"
			$SendToUserDisplayName = "NULL"
			$SendToUserEmail = "NULL"
			
			#[array]$sendtodluser = get-distributiongroup $grpprimarysmtp | select-object -expand AcceptMessagesOnlyFromSendersOrMembers
			[array]$sendtodluser = get-distributiongroup $grpprimarysmtp | select-object -expand AcceptMessagesOnlyFrom
			[array]$sendtodlgrp = get-distributiongroup $grpprimarysmtp | select-object -expand AcceptMessagesOnlyFromDLMembers
			#$acc = Get-ADGroup -filter 'mail -eq $grpprimarysmtp' -properties * | Select-Object *
			
			if ($null -ne $sendtodluser)
			{
				$restrictedsendtouser = "True"
				$restrictionusercount = $sendtodluser.count
				
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Group $grpprimarysmtp has $restrictionusercount restrictions on user(s) who can send to the group "
				
				
				if ($restrictionusercount -eq 1)
				{
					[string]$getsendtodluser = $sendtodluser
					
					if ($getsendtodluser -notmatch "^bp1\.ad\.bp\.com/Deletion/Deletions Pending Users/.*")
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Getting restricted send to user $getsendtodluser who can send to the group $grpprimarysmtp "
						$restrictionuserdetails = Get-User $getsendtodluser | Select-Object DisplayName, WindowsEmailAddress, RecipientTypeDetails, AccountDisabled, samaccountname
						
						
						$SendToUserNTID = $restrictionuserdetails.samaccountname
						$SendToUserDisplayName = $restrictionuserdetails.DisplayName
						$SendToUserEmail = $restrictionuserdetails.WindowsEmailAddress
						
						if (($SendToUserNTID -eq $Null) -or ($SendToUserNTID -eq "")) { $SendToUserNTID = "NULL" }
						if (($SendToUserDisplayName -eq $Null) -or ($SendToUserDisplayName -eq "")) { $SendToUserDisplayName = "NULL" }
						if (($SendToUserEmail -eq $Null) -or ($SendToUserEmail -eq "")) { $SendToUserEmail = "NULL" }
					}
					
				}
				
				
				
			}
			
			if ($null -ne $sendtodlgrp)
			{
				$restrictedsendtogrp = "True"
				$restrictionDLcount = $sendtodlgrp.count
				
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Group $grpprimarysmtp has $restrictionDLcount restrictions on groups(s) who can send to the group "
				
				if ($restrictionDLcount -eq 1)
				{
					$getsendtogrp = $sendtodlgrp
					$restrictiongrpdetails = Get-group $getsendtogrp | Select-Object DisplayName, WindowsEmailAddress, RecipientTypeDetails, AccountDisabled, samaccountname
					
					
					
					$SendToDLNTID = $restrictiongrpdetails.samaccountname
					$SendToDLDisplayName = $restrictiongrpdetails.DisplayName
					$SendToDLEmail = $restrictiongrpdetails.WindowsEmailAddress
					
					if (($SendToDLNTID -eq $Null) -or ($SendToDLNTID -eq "")) { $SendToDLNTID = "NULL" }
					if (($SendToDLDisplayName -eq $Null) -or ($SendToDLDisplayName -eq "")) { $SendToDLDisplayName = "NULL" }
					if (($SendToDLEmail -eq $Null) -or ($SendToDLEmail -eq "")) { $SendToDLEmail = "NULL" }
					
				}
				
			}
			
			if (!$sendtodluser)
			{
				$SendToUserNTID = "NULL"
				$SendToUserDisplayName = "NULL"
				$SendToUserEmail = "NULL"
			}
			
			if (!$sendtodlgrp)
			{
				$SendToDLNTID = "NULL"
				$SendToDLDisplayName = "NULL"
				$SendToDLEmail = "NULL"
				
			}
			
			if ($grpsenderauth -eq "False")
			{
				$grpsenderauth = "True"
			}
			
			if ($grpsenderauth -eq "True")
			{
				$grpsenderauth = "False"
			}
			
			#coOwner details 
			
			$CoOwnerNTID = "NULL"
			$CoOwnerDisplayName = "NULL"
			$CoOwnerEmail = "NULL"
			[int]$countofacls = 0
			
			[array]$outdlACL = (get-acl "AD:$($grpDN)").Access | where-object { ($_.IsInherited -ne "True" -and $_.Identityreference -like "BP1\*" -and $_.ActivedirectoryRights -like "*GenericWrite*") } | select IdentityReference
			$countofacls = $outdlACL.count
			
			$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
			add-content $logfile  "Group  $grpprimarysmtp has $countofacls CoOwners  | $now "
			
			if ($countofacls -eq 1)
			{
				$aclfrombp1 = $outdlACL.IdentityReference
				$aclownersamaccount = ($aclfrombp1 -split '\\')[1]
				
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				add-content $logfile  "Getting $aclfrombp1 for CoOwner information | $now "
				
				$aclcoowner = Get-User $aclownersamaccount | Select-Object DisplayName, WindowsEmailAddress, RecipientTypeDetails, AccountDisabled, samaccountname
				
				$CoOwnerNTID = $aclcoowner.samaccountname
				$CoOwnerDisplayName = $aclcoowner.DisplayName
				$CoOwnerEmail = $aclcoowner.WindowsEmailAddress
				
				if (!$CoOwnerEmail) { $CoOwnerEmail = "NULL" }
				if (!$CoOwnerDisplayName) { $CoOwnerDisplayName = "NULL" }
				if (!$CoOwnerEmail) { $CoOwnerEmail = "NULL" }
			}
			
			
			
			
			
			
			
			if ($null -ne $grpOwner)
			{
				$mgrout = Get-User $grpOwner | Select-Object DisplayName, WindowsEmailAddress, RecipientTypeDetails, AccountDisabled, samaccountname
				
				$OwnerDisp = $mgrout.DisplayName
				$OwnerEmail = $mgrout.WindowsEmailAddress
				$OwnerAlias = $mgrout.Samaccountname
				$OwnerAccState = $mgrout.AccountDisabled
			}
			
			if (! $OwnerDisp)
			{
				$OwnerDisp = "NULL"
			}
			
				
			
			if (!$OwnerEmail)
			{
				$OwnerEmail = "NULL"
			}
			
			
			if (!$OwnerAlias)
			{
				$OwnerAlias = "NULL"
			}
			
			
			
			if ($grpdescription)
			{
				$grpdescription = $grpdescription -replace '(`r`n|\r\n|`n|`t)', ' '
			}
			
			if ($grpDisplay)
			{
				$grpDisplay = $grpDisplay -replace '(`r`n|\r\n|`n|`t)', ' '
			}
			
			if (!$grpdescription)
			{
				$descript = "NULL"
			}
			
			
			
			if (! $OwnerDisp)
			{
				
				$OwnerDisp = "NULL"
				
			}
			
			if (! $OwnerEmail)
			{
				
				$OwnerEmail = "NULL"
				
			}
			
			if (! $OwnerAlias)
			{
				
				$OwnerAlias = "NULL"
				
			}
			
			
			if (! $HideAB)
			{
				$HideAB = "False"
			}
			
			
			
			#single sendto group for either groups or users			
			if (($grpDisplay -notmatch "System.Object*") -and ($null -ne $grpprimarysmtp) -and ($restrictionDLcount -le 1) -and ($restrictionusercount -le 1) -and ($countofacls -le 1))
			{
				[int]$GRPNo = $GRPNo + 1
				$ReportLine1 = [PSCustomObject][ordered] @{
					
					
					Samaccountname			      = $grpSamAccountName
					mail						  = $grpprimarysmtp
					DisplayName				      = $grpDisplay
					DN						      = $grpDN
					RecipientTypeDetails		  = $grprecipientdetailsEX
					OwnerNTID					  = $OwnerAlias
					OwnerDisplayName			  = $OwnerDisp
					OwnerEmail				      = $OwnerEmail
					CoOwnerNTID				      = $CoOwnerNTID
					CoOwnerDisplayName		      = $CoOwnerDisplayName
					CoOwnerEmail				  = $CoOwnerEmail
					SendToUserNTID			      = $SendToUserNTID
					SendToUserDisplayName		  = $SendToUserDisplayName
					SendToUserEmail			      = $SendToUserEmail
					SendToDLNTID				  = $SendToDLNTID
					SendToDLDisplayName		      = $SendToDLDisplayName
					SendToDLEmail				  = $SendToDLEmail
					Alias						  = $grpalias
					Description				      = $grpdescription
					Location					  = $Locale
					AcceptFromExternal		      = $grpsenderauth
					HiddenFromAddressListsEnabled = $HideAB
				}
				
				
				
				
				$exportReportPaths = @(
					$exportreportgrpcsv1,
					$exportreportgrpcsv2,
					$exportreportgrpcsv3,
					$exportreportgrpcsv4,
					$exportreportgrpcsv5,
					$exportreportgrpcsv6,
					$exportreportgrpcsv7,
					$exportreportgrpcsv8,
					$exportreportgrpcsv9,
					$exportreportgrpcsv10,
					$exportreportgrpcsv11,
					$exportreportgrpcsv12,
					$exportreportgrpcsv13,
					$exportreportgrpcsv14,
					$exportreportgrpcsv15,
					$exportreportgrpcsv16,
					$exportreportgrpcsv17,
					$exportreportgrpcsv18,
					$exportreportgrpcsv19,
					$exportreportgrpcsv20,
					$exportreportgrpcsv21,
					$exportreportgrpcsv22,
					$exportreportgrpcsv23,
					$exportreportgrpcsv24,
					$exportreportgrpcsv25,
					$exportreportgrpcsv26,
					$exportreportgrpcsv27,
					$exportreportgrpcsv28,
					$exportreportgrpcsv29,
					$exportreportgrpcsv30,
					$exportreportgrpcsv31,
					$exportreportgrpcsv32,
					$exportreportgrpcsv33,
					$exportreportgrpcsv34,
					$exportreportgrpcsv35,
					$exportreportgrpcsv36,
					$exportreportgrpcsv37,
					$exportreportgrpcsv38,
					$exportreportgrpcsv39,
					$exportreportgrpcsv40,
					$exportreportgrpcsv41,
					$exportreportgrpcsv42,
					$exportreportgrpcsv43,
					$exportreportgrpcsv44,
					$exportreportgrpcsv45,
					$exportreportgrpcsv46,
					$exportreportgrpcsv47,
					$exportreportgrpcsv48,
					$exportreportgrpcsv49,
					$exportreportgrpcsv50,
					$exportreportgrpcsv51,
					$exportreportgrpcsv52,
					$exportreportgrpcsv53,
					$exportreportgrpcsv54,
					$exportreportgrpcsv55,
					$exportreportgrpcsv56,
					$exportreportgrpcsv57,
					$exportreportgrpcsv58,
					$exportreportgrpcsv59,
					$exportreportgrpcsv60,
					$exportreportgrpcsv61,
					$exportreportgrpcsv62,
					$exportreportgrpcsv63,
					$exportreportgrpcsv64,
					$exportreportgrpcsv65,
					$exportreportgrpcsv66,
					$exportreportgrpcsv67,
					$exportreportgrpcsv68,
					$exportreportgrpcsv69,
					$exportreportgrpcsv70
				)
				
				# Example usage
				Export-ReportLine -GRPNumber $GRPNo -reportLine $ReportLine1 -exportReportPaths $exportReportPaths
				
				#$ReportLine1 | Export-CSV $exportreportgrpcsv100 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
			}
			
			
			#multiple sendtorestrictions for groups
			If (($grpDisplay -notmatch "System.Object*") -and ($null -ne $grpprimarysmtp) -and ($restrictionDLcount -ge 2))
			{
				
				foreach ($restrictDL in $sendtodlgrp)
				{
					[int]$GRPNo = $GRPNo + 1
					
					$restrictionuserdetails = $null
					$getsendtodlgrp = $null
					$SendToUserNTID = "NULL"
					$SendToUserDisplayName = "NULL"
					$SendToUserEmail = "NULL"
					
					$SendToDLNTID = "NULL"
					$SendToDLDisplayName = "NULL"
					$SendToDLEmail = "NULL"
					
					$CoOwnerNTID = "NULL"
					$CoOwnerDisplayName = "NULL"
					$CoOwnerEmail = "NULL"
					
					#$getsendtodlgrp = $restrictDL.AcceptMessagesOnlyFromDLMembers
					$getsendtodlgrp = $restrictDL
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Getting restricted send to user $getsendtodluser who can send to the group $grpprimarysmtp "
					
					$restrictiongrpdetails = Get-group $getsendtodlgrp | Select-Object DisplayName, WindowsEmailAddress, RecipientTypeDetails, AccountDisabled, samaccountname
					
					
					
					$SendToDLNTID = $restrictiongrpdetails.samaccountname
					$SendToDLDisplayName = $restrictiongrpdetails.DisplayName
					$SendToDLEmail = $restrictiongrpdetails.WindowsEmailAddress
					
					if (($SendToDLNTID -eq $Null) -or ($SendToDLNTID -eq "")) { $SendToDLNTID = "NULL" }
					if (($SendToDLDisplayName -eq $Null) -or ($SendToDLDisplayName -eq "")) { $SendToDLDisplayName = "NULL" }
					if (($SendToDLEmail -eq $Null) -or ($SendToDLEmail -eq "")) { $SendToDLEmail = "NULL" }
					
					$ReportLine2 = [PSCustomObject][ordered] @{
						
						
						Samaccountname			      = $grpSamAccountName
						mail						  = $grpprimarysmtp
						DisplayName				      = $grpDisplay
						DN						      = $grpDN
						RecipientTypeDetails		  = $grprecipientdetailsEX
						OwnerNTID					  = $OwnerAlias
						OwnerDisplayName			  = $OwnerDisp
						OwnerEmail				      = $OwnerEmail
						CoOwnerNTID				      = $CoOwnerNTID
						CoOwnerDisplayName		      = $CoOwnerDisplayName
						CoOwnerEmail				  = $CoOwnerEmail
						SendToUserNTID			      = $SendToUserNTID
						SendToUserDisplayName		  = $SendToUserDisplayName
						SendToUserEmail			      = $SendToUserEmail
						SendToDLNTID				  = $SendToDLNTID
						SendToDLDisplayName		      = $SendToDLDisplayName
						SendToDLEmail				  = $SendToDLEmail
						Alias						  = $grpalias
						Description				      = $grpdescription
						Location					  = $Locale
						AcceptFromExternal		      = $grpsenderauth
						HiddenFromAddressListsEnabled = $HideAB
					}
					
					
					
					
					$exportReportPaths = @(
						$exportreportgrpcsv1,
						$exportreportgrpcsv2,
						$exportreportgrpcsv3,
						$exportreportgrpcsv4,
						$exportreportgrpcsv5,
						$exportreportgrpcsv6,
						$exportreportgrpcsv7,
						$exportreportgrpcsv8,
						$exportreportgrpcsv9,
						$exportreportgrpcsv10,
						$exportreportgrpcsv11,
						$exportreportgrpcsv12,
						$exportreportgrpcsv13,
						$exportreportgrpcsv14,
						$exportreportgrpcsv15,
						$exportreportgrpcsv16,
						$exportreportgrpcsv17,
						$exportreportgrpcsv18,
						$exportreportgrpcsv19,
						$exportreportgrpcsv20,
						$exportreportgrpcsv21,
						$exportreportgrpcsv22,
						$exportreportgrpcsv23,
						$exportreportgrpcsv24,
						$exportreportgrpcsv25,
						$exportreportgrpcsv26,
						$exportreportgrpcsv27,
						$exportreportgrpcsv28,
						$exportreportgrpcsv29,
						$exportreportgrpcsv30,
						$exportreportgrpcsv31,
						$exportreportgrpcsv32,
						$exportreportgrpcsv33,
						$exportreportgrpcsv34,
						$exportreportgrpcsv35,
						$exportreportgrpcsv36,
						$exportreportgrpcsv37,
						$exportreportgrpcsv38,
						$exportreportgrpcsv39,
						$exportreportgrpcsv40,
						$exportreportgrpcsv41,
						$exportreportgrpcsv42,
						$exportreportgrpcsv43,
						$exportreportgrpcsv44,
						$exportreportgrpcsv45,
						$exportreportgrpcsv46,
						$exportreportgrpcsv47,
						$exportreportgrpcsv48,
						$exportreportgrpcsv49,
						$exportreportgrpcsv50,
						$exportreportgrpcsv51,
						$exportreportgrpcsv52,
						$exportreportgrpcsv53,
						$exportreportgrpcsv54,
						$exportreportgrpcsv55,
						$exportreportgrpcsv56,
						$exportreportgrpcsv57,
						$exportreportgrpcsv58,
						$exportreportgrpcsv59,
						$exportreportgrpcsv60,
						$exportreportgrpcsv61,
						$exportreportgrpcsv62,
						$exportreportgrpcsv63,
						$exportreportgrpcsv64,
						$exportreportgrpcsv65,
						$exportreportgrpcsv66,
						$exportreportgrpcsv67,
						$exportreportgrpcsv68,
						$exportreportgrpcsv69,
						$exportreportgrpcsv70
					)
					
					# Example usage
					Export-ReportLine -GRPNumber $GRPNo -reportLine $ReportLine2 -exportReportPaths $exportReportPaths
					#$ReportLine2 | Export-CSV $exportreportgrpcsv100 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
				}
			}
			
			#multiple sendtorestrictions for users
			If (($grpDisplay -notmatch "System.Object*") -and ($null -ne $grpprimarysmtp) -and ($restrictionusercount -ge 2))
			{
				foreach ($restrictuser in $sendtodluser)
				{
					$getsendtodluser = $Null
					[int]$GRPNo = $GRPNo + 1
					$SendToDLNTID = "NULL"
					$SendToDLDisplayName = "NULL"
					$SendToDLEmail = "NULL"
					$SendToUserNTID = "NULL"
					$SendToUserDisplayName = "NULL"
					$SendToUserEmail = "NULL"
					$CoOwnerNTID = "NULL"
					$CoOwnerDisplayName = "NULL"
					$CoOwnerEmail = "NULL"
					
					
					$getsendtodluser = $restrictuser
					if ($getsendtodluser -notmatch "^bp1\.ad\.bp\.com/Deletion/Deletions Pending Users/.*")
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Getting restricted send to user $getsendtodluser who can send to the group $grpprimarysmtp "
						
						$restrictionuserdetails = Get-User $getsendtodluser | Select-Object DisplayName, WindowsEmailAddress, RecipientTypeDetails, AccountDisabled, samaccountname
						
						
						$SendToUserNTID = $restrictionuserdetails.samaccountname
						$SendToUserDisplayName = $restrictionuserdetails.DisplayName
						$SendToUserEmail = $restrictionuserdetails.WindowsEmailAddress
						
						
						
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Adding user restriction to send to DL for $SendToUserEmail to group  $grpprimarysmtp | $now "
						
						$ReportLine3 = [PSCustomObject][ordered] @{
							
							
							Samaccountname			      = $grpSamAccountName
							mail						  = $grpprimarysmtp
							DisplayName				      = $grpDisplay
							DN						      = $grpDN
							RecipientTypeDetails		  = $grprecipientdetailsEX
							OwnerNTID					  = $OwnerAlias
							OwnerDisplayName			  = $OwnerDisp
							OwnerEmail				      = $OwnerEmail
							CoOwnerNTID				      = $CoOwnerNTID
							CoOwnerDisplayName		      = $CoOwnerDisplayName
							CoOwnerEmail				  = $CoOwnerEmail
							SendToUserNTID			      = $SendToUserNTID
							SendToUserDisplayName		  = $SendToUserDisplayName
							SendToUserEmail			      = $SendToUserEmail
							SendToDLNTID				  = $SendToDLNTID
							SendToDLDisplayName		      = $SendToDLDisplayName
							SendToDLEmail				  = $SendToDLEmail
							Alias						  = $grpalias
							Description				      = $grpdescription
							Location					  = $Locale
							AcceptFromExternal		      = $grpsenderauth
							HiddenFromAddressListsEnabled = $HideAB
						}
						
						
						
						
						$exportReportPaths = @(
							$exportreportgrpcsv1,
							$exportreportgrpcsv2,
							$exportreportgrpcsv3,
							$exportreportgrpcsv4,
							$exportreportgrpcsv5,
							$exportreportgrpcsv6,
							$exportreportgrpcsv7,
							$exportreportgrpcsv8,
							$exportreportgrpcsv9,
							$exportreportgrpcsv10,
							$exportreportgrpcsv11,
							$exportreportgrpcsv12,
							$exportreportgrpcsv13,
							$exportreportgrpcsv14,
							$exportreportgrpcsv15,
							$exportreportgrpcsv16,
							$exportreportgrpcsv17,
							$exportreportgrpcsv18,
							$exportreportgrpcsv19,
							$exportreportgrpcsv20,
							$exportreportgrpcsv21,
							$exportreportgrpcsv22,
							$exportreportgrpcsv23,
							$exportreportgrpcsv24,
							$exportreportgrpcsv25,
							$exportreportgrpcsv26,
							$exportreportgrpcsv27,
							$exportreportgrpcsv28,
							$exportreportgrpcsv29,
							$exportreportgrpcsv30,
							$exportreportgrpcsv31,
							$exportreportgrpcsv32,
							$exportreportgrpcsv33,
							$exportreportgrpcsv34,
							$exportreportgrpcsv35,
							$exportreportgrpcsv36,
							$exportreportgrpcsv37,
							$exportreportgrpcsv38,
							$exportreportgrpcsv39,
							$exportreportgrpcsv40,
							$exportreportgrpcsv41,
							$exportreportgrpcsv42,
							$exportreportgrpcsv43,
							$exportreportgrpcsv44,
							$exportreportgrpcsv45,
							$exportreportgrpcsv46,
							$exportreportgrpcsv47,
							$exportreportgrpcsv48,
							$exportreportgrpcsv49,
							$exportreportgrpcsv50,
							$exportreportgrpcsv51,
							$exportreportgrpcsv52,
							$exportreportgrpcsv53,
							$exportreportgrpcsv54,
							$exportreportgrpcsv55,
							$exportreportgrpcsv56,
							$exportreportgrpcsv57,
							$exportreportgrpcsv58,
							$exportreportgrpcsv59,
							$exportreportgrpcsv60,
							$exportreportgrpcsv61,
							$exportreportgrpcsv62,
							$exportreportgrpcsv63,
							$exportreportgrpcsv64,
							$exportreportgrpcsv65,
							$exportreportgrpcsv66,
							$exportreportgrpcsv67,
							$exportreportgrpcsv68,
							$exportreportgrpcsv69,
							$exportreportgrpcsv70
						)
						
						# Example usage
						Export-ReportLine -GRPNumber $GRPNo -reportLine $ReportLine3 -exportReportPaths $exportReportPaths
					}
					#	$ReportLine3 | Export-CSV $exportreportgrpcsv100 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
				}
			}
			
			#multiple coowners
			if (($grpDisplay -notmatch "System.Object*") -and ($null -ne $grpprimarysmtp) -and ($countofacls -ge 2))
			{
				
				
				
				foreach ($bp1userACL in $outdlACL)
				{
					[int]$GRPNo = $GRPNo + 1
					
					
					$SendToUserNTID = "NULL"
					$SendToUserDisplayName = "NULL"
					$SendToUserEmail = "NULL"
					
					$SendToDLNTID = "NULL"
					$SendToDLDisplayName = "NULL"
					$SendToDLEmail = "NULL"
					
					$CoOwnerNTID = "NULL"
					$CoOwnerDisplayName = "NULL"
					$CoOwnerEmail = "NULL"
					
					$aclfrombp1 = $bp1userACL.IdentityReference
					$aclownersamaccount = ($aclfrombp1 -split '\\')[1]
					
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					add-content $logfile  "Getting $aclfrombp1 for CoOwner information | $now "
					
					$aclcoowner = Get-User $aclownersamaccount | Select-Object DisplayName, WindowsEmailAddress, RecipientTypeDetails, AccountDisabled, samaccountname
					
					$CoOwnerNTID = $aclcoowner.samaccountname
					$CoOwnerDisplayName = $aclcoowner.DisplayName
					$CoOwnerEmail = $aclcoowner.WindowsEmailAddress
					
					if (!$CoOwnerEmail) { $CoOwnerEmail = "NULL" }
					if (!$CoOwnerDisplayName) { $CoOwnerDisplayName = "NULL" }
					if (!$CoOwnerEmail) { $CoOwnerEmail = "NULL" }
					
					if ($CoOwnerNTID -eq $OwnerAlias)
					{
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "CoOwner and Owner match so skipping output Owner: $OwnerAlias CoOwner: $CoOwnerNTID  | $now "
						
					}
					
					if ($CoOwnerNTID -ne $OwnerAlias)
					{
						
						$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
						add-content $logfile  "Adding coowner for Group  $grpprimarysmtp CoOwner : $CoOwnerEmail  | $now "
						
						
						$ReportLine1 = [PSCustomObject][ordered] @{
							
							
							Samaccountname			      = $grpSamAccountName
							mail						  = $grpprimarysmtp
							DisplayName				      = $grpDisplay
							DN						      = $grpDN
							RecipientTypeDetails		  = $grprecipientdetailsEX
							OwnerNTID					  = $OwnerAlias
							OwnerDisplayName			  = $OwnerDisp
							OwnerEmail				      = $OwnerEmail
							CoOwnerNTID				      = $CoOwnerNTID
							CoOwnerDisplayName		      = $CoOwnerDisplayName
							CoOwnerEmail				  = $CoOwnerEmail
							SendToUserNTID			      = $SendToUserNTID
							SendToUserDisplayName		  = $SendToUserDisplayName
							SendToUserEmail			      = $SendToUserEmail
							SendToDLNTID				  = $SendToDLNTID
							SendToDLDisplayName		      = $SendToDLDisplayName
							SendToDLEmail				  = $SendToDLEmail
							Alias						  = $grpalias
							Description				      = $grpdescription
							Location					  = $Locale
							AcceptFromExternal		      = $grpsenderauth
							HiddenFromAddressListsEnabled = $HideAB
						}
						
						
						
						
						$exportReportPaths = @(
							$exportreportgrpcsv1,
							$exportreportgrpcsv2,
							$exportreportgrpcsv3,
							$exportreportgrpcsv4,
							$exportreportgrpcsv5,
							$exportreportgrpcsv6,
							$exportreportgrpcsv7,
							$exportreportgrpcsv8,
							$exportreportgrpcsv9,
							$exportreportgrpcsv10,
							$exportreportgrpcsv11,
							$exportreportgrpcsv12,
							$exportreportgrpcsv13,
							$exportreportgrpcsv14,
							$exportreportgrpcsv15,
							$exportreportgrpcsv16,
							$exportreportgrpcsv17,
							$exportreportgrpcsv18,
							$exportreportgrpcsv19,
							$exportreportgrpcsv20,
							$exportreportgrpcsv21,
							$exportreportgrpcsv22,
							$exportreportgrpcsv23,
							$exportreportgrpcsv24,
							$exportreportgrpcsv25,
							$exportreportgrpcsv26,
							$exportreportgrpcsv27,
							$exportreportgrpcsv28,
							$exportreportgrpcsv29,
							$exportreportgrpcsv30,
							$exportreportgrpcsv31,
							$exportreportgrpcsv32,
							$exportreportgrpcsv33,
							$exportreportgrpcsv34,
							$exportreportgrpcsv35,
							$exportreportgrpcsv36,
							$exportreportgrpcsv37,
							$exportreportgrpcsv38,
							$exportreportgrpcsv39,
							$exportreportgrpcsv40,
							$exportreportgrpcsv41,
							$exportreportgrpcsv42,
							$exportreportgrpcsv43,
							$exportreportgrpcsv44,
							$exportreportgrpcsv45,
							$exportreportgrpcsv46,
							$exportreportgrpcsv47,
							$exportreportgrpcsv48,
							$exportreportgrpcsv49,
							$exportreportgrpcsv50,
							$exportreportgrpcsv51,
							$exportreportgrpcsv52,
							$exportreportgrpcsv53,
							$exportreportgrpcsv54,
							$exportreportgrpcsv55,
							$exportreportgrpcsv56,
							$exportreportgrpcsv57,
							$exportreportgrpcsv58,
							$exportreportgrpcsv59,
							$exportreportgrpcsv60,
							$exportreportgrpcsv61,
							$exportreportgrpcsv62,
							$exportreportgrpcsv63,
							$exportreportgrpcsv64,
							$exportreportgrpcsv65,
							$exportreportgrpcsv66,
							$exportreportgrpcsv67,
							$exportreportgrpcsv68,
							$exportreportgrpcsv69,
							$exportreportgrpcsv70
						)
						
						# Example usage
						Export-ReportLine -GRPNumber $GRPNo -reportLine $ReportLine1 -exportReportPaths $exportReportPaths
					}
					#$ReportLine4 | Export-CSV $exportreportgrpcsv100 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
				}
			}
			
		}
		
	}
	
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
		add-content $logfile  " Finished getting OnPremise ASecurity group details for the Temp GRPLookuptable moving to exports | $now "
		
	
	

DisconnectExchangeOnPrem
	
	$exportfiles = @(
		@{ Path = $exportreportgrpcsv1; BaseName = "EARL-TempGRPTable-OnPremSGgroups1" },
		@{ Path = $exportreportgrpcsv2; BaseName = "EARL-TempGRPTable-OnPremSGgroups2" },
		@{ Path = $exportreportgrpcsv3; BaseName = "EARL-TempGRPTable-OnPremSGgroups3" },
		@{ Path = $exportreportgrpcsv4; BaseName = "EARL-TempGRPTable-OnPremSGgroups4" },
		@{ Path = $exportreportgrpcsv5; BaseName = "EARL-TempGRPTable-OnPremSGgroups5" },
		@{ Path = $exportreportgrpcsv6; BaseName = "EARL-TempGRPTable-OnPremSGgroups6" },
		@{ Path = $exportreportgrpcsv7; BaseName = "EARL-TempGRPTable-OnPremSGgroups7" },
		@{ Path = $exportreportgrpcsv8; BaseName = "EARL-TempGRPTable-OnPremSGgroups8" },
		@{ Path = $exportreportgrpcsv9; BaseName = "EARL-TempGRPTable-OnPremSGgroups9" },
		@{ Path = $exportreportgrpcsv10; BaseName = "EARL-TempGRPTable-OnPremSGgroups10" },
		@{ Path = $exportreportgrpcsv11; BaseName = "EARL-TempGRPTable-OnPremSGgroups11" },
		@{ Path = $exportreportgrpcsv12; BaseName = "EARL-TempGRPTable-OnPremSGgroups12" },
		@{ Path = $exportreportgrpcsv13; BaseName = "EARL-TempGRPTable-OnPremSGgroups13" },
		@{ Path = $exportreportgrpcsv14; BaseName = "EARL-TempGRPTable-OnPremSGgroups14" },
		@{ Path = $exportreportgrpcsv15; BaseName = "EARL-TempGRPTable-OnPremSGgroups15" },
		@{ Path = $exportreportgrpcsv16; BaseName = "EARL-TempGRPTable-OnPremSGgroups16" },
		@{ Path = $exportreportgrpcsv17; BaseName = "EARL-TempGRPTable-OnPremSGgroups17" },
		@{ Path = $exportreportgrpcsv18; BaseName = "EARL-TempGRPTable-OnPremSGgroups18" },
		@{ Path = $exportreportgrpcsv19; BaseName = "EARL-TempGRPTable-OnPremSGgroups19" },
		@{ Path = $exportreportgrpcsv20; BaseName = "EARL-TempGRPTable-OnPremSGgroups20" },
		@{ Path = $exportreportgrpcsv21; BaseName = "EARL-TempGRPTable-OnPremSGgroups21" },
		@{ Path = $exportreportgrpcsv22; BaseName = "EARL-TempGRPTable-OnPremSGgroups22" },
		@{ Path = $exportreportgrpcsv23; BaseName = "EARL-TempGRPTable-OnPremSGgroups23" },
		@{ Path = $exportreportgrpcsv24; BaseName = "EARL-TempGRPTable-OnPremSGgroups24" },
		@{ Path = $exportreportgrpcsv25; BaseName = "EARL-TempGRPTable-OnPremSGgroups25" },
		@{ Path = $exportreportgrpcsv26; BaseName = "EARL-TempGRPTable-OnPremSGgroups26" },
		@{ Path = $exportreportgrpcsv27; BaseName = "EARL-TempGRPTable-OnPremSGgroups27" },
		@{ Path = $exportreportgrpcsv28; BaseName = "EARL-TempGRPTable-OnPremSGgroups28" },
		@{ Path = $exportreportgrpcsv29; BaseName = "EARL-TempGRPTable-OnPremSGgroups29" },
		@{ Path = $exportreportgrpcsv30; BaseName = "EARL-TempGRPTable-OnPremSGgroups30" },
		@{ Path = $exportreportgrpcsv31; BaseName = "EARL-TempGRPTable-OnPremSGgroups31" },
		@{ Path = $exportreportgrpcsv32; BaseName = "EARL-TempGRPTable-OnPremSGgroups32" },
		@{ Path = $exportreportgrpcsv33; BaseName = "EARL-TempGRPTable-OnPremSGgroups33" },
		@{ Path = $exportreportgrpcsv34; BaseName = "EARL-TempGRPTable-OnPremSGgroups34" },
		@{ Path = $exportreportgrpcsv35; BaseName = "EARL-TempGRPTable-OnPremSGgroups35" },
		@{ Path = $exportreportgrpcsv36; BaseName = "EARL-TempGRPTable-OnPremSGgroups36" },
		@{ Path = $exportreportgrpcsv37; BaseName = "EARL-TempGRPTable-OnPremSGgroups37" },
		@{ Path = $exportreportgrpcsv38; BaseName = "EARL-TempGRPTable-OnPremSGgroups38" },
		@{ Path = $exportreportgrpcsv39; BaseName = "EARL-TempGRPTable-OnPremSGgroups39" },
		@{ Path = $exportreportgrpcsv40; BaseName = "EARL-TempGRPTable-OnPremSGgroups40" },
		@{ Path = $exportreportgrpcsv41; BaseName = "EARL-TempGRPTable-OnPremSGgroups41" },
		@{ Path = $exportreportgrpcsv42; BaseName = "EARL-TempGRPTable-OnPremSGgroups42" },
		@{ Path = $exportreportgrpcsv43; BaseName = "EARL-TempGRPTable-OnPremSGgroups43" },
		@{ Path = $exportreportgrpcsv44; BaseName = "EARL-TempGRPTable-OnPremSGgroups44" },
		@{ Path = $exportreportgrpcsv45; BaseName = "EARL-TempGRPTable-OnPremSGgroups45" },
		@{ Path = $exportreportgrpcsv46; BaseName = "EARL-TempGRPTable-OnPremSGgroups46" },
		@{ Path = $exportreportgrpcsv47; BaseName = "EARL-TempGRPTable-OnPremSGgroups47" },
		@{ Path = $exportreportgrpcsv48; BaseName = "EARL-TempGRPTable-OnPremSGgroups48" },
		@{ Path = $exportreportgrpcsv49; BaseName = "EARL-TempGRPTable-OnPremSGgroups49" },
		@{ Path = $exportreportgrpcsv50; BaseName = "EARL-TempGRPTable-OnPremSGgroups50" },
		@{ Path = $exportreportgrpcsv51; BaseName = "EARL-TempGRPTable-OnPremSGgroups51" },
		@{ Path = $exportreportgrpcsv52; BaseName = "EARL-TempGRPTable-OnPremSGgroups52" },
		@{ Path = $exportreportgrpcsv53; BaseName = "EARL-TempGRPTable-OnPremSGgroups53" },
		@{ Path = $exportreportgrpcsv54; BaseName = "EARL-TempGRPTable-OnPremSGgroups54" },
		@{ Path = $exportreportgrpcsv55; BaseName = "EARL-TempGRPTable-OnPremSGgroups55" },
		@{ Path = $exportreportgrpcsv56; BaseName = "EARL-TempGRPTable-OnPremSGgroups56" },
		@{ Path = $exportreportgrpcsv57; BaseName = "EARL-TempGRPTable-OnPremSGgroups57" },
		@{ Path = $exportreportgrpcsv58; BaseName = "EARL-TempGRPTable-OnPremSGgroups58" },
		@{ Path = $exportreportgrpcsv59; BaseName = "EARL-TempGRPTable-OnPremSGgroups59" },
		@{ Path = $exportreportgrpcsv60; BaseName = "EARL-TempGRPTable-OnPremSGgroups60" },
		@{ Path = $exportreportgrpcsv61; BaseName = "EARL-TempGRPTable-OnPremSGgroups61" },
		@{ Path = $exportreportgrpcsv62; BaseName = "EARL-TempGRPTable-OnPremSGgroups62" },
		@{ Path = $exportreportgrpcsv63; BaseName = "EARL-TempGRPTable-OnPremSGgroups63" },
		@{ Path = $exportreportgrpcsv64; BaseName = "EARL-TempGRPTable-OnPremSGgroups64" },
		@{ Path = $exportreportgrpcsv65; BaseName = "EARL-TempGRPTable-OnPremSGgroups65" },
		@{ Path = $exportreportgrpcsv66; BaseName = "EARL-TempGRPTable-OnPremSGgroups66" },
		@{ Path = $exportreportgrpcsv67; BaseName = "EARL-TempGRPTable-OnPremSGgroups67" },
		@{ Path = $exportreportgrpcsv68; BaseName = "EARL-TempGRPTable-OnPremSGgroups68" },
		@{ Path = $exportreportgrpcsv69; BaseName = "EARL-TempGRPTable-OnPremSGgroups69" },
		@{ Path = $exportreportgrpcsv70; BaseName = "EARL-TempGRPTable-OnPremSGgroups70" }
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
	
Try
{
	Map-Filewatcher
	Copy-item -path $Fileout -destination $filewatcherout
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	Add-Content $logfile "group File Copied to FileWatcher $Fileout to $filewatcherout [1st try] | $now"
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
		Add-Content $logfile "Groups File Copied to FileWatcher $Fileout to $filewatcherout [2nd try] | $now"
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

#exportDLOnPrem
exportSGOnPrem




$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "All finished with exports for LDAP replacement Temp GRPLookupTable | $now"
DisconnectExchangeOnPrem
RemoveFilewatcher

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "Closing down script - bye $now"
Stop-Transcript

Get-PSSession | Remove-PSSession
Exit-PSSession
Exit



