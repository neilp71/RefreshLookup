



<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.148
	 Created on:    07/08/2023 14:00
	 Created by:   	Neil Poultney
	 Organization: 	NDP Consultancy Ltd - neil.poultney@ndpconsultancyltd.co.uk
	 Filename:     	Refesh-EARL-get-lookupTable-Export-priv.ps1
	===========================================================================
	.DESCRIPTION
		Exports for EARL MailDb to Lookup Table for priv users to Temp db.

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
$transcriptlog = "H:\EARLTranscripts\LookupTbl\Refresh-lookup-export-priv-" + $nowfiledate + ".log"

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
	$logfilelocation = $loglocation + "Refresh-Lookup-Table-PrivUsers-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
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
	$loglocation = "H:\EARLPSLogs\BulkExports\" # change to usetype RPMBCREATE etc
	Set-Variable -Name logfolder -Value $logslocation -Option ReadOnly -Scope Script -Force
	$logfilelocation = $loglocation + "Refresh-Lookup-Table-PrivUsers-" + $nowfiledate + ".log" # change to usetype RPMBCREATE etc
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
	$exportreportcsv1 = "H:\M365Reports\EARL-Prod-TempLookupTable-priv-1-" + $nowfiledate + ".csv"
	#$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv2 = "H:\M365Reports\EARL-Prod-TempLookupTable-priv-2-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv3 = "H:\M365Reports\EARL-Prod-TempLookupTable-priv-3-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv4 = "H:\M365Reports\EARL-Prod-TempLookupTable-priv-4-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv5 = "H:\M365Reports\EARL-Prod-TempLookupTable-priv-5-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv6 = "H:\M365Reports\\EARL-Prod-TempLookupTable-priv-6-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv7 = "H:\M365Reports\EARL-Prod-TempLookupTable-priv-7-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv8 = "H:\M365Reports\EARL-Prod-TempLookupTable-priv-8-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv9 = "H:\M365Reports\EARL-Prod-TempLookupTable-priv-9-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv10 = "H:\M365Reports\EARL-Prod-TempLookupTable-priv-10-" + $nowfiledate + ".csv"
	
	#temp
	$lookuptime1 = (get-date).adddays(-720)
	#$lookuptime = (get-date).addhours(-3)
	Set-Variable -Name lasthour1 -Value $lookuptime1 -Option ReadOnly -Scope Script -Force
	
	#$Mbx5 = Get-MailContact -ResultSize unlimited | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	#$Mbx5 = Get-MailContact -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	$accounts = Get-User -filter "samaccountname -like 'svc*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like '-svc*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like 'serv-*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like '-serv-*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like '-tsk*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like 'task*'-and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like 'tsk*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like '-task-*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like '-gbl*'" -resultsize unlimited | Select-Object *
	
	#"samaccountname -like 'svc*' -and WhenChanged -gt '$lookuptime1' "
	
	
	
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
		[int]$MbxNumber = 0
		
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
				
				
				if ((! $Display) -or ($Display -eq ""))
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
					add-content $logfile  " Priv User number : $MbxNumber | skipped because UserAccountcontrol is $UsrACCCtrl  | $now "
				}
				
				If ($UsrACCCtrl -eq "66050")
				{
					$skipped = "YES"
					$skippedpriv = $skippedpriv + 1
					add-content $logfile  " Priv User number : $MbxNumber | skipped because UserAccountcontrol is $UsrACCCtrl  | $now "
				}
				
				
				If ($NTID -like "*-dou")
					{
					$skipped = "YES"
					add-content $logfile  " Priv User number : $MbxNumber | skipped because UserAccountcontrol is $UsrACCCtrl  | $now "
					
				}
				
				
				
				If (($skipped -eq "NO") -and ($NTID))
				{
					
					$addedpriv = $addedpriv + 1
					
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
					
					
					if ($MbxNumber -le 2000)
					{
						$ReportLine4 | Export-CSV $exportreportcsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -ge 2000) -and ($MbxNumber -le 4000))
					{
						$ReportLine4 | Export-CSV $exportreportcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -ge 4000) -and ($MbxNumber -le 6000))
					{
						$ReportLine4 | Export-CSV $exportreportcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -ge 6000) -and ($MbxNumber -le 8000))
					{
						$ReportLine4 | Export-CSV $exportreportcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -ge 8000) -and ($MbxNumber -le 10000))
					{
						$ReportLine4 | Export-CSV $exportreportcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -ge 10000) -and ($MbxNumber -le 12000))
					{
						$ReportLine4 | Export-CSV $exportreportcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -ge 12000) -and ($MbxNumber -le 14000))
					{
						$ReportLine4 | Export-CSV $exportreportcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					if (($MbxNumber -ge 14000) -and ($MbxNumber -le 16000))
					{
						$ReportLine4 | Export-CSV $exportreportcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -ge 16000) -and ($MbxNumber -le 18000))
					{
						$ReportLine4 | Export-CSV $exportreportcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if (($MbxNumber -ge 18000) -and ($MbxNumber -le 20000))
					{
						$ReportLine4 | Export-CSV $exportreportcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					if ($MbxNumber -ge 20000)
					{
						$ReportLine4 | Export-CSV $exportreportcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
					}
					
					
				}
			}
			
		}
		
		
		
		
		
		DisconnectExchangeOnPrem
		
		
		Add-Content $logfile "Skipped $skippedpriv Priv users as they are disabled with no mail for import"
		Add-Content $logfile "Imported $addedpriv Priv Users as they are enabled and no mail currently in place"
		
		#sortoutput so no blank lines and no duplicates
		$inputCsv1 = Import-Csv $exportreportcsv1 -delimiter "|" | Sort-Object * -Unique
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$finaloutcsv1 = "H:\M365Reports\EARL-LookupTbl-privuser-1-" + $nowfiledate + "-1.csv"
		$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$Outfile1 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-1-" + $nowfiledate + "-2.csv"
		gc $finaloutcsv1 | ? { $_.trim() -ne "" } | set-content $Outfile1 -Encoding UTF8
		
		######
		
		Start-Sleep -s 2
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
		$inputCsv2 = Import-Csv $exportreportcsv2 -delimiter "|" | Sort-Object * -Unique
		$finaloutcsv2 = "H:\M365Reports\EARL-LookupTbl-privuser-2-" + $nowfiledate + "-1.csv"
		$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$Outfile2 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-2-" + $nowfiledate + "-2.csv"
		gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Outfile2 -Encoding UTF8
		
		Start-Sleep -s 2
		
		$inputCsv3 = Import-Csv $exportreportcsv3 -delimiter "|" | Sort-Object * -Unique
		$finaloutcsv3 = "H:\M365Reports\EARL-LookupTbl-privuser-3-" + $nowfiledate + "-1.csv"
		$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
		Start-Sleep -s 5
		$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
		$Outfile3 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-3-" + $nowfiledate + "-2.csv"
		gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Outfile3 -Encoding UTF8
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv4
		if ($checkfile -eq "True")
		{
			$inputCsv4 = Import-Csv $exportreportcsv4 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv4 = "H:\M365Reports\EARL-LookupTbl-privuser-4-" + $nowfiledate + "-1.csv"
			$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile4 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-4-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Outfile4 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv5
		if ($checkfile -eq "True")
		{
			$inputCsv5 = Import-Csv $exportreportcsv5 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv5 = "H:\M365Reports\EARL-LookupTbl-privuser-5-" + $nowfiledate + "-1.csv"
			$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile5 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-5-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Outfile5 -Encoding UTF8
			
		}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv6
		if ($checkfile -eq "True")
		{
			$inputCsv6 = Import-Csv $exportreportcsv6 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv6 = "H:\M365Reports\EARL-LookupTbl-privuser-6-" + $nowfiledate + "-1.csv"
			$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile6 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-6-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Outfile6 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv7
		if ($checkfile -eq "True")
		{
			$inputCsv7 = Import-Csv $exportreportcsv7 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv7 = "H:\M365Reports\EARL-LookupTbl-privuser-7-" + $nowfiledate + "-1.csv"
			$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile7 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-7-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Outfile7 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv8
		if ($checkfile -eq "True")
		{
			$inputCsv8 = Import-Csv $exportreportcsv8 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv8 = "H:\M365Reports\EARL-LookupTbl-privuser-8-" + $nowfiledate + "-1.csv"
			$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile8 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-8-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Outfile8 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv9
		if ($checkfile -eq "True")
		{
			$inputCsv9 = Import-Csv $exportreportcsv9 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv9 = "H:\M365Reports\EARL-LookupTbl-privuser-9-" + $nowfiledate + "-1.csv"
			$inputCsv9 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile9 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-9-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv9 | ? { $_.trim() -ne "" } | set-content $Outfile9 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv10
		if ($checkfile -eq "True")
		{
			$inputCsv10 = Import-Csv $exportreportcsv10 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv10 = "H:\M365Reports\EARL-LookupTbl-privuser-10-" + $nowfiledate + "-1.csv"
			$inputCsv10 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile10 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privuser-10-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv10 | ? { $_.trim() -ne "" } | set-content $Outfile10 -Encoding UTF8
			
		}
		
		
		######
		
		Map-Filewatcher
		
		if (Test-Path $Outfile1)
		{
			Try
			{
				
				Copy-item -path $Outfile1 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile1 to $filewatcherout [1st try] | $now"
				
				
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
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile1 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile1 | $now"
				}
			}
		}
		
		if (Test-Path $Outfile2)
		{
			
			
			Try
			{
				Start-Sleep -Seconds 60
				Copy-item -path $Outfile2 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile2 to $filewatcherout [1st try] | $now"
				
				
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
					Copy-item -path $Outfile2 -destination $filewatcherout
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile2 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile2 | $now"
				}
			}
			
		}
		
		if (Test-Path $Outfile3)
		{
			Try
			{
				Start-Sleep -Seconds 60
				Copy-item -path $Outfile3 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile3 to $filewatcherout [1st try] | $now"
				
				
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
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile3 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile3 | $now"
				}
			}
		}
		
		if (Test-Path $Outfile4)
		{
			Try
			{
				Start-Sleep -Seconds 60
				Copy-item -path $Outfile4 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile4 to $filewatcherout [1st try] | $now"
				
				
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
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile4 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile4 | $now"
				}
			}
		}
		
		if (Test-Path $Outfile5)
		{
			
			Try
			{
				Start-Sleep -Seconds 60
				Copy-item -path $Outfile5 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile5 to $filewatcherout [1st try] | $now"
				
				
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
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile5 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile5 | $now"
				}
			}
		}
		
		if (Test-Path $Outfile6)
		{
			
			Try
			{
				Start-Sleep -Seconds 60
				Copy-item -path $Outfile6 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile6 to $filewatcherout [1st try] | $now"
				
				
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
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile6 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile6 | $now"
				}
			}
		}
		
		if (Test-Path $Outfile7)
		{
			
			Try
			{
				Start-Sleep -Seconds 60
				Copy-item -path $Outfile7 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile7 to $filewatcherout [1st try] | $now"
				
				
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
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile7 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile7 | $now"
				}
			}
			
		}
		
		if (Test-Path $Outfile8)
		{
			
			Try
			{
				Start-Sleep -Seconds 60
				Copy-item -path $Outfile8 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile8 to $filewatcherout [1st try] | $now"
				
				
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
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile8 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile8 | $now"
				}
			}
			
		}
		
		if (Test-Path $Outfile9)
		{
			
			Try
			{
				Start-Sleep -Seconds 60
				Copy-item -path $Outfile9 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile9 to $filewatcherout [1st try] | $now"
				
				
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
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile9 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile9 | $now"
				}
			}
		}
		
		if (Test-Path $Outfile10)
		{
			
			##
			
			Try
			{
				Start-Sleep -Seconds 60
				Copy-item -path $Outfile10 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile10 to $filewatcherout [1st try] | $now"
				
				
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
					Add-Content $logfile "Priv Users File Copied to FileWatcher $Outfile10 to $filewatcherout [2nd try] | $now"
					
					#RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Cannot copy file for priv users to FileWatcher $Outfile10 | $now"
				}
			}
		}
		
		#>
		#cleanup files
		if (Test-Path $exportreportcsv1)
		{
			Remove-Item $exportreportcsv1
		}
		
		if (Test-Path $exportreportcsv2)
		{
			Remove-Item $exportreportcsv2
		}
		
		if (Test-Path $exportreportcsv3)
		{
			Remove-Item $exportreportcsv3
		}
		
		if (Test-Path $exportreportcsv4)
		{
			Remove-Item $exportreportcsv4
		}
		
		if (Test-Path $exportreportcsv5)
		{
			Remove-Item $exportreportcsv5
		}
		
		if (Test-Path $exportreportcsv6)
		{
			Remove-Item $exportreportcsv6
		}
		
		if (Test-Path $exportreportcsv7)
		{
			Remove-Item $exportreportcsv7
		}
		
		if (Test-Path $exportreportcsv8)
		{
			Remove-Item $exportreportcsv8
		}
		
		if (Test-Path $exportreportcsv9)
		{
			Remove-Item $exportreportcsv9
		}
		
		if (Test-Path $exportreportcsv10)
		{
			Remove-Item $exportreportcsv10
		}
		
	
		
		if (Test-Path $finaloutcsv1)
		{
			Remove-Item $finaloutcsv1
		}
		
		if (Test-Path $finaloutcsv2)
		{
			Remove-Item $finaloutcsv2
		}
		
		if (Test-Path $finaloutcsv3)
		{
			Remove-Item $finaloutcsv3
		}
		
		if (Test-Path $finaloutcsv4)
		{
			Remove-Item $finaloutcsv4
		}
		
		if (Test-Path $finaloutcsv5)
		{
			Remove-Item $finaloutcsv5
		}
		
		if (Test-Path $finaloutcsv6)
		{
			Remove-Item $finaloutcsv6
		}
		
		if (Test-Path $finaloutcsv7)
		{
			Remove-Item $finaloutcsv7
		}
		
		if (Test-Path $finaloutcsv8)
		{
			Remove-Item $finaloutcsv8
		}
		
	}
	
	
}

function exportprivsysops
{
	
	
	ConnectExchangeonPrem
	
	
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Running Function to export Privileged sysops accounts for export"
	
	$Mbx5 = ""
	$exportreportcsv5 = ""
	$M = ""
	$reconnect = 0
	$accounts = ""
	Start-Sleep -s 5
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv1 = "H:\M365Reports\EARL-Prod-TempLookupTable-privsysop-1-" + $nowfiledate + ".csv"
	#$Mbx3 = Get-Mailuser -ResultSize unlimited | Select DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv2 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-2-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv3 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-3-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv4 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-4-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv5 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-5-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv6 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-6-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv7 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-7-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv8 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-8-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv9 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-9-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv10 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-10-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv11 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-11-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv12 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-12-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv13 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-13-" + $nowfiledate + ".csv"
	
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv14 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-14-" + $nowfiledate + ".csv"
	Start-Sleep -s 2
	$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
	$exportreportcsv15 = "H:\M365Reports\\EARL-TempLookupTable-privsysop-15-" + $nowfiledate + ".csv"
	
	#temp
	$lookuptime1 = (get-date).adddays(-900)
	$lookuptime = (get-date).addhours(-3)
	Set-Variable -Name lasthour1 -Value $lookuptime1 -Option ReadOnly -Scope Script -Force
	Set-Variable -Name lasthour -Value $lookuptime -Option ReadOnly -Scope Script -Force
	
	#$Mbx5 = Get-MailContact -ResultSize unlimited | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	#$Mbx5 = Get-MailContact -ResultSize unlimited -filter "WhenChanged -gt '$lasthour'" | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress, whenchanged, whencreated
	#$accounts = Get-User -filter "samaccountname -like 'svc*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like '-svc*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like 'serv-*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like '-serv-*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like '-tsk*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like 'task*'-and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like 'tsk*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like '-task-*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts = Get-User -filter "samaccountname -like '-sysop-*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like 'sysop-*' -and WhenChanged -gt '$lasthour'" -resultsize unlimited | Select-Object *
	#$accounts = Get-User -filter "samaccountname -like '-sysop-*' -and WhenChanged -gt '$lookuptime1'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like 'sysop-*' -and WhenChanged -gt '$lookuptime1'" -resultsize unlimited | Select-Object *
	$accounts = Get-User -filter "samaccountname -like '-sysop-*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	$accounts += Get-User -filter "samaccountname -like 'sysop-*' -and WhenChanged -gt '$lasthour1'" -resultsize unlimited | Select-Object *
	#$accounts += Get-User -filter "samaccountname -like '-gbl*'" -resultsize unlimited | Select-Object *
	
	#"samaccountname -like 'svc*' -and WhenChanged -gt '$lookuptime1' "
	
	
	
	$count = $accounts.count
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Found: $count Privileged sysops accounts to process"
	
	add-content $logfile  "LookupTable Exporting to files for Privileged accounts - sysops "
	
	
	
	[regex]$pattern1 = '-sysop-......-z$'
	[regex]$pattern2 = 'sysop-......-z$'
	
	
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
		[int]$MbxNumber = 0
		
		$sysopzusr = $account.samaccountname
		
		
		
		ForEach ($account in $accounts)
		{
			
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
			$usr = $account.samaccountname
			$Disp = $account.DisplayName
			$recipientdetailsEX = $account.RecipientTypeDetails
			
			
			if (($usr -match $pattern1) -or ($usr -match $pattern2))
			{
				
				$MbxNumber++
				
				
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
					
					if ($descript)
					{
						$descript.replace('`r`n', ' ')
					}
					
					
					if ($descript)
					{
						$descript.replace('\r\n', ' ')
					}
					
					if ($descript)
					{
						$descript.replace('`n', ' ')
					}
					
					if ($descript)
					{
						$descript.replace('`t', ' ')
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
					
					
					
					If (($skipped -eq "NO") -and ($NTID))
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
							Description			       = $descript
							
						}
						
						
						if ($MbxNumber -le 1000)
						{
							$ReportLine4 | Export-CSV $exportreportcsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 1000) -and ($MbxNumber -le 2000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 2000) -and ($MbxNumber -le 3000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 3000) -and ($MbxNumber -le 4000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 4000) -and ($MbxNumber -le 5000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 5000) -and ($MbxNumber -le 6000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						
						if (($MbxNumber -ge 6000) -and ($MbxNumber -le 7000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						
						if (($MbxNumber -ge 7000) -and ($MbxNumber -le 8000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 8000) -and ($MbxNumber -le 9000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 9000) -and ($MbxNumber -le 10000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						
						if (($MbxNumber -ge 10000) -and ($MbxNumber -le 11000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 11000) -and ($MbxNumber -le 12000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 12000) -and ($MbxNumber -le 13000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if (($MbxNumber -ge 13000) -and ($MbxNumber -le 14000))
						{
							$ReportLine4 | Export-CSV $exportreportcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						if ($MbxNumber -ge 14000)
						{
							$ReportLine4 | Export-CSV $exportreportcsv15 -NoTypeInformation -delimiter "|" -Encoding UTF8 -Append -force
						}
						
						
					}
				}
			}
		}
		
		
		
		
		
		DisconnectExchangeOnPrem
		
		
		Add-Content $logfile "Skipped $skippedpriv Priv users as they are disabled with no mail for import"
		Add-Content $logfile "Imported $addedpriv Priv Users as they are enabled and no mail currently in place"
		
		#sortoutput so no blank lines and no duplicates
		
		$checkfile = Test-Path -Path $exportreportcsv1
		if ($checkfile -eq "True")
		{
			$inputCsv1 = Import-Csv $exportreportcsv1 -delimiter "|" | Sort-Object * -Unique
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$finaloutcsv1 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-1-" + $nowfiledate + "-1.csv"
			$inputCsv1 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $True } | Export-Csv $finaloutcsv1 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile1 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-1-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv1 | ? { $_.trim() -ne "" } | set-content $Outfile1 -Encoding UTF8
			
		}
		######
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv2
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm"
			$inputCsv2 = Import-Csv $exportreportcsv2 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv2 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-2-" + $nowfiledate + "-1.csv"
			$inputCsv2 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv2 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile2 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-2-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv2 | ? { $_.trim() -ne "" } | set-content $Outfile2 -Encoding UTF8
			
		}
		Start-Sleep -s 2
		
		$checkfile = Test-Path -Path $exportreportcsv3
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv3 = Import-Csv $exportreportcsv3 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv3 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-3-" + $nowfiledate + "-1.csv"
			$inputCsv3 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv3 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			
			$Outfile3 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-3-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv3 | ? { $_.trim() -ne "" } | set-content $Outfile3 -Encoding UTF8
			
		}
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv4
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv4 = Import-Csv $exportreportcsv4 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv4 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-4-" + $nowfiledate + "-1.csv"
			$inputCsv4 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv4 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile4 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-4-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv4 | ? { $_.trim() -ne "" } | set-content $Outfile4 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv5
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv5 = Import-Csv $exportreportcsv5 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv5 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-5-" + $nowfiledate + "-1.csv"
			$inputCsv5 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv5 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile5 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-5-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv5 | ? { $_.trim() -ne "" } | set-content $Outfile5 -Encoding UTF8
			
		}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv6
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv6 = Import-Csv $exportreportcsv6 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv6 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-6-" + $nowfiledate + "-1.csv"
			$inputCsv6 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv6 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile6 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-6-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv6 | ? { $_.trim() -ne "" } | set-content $Outfile6 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv7
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv7 = Import-Csv $exportreportcsv7 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv7 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-7-" + $nowfiledate + "-1.csv"
			$inputCsv7 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv7 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile7 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-7-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv7 | ? { $_.trim() -ne "" } | set-content $Outfile7 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv8
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv8 = Import-Csv $exportreportcsv8 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv8 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-8-" + $nowfiledate + "-1.csv"
			$inputCsv8 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv8 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile8 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-8-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv8 | ? { $_.trim() -ne "" } | set-content $Outfile8 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv9
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv9 = Import-Csv $exportreportcsv9 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv9 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-9-" + $nowfiledate + "-1.csv"
			$inputCsv9 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv9 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile9 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-9-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv9 | ? { $_.trim() -ne "" } | set-content $Outfile9 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv10
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv10 = Import-Csv $exportreportcsv10 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv10 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-10-" + $nowfiledate + "-1.csv"
			$inputCsv10 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv10 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile10 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-10-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv10 | ? { $_.trim() -ne "" } | set-content $Outfile10 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv11
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv11 = Import-Csv $exportreportcsv11 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv11 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-11-" + $nowfiledate + "-1.csv"
			$inputCsv11 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv11 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile11 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-11-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv11 | ? { $_.trim() -ne "" } | set-content $Outfile11 -Encoding UTF8
			
		}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv12
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv12 = Import-Csv $exportreportcsv12 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv12 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-12-" + $nowfiledate + "-1.csv"
			$inputCsv12 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv12 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile12 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-12-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv12 | ? { $_.trim() -ne "" } | set-content $Outfile12 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv13
		if ($checkfile -eq "True")
		{
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$inputCsv13 = Import-Csv $exportreportcsv13 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv13 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-13-" + $nowfiledate + "-1.csv"
			$inputCsv13 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv13 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile13 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-13-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv13 | ? { $_.trim() -ne "" } | set-content $Outfile13 -Encoding UTF8
			
		}
		
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv14
		if ($checkfile -eq "True")
		{
			$inputCsv14 = Import-Csv $exportreportcsv14 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv14 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-14-" + $nowfiledate + "-1.csv"
			$inputCsv14 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv14 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile14 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-14-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv14 | ? { $_.trim() -ne "" } | set-content $Outfile14 -Encoding UTF8
			
		}
		
		Start-Sleep -s 2
		$checkfile = Test-Path -Path $exportreportcsv15
		if ($checkfile -eq "True")
		{
			$inputCsv15 = Import-Csv $exportreportcsv15 -delimiter "|" | Sort-Object * -Unique
			$finaloutcsv15 = "H:\M365Reports\EARL-TempLookupTable-privusersysop-15-" + $nowfiledate + "-1.csv"
			$inputCsv15 | Sort -Property @{ Expression = { $_.Samaccountname }; Ascending = $False } | Export-Csv $finaloutcsv15 -NoTypeInformation -delimiter "|" -Encoding UTF8
			Start-Sleep -s 5
			$nowfiledate = get-date -f "yyyy-MM-dd-hh-mm-ss"
			$Outfile15 = "H:\M365Reports\TempLookupUploads\EARL-TempLookupTable-privusersysop-15-" + $nowfiledate + "-2.csv"
			gc $finaloutcsv15 | ? { $_.trim() -ne "" } | set-content $Outfile15 -Encoding UTF8
			
		}
		
		
		#copy to filewatcher
		
		Map-Filewatcher
		
		if (Test-Path $Outfile1)
		{
			Try
			{
				
				Copy-item -path $Outfile1 -destination $filewatcherout
				$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile1 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Fileout to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile2 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile2 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile3 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile3 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile4 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile4 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile5 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile5 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile6 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile6 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile7 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile7 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile8 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile8 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile9 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile9 to $filewatcherout [2nd try] | $now"
					Start-Sleep -Seconds 30
					
					RemoveFilewatcher
				}
				catch
				{
					$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					Add-Content $logfile "Priv sysop copy files to FileWatcher $Outfile9 | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile10 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile10 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "Priv sysop File Copied to FileWatcher $Outfile11 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile11 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile12 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile12 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile13 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile13 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile14 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile14 to $filewatcherout [2nd try] | $now"
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
				Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile15 to $filewatcherout [1st try] | $now"
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
					Add-Content $logfile "remotemailbox File Copied to FileWatcher $Outfile15 to $filewatcherout [2nd try] | $now"
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
		
		
		#cleanup files
		#cleanup files
		
		if (Test-Path $exportreportcsv1)
		{
			Remove-Item $exportreportcsv1
		}
		
		if (Test-Path $finaloutcsv1)
		{
			Remove-Item $finaloutcsv1
		}
		
		if (Test-Path $exportreportcsv2)
		{
			Remove-Item $exportreportcsv2
		}
		
		if (Test-Path $finaloutcsv2)
		{
			Remove-Item $finaloutcsv2
		}
		
		if (Test-Path $exportreportcsv3)
		{
			Remove-Item $exportreportcsv3
		}
		
		if (Test-Path $finaloutcsv3)
		{
			Remove-Item $finaloutcsv3
		}
		
		if (Test-Path $exportreportcsv4)
		{
			Remove-Item $exportreportcsv4
		}
		
		if (Test-Path $finaloutcsv4)
		{
			Remove-Item $finaloutcsv4
		}
		
		if (Test-Path $exportreportcsv5)
		{
			Remove-Item $exportreportcsv1
		}
		
		if (Test-Path $finaloutcsv5)
		{
			Remove-Item $finaloutcsv5
		}
		
		if (Test-Path $exportreportcsv6)
		{
			Remove-Item $exportreportcsv6
		}
		
		if (Test-Path $finaloutcsv6)
		{
			Remove-Item $finaloutcsv6
		}
		
		if (Test-Path $exportreportcsv7)
		{
			Remove-Item $exportreportcsv7
		}
		
		if (Test-Path $finaloutcsv7)
		{
			Remove-Item $finaloutcsv7
		}
		
		if (Test-Path $exportreportcsv8)
		{
			Remove-Item $exportreportcsv8
		}
		
		if (Test-Path $finaloutcsv8)
		{
			Remove-Item $finaloutcsv8
		}
		if (Test-Path $exportreportcsv9)
		{
			Remove-Item $exportreportcsv9
		}
		
		if (Test-Path $finaloutcsv9)
		{
			Remove-Item $finaloutcsv9
		}
		if (Test-Path $exportreportcsv10)
		{
			Remove-Item $exportreportcsv10
		}
		
		if (Test-Path $finaloutcsv10)
		{
			Remove-Item $finaloutcsv10
		}
		
		if (Test-Path $exportreportcsv11)
		{
			Remove-Item $exportreportcsv11
		}
		
		if (Test-Path $finaloutcsv11)
		{
			Remove-Item $finaloutcsv11
		}
		
		if (Test-Path $exportreportcsv12)
		{
			Remove-Item $exportreportcsv12
		}
		
		if (Test-Path $finaloutcsv12)
		{
			Remove-Item $finaloutcsv12
		}
		
		if (Test-Path $exportreportcsv13)
		{
			Remove-Item $exportreportcsv13
		}
		
		if (Test-Path $finaloutcsv13)
		{
			Remove-Item $finaloutcsv13
		}
		
		if (Test-Path $exportreportcsv14)
		{
			Remove-Item $exportreportcsv14
		}
		
		if (Test-Path $finaloutcsv14)
		{
			Remove-Item $finaloutcsv14
		}
		
		if (Test-Path $exportreportcsv15)
		{
			Remove-Item $exportreportcsv15
		}
		
		if (Test-Path $finaloutcsv15)
		{
			Remove-Item $finaloutcsv15
		}
		
		
	}
	
	DisconnectExchangeOnPrem
	$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
	add-content $logfile  "Finished Function to export sysops users (sysop-z) for export"
}



exportpriv
exportprivsysops





$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "All finished with exports for LDAP replacement LookupTable hourly for Priv Users | $now"
DisconnectExchangeOnPrem
RemoveFilewatcher

$now = [datetime]::Now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
Add-Content $logfile "Closing down script - bye $now"
Stop-Transcript

Disconnect-EXO
DisconnectExchangeOnPrem

Exit-PSSession
Exit



