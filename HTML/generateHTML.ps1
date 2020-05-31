#############################################################################################################################

#add exchange snapin
Add-PSSnapin *exchange* -ea 0

cd .\HTML
$rootfolder  = (Get-Item -Path "..\").FullName
$currentfolder  = (Get-Item -Path ".\").FullName

###############################################
function updatehistory{

	# Gives a list of all Microsoft Updates sorted by KB number/HotfixID
	# By Tom Arbuthnot. Lyncdup.com

	$wu = new-object -com "Microsoft.Update.Searcher"

	$totalupdates = $wu.GetTotalHistoryCount()
	$recent = (Get-Date).adddays(-7)
	$all = $wu.QueryHistory(0,$totalupdates) | where {$_.Date -gt $recent}
	
	# Define a new array to gather output
	$OutputCollection=  @()
			
    Foreach ($update in $all) {
        
		$string = $update.title
		$date = $update.Date

		$Regex = "KB\d*"
		$KB = $string | Select-String -Pattern $regex | Select-Object { $_.Matches }

		 $output = New-Object -TypeName PSobject
		 $output | add-member NoteProperty "HotFixID" -value $KB.' $_.Matches '.Value
		 $output | add-member NoteProperty "Date" -value $date
		 $output | add-member NoteProperty "Title" -value $string	
		 $OutputCollection += $output

	}

	$OutputCollection | Sort-Object HotFixID | Format-Table -AutoSize	
    $OutputCollection | Sort-Object HotFixID | Format-Table -AutoSize > "C:\PS_WINDOWS_UPDATE\LOG\installedUpdates.txt"
	$OutputCollection | Select HotFixID > "C:\PS_WINDOWS_UPDATE\LOG\KB.txt"

	#Write-Host "$($OutputCollection.Count) Updates Found"
    return $OutputCollection
}

function generateHTML {

	$Style =  @"
		<style>
			BODY{font-family: Segoe UI;font-size:12pt;text-align:center}
			H1{font-size: 16px;}
			H2{font-size: 14px;}
			H3{font-size: 12px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 10pt; width:1000px; }
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #FFFFFF; border-color: black;background-color:red}
			TD{border: 1px solid black; padding: 5px; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
		</style>
"@
	rm "$rootfolder\HtmlReport.html" -force -ea 0

	$global:outfile = "$rootfolder"
	$htmlfile = "$rootfolder\HtmlReport.html"
	$servername = $ENV:ComputerName

	$line = "<hr>"

	Write-Output "<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<style>
BODY{font-family: Segoe UI;font-size:12pt;}
TABLE{margin-left:auto;margin-right:auto;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;color:white;background-color:#FFFFFF }
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:red}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black}
</style>
</head><body>
<img src=".//logo.png" alt="display this">
</body></html>"	| Set-Content $htmlfile 
	
	
	$date = (Get-Date).ToShortDateString()
	
	$htmlheader = $null | ConvertTo-Html -Body "<h1>Exchange Health Check Report </h1>`n<h5>Generated on $date | powered by Eddie | https://exchangeblogonline.de</h5>" -Head $Style | Add-Content $htmlfile -ea 0 -wa 0 
  
	$day = (get-date).Day
	$today = (Get-Date).AddDays(-7)
        
	$a = updatehistory  | ConvertTo-HTML -PreContent "<h3>Last installed Windows Updates:</h3>"

	$b = Get-MailboxDatabase -Status | 	select Name, DatabaseSize, AvailableNewMailboxspace, Server, EdbFilePath, LastFullBackup | ConvertTo-HTML -PreContent "<h3>MailboxDatabase Status:</h3>"  

	$c = Get-DatabaseAvailabilityGroup | ForEach { 
		$_.Servers | 
		ForEach { Get-MailboxDatabaseCopyStatus -Server $_ | select Name, Status, CopyQueueLength, ReplayQueueLength, LastInspectedLogTime, ContentIndexState 
		} 
	} | ConvertTo-Html -PreContent "<h3>DatabaseAvailabilityGroup Status:</h3>"  

	$d = Get-Service msexchange* | ? { $_.Status -eq "Stopped" } | select Status, Name, DisplayName |
	ConvertTo-Html -PreContent "<h3>Stopped Exchange Services detected:</h3>" 

	$e = Get-MailboxServer | select Name, DatabaseAvailabilityGroup, DatabaseCopyAutoActivationPolicy | 
	ConvertTo-Html -PreContent "<h3>DatabaseActivationPolicy:</h3>"  

	$f = Get-TransportService | Get-Queue -ea 0 | select Identity, DeliveryType, Status, MessageCount, Velocity, RiskLevel, OutboundIPPool, NextHopDomain | 
	ConvertTo-Html -PreContent "<h3>MailQueueStatus:</h3>"  

	$g = Get-ServerComponentState $Servername | select ServerFqdn, Component, State | 
	ConvertTo-Html -PreContent "<h3>ServerComponentState:</h3>"  

	$i = (Get-DatabaseAvailabilityGroup) | ForEach { $_.Servers | ForEach { Test-ReplicationHealth -Server $_ -ea 0 | 
			select Server, Check, Result, Error } } |
	ConvertTo-Html -PreContent "<h3>ReplicationHealthStatus:</h3>" 
   
	$j = Get-MailboxDatabase -Status | select Name, Server, Recovery, ReplicationType | 
	ConvertTo-Html -PreContent "<h3>MailboxDatabaseStatus:</h3>" 
	 
	$k = Test-ServiceHealth -ea 0 | select Role, RequiredServicesRunning | 
	ConvertTo-Html -PreContent "<h3>ServiceHealthStatus</h3>"  
    
	$l = Test-ServiceHealth -ea 0 | select ServicesNotRunning -ExpandProperty ServicesNotRunning | 
	ConvertTo-Html -PreContent "<h3>ServiceHealth ServicesNotRunning Status</h3>"  

	$m = Get-TransportService | % { Test-Mailflow $_ -ea 0 } | 
	ConvertTo-Html -PreContent "<h3>MailFlowStatus:</h3>"  
	
	$n = Get-DatabaseAvailabilityGroup -Status | select PrimaryActiveManager | 
	ConvertTo-HTML -PreContent "<h3>PrimaryActiveManager:</h3>" 
	
	$o = excerts | ConvertTo-Html -PreContent "<h3>Exchange Certs:</h3>"

	$j = Get-ServerHealth -Identity $servername | ? { $_.AlertValue -eq "Unhealthy" } |	select Server, State, Name, TargetResource, HealthSetName, AlertValue, ServerComponent | 
	ConvertTo-HTML -PreContent "<h3>Get-ServerHealth :</h3>" 
	
	$k = Test-MAPIConnectivity -MAPIConnectivity -Server $servername  | select MailboxServer, Database,Result,Error | 
	ConvertTo-HTML -PreContent "<h3>Test-MAPIConnectivity:</h3>"

	
		ConvertTo-HTML -body "
		$line
		$a $line
		$g $line
		$j $line    
		$i $line     
		$g $line
		$k $line        
		$e $line
		$l $line
		$m $line
		$d $line
		$c $line
		$f $line
		$b $line
		$n $line
		$o $line
		$j $line		
		"  -Head $Style | Add-Content "$htmlfile"				
	
    $rootfolder  = (Get-Item -Path "..\").FullName
    $currentfolder  = (Get-Item -Path "..\").FullName
    $ReportingEvents = "C:\Windows\SoftwareDistribution\ReportingEvents.log"
    $ReportingEventsFiltered =  "$rootfolder\LOG\ReportingEvents.log"

	$updatelogs = gc $ReportingEvents | findstr (Get-Date -Format yyyy-MM-dd)
	$updatelogs | Out-File "$ReportingEventsFiltered" -Encoding UTF8 
   
	$tail = $null | ConvertTo-Html -Body "<h5>Generated on $date | powered by Eddie | https://exchangeblogonline.de</h5>" | Add-Content $htmlfile -ea 0 -wa 0 
   	   
}


generateHTML
cd ..\