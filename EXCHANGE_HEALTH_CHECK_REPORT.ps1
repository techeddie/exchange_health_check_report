<#############################################################################

     AUTHOR:         
     Eddie
     
     BLOG:			 
     https://exchangeblogonline.de
    
     E-MAIL:         
     eddie@directbox.de
	 
     VERSION:		 
     2.0

     COMMENT:        
     EXCHANGE HEALTH CHECK REPORT
     Script must be run on Exchange host
								   
##############################################################################>

Add-PSSnapin *exchange*

#declaration vars
#script folder 
$scriptfolder = (Get-Item -Path ".\").FullName
$Servername = $ENV:ComputerName
$htmlFile = "$scriptfolder\HtmlReport.html"
$generateHTMLFile = "$scriptfolder\HTML\generateHTML.ps1"

#smtp settings:
$configfiles = get-content "$scriptfolder\params.txt"
$global:smtpserver = ($configfiles.GetValue(0)).Split(":")[1] 
$global:smtpsender = $configfiles.GetValue(1).Split(":")[1]
$global:smtpport = $configfiles.GetValue(2).Split(":")[1]
$global:smtprecipient = ($configfiles.GetValue(3).Split(",")).replace("recipients:", "")

#create script and log folder 
$logfolder = "$scriptfolder\LOG"
mkdir $scriptfolder -ErrorAction SilentlyContinue
mkdir $logfolder -ErrorAction SilentlyContinue


#function create eventlog sucess result
function eventsucess {
    $eventlog = "ExchangeMaintenanceEvents"

    $check = Get-EventLog "$eventlog" -Newest 1

    if (!($check)) {
        New-EventLog -LogName "$eventlog" -Source "$eventlog"
    }

    Write-EventLog -LogName "$eventlog"  `
        -Source "$eventlog" `
        -EventId 2000 `
        -Message "MAINTENANCE_REPORT wurde erfolgreich abgeschlossen!"  `
        -EntryType Information

}

#function collect exchange certs
function excerts {
    #autor: from Paul Cunningham
    #https://gallery.technet.microsoft.com/office/Exchange-Certificate-91578ac4
    $exchangeservers = @(Get-ExchangeServer)

    foreach ($server in $exchangeservers) {
        $htmlsegment = @()
    
        $serverdetails = "Server: $($server.Name) ($($server.ServerRole))"
        Write-Host "Collecting cert information from Server: $serverdetails" -ForegroundColor Green
    
        $certificates = @(Get-ExchangeCertificate -Server $server)  # | ?{$_.services -match "IIS"} )

        $certtable = @()

        foreach ($cert in $certificates) {
        
            $iis = $null
            $smtp = $null
            $pop = $null
            $imap = $null
            $um = $null       
        
            $subject = ((($cert.Subject -split ",")[0]) -split "=")[1]
                
            if ($($cert.IsSelfSigned)) {
                $selfsigned = "Yes"
            } else {
                $selfsigned = "No"
            }

            $issuer = ((($cert.Issuer -split ",")[0]) -split "=")[1]

            #$domains = @($cert | Select -ExpandProperty:CertificateDomains)
            $certdomains = @($cert | Select -ExpandProperty:CertificateDomains)
            if ($($certdomains.Count) -gt 1) {
                $domains = $null
                $domains = $certdomains -join ", "
            } else {
                $domains = $certdomains[0]
            }

            #$services = @($cert | Select -ExpandProperty:Services)
            $services = $cert.ServicesStringForm.ToCharArray()

            if ($services -icontains "W") { $iis = "Yes" }
            if ($services -icontains "S") { $smtp = "Yes" }
            if ($services -icontains "P") { $pop = "Yes" }
            if ($services -icontains "I") { $imap = "Yes" }
            if ($services -icontains "U") { $um = "Yes" }

            $certObj = New-Object PSObject
            $certObj | Add-Member NoteProperty -Name "Subject" -Value $subject
            $certObj | Add-Member NoteProperty -Name "Status" -Value $cert.Status
            $certObj | Add-Member NoteProperty -Name "Expires" -Value $cert.NotAfter.ToShortDateString()
            $certObj | Add-Member NoteProperty -Name "Self Signed" -Value $selfsigned
            $certObj | Add-Member NoteProperty -Name "Issuer" -Value $issuer
            $certObj | Add-Member NoteProperty -Name "SMTP" -Value $smtp
            $certObj | Add-Member NoteProperty -Name "IIS" -Value $iis
            $certObj | Add-Member NoteProperty -Name "POP" -Value $pop
            $certObj | Add-Member NoteProperty -Name "IMAP" -Value $imap
            $certObj | Add-Member NoteProperty -Name "UM" -Value $um
            $certObj | Add-Member NoteProperty -Name "Thumbprint" -Value $cert.Thumbprint
            $certObj | Add-Member NoteProperty -Name "Domains" -Value $domains
        
            $certtable += $certObj
        }

        $htmlcerttable = $certtable 
        return $certtable
    }
}

#function generate html outfile
function generateHTML {
    & $generateHTMLFile
}

#function send email to each recipient

function sendmail {  
    	
    $addresses = $smtprecipient[0]
    $addresses | % { 
	
        #addHTML
        $htmlBody = get-Content "$scriptfolder\HtmlReport.html" | out-string
        $attachment = "$scriptfolder\logo.png"
        $date = Get-Date -Format dd.MM.yyyy
			
        Send-MailMessage -smtpServer $smtpserver -Port $smtpport -From $smtpsender -To $smtprecipient -Subject "$servername - Exchange Maintenance Report on $date" -BodyAsHtml $htmlBody  -Attachments "$scriptfolder\LOG\ReportingEvents.log", "$attachment"
        
    }
    
}


#run functions
generateHTML
sendmail
eventsucess

#open html file
start $htmlFile

#END