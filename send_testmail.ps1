<#############################################################################

	 AUTHOR:         Eddie
	 BLOG:			 https://exchangeblogonline.de
	 E-MAIL:         eddie@directbox.de
	 VERSION:		 1.0
	 KOMMENTAR:  	 SMTP TEST								  

##############################################################################>

#smtp settings:
$scriptfolder = (Get-Item -Path ".\").FullName
$configfiles = get-content "$scriptfolder\params.txt"
$global:smtpserver = ($configfiles.GetValue(0)).Split(":")[1]
$global:smtpsender = $configfiles.GetValue(1).Split(":")[1]
$global:smtprecipient = ($configfiles.GetValue(3).Split(",")).replace("recipients:", "")

#function for send mail
function sendmail {  
    	
    $addresses = $smtprecipient[0]
    $addresses | % { 
		
        $servername = $ENV:ComputerName
        $date = Get-Date -Format "dd.MM.yyyy HH:mm"
		
        Send-MailMessage -smtpServer $smtpserver -From $smtpsender -To $smtprecipient -Subject "Test mail sent from host: $servername on $date" -Body "Just for testing purposes"
        
    }
    
}

#run sendmail function
sendmail