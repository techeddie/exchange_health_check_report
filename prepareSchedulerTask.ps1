################################################
# 
# AUTHOR:  Eddie
# EMAIL:   eddie@directbox.de
# BLOG:    https://exchangeblogonline.de
# COMMENT: EXCHANGE_HEALTH_CHECK_REPORT
#
################################################

#create script and log folder 
$scriptfolder = (Get-Item -Path ".\").FullName

#create scheduler as system account - admin rights needed
function Invoke-PrepareScheduledTaskStartExchangeHealthREPORT {
    $taskName = "EXCHANGE_HEALTH_CHECK_REPORT"
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($task -ne $null) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false 
    }

    # TODO: EDIT THIS STUFF AS NEEDED...
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-File "EXCHANGE_HEALTH_CHECK_REPORT.ps1"' -WorkingDirectory $scriptfolder
    $date = (Get-Date).AddYears(1990)
    $trigger =  New-ScheduledTaskTrigger -Once -At $date
    $settings = New-ScheduledTaskSettingsSet -Compatibility Win8

    $principal = New-ScheduledTaskPrincipal -UserId SYSTEM -LogonType ServiceAccount -RunLevel Highest

    $definition = New-ScheduledTask -Action $action -Principal $principal -Trigger $trigger -Settings $settings -Description "Run $($taskName) at startup"

    Register-ScheduledTask -TaskName $taskName -InputObject $definition

    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    # TODO: LOG AS NEEDED...
    if ($task -ne $null) {
        Write-Output "Created scheduled task: '$($task.ToString())'." 
    } else {
        Write-Output "Created scheduled task: FAILED."
    }
}

#execute scheduler function
Invoke-PrepareScheduledTaskStartExchangeHealthREPORT 

#open task scheduler
taskschd.msc