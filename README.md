## PBIRS Alerts
The goal of this projet is to create and email automated failure notifications if a scheduled refresh failed the last time it ran.<br><br>
**Three types of notification are generated:<br>**
1. AD is queried for the EmailAddress attribute of the original creator and the last account to modify a given schedule
2. Custom E-mail addresses can be mapped to all schedules for a given report via the included PowerShell Data File
3. A summary of all failed schedules are sent to the PBIRS manager/DL

## Setup
1. Deploy the script and data file to a local location can be reached by your automation solution of choice.<br>
2. Update Custom-Alerts.psd1 to include any custom alerting<br>
   ex: If you wanted (myteam@contoso.com and myboss@contoso.com to be alerted for all schedules on the report located at https://pbirs.contoso.com/Reports/Folder/MyReport) and (yourteam@contoso.com and yourboss@contoso.com to be alerted for all schedules on the report located at https://pbirs.contoso.com/Reports/Folder/YourReport)
```
@{
    Alert = @(
      @{ 
        Path = '/Folder/MyReport'
        Email = 'myteam@contoso.com','myboss@contoso.com'
       }
      @{ 
        Path = '/Folder/YourReport'
        Email = 'yourteam@contoso.com','yourboss@contoso.com'
       }     
    )
}
```
3. Update PBIRS-Alerts.ps1 (Lines 129-138) in the following manner to fit your environment:<br>

| Variable | Description |
| ----- | ----- |
| $SMTPRelay | dns or ip address of your SMTP Relay |
| $SMTPUsername | Username for SMTP, default anonymous |
| $SMTPPassword | Password for SMTP, default anonymous |
| $Subject | Email subject |
| $ServerInstance | SQL Server which hosts the PBIRS Database |
| $Database | PBIRS Database name |
| $Alerts | Physical location of Custom-Alerts.psd1 |

## Script use:

.\PBIRS-Alerts.ps1

| Parameter | Explanation |
| ---- | ---- |
| `None` | None |

## Greater Specificity on alerting:
**TBD**<br>
- TBD
> ```TBD```


Scheduling:
We currently schedule using a SQL Agent Job which runs daily:
| Step Type | Step Text |
| ---- | ---- |
| Operating System (CmdExec) | C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe "C:\AuthorizedScripts\PBIRSrefresh\PBIRS-Alerts.ps1" |
