<#
    PBIRS-Alerts.ps1
    - Version: 1.0
    - Last Modified: 7/9/2024
    - SECOUGHL : File creation and initial alert additions  
#>
param (
)

function Invoke-SqlQuery {
    [cmdletbinding()]
    param (
        [string]$ServerInstance,
        [string]$Database = 'master',
        [string]$Query,
        [string]$Description
    )
    
    try {
       
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = "Server=$ServerInstance;Database=$Database;Integrated Security=True"
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = $Query
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet) | Out-Null
        $SqlConnection.Close()

        If ($DataSet.Tables[0].Item.Count -gt 0) {
            Write-Verbose "$($Description): Returned $($DataSet.Tables[0].Item.Count) Rows"
        }

        return $DataSet.Tables[0]
    }
    catch {
        if ($Description) {
            $FriendlyError = "Couldn't connect to $ServerInstance while executing $Description"
        }
        else {
            $FriendlyError = "Couldn't connect to $ServerInstance"
        }
        Write-Host $FriendlyError -foregroundcolor Red
        return
    }

}

$CSS = @"
<style type="text/css">
table
{
Margin: 0px 0px 0px 4px;
Border: 1px solid rgb(190, 190, 190);
Font-Family: Tahoma;
Font-Size: 8pt;
Background-Color: rgb(252, 252, 252);
}
tr:hover td
{
Background-Color: rgb(0, 127, 195);
Color: rgb(255, 255, 255);
}
tr:nth-child(even)
{
Background-Color: rgb(110, 122, 130);
}
th
{
Text-Align: Left;
Color: rgb(150, 150, 220);
Padding: 1px 4px 1px 4px;
}
td
{
Vertical-Align: Top;
Padding: 1px 4px 1px 4px;
}
footer
{ color:black;
  margin-left:4px;
  font-family:Tahoma;
  font-size:8pt;
}
</style>
<h1> Failed Refresh Report</h1>
<h4>Hello,</h4>
<p>You are receiving this report because you are either the original creator of, or the last person to modify, a scheduled refresh in PowerBI Report Server which has failed to run.</p>
<p>Please see the errors below and resolve or remove the schedule if it is no longer needed.</p> 
"@

$Footer = @"
<footer>
<p>Please verify the following before opening a ticket with the Data Team:</p>
<ul style="list-style-type:disc;">
<li>The data source is still valid</li>
<li>The refresh account information is correct and up-to-date</li>
<li>The refresh account is enabled in Active Directory</li>
</ul>
<p>Respectfully,<br>
<a href="mailto:PBIManagers@contoso.com">Contoso Data Team</a>
</footer>
"@

$DebugQuery = @"
select 
              ModifiedBy.UserName ModifiedBy ,
			  CreatedBy.UserName CreatedBy,
              sh.[Message],
              s.[description],
              [path],
              [lastruntime]
              from [catalog] c
              inner join subscriptions s on c.itemid = s.report_oid
              inner join users on s.OwnerID = users.UserID
              inner join SubscriptionHistory sh on sh.SubscriptionID = s.SubscriptionID and sh.EndTime = s.LastRunTime
              INNER JOIN Users ModifiedBy ON c.ModifiedByID = ModifiedBy.UserID
			  INNER JOIN Users CreatedBy ON c.CreatedByID = CreatedBy.UserID
              where LastStatus like '%failed%'
			 and( 
			  c.ModifiedByID in  (
				'CC59A546-51A1-4898-A7D9-561B0B9B06C5',
				'5F480C0C-81C4-4EEB-B504-61858AFF9821',
				'17DC8EDF-0AF4-4DF3-845B-50788BA9AB0E',
				'BE6E6AE2-F16B-4101-A41C-D74B5B8EF2FF',
				'6B82E39A-ABE3-4417-B1A0-FA3B681E2BBE',
				'16E6E69C-5948-4283-BC89-7F36A2BD619D')
									or
									
			  c.CreatedByID in  (
				'CC59A546-51A1-4898-A7D9-561B0B9B06C5',
				'5F480C0C-81C4-4EEB-B504-61858AFF9821',
				'17DC8EDF-0AF4-4DF3-845B-50788BA9AB0E',
				'BE6E6AE2-F16B-4101-A41C-D74B5B8EF2FF',
				'6B82E39A-ABE3-4417-B1A0-FA3B681E2BBE',
				'16E6E69C-5948-4283-BC89-7F36A2BD619D')
				)
		
"@

$Query = @"
		        
SELECT
	ModifiedBy.UserName ModifiedBy
	, CreatedBy.UserName CreatedBy
	, sh.[Message]
	, s.[description]
	, [path]
	, [lastruntime]
FROM 
	[catalog] c
		JOIN subscriptions s on c.itemid = s.report_oid
		JOIN users on s.OwnerID = users.UserID
		JOIN SubscriptionHistory sh on sh.SubscriptionID = s.SubscriptionID and dateadd(MILLISECOND, -DATEPART(millisecond,sh.EndTime),sh.EndTime) = dateadd(MILLISECOND, -DATEPART(millisecond,s.LastRunTime),s.LastRunTime)
		JOIN Users ModifiedBy ON c.ModifiedByID = ModifiedBy.UserID
		JOIN Users CreatedBy ON c.CreatedByID = CreatedBy.UserID
WHERE 
	s.LastStatus like '%failed%'
	-- and path like ''
ORDER BY 
	c.[Path] desc
"@


$SMTPRelay = "smtp-relay.contoso.com"
$From = "PBIRS Alerts<noreply@contoso.com>"
$PBIManagers = "PBIManagers@contoso.com"
$SMTPUsername = "anonymous"
$SMTPPassword = ConvertTo-SecureString "anonymous" -AsPlainText -Force
$SMTPCredential = New-Object System.Management.Automation.PSCredential($SMTPUsername, $SMTPPassword)
$Subject = "PowerBI Report Server Failed Refreshes"

$ServerInstance = 'SQLServer.contoso.com\Instance'
$Database = 'PBI_ReportServer'

#Import bespoke alerts
$Alerts = Import-PowerShellDataFile "C:\AuthorizedScripts\PBIRSrefresh\PBIRS Bespoke Alerts.psd1"

$FailedRefreshes = Invoke-SqlQuery -ServerInstance $ServerInstance -Database $Database -Query $Query | Group-Object username

######### Individual Notifications #########

#Populate all users from the query into an arraylist
$Users = @()
$Users += $FailedRefreshes[0].Group.CreatedBy
$Users += $FailedRefreshes[0].Group.ModifiedBy 


#Have to pad the hashtable because we index it as an array later, and PS is too "helpful" with that interaction
$UserHash=@{'pad'='pad'}

For($i=0;$i -lt $Users.Count;$i++){
    
    if (!$UserHash.ContainsKey($Users[$i]))
        {
            Try{
                $UserTemp = Get-ADUser $Users[$i].Replace('contoso\','') -Properties EmailAddress
                $UserHash.Add($Users[$i],$UserTemp.EmailAddress)
            }
            catch{
                Write-Host "Could not resolve $($Users[$i])" -ForegroundColor Red
            }
            
        }
    }

# Process Indivual (automatic) Notifications
For($i=0;$i -lt $UserHash.Count;$i++){
    If ($($UserHash.Keys)[$i] -notlike 'pad'){   
        $Notifications = $null
        Write-Verbose "Failed Refresh Notification for: " $($UserHash.Keys)[$i] -ForegroundColor Cyan
        $Notifications = $FailedRefreshes.Group | Where-Object {($_.ModifiedBy -Like $($UserHash.Keys)[$i]) -or ($_.CreatedBy -Like $($UserHash.Keys)[$i])}
    
        $HTML = $Notifications | Select-Object * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | ConvertTo-Html -Head $CSS -PostContent $Footer
        $FileName = $($UserHash.Keys)[$i].ToString().Replace('contoso\','')
     
        Send-MailMessage -From $From -to $($UserHash.Values)[$i] -Subject $Subject -Bodyashtml ($HTML|out-string) -SmtpServer $SMTPRelay -Credential $SMTPCredential -ErrorAction Stop
        }
    }

# Process individual/group (manual) notifications
For($i=0;$i -lt $Alerts.Alert.Count;$i++){
    if ($Alerts.Alert[$i].Path -in $FailedRefreshes.Group.Path){
        Write-Verbose "Failed Refresh Notification for: $($Alerts.Alert[$i].Email) regarding Path $($Alerts.Alert[$i].Path)"
        $Notifications = $FailedRefreshes.Group | Where-Object {$_.Path -like $Alerts.Alert[$i].Path}
        $HTML = $Notifications | Select-Object * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | ConvertTo-Html -Head $CSS -PostContent $Footer
        Send-MailMessage -From $From -to $($Alerts.Alert[$i].Email) -Subject $Subject -Bodyashtml ($HTML|out-string) -SmtpServer $SMTPRelay -Credential $SMTPCredential -ErrorAction Stop
        }

    }


########## Team notification for all failures ##########
$HTML = $FailedRefreshes.Group | Select-Object * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | ConvertTo-Html -Head $CSS -PostContent $Footer
Send-MailMessage -From $From -to $PBIManagers -Subject $Subject -Bodyashtml ($HTML|out-string) -SmtpServer $SMTPRelay -Credential $SMTPCredential -ErrorAction Stop
