#Connect to SQL and run QUERY 
$SQLServer = "server" 
$SQLDBName = "db"
$SQLUsername = "it"
$SQLPassword = "pw"

$OutputFile = "\\10.0.1.23\C$\FollowUp.csv"

$SqlQuery = "Select ID, CreatedDT, rtrim(FollowUpBy)[FollowUpBy], rtrim(TM_Email)[TM_Email], rtrim(Subject)[Subject] from complaints where Followupdate <= getdate() and Complete = 'N'"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection

$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd
  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close() 
 
#Output RESULTS to CSV
$DataSet.Tables[0] | select "ID", "CreatedDT", "FollowUpBy","TM_Email","Subject" | Export-Csv $OutputFile -NoTypeInformation

$From = "notifications@coldpoint.us"
$To = ""
$Subject = "The following complaints are due."
$Attachment = $OutputFile
$Body = "Please see the attachment for complaints that are past due or due for review today."
$SMTPServer = "smtp.office365.com"
$SMTPPort = "587"
$secpasswd = ConvertTo-SecureString "pw" -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ("notifications@coldpoint.us", $secpasswd)

Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $cred -Attachments $Attachment

Remove-Item \\10.0.1.23\C$\FollowUp.csv