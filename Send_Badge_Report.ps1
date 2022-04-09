#Connect to SQL and run QUERY 
$SQLServer = "server" 
$SQLDBName = "db"
$SQLUsername = "it"
$SQLPassword = "pw"
$date = Get-Date -Format "MM_dd_yyyy"

##Delete the output file if it already exists
If (Test-Path \\10.0.1.23\C$\VendingReport_Logs\cp_badges_$date.csv ){
    Remove-Item \\10.0.1.23\C$\VendingReport_Logs\cp_badges_$date.csv
}

$OutputFile = "\\10.0.1.23\C$\VendingReport_Logs\cp_badges.csv"
 
$SqlQuery = "exec rpt_badge_sheet"
  
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
$DataSet.Tables[0] | select "Last","First","Badge_ID","Vending_ID","Allowance" | Export-Csv $OutputFile -NoTypeInformation

sleep 3

$From = ""
$To = ""
$Attachment = $OutputFile
$Subject = "Coldpoint Badges"
$Body = "This is an automated message - Please see the attached."
$SMTPServer = "smtp.office365.com"
$SMTPPort = "587"
$secpasswd = ConvertTo-SecureString "pw" -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ("", $secpasswd)

Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $cred -Attachments $Attachment

Rename-Item -Path \\10.0.1.23\C$\VendingReport_Logs\cp_badges.csv -NewName \\10.0.1.23\C$\VendingReport_Logs\cp_badges_$date.csv