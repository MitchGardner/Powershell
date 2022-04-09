#Connect to SQL and run QUERY 
$SQLServer = "server" 
$SQLDBName = "db"
$SQLUsername = "it"
$SQLPassword = "pw"
 
$SqlQuery = "exec Export_ASN"

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

$Cust_PO = $DataSet.Tables[0].Rows[0].ExternalUID
$OutputFile = "\\10.0.1.8\c$\MT Share\Cadence_Import\$Cust_PO.csv"

#Output RESULTS to CSV
$DataSet.Tables[0] | select "ExternalUID","ClientID","Lot","SKU","UPC", "Actual_Qty", "ParentMUID", "Mfg_Date" | Export-Csv $OutputFile -NoTypeInformation