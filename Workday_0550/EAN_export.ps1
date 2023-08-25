$SQLDBName = "DPD_DB"
$table="EAN"
$path="\\10.47.17.20\pmi-dbo\\Ecomm\Reitmaier\Kuehne_Nagel\Sklady_Lokace\Portfolio_zboží_aktual\"
$file="EAN"
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\PMIdblog.txt'
$Date = Get-Date 
$Datum = Get-date


$SQLlogpath = "\\10.47.17.20\pmi-dbo\SQL_script\safe\sql.txt"

$SQLlog = Get-Content -Path $SQLlogpath
$SQLServer = $SQLlog[0]
$uid =$SQLlog[1]
$base64Encoded = $SQLlog[2]
$bytes = [System.Convert]::FromBase64String($base64Encoded)
$passw  = [System.Text.Encoding]::UTF8.GetString($bytes)

$xlsxFile = $path + $file +'.xlsx' ;
$SQL= "SELECT [EAN PK],[EAN CT],[EAN BX],[MATNR],[MAKTX] FROM [dbo].[$table]"
 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw ;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlQuery = "SELECT [EAN PK],[EAN CT],[EAN BX],[MATNR],[MAKTX] FROM [dbo].[$table]"
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlConnection.open()
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$Row =$SqlAdapter.Fill($DataSet)

if ( $Row -ne 0)
{
  #$SqlAdapter.Fill($DataSet) >$null | Out-Null
  $SqlConnection.close()

        if (Test-Path $xlsxFile)
        {
            Remove-Item $xlsxFile
        }    
            $Sheet='EAN'
            $xlsx= Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDBName -user $uid -Password $passw  -Query $SQL |
            Select-Object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors |
            export-Excel -Path $xlsxFile -WorksheetName $Sheet -Autosize  -BoldTopRow -PassThru 
  
            $ws = $xlsx.Workbook.Worksheets[$Sheet]
            Set-ExcelColumn -Worksheet $ws -Column 1 -NumberFormat '0' -AutoSize
            Set-ExcelColumn -Worksheet $ws -Column 2 -NumberFormat '0' -AutoSize
            Set-ExcelColumn -Worksheet $ws -Column 3 -NumberFormat '0' -AutoSize
            Close-ExcelPackage $xlsx
  
              #$DataSet.Tables[0] | export-csv -Delimiter ";" -Path $xlsFile -NoTypeInformation -Encoding UTF8 
              $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
              $LogString = $LogClause + $file + " export;" 
              write-host $file'.xlsx Exported' - $date
              Write-output $LogString | Out-File $Logpath -Append  

}
else
{
              $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
              $LogString = $LogClause + $table +" no data rows" 
              write-host $LogString
              Write-output $LogString | Out-File $Logpath -Append  
}
  
  $SqlConnection.Close()

$MyFileName = "EAN_import.ps1"
 $filebase = Join-Path $PSScriptRoot $MyFileName

Powershell.exe -File  $filebase
## - End of Script - ##