$SQLDBName = "DPD_DB"
$table="Direct_View"
$path="\\10.47.17.20\pmi-dbo\css\export\"
$file="PMICSS_"
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

do

{
  $date = $date.AddDays(-1)
} while ($date.DayOfWeek -eq [System.DayOfWeek]::Saturday -or 
     $date.DayOfWeek -eq [System.DayOfWeek]::Sunday)
write-host $Date
$xlsxFile = $path + $file +  $date.ToString('yyyy_MM_dd')+'.xlsx' ;
$SQL= "SELECT *  FROM [dbo].[$table] WHERE left([D&T],10) = '" + $date.ToString('yyyy-MM-dd') + "'"
 

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw ;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
 $SqlQuery = "SELECT *  FROM [dbo].[$table] WHERE left([D&T],10) = '" + $date.ToString('yyyy-MM-dd') + "'"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlConnection.open()
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$Row =$SqlAdapter.Fill($DataSet)

if ( $Row -ne 0){
  #$SqlAdapter.Fill($DataSet) >$null | Out-Null
  $SqlConnection.close()                
  
  
              if (Test-Path $xlsxFile)
              {
                          $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
                          $LogString = $LogClause + $file + " Non_exported;" 
                          write-host $file+$date.ToString('yyyy_MM_dd')'.xlsx Non-exported, file exist'
                          Write-output $LogString | Out-File $Logpath -Append    
              }
              else
              {
                          $Sheet='PMICSS'
                          $xlsx= Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDBName -user $uid -Password $passw  -Query $SQL |
                          Select-Object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors |
                          export-Excel -Path $xlsxFile -WorksheetName $Sheet -Autosize  -BoldTopRow -PassThru 
  
                          $ws = $xlsx.Workbook.Worksheets[$Sheet]
                          Set-ExcelColumn -Worksheet $ws -Column 4 -NumberFormat '0' -AutoSize
                          Set-ExcelColumn -Worksheet $ws -Column 5 -Width 17
                          Set-ExcelColumn -Worksheet $ws -Column 6 -NumberFormat '0' -AutoSize                          
                          Close-ExcelPackage $xlsx
  
              #$DataSet.Tables[0] | export-csv -Delimiter ";" -Path $xlsFile -NoTypeInformation -Encoding UTF8 
              $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
              $LogString = $LogClause + $file + " export;" 
              write-host $file$date.ToString('yyyy_MM_dd')'.xlsx Exported'
              Write-output $LogString | Out-File $Logpath -Append  
              };
  }
  else{
              $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
              $LogString = $LogClause + $table +" no data rows" 
              write-host $LogString
              Write-output $LogString | Out-File $Logpath -Append  
  
  }
  
  $SqlConnection.Close()

$MyFileName = "EAN_export.ps1"
 $filebase = Join-Path $PSScriptRoot $MyFileName

Powershell.exe -File  $filebase
## - End of Script - ##