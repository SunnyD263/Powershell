$SQLDBName = "DPD_DB"
$table="NonDlv_Dvc_sum_View_export"
$path="\\10.47.17.20\pmi-dbo\Ecomm\NonDlv_device\"
$file="PMINONDELDVC_"
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\PMIdblog.txt'
$Date = Get-Date
$Datum = Get-Date 

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
     $xlsxFile = $path + $file +  $date.ToString('yyyy_MM_dd')+'.xlsx';
     $SQL= "SELECT *   FROM [dbo].[$table] where SUM = 0 and left(Scantime,10)= '" + $date.ToString('yyyy-MM-dd') + "' or  SUM is not null and left(Scantime,10)= '" + $date.ToString('yyyy-MM-dd') + "' order by Reference,Material"


$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw ;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlQuery = $SQL
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
                          write-host $file $date.ToString('yyyy_MM_dd')'.xlsx Non-exported, file exist'
                          Write-output $LogString | Out-File $Logpath -Append    
              }
              else
              {
                          $Sheet='PMINONDELDVC'
                          $xlsx= Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDBName -user $uid -Password $passw  -Query $SQL |
                          Select-Object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors |
                          export-Excel -Path $xlsxFile -WorksheetName $Sheet -Autosize  -BoldTopRow -PassThru 

                          $ws = $xlsx.Workbook.Worksheets[$Sheet]   
                          Set-ExcelColumn -Worksheet $ws -Column 4 -NumberFormat '0' -AutoSize                     
                          Set-ExcelColumn -Worksheet $ws -Column 10 -NumberFormat '0' -AutoSize
                          Set-ExcelColumn -Worksheet $ws -Column 9 -Width 17    
                          Set-ExcelColumn -Worksheet $ws -Column 8 -Width 17                         
                          Close-ExcelPackage $xlsx
  
              #$DataSet.Tables[0] | export-csv -Delimiter ";" -Path $xlsFile -NoTypeInformation -Encoding UTF8 
              $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
              $LogString = $LogClause + $file + " export;" 
              write-host $file $date.ToString('yyyy_MM_dd')'.xlsx Exported'
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

$MyFileName = "Ecomm_SWAP_del_export.ps1"
 $filebase = Join-Path $PSScriptRoot $MyFileName

Powershell.exe -File  $filebase