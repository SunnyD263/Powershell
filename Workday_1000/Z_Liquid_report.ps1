$SQLServer = "DENOTSQ161"
$SQLDBName = "Liquid"
$uid = "sqczpmip1_app"
$passw = "lwjfh/ezt.34Hf"
$table = "Export_Over_View"
$path = "\\10.47.17.20\pmi-dbo\SQL_script\PS_Script\Report\"
$file = "Liquid_KN_report_"
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_1000\Log\PMIdblog.txt'
$Email = 'Petra.Knizova@pmi.com'
$EmailCopy = 'jan.sonbol@kuehne-nagel.com;Tomas.Burian@pmi.com;vlastimil.zika@kuehne-nagel.com'
$Date = Get-Date 

do
{
  $date = $date.AddDays(0)
} while ($date.DayOfWeek -eq [System.DayOfWeek]::Saturday -or 
  $date.DayOfWeek -eq [System.DayOfWeek]::Sunday)

$xlsxFile = $path + $file + $date.ToString("yyyy_MM_dd") + '.xlsx';
$xlsxF = $file + $date.ToString("yyyy_MM_dd") + '.xlsx'
$SQL = "SELECT * FROM [dbo].[$table]"
#$SQL= "SELECT *  FROM [dbo].[TradeIN_View] WHERE convert(varchar,[RCVDate],23) = '" + $date.ToString("yyyy-MM-dd") + "'"
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlQuery = $SQL
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlConnection.open()
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$Row = $SqlAdapter.Fill($DataSet)

if ( $Row -ne 0) {
  #$SqlAdapter.Fill($DataSet) >$null | Out-Null
  #$SqlConnection.close()                

  write-host $xlsxFile
  if (Test-Path $xlsxFile) {
    $LogClause = $Date.ToString("dd/MM/yyyy HH:mm:ss") + "; " + $env:UserDomain + "\" + $env:UserName + "; " + $env:ComputerName + "; "
    $LogString = $LogClause + $file + " Non_exported;" 
    write-host $file $date.ToString("yyyy_MM_dd")'.xlsx Non-exported, file exist'
    Write-output $LogString | Out-File $Logpath -Append    
  }
  else {


   
    $Sheet = 'Overview'
    $xlsx = Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDBName -user $uid -Password $passw -Query $SQL |
    Select-Object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors  |
    export-Excel -Path $xlsxFile -WorksheetName $Sheet -Autosize  -BoldTopRow -PassThru 
    
    $ws = $xlsx.Workbook.Worksheets[$Sheet]
    Set-ExcelColumn -Worksheet $ws -Column 6 -Width 17
    Set-ExcelColumn -Worksheet $ws -Column 7 -Width 17
    Set-ExcelColumn -Worksheet $ws -Column 8 -Width 17
    Set-ExcelColumn -Worksheet $ws -Column 9 -Width 17
    Set-ExcelColumn -Worksheet $ws -Column 10 -Width 17
    Close-ExcelPackage $xlsx
  
   
    # $DataSet.Tables[0] | export-csv -Delimiter ";" -Path $xlsxFile -NoTypeInformation -Encoding UTF8 
    $LogClause = $Date.ToString("dd/MM/yyyy HH:mm:ss") + '; ' + $env:UserDomain + '\' + $env:UserName + '; ' + $env:ComputerName + ';'
    $LogString = $LogClause + $file + ' export;'
    write-host $file $date.ToString("yyyy_MM_dd")".xlsx Exported"
    Write-output $LogString | Out-File $Logpath -Append 
    $LogClause = $Date.ToString("dd/MM/yyyy HH:mm:ss") + '; ' + $env:UserDomain + '\' + $env:UserName + '; ' + $env:ComputerName + '; '
    $LogString = $LogClause + $xlsxF + " Upload;"
    Write-output $LogString | Out-File $Logpath -Append    

    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = $Email
    $Mail.Cc = $EmailCopy
    $Mail.Subject = $xlsxF
    $Mail.Body = 'Pravidelny report stavu likvidace cigaret'
    $Mail.Attachments.Add($xlsxFile)
    $Mail.Send()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

  }
}
else {
  $LogClause = $Date.ToString("dd/MM/yyyy HH:mm:ss") + '; ' + $env:UserDomain + '\' + $env:UserName + '; ' + $env:ComputerName + '; '
  $LogString = $LogClause + $table + ' no data rows'
  write-host $LogString
  Write-output $LogString | Out-File $Logpath -Append  

}

$SqlConnection.Close()

#$MyFileName = 'TradeIN_report.ps1'
#$filebase = Join-Path $PSScriptRoot $MyFileName
#Powershell.exe -File  $filebase

## - End of Script - ##



  