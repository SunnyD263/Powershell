$SQLDBName = "DPD_DB"
$table = "SWAP_Dvc_sum_View_Export"
$path="\\10.47.17.20\pmi-dbo\Ecomm\SWAP_device_miss\"
$file="PMISWAPMISSDVC_"
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\PMIdblog.txt'
$Date = Get-Date
$Datum = Get-Date 

$Emaillogpath = "\\10.47.17.20\pmi-dbo\SQL_script\safe\email.txt"
$Emaillog = Get-Content -Path $Emaillogpath
$Email = $Emaillog[3] + ";" + $Emaillog[4]
$EmailCopy = $Emaillog[0]

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
  $SQL= "SELECT *  FROM [dbo].[$table] WHERE SUM <> 0 and left(Scantime,10)= '" + $date.ToString('yyyy-MM-dd') + "' or  SUM is null and left(Scantime,10)= '" + $date.ToString('yyyy-MM-dd') + "' order by REFERENCE"

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
    $LogClause = $Datum.ToString("dd/MM/yyyy HH:mm:ss") + "; " + $env:UserDomain + "\" + $env:UserName + "; " + $env:ComputerName + "; "
    $LogString = $LogClause + $file + " Non_exported;" 
    write-host $file $date.ToString("yyyy_MM_dd")'.xlsx Non-exported, file exist'
    Write-output $LogString | Out-File $Logpath -Append    
  }
  else {


   
    $Sheet = 'PMISWAPMISSDVC'
    $xlsx = Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDBName -user $uid -Password $passw -Query $SQL |
    Select-Object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors  |
    export-Excel -Path $xlsxFile -WorksheetName $Sheet -Autosize  -BoldTopRow -PassThru 
    
    $ws = $xlsx.Workbook.Worksheets[$Sheet]
    Set-ExcelColumn -Worksheet $ws -Column 3 -NumberFormat '0' -AutoSize
    Set-ExcelColumn -Worksheet $ws -Column 13 -NumberFormat '0' -AutoSize
    Set-ExcelColumn -Worksheet $ws -Column 12 -Width 17   
    Set-ExcelColumn -Worksheet $ws -Column 11 -Width 17    
    Close-ExcelPackage $xlsx
  
    
    # $DataSet.Tables[0] | export-csv -Delimiter ";" -Path $xlsxFile -NoTypeInformation -Encoding UTF8 
    $LogClause = $Datum.ToString("dd/MM/yyyy HH:mm:ss") + '; ' + $env:UserDomain + '\' + $env:UserName + '; ' + $env:ComputerName + ';'
    $LogString = $LogClause + $file + ' export;'
    write-host $file $date.ToString("yyyy_MM_dd")".xlsx Exported"
    Write-output $LogString | Out-File $Logpath -Append 
    $LogClause = $Datum.ToString("dd/MM/yyyy HH:mm:ss") + '; ' + $env:UserDomain + '\' + $env:UserName + '; ' + $env:ComputerName + '; '
    $LogString = $LogClause + $xlsxF + " Upload;"
    Write-output $LogString | Out-File $Logpath -Append    

    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = $Email
    $Mail.Cc = $EmailCopy
    $Mail.Subject = 'Pravidelny report chybejicich zarizeni ze SWAP zasilek'
    $Mail.Body = 'Pravidelny report chybejicich zarizeni ze SWAP zasilek'
    $Mail.Attachments.Add($xlsxFile)
    $Mail.Send()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

  }
}
else {

    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = $Email
    $Mail.Cc = $EmailCopy
    $Mail.Subject = 'Pravidelny report chybejicich zarizeni ze SWAP zasilek'
    $Mail.Body = 'Zadna chybejici zarizeni'
    $Mail.Send() 

  $LogClause = $Datum.ToString("dd/MM/yyyy HH:mm:ss") + '; ' + $env:UserDomain + '\' + $env:UserName + '; ' + $env:ComputerName + '; '
  $LogString = $LogClause + $table + ' no data rows'
  write-host $LogString
  Write-output $LogString | Out-File $Logpath -Append  

}

$SqlConnection.Close()

$MyFileName = 'Ecomm_SWAP_Dvc.ps1'
$filebase = Join-Path $PSScriptRoot $MyFileName
Powershell.exe -File  $filebase