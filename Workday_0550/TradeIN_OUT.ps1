$SQLDBName = "DPD_DB"
$table="TradeIN_View"
$path="\\10.47.17.20\pmi-dbo\Data\PCIP\DPD\TRADEIN\OUT\"
$file="TRADEIN_"
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\PMIdblog.txt'
$remotefilepath = "/TRADEIN/OUT/"
$Date = Get-Date 
$Datum = Get-Date 

$FTPlogpath = "\\10.47.17.20\pmi-dbo\SQL_script\safe\wedos.txt"
$SQLlogpath = "\\10.47.17.20\pmi-dbo\SQL_script\safe\sql.txt"

$FTPlog = Get-Content -Path $FTPlogpath
$server = $FTPlog[0]
$user = $FTPlog[1]
$base64Encoded = $FTPlog[2]
$bytes = [System.Convert]::FromBase64String($base64Encoded)
$pass = [System.Text.Encoding]::UTF8.GetString($bytes)


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

$xlsxFile = $path + $file +  $date.ToString("yyyy_MM_dd")+'.xlsx';
$xlsxF =  $file +  $date.ToString("yyyy_MM_dd")+'.xlsx'
$SQL= "SELECT [REFERENCE],[CRTDate],[SHPDate],[RCVDate],[PARCELNO_ST],[PARCELNO],[STATUS],[CdfCharger],[CdfHolder]  FROM [dbo].[$table] WHERE convert(varchar,[RCVDate],23) = '" + $date.ToString("yyyy-MM-dd") + "'"
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
$Row =$SqlAdapter.Fill($DataSet)

if ( $Row -ne 0)
{
#$SqlAdapter.Fill($DataSet) >$null | Out-Null
#$SqlConnection.close()                

write-host $xlsxFile
            if (Test-Path $xlsxFile)
            {
                        $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
                        $LogString = $LogClause + $file+$date.ToString("yyyy_MM_dd")+'.xlsx Non-exported, file exist;'
                        write-host $file $date.ToString("yyyy_MM_dd")'.xlsx Non-exported, file exist'
                        Write-output $LogString | Out-File $Logpath -Append    
            }
            else
            {
                        $Sheet='TRADEIN'
                        $xlsx= Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDBName -user $uid -Password $passw -Query $SQL |
                        Select-Object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors |
                        export-Excel -Path $xlsxFile -WorksheetName $Sheet -Autosize  -BoldTopRow -PassThru -NoNumberConversion '*'

                        $ws = $xlsx.Workbook.Worksheets[$Sheet]
                        Set-ExcelColumn -Worksheet $ws -Column 1 -NumberFormat '0' -AutoSize
                        Set-ExcelColumn -Worksheet $ws -Column 5 -NumberFormat '0' -AutoSize
                        Set-ExcelColumn -Worksheet $ws -Column 6 -NumberFormat '0' -AutoSize
                        Close-ExcelPackage $xlsx

           # $DataSet.Tables[0] | export-csv -Delimiter ";" -Path $xlsxFile -NoTypeInformation -Encoding UTF8 
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss") + '; ' + $env:UserDomain + '\' + $env:UserName + '; '+ $env:ComputerName + ';'
            $LogString = $LogClause + $file+$date.ToString("yyyy_MM_dd")+".xlsx Exported"
            write-host $file $date.ToString("yyyy_MM_dd")".xlsx Exported"
            Write-output $LogString | Out-File $Logpath -Append 

            $webclient = New-Object System.Net.WebClient
            $webclient.Proxy=$null
            $webclient.Credentials = New-Object System.Net.NetworkCredential($user, $pass)
              $fileuri = New-Object System.Uri("ftp://$server$remotefilepath$xlsxF")
             $webclient.UploadFile($fileuri, $xlsxFile)
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ '; '+ $env:UserDomain + '\' + $env:UserName + '; '+ $env:ComputerName + '; '
            $LogString = $LogClause + $xlsxF + " Upload;"
            Write-output $LogString | Out-File $Logpath -Append    
            }
}
else{
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss") + '; '+ $env:UserDomain + '\' + $env:UserName + '; ' + $env:ComputerName + '; '
            $LogString = $LogClause + $table + ' no data rows'
            write-host $LogString
            Write-output $LogString | Out-File $Logpath -Append  

}

$SqlConnection.Close()

#$MyFileName = 'PMX_DSS_VOL_export.ps1'
#$filebase = Join-Path $PSScriptRoot $MyFileName
#Powershell.exe -File  $filebase

## - End of Script - ##