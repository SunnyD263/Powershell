$SQLDBName = "Produktivita"
$table="DSSMITH_Time_View"
$path="\\10.47.17.20\pmi-dbo\PMX\Time\"
$file="PMX_Time_"
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\PMIdblog.txt'

$Datum = Get-Date

$SQLlogpath = "\\10.47.17.20\pmi-dbo\SQL_script\safe\sql.txt"

$SQLlog = Get-Content -Path $SQLlogpath
$SQLServer = $SQLlog[0]
$uid =$SQLlog[1]
$base64Encoded = $SQLlog[2]
$bytes = [System.Convert]::FromBase64String($base64Encoded)
$passw  = [System.Text.Encoding]::UTF8.GetString($bytes)


for ($var = -2; $var -gt -30; $var--){
$date = (Get-Date).AddDays($var)

#do
#{
#  $date = $date.AddDays(-2)
#} while ($date.DayOfWeek -eq [System.DayOfWeek]::Saturday -or 
#         $date.DayOfWeek -eq [System.DayOfWeek]::Sunday)

$xlsxFile = $path + $file +  $date.ToString('yyyy_MM_dd')+'.xlsx';
$SQL= "SELECT *  FROM [dbo].[DSSMITH_Time_View] WHERE convert(varchar,[Process Start Date],112) =" + $date.ToString('yyyyMMdd')

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw ;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
 $SqlQuery = "SELECT *  FROM [dbo].[DSSMITH_Time_View] WHERE convert(varchar,[Process Start Date],112) =" + $date.ToString('yyyyMMdd')
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
                        $LogString = $LogClause + $file+$date.ToString('yyyy_MM_dd')+'.xlsx Non-exported, file exist;'
                        write-host $file $date.ToString('yyyy_MM_dd')'.xlsx Non-exported, file exist'
                        Write-output $LogString | Out-File $Logpath -Append    
            }
            else
            {
                        $Sheet='Time'
                        $xlsx= Invoke-Sqlcmd -ServerInstance $SQLServer -Database $SQLDBName -user $uid -Password $passw  -Query $SQL |
                        Select-Object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors |
                        export-Excel -Path $xlsxFile -WorksheetName $Sheet -Autosize  -BoldTopRow –PassThru -NoNumberConversion '*'

                                    $xlsx.Workbook.Worksheets[$Sheet]
                                     # $ws.Dimension.Columns  #number of columns
                                     # $ws.Dimension.Rows     #number of rows
                                     Close-ExcelPackage $xlsx

            #$DataSet.Tables[0] | export-csv -Delimiter ";" -Path $xlsFile -NoTypeInformation -Encoding UTF8 
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $file+$Date.ToString('yyyy_MM_dd')+'.xlsx Exported;'
            write-host $file $Date.ToString('yyyy_MM_dd')'.xlsx Exported'
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
}

$MyFileName = "PMX_DSS_VOL_export.ps1"
$filebase = Join-Path $PSScriptRoot $MyFileName
Powershell.exe -File  $filebase

## - End of Script - ##