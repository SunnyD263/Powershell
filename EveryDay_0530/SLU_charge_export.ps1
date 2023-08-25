$SQLServer = "10.58.164.129"
$SQLDBName = "SLU"
$uid ="sqczpmip1_app"
$pswd = 'lwjfh/ezt.34Hf'
$table="Charging_View"
$path="\\10.47.17.20\pmi-dbo\SLU-stáří\data\"
$file="SLU_charging_"
#$Updpath= '\\10.47.17.20\pmi-dbo\SQL_script\EveryDay_0530\Log\lastupd.txt'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\EveryDay_0530\Log\PMIdblog.txt'
$Date = Get-Date 
$month =$Date.adddays(-1).tostring("yyyy_MM")
## EMAIL ##
$recipients = 'jan.sonbol@kuehne-nagel.com'
$msg = ''


$FirstDay = get-date -format "dd"
if ($FirstDay -eq '03'){


$xlsFile = $path + $file + $month +'.csv' ;


$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$pswd;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlQuery = "SELECT *  FROM [dbo].[$table]"
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlConnection.open()
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$Row =$SqlAdapter.Fill($DataSet)

if ( $Row -ne 0){
##$SqlAdapter.Fill($DataSet) >$null | Out-Null
$SqlConnection.close()                


            if (Test-Path $xlsFile)
            {
                        $LogClause=  $Date.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
                        $LogString = $LogClause + $file + " Non_exported;" 
                        write-host $file $date.ToString('yyyy_MM_dd')'.csv Non-exported, file exist'
                        Write-output $LogString | Out-File $Logpath -Append    
            }
            else
            {
            $DataSet.Tables[0] | export-csv -Delimiter ';' -Path $xlsFile -NoTypeInformation -Encoding UTF8 
            $LogClause=  $Date.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $file + " export;" 
            write-host $file $date.ToString('yyyy_MM_dd')'.csv Exported'
            Write-output $LogString | Out-File $Logpath -Append  
            };
}
else{
            $LogClause=  $Date.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $table +" no data rows" 
            write-host $LogString
            Write-output $LogString | Out-File $Logpath -Append  

}

$SqlConnection.Close()



## - End of Script - ##




# EMAIL
###########################################################

Start-Process Outlook  -wait
$o = New-Object -ComObject Outlook.Application 
$mail = $o.CreateItem(0)
$mail.To = $recipients
$mail.Subject = 'SLU_charging_' + $month
$mail.Body = $msg

$mail.Attachments.Add($xlsFile)

$mail.Send()
$o.Quit()

            $LogClause=  $Date.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $file + $month +'.csv' 
            write-host $LogString
            Write-output $LogString | Out-File $Logpath -Append  

}
else{

write-host 'This code export only firts day of month'
}

#$MyFileName = "Ecomm_non_del_Export.ps1"
#$filebase = Join-Path $PSScriptRoot $MyFileName
#Powershell.exe -File  $filebase