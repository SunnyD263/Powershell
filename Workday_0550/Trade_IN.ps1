$localpath = '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\'
$endpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\imported\'
$Updpath= '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\lastupd.txt'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\PMIdblog.txt'
$localfilename = 'CZ_TRADEIN_ILU_Daily.txt'
$table="NonDlv_Dvc_View_export"
$LastUP = Get-Content -Path $Updpath
$SQLDBName = "DPD_DB"
$table="TRADE_IN"
$remotefilepath = ""
$Date = Get-Date
$Datum = Get-Date
$Day = New-TimeSpan -Start $Date  -End $LastUP
$Day= $Day.Days - 1

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

Write-Host "Please wait while your file downloads"
        
            $fileuri = New-Object System.Uri(“ftp://$server/$remotefilepath/$localfilename”)
            $localfilelocation = $localpath + $localfilename
            write-host  $localfilelocation
            $webclient = New-Object System.Net.WebClient
            $webclient.Credentials = New-Object System.Net.NetworkCredential($user, $pass)
            $webclient.DownloadFile($fileuri, $localfilelocation)
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $localfilename + " Download;" 
            write-host $localfilename  ' Download'
            Write-output $LogString | Out-File $Logpath -Append           
    
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw ;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand

$file_name = $file.name
$check = Test-Path -Path $endpath$file_name -PathType Leaf
$SumRow=0  

    if ($check -eq $true)
            {
            Remove-Item $file -Verbose
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $file_name + " has already been imported to SQLdB;" 
            write-host $LogString
            }
    else 
            {
              
            $data=Import-CSV -Path $localpath$localfilename -Delimiter '	' 

            [System.Data.SqlClient.SqlConnection]::ClearAllPools()
            ForEach ($rec in $data) 
                {
                $REFERENCE=$rec.KJCORF.Trim()
		        if (-not ([string]::IsNullOrEmpty( $REFERENCE)))
                {
                $CRTDate = $rec.KJCRTD.Substring(0,2) + $rec.KJCRTD.Substring(3,2)  +   '-' + $rec.KJCRTD.Substring(5,1) +  $rec.KJCRTD.Substring(7,1) + '-' + $rec.KJCRTD.Substring(8,2)
                $SHPDate = $rec.KMCDAT.Substring(0,2) + $rec.KMCDAT.Substring(3,2)  +   '-' + $rec.KMCDAT.Substring(5,1) +  $rec.KMCDAT.Substring(7,1) + '-' + $rec.KMCDAT.Substring(8,2)

                $SqlQuery="INSERT INTO $table (REFERENCE, CRTDate, SHPDate)`
		        VALUES  ('$REFERENCE','$CRTDate', '$SHPDate')"
                $SqlCmd.CommandText = $SqlQuery
                $SqlCmd.Connection = $SqlConnection  
                $SqlConnection.open()                

               
                try
                    {
                    #$SqlCmd.ExecuteNonQuery() | Out-Null 
                    $SumRow = $SumRow +  [int]$SqlCmd.ExecuteNonQuery()
                     }
                catch [system.exception]
                     {
                    $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "                    
			$msg =$LogClause + $file_name  + " ; "  + $($_.Exception.Message) + ";"

                    If ($null -ne $_.Exception.Message)

                            {
                	    $msg | Out-File $Logpath -Append
                            break;
                            }
   
                     }
                finally
                    {
                    $SqlConnection.close() 
                    }
                }
                }
              
                $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
                $LogString = $LogClause + $file_name + " imported to SQLdB " + $SumRow + " rows added;" 
                write-host $LogString
                Remove-Item $localpath$localfilename -Verbose	 
            
             }

If ($Sumrow -ne 0) 
{
Write-Host "Start delete duplicates"

 $SqlQuery = " WITH CTE AS (SELECT [REFERENCE],ROW_NUMBER() OVER (PARTITION BY [REFERENCE] ORDER BY [ID] ASC) row_num FROM dbo.[TRADE_IN]) DELETE FROM CTE WHERE row_num > 1"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
 $SqlAdapter.SelectCommand = $SqlCmd
 $DataSet = New-Object System.Data.DataSet
 $Row =$SqlAdapter.Fill($DataSet)
$SqlConnection.Close()


Write-Host "Delete duplicates finnished"
}


Write-Host 'Importing finished.'$SumRow' rows added to Db.'       

 $SqlQuery = "SELECT * from TRADE_IN"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
 $SqlAdapter.SelectCommand = $SqlCmd
 $DataSet = New-Object System.Data.DataSet
 $Row =$SqlAdapter.Fill($DataSet)
$SqlConnection.Close()



$LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; " + $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
$LogString = $LogClause + $file_name + "; Db have " + $Row + " rows; Trade_IN Import job done"
Write-output $LogString | Out-File $Logpath -Append                
write-host $LogString   




#$MyFileName = "Z_Liquid_report.ps1"
#$filebase = Join-Path $PSScriptRoot $MyFileName
#Powershell.exe -File  $filebase