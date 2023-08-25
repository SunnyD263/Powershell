$localpath = '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\'
$endpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\imported\'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\EveryDay_0530\Log\PMIdblog.txt'
$localfilename = 'CZ_PMI_PrclNum_PD2.txt'
$SQLDBName = "DPD_DB"
$table="PD2"
$remotefilepath = ""
$Datum = Get-Date
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
        
            $fileuri = New-Object System.Uri("ftp://$server/$remotefilepath/$localfilename")
            $localfilelocation = $localpath + $localfilename
            write-host  $localfilelocation
            $webclient = New-Object System.Net.WebClient
            $webclient.Credentials = New-Object System.Net.NetworkCredential($user, $pass)
            $webclient.DownloadFile($fileuri, $localfilelocation)
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $localfilename + " Download;" 
            write-host $localfilename  ' Download'
            Write-output $LogString | Out-File $Logpath -Append           
       
Write-Host "Start importing to SQL"

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw ;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand

$file_name = $localfilename
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

                $REFERENCE=$rec.KMCLOR.Trim()
                $Parcel=$rec.ORREFN.Trim()

                $ChngDT = $rec.KMCRTD.Substring(0,4) + '-' + $rec.KMCRTD.Substring(4,2)  + '-' + $rec.KMCRTD.Substring(6,2)
                                    
		        if (-not ([string]::IsNullOrEmpty( $REFERENCE))) {               

                    $SqlQuery="SET ANSI_NULLS OFF; `
                    INSERT INTO $table (ParcelNO, REFERENCE, EVENT_DATE_TIME)`
                    VALUES  ('$Parcel', '$REFERENCE','$ChngDT')"
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

    
Write-Host "Start delete duplicates"

 $SqlQuery = " WITH CTE AS (SELECT [REFERENCE],[PARCELNO],ROW_NUMBER() OVER (PARTITION BY [PARCELNO] ORDER BY [ID] ASC) row_num FROM dbo.PD2) DELETE FROM CTE WHERE row_num > 1"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
 $SqlAdapter.SelectCommand = $SqlCmd
 $DataSet = New-Object System.Data.DataSet
 $Row =$SqlAdapter.Fill($DataSet)
$SqlConnection.Close()

Write-Host "Delete duplicates finnished"

Write-Host 'Importing finished.'$SumRow' rows added to Db.'       

 $SqlQuery = "SELECT * from PD2"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
 $SqlAdapter.SelectCommand = $SqlCmd
 $DataSet = New-Object System.Data.DataSet
 $Row =$SqlAdapter.Fill($DataSet)
$SqlConnection.Close()



$LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; " + $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
$LogString = $LogClause + $file_name + "; Db have " + $Row + " rows; PD2_scr Import job done"
Write-output $LogString | Out-File $Logpath -Append                
write-host $LogString   


# $SqlQuery = "Update [dbo].[Update] Set [Date]='$Datum' from [Update] where [ID] = 4"
 #$SqlCmd.CommandText = $SqlQuery
 #$SqlCmd.Connection = $SqlConnection
#$SqlConnection.Close()
#}
#else{
#Write-host 'Today`s data has already been imported'
#}

 $MyFileName = "PMI_OrdItems_PB4_cntf.ps1"
 $filebase = Join-Path $PSScriptRoot $MyFileName

Powershell.exe -File  $filebase