$localpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_1000\'
$endpath=  '\\10.47.17.20\pmi-dbo\SQL_script\Workday_1000\imported\'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\WorkDay_1000\Log\PMIdblog.txt'
$localfilename = 'CZ_Detail_Stock Level_PMI NCI.txt'
$SQLDBName = "Liquid"
$table="NCI_Stock"
$remotefilepath = ""
$Datum = Get-Date
$Day= $Day.Days - 1

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

$SqlQuery = "DELETE FROM $table"
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$Row =$SqlAdapter.Fill($DataSet)
$SqlConnection.Close()

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
                    $Article=$rec.Article.Trim()
                    $Batch=$rec.Batch.Trim()
                    $Description=$rec.Description.Trim()
                    $Zone=$rec.KPWHZO.Trim()
                    $Location=$rec.KPWHLO.Trim()
                    $KJCORF=$rec.KJCORF.Trim()
                    $KJORDT=$rec.KJORDT.Trim()
                    $KPCASE=$rec.KPCASE.Trim()
                    $Status=$rec.Status.Trim()
                    $KPAVAL=$rec.KPAVAL.Trim()
                    $Quantity=$rec.Quantity.Trim()
    
                     if (-not ([string]::IsNullOrEmpty($Quantity))) {                                             
                    
    
                        $SqlQuery="SET ANSI_NULLS OFF; `
                        INSERT INTO $table ([Article],[Batch],[Description],[KPWHZO],[KPWHLO],[KJCORF],[KJORDT],[KPCASE],[Status],[KPAVAL],[Quantity])`
                        VALUES  ('$Article','$Batch','$Description','$Zone','$Location','$KJCORF','$KJORDT','$KPCASE','$Status',$KPAVAL,$Quantity)"
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

    

Write-Host 'Importing finished.'$SumRow' rows added to Db.'       

 $SqlQuery = "SELECT * from $table"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
 $SqlAdapter.SelectCommand = $SqlCmd
 $DataSet = New-Object System.Data.DataSet
 $Row =$SqlAdapter.Fill($DataSet)
$SqlConnection.Close()



$LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; " + $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
$LogString = $LogClause + $file_name + "; Db have " + $Row + " rows; OrdItems_PB4 Import job done"
Write-output $LogString | Out-File $Logpath -Append                
write-host $LogString   


