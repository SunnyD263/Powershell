$localpath = '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\'
$endpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\imported\'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\EveryDay_0530\Log\PMIdblog.txt'
$localfilename = 'PMI_ORDITEM_PD_cntf.txt'
$SQLDBName = "DPD_DB"
$table="OrderItems"
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
                    $ORDTyP=$rec.KMORTY.Trim()
                    $Material=$rec.KPARTC.Trim()
                    $Quantity=$rec.KPAQST
                    $Codentify=$rec.Codentify.Trim()
                    $Task=$rec.KPCASE.Trim()
    
                     if (-not ([string]::IsNullOrEmpty( $REFERENCE))) {                                             
                    
    
                        $SqlQuery="SET ANSI_NULLS OFF; `
                        INSERT INTO $table ([Reference],[ORDTyp],[Material],[Quantity],[Codentify],[Task])`
                        VALUES  ('$REFERENCE','$ORDTyp','$Material',$Quantity,'$Codentify','$Task')"
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

 $SqlQuery = "WITH CTE AS (SELECT [REFERENCE],[Material],[ORDTyp],[Codentify],[Quantity],[Task], ROW_NUMBER() OVER (PARTITION BY [Task],[Codentify]  ORDER BY [Task] desc) row_num FROM $table) delete FROM CTE WHERE row_num > 1 and left(ORDTyp,2)='PD'"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
 $SqlAdapter.SelectCommand = $SqlCmd
 $DataSet = New-Object System.Data.DataSet
 $Row =$SqlAdapter.Fill($DataSet)
$SqlConnection.Close()

Write-Host "Delete duplicates finnished"



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
$LogString = $LogClause + $file_name + "; Db have " + $Row + " rows; OrdItems_PD_cntf Import job done"
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

 $MyFileName = "SWAP_reference.ps1"
 $filebase = Join-Path $PSScriptRoot $MyFileName

Powershell.exe -File  $filebase