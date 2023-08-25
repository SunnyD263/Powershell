$localpath = '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\'
$endpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\imported\'
$Updpath= '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\lastupd.txt'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\PMIdblog.txt'
$localfilename = 'PMX_DSSMITH.txt'
$LastUP = Get-Content -Path $Updpath

$SQLDBName = "Produktivita"
$table="PMX_DSSMITH"
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
       




Write-Host "Start importing to SQL"

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand


$file_name = $file.name
$check = Test-Path -Path $endpath$file_name -PathType Leaf
$SumRow=0  

    if ($check -eq $true)
            {
            Remove-Item $file -Verbose
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $file_name+$Datum.ToString("dd/MM/yyyy")+ " has already been imported to SQLdB;" 
            write-host $LogString
            }
    else 
            {
              
            $data=Import-CSV -Path $localpath$localfilename -Delimiter '	' 
            [System.Data.SqlClient.SqlConnection]::ClearAllPools()
            ForEach ($rec in $data) 
                {
		    if (-not ([string]::IsNullOrEmpty($rec))) {        
                $Depo=$rec.Depo.Trim()
                $Client = $rec.Client.Trim()
                $Operator=$rec.Operator.Trim()
                $ChngDT = $rec.Date.Substring(6,4)   +   '-' + $rec.Date.Substring(3,2) + '-' + $rec.Date.Substring(0,2)
                $ChngTM=$rec.Time.Trim()
                $ORDTYP =$rec.OrderType.Trim()
                $MO =$rec.MovementOrder.Trim()
                $ZoneCode =$rec.ZoneCode.Trim()
                $PalletID =$rec.PalletID.Trim()
                $Qty =$rec.Qty.Trim()                

                $SqlQuery="SET ANSI_NULLS OFF; `
		        INSERT INTO $table ([Depo], [Client], [Operator], [Date], [Time],[Order type],[Movement Order],[ZoneCode],[Pallet ID],[Qty])`
		        VALUES  ('$Depo','$Client', '$Operator','$ChngDT','$ChngTM','$ORDTYP','$MO', '$ZoneCode', '$PalletID', '$Qty')"
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
			$msg =$LogClause + $file_name+$Datum.ToString("dd/MM/yyyy")+ " ; "  + $($_.Exception.Message) + ";"

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
                $LogString = $LogClause + $file_name+$Datum.ToString("dd/MM/yyyy")+ " imported to SQLdB " + $SumRow + " rows added;" 
                write-host $LogString
                Remove-Item $localpath$localfilename -Verbose	 
            
           }


 $SqlQuery = "SELECT * from PMX_DSSMITH"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
 $SqlAdapter.SelectCommand = $SqlCmd
 $DataSet = New-Object System.Data.DataSet
 $Row =$SqlAdapter.Fill($DataSet)

$LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; " + $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
$LogString = $LogClause + $file_name+$Datum.ToString("dd/MM/yyyy") + "; Db have " + $Row + " rows; PMX_scr Import job done"
Write-output $LogString | Out-File $Logpath -Append                
Write-host $LogString   
$SqlConnection.Close()


 $MyFileName = "PMX_DSS_Time_export.ps1"
 $filebase = Join-Path $PSScriptRoot $MyFileName

Powershell.exe -File  $filebase