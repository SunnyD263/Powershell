$filemask =  'STATUSDATA_cz10145787_D'
$localpath = '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\'
$endpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\imported\'
$errpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\OUT\problem\'
$Updpath= '\\10.47.17.20\pmi-dbo\SQL_script\DPD_Import\LOG\lastupd.txt'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\DPD_Import\LOG\PMIdblog.txt'
$LastUP = Get-Content -Path $Updpath
$Datum = Get-Date 
$Day = New-TimeSpan -Start $Datum  -End $LastUP
$Day= $Day.Days - 1
$SQLDBName = "DPD_DB"
$table="PMIdb"

$SQLlogpath = "\\10.47.17.20\pmi-dbo\SQL_script\safe\sql.txt"
$SQLlog = Get-Content -Path $SQLlogpath
$SQLServer = $SQLlog[0]
$uid =$SQLlog[1]
$base64Encoded = $SQLlog[2]
$bytes = [System.Convert]::FromBase64String($base64Encoded)
$passw  = [System.Text.Encoding]::UTF8.GetString($bytes)

$DPDlogpath = "\\10.47.17.20\pmi-dbo\SQL_script\safe\dpd.txt"
$DPDlog = Get-Content -Path $DPDlogpath
$proxyAddress = $DPDlog[0]
$server = $DPDlog[1]
$user = $DPDlog[6]
$base64Encoded = $DPDlog[7]
$bytes = [System.Convert]::FromBase64String($base64Encoded)
$pass  = [System.Text.Encoding]::UTF8.GetString($bytes)

Write-Host "Please wait while your file downloads"

#Function to get all files
function Get-FtpDir ($url, $credentials)
{
    $request = [System.Net.FtpWebRequest]::Create($url)
    $request.Credentials = $credentials
    $request.Method = [System.Net.WebRequestMethods+FTP]::ListDirectoryDetails
    $request.Proxy = $proxy   
    $response = $request.GetResponse()
    $reader = New-Object IO.StreamReader $response.GetResponseStream()
    $readline = $reader.ReadLine()
    $output = New-Object System.Collections.Generic.List[System.Object]
    while ($null -ne $readline)
    {        
        $readline = $reader.ReadLine()
        $output.Add($readline)
   
    }
    $reader.Close()
    $response.Close()
    $output
}


$remotefilepath = ""
            $proxy = New-Object System.Net.WebProxy($proxyAddress)
            $proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials 

$url = New-Object System.Uri(“ftp://$server/$remotefilepath/”)

#List of all files on FTP-Server
$files = Get-FTPDir $url -credentials (New-Object System.Net.NetworkCredential($user, $pass))

foreach($file in $files)
{
    for ($i=$Day; $i -le 0;$i++) 
    {
    $ImpDate = $Datum.AddDays($i).ToString("yyyyMMdd")
    $ftpFile = '*' + $filemask + $ImpDate + '*'
        if ($file -like $ftpFile)

        {   
            $file = $file.replace('[FILE] <A HREF="','')
            $localfilename = $file.Substring( 0, $file.IndexOf('">'))
            $fileuri = New-Object System.Uri(“ftp://$server/$remotefilepath/$localfilename”)
            $localfilelocation = "$localpath$localfilename"

            $webclient = New-Object System.Net.WebClient

            $webclient.Credentials = New-Object System.Net.NetworkCredential($user, $pass)
                        $webclient.Proxy = $proxy
            $webclient.DownloadFile($fileuri, $localfilelocation)
            $Datum = Get-Date 
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $localfilename + " Download;" 
            write-host $localfilename  ' Download'
        }
    }
}

$files = Get-ChildItem $localpath$filemask*.sem | Remove-Item -Verbose
$files = Get-ChildItem $localpath$filemask*
Write-Host "Start importing to SQL"

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand

foreach ($file in $files) 
{
    $file_name = $file.name
    $check = Test-Path -Path $endpath\$file_name -PathType Leaf
    $SumRow=0  
    $ImpError = 0   
    if ($check -eq $true)
        {
        Remove-Item $file -Verbose
        $Datum = Get-Date 
        $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
        $LogString = $LogClause + $file_name + " has already been imported to SQLdB;" 
        write-host $LogString
        }
    else 
        {

        $data=Import-CSV -Path  $file -Delimiter ';'   
        [System.Data.SqlClient.SqlConnection]::ClearAllPools()

        ForEach ($rec in $data) 
            {
                If (($rec.SCAN_CODE -ne 07) -and ($rec.SCAN_CODE -ne 18))
                { 
                $ChngDT = $rec.EVENT_DATE_TIME.Substring(0,4) + '-' + $rec.EVENT_DATE_TIME.Substring(4,2) + '-' + $rec.EVENT_DATE_TIME.Substring(6,2) + ' ' + $rec.EVENT_DATE_TIME.Substring(8,2) + ':' + $rec.EVENT_DATE_TIME.Substring(10,2) + ':' + $rec.EVENT_DATE_TIME.Substring(12,2)
                $Reference= $rec.CUSTOMER_REFERENCE
                If ($Reference.length -eq 9 )
                    {
                    $ReferenceChk= $Reference.substring(0,1)
                    if ($ReferenceChk -eq 'v' -or  $ReferenceChk -eq 't' -or  $ReferenceChk -eq 'z')
                        {
                        $Reference= $Reference.substring(1,8)
                        } 
                   }                   
              elseif ($Reference.length -le 14)
                    {
                    $Reference= $Reference.substring(0,$Reference.length-1)                
                    }              
                else
                    {
                    $Reference= $Reference.substring(0,15) 
                    }
                $SqlQuery="SET ANSI_NULLS OFF; `
                INSERT INTO $table (PARCELNO, SCAN_CODE, EVENT_DATE_TIME, SERVICE, ZIP, REFERENCE, CUSTOMER, Source,KN) `
                VALUES  ('$($rec.PARCELNO)','$($rec.SCAN_CODE)','$ChngDT','$($rec.SERVICE)','$($rec.CONSIGNEE_ZIP)','$Reference','$($rec.RECEIVER_NAME)','DPD-POSM','Import')"

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
                    $Datum = Get-Date 
                    $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "                    
			        $msg =$LogClause + $file_name  + " ; "  + $($_.Exception.Message) + ";"

                    If ($null -ne $_.Exception.Message)

                            {
                	    $msg | Out-File $Logpath -Append
                            $ImpError = 1
                            break;
                            }
   
                    }
                finally
                    {
                    $SqlConnection.close() 
                    }
                }
            }
            
        if ($ImpError -eq 0) 
            { 
            Move-Item -Path  $file -Destination $endpath$file_name
            }
        else
            {
            Move-Item -Path  $file -Destination $errpath$file_name
            }                
               
        $Datum = Get-Date 
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $file_name + " imported to SQLdB " + $SumRow + " rows added;" 
            write-host $LogString
            Write-output $LogString | Out-File $Logpath -Append
	 
          
        }

$allrows= $allrows + $SumRow
}

Write-Host 'Importing finished.' $Allrows ' rows added to Db.'       

$Datum = Get-Date 
$LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; " + $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
$LogString = $LogClause + $file_name + "; Db have " + $Row + " rows;  Import job done"  
write-host $LogString   
 
$MyFileName = "10645787115_FC.ps1"
$filebase = Join-Path $PSScriptRoot $MyFileName
Powershell.exe -File  $filebase
  


