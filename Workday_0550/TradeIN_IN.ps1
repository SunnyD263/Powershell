$localpath = '\\10.47.17.20\pmi-dbo\Data\PCIP\DPD\TRADEIN\IN\'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\PMIdblog.txt'
$endpath=  '\\10.47.17.20\pmi-dbo\Data\PCIP\DPD\TRADEIN\IN\imported\'
$errpath=  '\\10.47.17.20\pmi-dbo\Data\PCIP\DPD\TRADEIN\IN\problem\'
$SQLServer = "DENOTSQ161"
$SQLDBName = "DPD_DB"
$remotefilepath = "/TRADEIN/IN/"
$filemask = 'ExportTR_IN_'
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

Write-Host "Please wait while your file downloads"

#Function to get all files
function Get-FtpDir ($url,$credentials)
{
    $request = [System.Net.FtpWebRequest]::Create($url)
    $request.Credentials = $credentials
    $request.Method = [System.Net.WebRequestMethods+FTP]::ListDirectoryDetails
    
   
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



$url = New-Object System.Uri("ftp://$server/$remotefilepath/")

#List of all files on FTP-Server
$files = Get-FTPDir $url -credentials (New-Object System.Net.NetworkCredential($user, $pass)) 
foreach($file in $files)

{          $allrows = 0
   $ftpFile = '*' + $filemask + '*'
            if ($file -like $ftpFile)
            {

            $file = $file.replace('[FILE] <A HREF="','')
            $localfilename = $file.Substring( 0, $file.IndexOf('">'))
            $fileuri = New-Object System.Uri("ftp://$server/$remotefilepath/$localfilename")
            $localfilelocation = "$localpath$localfilename"
            $webclient = New-Object System.Net.WebClient
            $webclient.Credentials = New-Object System.Net.NetworkCredential($user, $pass)
            $webclient.DownloadFile($fileuri, $localfilelocation)
            $Datum = Get-Date 
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $localfilename + " Download;" 
            Write-output $LogString | Out-File $Logpath -Append    
            write-host $localfilename  ' Download'

            $ImportExcel = Import-Excel -Path $localfilelocation   -WorksheetName 'Lc_Export'

            Write-Host "Start importing to SQL"

            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw;"
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            
  
            $SumRow=0  
            $ImpError = 0   
             
                ForEach ($rec in $ImportExcel)
                            {
                            
                            $id=$rec.Reference
                            $Status= $rec.Status
                            $CdfCharger = $rec.CdfCharger
                            $CdfHolder = $rec.CdfHolder
                         
                            $SqlQuery="UPDATE [DPD_DB].[dbo].[TRADE_IN] SET [STATUS] = '$Status' ,[CdfCharger] =  '$CdfCharger' ,[CdfHolder] ='$CdfHolder'  WHERE REFERENCE = '$id'"
            
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
            
                            if ($ImpError -eq 0) 
                            { 
                            Move-Item -Path  $localfilelocation -Destination $endpath$file_name
                            }
                            else
                            {
                            Move-Item -Path  $localfilelocation -Destination $errpath$file_name
                            }
                            
                            $Datum = Get-Date 
                            $LogClause=  $Date.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
                            $LogString = $LogClause + $file_name + " updated to SQLdB " + $SumRow + " rows updated;" 
                            Write-output $LogString | Out-File $Logpath -Append
                            write-host $LogString
                 
                      
            }
            
$allrows= $allrows + $SumRow
}
            
            Write-Host 'Upadating finished.' $Allrows ' rows added to Db.'


            $ftprequest = [System.Net.FtpWebRequest]::create($fileuri)
            $ftprequest.Credentials = New-Object System.Net.NetworkCredential($user, $pass)
            $ftprequest.Proxy=$null
            try
            {
               $ftprequest.Method = [System.Net.WebRequestMethods+Ftp]::DeleteFile
               $ftprequest.GetResponse() | Out-Null

               Write-Host ("File {0} deleted." -f $fileuri)
            }
            catch
            {
                if ($_.Exception.InnerException.Response.StatusCode -eq 550)
                {
                    Write-Host ("File {0} does not exist." -f $fileuri)
                }
                else
                {
                    Write-Host $_.Exception.Message
                }
            }
        


#$MyFileName = "10145787115_POSM.ps1"
#$filebase = Join-Path $PSScriptRoot $MyFileName
#Powershell.exe -File  $filebase

