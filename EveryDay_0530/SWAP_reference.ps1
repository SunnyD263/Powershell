$filemask =  'parcel_list_'
$localpath = '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\SWAP\'
$endpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\SWAP\imported\'
$errpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\SWAP\problem\'
$Updpath= '\\10.47.17.20\pmi-dbo\SQL_script\EveryDay_0530\Log\lastupd.txt'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\EveryDay_0530\Log\PMIdblog.txt'
$LastUP = Get-Content -Path $Updpath
$SQLDBName = "DPD_DB"
$table="PD4"
$Date = Get-Date
$Datum = Get-Date 
$Day = New-TimeSpan -Start $Date  -End $LastUP
$Day= $Day.Days - 1

$SQLlogpath = "\\10.47.17.20\pmi-dbo\SQL_script\safe\sql.txt"

$SQLlog = Get-Content -Path $SQLlogpath
$SQLServer = $SQLlog[0]
$uid =$SQLlog[1]
$base64Encoded = $SQLlog[2]
$bytes = [System.Convert]::FromBase64String($base64Encoded)
$passw  = [System.Text.Encoding]::UTF8.GetString($bytes)

Write-Host "Start importing to SQL"
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlConnection.open()       

$files = Get-ChildItem $localpath$filemask*
foreach ($file in $files) 
{

$file_name = $file.name
$check = Test-Path -Path $endpath$file_name -PathType Leaf
$SumRow=0  
$ImpError = 0
$NotToday=  get-date -uformat "%d-%m-%Y"

If($file_name -ne $filemask+$NotToday+'.csv') 
     {
    if ($check -eq $true)
          {
            Remove-Item $file -Verbose
            $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $file_name+$Datum.ToString("dd/MM/yyyy")+ " has already been imported to SQLdB;" 
            write-host $LogString
          }
       else 
          {

          $data=Import-CSV -Path $localpath$file_name -Delimiter ';' 
          [System.Data.SqlClient.SqlConnection]::ClearAllPools()
          ForEach ($rec in $data) 
               {
                    $PARCELNO =  $rec.return
                    $PARCELNO_ST =$rec.forward
                    $REFERENCE = $rec.ORDER
                    $ChngDT =$rec.date.Substring(6,4) + '-'  + $rec.date.Substring(3,2) + '-' +  $rec.date.Substring(0,2) 

		        if (-not ([string]::IsNullOrEmpty( $REFERENCE))) {         

                    $SqlQuery="INSERT INTO $table ([PARCELNO],[PARCELNO_ST],[EVENT_DATE_TIME],[REFERENCE])`
		          VALUES  ('$PARCELNO','$PARCELNO_ST','$ChngDT','$REFERENCE')"
                    $SqlCmd.CommandText = $SqlQuery
                    $SqlCmd.Connection = $SqlConnection                                
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
                            $ImpError = 1
                            break;
                            }
                    }
   
                     }
                     

            
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
                $LogClause=  $Datum.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
                $LogString = $LogClause + $file_name+$Datum.ToString("dd/MM/yyyy") + " imported to SQLdB " + $SumRow + " rows added;" 
                write-host $LogString


}
$SqlConnection.close() 
Write-Host "Start delete duplicates"

 $SqlQuery = " WITH CTE AS (SELECT [REFERENCE],ROW_NUMBER() OVER (PARTITION BY [REFERENCE] ORDER BY [EVENT_DATE_TIME] desc) row_num FROM dbo.PD4) DELETE FROM CTE WHERE row_num > 1"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
 $SqlAdapter.SelectCommand = $SqlCmd
 $DataSet = New-Object System.Data.DataSet
 $Row =$SqlAdapter.Fill($DataSet)


Write-Host "Delete duplicates finnished"



Write-Host 'Importing finished.$SumRow rows added to Db.'       

 $SqlQuery = "SELECT * from PD4"
 $SqlCmd.CommandText = $SqlQuery
 $SqlCmd.Connection = $SqlConnection
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
 $SqlAdapter.SelectCommand = $SqlCmd
 $DataSet = New-Object System.Data.DataSet
 $Row =$SqlAdapter.Fill($DataSet)
$SqlConnection.Close()



$LogClause=  $Date.ToString("dd/MM/yyyy HH:mm:ss")+ "; " + $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
$LogString = $LogClause + "; Db have " + $Row + " rows; SWAP_reference Import job done"
Write-output $LogString | Out-File $Logpath -Append                
write-host $LogString   


#$MyFileName = "Trade_IN.ps1"
#$filebase = Join-Path $PSScriptRoot $MyFileName
#Powershell.exe -File  $filebase