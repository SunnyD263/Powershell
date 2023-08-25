
$filemask =  'ShipNumbers_'
$localpath = '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\SWAP\'
$endpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\SWAP\imported\'
$errpath=  '\\10.47.17.20\pmi-dbo\data\PCIP\DPD\SWAP\problem\'
$Updpath= '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\lastupd.txt'
$Logpath = '\\10.47.17.20\pmi-dbo\SQL_script\Workday_0550\Log\PMIdblog.txt'
$LastUP = Get-Content -Path $Updpath
$SQLServer = "10.58.164.129"
$SQLDBName = "DPD_DB"
$uid ="sqczpmip1_app"
$passw = "lwjfh/ezt.34Hf"
$table="PD4"
$Day = New-TimeSpan -Start $Date  -End $LastUP
$Day= $Day.Days - 1

$files = Get-ChildItem $localpath$filemask*
Write-Host "Start importing to SQL"

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User Id=$uid;Password=$passw;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlConnection.open()       


$files = Get-ChildItem $localpath$filemask*
foreach ($file in $files) 
{


$allrows = 0
$file_name = $file.name
$check = Test-Path -Path $endpath$file_name -PathType Leaf
$SumRow=0  
$ImpError = 0
$NotToday=  get-date -uformat "%d_%m_%Y"

If($file_name -ne $filemask+$NotToday+'.xlsx') 
{
    if ($check -eq $true)
            {
            Remove-Item $file -Verbose
            $LogClause=  $Date.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
            $LogString = $LogClause + $file_name + " has already been imported to SQLdB;" 
            write-host $LogString
            }
       else 
            {

            $ExcelFile =  $File.FullName
            [System.Data.SqlClient.SqlConnection]::ClearAllPools()
            $Excel = New-Object -ComObject Excel.Application
            $Excel.Visible = $false
            $Excel.DisplayAlerts = $false

            $wb = $Excel.Workbooks.Open($ExcelFile)
            $ws = $wb.Worksheets.Item(1)
            $rowMax = ($ws.usedRange.rows).count
            $ws.columns(1).NumberFormat = "#"
            $ws.columns(2).NumberFormat = "#"
            $ws.Cells.Item(1,1)  = 'PARCELNO'
            $ws.Cells.Item(1,2) = 'PARCELNO_ST'
            $ws.Cells.Item(1,3) = 'EVENT_DATE_TIME'
            $ws.Cells.Item(1,4) = 'REFERENCE'
                
                 for($j=2; $j -le $rowMax-1; $j++)
                    {
                    $PARCELNO =$ws.cells.Item($j,1).text
                    $PARCELNO =  $PARCELNO.replace("	","")
                    $PARCELNO =  $PARCELNO.substring(0,14)
                    $PARCELNO_ST =[bigint] $ws.cells.Item($j,2).value2
                    $EVENT_DATE_TIME =$ws.cells.Item($j,3).value2
                    $REFERENCE = $ws.cells.Item($j,4).text     
                    		
                    $ChngDT = $EVENT_DATE_TIME.Substring(6,4)  + '-' + $EVENT_DATE_TIME.Substring(3,2)  +  '-' + $EVENT_DATE_TIME.Substring(0,2)
                    $SqlQuery="SET ANSI_NULLS OFF; `
		            INSERT INTO $table ([PARCELNO],[PARCELNO_ST],[EVENT_DATE_TIME],[REFERENCE])`
		            VALUES  ('$PARCELNO','$PARCELNO_ST', '$ChngDT','$REFERENCE')"
                    $SqlCmd.CommandText = $SqlQuery
                    $SqlCmd.Connection = $SqlConnection  
  
                               
                try
                    {
                    #$SqlCmd.ExecuteNonQuery() | Out-Null 
                    $SumRow = $SumRow +  [int]$SqlCmd.ExecuteNonQuery()  
                     }
                catch [system.exception]
                     {
                    $LogClause=  $Date.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "                    
			        $msg =$LogClause + $file_name  + " ; "  + $($_.Exception.Message) + ";"

                    If ($null -ne $_.Exception.Message)

                            {
                	    $msg | Out-File $Logpath -Append
                            $ImpError = 1
                            break;
                            }
   
                     }

                }
                $wb.close()
                 if ($ImpError -eq 0) 
                { 
                Move-Item -Path  $file -Destination $endpath$file_name
                }
                else
                {
                Move-Item -Path  $file -Destination $errpath$file_name
                }

              
                $LogClause=  $Date.ToString("dd/MM/yyyy HH:mm:ss")+ "; "+ $env:UserDomain + "\" + $env:UserName + "; "+ $env:ComputerName + "; "
                $LogString = $LogClause + $file_name + " imported to SQLdB " + $SumRow + " rows added;" 
                write-host $LogString
            
             }
    }
$allrows= $allrows + $SumRow
}
Write-Host "Start delete duplicates"

 $SqlQuery = " WITH CTE AS (SELECT [REFERENCE],ROW_NUMBER() OVER (PARTITION BY [REFERENCE] ORDER BY [EVENT_DATE_TIME] ASC) row_num FROM dbo.PD4) DELETE FROM CTE WHERE row_num > 1"
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
$LogString = $LogClause + "; Db have " + $Row + " rows;  Import job done"
Write-output $LogString | Out-File $Logpath -Append                
write-host $LogString   


 $MyFileName = "PD2_reference.ps1"
 $filebase = Join-Path $PSScriptRoot $MyFileName

Powershell.exe -File  $filebase