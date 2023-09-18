# *** SQL QUERY TO CSV ***
#
# DESCRIPTION:
#    Performs query against SQL database and exports to a CSV file
#
# PARAMETERS:
#    SQLServer = Server to connect to (FQDN is allowed)
#    SQLPort = Port to connect to if using a non-standard port (NOTE: Not setup for integrated security connections)
#    SQLDatabase = Default database of query
#    InputFile = SQL Query File
#    OutputFile = CSV File to Export to
#    KeepHeader = True/False to keep the first header row in output file
#    SQLAccount = Enter SQL account information, if not using Windows Authentication
#    SQLPasswordFile = Enter path to secure PASSWORD FILE for SQL account, if using SQL authentication [NOTE: AD account used to retrieve *must* match AD account under which the password file was generated]
#    USETabs = Create TAB delimited, instead of COMMA
#    USEVerticalBar = Create VERTICAL BAR delimited, instead of COMMA
#    RemoveQuotes = Removes quote characters from output
#    ReplaceChar = Replace X character with Y character in final output
#    DebugMode = Output further details for troubleshooting purposes
#    Silent = Turn on to stop all screen output
#    NoLog = Do not log summary results
#    Encoding = ASCII (default) or other format like UTF8
#    
# INFO:
#    Uses a 10-minute query timeout
#    
# USES:
#    Temporary working file in current directory (needs read/write/delete access)
#    Windows PowerShell v5 (part of Windows Managment Framework v5.1) for -NoNewLine option
#
# SCRIPT INFO:
#    Modified: Dec 2022
#    Created: Jaime
#

                              # Parameter list
Param(
  [string]$SQLServer,
  [string]$SQLPort,
  [string]$SQLDatabase,
  [string]$InputFile,
  [string]$OutputFile,
  [string]$KeepHeader,
  [string]$SQLAccount,
  [string]$SQLPasswordFile,
  [bool]$USETabs,
  [bool]$USEVerticalBar,
  [bool]$RemoveQuotes,
  [bool]$DebugMode,
  [bool]$Silent,
  [bool]$NoLog,
  [string]$ReplaceChar,
  [string]$Encoding
)

if($DebugMode -ne $true){$ErrorActionPreference = "SilentlyContinue"}

if($Encoding -eq ''){$Encoding = "ASCII"}

if($Silent -ne $true)
{
   cls
   Write-Host "SQL QUERY TO CSV FILE Script:" -foregroundcolor Green
   Write-Host "By: Jaime" -foregroundcolor Green
   Write-Host " "
}

if($SQLServer -eq "")
{
if($Silent -ne $true)
   {   
      write-host "You need to specify the SQL server." -foregroundcolor Red
   }
   exit
}

if($SQLDatabase -eq "")
{
   if($Silent -ne $true)
   {
      write-host "You need to specify the SQL database." -foregroundcolor Red
   }
   exit
}

if($InputFile -eq "")
{
   if($Silent -ne $true)
   {
      write-host "You need to specify an input query file." -foregroundcolor Red
   }  
   exit
}

if($OutputFile -eq "")
{
   if($Silent -ne $true)
   {
      write-host "You need to specify an output file." -foregroundcolor Red
   }
   exit
}

if($SQLPasswordFile -ne "")
{
   if($Silent -ne $true)
   {
      write-host "Reading in secure password..." -nonewline
   }
   $securePassword = Get-Content $SQLPasswordFile | ConvertTo-SecureString
   $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
   $myPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
   if($Silent -ne $true)
   {
      write-host "Done" -foregroundcolor Green
      write-host " "
   }
}

if($Silent -ne $true)
{
   write-host "SQL Server = "$SQLServer
   write-host "SQL Database = "$SQLDatabase
}
if($KeepHeader -eq $true)
{
   if($Silent -ne $true)
   {
      write-host "OPTION = Keep header row enabled" -foregroundcolor Yellow
   }
}
if($Silent -ne $true)
{
   write-host " "
}

                              # Input data
if($Silent -ne $true)
{
   write-host "Opening: "$InputFile"..." -nonewline
}
$file_data = get-content $InputFile
if($DebugMode -eq $true){if($Silent -ne $true){write-host " ";write-host $file_data -foregroundcolor Gray}}
$my_query = ""
$info = $file_data | Measure-Object
for ($i=0; $i -le $info.count; $i++)
{
   $my_query = $my_query + $file_data[$i] + "`n"
}
if($Silent -ne $true)
{
   write-host "Done" -foregroundcolor Green
}

                              # Log file
$myLog = $MyInvocation.MyCommand.Path.tostring().replace(".ps1",".log")
$oldLog = "SQL-Query-To-CSV.bak"
if((test-path -Path $myLog) -eq $true)
{
   if((get-item $myLog).length -ge 2*1024*1024)
   {
      if((test-path -Path $oldLog) -eq $true)
      {
         remove-item -path $oldLog
      }
      rename-item -path $myLog -newname $oldLog
   }
}

                              # Temporary Working File
$temp = ".\temp"+((Get-Random -Minimum 1 -Maximum 99999999).tostring())+".tmp"
                              # Get SQL Connection Setup
$Database = $SQLDatabase
$Server = $SQLServer
                              # Export File
$AttachmentPath = $OutputFile
                              # Remove any previous file
if((test-path $AttachmentPath) -eq $true)
{
   remove-item $AttachmentPath
}

                              # Connect to SQL and query 
if($Silent -ne $true)
{
   write-host "Running SQL query..." -nonewline
}
$SqlQuery = $my_query
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
                              # Are we using a SQL account or Integrated Windows Authentication?
if($SQLAccount -ne "")
{
   $SqlConnection.ConnectionString = "Data Source=$Server;Initial Catalog=$Database;User Id=$SQLAccount;Password=$myPassword"
   if($SQLPort -ne "")
   {
      [string]$mySource = $Server+","+$SQLPort
      $SqlConnection.ConnectionString = "Data Source=$mySource;Initial Catalog=$Database;User Id=$SQLAccount;Password=$myPassword"
   }
}
else
{
   $SqlConnection.ConnectionString = "Data Source=$Server;Initial Catalog=$Database;Integrated Security = True"
}

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlCmd.CommandTimeout = 600
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
                              # Store query output
$DataSet = New-Object System.Data.DataSet
$nRecs = $SqlAdapter.Fill($DataSet)
$nRecs | Out-Null
                              # Populate Hash Table
$objTable = $DataSet.Tables[0]
                              # Export Hash Table to (Temp) CSV File
$objTable | Export-CSV $temp
if($Silent -ne $true)
{
   write-host "Done" -foregroundcolor Green
   write-host " " 
}

                              # Reprocess the temp file to get rid of header lines (top 2)
$file_data = get-content $temp
$info = $file_data | Measure-Object

                              # Keep an extra line if the KeepHeader flag is set to true
if($Silent -ne $true)
{
   write-host "Line Returned = " -nonewline
}
$start = 2
if($KeepHeader -eq $true)
{
   $start = 1
}
if($Silent -ne $true)
{
   write-host ($info.count-$start) -foregroundcolor Green
   write-host " "
}

                              # Log activity, unless requested not to (temp queries, etc.)
if($NoLog -ne $true)
{
   (((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : ["+$InputFile+"] -- "+[string]($info.count-$start)+" records") | out-file -FilePath $myLog -Encoding:UTF8 -Append:$true
}

$myFileContents = New-Object System.Collections.Generic.List[System.Object]
if($Silent -ne $true)
{
   write-host "Writing output to file..." -nonewline
}
$counter = 0
$myTotal = 0
$max_buffer = 20
$buffer = $null
$buffer_count = 0
for ($i=$start; $i -lt $info.count; $i++)
{
   $counter = $counter + 1
   if( ($counter/40) -eq ([math]::truncate($counter/40)) )
   {
      if($Silent -ne $true)
      {
         write-host "." -nonewline
      }
   }
   $myOut = $file_data[$i].tostring().trim()
   $myTotal = $myTotal + 1

                              # Ignore blank lines
   if($myOut -ne "" -and $myOut.length -gt 4)
   {
                              # Convert to TAB or VERTICAL BAR format, if requested
      if($USETabs -eq $true -or $USEVerticalBar -eq $true)
      {
         $buildOut = ""
         $inQuotes = $false
         $z = 0
         do
         {
            $letter = $myOut[$z]
            if($letter -eq '"')
            {
               if($inQuotes -eq $false){$inQuotes = $true}else{$inQuotes = $false}
            }
            if($inQuotes -eq $false -and $letter -eq ",")
            {
               if($USETabs -eq $true)
               {
                  $buildOut = $buildOut + "`t"   
               }
               if($USEVerticalBar -eq $true)
               {
                  $buildOut = $buildOut + "|"   
               }
            }
            if($inQuotes -eq $true -and $letter -eq ",")
            {
               $buildOut = $buildOut + ","   
            }
            if($letter -ne '"' -and $letter -ne ",")
            {
               $buildOut = $buildOut + $letter
            }
            $z += 1
         }
         while($z -lt $myOut.length)
         $myOut = $buildOut
      }
                             # Write all but the last line to file
      if($i -lt ($info.count-1))
      {
                             # Buffer output so that writes are handled in chunks of data for efficient write speeds
          if($buffer_count -le $max_buffer)
          {
             if($buffer_count -ne 0){$buffer += "`r`n"}
             $buffer += $myOut
             $buffer_count += 1
          }
          else
          {
             if($buffer_count -ne 0){$buffer += "`r`n"}
             $buffer += $myOut
             if($RemoveQuotes -eq $true)
             {
                $buffer = $buffer.replace('"','')
             }
             if($ReplaceChar -ne $null)
             {
                $rc1 = $ReplaceChar[0]
                $rc2 = $ReplaceChar[1]
                $buffer = $buffer.replace($rc1,$rc2)
             }
             $buffer | out-file -filepath $AttachmentPath -Encoding:$Encoding -Append
             $buffer = $null
             $buffer_count = 0
          }   
      }
   }
}

                             # Flush any output in the buffer
if($buffer_count -ne 0)
{
   if($RemoveQuotes -eq $true)
   {
      $buffer = $buffer.replace('"','')
   }
   if($ReplaceChar -ne $null)
   {
      $rc1 = $ReplaceChar[0]
      $rc2 = $ReplaceChar[1]
      $buffer = $buffer.replace($rc1,$rc2)
   }
   $buffer | out-file -filepath $AttachmentPath -Encoding:$Encoding -Append
}

                             # Write the last line using a different method to avoid an extra line at end of file
if($RemoveQuotes -eq $true)
{
   $myOut = $myOut.replace('"','')
}
$myOut | out-file -filepath $AttachmentPath -Encoding:$Encoding -Append -NoNewline
if($Silent -ne $true)
{
   write-host "Done" -foregroundcolor Green
   write-host " "
}

                             # Delete the temp file
remove-item $temp

                             # Confirm the record count
if($Silent -ne $true)
{
   write-host "Validating lines written: " -nonewline
}
$myConfirm = get-content $AttachmentPath
if(($myConfirm.count) -eq ($info.count-$start))
{
   if($Silent -ne $true)
   {
      write-host "OK" -foregroundcolor Green
   }
}
else
{
   if($Silent -ne $true)
   {
      write-host ("FAILED -- "+($myConfirm.count)+" lines written") -foregroundcolor Yellow
   }
}

if($Silent -ne $true)
{
   Write-Host " "
   Write-Host "OPERATION COMPLETED" -foregroundcolor Green
}

# *** END
