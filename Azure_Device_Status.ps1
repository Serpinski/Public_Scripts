# *** AZURE DEVICE STATUS ***
#
# DESCRIPTION:
#    Shows relevant output from dsregcmd /status command
#
# NOTES:
#    - NgcSet = If Windows Hello key is assigned to current user, then YES
#
# SCRIPT INFO:
#    Modified: Apr 2025
#    Created: Jaime
#

$ErrorActionPreference = "SilentlyContinue"

                                               # Initialization
$myStatus = $null
$myDate = get-date
$logFile = "c:\windows\dsregcmd.log"
$myDateFormatted = $myDate.year.tostring("0000")+"-"+$myDate.month.tostring("00")+"-"+$myDate.day.tostring("00")+" "+$myDate.hour.tostring("00")+":"+$myDate.minute.tostring("00")
$myStatus = @()

$info = dsregcmd /status

                                               # Dump to logfile for reference/troubleshooting
$info | out-file -filepath $logFile -Encoding:ASCII -Append:$false

                                               # Alternate save location
if((test-path -path $logFile) -ne $true)
{
   $info | out-file -filepath ($env:temp+"\dsregcmd.log") -Encoding:ASCII -Append:$false
}

foreach($line in $info)
{
   $line = $line.replace(" ","")
   if($line.tolower().contains("azureadjoined") -eq $true){$dsAdJoined = $line.length -ge 3 -and $line.substring($line.length-3,3) -eq "YES"}
   if($line.tolower().contains("domainjoined") -eq $true){$dsDomainJoined = $line.length -ge 3 -and $line.substring($line.length-3,3) -eq "YES"}
   if($line.tolower().contains("tpmprotected") -eq $true){$dsTpm = $line.length -ge 3 -and $line.substring($line.length-3,3) -eq "YES"}
   if($line.tolower().contains("deviceauthstatus") -eq $true){$dsAuthStatus = $line.length -ge 7 -and $line.substring($line.length-7,7) -eq "SUCCESS"}
   if($line.tolower().contains("domainname") -eq $true){$dsDomain = $line.length -ge 4 -and $line.substring($line.length-4,4) -eq "MAHC"}
   if($line.tolower().contains("tenantname") -eq $true){$dsTenantName = $line.length -ge 17 -and $line.substring($line.length-17,17) -eq "MedicalAssociates"}
   if($line.tolower().contains("tenantid") -eq $true){$dsTenantId = $line.length -ge 36 -and $line.substring($line.length-36,36) -eq "d5033e81-d8e7-4063-8687-8cbdba75652f"}
   if($line.tolower().contains("ngcset") -eq $true){$dsNgcSet = $line.length -ge 3 -and $line.substring($line.length-3,3) -eq "YES"}
   if($line.tolower().contains("azureadprt:") -eq $true){$dsAzurePrt = $line.length -ge 3 -and $line.substring($line.length-3,3) -eq "YES"}
}

$myStatus += [PSCustomObject]@{
   Date = $myDateFormatted
   AzureAdJoined = $dsAdJoined
   DomainJoined = $dsDomainJoined
   TpmStatusEnabled = $dsTpm
   DeviceAuthSuccessful = $dsAuthStatus
   DomainNameMatch = $dsDomain
   TenantNameMatch = $dsTenantName
   TenantIDMatch = $dsTenantId
   NgcSet = $dsNgcSet
   AzureADPrtSet = $dsAzurePrt
}

if($myStatus -ne $null)
{
   $myStatus
}