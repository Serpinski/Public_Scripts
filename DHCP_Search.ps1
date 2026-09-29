# *** DHCP SEARCH ***
#
# DESCRIPTION:
#    Searches DHCP server for matching lease
#
# SCRIPT INFO:
#    Modified: Nov 2023
#    Created: Jaime
#

$ErrorActionPreference = "SilentlyContinue"

                              # Parameter list
Param(
  [string]$DHCPServer,
  [string]$byIP,
  [string]$byMAC,
  [string]$byName
)

                                           # Global notification variables
cls
Write-Host "DHCP SEARCH Script:" -foregroundcolor Green
Write-Host " "

if($byMAC -eq "")
{
   $byMAC = "dontfindthis"
}
else
{
   $byMAC = $byMAC.tolower().replace("-","").replace(":","")
}

if($byName -eq "")
{
   $byName = "dontfindthis"
}

write-host "Searching..."
write-host " "

$scopes = Get-DHCPServerv4scope -ComputerName $DHCPServer
foreach($entry in $scopes)
{
   $found = $false
   $matchRecord = $null
   $matchRecord = new-object System.Collections.ArrayList
   $leases = Get-DHCPServerv4Lease -ComputerName $DHCPServer -ScopeId $entry.scopeid
   foreach($record in $leases)
   {
      if($record.hostname -ne $null)
      {
         $rHost = $record.hostname.tolower()
      }
      else
      {
         $rHost = "blank"
      }
      if($record.clientid -ne $null)
      {
         $rClient = $record.clientid.tolower().replace("-","")
      }
      else
      {
         $rClient = "nonespecified"
      }
      if($record.ipaddress -eq $byIP -or $rClient.contains($byMAC.tolower()) -eq $true -or $rHost.contains($byName.tolower()) -eq $true)
      {
         $found = $true
         $matchRecord += $record
      }
   }
   if($found -eq $true)
   {
      write-host "DHCP Scope: " -nonewline
      write-host $entry.name -foregroundcolor Green
      $matchRecord
   }
}