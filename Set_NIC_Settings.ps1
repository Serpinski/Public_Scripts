# *** SET NIC SETTINGS ***
#
# DESCRIPTION:
#    Sets NIC Settings from registry (aka advanced settings not via GPO)
#
# PARAMETERS:
#    pNIC = NIC card to search for.  Wildcards are allowed, like "*8260*"
#    pOption = Option to change.  Uses the shortname, like IEEE11nMode instead of "HT Mode"
#    pValue = Value to change to.  Used the friendly name, like "VHT Mode"
#    DisplayOnly = Do not make changes
#    DumpSettings = Display ALL FILTERED settings [DEFAULT]
#
# SCRIPT INFO:
#    Modified: July 2026
#    Created: Jaime
#

                              # Parameter list
Param(
  [string]$pNIC,
  [string]$pOption,
  [string]$pValue,
  [bool]$DisplayOnly
)

$ErrorActionPreference = "SilentlyContinue"

function ObjectDump
{
   $debugMe = ""
   $theStatus = @()
   $expectedMatches = ((get-netadapter | where {$_.status -eq "Up" -and $_.interfacedescription -notlike "*Virtual*"}) | measure-object).count
   $matchCount = 0
   $myWiredCards = (get-netadapter | where {$_.status -eq "Up" -and $_.mediatype -like "*802.3*" -and $_.interfacedescription -notlike "*Virtual*"} | select interfaceguid).interfaceguid
   $myWirelessCards = (get-netadapter | where {$_.status -eq "Up" -and $_.mediatype -like "*802.11*" -and $_.interfacedescription -notlike "*Virtual*"} | select interfaceguid).interfaceguid
   $myDrivers = get-wmiobject Win32_PnPSignedDriver | where {$_.deviceclass -eq "NET"} | select DeviceName,DriverVersion,DriverDate
   $myReg = get-childitem -path "HKLM:\System\CurrentControlSet\Control\Class"
   $goodKey = $null
   foreach($subKey in $myReg)
   {
      if($subKey.name -notlike "*backup*")
      {
         $myKey = [string]'HKLM:\'+$subKey.name.substring(19,$subKey.name.length-19)
         if(((get-itemproperty -path $myKey)."Class") -eq "Net")
         {
            $goodKey = $myKey
            $debugMe = $debugMe + "/goodkey/"
         }
      }
   }
   $newReg = get-childitem -path $goodKey -erroraction:silentlycontinue
   foreach($instance in $newReg)
   {
      $convertName = $instance.name.substring($instance.name.length-4,4)
      $myNIC = get-itemproperty -path ($goodKey+"\"+$convertName)
      $myGUID = $null
      [string]$myGUID = $myNIC."NetCfgInstanceId"
      if($myGUID -ne $null)
      {
         $debugMe = $debugMe + "/guid/"
      }
      $foundMatch = $false
      foreach($entry in $myWiredCards)
      {
         if($entry -eq $myGUID){$foundMatch = $true}
      }
      foreach($entry in $myWirelessCards)
      {
         if($entry -eq $myGUID){$foundMatch = $true}
      }
      if($foundMatch -eq $true -and $myGUID -ne $null)
      {
         $debugMe = $debugMe + "/match/"
         if($myWirelessCards.contains($myGUID) -eq $true)
         {
            $theStatus += [PSCustomObject]@{
               Parameter = "WIRELESS CARD"
               Value = $myNIC."DriverDesc"
            }
            $theStatus += [PSCustomObject]@{
               Parameter = "WIRELESS DRIVER VERSION"
               Value = ($myDrivers | where {$_.devicename -eq $myNIC."DriverDesc"}).driverversion
            }
            $theStatus += [PSCustomObject]@{
               Parameter = "WIRELESS DRIVER DATE"
               Value = ($myDrivers | where {$_.devicename -eq $myNIC."DriverDesc"}).driverdate
            }
            $matchCount = $matchCount + 1
         }
         if($myWiredCards.contains($myGUID) -eq $true)         {
            $theStatus += [PSCustomObject]@{
               Parameter = "ETHERNET CARD"
               Value = $myNIC."DriverDesc"
            }
            $theStatus += [PSCustomObject]@{
               Parameter = "ETHERNET DRIVER VERSION"
               Value = ($myDrivers | where {$_.devicename -eq $myNIC."DriverDesc"}).driverversion
            }
            $theStatus += [PSCustomObject]@{
               Parameter = "ETHERNET DRIVER DATE"
               Value = ($myDrivers | where {$_.devicename -eq $myNIC."DriverDesc"}).driverdate
            }
            $matchCount = $matchCount + 1
         }
         $paramList = $goodKey+"\"+$convertName+"\Ndi\Params"
         $paramNames = get-childitem $paramList -erroraction:silentlycontinue
         foreach($entry in $paramNames)
         {
            [string]$myParam = $entry.name.split("\")[-1]
            $myValue = $myNic.$myParam
            $friendlyEnumValueKey = $goodKey+"\"+$convertName+"\Ndi\Params\"+$myParam+"\enum"
            $friendlyNames = get-itemproperty -path $friendlyEnumValueKey
            [string]$myFriendlyName = $friendlyNames.$myValue
            $theStatus += [PSCustomObject]@{
               Parameter = $myParam
               Value = $myFriendlyName
            }
         }
      }
   }
   if($expectedMatches -ne $matchCount)
   {
      $theStatus += [PSCustomObject]@{
         Parameter = "ERROR"
         Value = ("[Debug][E:"+(($myWiredCards | measure-object).count)+":"+$myWiredCards+"][W:"+(($myWirelessCards | measure-object).count)+":"+$myWirelessCards+"][M:"+$matchCount+"]["+$debugMe+"]")
      }
   }
   return $theStatus
}

if($pNIC -eq $null -or $pNIC -eq "")
{
   $myStatus = ObjectDump
   $myStatus
   exit
}

cls
Write-Host "SET NIC SETTINGS Script:" -foregroundcolor Green
Write-Host " "

                              # Display parameter info
write-host "NIC Query: " -nonewline
write-host $pNIC -foregroundcolor Green
write-host "Option: " -nonewline
write-host $pOption -foregroundcolor Green
write-host "Set Value: " -nonewline
write-host $pValue -foregroundcolor Green
if($displayOnly -eq $true)
{
   write-host "Operation Mode = DisplayOnly" -foregroundcolor Green
}
write-host " "

                              # Find the correct REGISTRY KEY
$myReg = get-childitem -path "HKLM:\System\CurrentControlSet\Control\Class"
$goodKey = $null
foreach($subKey in $myReg)
{
                              # Ignore the backup registry keys in hive as they may be duplicates of active NIC ones
   if($subKey.name -notlike "*backup*")
   {
                              # Reformat the stored registry key names to the HKLM format that PowerShell needs to use (aka HKEY_LOCAL_MACHINE to HKLM)
      $myKey = [string]'HKLM:\'+$subKey.name.substring(19,$subKey.name.length-19)
                              # Search for the NET one to find the correct GUID
      if(((get-itemproperty -path $myKey)."Class") -eq "Net")
      {
         $goodKey = $myKey
      }
   }
}
write-host "Dynamic NIC key = "$goodKey -foregroundcolor Gray
write-host " "
   
write-host "Searching NIC Drivers..."
$newReg = get-childitem -path $goodKey -erroraction:silentlycontinue
                              # Parse thru all the adapters
$found = $false
$changed = $false
foreach($instance in $newReg)
{
                              # All of the instances start with 0000, 0001, 0002, etc.
   $convertName = $instance.name.substring($instance.name.length-4,4)
   $myNIC = get-itemproperty -path ($goodKey+"\"+$convertName)
                              # Get the driver friendly name
   $myDESC = $null
   $myDESC = $myNIC."DriverDesc"
                              # Does it match what we queried for?
   if($myDESC -like $pNIC -and $myDESC -ne $null)
   {
      $found = $true
      write-host $myDESC -foregroundcolor Cyan -nonewline
                              # Convert the current setting to the FRIENDLY NAME of the setting's value (aka convert 2 to 'HT Mode', etc.)
      $descPath = $goodKey+"\"+$convertName+"\Ndi\Params\"+$pOption+"\enum"
      $results = $null
      $results = ((get-itemproperty -path $descPath).[string]($myNIC.$pOption))
      if($results -eq $null)
      {
         $results = ""
      }
      write-host " -- " -nonewline
                              # Does the current setting match the desired setting?
      if($results.contains($pValue) -eq $true)
      {
         write-host $results -nonewline -foregroundcolor Green
         if($displayOnly -ne $true)
         {
            write-host " already set" -nonewline -foregroundcolor Green
         }
         write-host " "
      }
      else
      {
                              # Setting does not match, so figure out what the correct setting should be
         $newValue = $null
         $tmpReg = get-itemproperty -path $descPath
                              # Generic script section.  Parses thru first 25 listed options, searching for match.  Not the best approach, but easiest to code
         for($i=0;$i -lt 25;$i++)
         {
            $seek = $tmpReg.([string]$i)
                              # Does the FRIENDLY VALUE match the desired FRIENDLY NAME passed?
            if($seek.contains($pValue) -eq $true)
            {
                              # Yes, so save the index number
               $newValue = $i
            }
         }
                              # Did we find a match in the indexes?
         if($newValue -ne $null)
         {
            write-host $results -foregroundcolor Yellow -nonewline
            if($displayOnly -ne $true)
            {
                              # Set the registry key to the updated index value
               set-itemproperty -Path ($goodKey+"\"+$convertName) -Name $pOption -Value $newValue
               write-host " (Changed to "$pValue")"
               $changed = $true
            }
            else
            {
               write-host " "
            }
         }
         else
         {
            write-host "Option value not found"
         }
      }
   }
}

if($found -eq $false)
{
   write-host "No matches found" -foregroundcolor Yellow
}

Write-Host " "
Write-Host "OPERATION COMPLETED" -foregroundcolor Green
