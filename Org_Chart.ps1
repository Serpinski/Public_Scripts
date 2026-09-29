# *** ORG CHART ***
#
# DESCRIPTION:
#    Drills down thru the manager listings in AD to flush out a list of users
#
# INPUTS:
#    - rootEmail = Email address of the root to search from
#    - showDepth = How many levels down from the root to show
#    - showNonManagers = Exclude showing direct reports who do not have their own direct reports
#
# SCRIPT INFO:
#    Modified: Sept 2026
#    Created: Jaime
#

                              # Parameter list
Param(
  [string]$rootEmail,
  [int]$showDepth,
  [bool]$showNonManagers
)

$ErrorActionPreference = 'SilentlyContinue'

                              # Get a count of number of reports of a person (total tree)
function Get-ReportCount($lPersonID)
{
   $myChildren = $global:directReports[[string]$lPersonID]
   if($myChildren -eq $null)
   {
      return 0
   }
   $count = $myChildren.count
   foreach($child in $myChildren)
   {
      $count += Get-ReportCount $child.PersonID
   }
   return $count
}

                              # Recursively shows organizational chart (up to depth Level)
function Show-OrgChart($lPerson, $lPrefix="", $lIsLast=$true)
{
                              # Generate connector lines
   if($lPrefix -eq "")
   {
      $connector = ""
   }
   elseif($lIsLast)
   {
      $connector = "$global:dCorner$global:dLine$global:dLine "
   }
   else
   {
      $connector = "$global:dTee$global:dLine$global:dLine "
   }
                              # Only show up to depth level
   if([int]($lPerson.ManagerLevel) -le $showDepth)
   {
      $numReports = Get-ReportCount $lPerson.PersonID
      if($numReports -gt 0)
      {
         write-host ($lPrefix+$connector+($lPerson.DisplayName)+" ["+$numReports+" employees]")
      }
      else
      {
         if($showNonManagers -eq $true)
         {
            write-host ($lPrefix+$connector+($lPerson.DisplayName))
         }
      }
   }
                              # Generate child list and sort by display name
   $children = $global:directReports[[string]$lPerson.PersonID] | sort-object displayName
                              # Cut off blank leaf nodes
   if($null -eq $children -or $children.Count -eq 0)
   {
      return
   }
   for($i = 0;$i -lt $children.count;$i++)
   {
      $child = $children[$i]
      if($null -eq $child)
      {
                              # Sanity check
         continue
      }
      if($lPrefix -eq "")
      {
                              # Root level spacing but no vertical bar
         $childPrefix = "    "
      }
      elseif($lIsLast)
      {
         $childPrefix = "$lPrefix    "
      }
      else
      {
         $childPrefix = "$lPrefix$global:dVert   "
      }
                              # Recursively call
      Show-OrgChart $child $childPrefix ($i -eq ($children.Count - 1))
   }
}

                              # MAIN SCRIPT ----------------------------------------------------
cls
write-host "ORG CHART Script:" -foregroundcolor Green
write-host " "

                              # Show all if depth not specified
if($showDepth -eq "")
{
   [int]$showDepth = 999
}

                              # Special characters to show
$global:dCorner = [char]0x2514
$global:dLine   = [char]0x2500
$global:dTee    = [char]0x251C
$global:dVert   = [char]0x2502

                              # Create working array lists
$peopleList = new-object System.Collections.ArrayList

                              # Queue to watch for loop detection
$loopQueue = new-object System.Collections.ArrayList

                              # Add first person into array list from root email
                              # Second number is the persons generated (PID = person identifier), starting with 0
                              # Third number is the person they report to
                              # Fourth number is the org level of the person (1=top)
                              # Fifth is the display name of the person (*** DO NOT ADD ANOTHER FIELD AFTER THIS ONE DUE TO COMMAS ***)
$myList = get-aduser -filter {mail -eq $rootEmail} -properties mail,displayname
$rootDisplay = $myList.displayname.tostring().replace(",","")
$myQueue = new-object System.Collections.ArrayList
foreach($person in $myList)
{
   if($person.mail -ne "")
   {
      $ret = $loopQueue.add($person.mail.tolower())
      $ret = $myQueue.add(($person.mail.tolower()+",1,-1,1,"+($person.displayname.tostring().replace(",",""))))
   }
}

                              # Maximum person ID number, starting at 1
$pidMax = 1

                              # Parse thru list until all leaf nodes are flushed out
write-host "Generating..." -nonewline
$count = 0
do
{
   $count = $count + 1
   if($count/10 -eq [int]($count/10))
   {
      write-host "." -nonewline
   }
   if($count/25 -eq [int]($count/25))
   {
      write-host ("[q:"+$myQueue.count+"/p:"+$loopQueue.count+"]") -nonewline
   }

   $person,$myPID,$myReportTo,$myLevel,$myDisplay = $myQueue[0].split(",")
   $me = get-aduser -filter {mail -eq $person} -properties mail,manager,displayname

                              # Determine if person has any people who report to them in AD
   $myReports = $null
   $myDN = $me.distinguishedname.tolower()
   $myReports = get-aduser -filter {manager -eq $myDN} -properties mail,displayname

                              # Is this a manager?
   if($myReports -ne $null)
   {
      if($person -ne $rootEmail)
      {
         $zTmp = $me.mail.tolower() + ","+$myPID+","+$myReportTo+","+$myLevel+","+$me.displayname.tostring().replace(",","")
         $ret = $peopleList.add($zTmp)
      }
                              # Add all direct reports to queue so that they can be explored as well
      foreach($record in $myReports)
      {
         if($loopQueue.contains($record.mail.tolower()) -eq $false)
         {
            $pidMax= $pidMax + 1
            $zTmp = $record.mail.tolower() + ","+$pidMax+","+$myPID+","+[string](([int]$myLevel)+1)+","+$record.displayname.tostring().replace(",","")
            $ret = $myQueue.add($zTmp)
            $ret = $loopQueue.add($record.mail.tolower())
         }
      }
   }
   else
   {
      $zTmp = $me.mail.tolower() + ","+$myPID+","+$myReportTo+","+$myLevel+","+$me.displayname.tostring().replace(",","")
      $ret = $peopleList.add($zTmp)
   }
   $ret = $myQueue.remove($myQueue[0])
}
while($myQueue.Count -gt 0)
write-host " "
write-host " "

                              # Reformat generated data into an object array
$people = @()

                              # Root entry
$people += [PSCustomObject]@{
   Email = $rootEmail
   PersonID = 1
   ManagerPersonID = $null
   ManagerLevel = 1
   DisplayName = $rootDisplay
}
                              # Everyone under root entry
foreach($rec in $peopleList)
{
   $rPerson,$rID,$rManager,$rMgrLevel,$rDisplay = $rec.split(",")
   $people += [PSCustomObject]@{
      Email = $rPerson
      PersonID = $rID
      ManagerPersonID = $rManager
      ManagerLevel = $rMgrLevel
      DisplayName = $rDisplay
   }
}

                              # Index Person IDs
$personLookup = @()
$people | foreach-object {$personLookup[$_.PersonID] = $_ }

                              # Sort by display name
$people = $people | sort-object $_.DisplayName

                              # Group direct reports by manager
$global:directReports = @{}
foreach($person in $people)
{
   if($person.ManagerPersonID)
   {
      if(-not $global:directReports.ContainsKey($person.ManagerPersonId))
      {
         $global:directReports[$person.ManagerPersonId] = [System.Collections.ArrayList]::new()
      }
      [void]$global:directReports[$person.ManagerPersonID].Add($person)
   }
}

                              # Find the root and start the org chart
$roots = @($people | where-object { $_.ManagerPersonID -eq $null })
foreach($root in $roots)
{
   Show-OrgChart $root
}