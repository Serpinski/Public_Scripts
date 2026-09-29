# *** ENUMERATE NTFS PERMISSIONS ***
#
# DESCRIPTION:
#    - Generates report of USER or GROUP access to FOLDER tree and outputs report
#
# PARAMETER:
#    searchObject - Object (user or group) to search for (jdoe, operators)
#    rootPath - UNC path to search for matching access rights
#    includeGlobalGroups - Includes (Domain Users) and (Everyone) access rights on report
#    reportName - Output CSV report to this file
#
# SCRIPT INFO:
#    Modified: Sept 2026
#    Created: Jaime
# 

                              # Parameter list
Param(
  [string]$searchObject,
  [string]$rootPath,
  [bool]$includeGlobalGroups,
  [string]$reportName
)

                              # Recursive traversal of folder path using multithreading
function getFolderTree($lPath)
{
   $rData = @($lPath)
                              # Clear jobs
   get-job | remove-job
                              # Script to call within jobs (recurse one folder)
   $scriptBlock =
      {
         param($sFolder,$dLog)
         $sFolder
         try
         {
            [System.IO.Directory]::EnumerateDirectories($sFolder,"*",[System.IO.SearchOption]::AllDirectories)
         }
         catch
         {
                              # Ignore if access is denied
            (((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : Read Error ["+$sFolder+"]") | out-file -FilePath $dLog -Encoding:UTF8 -Append:$true
         }
      }

   $jobNum = 0
   $activeThrottle = 12
   foreach($child in [System.IO.Directory]::EnumerateDirectories($lPath))
   {
      $jobNum = $jobNum + 1

      while((get-job -state Running).count -ge $activeThrottle)
      {
         start-sleep -milliseconds 250
      }
                              # One job per top-level folder
      $list = start-job -name ("enumFolders-"+$jobNum) -scriptblock $scriptBlock -argumentlist $child,$global:debugLog

                              # Show Status
      $completedThreads = (get-job -State Completed).count
      $totalThreads = (get-job).count
      $percent = [math]::Round(($completedThreads/$totalThreads)*100,2)
      write-progress -activity "Scanning Folders" -status ("Completed "+$completedThreads+" of "+$totalThreads+" threads") -percentcomplete $percent
   }
   do
   {
                              # Show Status
      $completedThreads = (get-job -State Completed).count
      $totalThreads = (get-job).count
      $percent = [math]::Round(($completedThreads/$totalThreads)*100,2)
      write-progress -activity "Scanning Folders" -status ("Completed "+$completedThreads+" of "+$totalThreads+" threads") -percentcomplete $percent

                              # Waiting for all jobs to finish
      start-sleep -milliseconds 500
   }
   while((get-job -State Running).count -gt 0)
   foreach($job in get-job)
   {
                              # Populate all job return data to singular variable
      $rData += receive-job $job
   }
                              # Job cleanup
   get-job | remove-job
   return $rData
}

                              # ----------------------------- MAIN SCRIPT
cls
write-host "ENUMERATE NTFS PERMISSIONS Script:" -foregroundcolor Green
write-host " "

                              # Debug log
$global:debugLog = ".\Enumerate_NTFS_Permissions.log"
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : ---------- Starting Script ----------") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true

if($searchObject -eq "")
{
   write-host "-searchObject = Object (user or group) to search for (jdoe, operators)"
   write-host "-rootPath - UNC path to search for matching access rights (\\vsfs5\data)"
   write-host "-includeGlobalGroups - Includes (Domain Users) and (Everyone) access rights on USER report"
   write-host "-reportName - Output CSV report to this file"
   exit
}

import-module ActiveDirectory

write-host "Credentials to use for search [Needs access to search folder tree]: " -foregroundcolor Yellow
$myUsername = read-host "Username [in DOMAIN\username format]"
$myPassword = read-host "Password" -AsSecureString
$myCred = New-Object System.Management.Automation.PSCredential ($myUsername, $myPassword)
write-host " "
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : Credentials Entered ["+$myUsername+"]") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : Root Path ["+$rootPath+"]") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true

                              # To determine total runtime
$myDate = get-date

                              # Determine type of object we are looking at (USER or GROUP)
$obj = Get-ADObject -LDAPFilter "(sAMAccountName=$searchObject)" -properties objectClass

if($obj.ObjectClass -eq "user")
{
   write-host "-----> Gathering AD membership for USER [" -nonewline
   write-host $searchObject -foregroundcolor Cyan -nonewline
   write-host "]..."
   if($includeGlobalGroups -eq $true)
   {
      write-host ("-----> Including Global Groups (Domain Users/Everyone)")
   }
   $user = get-aduser $searchObject -properties SID
   if(!$user)
   {
      write-host ("Object not found: "+$searchObject) -foregroundcolor Yellow
      exit
   }
                              # Enumerate groups (into SID names)
   $groupSIDs = get-aduser $searchObject -properties MemberOf -credential $myCred | select-object -ExpandProperty MemberOf | get-adgroup -credential $myCred | select-object SID

                              # Full SID list (i.e., all groups plus user SID).  Include "Domain Users" and "Everyone" if requested
   if($includeGlobalGroups -eq $true)
   {
      (((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : Including global groups") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true

      $allSIDs = @(
         $groupSIDs | ForEach-Object { $_.SID.ToString() }
         $user.SID.Value
         (Get-ADGroup "Domain Users").SID.ToString()
         "S-1-1-0"
      )
   }
   else
   {
      $allSIDs = @(
         $groupSIDs | ForEach-Object { $_.SID.ToString() }
         $user.SID.Value
      )
   }
}
else
{
   write-host "-----> Gathering AD membership for GROUP [" -nonewline
   write-host $searchObject -foregroundcolor Cyan -nonewline
   write-host "]..."
   $group = get-adgroup $searchObject
   if(!$group)
   {
      write-host ("Object not found: "+$searchObject) -foregroundcolor Yellow
      exit
   }

                              # Full SID list
   $allSIDs = @(
      (get-adgroup -LDAPFilter "(member:1.2.840.113556.1.4.1941:=$($group.DistinguishedName))").sid.value
      $group.SID.Value
   )
}
write-host ("<----- Found "+($allSIDs.Count)+" Security Principals") -foregroundcolor Green
write-host " "
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : Search Object ["+$searchObject+"]") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : ["+($allSIDs.Count)+"] SIDS Retrieved") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true

                              # ----- Create FOLDER TREE (using .NET for speed of enumeration)
$results = @()
write-host "-----> Scanning folders under [" -nonewline
write-host $rootPath -foregroundcolor Cyan -nonewline
write-host "]..."
$folderList = getFolderTree $rootPath

write-host "<----- Scan Complete [" -foregroundcolor Green -nonewline
write-host ([string]($folderList.count)+" folders]") -foregroundcolor Green
write-host " "
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : Folder Scan Completed") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : ["+($folderList.count)+"] Folders Detected") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true

write-host "-----> Parsing folder ACE Rights" -nonewline

                              # <<<<<<<<<<<<<<<<<<<< NOTES >>>>>>>>>>>>>>>>>>>>
                              # $sidLookup is list of all USER SIDS in hashtable (for search)
                              # $sidCache is cached translation table [BUILTIN\Administrators] to [S-1-5-32-544] to speed process
                              # $matchingACEs is a running list of folders which have matching permissions (non-inherited)

                              # Cache into HASH TABLE for faster lookups
$sidLookup = [System.Collections.Generic.HashSet[string]]::new()
foreach($sid in $allSIDs)
{
   $ret = $sidLookup.Add($sid)
}
                              # Cache IDENTITY LOOKUPS for faster processing
$sidCache = @{}
$matchingACEs = [System.Collections.Generic.List[object]]::new()
$myCount = 0
$myTotal = $folderList.count
foreach($folder in $folderList)
{
   $myCount = $myCount + 1
   if(($myCount % 85) -eq 0)
   {
      $percent = [math]::Round(($myCount/$myTotal)*100,2)
      write-progress -activity "Parsing ACE Rights" -status ("Processing folder "+$myCount+" of "+$myTotal) -percentcomplete $percent
   }
   try
   {
      if($folder.length -le 256)
      {
         $acl = [System.IO.Directory]::GetAccessControl($folder,[System.Security.AccessControl.AccessControlSections]::Access)
      }
      else
      {
         $acl = $null
      }
   }
   catch
   {
      continue
   }
   foreach($ace in $acl.Access)
   {
                              # Filter out folders which are inherited from above EXCEPT on rootPath folder
      if($ace.IsInherited){if($folder -ne $rootPath){continue}}
      $identity = $ace.IdentityReference.Value
      if(-not $sidCache.ContainsKey($identity))
      {
         if($ace.IdentityReference -is [System.Security.Principal.SecurityIdentifier])
         {
            $sidCache[$identity] = $identity
         }
         else
         {
            try
            {
               $sidCache[$identity] = $ace.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value
            }
            catch
            {
               continue
            }
         }
      }
      $sid = $sidCache[$identity]
      if($sidLookup.Contains($sid))
      {
                              # MAP GENERIC_ALL permissions to FullControl rights (like the GUI does)
         $sPerm = $ace.FileSystemRights
         if($ace.FileSystemRights -eq "268435456")
         {
            $sPerm = "FullControl"
         }
                              # Gather Object and ACL data for later report generation
         $matchingACEs.Add([PSCustomObject]@{
            Folder      = $folder
            Identity    = $ace.IdentityReference
            AccessType  = $ace.AccessControlType
            Rights      = $sPerm
            Inherited   = $ace.IsInherited
         })
      }
   }
}
write-host " "
write-host ("<----- Parse Complete ["+($matchingACEs.Count)+" Matching ACEs]") -foregroundcolor Green
write-host " "
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : ACE Scan Completed") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : ["+($matchingACEs.Count)+"] Matching ACEs") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true

                              # ----- Any ACE data?  If so, reparse into a smaller object for CSV export
write-host "-----> Generating Output Report"
if($matchingACEs)
{
                              # Only look for ALLOW access
   $effectiveAllow = $matchingACEs | where-object {$_.AccessType -eq "Allow"}
   if($effectiveAllow)
   {
      foreach($allowEntry in $effectiveAllow)
      {
                              # Objects to include in report
         $results += [PSCustomObject]@{
            FolderPath      = $allowEntry.Folder
            MatchingEntries = ($allowEntry.Identity -join "; ")
            Rights          = (($allowEntry.Rights | select-object -unique) -join "; ")
         }
      }
   }
}
$myTime = [int]((new-timespan -start $myDate -end (get-date)).totalseconds)
write-host ("<----- Report Generation Completed ["+$myTime+" seconds]") -foregroundcolor Green
(((get-date -format "yyyy-MM-dd HH:mm:ss").tostring())+" : Report Completed ["+$myTime+" seconds]") | out-file -FilePath $global:debugLog -Encoding:UTF8 -Append:$true
write-host " "

                              # Write report to FILE
$results | sort-object FolderPath -unique | export-csv $reportName -NoTypeInformation
write-host " "

                              # SUMMARY
write-host ("Report Saved At: "+$reportName) -foregroundcolor Cyan
write-host " "
