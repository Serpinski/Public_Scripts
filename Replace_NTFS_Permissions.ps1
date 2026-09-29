# *** REPLACE NTFS PERMISSIONS ***
#
# DESCRIPTION:
#    - Replace one GUID for another GUID in NTFS permissions for file/folder
#
# PARAMETERS:
#    - myObject = Folder (including subfolders) to evaluate
#    - origName = Search for this OLD access holder ("MAHC\jhabel", etc.)
#    - newName = Replace with this NEW access holder ("MAHC\Domain Admins", "NONE", etc.)
#
# SCRIPT INFO:
#    Modified: April 2024
#    Created: Jaime
# 

                              # Parameter list
Param(
  [string]$myObject,
  [string]$origName,
  [string]$newName
)

$ErrorActionPreference = "SilentlyContinue"

                              # Replaces OBJECT permissions from A identifier to B identifier, with same access rights
function Replace_Permissions($theObject,$theOrigName,$theNewName)
{
   $myObject = $theObject.tolower()
   if((test-path -path $myObject) -eq $true)
   {
                              # Check current permissions
      $myPerms = get-acl -path $myObject

                              # Loop thru access permissions looking for identifier A
      $isFound = $false
      foreach($aclCheck in $myPerms.access)
      {
         if($aclCheck.IdentityReference.tostring().tolower() -eq $theOrigName.tostring().tolower())
         {
                              # Ignore inherited rights on sub-levels, which should be fixed at the parent level
            if($aclCheck.IsInherited -ne $true)
            {
                              # FOUND --- Identifier A
               $isFound = $true
               $oldAccessRule = $aclCheck

                              # Create new access object using Identifier B user and Identifier A access rights
               $accessID = $theNewName
               $accessRights = $aclCheck.FileSystemRights

                              # Treat folders different for access inheritance
               if((get-item $theObject) -is [System.IO.DirectoryInfo])
               {
                  $accessInheritance = "ContainerInherit, ObjectInherit"
               }
               else
               {
                  $accessInheritance = "None"
               }
               $accessPropagation = $aclCheck.PropagationFlags
               $accessType = $aclCheck.AccessControlType
               $accessARG = $accessID, $accessRights, $accessInheritance, $accessPropagation, $accessType
               $fileSystemAccessRule = new-object -typeName System.Security.AccessControl.FileSystemAccessRule -argumentlist $accessARG

                              # Remove Identifier A from access
               $results = $myPerms.removeaccessrule($oldAccessRule)
            }
         }
      }
                              # If FOUND, then perform permissions changes
      if($isFound -eq $true)
      {
                              # Running count of changes
         $global:fixCount = $global:fixCount + 1

                              # Add Identifier B, unless NONE was specified
         if($theNewName -ne "NONE")
         {
            $results = $myPerms.addaccessrule($fileSystemAccessRule)
         }
                              # Commit permissions change to object
         $results = set-acl -path $myObject -aclobject $myPerms
         write-host "Adjusted" -foregroundcolor Yellow
      }
      else
      {
         write-host "No change" -foregroundcolor Green
      }
   }
}

                              # ----------------------------- MAIN SCRIPT
cls
write-host "REPLACE NTFS PERMISSIONS Script:" -foregroundcolor Green
write-host " "

if($myObject -eq "" -or $origName -eq "" -or $newName -eq "")
{
   write-host "Please specify MYOBJECT, ORIGNAME, and NEWNAME Parameters" -foregroundcolor Yellow
   write-host "MYOBJECT = Directory (U:\Test)" -foregroundcolor Yellow
   write-host "ORIGNAME = ID to replace (DOMAIN\jaime)" -foregroundcolor Yellow
   write-host "NEWNAME = ID to replace with (DOMAIN\jdoe)" -foregroundcolor Yellow
   exit
}

write-host ("Searching in [") -nonewline -foregroundcolor Cyan
write-host $myObject -nonewline -foregroundcolor Yellow
write-host ("] folder and subfolders") -foregroundcolor Cyan
write-host ("Replacing all instances of [") -nonewline -foregroundcolor Cyan
write-host $origName -nonewline -foregroundcolor Yellow
write-host "] with [" -nonewline -foregroundcolor Cyan
write-host $newName -nonewline -foregroundcolor Yellow
write-host ("], using the same permissions level") -foregroundcolor Cyan
write-host " "

$global:fixCount = 0
write-host ":: Searching ::" -foregroundcolor Green
                              # Get directory AND all child objects
$myList = (get-childitem -path $myObject -recurse)+(get-item -path $myObject)
foreach($entry in $myList)
{
   write-host (($entry.fullname)+"...") -nonewline
   Replace_Permissions ($entry.fullname) $origName $newName
}

write-host " "
write-host ("["+$global:fixCount+"] objects updated") -foregroundcolor Cyan

Write-Host " "
Write-Host "OPERATION COMPLETED" -foregroundcolor Green
