# *** COPY NTFS PERMISSIONS ***
#
# DESCRIPTION:
#    - Copy NTFS permissions from one file/folder to another (replacing existing permissions)
#
# PARAMETERS:
#    - origObject = Folder (including subfolders) to copy FROM
#    - destObject = Folder (including subfolders) to copy TO
#
# SCRIPT INFO:
#    Modified: Oct 2024
#    Created: Jaime
# 

                              # Parameter list
Param(
  [string]$origObject,
  [string]$destObject
)

$ErrorActionPreference = "SilentlyContinue"

                              # Replaces COPY permissions from A identifier to B identifier
function Copy_Permissions($theOrigObject,$theDestObject)
{
   $myObject = $theOrigObject.tolower()
   if((test-path -path $myObject) -eq $true)
   {
      if((test-path -path $theDestObject) -eq $true)
      {
                              # Get current permissions
         $myPerms = get-acl -path $myObject
         write-host "Permissions:"
         foreach($entry in $myPerms.access)
         {
            if($entry.FileSystemRights -ne "268435456")
            {
               write-host ("   ["+$entry.IdentityReference+"] has ["+$entry.FileSystemRights+"] access of control type ["+$entry.AccessControlType+"]") -foregroundcolor Cyan
            }
         }

                              # Save to destination
         $results = set-acl -path $theDestObject -aclobject $myPerms
      }
   }
}

                              # ----------------------------- MAIN SCRIPT
cls
write-host "COPY NTFS PERMISSIONS Script:" -foregroundcolor Green
write-host " "

if($origObject.length -eq 0 -or $destObject.length -eq 0)
{
   write-host "Please specify 'origObject' and 'destObject' parameters" -foregroundcolor Yellow
   write-host "ORIGOBJECT = Directory (U:\myFolder)" -foregroundcolor Yellow
   write-host "DESTOBJECT = Directory (U:\myNewFolder)" -foregroundcolor Yellow
   exit
}

write-host "Source Folder:"
write-host ("   "+$origObject) -foregroundcolor Cyan
write-host " "
write-host "Destination Folder:"
write-host ("   "+$destObject) -foregroundcolor Cyan
write-host " "

Copy_Permissions $origObject $destObject

Write-Host " "
Write-Host "OPERATION COMPLETED" -foregroundcolor Green