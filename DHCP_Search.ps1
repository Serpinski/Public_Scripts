# *** DHCP SEARCH ***
#
# DESCRIPTION:
#    Searches DHCP server for matching lease
#
# SCRIPT INFO:
#    Modified: Nov 2023
#    Created: Jaime Habel
#

#$ErrorActionPreference = "SilentlyContinue"

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
Write-Host "By Jaime Habel" -foregroundcolor Green
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


# SIG # Begin signature block
# MIIIjwYJKoZIhvcNAQcCoIIIgDCCCHwCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUtcGUaHt8+6xf+SAZheoFwUA4
# dIKgggX7MIIF9zCCBN+gAwIBAgITFQAAPG2rWrykJPNqlQADAAA8bTANBgkqhkiG
# 9w0BAQsFADBGMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxFDASBgoJkiaJk/IsZAEZ
# FgRtYWhjMRcwFQYDVQQDEw5tYWhjLVZTQ0VSVC1DQTAeFw0yNTExMDMxNDUwMDZa
# Fw0yNzExMDMxNTAwMDZaMBYxFDASBgNVBAMTC0NvZGVTaWduaW5nMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAxv70ol0jVrpiXGPEOlNdps9lZ7SGtEAD
# nDOOQyWKLEjKV4AES/uVW7l7wJLNDG5ChxOOlJKmQ7jsGZBvkLs6SNXiSBLBO8Fo
# WIB20anZf3xBpzVPhqc6lc261imU3NNjIz4QSVpeDbIoSyvIi63ldu5r5qmMWKHF
# +nMQusvPUJHdOso5IA+7KCmh2AgKHDFzQL28toAk03Nr/JvuXqM/hLnmk76dvSGf
# PV/FVZUuL4Uo+r69WUfbNYY/+vo90nWcCdCR/9xcxPCePZ/5s8WFBMYZV7j50M1p
# PSf5bgVsCQ7E5mT9EcRq71238+5A1f2lT/JmFLajj9De0sWdHAbdkQIDAQABo4ID
# DDCCAwgwPgYJKwYBBAGCNxUHBDEwLwYnKwYBBAGCNxUIgbHAMIee9BeFjZEUgtvq
# ToO94luBXIWBhBqEq5lRAgFkAgEIMBMGA1UdJQQMMAoGCCsGAQUFBwMDMA4GA1Ud
# DwEB/wQEAwIHgDAbBgkrBgEEAYI3FQoEDjAMMAoGCCsGAQUFBwMDMB0GA1UdDgQW
# BBSlmUN93mNKTmRkxZWLt7Vf0XH2hjAfBgNVHSMEGDAWgBQfA1Qubg5/B7nWBrrK
# BtMTtgbmujCB+AYDVR0fBIHwMIHtMIHqoIHnoIHkhoG2bGRhcDovLy9DTj1tYWhj
# LVZTQ0VSVC1DQSgzKSxDTj1WU0NFUlQsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUy
# MFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9bWFoYyxE
# Qz1sb2NhbD9jZXJ0aWZpY2F0ZVJldm9jYXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xh
# c3M9Y1JMRGlzdHJpYnV0aW9uUG9pbnSGKVxcdnNjZXJ0XENlcnRFbnJvbGxcbWFo
# Yy1WU0NFUlQtQ0EoMykuY3JsMIG/BggrBgEFBQcBAQSBsjCBrzCBrAYIKwYBBQUH
# MAKGgZ9sZGFwOi8vL0NOPW1haGMtVlNDRVJULUNBLENOPUFJQSxDTj1QdWJsaWMl
# MjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERD
# PW1haGMsREM9bG9jYWw/Y0FDZXJ0aWZpY2F0ZT9iYXNlP29iamVjdENsYXNzPWNl
# cnRpZmljYXRpb25BdXRob3JpdHkwNwYDVR0RBDAwLqAsBgorBgEEAYI3FAIDoB4M
# HENvZGVTaWduaW5nQG1haGVhbHRoY2FyZS5jb20wTgYJKwYBBAGCNxkCBEEwP6A9
# BgorBgEEAYI3GQIBoC8ELVMtMS01LTIxLTE0NTk4Mzg2MDUtNzc2MzQwNzcwLTky
# NjcwOTA1NC02ODMzODANBgkqhkiG9w0BAQsFAAOCAQEAoW6QUR8zvKlRxFF1YIbn
# 8DWq4Thjv+kSOU0eVVemi20wZL6sMUDyHy/n4Apynh4cpQPq1VTQl6rtI8g/7QAE
# Y4mehdMHVkEOXslIiqqfuJaxBEBHhqrUWQn5E1FJyrQBSG6T2iEmdCBegqqgHiwI
# bY59tAsiCG90bq3effz6UWgFOxrnGrmVwciLtvR9chqEC7dG/62hf/X8bVL/CbGo
# +sFPYexKVZLjzJYFsWRvysr6tGhQNdS7fJUGeDeDw9oetLZMLJbZAgeVcVDUCx/Y
# +0txsPthg7ygLsZQJI5cvQSX55AqtRbskoO1/x1GXhx4QNdKy3BQTO30oP0rIkCC
# YzGCAf4wggH6AgEBMF0wRjEVMBMGCgmSJomT8ixkARkWBWxvY2FsMRQwEgYKCZIm
# iZPyLGQBGRYEbWFoYzEXMBUGA1UEAxMObWFoYy1WU0NFUlQtQ0ECExUAADxtq1q8
# pCTzapUAAwAAPG0wCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKEC
# gAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwG
# CisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFIc3K7YW8joLzSHmBLrltzw89gCm
# MA0GCSqGSIb3DQEBAQUABIIBABnNnznddbkfLxADwDXJrc0xF8kMWkILMSsBo9Lu
# /dOZzagsROmg/HvUKbYm/LVxJxjRCHQAq3m8IcmvAX2kgYVVDJG9G5caqcC5A19O
# 83oflEr8DhyXZCRkk6zw3mbJ65RBcJvyGIfn8mA7ZbRqkZxPiR3YqOXZ8Xzm1Lqe
# O7X6EAP/uDlgYMN1rJu7zlYyprd1fi3hbT2MuT6lguOr9SCya2ZYl2/bzlvXdGUh
# 8Lla4poNnR+MUB988Wn/E7KV/r+Ku5wCSqOBKZXKxGk3pPvAcmC+2w1NaYypktOe
# bU1PwjkDJmYL4oFML/WT6JJhIGxNMFwRoHZP6xyrggjKEOo=
# SIG # End signature block
