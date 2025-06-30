# Enter the computer name with -h suffix
$computerName = "starchenko-h"
# Enter the number of days to get logs for (default is 1)
$StartTime = (Get-Date).AddDays(-7)

$domain = "YOU.DOMAIN.COM"
$fullName = "$computerName.$domain"
$userName = $computerName -replace "-h$"  # remove suffix

# Check host availability
if (-not (Test-Connection $fullName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
    Write-Host "Computer $fullName is offline" -ForegroundColor Red
    exit
}

Write-Host "Computer $fullName is online" -ForegroundColor Green

# Check RDP port
try {
    $rdpPortAvailable = Test-NetConnection $fullName -Port 3389 -InformationLevel Quiet -WarningAction SilentlyContinue
    Write-Host " - RDP port available: $($rdpPortAvailable)" -ForegroundColor $(if ($rdpPortAvailable) {'Green'} else {'Red'})
}
catch {
    Write-Host " - Error checking RDP port: $($_.Exception.Message)" -ForegroundColor Red
}

# Check local RDP users
try {
    $params = @{
        ComputerName  = $fullName
        ErrorAction   = 'Stop'
        ScriptBlock   = { Get-LocalGroupMember "Remote Desktop Users" }
    }

    if ($RDPUsers = Invoke-Command @params) {
        $users = $RDPUsers.Name -join '; '
        Write-Host " - Local RDP users:`n   $users" -ForegroundColor Green
    } else {
        Write-Host " - No users in RDP group" -ForegroundColor Red
    }
}
catch {
    Write-Host " - Error checking local RDP users: $($_.Exception.Message)" -ForegroundColor Red
}

# Check user membership in AD group
try {
    $searcher = [ADSISearcher]"(&(objectCategory=User)(sAMAccountName=$userName))"
    $searchResult = $searcher.FindOne()

    if (-not $searchResult) {
        throw "User $userName not found in Active Directory"
    }

    $memberOf = $searchResult.Properties["memberOf"] -join ';'
    $inGroup = $memberOf -match "YOU RDP ACCES GROUP"

    Write-Host " - User $userName is in group 'YOU RDP ACCES GROUP': $inGroup" -ForegroundColor $(if ($inGroup) {'Green'} else {'Red'})
}
catch {
    Write-Host " - Error checking AD: $($_.Exception.Message)" -ForegroundColor Red
}

# Analyze RDP connection events
$filter = @{
    LogName     = "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational"
    StartTime   = $StartTime
    ID          = 21,23,24,25,39
}

try {
    $events = Get-WinEvent -ComputerName $fullName -FilterHashtable $filter -ErrorAction Stop |
              Sort-Object TimeCreated -Descending

    if (-not $events) {
        Write-Host " - No connection events found for the last week" -ForegroundColor Yellow
    } else {
        $remoteConnectionsFound = $false
        Write-Host "Remote connection history for the last week:" -ForegroundColor Green

        foreach ($ev in $events) {
            $ip = $ev.Properties[2].Value  # IP address is in Properties[2]
            $user = $ev.Properties[0].Value

            # Skip local connections
            if ([string]::IsNullOrEmpty($ip) -or $ip -in '-', '::1', '127.0.0.1', 'LOCAL') {
                continue
            }

            $action = switch ($ev.Id) {
                21 { "Logon"      }
                23 { "Logoff"     }
                24 { "Disconnect" }
                25 { "Reconnect"  }
                39 { "Logout"     }
            }

            $foregroundColor = switch ($action) {
                "Logon"      { "Green"   }
                "Logoff"     { "Red"     }
                "Disconnect" { "Magenta" }
                "Reconnect"  { "Cyan"    }
                "Logout"     { "Red"     }
                default      { "White"   }
            }

            Write-Host ("[{0}] > [{1}]" -f `
                $ev.TimeCreated.ToString("dd-MM-yyyy HH:mm"), $action) `
                -ForegroundColor $foregroundColor
            $remoteConnectionsFound = $true
        }

        if (-not $remoteConnectionsFound) {
            Write-Host " - No remote connections detected during the week" -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host " - Error retrieving events: $($_.Exception.Message)" -ForegroundColor Red
}
