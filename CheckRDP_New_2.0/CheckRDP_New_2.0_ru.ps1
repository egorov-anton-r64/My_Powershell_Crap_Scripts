# Ввести имя пк-h
$computerName = "ИмяПк-h"
# Ввести количество дней за которое хотим получить лог ( 1 по умолчанию )
$StartTime = (Get-Date).AddDays(-7)

$domain = "ТВОЙДОМЕН"
$fullName = "$computerName.$domain"
$userName = $computerName -replace "-h$"  #удаление суффикса, если не нужно удали или закоментируй 

# Проверка доступности хоста
if (-not (Test-Connection $fullName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
    Write-Host "Компьютер $fullName выключен" -ForegroundColor Red
    exit
}

Write-Host "Компьютер $fullName включен" -ForegroundColor Green

# Проверка RDP порта
try {
    $rdpPortAvailable = Test-NetConnection $fullName -Port 3389 -InformationLevel Quiet -WarningAction SilentlyContinue
    Write-Host " - RDP порт доступен: $($rdpPortAvailable)" -ForegroundColor $(if ($rdpPortAvailable) {'Green'} else {'Red'})
}
catch {
    Write-Host " - Ошибка проверки порта RDP: $($_.Exception.Message)" -ForegroundColor Red
}

# Проверка локальных пользователей RDP
try {
    $params = @{
        ComputerName  = $fullName
        ErrorAction   = 'Stop'
        ScriptBlock   = { Get-LocalGroupMember "Пользователи удаленного рабочего стола" }
    }

    if ($RDPUsers = Invoke-Command @params) {
        $users = $RDPUsers.Name -join '; '
        Write-Host " - Локальные пользователи RDP:   $users" -ForegroundColor Green
    } else {
        Write-Host " - Нет пользователей в группе RDP" -ForegroundColor Red
    }
}
catch {
    Write-Host " - Ошибка проверки локальных пользователей: $($_.Exception.Message)" -ForegroundColor Red
}

# Проверка принадлежности пользователя к группе AD
try {
    $searcher = [ADSISearcher]"(&(objectCategory=User)(sAMAccountName=$userName))"
    $searchResult = $searcher.FindOne()

    if (-not $searchResult) {
        throw "Пользователь $userName не найден в Active Directory"
    }

    $memberOf = $searchResult.Properties["memberOf"] -join ';'
    $inGroup = $memberOf -match "ТВОЯRDPГРУППА"

    Write-Host " - Пользователь $userName в группе 'ТВОЯRDPГРУППА': $inGroup" -ForegroundColor $(if ($inGroup) {'Green'} else {'Red'})
}
catch {
    Write-Host " - Ошибка проверки AD: $($_.Exception.Message)" -ForegroundColor Red
}

# Анализ событий подключения через RDP
$filter = @{
    LogName     = "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational"
    StartTime = $StartTime
    ID          = 21,23,24,25,39
}

try {
    $events = Get-WinEvent -ComputerName $fullName -FilterHashtable $filter -ErrorAction Stop |
              Sort-Object TimeCreated -Descending

    if (-not $events) {
        Write-Host " - Событий подключений за неделю не найдено" -ForegroundColor Yellow
    } else {
        $remoteConnectionsFound = $false
        Write-Host "История удаленных подключений за неделю:" -ForegroundColor Green

        foreach ($ev in $events) {
            $ip = $ev.Properties[2].Value  # IP-адрес находится в Properties[2]
            $user = $ev.Properties[0].Value

            # Пропускаем локальные подключения
            if ([string]::IsNullOrEmpty($ip) -or $ip -in '-', '::1', '127.0.0.1', 'ЛОКАЛЬНЫЕ') {
                continue
            }

            $action = switch ($ev.Id) {
                21 { "Вход в систему"   }
                23 { "Выход из системы" }
                24 { "Отклчение"        }
                25 { "Подключение"      }
                39 { "Логофф"           }
            }

            $foregroundColor = switch ($action) {
                "Вход в систему"  { "Green"   }
                "Выход из системы"{ "Red"     }
                "Отклчение"       { "Magenta" }
                "Подключение"     { "Cyan"    }
                "Логофф"          { "Red"     }
                default           { "White"   }
            }

            Write-Host ("[{0}] > [{1}]" -f `
                $ev.TimeCreated.ToString("dd-MM-yyyy HH:mm"), $action) `
                -ForegroundColor $foregroundColor
            $remoteConnectionsFound = $true
        }

        if (-not $remoteConnectionsFound) {
            Write-Host " - Удаленных подключений за неделю не обнаружено" -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host " - Ошибка получения событий: $($_.Exception.Message)" -ForegroundColor Red
}
