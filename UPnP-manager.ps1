# miniUPnP Manager by INotxddd / mr_sartok
# Требуются права администратора
#Requires -RunAsAdministrator

# Установка кодировки для корректного отображения русского языка
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Переход в нужный раздел реестра и установка параметров консоли
Set-Location HKCU:\Console
$psProfileKey = '.\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe'
if (-not (Test-Path $psProfileKey)) { New-Item $psProfileKey | Out-Null }
Set-Location $psProfileKey
$size                 = $host.UI.RawUI.WindowSize
$size.Width           = 100
$size.Height          = 25
$host.UI.RawUI.WindowSize = $size
New-ItemProperty -Path . -Name FaceName   -Value "Consolas" -PropertyType String -Force
New-ItemProperty -Path . -Name FontFamily -Value 0x00000036 -PropertyType DWord  -Force
New-ItemProperty -Path . -Name FontSize   -Value 0x00100000 -PropertyType DWord  -Force # 16 px
New-ItemProperty -Path . -Name FontWeight -Value 0x00000190 -PropertyType DWord  -Force # 400
Set-ConsoleFont 16
Set-Location $env:USERPROFILE

# Визуальные параметры
(Get-Host).UI.RawUI.ForegroundColor = "White"
(Get-Host).UI.RawUI.BackgroundColor = "Black"
(Get-Host).UI.RawUI.CursorSize      = 10
(Get-Host).UI.RawUI.WindowTitle     = "miniUPnP Manager by INotxddd / mr_sartok"
Clear-Host

# Пути и файлы
$PORTS_FILE   = Join-Path $env:USERPROFILE "UPnP\inot_ports.txt"
$upnpcPath    = Join-Path $env:USERPROFILE "UPnP\upnpc-static.exe"
$upnpfolder   = Join-Path $env:USERPROFILE "UPnP\"

function CreateFolder {
    if (-not (Test-Path $upnpfolder)) {
        New-Item -Path $upnpfolder -ItemType Directory | Out-Null
        Write-Host "Папка 'UPnP' создана." -ForegroundColor Green
    }
}
CreateFolder

#################### UI helpers ####################
function Show-Header {
    param([string]$Title)
    $width = $host.UI.RawUI.WindowSize.Width
    Write-Host "$Title" -ForegroundColor Yellow
}

function Show-Menu {
    param([string]$Title = 'miniUPnP Manager')
    Clear-Host
    Show-Header $Title

    Write-Host "[ 1 ]" -ForegroundColor Green -NoNewline; Write-Host " Использовать стандартные порты Minecraft" -ForegroundColor White
    Write-Host "     └─ 25565 TCP/UDP" -ForegroundColor Gray
    Write-Host
    Write-Host "[ 2 ]" -ForegroundColor Green -NoNewline; Write-Host " Использовать стандартные порты Factorio" -ForegroundColor White
    Write-Host "     └─ 34197 UDP" -ForegroundColor Gray
    Write-Host
    Write-Host "[ 3 ]" -ForegroundColor Green -NoNewline; Write-Host " Настроить свои порты" -ForegroundColor White
    Write-Host "[ 4 ]" -ForegroundColor Green -NoNewline; Write-Host " Проверить статус подключения" -ForegroundColor White
    Write-Host "[ 5 ]" -ForegroundColor Green -NoNewline; Write-Host " Показать текущие правила" -ForegroundColor White
    Write-Host "[ 6 ]" -ForegroundColor Green -NoNewline; Write-Host " Удалить все правила" -ForegroundColor White
    Write-Host "[ 7 ]" -ForegroundColor Red  -NoNewline; Write-Host " Выход" -ForegroundColor White
    Write-Host
}

#################### upnpc helpers ####################
function Test-Upnpc {
    if (-not (Test-Path $upnpcPath)) {
        Write-Host "upnpc не установлен. Начинаем установку..." -ForegroundColor Yellow
        Install-Upnpc
    }
}

function Install-Upnpc {
    try {
        Write-Host "Скачивание upnpc..." -ForegroundColor Cyan
        $url          = "http://miniupnp.free.fr/files/upnpc-exe-win32-20220515.zip"
        $zipPath      = Join-Path $env:TEMP "upnpc.zip"
        $extractPath  = Join-Path $env:TEMP "temp_upnpc"
        Invoke-WebRequest -Uri $url -OutFile $zipPath
        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        Copy-Item "$extractPath\upnpc-static.exe" "$upnpfolder" -Force
        Copy-Item "$extractPath\upnpc-shared.exe" "$upnpfolder" -Force
        Copy-Item "$extractPath\miniupnpc.dll"   "$upnpfolder" -Force
        Remove-Item $zipPath -Force
        Remove-Item $extractPath -Recurse -Force
        Write-Host "Установка завершена успешно!" -ForegroundColor Green
        Start-Sleep -Seconds 2
    } catch {
        Write-Host "Ошибка при установке: $_" -ForegroundColor Red
        pause; exit
    }
}

#################### Presets ####################
function Set-MinecraftPorts {
    Write-Host "Настройка портов Minecraft..." -ForegroundColor Cyan
    $ports = @(
        @{Port="25565"; Protocol="tcp"; Description="Minecraft TCP"},
        @{Port="25565"; Protocol="udp"; Description="Minecraft UDP"}
    )
    Apply-Preset -ports $ports
}

function Set-FactorioPorts {
    Write-Host "Настройка портов Factorio..." -ForegroundColor Cyan
    $ports = @(
        @{Port="34197"; Protocol="udp"; Description="Factorio UDP"}
    )
    Apply-Preset -ports $ports
}

function Apply-Preset {
    param([array]$ports)
    foreach ($p in $ports) {
        Write-Host "Настройка порта $($p.Port) $($p.Protocol)..." -ForegroundColor Yellow
        & $upnpcPath -e $p.Description -d $p.Port $p.Protocol | Out-Null
        & $upnpcPath -e $p.Description -a "@" $p.Port $p.Port $p.Protocol | Out-Null
    }
    # merge into stored file for clean removal later
    $existing = @()
    if (Test-Path $PORTS_FILE) {
        try { $existing = Get-Content $PORTS_FILE | ConvertFrom-Json } catch { $existing = @() }
    }
    foreach ($p in $ports) {
        if (-not ($existing | Where-Object { $_.Port -eq $p.Port -and $_.Protocol -eq $p.Protocol -and $_.Description -eq $p.Description })) {
            $existing += $p
        }
    }
    $existing | ConvertTo-Json | Set-Content $PORTS_FILE
    Write-Host "Текущие правила:" -ForegroundColor Green
    & $upnpcPath -l
    pause
}

#################### Custom ports workflow ####################
function Set-CustomPorts {
    do {
        Clear-Host; Show-Header "Настройка пользовательских портов"
        $current = @(); if (Test-Path $PORTS_FILE) { $current = Get-Content $PORTS_FILE | ConvertFrom-Json -ErrorAction SilentlyContinue }
        if ($current.Count) {
            Write-Host "Текущие настройки портов:" -ForegroundColor Cyan
            foreach ($p in $current) { Write-Host "  ├─ $($p.Port) $($p.Protocol.ToUpper()) - $($p.Description)" -ForegroundColor Gray }
        } else { Write-Host "Нет настроенных портов" -ForegroundColor Yellow }

        Write-Host "Доступные действия:" -ForegroundColor White
        Write-Host "[ 1 ]" -ForegroundColor Green -NoNewline; Write-Host " Добавить порт" -ForegroundColor White
        Write-Host "[ 2 ]" -ForegroundColor Green -NoNewline; Write-Host " Применить настройки" -ForegroundColor White
        Write-Host "[ 3 ]" -ForegroundColor Yellow -NoNewline; Write-Host " Удалить порт" -ForegroundColor White
        Write-Host "[ 4 ]" -ForegroundColor Red  -NoNewline; Write-Host " Вернуться в главное меню" -ForegroundColor White
        $choice = Read-Host "Выберите действие (1-4)"
        switch ($choice) {
            '1' { Add-CustomPort   -current $current }
            '2' { if ($current.Count) { Apply-CustomPorts -ports $current } else { Write-Host "Нет портов для применения!" -ForegroundColor Red; Start-Sleep 2 } }
            '3' { if ($current.Count) { Remove-CustomPort -current $current } else { Write-Host "Нет портов для удаления!" -ForegroundColor Red; Start-Sleep 2 } }
            '4' { return }
            default { Write-Host "Неверный выбор!" -ForegroundColor Red; Start-Sleep 1 }
        }
    } while ($true)
}

function Add-CustomPort {
    param([array]$current)
    do {
        $port = Read-Host "Введите номер порта (1-65535)"
        if ($port -match '^\d+$' -and [int]$port -ge 1 -and [int]$port -le 65535) { break }
        Write-Host "Неверный номер порта!" -ForegroundColor Red
    } while ($true)

    Write-Host "Выберите протокол:"; Write-Host "[ 1 ]" -ForegroundColor Cyan -NoNewline; Write-Host " TCP" -ForegroundColor White
    Write-Host "[ 2 ]" -ForegroundColor Cyan -NoNewline; Write-Host " UDP" -ForegroundColor White
    Write-Host "[ 3 ]" -ForegroundColor Cyan -NoNewline; Write-Host " TCP и UDP" -ForegroundColor White
    do { $protoChoice = Read-Host "Ваш выбор (1-3)"; if ($protoChoice -match '^[1-3]$') { break }; Write-Host "Неверный выбор!" -ForegroundColor Red } while ($true)
    $desc = Read-Host "Введите описание правила"
    switch ($protoChoice) {
        '1' { $current += @{Port=$port; Protocol='tcp'; Description=$desc} }
        '2' { $current += @{Port=$port; Protocol='udp'; Description=$desc} }
        '3' { $current += @{Port=$port; Protocol='tcp'; Description="$desc (TCP)"}; $current += @{Port=$port; Protocol='udp'; Description="$desc (UDP)"} }
    }
    $current | ConvertTo-Json | Set-Content $PORTS_FILE
    Write-Host "Порт успешно добавлен!" -ForegroundColor Green; Start-Sleep 1
}

function Apply-CustomPorts { param([array]$ports)
    Write-Host "Применение настроек портов..." -ForegroundColor Cyan
    foreach ($p in $ports) {
        Write-Host "Настройка порта $($p.Port) $($p.Protocol)..." -ForegroundColor Yellow
        & $upnpcPath -e $p.Description -d $p.Port $p.Protocol | Out-Null
        & $upnpcPath -e $p.Description -a "@" $p.Port $p.Port $p.Protocol | Out-Null
    }
    Write-Host "Текущие правила:" -ForegroundColor Green; & $upnpcPath -l; pause
}

function Remove-CustomPort {
    param([array]$current)
    Clear-Host; Show-Header "Удаление пользовательских портов"
    if (-not $current.Count) { Write-Host "Нет настроенных портов" -ForegroundColor Yellow; Start-Sleep 2; return }
    Write-Host "Доступные порты для удаления:" -ForegroundColor Cyan
    for ($i=0; $i -lt $current.Count; $i++) {
        $p=$current[$i]; Write-Host "  ├─ $($i+1). $($p.Port) $($p.Protocol.ToUpper()) - $($p.Description)" -ForegroundColor Gray
    }
    Write-Host "n[ 0 ]" -ForegroundColor Red -NoNewline; Write-Host " Отмена" -ForegroundColor White
    $sel = Read-Host "Выберите порт для удаления (0-$($current.Count))"
    if ($sel -eq '0') { return }
    if ($sel -match '^\d+$' -and [int]$sel -ge 1 -and [int]$sel -le $current.Count) {
        $idx  = [int]$sel - 1; $port = $current[$idx]
        Write-Host "Вы уверены, что хотите удалить порт $($port.Port) $($port.Protocol)?" -ForegroundColor Yellow
        if ((Read-Host "Подтвердите (y/n)") -eq 'y') {
            & $upnpcPath -e $port.Description -d $port.Port $port.Protocol | Out-Null
            $current = $current | Where-Object { -not($_.Port -eq $port.Port -and $_.Protocol -eq $port.Protocol -and $_.Description -eq $port.Description) }
            if ($current.Count) { $current | ConvertTo-Json | Set-Content $PORTS_FILE } else { if (Test-Path $PORTS_FILE) { Remove-Item $PORTS_FILE -Force } }
            Write-Host "Порт и правило успешно удалены!" -ForegroundColor Green; Start-Sleep 1
        }
    } else { Write-Host "Неверный выбор!" -ForegroundColor Red; Start-Sleep 1 }
}

#################### Remove all ####################
function Remove-AllRules {
    Write-Host "Удаление всех правил..." -ForegroundColor Yellow
    $defaultPorts = @(
        @{Port="25565"; Protocol="tcp"; Desc="Minecraft TCP"},
        @{Port="25565"; Protocol="udp"; Desc="Minecraft UDP"},
        @{Port="34197"; Protocol="udp"; Desc="Factorio UDP"}
    )
    foreach ($p in $defaultPorts) { & $upnpcPath -e $p.Desc -d $p.Port $p.Protocol | Out-Null }
    if (Test-Path $PORTS_FILE) {
        $custom = Get-Content $PORTS_FILE | ConvertFrom-Json -ErrorAction SilentlyContinue
        foreach ($p in $custom) { & $upnpcPath -e $p.Description -d $p.Port $p.Protocol | Out-Null }
        Remove-Item $PORTS_FILE -Force
    }
    Write-Host "Все правила были удалены!" -ForegroundColor Green; pause
}

#################### Main loop ####################
function Start-MainLoop {
    Test-Upnpc
    do {
        Show-Menu
        $choice = Read-Host "Выберите действие (1-7)"
        switch ($choice) {
            '1' { Set-MinecraftPorts }
            '2' { Set-FactorioPorts  }
            '3' { Set-CustomPorts   }
            '4' { Write-Host "Проверка статуса подключения..." -ForegroundColor Yellow; & $upnpcPath -s; pause }
            '5' { Write-Host "Текущие правила:" -ForegroundColor Green; & $upnpcPath -l; pause }
            '6' { Remove-AllRules }
            '7' { return }
            default { Write-Host "Неверный выбор!" -ForegroundColor Red; Start-Sleep 1 }
        }
    } while ($true)
}

#################### Entry ####################
Start-MainLoop
