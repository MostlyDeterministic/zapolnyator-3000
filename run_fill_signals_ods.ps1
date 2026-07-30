#requires -Version 5.1
[CmdletBinding()]
param()

# Настройка кодировки для корректного русского текста из Python в Windows PowerShell 5.1
try {
    chcp 65001 | Out-Null
}
catch {
}

try {
    $Utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [Console]::InputEncoding = $Utf8NoBom
    [Console]::OutputEncoding = $Utf8NoBom
    $OutputEncoding = $Utf8NoBom
}
catch {
}

$env:PYTHONIOENCODING = 'utf-8'
$env:PYTHONUTF8 = '1'

$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Cyan
}

function Write-WarnMsg {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Yellow
}

function Write-ErrMsg {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Red
}

function Pause-BeforeExit {
    Write-Host ""
    Read-Host "Нажмите Enter для закрытия окна"
}

function Test-CommandExists {
    param([string]$CommandName)
    try {
        $cmd = Get-Command $CommandName -ErrorAction Stop
        return $null -ne $cmd
    }
    catch {
        return $false
    }
}

function Get-PythonCommand {
    if (Test-CommandExists -CommandName 'python') {
        try {
            & python --version *> $null
            if ($LASTEXITCODE -eq 0) { return 'python' }
        }
        catch {}
    }

    if (Test-CommandExists -CommandName 'py') {
        try {
            & py -3 --version *> $null
            if ($LASTEXITCODE -eq 0) { return 'py' }
        }
        catch {}
    }

    return $null
}

function Invoke-Python {
    param(
        [Parameter(Mandatory = $true)][string]$PythonCmd,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [switch]$Silent
    )

    if ($PythonCmd -eq 'py') {
        if ($Silent) { & py -3 @Arguments *> $null } else { & py -3 @Arguments }
    }
    else {
        if ($Silent) { & python @Arguments *> $null } else { & python @Arguments }
    }

    return $LASTEXITCODE
}

function Install-PythonIfNeeded {
    $pythonCmd = Get-PythonCommand
    if ($pythonCmd) { return $pythonCmd }

    Write-WarnMsg "Python 3 не найден."
    $answer = Read-Host "Установить Python 3 автоматически через winget? (Y/N)"
    if ($answer -notmatch '^(Y|y|Д|д)$') {
        throw "Python 3 не установлен. Установка отменена пользователем."
    }

    if (-not (Test-CommandExists -CommandName 'winget')) {
        throw "Не найден winget. Установите Python 3 вручную и повторите запуск."
    }

    Write-Info "Устанавливаю Python 3..."
    & winget install --exact --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) {
        throw "Не удалось установить Python 3 через winget."
    }

    $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [System.Environment]::GetEnvironmentVariable('Path', 'User')

    $pythonCmd = Get-PythonCommand
    if (-not $pythonCmd) {
        throw "Python установлен, но команда python/py не найдена. Закройте и заново откройте PowerShell."
    }

    return $pythonCmd
}

function Test-PythonModule {
    param(
        [Parameter(Mandatory = $true)][string]$PythonCmd,
        [Parameter(Mandatory = $true)][string]$ModuleName
    )

    $code = "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$ModuleName') else 1)"
    $exitCode = Invoke-Python -PythonCmd $PythonCmd -Arguments @('-c', $code) -Silent
    return ($exitCode -eq 0)
}

function Ensure-Pip {
    param([Parameter(Mandatory = $true)][string]$PythonCmd)

    Write-Info "Проверяю pip..."
    $exitCode = Invoke-Python -PythonCmd $PythonCmd -Arguments @('-m', 'pip', '--version') -Silent
    if ($exitCode -eq 0) {
        Write-Info "pip найден."
        return
    }

    Write-WarnMsg "pip не найден. Пытаюсь включить его через ensurepip..."
    $exitCode = Invoke-Python -PythonCmd $PythonCmd -Arguments @('-m', 'ensurepip', '--upgrade') -Silent
    if ($exitCode -ne 0) {
        throw "В выбранном Python недоступен pip, и не удалось включить его через ensurepip."
    }

    $exitCode = Invoke-Python -PythonCmd $PythonCmd -Arguments @('-m', 'pip', '--version') -Silent
    if ($exitCode -ne 0) {
        throw "pip не удалось подготовить после ensurepip."
    }

    Write-Info "pip успешно подготовлен."
}

function Install-MissingPythonPackages {
    param([Parameter(Mandatory = $true)][string]$PythonCmd)

    Ensure-Pip -PythonCmd $PythonCmd

    $requiredModules = @(
        @{ Module = 'odf';      Package = 'odfpy'    },
        @{ Module = 'openpyxl'; Package = 'openpyxl' }
    )

    $missingPackages = New-Object System.Collections.Generic.List[string]

    foreach ($item in $requiredModules) {
        $moduleName = $item.Module
        $packageName = $item.Package

        if (Test-PythonModule -PythonCmd $PythonCmd -ModuleName $moduleName) {
            Write-Info ("Пакет уже установлен: " + $packageName)
        }
        else {
            Write-WarnMsg ("Не найден пакет: " + $packageName)
            $missingPackages.Add($packageName)
        }
    }

    if ($missingPackages.Count -eq 0) {
        Write-Info "Все необходимые Python-пакеты уже установлены."
        return
    }

    $pkgText = ($missingPackages -join ', ')
    $answer = Read-Host ("Установить отсутствующие пакеты: " + $pkgText + " ? (Y/N)")
    if ($answer -notmatch '^(Y|y|Д|д)$') {
        throw "Установка Python-пакетов отменена пользователем."
    }

    Write-Info ("Устанавливаю пакеты: " + $pkgText)
    $installArgs = @('-m', 'pip', 'install')
    $installArgs += $missingPackages

    $exitCode = Invoke-Python -PythonCmd $PythonCmd -Arguments $installArgs
    if ($exitCode -ne 0) {
        throw "Не удалось установить необходимые Python-пакеты."
    }
}

function Write-PythonOutputLine {
    param([string]$Text)

    if ($Text -like 'NEW_MODULES_COUNT:*') {
        if ($Text -match 'NEW_MODULES_COUNT:\s*0') {
            Write-Host "Новых модулей в XLSX, отсутствующих в ODS: нет" -ForegroundColor Green
        }
        else {
            $countText = $Text -replace '^NEW_MODULES_COUNT:\s*', ''
            Write-Host ("Новых модулей в XLSX, отсутствующих в ODS: " + $countText) -ForegroundColor Yellow
        }
        return
    }

    if ($Text -like 'NEW_MODULE:*') {
        $payload = $Text -replace '^NEW_MODULE:\s*', ''
        $parts = $payload -split '\s*\|\s*', 2
        if ($parts.Count -eq 2) {
            Write-Host "  - " -NoNewline
            Write-Host $parts[0] -ForegroundColor Yellow -NoNewline
            Write-Host ": " -NoNewline
            Write-Host $parts[1] -ForegroundColor Red
        }
        else {
            Write-Host $Text -ForegroundColor Red
        }
        return
    }

    if ($Text -eq 'Список новых модулей:') {
        Write-Host $Text -ForegroundColor Yellow
        return
    }

    if ($Text -like '  - *') {
        return
    }

    if ($Text -like 'DUPLICATE_KEYS_COUNT:*') {
        if ($Text -match 'DUPLICATE_KEYS_COUNT:\s*0') {
            Write-Host "Дубликатов ключей в XLSX: нет" -ForegroundColor Green
        }
        else {
            $countText = $Text -replace '^DUPLICATE_KEYS_COUNT:\s*', ''
            Write-Host ("Дубликатов ключей в XLSX: " + $countText) -ForegroundColor Magenta
        }
        return
    }

    if ($Text -like 'DUPLICATE_KEY:*') {
        $payload = $Text -replace '^DUPLICATE_KEY:\s*', ''
        $parts = $payload -split '\s*\|\s*', 5
        if ($parts.Count -eq 5) {
            Write-Host "  - " -NoNewline
            Write-Host $parts[0] -ForegroundColor Magenta -NoNewline
            Write-Host ": " -NoNewline
            Write-Host $parts[1] -ForegroundColor Yellow -NoNewline
            Write-Host ", канал " -NoNewline
            Write-Host $parts[2] -ForegroundColor Cyan -NoNewline
            Write-Host ", тип " -NoNewline
            Write-Host $parts[3] -ForegroundColor Cyan -NoNewline
            Write-Host ", строки XLSX: " -NoNewline
            Write-Host $parts[4] -ForegroundColor Red
        }
        else {
            Write-Host $Text -ForegroundColor Magenta
        }
        return
    }

    if ($Text -eq 'Список дубликатов:') {
        Write-Host $Text -ForegroundColor Magenta
        return
    }

    Write-Host $Text
}

try {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    Set-Location -Path $ScriptDir

    Write-Info ("Рабочая папка: " + $ScriptDir)

    $OdsFiles = Get-ChildItem -Path $ScriptDir -File -Filter *.ods |
        Where-Object { $_.Name -notlike '*.backup_before_fill.ods' } |
        Sort-Object -Property Name

    if (($null -eq $OdsFiles) -or ($OdsFiles.Count -eq 0)) {
        throw "В папке не найдено ни одного .ods файла."
    }

    Write-Host "Найденные ODS-файлы:"
    for ($i = 0; $i -lt $OdsFiles.Count; $i++) {
        $num = $i + 1
        Write-Host ("{0}. {1}" -f $num, $OdsFiles[$i].Name)
    }

    $Choice = Read-Host "Выберите номер ODS-файла"

    [int]$ChoiceNumber = 0
    if (-not [int]::TryParse($Choice, [ref]$ChoiceNumber)) {
        throw "Нужно ввести число."
    }

    $Index = $ChoiceNumber - 1
    if (($Index -lt 0) -or ($Index -ge $OdsFiles.Count)) {
        throw "Номер вне диапазона."
    }

    $Selected = $OdsFiles[$Index]
    $PythonScript = Join-Path -Path $ScriptDir -ChildPath 'fill_signals_ods.py'

    if (-not (Test-Path -Path $PythonScript)) {
        throw ("Не найден файл: " + $PythonScript)
    }

    $PythonCmd = Install-PythonIfNeeded
    Install-MissingPythonPackages -PythonCmd $PythonCmd

    Write-Info ("Выбран файл: " + $Selected.Name)
    Write-Info "Запуск заполнения..."

    $Output = @()
    if ($PythonCmd -eq 'py') {
        $Output = & py -3 $PythonScript $Selected.FullName 2>&1
        $PythonExitCode = $LASTEXITCODE
    }
    else {
        $Output = & python $PythonScript $Selected.FullName 2>&1
        $PythonExitCode = $LASTEXITCODE
    }

    foreach ($line in $Output) {
        Write-PythonOutputLine -Text ([string]$line)
    }

    if ($PythonExitCode -ne 0) {
        throw ("Скрипт Python завершился с ошибкой. Код: " + $PythonExitCode)
    }

    Write-Info "Готово. Проверьте ODS и резервную копию *.backup_before_fill.ods"
}
catch {
    Write-ErrMsg $_.Exception.Message
}
finally {
    Pause-BeforeExit
}
