#requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$NoPause
)

# Настройка кодировки для корректного русского текста в Windows PowerShell 5.1.
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
        $Command = Get-Command $CommandName -ErrorAction Stop
        return $null -ne $Command
    }
    catch {
        return $false
    }
}

function Refresh-ProcessPath {
    $MachinePath = [System.Environment]::GetEnvironmentVariable('Path', 'Machine')
    $UserPath = [System.Environment]::GetEnvironmentVariable('Path', 'User')

    $PathParts = @()
    if ($MachinePath) {
        $PathParts += $MachinePath
    }
    if ($UserPath) {
        $PathParts += $UserPath
    }

    $env:Path = $PathParts -join ';'
}

function Test-PythonCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Candidate
    )

    try {
        $CandidateName = [System.IO.Path]::GetFileNameWithoutExtension($Candidate)

        if ($CandidateName -ieq 'py') {
            & $Candidate -3 --version *> $null
        }
        else {
            & $Candidate --version *> $null
        }

        return ($LASTEXITCODE -eq 0)
    }
    catch {
        return $false
    }
}

function Get-PythonCommand {
    $Candidates = New-Object System.Collections.Generic.List[string]

    if (Test-CommandExists -CommandName 'py') {
        [void]$Candidates.Add('py')
    }

    if (Test-CommandExists -CommandName 'python') {
        [void]$Candidates.Add('python')
    }

    $SearchRoots = @()

    if ($env:LOCALAPPDATA) {
        $SearchRoots += (Join-Path $env:LOCALAPPDATA 'Programs\Python')
    }

    if ($env:ProgramFiles) {
        $SearchRoots += $env:ProgramFiles
    }

    if (${env:ProgramFiles(x86)}) {
        $SearchRoots += ${env:ProgramFiles(x86)}
    }

    foreach ($Root in $SearchRoots) {
        if (-not (Test-Path -LiteralPath $Root)) {
            continue
        }

        $PythonExecutables = @()

        if ($Root -like '*Programs\Python') {
            $PythonExecutables = Get-ChildItem `
                -LiteralPath $Root `
                -Directory `
                -Filter 'Python*' `
                -ErrorAction SilentlyContinue |
                Sort-Object -Property Name -Descending |
                ForEach-Object {
                    Join-Path $_.FullName 'python.exe'
                }
        }
        else {
            $PythonExecutables = Get-ChildItem `
                -LiteralPath $Root `
                -Directory `
                -Filter 'Python*' `
                -ErrorAction SilentlyContinue |
                Sort-Object -Property Name -Descending |
                ForEach-Object {
                    Join-Path $_.FullName 'python.exe'
                }
        }

        foreach ($PythonExecutable in $PythonExecutables) {
            if (Test-Path -LiteralPath $PythonExecutable) {
                [void]$Candidates.Add($PythonExecutable)
            }
        }
    }

    $Seen = @{}

    foreach ($Candidate in $Candidates) {
        $Key = $Candidate.ToLowerInvariant()

        if ($Seen.ContainsKey($Key)) {
            continue
        }

        $Seen[$Key] = $true

        if (Test-PythonCandidate -Candidate $Candidate) {
            return [string]$Candidate
        }
    }

    return $null
}

function Invoke-Python {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonCmd,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [switch]$Silent
    )

    [string[]]$InvocationArguments = @()

    $CandidateName = [System.IO.Path]::GetFileNameWithoutExtension($PythonCmd)
    if ($CandidateName -ieq 'py') {
        $InvocationArguments += '-3'
    }

    $InvocationArguments += $Arguments

    if ($Silent) {
        & $PythonCmd @InvocationArguments *> $null
    }
    else {
        # Выводим сообщения программы пользователю, но не возвращаем их из функции.
        # Иначе PowerShell превратит результат функции в массив вместе с кодом возврата.
        & $PythonCmd @InvocationArguments 2>&1 |
            ForEach-Object {
                Write-Host ([string]$_)
            }
    }

    $NativeExitCode = $LASTEXITCODE
    return [int]$NativeExitCode
}

function Install-PythonIfNeeded {
    $PythonCmd = Get-PythonCommand
    if ($PythonCmd) {
        return [string]$PythonCmd
    }

    Write-WarnMsg "Python 3 не найден."

    $Answer = Read-Host "Установить Python 3 автоматически через winget? (Y/N)"
    if ($Answer -notmatch '^(Y|y|Д|д)$') {
        throw "Python 3 не установлен. Установка отменена пользователем."
    }

    if (-not (Test-CommandExists -CommandName 'winget')) {
        throw "Не найден winget. Установите Python 3 вручную и повторите запуск."
    }

    Write-Info "Устанавливаю Python 3..."

    [string[]]$WingetArguments = @(
        'install',
        '--exact',
        '--id', 'Python.Python.3.12',
        '--source', 'winget',
        '--silent',
        '--disable-interactivity',
        '--accept-package-agreements',
        '--accept-source-agreements'
    )

    # Вывод winget направляется прямо в консоль и не становится частью
    # возвращаемого значения Install-PythonIfNeeded.
    & winget @WingetArguments 2>&1 |
        ForEach-Object {
            Write-Host ([string]$_)
        }

    $WingetExitCode = $LASTEXITCODE

    # Обновляем PATH текущего процесса и несколько раз проверяем установку.
    # Дополнительно Get-PythonCommand ищет python.exe по типовым каталогам,
    # поэтому перезапуск окна после установки обычно не требуется.
    for ($Attempt = 1; $Attempt -le 10; $Attempt++) {
        Refresh-ProcessPath
        $PythonCmd = Get-PythonCommand

        if ($PythonCmd) {
            Write-Info "Python 3 успешно установлен и найден."
            return [string]$PythonCmd
        }

        Start-Sleep -Seconds 1
    }

    if ($WingetExitCode -ne 0) {
        throw (
            "Не удалось установить Python 3 через winget. " +
            "Код завершения: " + $WingetExitCode
        )
    }

    throw (
        "Python установлен, но исполняемый файл не удалось найти. " +
        "Закройте окно, откройте его снова и повторите запуск."
    )
}

function Test-PythonModule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonCmd,

        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    $Code = (
        "import importlib.util, sys; " +
        "sys.exit(0 if importlib.util.find_spec('" + $ModuleName + "') else 1)"
    )

    $ExitCode = Invoke-Python `
        -PythonCmd $PythonCmd `
        -Arguments @('-c', $Code) `
        -Silent

    return ($ExitCode -eq 0)
}

function Ensure-Pip {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonCmd
    )

    Write-Info "Проверяю pip..."

    $ExitCode = Invoke-Python `
        -PythonCmd $PythonCmd `
        -Arguments @('-m', 'pip', '--version') `
        -Silent

    if ($ExitCode -eq 0) {
        Write-Info "pip найден."
        return
    }

    Write-WarnMsg "pip не найден. Пытаюсь включить его через ensurepip..."

    $ExitCode = Invoke-Python `
        -PythonCmd $PythonCmd `
        -Arguments @('-m', 'ensurepip', '--upgrade') `
        -Silent

    if ($ExitCode -ne 0) {
        throw "В выбранном Python недоступен pip, и не удалось включить его через ensurepip."
    }

    $ExitCode = Invoke-Python `
        -PythonCmd $PythonCmd `
        -Arguments @('-m', 'pip', '--version') `
        -Silent

    if ($ExitCode -ne 0) {
        throw "pip не удалось подготовить после ensurepip."
    }

    Write-Info "pip успешно подготовлен."
}

function Get-MissingPythonPackages {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonCmd,

        [Parameter(Mandatory = $true)]
        [object[]]$RequiredModules
    )

    $MissingPackages = New-Object System.Collections.Generic.List[string]

    foreach ($Item in $RequiredModules) {
        if (-not (Test-PythonModule `
            -PythonCmd $PythonCmd `
            -ModuleName ([string]$Item.Module))) {

            [void]$MissingPackages.Add([string]$Item.Package)
        }
    }

    return @($MissingPackages)
}

function Install-MissingPythonPackages {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonCmd
    )

    Ensure-Pip -PythonCmd $PythonCmd

    $RequiredModules = @(
        @{ Module = 'odf';      Package = 'odfpy'    },
        @{ Module = 'openpyxl'; Package = 'openpyxl' }
    )

    $MissingPackages = New-Object System.Collections.Generic.List[string]

    foreach ($Item in $RequiredModules) {
        $ModuleName = [string]$Item.Module
        $PackageName = [string]$Item.Package

        if (Test-PythonModule -PythonCmd $PythonCmd -ModuleName $ModuleName) {
            Write-Info ("Пакет уже установлен: " + $PackageName)
        }
        else {
            Write-WarnMsg ("Не найден пакет: " + $PackageName)
            [void]$MissingPackages.Add($PackageName)
        }
    }

    if ($MissingPackages.Count -eq 0) {
        Write-Info "Все необходимые Python-пакеты уже установлены."
        return
    }

    $PackageText = $MissingPackages -join ', '
    $Answer = Read-Host (
        "Установить отсутствующие пакеты: " + $PackageText + " ? (Y/N)"
    )

    if ($Answer -notmatch '^(Y|y|Д|д)$') {
        throw "Установка Python-пакетов отменена пользователем."
    }

    Write-Info ("Устанавливаю пакеты: " + $PackageText)

    [string[]]$InstallArguments = @(
        '-m',
        'pip',
        'install',
        '--disable-pip-version-check'
    )

    foreach ($PackageName in $MissingPackages) {
        $InstallArguments += [string]$PackageName
    }

    $FirstExitCode = Invoke-Python `
        -PythonCmd $PythonCmd `
        -Arguments $InstallArguments

    $RemainingPackages = Get-MissingPythonPackages `
        -PythonCmd $PythonCmd `
        -RequiredModules $RequiredModules

    # Критерий успеха — возможность импортировать модули, а не только код pip.
    if ($RemainingPackages.Count -eq 0) {
        Write-Info "Необходимые Python-пакеты успешно установлены."
        return
    }

    Write-WarnMsg (
        "Не все пакеты доступны после обычной установки. " +
        "Повторяю установку только для текущего пользователя..."
    )

    [string[]]$UserInstallArguments = @(
        '-m',
        'pip',
        'install',
        '--user',
        '--disable-pip-version-check'
    )

    foreach ($PackageName in $RemainingPackages) {
        $UserInstallArguments += [string]$PackageName
    }

    $SecondExitCode = Invoke-Python `
        -PythonCmd $PythonCmd `
        -Arguments $UserInstallArguments

    $RemainingPackages = Get-MissingPythonPackages `
        -PythonCmd $PythonCmd `
        -RequiredModules $RequiredModules

    if ($RemainingPackages.Count -ne 0) {
        throw (
            "Не удалось установить Python-пакеты: " +
            ($RemainingPackages -join ', ') +
            ". Коды pip: " + $FirstExitCode + ", " + $SecondExitCode
        )
    }

    Write-Info "Необходимые Python-пакеты успешно установлены."
}

function Write-PythonOutputLine {
    param([string]$Text)

    if ($Text.StartsWith('[ОШИБКА]')) {
        Write-Host $Text -ForegroundColor Red
        return
    }

    if ($Text.StartsWith('[ПРЕДУПРЕЖДЕНИЕ]')) {
        Write-Host $Text -ForegroundColor Yellow
        return
    }

    if ($Text -like 'NEW_MODULES_COUNT:*') {
        if ($Text -match 'NEW_MODULES_COUNT:\s*0') {
            Write-Host `
                "Новых модулей в XLSX, отсутствующих в ODS: нет" `
                -ForegroundColor Green
        }
        else {
            $CountText = $Text -replace '^NEW_MODULES_COUNT:\s*', ''
            Write-Host (
                "Новых модулей в XLSX, отсутствующих в ODS: " + $CountText
            ) -ForegroundColor Yellow
        }
        return
    }

    if ($Text -like 'NEW_MODULE:*') {
        $Payload = $Text -replace '^NEW_MODULE:\s*', ''
        $Parts = $Payload -split '\s*\|\s*', 2

        if ($Parts.Count -eq 2) {
            Write-Host "  - " -NoNewline
            Write-Host $Parts[0] -ForegroundColor Yellow -NoNewline
            Write-Host ": " -NoNewline
            Write-Host $Parts[1] -ForegroundColor Red
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
            $CountText = $Text -replace '^DUPLICATE_KEYS_COUNT:\s*', ''
            Write-Host (
                "Дубликатов ключей в XLSX: " + $CountText
            ) -ForegroundColor Magenta
        }
        return
    }

    if ($Text -like 'DUPLICATE_KEY:*') {
        $Payload = $Text -replace '^DUPLICATE_KEY:\s*', ''
        $Parts = $Payload -split '\s*\|\s*', 5

        if ($Parts.Count -eq 5) {
            Write-Host "  - " -NoNewline
            Write-Host $Parts[0] -ForegroundColor Magenta -NoNewline
            Write-Host ": " -NoNewline
            Write-Host $Parts[1] -ForegroundColor Yellow -NoNewline
            Write-Host ", канал " -NoNewline
            Write-Host $Parts[2] -ForegroundColor Cyan -NoNewline
            Write-Host ", тип " -NoNewline
            Write-Host $Parts[3] -ForegroundColor Cyan -NoNewline
            Write-Host ", строки XLSX: " -NoNewline
            Write-Host $Parts[4] -ForegroundColor Red
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

$ScriptExitCode = 0

try {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    Set-Location -LiteralPath $ScriptDir

    Write-Info ("Рабочая папка: " + $ScriptDir)

    $OdsFiles = @(
        Get-ChildItem `
            -LiteralPath $ScriptDir `
            -File `
            -Filter '*.ods' |
            Where-Object {
                $_.Name -notlike '*.backup_before_fill.ods'
            } |
            Sort-Object -Property Name
    )

    if ($OdsFiles.Count -eq 0) {
        throw "В папке не найдено ни одного .ods файла."
    }

    Write-Host "Найденные ODS-файлы:"

    for ($Index = 0; $Index -lt $OdsFiles.Count; $Index++) {
        Write-Host ("{0}. {1}" -f ($Index + 1), $OdsFiles[$Index].Name)
    }

    $Choice = Read-Host "Выберите номер ODS-файла"

    [int]$ChoiceNumber = 0
    if (-not [int]::TryParse($Choice, [ref]$ChoiceNumber)) {
        throw "Нужно ввести число."
    }

    $SelectedIndex = $ChoiceNumber - 1
    if (($SelectedIndex -lt 0) -or ($SelectedIndex -ge $OdsFiles.Count)) {
        throw "Номер вне диапазона."
    }

    $Selected = $OdsFiles[$SelectedIndex]
    $PythonScript = Join-Path $ScriptDir 'fill_signals_ods.py'

    if (-not (Test-Path -LiteralPath $PythonScript)) {
        throw ("Не найден файл: " + $PythonScript)
    }

    $PythonCmd = Install-PythonIfNeeded
    Install-MissingPythonPackages -PythonCmd $PythonCmd

    Write-Info ("Выбран файл: " + $Selected.Name)
    Write-Info "Запуск заполнения..."

    [string[]]$PythonArguments = @(
        $PythonScript,
        $Selected.FullName,
        '--structured-output'
    )

    $Output = @()

    if ([System.IO.Path]::GetFileNameWithoutExtension($PythonCmd) -ieq 'py') {
        $Output = & $PythonCmd -3 @PythonArguments 2>&1
    }
    else {
        $Output = & $PythonCmd @PythonArguments 2>&1
    }

    $PythonExitCode = $LASTEXITCODE

    foreach ($Line in $Output) {
        Write-PythonOutputLine -Text ([string]$Line)
    }

    if ($PythonExitCode -ne 0) {
        throw (
            "Скрипт Python завершился с ошибкой. Код: " +
            $PythonExitCode
        )
    }

    Write-Info (
        "Готово. Проверьте ODS и резервную копию " +
        "*.backup_before_fill.ods"
    )
}
catch {
    $ScriptExitCode = 1
    Write-ErrMsg ("[ОШИБКА] " + $_.Exception.Message)
}
finally {
    if (-not $NoPause) {
        Pause-BeforeExit
    }
}

# Передаём CMD-оболочке корректный код завершения,
# не закрывая родительский интерактивный PowerShell при прямом запуске.
[System.Environment]::ExitCode = $ScriptExitCode
