<# Логгер
    Автор: TymoshchukMN
    Создан 28.08.2023
#>

param(
    # строка лога
    [Parameter(Mandatory=$true)]$log
)

<#
.SYNOPSIS
    Удалить старый файл
.DESCRIPTION
    Удалене стрейшего файла в каталоге
#>
function Remove-OlderFile
{
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $logsFolderPath
    )
    
    $files = Get-ChildItem -Path $logsFolderPath -Filter "*.log"

    [uint16]$indexOlderFile = 0
    [datetime]$creationTime = $files[$indexOlderFile].CreationTime

    for ([uint16]$i = 1; $i -lt [uint16]$files.Count; ++$i)
    {
        if ($creationTime -gt $files[$i].CreationTime)
        {
            $creationTime = $files[$i].CreationTime
            $indexOlderFile = $i
        }
    }
    
    Remove-Item -Path $files[$indexOlderFile].FullName -Force
}

<#
.SYNOPSIS
    Запись лога в файл
#>
function Write-LogsToFile
{
    param(
        # лог для записи в файл
        [Parameter(Mandatory=$true)][string[]]$log,
        # каталог для создания файла логов
        [Parameter(Mandatory=$true)][string]$logsFolderPath
    )
    
    [string]$logFile = "$($logsFolderPath)\$((Get-date).ToString("dd.MM.yyyy HH.mm.ss")).log" 
    $log | Set-Content -Path $logFile
}

# расположение каталога для хранения логов
[string]$logPath = "C:\Users\$($env:USERNAME)\Desktop\"

# название каталога с логами
[string]$folderName = "MountPrinterLogs"
[string]$logsFolderPath = "$($logPath)$($folderName)"

# лимит файлов с логами. Если логов больще - старые удаляем
[uint16]$limitFiles = 3

if (Test-Path ($logsFolderPath))
{
    [uint16]$countExistFiles = (Get-ChildItem `
        -Path $logsFolderPath `
        -Filter "*.log" | Measure-Object).Count

    if ($countExistFiles -lt $limitFiles)
    {
        Write-LogsToFile -log $log -logsFolderPath $logsFolderPath
    }
    else
    {
        Remove-OlderFile -logsFolderPath $logsFolderPath
        Write-LogsToFile -log $log -logsFolderPath $logsFolderPath
    }
}
else
{
    New-Item -Path $logPath -Name $folderName -ItemType Directory
    Write-LogsToFile -log $log -logsFolderPath $logsFolderPath
}