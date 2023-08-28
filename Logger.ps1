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
        if ($creationTime -lt $files[$i].CreationTime)
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
    # лог для записи в файл
    [Parameter(Mandatory=$true)]$log
    # каталог для создания файла логов
    [Parameter(Mandatory=$true)]$logsFolderPath

    
    [string]$logFile = "$($logsFolderPath)\$((Get-date).ToString("dd.MM.yyyy HH.MM")).log" 
    $log | Set-Content -Path $logFile
}

# расположение каталога для хранения логов
[string]$logPath = "C:\Users\$($env:USERNAME)\Desktop\"

# название каталога с логами
[string]$folderName = "MountPrinterLogs"
[string]$logsFolderPath = "$($logPath)$($folderName)"

# лимит файлов с логами. Если логов больще - старые удаляем
[uint16]$limitFiles = 3

if (Test-Path ($logPath))
{
    [uint16]$countExistFiles = (Get-ChildItem `
        -Path $logsFolderPath `
        -Filter "*.log" | Measure-Object).Count

    if ($countExistFiles -lt $limitFiles)
    {
        
    }
    else
    {
        Remove-OlderFile -logsFolderPath $logsFolderPath 
    }

}
else
{
    New-Item -Path $logPath -Name $folderName -ItemType Directory
}