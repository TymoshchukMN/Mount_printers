<#
    Скрипт для монтирвоания принтеров в магазинах.
    Монтирование выполняется на основе групп "*printers access*"
    Например "Терра printers access", или "Апполо printers access"
    
    Автор: TymoshchukMN
    Создан: 02.08.2023
    Изменен 28.08.2023
#>

<#
.Synopsis
   Получение списка групп по маске "*printers access*"
.DESCRIPTION
   Получение списка групп в которых состоит пользователь. 
   Если список групп в которых состоит пользователь,  пустой, 
   или возникла ошибка при выполнении LDAP-запроса, бросаем исключение.
   Иначе, проверяем есть ли группы с маской "*printers access*" в названии
#>
function Get-UsersGroupsList()
{
    param(
        [Parameter(Mandatory=$true)][string]$userName
    )

    $LDAP = New-Object System.DirectoryServices.DirectorySearcher
    $LDAP.Filter = ("(&(objectCategory=User)(samAccountName=$($userName)))")

    # получаем список групп пользователя
    $groupList =($LDAP.FindOne()).Properties.memberof

    if([string]::IsNullOrEmpty($groupList) -or $groupList.Count -lt 1)
    {
        $LDAP.Dispose()
        Remove-Variable -Name LDAP -ErrorAction SilentlyContinue
        
        throw "Не удалось получить список групп по пользователю <font color=red>$($userName) используя LDAP</font>"
    }
    else
    {
        [string]$groupMask = "*printers access*"

        return $groupList | Where-Object {$_ -like "$($groupMask)"}
    } 
}

<#
.Synopsis
   Получение кода магазина
.DESCRIPTION
   Получение кода магазина на основании distinguishedName группы.
   Если не удалось получить группу по distinguishedName
   или возникла ошибка при выполнении LDAP-запроса, бросаем исключение
#>
function Get-CodeByGroup
{
    param(
        [Parameter(Mandatory=$true)][string]$distinguishedName
    )

    # LDAP-фильтр для посика группы
    [string]$ldapFilter = "(&(objectCategory=group)(objectClass=group)(distinguishedName=$($distinguishedName)))"

    $LDAP = New-Object System.DirectoryServices.DirectorySearcher
    $LDAP.Filter = $ldapFilter

    # получаем samaccountname группы по distinguishedName
    $group = $LDAP.Findall().Properties.samaccountname

    if([string]::IsNullOrEmpty($group) -or $group.Count -lt 1)
    {
        $LDAP.Dispose()
        Remove-Variable -Name LDAP -ErrorAction SilentlyContinue
        
        throw "Не удалось получить группу по distinguishedName.
            <font color=red>$($distinguishedName)</font>"
    }
    else
    {
        # получаем код магазина. Это первые 3 символа 
        # в distinguishedName группы
        [string]$code = ([string]$group).Substring(0,3)
    }

    return $code
}

<#
.DESCRIPTION
   Создание таблицы для отправки отчета об ошибке
#>
function Create-HTMLTable()
{
    $HtmlTable = "<table border='1' align='Left' cellpadding='2' cellspacing='0' style='color:black;font-family:arial,helvetica,sans-serif;text-align:left;'>
    <tr style ='font-size:12px;font-weight: normal;background: #FFFFFF;background-color: #C0E979;'>
        <th align=left>
            <b>
                Login
            </b>
        </th>
        <th align=left>
            <b>
                Server name
            </b>
        </th>
        <th align=left>
            <b>
                Script path
            </b>
        </th>
        <th align=left>
            <b>
               Подключение по RDP с ПК
            </b>
        </th>
        <th align=left>
            <b>
               Время возникновения
            </b>
        </th>
     </tr>"
    
    
    $HtmlTable += 
    "<tr style='font-size:12px;background-color:#FFFFFF'>
        <td>" + $env:USERNAME + "</td>
        <td>" + $env:COMPUTERNAME  + "</td>
        <td>" + "$($MyInvocation.ScriptName)" + "</td>
        <td>" + $env:CLIENTNAME + "</td>
        <td>" + $(get-date) + "</td>           
    </tr>"

    return $HtmlTable
}

<#
.Synopsis
   Получение принтеров по коду для монтирования
.DESCRIPTION
   Получение принтеров по коду для монтирования с принт-сервера lt-printsrv2.
   Если принтеров нет, бросаем исключение
   Если группа есть, возвращаем код магазина
#>
function Get-Printers(
    [string]$shopcode,
    [string]$printSRV,
    $log)
{
    [array]$printers =
        @(Get-Printer -ComputerName $printSRV | `
            ? {($_.Name -like ("*_"+$shopcode))})

    # проверяем количество принтеров, если меньше 1,
    # т.е. принтеров нет, бросаем исключение
    if($printers.Count -lt 1)
    {
        $log += "На принт сервере $($printSRV) нет принтера с кодом $($shopcode)"
    }
    else
    {
        return $printers
    }
}

<#
.Synopsis
   Отправка сообщения об ошибке
.DESCRIPTION
   Отправка сообщения об ошибке выполнения скрипта. 
   Отправка выполняеся от имени error-sender
#>
function Send-Mail()
{
    param(
        [Parameter(Mandatory=$true)][string]$message
    )

    $encoding = [System.Text.Encoding]::UTF8;    

    # получаем конфиг для отпарвки письма
    $json = Get-Content .\mailConfig.json
    $mailConfig = $json | ConvertFrom-Json

    [string]$keyPath = $mailConfig.keyPath
    [string]$cryptedPass =  $mailConfig.cryptedPass
    
    $password = Get-Content $cryptedPass | `
        ConvertTo-SecureString -Key (Get-Content $keyPath)

    $EmailCredential = 
        New-Object -TypeName System.Management.Automation.PSCredential `
        -ArgumentList $mailConfig.sender,$Password
    
    Send-MailMessage -From error-sender@comfy.ua `
        -To GP-script-processing@comfy.ua `
        -Subject "Ошибка выполнения скрипта" `
        -Body $($message | Out-String) -BodyAsHtml -Encoding $encoding `
        -SmtpServer $mailConfig.SmtpServer -Credential $EmailCredential
}

# ============= начало скрипта ================

[string]$username = $env:USERNAME
[array]$logs = @()
$logs += "Текущий пользователь - $($username)"
#region получение списка групп пользователя

try
{
    # получение списка групп пользователя по маске "*printers access*"
    [array]$groupList = Get-UsersGroupsList -userName $username

    # если в списке нет групп с маской  "*printers access*"
    # завершаем выполнение скрипта
    if([string]::IsNullOrEmpty($groupList) -or $groupList.Count -lt 1)
    {
        $logs += "Пользоавтель $($username) не состоит ни в одной группе 'printers access'"
        .\Logger.ps1 -log $logs
        exit
    }
    else
    {
        $logs += "Пользоавтель состоит в группе $($groupList)"
    }
}
catch
{
    $HtmlTable = Create-HTMLTable

    # Добавляем в тело сообщения, тескт исключения
    $message = "<p>$($_)<br></p>"
     
    $message += $HtmlTable
    Send-Mail -message $message
}

#endregion получение списка групп пользователя

#region получение кода магазина

[array]$codes = @()
try
{
    # получение кода магазина из группы
    foreach ($distinguishedName in $groupList)
    {
        $codes += Get-CodeByGroup -distinguishedName $distinguishedName
        $logs += "По группе получен код - $($codes)"
    }
}
catch
{
    $logs += "Не удалось получить группу по distinguishedName - $($distinguishedName)."
    .\Logger.ps1 -log $logs
    $HtmlTable = Create-HTMLTable

    # Добавляем в тело сообщения, тескт исключения
    $message = "<p>$($_)<br></p>"
     
    $message += $HtmlTable
    Send-Mail -message $message
}

#endregion получение кода магазина

#region Обработка принтеров

[string]$printSRV = "lt-printsrv2" 

$printers = @()
for ([int16]$i = 0; $i -lt $codes.Count; ++$i)
{
    $printers += Get-Printers -shopcode $codes[$i] -printSRV $printSRV -log $logs
}

$logs += "Полученные принтера: $($printers)"

# удаляем все подключенные по сети принтера
Get-Printer | ? {$_.Type -eq "Connection"} | Remove-Printer

# добавление принтеров
foreach ($printer in $printers)
{
    $printername=("\\"+ $printSRV +"\"+$printer.Name)

    Add-Printer -ConnectionName $printername -Confirm:$false

    if ($username -match "touch" -or $username -match "zal")
    {
		(New-Object -ComObject WScript.Network).SetDefaultPrinter($printername)
	}
}

#endregion Обработка принтеров

.\Logger.ps1 -log $logs