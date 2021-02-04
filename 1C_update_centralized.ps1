#Created by https://github.com/VladimirKosyuk

# Обновляет 1С клиент в соответствии с версией сервера централизованно

# Build date: 04/02/2020

#About:

<# 

Функциональное описание: 

-Создает файл лога скрипта, если он не создан
-Собирает массив имен клиентских ПК windows из AD, далее делает для каждого из участников массива
-Копирует установочные файлы 1С на диск С:
-Ищет установленные программы, название которых начинается с 1С
-Если находит - сверяет версию с текущей версией сервера
-Если версия различается - пишет в лог, удаляет текущую версию, устанавливает актуальную версию в зависимости от разрядности ОС (если установочный файл доступен)
-Если версия не различается - удаляет установочные файлы и переходит к следующему ПК
-Если выполнение шагов выше занимает больше 10 минут - отменяет задание и удаляет установочные файлы, потом переходит к следующему ПК

Требования к внедрению:

-Powersheel минимально версии 3;
-Powersheel ExecutionPolicy Unrestricted;
-Запуск с повышенными привилегиями;
-Доступ к $Repo, setup.exe, серверу 1С;
-Минимум 500 МБ свободного места на диске C:
-Путь к файлу установки должны быть вида $Repo+"*-64\setup.exe или $Repo+"*-32\setup.exe, где * любое допустимое множество символов.

#>	

#DCs server IP
$DC_srv = 
#1C server DNS name
$SrvAddr = 
#Where 1C installer is placed
$Repo = 
#where to output script log
$Log = 

$Date = Get-Date -Format "dd_MM_yyyy"
$GlobalLog = $Log +$Date+"_client_update.log"
#create global log file, if not exist
if (!(Get-ChildItem $GlobalLog -ErrorAction SilentlyContinue)) {New-Item  $GlobalLog -ErrorAction SilentlyContinue | Out-Null}
#get server version
$Srv = Get-WmiObject Win32_Product -ComputerName $SrvAddr -ErrorAction SilentlyContinue | Where-Object {$_.Name -like "*1C:Предприятие 8 (x86-64)*"}

#collect PCs array, need to be limited
$List_SRVs = Get-ADComputer -server $DC_srv  -Filter * -properties *|
Where-Object {$_.enabled -eq $true}|
Where-Object {($_.OperatingSystem -notlike "*Windows Server*") -and ($_.OperatingSystem -like "*Windows*")}|
Select-Object -ExpandProperty "DNSHostName"

foreach ($List_SRV in $List_SRVs){

#coping installation files
$SMB_Repo = "\\"+$List_SRV+"\C$\Temp_1C"
Copy-Item -Path $Repo -Destination $SMB_Repo -Recurse

$1C_Update = Invoke-Command -ErrorAction SilentlyContinue -ComputerName $List_SRV  -ArgumentList $Srv, $GlobalLog -ScriptBlock {
param($Srv, $GlobalLog)

#create setup.exe path vars
$Repo_local = "C:\Temp_1C\"
$Path64 = $Repo_local+"*-64\setup.exe"
$Client64 = (Get-ChildItem $Path64).FullName
$Path32 = $Repo_local+"*-32\setup.exe"
$Client32 = (Get-ChildItem $Path32).FullName

$IsInstalled = Get-WmiObject Win32_Product -ErrorAction SilentlyContinue | Where-Object {$_.Name -match "^(1С|1C)"}   
    #check if 1C is installed and not match server version
    if (($IsInstalled) -and ($IsInstalled.Version -notmatch $Srv.Version)){
    #check if global log is reachable
        if ((!(Test-Path $GlobalLog) -eq "True")){
            Write-Output ($env:computername+" "+$GlobalLog+" "+"is unreachable, output will be local")
            $GlobalLog = "C:\1C_Update.log"
            }
    Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+" "+"current version"+" "+($IsInstalled.Version)+" "+"not matching server version"+" "+($Srv.Version)) | Out-File $GlobalLog -Append
    Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+" "+"deleting current version") | Out-File $GlobalLog -Append
    #delete current 1C
    ForEach($Installed in $IsInstalled){$Installed.Uninstall()| Out-Null}
    #get OS architecture
    $Architec = Get-WmiObject Win32_OperatingSystem | Select-Object -Property OSArchitecture
        if ($Architec.OSArchitecture -like "64*") {
                if ((Test-Path $Client64) -eq "True"){
                Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+" "+"installing new version from "+" "+$Client64) | Out-File $GlobalLog -Append
                    #install 1C 64
                    try {Start-Process -FilePath $Client64  /S -NoNewWindow -Wait -PassThr
                        Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+" "+"updated, current version is"+" "+((Get-WmiObject Win32_Product -ErrorAction SilentlyContinue | Where-Object {$_.Name -match "^(1С|1C)"}).Version)) | Out-File $GlobalLog -Append}
                    catch{Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+($Error[0].Exception.Message ))| Out-File $GlobalLog -Append}
                    }
                    else {Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+" "+$Client64+" "+"is unreachable")| Out-File $GlobalLog -Append}
            }
        else {
            if ((Test-Path $Client32) -eq "True"){
            Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+" "+"installing new version from "+" "+$Client32) | Out-File $GlobalLog -Append
                #install 1C 32
                try {Start-Process -FilePath $Client32  /S -NoNewWindow -Wait -PassThr
                    Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+" "+"updated, current version is"+" "+((Get-WmiObject Win32_Product -ErrorAction SilentlyContinue | Where-Object {$_.Name -match "^(1С|1C)"}).Version)) | Out-File $GlobalLog -Append}
                catch{Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+($Error[0].Exception.Message ))| Out-File $GlobalLog -Append}
                }
                else {Write-Output ($env:COMPUTERNAME+" "+(Get-Date)+" "+$Client32+" "+"is unreachable")| Out-File $GlobalLog -Append}
            }
    }
}  -AsJob

#timeout set up to 10 minutes to execute invoke job
Wait-Job $1C_Update -Timeout 600
if (!($1C_Update.State -eq 'Completed')){Stop-Job -Id $1C_Update.Id}

#removing installation files
remove-item $SMB_Repo -Recurse -force

}
Remove-Variable -Name *  -Force -ErrorAction SilentlyContinue
