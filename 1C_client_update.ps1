#Created by https://github.com/VladimirKosyuk

# Обновляет 1С клиент в соответствии с версией сервера 

# Build date: 02/02/2020

#About:

<# 

Функциональное описание: 

-На локальном ПК ищет установленные программы, название которых начинается с 1С;
-Если находит - сверяет версию с текущей версией сервера;
-Если версия различается - пишет в лог, если нет - далее ничего не делает;
-Удаляет текущую версию;
-В зависимости от разрядности ОС и наличия доступа к $Repo устанавливает актуальный клиент.

Требования к внедрению:

-Powersheel минимально версии 3;
-Powersheel ExecutionPolicy Unrestricted;
-Запуск с повышенными привилегиями;
-Доступ к $Repo, setup.exe, серверу 1С;
-Путь к файлу установки должны быть вида $Repo+"*-64\setup.exe или $Repo+"*-32\setup.exe, где * любое допустимое множество символов.

#>	

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
#create setup path
$Path64 = $Repo+"*-64\setup.exe"
$Client64 = (Get-ChildItem $Path64).FullName
$Path32 = $Repo+"*-32\setup.exe"
$Client32 = (Get-ChildItem $Path32).FullName

$IsInstalled = Get-WmiObject Win32_Product -ErrorAction SilentlyContinue | Where-Object {$_.Name -match "^(1С|1C)"}   
    #check if 1C is installed and not match server version
    if (($IsInstalled) -and ($IsInstalled.Version -notmatch $Srv.Version)){
    #check if global log is reachable
        if ((!(Test-Path $GlobalLog) -eq "True")){
            Write-Output ($env:computername+" "+$GlobalLog+" "+"is unreachable, output will be local")
            $GlobalLog = "C:\1C_Update.log"
            }
    Write-Output ($env:COMPUTERNAME+" "+(Get-Date -Format "HH:mm")+" "+"current version"+" "+($IsInstalled.Version)+" "+"not matching server version"+" "+($Srv.Version)) | Out-File $GlobalLog -Append
    Write-Output ($env:COMPUTERNAME+" "+(Get-Date -Format "HH:mm")+" "+"deleting current version") | Out-File $GlobalLog -Append
    #delete current 1C
    ForEach($Installed in $IsInstalled){$Installed.Uninstall()| Out-Null}
    #get OS architecture
    $Architec = Get-WmiObject Win32_OperatingSystem | Select-Object -Property OSArchitecture
        if ($Architec.OSArchitecture -like "64*") {
                if ((Test-Path $Client64) -eq "True"){
                Write-Output ($env:COMPUTERNAME+" "+(Get-Date -Format "HH:mm")+" "+"installing new version from "+" "+$Client64) | Out-File $GlobalLog -Append
                    #install 1C 64
                    try {Start-Process -FilePath $Client64  /S -NoNewWindow -Wait -PassThr}
                    catch{Write-Output ($env:COMPUTERNAME+" "+(Get-Date -Format "HH:mm")+($Error[0].Exception.Message ))| Out-File $GlobalLog -Append}
                    }
                    else {Write-Output ($env:COMPUTERNAME+" "+(Get-Date -Format "HH:mm")+" "+$Client64+" "+"is unreachable")| Out-File $GlobalLog -Append}
            }
        else {
            if ((Test-Path $Client32) -eq "True"){
            Write-Output ($env:COMPUTERNAME+" "+(Get-Date -Format "HH:mm")+" "+"installing new version from "+" "+$Client32) | Out-File $GlobalLog -Append
                #install 1C 32
                try {Start-Process -FilePath $Client32  /S -NoNewWindow -Wait -PassThr}
                catch{Write-Output ($env:COMPUTERNAME+" "+(Get-Date -Format "HH:mm")+($Error[0].Exception.Message ))| Out-File $GlobalLog -Append}
                }
                else {Write-Output ($env:COMPUTERNAME+" "+(Get-Date -Format "HH:mm")+" "+$Client32+" "+"is unreachable")| Out-File $GlobalLog -Append}
            }
    Write-Output ($env:COMPUTERNAME+" "+(Get-Date -Format "HH:mm")+" "+"updated, current version is"+" "+((Get-WmiObject Win32_Product -ErrorAction SilentlyContinue | Where-Object {$_.Name -match "^(1С|1C)"}).Version)) | Out-File $GlobalLog -Append
    }

Remove-Variable -Name *  -Force -ErrorAction SilentlyContinue   