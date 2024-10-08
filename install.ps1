function InstallProgramsViaChocolatey {
   
    $programs = @(
        "lightshot.install",
        "notepadplusplus",
        "winrar",
        "anydesk",
        "googlechrome",
        "firefox",
        "openvpn",
        "vnc-viewer --version=6.21.1109",
        "filezilla",
        "obsidian",
        "virtualbox",
        "microsip",
        "telegram.install",
        "mobaxterm",
        "vcredist-all",
        "obsidian",
        "syncthing"

    )

    
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    }


    foreach ($obj in $programs) { 
        choco install $obj -y
    }
    Set-ExecutionPolicy default -Scope Process -Force
    
}

function Install-MangoTalker {
    $arguments = "--silent -l=ru --sh"
    & { (New-Object System.Net.WebClient).DownloadFile('https://mtalker.mango-office.ru/MangoTalkerUpdater/download?path=versions/windows/latest/&utm_source=letter_wind&utm_medium=e-mail&utm_campaign=vats_talker', "$env:UserProfile\Downloads\mango_talker.exe") }
    $mangopath = "$env:UserProfile\Downloads\mango_talker.exe"
    Start-Process -FilePath $mangopath -ArgumentList $arguments -Wait
}

function Install-MicrosoftOffice {
    $odtFolderPath = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop\apps\ODT')
    $setupPath = [System.IO.Path]::Combine($odtFolderPath, 'setup.exe')
    $configPath = [System.IO.Path]::Combine($odtFolderPath, 'config_365_enterprise_fullsupport_noAccess.xml')

    if (Test-Path $setupPath) {
        if (Test-Path $configPath) {
            reg add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /v "CountryCode" /t REG_SZ /d "std::wstring|US" /f
            Start-Process -FilePath $setupPath -ArgumentList @("/configure", $configPath) -Wait
        } else {
            Write-Host "Конфигурационный файл не найден."
        }
    } else {
        Write-Host "Файл установки Microsoft Office не найден."
    }
    & ([ScriptBlock]::Create((irm https://get.activated.win))) /HWID /S
}

function Replace1cFolder {
    $sourceFolder = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop\apps\1cv8')
    $destinationFolder = "C:\Program Files\1cv8"

    if (Test-Path $sourceFolder) {
        Move-Item -Path $sourceFolder -Destination $destinationFolder -Force
    } else {
        Write-Host "Папка '1cv8' не найдена на рабочем столе всех пользователей."
    }
}

#Вызов функций
InstallProgramsViaChocolatey
Install-MangoTalker
Install-MicrosoftOffice
Replace1cFolder