cls
#Forbereder system
write-host "Tester forbindelse til printer 40 (Printer ved Booking-PC).. " -NoNewline; Sleep -s 3
if (Test-Connection  192.168.1.40 -Quiet) {
    write-host "[Forbindelse verificeret]".toUpper() -f green
    write-host "`t- Begynder installation af Printer 40:"; Sleep -s 5

    write-host "`t`t- Forbereder system.."
        $printername = "Printer 40 - Lager"
        $printdriverlink = "https://download.brother.com/welcome/dlf100988/Y14A_C1-hostm-1110.EXE"
        $printerinf = "C:\Printer\Printer 40 - Lager\32_64\BROHL13A.INF"
        $printerdriver = "Brother HL-L2360D series"
        $printerip = "192.168.1.40"
        $printerlocation = "Printer ved Booking-PC"
    
        $printerfolder = "C:\Printer\$printername"
        $file = Split-Path $printdriverlink -Leaf
        $portNumber = "91"+$printerip.Split(".")[-1]

        Stop-Service "Spooler"
        Remove-Item "C:\Windows\System32\spool\PRINTERS\*.*" -Force | Out-Null
        Start-Service "Spooler"

        # deaktiver automatisk installation af netværksprintere
        if((get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" | Select-Object -ExpandProperty AutoSetup) -ne 0)
        {Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Name "AutoSetup" -Type DWord -Value 0}
        # fjerner allerede installerede printere
        Get-Printer | ? Name -cMatch "OneNote (Desktop)|OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer
        Get-Printer | ? Name -Match "2365|$printername" | Remove-Printer -ea SilentlyContinue
        Get-PrinterPort | ? Name -match $printerip | Remove-PrinterPort -ea SilentlyContinue
        Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue
        # installér 7-zip hvis den ikke allerede er installeret. bruges til stabil driver udpakning.
        if(!(Test-Path "$env:ProgramFiles\7-Zip\7z.exe")){
            $dlurl = 'https://7-zip.org/' + (Invoke-WebRequest -Uri 'https://7-zip.org/' | Select-Object -ExpandProperty Links | Where-Object {($_.innerHTML -eq 'Download') -and ($_.href -like "a/*") -and ($_.href -like "*-x64.exe")} | Select-Object -First 1 | Select-Object -ExpandProperty href)
            $installerPath = Join-Path $env:TEMP (Split-Path $dlurl -Leaf)
            Invoke-WebRequest $dlurl -OutFile $installerPath -UseBasicParsing
            Start-Process -FilePath $installerPath -Args "/S" -Verb RunAs -Wait
            Remove-Item $installerPath} 

        new-item -ItemType Directory -Path $printerfolder -Force | out-null

    #Downloader driver
    write-host "`t`t- Downloader driver.."
        Remove-item -Path $printerfolder\* -Force -recurse | out-null
        (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$file")

    #Udpakker driver
    write-host "`t`t- Udpakker driver.."
        & ${env:ProgramFiles}\7-Zip\7z.exe x "$printerfolder\$file" "-o$($printerfolder)" -y | out-null; ; sleep -s 5

    #Installer Printer
    write-host "`t`t- Konfigurer Printer:"; sleep -s 5
        write-host "`t`t`t`t- Driverbiblotek"
        pnputil.exe -i -a $printerinf | out-null ; sleep -s 5
        write-host "`t`t`t`t- Driver"
        Add-PrinterDriver -Name $printerdriver | out-null; sleep -s 5
        write-host "`t`t`t`t- Printerport"
        Add-PrinterPort -Name $printerip -PrinterHostAddress $printerip -PortNumber $portnumber | out-null; sleep -s 5
        write-host "`t`t`t`t- Printer"
        Add-Printer -Name $printername -PortName $printerip -DriverName $printerdriver -PrintProcessor winprint -Location $printerlocation -Comment "automatiseret af Andreas" | out-null; sleep -s 5
    #Oprydning
        Remove-item  -Path "$printerfolder\" -Exclude $file -Recurse -Force

    write-host "`t- Printeren er installeret!" -f Green
}else {write-host "[INGEN FORBINDELSE]" -f red; write-host "`tDer er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}


#Forbereder system
write-host "Tester forbindelse til printer 30 (Printer ved Lones bord).. " -NoNewline; Sleep -s 3
if (Test-Connection  192.168.1.30 -Quiet) {
    write-host "[Forbindelse verificeret]".toUpper() -f green
    write-host "`t- Begynder installation af Printer 30:"; Sleep -s 5

    write-host "`t`t- Forbereder system.."
        $printername = "Printer 30 - Lager"
        $printdriverlink = "https://download.brother.com/welcome/dlf005249/MFC-9330CDW-inst-B1-ASA.EXE"
        $printerinf = "C:\Printer\Printer 30 - Lager\install\driver\gdi\32_64\BRPRC12A.INF"
        $printerdriver = "Brother MFC-9330CDW Printer"
        $printerip = "192.168.1.30"
        $printerlocation = "Printer ved Lones bord"
    
        $printerfolder = "C:\Printer\$printername"
        $file = Split-Path $printdriverlink -Leaf
        $portNumber = "91"+$printerip.Split(".")[-1]

        Stop-Service "Spooler"
        Remove-Item "C:\Windows\System32\spool\PRINTERS\*.*" -Force | Out-Null
        Start-Service "Spooler"

        # deaktiver automatisk installation af netværksprintere
        if (!(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private")) {
            New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Force | Out-Null}
            Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Name "AutoSetup" -Type DWord -Value 0

        # fjerner allerede installerede printere
        Get-Printer | ? Name -cMatch "OneNote (Desktop)|OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer 
        Get-Printer | ? Name -Match "9330|$printername" | Remove-Printer -ea SilentlyContinue
        Get-PrinterPort | ? Name -match $printerip | Remove-PrinterPort -ea SilentlyContinue
        Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue
        # installér 7-zip hvis den ikke allerede er installeret. bruges til stabil driver udpakning.
        if(!(Test-Path "$env:ProgramFiles\7-Zip\7z.exe")){
            $dlurl = 'https://7-zip.org/' + (Invoke-WebRequest -Uri 'https://7-zip.org/' | Select-Object -ExpandProperty Links | Where-Object {($_.innerHTML -eq 'Download') -and ($_.href -like "a/*") -and ($_.href -like "*-x64.exe")} | Select-Object -First 1 | Select-Object -ExpandProperty href)
            $installerPath = Join-Path $env:TEMP (Split-Path $dlurl -Leaf)
            Invoke-WebRequest $dlurl -OutFile $installerPath -UseBasicParsing
            Start-Process -FilePath $installerPath -Args "/S" -Verb RunAs -Wait
            Remove-Item $installerPath} 

        new-item -ItemType Directory -Path $printerfolder -Force | out-null

    #Downloader driver
    write-host "`t`t- Downloader driver.."
        Remove-item -Path $printerfolder\* -Force -recurse | out-null
        (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$file")

    #Udpakker driver
    write-host "`t`t- Udpakker driver.."
        & ${env:ProgramFiles}\7-Zip\7z.exe x "$printerfolder\$file" "-o$($printerfolder)" -y | out-null; ; sleep -s 5

    #Installer Printer
    write-host "`t`t- Konfigurer Printer:"; sleep -s 5
        write-host "`t`t`t`t- Driverbiblotek"
        pnputil.exe -i -a $printerinf | out-null ; sleep -s 5
        write-host "`t`t`t`t- Driver"
        Add-PrinterDriver -Name $printerdriver | out-null; sleep -s 5
        write-host "`t`t`t`t- Printerport"
        Add-PrinterPort -Name $printerip -PrinterHostAddress $printerip -PortNumber $portnumber | out-null; sleep -s 5
        write-host "`t`t`t`t- Printer"
        Add-Printer -Name $printername -PortName $printerip -DriverName $printerdriver -PrintProcessor winprint -Location $printerlocation -Comment "automatiseret af Andreas" | out-null; sleep -s 5
    #Oprydning
        Remove-item  -Path "$printerfolder\" -Exclude $file -Recurse -Force

    write-host "`t- Printeren er installeret!" -f Green
}else {write-host "[INGEN FORBINDELSE]" -f red; write-host "`tDer er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}






























