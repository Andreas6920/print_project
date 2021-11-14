# Need-to-be-fixed list:
    # Empty.

function printer_kontor {
    
    # Inden afdelingens printere påbegynder installation:
        # Deaktiver automatisk installation af netværksprintere
        if (!(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private")) {
            New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Force | Out-Null}
            Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Name "AutoSetup" -Type DWord -Value 0

        # Installér 7-zip hvis den ikke allerede er installeret. bruges til stabil driver udpakning.
        if(!(Test-Path "$env:ProgramFiles\7-Zip\7z.exe")){
            $dlurl = 'https://7-zip.org/' + (Invoke-WebRequest -Uri 'https://7-zip.org/' | Select-Object -ExpandProperty Links | Where-Object {($_.innerHTML -eq 'Download') -and ($_.href -like "a/*") -and ($_.href -like "*-x64.exe")} | Select-Object -First 1 | Select-Object -ExpandProperty href)
            $installerPath = Join-Path $env:TEMP (Split-Path $dlurl -Leaf)
            Invoke-WebRequest $dlurl -OutFile $installerPath -UseBasicParsing
            Start-Process -FilePath $installerPath -Args "/S" -Verb RunAs -Wait
            Remove-Item $installerPath} 
    
    # Printer 10 - Kontor
    write-host "Forbinder til printer 10 (Printer bag Lones plads).." -NoNewline; Sleep -s 3
    if (Test-Connection  192.168.1.10 -Quiet) {
        write-host "[Forbindelse verificeret]".toUpper() -f green
        write-host "`t- Begynder installation af Printer 10:"; Sleep -s 5
        write-host "`t`t- Forbereder system.."
        # Variabler klargøres
            $printername = "Printer 10 - Kontor"
            $printerfolder = "$env:SystemDrive\Printer\$printername"
            $printerinf = "$env:SystemDrive\Printer\Printer 10 - Kontor\OKW3X055114\Driver\OKW3X055.INF"
            $printerdriver = "ES4132(PCL6)"
            $printerip = "192.168.1.10"
            $printerlocation = "Printer bag Lone B's bord"

        # System klargøres
            # Renser spooler
            Stop-Service "Spooler" | out-null; sleep -s 3
            Remove-Item "$env:SystemRoot\System32\spool\PRINTERS\*.*" -Force | Out-Null
            Start-Service "Spooler"

            # Fjerner allerede installerede printere
            Get-Printer | ? Name -cMatch "OneNote (Desktop)|OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer 
            Get-Printer | ? Name -Match "9310|$printername" | Remove-Printer -ea SilentlyContinue
            Get-PrinterPort | ? Name -match $printerip | Remove-PrinterPort -ea SilentlyContinue
            Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue

        # Mappe oprettes til driver
            new-item -ItemType Directory -Path $printerfolder -Force | out-null

        # Downloader driver
          write-host "`t`t- Downloader driver.."
            Remove-item -Path $printerfolder\* -Force -recurse | out-null
            $FileDestination = "C:\Printer\Printer 10 - Kontor\OKW3X055114_254753.exe"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri "https://drive.google.com/uc?export=download&id=1R5xYNtAbOtm_8qX0mDcn8YeY945eLjbn" -OutFile "_tmp.txt" -SessionVariable googleDriveSession
            $searchString = Select-String -Path "_tmp.txt" -Pattern "confirm="
            $searchString -match "confirm=(?<content>.*)&amp;id=" | Out-Null
            $confirmCode = $matches['content']
            Remove-Item "_tmp.txt"
            Invoke-WebRequest -Uri "https://drive.google.com/uc?export=download&confirm=${confirmCode}&id=1R5xYNtAbOtm_8qX0mDcn8YeY945eLjbn" -OutFile $FileDestination -WebSession $googleDriveSession
        
        # Udpakker driver
          write-host "`t`t- Udpakker driver.."
            $file = (Get-ChildItem $printerfolder | sort LastWriteTime | select -last 1).Name
            & ${env:ProgramFiles}\7-Zip\7z.exe x "$printerfolder\$file" "-o$($printerfolder)" -y | out-null; ; sleep -s 5

        # Installer Printer
          write-host "`t`t- Konfigurer Printer:"; sleep -s 5
            write-host "`t`t`t`t- Driverbiblotek"
            pnputil.exe -i -a $printerinf | out-null ; sleep -s 5
            write-host "`t`t`t`t- Driver"
            Add-PrinterDriver -Name $printerdriver | out-null; sleep -s 5
            write-host "`t`t`t`t- Printerport"
            Add-PrinterPort -Name $printerip -PrinterHostAddress $printerip | out-null; sleep -s 5
            write-host "`t`t`t`t- Printer"
            Add-Printer -Name $printername -PortName $printerip -DriverName $printerdriver -PrintProcessor winprint -Location $printerlocation -Comment "automatiseret af Andreas" | out-null; sleep -s 5
            Stop-Service "Spooler" | Out-Null; sleep -s 5
            Start-Service "Spooler" | Out-Null

        write-host "`t- Printeren er installeret!" -f Green
    }else {write-host "[INGEN FORBINDELSE]" -f red; write-host "`tDer er ikke forbindelse til printeren, test om den er i dvale eller om du/printeren har internet!" -f red}
    
    # Printer 20 - Kontor
    write-host "Forbinder til printer 20 (Scanner ved indgangen).." -NoNewline; Sleep -s 3
    if (Test-Connection  192.168.1.20 -Quiet) {
        write-host "[Forbindelse verificeret]".toUpper() -f green
        write-host "`t- Begynder installation af Printer 20:"; Sleep -s 5

        write-host "`t`t- Forbereder system.."
            $printername = "Printer 20 - Kontor"
            $printdriverlink = "https://gdlp01.c-wss.com/gds/1/0100009371/02/MF429MFDriverV580WPEN.exe"
            $printerinf = "$env:SystemDrive\Printer\Printer 20 - Kontor\intdrv\PCL6\x64\etc\Cnp60MA64.INF"
            
            $printerdriver = "Canon Generic Plus PCL6 V130"
            $printerip = "192.168.1.20"
            $printerlocation = "Den canon printer med scanner"
            
            $printerfolder = "$env:SystemDrive\Printer\$printername"
            $file = Split-Path $printdriverlink -Leaf

            # renser spooler
            Stop-Service "Spooler" | out-null; sleep -s 3
            Remove-Item "$env:SystemRoot\System32\spool\PRINTERS\*.*" -Force | Out-Null
            Start-Service "Spooler"

            # fjerner allerede installerede printere
            Get-Printer | ? Name -cMatch "OneNote (Desktop)|OneNote for Windows 20|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer 
            Get-Printer | ? Name -Match "MF420|$printername" | Remove-Printer -ea SilentlyContinue
            Get-PrinterPort | ? Name -match $printerip | Remove-PrinterPort -ea SilentlyContinue
            Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue

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
            Add-PrinterPort -Name $printerip -PrinterHostAddress $printerip | out-null; sleep -s 5
            write-host "`t`t`t`t- Printer"
            Add-Printer -Name $printername -PortName $printerip -DriverName $printerdriver -PrintProcessor winprint -Location $printerlocation -Comment "automatiseret af Andreas" | out-null; sleep -s 5
            Stop-Service "Spooler" | Out-Null; sleep -s 5
            Start-Service "Spooler" | Out-Null

        write-host "`t- Printeren er installeret!" -f Green
    }else {write-host "[INGEN FORBINDELSE]" -f red; write-host "`tDer er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}

    # Printer 50 - Kontor
    write-host "Forbinder til printer 50 (HP printeren).." -NoNewline; Sleep -s 3
    if (Test-Connection  192.168.1.50 -Quiet) {
        write-host "[Forbindelse verificeret]".toUpper() -f green
        write-host "`t- Begynder installation af Printer 50:"; Sleep -s 5

        write-host "`t`t- Forbereder system.."
        # Variabler     
            $printername = "Printer 50 - Kontor"
            $printdriverlink = "https://ftp.ext.hp.com/pub/softlib/software13/printers/LJE/M507/LJM507_Full_WebPack_49.1.4431.exe"
            $printerinf = "$env:SystemDrive\Printer\Printer 50 - Kontor\hpkoca2a_x64.inf"
            
            $printerdriver = "HP LaserJet M507 PCL 6 (V3)"
            $printerip = "192.168.1.50"
            $printerlocation = "HP Printeren i midten af kontoret"
            
            $printerfolder = "$env:SystemDrive\Printer\$printername"
            $file = Split-Path $printdriverlink -Leaf

        # Klargøring
            # renser spooler
            Stop-Service "Spooler" | out-null; sleep -s 3
            Remove-Item "$env:SystemRoot\System32\spool\PRINTERS\*.*" -Force | Out-Null
            Start-Service "Spooler"

            # deaktiver automatisk installation af netværksprintere
            if (!(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private")) {
                New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Force | Out-Null}
                Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Name "AutoSetup" -Type DWord -Value 0

            # fjerner allerede installerede printere
            Get-Printer | ? Name -cMatch "OneNote (Desktop)|OneNote for Windows 50|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer 
            Get-Printer | ? Name -Match "M507|$printername" | Remove-Printer -ea SilentlyContinue
            Get-PrinterPort | ? Name -match $printerip | Remove-PrinterPort -ea SilentlyContinue
            Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue

        # Opretter mappe til driver
        new-item -ItemType Directory -Path $printerfolder -Force | out-null

        # Downloader driver
        write-host "`t`t- Downloader driver.."
            # Downloader driver
          write-host "`t`t- Downloader driver.."
            Remove-item -Path $printerfolder\* -Force -recurse | out-null
            $FileDestination = "C:\Printer\Printer 50 - Kontor\LJM507_Full_WebPack_49.1.4431.exe"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri "https://drive.google.com/uc?export=download&id=1ReVrnos3kCbMtkhLv988hydfoJUiZAmf" -OutFile "_tmp.txt" -SessionVariable googleDriveSession
            $searchString = Select-String -Path "_tmp.txt" -Pattern "confirm="
            $searchString -match "confirm=(?<content>.*)&amp;id=" | Out-Null
            $confirmCode = $matches['content']
            Remove-Item "_tmp.txt"
            Invoke-WebRequest -Uri "https://drive.google.com/uc?export=download&confirm=${confirmCode}&id=1ReVrnos3kCbMtkhLv988hydfoJUiZAmf" -OutFile $FileDestination -WebSession $googleDriveSession
        # Udpakker driver
        write-host "`t`t- Udpakker driver.."
            & ${env:ProgramFiles}\7-Zip\7z.exe x "$printerfolder\$file" "-o$($printerfolder)" -y | out-null; ; sleep -s 5

        # Installer Printer
        write-host "`t`t- Konfigurer Printer:"; sleep -s 5
            write-host "`t`t`t`t- Driverbiblotek"
            pnputil.exe -i -a $printerinf | out-null ; sleep -s 5
            write-host "`t`t`t`t- Driver"
            Add-PrinterDriver -Name $printerdriver | out-null; sleep -s 5
            write-host "`t`t`t`t- Printerport"
            Add-PrinterPort -Name $printerip -PrinterHostAddress $printerip | out-null; sleep -s 5
            write-host "`t`t`t`t- Printer"
            Add-Printer -Name $printername -PortName $printerip -DriverName $printerdriver -PrintProcessor winprint -Location $printerlocation -Comment "automatiseret af Andreas" | out-null; sleep -s 5
            Stop-Service "Spooler" | Out-Null; sleep -s 5
            Start-Service "Spooler" | Out-Null

        write-host "`t- Printeren er installeret! `n" -f Green
    }else {write-host "[INGEN FORBINDELSE]" -f red; write-host "`tDer er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}

    # Når afdelingen printere er installeret:
        
        # List alle printer og sæt dem til en-sidet print
        Get-Printer * | Set-PrintConfiguration -DuplexingMode OneSided

        # Slet udpakkede filer, for besparelse af diskplads. driver bibeholdes.
        $newfolders = Get-ChildItem -Directory -path C:\Printer | where name -match kontor
        Foreach ($folder in $newfolders){cd c:\printer\$folder; remove-item * -Exclude "*.exe" -Recurse -Force }


}

function printer_butik {


    write-host "Forbinder til printer 60 (Printer ved kassen).." -NoNewline; Sleep -s 3
    if (Test-Connection  192.168.1.60 -Quiet) {
        write-host "[Forbindelse verificeret]".toUpper() -f green
        write-host "`t- Begynder installation af Printer 60:"; Sleep -s 5

        write-host "`t`t- Forbereder system.."
            $printername = "Printer 60 - Butik"
            $printdriverlink = "https://www.oki.com/be/printing/en/download/OKW3X04V101_22029.exe"
            $printerinf = "$env:SystemDrive\Printer\Printer 60 - Butik\OKW3X04V101\driver\OKW3X04V.INF"
            $printerdriver = "ES7131(PCL6)"
            $printerip = "192.168.1.60"
            $printerlocation = "Printeren ved kassen"
            
            $printerfolder = "$env:SystemDrive\Printer\$printername"
            $file = Split-Path $printdriverlink -Leaf

            # renser spooler
            Stop-Service "Spooler" | out-null; sleep -s 3
            Remove-Item "$env:SystemRoot\System32\spool\PRINTERS\*.*" -Force | Out-Null
            Start-Service "Spooler"

            # deaktiver automatisk installation af netværksprintere
            if (!(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private")) {
                New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Force | Out-Null}
                Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Name "AutoSetup" -Type DWord -Value 0

            # fjerner allerede installerede printere
            Get-Printer | ? Name -cMatch "OneNote (Desktop)|OneNote for Windows 60|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer 
            Get-Printer | ? Name -Match "ES7131|$printername" | Remove-Printer -ea SilentlyContinue
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
            Add-PrinterPort -Name $printerip -PrinterHostAddress $printerip | out-null; sleep -s 5
            write-host "`t`t`t`t- Printer"
            Add-Printer -Name $printername -PortName $printerip -DriverName $printerdriver -PrintProcessor winprint -Location $printerlocation -Comment "automatiseret af Andreas" | out-null; sleep -s 5
        #Oprydning
            Remove-item  -Path "$printerfolder\" -Exclude $file -Recurse -Force
            Stop-Service "Spooler" | Out-Null; sleep -s 5
            Start-Service "Spooler" | Out-Null
            # undgå dobbelsidet udskrift
            Get-Printer | ? Name -match $printername | Set-PrintConfiguration -DuplexingMode OneSided;
        write-host "`t- Printeren er installeret!" -f Green
    }else {write-host "[INGEN FORBINDELSE]" -f red; write-host "`tDer er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}


}
    
function printer_lager {
    
    write-host "Forbinder til printer 30 (Printer ved Lones bord).. " -NoNewline; Sleep -s 3
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
            Add-PrinterPort -Name $printerip -PrinterHostAddress $printerip | out-null; sleep -s 5
            write-host "`t`t`t`t- Printer"
            Add-Printer -Name $printername -PortName $printerip -DriverName $printerdriver -PrintProcessor winprint -Location $printerlocation -Comment "automatiseret af Andreas" | out-null; sleep -s 5
        #Oprydning
            Remove-item  -Path "$printerfolder\" -Exclude $file -Recurse -Force
            Stop-Service "Spooler" | Out-Null; sleep -s 5
            Start-Service "Spooler" | Out-Null
            # undgå dobbelsidet udskrift
            Get-Printer | ? Name -match $printername | Set-PrintConfiguration -DuplexingMode OneSided;
        write-host "`t- Printeren er installeret!" -f Green
    }else {write-host "[INGEN FORBINDELSE]" -f red; write-host "`tDer er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}
    
    write-host "Forbinder til printer 40 (Printer ved Booking-PC).. " -NoNewline; Sleep -s 3
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
        
                    Stop-Service "Spooler" | Out-Null
                    Remove-Item "C:\Windows\System32\spool\PRINTERS\*.*" -Force | Out-Null
                    Start-Service "Spooler" | Out-Null
        
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
                    Add-PrinterPort -Name $printerip -PrinterHostAddress $printerip | Out-null; sleep -s 5
                    write-host "`t`t`t`t- Printer"
                    Add-Printer -Name $printername -PortName $printerip -DriverName $printerdriver -PrintProcessor winprint -Location $printerlocation -Comment "automatiseret af Andreas" | out-null; sleep -s 5
                #Oprydning
                    Remove-item  -Path "$printerfolder\" -Exclude $file -Recurse -Force
                    Stop-Service "Spooler" | Out-Null; sleep -s 5
                    Start-Service "Spooler" | Out-Null
                    # undgå dobbelsidet udskrift
                    Get-Printer | ? Name -match $printername | Set-PrintConfiguration -DuplexingMode OneSided;
        
        
        
                write-host "`t- Printeren er installeret!" -f Green
    }else {write-host "[INGEN FORBINDELSE]" -f red; write-host "`tDer er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}


}
        
       


    #front-end begynd
    #tjek efter admin rettigheder
    $admin_permissions_check = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $admin_permissions_check = $admin_permissions_check.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($admin_permissions_check) {


    do {
        cls
        "";"";Write-host "VÆLG EN AF FØLGENDE MULIGHEDER VED AT INDTASTE NUMMERET:" -f yellow
        Write-host ""; Write-host "";
        Write-host "Printer installation:"
        Write-host "`t1  - Kontor afdeling`t(printer 10, 20, 50)"
        Write-host "`t2  - Lager afdeling`t`t(printer 30, 40)"
        Write-host "`t3  - Butiks afdeling`t(printer 60)"
        #"";"";Write-host "Andet:"
        #Write-host "        [4] - Installation af helt ny PC"
        "";Write-host "`t0 - EXIT"
        Write-host ""; Write-host "";
        Write-Host "INDTAST DIT NUMMER HER: " -f yellow -nonewline; ; ;
        $option = Read-Host
        Switch ($option) { 
            0 {exit}
            1 {printer_kontor;}
            2 {printer_lager;}
            3 {printer_butik;}
        }}
    
    while ($option -notin 1..3 )
                            }

else {
    1..99 | % {
        $Warning_message = "POWERSHELL IS NOT RUNNING AS ADMINISTRATOR. Please close this and run this script as administrator."
        cls; ""; ""; ""; ""; ""; write-host $Warning_message -ForegroundColor White -BackgroundColor Red; ""; ""; ""; ""; ""; Start-Sleep 1; cls
        cls; ""; ""; ""; ""; ""; write-host $Warning_message -ForegroundColor White; ""; ""; ""; ""; ""; Start-Sleep 1; cls
    }    
}