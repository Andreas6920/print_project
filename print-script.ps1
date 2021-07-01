function printer_kontor {

    write-host "Tester forbindelse til printer 10 (Printer bag Lone B's plads).."
    if (Test-Connection  192.168.1.10 -Quiet) 
    {
        
        write-host "        - Forbindelse verificeret."
        #Step 1 - forbereder system

        write-host "        - Installere Printer 10..."
        $printername = "Printer 10 - Kontor"
        $printerdriver = "ES4132(PCL6)"
        $printdriverlink = "https://www.oki.com/be/printing/en/download/OKW3X055114_254753.exe"
        $file = $printdriverlink.Split("/")[-1].Split(".*")[0]
        $printerip = "192.168.1.10"
        $printerfolder = "C:\Printer\$printername"
        
        if (!(Get-Module -ListAvailable -Name 7Zip4PowerShell)){Install-Module -Name 7Zip4PowerShell -Force | out-null}
        Get-Printer | ? Name -cMatch "OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer
        Get-Printer | ? Name -Match "426|$printername" | Remove-Printer -ea SilentlyContinue
        Get-PrinterPort | Where-Object PrinterHostAddress -match $printerip | Remove-PrinterPort -ea SilentlyContinue
        Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue
    
        #Step 2 - Opret mappe
    write-host "                - Opretter undermappe..."
            new-item -ItemType Directory -Path $printerfolder -Force | out-null
            Remove-item $printerfolder\* -Recurse -Force; sleep -s 3     
            
        #Step 3 - download driver
    write-host "                - Downloader driver..."
        (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$file.zip")
    write-host "                - Udpakker driver..."
        #Step 4 - udpak driver
        Expand-7Zip -ArchiveFileName $printerfolder\$file.zip -TargetPath $printerfolder
    write-host "                - Installere driver... "
        #Step 5 - Lokaliser inf fil, installer driver
        start-process "printui.exe" -ArgumentList '/ia /m "ES4132(PCL6)" /h "x64" /v "Type 3 - User Mode" /f "C:\Printer\Printer 10 - Kontor\OKW3X055114\Driver\OKW3X055.INF"'; sleep -s 5
    write-host "                - Opretter printerport..."
        #Step 6 - opret port til printer
        $Port = ([wmiclass]"win32_tcpipprinterport").createinstance()

        $Port.Name = $printerip
        $Port.HostAddress = $printerip
        $Port.Protocol = "1"
        $Port.PortNumber = "91"+$printerip.Split(".")[-1]
        $Port.SNMPEnabled = $false
        $Port.Description = "Created by Andreas Mouritsen"
        $Port.Put() | Out-Null
        
    write-host "                - opretter printer..."
        #Step 7 - Opret printer med driver og port
        $Printer = ([wmiclass]"win32_Printer").createinstance()
        $Printer.Name = $printername
        $Printer.DriverName = $printerdriver
        $Printer.DeviceID = $printername
        $Printer.Shared = $false
        $Printer.PortName = $printerip
        $Printer.Comment = "Automatiseret af Andreas Mouritsen"
        $printer.Location = "Printer bag Lone B's bord"
        $Printer.Put() | Out-Null
        #undgå dobbelsidet udskrift
        Get-Printer | ? Name -match $printername | Set-PrintConfiguration -DuplexingMode OneSided
        #pladsoprydning
        Remove-Item $printerfolder\* -Exclude "$file.zip" -recurse
        start-sleep -s 5
    write-host "                - $printername er nu installeret!" -f green;
    }
    else {write-host "Der er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}
    
    write-host "Tester forbindelse til printer 20 (Scanner ved indgangen).."
    if (Test-Connection  192.168.1.20 -Quiet) 
    {
        write-host "        - Forbindelse verificeret."
        #Step 1 - forbereder system

        write-host "        - Installere Printer 20..."
        $printername = "Printer 20 - Kontor"
        $printerdriver = "Canon Generic Plus PCL6"
        $printdriverlink = "https://gdlp01.c-wss.com/gds/1/0100009371/02/MF429MFDriverV580WPEN.exe"
        $file = $printdriverlink.Split("/")[-1].Split(".*")[0]
        $printerip = "192.168.1.20"
        $printerfolder = "C:\Printer\$printername"
        
        if (!(Get-Module -ListAvailable -Name 7Zip4PowerShell)){Install-Module -Name 7Zip4PowerShell -Force | out-null}
        Get-Printer | ? Name -cMatch "OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer
        Get-Printer | ? Name -Match "426|$printername" | Remove-Printer -ea SilentlyContinue
        Get-PrinterPort | Where-Object PrinterHostAddress -match $printerip | Remove-PrinterPort -ea SilentlyContinue
        Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue
        
    
        #Step 2 - Opret mappe
    write-host "                - Opretter undermappe..."
            new-item -ItemType Directory -Path $printerfolder -Force | out-null
            Remove-item $printerfolder\* -Recurse -Force; sleep -s 3     
            
        #Step 3 - download driver
    write-host "                - Downloader driver..."
        (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$file.zip")
    write-host "                - Udpakker driver..."
        #Step 4 - udpak driver
        Expand-7Zip -ArchiveFileName $printerfolder\$file.zip -TargetPath $printerfolder
    write-host "                - Installere driver... "
        #Step 5 - Lokaliser inf fil, installer driver
        start-process "printui.exe" -ArgumentList '/ia /m "Canon Generic Plus PCL6" /h "x64" /v "Type 3 - User Mode" /f "C:\Printer\Printer 20 - Kontor\intdrv\PCL6\x64\Driver\Cnp60MA64.INF"'; sleep -s 5
    write-host "                - Opretter printerport..."
        #Step 6 - opret port til printer
        $Port = ([wmiclass]"win32_tcpipprinterport").createinstance()

        $Port.Name = $printerip
        $Port.HostAddress = $printerip
        $Port.Protocol = "1"
        $Port.PortNumber = "91"+$printerip.Split(".")[-1]
        $Port.SNMPEnabled = $false
        $Port.Description = "Created by Andreas Mouritsen"
        $Port.Put() | Out-Null
        
    write-host "                - opretter printer..."
        #Step 7 - Opret printer med driver og port
        $Printer = ([wmiclass]"win32_Printer").createinstance()
        $Printer.Name = $printername
        $Printer.DriverName = $printerdriver
        $Printer.DeviceID = $printername
        $Printer.Shared = $false
        $Printer.PortName = $printerip
        $Printer.Comment = "Automatiseret af Andreas Mouritsen"
        $printer.Location = "Den canon printer med scanner"
        $Printer.Put() | Out-Null
        #undgå dobbelsidet udskrift
        Get-Printer | ? Name -match $printername | Set-PrintConfiguration -DuplexingMode OneSided
        #pladsoprydning
        Remove-Item $printerfolder\* -Exclude "$file.zip" -recurse
        start-sleep -s 5
    write-host "                - $printername er nu installeret!" -f green;
    }
    else {write-host "Der er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}
    
    write-host "Tester forbindelse til printer 50 (HP printeren).."
    if (Test-Connection  192.168.1.50 -Quiet) 
    {
     
        write-host "        - Forbindelse verificeret."
        #Step 1 - forbereder system

        write-host "        - Installere Printer 50..."
        $printername = "Printer 50 - Kontor"
        $printerdriver = "HP LaserJet M507 PCL 6 (V3)"
        $printdriverlink = "https://ftp.ext.hp.com/pub/softlib/software13/printers/LJE/M507/LJM507_Full_WebPack_49.1.4431.exe"
        $file = $printdriverlink.Split("/")[-1].Split(".*")[0]
        $printerip = "192.168.1.50"
        $printerfolder = "C:\Printer\$printername"
        
        if (!(Get-Module -ListAvailable -Name 7Zip4PowerShell)){Install-Module -Name 7Zip4PowerShell -Force | out-null}
        Get-Printer | ? Name -cMatch "OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer
        Get-Printer | ? Name -Match "507|$printername" | Remove-Printer -ea SilentlyContinue
        Get-PrinterPort | Where-Object PrinterHostAddress -match $printerip | Remove-PrinterPort -ea SilentlyContinue
        Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue
        
    
        #Step 2 - Opret mappe
    write-host "                - Opretter undermappe..."
            new-item -ItemType Directory -Path $printerfolder -Force | out-null
            Remove-item $printerfolder\* -Recurse -Force; sleep -s 3     
            
        #Step 3 - download driver
    write-host "                - Downloader driver..."
        (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$file.exe"); sleep -s 5
        #step 4 - Udpakker driver
    write-host "                - Udpakker driver..."
        ## Downloader + installér 7-zip
        $dlurl = 'https://7-zip.org/' + (Invoke-WebRequest -Uri 'https://7-zip.org/' | Select-Object -ExpandProperty Links | Where-Object {($_.innerHTML -eq 'Download') -and ($_.href -like "a/*") -and ($_.href -like "*-x64.exe")} | Select-Object -First 1 | Select-Object -ExpandProperty href)
        $installerPath = Join-Path $env:TEMP (Split-Path $dlurl -Leaf)
        Invoke-WebRequest $dlurl -OutFile $installerPath #-UseBasicParsing ??
        Start-Process -FilePath $installerPath -Args "/S" -Verb RunAs -Wait
        Remove-Item $installerPath
        ## udpakker med 7-zip
        & ${env:ProgramFiles}\7-Zip\7z.exe x "$printerfolder\$file.exe" "-o$($printerfolder)" -y | out-null
        #Step 5 - Lokaliser inf fil, installer driver
    write-host "                - Installere driver... "
        start-process "printui.exe" -ArgumentList '/ia /m "HP LaserJet M507 PCL 6 (V3)" /h "x64" /v "Type 3 - User Mode" /f "C:\Printer\Printer 50 - Kontor\hpkoca2a_x64.inf"'; sleep -s 5
    write-host "                - Opretter printerport..."
        #Step 6 - opret port til printer
        $Port = ([wmiclass]"win32_tcpipprinterport").createinstance()

        $Port.Name = $printerip
        $Port.HostAddress = $printerip
        $Port.Protocol = "1"
        $Port.PortNumber = "91"+$printerip.Split(".")[-1]
        $Port.SNMPEnabled = $false
        $Port.Description = "Created by Andreas Mouritsen"
        $Port.Put() | Out-Null
        
    write-host "                - opretter printer..."
        #Step 7 - Opret printer med driver og port
        $Printer = ([wmiclass]"win32_Printer").createinstance()
        $Printer.Name = $printername
        $Printer.DriverName = $printerdriver
        $Printer.DeviceID = $printername
        $Printer.Shared = $false
        $Printer.PortName = $printerip
        $Printer.Comment = "Automatiseret af Andreas Mouritsen"
        $printer.Location = "HP Printeren i midten af kontoret"
        $Printer.Put() | Out-Null
        #undgå dobbelsidet udskrift
        Get-Printer | ? Name -match $printername | Set-PrintConfiguration -DuplexingMode OneSided
        #pladsoprydning
        Remove-Item $printerfolder\* -Exclude "$file.exe" -recurse -Force
        start-sleep -s 5
    write-host "                - $printername er nu installeret!" -f green;
      
    }
    else {write-host "Der er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}


}



function printer_lager {

    
    write-host "Tester forbindelse til printer 30 (Printer ved Bentes plads).."
    if (Test-Connection  192.168.1.30 -Quiet) 
    {
    write-host "        - Forbindelse verificeret."
        #Step 1 - forbereder system

        write-host "        - Installere Printer 30..."
        $printername = "Printer 30 - Lager"
        $printerdriver = "Brother MFC-9330CDW Printer"
        $printdriverlink = "https://download.brother.com/welcome/dlf005249/MFC-9330CDW-inst-B1-ASA.EXE"
        $file = $printdriverlink.Split("/")[-1].Split(".*")[0]
        $printerip = "192.168.1.30"
        $printerfolder = "C:\Printer\$printername"
        
        if (!(Get-Module -ListAvailable -Name 7Zip4PowerShell)){Install-Module -Name 7Zip4PowerShell -Force | out-null}
        Get-Printer | ? Name -cMatch "OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer
        Get-Printer | ? Name -Match "9330|$printername" | Remove-Printer -ea SilentlyContinue
        Get-PrinterPort | Where-Object PrinterHostAddress -match $printerip | Remove-PrinterPort -ea SilentlyContinue
        Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue
    
        #Step 2 - Opret mappe
    write-host "                - Opretter undermappe..."
            new-item -ItemType Directory -Path $printerfolder -Force | out-null
            Remove-item $printerfolder\* -Recurse -Force; sleep -s 3     
            
        #Step 3 - download driver
    write-host "                - Downloader driver..."
        (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$file.zip"); sleep -s 5
    write-host "                - Udpakker driver..."
        #Step 4 - udpak driver
        Expand-7Zip -ArchiveFileName $printerfolder\$file.zip -TargetPath $printerfolder
    write-host "                - Installere driver... "
        #Step 5 - Lokaliser inf fil, installer driver
        start-process "printui.exe" -ArgumentList '/ia /m "Brother MFC-9330CDW Printer" /h "x64" /v "Type 3 - User Mode" /f "C:\Printer\Printer 30 - Lager\install\driver\gdi\32_64\BRPRC12A.INF"'; sleep -s 5
    write-host "                - Opretter printerport..."
        #Step 6 - opret port til printer
        $Port = ([wmiclass]"win32_tcpipprinterport").createinstance()

        $Port.Name = $printerip
        $Port.HostAddress = $printerip
        $Port.Protocol = "1"
        $Port.PortNumber = "91"+$printerip.Split(".")[-1]
        $Port.SNMPEnabled = $false
        $Port.Description = "Created by Andreas Mouritsen"
        $Port.Put() | Out-Null
        
    write-host "                - opretter printer..."
        #Step 7 - Opret printer med driver og port
        $Printer = ([wmiclass]"win32_Printer").createinstance()
        $Printer.Name = $printername
        $Printer.DriverName = $printerdriver
        $Printer.DeviceID = $printername
        $Printer.Shared = $false
        $Printer.PortName = $printerip
        $Printer.Comment = "Automatiseret af Andreas Mouritsen"
        $printer.Location = "Bentes' printer"
        $Printer.Put() | Out-Null
        #undgå dobbelsidet udskrift
        Get-Printer | ? Name -match $printername | Set-PrintConfiguration -DuplexingMode OneSided
        #pladsoprydning
        Remove-Item $printerfolder\* -Exclude "$file.zip" -recurse
        start-sleep -s 5
    write-host "                - $printername er nu installeret!" -f green;
    }
    else {write-host "Der er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}
    
    write-host "Tester forbindelse til printer 40 (Printer ved booking).."
    if (Test-Connection  192.168.1.40 -Quiet) 
    {
        write-host "        - Forbindelse verificeret."
        #Step 1 - forbereder system

        write-host "        - Installere Printer 40..."
        $printername = "Printer 40 - Kontor"
        $printerdriver = "Brother HL-L2360D series"
        $printdriverlink = "https://download.brother.com/welcome/dlf100988/Y14A_C1-hostm-1110.EXE"
        $file = $printdriverlink.Split("/")[-1].Split(".*")[0]
        $printerip = "192.168.1.40"
        $printerfolder = "C:\Printer\$printername"
        
        if (!(Get-Module -ListAvailable -Name 7Zip4PowerShell)){Install-Module -Name 7Zip4PowerShell -Force | out-null}
        Get-Printer | ? Name -cMatch "OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer
        Get-Printer | ? Name -Match "2365|$printername" | Remove-Printer -ea SilentlyContinue
        Get-PrinterPort | Where-Object PrinterHostAddress -match $printerip | Remove-PrinterPort -ea SilentlyContinue
        Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue
    
        #Step 2 - Opret mappe
    write-host "                - Opretter undermappe..."
            new-item -ItemType Directory -Path $printerfolder -Force | out-null
            Remove-item $printerfolder\* -Recurse -Force; sleep -s 3     
            
        #Step 3 - download driver
    write-host "                - Downloader driver..."
        (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$file.zip")
    write-host "                - Udpakker driver..."
        #Step 4 - udpak driver
        Expand-7Zip -ArchiveFileName $printerfolder\$file.zip -TargetPath $printerfolder
    write-host "                - Installere driver... "
        #Step 5 - Lokaliser inf fil, installer driver
        start-process "printui.exe" -ArgumentList '/ia /m "Brother HL-L2360D series" /h "x64" /v "Type 3 - User Mode" /f "C:\Printer\Printer 40 - Kontor\32_64\BROHL13A.INF"'; sleep -s 5
    write-host "                - Opretter printerport..."
        #Step 6 - opret port til printer
        $Port = ([wmiclass]"win32_tcpipprinterport").createinstance()

        $Port.Name = $printerip
        $Port.HostAddress = $printerip
        $Port.Protocol = "1"
        $Port.PortNumber = "91"+$printerip.Split(".")[-1]
        $Port.SNMPEnabled = $false
        $Port.Description = "Created by Andreas Mouritsen"
        $Port.Put() | Out-Null
        
    write-host "                - opretter printer..."
        #Step 7 - Opret printer med driver og port
        $Printer = ([wmiclass]"win32_Printer").createinstance()
        $Printer.Name = $printername
        $Printer.DriverName = $printerdriver
        $Printer.DeviceID = $printername
        $Printer.Shared = $false
        $Printer.PortName = $printerip
        $Printer.Comment = "Automatiseret af Andreas Mouritsen"
        $printer.Location = "Printer ved booking"
        $Printer.Put() | Out-Null
        #undgå dobbelsidet udskrift
        Get-Printer | ? Name -match $printername | Set-PrintConfiguration -DuplexingMode OneSided
        #pladsoprydning
        Remove-Item $printerfolder\* -Exclude "$file.zip" -recurse
        start-sleep -s 5
    write-host "                - $printername er nu installeret!" -f green;    
    }
    else {write-host "Der er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}    


}


function printer_butik {


        write-host "Tester forbindelse til printer 60 (Printer ved kassen).."
        if (Test-Connection  192.168.1.60 -Quiet) 
        {
            write-host "        - Forbindelse verificeret."
            #Step 1 - forbereder system
    
            write-host "        - Installere Printer 60..."
            $printername = "Printer 60 - Butik"
            $printerdriver = "ES7131(PCL6)"
            $printdriverlink = "https://www.oki.com/be/printing/en/download/OKW3X04V101_22029.exe"
            $file = $printdriverlink.Split("/")[-1].Split(".*")[0]
            $printerip = "192.168.1.60"
            $printerfolder = "C:\Printer\$printername"
            
            if (!(Get-Module -ListAvailable -Name 7Zip4PowerShell)){Install-Module -Name 7Zip4PowerShell -Force | out-null}
            Get-Printer | ? Name -cMatch "OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer
            Get-Printer | ? Name -Match "7131|$printername" | Remove-Printer -ea SilentlyContinue
            Get-PrinterPort | Where-Object PrinterHostAddress -match $printerip | Remove-PrinterPort -ea SilentlyContinue
            Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue
        
            #Step 2 - Opret mappe
        write-host "                - Opretter undermappe..."
                new-item -ItemType Directory -Path $printerfolder -Force | out-null
                Remove-item $printerfolder\* -Recurse -Force; sleep -s 3     
                
            #Step 3 - download driver
        write-host "                - Downloader driver..."
            (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$file.zip"); sleep -s 5
        write-host "                - Udpakker driver..."
            #Step 4 - udpak driver
            Expand-7Zip -ArchiveFileName $printerfolder\$file.zip -TargetPath $printerfolder; sleep -s 5
        write-host "                - Installere driver... "
            #Step 5 - Lokaliser inf fil, installer driver
            start-process "printui.exe" -ArgumentList '/ia /m "ES7131(PCL6)" /h "x64" /v "Type 3 - User Mode" /f "C:\Printer\Printer 60 - Butik\OKW3X04V101\driver\OKW3X04V.INF"'; sleep -s 5
        write-host "                - Opretter printerport..."
            #Step 6 - opret port til printer
            $Port = ([wmiclass]"win32_tcpipprinterport").createinstance()
    
            $Port.Name = $printerip
            $Port.HostAddress = $printerip
            $Port.Protocol = "1"
            $Port.PortNumber = "91"+$printerip.Split(".")[-1]
            $Port.SNMPEnabled = $false
            $Port.Description = "Created by Andreas Mouritsen"
            $Port.Put() | Out-Null
            
        write-host "                - opretter printer..."
            #Step 7 - Opret printer med driver og port
            $Printer = ([wmiclass]"win32_Printer").createinstance()
            $Printer.Name = $printername
            $Printer.DriverName = $printerdriver
            $Printer.DeviceID = $printername
            $Printer.Shared = $false
            $Printer.PortName = $printerip
            $Printer.Comment = "Automatiseret af Andreas Mouritsen"
            $printer.Location = "printeren ved kassen"
            $Printer.Put() | Out-Null
            #undgå dobbelsidet udskrift
            Get-Printer | ? Name -match $printername | Set-PrintConfiguration -DuplexingMode OneSided
            #pladsoprydning
            Remove-Item $printerfolder\* -Exclude "$file.zip" -recurse
            start-sleep -s 5
            write-host "                - $printername er nu installeret!" -f green;
        }
        else {write-host "Der er ikke forbindelse til printeren, test om den er slukket eller om du/printeren har internet!" -f red}
        }
    





    #front-end begynd
    #tjek efter admin rettigheder
    $admin_permissions_check = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $admin_permissions_check = $admin_permissions_check.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($admin_permissions_check) {


    do {
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
            Default { cls; ""; ""; Write-host "UGYLDIGT VALG.." -f red; ""; ""; Start-Sleep 1; cls; ""; "" } 
        }
         
    }while ($option -ne 5 )
                            }

else {
    1..99 | % {
        $Warning_message = "POWERSHELL IS NOT RUNNING AS ADMINISTRATOR. Please close this and run this script as administrator."
        cls; ""; ""; ""; ""; ""; write-host $Warning_message -ForegroundColor White -BackgroundColor Red; ""; ""; ""; ""; ""; Start-Sleep 1; cls
        cls; ""; ""; ""; ""; ""; write-host $Warning_message -ForegroundColor White; ""; ""; ""; ""; ""; Start-Sleep 1; cls
    }    
}





#$path = "C:\Printer\Printer 40 - Kontor\32_64\BROHL13A.INF"

#pnputil.exe -i -a "C:\Printer\Printer 40 - Kontor\32_64\BROHL13A.INF"
#Add-PrinterDriver -Name "Brother HL-L2360D series"
#Add-PrinterPort -Name "test123_123" -PrinterHostAddress "192.168.1.40"
#Add-Printer -Name "Printer 40 - Lager" -PortName "test123_123" -DriverName "Brother HL-L2360D series" -PrintProcessor winprint