cls
#Forbereder system
write-host "Tester forbindelse til printer 40 (Printer ved Booking-PC).. " -NoNewline; Sleep -s 3
write-host "[Forbindelse verificeret]".toUpper() -f green
write-host "`t- Installere Printer 40:"; Sleep -s 5

write-host "`t`t- Forbereder system.."
    $printername = "Printer 40 - Lager"
    $printdriverlink = "https://download.brother.com/welcome/dlf100988/Y14A_C1-hostm-1110.EXE"
    $printerinf = "C:\Printer\Printer 40 - Kontor\32_64\BROHL13A.INF"
    $printerdriver = "Brother HL-L2360D series"
    $printerip = "192.168.1.40"
    $printerlocation = "I skuret under urtepotterne"
    
    $printerfolder = "C:\Printer\$printername"
    $file = Split-Path $printdriverlink -Leaf
    $portNumber = "91"+$printerip.Split(".")[-1]

    Get-Printer | ? Name -cMatch "OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer
    Get-Printer | ? Name -Match "2365|$printername" | Remove-Printer -ea SilentlyContinue
    Get-PrinterPort | ? Name -match "192.168.1.40" | Remove-PrinterPort -ea SilentlyContinue
    Get-PrinterDriver | ? Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue

    if(!(Test-Path "$env:ProgramFiles\7-Zip\7z.exe")){
        $dlurl = 'https://7-zip.org/' + (Invoke-WebRequest -Uri 'https://7-zip.org/' | Select-Object -ExpandProperty Links | Where-Object {($_.innerHTML -eq 'Download') -and ($_.href -like "a/*") -and ($_.href -like "*-x64.exe")} | Select-Object -First 1 | Select-Object -ExpandProperty href)
        $installerPath = Join-Path $env:TEMP (Split-Path $dlurl -Leaf)
        Invoke-WebRequest $dlurl -OutFile $installerPath -UseBasicParsing
        Start-Process -FilePath $installerPath -Args "/S" -Verb RunAs -Wait
        Remove-Item $installerPath} 

    new-item -ItemType Directory -Path $printerfolder -Force | out-null

#Downloader driver
write-host "`t`t- Downloader driver"
    Remove-item -Path $printerfolder\* -Force -recurse | out-null
    (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$file")

#Udpakker driver
write-host "`t`t- Udpakker driver"
    & ${env:ProgramFiles}\7-Zip\7z.exe x "$printerfolder\$file" "-o$($printerfolder)" -y | out-null; ; sleep -s 5

#Installer Printer
write-host "`t`t- Konfigurer Printer:"
    write-host "`t`t`t- Driverbiblotek"
    pnputil.exe -i -a $printerinf | out-null ; sleep -s 5
    write-host "`t`t`t- Driver"
    Add-PrinterDriver -Name $printerdriver | out-null; sleep -s 5
    write-host "`t`t`t- Printerport"
    Add-PrinterPort -Name $printerip -PrinterHostAddress $printerip -PortNumber $portnumber | out-null; sleep -s 5
    write-host "`t`t`t- Printer"
    Add-Printer -Name "Printer 40 - Lager" -PortName $printerip -DriverName $printerdriver -PrintProcessor winprint -Location $printerlocation -Comment "automatiseret af Andreas" | out-null; sleep -s 5
#Oprydning
    Remove-item  -Path "$printerfolder\" -Exclude $file -Recurse -Force

write-host "`t- Printeren er installeret!" -f Green