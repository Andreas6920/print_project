Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Relaunch script as an elevated process
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    # Relaunch as an elevated process
    $Script = $MyInvocation.MyCommand.Path
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-ExecutionPolicy RemoteSigned", "-File `"$Script`""
    
}

Function Get-LogDate {return (Get-Date -f "[yyyy/MM/dd HH:MM:ss]")}

Function Start-PrinterConfiguration {

 param (
     [Parameter(Mandatory=$true)]
     [string]$Name,
     [Parameter(Mandatory=$true)]
     [string]$IPv4,
     [Parameter(Mandatory=$false)]
     [string]$Driverlink,
     [Parameter(Mandatory=$true)]
     [string]$Drivername,
     [Parameter(Mandatory=$true)]
     [string]$Driverfilename,
     [Parameter(Mandatory=$true)]
     [string]$Location)

# Kontrollér forbindelse til Printer
    Write-Host "$(Get-LogDate)    - Tester Internet forbindelse..." -NoNewline -ForegroundColor Yellow
    if (Test-Connection  1.1.1.1 -Quiet) {
        Start-Sleep -s 3
        Write-Host "[FORBINDELSE ETABLERET]" -ForegroundColor Green

# Pre-install
    # Variabler
    Write-Host "$(Get-LogDate)`t- Systemet forberedes:"
    Start-Sleep -s 3
    $system = $env:SystemDrive
    $system32 = [Environment]::GetFolderPath("System")
    $printerfolder = Join-path -Path $system -ChildPath "Printer\$name"
    $printerdriverfile = Join-path -path $printerfolder -ChildPath "$Name.zip"
    $spoolfolder = Join-path -path $system32 -ChildPath "spool\PRINTERS"

# Installation
     
    # Clean spooler
    if (Get-ChildItem $spoolfolder){
        Write-Host "$(Get-LogDate)`t`t- Renser spooler"
        Stop-Service "Spooler" | out-null 
        Start-Sleep -s 3
        Get-ChildItem -Path $spoolfolder | Remove-Item -Recurse -Force
        Start-Service "Spooler" | out-null
        Start-Sleep -s 3}
        
    # Deaktiver internet explorer first run
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (!(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main").DisableFirstRunCustomize){
        Write-Host "$(Get-LogDate)`t`t- Deaktiver IE wizard"
        Start-Sleep -s 3
        New-Item -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Force | Out-Null
        Start-Sleep -s 3
        Set-ItemProperty -Path  "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize"  -Value 1}
        
    # Deaktiver automatisk installation af netværksprintere
    if (!(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private")) {
        Write-Host "$(Get-LogDate)`t`t- Deaktiver auto install"
        New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Name "AutoSetup" -Type DWord -Value 0}
    
    # Fjern gamle installation af aktuelle printer
    $printertags = $Name
    foreach ($tag in $printertags){
    if (Get-Printer | ? Name -match $tag){
        $printername = (Get-printer | ? name -match $tag).Name
        $printerport = (Get-printer | ? name -match $tag).PortName
        Write-Host "$(Get-LogDate)`t`t- Fjerner gamle installation"
        Remove-Printer -Name $printername
        Start-Sleep -S 2
        Remove-PrinterPort -Name $printerport}}
     
    # Fjern gamle windows printere
    $printertags = "Fax", "OneNote for Windows 10", "Microsoft XPS Document Writer", "Microsoft Print to PDF" 
    foreach ($tag in $printertags){
    if (Get-Printer | ? Name -cmatch $tag){
        $printername = (Get-printer | ? name -match $tag).Name
        Write-Host "$(Get-LogDate)`t`t- Fjerner printer: $printername"
        Remove-Printer -Name $printername}}

    # Fjern gamle jensen company printere
    $printertags = "9310", "4132", "M507", "7131", "9330", "2365", "Printer 10 - Kontor"
    foreach ($tag in $printertags){
    if (Get-Printer | ? Name -match $tag){
        $printername = (Get-printer | ? name -match $tag).Name
        $printerport = (Get-printer | ? name -match $tag).PortName
        Write-Host "$(Get-LogDate)`t`t- Fjerner printer: $printername"
        Remove-Printer -Name $printername
        Start-Sleep -S 2
        Remove-PrinterPort -Name $printerport}}
             
     # Mappe oprettes til driver
    if(!(test-path $printerfolder)){
        Write-Host "$(Get-LogDate)`t`t- Opretter printermappe"
        new-item -ItemType Directory $printerfolder | Out-Null}
    else{
        Remove-Item "$printerfolder\*" -Recurse -Exclude "$Name.zip" -Force | Out-Null
        if((test-path $printerdriverfile)){
        $backup = Get-Date (Get-ChildItem $printerdriverfile).CreationTime.ToShortDateString() -Format "yyyy.MM.dd"
        Rename-Item -Path $printerdriverfile -NewName "$backup.zip"}}
    Start-Sleep -S 2
     
     # Downloader driver
    Write-Host "$(Get-LogDate)`t`t- Downloader driver"
    (New-Object net.webclient).Downloadfile($Driverlink, $printerdriverfile)   
     
     # Udpakker driver
    Write-Host "$(Get-LogDate)`t`t- Udpakker driver"
    Expand-Archive -Path $printerdriverfile -DestinationPath $printerfolder
    $printerdriverinf = (get-childitem $printerfolder -include "*.inf" -Recurse | ? Name -eq $Driverfilename)[0].FullName
    Start-Sleep -S 2
         
    # Opsæt printer
    Write-Host "$(Get-LogDate)`t- Printeren Opsættes:"
    $ProgressPreference = "SilentlyContinue" # hide progressbar
    Start-Sleep -s 3
    Write-Host "$(Get-LogDate)`t`t- Tilføjer driver"
    pnputil.exe -i -a $printerdriverinf | out-null
    Start-Sleep -s 3
    Write-Host "$(Get-LogDate)`t`t- Installér driver:"$Drivername
    Add-PrinterDriver -Name $Drivername | out-null
    Start-Sleep -s 3
    Write-Host "$(Get-LogDate)`t`t- Opretter printerport:"$IPv4
    Add-PrinterPort -Name $IPv4 -PrinterHostAddress $IPv4 -ErrorAction Ignore | out-null
    Start-Sleep -s 3
    Write-Host "$(Get-LogDate)`t`t- Opsætter printer"
    Add-Printer -Name $Name -PortName $IPv4 -DriverName $Drivername -PrintProcessor winprint -Location $Location -Comment "automatiseret af Andreas" | out-null; sleep -s 5
    Start-sleep -S 3;
    Write-Host "$(Get-LogDate)`t`t- Indstiller én sides udskrift fremfor dobbelsiddet"
    Get-Printer -Name $Name | Set-PrintConfiguration -DuplexingMode OneSided
    Start-sleep -S 3;
    Write-Host "$(Get-LogDate)`t`t- Rengør disk"
    Get-childitem -path $printerfolder -Directory | Remove-Item -Recurse -Force | Out-Null
    Get-childitem -path $printerfolder | ? Name -notmatch "$Name|\d{1,4}\.\d{1,2}\.\d{1,2}.zip" | Remove-Item -Force | Out-Null
    Start-sleep -S 3;
    Write-Host "$(Get-LogDate)`t`t- Dobbelt-tjekker at printer processen kører"
    Start-Service  -Name "Spooler"
    Start-sleep -S 3;
    Write-Host "$(Get-LogDate)`t- $Name er nu installeret.`n" -f Green
    $ProgressPreference = "Continue"}

Else {
        Write-Host "[INGEN FORBINDELSE]" -f Red
        Write-Host "$(Get-LogDate)`tDer er ikke forbindelse til printeren"  -BackgroundColor Red -f White
        Write-Host "$(Get-LogDate)`tTest om printeren er i dvale eller om du/printeren har internet!" -BackgroundColor Red -f White
        Write-host ""
    }




<# End of Start-PrinterConfiguration function #>}


Function Install-Printer {
    param (
        [ValidateSet(11, 20, 30, 40, 50, 60, 70, 80)]
        [int]$PrinterNummer,   
        [ValidateSet("Kontor", "Lager", "Butik")] 
        [string]$Afdeling,
        [switch]$Alle,
        [switch]$NavisionPrinter)

if(($Afdeling -eq "Kontor") -or ($PrinterNummer -eq "11") -or ($Alle)){
    Start-PrinterConfiguration -Name "Printer 11 - Kontor" `
    -IPv4 "192.168.1.11" `
    -Driverlink "https://www.googleapis.com/drive/v3/files/1aAFlSwdaEXwYMnZm-7G-rDQcQZX45R4a?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" `
    -Location "Printer bag Lone B" `
    -Drivername "HP LaserJet M507 PCL 6 (V3)" `
    -Driverfilename "hpkoca2a_x64.inf"}

if(($Afdeling -eq "Kontor") -or ($PrinterNummer -eq "20") -or ($Alle)){
    Start-PrinterConfiguration -Name "Printer 20 - Kontor" `
    -IPv4 "192.168.1.20" `
    -Driverlink "https://www.googleapis.com/drive/v3/files/1mW3MC4ODo77bfyWa3sGotITFsaZICvwi?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" `
    -Location "Canon printer med scanner" `
    -Drivername "Canon Generic Plus PCL6" `
    -Driverfilename "Cnp60MA64.INF"}

if(($Afdeling -eq "Lager") -or ($PrinterNummer -eq "30") -or ($Alle)){
    Start-PrinterConfiguration -Name "Printer 30 - Lager" `
    -IPv4 "192.168.1.30" `
    -Driverlink "https://www.googleapis.com/drive/v3/files/1s2o8FHiJ6f4dNW7AyPkWRqJxJ_dFhu6U?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" `
    -Location "Lagerprinter med scanner" `
    -Drivername "Brother MFC-9330CDW Printer" `
    -Driverfilename "BRPRC12A.INF"}

if(($Afdeling -eq "Lager") -or ($PrinterNummer -eq "40") -or ($Alle)){
    Start-PrinterConfiguration -Name "Printer 40 - Lager" `
    -IPv4 "192.168.1.40" `
    -Driverlink "https://www.googleapis.com/drive/v3/files/1uzIMA03CMIvebVwyE7dljLBlrN-fJINl?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" `
    -Location "Lagerprinter ved booking" `
    -Drivername "Brother HL-L2360D series" `
    -Driverfilename "BROHL13A.INF"}

if(($Afdeling -eq "Kontor") -or ($PrinterNummer -eq "50") -or ($Alle)){
    Start-PrinterConfiguration -Name "Printer 50 - Kontor" `
    -IPv4 "192.168.1.50" `
    -Driverlink "https://www.googleapis.com/drive/v3/files/1aAFlSwdaEXwYMnZm-7G-rDQcQZX45R4a?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" `
    -Location "HP Printeren i midten af kontoret" `
    -Drivername "HP LaserJet M507 PCL 6 (V3)" `
    -Driverfilename "hpkoca2a_x64.inf"}

if(($Afdeling -eq "Butik") -or ($PrinterNummer -eq "60") -or ($Alle)){
    Start-PrinterConfiguration -Name "Printer 60 - Butik" `
    -IPv4 "192.168.1.60" `
    -Driverlink "https://www.googleapis.com/drive/v3/files/1mURq7zSc6e4o85_IRjXV5k9nuWT1fCk8?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" `
    -Location "Printeren ved kassen" `
    -Drivername "ES7131(PCL6)" `
    -Driverfilename "OKW3X04V.INF"}

if(($Afdeling -eq "Lager") -or ($PrinterNummer -eq "70") -or ($Alle)){
    Start-PrinterConfiguration -Name "Printer 70 - Lager" `
    -IPv4 "192.168.1.70" `
    -Driverlink "https://www.googleapis.com/drive/v3/files/1OoS5fvj9S_J8tyYY-U7wY4C_QcuXR3dD?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" `
    -Location "Printeren til følgesedler" `
    -Drivername "Lexmark MS820 Series" `
    -Driverfilename "LMU03o40.inf"}

if(($Afdeling -eq "Butik") -or ($PrinterNummer -eq "80") -or ($Alle)){
    Start-PrinterConfiguration -Name "Printer 80 - Butik" `
    -IPv4 "192.168.1.80" `
    -Driverlink "https://www.googleapis.com/drive/v3/files/15OTs4jA9-c6xgS1xjU3Vk4yEwkLHfJm9?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" `
    -Location "Farveprinteren ved kassen" `
    -Drivername "HP Color Laser MFP 178 179" `
    -Driverfilename "sht13c.INF"}

if($NavisionPrinter){

    # Variables 
    Write-Host "    Opsætter printer til Navision.."
    $link = "https://nphardwareconnector.blob.core.windows.net/production/Setup.exe"
    $path = Join-Path -Path $env:TMP -ChildPath (Split-Path $link -Leaf)
    $desktop = [Environment]::GetFolderPath("Desktop")
    $startup = [Environment]::GetFolderPath("Startup")
    $shortcut = Join-Path $desktop -Childpath "NP Hardware Connector.lnk"

    # Install
    if (!(Test-Path $shortcut)) {
    Write-host "$(Get-LogDate)`t- Downloader Programmet.."
    Start-Sleep -S 1
    (New-Object net.webclient).Downloadfile("$link", "$path")
    Write-host "$(Get-LogDate)`t- Installere Programmet.."
    Get-Process | Where-Object { $_.Name -match "NP Hardware Connector" } | Select-Object -First 1 | Stop-Process
    Start-Sleep -S 1
    Start $path}

    # Create startup task
    Write-host "$(Get-LogDate)`t- Sætter til at starte automatisk.."
    Start-Sleep -S 3
    while (!(Test-Path $shortcut)) { Start-Sleep -S 1 }
    Copy-Item $shortcut $startup  
    Write-Host "$(Get-LogDate)`t- Navision printer opsætning er nu installeret." -f Green
    Start-Sleep -S 3}
        
    else{
        do {
            Clear-Host
            Write-Host ""
            Write-Host ""
            Write-Host "PRINTERPROGRAM, version 2.5" -f yellow
            Write-Host ""
            Write-Host ""
            Write-Host "$(Get-LogDate)`tValgmuligheder:"
            Write-Host ""
            Write-Host "$(Get-LogDate)`t1    -    Installér alle printere"
            Write-Host "$(Get-LogDate)`t2    -    Installér specific printer"
            Write-Host "$(Get-LogDate)`t3    -    Installér kontor printere"
            Write-Host "$(Get-LogDate)`t4    -    Installér lager printere"
            Write-Host "$(Get-LogDate)`t5    -    Installér butik printere"
            Write-Host "$(Get-LogDate)`t6    -    Installér navision printer"
            Write-Host "$(Get-LogDate)`t0    -    EXIT"
            Write-Host ""
            Write-Host ""
            Write-Host "INDTAST NUMMERET HER (1-4), EFTERFUGLT AF ENTER: " -f yellow -nonewline; ; ;
            $option = Read-Host
            "";
            Switch ($option) { 
                0 {exit}

                1 { Install-Printer -Alle }
                2 { $number = Read-Host "Indtast Nummer", Install-Printer -PrinterNummer $number }
                3 { Install-Printer -Afdeling "Kontor" }
                4 { Install-Printer -Afdeling "Lager" }
                5 { Install-Printer -Afdeling "Butik" }
                6 { Install-Printer -NavisionPrinter }
                
                            }}
        while ($option -notin 1..2 )}
        
<# End of Install-Printer function #>}