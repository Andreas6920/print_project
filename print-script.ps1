function Install-Naviprinter {
    
    # Variables 
    Write-Host "Opsætter printer til Navision.."
    $link = "https://nphardwareconnector.blob.core.windows.net/production/Setup.exe"
    $path = $($env:TMP)+"\"+(Split-Path $link -Leaf)
    $desktop = [Environment]::GetFolderPath("Desktop")
    $startup = [Environment]::GetFolderPath("Startup")
    $shortcut = Join-Path $desktop "NP Hardware Connector.lnk"
     # [Environment]::GetFolderPath("LocalApplicationData")

    # Install
    if (!(Test-Path $shortcut)) {
    Write-host "`t- Downloader Programmet.."
    Start-Sleep -S 1
    (New-Object net.webclient).Downloadfile("$link", "$path")
    Write-host "`t- Installere Programmet.."
    Get-Process | Where-Object { $_.Name -match "NP Hardware Connector" } | Select-Object -First 1 | Stop-Process
    Start-Sleep -S 1
    Start $path}

    # Create startup task
    Write-host "`t- Sætter til at starte automatisk.."
    Start-Sleep -S 3
    while (!(Test-Path $shortcut)) { Start-Sleep -S 1 }
    Copy-Item $shortcut $startup  
    Write-Host "`t- Navision printer opsætning er nu installeret." -f Green
    Start-Sleep -S 3       

}

function Invoke-PrinterMsgbox {

    # Hvor mange printer er installeret
    [int]$option
    if ([int]$option -eq 1){$afdeling = "Kontor"; $printer_all=3}
    if ([int]$option -eq 2){$afdeling = "Lager"; $printer_all=3}
    if ([int]$option -eq 3){$afdeling = "Butik"; $printer_all=1}
 
    # i forhold til hvor mange printere der er i afdelingen
    $printer_complete = [int](get-printer | Where-Object Name -match $afdeling | Measure-Object).Count
    if ($printer_complete -eq $printer_all) {$msg = "Alle printere er installeret"}
    else {$msg = "$printer_complete ud af $printer_all printer(e) er installeret."}

    msg * $msg
}  

function Install-Printer {

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


 # Tjek forbindelse til Printer
 Write-Host "Forbinder til $Name..." -NoNewline
 if (Test-Connection  $IPv4 -Quiet) {
     Start-Sleep -s 3
     Write-Host "[FORBINDELSE ETABLERET]"


 # Pre-install
 Write-Host "`t- Systemet forberedes:"
 Start-Sleep -s 3
     
     # Variabler
         $printerfolder = "$env:SystemDrive\Printer\$Name"
         $printerdriverfile = $printerfolder + "\$Name.zip"
     
     # Clean spooler
     if(Get-ChildItem "$env:SystemRoot\System32\spool\PRINTERS\" ){
         Write-Host "`t    - Renser spooler"
         Stop-Service "Spooler" | out-null 
         Start-Sleep -s 3
         Remove-Item "$env:SystemRoot\System32\spool\PRINTERS\*.*" -Force | Out-Null
         Start-Service "Spooler" | out-null
         Start-Sleep -s 3}
     
     # Deaktiver internet explorer first run
     [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
     if (!(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main").DisableFirstRunCustomize){
         Write-Host "`t    - Deaktiver IE wizard"
         Start-Sleep -s 3
         New-Item -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Force | Out-Null
         Start-Sleep -s 3
         Set-ItemProperty -Path  "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize"  -Value 1}
     
     # Deaktiver automatisk installation af netværksprintere
     if (!(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private")) {
         Write-Host "`t    - Deaktiver auto install"
         New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Force | Out-Null
         Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Name "AutoSetup" -Type DWord -Value 0}
     
     # Fjern aktuelle printer
     $printertags = $Name
     foreach ($tag in $printertags){
     if (Get-Printer | ? Name -match $tag){
         $printername = (Get-printer | ? name -match $tag).Name
         $printerport = (Get-printer | ? name -match $tag).PortName
         Write-Host "`t    - Fjerner gamle installation"
         Remove-Printer -Name $printername
         Start-Sleep -S 2
         Remove-PrinterPort -Name $printerport}}
     
     # Fjern gamle windows printere
     $printertags = "Fax", "OneNote for Windows 10", "Microsoft XPS Document Writer", "Microsoft Print to PDF" 
     foreach ($tag in $printertags){
     if (Get-Printer | ? Name -cmatch $tag){
         $printername = (Get-printer | ? name -match $tag).Name
         Write-Host "`t    - Fjerner printer: $printername"
         Remove-Printer -Name $printername}}

     # Fjern gamle jensen company printere
     $printertags = "9310", "4132", "M507", "7131", "9330", "2365", "Printer 10 - Kontor"
     foreach ($tag in $printertags){
     if (Get-Printer | ? Name -match $tag){
         $printername = (Get-printer | ? name -match $tag).Name
         $printerport = (Get-printer | ? name -match $tag).PortName
         Write-Host "`t    - Fjerner printer: $printername"
         Remove-Printer -Name $printername
         Start-Sleep -S 2
         Remove-PrinterPort -Name $printerport}}
             
     # Mappe oprettes til driver
         if(!(test-path $printerfolder)){
             Write-Host "`t    - Opretter printermappe"
             new-item -ItemType Directory $printerfolder | Out-Null}
         else{
             Remove-Item "$printerfolder\*" -Recurse -Exclude "$Name.zip" -Force | Out-Null
             if((test-path $printerdriverfile)){
             $backup = Get-Date (Get-ChildItem $printerdriverfile).CreationTime.ToShortDateString() -Format "yyyy.MM.dd"
             Rename-Item -Path $printerdriverfile -NewName "$backup.zip"}
         }
         Start-Sleep -S 2
     
     # Downloader driver
         Write-Host "`t    - Downloader driver"
         (New-Object net.webclient).Downloadfile($Driverlink, $printerdriverfile)   
     
     # Udpakker driver
         Write-Host "`t    - Udpakker driver"
         Expand-Archive -Path $printerdriverfile -DestinationPath $printerfolder
         $printerdriverinf = (get-childitem $printerfolder -include "*.inf" -Recurse | ? Name -eq $Driverfilename)[0].FullName
         Start-Sleep -S 2
 
         
 # Install
 Write-Host "`t- Printeren installeres:"
         $ProgressPreference = "SilentlyContinue" # hide progressbar
         Start-Sleep -s 3
         Write-Host "`t    - Tilføjer driver"
         pnputil.exe -i -a $printerdriverinf | out-null
         Start-Sleep -s 3
         Write-Host "`t    - Installér driver:"$Drivername
         Add-PrinterDriver -Name $Drivername | out-null
         Start-Sleep -s 3
         Write-Host "`t    - Opretter printerport:"$IPv4
         Add-PrinterPort -Name $IPv4 -PrinterHostAddress $IPv4 -ErrorAction Ignore | out-null
         Start-Sleep -s 3
         Write-Host "`t    - Opsætter printer"
         Add-Printer -Name $Name -PortName $IPv4 -DriverName $Drivername -PrintProcessor winprint -Location $Location -Comment "automatiseret af Andreas" | out-null; sleep -s 5
         Start-sleep -S 3;
         Write-Host "`t    - Indstiller én sides udskrift fremfor dobbelsiddet"
         Get-Printer -Name $Name | Set-PrintConfiguration -DuplexingMode OneSided
         Start-sleep -S 3;
         Write-Host "`t    - Rengør disk"
         Get-childitem -path $printerfolder -Directory | Remove-Item -Recurse -Force | Out-Null
         Get-childitem -path $printerfolder | ? Name -notmatch "$Name|\d{1,4}\.\d{1,2}\.\d{1,2}.zip" | Remove-Item -Force | Out-Null
         Start-sleep -S 3;
         Write-Host "`t    - Dobbelt-tjekker at printer processen kører"
         Start-Service  -Name "Spooler"
         Start-sleep -S 3;
         Write-Host "`t- $Name er nu installeret.`n" -f Green
         $ProgressPreference = "Continue"}
 Else {
     
    Write-Host "[INGEN FORBINDELSE]" -f Red
    Write-Host "`tDer er ikke forbindelse til printeren"  -BackgroundColor Red -f White
    Write-Host "`tTest om printeren er i dvale eller om du/printeren har internet!" -BackgroundColor Red -f White}
}



 #tjek efter admin rettigheder
 $admin_permissions_check = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
 $admin_permissions_check = $admin_permissions_check.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
 if ($admin_permissions_check) {

     do {
         Clear-Host
         "";"";Write-Host "VÆLG EN AF FØLGENDE MULIGHEDER VED AT INDTASTE NUMMERET:" -f yellow
         Write-Host ""; Write-Host "";
         Write-Host "Printer installation " -nonewline; Write-host "Version 2.2:" -f Gray;"";
         Write-Host "`t1`t-`tKontor afdeling`t(printer 11, 20, 50)"
         Write-Host "`t2 `t-`tLager afdeling`t(printer 30, 40, 70)"
         Write-Host "`t3`t-`tButiks afdeling`t(printer 60)"
         Write-Host "`t4`t-`tNavision printer integration"
         "";Write-Host "`t0`t-`tEXIT"
         "";"";
         Write-Host "INDTAST DIT NUMMER HER: " -f yellow -nonewline; ; ;
         $option = Read-Host
         "";
         Switch ($option) { 
             0 {exit}
             1 { # Kontor Afdeling
                 Install-Printer -Name "Printer 11 - Kontor" `
                    -IPv4 "192.168.1.11" `
                    -Driverlink "https://drive.google.com/uc?export=download&confirm=uc-download-link&id=1aAFlSwdaEXwYMnZm-7G-rDQcQZX45R4a" `
                    -Location "Printer bag Lone B" `
                    -Drivername "HP LaserJet M507 PCL 6 (V3)" `
                    -Driverfilename "hpkoca2a_x64.inf";

                Install-Printer -Name "Printer 20 - Kontor" `
                    -IPv4 "192.168.1.20" `
                    -Driverlink "https://drive.google.com/uc?export=download&confirm=uc-download-link&id=1mW3MC4ODo77bfyWa3sGotITFsaZICvwi" `
                    -Location "Canon printer med scanner" `
                    -Drivername "Canon Generic Plus PCL6" `
                    -Driverfilename "Cnp60MA64.INF";

                Install-Printer -Name "Printer 50 - Kontor" `
                    -IPv4 "192.168.1.50" `
                    -Driverlink "https://drive.google.com/uc?export=download&confirm=uc-download-link&id=1aAFlSwdaEXwYMnZm-7G-rDQcQZX45R4a" `
                    -Location "HP Printeren i midten af kontoret" `
                    -Drivername "HP LaserJet M507 PCL 6 (V3)" `
                    -Driverfilename "hpkoca2a_x64.inf";
                Install-Naviprinter; 
                Invoke-PrinterMsgbox;
                exit;}
             
             2 { # Lager afdeling
                 Install-Printer -Name "Printer 30 - Lager" `
                    -IPv4 "192.168.1.30" `
                    -Driverlink "https://drive.google.com/uc?export=download&confirm=uc-download-link&id=1s2o8FHiJ6f4dNW7AyPkWRqJxJ_dFhu6U" `
                    -Location "Lagerprinter med scanner" `
                    -Drivername "Brother MFC-9330CDW Printer" `
                    -Driverfilename "BRPRC12A.INF";
             
                 Install-Printer -Name "Printer 40 - Lager" `
                    -IPv4 "192.168.1.40" `
                    -Driverlink "https://drive.google.com/uc?export=download&confirm=uc-download-link&id=1uzIMA03CMIvebVwyE7dljLBlrN-fJINl" `
                    -Location "Lagerprinter ved booking" `
                    -Drivername "Brother HL-L2360D series" `
                    -Driverfilename "BROHL13A.INF";

                 Install-Printer -Name "Printer 70 - Lager" `
                    -IPv4 "192.168.1.70" `
                    -Driverlink "https://drive.google.com/uc?export=download&confirm=uc-download-link&id=1OoS5fvj9S_J8tyYY-U7wY4C_QcuXR3dD" `
                    -Location "Printeren til følgesedler" `
                    -Drivername "Lexmark MS820 Series" `
                    -Driverfilename "LMU03o40.inf";                    
                Install-Naviprinter;
                Invoke-PrinterMsgbox;
                exit;}
             
             3 { # Butiks afdeling
                Install-Printer -Name "Printer 60 - Butik" `
                    -IPv4 "192.168.1.60" `
                    -Driverlink "https://drive.google.com/uc?export=download&confirm=uc-download-link&id=1mURq7zSc6e4o85_IRjXV5k9nuWT1fCk8" `
                    -Location "Printeren ved kassen" `
                    -Drivername "ES7131(PCL6)" `
                    -Driverfilename "OKW3X04V.INF";                    
                Install-Naviprinter;
                Invoke-PrinterMsgbox;
                 exit;}

             4 { # Installer Navision printer integration
                Install-Naviprinter;
                msg * "Navision printer integration installeret"
                exit;}
                          }}
     while ($option -notin 1..4 )}
     
 else {
        1..99 | % {$Warning_message = "POWERSHELL IS NOT RUNNING AS ADMINISTRATOR. Please close this and run this script as administrator."
        cls; ""; ""; ""; ""; ""; write-host $Warning_message -ForegroundColor White -BackgroundColor Red; ""; ""; ""; ""; ""; Start-Sleep 1; cls
        cls; ""; ""; ""; ""; ""; write-host $Warning_message -ForegroundColor White; ""; ""; ""; ""; ""; Start-Sleep 1; cls} }

