function Start-PrinterPreparation {
    
    # Variabler
        Write-Host "`t- Systemet forberedes:"
        Start-Sleep -s 3
        $system = $env:SystemDrive
        $system32 = [Environment]::GetFolderPath("System")
        $printerfolder = Join-path -Path $system -ChildPath "Printer\$name"
        $printerdriverfile = Join-path -path $printerfolder -ChildPath "$Name.zip"
        $spoolfolder = Join-path -path $system32 -ChildPath "spool\PRINTERS"

    # Deaktiver internet explorer first run
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Host "`t    - Deaktiver IE wizard"
        $RegistryPath = "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main"
        If(!(Get-Item $RegistryPath | ? Property -EQ "DisableFirstRunCustomize")){Set-ItemProperty -Path  $RegistryPath -Name "DisableFirstRunCustomize" -Value 1}
        Start-Sleep -s 3

    # Clean spooler
        if (Get-ChildItem $spoolfolder){
        Write-Host "`t    - Renser spooler"
        Stop-Service "Spooler" | out-null 
        Start-Sleep -s 3
        Get-ChildItem -Path $spoolfolder | Remove-Item -Recurse -Force
        Start-Service "Spooler" | out-null
        Start-Sleep -s 3}

    # Deaktiver automatisk installation af netværksprintere
        if (!(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private")) {
        Write-Host "`t    - Deaktiver auto install"
        New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Force | Out-Null
        Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Name "AutoSetup" -Type DWord -Value 0}

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



}

function Install-Naviprinter {
    
    # Variables 
        Write-Host "    Opsætter printer til Navision.."
        $link = "https://nphardwareconnector.blob.core.windows.net/production/Setup.exe"
        $path = Join-Path -Path $env:TMP -ChildPath (Split-Path $link -Leaf)
        $desktop = [Environment]::GetFolderPath("Desktop")
        $startup = [Environment]::GetFolderPath("Startup")
        $shortcut = Join-Path $desktop -Childpath "NP Hardware Connector.lnk"


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
   
   # Kontrollér forbindelse til Printer
       Write-Host "    Forbinder til $Name..." -NoNewline
       if (Test-Connection  $IPv4 -Quiet) {
           Start-Sleep -s 3
           Write-Host "[FORBINDELSE ETABLERET]"
   
   # Pre-install
       # Variabler
           Write-Host "`t- Systemet forberedes:"
           Start-Sleep -s 3
           $system = $env:SystemDrive
           $system32 = [Environment]::GetFolderPath("System")
           $printerfolder = Join-path -Path $system -ChildPath "Printer\$name"
           $printerdriverfile = Join-path -path $printerfolder -ChildPath "$Name.zip"
           $spoolfolder = Join-path -path $system32 -ChildPath "spool\PRINTERS"
   
   # Installation
           
       # Fjern gamle installation af aktuelle printer
           $printertags = $Name
           foreach ($tag in $printertags){
           if (Get-Printer | ? Name -match $tag){
               $printername = (Get-printer | ? name -match $tag).Name
               $printerport = (Get-printer | ? name -match $tag).PortName
               Write-Host "`t    - Fjerner gamle installation"
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
               Rename-Item -Path $printerdriverfile -NewName "$backup.zip"}}
           Start-Sleep -S 2
        
        # Downloader driver
           Write-Host "`t    - Downloader driver"
           (New-Object net.webclient).Downloadfile($Driverlink, $printerdriverfile)   
        
        # Udpakker driver
           Write-Host "`t    - Udpakker driver"
           Expand-Archive -Path $printerdriverfile -DestinationPath $printerfolder
           $printerdriverinf = (get-childitem $printerfolder -include "*.inf" -Recurse | ? Name -eq $Driverfilename)[0].FullName
           Start-Sleep -S 2
            
       # Opsæt printer
           Write-Host "`t- Printeren Opsættes:"
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
           Write-Host "`tTest om printeren er i dvale eller om du/printeren har internet!" -BackgroundColor Red -f White
           Write-host ""
       }
}

function Start-PrinterScript {
    param (
        [Parameter(Mandatory=$false)]
        [string]$Department,
        [Parameter(Mandatory=$false)]
        [string]$Number,
        [Parameter(Mandatory=$false)]
        [string]$All)


    $Master = (irm -useb "https://raw.githubusercontent.com/Andreas6920/print_project/main/res/master_beta.txt").Split([Environment]::NewLine)
    

    if($Department){
        Start-Job -Name "Preparation" -Scriptblock  {Start-PrinterPreparation}
        Wait-Job -Name "Preparation"
        $Master | select-string -pattern $Department | ForEach-Object { Start-Job -Scriptblock  {$_}  }}

    if($All){
        Start-Job -Name "Preparation" -Scriptblock  {Start-PrinterPreparation}
        Wait-Job -Name "Preparation"
        Start-PrinterPreparation; $Master | select-string -pattern $Department | ForEach-Object { Start-Job -Scriptblock  {$_}  }}


    }



#Start-PrinterScript -Department Butik