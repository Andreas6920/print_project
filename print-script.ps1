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
    if (Test-Connection  $IPv4 -Quiet) {
        Write-Host "Forbinder til $Name...`t" -NoNewline
        Start-Sleep -s 3
        Write-Host "[FORBINDELSE VERIFICERET]"


    # Pre-install
    Write-Host "`t`t- Systemet forberedes:"
    Start-Sleep -s 3
        
        # Variabler
            $printerfolder = "$env:SystemDrive\Printer\$Name"
            $printerdriverfile = $printerfolder + "\$Name.zip"
        
        # Clean spooler
        if(Get-ChildItem "$env:SystemRoot\System32\spool\PRINTERS\" ){
            Write-Host "`t`t`t`t- Renser spooler"
            Stop-Service "Spooler" | out-null 
            Start-Sleep -s 3
            Remove-Item "$env:SystemRoot\System32\spool\PRINTERS\*.*" -Force | Out-Null
            Start-Service "Spooler" | out-null
            Start-Sleep -s 3}
        
        # Deaktiver internet explorer first run
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        if (!(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main").DisableFirstRunCustomize){
            Write-Host "`t`t`t`t- Deaktiver IE wizard"
            Start-Sleep -s 3
            New-Item -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Force | Out-Null
            Start-Sleep -s 3
            Set-ItemProperty -Path  "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize"  -Value 1}
        
        # Deaktiver automatisk installation af netværksprintere
        if (!(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private")) {
            Write-Host "`t`t`t`t- Deaktiver auto install"
            New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Force | Out-Null
            Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\NcdAutoSetup\Private" -Name "AutoSetup" -Type DWord -Value 0}
        
        # Fjern aktuelle printer
        $printertags = $Name
        foreach ($tag in $printertags){
        if (Get-Printer | ? Name -match $tag){
            $printername = (Get-printer | ? name -match $tag).Name
            $printerport = (Get-printer | ? name -match $tag).PortName
            Write-Host "`t`t`t`t- Fjerner printer: $printername"
            Remove-Printer -Name $printername
            Start-Sleep -S 2
            Remove-PrinterPort -Name $printerport}}
        
        # Fjern gamle windows printere
        $printertags = "Fax","OneNote for Windows 10","Microsoft XPS Document Writer", "Microsoft Print to PDF" 
        foreach ($tag in $printertags){
        if (Get-Printer | ? Name -cmatch $tag){
            $printername = (Get-printer | ? name -match $tag).Name
            Write-Host "`t`t`t`t- Fjerner printer: $printername"
            Remove-Printer -Name $printername}}

        # Fjern gamle jensen company printere
        $printertags = "9310","4132","M507", "7131","9330", "2365", "Printer 10 - Kontor"
        foreach ($tag in $printertags){
        if (Get-Printer | ? Name -match $tag){
            $printername = (Get-printer | ? name -match $tag).Name
            $printerport = (Get-printer | ? name -match $tag).PortName
            Write-Host "`t`t`t`t- Fjerner printer: $printername"
            Remove-Printer -Name $printername
            Start-Sleep -S 2
            Remove-PrinterPort -Name $printerport}}
                
        # Mappe oprettes til driver
            if(!(test-path $printerfolder)){
                Write-Host "`t`t`t`t- Opretter printermappe"
                new-item -ItemType Directory $printerfolder | Out-Null}
            else{
                Remove-Item "$printerfolder\*" -Recurse -Exclude "$Name.zip" -Force | Out-Null
                if((test-path $printerdriverfile)){
                $backup = Get-Date (Get-ChildItem $printerdriverfile).CreationTime.ToShortDateString() -Format "yyyy.MM.dd"
                Rename-Item -Path $printerdriverfile -NewName "$backup.zip"}
            }
            Start-Sleep -S 2
        
        # Downloader driver
            Write-Host "`t`t`t`t- Downloader driver"
            (New-Object net.webclient).Downloadfile($Driverlink, $printerdriverfile)   
        
        # Udpakker driver
            Write-Host "`t`t`t`t- Udpakker driver"
            Expand-Archive -Path $printerdriverfile -DestinationPath $printerfolder
            $printerdriverinf = (get-childitem $printerfolder -include "*.inf" -Recurse | ? Name -eq $Driverfilename)[0].FullName
            Start-Sleep -S 2
    
            
    # Install
    Write-Host "`t`t- Printeren installeres:"
            $ProgressPreference = "SilentlyContinue" # hide progressbar
            Start-Sleep -s 3
            Write-Host "`t`t`t`t- Tilføjer driver"
            pnputil.exe -i -a $printerdriverinf | out-null
            Start-Sleep -s 3
            Write-Host "`t`t`t`t- Installér driver:"$Drivername
            Add-PrinterDriver -Name $Drivername | out-null
            Start-Sleep -s 3
            Write-Host "`t`t`t`t- Opretter printerport:"$IPv4
            Add-PrinterPort -Name $IPv4 -PrinterHostAddress $IPv4 -ErrorAction Ignore | out-null
            Start-Sleep -s 3
            Write-Host "`t`t`t`t- Opsætter printer"
            Add-Printer -Name $Name -PortName $IPv4 -DriverName $Drivername -PrintProcessor winprint -Location $Location -Comment "automatiseret af Andreas" | out-null; sleep -s 5
            Start-sleep -S 3;
            Write-Host "`t`t`t`t- Indstiller en sides udskrift fremfor dobbelsiddet"
            Get-Printer -Name $Name | Set-PrintConfiguration -DuplexingMode OneSided
            Start-sleep -S 3;
            Write-Host "`t`t`t`t- Fjerner downloadede filer"
            Get-childitem -path $printerfolder -Directory | Remove-Item -Recurse -Force | Out-Null
            Get-childitem -path $printerfolder | ? Name -notmatch "$Name|\d{1,4}\.\d{1,2}\.\d{1,2}.zip" | Remove-Item -Force | Out-Null
            Start-sleep -S 3;
            Write-Host "`t`t`t`t- Dobbelt-tjekker at printer servicen kører"
            Start-Service  -Name "Spooler"
            Start-sleep -S 3;
            Write-Host "`t`t- $Name er nu installeret.`n" -f Green
            $ProgressPreference = "Continue" #unhide progressbar

    }
}
function Menu-Printer {

    #tjek efter admin rettigheder
    $admin_permissions_check = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $admin_permissions_check = $admin_permissions_check.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($admin_permissions_check) {

        do {
            cls
            "";"";Write-Host "VÆLG EN AF FØLGENDE MULIGHEDER VED AT INDTASTE NUMMERET:" -f yellow
            Write-Host ""; Write-Host "";
            Write-Host "Printer installation " -nonewline; Write-host "Version 2.0:" -f Gray;"";
            Write-Host "`t1`t-`tKontor afdeling`t(printer 11, 20, 50)"
            Write-Host "`t2 `t-`tLager afdeling`t(printer 30, 40)"
            Write-Host "`t3`t-`tButiks afdeling`t(printer 60)"
            #"";"";Write-Host "Andet:"
            #Write-Host "        [4] - Installation af helt ny PC"
            "";Write-Host "`t0`t-`tEXIT"
            "";"";
            Write-Host "INDTAST DIT NUMMER HER: " -f yellow -nonewline; ; ;
            $option = Read-Host
            "";
            Switch ($option) { 
                0 {exit}
                1 {
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
                    -Driverfilename "hpkoca2a_x64.inf";}
                
                2 {
                    Install-Printer -Name "Printer 60 - Butik" `
                    -IPv4 "192.168.1.60" `
                    -Driverlink "https://drive.google.com/uc?export=download&confirm=uc-download-link&id=1mURq7zSc6e4o85_IRjXV5k9nuWT1fCk8" `
                    -Location "Printeren ved kassen" `
                    -Drivername "ES7131(PCL6)" `
                    -Driverfilename "OKW3X04V.INF";}
                
                3 {
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
                    -Driverfilename "BROHL13A.INF";}
                             }}
        while ($option -notin 1..3 )}
        
    else {
           1..99 | % {$Warning_message = "POWERSHELL IS NOT RUNNING AS ADMINISTRATOR. Please close this and run this script as administrator."
           cls; ""; ""; ""; ""; ""; write-host $Warning_message -ForegroundColor White -BackgroundColor Red; ""; ""; ""; ""; ""; Start-Sleep 1; cls
           cls; ""; ""; ""; ""; ""; write-host $Warning_message -ForegroundColor White; ""; ""; ""; ""; ""; Start-Sleep 1; cls} }
}

Menu-Printer