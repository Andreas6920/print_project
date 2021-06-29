    
function printer_lager {


    write-host "Tester forbindelse til printer 30 (Printer ved Bentes plads).."
    if (Test-Connection  192.168.1.30 -Quiet) 
    {
    write-host "        - Forbindelse verificeret."
        #Step 1 - forbereder system

        $printername = "Printer 30 - Lager"
        $printerdriver = "Brother MFC-9330CDW Printer"
        $printdriverlink = "https://download.brother.com/welcome/dlf005249/MFC-9330CDW-inst-B1-ASA.EXE"
        $printerip = "192.168.1.30"
        $printerfolder = "C:\Printer\$printername"
        
        if (!(Get-Module -ListAvailable -Name 7Zip4PowerShell)){Install-Module -Name 7Zip4PowerShell -Force | out-null}
        Get-Printer | Where-Object Name -cMatch "OneNote for Windows 10|Microsoft XPS Document Writer|Microsoft Print to PDF|Fax" | Remove-Printer
        Get-Printer | Where-Object Name -Match "9330|$printername" | Remove-Printer -ea SilentlyContinue
        Get-PrinterPort | Where-Object PrinterHostAddress -match $printerip | Remove-PrinterPort -ea SilentlyContinue
        Get-PrinterDriver | Where-Object Name -match $printerdriver | Remove-PrinterDriver -ea SilentlyContinue
    
    write-host "        - Installere Printer 30..."
        #Step 2 - Opret mappe
    write-host "                - Opretter undermappe..."
            new-item -ItemType Directory -Path $printerfolder -Force | out-null
            Remove-item $printerfolder\* -Recurse; sleep -s 3     
            
        #Step 3 - download driver
    write-host "                - Downloader driver..."
        (New-Object Net.WebClient).DownloadFile($printdriverlink, "$printerfolder\$printerdriver.zip")
    write-host "                - Udpakker driver..."
        #Step 4 - udpak driver
        Expand-7Zip -ArchiveFileName $printerfolder\$printerdriver.zip -TargetPath $printerfolder
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
            start-sleep -s 5
    write-host "                - $printername er nu installeret!" -f green;"";"";"";
    }
    else {write-host "Der er ikke forbindelse til Printer 30, test om den er slukket eller om du/printeren har internet!" -f red}
    }




    #front-end begynd
    #tjek efter admin rettigheder
    $admin_permissions_check = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $admin_permissions_check = $admin_permissions_check.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($admin_permissions_check) {

    $intro ="
           __________                                 
         .'----------`.                              
         | .--------. |                             
         | |########| |       __________              
         | |########| |      /__________\             
.--------| `--------'  |------|   --=--  |------------.
|         `----,-.-----'      |o ======  |             | 
|       ______|_|_______     |__________|             | 
|      /  %%%%%%%%%%%%  \                             | 
|     /  %%%%%%%%%%%%%%  \                            | 
|     ^^^^^^^^^^^^^^^^^^^^                            | 
+-----------------------------------------------------+
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 

"


    do {
         Write-host $intro -f Yellow 
        "";"";Write-host "VÆLG EN AF FØLGENDE MULIGHEDER VED AT INDTASTE NUMMERET:" -f yellow
        Write-host ""; Write-host "";
        Write-host "Printer installation:"
        Write-host "        [1] - Kontor printere"
        Write-host "        [2] - Lager printere"
        Write-host "        [3] - Butiks printere"
        #"";"";Write-host "Andet:"
        #Write-host "        [4] - Installation af helt ny PC"
        "";"";Write-host "        [0] - EXIT"
        Write-host ""; Write-host "";
        Write-Host "INDTAST DIT NUMMER HER: " -f yellow -nonewline; ; ;
        $option = Read-Host
        Switch ($option) { 
            0 {exit}
            1 {printer_lager;}
            2 {}
            3 {}
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
