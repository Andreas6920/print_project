Write-host "Installing module"
$modulepath = $env:PSmodulepath.split(";")[1]
$modules = @("https://paste.ee/r/Sgcwz")
Foreach ($module in $modules) {
# prepare folder
	$file = "PrinterScript.psm1"
    $filename = $file.Replace(".psm1","").Replace(".ps1","").Replace(".psd","")
        if ($file -notmatch '\.psm1$'){$file = $file+".psm1"}
    $filedestination = "$modulepath/$filename/$file"
    $filesubfolder = split-path $filedestination -Parent
    If (!(Test-Path $filesubfolder )) {New-Item -ItemType Directory -Path $filesubfolder -Force | Out-Null; Start-Sleep -S 1}
# Download module
    (New-Object net.webclient).Downloadfile($module, $filedestination)
# Install module
    if (Get-Module -ListAvailable -Name $filename){ Import-module -name $filename; Write-Host "." -NoNewline}
 
 else {Write-Host "!"}}
 
 
Start-job -Name "Printer 60" -ScriptBlock {Install-Printer -Name "Printer 60 - Butik" -IPv4 "8.8.4.4" -Driverlink "https://www.googleapis.com/drive/v3/files/1mURq7zSc6e4o85_IRjXV5k9nuWT1fCk8?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" -Location "Printeren ved kassen" -Drivername "ES7131(PCL6)" -Driverfilename "OKW3X04V.INF";}
Start-job -Name "Printer 80" -ScriptBlock {Install-Printer -Name "Printer 80 - Butik" -IPv4 "8.8.8.8" -Driverlink "https://www.googleapis.com/drive/v3/files/15OTs4jA9-c6xgS1xjU3Vk4yEwkLHfJm9?alt=media&key=AIzaSyDCXkesTBEsIlxySObsDb2j5-44AsTtqXk" -Location "Farveprinteren ved kassen" -Drivername "HP Color Laser MFP 178 179" -Driverfilename "sht13c.INF";}
