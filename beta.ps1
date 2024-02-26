Write-host "Installing module"
    $modulepath = $env:PSmodulepath.split(";")[1]
    $module = "https://raw.githubusercontent.com/Andreas6920/print_project/main/res/PrinterScript.psm1"
	$file = "PrinterScript.psm1"
    $filename = $file.Replace(".psm1","").Replace(".ps1","").Replace(".psd","")
    #if ($file -notmatch '\.psm1$'){$file = $file+".psm1"}
    $filedestination = "$modulepath/$filename/$file"
    $filesubfolder = split-path $filedestination -Parent
    If (!(Test-Path $filesubfolder )) {New-Item -ItemType Directory -Path $filesubfolder -Force | Out-Null; Start-Sleep -S 1}
    (New-Object net.webclient).Downloadfile($module, $filedestination)
    if (Get-Module -ListAvailable -Name $filename){ Import-module -name $filename}

    Start-PrinterScript -Department Lager
    
