#Check difference
    $new = (Invoke-WebRequest -uri "https://raw.githubusercontent.com/Andreas6920/print_project/main/checker.ps1").content
    $old = get-content "$($env:ProgramData)\Admin\run.ps1"
    if(!($old -cmatch $old)){
    
    $new | -OutFile "$($env:ProgramData)\Admin\run.ps1" -UseBasicParsing; powershell -ep bypass "$($env:ProgramData)\Admin\run.ps1"
    
    }