Invoke-WebRequest -uri "https://raw.githubusercontent.com/Andreas6920/print_project/main/print-script.ps1" -OutFile "$env:ProgramData\print_script.ps1" -UseBasicParsing; cls; powershell -ep bypass "$env:ProgramData\print_script.ps1"


test
