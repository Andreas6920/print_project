#Check difference
$a = "$($env:ProgramData)\Admin\run.ps1"
    if(!(test-path $a)){new-item $a}
$b = "https://raw.githubusercontent.com/Andreas6920/print_project/main/res/run.ps1"
$old = get-content $a
$new = (Invoke-WebRequest -uri $b).content

if($old -ne $new){

    Invoke-WebRequest -uri $b -OutFile $a -UseBasicParsing;
    PowerShell -ep Bypass $a

}
