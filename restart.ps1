
$JLR = @()
$JLR = Get-ADComputer -SearchBase "OU=OU-POL LandRover,OU=OU-Piter,DC=int,DC=rolfcorp,DC=ru" -SearchScope 2 -Filter '*'|Where-Object {$_.name -like "pollr-aiobf*" -or $_.name  -like "pollr-aioins*" -or $_.name  -like "pollr-aiorecep*" -or $_.name  -like "pollr-aiosale*" -or $_.name  -like "pollr-bf*" -or $_.name  -like "pollr-evolute*" -or $_.name  -like "pollr-sale*" -or $_.name  -like "polr-aiobf*" -or $_.name  -like "polr-aioins*" -or $_.name  -like "polr-bf*" -and $_.name -notlike "polr-reserv03"}  |Select-Object -ExpandProperty Name | Sort-Object
$JLR | Restart-Computer -Force

$POL = @()
$POL = Get-ADComputer -SearchBase "OU=OU-POL Renault,OU=OU-Piter,DC=int,DC=rolfcorp,DC=ru" -SearchScope 2 -Filter '*'|Where-Object {$_.name -like "pollr-aiobf*" -or $_.name -like "pollr-aioins*" -or $_.name -like "pollr-sale*" -or $_.name -like "polr-aiobf*" -or $_.name -like "polr-aiofleet*" -or $_.name -like "polr-aioconf*" -or $_.name -like "polr-aioins*" -or $_.name -like "polr-aiosale*" -or $_.name -like "polr-aioyyy*" -or $_.name -like "polr-bf*" -or $_.name -like "polr-reserv*" -or $_.name -like "polr-aioreserv*" -or $_.name -like "pollr-hr01" -and $_.name -notlike "polr-reserv03"} |Select-Object -ExpandProperty Name | Sort-Object
$POL | Restart-Computer -Force

$y = Get-Date
foreach ($PC in $JLR) {$x = Invoke-Command $PC {Get-CimInstance Win32_OperatingSystem | select LastBootUpTime}
if ($x.LastBootUpTime.date -lt $y.date ){Invoke-Command $PC {Restart-Computer -Force}}

}
foreach ($PC in $POL) {$x = Invoke-Command $PC {Get-CimInstance Win32_OperatingSystem | select LastBootUpTime}
if ($x.LastBootUpTime.date -lt $y.date ){Invoke-Command $PC {Restart-Computer -Force}}

}


Remove-Item "F:\_temp\PS\Restart\dontrebooted.txt"
Remove-Item "F:\_temp\PS\Restart\rebooted.txt"


<#foreach ($PC in $JLR) {$x = Invoke-Command $PC {Get-CimInstance Win32_OperatingSystem | select LastBootUpTime}
if ($x.LastBootUpTime.date -lt $y.date ){$x |Out-File -filepath "D:\_temp\PS\Restart\dontrebooted.txt" -append}
if ($x.LastBootUpTime.date -eq $y.date ){$x |Out-File -filepath "D:\_temp\PS\Restart\rebooted.txt" -append}
}
#>
foreach ($PC in $POL) {$x = Invoke-Command $PC {Get-CimInstance Win32_OperatingSystem | select LastBootUpTime}
if ($x.LastBootUpTime.date -lt $y.date ){$x |Out-File -filepath "D:\_temp\PS\Restart\dontrebooted.txt" -append}
if ($x.LastBootUpTime.date -eq $y.date ){$x |Out-File -filepath "D:\_temp\PS\Restart\rebooted.txt" -append}
}