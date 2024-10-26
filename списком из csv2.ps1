# Отредактируй $name - Брэнд для почтового домена и $OU - оушка в AD для учетки

# список в формате  Котышов Николай Андреевич <NAKotyshov@rolf.ru>; закинь в файл \\polr-it01\_temp\PS\Создание п.я\csv2.csv

# Из \\polr-it01\_temp\PS\Создание п.я\output.txt возьми ответ для инжитов
############################################################################
$name="evolute" # voyah evolute jaecoo omoda
$OU = “OU=AnotherUsers,OU=OU-POL LandRover,OU=OU-Piter,DC=int,DC=rolfcorp,DC=ru" # "OU=AnotherUsers,OU=OU-POL LandRover,OU=OU-Piter,DC=int,DC=rolfcorp,DC=ru"  "OU=AnotherUsers,OU=OU-POL Renault,OU=OU-Piter,DC=int,DC=rolfcorp,DC=ru"
#############################################################################

if ($name -eq "omoda")

{$domain = "omodarolf"
 $mail="@omoda-rolf.ru"} 

if ($name -eq "jaecoo")

{$domain = "jaecoorolf"
 $mail="@jaecoo-rolf.ru"} 

if ($name -eq "voyah")
 
 {$domain = "rolfvoyah"
 $mail="@rolf-voyah.ru"} 

if ($name -eq "evolute") 

{$domain = "evoluterolf"
$mail="@evolute-rolf.ru"} 

$CSV=Import-Csv -Path "\\polr-it01\_temp\PS\Создание п.я\csv2.csv" -Delimiter ";" -Encoding default

$CSV | foreach-object {

if ($_.in.split(" ").count -eq 1){
if ($_.in.split("@").count -eq 2){$login = $_.in.split("@")[0];$fio = (get-aduser $login).name}


else{$login=$_.in;$fio = (get-aduser $login).name}
}
 else
 {
if ((($_.in.Split(" ") -like '*@*').split("@")[0])[0] -eq "<")

{$login = (($_.in.Split(" ") -like '*@*').split("@")[0]).Remove(0,1)
$fio = $_.in.split(" ")[0]+" "+$_.in.split(" ")[1]+" "+$_.in.Split(" ")[2]
}

else
{$login = ($_.in.Split(" ") -like '*@*').split("@")[0]
$fio = $_.in.split(" ")[0]+" "+$_.in.split(" ")[1]+" "+$_.in.Split(" ")[2]
}
}
#for($login; $login -match "\d$";$login=$login -replace ".$"){}

$users=Get-ADUser -Filter "(samaccountname -like '$login*')" |select SamAccountName
if (($users.SamAccountName).count -gt 1)
 {$newlogin= $login+($users.SamAccountName).count}
else
{$newlogin= $login+"1"}




$newname= $fio + " ($name)"
$DN="@int.rolfcorp.ru"
$UPN=$newlogin+$DN

$GN=(Get-ADUser -Identity $login|select GivenName).givenname
$SN=(Get-ADUser -Identity $login|select SurName).surname
$DispN=((Get-ADUser -Identity $login|select name).name)+" ("+$name+")"

New-ADUser -Name $DispN -AccountPassword (ConvertTo-SecureString -String 'Aa12345' -AsPlainText -Force)  -GivenName $GN -Surname $SN -SamAccountName $newlogin -Description "для дополнительного п/я"  -Path $OU -Enabled $true -DisplayName $DispN -UserPrincipalName $UPN
Set-ADUser -identity $newlogin -add @{"extensionattribute13"="$domain"}

Set-ADUser -identity $newlogin -add @{"extensionattribute15"=6}
$output ="почтовый ящик $newlogin$mail к учетной записи $login" >> "\\polr-it01\_temp\PS\Создание п.я\output.txt"
}