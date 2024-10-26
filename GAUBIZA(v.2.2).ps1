                   FUNCTION browser_and_cash
####################    КНОПКА IE   ####################
{
param ([string]$USER,[string]$PC)
browser $USER $PC
cash $USER $PC
}
###############################################################################################
                          Function Browser
#####################   СБРОС НАСТРОЕК БРАУЗЕРА    ########################
{

param ([string]$USER,[string]$PC)

[string]$auto="http://proxy/wpad.pac" #Сценарий автоконфигурации

[int]$proxyEn=0 #Флажок использовать прокси

[int]$AutoDetect=0 #Флажок автоматическое определение параметров

 

$id=get-AdUser $USER -Properties SID

$regKey="Registry::HKEY_USERS\" + $id.SID +"\Software\Microsoft\Windows\CurrentVersion\Internet Settings"

$session = new-pssession -computername $PC

$label_status.Text = "Получаем настройки proxy server от $PC  …"

# $proxyServerOld= Invoke-Command -session $session –ScriptBlock {param($regKey)Get-ItemProperty -path $regKey -ErrorAction SilentlyContinue} -arg $regKey


$label_status.Text = "Применение стандартных настроек..."

Invoke-Command -session $session –ScriptBlock {param($regkey,$auto) Set-ItemProperty -Path $regkey -Name AutoConfigURL -Value $auto} -arg $regKey, $auto

Invoke-Command -session $session –ScriptBlock {param($regkey,$proxyEn) Set-ItemProperty -Path $regkey -Name ProxyEnable -Value $proxyEn} -arg $regKey, $proxyEn

Invoke-Command -session $session –ScriptBlock {param($regkey,$AutoDetect) New-Itemproperty -Path $regkey -Name AutoDetect  -Value $AutoDetect -Force} -arg $regKey, $AutoDetect

$label_status.Text = "Настройки Применены"

# $proxyServerNew= Invoke-Command -session $session –ScriptBlock {param($regKey)Get-ItemProperty -path $regKey -ErrorAction SilentlyContinue} -arg $regKey

 
$label_status.Text = "Отключение блокировки всплывающих окон..."

[int]$temp="1"
for($temp=1; $temp -le 3; $temp++){
$regKey="Registry::HKEY_USERS\" + $id.SID + "\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" + $temp
 Invoke-Command -session $session –ScriptBlock {param($regkey) Set-ItemProperty -Path $regkey -Name 1809 -Value 3 -Force} -arg $regKey
 
 }
 remove-pssession -computername $PC
$label_status.Text ="Блокировка всплывающих окон отключена."
 Add-ADGroupMember -Identity SQ_SWITCHED_USERS -Members $USER
 Add-ADGroupMember -Identity SQ_Internet_WL -Members $USER
 $label_status.Text ="Пользователь $USER добавлен в группы:SQ_SWITCHED_USERS и SQ_Internet_WL"
}

################################################################################################

                             Function Cash
######################   ЧИСТКА ВРЕМЕННЫХ ФАЙЛОВ   ########################
{

Param([string]$user,[string]$PC)

 $label_status.Text = "Очистка кэша..."
$session = new-pssession -computername $PC

Invoke-Command -session $session –ScriptBlock {$disk=$env:appdata.Chars(0)+":"

Remove-Item -Path "$disk\Users\$user\AppData\Local\Microsoft\Windows\Temporary Internet Files\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose

Remove-Item -Path "$disk\Users\$user\AppData\Local\Microsoft\Windows\WER\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose

Remove-Item -Path "$disk\Users\$user\AppData\Local\Microsoft\Windows\INetCache\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose

Remove-Item -Path "$disk\Users\$user\AppData\Local\Microsoft\Windows\WebCache\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose

Remove-Item -Path "$disk\Users\$user\AppData\Local\Microsoft\Windows\WebCache\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose

Remove-Item -Path "$disk\Users\$user\AppData\Local\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose

Remove-Item -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose

Remove-Item -Path "$disk\Users\$user\AppData\Local\Microsoft\Windows\AppCache\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose

Remove-Item -Path "C:\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose

Remove-Item -Path "$disk\Users\$user\Recent\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
}
 
$label_status.Text = "Очистка завершена. Пользователю необходимо перезапустить браузер"
 
remove-pssession -computername $PC
}

##################################################################################################
                             Function Printers
###################### ОШИБКА ПОДКЛЮЧЕНИЯ ПРИНТЕРА  ########################
{
Param([string]$user,[string]$PC)
$label_status.Text = "Пользователь : $user  Компьютер : $PC"
$id=get-AdUser $USER -Properties SID
$regKey="Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows NT\Printers\PointAndPrint"
$paramName="RestrictDriverInstallationToAdministrators"
$session = new-pssession -computername $PC

$label_status.Text = "Изменение реестра Компьютера: $PC ..."
Invoke-Command -session $session –ScriptBlock {param($regKey, $paramName) New-Itemproperty -Path $regKey -Name $paramName -Value 0 -Force} -arg $regKey, $paramName
$label_status.Text = "Изменения применены, пользователь может перезагрузить ПК."
}
##################################################################################################
                                
                             Function Accept
###################### Подключение без подтверждения ##########################
{
Param([string]$PC)
$s = new-pssession -computername $PC
Invoke-Command -session $s -scriptblock {Set-ItemProperty -Path registry::"HKlm\SOFTWARE\DameWare Development\Mini Remote Control Service\Settings" -Name "Permission Required" -Value 0}
$label_status.Text = "Изменения применены. Подключение к $PC не требует подтверждения."
}
##################################################################################################                                
                                  Function LOGON
#######################   Получение имени текущего пользователя    ###############################
{
param([string]$PC)
$user=Get-WMIObject -Class Win32_ComputerSystem -Computer $PC |Select-Object Username
[string]$textbox=$user.username -replace ("(.+?)\\")
$label_status.Text = $textbox
$textBox_USER_NAME.Text=$textbox
}
##################################################################################################                                  
                               Function arms_print
 ######################    Копирование на рабочий стол ярлыков ARMS и Принтеры    ################ 
 {param([string]$PC,[string]$USER)
 $s = new-pssession -computername $PC
 [boolean]$LD = Invoke-Command -session $s –ScriptBlock {Test-Path D:\users}
if($LD){$LocDisk="d$"}
else {$LocDisk="c$"}
 Copy-Item "\\dlc-fs-05\#cfs$\Pol-Renault\Отдел_информатизации\ARMS\ARMS.lnk" "\\$PC\$LocDisk\Users\$USER\Desktop" -Force
 $label_status.Text = "Ярлык ARMS скопирован" 
  Copy-Item "\\dlc-fs-05\#cfs$\Pol-Renault\Отдел_информатизации\ARMS\ПРИНТЕРЫ.lnk" "\\$PC\$LocDisk\Users\$USER\Desktop" -Force
  $label_status.Text = "Ярлык ПРИНТЕРЫ скопирован"                         
  }
                                                             
##################################################################################################                                   
									FUNCTION StandartPass
 ######################    Установка стандартного пароля локального администратора    ################ 
                     
{
param([string]$PC)
$s = new-pssession -computername $PC
 Invoke-Command -session $s –ScriptBlock {Set-LocalUser -Name администратор -Password (ConvertTo-SecureString "6548624" -AsPlainText -Force)    }

  $label_status.Text = "Стандартный пароль установлен"     
}
						 Function Menu
######################    Графический ИНТЕРФЕЙС   ##########################
{
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form # - создаём форму
$form.Text = 'ГАУБИЦА V2.2' # -Заголовок
$form.Size = New-Object System.Drawing.Size(300,400) # - размер
$form.StartPosition = 'CenterScreen' # - расположение на экране

$OKButton = New-Object System.Windows.Forms.Button # - создаём кнопку
$OKButton.Location = New-Object System.Drawing.Point(15,120) # - расположение кнопки
$OKButton.Size = New-Object System.Drawing.Size(75,23) # -  размер кнопки
$OKButton.Text = "IE" # - текст кнопки
$OKButton.ADD_click({browser_and_cash $textBox_USER_NAME.text $textBox_PC_NAME.text})
$form.Controls.Add($OKButton) # - добавляем кнопку на форму

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(190,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Выход'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$AcceptButton = New-Object System.Windows.Forms.Button # - создаём кнопку
$AcceptButton.Location = New-Object System.Drawing.Point(15,150) # - расположение кнопки
$AcceptButton.Size = New-Object System.Drawing.Size(75,23) # -  размер кнопки
$AcceptButton.Text = "NO Accept" # - текст кнопки
$AcceptButton.ADD_click({Accept $textBox_PC_NAME.text})
$form.Controls.Add($AcceptButton) # - добавляем кнопку на форму

$LOGONButton = New-Object System.Windows.Forms.Button # - создаём кнопку
$LOGONButton.Location = New-Object System.Drawing.Point(100,150) # - расположение кнопки
$LOGONButton.Size = New-Object System.Drawing.Size(75,23) # -  размер кнопки
$LOGONButton.Text = "Logon User" # - текст кнопки
$LOGONButton.ADD_click({LOGON $textBox_PC_NAME.text})
$form.Controls.Add($LOGONButton) # - добавляем кнопку на форму

$label_PC_NAME = New-Object System.Windows.Forms.Label # - создаём надпись
$label_PC_NAME.Location = New-Object System.Drawing.Point(10,20) # - расположение
$label_PC_NAME.Size = New-Object System.Drawing.Size(280,20)
$label_PC_NAME.Text = 'Введите имя ПК:'
$form.Controls.Add($label_PC_NAME)

$textBox_PC_NAME = New-Object System.Windows.Forms.TextBox
$textBox_PC_NAME.Location = New-Object System.Drawing.Point(10,40)
$textBox_PC_NAME.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox_PC_NAME)

$label_USER_NAME = New-Object System.Windows.Forms.Label
$label_USER_NAME.Location = New-Object System.Drawing.Point(10,70)
$label_USER_NAME.Size = New-Object System.Drawing.Size(280,20)
$label_USER_NAME.Text = 'Введите имя Пользователя:'
$form.Controls.Add($label_USER_NAME)

$textBox_USER_NAME = New-Object System.Windows.Forms.TextBox
$textBox_USER_NAME.Location = New-Object System.Drawing.Point(10,90)
$textBox_USER_NAME.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox_USER_NAME)


$PrintersButton = New-Object System.Windows.Forms.Button # - создаём кнопку
$PrintersButton.Location = New-Object System.Drawing.Point(100,120) # - расположение кнопки
$PrintersButton.Size = New-Object System.Drawing.Size(75,23) # -  размер кнопки
$PrintersButton.Text = "Printers" # - текст кнопки
$PrintersButton.ADD_click({Printers $textBox_USER_NAME.text $textBox_PC_NAME.text})
$form.Controls.Add($PrintersButton) # - добавляем кнопку на форму

$arms_printButton = New-Object System.Windows.Forms.Button # - создаём кнопку
$arms_printButton.Location = New-Object System.Drawing.Point(190,150) # - расположение кнопки
$arms_printButton.Size = New-Object System.Drawing.Size(75,23) # -  размер кнопки
$arms_printButton.Text = "Arms+Print" # - текст кнопки
$arms_printButton.ADD_click({arms_print $textBox_PC_NAME.text $textBox_USER_NAME.text})
$form.Controls.Add($arms_printButton) # - добавляем кнопку на форму


$StandartPass = New-Object System.Windows.Forms.Button # - создаём кнопку
$StandartPass.Location = New-Object System.Drawing.Point(15,180) # - расположение кнопки
$StandartPass.Size = New-Object System.Drawing.Size(75,23) # -  размер кнопки
$StandartPass.Text = "Password" # - текст кнопки
$StandartPass.ADD_click({StandartPass $textBox_PC_NAME.text})
$form.Controls.Add($StandartPass) # - добавляем кнопку на форму

$label_status = New-Object System.Windows.Forms.Label # - создаём надпись
$label_status.Location = New-Object System.Drawing.Point(10,240) # - расположение
$label_status.Size = New-Object System.Drawing.Size(280,60)
$label_status.Text = ''
$form.Controls.Add($label_status)

$form.Topmost = $true

$form.Add_Shown({$textBox_PC_NAME.Select()})
$result = $form.ShowDialog()

#if ($result -eq [System.Windows.Forms.DialogResult]::OK)
#{
#    $x = $textBox.Text
 #   $x
#}
}

##################START##################
                   Menu