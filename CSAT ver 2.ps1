<#
ОПИСАНИЕ:
Скрипт загружает данные из файла \\dlc-fs-02\CSAT\Login CSAT.csv;
Проверяет значения CSAT в выгрузке на пригодность (значения от 0.1 до 5.0);
Изменяет атрибуты (extensionattribute11, DisplayName) в карточках пользователей в AD  согласно файлу выгрузки CSAT.
События связанные с ошибками во время работы скрипта регистрируются в файле журнала. \\dlc-fs-02\CSAT\Log\SET_DispName
При возникновении ошибок предусмотрена отправка уведомления на п.я. ответственного лица. На данный момент anmoiseenko@rolf.ru
#>

Import-Module ActiveDirectory

$EMAIL = 'anmoiseenko@rolf.ru' # Email ответственного лица для отправки alert
$ErrorCheck = $false # Проверка на событие ошибок 
$CSATinput = "\\dlc-fs-02\CSAT\Login CSAT.csv" # Файл Выгрузки CSAT



function SendAlert # Функция отправки alert ответственному лицу в случае возникновения ошибок
	{
	Send-MailMessage -FROM 'alert_CSAT_script@rolf.ru' -to $EMAI -Subject "Возникла ошибка при выполнении скрипта CSAT.ps1 на хосте rolf-hpsim." -body "Возникла ошибка при выполнении скрипта CSAT.ps1 на хосте rolf-hpsim. Файл журнала \\dlc-fs-02\CSAT\Log\SET_DispName\" –SmtpServer 'smtprelay.rolf.ru' -Encoding 'UTF8' 
	}

#Получение Данных из файла выгрузки CSAT
Try {$CSATData = Import-CSV -Path $CSATinput -delimiter ";" -Encoding Default}
catch [System.IO.FileNotFoundException]{$ErrorCheck = $true; "Отсутствует файл выгрузки CSAT \\dlc-fs-02\CSAT\Login CSAT.csv" | Out-File -FilePath "\\dlc-fs-02\CSAT\Log\SET_DispName\CSAT_Error_$(get-date -f yyyy-MM-dd).txt" -Append}

#Обработка файла выгрузки
$CSATData.ForEach({
	Try {
		$CatchCsat=$_.CSAT # Переменная используется в выводе catch
		$CatchLogin=$_.login # Переменная используется в выводе catch
		$UserData = Get-ADUser $_.Login -Properties extensionattribute11, CN # Выгрузка из AD
		If ($UserData.extensionattribute11 -eq $_.CSAT){} # Проверка на совпадение значения csat c 11м атрибутом. В случае совпадения изменение DisplayName не требуется. Отключить проверку при первом запуске скрипта.
		else 
			{    
			$CSAT = [single]$_.CSAT 
			$DisplayCSAT = $UserData.CN+ " (CSAT: "+$_.CSAT+" ⭐)" # Формирование DisplayName c CSAT
				if ($CSAT -lt 0.1 -or $CSAT -gt 5 -or $null ) # проверка на соответствие значения CSAT
					{
					$ErrorCheck = $true; "У пользователя $($_.Login) некорректное значение CSAT: $($_.CSAT) в выгрузке \\dlc-fs-02\CSAT\Login CSAT.csv  Значение CSAT должно быть в промежутке от 0.1 до 5.0" | Out-File -FilePath "\\dlc-fs-02\CSAT\Log\SET_DispName\CSAT_Error_$(get-date -f yyyy-MM-dd).txt" -Append
					}
				else 
					{ 
					Set-ADUser –Identity $_.Login -replace @{"extensionattribute11"=$_.CSAT}
					Set-ADUser –Identity $_.Login -DisplayName "$DisplayCSAT"
					}
			}
		}
	Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] 
		{ 
		$ErrorCheck = $true;
		"Ошибка получения данных пользователя $($CatchLogin) из AD: $($error[0])" | Out-File -FilePath "\\dlc-fs-02\CSAT\Log\SET_DispName\CSAT_Error_$(get-date -f yyyy-MM-dd).txt" -Append
		}
	Catch [System.Management.Automation.RuntimeException]
		{
		$ErrorCheck = $true;
		"У пользователя $CatchLogin некорректное значение CSAT: $CatchCsat в выгрузке \\dlc-fs-02\CSAT\Login CSAT.csv  Значение CSAT должно быть в промежутке от 0.1 до 5.0" | Out-File -FilePath "\\dlc-fs-02\CSAT\Log\SET_DispName\CSAT_Error_$(get-date -f yyyy-MM-dd).txt" -Append
		}
})

###Отправки alert ответственному лицу в случае возникновения ошибок

if ($errorCheck) {SendAlert}