<#
ОПИСАНИЕ:
Скрипт выполняет выгрузки у.з. пользователей из базы SAP (sql запрос) и Active Directory и проверяет их на соответствие ранее установленным критериям. В результате формируется отчёт SAP-AD.xlsx и рассылается по указанным адресам вместе с инструкцие по обработке отчёта.

Исключения из проверки:
1)"ФИО SAP -ne ФИО AD" -не учитывается значение (CSAT: Х.Х ⭐) в атрибуте displayName
2)УЗ в AD без email и ext15 - проверка только по OU-Piter,OU-Moscow
3) УЗ в AD отключены и не находятся в disabled users - исключения из проверки:int.rolfcorp.ru/OU-Moscow/OU-Holding/Public Objects/Conference_Rooms/* ;int.rolfcorp.ru/Microsoft Exchange System Objects/* 

УКАЗАТЕЛИ:

    УКАЗАТЕЛИ РАСПОЛОЖЕНИЯ ФАЙЛОВ

    function SendReport # Функция отправки отчета

    ВЫГРУЗКА из SAP
	       function sap_status #Ф-я определяет, сотрудник работает\уволен\в декрете. Столбец Sap_status (Значения:Enabled,Disabled,Pregnancy)
	    function sap_fio # ф-я форматирует переменную $Sapfio

    Выгрузка из AD
	    Function ConvertSupportTeam #Ф-я определяет команду поддержки по OU. Столбец "Команда поддержки" (Значения:'Администраторы Аэропорт'..'Администраторы Ясенево',ИНЖИТ,Unassigned Users)
	    Function ConvertUserAccountControl  #Ф-я определяет включена\отключена учетная запись в AD столбец AD_Status (Значения: Enabled\Disabled)
	    Function CsatOFF  #Ф-я очищает свойство displayName от значения CSAT столбец displayName без CSAT

    Выгрузки ФИО SAP -ne ФИО AD; TN не найден; ДУБЛИ TN

    Выгрузка УЗ в AD без email и ext15 пуст

    Выгрузка УЗ в AD отключены и не находятся в disabled users

    Выгрузка УЗ в AD с незаполненными атрибутами

    Выгрузка УЗ в SAP - Disabled, в AD- ENABLED

    Выгрузка УЗ в SAP - Pregnancy, в AD- не в Pregnancy
#>


Import-Module activedirectory
Import-Module ImportExcel

#############################################################----УКАЗАТЕЛИ РАСПОЛОЖЕНИЯ ФАЙЛОВ--------######################################################

$WorkPath="$PSScriptRoot\" # Рабочий каталог
$SapDataPath=$WorkPath+"Sap-Data.csv" # Файл выгрузки из SAP
$AdDataPath=$WorkPath+"Ad-Data.csv" # Файл выгрузки из AD
$SapADPath=$WorkPath+"Sap-AD.xlsx" # Главный файл выгрузки
$ManualPath=$WorkPath+"SAP-AD инструкция.docx" # Файл с инструкцией
$EMails=@("AIPrygunov@rolftech.ru","SASilverstov@rolftech.ru","anmoiseenko@rolftech.ru") #Получатели отчёта
$CarbonCopy=@("smpolykov@rolftech.ru","daelezov@rolftech.ru") #Получатели отчёта в копии
$GuidPath=$WorkPath+"guid.csv" # Файл с guid команд поддержки

#############################################################----УКАЗАТЕЛИ РАСПОЛОЖЕНИЯ ФАЙЛОВ КОНЕЦ--------######################################################

Remove-Item -Path "$SapADPath" -Force

function SendReport # Функция отправки отчета
	{
	$UserName = "rolfnet\meow"
	$SecurePassword = "woof"|ConvertTo-SecureString -AsPlainText -Force
	$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
	Send-MailMessage -Credential $creds –SmtpServer smtprelay.rolf.ru -FROM 'msk-sysadmins@rolf.ru' -To $EMails -Cc $CarbonCopy -Subject "Отчет по проблемным учетным записям пользователей в AD." -Body "Коллеги, доброе утро!`n`nВо вложении отчет по проблемным учетным записям пользователей в AD и инструкция по обработке отчёта.`nПросьба обработать файл SAP-AD.xlsx согласно инструкции.`nХорошего дня! " -Encoding 'UTF8' -Attachments  $SapADPath, $ManualPath
	}


#############################################################----ВЫГРУЗКА из SAP--------######################################################

Write-Host "Выгрузка пользователей SAP начало"
Get-Date | Write-Host

$server = "******"
$database = "lite"
$query = "select a.id,a.surname,a.name,a.patronymic,a.netname,a.CFOname,a.out_date,a.busy,a.persg,b.werks_id,c.name2  from lite.dbo.user_rp_persons_sap a inner join lite.dbo.user_rp_employees_sap b on a.id =b.person_id join lite.dbo.user_rp_werks_sap c  on b.werks_id = c.id ORDER BY a.id"
$connectionTemplate = "Data Source={0};Integrated Security=SSPI;Initial Catalog={1};"
$connectionString = [string]::Format($connectionTemplate, $server, $database)
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
$command = New-Object System.Data.SqlClient.SqlCommand
$command.CommandText = $query
$command.Connection = $connection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $command
$DataSet = New-Object System.Data.DataSet
$row_count = $SqlAdapter.Fill($DataSet)
$connection.Close()
$SAPdata = @()

function sap_status ([string]$out_date,[string]$busy) #Ф-я определяет, сотрудник работает\уволен\в декрете. Столбец Sap_status (Значения:Enabled,Disabled,Pregnancy)
	{
		if ($out_date -eq " ")
		{
			if ($busy -eq "0"){$Sap_status ="Pregnancy"}
			else {$Sap_status ="Enabled"} 
		}
		else {$Sap_status ="Disabled"}
	return $Sap_status
	}
function sap_fio ([string]$surname,[string]$name,[string]$patronymic) # ф-я форматирует переменную $Sapfio
	{
	$SapFIO =$surname, $name,$patronymic -join " " 
	while ($Sapfio[-1] -eq " ")
		{$Sapfio=$Sapfio.Remove($Sapfio.Length - 1)}
	return $Sapfio
	}

foreach ($Row in ($DataSet.Tables[0].Rows|Where-Object persg -in "1","5","9"))
	{
	$SAPitem =  [PSCustomObject]@{
		Sap_Id = $Row.Item('id')
        Sap_FIO = sap_fio $Row.Item('surname'), $Row.Item('name'),$Row.Item('patronymic')
        Sap_NetName = $Row.Item('netname')
        Sap_Location = $Row.Item('name2')
        Sap_Department=$Row.Item('CFOname')
        Sap_Status=sap_status $Row.Item('out_date') $Row.Item('busy') 
        Sap_persg  =$Row.Item('persg')
        }
    $SAPdata += $SAPitem
	}

$SAPdata|export-csv -Path $SapDataPath -Encoding Default -Delimiter ";"
Write-Host "Выгрузка пользователей SAP конец"
Get-Date | Write-Host
#############################################################----ВЫГРУЗКА из SAP КОНЕЦ--------######################################################

############################################---------Выгрузка из AD-----------------#######################################################

$GUIDS=import-csv -Path $GuidPath -Encoding Default -Delimiter ";"
$GuidHash = @{}
foreach ($GUID in $GUIDS)
	{
    $GuidHash[$GUID.guid] = $GUID.SupportTeam
	}

Function ConvertSupportTeam ([string]$canonicalName) #Ф-я определяет команду поддержки по OU. Столбец "Команда поддержки" (Значения:'Администраторы Аэропорт'..'Администраторы Ясенево',ИНЖИТ,Unassigned Users)
	{
	$Cities=@("OU-Piter","OU-Moscow")
	$SupportTeam=$NULL
	if($canonicalName.split("/")[1] -in $Cities )
		{
		$x=($canonicalName.split("/")[2])
		$TempGuid=(Get-ADObject -Filter { Name -eq $x })
		$SupportTeam= $GuidHash[$TempGuid.ObjectGUID.Guid]
		if ($SupportTeam.Count -gt 1){$SupportTeam=$SupportTeam[0]}
		if ($SupportTeam -eq $NULL) {$SupportTeam="ИНЖИТ"}
		}
	elseif ($canonicalName.split("/")[1] -eq 'Disabled users' ){$SupportTeam ='Disabled users'}
	elseif ($canonicalName.split("/")[1] -eq 'Unassigned_Users' ){$SupportTeam ='Unassigned Users'}
	else {$SupportTeam ='ИНЖИТ'}
	return $SupportTeam
	}

Function ConvertUserAccountControl ([int]$UAC) #Ф-я определяет включена\отключена учетная запись в AD столбец AD_Status (Значения: Enabled\Disabled)
	{
	$status=$NULL
	$UACPropertyFlags = @(
	"SCRIPT",
	"ACCOUNTDISABLE",
	"RESERVED",
	"HOMEDIR_REQUIRED",
	"LOCKOUT",
	"PASSWD_NOTREQD",
	"PASSWD_CANT_CHANGE",
	"ENCRYPTED_TEXT_PWD_ALLOWED",
	"TEMP_DUPLICATE_ACCOUNT",
	"ENABLED",
	"RESERVED",
	"INTERDOMAIN_TRUST_ACCOUNT",
	"WORKSTATION_TRUST_ACCOUNT",
	"SERVER_TRUST_ACCOUNT",
	"RESERVED",
	"RESERVED",
	"DONT_EXPIRE_PASSWORD",
	"MNS_LOGON_ACCOUNT",
	"SMARTCARD_REQUIRED",
	"TRUSTED_FOR_DELEGATION",
	"NOT_DELEGATED",
	"USE_DES_KEY_ONLY",
	"DONT_REQ_PREAUTH",
	"PASSWORD_EXPIRED",
	"TRUSTED_TO_AUTH_FOR_DELEGATION",
	"RESERVED",
	"PARTIAL_SECRETS_ACCOUNT"
	"RESERVED"
	"RESERVED"
	"RESERVED"
	"RESERVED"
	"RESERVED"
	)
	$status = (0..($UACPropertyFlags.Length) | ?{$UAC -bAnd [math]::Pow(2,$_)} | %{$UACPropertyFlags[$_]}) -join ” | ”
	if ($status -match "ACCOUNTDISABLE"){$status="DISABLED"}
	if ($status -match "ENABLED"){$status="ENABLED"}
	return $status
	}

Function CsatOFF ([string]$displayName) #Ф-я очищает свойство displayName от значения CSAT столбец displayName без CSAT
	{
	if ($displayName -match "CSAT")
		{
		$displayName=($displayname -replace " \(CSAT.*")
		}
	return $displayName
	}

Write-Host "Выгрузка пользователей AD начало"
Get-Date | Write-Host

$ADUsersData = get-aduser -filter * -SearchBase "DC=int,DC=rolfcorp,DC=ru" -Properties ExtensionAttribute14, extensionattribute15, sAMAccountName,  displayName, cn, Name, mail, canonicalName, userAccountControl, description |Select-Object  @{name='Команда поддержки' ;e={ConvertSupportTeam $_.canonicalName}}, ExtensionAttribute14, extensionattribute15,sAMAccountName,   @{name='displayName без CSAT' ;e={CsatOFF $_.displayName}}, Name, cn, mail,  @{name='AD_Status' ;e={ConvertUserAccountControl $_.userAccountControl}}, description, canonicalName | Sort-Object ExtensionAttribute14
$ADUsersData|export-csv -Path $AdDataPath -Encoding Default -Delimiter ";"

Write-Host "Выгрузка пользователей AD конец"
Get-Date | Write-Host 

############################################---------Выгрузка из AD Конец-----------------#######################################################

#######################################----Выгрузки ФИО SAP -ne ФИО AD; TN не найден; ДУБЛИ TN----#######################################

Write-Host "Выгрузка  SAP_FIO_NE_AD_FIO начало"
Get-Date | Write-Host

$AdEnableData= ($ADUsersData| Where-Object ExtensionAttribute14 -ne $NULL|where-object {$_.ExtensionAttribute14 -notin "0".."9"} )
$SapOnData=($SAPdata|where-object Sap_Status -eq  "Enabled")
$TNgt1=($AdEnableData | Group-Object ExtensionAttribute14 | Where-Object { $_.Count -gt 1 }).group

$SAPHash = @{}

foreach ($SAPitem in $SAPdata) 
	{
    $SAPHash[[int]$SAPitem.Sap_Id] = $SAPitem
	}
$SapJoinTNgt1 = $TNgt1 | ForEach-Object -Process
	{
    [int]$ExtensionAttribute14=$_.ExtensionAttribute14
	[pscustomobject]@{
        "Команда поддержки"=$_."Команда поддержки"
        "AD ExtensionAttribute14"=$_.ExtensionAttribute14
        "AD Extensionattribute15"=$_.ExtensionAttribute15 
        "AD sAMAccountName"=$_.sAMAccountName
        "AD cn"=$_.cn
        "AD mail"=$_.mail
        "AD canonicalName"=$_.canonicalName
        "AD Status"=$_.AD_Status
        "AD description"=$_.description
        "SAP Id"= $SAPHash[$ExtensionAttribute14].Sap_Id
        "SAP ФИО"=$SAPHash[$ExtensionAttribute14].Sap_FIO
        "SAP Логин"=$SAPHash[$ExtensionAttribute14].Sap_NetName
        "SAP Локация"=$SAPHash[$ExtensionAttribute14].Sap_Location
        "SAP Департамент"=$SAPHash[$ExtensionAttribute14].Sap_Department
		}
	}
$SapJoinTNgt1|Export-Excel -Path $SapADPath -AutoSize  -WorksheetName 'ДУБЛИ ext14 в AD' 
$ADwithoutD=($AdEnableData|Where-Object ExtensionAttribute14 -NotIn $TNgt1.ExtensionAttribute14)
 
$ADHash = @{}

foreach ($ADitem in $ADwithoutD)
	{
    $ADHash[[int]$ADitem.ExtensionAttribute14] = $ADitem
	}

$SapJoinAD = $SapOnData | ForEach-Object -Process {
    [int]$Sap_id=$_.Sap_id
    [pscustomobject]@{
         "Команда поддержки"=$ADHash[$Sap_id]."Команда поддержки"   
         "SAP Табельный номер" = $_.Sap_Id
         "AD ExtensionAttribute14"= $ADHash[$Sap_id].ExtensionAttribute14
         "AD Extensionattribute15"= $ADHash[$Sap_id].Extensionattribute15    
         "SAP ФИО" = $_.Sap_FIO
         "AD displayName без CSAT"= $ADHash[$Sap_id]."displayName без CSAT"
         "AD cn"= $ADHash[$Sap_id].cn
         "AD Name"= $ADHash[$Sap_id].Name   
         "SAP Юр.лицо"=$_.Sap_Location
         "SAP Отдел"=$_.Sap_Department
         "AD должность"= $ADHash[$Sap_id].description   
         "SAP Логин" = $_.Sap_NetName
         "AD sAMAccountName"= $ADHash[$Sap_id].sAMAccountName
         "AD mail"= $ADHash[$Sap_id].mail
         "AD Enabled\disabled"= $ADHash[$Sap_id].AD_Status
         "AD canonicalName"= $ADHash[$Sap_id].canonicalName
		}
	}

$SapJoinAD|where-object "AD sAMAccountName" -eq $NULL|Select-Object "SAP Табельный номер","SAP ФИО","SAP Логин","SAP Юр.лицо", "SAP Отдел" |Export-Excel -Path $SapADPath -AutoSize  -WorksheetName 'ТН не найден в AD'
$SapJoinAD|where-object "AD sAMAccountName" -ne $NULL|Where-Object {$_."SAP ФИО" -ne $_."AD displayName без CSAT" -or $_."SAP ФИО" -ne $_."AD cn" -or $_."SAP ФИО" -ne $_."AD Name" -or $_."AD displayName без CSAT" -ne $_."AD cn" -or $_."AD displayName без CSAT" -ne $_."AD Name" -or $_."AD cn" -ne $_."AD Name"  }  |Select-Object 'Команда поддержки','SAP ФИО','AD displayName без CSAT','AD cn','AD Name','AD sAMAccountName','SAP Логин','AD canonicalName','SAP Юр.лицо','SAP Отдел','AD должность'|Export-Excel -Path $SapADPath -AutoSize  -WorksheetName 'ФИО SAP -ne ФИО AD'

$ADUsersData|Export-Excel -Path $SapADPath -AutoSize  -WorksheetName 'Справочник ADusers(all)' -HideSheet 'Справочник ADusers(all)' 
$SAPdata|Export-Excel -Path $SapADPath -AutoSize  -WorksheetName 'Справочник Sap(all)' -HideSheet 'Справочник Sap(all)'
$SapJoinAD|Export-Excel -Path $SapADPath -AutoSize  -WorksheetName 'Справочник Sap(on)-AD' -HideSheet 'Справочник Sap(on)-AD'

Write-Host "Выгрузка SAP_FIO_NE_AD_FIO конец"
Get-Date | Write-Host

#######################################----Выгрузки ФИО SAP -ne ФИО AD; TN не найден; ДУБЛИ TN Конец----#######################################

#######################################---Выгрузка УЗ в AD без email и ext15 пуст-----######################################

$Cities=@("OU-Piter","OU-Moscow")
$AdOn=(Import-Csv $AdDataPath -Delimiter ";" -Encoding Default|where-object AD_Status -eq  "ENABLED"|Where-Object {($_.canonicalName.split("/")[1]) -in $Cities})
$AD_on_without_mail=($AdOn|where-object {$_.mail -EQ "" -and ($_.extensionattribute15 -eq "0" -or $_.extensionattribute15 -eq "")})|Select-Object 'Команда поддержки','ExtensionAttribute14','extensionattribute15','sAMAccountName','cn','description','canonicalName'|Export-Excel -Path $SapADPath -AutoSize -WorksheetName 'УЗ без Email и ext15'

#######################################---Выгрузка УЗ в AD без email и ext15 пуст конец-----######################################

#######################################---Выгрузка УЗ в AD отключены и не находятся в disabled users-----######################################

$AdOff=($ADUsersData|where-object AD_Status -eq  "DISABLED"|Where-Object {($_.canonicalName.split("/")[1]) -ne "Microsoft Exchange System Objects"}|Where-Object {$_.canonicalName -notmatch "int.rolfcorp.ru/OU-Moscow/OU-Holding/Public Objects/Conference_Rooms"})
($AdOff|where-object "Команда поддержки" -ne "Disabled users")|Select-Object 'Команда поддержки','ExtensionAttribute14','ExtensionAttribute15','sAMAccountName','cn','mail','description','canonicalName'|Export-Excel -Path $SapADPath -AutoSize -WorksheetName 'УЗ отключена и не в Disabled us'

#######################################---Выгрузка УЗ в AD отключены и не находятся в disabled users конец-----######################################

######################################---Выгрузка УЗ в AD с незаполненными атрибутами---#######################################

$AdOn=(Import-Csv $AdDataPath -Delimiter ";" -Encoding Default|where-object AD_Status -eq  "ENABLED")
($AdOn|where-object {$_."Команда поддержки" -ne "Disabled users"}|where-object  {($_.ExtensionAttribute14 -notin "0".."6" -and $_.ExtensionAttribute15 -notin "0","2","3") -or ($_.ExtensionAttribute14 -eq "" -and $_.ExtensionAttribute15 -notin "5","6") -or ($_.ExtensionAttribute14 -eq "0" -and $_.ExtensionAttribute15 -ne "1") -or ($_.ExtensionAttribute14 -in "1".."6") })|Select-Object 'Команда поддержки','ExtensionAttribute14','ExtensionAttribute15','sAMAccountName','cn','mail','AD_Status','description','canonicalName'|Export-Excel -Path $SapADPath -AutoSize  -WorksheetName 'УЗ AD c незаполненными атриб'

######################################---Выгрузка УЗ в AD с незаполненными атрибутами конец---#######################################

######################################---Выгрузка УЗ в SAP - Disabled, в AD- ENABLED ---#######################################

Write-Host "Начало работы выгрузки SAP - Disabled, в AD- ENABLED"
Get-Date | Write-Host

$SapOffData=(Import-Csv $SapDataPath -Delimiter ";" -Encoding Default|where-object Sap_Status -eq  "Disabled"|Select-Object sap_id,sap_fio,Sap_Location,Sap_Department)
$ADenabled=(Import-Csv $AdDataPath -Delimiter ";" -Encoding Default|where-object {$_.AD_Status -eq  "ENABLED" -and $_.ExtensionAttribute14 -ne $NULL} |select-object ExtensionAttribute14,"Команда поддержки",sAMAccountName,description,canonicalName)

$LookupHash = @{}

foreach ($ADitem in $ADenabled) {
    $LookupHash[$ADitem.ExtensionAttribute14] = $ADitem
	}
$SapOffJoinAd = $SapOffData | ForEach-Object -Process {
    [pscustomobject]@{
        "Команда поддержки"=$LookupHash[$_.sap_id]."Команда поддержки"
        "SAP Id"= $_.sap_id
        "SAP ФИО"=$_.sap_fio
        "AD sAMAccountName" =$LookupHash[$_.sap_id].sAMAccountName
        "AD canonicalName"=$LookupHash[$_.sap_id].canonicalName
        "AD description"=$LookupHash[$_.sap_id].description
        "SAP Локация"=$_.Sap_Location
        "SAP Департамент"=$_.Sap_Department
		}
	}
$SapOffJoinAd|Where-Object "AD sAMAccountName" -ne $NULL|Export-Excel -Path $SapADPath -AutoSize  -WorksheetName 'SAP disabled AD enabled'

######################################---Выгрузка УЗ в SAP - Disabled, в AD- ENABLED конец---#######################################

Write-Host "Окончание  работы выгрузки SAP - Disabled, в AD- ENABLED"
Get-Date | Write-Host

######################################---Выгрузка УЗ в SAP - Pregnancy, в AD- не в Pregnancy ---#######################################

$Sappregnacy=($SAPdata|where-object Sap_Status -eq  "Pregnancy"|Select-Object sap_id,sap_fio,Sap_Location,Sap_Department)

$SapPRegJoinAd = $Sappregnacy | ForEach-Object -Process {
    [int]$sap_id=$_.sap_id
    [pscustomobject]@{
       "Команда поддержки"=$ADHash[$sap_id]."Команда поддержки"
        "SAP Id"= $_.sap_id
        "SAP ФИО"=$_.sap_fio
        sAMAccountName =$ADHash[$sap_id].sAMAccountName
        ExtensionAttribute15 = $ADHash[$sap_id].ExtensionAttribute15
        canonicalName=$ADHash[$sap_id].canonicalName
        description=$ADHash[$sap_id].description
        "SAP Локация"=$_.Sap_Location
        "SAP Департамент"=$_.Sap_Department
		}
	}
$SapPRegJoinAd|where-object {$_.canonicalName.split("/")[2] -notin 'Pregnancy','Pregnancy_mail','Pregnancy_Unmail' -or $_.ExtensionAttribute15 -ne "2"}|Export-Excel -Path $SapADPath -AutoSize  -WorksheetName 'SAP Pregnancy' -HideSheet 'Справочник Sap(all)'

SendReport #Отправка отчёта

######################################---Выгрузка УЗ в SAP - Pregnancy, в AD- не в Pregnancy конец---#######################################

Write-Host "Окончание  работы скрипта"
Get-Date | Write-Host