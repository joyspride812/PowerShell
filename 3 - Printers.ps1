<#
ОПИСАНИЕ:
	Скрипт - приложение с GUI для  отображения инфо (Имя, Модель картриджа, Состояние, Расположение) о принтерах опубликованных на принтсервере с возможностями подключения в текущем сеансе и перехода на "WEB морду принтера"

Возможные улучшения:
	Ф-ия определения локации (принтсервера, префикса принтера) по IP ($loc)

УКАЗАТЕЛИ:
	function web # Функция запуска браузера с адресом WEB морды выбранного принтера
	Функция определение локации - в разработке
	Параметры главного окна
	Параметры таблицы
	Параметры элементов GUI
	function SortListView # функция сортировки столбцов таблицы по возрастанию\убыванию
#>

############################ФУНКЦИЯ WEB ###########################################################

function web ([parameter(Mandatory=$true, HelpMessage="Выберите принтер")][string]$SelectedItem,[string]$PrintServer) # Функция запуска браузера с адресом WEB морды выбранного принтера
	{
	lable.text ="Открывается WEB интерфейс"
	$PrintServer= $PrintServer.Replace("\","")
	$PortName=(Get-Printer -ComputerName $PrintServer -Name $SelectedItem).PortName
	$hostaddress=(Get-CimInstance -ClassName Win32_TCPIPPrinterPort -ComputerName $PrintServer |Where-Object -Property Name -EQ "$PortName" | Select-Object -Property Name, HostAddress).hostaddress
	$IP="http://$hostaddress";& "C:\Program Files\Google\Chrome\Application\chrome.exe" $IP 
	$lable.text ="Открывается WEB интерфейс"
	}
############################ФУНКЦИЯ WEB конец###########################################################


######################################Функция определение локации###########################################
	$loc="Полюстрово\Автопрайм"
	$PrintServer="**********"
######################################Функция определение локации КОНЕЦ#####################################







Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()


$printers=Get-Printer -ComputerName linx-print-01|Where-Object name -Match "pol-*" |Select-Object NAME, Comment, PrinterStatus


$LastColumnClicked = 0 # последний номер столбца, по которому был клик
$LastColumnAscending = $false # направление последней сортировки этого столбца
 


############################################# Параметры главного окна ####################################################
$Form = New-Object System.Windows.Forms.Form
$ListView = New-Object System.Windows.Forms.ListView
$Form.Text = "Принтеры"
$FORM.Font = New-Object System.Drawing.Font("Bahnschrift SemiBold SemiConden",10,[System.Drawing.FontStyle]::Regular)
$Form.Height = 500
$Form.Width = 730
$Form.DataBindings.DefaultDataSourceUpdateMode = 0
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.Icon = New-Object System.Drawing.Icon("\\int.rolfcorp.ru\sfc\RETAIL\DC_SPb\РОЛЬФ-Рено-Полюстровский\Отдел_информатизации\Printers\print2.ico")

############################################ Параметры главного окна конец ####################################################

$timer1 = New-Object System.Windows.Forms.Timer
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$MyShared = [HashTable]::Synchronized(@{}) # синхронизированная хэш-таблица



# "длительная" команда подключаем принтер
$bgCommand = 
	{
	$MyShared.Stop = $false
	$MyShared.Form.Controls["label"].Text="Проверяем принтер " + $MyShared.SelectedItem
		if((get-printer|Select-Object Name|? NAME -EQ $MyShared.PrinterName) -ne $NULL)
			{
			$MyShared.Form.Controls["label"].Text=$MyShared.SelectedItem+" уже подключен."
			} 
		else
			{
			$MyShared.Form.Controls["label"].Text="Подключаем принтер " + $MyShared.SelectedItem
			(New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($MyShared.PrinterName)
			if((get-printer|Select-Object Name|? NAME -EQ $MyShared.PrinterName) -eq $NULL)
				{
				$MyShared.Form.Controls["label"].Text="Непредвиденная ошибка, обратитесь к системному администратору.";
				} 
			else
				{
				if ($MyShared.check -eq "$true")  
					{
					[string]$ShareName=$MyShared.SelectedItem
					(Get-WmiObject -class win32_printer -Filter "ShareName='$ShareName'").SetDefaultPrinter()
					$MyShared.Form.Controls["label"].Text=$MyShared.SelectedItem + " подключен и установлен по умолчанию.";
					}  
				else  
					{  
					$MyShared.Form.Controls["label"].Text=$MyShared.SelectedItem + " подключен.";
					}  
				}
			}   
       
       $MyShared.Stop = $true
	}

$handler_timer1_Tick=
	{
	if ($MyShared.Stop)
       {
           $bgRunspace.dispose()
           $timer1.Enabled = $false
           $ProgressBar.hide()
           $ADDButton.Enabled = $True;
           $ListView.Enabled=$True;
           $CheckBox.Enabled=$True
       }
	}	

$handler_button1_Click=
	{
	$ADDButton.Enabled = $False
    $ListView.Enabled= $False
    $CheckBox.Enabled=$False
    $progressBar.Show()
    $MyShared.SelectedItem=$listView.SelectedItems[0].Text
    $MyShared.PrinterName=$PrintServer+$listView.SelectedItems[0].Text
    $MyShared.check=$CheckBox.Checked
    $bgRunspace = [RunSpaceFactory]::CreateRunspace()
    $bgRunspace.ApartmentState = "STA"
    $bgRunspace.ThreadOptions = "ReuseThread"
    $bgRunspace.Open()
    $bgRunspace.SessionStateProxy.setVariable("MyShared", $MyShared)
    $bgPowerShell = [PowerShell]::Create()
    $bgPowerShell.Runspace = $bgRunspace
    $bgPowerShell.AddScript($bgCommand).BeginInvoke()
    $timer1.Enabled = $true
	}

$OnLoadForm_StateCorrection=
	{
    $form.WindowState = $InitialFormWindowState
	}

###################################### Параметры таблицы ###########################################

$ListView.View = [System.Windows.Forms.View]::Details
$ListView.Width = $Form.ClientRectangle.Width
$ListView.Height = $Form.ClientRectangle.Height
$ListView.Anchor = "Top, Left, Right, Bottom"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 15
$System_Drawing_Point.Y = 32
$ListView.Location = $System_Drawing_Point
$ListView.Name = "Локация"
$ListView.View = 'Details'
$ListView.Height = 300
$ListView.Width = 680 
# Картинки
$imageList = new-Object System.Windows.Forms.ImageList 
$imageList.ImageSize = New-Object System.Drawing.Size(30,30) # Size of the pictures
$bitm1=[System.Drawing.Image]::FromFile("\\int.rolfcorp.ru\sfc\RETAIL\DC_SPb\РОЛЬФ-Рено-Полюстровский\Отдел_информатизации\Printers\print.ico")
$imageList.Images.Add("PI",$bitm1) 
$ListView.SmallImageList = $imageList
# Добавление таблицы в окно
$Form.Controls.Add($ListView)
 
# Добавление столбцов в таблицу
$ListView.Columns.Add("Имя",150) | Out-Null
$ListView.Columns.Add("Картридж",90) | Out-Null
$ListView.Columns.Add("Состояние",90) | Out-Null
$ListView.Columns.Add("Расположение",310) | Out-Null
$handler =
	{
    $selectedItem = $listView.SelectedItems[0]
    $SI= $($selectedItem.Text)
    return $SI
	}

# Подписываемся на событие выделения элемента
$listView.Add_SelectedIndexChanged($handler)

# Добавление строк
($printers|Sort-Object name).foreach({
	$ListViewItem = New-Object System.Windows.Forms.ListViewItem($_.name,0)
	$ListViewItem.Subitems.Add(($_.comment.split("(")[1]).Replace(")","")) | Out-Null
	$ListViewItem.Subitems.Add([string]$_.printerstatus) | Out-Null
	$ListViewItem.Subitems.Add($_.comment.split("(")[0]) | Out-Null
	$ListView.Items.Add($ListViewItem) | Out-Null
	})                                                                                                      


################################# Параметры элементов GUI ############################################################

# Надпись "Локация:" параметры

$label_location = New-Object System.Windows.Forms.Label # - объявляем
$label_location.Location = New-Object System.Drawing.Point(15,5) # - расположение
$label_location.Size = New-Object System.Drawing.Size(280,20) # - размер
$label_location.Text = "Локация: $loc" # - Текст
$label_location.Font = New-Object System.Drawing.Font("Arial Black",10,[System.Drawing.FontStyle]::Regular)
$form.Controls.Add($label_location)

# Чекбокс по-умолчанию" параметры
$CheckBox = New-Object System.Windows.Forms.CheckBox
$CheckBox.Text = 'Использовать по умолчанию'
$CheckBox.AutoSize = $true
$CheckBox.Checked = $true
$CheckBox.Location  = New-Object System.Drawing.Point(18,400)
$form.Controls.Add($CheckBox)


#Интерактивная строка параметры
$label = New-Object System.Windows.Forms.Label # - создаём надпись
$label.Location = New-Object System.Drawing.Point(18,350) # - расположение
$label.Size = New-Object System.Drawing.Size(380,20)
$label.Text = "Выберите принтер"
$label.name = "label"
$label.DataBindings.DefaultDataSourceUpdateMode = 0
$form.Controls.Add($label)

#Кнопка Подключить параметры
$ADDButton = New-Object System.Windows.Forms.Button # - создаём кнопку
$ADDButton.DataBindings.DefaultDataSourceUpdateMode = 0
$ADDButton.Location = New-Object System.Drawing.Point(510,385) # - расположение кнопки
$ADDButton.Size = New-Object System.Drawing.Size(95,33) # -  размер кнопки
$ADDButton.Text = "Подключить" # - текст кнопки
$ADDButton.TabIndex = 0
$ADDButton.UseVisualStyleBackColor = $True
$ADDButton.add_Click($handler_button1_Click)
$form.Controls.Add($ADDButton) # - добавляем кнопку на форму

#Кнопка Выход
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(610,385)
$CancelButton.Size = New-Object System.Drawing.Size(95,33)
$CancelButton.Text = 'Выход'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

#Кнопка WEB морда
$WEBButton = New-Object System.Windows.Forms.Button
$WEBButton.Location = New-Object System.Drawing.Point(410,385)
$WEBButton.Size = New-Object System.Drawing.Size(95,33)
$WEBButton.Text = 'WEB морда'
$WEBButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$WEBButton.add_Click({WEB $listView.SelectedItems[0].Text $PrintServer})

$form.CancelButton = $WEBButton
$form.Controls.Add($WEBButton)

#Банер Rolfтех
$PictureBox = New-Object System.Windows.Forms.PictureBox
$PictureBox.Load('\\int.rolfcorp.ru\sfc\RETAIL\DC_SPb\РОЛЬФ-Рено-Полюстровский\Отдел_информатизации\Printers\ROLFtech.jpg')
$PictureBox.Location  = New-Object System.Drawing.Point(556,3)
$PictureBox.Size = New-Object System.Drawing.Size (200,100)
$form.Controls.add($PictureBox)



#ProgressBar
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Size(18,370)
$ProgressBar.Size = New-Object System.Drawing.Size(300,20)
$ProgressBar.Style = "Marquee"
$ProgressBar.MarqueeAnimationSpeed = 25
$ProgressBar.visible
$ProgressBar.DataBindings.DefaultDataSourceUpdateMode = 0
$Form.Controls.Add($ProgressBar)
$ProgressBar.Hide()
 

$ListView.add_ColumnClick({SortListView $_.Column})
 



## Обработчик нажатия
function SortListView # функция сортировки столбцов таблицы по возрастанию\убыванию
	{
	param([parameter(Position=0)][UInt32]$Column)
	$Numeric = $true # определение метода сортировки возраст\убыв
 
	# если нажать на столбец дважды - меняется порядок сортировки, при первом нажатии сортировка по возрастанию
	if($Script:LastColumnClicked -eq $Column)
		{
		$Script:LastColumnAscending = -not $Script:LastColumnAscending
		}
	else
		{
		$Script:LastColumnAscending = $true
		}
	$Script:LastColumnClicked = $Column
	$ListItems = @(@(@(@()))) # трёхмерный массив; столбец 1 индексирует остальные столбцы ,столбец 2 - значение для сортировки, столбец 3 - объект System.Windows.Forms.ListViewItem
 
	foreach($ListItem in $ListView.Items)
		{
		# если все элементы числовые, можно использовать числовую сортировку
		if($Numeric -ne $false) # nothing can set this back to true, so don't process unnecessarily
			{
			try
				{
				$Test = [Double]$ListItem.SubItems[[int]$Column].Text
				}
			catch
				{
				$Numeric = $false # a non-numeric item was found, so sort will occur as a string
				}
			}		
		$ListItems += ,@($ListItem.SubItems[[int]$Column].Text,$ListItem)
		}
 
	$EvalExpression = 
		{
		if($Numeric) {return [Double]$_[0] }
		else {return [String]$_[0] }
		}
 
	$ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Ascending=$Script:LastColumnAscending}
 
	$ListView.BeginUpdate()
	$ListView.Items.Clear()
	
	foreach($ListItem in $ListItems)
		{
		$ListView.Items.Add($ListItem[1])
		}
	$ListView.EndUpdate()
	}
 

$MyShared.Form =$form
$timer1.Interval = 1000
$timer1.add_Tick($handler_timer1_Tick)



$InitialFormWindowState = $form.WindowState
$form.add_Load($OnLoadForm_StateCorrection)
$Response = $Form.ShowDialog()


