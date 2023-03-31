<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MDIParent1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDIParent1))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ФайлToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.АвторизоватьсяToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ВыходToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВидToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СтрокаСостоянияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.БДToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.АрхивированиеБДToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВосстановлениеРезервнойКопииБДToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВосстановлениеВИсходныйВидБДToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИнструментыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.МастерПостроенияОтчетовToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоставитьНаКонтрольToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.НастройкиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЖурналСобытийToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.КнигиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИсходящиеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВходящиеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДоговорыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПриказыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПриказыПоСтудентамToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПриказыПоАХДToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДокладныеИСлужебныеЗапискиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.АктыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.РаспоряженияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДоверенностиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОкнаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.КаскадомToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СлеваНаправоToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СверхуВнизToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ЗакрытьВсеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СправкаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОПрограммеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОАвтореToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЛицензионноеСоглашениеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.НаписатьПисьмоРазработчикуToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel4 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel5 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel6 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 10
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ФайлToolStripMenuItem, Me.ВидToolStripMenuItem, Me.БДToolStripMenuItem, Me.ИнструментыToolStripMenuItem, Me.КнигиToolStripMenuItem, Me.ОкнаToolStripMenuItem, Me.СправкаToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(632, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ФайлToolStripMenuItem
        '
        Me.ФайлToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.АвторизоватьсяToolStripMenuItem, Me.ToolStripMenuItem1, Me.ВыходToolStripMenuItem})
        Me.ФайлToolStripMenuItem.Name = "ФайлToolStripMenuItem"
        Me.ФайлToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.ФайлToolStripMenuItem.Text = "Файл"
        '
        'АвторизоватьсяToolStripMenuItem
        '
        Me.АвторизоватьсяToolStripMenuItem.Name = "АвторизоватьсяToolStripMenuItem"
        Me.АвторизоватьсяToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.АвторизоватьсяToolStripMenuItem.Text = "Авторизоваться"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(158, 6)
        '
        'ВыходToolStripMenuItem
        '
        Me.ВыходToolStripMenuItem.Name = "ВыходToolStripMenuItem"
        Me.ВыходToolStripMenuItem.Size = New System.Drawing.Size(161, 22)
        Me.ВыходToolStripMenuItem.Text = "Выход"
        '
        'ВидToolStripMenuItem
        '
        Me.ВидToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.СтрокаСостоянияToolStripMenuItem})
        Me.ВидToolStripMenuItem.Name = "ВидToolStripMenuItem"
        Me.ВидToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.ВидToolStripMenuItem.Text = "Вид"
        '
        'СтрокаСостоянияToolStripMenuItem
        '
        Me.СтрокаСостоянияToolStripMenuItem.Name = "СтрокаСостоянияToolStripMenuItem"
        Me.СтрокаСостоянияToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.СтрокаСостоянияToolStripMenuItem.Text = "Строка состояния"
        '
        'БДToolStripMenuItem
        '
        Me.БДToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.АрхивированиеБДToolStripMenuItem, Me.ВосстановлениеРезервнойКопииБДToolStripMenuItem, Me.ВосстановлениеВИсходныйВидБДToolStripMenuItem})
        Me.БДToolStripMenuItem.Name = "БДToolStripMenuItem"
        Me.БДToolStripMenuItem.Size = New System.Drawing.Size(34, 20)
        Me.БДToolStripMenuItem.Text = "БД"
        '
        'АрхивированиеБДToolStripMenuItem
        '
        Me.АрхивированиеБДToolStripMenuItem.Name = "АрхивированиеБДToolStripMenuItem"
        Me.АрхивированиеБДToolStripMenuItem.Size = New System.Drawing.Size(280, 22)
        Me.АрхивированиеБДToolStripMenuItem.Text = "Архивирование БД"
        '
        'ВосстановлениеРезервнойКопииБДToolStripMenuItem
        '
        Me.ВосстановлениеРезервнойКопииБДToolStripMenuItem.Name = "ВосстановлениеРезервнойКопииБДToolStripMenuItem"
        Me.ВосстановлениеРезервнойКопииБДToolStripMenuItem.Size = New System.Drawing.Size(280, 22)
        Me.ВосстановлениеРезервнойКопииБДToolStripMenuItem.Text = "Восстановление резервной копии БД"
        '
        'ВосстановлениеВИсходныйВидБДToolStripMenuItem
        '
        Me.ВосстановлениеВИсходныйВидБДToolStripMenuItem.Name = "ВосстановлениеВИсходныйВидБДToolStripMenuItem"
        Me.ВосстановлениеВИсходныйВидБДToolStripMenuItem.Size = New System.Drawing.Size(280, 22)
        Me.ВосстановлениеВИсходныйВидБДToolStripMenuItem.Text = "Восстановление в исходный вид БД"
        '
        'ИнструментыToolStripMenuItem
        '
        Me.ИнструментыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.МастерПостроенияОтчетовToolStripMenuItem, Me.ПоставитьНаКонтрольToolStripMenuItem, Me.НастройкиToolStripMenuItem, Me.ЖурналСобытийToolStripMenuItem})
        Me.ИнструментыToolStripMenuItem.Name = "ИнструментыToolStripMenuItem"
        Me.ИнструментыToolStripMenuItem.Size = New System.Drawing.Size(95, 20)
        Me.ИнструментыToolStripMenuItem.Text = "Инструменты"
        '
        'МастерПостроенияОтчетовToolStripMenuItem
        '
        Me.МастерПостроенияОтчетовToolStripMenuItem.Name = "МастерПостроенияОтчетовToolStripMenuItem"
        Me.МастерПостроенияОтчетовToolStripMenuItem.Size = New System.Drawing.Size(229, 22)
        Me.МастерПостроенияОтчетовToolStripMenuItem.Text = "Мастер построения отчетов"
        '
        'ПоставитьНаКонтрольToolStripMenuItem
        '
        Me.ПоставитьНаКонтрольToolStripMenuItem.Name = "ПоставитьНаКонтрольToolStripMenuItem"
        Me.ПоставитьНаКонтрольToolStripMenuItem.Size = New System.Drawing.Size(229, 22)
        Me.ПоставитьНаКонтрольToolStripMenuItem.Text = "Поставить на контроль"
        '
        'НастройкиToolStripMenuItem
        '
        Me.НастройкиToolStripMenuItem.Name = "НастройкиToolStripMenuItem"
        Me.НастройкиToolStripMenuItem.Size = New System.Drawing.Size(229, 22)
        Me.НастройкиToolStripMenuItem.Text = "Настройки"
        '
        'ЖурналСобытийToolStripMenuItem
        '
        Me.ЖурналСобытийToolStripMenuItem.Name = "ЖурналСобытийToolStripMenuItem"
        Me.ЖурналСобытийToolStripMenuItem.Size = New System.Drawing.Size(229, 22)
        Me.ЖурналСобытийToolStripMenuItem.Text = "Журнал событий"
        '
        'КнигиToolStripMenuItem
        '
        Me.КнигиToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ИсходящиеToolStripMenuItem, Me.ВходящиеToolStripMenuItem, Me.ДоговорыToolStripMenuItem, Me.ПриказыToolStripMenuItem, Me.ПриказыПоСтудентамToolStripMenuItem, Me.ПриказыПоАХДToolStripMenuItem, Me.ДокладныеИСлужебныеЗапискиToolStripMenuItem, Me.АктыToolStripMenuItem, Me.РаспоряженияToolStripMenuItem, Me.ДоверенностиToolStripMenuItem})
        Me.КнигиToolStripMenuItem.Name = "КнигиToolStripMenuItem"
        Me.КнигиToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.КнигиToolStripMenuItem.Text = "Книги"
        '
        'ИсходящиеToolStripMenuItem
        '
        Me.ИсходящиеToolStripMenuItem.Name = "ИсходящиеToolStripMenuItem"
        Me.ИсходящиеToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ИсходящиеToolStripMenuItem.Text = "Исходящие"
        '
        'ВходящиеToolStripMenuItem
        '
        Me.ВходящиеToolStripMenuItem.Name = "ВходящиеToolStripMenuItem"
        Me.ВходящиеToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ВходящиеToolStripMenuItem.Text = "Входящие"
        '
        'ДоговорыToolStripMenuItem
        '
        Me.ДоговорыToolStripMenuItem.Name = "ДоговорыToolStripMenuItem"
        Me.ДоговорыToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ДоговорыToolStripMenuItem.Text = "Договоры"
        '
        'ПриказыToolStripMenuItem
        '
        Me.ПриказыToolStripMenuItem.Name = "ПриказыToolStripMenuItem"
        Me.ПриказыToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ПриказыToolStripMenuItem.Text = "Приказы"
        '
        'ПриказыПоСтудентамToolStripMenuItem
        '
        Me.ПриказыПоСтудентамToolStripMenuItem.Name = "ПриказыПоСтудентамToolStripMenuItem"
        Me.ПриказыПоСтудентамToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ПриказыПоСтудентамToolStripMenuItem.Text = "Приказы по студентам"
        '
        'ПриказыПоАХДToolStripMenuItem
        '
        Me.ПриказыПоАХДToolStripMenuItem.Name = "ПриказыПоАХДToolStripMenuItem"
        Me.ПриказыПоАХДToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ПриказыПоАХДToolStripMenuItem.Text = "Приказы по АХД"
        '
        'ДокладныеИСлужебныеЗапискиToolStripMenuItem
        '
        Me.ДокладныеИСлужебныеЗапискиToolStripMenuItem.Name = "ДокладныеИСлужебныеЗапискиToolStripMenuItem"
        Me.ДокладныеИСлужебныеЗапискиToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ДокладныеИСлужебныеЗапискиToolStripMenuItem.Text = "Докладные и служебные записки"
        '
        'АктыToolStripMenuItem
        '
        Me.АктыToolStripMenuItem.Name = "АктыToolStripMenuItem"
        Me.АктыToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.АктыToolStripMenuItem.Text = "Акты"
        '
        'РаспоряженияToolStripMenuItem
        '
        Me.РаспоряженияToolStripMenuItem.Name = "РаспоряженияToolStripMenuItem"
        Me.РаспоряженияToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.РаспоряженияToolStripMenuItem.Text = "Распоряжения"
        '
        'ДоверенностиToolStripMenuItem
        '
        Me.ДоверенностиToolStripMenuItem.Name = "ДоверенностиToolStripMenuItem"
        Me.ДоверенностиToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ДоверенностиToolStripMenuItem.Text = "Доверенности"
        '
        'ОкнаToolStripMenuItem
        '
        Me.ОкнаToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.КаскадомToolStripMenuItem, Me.СлеваНаправоToolStripMenuItem, Me.СверхуВнизToolStripMenuItem, Me.ToolStripMenuItem2, Me.ЗакрытьВсеToolStripMenuItem})
        Me.ОкнаToolStripMenuItem.Name = "ОкнаToolStripMenuItem"
        Me.ОкнаToolStripMenuItem.Size = New System.Drawing.Size(47, 20)
        Me.ОкнаToolStripMenuItem.Text = "Окна"
        '
        'КаскадомToolStripMenuItem
        '
        Me.КаскадомToolStripMenuItem.Name = "КаскадомToolStripMenuItem"
        Me.КаскадомToolStripMenuItem.Size = New System.Drawing.Size(156, 22)
        Me.КаскадомToolStripMenuItem.Text = "Каскадом"
        '
        'СлеваНаправоToolStripMenuItem
        '
        Me.СлеваНаправоToolStripMenuItem.Name = "СлеваНаправоToolStripMenuItem"
        Me.СлеваНаправоToolStripMenuItem.Size = New System.Drawing.Size(156, 22)
        Me.СлеваНаправоToolStripMenuItem.Text = "Слева направо"
        '
        'СверхуВнизToolStripMenuItem
        '
        Me.СверхуВнизToolStripMenuItem.Name = "СверхуВнизToolStripMenuItem"
        Me.СверхуВнизToolStripMenuItem.Size = New System.Drawing.Size(156, 22)
        Me.СверхуВнизToolStripMenuItem.Text = "Сверху вниз"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(153, 6)
        '
        'ЗакрытьВсеToolStripMenuItem
        '
        Me.ЗакрытьВсеToolStripMenuItem.Name = "ЗакрытьВсеToolStripMenuItem"
        Me.ЗакрытьВсеToolStripMenuItem.Size = New System.Drawing.Size(156, 22)
        Me.ЗакрытьВсеToolStripMenuItem.Text = "Закрыть все"
        '
        'СправкаToolStripMenuItem
        '
        Me.СправкаToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОПрограммеToolStripMenuItem, Me.ОАвтореToolStripMenuItem, Me.ЛицензионноеСоглашениеToolStripMenuItem, Me.НаписатьПисьмоРазработчикуToolStripMenuItem})
        Me.СправкаToolStripMenuItem.Name = "СправкаToolStripMenuItem"
        Me.СправкаToolStripMenuItem.Size = New System.Drawing.Size(65, 20)
        Me.СправкаToolStripMenuItem.Text = "Справка"
        '
        'ОПрограммеToolStripMenuItem
        '
        Me.ОПрограммеToolStripMenuItem.Name = "ОПрограммеToolStripMenuItem"
        Me.ОПрограммеToolStripMenuItem.Size = New System.Drawing.Size(250, 22)
        Me.ОПрограммеToolStripMenuItem.Text = "О программе"
        '
        'ОАвтореToolStripMenuItem
        '
        Me.ОАвтореToolStripMenuItem.Name = "ОАвтореToolStripMenuItem"
        Me.ОАвтореToolStripMenuItem.Size = New System.Drawing.Size(250, 22)
        Me.ОАвтореToolStripMenuItem.Text = "О авторе"
        '
        'ЛицензионноеСоглашениеToolStripMenuItem
        '
        Me.ЛицензионноеСоглашениеToolStripMenuItem.Name = "ЛицензионноеСоглашениеToolStripMenuItem"
        Me.ЛицензионноеСоглашениеToolStripMenuItem.Size = New System.Drawing.Size(250, 22)
        Me.ЛицензионноеСоглашениеToolStripMenuItem.Text = "Лицензионное соглашение"
        '
        'НаписатьПисьмоРазработчикуToolStripMenuItem
        '
        Me.НаписатьПисьмоРазработчикуToolStripMenuItem.Name = "НаписатьПисьмоРазработчикуToolStripMenuItem"
        Me.НаписатьПисьмоРазработчикуToolStripMenuItem.Size = New System.Drawing.Size(250, 22)
        Me.НаписатьПисьмоРазработчикуToolStripMenuItem.Text = "Написать письмо разработчику"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1, Me.ToolStripStatusLabel2, Me.ToolStripStatusLabel3, Me.ToolStripStatusLabel4, Me.ToolStripStatusLabel5, Me.ToolStripStatusLabel6})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 431)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(632, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(66, 17)
        Me.ToolStripStatusLabel1.Text = "Состояние"
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(32, 17)
        Me.ToolStripStatusLabel2.Text = "Дата"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(12, 17)
        Me.ToolStripStatusLabel3.Text = "-"
        '
        'ToolStripStatusLabel4
        '
        Me.ToolStripStatusLabel4.Name = "ToolStripStatusLabel4"
        Me.ToolStripStatusLabel4.Size = New System.Drawing.Size(42, 17)
        Me.ToolStripStatusLabel4.Text = "Время"
        '
        'ToolStripStatusLabel5
        '
        Me.ToolStripStatusLabel5.Name = "ToolStripStatusLabel5"
        Me.ToolStripStatusLabel5.Size = New System.Drawing.Size(12, 17)
        Me.ToolStripStatusLabel5.Text = "-"
        '
        'ToolStripStatusLabel6
        '
        Me.ToolStripStatusLabel6.Name = "ToolStripStatusLabel6"
        Me.ToolStripStatusLabel6.Size = New System.Drawing.Size(200, 17)
        Me.ToolStripStatusLabel6.Text = "Авторизованы права пользователя"
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'MDIParent1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(632, 453)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MDIParent1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Учет документов"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Timer1 As Timer
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ФайлToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents АвторизоватьсяToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents ВыходToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВидToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СтрокаСостоянияToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents БДToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents АрхивированиеБДToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВосстановлениеРезервнойКопииБДToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВосстановлениеВИсходныйВидБДToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ИнструментыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents МастерПостроенияОтчетовToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПоставитьНаКонтрольToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents НастройкиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЖурналСобытийToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents КнигиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ИсходящиеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВходящиеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДоговорыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПриказыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПриказыПоСтудентамToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПриказыПоАХДToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДокладныеИСлужебныеЗапискиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents АктыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents РаспоряженияToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДоверенностиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОкнаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents КаскадомToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СлеваНаправоToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СверхуВнизToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As ToolStripSeparator
    Friend WithEvents ЗакрытьВсеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СправкаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОПрограммеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОАвтореToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЛицензионноеСоглашениеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents НаписатьПисьмоРазработчикуToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents ToolStripProgressBar1 As ToolStripProgressBar
    Friend WithEvents ToolStripStatusLabel2 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel3 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel4 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel5 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel6 As ToolStripStatusLabel
    Friend WithEvents Timer2 As Timer
    Friend WithEvents ToolTip1 As ToolTip
End Class
