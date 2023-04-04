<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form17
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form17))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.МенюИсходящихToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОбновитьToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ФильтрToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УбратьКолонкуToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоискПоНомеруToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоискПоДатеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВыключитьФильтрToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЭкспортWordToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЭкспортExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ВыходToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ОбновитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДобавитьЗаписьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.УдалитьЗаписьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripTextBox1 = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripComboBox1 = New System.Windows.Forms.ToolStripComboBox()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.MenuStrip1.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.МенюИсходящихToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(9, 1)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(145, 24)
        Me.MenuStrip1.TabIndex = 12
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'МенюИсходящихToolStripMenuItem
        '
        Me.МенюИсходящихToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОбновитьToolStripMenuItem1, Me.ФильтрToolStripMenuItem, Me.ЭкспортWordToolStripMenuItem, Me.ЭкспортExcelToolStripMenuItem, Me.ToolStripMenuItem1, Me.ВыходToolStripMenuItem})
        Me.МенюИсходящихToolStripMenuItem.Name = "МенюИсходящихToolStripMenuItem"
        Me.МенюИсходящихToolStripMenuItem.Size = New System.Drawing.Size(137, 20)
        Me.МенюИсходящихToolStripMenuItem.Text = "Меню Распоряжения"
        '
        'ОбновитьToolStripMenuItem1
        '
        Me.ОбновитьToolStripMenuItem1.Name = "ОбновитьToolStripMenuItem1"
        Me.ОбновитьToolStripMenuItem1.ShortcutKeys = System.Windows.Forms.Keys.F5
        Me.ОбновитьToolStripMenuItem1.Size = New System.Drawing.Size(151, 22)
        Me.ОбновитьToolStripMenuItem1.Text = "Обновить"
        '
        'ФильтрToolStripMenuItem
        '
        Me.ФильтрToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.УбратьКолонкуToolStripMenuItem, Me.ПоискПоНомеруToolStripMenuItem, Me.ВыключитьФильтрToolStripMenuItem})
        Me.ФильтрToolStripMenuItem.Name = "ФильтрToolStripMenuItem"
        Me.ФильтрToolStripMenuItem.Size = New System.Drawing.Size(151, 22)
        Me.ФильтрToolStripMenuItem.Text = "Фильтр"
        '
        'УбратьКолонкуToolStripMenuItem
        '
        Me.УбратьКолонкуToolStripMenuItem.Name = "УбратьКолонкуToolStripMenuItem"
        Me.УбратьКолонкуToolStripMenuItem.Size = New System.Drawing.Size(211, 22)
        Me.УбратьКолонкуToolStripMenuItem.Text = "Убрать колонку"
        '
        'ПоискПоНомеруToolStripMenuItem
        '
        Me.ПоискПоНомеруToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПоискПоДатеToolStripMenuItem})
        Me.ПоискПоНомеруToolStripMenuItem.Image = Global.Учет_документов.My.Resources.Resources.Поиск
        Me.ПоискПоНомеруToolStripMenuItem.Name = "ПоискПоНомеруToolStripMenuItem"
        Me.ПоискПоНомеруToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.ПоискПоНомеруToolStripMenuItem.Size = New System.Drawing.Size(211, 22)
        Me.ПоискПоНомеруToolStripMenuItem.Text = "Поиск по номеру"
        '
        'ПоискПоДатеToolStripMenuItem
        '
        Me.ПоискПоДатеToolStripMenuItem.Name = "ПоискПоДатеToolStripMenuItem"
        Me.ПоискПоДатеToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ПоискПоДатеToolStripMenuItem.Text = "Поиск по дате"
        '
        'ВыключитьФильтрToolStripMenuItem
        '
        Me.ВыключитьФильтрToolStripMenuItem.Name = "ВыключитьФильтрToolStripMenuItem"
        Me.ВыключитьФильтрToolStripMenuItem.Size = New System.Drawing.Size(211, 22)
        Me.ВыключитьФильтрToolStripMenuItem.Text = "Выключить фильтр"
        '
        'ЭкспортWordToolStripMenuItem
        '
        Me.ЭкспортWordToolStripMenuItem.Image = Global.Учет_документов.My.Resources.Resources.Word
        Me.ЭкспортWordToolStripMenuItem.Name = "ЭкспортWordToolStripMenuItem"
        Me.ЭкспортWordToolStripMenuItem.Size = New System.Drawing.Size(151, 22)
        Me.ЭкспортWordToolStripMenuItem.Text = "Экспорт Word"
        '
        'ЭкспортExcelToolStripMenuItem
        '
        Me.ЭкспортExcelToolStripMenuItem.Image = Global.Учет_документов.My.Resources.Resources.Excel
        Me.ЭкспортExcelToolStripMenuItem.Name = "ЭкспортExcelToolStripMenuItem"
        Me.ЭкспортExcelToolStripMenuItem.Size = New System.Drawing.Size(151, 22)
        Me.ЭкспортExcelToolStripMenuItem.Text = "Экспорт Excel"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(148, 6)
        '
        'ВыходToolStripMenuItem
        '
        Me.ВыходToolStripMenuItem.Name = "ВыходToolStripMenuItem"
        Me.ВыходToolStripMenuItem.Size = New System.Drawing.Size(151, 22)
        Me.ВыходToolStripMenuItem.Text = "Выход"
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОбновитьToolStripMenuItem, Me.ДобавитьЗаписьToolStripMenuItem})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(167, 48)
        '
        'ОбновитьToolStripMenuItem
        '
        Me.ОбновитьToolStripMenuItem.Image = Global.Учет_документов.My.Resources.Resources.Обновить
        Me.ОбновитьToolStripMenuItem.Name = "ОбновитьToolStripMenuItem"
        Me.ОбновитьToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.ОбновитьToolStripMenuItem.Text = "Обновить"
        '
        'ДобавитьЗаписьToolStripMenuItem
        '
        Me.ДобавитьЗаписьToolStripMenuItem.Image = Global.Учет_документов.My.Resources.Resources.Добавить
        Me.ДобавитьЗаписьToolStripMenuItem.Name = "ДобавитьЗаписьToolStripMenuItem"
        Me.ДобавитьЗаписьToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.ДобавитьЗаписьToolStripMenuItem.Text = "Добавить запись"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.УдалитьЗаписьToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(159, 26)
        '
        'УдалитьЗаписьToolStripMenuItem
        '
        Me.УдалитьЗаписьToolStripMenuItem.Image = Global.Учет_документов.My.Resources.Resources.закрыть
        Me.УдалитьЗаписьToolStripMenuItem.Name = "УдалитьЗаписьToolStripMenuItem"
        Me.УдалитьЗаписьToolStripMenuItem.Size = New System.Drawing.Size(158, 22)
        Me.УдалитьЗаписьToolStripMenuItem.Text = "Удалить запись"
        '
        'ToolStripTextBox1
        '
        Me.ToolStripTextBox1.Name = "ToolStripTextBox1"
        Me.ToolStripTextBox1.Size = New System.Drawing.Size(100, 25)
        Me.ToolStripTextBox1.ToolTipText = "Текст для фильтра"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2, Me.ToolStripComboBox1, Me.ToolStripTextBox1, Me.ToolStripButton3})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(870, 25)
        Me.ToolStrip1.TabIndex = 13
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = Global.Учет_документов.My.Resources.Resources.Добавить
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "ToolStripButton1"
        Me.ToolStripButton1.ToolTipText = "Добавить запись в БД"
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = Global.Учет_документов.My.Resources.Resources.Обновить
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "ToolStripButton2"
        Me.ToolStripButton2.ToolTipText = "Обновить таблицу"
        '
        'ToolStripComboBox1
        '
        Me.ToolStripComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ToolStripComboBox1.Items.AddRange(New Object() {"Регистрационный номер", "Дата регистрации", "Краткое содержание", "Количество листов", "Исполнитель", "Куда передан приказ", "Отметка о помещении в дело"})
        Me.ToolStripComboBox1.Name = "ToolStripComboBox1"
        Me.ToolStripComboBox1.Size = New System.Drawing.Size(121, 25)
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton3.Image = Global.Учет_документов.My.Resources.Resources.закрыть
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton3.Text = "ToolStripButton3"
        Me.ToolStripButton3.ToolTipText = "Удалить фильтр"
        '
        'ListView1
        '
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(0, 25)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(870, 442)
        Me.ListView1.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.ListView1.TabIndex = 14
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'Form17
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(870, 467)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form17"
        Me.Text = "Распоряжения"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents МенюИсходящихToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОбновитьToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ФильтрToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УбратьКолонкуToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПоискПоНомеруToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПоискПоДатеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВыключитьФильтрToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЭкспортWordToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЭкспортExcelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents ВыходToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents ОбновитьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДобавитьЗаписьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents УдалитьЗаписьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripTextBox1 As ToolStripTextBox
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents ToolStripButton2 As ToolStripButton
    Friend WithEvents ToolStripComboBox1 As ToolStripComboBox
    Friend WithEvents ToolStripButton3 As ToolStripButton
    Friend WithEvents ListView1 As ListView
End Class
