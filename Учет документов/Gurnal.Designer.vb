<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Gurnal
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Gurnal))
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripTextBox1 = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStripButton4 = New System.Windows.Forms.ToolStripButton()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.МенюИсToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОбновитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЭкспортWordToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЭкспортExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ВыходToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListView1
        '
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(0, 25)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(784, 536)
        Me.ListView1.TabIndex = 16
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2, Me.ToolStripButton3, Me.ToolStripTextBox1, Me.ToolStripButton4})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(784, 25)
        Me.ToolStrip1.TabIndex = 14
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = Global.Учет_документов.My.Resources.Resources.Word
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "ToolStripButton1"
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = Global.Учет_документов.My.Resources.Resources.Excel
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "ToolStripButton2"
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton3.Image = Global.Учет_документов.My.Resources.Resources.Обновить
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton3.Text = "ToolStripButton3"
        '
        'ToolStripTextBox1
        '
        Me.ToolStripTextBox1.Name = "ToolStripTextBox1"
        Me.ToolStripTextBox1.Size = New System.Drawing.Size(100, 25)
        '
        'ToolStripButton4
        '
        Me.ToolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton4.Image = Global.Учет_документов.My.Resources.Resources.закрыть
        Me.ToolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton4.Name = "ToolStripButton4"
        Me.ToolStripButton4.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton4.Text = "ToolStripButton4"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.МенюИсToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(9, 1)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(128, 24)
        Me.MenuStrip1.TabIndex = 15
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'МенюИсToolStripMenuItem
        '
        Me.МенюИсToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОбновитьToolStripMenuItem, Me.ЭкспортWordToolStripMenuItem, Me.ЭкспортExcelToolStripMenuItem, Me.ToolStripMenuItem1, Me.ВыходToolStripMenuItem})
        Me.МенюИсToolStripMenuItem.Name = "МенюИсToolStripMenuItem"
        Me.МенюИсToolStripMenuItem.Size = New System.Drawing.Size(120, 20)
        Me.МенюИсToolStripMenuItem.Text = "Меню Исходящих"
        '
        'ОбновитьToolStripMenuItem
        '
        Me.ОбновитьToolStripMenuItem.Name = "ОбновитьToolStripMenuItem"
        Me.ОбновитьToolStripMenuItem.ShortcutKeyDisplayString = "F5"
        Me.ОбновитьToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.ОбновитьToolStripMenuItem.Text = "Обновить"
        '
        'ЭкспортWordToolStripMenuItem
        '
        Me.ЭкспортWordToolStripMenuItem.Image = CType(resources.GetObject("ЭкспортWordToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ЭкспортWordToolStripMenuItem.Name = "ЭкспортWordToolStripMenuItem"
        Me.ЭкспортWordToolStripMenuItem.ShortcutKeyDisplayString = "Ctrl+W"
        Me.ЭкспортWordToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.ЭкспортWordToolStripMenuItem.Text = "Экспорт Word"
        '
        'ЭкспортExcelToolStripMenuItem
        '
        Me.ЭкспортExcelToolStripMenuItem.Image = CType(resources.GetObject("ЭкспортExcelToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ЭкспортExcelToolStripMenuItem.Name = "ЭкспортExcelToolStripMenuItem"
        Me.ЭкспортExcelToolStripMenuItem.ShortcutKeyDisplayString = "Ctrl+E"
        Me.ЭкспортExcelToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.ЭкспортExcelToolStripMenuItem.Text = "Экспорт Excel"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(193, 6)
        '
        'ВыходToolStripMenuItem
        '
        Me.ВыходToolStripMenuItem.Name = "ВыходToolStripMenuItem"
        Me.ВыходToolStripMenuItem.Size = New System.Drawing.Size(196, 22)
        Me.ВыходToolStripMenuItem.Text = "Выход"
        '
        'Gurnal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Gurnal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Журнал событий"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListView1 As ListView
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents ToolStripButton2 As ToolStripButton
    Friend WithEvents ToolStripButton3 As ToolStripButton
    Friend WithEvents ToolStripTextBox1 As ToolStripTextBox
    Friend WithEvents ToolStripButton4 As ToolStripButton
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents МенюИсToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОбновитьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЭкспортWordToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЭкспортExcelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents ВыходToolStripMenuItem As ToolStripMenuItem
End Class
