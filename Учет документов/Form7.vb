Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Form7
    Sub Dobavlenie_zap()
        lvwColumnSorter = New ListViewColumnSorter()
        Me.ListView1.ListViewItemSorter = lvwColumnSorter
        Try
            Dim Command As New OleDbCommand("Insert Into [Приказы] ([Регистрационный номер], [Дата регистрации], [Краткое содержание], [Количество листов], [Исполнитель], [Куда передан приказ], [Отметка о помещении в дело]) values ( '" & p1 & "', '" & p2 & "', '" & p3 & "', '" & p4 & "','" & p5 & "', '" & p6 & "', '" & p7 & "')", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            ListView1.Items.Add("[Peгиcтpaциoнный номер]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Дaтa регистрации]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Kpaткoe содержание]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Koличecтвo листов]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Исполнитель]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Kyдa передан приказ]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Oтмeткa о помещении в дело]")
            Form7_Load(Me, New EventArgs)
            For Each columnheader In Me.ListView1.Columns
                columnheader.Width = -2
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub Editor_zap()
        Try
            Dim Command As New OleDbCommand("UPDATE [Приказы] SET [Регистрационный номер]='" & p1 & "', [Дата регистрации]='" & p2 & "', [Краткое содержание]='" & p3 & "', [Количество листов]='" & p4 & "', [Исполнитель]='" & p5 & "', [Куда передан приказ]='" & p6 & "', [Отметка о помещении в дело]='" & p7 & "' WHERE ([Регистрационный номер] Like '" & ListView1.SelectedItems.Item(0).Text & "')", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            ListView1.SelectedItems.Item(0).Text = p1
            ListView1.SelectedItems.Item(0).SubItems.Item(1).Text = p2
            ListView1.SelectedItems.Item(0).SubItems.Item(2).Text = p3
            ListView1.SelectedItems.Item(0).SubItems.Item(3).Text = p4
            ListView1.SelectedItems.Item(0).SubItems.Item(4).Text = p5
            ListView1.SelectedItems.Item(0).SubItems.Item(5).Text = p6
            ListView1.SelectedItems.Item(0).SubItems.Item(6).Text = p7
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub Удалить()
        Dim МВох As DialogResult = MessageBox.Show("Удaлить текущую запись", "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        Try
            If МВох = DialogResult.No Then Exit Sub
            If МВох = DialogResult.Yes Then
                Dim Command As New OleDbCommand("DELETE FROM [Приказы] WHERE [Регистрационный номер]='" & ListView1.SelectedItems.Item(0).Text & "'", Connector)
                Connector.Open()
                Command.ExecuteNonQuery()
                Connector.Close()
                ListView1.SelectedItems.Item(0).Remove()
                Form7_Load(Me, New EventArgs)
                MessageBox.Show("Запись была успешна удалена из БД", "Операция успешна завершина", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MDIParent1
        Prava()
        lvwColumnSorter = New ListViewColumnSorter()
        Me.ListView1.ListViewItemSorter = lvwColumnSorter
        ListView1.Columns.Clear()
        ListView1.Columns.Add("Peгиcтpaциoнный номер")
        ListView1.Columns.Add("Дaтa регистрации")
        ListView1.Columns.Add("Kpaткoe содержание")
        ListView1.Columns.Add("Koличecтвo листов")
        ListView1.Columns.Add("Иcпoлнитeль")
        ListView1.Columns.Add("Kyдa передан приказ")
        ListView1.Columns.Add("Oтмeткa о помещении в дело")
        ListView1.Items.Clear()
        Try
            Dim Poisk As String
            If ToolStripTextBox1.Text = "" Then
                Poisk = "SELECT * FROM Приказы ORDER BY Приказы.[Код] DESC"
            Else
                Poisk = "SELECT * FROM Приказы WHERE ((Приказы.[" & ToolStripComboBox1.Text & "]) Like '%" & ToolStripTextBox1.Text & "%')"
            End If
            Dim DataReader As OleDbDataReader
            Dim Command As New OleDbCommand(Poisk, Connector)
            Connector.Open()
            DataReader = Command.ExecuteReader
            While DataReader.Read() = True
                ListView1.Items.Add(DataReader.GetValue(0))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(1))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(2))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(3))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(4))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(5))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(6))
                CountForm4 = ListView1.Items.Count
            End While
            DataReader.Close()
            Connector.Close()
            ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
            For Each columnheader In Me.ListView1.Columns
                columnheader.Width = -2
            Next
        Catch ex As Exception
            Connector.Close()
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ОбновитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбновитьToolStripMenuItem.Click, ToolStripMenuItem3.Click
        Form7_Load(Me, New EventArgs)
    End Sub

    Private Sub УбратьКолонкуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УбратьКолонкуToolStripMenuItem.Click, ToolStripMenuItem5.Click
        Dim Kolonka As Integer
        Try
            Kolonka = InputBox("Введите номер колонки которую нужно скрыть", "Запрос")
            ListView1.Columns(Kolonka - 1).Dispose()
        Catch ех As Exception
            MessageBox.Show("Cлeдyeт вводить только числа", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ПоискПоНомеруToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоискПоНомеруToolStripMenuItem.Click, ToolStripMenuItem6.Click
        Dim Nomer As String
        Try
            Nomer = InputBox("Введите номер для поиска", "Запрос")
            lvwColumnSorter = New ListViewColumnSorter()
            Me.ListView1.ListViewItemSorter = lvwColumnSorter
            ListView1.Columns.Clear()
            ListView1.Columns.Add("№ n/n")
            ListView1.Columns.Add("№ и дата договора")
            ListView1.Columns.Add("Пocтaвщик, исполнитель, подрядчик")
            ListView1.Columns.Add("Пpeдмeт договора")
            ListView1.Columns.Add("Cyммa договора")
            ListView1.Columns.Add("Cpoк исполнения обязательств")
            ListView1.Columns.Add("Koличecтвo листов")
            ListView1.Columns.Add("Haличиe приложений")
            ListView1.Columns.Add("Инициaтop (ответственный)")
            ListView1.Columns.Add("Пpимeчaниe")
            ListView1.Items.Clear()
            Dim DataReader As OleDbDataReader
            Dim Command As New OleDbCommand("SELECT * FROM Приказы WHERE ((([Регистрационный номер]) Like '%" & Nomer & "%'));", Connector)
            Connector.Open()
            DataReader = Command.ExecuteReader
            While DataReader.Read() = True
                ListView1.Items.Add(DataReader.GetValue(0))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(1))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(2))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(3))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(4))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(5))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(6))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(7))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(8))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(9))
            End While
            DataReader.Close()
            Connector.Close()
            ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
            For Each columnheader In Me.ListView1.Columns
                columnheader.Width = -2
            Next
        Catch ex As Exception
            MessageBox.Show("Cлeдyeт вводить только числа", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ПоискПоДатеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоискПоДатеToolStripMenuItem.Click, ToolStripMenuItem7.Click
        Dim Nomer As String
        Try
            Nomer = InputBox("Введите дату поиска", "Запрос")
            lvwColumnSorter = New ListViewColumnSorter()
            Me.ListView1.ListViewItemSorter = lvwColumnSorter
            ListView1.Columns.Clear()
            ListView1.Columns.Add("№ n/n")
            ListView1.Columns.Add("№ и дата договора")
            ListView1.Columns.Add("Пocтaвщик, исполнитель, подрядчик")
            ListView1.Columns.Add("Пpeдмeт договора")
            ListView1.Columns.Add("Cyммa договора)")
            ListView1.Columns.Add("Cpoк исполнения обязательств")
            ListView1.Columns.Add("Koличecтвo листов")
            ListView1.Columns.Add("Haличиe приложений")
            ListView1.Columns.Add("Инициaтop (ответственный)")
            ListView1.Columns.Add("Пpимeчaниe")
            ListView1.Items.Clear()
            Dim DataReader As OleDbDataReader
            Dim Command As New OleDbCommand("SELECT * FROM Приказы WHERE ((([Дата регистрации]) Like '%" & Nomer & "%'));", Connector)
            Connector.Open()
            DataReader = Command.ExecuteReader
            While DataReader.Read() = True
                ListView1.Items.Add(DataReader.GetValue(0))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(1))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(2))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(3))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(4))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(5))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(6))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(7))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(8))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(9))
            End While
            DataReader.Close()
            Connector.Close()
            ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
            For Each columnheader In Me.ListView1.Columns
                columnheader.Width = -2
            Next
        Catch ех As Exception
            MessageBox.Show("Cлeдyeт вводить только числа", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ВыключитьФильтрToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыключитьФильтрToolStripMenuItem.Click, ToolStripMenuItem8.Click
        Form7_Load(Me, New EventArgs)
    End Sub

    Private Sub ЭкспортWordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортWordToolStripMenuItem.Click, ToolStripMenuItem9.Click
        Try
            Dim Дата As String = Format(Now, "d ММММ уууу")
            Dim W = New Word.Application
            W.VisiЫe = True
            W.Documents.Add()
            W.Selection.TypeText("Выписка из БД: " & Дата & Chr(13) & Chr(10))
            For i As Short = 0 To ListView1.Items.Count - 1
                W.Selection.TypeText(ListView1.Items(i).SubItems.Item(0).Text & " " & ListView1.Items(i).SubItems.Item(1).Text & " " & ListView1.Items(i).SubItems.Item(2).Text & " " & ListView1.Items(i).SubItems.Item(3).Text & " " & ListView1.Items(i).SubItems.Item(4).Text & "" & ListView1.Items(i).SubItems.Item(5).Text & "" & ListView1.Items(i).SubItems.Item(6).Text & " " & ListView1.Items(i).SubItems.Item(7).Text & " " & ListView1.Items(i).SubItems.Item(8).Text & Chr(13) & Chr(10))
            Next i
            W = Nothing
            OtchetGurnal = "Экспорт исходящих в WORD " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ЭкспортExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортExcelToolStripMenuItem.Click, ToolStripMenuItem10.Click
        Try
            Dim myXL As Excel.Application, myWB As Excel.Workbook, myWS As Excel.Worksheet
            Dim i As Integer
            Dim Y As Integer
            Dim z As Integer
            myXL = New Excel.Application
            myWB = myXL.Workbooks.Add
            myWS = myWB.Worksheets(1)
            z = 2
            myXL.Visible = True
            For i = 1 To Me.ListView1.Items.Count
                For У = 1 To 9
                    myWS.Cells(1, Y) = ListView1.Columns(Y - 1).Text
                Next У
                myWS.Cells(z, 1) = ListView1.Items.Item(i - 1).SubItems.Item(0).Text
                myWS.Cells(z, 2) = ListView1.Items.Item(i - 1).SubItems.Item(1).Text
                myWS.Cells(z, 3) = ListView1.Items.Item(i - 1).SubItems.Item(2).Text
                myWS.Cells(z, 4) = ListView1.Items.Item(i - 1).SubItems.Item(3).Text
                myWS.Cells(z, 5) = ListView1.Items.Item(i - 1).SubItems.Item(4).Text
                myWS.Cells(z, 6) = ListView1.Items.Item(i - 1).SubItems.Item(5).Text
                myWS.Cells(z, 7) = ListView1.Items.Item(i - 1).SubItems.Item(6).Text
                myWS.Cells(z, 8) = ListView1.Items.Item(i - 1).SubItems.Item(7).Text
                myWS.Cells(z, 9) = ListView1.Items.Item(i - 1).SubItems.Item(8).Text
                z = z + 1
            Next i
            myWS.Columns(1).ColumnWidth = 20
            myWS.Columns(2).ColumnWidth = 20
            myWS.Columns(3).ColumnWidth = 20
            myWS.Columns(4).ColumnWidth = 50
            myWS.Columns(5).ColumnWidth = 20
            myWS.Columns(6).ColumnWidth = 20
            myWS.Columns(7).ColumnWidth = 20
            myWS.Columns(8).ColumnWidth = 20
            myWS.Columns(9).ColumnWidth = 20
            myXL = Nothing
            myWB = Nothing
            myWS = Nothing
            OtchetGurnal = "Экспорт исходящих в Excel " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click, ToolStripMenuItem11.Click
        Close()
    End Sub

    Private Sub УдалитьЗаписьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьЗаписьToolStripMenuItem.Click
        Удалить()
    End Sub

    Private Sub ContextMenuStrip2_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip2.Opening
        Form7_Load(Me, New EventArgs)
    End Sub

    Private Sub ОбновитьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ОбновитьToolStripMenuItem1.Click
        Form7_Load(Me, New EventArgs)
    End Sub

    Private Sub ДобавитьЗаписьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДобавитьЗаписьToolStripMenuItem.Click
        Form8.Close()
        Form8.Text = "Добавление записи"
        Form8.Show()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        ДобавитьЗаписьToolStripMenuItem_Click(Me, New EventArgs)
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Form7_Load(Me, New EventArgs)
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        ToolStripTextBox1.Text = ""
        ToolStripComboBox1.Text = ToolStripComboBox1.Items(0)
        Form7_Load(Me, New EventArgs)
    End Sub

    Private Sub ToolStripTextBoxl_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            If ToolStripComboBox1.Text = "" Then
                MessageBox.Show("Heвыбpaнo поля для фильтра", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                Form7_Load(Me, New EventArgs)
            End If
        End If
    End Sub
    Private Sub ToolStripComboBoxl_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Form7_Load(Me, New EventArgs)
        End If
    End Sub

    Private Sub ListViewl_ColumnClick(sender As Object, e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
        If (e.Column = lvwColumnSorter.SortColumn) Then
            If (lvwColumnSorter.Order = SortOrder.Ascending) Then
                lvwColumnSorter.Order = SortOrder.Descending
            Else
                lvwColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            lvwColumnSorter.SortColumn = e.Column
            lvwColumnSorter.Order = SortOrder.Ascending
        End If
        Me.ListView1.Sort()
    End Sub
    Private Sub ListViewl_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseUp
        Dim ItemIndex As Integer
        If e.Button = MouseButtons.Right Then
            ItemIndex = ListView1.SelectedIndices.Count
            If ItemIndex = 1 Then
                ContextMenuStrip1.Show(Location.X + e.X, Location.Y + e.Y + ContextMenuStrip1.Height)
            Else
                ContextMenuStrip2.Show(Location.X + e.X, Location.Y + e.Y + ContextMenuStrip1.Height)
            End If
        End If
    End Sub

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        Try
            p1 = ListView1.SelectedItems.Item(0).Text
            p2 = ListView1.SelectedItems.Item(0).SubItems.Item(1).Text
            p3 = ListView1.SelectedItems.Item(0).SubItems.Item(2).Text
            p4 = ListView1.SelectedItems.Item(0).SubItems.Item(3).Text
            p5 = ListView1.SelectedItems.Item(0).SubItems.Item(4).Text
            p6 = ListView1.SelectedItems.Item(0).SubItems.Item(5).Text
            p7 = ListView1.SelectedItems.Item(0).SubItems.Item(6).Text
            p8 = ListView1.SelectedItems.Item(0).SubItems.Item(7).Text
            p9 = ListView1.SelectedItems.Item(0).SubItems.Item(8).Text
            p10 = ListView1.SelectedItems.Item(0).SubItems.Item(9).Text
            Form8.Close()
            Form8.Text = "Редактирование записи"
            Form8.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class