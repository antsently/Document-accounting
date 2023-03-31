﻿Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Form3

    Sub Dobavlenie_zap()
        lvwColumnSorter = New ListViewColumnSorter()
        Me.ListView1.ListViewItemSorter = lvwColumnSorter
        Try
            Dim Command As New OleDbCommand("Insert Into [Входящая] ([Дата поступления], [Регистрационный номер], [Корреспондент документа], [Исходящий номер и дата документа], [Заголовок (краткое содержание документа)], [Количество листов с приложением], [Резолюция или кому направлен документ], [Роспись исполнителя в получении], [Срок исполнения], [Отметка о постановке на контроль], [Отметка об исполнении], 
[Примечание]) values ( '" & p1 & "', '" & p2 & "', '" & p3 & "', '" & p4 & "', '" & p5 & "', '" & p6 & "', '" & p7 & "', '" & p8 & "', '" & p9 & "', '" & p10 & "', '" & p11 & "', '" & p12 & "') ", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            ListView1.Items.Add("[Дaтa поступления]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Регистрационный номер]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Koppecnoндeнт документа]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Исходящий номер и дата документа]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Зaгoлoвoк (краткое содержание документа)]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Koличecтвo листов с приложением]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Резолюция или кому направлен документ]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Исnолнитель]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Cpoк исполнения]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Oтмeткa о постановке на контроль]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("[Oтмeткa об исполнении]")
            ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add("Примечание")
            Form3_Load(Me, New EventArgs)
            For Each columnheader In Me.ListView1.Columns
                columnheader.Width = -3
            Next
            OtchetGurnal = "Загрузка Входящая " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Sub Editor_zap()
        Try
            Dim Command As New OleDbCommand("UPDATE [Входящая] SET [Дата поступления]='" & p1 & "', [Регистрационный номер]='" & p2 & "', [Корреспондент документа]='" & p3 & "', [Исходящий номер и дата документа]='" & p4 & "', [Заголовок (краткое содержание документа)]='" & p5 & "', [Количество листов с приложением]='" & p6 & "', [Резолюция или кому направлен документ]='" & p7 & "', [Роспись исполнителя в получении]='" & p8 & "', [Срок исполнения]='" & p9 & "', [Отметка о постановке на контроль]='" & p10 & "', [Отметка об исполнении]='" & p11 & "',[Примечание] ='" & p12 & "' WHERE ([Регистрационный номер] Like '" & ListView1.SelectedItems.Item(0).SubItems.Item(1).Text & "')", Connector)
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
            ListView1.SelectedItems.Item(0).SubItems.Item(7).Text = p8
            ListView1.SelectedItems.Item(0).SubItems.Item(8).Text = p9
            ListView1.SelectedItems.Item(0).SubItems.Item(9).Text = p10
            ListView1.SelectedItems.Item(0).SubItems.Item(10).Text = p11
            ListView1.SelectedItems.Item(0).SubItems.Item(11).Text = p12
            Form3_Load(Me, New EventArgs)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub Удалить()
        Dim МВох As DialogResult = MessageBox.Show("Yдaлить текущую запись", "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        Try
            If МВох = DialogResult.No Then Exit Sub
            If МВох = DialogResult.Yes Then
                Dim Command As New OleDbCommand("DELETE FROM [Входящая] WHERE [Регистрационный номер]='" & ListView1.SelectedItems.Item(0).SubItems.Item(1).Text & "'", Connector)
                Connector.Open()
                Command.ExecuteNonQuery()
                Connector.Close()
                ListView1.SelectedItems.Item(0).Remove()
                Form3_Load(Me, New EventArgs)
                MessageBox.Show("Зanиcь была успешна удалена из БД", "Операция успешна завершина", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MDIParent1
        Prava()
        lvwColumnSorter = New ListViewColumnSorter()
        Me.ListView1.ListViewItemSorter = lvwColumnSorter
        ListView1.Columns.Clear()
        ListView1.Columns.Add("Дaтa поступления")
        ListView1.Columns.Add("Peгиcтpaциoнный номер")
        ListView1.Columns.Add("Koppecnoндeнт документа")
        ListView1.Columns.Add("Иcxoдящий номер и дата документа")
        ListView1.Columns.Add("Зaгoлoвoк (краткое содержание документа)")
        ListView1.Columns.Add("Koличecтвo листов с приложением")
        ListView1.Columns.Add("Peзoлюция или кому направлен документ")
        ListView1.Columns.Add("Иcпoлнитeль")
        ListView1.Columns.Add("Cpoк исполнения")
        ListView1.Columns.Add("Oтмeткa о постановке на контроль")
        ListView1.Columns.Add("Oтмeткa об исполнении")
        ListView1.Columns.Add("Пpимeчaниe")
        ListView1.Items.Clear()
        Try
            Dim Poisk As String
            If ToolStripTextBox1.Text = "" Then
                Poisk = "SELECT * FROM Входящая ORDER BY Входящая.[Код] DESC"
            Else
                Poisk = "SELECT * FROM Входящая where ((Входящая.[" & ToolStripComboBox1.Text & "]) Like '%" & ToolStripTextBox1.Text & "%')"
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
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(7))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(8))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(9))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(10))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(11))
                CountForm2 = ListView1.Items.Count
            End While
            DataReader.Close()
            Connector.Close()
            ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
            For Each columnheader In Me.ListView1.Columns
                columnheader.Width = -3
            Next
        Catch ex As Exception
            Connector.Close()
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub ОбновитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбновитьToolStripMenuItem.Click
        Form3_Load(Me, New EventArgs)
    End Sub

    Private Sub ОбновитьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ОбновитьToolStripMenuItem1.Click
        Form3_Load(Me, New EventArgs)
    End Sub

    Private Sub ДобавитьЗаписьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДобавитьЗаписьToolStripMenuItem.Click
        Form4.Close()
        Form4.Text = "Добавление записи"
        Form4.Show()
    End Sub

    Private Sub ListView1_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView1.ColumnClick
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

    Private Sub ListView1_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView1.MouseUp
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
            p11 = ListView1.SelectedItems.Item(0).SubItems.Item(10).Text
            p12 = ListView1.SelectedItems.Item(0).SubItems.Item(11).Text
            Form4.Close()
            Form4.Text = "Редактирование записи"
            Form4.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub УдалитьЗаписьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьЗаписьToolStripMenuItem.Click
        Удалить()
    End Sub

    Private Sub ВыключитьФильтрToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыключитьФильтрToolStripMenuItem.Click
        Form3_Load(Me, New EventArgs)
    End Sub

    Private Sub УбратьКолонкуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УбратьКолонкуToolStripMenuItem.Click
        Dim Kolonka As Integer
        Try
            Kolonka = InputBox("Введите номер колонки которую нужно скрыть", "Запрос")
            ListView1.Columns(Kolonka - 1).Dispose()
        Catch ех As Exception
            MessageBox.Show("Cлeдyeт вводить только числа", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ПоискПоНомеруToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоискПоНомеруToolStripMenuItem.Click
        Dim Nomer As String
        Try
            Nomer = InputBox("Введите регистрационный номер для поиска", "Запрос")
            lvwColumnSorter = New ListViewColumnSorter()
            Me.ListView1.ListViewItemSorter = lvwColumnSorter
            ListView1.Columns.Clear()
            ListView1.Columns.Add("Дaтa поступления")
            ListView1.Columns.Add("Perиcтpaциoнный номер")
            ListView1.Columns.Add("Koppecnoндeнт документа")
            ListView1.Columns.Add("Иcxoдящий номер и дата документа")
            ListView1.Columns.Add("Зaroлoвoк (краткое содержание документа)")
            ListView1.Columns.Add("Koличecтвo листов с приложением")
            ListView1.Columns.Add("Peзoлюция или кому направлен документ")
            ListView1.Columns.Add("Иcпoлнитeль")
            ListView1.Columns.Add("Cpoк исполнения")
            ListView1.Columns.Add("Oтмeткa о постановке на контроль")
            ListView1.Columns.Add("Oтмeткa об исполнении")
            ListView1.Columns.Add("Пpимeчaниe")
            ListView1.Items.Clear()
            Dim DataReader As OleDbDataReader
            Dim Command As New OleDbCommand("SELECT * FROM Входящая WHERE ((([Регистрационный номер]) Like '%" & Nomer & "%'));", Connector)
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
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(10))
                ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(11))
            End While
            DataReader.Close()
            Connector.Close()
            ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
            For Each columnheader In Me.ListView1.Columns
                columnheader.Width = -3
            Next
        Catch ex As Exception
            MessageBox.Show("Cлeдyeт вводить только числа", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ЭкспортWordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортWordToolStripMenuItem.Click
        Try
            Dim Дата As String = Format(Now, "d ММММ уууу")
            Dim W = New Word.Application
            W.VisiЫe = True
            W.Documents.Add()
            W.Selection.TypeText("Bыnиcкa из БД: " & Дата & Chr(13) & Chr(10))
            For i As Short = 0 To ListView1.Items.Count - 1
                W.Selection.TypeText(ListView1.Items(i).SubItems.Item(0).Text & " " &
ListView1.Items(i).SubItems.Item(1).Text & " " & ListView1.Items(i).SubItems.Item(2).Text & " " & ListView1.Items(i).SubItems.Item(3).Text & " " & ListView1.Items(i).SubItems.Item(4).Text & "" & ListView1.Items(i).SubItems.Item(5).Text & "" &
ListView1.Items(i).SubItems.Item(6).Text & " " & ListView1.Items(i).SubItems.Item(7).Text & " " & ListView1.Items(i).SubItems.Item(8).Text & "" & ListView1.Items(i).SubItems.Item(9).Text & " " & ListView1.Items(i).SubItems.Item(10).Text & " " &
ListView1.Items(i).SubItems.Item(11).Text & Chr(13) & Chr(10))
            Next i
            W = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ЭкспортExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортExcelToolStripMenuItem.Click
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

    Private Sub ПолучитьКарточкуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПолучитьКарточкуToolStripMenuItem.Click
        Try
            Dim W = New Word.Application
            W.Visible = True
            W.Documents.Add()
            With W.ActiveWindow.Selection
                .Font.Color = Word.WdColor.wdColorBlack
                .Font.Name = "Times New Roman"
                .Font.Bold = False
                .Font.Size = 12
                .Tables.Add(W.Selection.Range, 8, 2, Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                Word.WdAutoFitBehavior.wdAutoFitContent)
                .Tables(1).Cell(Row:=1, Column:=1).Range.InsertAfter("Oбщий отдел" & Chr(13) & Chr(10) &
                "(внутренний документ)")
                .Tables(1).Cell(Row:=1, Column:=2).Range.InsertAfter("KOHTPOЛbHO-PEГИCTPAЦИOHHAЯ КАРТОЧКА")
                .Tables(1).Cell(Row:=2, Column:=1).Range.InsertAfter("Haимeнoвaниe вида документа")
                .Tables(1).Cell(Row:=2, Column:=2).Range.InsertAfter(ListView1.SelectedItems.Item(0).Text)
                .Tables(1).Cell(Row:=3, Column:=1).Range.InsertAfter("Kpaткoe содержание")
                .Tables(1).Cell(Row:=3,
                Column:=2).Range.InsertAfter(ListView1.SelectedItems.Item(0).SubItems.Item(3).Text)
                .Tables(1).Cell(Row:=4, Column:=1).Range.InsertAfter("Aдpecaт документа")
                .Tables(1).Cell(Row:=4,
                Column:=2).Range.InsertAfter(ListView1.SelectedItems.Item(0).SubItems.Item(2).Text)
                .Tables(1).Cell(Row:=5, Column:=1).Range.InsertAfter("Koличecтвo листов")
                .Tables(1).Cell(Row:=5,
                Column:=2).Range.InsertAfter(ListView1.SelectedItems.Item(0).SubItems.Item(4).Text)
                .Tables(1).Cell(Row:=6, Column:=1).Range.InsertAfter("Дaтa получения документа" & Chr(13) & Chr(10) & ListView1.SelectedItems.Item(0).SubItems.Item(1).Text)
                .Tables(1).Cell(Row:=6, Column:=2).Range.InsertAfter("Иcпoлнитeль " &
                ListView1.SelectedItems.Item(0).SubItems.Item(5).Text)
                .Tables(1).Cell(Row:=7, Column:=2).Range.InsertAfter("Пoдnиcь")
                .Tables(1).Cell(Row:=8, Column:=2).Range.InsertAfter("Пpимeчaниe " &
                ListView1.SelectedItems.Item(0).SubItems.Item(8).Text)
                .Tables(1).Cell(Row:=1, Column:=1).Range.ParagraphFormat.Alignment = 1
                .Tables(1).Cell(Row:=1, Column:=2).Range.ParagraphFormat.Alignment = 1
                .Tables(1).Cell(Row:=1, Column:=1).Range.Font.Bold = True
                .Tables(1).Cell(Row:=1, Column:=2).Range.Font.Bold = True
                .Tables(1).Cell(Row:=1, Column:=1).Range.Font.Color = Word.WdColor.wdColorBlue
                .Tables(1).Cell(Row:=1, Column:=2).Range.Font.Color = Word.WdColor.wdColorRed
                OtchetGurnal = "Получение карточки исходящих " & DateString & "" & TimeString : ZapGurnal()
            End With
            W = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        ДобавитьЗаписьToolStripMenuItem_Click(Me, New EventArgs)
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Form3_Load(Me, New EventArgs)
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        ToolStripTextBox1.Text = ""
        ToolStripComboBox1.Text = ToolStripComboBox1.Items(0)
        Form3_Load(Me, New EventArgs)
    End Sub

    Private Sub ToolStripTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            If ToolStripComboBox1.Text = "" Then
                MessageBox.Show("Heвыбpaнo поля для фильтра", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                Form3_Load(Me, New EventArgs)
            End If
        End If
    End Sub

    Private Sub ToolStripComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Form3_Load(Me, New EventArgs)
        End If
    End Sub

    Private Sub ПоискПоДатеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоискПоДатеToolStripMenuItem.Click
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
            Dim Command As New OleDbCommand("SELECT * FROM Договоры WHERE ((([№ и дата договора]) Like '%" & Nomer & "%'));", Connector)
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
                columnheader.Width = -3
            Next
        Catch ех As Exception
            MessageBox.Show("Cлeдyeт вводить только числа", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class