Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Kontrol

    Private Sub Kontrol_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MDIParent1
        Prava()
        lvwColumnSorter = New ListViewColumnSorter()
        Me.ListView1.ListViewItemSorter = lvwColumnSorter
        ListView1.Columns.Clear()
        ListView1.Columns.Add("№ п/п")
        ListView1.Columns.Add("№ и дата договора")
        ListView1.Columns.Add("Поставщик, исполнитель, подрядчик")
        ListView1.Columns.Add("Инициатор (ответственный)")
        ListView1.Columns.Add("Дата исполнения")
        ListView1.Items.Clear()
        'Попытка реализации работы с датой
        Try
            Dim Poisk As String
            Dim endDate As DateTime = DateTime.Now.AddDays(7)
            If ToolStripTextBox1.Text = "" Then
                Poisk = "SELECT Договоры.[№ п/п], Договоры.[№ и дата договора], Договоры.[Поставщик, исполнитель, подрядчик], Договоры.[Инициатор (ответственный)], Договоры.[Дата исполнения] FROM Договоры ORDER BY Договоры.[№ п/п] DESC"
            Else
                Poisk = "SELECT Договоры.[№ п/п], Договоры.[№ и дата договора], Договоры.[Поставщик, исполнитель, подрядчик], Договоры.[Инициатор (ответственный)], Договоры.[Дата исполнения] FROM Договоры WHERE ((Договоры.[" & ToolStripComboBox1.Text & "]) Like '%" & ToolStripTextBox1.Text & "%')"
            End If
            Dim DataReader As OleDbDataReader
            Dim Command As New OleDbCommand(Poisk, Connector)
            Connector.Open()
            DataReader = Command.ExecuteReader
            While DataReader.Read()
                Dim item As New ListViewItem(DataReader("№ п/п").ToString())
                item.SubItems.Add(DataReader("№ и дата договора").ToString())
                item.SubItems.Add(DataReader("Поставщик, исполнитель, подрядчик").ToString())
                item.SubItems.Add(DataReader("Инициатор (ответственный)").ToString())
                Dim dateStr As String = DataReader("Дата исполнения").ToString()
                Dim dateValue As DateTime
                If DateTime.TryParse(dateStr, dateValue) AndAlso dateValue <= endDate Then
                    item.SubItems.Add(dateValue.ToString("dd.MM.yyyy HH:mm"))
                Else
                    item.SubItems.Add("")
                End If
                ListView1.Items.Add(item)
            End While
            DataReader.Close()
            Connector.Close()
        Catch ex As Exception
            Connector.Close()
            MessageBox.Show(ex.Message)
        End Try
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        For Each columnheader In Me.ListView1.Columns
            columnheader.Width = -2
            columnheader.TextAlign = HorizontalAlignment.Center
        Next
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Kontrol_Load(Me, New EventArgs)
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Kontrol_Load(Me, New EventArgs)
    End Sub
End Class