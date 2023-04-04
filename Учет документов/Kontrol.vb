Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Kontrol
    Private Sub Kontrol_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MDIParent1
        Prava()
        lvwColumnSorter = New ListViewColumnSorter()
        Me.ListView1.ListViewItemSorter = lvwColumnSorter
        ListView1.Columns.Clear()
        ListView1.Columns.Add("ID")
        ListView1.Columns.Add("Тип документа")
        ListView1.Columns.Add("Дата добавления события")
        ListView1.Columns.Add("Примечание")
        ListView1.Items.Clear()
        Try
            Dim Poisk As String
            If ToolStripTextBox1.Text = "" Then
                Poisk = "SELECT * FROM События ORDER BY Входящая.[Код] DESC"
            Else
                Poisk = "SELECT * FROM События WHERE ((Входящая.[" & ToolStripComboBox1.Text & "]) Like '%" & ToolStripTextBox1.Text & "%')"
            End If
            Dim DataReader As OleDbDataReader
            Dim Command As New OleDbCommand(Poisk, Connector)
            Connector.Open()
            DataReader = Command.ExecuteReader
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
End Class