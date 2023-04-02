Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Public Class Master

    Private Sub Master_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MDIParent1
        Prava()
        Dim da As New OleDb.OleDbDataAdapter
        Dim tbl As DataTable
        Dim cn As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBaseFileName)
        cn.Open()
        da.SelectCommand = New OleDb.OleDbCommand("Select Name from Msysobjects where type = 1", cn)
        tbl = cn.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "Table"})
        For i = 2 To tbl.Rows.Count - 1
            ComboBox1.Items.Add(tbl.Rows(i) !Table_Name.ToString)
        Next
        cn.Close()
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListView1.Columns.Clear()
        ListView1.Items.Clear()
        Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListView1.Columns.Clear()
        ListView1.Items.Clear()
        Dim cn As Object
        Dim rs As Object
        Dim fld As Object
        cn = CreateObject("ADODB.Connection")
        rs = CreateObject("ADODB.Recordset")
        cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBaseFileName)
        rs.Open("select * from [" & ComboBox1.Text & "] where 1=2", cn)
        For Each fld In rs.Fields
            ListBox1.Items.Add(fld.Name)
        Next
        rs.Close()
        rs = Nothing
        cn.Close()
        cn = Nothing
    End Sub

    Private Sub ListView1_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView1.ColumnClick
        If (e.Column = lvwColumnSorter.SortColumn) Then
            If (lvwColumnSorter.Order = SortOrder.Ascending) Then
                lvwColumnSorter.Order =
                SortOrder.Descending
            Else lvwColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            lvwColumnSorter.SortColumn = e.Column
            lvwColumnSorter.Order = SortOrder.Ascending
        End If
        Me.ListView1.Sort()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox2.Items.Count = 0 Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub ListBox1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseDoubleClick
        ListBox2.Items.Add(ListBox1.SelectedItem)
        ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
    End Sub

    Private Sub ListBox2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListBox2.MouseDoubleClick
        ListBox1.Items.Add(ListBox2.SelectedItem)
        ListBox2.Items.RemoveAt(ListBox2.SelectedIndex)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            MDIParent1.ToolStripProgressBar1.Value = 10
            Dim k1 As Integer
            Dim status As Integer
            Dim k2 As Integer
            status = ListBox2.Items.Count
            lvwColumnSorter = New ListViewColumnSorter()
            Me.ListView1.ListViewItemSorter = lvwColumnSorter
            For k1 = 1 To status
                ListView1.Columns.Add(k1)
            Next k1
            MDIParent1.ToolStripProgressBar1.Value = 20
            Dim Polya As String
            Polya = ""
            For k3 = 0 To ListBox2.Items.Count - 1
                Polya = Polya + "[" + ListBox2.Items(k3) + "]"
                If k3 <> ListBox2.Items.Count - 1 Then Polya = Polya + ", "
            Next
            MDIParent1.ToolStripProgressBar1.Value = 30
            Dim DataReader As OleDbDataReader
            Dim Zapros As String
            Zapros = TextBox1.Text
            If Zapros = " " Then

                Dim Command As New OleDbCommand("SELECT " & Polya & " FROM " & ComboBox1.Text, Connector)
                Connector.Open()
                DataReader = Command.ExecuteReader
            Else
                Dim Command As New OleDbCommand("SELECT " & Polya & " FROM [" & ComboBox1.Text & "] WHERE [" & ListBox2.SelectedItem & "] Like '%" & Zapros & "%';", Connector)
                Connector.Open()
                DataReader = Command.ExecuteReader
            End If
            MDIParent1.ToolStripProgressBar1.Value = 40
            While DataReader.Read() = True
                ListView1.Items.Add(DataReader.GetValue(0))
                For k2 = 1 To status - 1
                    ListView1.Items.Item(ListView1.Items.Count - 1).SubItems.Add(DataReader.GetValue(k2))
                Next k2
            End While
            MDIParent1.ToolStripProgressBar1.Value = 60
            DataReader.Close()
            Connector.Close()
            Dim Дата As String = Format(Now, "d MMMM уууу")
            Dim W = New Word.Application
            W.Visible = True
            W.Documents.Add()
            W.Selection.TypeText("Выписка из БД: " & Дата & Chr(13) & Chr(10))
            MDIParent1.ToolStripProgressBar1.Value = 70
            For i As Short = 0 To ListView1.Items.Count - 1
                For j = 0 To status - 1
                    W.Selection.TypeText(ListView1.Items(i).SubItems.Item(j).Text & " ")
                Next
                W.Selection.TypeText(Chr(13) & Chr(10))
            Next i
            W = Nothing
            MDIParent1.ToolStripProgressBar1.Value = 100
        Catch ex As Exception
            Connector.Close()
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        If ListBox2.Items.Count = 0 Then
            TextBox1.Enabled = False
            TextBox1.Text =
            Button1.Enabled = False
        Else
            TextBox1.Enabled = True
            TextBox1.Text = " "
            Button1.Enabled = True
        End If
    End Sub
End Class