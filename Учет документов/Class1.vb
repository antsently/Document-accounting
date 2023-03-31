Imports System.Collections
Imports System.Windows.Forms
Public Class ListViewColumnSorter
    Implements System.Collections.IComparer
    Private ColumnToSort As Integer
    Private OrderOfSort As SortOrder
    Private ObjectCompare As CaseInsensitiveComparer
    Public Sub New()
        ColumnToSort = 0
        OrderOfSort = SortOrder.None
        ObjectCompare = New CaseInsensitiveComparer()
    End Sub

    Public Function Compare(ByVal х As Object, ByVal у As Object) As Integer Implements IComparer.Compare
        Dim compareResult As Integer
        Dim listviewX As ListViewItem
        Dim listviewY As ListViewItem
        listviewX = CType(х, ListViewItem)
        listviewY = CType(у, ListViewItem)
        compareResult = ObjectCompare.Compare(listviewX.SubItems(ColumnToSort).Text, listviewY.SubItems(ColumnToSort).Text)
        If (OrderOfSort = SortOrder.Ascending) Then
            Return compareResult
        ElseIf (OrderOfSort = SortOrder.Descending) Then
            Return (-compareResult)
        Else
            Return 0
        End If
    End Function

    Public Property SortColumn() As Integer
        Set(ByVal Value As Integer)
            ColumnToSort = Value
        End Set
        Get
            Return ColumnToSort
        End Get
    End Property
    Public Property Order() As SortOrder
        Set(ByVal Value As SortOrder)
            OrderOfSort = Value
        End Set
        Get
            Return OrderOfSort
        End Get
    End Property
End Class
