Module Module1
    Public SettingsGurnal As String
    Public SettingsOprPolzovatelya As String
    Public SettingsObnovlenie As String
    Public SettingsIP As String
    Public SettingsArhivirovanie As String
    Public SettingsVostanovlenieRezervnoiKopii As String
    Public SettingsOtkatBD As String
    Public SettingsAutoObnovlenie As String
    Public OtchetGurnal As String
    Public DataBaseFileName As String = Application.StartupPath & "\bd.mdb"
    Public Connector As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBaseFileName)
    Public PravaPolzovatelyaKod As String
    Public PravaPolzovatelyaFamiliya As String
    Public PravaPolzovatelyaImya As String
    Public PravaPolzovatelyaOtchestvo As String
    Public PravaPolzovatelyaPrava As String = "Пользователь"
    Public PravaPolzovatelyaPassword As String
    Public CountForm1 As String
    Public CountForm2 As String
    Public CountForm3 As String
    Public CountForm4 As String
    Public CountForm5 As String
    Public CountForm6 As String
    Public CountForm7 As String
    Public CountForm8 As String
    Public CountForm9 As String
    Public CountForm10 As String
    Public p1 As String
    Public p2 As String
    Public p3 As String
    Public p4 As String
    Public p5 As String
    Public p6 As String
    Public p7 As String
    Public p8 As String
    Public p9 As String
    Public p10 As String
    Public p11 As String
    Public p12 As String
    Public lvwColumnSorter As ListViewColumnSorter

    Public Sub Prava()
        'код позже'
    End Sub
End Module

