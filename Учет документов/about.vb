Public Class about
    Private Sub about_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim url As String = "https://github.com/Antsently"
        Process.Start(url)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim url As String = "https://vk.com/antsently"
        Process.Start(url)
    End Sub
End Class