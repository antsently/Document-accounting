Public Class Licenses
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)
        Button1.Left = CInt((Me.ClientSize.Width - Button1.Width) / 2)
        Button1.Top = Me.ClientSize.Height - Button1.Height - 15
    End Sub
End Class