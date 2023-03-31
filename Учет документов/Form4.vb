Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MDIParent1
        Prava()
        If Me.Text = "Добавление записи" Then
            DateTimePicker1.Text = Now.ToLongDateString
            TextBox2.Text = "122/01-29" & " " & CountForm2 + 1
        Else
            DateTimePicker1.Value = p1
            TextBox2.Text = p2
            TextBox3.Text = p3
            TextBox4.Text = p4
            TextBox5.Text = p5
            TextBox6.Text = p6
            TextBox7.Text = p7
            TextBox8.Text = p8
            TextBox9.Text = p9
            TextBox10.Text = p10
            TextBox11.Text = p11
            TextBox12.Text = p12
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Text = "Добавление записи" Then
            If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Then
                MessageBox.Show("He все строки заnолненны! ! !", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                p1 = DateTimePicker1.Value
                p2 = TextBox2.Text
                p3 = TextBox3.Text
                p4 = TextBox4.Text
                p5 = TextBox5.Text
                p6 = TextBox6.Text
                p7 = TextBox7.Text
                p8 = TextBox8.Text
                p9 = TextBox9.Text
                p10 = TextBox10.Text
                p11 = TextBox11.Text
                p12 = TextBox12.Text
                Form3.Dobavlenie_zap()
                Me.Close()
            End If
        Else
            p1 = DateTimePicker1.Value
            p2 = TextBox2.Text
            p3 = TextBox3.Text
            p4 = TextBox4.Text
            p5 = TextBox5.Text
            p6 = TextBox6.Text
            p7 = TextBox7.Text
            p8 = TextBox8.Text
            p9 = TextBox9.Text
            p10 = TextBox10.Text
            p11 = TextBox11.Text
            p12 = TextBox12.Text
            Form3.Editor_zap()
            Me.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim МВох As DialogResult = MessageBox.Show("Bы уверены что хотите изменить ключевое поле в БД которое может нарушить работу программы и целостность данных???", "Запрос", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        Try
            If МВох = DialogResult.No Then Exit Sub
            If МВох = DialogResult.Yes Then
                TextBox2.Enabled = True
                Button3.Visible = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class