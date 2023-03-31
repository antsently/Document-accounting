Public Class Form18
    Private Sub Form18_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MDIParent1
        Prava()
        If Me.Text = "Добавление записи" Then
            DateTimePicker1.Text = Now.ToLongDateString
            TextBox1.Text = CountForm9 + 1
            OtchetGurnal = "Загрузка данных для добавления Распоряжения " & DateString & " " & TimeString : ZapGurnal()
        Else
            TextBox1.Text = p1
            DateTimePicker1.Value = p2
            TextBox3.Text = p3
            TextBox4.Text = p4
            TextBox5.Text = p5
            TextBox6.Text = p6
            TextBox7.Text = p7
            OtchetGurnal = "Загрузка данных для редактирования Распоряжения " & DateString & " " & TimeString : ZapGurnal()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Text = "Добавление записи" Then
            If TextBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
                MessageBox.Show("He все строки заполненны! ! !", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                OtchetGurnal = "Не все строки заполнили! " & DateString & " " & TimeString : ZapGurnal()
            Else
                p1 = TextBox1.Text
                p2 = DateTimePicker1.Value
                p3 = TextBox3.Text
                p4 = TextBox4.Text
                p5 = TextBox5.Text
                p6 = TextBox6.Text
                p7 = TextBox7.Text
                Form17.Dobavlenie_zap()
                Me.Close()
                OtchetGurnal = "Передача данных на выполнения добавления Распоряжения " & DateString & " " & TimeString : ZapGurnal()
            End If
        Else
            p1 = TextBox1.Text
            p2 = DateTimePicker1.Value
            p3 = TextBox3.Text
            p4 = TextBox4.Text
            p5 = TextBox5.Text
            p6 = TextBox6.Text
            p7 = TextBox7.Text
            Form17.Editor_zap()
            Me.Close()
            OtchetGurnal = "Передача данных на выполнения редактирования Распоряжения " & DateString & " " & TimeString : ZapGurnal()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OtchetGurnal = "Отмена добавления/редактирования Распоряжения " & DateString & " " & TimeString : ZapGurnal()
        Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim МВох As DialogResult = MessageBox.Show("Bы уверены что хотите изменить ключевое поле в БД которое может нарушить работу программы и целостность данных???", "Запрос", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        Try
            If МВох = DialogResult.No Then Exit Sub
            If МВох = DialogResult.Yes Then
                TextBox1.Enabled = True
                Button3.Visible = False
                OtchetGurnal = "Вмешательства в ключевое поле Распоряжения " & DateString & " " & TimeString : ZapGurnal()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub
End Class