Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Settings
    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MDIParent1
        Try
            ListBox1.Items.Clear()
            Dim DataReader As OleDbDataReader
            Dim Command As New OleDbCommand("SELECT * FROM Пользователи", Connector)
            Connector.Open()
            DataReader = Command.ExecuteReader
            While DataReader.Read() = True
                ListBox1.Items.Add(DataReader.GetValue(0) & " " & DataReader.GetValue(1) & " " & DataReader.GetValue(2) & " " & DataReader.GetValue(3) & " - " & DataReader.GetValue(4))
            End While
            DataReader.Close()
            Connector.Close()
            TextBox5.Text = SettingsGurnal
            ComboBox1.Text = SettingsOprPolzovatelya
            ComboBox2.Text = SettingsIP
            ComboBox3.Text = SettingsArhivirovanie
            ComboBox4.Text = SettingsVostanovlenieRezervnoiKopii
            ComboBox5.Text = SettingsOtkatBD
            ComboBox6.Text = SettingsAutoObnovlenie
            OtchetGurnal = "Настройки успешно запущены " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Sub Dobavlenie_zap()
        Try
            Dim Command As New OleDbCommand("Insert Into [Пользователи] ([Код], [Фамилия], [Имя],[Отчества], [Права], [Пароль]) values ('" & p1 & "', '" & p2 & "','" & p3 & "','" & p4 & "', '" & p5 & "', '" & p6 & "')", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            Settings_Load(Me, New EventArgs)
            OtchetGurnal = "Добавлен новый пользователь " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Sub Editor_zap()
        Try
            Dim Command As New OleDbCommand("UPDATE [Пользователи] SET [Фамилия]='" & p2 & "', [Имя]='" & p3 & "', [Отчества]='" & p4 & "', [Права]='" & p5 & "', [Пароль]='" & p6 & "' WHERE ([Код] Like '" & ListBox1.SelectedIndex + 1 & "')", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            Settings_Load(Me, New EventArgs)
            OtchetGurnal = "Произведено редактирование пользователя " & DateString & " " & TimeString
            ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub
    Sub Удалить()
        Try
            If ListBox1.SelectedIndex < 0 Then
                MessageBox.Show("He выбрано ни одного пользователя", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            Else
                Dim MBox As DialogResult = MessageBox.Show("Удалить текущего пользователя", "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
                If MBox = DialogResult.No Then Exit Sub
                If MBox = DialogResult.Yes Then
                    Dim Command As New OleDbCommand("DELETE FROM [Пользователи] WHERE [Код] =" &
                    ListBox1.SelectedIndex + 1, Connector)
                    Connector.Open()
                    Command.ExecuteNonQuery()
                    Connector.Close()
                    Settings_Load(Me, New EventArgs)
                    MessageBox.Show("Запись была успешна удалена из БД", "Операция успешна завершина", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    OtchetGurnal = "Пользователь удален " & DateString & " " & TimeString : ZapGurnal()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Settings2.Close()
        Settings2.Text = "Добавление записи"
        Settings2.Show()
        Settings2.Focus()
        OtchetGurnal = "Запущено добавление нового пользователя " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Try
            If ListBox1.SelectedIndex < 0 Then
                MessageBox.Show("He выбрано ни одного пользователя", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                OtchetGurnal = "Не выбрано ни одного пользователя " & DateString & " " & TimeString
                ZapGurnal()
                Exit Sub
            Else
                p1 = ListBox1.SelectedIndex + 1
                Settings2.Close()
                Settings2.Text = "Редактирование записи"
                Settings2.Show()
                Settings2.Focus()
                OtchetGurnal = "Передача данных для редактирования записи " & DateString & " " & TimeString
                ZapGurnal()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Settings_Load(Me, New EventArgs)
        OtchetGurnal = "Обновление параметров настроек " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Удалить()
        OtchetGurnal = "Запущена процедура удаления пользователя " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub TextBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TextBox5.SelectedIndexChanged
        Try
            Dim Command As New OleDbCommand("UPDATE [Параметры] SET [Журнал]='" & TextBox5.Text & "'", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            SettingsGurnal = TextBox5.Text
            OtchetGurnal = "Журнал " & SettingsGurnal & " " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim Command As New OleDbCommand("UPDATE [Параметры] SET [Определять пользователя]='" & ComboBox1.Text & "'", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            SettingsOprPolzovatelya = ComboBox1.Text
            OtchetGurnal = "Определять пользователя " & SettingsOprPolzovatelya & " " & DateString & " " &
            TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            Dim Command As New OleDbCommand("UPDATE [Параметры] SET [IP]='" & ComboBox2.Text & "'", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            SettingsIP = ComboBox2.Text
            OtchetGurnal = "Определять IP " & SettingsIP & " " & DateString & " " & TimeString
            ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            Dim Command As New OleDbCommand("UPDATE [Параметры] SET [Архивировать БД]='" & ComboBox3.Text & "'", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            SettingsArhivirovanie = ComboBox3.Text
            OtchetGurnal = "Архивирование БД " & SettingsArhivirovanie & " " & DateString & " " &
            TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Try
            Dim Command As New OleDbCommand("UPDATE [Параметры] SET [Востановление резервной копии БД]='" & ComboBox4.Text & "'", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            SettingsVostanovlenieRezervnoiKopii = ComboBox4.Text
            OtchetGurnal = "Востановление резервной копии БД " & SettingsVostanovlenieRezervnoiKopii & " " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Try
            Dim Command As New OleDbCommand("UPDATE [Параметры] SET [Востановление в исходный вид БД]='" & ComboBox5.Text & "'", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            SettingsOtkatBD = ComboBox5.Text
            OtchetGurnal = "Востановление в исходный вид БД " & SettingsOtkatBD & " " & DateString & " " &
            TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        Try
            Dim Command As New OleDbCommand("UPDATE [Параметры] SET [Автообновление]='" & ComboBox6.Text & "'", Connector)
            Connector.Open()
            Command.ExecuteNonQuery()
            Connector.Close()
            SettingsAutoObnovlenie = ComboBox6.Text
            OtchetGurnal = "Автообновление " & SettingsAutoObnovlenie & " " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ListBox1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseDoubleClick
        Try
            ToolStripButton2_Click(Me, New EventArgs)
            OtchetGurnal = "Запуск редактирования пользователя " & DateString & " " & TimeString : ZapGurnal()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub
End Class