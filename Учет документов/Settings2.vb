Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Public Class Settings2
    Private Sub Settings2_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.MdiParent = MDIParent1
        Try
            If Me.Text = "Редактирование записи" Then
                Dim DataReader As OleDbDataReader
                Dim Command As New OleDbCommand("SELECT Пользователи.Код, Пользователи.Фамилия, Пользователи.Имя, Пользователи.Отчество, Пользователи.Права, Пользователи.Пароль FROM Пользователи WHERE (((Пользователи.Код)=" & p1 & "));", Connector)
                Connector.Open()
                DataReader = Command.ExecuteReader
                While DataReader.Read() = True
                    TextBox1.Text = (DataReader.GetValue(0))
                    TextBox2.Text = (DataReader.GetValue(1))
                    TextBox3.Text = (DataReader.GetValue(2))
                    TextBox4.Text = (DataReader.GetValue(3))
                    TextBox5.Text = (DataReader.GetValue(4))
                    TextBox6.Text = (DataReader.GetValue(5))
                End While
                DataReader.Close()
                Connector.Close()
                OtchetGurnal = "Загрузка редактор пользователя успешна " & DateString & " " & TimeString : ZapGurnal()
            Else
                Dim DataReader As OleDbDataReader
                Dim Command As New OleDbCommand("SELECT Count([Пользователи]![Код]) AS [Count] FROM Пользователи;", Connector)
                Connector.Open()
                DataReader = Command.ExecuteReader
                While DataReader.Read() = True
                    TextBox1.Text = (DataReader.GetValue(0)) + 1
                End While
                DataReader.Close()
                Connector.Close()
                OtchetGurnal = "Загрузка добавления пользователя успешна " & DateString & " " & TimeString : ZapGurnal()
            End If
        Catch ex As Exception
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Enabled = True
        OtchetGurnal = "Изменение системного регистра " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox6.PasswordChar = ""
        OtchetGurnal = "Пароль теперь виден всем " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Text = "Добавление записи" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
                MessageBox.Show("He все строки заполненны!!!", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            Else
                p1 = TextBox1.Text
                p2 = TextBox2.Text
                p3 = TextBox3.Text
                p4 = TextBox4.Text
                p5 = TextBox5.Text
                p6 = TextBox6.Text
                Settings.Dobavlenie_zap()
                Me.Close()
                OtchetGurnal = "Запущена процедура добавления пользователя в БД " & DateString & " " & TimeString : ZapGurnal()
            End If
        Else
            p1 = TextBox1.Text
            p2 = TextBox2.Text
            p3 = TextBox3.Text
            p4 = TextBox4.Text
            p5 = TextBox5.Text
            p6 = TextBox6.Text
            Settings.Editor_zap()
            Me.Close()
            OtchetGurnal = "Запущена процедура редактирования пользователя в БД " & DateString & " " & TimeString : ZapGurnal()
        End If
    End Sub
End Class