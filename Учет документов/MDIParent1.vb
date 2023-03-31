Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Public Class MDIParent1
    Private Sub MDIParent1_ContextMenuStripChanged(sender As Object, e As EventArgs) Handles Me.ContextMenuStripChanged
        ToolStripStatusLabel2.Text = Now.ToLongDateString
        ToolStripStatusLabel4.Text = Now.ToLongTimeString
        Dim DataReader1 As OleDbDataReader
        Dim Command1 As New OleDbCommand("SELECT * FROM Параметры", Connector)
        Connector.Open()
        DataReader1 = Command1.ExecuteReader

        While DataReader1.Read() = True
            SettingsGurnal = DataReader1.GetValue(0)
            SettingsOprPolzovatelya = DataReader1.GetValue(1)
            SettingsObnovlenie = DataReader1.GetValue(2)
            SettingsIP = DataReader1.GetValue(3)
            SettingsArhivirovanie = DataReader1.GetValue(4)
            SettingsVostanovlenieRezervnoiKopii = DataReader1.GetValue(5)
            SettingsOtkatBD = DataReader1.GetValue(6)
            SettingsAutoObnovlenie = DataReader1.GetValue(7)
        End While
        DataReader1.Close()
        Connector.Close()
        Prava()
        OtchetGurnal = "Программа успешно запущена " & DateString & " " & TimeString : ZapGurnal()

        Me.LayoutMdi(MdiLayout.Cascade)
        Me.LayoutMdi(MdiLayout.TileVertical)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
        MsgBox(Chr(10) & "Программа для ведения учета регистрационных документов 3Ф_РАНХиГС", vbOKOnly + vbInformation, "Информация")
        MsgBox(Chr(10) & "Автор: Алехин Денис Владимирович", vbOKOnly + vbInformation, "Информация")
        MsgBox(Chr(10) & "Данная программа защищена законом об авторских правах и международными соглашениями. Незаконное воспроизведение или распространение данной программы или любой ее части влечет гражданску и уголовную ответственность" & Chr(13) & "Данная программа специально написана для Общего Отдела - 3Ф_РАНХиГС", vbOKOnly + vbInformation, "Информация")
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        OtchetGurnal = "Программа завершена " & DateString & " " & TimeString
        Me.Close()
    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked
        'Me.StatusStrip1.Visible = Me.StatusBarToolStripMenuItem.Checked
    End Sub

    'Private Sub StatusStrip1_Validating(sender As Object, e As CancelEventArgs) Handles StatusStrip1.Validating
    '    Me.StatusStrip1.Visible = Me.StatusBarToolStripMenuItem.Checked
    'End Sub

    Private Sub ИсходящиеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ИсходящиеToolStripMenuItem.Click
        Form1.Close()
        Form1.Show()
        OtchetGurnal = "Запущена книга Исходящие " & DateString & " " & TimeString : ZapGurnal()
    End Sub


    Private Sub ВходящиеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВходящиеToolStripMenuItem.Click
        Form3.Close()
        Form3.Show()
        OtchetGurnal = "Запущена книга Входящие " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ДоговорыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДоговорыToolStripMenuItem.Click
        Form5.Close()
        Form5.Show()
        OtchetGurnal = "Запущена книга Договоры " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ПриказыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПриказыToolStripMenuItem.Click
        Form7.Close()
        Form7.Show()
        OtchetGurnal = "Запущена книга Приказы " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ПриказыПоСтудентамToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПриказыПоСтудентамToolStripMenuItem.Click
        Form9.Close()
        Form9.Show()
        OtchetGurnal = "Запущена книга Приказы по студентам " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ПриказыПоАХДToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПриказыПоАХДToolStripMenuItem.Click
        Form11.Close()
        Form11.Show()
        OtchetGurnal = "Запущена книга Приказы по АХД " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ДокладныеИСлужебныеЗапискиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДокладныеИСлужебныеЗапискиToolStripMenuItem.Click
        Form13.Close()
        Form13.Show()
        OtchetGurnal = "Запущена книга Докладные и служебные записки " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub АктыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles АктыToolStripMenuItem.Click
        Form15.Close()
        Form15.Show()
        OtchetGurnal = "Запущена книга Акты " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub РаспоряженияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles РаспоряженияToolStripMenuItem.Click
        Form17.Close()
        Form17.Show()
        OtchetGurnal = "Запущена книга Распоряжения " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ДоверенностиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДоверенностиToolStripMenuItem.Click
        Form19.Close()
        Form19.Show()
        OtchetGurnal = "Запущена книга Доверенности " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ToolStripProgressBar1.Value = ToolStripProgressBar1.Value + 1
        If ToolStripProgressBar1.Value = 100 Then
            Timer1.Enabled = False
            ToolStripStatusLabel2.Text = Now.ToLongDateString
            ToolStripStatusLabel4.Text = Now.ToLongTimeString
        End If
    End Sub

    Private Sub МастерПостроенияОтчетовToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles МастерПостроенияОтчетовToolStripMenuItem.Click
        'Master.Close()
        'Master.Show()
        OtchetGurnal = "Открыт мастер запросов " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ПоставитьНаКонтрольToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоставитьНаКонтрольToolStripMenuItem.Click
        'Kontrol.Close()
        'Kontrol.Show()
        OtchetGurnal = "Контроль открыт " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub НастройкиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НастройкиToolStripMenuItem.Click
        'Settings.Close()
        'Settings.Show()
        OtchetGurnal = "Открыты настройки " & DateString & " " & TimeString : ZapGurnal()
    End Sub

    Private Sub ЖурналСобытийToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЖурналСобытийToolStripMenuItem.Click
        'Gurnal.Close()
        'Gurnal.Show()
        OtchetGurnal = "Открыт журнал " & DateString & " " & TimeString : ZapGurnal() : ZapGurnal()
    End Sub

    Private Sub АвторизоватьсяToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles АвторизоватьсяToolStripMenuItem.Click
        Try
            Dim Proverka As String
            Proverka = InputBox("Введите пароль для авторизации пользователя", "Запрос авторизации")
            If Proverka = "" Then
                Exit Sub
            Else
                Dim DataReader As OleDbDataReader
                Dim Command As New OleDbCommand("SELECT Пользователи.[Код], Пользователи.[Фамилия], Пользователи.[Имя], Пользователи.[Отчество], Пользователи.[Права], Пользователи.[Пароль] FROM Пользователи WHERE (((Пользователи.[Пароль])='" & Proverka & "'));", Connector)
                Connector.Open()
                DataReader = Command.ExecuteReader
                While DataReader.Read() = True
                    PravaPolzovatelyaKod = DataReader.GetValue(0)
                    PravaPolzovatelyaFamiliya = DataReader.GetValue(1)
                    PravaPolzovatelyaImya = DataReader.GetValue(2)
                    PravaPolzovatelyaOtchestvo = DataReader.GetValue(3)
                    PravaPolzovatelyaPrava = DataReader.GetValue(4)
                    PravaPolzovatelyaPassword = DataReader.GetValue(5)
                End While
                If PravaPolzovatelyaKod = "" Then MessageBox.Show("Авторизация не удалась", "Не удача", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                If PravaPolzovatelyaKod <> "" Then
                    MessageBox.Show("Дoбpo пожаловать в программу: " & PravaPolzovatelyaFamiliya & " " & PravaPolzovatelyaImya & " " & PravaPolzovatelyaOtchestvo, "Авторизация", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ToolStripStatusLabel5.Text = "Авторизован: " & PravaPolzovatelyaFamiliya & " " & PravaPolzovatelyaImya & " " & PravaPolzovatelyaOtchestvo
                End If
                DataReader.Close()
                Connector.Close()
                Prava()
                OtchetGurnal = "Попытка авторизации в программе " & DateString & " " & TimeString
                ZapGurnal()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub АрхивированиеБДToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles АрхивированиеБДToolStripMenuItem.Click
        Try
            Dim МВох As DialogResult = MessageBox.Show("Bы уверенны что хотите запустить архивирование всех данных при этом старый архив будет удален???", "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If МВох = DialogResult.No Then Exit Sub
            If МВох = DialogResult.Yes Then
                ToolStripProgressBar1.Value = 25
                For Each ChildForm As Form In Me.MdiChildren
                    ChildForm.Close()
                Next
                ToolStripProgressBar1.Value = 50
                IO.File.Delete(Application.StartupPath & "/bd - архив")
                ToolStripProgressBar1.Value = 75
                IO.File.Copy(Application.StartupPath & "/bd.mdb", Application.StartupPath & "/bd - архив")
                ToolStripProgressBar1.Value = 100
                MessageBox.Show("Архивирование успешно завершено!!!", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information)
                OtchetGurnal = "Запущена архивирование БД" & DateString & " " & TimeString : ZapGurnal()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ВосстановлениеРезервнойКопииБДToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВосстановлениеРезервнойКопииБДToolStripMenuItem.Click
        Try
            Dim МВох As DialogResult = MessageBox.Show("Bы уверенны что хотите восстановить БД из архива при этом текущие данные будут утерены???", "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If МВох = DialogResult.No Then Exit Sub
            If МВох = DialogResult.Yes Then
                ToolStripProgressBar1.Value = 25
                For Each ChildForm As Form In Me.MdiChildren
                    ChildForm.Close()
                Next
                ToolStripProgressBar1.Value = 50
                IO.File.Delete(Application.StartupPath & "/bd.mdb")
                ToolStripProgressBar1.Value = 75
                IO.File.Copy(Application.StartupPath & "/bd - архив", Application.StartupPath & "/bd.mdb")
                ToolStripProgressBar1.Value = 100
                MessageBox.Show("Восстановление из архива выnолненно успешно!!!", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information)
                OtchetGurnal = "Восстановление из архива БД" & DateString & " " & TimeString : ZapGurnal()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ВосстановлениеВИсходныйВидБДToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВосстановлениеВИсходныйВидБДToolStripMenuItem.Click
        Try
            Dim МВох As DialogResult = MessageBox.Show("Bы уверенны что хотите БД обнулить, что приведет к потери всех данных???", "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If МВох = DialogResult.No Then Exit Sub
            If МВох = DialogResult.Yes Then
                ToolStripProgressBar1.Value = 25
                For Each ChildForm As Form In Me.MdiChildren
                    ChildForm.Close()
                Next
                ToolStripProgressBar1.Value = 50
                IO.File.Delete(Application.StartupPath & "/bd.mdb")
                ToolStripProgressBar1.Value = 75
                IO.File.Copy(Application.StartupPath & "/bd - пустая", Application.StartupPath & "/bd.mdb")
                ToolStripProgressBar1.Value = 100
                MessageBox.Show("БД возвращена в исходный вид и все записи успешно удалены!!!", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information)
                OtchetGurnal = "Запущена процедура восстановление бд в исходный вид " & DateString & "" & TimeString : ZapGurnal()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub НаписатьПисьмоРазработчикуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НаписатьПисьмоРазработчикуToolStripMenuItem.Click
        Try
            Dim Proverka As String
            Proverka = InputBox("Введите текст сообщения который будет автоматически отправлен… ", "Отправка сообщения")
            If Proverka = "" Then
                Exit Sub
            Else
                Dim olApp As Outlook.Application
                olApp = CreateObject("Outlook.Application")
                Dim olItem As Outlook.ContactItem
                olItem = olApp.CreateItem(Outlook.OlItemType.olContactItem)
                With olItem
                    .FullName = "Алехин Денис Владимирович"
                    .Birthday = "09/01/2003"
                    .CompanyName = "3Ф РАНХиГС"
                    .Email1Address = " stupidlymymail@gmail.com "
                    .JobTitle = "Разработчик"
                End With
                olItem.Save()
                Dim olMail As Outlook.MailItem
                olMail = olApp.CreateItem(Outlook.OlItemType.olMailItem)
                olMail.To = olItem.Email1Address
                olMail.Subject = "Есть вопрос по программе..."
                olMail.Body = "Уважаемый " & olItem.FirstName & ", " & vbCr & vbCr & vbTab & Proverka
                olMail.Send()
                MsgBox("Сообщение успешно отправлено!!!", vbMsgBoxSetForeground)
                olMail = Nothing
                olItem = Nothing
                olApp = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            OtchetGurnal = "Ошибка " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ОПрограммеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОПрограммеToolStripMenuItem.Click
        Try
            about.Show()
        Catch ex As Exception
            OtchetGurnal = "Запущена форма О программе " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ОАвтореToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОАвтореToolStripMenuItem.Click
        Try
            about_me.Show()
        Catch ex As Exception
            OtchetGurnal = "Запущена форма О Авторе " & ex.Message & DateString & " " & TimeString : ZapGurnal()
        End Try
    End Sub

    Private Sub ЛицензионноеСоглашениеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЛицензионноеСоглашениеToolStripMenuItem.Click
        Licenses.Show()
    End Sub
End Class