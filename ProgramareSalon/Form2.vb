Imports System.Data.OleDb
Imports System.Globalization
Imports System.Net.Mail

Public Class Form2

    Dim mailaddress As MailAddress
    Dim name As String
    Dim lastName As String
    Dim email As String
    Dim obs As String
    Dim data_plan As Date
    Dim hour As String
    Dim phone As String

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.TextBox1.AppendText("")
        Me.DateTimePicker1.Value = DateTime.Today.AddDays(-1)
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
        DateTimePicker1.Value = DateTime.Today.ToString("yyyy-MM-dd")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        HomePage.Show()
        Me.Hide()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        Dim date1 As DateTime = DateTime.ParseExact(DateTime.Today.ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture)
        Dim date2 As DateTime = DateTime.ParseExact(DateTimePicker1.Value().ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture)
        If date1 <= date2 Then
            ' complet withe possible hour in list
            ' search for hour disponible
            Label9.Text = ""
            ComboBox1.Items.Clear()
            ComboBox1.Items.Insert(0, "07:11")
            ComboBox1.Items.Insert(0, "10:11")
            ComboBox1.Items.Insert(1, "10:12")
            ComboBox1.Items.Insert(2, "10:13")
            ComboBox1.Items.Insert(3, "10:15")
        Else
            Label9.Text = "Setează o dată validă"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'ToDo
        'Step 1: check complete all information
        ' CheckData()

        'Step 2: add the databse
        'Create connection string
        If CheckData() Then
            Dim con As OleDbConnection = New OleDbConnection
            Try
                'To Connect the Database
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=A:\CristinaProject\ProgramareSalon\SalonDB.accdb"
                con.Open()

                Dim query As String
                query = "Insert into Programari  (Nume, Prenume, [E-mail], Observatii, Data, Ora, Telefon, Timeinserted)
                            VALUES ('" & name & "', '" & lastName & "', '" & email & "', '" & obs & "', #" & data_plan & "#, #" & hour & "#, '" & phone & "', #" & Date.Now & "#)"

                Dim command As OleDbCommand = New OleDbCommand(query, con)
                command.ExecuteNonQuery()
                command.Dispose()

                Label9.ForeColor = Color.Green
                Label9.Text = "Adaugată cu succes!"
            Catch ex As Exception
                MsgBox(Convert.ToString(ex))
            Finally
                con.Close()
            End Try
        End If

        'Step 3: message to check if want to add another or getout

        'Step 4: clean forms



    End Sub

    Private Function getTimeNow() As String
        Dim now As DateTime = DateTime.Now
        Dim hour As String = now.ToString("HH")
        Dim minute As String = now.ToString("mm")
        Return hour & ":" & minute
    End Function

    Private Function CheckData() As Boolean
        name = Me.TextBox3.Text
        lastName = Me.TextBox4.Text
        email = Me.TextBox6.Text
        obs = Me.TextBox5.Text
        data_plan = Me.DateTimePicker1.Value
        hour = Me.ComboBox1.SelectedItem
        phone = Me.TextBox8.Text

        If name.Length <= 3 Then
            Label9.Text = "Numele nu a fost setat sau nu este valid"
            Return False
        End If
        If lastName.Length <= 3 Then
            Label9.Text = "Prenumele nu a fost setat sau este prea scurt"
            Return False
        End If
        'check email
        If email.Length = 0 Then
            Label9.Text = "E-mail-ul nu a fost setat"
            Return False
        Else
            Try
                mailaddress = New MailAddress(email)
            Catch ex As Exception
                Label9.Text = "E-mail-ul nu este valid!"
                Return False
            End Try
        End If
        'If data_plan.Date < Now.Date Then
        '    Label9.Text = "Setează o dată validă"
        '    Return False
        'End If
        'If hour < Date.Now Then
        '    Label9.Text = "Ora nu a fost selectată"
        '    Return False
        'End If
        Dim time1 As TimeSpan = TimeSpan.ParseExact(getTimeNow.ToString, "hh\:mm", Nothing)
        Dim time2 As TimeSpan = TimeSpan.ParseExact(hour, "hh\:mm", Nothing)
        If time2 < time1 Then
            Label9.Text = "Nu a fost selectată o ora valida"
            Return False
        End If
        If phone.Length <> 10 Then
            Label9.Text = "Adaugă numărul de telefon sau  introdu unul valid"
            Return False
        End If


        Return True

    End Function

End Class