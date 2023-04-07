Imports System.Data.OleDb
Public Class Form2
    Dim name As String
    Dim lastName As String
    Dim email As String
    Dim obs As String
    Dim data_plan As Date
    Dim hour As Date
    Dim phone As String

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.TextBox1.AppendText("")
        Me.DateTimePicker1.Value = DateTime.Today.AddDays(-1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        HomePage.Show()
        Me.Hide()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If Date.Now.ToShortDateString <= DateTimePicker1.Value().ToShortDateString Then
            ' search for hour disponible
            ComboBox1.Items.Insert(0, "4/10/2023 10:11 PM")
            ComboBox1.Items.Insert(1, "4/10/2023 10:11 PM")
            ComboBox1.Items.Insert(2, "4/10/2023 10:11 PM")
            ComboBox1.Items.Insert(3, "4/10/2023 10:11 PM")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'ToDo
        'Step 1: check complete all information
        Dim test As Boolean
        test = CheckData()

        'Step 2: add the databse
        'Create connection string
        If test Then
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

                MsgBox("Added Successfully")
            Catch ex As Exception
                MsgBox(Convert.ToString(ex))
            Finally
                con.Close()
            End Try
        End If

        'Step 3: message to check if want to add another or getout

        'Step 4: clean forms



    End Sub

    Private Function CheckData() As Boolean
        name = Me.TextBox3.Text
        lastName = Me.TextBox4.Text
        email = Me.TextBox6.Text
        obs = Me.TextBox5.Text
        data_plan = Me.DateTimePicker1.Value
        hour = Me.ComboBox1.SelectedItem
        phone = Me.TextBox8.Text

        If name! = "" Then
            Label9.Text = "Numele nu a fost setat"
            Return False
        End If
        If lastName! = "" Then
            Label9.Text = "Prenumele nu a fost setat"
            Return False
        End If
        If email! = "" Then
            Label9.Text = "E-mail-ul nu a fost setat"
            Return False
        End If
        If data_plan.Date < Now.Date Then
            Label9.Text = "Setează o dată validă"
            Return False
        End If
        If hour < Date.Now Then
            Label9.Text = "Ora nu a fost selectată"
            Return False
        End If
        If phone = "" Then
            Label9.Text = "Adaugă numărul de telefon"
            Return False
        End If

        Return True
    End Function

End Class