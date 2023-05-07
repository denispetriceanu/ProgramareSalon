Imports System.Data.OleDb
Imports System.Globalization
Imports System.Net.Mail

Public Class Form2
    Dim orar(20) As String 'definim un array cu 24 de elemente, pentru orele de la 8: 00 la 18:00
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
        ArrHours()
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
            Label9.Text = ""
            ' complet withe possible hour in list
            ' get our with ocupated
            Dim con As OleDbConnection = New OleDbConnection
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=A:\CristinaProject\ProgramareSalon\SalonDB.accdb"
            con.Open()

            Dim query As String = "SELECT Ora FROM Programari where Data = #" + date2 + "#"
            Dim command As New OleDbCommand(query, con)

            Dim adapter As New OleDbDataAdapter(command)
            ' save data in array
            Dim dataset As New DataSet()
            adapter.Fill(dataset, "Programari")
            Dim table As DataTable = dataset.Tables("Programari")
            Dim array(table.Rows.Count - 1) As String

            For i As Integer = 0 To table.Rows.Count - 1
                Dim row As DataRow = table.Rows(i)
                For j As Integer = 0 To table.Columns.Count - 1
                    Dim timp As DateTime = DateTime.ParseExact(row(j), "h:mm:ss tt", CultureInfo.InvariantCulture)
                    Dim timpFormatat As String = timp.ToString("hh:mm tt")
                    array(i) = timpFormatat
                Next
            Next

            ComboBox1.Items.Clear()
            array.Reverse(orar)
            For Each ora As String In orar
                If Not array.Contains(ora) Then
                    'If Not ora Is Nothing Then
                    ComboBox1.Items.Insert(0, ora)
                    'End If
                End If
            Next
        Else
            Label9.Text = "Setează o dată validă"
        End If
    End Sub

    Private Function ArrHours()
        Dim startTime As New TimeSpan(8, 0, 0) ' Ora de inceput
        Dim endTime As New TimeSpan(18, 0, 0) ' Ora de sfarsit
        Dim interval As New TimeSpan(0, 30, 0) ' Intervalul de timp

        'Dim orar((CInt((endTime - startTime).TotalMinutes) \ CInt(interval.TotalMinutes)) + 1) As String


        Dim oraCurenta As TimeSpan = startTime
        For i As Integer = 0 To orar.Length - 1
            Dim timp As DateTime = DateTime.ParseExact(oraCurenta.ToString(), "HH:mm:ss", CultureInfo.InvariantCulture)
            Dim timpConvert As String = DateTime.Parse(timp.ToString).ToString("hh:mm tt")
            orar(i) = timpConvert
            oraCurenta = oraCurenta.Add(interval)
        Next
    End Function


    Private Function CompareDate()
        Dim date1 As DateTime = DateTime.ParseExact(DateTime.Today.ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture)
        Dim date2 As DateTime = DateTime.ParseExact(DateTimePicker1.Value().ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture)
        If date1 = date2 Then
            Return True
        Else
            Return False
        End If

    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'ToDo
        'Step 1: check complete all information
        Dim test As Boolean = CheckData()

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
        If test Then
            Dim result As DialogResult = MessageBox.Show("Mai adăugați alte programări?", "Programare", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                Me.TextBox3.Text = ""
                Me.TextBox4.Text = ""
                Me.TextBox6.Text = ""
                Me.TextBox5.Text = ""
                Me.TextBox8.Text = ""
            Else
                HomePage.Show()
                Me.Close()
            End If
        End If


    End Sub

    Private Function getTimeNow() As String
        Dim now As DateTime = DateTime.Now
        Dim hour As String = now.ToString("HH")
        Dim minute As String = now.ToString("mm")
        Dim period As String = now.ToString("tt")
        Return hour & ":" & minute & " " & period
    End Function

    Private Function CheckData() As Boolean
        name = Me.TextBox3.Text
        lastName = Me.TextBox4.Text
        email = Me.TextBox6.Text
        obs = Me.TextBox5.Text
        data_plan = Me.DateTimePicker1.Value
        hour = Me.ComboBox1.SelectedItem
        phone = Me.TextBox8.Text

        'Me.TextBox3.Text = "Test"
        'Me.TextBox4.Text = "Test"
        'Me.TextBox6.Text = "test@test.com"
        'Me.TextBox5.Text = "test-obs"
        'Me.TextBox8.Text = "0740663926"


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
        Dim time1 As DateTime = DateTime.ParseExact(getTimeNow(), "h:mm tt", CultureInfo.InvariantCulture)
        If Not hour = Nothing Then
            Dim time2 As DateTime = DateTime.ParseExact(hour, "hh\:mm tt", CultureInfo.InvariantCulture)
            If time2 < time1 And CompareDate() Then
                Label9.Text = "nu a fost selectată o ora valida"
                Return False
            End If
        Else
            Label9.Text = "Nu a fost selectată ora"
        End If

        If phone.Length <> 10 Then
            Label9.Text = "Adaugă numărul de telefon sau  introdu unul valid"
            Return False
        End If


        Return True

    End Function

End Class