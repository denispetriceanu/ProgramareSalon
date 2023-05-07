Imports System.Data.OleDb

Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Adăugarea coloanelor la obiectul DataGridView
        'DataGridView1.Columns.Add("Nume", "Nume")
        'DataGridView1.Columns.Add("Prenume", "Prenume")
        'DataGridView1.Columns.Add("Telefon", "Telefon")
        'DataGridView1.Columns.Add("Data", "Data")
        'DataGridView1.Columns.Add("Ora", "Ora")

        ' preluarea datelor din db
        ' realizarea conexiunii
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=A:\CristinaProject\ProgramareSalon\SalonDB.accdb"
        con.Open()

        Dim query As String = "SELECT Nume, Prenume, Telefon, Data, Ora FROM Programari"
        Dim command As New OleDbCommand(query, con)

        Dim adapter As New OleDbDataAdapter(command)
        Dim dataSet As New DataSet()

        adapter.Fill(dataSet)

        DataGridView1.DataSource = dataSet.Tables(0)

        'For i As Integer = 1 To 5
        '    ' Adăugarea unui rând nou cu date
        '    Dim row As New DataGridViewRow()
        '    row.CreateCells(DataGridView1)
        '    row.Cells(0).Value = "Popescu"
        '    row.Cells(1).Value = "Alexandru"
        '    row.Cells(2).Value = 25

        '    ' Adaugarea randului nou in tabel
        '    DataGridView1.Rows.Add(row)
        'Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        HomePage.Show()
        Me.Hide()
    End Sub

End Class