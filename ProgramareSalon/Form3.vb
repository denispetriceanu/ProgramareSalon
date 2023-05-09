Imports System.Data.OleDb

Public Class Form3
    Dim con As OleDbConnection
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' adaugare buton stergere
        Dim deleteButtonColumn As New DataGridViewButtonColumn()
        deleteButtonColumn.HeaderText = "Stergere"
        deleteButtonColumn.Text = "Stergere"
        deleteButtonColumn.Name = "deleteButtonColumn"
        deleteButtonColumn.DefaultCellStyle.ForeColor = Color.Red
        deleteButtonColumn.UseColumnTextForButtonValue = True

        DataGridView1.Columns.Add(deleteButtonColumn)


        ' preluarea datelor din db
        ' realizarea conexiunii
        con = New OleDbConnection
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=A:\CristinaProject\ProgramareSalon\SalonDB.accdb"
        con.Open()

        Dim query As String = "SELECT Id, Nume, Prenume, Telefon, Data, Ora FROM Programari"
        Dim command As New OleDbCommand(query, con)

        Dim adapter As New OleDbDataAdapter(command)
        Dim dataSet As New DataSet()

        adapter.Fill(dataSet)

        DataGridView1.DataSource = dataSet.Tables(0)
        DataGridView1.Columns("Ora").DefaultCellStyle.Format = "hh:mm tt"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        HomePage.Show()
        Me.Hide()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("deleteButtonColumn").Index AndAlso e.RowIndex >= 0 Then
            Dim confirmResult As DialogResult = MessageBox.Show("Sunteti sigur ca doriti sa stergeti aceasta inregistrare?", "Confirmare stergere", MessageBoxButtons.YesNo)
            If confirmResult = DialogResult.Yes Then
                Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
                Dim idProgramare As Integer = row.Cells("ID").Value
                Dim query As String = "DELETE FROM Programari WHERE ID = @idProgramare"
                Dim command As New OleDbCommand(query, con)
                command.Parameters.AddWithValue("@idProgramare", idProgramare)
                command.ExecuteNonQuery()
                DataGridView1.Rows.Remove(row)
            End If
        End If
    End Sub


End Class