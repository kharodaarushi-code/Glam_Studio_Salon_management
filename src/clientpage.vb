Imports System.Data.OleDb

Public Class clientpage
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\GlamStudio\salon.accdb")
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim dt As DataTable

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            con.Open()
            cmd = New OleDbCommand("INSERT INTO client (First_Name, Last_Name, Email, Phone_Number, Address, Pre_Services, Total_Price) VALUES (@fn, @ln, @em, @ph, @add, @ps, @tp)", con)
            cmd.Parameters.AddWithValue("@fn", txtFirstName.Text)
            cmd.Parameters.AddWithValue("@ln", txtLastName.Text)
            cmd.Parameters.AddWithValue("@em", txtEmail.Text)
            cmd.Parameters.AddWithValue("@ph", txtPhone.Text)
            cmd.Parameters.AddWithValue("@add", txtAddress.Text)
            cmd.Parameters.AddWithValue("@ps", txtServices.Text)
            cmd.Parameters.AddWithValue("@tp", Convert.ToDecimal(txtTotalPrice.Text))
            cmd.ExecuteNonQuery()
            MessageBox.Show("Client Details Saved Successfully!")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
End Class
