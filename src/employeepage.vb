Imports System.Data.OleDb

Public Class employeepage
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\GlamStudio\salon.accdb")
    Dim cmd As OleDbCommand

    Private Sub btnAddEmp_Click(sender As Object, e As EventArgs) Handles btnAddEmp.Click
        Try
            con.Open()
            cmd = New OleDbCommand("INSERT INTO employee (First_Name, Last_Name, Department, Salary, Email, Phone_Number, Address) VALUES (@fn, @ln, @dept, @sal, @em, @ph, @add)", con)
            cmd.Parameters.AddWithValue("@fn", txtEmpFN.Text)
            cmd.Parameters.AddWithValue("@ln", txtEmpLN.Text)
            cmd.Parameters.AddWithValue("@dept", cmbDept.Text)
            cmd.Parameters.AddWithValue("@sal", Convert.ToDecimal(txtSalary.Text))
            cmd.Parameters.AddWithValue("@em", txtEmpEmail.Text)
            cmd.Parameters.AddWithValue("@ph", txtEmpPhone.Text)
            cmd.Parameters.AddWithValue("@add", txtEmpAddress.Text)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Employee Record Created Successfully!")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
End Class
