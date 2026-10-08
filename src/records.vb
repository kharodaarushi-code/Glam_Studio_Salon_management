Imports System.Data.OleDb

Public Class records
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\GlamStudio\salon.accdb")

    Private Sub btnCalculateProfit_Click(sender As Object, e As EventArgs) Handles btnCalculateProfit.Click
        Dim totalIncome As Decimal = GetTotal("SELECT SUM(Income) FROM income WHERE Month=" & txtMonth.Text & " AND Year=" & txtYear.Text)
        Dim totalExpense As Decimal = GetTotal("SELECT SUM(Total_Amount) FROM expense WHERE Month=" & txtMonth.Text & " AND Year=" & txtYear.Text)
        Dim totalSalaries As Decimal = GetTotal("SELECT SUM(Salary) FROM employee")

        Dim netProfit As Decimal = totalIncome - (totalExpense + totalSalaries)

        If netProfit >= 0 Then
            MessageBox.Show("Net Profit for the Month: Rs. " & netProfit.ToString(), "Profit Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Net Loss for the Month: Rs. " & Math.Abs(netProfit).ToString(), "Loss Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Function GetTotal(query As String) As Decimal
        Dim result As Decimal = 0
        Try
            con.Open()
            Dim cmd As New OleDbCommand(query, con)
            Dim res = cmd.ExecuteScalar()
            If Not IsDBNull(res) AndAlso res IsNot Nothing Then
                result = Convert.ToDecimal(res)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            con.Close()
        End Try
        Return result
    End Function
End Class
