Public Class records 
Dim dbcon   As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data 
Source=C:\Users\ancha\OneDrive\Documents\finalprojectDB.accdb") 
    Private Sub records_Load(sender As Object, e As EventArgs) Handles MyBase.Load 
        Me.Size = New Size(1024, 768) ' Adjust size here 
        dbcon.Open() 
        LoadIncome() 
        LoadSalariesOnly() 
        LoadExpense() 
        dbcon.Close() 
        ' Add month numbers (1–12) to ComboBox 
        cmbmonth.Items.Add("All") 
        For i As Integer = 1 To 12 
            cmbmonth.Items.Add(i.ToString()) 
        Next 
        cmbmonth.SelectedIndex = 0 ' Set default to "All" 
        ' Add "All" to year ComboBox 
        cmbyear.Items.Add("All") 
        For year As Integer = DateTime.Now.Year To DateTime.Now.Year -                 3 Step -1 
        cmbyear.Items.Add(year.ToString()) 
        Next 
        cmbyear.SelectedIndex = 0 ' Set default to "All" 
    End Sub 
    Private Sub btnshow_Click(sender As Object, e As EventArgs) Handles btnshow.Click 
        Dim selectedMonth As Integer = 0 
        Dim selectedYear As Integer = 0 
        If cmbmonth.SelectedIndex > 0 Then 
            selectedMonth = Integer.Parse(cmbmonth.SelectedItem.ToString()) 
        End If 
        If cmbyear.SelectedIndex > 0 Then 
            selectedMonth = Integer.Parse(cmbmonth.SelectedItem.ToString()) 
        End If 
        If dbcon.State = ConnectionState.Closed Then dbcon.Open() 
        LoadFilteredData("income", DataGridView1, selectedMonth, selectedYear) 
        LoadFilteredData("expense", DataGridView3, selectedMonth, selectedYear) 
        dbcon.Close() 
    End Sub 
    Private Sub LoadFilteredData(tableName As String, dgv As DataGridView, Optional 
filterMonth As Integer = 0, Optional filterYear As Integer = 0) 
        Dim query As String = $"SELECT * FROM {tableName}" 
        Dim conditions As New List(Of String) 
        Dim adapter As OleDbDataAdapter 
        If filterMonth > 0 Then 
            conditions.Add("[Month] = ?") 
        End If 
        If filterYear > 0 Then 
            conditions.Add("[Year] = ?") 
        End If 
        If conditions.Count > 0 Then
             query &= " WHERE " & String.Join(" AND ", conditions) 
        End If 
        adapter = New OleDbDataAdapter(query, dbcon) 
        If filterMonth > 0 Then 
            adapter.SelectCommand.Parameters.AddWithValue("?", filterMonth) 
        End If 
        If filterYear > 0 Then 
            adapter.SelectCommand.Parameters.AddWithValue("?", filterYear) 
        End If 
        Dim dt As New DataTable() 
        adapter.Fill(dt) 
        dgv.DataSource = dt 
    End Sub 
    Private Sub LoadIncome() 
        Dim query As String = "SELECT * FROM income" 
        Dim adapter As New OleDbDataAdapter(query, dbcon) 
        Dim dt As New DataTable() 
        adapter.Fill(dt) 
        DataGridView1.DataSource = dt 
    End Sub 
    Private Sub LoadSalariesOnly() 
        Try 
            Call pConnectDB() 
            ' Select first_name, last_name, and salary separately 
            Dim query As String = "SELECT First_Name, Last_Name, Salary FROM employee" 
            Dim adp As New OleDbDataAdapter(query, dbcon) 
            Dim dt As New DataTable() 
            adp.Fill(dt) 
            DataGridView2.DataSource = dt 
            Call pDisconnectDB() 
        Catch ex As Exception 
            MessageBox.Show("Error loading salaries: " & ex.Message) 
        End Try 
    End Sub 
    Private Sub LoadExpense() 
        Dim query As String = "SELECT * FROM expense" 
        Dim adapter As New OleDbDataAdapter(query, dbcon) 
        Dim dt As New DataTable() 
        adapter.Fill(dt) 
        DataGridView3.DataSource = dt 
    End Sub 
    Private Sub btncalculate_Click(sender As Object, e As EventArgs) Handles 
btncalculate.Click 
        Try 
            ' Total Income from DataGridView1 only (filtered view) 
            Dim totalIncome As Decimal = 0 
            For Each row As DataGridViewRow In DataGridView1.Rows 
                If Not row.IsNewRow Then 
                    totalIncome += Convert.ToDecimal(row.Cells("income").Value) 
                End If
            Next 
txtincome.Text = totalIncome.ToString("F2") 
' Salary 
If dbcon.State = ConnectionState.Closed Then dbcon.Open() 
Dim salaryCmd As New OleDbCommand("SELECT SUM(salary) FROM employee",      
dbcon) 
Dim salaryResult As Object = salaryCmd.ExecuteScalar() 
txtsalary.Text=If(IsDBNull(salaryResult),"0",Convert.ToDecimal(salaryResult).ToString("F2
")) 
' Expense 
Dim expenseCmd As New OleDbCommand("SELECT SUM(Total_Amount) FROM 
expense", dbcon) 
Dim 
expenseResult 
As 
Object 
= 
expenseCmd.ExecuteScalar() 
txtexpense.Text=If(IsDBNull(expenseResult),"0",Convert.ToDecimal(expenseResult).ToStri
ng("F2")) 
' Profit = Income - Salary - Expense 
Dim totalSalary As Decimal = Convert.ToDecimal(txtsalary.Text) 
Dim totalExpense As Decimal = Convert.ToDecimal(txtexpense.Text) 
Dim profit As Decimal = totalIncome - totalSalary - totalExpense 
txtprofit.Text = profit.ToString("F2") 
'Show a dialog depending on profit status 
If profit > 0 Then 
MessageBox.Show("Congratulations! You are in profit this month.", "Profit Status", 
MessageBoxButtons.OK, MessageBoxIcon.Information) 
Else 
MessageBox.Show("Not in profit this month. Consider reducing expenses or 
boosting income.", "Profit Status", MessageBoxButtons.OK, MessageBoxIcon.Warning) 
End If 
dbcon.Close() 
Catch ex As Exception 
MessageBox.Show("Error: " & ex.Message) 
If dbcon.State = ConnectionState.Open Then dbcon.Close() 
End Try 
End Sub 
End Class
 
            
