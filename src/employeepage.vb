Private Sub employeepage_Load(sender As Object, e As EventArgs) Handles MyBase.Load 
LoadDgvemp() 
adp = New OleDbDataAdapter("select * from employee ORDER BY EmpID ASC", dbcon) 
ds = New DataSet() 
adp.Fill(ds) 
cmbempID.DataSource = ds.Tables(0) 
cmbempID.ValueMember = "EmpID" 
cmbempID.DisplayMember = "EmpID" 
End Sub 
Private Sub LoadDgvemp() 
Call pConnectDB() 
adp = New OleDbDataAdapter("Select * from employee", dbcon) 
dt = New DataTable() 
adp.Fill(dt) 
DataGridView1.DataSource = dt 
DataGridView1.Sort(DataGridView1.Columns("EmpID"), 
System.ComponentModel.ListSortDirection.Ascending) 
Call pDisconnectDB() 
End Sub 
Private Sub btnempinsert_Click(sender As Object, e As EventArgs) Handles 
btnempinsert.Click 
Dim psql As String 
Dim EID As String 
Call pConnectDB() 
psql = "select max(EmpID) from employee" 
cmd = New OleDbCommand(psql, dbcon) 
If IsDBNull(cmd.ExecuteScalar()) Then 
EID = 101 
Else 
EID = cmd.ExecuteScalar() + 1 
End If 
Dim sSql As String 
sSql = "Insert into employee values(" & EID & ",'" & txtempfirstname.Text & "','" & 
txtemplastname.Text & "','" & txtempdepartment.Text & "'," & txtempsalary.Text & ",'" & 
txtempemail.Text & "'," & txtempnumber.Text & ",'" & txtempaddress.Text & "')" 
cmd = New OleDbCommand(sSql, dbcon) 
cmd.ExecuteNonQuery() 
MessageBox.Show("Record Added Successfully") 
cmbempID.Text = "" 
txtempfirstname.Text = "" 
txtemplastname.Text = "" 
txtempdepartment.Text = "" 
txtempsalary.Text = "" 
txtempemail.Text = "" 
txtempnumber.Text = "" 
txtempaddress.Text = "" 
Call pDisconnectDB() 
Call LoadDgvemp() 
End Sub 
    Private Sub empID_SelectionChangeCommitted(sender As Object, e As EventArgs) 
Handles cmbempID.SelectionChangeCommitted 
        pConnectDB() 
        sql = "select * from employee where EmpID=" & cmbempID.SelectedValue & "" 
        cmd = New OleDbCommand(sql, dbcon) 
        dr = cmd.ExecuteReader() 
        dr.Read() 
        txtempfirstname.Text = dr.Item("First_Name") 
        txtemplastname.Text = dr.Item("Last_Name") 
        txtempdepartment.Text = dr.Item("Department") 
        txtempsalary.Text = dr.Item("Salary") 
        txtempemail.Text = dr.Item("Email") 
        txtempnumber.Text = dr.Item("Phone_Number") 
        txtempaddress.Text = dr.Item("Address") 
        pDisconnectDB() 
        Call LoadDgvemp() 
    End Sub 
 
    Private Sub btnempdelete_Click(sender As Object, e As EventArgs) Handles 
btnempdelete.Click 
        Call pConnectDB() 
        Dim Uid As Integer 
        Uid = cmbempID.SelectedValue 
 
        Dim sSql As String 
        sSql = "Delete from  employee where EmpID = " & Uid & "" 
        MessageBox.Show(sSql) 
        Dim cmd As OleDbCommand 
        cmd = New OleDbCommand(sSql, dbcon) 
        cmd.ExecuteNonQuery() 
        MessageBox.Show("Record Deleted Successfully") 
        cmbempID.Text = "" 
        txtempfirstname.Text = "" 
        txtemplastname.Text = "" 
        txtempdepartment.Text = "" 
        txtempsalary.Text = "" 
        txtempemail.Text = "" 
        txtempnumber.Text = "" 
        txtempaddress.Text = "" 
        Call pDisconnectDB() 
        Call LoadDgvemp() 
    End Sub 
 
    Private Sub btnempupdate_Click(sender As Object, e As EventArgs) Handles 
btnempupdate.Click 
        pConnectDB() 
        Dim Sql1 As String 
        Dim Uid1 As Integer 
        Uid1 = cmbempID.SelectedValue 
Sql1 = "Update employee Set First_Name = '" & txtempfirstname.Text & "', Last_Name 
= '" & txtemplastname.Text & "', Department = '" & txtempdepartment.Text & "', Salary = " & 
txtempsalary.Text & ", Email = '" & txtempemail.Text & "', Phone_Number = " & 
txtempnumber.Text & ", Address ='" & txtempaddress.Text & "'  where EmpID = " & Uid1 & 
"" 
MessageBox.Show(Sql1) 
Dim cmd As OleDbCommand 
cmd = New OleDbCommand(Sql1, dbcon) 
cmd.ExecuteNonQuery() 
MessageBox.Show("Record updated") 
cmbempID.Text = "" 
txtempfirstname.Text = "" 
txtemplastname.Text = "" 
txtempdepartment.Text = "" 
txtempsalary.Text = "" 
txtempemail.Text = "" 
txtempnumber.Text = "" 
txtempaddress.Text = "" 
Call pDisconnectDB() 
Call LoadDgvemp() 
End Sub 
End Class
