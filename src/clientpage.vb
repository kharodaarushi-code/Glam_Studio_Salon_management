Private Sub clientpage_Load(sender As Object, e As EventArgs) Handles    MyBase.Load 
LoadDgvclient() 
adp = New OleDbDataAdapter("select * from client", dbcon) 
ds = New DataSet() 
adp.Fill(ds) 
cmbclientID.DataSource = ds.Tables(0) 
cmbclientID.ValueMember = "ClientID" 
cmbclientID.DisplayMember = "ClientID" 
End Sub 
Private Sub LoadDgvclient() 
Call pConnectDB() 
adp = New OleDbDataAdapter("Select * from client", dbcon) 
dt = New DataTable() 
adp.Fill(dt) 
DataGridView1.DataSource = dt 
Call pDisconnectDB() 
End Sub 
Private Sub btnclientinsert_Click(sender As Object, e As EventArgs) Handles   
btnclientinsert.Click 
Dim psql As String 
Dim CID As String 
Call pConnectDB() 
psql = "select max(ClientID) from client" 
cmd = New OleDbCommand(psql, dbcon) 
If IsDBNull(cmd.ExecuteScalar()) Then 
CID = 101 
Else 
CID = cmd.ExecuteScalar() + 1 
End If 
Dim sSql As String = "INSERT INTO client (ClientID, First_Name,    
Phone_Number, Address) " &"VALUES (?, ?, ?, ?, ?, ?)" 
cmd = New OleDbCommand(sSql, dbcon) 
' Add values to parameters (in order) 
cmd.Parameters.AddWithValue("?", CID) 
cmd.Parameters.AddWithValue("?", txtclientfirstname.Text.Trim()) 
cmd.Parameters.AddWithValue("?", txtclientlastname.Text.Trim()) 
cmd.Parameters.AddWithValue("?", txtclientemail.Text.Trim()) 
Last_Name, Email, 
cmd.Parameters.AddWithValue("?",Convert.ToInt64(txtclientnumber.Text.Trim())) 
cmd.Parameters.AddWithValue("?", txtclientaddress.Text.Trim()) 
cmd.ExecuteNonQuery() 
MessageBox.Show("Record Added Successfully") 
cmbclientID.Text = "" 
txtclientfirstname.Text = "" 
txtclientlastname.Text = "" 
txtclientemail.Text = "" 
txtclientnumber.Text = ""
txtclientaddress.Text = "" 
Call pDisconnectDB() 
LoadDgvclient() 
End Sub 
Private Sub clientID_SelectionChangeCommitted(sender As Object, e As EventArgs) 
Handles cmbclientID.SelectionChangeCommitted 
pConnectDB() 
sql = "select * from client where ClientID=" & cmbclientID.SelectedValue    & "" 
cmd = New OleDbCommand(sql, dbcon) 
dr = cmd.ExecuteReader() 
dr.Read() 
txtclientfirstname.Text = dr.Item("First_Name") 
txtclientlastname.Text = dr.Item("Last_Name") 
txtclientemail.Text = dr.Item("Email") 
txtclientnumber.Text = dr.Item("Phone_Number") 
txtclientaddress.Text = dr.Item("Address") 
pDisconnectDB() 
Call LoadDgvclient() 
End Sub 
Private Sub btnclientdelete_Click(sender As Object, e As EventArgs) Handles 
btnclientdelete.Click 
Call pConnectDB() 
Dim Uid As Integer 
Uid = cmbclientID.SelectedValue 
Dim sSql As String 
sSql = "Delete from  client where ClientID = " & Uid & "" 
MessageBox.Show(sSql) 
Dim cmd As OleDbCommand 
cmd = New OleDbCommand(sSql, dbcon) 
cmd.ExecuteNonQuery() 
MessageBox.Show("Record Deleted Successfully") 
cmbclientID.Text = "" 
txtclientfirstname.Text = "" 
txtclientlastname.Text = "" 
txtclientemail.Text = "" 
txtclientnumber.Text = "" 
txtclientaddress.Text = "" 
Call pDisconnectDB() 
Call LoadDgvclient() 
End Sub 
Private Sub btnclientupdate_Click(sender As Object, e As EventArgs) Handles   
btnclientupdate.Click 
pConnectDB() 
Dim Sql1 As String 
Dim Uid1 As Integer 
Uid1 = cmbclientID.SelectedValue 
Sql1 = "Update client Set First_Name = '" & txtclientfirstname.Text & "', Last_Name = '" 
& txtclientlastname.Text & "', Email = '" & txtclientemail.Text & "', Phone_Number = " & 
txtclientnumber.Text & ", Address ='" & txtclientaddress.Text & "'  where ClientID = " & Uid1
& "" 
        MessageBox.Show(Sql1) 
        Dim cmd As OleDbCommand 
        cmd = New OleDbCommand(Sql1, dbcon) 
        cmd.ExecuteNonQuery() 
        MessageBox.Show("Record updated") 
        cmbclientID.Text = "" 
        txtclientfirstname.Text = "" 
        txtclientlastname.Text = "" 
        txtclientemail.Text = "" 
        txtclientnumber.Text = "" 
        txtclientaddress.Text = "" 
        pDisconnectDB() 
        LoadDgvclient() 
    End Sub 
    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click 
        If txtsearch.Text.Trim() = "" Then 
            MessageBox.Show("Please enter a name to search.") 
            Return 
        End If 
        pConnectDB() 
        Dim searchName As String = txtsearch.Text.Trim() 
Dim sqlSearch As String = "SELECT * FROM client WHERE First_Name LIKE '%" & 
searchName & "%'" 
        adp = New OleDbDataAdapter(sqlSearch, dbcon) 
        dt = New DataTable() 
        adp.Fill(dt) 
        If dt.Rows.Count > 0 Then 
            DataGridView1.DataSource = dt 
        Else 
            MessageBox.Show("No records found with that name.") 
            DataGridView1.DataSource = Nothing 
        End If 
        pDisconnectDB() 
    End Sub 
    Private Sub txtsearch_TextChanged(sender As Object, e As EventArgs) Handles 
txtsearch.TextChanged 
        If txtsearch.Text.Trim() = "" Then 
            LoadDgvclient() ' reload full list  
        End If 
    End Sub 
    Private Sub btnpreservices_Click(sender As Object, e As EventArgs) Handles   
btnPreServices.Click 
        If cmbclientID.SelectedIndex = -1 Then 
         MessageBox.Show("Please select a Client ID first.") 
          Return 
        End If 
        Dim selectedClientID As Integer = 
Convert.ToInt32(cmbclientID.SelectedValue) ' Use .Text if not bound 
        Dim query As String = "SELECT Pre_Services, Total_Price FROM client WHERE ClientID = ?" 
Dim cmd As New OleDbCommand(query, dbcon) 
cmd.Parameters.AddWithValue("?", selectedClientID) 
Try 
dbcon.Open() 
Dim reader As OleDbDataReader = cmd.ExecuteReader() 
If reader.Read() Then 
Dim services As String = If(reader.IsDBNull(reader.GetOrdinal("Pre_Services")), 
"", reader("Pre_Services").ToString()) 
Dim total As String = If(reader.IsDBNull(reader.GetOrdinal("Total_Price")), "", 
reader("Total_Price").ToString()) 
MessageBox.Show("Previous Services: " & services & vbCrLf & "Total Cost: ₹" & 
total, "Client History") 
Else 
MessageBox.Show("No records found for the selected Client ID.") 
End If 
reader.Close() 
Catch ex As Exception 
MessageBox.Show("Error fetching data: " & ex.Message) 
Finally 
dbcon.Close() 
End Try 
End Sub 
End Class 
