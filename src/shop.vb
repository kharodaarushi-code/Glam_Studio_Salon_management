 Private Sub shop_Load(sender As Object, e As EventArgs) Handles MyBase.Load 
        Call pConnectDB() 
        adp = New OleDbDataAdapter() 
        ds = New DataSet() 
        adp.SelectCommand = New OleDbCommand("SELECT * FROM shop",  dbcon) 
        adp.Fill(ds) 
        lstproducts.DataSource = ds.Tables(0) 
        lstproducts.ValueMember = "ItemID" 
        lstproducts.DisplayMember = "Item_Name" 
        Call pDisconnectDB() 
    End Sub 
    Private Sub btnshopadd_Click(sender As Object, e As EventArgs) Handles 
btnshopadd.Click 
        lstprice.Items.Clear() 
        If dbcon.State = ConnectionState.Closed Then 
            dbcon.Open() 
        End If 
        ' Loop through selected items 
        For Each selectedItem As Object In lstproducts.SelectedItems 
            Dim drv As DataRowView = CType(selectedItem, DataRowView) 
            Dim itemName As String = drv("Item_Name").ToString() 
            ' Prepare and run query to get price for the selected product 
            Dim query As String = "SELECT Item_Price FROM shop WHERE Item_Name = ?" 
            Dim cmd As New OleDbCommand(query, dbcon) 
            cmd.Parameters.AddWithValue("?", itemName) 
            Dim dr As OleDbDataReader = cmd.ExecuteReader() 
            While dr.Read() 
                lstprice.Items.Add(itemName & " - ₹" & dr("Item_Price").ToString()) 
            End While 
            dr.Close() 
        Next 
        dbcon.Close() 
    End Sub 
    Private Function GetNextIncomeID() As Integer 
        Dim newID As Integer = 1 
        Dim cmd As New OleDbCommand("SELECT MAX([Income_ID]) FROM income", 
dbcon) 
        Dim result = cmd.ExecuteScalar() 
        If Not IsDBNull(result) Then 
            newID = Convert.ToInt32(result) + 1 
        End If 
        Return newID 
    End Function 
    Private Sub btnshopcost_Click(sender As Object, e As EventArgs) Handles 
btnshopcost.Click 
        Dim total As Double = 0 
        For Each item As String In lstprice.Items 
            ' Split the string to get the price part after "₹" 
            Dim parts() As String = item.Split("₹") 
If parts.Length = 2 Then 
Dim priceStr As String = parts(1).Trim() 
Dim price As Double 
If Double.TryParse(priceStr, price) Then 
total += price 
End If 
End If 
Next 
' Show total in a textbox 
txtshopcost.Text = "₹" & total.ToString("F2") 
' Insert total into income table 
Try 
If dbcon.State = ConnectionState.Closed Then dbcon.Open() 
Dim totalValue As Decimal 
If Decimal.TryParse(txtshopcost.Text.Replace("₹", "").Trim(), totalValue) Then 
Dim IncomeID As Integer = GetNextIncomeID()  'id generation 
' Get today's date values 
Dim today As Date = Date.Today 
Dim currentMonth As String = today.Month 
Dim currentYear As Integer = today.Year 
' Insert income, type, date, month, year 
Dim insertCmd As New OleDbCommand("INSERT INTO income ([Income_ID], 
[Income], [Type], [Date], [Month], [Year]) VALUES (?, ?, ?, ?, ?, ?)", dbcon) 
insertCmd.Parameters.AddWithValue("?", IncomeID) 
insertCmd.Parameters.AddWithValue("?", totalValue) 
insertCmd.Parameters.AddWithValue("?", "product") 
insertCmd.Parameters.AddWithValue("?", today) 
insertCmd.Parameters.AddWithValue("?", currentMonth) 
insertCmd.Parameters.AddWithValue("?", currentYear) 
insertCmd.ExecuteNonQuery() 
Else 
MessageBox.Show("Failed to parse total cost value.") 
End If 
dbcon.Close() 
Catch ex As Exception 
MessageBox.Show("Error inserting income: " & ex.Message) 
If dbcon.State = ConnectionState.Open Then dbcon.Close() 
End Try 
End Sub 
Private Sub btnshopdelete_Click(sender As Object, e As EventArgs) Handles 
btnshopdelete.Click 
' Loop backwards to safely remove multiple items 
For i As Integer = lstprice.SelectedIndices.Count - 1 To 0 Step -1 
lstprice.Items.RemoveAt(lstprice.SelectedIndices(i)) 
Next 
End Sub 
End Class
