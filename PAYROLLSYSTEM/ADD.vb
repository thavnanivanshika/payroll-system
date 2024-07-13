Imports System.Data.SqlClient

Public Class ADD
    ' Define a function to generate a unique 5-digit employee code
    Private connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
    Private Function GenerateEmployeeCode() As String
        Dim rnd As New Random()
        Dim code As String
        Do
            code = rnd.Next(10000, 99999).ToString() ' Generate a random 5-digit number
        Loop While IsEmployeeCodeExists(code)
        Return code
    End Function

    ' Check if the generated employee code already exists in the database
    Private Function IsEmployeeCodeExists(code As String) As Boolean
        Dim con As SqlConnection = New SqlConnection("Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False")
        Dim cmd As SqlCommand = New SqlCommand("SELECT COUNT(*) FROM EMPLOYEE WHERE EMPLOYEE_CODE = @Code", con)
        cmd.Parameters.AddWithValue("@Code", code)
        con.Open()
        Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
        con.Close()
        Return count > 0
    End Function

    Private Sub ADD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Populate the ComboBox with shift timings
        ComboBox3.Items.Add("9-12")
        ComboBox3.Items.Add("12-4")
        ComboBox3.Items.Add("4-7")
        LoadCompanyDetails()

        ' Add event handlers for all input controls
        AddHandler TextBox1.TextChanged, AddressOf CheckFields
        AddHandler ComboBox1.SelectedIndexChanged, AddressOf CheckFields
        AddHandler TextBox3.TextChanged, AddressOf CheckFields
        AddHandler DateTimePicker1.ValueChanged, AddressOf CheckFields
        AddHandler TextBox4.TextChanged, AddressOf CheckFields
        AddHandler TextBox8.TextChanged, AddressOf CheckFields
        AddHandler DateTimePicker2.ValueChanged, AddressOf CheckFields
        AddHandler ComboBox3.SelectedIndexChanged, AddressOf CheckFields
        AddHandler TextBox5.TextChanged, AddressOf CheckFields
        AddHandler TextBox9.TextChanged, AddressOf CheckFields
        AddHandler TextBox10.TextChanged, AddressOf CheckFields
        AddHandler ComboBox2.SelectedIndexChanged, AddressOf CheckFields
        AddHandler TextBox11.TextChanged, AddressOf CheckFields
        AddHandler ComboBox4.SelectedIndexChanged, AddressOf CheckFields

        ' Initially check fields
        CheckFields()
    End Sub

    Private Sub CheckFields()
        ' Check if all required fields are filled
        If Not String.IsNullOrEmpty(TextBox1.Text) AndAlso
           ComboBox1.SelectedIndex <> -1 AndAlso
           Not String.IsNullOrEmpty(TextBox3.Text) AndAlso
           Not String.IsNullOrEmpty(TextBox4.Text) AndAlso
           Not String.IsNullOrEmpty(TextBox8.Text) AndAlso
           ComboBox3.SelectedIndex <> -1 AndAlso
           Not String.IsNullOrEmpty(TextBox5.Text) AndAlso
           Not String.IsNullOrEmpty(TextBox9.Text) AndAlso
           Not String.IsNullOrEmpty(TextBox10.Text) AndAlso
           ComboBox2.SelectedIndex <> -1 AndAlso
           Not String.IsNullOrEmpty(TextBox11.Text) AndAlso
           ComboBox4.SelectedIndex <> -1 Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        HOME.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim empCode As String = GenerateEmployeeCode()
        Dim empName As String = TextBox1.Text
        Dim Department As String = ComboBox4.SelectedItem.ToString()
        Dim con As SqlConnection = New SqlConnection("Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False")

        Dim cmd As SqlCommand = New SqlCommand("INSERT INTO [dbo].[EMPLOYEE1]
           ([EMPLOYEE_CODE]
           ,[EMPLOYEE_NAME]
           ,[GENDER]
           ,[ADDRESS]
           ,[DOB]
           ,[CITY]
           ,[QUALIFICATION]
           ,[DOJ]
           ,[SHIFT]
           ,[PINCODE]
           ,[PHONE]
           ,[EMAIL]
           ,[PF_NUMBER]
           ,[ESI_NUMBER]
           ,[DEPARTMENT])
       VALUES
       (@EmpCode, @EmpName, @Gender, @Address, @DOB, @City, @Qualification, @DOJ, @Shift, @Pincode, @Phone, @Email, @PFNumber, @ESINumber,@DEPARTMENT)", con)
        cmd.Parameters.AddWithValue("@EmpCode", empCode)
        cmd.Parameters.AddWithValue("@EmpName", TextBox1.Text)
        cmd.Parameters.AddWithValue("@Gender", ComboBox1.SelectedItem.ToString())
        cmd.Parameters.AddWithValue("@Address", TextBox3.Text)
        cmd.Parameters.AddWithValue("@DOB", DateTimePicker1.Value)
        cmd.Parameters.AddWithValue("@City", TextBox4.Text)
        cmd.Parameters.AddWithValue("@Qualification", TextBox8.Text)
        cmd.Parameters.AddWithValue("@DOJ", DateTimePicker2.Value)
        cmd.Parameters.AddWithValue("@Shift", TimeSpan.Parse(GetSelectedShiftTime()))
        cmd.Parameters.AddWithValue("@Pincode", TextBox5.Text)
        cmd.Parameters.AddWithValue("@Phone", TextBox9.Text)
        cmd.Parameters.AddWithValue("@Email", TextBox10.Text)
        If ComboBox2.SelectedItem IsNot Nothing Then
            If ComboBox2.SelectedItem.ToString() = "YES" Then
                ' Assuming cmd is your SQL command object
                cmd.Parameters.AddWithValue("@PFNumber", TextBox2.Text)
            Else
                ' Assuming cmd is your SQL command object
                cmd.Parameters.AddWithValue("@PFNumber", DBNull.Value)
            End If
        End If
        ' You can set this value according to your logic
        cmd.Parameters.AddWithValue("@ESINumber", TextBox11.Text)
        cmd.Parameters.AddWithValue("@DEPARTMENT", ComboBox4.SelectedItem.ToString()) ' You can set this value according to your logic
        con.Open()
        cmd.ExecuteNonQuery()
        Dim salaryForm As New SAL(empCode, empName, Department)
        con.Close()
        MessageBox.Show("Employee added successfully")
        salaryForm.Show()
        Me.Close()
    End Sub

    Private Function GetSelectedShiftTime() As String
        ' Retrieve the selected shift timing from the ComboBox
        Select Case ComboBox3.SelectedItem.ToString()
            Case "9-12"
                Return "09:00:00" ' Start time of the shift
            Case "12-4"
                Return "12:00:00"
            Case "4-7"
                Return "16:00:00"
            Case Else
                Return "00:00:00" ' Default value
        End Select
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        HOME.Show()
    End Sub
    Private Sub LoadCompanyDetails()
        ' SQL query to get the company name and address
        Dim query As String = "SELECT [COMPANY NAME], [PHONE] FROM [dbo].[COMPANY]"

        ' Create a connection and command object
        Using connection As New SqlConnection(connectionString),
              command As New SqlCommand(query, connection)
            Try
                ' Open the connection
                connection.Open()

                ' Execute the command and get the data reader
                Using reader As SqlDataReader = command.ExecuteReader()
                    ' Check if there is a row to read
                    If reader.Read() Then
                        ' Get the name and address from the reader
                        Dim companyName As String = reader("COMPANY NAME").ToString()
                        Dim companyPHONE As String = reader("PHONE").ToString()

                        ' Set the labels' text to the company details
                        Label19.Text = companyName
                        Label18.Text = companyPHONE
                    Else
                        ' Handle the case where no data is returned
                        MessageBox.Show("No company details found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using

            Catch ex As Exception
                ' Handle any errors that might have occurred
                MessageBox.Show("An error occurred while fetching company details: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

End Class
