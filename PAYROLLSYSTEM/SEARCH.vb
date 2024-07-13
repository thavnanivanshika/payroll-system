Imports System.Data.SqlClient

Public Class SEARCH
    Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim employeeCode As String = TextBox8.Text.Trim()

        If IsNumeric(employeeCode) Then
            If EmployeeCodeExists(Convert.ToInt32(employeeCode)) Then
                GroupBox1.Visible = True
                Panel2.Visible = True
                Panel1.Visible = False
                FetchEmployeeData(employeeCode)
                ' Fetch department and add data to the salary table
            Else
                MessageBox.Show("Employee code does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Please enter a valid numeric employee code.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    Private Sub FetchEmployeeData(employeeCode As String)
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim query As String = "SELECT EMPLOYEE_CODE,EMPLOYEE_NAME, DEPARTMENT FROM EMPLOYEE1 WHERE EMPLOYEE_CODE = @EmployeeCode"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                connection.Open()
                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.HasRows Then
                        reader.Read()
                        Label11.Text = reader("EMPLOYEE_CODE").ToString()
                        Label13.Text = reader("EMPLOYEE_NAME").ToString()
                        Label15.Text = reader("DEPARTMENT").ToString()

                        GroupBox1.Visible = True
                    Else
                        MessageBox.Show("Employee not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        GroupBox1.Visible = False
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Function EmployeeCodeExists(employeeCode As Integer) As Boolean
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim query As String = "SELECT COUNT(*) FROM EMPLOYEE1 WHERE EMPLOYEE_CODE = @EmployeeCode"
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                connection.Open()
                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

    Private Function GetEmployeeDepartment(employeeCode As Integer) As String
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim query As String = "SELECT DEPARTMENT FROM EMPLOYEE1 WHERE EMPLOYEE_CODE = @EmployeeCode"
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                connection.Open()
                Dim department As Object = command.ExecuteScalar()
                If department IsNot Nothing Then
                    Return department.ToString()
                Else
                    Return String.Empty
                End If
            End Using
        End Using
    End Function

    Private Function SalaryDataExists(employeeCode As Integer) As Boolean
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim query As String = "SELECT COUNT(*) FROM SALARY WHERE EMPLOYEE_CODE = @EmployeeCode"
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                connection.Open()
                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

    Private Sub AddSalaryData(empCode As Integer, Department As String, Designation As String, BasicSalary As String, DA As String, HRA As String, Medical As String, Performance As String, Conveyance As String)
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()
                Dim sql As String = "INSERT INTO [dbo].[SALARY] ([EMPLOYEE_CODE], [DEPARTMENT], [DESIGNATION], [BASIC_SALARY], [DA], [HRA], [MEDICAL], [PERFORMANCE], [CONVEYANCE]) VALUES (@EmpCode, @Department, @Designation, @BasicSalary, @DA, @HRA, @Medical, @Performance, @Conveyance)"

                Using cmd As New SqlCommand(sql, con)
                    cmd.Parameters.AddWithValue("@EmpCode", empCode)
                    cmd.Parameters.AddWithValue("@Department", Department)
                    cmd.Parameters.AddWithValue("@Designation", Designation)
                    cmd.Parameters.AddWithValue("@BasicSalary", If(Decimal.TryParse(BasicSalary, New Decimal), Decimal.Parse(BasicSalary), DBNull.Value))
                    cmd.Parameters.AddWithValue("@DA", If(Decimal.TryParse(DA, New Decimal), Decimal.Parse(DA), DBNull.Value))
                    cmd.Parameters.AddWithValue("@HRA", If(Decimal.TryParse(HRA, New Decimal), Decimal.Parse(HRA), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Medical", If(Decimal.TryParse(Medical, New Decimal), Decimal.Parse(Medical), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Performance", If(Decimal.TryParse(Performance, New Decimal), Decimal.Parse(Performance), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Conveyance", If(Decimal.TryParse(Conveyance, New Decimal), Decimal.Parse(Conveyance), DBNull.Value))

                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Salary details added successfully")
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error inserting salary details: " & ex.Message)
        End Try
    End Sub

    Private Sub SEARCH_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GroupBox1.Visible = False
        Panel2.Visible = False
        If GroupBox1.Visible = True Then
            Panel1.Visible = False
        Else
            Panel1.Visible = True
        End If
        LoadCompanyDetails()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim employeeCode As String = TextBox8.Text.Trim()
        If IsNumeric(employeeCode) Then
            Dim empCode As Integer = Convert.ToInt32(employeeCode)
            If Not SalaryDataExists(empCode) Then
                Dim department As String = GetEmployeeDepartment(empCode)
                If Not String.IsNullOrEmpty(department) Then
                    AddSalaryData(empCode, department, TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, TextBox7.Text)
                Else
                    MessageBox.Show("Department not found for the given employee code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Salary data already exists for the given employee code.", "Data Exists", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            MessageBox.Show("Please enter a valid numeric employee code.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

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
                        Label18.Text = companyName
                        Label17.Text = companyPHONE
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
