Imports System.Data.SqlClient
Imports System.Reflection.Emit

Public Class SAL
    Private empCode As String
    Private empName As String
    Private Department As String
    Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
    Public Sub New(empCode As String, empName As String, Department As String)
        InitializeComponent()
        Me.empCode = empCode
        Me.empName = empName
        Me.Department = Department
        Label12.Text = empCode
        Label13.Text = empName
        Label14.Text = Department
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"

            Using con As New SqlConnection(connectionString)
                con.Open()
                Dim sql As String = "INSERT INTO [dbo].[SALARY] ([EMPLOYEE_CODE], [DEPARTMENT], [DESIGNATION], [BASIC_SALARY], [DA], [HRA], [MEDICAL], [PERFORMANCE], [CONVEYANCE]) VALUES (@EmpCode, @Department, @Designation, @BasicSalary, @DA, @HRA, @Medical, @Performance, @Conveyance)"

                Using cmd As New SqlCommand(sql, con)
                    cmd.Parameters.AddWithValue("@EmpCode", empCode)
                    cmd.Parameters.AddWithValue("@Department", Department)
                    cmd.Parameters.AddWithValue("@Designation", TextBox1.Text)
                    cmd.Parameters.AddWithValue("@BasicSalary", If(Decimal.TryParse(TextBox2.Text, New Decimal), Decimal.Parse(TextBox2.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@DA", If(Decimal.TryParse(TextBox3.Text, New Decimal), Decimal.Parse(TextBox3.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@HRA", If(Decimal.TryParse(TextBox4.Text, New Decimal), Decimal.Parse(TextBox4.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Medical", If(Decimal.TryParse(TextBox5.Text, New Decimal), Decimal.Parse(TextBox5.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Performance", If(Decimal.TryParse(TextBox6.Text, New Decimal), Decimal.Parse(TextBox6.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@Conveyance", If(Decimal.TryParse(TextBox7.Text, New Decimal), Decimal.Parse(TextBox7.Text), DBNull.Value))

                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Salary details added successfully")
                    Me.Close()
                    HOME.Show()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error inserting salary details: " & ex.Message)
        End Try
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        ADD.Show()
    End Sub

    Private Sub SAL_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCompanyDetails()
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

