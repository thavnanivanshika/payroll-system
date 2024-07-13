Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class FINAL
    Private _selectedMonth As String
    Private _selectedYear As Integer
    Private connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
    Public Sub New(selectedMonth As String, selectedYear As Integer)
        InitializeComponent()
        _selectedMonth = selectedMonth
        _selectedYear = selectedYear
        Label18.Text = _selectedMonth
        Label19.Text = _selectedYear.ToString()  ' Ensure conversion to string for display
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        REPORT.Show()
    End Sub

    Private Sub FINAL_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label18.Visible = False
        Label19.Visible = False
        LoadCompanyDetails()
        Dim selectedMonth As String = Label18.Text
        Dim selectedYear As Integer = 0
        If Not Integer.TryParse(Label19.Text, selectedYear) Then
            MessageBox.Show("Invalid year format. Please enter a valid number.")
            Return  ' Exit the function if conversion fails
        End If
        Dim hrData As Tuple(Of Integer, Integer) = GetHREmployeeCount(selectedMonth, selectedYear)

        ' Check for data retrieval success (optional)
        If hrData Is Nothing Then
            MessageBox.Show("Error retrieving data")
        Else
            Label6.Text = hrData.Item1.ToString()  ' Display HR employee count
            Label15.Text = hrData.Item2.ToString()  ' Display total working days
        End If
        Dim PRData As Tuple(Of Integer, Integer) = GetPRODUCTIONEmployeeCount(selectedMonth, selectedYear)

        ' Check for data retrieval success (optional)
        If PRData Is Nothing Then
            MessageBox.Show("Error retrieving data")
        Else
            Label7.Text = PRData.Item1.ToString()  ' Display HR employee count
            Label14.Text = PRData.Item2.ToString()  ' Display total working days
        End If
        Dim SLData As Tuple(Of Integer, Integer) = GetSALESEmployeeCount(selectedMonth, selectedYear)

        ' Check for data retrieval success (optional)
        If SLData Is Nothing Then
            MessageBox.Show("Error retrieving data")
        Else
            Label8.Text = SLData.Item1.ToString()  ' Display HR employee count
            Label13.Text = SLData.Item2.ToString()  ' Display total working days
        End If
        Dim FIData As Tuple(Of Integer, Integer) = GetFINANCEEmployeeCount(selectedMonth, selectedYear)

        ' Check for data retrieval success (optional)
        If FIData Is Nothing Then
            MessageBox.Show("Error retrieving data")
        Else
            Label9.Text = FIData.Item1.ToString()  ' Display HR employee count
            Label12.Text = FIData.Item2.ToString()  ' Display total working days
        End If
        Dim ITData As Tuple(Of Integer, Integer) = GetITEmployeeCount(selectedMonth, selectedYear)
        Dim sum As Integer = 0 ' Initialize sum variable

        ' Check for data retrieval success (optional)
        If ITData Is Nothing Then
            MessageBox.Show("Error retrieving data")
        Else
            Label10.Text = ITData.Item1.ToString()  ' Display HR employee count
            Label11.Text = ITData.Item2.ToString()  ' Display total working days

            ' Convert label texts to integers and sum them up
            sum = Convert.ToInt32(Label15.Text) + Convert.ToInt32(Label14.Text) + Convert.ToInt32(Label13.Text) + Convert.ToInt32(Label12.Text) + Convert.ToInt32(Label11.Text)
        End If

        Label16.Text = sum.ToString()

    End Sub

    Public Function GetHREmployeeCount(ByVal selectedMonth As String, ByVal selectedYear As Integer) As Tuple(Of Integer, Integer)
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"  ' Update with your actual connection string

        ' SQL query to count HR employees and total working days, filtered by month and year
        Dim sql As String = "Select COUNT(e.Employee_CODE) As EmployeeCount, " &
"SUM(T.TOTAL_SALARY) As TotalSalary " &
"FROM EMPLOYEE1 e " &
"INNER JOIN SALARY s On e.Employee_CODE = s.Employee_CODE " &
"INNER JOIN [TRANSACTION] t On e.Employee_CODE = t.Employee_CODE " &
"WHERE t.MONTH = @Month And t.YEAR = @Year And e.Department = 'HUMAN RESOURCE'"

        ' Execute the query and retrieve data
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(sql, connection)
                command.Parameters.AddWithValue("@Month", selectedMonth)
                command.Parameters.AddWithValue("@Year", selectedYear)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.Read() Then
                    Dim employeeCount As Integer = If(IsDBNull(reader("EmployeeCount")), 0, Convert.ToInt32(reader("EmployeeCount")))
                    Dim TotalSalary As Integer = If(IsDBNull(reader("TotalSalary")), 0, Convert.ToInt32(reader("TotalSalary")))
                    Return New Tuple(Of Integer, Integer)(employeeCount, TotalSalary)
                Else
                    ' Handle case where no data is found (optional)
                    Return New Tuple(Of Integer, Integer)(0, 0)  ' Return 0 for both counts
                End If

                reader.Close()
            End Using
        End Using

        ' Should not reach here if data is retrieved successfully
        Return Nothing  ' Indicate an error (optional)
    End Function
    Public Function GetPRODUCTIONEmployeeCount(ByVal selectedMonth As String, ByVal selectedYear As Integer) As Tuple(Of Integer, Integer)
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"  ' Update with your actual connection string

        ' SQL query to count HR employees and total working days, filtered by month and year
        Dim sql As String = "SELECT COUNT(e.Employee_CODE) AS EmployeeCount, " &
"SUM(T.TOTAL_SALARY) AS TotalSalary " &
"FROM EMPLOYEE1 e " &
"INNER JOIN SALARY s ON e.Employee_CODE = s.Employee_CODE " &
"INNER JOIN [TRANSACTION] t ON e.Employee_CODE = t.Employee_CODE " &
"WHERE t.MONTH = @Month AND t.YEAR = @Year AND e.Department = 'PRODUCTION'"


        ' Execute the query and retrieve data
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(sql, connection)
                command.Parameters.AddWithValue("@Month", selectedMonth)
                command.Parameters.AddWithValue("@Year", selectedYear)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.Read() Then
                    Dim employeeCount As Integer = If(IsDBNull(reader("EmployeeCount")), 0, Convert.ToInt32(reader("EmployeeCount")))
                    Dim TotalSalary As Integer = If(IsDBNull(reader("TotalSalary")), 0, Convert.ToInt32(reader("TotalSalary")))
                    Return New Tuple(Of Integer, Integer)(employeeCount, TotalSalary)
                Else
                    ' Handle case where no data is found (optional)
                    Return New Tuple(Of Integer, Integer)(0, 0)  ' Return 0 for both counts
                End If

                reader.Close()
            End Using
        End Using

        ' Should not reach here if data is retrieved successfully
        Return Nothing  ' Indicate an error (optional)
    End Function
    Public Function GetSALESEmployeeCount(ByVal selectedMonth As String, ByVal selectedYear As Integer) As Tuple(Of Integer, Integer)
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"  ' Update with your actual connection string

        ' SQL query to count HR employees and total working days, filtered by month and year
        Dim sql As String = "SELECT COUNT(e.Employee_CODE) AS EmployeeCount, " &
"SUM(T.TOTAL_SALARY) AS TotalSalary " &
"FROM EMPLOYEE1 e " &
"INNER JOIN SALARY s ON e.Employee_CODE = s.Employee_CODE " &
"INNER JOIN [TRANSACTION] t ON e.Employee_CODE = t.Employee_CODE " &
"WHERE t.MONTH = @Month AND t.YEAR = @Year AND e.Department = 'SALES'"


        ' Execute the query and retrieve data
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(sql, connection)
                command.Parameters.AddWithValue("@Month", selectedMonth)
                command.Parameters.AddWithValue("@Year", selectedYear)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.Read() Then
                    Dim employeeCount As Integer = If(IsDBNull(reader("EmployeeCount")), 0, Convert.ToInt32(reader("EmployeeCount")))
                    Dim TotalSalary As Integer = If(IsDBNull(reader("TotalSalary")), 0, Convert.ToInt32(reader("TotalSalary")))
                    Return New Tuple(Of Integer, Integer)(employeeCount, TotalSalary)
                Else
                    ' Handle case where no data is found (optional)
                    Return New Tuple(Of Integer, Integer)(0, 0)  ' Return 0 for both counts
                End If

                reader.Close()
            End Using
            Return Nothing
        End Using ' Indicate an error (optional)
    End Function

    Public Function GetFINANCEEmployeeCount(ByVal selectedMonth As String, ByVal selectedYear As Integer) As Tuple(Of Integer, Integer)
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"  ' Update with your actual connection string

        ' SQL query to count HR employees and total working days, filtered by month and year
        Dim sql As String = "SELECT COUNT(e.Employee_CODE) AS EmployeeCount, " &
"SUM(T.TOTAL_SALARY) AS TotalSalary " &
"FROM EMPLOYEE1 e " &
"INNER JOIN SALARY s ON e.Employee_CODE = s.Employee_CODE " &
"INNER JOIN [TRANSACTION] t ON e.Employee_CODE = t.Employee_CODE " &
"WHERE t.MONTH = @Month AND t.YEAR = @Year AND e.Department = 'FINANCE'"


        ' Execute the query and retrieve data
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(sql, connection)
                command.Parameters.AddWithValue("@Month", selectedMonth)
                command.Parameters.AddWithValue("@Year", selectedYear)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.Read() Then
                    Dim employeeCount As Integer = If(IsDBNull(reader("EmployeeCount")), 0, Convert.ToInt32(reader("EmployeeCount")))
                    Dim TotalSalary As Integer = If(IsDBNull(reader("TotalSalary")), 0, Convert.ToInt32(reader("TotalSalary")))
                    Return New Tuple(Of Integer, Integer)(employeeCount, TotalSalary)
                Else
                    ' Handle case where no data is found (optional)
                    Return New Tuple(Of Integer, Integer)(0, 0)  ' Return 0 for both counts
                End If

                reader.Close()
            End Using
        End Using

        ' Should not reach here if data is retrieved successfully

        ' Should not reach here if data is retrieved successfully
        Return Nothing  ' Indicate an error (optional)
    End Function
    Public Function GetITEmployeeCount(ByVal selectedMonth As String, ByVal selectedYear As Integer) As Tuple(Of Integer, Integer)
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"  ' Update with your actual connection string

        ' SQL query to count HR employees and total working days, filtered by month and year
        Dim sql As String = "SELECT COUNT(e.Employee_CODE) AS EmployeeCount, " &
"SUM(T.TOTAL_SALARY) AS TotalSalary " &
"FROM EMPLOYEE1 e " &
"INNER JOIN SALARY s ON e.Employee_CODE = s.Employee_CODE " &
"INNER JOIN [TRANSACTION] t ON e.Employee_CODE = t.Employee_CODE " &
"WHERE t.MONTH = @Month AND t.YEAR = @Year AND e.Department = 'INFORMATION AND TECHNOLOGY'"

        ' Execute the query and retrieve data
        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(sql, connection)
                command.Parameters.AddWithValue("@Month", selectedMonth)
                command.Parameters.AddWithValue("@Year", selectedYear)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.Read() Then
                    Dim employeeCount As Integer = If(IsDBNull(reader("EmployeeCount")), 0, Convert.ToInt32(reader("EmployeeCount")))
                    Dim TotalSalary As Integer = If(IsDBNull(reader("TotalSalary")), 0, Convert.ToInt32(reader("TotalSalary")))
                    Return New Tuple(Of Integer, Integer)(employeeCount, TotalSalary)
                Else
                    ' Handle case where no data is found (optional)
                    Return New Tuple(Of Integer, Integer)(0, 0)  ' Return 0 for both counts
                End If

                reader.Close()
            End Using
        End Using

        ' Should not reach here if data is retrieved successfully
        Return Nothing  ' Indicate an error (optional)
    End Function
    Private Sub LoadCompanyDetails()
        ' SQL query to get the company name and address
        Dim query As String = "SELECT [COMPANY NAME], [ADDRESS] FROM [dbo].[COMPANY]"

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
                        Dim companyAddress As String = reader("ADDRESS").ToString()

                        ' Set the labels' text to the company details
                        Label23.Text = companyName
                        Label24.Text = companyAddress
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

