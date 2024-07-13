Imports System.Data.SqlClient
Imports System.Globalization

Public Class TRANSACTION
    Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
    Private Sub TRANSACTION_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GroupBox1.Visible = False
        Label33.Visible = False
        Label34.Visible = False
        Label35.Visible = False
        Label39.Visible = False
        Label41.Visible = False ' Make Label41 initially invisible
        LoadCompanyDetails()
        ' Populate the ComboBox with months
        ComboBox1.Items.AddRange(New String() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})

        ' Set the default selected item to the current month
        Dim currentMonth As Integer = DateTime.Now.Month
        ComboBox1.SelectedIndex = currentMonth - 1

        ' Set the default value of the TextBox to the current year
        TextBox1.Text = DateTime.Now.Year.ToString()

        ' Set up event handlers
        AddHandler ComboBox1.SelectedIndexChanged, AddressOf UpdateDaysInMonth
        AddHandler TextBox1.TextChanged, AddressOf UpdateDaysInMonth

        ' Update days in the current month and year
        UpdateDaysInMonth(Nothing, Nothing)
    End Sub

    Private Sub UpdateDaysInMonth(sender As Object, e As EventArgs)
        Dim month As Integer = ComboBox1.SelectedIndex + 1
        Dim year As Integer

        ' Validate the year input
        If Integer.TryParse(TextBox1.Text, year) Then
            ' Calculate working days, Sundays, and total days
            Dim workingDays As Integer = CalculateWorkingDays(year, month)
            Dim sundays As Integer = CalculateSundays(year, month)
            Dim totalDays As Integer = DateTime.DaysInMonth(year, month)

            ' Update labels
            Label2.Text = $"Working days in {ComboBox1.SelectedItem} {year}: {workingDays}"
            Label40.Visible = True ' Make Sundays label visible
            Label40.Text = $"Sundays: {sundays}"
            Label41.Visible = True ' Make Total Days label visible
            Label41.Text = $"Total Days: {totalDays}"
        Else
            Label2.Text = "Invalid year"
        End If
    End Sub

    Private Function CalculateWorkingDays(year As Integer, month As Integer) As Integer
        Dim workingDays As Integer = 0
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month)

        For day As Integer = 1 To daysInMonth
            Dim currentDay As New DateTime(year, month, day)
            ' Exclude Sundays
            If currentDay.DayOfWeek <> DayOfWeek.Sunday Then
                workingDays += 1
            End If
        Next

        Return workingDays
    End Function

    Private Function CalculateSundays(year As Integer, month As Integer) As Integer
        Dim sundays As Integer = 0
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month)

        For day As Integer = 1 To daysInMonth
            Dim currentDay As New DateTime(year, month, day)
            ' Count Sundays
            If currentDay.DayOfWeek = DayOfWeek.Sunday Then
                sundays += 1
            End If
        Next

        Return sundays
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim employeeCode As String = TextBox2.Text
        If Not String.IsNullOrEmpty(employeeCode) Then
            FetchEmployeeData(employeeCode)
        Else
            MessageBox.Show("Please enter a valid employee code.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim query As String = "SELECT [PF EMPLOYEE RATE], [ESI EMPLOYEE RATE] FROM [dbo].[PARAMETER]"

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand(query, connection)
                    Using reader As SqlDataReader = command.ExecuteReader()
                        If reader.Read() Then
                            ' Fetch the values from the database
                            Dim pfRate As Decimal = reader.GetDecimal(0) / 100
                            Dim esiRate As Decimal = reader.GetDecimal(1) / 100

                            ' Display the values in the labels
                            Label31.Text = pfRate.ToString()
                            Label32.Text = esiRate.ToString()

                            ' Make the labels visible
                            Label31.Visible = True
                            Label32.Visible = True
                        Else
                            MessageBox.Show("No data found in the PARAMETER table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while fetching data from the database: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FetchEmployeeData(employeeCode As String)
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim query As String = "SELECT EMPLOYEE_NAME, DEPARTMENT FROM EMPLOYEE1 WHERE EMPLOYEE_CODE = @EmployeeCode"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                connection.Open()
                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.HasRows Then
                        reader.Read()
                        Label5.Text = reader("EMPLOYEE_NAME").ToString()
                        Label7.Text = reader("DEPARTMENT").ToString()

                        GroupBox1.Visible = True
                    Else
                        MessageBox.Show("Employee not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        GroupBox1.Visible = False
                    End If
                End Using
            End Using
        End Using

        Dim salaryQuery As String = "SELECT BASIC_SALARY, DA, HRA, MEDICAL, PERFORMANCE, CONVEYANCE FROM SALARY WHERE EMPLOYEE_CODE = @EmployeeCode"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(salaryQuery, connection)
                command.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                connection.Open()
                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.HasRows Then
                        reader.Read()
                        ' Update labels with salary information
                        Label8.Text = reader("BASIC_SALARY").ToString()
                        Label9.Text = reader("DA").ToString()
                        Label10.Text = reader("HRA").ToString()
                        Label11.Text = reader("MEDICAL").ToString()
                        Label12.Text = reader("PERFORMANCE").ToString()
                        Label13.Text = reader("CONVEYANCE").ToString()
                        ' Make salary labels visible
                        GroupBox1.Visible = True
                    Else
                        MessageBox.Show("Salary information not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ' Hide salary labels
                        GroupBox1.Visible = False
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim employeeCode As String = TextBox2.Text
        Dim month As String = ComboBox1.SelectedItem.ToString()
        Dim year As Integer

        If Integer.TryParse(TextBox1.Text, year) Then
            ' Calculate working days
            Dim workingDays As Integer = Integer.Parse(TextBox3.Text)

            ' Fetch other values from text boxes
            Dim holiday As Integer = Integer.Parse(TextBox4.Text)
            Dim earlyLeave As Integer = Integer.Parse(TextBox5.Text)
            Dim casualLeave As Integer = Integer.Parse(TextBox6.Text)
            Dim leaveWithoutPay As Integer = Integer.Parse(TextBox7.Text)
            Dim tds As Integer = Integer.Parse(TextBox8.Text)
            Dim advance As Integer = Integer.Parse(TextBox9.Text)
            Dim loanFromBank As Decimal = Decimal.Parse(TextBox10.Text)
            Dim loanFromCompany As Decimal = Decimal.Parse(TextBox11.Text)
            Dim totalDaysWorked As Integer = workingDays + holiday + earlyLeave + casualLeave - leaveWithoutPay
            Dim pfRate As Decimal = Decimal.Parse(Label31.Text)
            Dim esiRate As Decimal = Decimal.Parse(Label32.Text)
            Dim totalAmountReduced As Decimal = tds + advance + loanFromBank + loanFromCompany + pfRate + esiRate

            ' Fetch salary values from the database
            Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
            Dim salaryQuery As String = "SELECT BASIC_SALARY, DA, HRA, MEDICAL, PERFORMANCE, CONVEYANCE FROM SALARY WHERE EMPLOYEE_CODE = @EmployeeCode"

            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Using command As New SqlCommand(salaryQuery, connection)
                    command.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                    Using reader As SqlDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            reader.Read()
                            ' Get salary values from the database
                            Dim basicSalary As Decimal = reader.GetDecimal(0)
                            Dim da As Decimal = reader.GetDecimal(1)
                            Dim hra As Decimal = reader.GetDecimal(2)
                            Dim medical As Decimal = reader.GetDecimal(3)
                            Dim performance As Decimal = reader.GetDecimal(4)
                            Dim conveyance As Decimal = reader.GetDecimal(5)

                            ' Calculate total salary
                            Dim totalSalary As Decimal = basicSalary + da + hra + medical + performance + conveyance
                            Dim dailySalary As Decimal = totalSalary / CalculateWorkingDays(year, ComboBox1.SelectedIndex + 1)
                            Dim totalEarnings As Decimal = (dailySalary * totalDaysWorked) - totalAmountReduced

                            ' Display total salary in label in rupee format
                            Dim cultureInfo As New CultureInfo("en-IN")
                            Label33.Text = dailySalary.ToString("C", cultureInfo) ' Display as currency format with Indian English culture
                            Label34.Text = totalDaysWorked.ToString()
                            Label35.Text = totalEarnings.ToString("C", cultureInfo)
                            Label35.Visible = True
                            Label39.Visible = True
                            reader.Close() ' Close the reader before executing the next command

                            ' Check if the record already exists
                            Dim checkQuery As String = "SELECT COUNT(*) FROM [dbo].[TRANSACTION] WHERE EMPLOYEE_CODE = @EmployeeCode AND [MONTH] = @Month AND [YEAR] = @Year"
                            Dim recordExists As Boolean

                            Using checkCommand As New SqlCommand(checkQuery, connection)
                                checkCommand.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                                checkCommand.Parameters.AddWithValue("@Month", month)
                                checkCommand.Parameters.AddWithValue("@Year", year)
                                recordExists = Convert.ToInt32(checkCommand.ExecuteScalar()) > 0
                            End Using

                            If recordExists Then
                                ' Ask user for confirmation to update
                                Dim result As DialogResult = MessageBox.Show("Record already exists. Do you want to update it?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                If result = DialogResult.Yes Then
                                    ' Update the record
                                    Dim updateQuery As String = "UPDATE [dbo].[TRANSACTION] SET WORKING_DAY = @WorkingDays, HOLIDAY = @Holiday, EARLY_LEAVE = @EarlyLeave, CASUAL_LEAVE = @CasualLeave, LEAVE_WITHOUT_PAY = @LeaveWithoutPay, TDS = @TDS, ADVANCE = @Advance, LOAN_FROM_BANK = @LoanFromBank, LOAN_FROM_COMPANY = @LoanFromCompany, TOTAL_SALARY = @TotalSalary WHERE EMPLOYEE_CODE = @EmployeeCode AND [MONTH] = @Month AND [YEAR] = @Year"

                                    Using updateCommand As New SqlCommand(updateQuery, connection)
                                        updateCommand.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                                        updateCommand.Parameters.AddWithValue("@Month", month)
                                        updateCommand.Parameters.AddWithValue("@Year", year)
                                        updateCommand.Parameters.AddWithValue("@WorkingDays", workingDays)
                                        updateCommand.Parameters.AddWithValue("@Holiday", holiday)
                                        updateCommand.Parameters.AddWithValue("@EarlyLeave", earlyLeave)
                                        updateCommand.Parameters.AddWithValue("@CasualLeave", casualLeave)
                                        updateCommand.Parameters.AddWithValue("@LeaveWithoutPay", leaveWithoutPay)
                                        updateCommand.Parameters.AddWithValue("@TDS", tds)
                                        updateCommand.Parameters.AddWithValue("@Advance", advance)
                                        updateCommand.Parameters.AddWithValue("@LoanFromBank", loanFromBank)
                                        updateCommand.Parameters.AddWithValue("@LoanFromCompany", loanFromCompany)
                                        updateCommand.Parameters.AddWithValue("@TotalSalary", totalEarnings)

                                        updateCommand.ExecuteNonQuery()
                                    End Using

                                    MessageBox.Show("Record updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                End If
                            Else
                                ' Insert data into TRANS table
                                Dim insertQuery As String = "INSERT INTO [dbo].[TRANSACTION] " &
                                                            "(EMPLOYEE_CODE, [MONTH], [YEAR], WORKING_DAY, HOLIDAY, EARLY_LEAVE, CASUAL_LEAVE, LEAVE_WITHOUT_PAY, TDS, ADVANCE, LOAN_FROM_BANK, LOAN_FROM_COMPANY, TOTAL_SALARY) " &
                                                            "VALUES (@EmployeeCode, @Month, @Year, @WorkingDays, @Holiday, @EarlyLeave, @CasualLeave, @LeaveWithoutPay, @TDS, @Advance, @LoanFromBank, @LoanFromCompany, @TotalSalary)"

                                Using insertCommand As New SqlCommand(insertQuery, connection)
                                    insertCommand.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                                    insertCommand.Parameters.AddWithValue("@Month", month)
                                    insertCommand.Parameters.AddWithValue("@Year", year)
                                    insertCommand.Parameters.AddWithValue("@WorkingDays", workingDays)
                                    insertCommand.Parameters.AddWithValue("@Holiday", holiday)
                                    insertCommand.Parameters.AddWithValue("@EarlyLeave", earlyLeave)
                                    insertCommand.Parameters.AddWithValue("@CasualLeave", casualLeave)
                                    insertCommand.Parameters.AddWithValue("@LeaveWithoutPay", leaveWithoutPay)
                                    insertCommand.Parameters.AddWithValue("@TDS", tds)
                                    insertCommand.Parameters.AddWithValue("@Advance", advance)
                                    insertCommand.Parameters.AddWithValue("@LoanFromBank", loanFromBank)
                                    insertCommand.Parameters.AddWithValue("@LoanFromCompany", loanFromCompany)
                                    insertCommand.Parameters.AddWithValue("@TotalSalary", totalEarnings)

                                    insertCommand.ExecuteNonQuery()
                                End Using

                                MessageBox.Show("Data inserted into TRANS table.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        Else
                            MessageBox.Show("Salary information not found for the employee.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End Using
                End Using
            End Using
        Else
            MessageBox.Show("Invalid year", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                        Label44.Text = companyName
                        Label43.Text = companyPHONE
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
