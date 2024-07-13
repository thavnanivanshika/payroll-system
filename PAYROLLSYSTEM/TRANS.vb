Imports System.Data.SqlClient
Imports System.Globalization

Public Class TRANS
    Private Sub Form_Load()
        ' Populate Month ComboBox
        Label45.Text = "0"
        For i As Integer = 1 To 12
            ComboBox1.Items.Add(MonthName(i))
        Next
        ' Set default values to current month and year
        ComboBox1.SelectedIndex = DateTime.Now.Month - 1
        TextBox1.Text = DateTime.Now.Year.ToString()
    End Sub

    Private Sub TRANS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GroupBox1.Visible = False
        Label35.Visible = False
        Form_Load()
        LoadCompanyDetails()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim month As Integer = ComboBox1.SelectedIndex + 1
        Dim year As Integer = Integer.Parse(TextBox1.Text)
        DisplayPFAndESIRates()
        Dim employeeCode As String = TextBox2.Text
        GroupBox1.Visible = True
        If String.IsNullOrEmpty(employeeCode) Then
            MessageBox.Show("Please enter the employee code.")
            Return
        End If
        ' Fetch and display employee and salary details
        FetchEmployeeDetails(employeeCode)
        FetchSalaryDetails(employeeCode)
        DisplayMonthDetails(month, year)
        UpdatePF()
        UpdateESI()

    End Sub

    Private Sub FetchEmployeeDetails(employeeCode As String)
        Dim connString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim query As String = "SELECT EMPLOYEE_NAME, Department FROM Employee1 WHERE Employee_Code = @EmployeeCode"

        Using conn As New SqlConnection(connString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@EmployeeCode", employeeCode)
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Label5.Text = reader("EMPLOYEE_NAME").ToString()
                        Label7.Text = reader("Department").ToString()
                    Else
                        MessageBox.Show("Employee not found.")
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub FetchSalaryDetails(employeeCode As String)
        Dim connString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim query As String = "SELECT Basic_Salary, DA, HRA, Medical, Performance, Conveyance FROM Salary WHERE Employee_Code = @EmployeeCode"

        Using conn As New SqlConnection(connString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@EmployeeCode", employeeCode)

                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Label8.Text = reader("Basic_Salary").ToString()
                        Label9.Text = reader("DA").ToString()
                        Label10.Text = reader("HRA").ToString()
                        Label11.Text = reader("Medical").ToString()
                        Label12.Text = reader("Performance").ToString()
                        Label13.Text = reader("Conveyance").ToString()
                    Else
                        MessageBox.Show("Salary details not found.")
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        HOME.Show()
    End Sub
    Private Sub LoadCompanyDetails()
        Dim connString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        ' SQL query to get the company name and address
        Dim query As String = "SELECT [COMPANY NAME], [PHONE] FROM [dbo].[COMPANY]"

        ' Create a connection and command object
        Using connection As New SqlConnection(connString),
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
    Private Sub DisplayMonthDetails(month As String, year As Integer)
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month)
        Dim sundays As Integer = CountSundays(month, year)
        Dim daysWithoutSundays As Integer = daysInMonth - sundays

        Label2.Text = daysInMonth.ToString()
        Label40.Text = sundays.ToString()
        Label41.Text = daysWithoutSundays.ToString()
    End Sub

    Private Function CountSundays(month As String, year As Integer) As Integer
        Dim sundays As Integer = 0
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month)

        For day As Integer = 1 To daysInMonth
            Dim currentDate As New DateTime(year, month, day)
            If currentDate.DayOfWeek = DayOfWeek.Sunday Then
                sundays += 1
            End If
        Next

        Return sundays
    End Function
    Private Sub UpdateUnpaidLeave()
        Try
            ' Convert TextBox inputs to numbers
            Dim dutyDay As Integer = If(String.IsNullOrEmpty(TextBox3.Text), 0, Integer.Parse(TextBox3.Text))
            Dim earlyLeave As Integer = If(String.IsNullOrEmpty(TextBox4.Text), 0, Integer.Parse(TextBox4.Text))
            Dim casualLeave As Integer = If(String.IsNullOrEmpty(TextBox5.Text), 0, Integer.Parse(TextBox5.Text))
            Dim holiday As Integer = If(String.IsNullOrEmpty(TextBox6.Text), 0, Integer.Parse(TextBox6.Text))
            Dim deductedValue As Integer = If(String.IsNullOrEmpty(Label41.Text), 0, Integer.Parse(Label41.Text))

            ' Calculate total leave taken
            Dim totalLeaveTaken As Integer = dutyDay + earlyLeave + casualLeave + holiday

            ' Calculate unpaid leave
            Dim unpaidLeave As Integer = deductedValue - totalLeaveTaken

            ' Display unpaid leave
            Label45.Text = unpaidLeave.ToString()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged, TextBox4.TextChanged, TextBox5.TextChanged, TextBox6.TextChanged
        UpdateUnpaidLeave()
    End Sub

    Private Sub DisplayPFAndESIRates()
        Dim connString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"



        ' Create a connection and command object
        Dim connection As New SqlConnection(connString)


        Try
            ' Open connection
            connection.Open()

            ' Query to fetch PF rate from parameter table
            Dim queryPF As String = "SELECT [PF EMPLOYEE RATE] FROM Parameter"
            Dim commandPF As SqlCommand = New SqlCommand(queryPF, connection)
            Dim pfRate As Decimal = Convert.ToDecimal(commandPF.ExecuteScalar()) / 100 ' Divide by 100
            Label33.Text = pfRate.ToString()

            ' Query to fetch ESI rate from parameter table
            Dim queryESI As String = "SELECT [ESI EMPLOYEE RATE] FROM Parameter"
            Dim commandESI As SqlCommand = New SqlCommand(queryESI, connection)
            Dim esiRate As Decimal = Convert.ToDecimal(commandESI.ExecuteScalar()) / 100 ' Divide by 100
            Label34.Text = esiRate.ToString()

        Catch ex As Exception
            MessageBox.Show("Error fetching rates from database: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Close connection
            connection.Close()
        End Try
    End Sub
    Private Sub UpdatePF()
        Dim connString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"



        ' Create a connection and command object
        Dim connection As New SqlConnection(connString)
        Try
            connection.Open()
            ' Convert Label8 and Label9 values to decimals
            Dim label8Value As Decimal = Decimal.Parse(Label8.Text)
            Dim label9Value As Decimal = Decimal.Parse(Label9.Text)

            ' Check if Label8 value is greater than 21000
            If label8Value > 21000 Then
                ' Query to fetch PF rate from parameter table
                Dim queryPF As String = "SELECT [PF EMPLOYEE RATE]  FROM Parameter"
                Dim commandPF As SqlCommand = New SqlCommand(queryPF, connection)
                Dim pfRate As Decimal = Convert.ToDecimal(commandPF.ExecuteScalar()) / 100 ' Divide by 100

                ' Calculate PF amount
                Dim pfAmount As Decimal = (label8Value + label9Value) * pfRate

                ' Display PF amount in Label31
                Label31.Text = pfAmount.ToString()
            Else
                ' If Label8 value is not greater than 21000, set Label31 value to zero
                Label31.Text = "0"
            End If
        Catch ex As Exception
            MessageBox.Show("Error calculating PF: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub UpdateESI()
        Dim connString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim connection As New SqlConnection(connString)
        Try
            connection.Open()
            Dim label8Value As Decimal = Decimal.Parse(Label8.Text)
            Dim label9Value As Decimal = Decimal.Parse(Label9.Text)

            ' Query to fetch ESI rate from parameter table
            Dim queryESI As String = "SELECT [ESI EMPLOYEE RATE] FROM Parameter"
            Dim commandESI As SqlCommand = New SqlCommand(queryESI, connection)
            Dim esiRate As Decimal = Convert.ToDecimal(commandESI.ExecuteScalar()) / 100 ' Divide by 100

            ' Calculate ESI amount
            Dim esiAmount As Decimal = (label8Value + label9Value) * esiRate

            ' Display ESI amount in Label32
            Label32.Text = esiAmount.ToString()
        Catch ex As Exception
            MessageBox.Show("Error calculating ESI: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub UpdateTotalDaysWorked()
        Try
            ' Convert TextBox inputs to integers
            Dim textBox1Value As Integer = If(String.IsNullOrEmpty(TextBox3.Text), 0, Integer.Parse(TextBox3.Text))
            Dim textBox2Value As Integer = If(String.IsNullOrEmpty(TextBox4.Text), 0, Integer.Parse(TextBox4.Text))
            Dim textBox3Value As Integer = If(String.IsNullOrEmpty(TextBox5.Text), 0, Integer.Parse(TextBox5.Text))
            Dim textBox4Value As Integer = If(String.IsNullOrEmpty(TextBox6.Text), 0, Integer.Parse(TextBox6.Text))

            ' Convert Label45 value to integer
            Dim label45Value As Integer = If(String.IsNullOrEmpty(Label45.Text), 0, Integer.Parse(Label45.Text))

            ' Calculate total days worked
            Dim totalDaysWorked As Integer = (textBox1Value + textBox2Value + textBox3Value + textBox4Value)

            ' Display total days worked in LabelXX (replace XX with the appropriate label number)


            ' Calculate salary
            Dim salary As Decimal = 0
            Dim labelValuesSum As Decimal = Decimal.Parse(Label8.Text) + Decimal.Parse(Label9.Text) + Decimal.Parse(Label10.Text) +
                                        Decimal.Parse(Label11.Text) + Decimal.Parse(Label12.Text) + Decimal.Parse(Label13.Text)
            Dim label2Value As Decimal = If(String.IsNullOrEmpty(Label41.Text), 0, Decimal.Parse(Label41.Text))

            If label2Value <> 0 Then
                salary = (labelValuesSum / label2Value) * totalDaysWorked

            End If

            ' Display salary in LabelYY (replace YY with the appropriate label number)

            Dim textBox7Value As Integer = If(String.IsNullOrEmpty(TextBox8.Text), 0, Integer.Parse(TextBox8.Text))
            Dim textBox8Value As Integer = If(String.IsNullOrEmpty(TextBox9.Text), 0, Integer.Parse(TextBox9.Text))
            Dim textBox5Value As Integer = If(String.IsNullOrEmpty(TextBox10.Text), 0, Integer.Parse(TextBox10.Text))
            Dim textBox6Value As Integer = If(String.IsNullOrEmpty(TextBox11.Text), 0, Integer.Parse(TextBox11.Text))

            ' Convert Label45 value to integer
            Dim label8Value As Integer = If(String.IsNullOrEmpty(Label31.Text), 0, Decimal.Parse(Label31.Text))
            Dim label9Value As Integer = If(String.IsNullOrEmpty(Label32.Text), 0, Decimal.Parse(Label32.Text))
            Dim totalDEDUCTION As Integer = (textBox7Value + textBox8Value + textBox5Value + textBox6Value + label8Value + label9Value)

            Dim TOTALSALARY As Integer = salary - totalDEDUCTION

            Label35.Text = TOTALSALARY.ToString()
        Catch ex As Exception
            MessageBox.Show("Error calculating total days worked and salary: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label35.Visible = True
        UpdateTotalDaysWorked()
        AddOrUpdateTransaction()
    End Sub
    Private Sub AddOrUpdateTransaction()
        Dim connString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
        Dim connection As New SqlConnection(connString)
        Try

            ' Convert inputs to appropriate data types
            Dim employeeCode As Integer
            If Integer.TryParse(TextBox2.Text, employeeCode) = False Then employeeCode = 0

            Dim month As String = If(ComboBox1.SelectedItem IsNot Nothing, ComboBox1.SelectedItem.ToString(), String.Empty)

            Dim year As Integer
            If Integer.TryParse(TextBox1.Text, year) = False Then year = 0

            Dim workingDay As Integer
            If Integer.TryParse(TextBox3.Text, workingDay) = False Then workingDay = 0

            Dim holiday As Integer
            If Integer.TryParse(TextBox4.Text, holiday) = False Then holiday = 0

            Dim earlyLeave As Integer
            If Integer.TryParse(TextBox5.Text, earlyLeave) = False Then earlyLeave = 0

            Dim casualLeave As Integer
            If Integer.TryParse(TextBox6.Text, casualLeave) = False Then casualLeave = 0

            Dim leaveWithoutPay As Integer
            If Integer.TryParse(Label45.Text, leaveWithoutPay) = False Then leaveWithoutPay = 0

            Dim tds As Decimal
            If Decimal.TryParse(TextBox8.Text, tds) = False Then tds = 0

            Dim advance As Decimal
            If Decimal.TryParse(TextBox9.Text, advance) = False Then advance = 0

            Dim loanFromBank As Decimal
            If Decimal.TryParse(TextBox10.Text, loanFromBank) = False Then loanFromBank = 0

            Dim loanFromCompany As Decimal
            If Decimal.TryParse(TextBox11.Text, loanFromCompany) = False Then loanFromCompany = 0

            Dim totalSalary As Decimal
            If Decimal.TryParse(Label35.Text, totalSalary) = False Then totalSalary = 0

            ' Open the connection
            connection.Open()

            ' Check if record exists
            Dim queryCheck As String = "SELECT COUNT(*) FROM [TRANSACTION] WHERE Employee_Code = @Employee_Code AND [Month] = @Month AND [Year] = @Year"
            Dim commandCheck As SqlCommand = New SqlCommand(queryCheck, connection)
            commandCheck.Parameters.AddWithValue("@Employee_Code", employeeCode)
            commandCheck.Parameters.AddWithValue("@Month", month)
            commandCheck.Parameters.AddWithValue("@Year", year)

            Dim recordCount As Integer = Convert.ToInt32(commandCheck.ExecuteScalar())

            If recordCount > 0 Then
                ' Record exists, ask for update
                Dim dialogResult As DialogResult = MessageBox.Show("Record for this Employee Code, Month, and Year already exists. Do you want to update it?", "Update Record", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If dialogResult = DialogResult.Yes Then
                    ' Update existing record
                    Dim queryUpdate As String = "UPDATE [Transaction] SET Working_Day = @Working_Day, Holiday = @Holiday, Early_Leave = @Early_Leave, Casual_Leave = @Casual_Leave, Leave_Without_Pay = @Leave_Without_Pay, TDS = @TDS, Advance = @Advance, Loan_From_Bank = @Loan_From_Bank, Loan_From_Company = @Loan_From_Company, Total_Salary = @Total_Salary WHERE Employee_Code = @Employee_Code AND [Month] = @Month AND [Year] = @Year"
                    Dim commandUpdate As SqlCommand = New SqlCommand(queryUpdate, connection)
                    commandUpdate.Parameters.AddWithValue("@Working_Day", workingDay)
                    commandUpdate.Parameters.AddWithValue("@Holiday", holiday)
                    commandUpdate.Parameters.AddWithValue("@Early_Leave", earlyLeave)
                    commandUpdate.Parameters.AddWithValue("@Casual_Leave", casualLeave)
                    commandUpdate.Parameters.AddWithValue("@Leave_Without_Pay", leaveWithoutPay)
                    commandUpdate.Parameters.AddWithValue("@TDS", tds)
                    commandUpdate.Parameters.AddWithValue("@Advance", advance)
                    commandUpdate.Parameters.AddWithValue("@Loan_From_Bank", loanFromBank)
                    commandUpdate.Parameters.AddWithValue("@Loan_From_Company", loanFromCompany)
                    commandUpdate.Parameters.AddWithValue("@Total_Salary", totalSalary)
                    commandUpdate.Parameters.AddWithValue("@Employee_Code", employeeCode)
                    commandUpdate.Parameters.AddWithValue("@Month", month)
                    commandUpdate.Parameters.AddWithValue("@Year", year)

                    commandUpdate.ExecuteNonQuery()

                    MessageBox.Show("Record updated successfully.", "Update Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                ' Insert new record
                Dim queryInsert As String = "INSERT INTO [Transaction] (Employee_Code, [Month], [Year], Working_Day, Holiday, Early_Leave, Casual_Leave, Leave_Without_Pay, TDS, Advance, Loan_From_Bank, Loan_From_Company, Total_Salary) VALUES (@Employee_Code, @Month, @Year, @Working_Day, @Holiday, @Early_Leave, @Casual_Leave, @Leave_Without_Pay, @TDS, @Advance, @Loan_From_Bank, @Loan_From_Company, @Total_Salary)"
                Dim commandInsert As SqlCommand = New SqlCommand(queryInsert, connection)
                commandInsert.Parameters.AddWithValue("@Employee_Code", employeeCode)
                commandInsert.Parameters.AddWithValue("@Month", month)
                commandInsert.Parameters.AddWithValue("@Year", year)
                commandInsert.Parameters.AddWithValue("@Working_Day", workingDay)
                commandInsert.Parameters.AddWithValue("@Holiday", holiday)
                commandInsert.Parameters.AddWithValue("@Early_Leave", earlyLeave)
                commandInsert.Parameters.AddWithValue("@Casual_Leave", casualLeave)
                commandInsert.Parameters.AddWithValue("@Leave_Without_Pay", leaveWithoutPay)
                commandInsert.Parameters.AddWithValue("@TDS", tds)
                commandInsert.Parameters.AddWithValue("@Advance", advance)
                commandInsert.Parameters.AddWithValue("@Loan_From_Bank", loanFromBank)
                commandInsert.Parameters.AddWithValue("@Loan_From_Company", loanFromCompany)
                commandInsert.Parameters.AddWithValue("@Total_Salary", totalSalary)

                commandInsert.ExecuteNonQuery()

                MessageBox.Show("Record added successfully.", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Finally
            ' Close the connection
            connection.Close()
        End Try
    End Sub

End Class
