Imports System.Data.SqlClient

Public Class REPORT
    ' Replace "connectionstring" with your actual connection string
    Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"

    Private Sub REPORT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCompanyDetails()
        ComboBox1.Items.AddRange(New String() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})

        ' Set the default selected item to the current month
        Dim currentMonth As Integer = DateTime.Now.Month
        ComboBox1.SelectedIndex = currentMonth - 1

        ' Set the default value of the TextBox to the current year
        TextBox1.Text = DateTime.Now.Year.ToString()

        ' Initially hide DataGridViews and Labels
        DataGridView1.Visible = False
        DataGridView2.Visible = False
        DataGridView3.Visible = False
        DataGridView4.Visible = False
        DataGridView5.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False

        ' Initialize empty DataTables for each department
        DataGridView1.DataSource = dtHR
        DataGridView2.DataSource = dtProduction
        DataGridView3.DataSource = dtIT
        DataGridView4.DataSource = dtFinance
        DataGridView5.DataSource = dtSales
    End Sub

    Dim dtHR As New DataTable("HR")
    Dim dtProduction As New DataTable("Production")
    Dim dtIT As New DataTable("IT")
    Dim dtFinance As New DataTable("Finance")
    Dim dtSales As New DataTable("Sales")

    Sub LoadData()
        Dim selectedMonth As String = ComboBox1.SelectedItem.ToString()
        Dim selectedYear As Integer = Integer.Parse(TextBox1.Text)

        ' Make DataGridViews and Labels visible
        DataGridView1.Visible = True
        DataGridView2.Visible = True
        DataGridView3.Visible = True
        DataGridView4.Visible = True
        DataGridView5.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True

        Dim connection As SqlConnection = New SqlConnection(connectionString)

        ' Build the base SQL statement
        Dim sql = "SELECT e.EMPLOYEE_CODE, e.EMPLOYEE_NAME, e.Department, " &
          " s.[DESIGNATION], s.[BASIC_SALARY], s.[DA], " &
          "s.[HRA], s.[MEDICAL], s.[PERFORMANCE], s.[CONVEYANCE], " &
          "t.MONTH, t.YEAR, t.WORKING_DAY, t.HOLIDAY, t.EARLY_LEAVE, t.CASUAL_LEAVE, " &
          "t.LEAVE_WITHOUT_PAY, t.TDS, t.ADVANCE, t.LOAN_FROM_BANK, t.LOAN_FROM_COMPANY, " &
          "t.TOTAL_SALARY " &
          "FROM EMPLOYEE1 e " &
          "INNER JOIN SALARY s ON e.Employee_CODE = s.Employee_CODE " &
          "INNER JOIN [TRANSACTION] t ON e.Employee_CODE = t.Employee_CODE " &
          "WHERE t.MONTH = @Month AND t.YEAR = @Year " &
          "AND e.Department = @Department " &
          "ORDER BY e.EMPLOYEE_NAME"


        ' Create separate DataAdapters for each department
        Dim adapterHR As New SqlDataAdapter(sql, connection)
        Dim adapterProduction As New SqlDataAdapter(sql, connection)
        Dim adapterIT As New SqlDataAdapter(sql, connection)
        Dim adapterFinance As New SqlDataAdapter(sql, connection)
        Dim adapterSales As New SqlDataAdapter(sql, connection)

        ' Add parameters for month and year to each adapter
        adapterHR.SelectCommand.Parameters.AddWithValue("@Month", selectedMonth)
        adapterHR.SelectCommand.Parameters.AddWithValue("@Year", selectedYear)
        adapterHR.SelectCommand.Parameters.AddWithValue("@Department", "HUMAN RESOURCE")

        adapterProduction.SelectCommand.Parameters.AddWithValue("@Month", selectedMonth)
        adapterProduction.SelectCommand.Parameters.AddWithValue("@Year", selectedYear)
        adapterProduction.SelectCommand.Parameters.AddWithValue("@Department", "PRODUCTION")

        adapterIT.SelectCommand.Parameters.AddWithValue("@Month", selectedMonth)
        adapterIT.SelectCommand.Parameters.AddWithValue("@Year", selectedYear)
        adapterIT.SelectCommand.Parameters.AddWithValue("@Department", "INFORMATION AND TECHNOLOGY")

        adapterFinance.SelectCommand.Parameters.AddWithValue("@Month", selectedMonth)
        adapterFinance.SelectCommand.Parameters.AddWithValue("@Year", selectedYear)
        adapterFinance.SelectCommand.Parameters.AddWithValue("@Department", "FINANCE")

        adapterSales.SelectCommand.Parameters.AddWithValue("@Month", selectedMonth)
        adapterSales.SelectCommand.Parameters.AddWithValue("@Year", selectedYear)
        adapterSales.SelectCommand.Parameters.AddWithValue("@Department", "SALES")

        ' Clear previous data from DataTables
        dtHR.Clear()
        dtProduction.Clear()
        dtIT.Clear()
        dtFinance.Clear()
        dtSales.Clear()

        ' Fill each DataTable with data from its corresponding adapter
        adapterHR.Fill(dtHR)
        adapterProduction.Fill(dtProduction)
        adapterIT.Fill(dtIT)
        adapterFinance.Fill(dtFinance)
        adapterSales.Fill(dtSales)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        LoadData()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        HOME.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Get selected month
        Dim selectedMonth As String = If(ComboBox1.SelectedItem IsNot Nothing, ComboBox1.SelectedItem.ToString(), "")
        If String.IsNullOrEmpty(selectedMonth) Then
            MessageBox.Show("Please select a month.")
            Return ' Exit the function if no month is selected
        End If

        ' Convert year (assuming numeric)
        Dim selectedYear As Integer
        If Not Integer.TryParse(TextBox1.Text, selectedYear) Then
            MessageBox.Show("Invalid year format. Please enter a valid number.")
            Return ' Exit the function if conversion fails
        End If

        ' Pass month and year to FINAL form
        Dim finalForm As New FINAL(selectedMonth, selectedYear)
        finalForm.Show()
        Me.Close()
    End Sub
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
                        Label8.Text = companyName
                        Label9.Text = companyAddress
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
