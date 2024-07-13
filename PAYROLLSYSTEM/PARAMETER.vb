Imports System.Data.SqlClient

Public Class PARAMETER
    Private connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()
                Dim sql As String = "UPDATE [dbo].[PARAMETER] SET 
                                    [PF LIMIT] = @PF_Limit,
                                    [PF EMPLOYEE RATE] = @PF_Employee_Rate,
                                    [PF EMPLOYER RATE] = @PF_Employer_Rate,
                                    [ESI LIMIT] = @ESI_Limit,
                                    [ESI EMPLOYEE RATE] = @ESI_Employee_Rate,
                                    [ESI EMPLOYER RATE] = @ESI_Employer_Rate"

                Using cmd As New SqlCommand(sql, con)
                    ' Parsing and adding parameters
                    cmd.Parameters.AddWithValue("@PF_Limit", If(Decimal.TryParse(TextBox1.Text, New Decimal), Decimal.Parse(TextBox1.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@PF_Employee_Rate", If(Decimal.TryParse(TextBox2.Text, New Decimal), Decimal.Parse(TextBox2.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@PF_Employer_Rate", If(Decimal.TryParse(TextBox3.Text, New Decimal), Decimal.Parse(TextBox3.Text), DBNull.Value))

                    cmd.Parameters.AddWithValue("@ESI_Limit", If(Decimal.TryParse(TextBox4.Text, New Decimal), Decimal.Parse(TextBox4.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ESI_Employee_Rate", If(Decimal.TryParse(TextBox5.Text, New Decimal), Decimal.Parse(TextBox5.Text), DBNull.Value))
                    cmd.Parameters.AddWithValue("@ESI_Employer_Rate", If(Decimal.TryParse(TextBox6.Text, New Decimal), Decimal.Parse(TextBox6.Text), DBNull.Value))

                    ' Execute the update command
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Compensation parameters updated successfully")
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error updating compensation parameters: " & ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        HOME.Show()
    End Sub

    Private Sub PARAMETER_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCompanyDetails()
        LoadParameterValues()
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
                        Label10.Text = companyName
                        Label9.Text = companyPHONE
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

    Private Sub LoadParameterValues()
        ' SQL query to get the parameter values
        Dim query As String = "SELECT [PF LIMIT], [PF EMPLOYEE RATE], [PF EMPLOYER RATE], [ESI LIMIT], [ESI EMPLOYEE RATE], [ESI EMPLOYER RATE] FROM [dbo].[PARAMETER]"

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
                        ' Get the parameter values from the reader
                        TextBox1.Text = reader("PF LIMIT").ToString()
                        TextBox2.Text = reader("PF EMPLOYEE RATE").ToString()
                        TextBox3.Text = reader("PF EMPLOYER RATE").ToString()
                        TextBox4.Text = reader("ESI LIMIT").ToString()
                        TextBox5.Text = reader("ESI EMPLOYEE RATE").ToString()
                        TextBox6.Text = reader("ESI EMPLOYER RATE").ToString()
                    Else
                        ' Handle the case where no data is returned
                        MessageBox.Show("No parameter values found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using

            Catch ex As Exception
                ' Handle any errors that might have occurred
                MessageBox.Show("An error occurred while fetching parameter values: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub
End Class
