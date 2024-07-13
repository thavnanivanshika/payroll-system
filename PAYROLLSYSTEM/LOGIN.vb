Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Public Class LOGIN
    Private connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"

    Private Sub LOGIN_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        If CheckBox1.Checked = False Then
            TextBox2.UseSystemPasswordChar = True
        End If
        LoadCompanyDetails()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim username As String = TextBox1.Text.Trim()
        Dim password As String = TextBox2.Text.Trim()
        If String.IsNullOrEmpty(username) OrElse String.IsNullOrEmpty(password) Then
            MessageBox.Show("Please enter both username and password.", "warning",
            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If TextBox1.Text = ("VANSHIKA") And TextBox2.Text = ("123456") Then
            HOME.Show()
            Me.Hide()
        Else
            MessageBox.Show("Invalid username or password.", "warning", MessageBoxButtons.OK,
            MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox2.UseSystemPasswordChar = False
        ElseIf CheckBox1.Checked = False Then
            TextBox2.UseSystemPasswordChar = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Form1.Show()
    End Sub


    Private Sub LoadCompanyDetails()
        ' SQL query to get the company name and address
        Dim query As String = "SELECT [COMPANY NAME] FROM [dbo].[COMPANY]"

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


                        ' Set the labels' text to the company details
                        Label4.Text = companyName

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