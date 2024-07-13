Imports System.Data.SqlClient

Public Class HOME
    Private connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"
    Private Sub LOGOUTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LOGOUTToolStripMenuItem.Click
        Dim result As DialogResult
        result = MessageBox.Show("ARE YOU SURE YOU WANT TO LOGOUT", "Confirmation",
MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If result = DialogResult.OK Then
            Me.Close()
            Form1.Show()
        End If
    End Sub

    Private Sub ADDNEWEMPLOYEEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDNEWEMPLOYEEToolStripMenuItem.Click
        Me.Close()
        ADD.Show()
    End Sub

    Private Sub SHOWEMPLOYEEDETAILSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SHOWEMPLOYEEDETAILSToolStripMenuItem.Click
        Me.Close()
        SHOWW.Show()
    End Sub

    Private Sub ADDSALARYSTRUCTUREToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDSALARYSTRUCTUREToolStripMenuItem.Click
        Me.Close()
        SEARCH.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        LEARN.Show()
    End Sub

    Private Sub COMPENSATIONFRAMEWORKToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles COMPENSATIONFRAMEWORKToolStripMenuItem.Click
        Me.Close()
        PARAMETER.Show()
    End Sub

    Private Sub TRANSACTIONToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TRANSACTIONToolStripMenuItem.Click
        Me.Close()
        TRANS.Show()
    End Sub

    Private Sub REPORToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles REPORToolStripMenuItem.Click
        Me.Close()
        REPORT.Show()
    End Sub

    Private Sub EDITSALARYSTRUCTUREToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EDITSALARYSTRUCTUREToolStripMenuItem.Click
        Me.Close()
        SALARY_EDIT.Show()
    End Sub

    Private Sub CompanyBrandingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompanyBrandingToolStripMenuItem.Click
        Me.Close()
        COMPANY.Show()
    End Sub

    Private Sub HOME_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
                        Label2.Text = companyName
                        Label3.Text = companyPHONE
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