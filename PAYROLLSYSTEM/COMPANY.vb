Imports System.Data.SqlClient

Public Class COMPANY
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        HOME.Show()

    End Sub




    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Define your connection string (update as needed)
        Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"

        ' Initialize variables for the new values


        ' Initialize variables for the new values
        Dim newCompanyName As String = TextBox1.Text.Trim()
        Dim newAddress As String = TextBox2.Text.Trim()
        Dim newPhoneNumber As String = TextBox3.Text.Trim()

        ' Initialize variables to hold current values
        Dim currentCompanyName As String = String.Empty
        Dim currentAddress As String = String.Empty
        Dim currentPhoneNumber As String = String.Empty

        ' Retrieve current values from the database
        Using connection As New SqlConnection(connectionString)
            Dim selectCommand As New SqlCommand("SELECT [COMPANY NAME], [ADDRESS], [PHONE] FROM COMPANY", connection)

            connection.Open()
            Using reader As SqlDataReader = selectCommand.ExecuteReader()
                If reader.Read() Then
                    currentCompanyName = reader("COMPANY NAME").ToString()
                    currentAddress = reader("ADDRESS").ToString()
                    currentPhoneNumber = reader("PHONE").ToString()
                End If
            End Using
        End Using

        ' Update the values only if the textboxes are not empty
        If String.IsNullOrEmpty(newCompanyName) Then
            newCompanyName = currentCompanyName
        End If
        If String.IsNullOrEmpty(newAddress) Then
            newAddress = currentAddress
        End If
        If String.IsNullOrEmpty(newPhoneNumber) Then
            newPhoneNumber = currentPhoneNumber
        End If

        ' Update the database with the new values
        Using connection As New SqlConnection(connectionString)
            Dim updateCommand As New SqlCommand("UPDATE COMPANY SET [COMPANY NAME] = @CompanyName, [ADDRESS] = @Address, [PHONE] = @Phone", connection)
            updateCommand.Parameters.AddWithValue("@CompanyName", newCompanyName)
            updateCommand.Parameters.AddWithValue("@Address", newAddress)
            updateCommand.Parameters.AddWithValue("@Phone", newPhoneNumber)

            connection.Open()
            updateCommand.ExecuteNonQuery()
        End Using

        ' Show a message indicating success
        MessageBox.Show("Company details updated successfully.")
    End Sub
End Class

