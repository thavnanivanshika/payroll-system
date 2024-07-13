Imports System.Data.SqlClient

Public Class SHOWW
    Dim connectionString As String = "Data Source=DESKTOP-DL4DPHJ\SQLEXPRESS;Initial Catalog=PAYROLL;Integrated Security=True;Encrypt=False"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        HOME.Show()
    End Sub

    Private Sub SHOWW_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadDataIntoDataGridView()
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ReadOnly = False
        LoadCompanyDetails()
    End Sub

    Private Sub LoadDataIntoDataGridView()
        Try
            Dim connection As New SqlConnection(connectionString)
            connection.Open()
            Dim query As String = "SELECT * FROM EMPLOYEE1"
            Dim adapter As New SqlDataAdapter(query, connection)
            Dim commandBuilder As New SqlCommandBuilder(adapter)
            Dim dataSet As New DataSet()
            adapter.Fill(dataSet, "EMPLOYEE")
            DataGridView1.DataSource = dataSet.Tables("EMPLOYEE")
            connection.Close()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Update the selected row
        If DataGridView1.SelectedRows.Count > 0 Then
            Try
                Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim employeeCode As String = row.Cells("EMPLOYEE_CODE").Value.ToString() ' Assuming "EMPLOYEE_CODE" is the unique column

                Dim updateQuery As String = "UPDATE EMPLOYEE1 SET "
                Dim parameters As New List(Of SqlParameter)

                For Each cell As DataGridViewCell In row.Cells
                    If cell.OwningColumn.Name <> "EMPLOYEE_CODE" Then ' Skip the unique column
                        updateQuery &= cell.OwningColumn.Name & " = @" & cell.OwningColumn.Name & ", "
                        parameters.Add(New SqlParameter("@" & cell.OwningColumn.Name, cell.Value))
                    End If
                Next

                ' Remove the last comma and add the WHERE clause
                updateQuery = updateQuery.TrimEnd(","c, " "c) & " WHERE EMPLOYEE_CODE = @EMPLOYEE_CODE"
                parameters.Add(New SqlParameter("@EMPLOYEE_CODE", employeeCode))

                Dim connection As New SqlConnection(connectionString)
                Dim command As New SqlCommand(updateQuery, connection)
                command.Parameters.AddRange(parameters.ToArray())
                connection.Open()
                command.ExecuteNonQuery()
                connection.Close()
                MessageBox.Show("Record updated successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("Please select a row to update", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Delete the selected row
        If DataGridView1.SelectedRows.Count > 0 Then
            Try
                Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim employeeCode As String = row.Cells("EMPLOYEE_CODE").Value.ToString() ' Assuming "EMPLOYEE_CODE" is the unique column

                Dim connection As New SqlConnection(connectionString)
                Dim command As New SqlCommand("DELETE FROM EMPLOYEE1 WHERE EMPLOYEE_CODE = @EMPLOYEE_CODE", connection)
                command.Parameters.AddWithValue("@EMPLOYEE_CODE", employeeCode)
                connection.Open()
                command.ExecuteNonQuery()
                connection.Close()
                DataGridView1.Rows.Remove(row)
                MessageBox.Show("Record deleted successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("Please select a row to delete", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
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
