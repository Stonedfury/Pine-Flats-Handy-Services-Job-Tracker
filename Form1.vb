Imports System.Globalization
Imports OfficeOpenXml

Public Class Form1
    Private excelPackage As ExcelPackage

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btn_AddJob.Enabled = False
        ' Set the license context for EPPlus
        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial

        ' Show the OpenFileDialog when the form loads
        OpenFileDialog1.Title = "Select Excel Workbook"
        OpenFileDialog1.Filter = "Excel Files (*.xlsx; *.xls)|*.xlsx; *.xls; *.xlsm|All files (*.*)|*.*"

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim selectedFilePath As String = OpenFileDialog1.FileName

            ' Update Label_File with the name of the selected file
            Label_File.Text = "Selected Excel File: " & System.IO.Path.GetFileName(selectedFilePath)

            ' Load the Excel file and enable the "Add Job" button
            excelPackage = New OfficeOpenXml.ExcelPackage(New System.IO.FileInfo(selectedFilePath))

        Else
            ' User cancelled the dialog
            MessageBox.Show("No file selected. Please select an Excel workbook.")
        End If

        ' Generate a new invoice number
        GenerateInvoiceNumber()

        btn_AddJob.Enabled = True

        ' Set the text of TextBox_AddToday to today's date
        TextBox_AddToday.Text = DateTime.Now.ToString("MM/dd/yyyy")
        TextBox_TodaysDate.Text = DateTime.Now.ToString("MM/dd/yyyy")

        ' Set default values and formatting for textboxes
        TextBox_AddRatePerHour.Text = "$0.00"
        TextBox_AddMaterialCost.Text = "$0.00"
        TextBox_AddDownPayment.Text = "$0.00"
        TextBox_AddState.Text = "CA"
        TextBox_AddLaborHours.Text = "0"
        ' ******************** script for Tab 2 ******************
        LoadInvoices()
        AddHandler TextBox_EditDownPayment.Leave, AddressOf TextBox_EditDownPayment_Leave
    End Sub
    Private Function Today() As String
        Return DateTime.Now.ToString("MM/dd/yyyy")
    End Function

    ' clear the form details on click
    Private Sub ClearJobInformation()
        TextBox_AddClientName.Text = ""
        TextBox_AddCompany.Text = ""
        TextBox_AddPhone1.Text = ""
        TextBox_AddPhone2.Text = ""
        TextBox_AddJobName.Text = ""
        TextBox_AddStreet1.Text = ""
        TextBox_AddStreet2.Text = ""
        TextBox_AddCity.Text = ""
        TextBox_AddZip.Text = ""
        TextBox_AddLaborHours.Text = "0"
        TextBox_AddMaterialCost.Text = "$0.00"
        TextBox_AddRatePerHour.Text = "$0.00"
        TextBox_AddTotalEstimate.Text = "$0.00"
        TextBox_AddDownPayment.Text = "$0.00"
        TextBox_AddStartDate.Text = Today()
        chkbx_Paid.Checked = False
        TextBox_AddJobDetails.Text = ""
        TextBox_AddJobNotes.Text = ""
        TextBox_AddEmail.Text = ""
        TextBox_AddClientName.Focus()
    End Sub
    ' end of clear form details

    'Added to generate a new invoice number on load and on click
    Private Sub GenerateInvoiceNumber()
        ' Use EPPlus to read from the Excel file
        Dim jobsWorksheet As OfficeOpenXml.ExcelWorksheet = excelPackage.Workbook.Worksheets("Jobs")

        ' Find the maximum existing invoice number
        Dim maxInvoiceNumber As Integer = 0
        Dim rowCount As Integer = jobsWorksheet.Dimension.End.Row

        For row As Integer = 2 To rowCount ' Assuming the data starts from row 2
            Dim currentInvoice As String = jobsWorksheet.Cells(row, 2).Text ' Invoice number is in Column B
            If Not String.IsNullOrEmpty(currentInvoice) Then
                ' Extract the numeric part of the invoice number
                Dim numericPart As String = currentInvoice.Substring(currentInvoice.LastIndexOf("-") + 1)
                Dim invoiceNumber As Integer
                If Integer.TryParse(numericPart, invoiceNumber) Then
                    If invoiceNumber > maxInvoiceNumber Then
                        maxInvoiceNumber = invoiceNumber
                    End If
                End If
            End If
        Next

        ' Generate the new invoice number
        Dim newInvoiceNumber As String = "PFHS-" & (maxInvoiceNumber + 1).ToString("0000")

        ' Set the text of TextBox_AddInvoiceNumber to the new invoice number
        TextBox_AddInvoiceNumber.Text = newInvoiceNumber
    End Sub

    ' End of the new sub So I can remove if needed.

    Private Sub LinkLabel_Domain_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel_Domain.LinkClicked
        ' Open the website in the default web browser
        System.Diagnostics.Process.Start("https://pineflatshandyservices.com")
    End Sub

    Private Sub btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click
        Me.Close()
    End Sub

    Private Sub TextBox_AddJobDetails_TextChanged(sender As Object, e As EventArgs) Handles TextBox_AddJobDetails.TextChanged
        ' Update the character count label
        Label_JobDetailsCount.Text = "Character Count: " & TextBox_AddJobDetails.TextLength.ToString()
    End Sub

    Private Sub TextBox_AddJobNotes_TextChanged(sender As Object, e As EventArgs) Handles TextBox_AddJobNotes.TextChanged
        ' Update the character count label
        Label_JobNotesCount.Text = "Character Count: " & TextBox_AddJobNotes.TextLength.ToString()
    End Sub

    Private Sub TextBox_AddLaborHours_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_AddLaborHours.KeyPress
        ' Ensure that only numeric values are allowed
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox_AddRatePerHour_Leave(sender As Object, e As EventArgs) Handles TextBox_AddRatePerHour.Leave
        ' Format as currency when leaving the textbox
        Dim ratePerHour As Double
        If Double.TryParse(TextBox_AddRatePerHour.Text, ratePerHour) Then
            TextBox_AddRatePerHour.Text = ratePerHour.ToString("C2")
        Else
            TextBox_AddRatePerHour.Text = "$0.00"
        End If
    End Sub

    Private Sub TextBox_AddMaterialCost_Leave(sender As Object, e As EventArgs) Handles TextBox_AddMaterialCost.Leave
        ' Format as currency when leaving the textbox
        Dim materialCost As Double
        If Double.TryParse(TextBox_AddMaterialCost.Text, materialCost) Then
            TextBox_AddMaterialCost.Text = materialCost.ToString("C2")
        Else
            TextBox_AddMaterialCost.Text = "$0.00"
        End If
    End Sub

    Private Sub TextBox_AddDownPayment_Leave(sender As Object, e As EventArgs) Handles TextBox_AddDownPayment.Leave
        ' Format as currency when leaving the textbox
        Dim downPayment As Double
        If Double.TryParse(TextBox_AddDownPayment.Text, downPayment) Then
            TextBox_AddDownPayment.Text = downPayment.ToString("C2")
        Else
            TextBox_AddDownPayment.Text = "$0.00"
        End If
    End Sub

    Private Sub btn_CalculateTotal_Click(sender As Object, e As EventArgs) Handles btn_CalculateTotal.Click
        ' Parse values from textboxes
        Dim laborHours As Integer
        Dim ratePerHour As Double
        Dim materialCost As Double

        ' Trim any leading or trailing whitespace from the input
        Dim laborHoursText As String = TextBox_AddLaborHours.Text.Trim()
        Dim ratePerHourText As String = TextBox_AddRatePerHour.Text.Trim().Replace("$", "")
        Dim materialCostText As String = TextBox_AddMaterialCost.Text.Trim().Replace("$", "")

        ' Validate and parse Labor Hours
        If Not Integer.TryParse(laborHoursText, laborHours) Then
            MessageBox.Show("Please enter a valid numeric value for Labor Hours.")
            Exit Sub ' Exit the sub if parsing fails
        End If

        ' Validate and parse Rate Per Hour
        If Not Double.TryParse(ratePerHourText, ratePerHour) Then
            MessageBox.Show("Please enter a valid numeric value for Rate Per Hour.")
            Exit Sub ' Exit the sub if parsing fails
        End If

        ' Validate and parse Material Cost
        If Not Double.TryParse(materialCostText, materialCost) Then
            MessageBox.Show("Please enter a valid numeric value for Material Cost.")
            Exit Sub ' Exit the sub if parsing fails
        End If

        ' Calculate labor cost
        Dim laborCost As Double = laborHours * ratePerHour

        ' Calculate total estimate
        Dim totalEstimate As Double = laborCost + materialCost

        ' Display the total estimate in the appropriate textbox
        TextBox_AddTotalEstimate.Text = totalEstimate.ToString("C2") ' Format as currency with 2 decimal places
    End Sub


    Private Sub AddJob(clientID As Integer, invoiceNumber As String, clientName As String)
        Try
            ' Add a new row to the "Jobs" worksheet
            Dim jobsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Jobs")
            Dim nextRow As Integer = 2 ' Start from row 2 (assuming headers are in row 1)

            ' Find the next available row
            While Not String.IsNullOrEmpty(jobsWorksheet.Cells(nextRow, 1).Text)
                nextRow += 1
            End While

            ' Format phone numbers
            Dim formattedPhone1 As String = FormatPhoneNumber(TextBox_AddPhone1.Text)
            Dim formattedPhone2 As String = FormatPhoneNumber(TextBox_AddPhone2.Text)

            ' Parse labor hours and material cost
            Dim laborHours As Integer
            If Not Integer.TryParse(TextBox_AddLaborHours.Text, laborHours) Then
                MessageBox.Show("Labor hours must be a valid integer.")
                Return
            End If

            ' Remove currency symbols and formatting characters from the material cost textbox
            Dim materialCostText As String = TextBox_AddMaterialCost.Text.Replace("$", "").Replace(",", "")

            ' Parse the cleaned material cost value
            Dim materialCost As Double
            If Not Double.TryParse(materialCostText, NumberStyles.Currency, CultureInfo.CurrentCulture, materialCost) Then
                MessageBox.Show("Material cost must be a valid number.")
                Return
            End If

            ' My edit for rate per hour and down payment
            ' *****************************************
            ' *****************************************
            ' Remove currency symbols and formatting characters from the material cost textbox
            Dim ratePerHourText As String = TextBox_AddRatePerHour.Text.Replace("$", "").Replace(",", "")

            ' Parse the cleaned material cost value
            Dim ratePerHour As Double
            If Not Double.TryParse(ratePerHourText, NumberStyles.Currency, CultureInfo.CurrentCulture, ratePerHour) Then
                MessageBox.Show("Rate per hour must be a valid number.")
                Return
            End If

            ' ******************************************
            ' ******************************************
            ' *****************************************
            ' *****************************************
            ' Remove currency symbols and formatting characters from the material cost textbox
            Dim downPaymentText As String = TextBox_AddDownPayment.Text.Replace("$", "").Replace(",", "")

            ' Parse the cleaned material cost value
            Dim downPayment As Double
            If Not Double.TryParse(downPaymentText, NumberStyles.Currency, CultureInfo.CurrentCulture, downPayment) Then
                MessageBox.Show("Down Payment must be a valid number.")
                Return
            End If

            ' ******************************************
            ' ******************************************
            Dim totalEstimateText As String = TextBox_AddTotalEstimate.Text.Replace("$", "").Replace(",", "")

            ' Parse the cleaned material cost value
            Dim totalEstimate As Double
            If Not Double.TryParse(totalEstimateText, NumberStyles.Currency, CultureInfo.CurrentCulture, totalEstimate) Then
                MessageBox.Show("Total Estimate must be a valid number.")
                Return
            End If

            ' ******************************************
            ' Format the material cost as currency with two decimal places
            Dim formattedMaterialCost As String = materialCost.ToString("C2", CultureInfo.CurrentCulture)


            ' Add data to the new row
            jobsWorksheet.Cells(nextRow, 1).Value = clientID ' Client ID in Column A
            jobsWorksheet.Cells(nextRow, 2).Value = invoiceNumber ' Invoice number in Column B
            jobsWorksheet.Cells(nextRow, 3).Value = clientName ' Client name in Column C
            jobsWorksheet.Cells(nextRow, 4).Value = TextBox_AddCompany.Text ' Company in Column D
            jobsWorksheet.Cells(nextRow, 5).Value = formattedPhone1 ' Phone1 in Column E
            jobsWorksheet.Cells(nextRow, 6).Value = formattedPhone2 ' Phone2 in Column F
            jobsWorksheet.Cells(nextRow, 7).Value = TextBox_AddJobName.Text ' JobName in Column G
            jobsWorksheet.Cells(nextRow, 8).Value = TextBox_AddStreet1.Text ' Street1 in Column H
            jobsWorksheet.Cells(nextRow, 9).Value = TextBox_AddStreet2.Text ' Street2 in Column I
            jobsWorksheet.Cells(nextRow, 10).Value = TextBox_AddCity.Text ' City in Column J
            jobsWorksheet.Cells(nextRow, 11).Value = TextBox_AddState.Text ' State in Column K
            jobsWorksheet.Cells(nextRow, 12).Value = TextBox_AddZip.Text ' Zip in Column L
            jobsWorksheet.Cells(nextRow, 13).Value = laborHours ' Labor Hours in Column M
            Dim material As ExcelRange = jobsWorksheet.Cells(nextRow, 14)
            material.Value = materialCost ' Materials Cost Column N
            material.Style.Numberformat.Format = "$#,##0.00" ' Set cell format to currency with two decimal places
            '********************************
            ' jobsWorksheet.Cells(nextRow, 15).Value = TextBox_AddRatePerHour.Text ' Current Rate per hour Column o
            Dim rate As ExcelRange = jobsWorksheet.Cells(nextRow, 15)

            rate.Value = ratePerHour ' Rate Per Hour Column O
            rate.Style.Numberformat.Format = "$#,##0.00" ' Set cell format to currency with two decimal places
            ' *******************************

            ' jobsWorksheet.Cells(nextRow, 16).Value = TextBox_AddTotalEstimate.Text ' Estimate Totalled Column p
            Dim estimate As ExcelRange = jobsWorksheet.Cells(nextRow, 16)

            estimate.Value = totalEstimate ' Total Estimate Column P
            estimate.Style.Numberformat.Format = "$#,##0.00" ' Set cell format to currency with two decimal places
            ' *******************************
            ' *******************************
            'jobsWorksheet.Cells(nextRow, 17).Value = TextBox_AddDownPayment.Text ' Down payment towards total Column q
            Dim downPay As ExcelRange = jobsWorksheet.Cells(nextRow, 17)
            downPay.Value = downPayment ' Down Payment Column P
            downPay.Style.Numberformat.Format = "$#,##0.00" ' Set cell format to currency with two decimal places
            ' *******************************

            ' Parse and set the date value
            Dim startDate As Date
            If Date.TryParse(TextBox_AddStartDate.Text, startDate) Then
                Dim startDateCell As ExcelRange = jobsWorksheet.Cells(nextRow, 18)
                startDateCell.Value = startDate ' Start Date in Column R
                startDateCell.Style.Numberformat.Format = "mm/dd/yyyy" ' Set cell format to date
            Else
                MessageBox.Show("Invalid start date.")
                Return
            End If

            jobsWorksheet.Cells(nextRow, 19).Value = If(chkbx_Paid.Checked, "Yes", "No") ' Paid in full Column S
            jobsWorksheet.Cells(nextRow, 20).Value = TextBox_AddJobDetails.Text ' Job Details Column T
            jobsWorksheet.Cells(nextRow, 21).Value = TextBox_AddJobNotes.Text ' Job Notes Column U

            ' Set the date added
            Dim dateAddedCell As ExcelRange = jobsWorksheet.Cells(nextRow, 22)
            dateAddedCell.Value = DateTime.Today ' Date added column V
            dateAddedCell.Style.Numberformat.Format = "mm/dd/yyyy" ' Set cell format to date

            jobsWorksheet.Cells(nextRow, 24).Value = TextBox_AddEmail.Text ' Client Email Column X

            ' Save changes to the Excel file
            excelPackage.Save()

            ' Inform the user that the job has been added successfully
            MessageBox.Show("Job added successfully.")

        Catch ex As Exception
            MessageBox.Show("An error occurred while adding the job: " & ex.Message)
        End Try
    End Sub



    Private Function FindClientID(firstName As String, lastName As String) As Integer?
        ' Find client ID based on first and last names in the "Clients" worksheet
        Dim clientsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Clients")

        ' Check if the worksheet exists
        If clientsWorksheet Is Nothing Then
            MessageBox.Show("Clients worksheet not found.")
            Return Nothing
        End If

        ' Check if the worksheet has dimensions
        If clientsWorksheet.Dimension Is Nothing Then
            MessageBox.Show("Clients worksheet is empty or not properly formatted.")
            Return Nothing
        End If

        Dim lastRow As Integer = clientsWorksheet.Dimension.End.Row

        For row As Integer = 3 To lastRow ' Assuming data starts from row 3
            Dim currentFirstName As String = clientsWorksheet.Cells(row, 2).Text ' Use .Text to avoid .ToString() on null values
            Dim currentLastName As String = clientsWorksheet.Cells(row, 3).Text ' Use .Text to avoid .ToString() on null values

            ' Log the current row values for debugging
            Debug.WriteLine($"Row {row}: FirstName={currentFirstName}, LastName={currentLastName}")

            ' Ensure that the current row cells are not empty
            If String.IsNullOrEmpty(currentFirstName) Or String.IsNullOrEmpty(currentLastName) Then
                Continue For
            End If

            If currentFirstName.ToLower() = firstName.ToLower() AndAlso currentLastName.ToLower() = lastName.ToLower() Then
                ' Match found, return the client ID from column A
                Dim clientID As Integer
                If Integer.TryParse(clientsWorksheet.Cells(row, 1).Text, clientID) Then
                    Return clientID ' Column A
                Else
                    MessageBox.Show("Client ID is not a valid integer.")
                    Return Nothing ' Indicate an error with Nothing
                End If
            End If
        Next

        ' No match found
        Return Nothing
    End Function

    Private Function AddNewClient(firstName As String, lastName As String, phone1 As String, company As String) As Integer
        Try
            Dim clientsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Clients")

            ' Check if the worksheet exists
            If clientsWorksheet Is Nothing Then
                MessageBox.Show("Clients worksheet not found.")
                Return -1
            End If

            ' Check if the worksheet has dimensions
            If clientsWorksheet.Dimension Is Nothing Then
                MessageBox.Show("Clients worksheet is empty or not properly formatted.")
                Return -1
            End If

            ' Find the last row with data
            Dim lastRow As Integer = clientsWorksheet.Dimension.End.Row
            Dim foundLastRow As Integer = 1 ' Start from the first row after headers (assuming headers are in row 1)

            ' Check if the client already exists
            For row As Integer = 2 To lastRow ' Assuming data starts from row 2
                Dim existingFirstName As String = clientsWorksheet.Cells(row, 2).Text
                Dim existingLastName As String = clientsWorksheet.Cells(row, 3).Text

                If existingFirstName.ToLower() = firstName.ToLower() AndAlso existingLastName.ToLower() = lastName.ToLower() Then
                    MessageBox.Show("Found customer: " & firstName & " " & lastName)
                    Return -1 ' Indicate that the client was found and no new entry was made
                End If
            Next

            ' Find the last row with actual data
            For row As Integer = lastRow To 1 Step -1
                If Not String.IsNullOrEmpty(clientsWorksheet.Cells(row, 1).Text) Then
                    foundLastRow = row
                    Exit For
                End If
            Next

            ' Determine the new row to add data
            Dim newRow As Integer = foundLastRow + 1
            MessageBox.Show("Adding new client to row: " & newRow)

            ' Find the last Client ID and increment it, or start at 1 if no valid Client ID is found
            Dim lastClientID As Integer
            If foundLastRow > 1 AndAlso Integer.TryParse(clientsWorksheet.Cells(foundLastRow, 1).Value.ToString(), lastClientID) Then
                Dim newClientID As Integer = lastClientID + 1

                ' Add the new client details to the next row
                clientsWorksheet.Cells(newRow, 1).Value = newClientID ' Client ID in Column A
                clientsWorksheet.Cells(newRow, 2).Value = firstName ' First Name in Column B
                clientsWorksheet.Cells(newRow, 3).Value = lastName ' Last Name in Column C
                clientsWorksheet.Cells(newRow, 4).Value = phone1 ' Phone1 in Column D
                clientsWorksheet.Cells(newRow, 5).Value = company ' Company in Column E
                clientsWorksheet.Cells(newRow, 6).Value = TextBox_AddEmail.Text ' Email in Column F
                clientsWorksheet.Cells(newRow, 13).Value = Today() ' Date Client is added Column M

                ' Save the changes to the Excel file
                excelPackage.Save()

                Return newClientID
            Else
                ' If no valid Client ID is found, start from 1
                Dim newClientID As Integer = 1

                ' Add the new client details to the next row
                clientsWorksheet.Cells(newRow, 1).Value = newClientID ' Client ID in Column A
                clientsWorksheet.Cells(newRow, 2).Value = firstName ' First Name in Column B
                clientsWorksheet.Cells(newRow, 3).Value = lastName ' Last Name in Column C
                clientsWorksheet.Cells(newRow, 4).Value = phone1 ' Phone1 in Column D
                clientsWorksheet.Cells(newRow, 5).Value = company ' Company in Column E
                clientsWorksheet.Cells(newRow, 6).Value = TextBox_AddEmail.Text ' Email in Column F
                clientsWorksheet.Cells(newRow, 13).Value = Today() ' Date Client is added Column M
                ' Save the changes to the Excel file
                excelPackage.Save()

                Return newClientID
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred while adding the new client: " & ex.Message)
            Return -1
        End Try
    End Function

    Private Function FormatPhoneNumber(phoneNumber As String) As String
        ' Remove non-numeric characters
        Dim digitsOnly As String = New String(phoneNumber.Where(AddressOf Char.IsDigit).ToArray())

        ' Check if the number of digits is 10 (standard US phone number)
        If digitsOnly.Length = 10 Then
            ' Format as (123) 456-7890
            Return String.Format("({0}) {1}-{2}", digitsOnly.Substring(0, 3), digitsOnly.Substring(3, 3), digitsOnly.Substring(6, 4))
        Else
            ' Return the original input if it's not a valid 10-digit phone number
            Return phoneNumber
        End If
    End Function

    Private Sub FormatPhoneNumberWhileTyping(ByRef textBox As TextBox)
        ' Get the digits from the textbox
        Dim digitsOnly As String = New String(textBox.Text.Where(AddressOf Char.IsDigit).ToArray())

        ' Format the digits into a phone number format
        If digitsOnly.Length <= 3 Then
            textBox.Text = digitsOnly
        ElseIf digitsOnly.Length <= 6 Then
            textBox.Text = String.Format("({0}) {1}", digitsOnly.Substring(0, 3), digitsOnly.Substring(3))
        ElseIf digitsOnly.Length <= 10 Then
            textBox.Text = String.Format("({0}) {1}-{2}", digitsOnly.Substring(0, 3), digitsOnly.Substring(3, 3), digitsOnly.Substring(6))
        Else
            textBox.Text = String.Format("({0}) {1}-{2}", digitsOnly.Substring(0, 3), digitsOnly.Substring(3, 3), digitsOnly.Substring(6, 4))
        End If

        ' Move the cursor to the end of the text
        textBox.SelectionStart = textBox.Text.Length
    End Sub

    Private Sub TextBox_AddPhone1_TextChanged(sender As Object, e As EventArgs) Handles TextBox_AddPhone1.TextChanged
        FormatPhoneNumberWhileTyping(TextBox_AddPhone1)
    End Sub

    Private Sub TextBox_AddPhone2_TextChanged(sender As Object, e As EventArgs) Handles TextBox_AddPhone2.TextChanged
        FormatPhoneNumberWhileTyping(TextBox_AddPhone2)
    End Sub

    Private Sub btn_AddJob_Click(sender As Object, e As EventArgs) Handles btn_AddJob.Click
        ' Split the full name into first and last names
        Dim fullName As String = TextBox_AddClientName.Text.Trim()
        Dim names() As String = fullName.Split(New Char() {" "}, StringSplitOptions.RemoveEmptyEntries)

        If names.Length >= 2 Then ' Ensure we have at least a first and last name
            Dim firstName As String = names(0)
            Dim lastName As String = names(1)

            ' Search for the matching client in the "Clients" worksheet
            Dim clientID As Integer? = FindClientID(firstName, lastName)

            If Not clientID.HasValue Then ' Client not found, add a new client
                Dim formattedPhone1 As String = FormatPhoneNumber(TextBox_AddPhone1.Text)
                Dim newClientID As Integer = AddNewClient(firstName, lastName, formattedPhone1, TextBox_AddCompany.Text)

                If newClientID <> -1 Then ' New client added successfully
                    AddJob(newClientID, TextBox_AddInvoiceNumber.Text.Trim(), fullName)
                End If
            Else
                ' Client found, add job
                AddJob(clientID.Value, TextBox_AddInvoiceNumber.Text.Trim(), fullName)
            End If
        Else
            MessageBox.Show("Please enter both first and last names.")
        End If
        If String.IsNullOrEmpty(TextBox_AddLaborHours.Text) OrElse Val(TextBox_AddLaborHours.Text) = 0 Then
            MessageBox.Show("Please enter a valid number for labor hours.")
            Return ' Exit the method if labor hours are not provided or zero
        End If
        GenerateInvoiceNumber()
        ClearJobInformation()
        LoadInvoices()
    End Sub
    Private Sub TextBox_AddStartDate_Click(sender As Object, e As EventArgs) Handles TextBox_AddStartDate.Click
        ' Show the DateTimePicker near the TextBox
        DateTimePicker_AddStartDate.Location = New Point(TextBox_AddStartDate.Left, TextBox_AddStartDate.Top + TextBox_AddStartDate.Height)
        DateTimePicker_AddStartDate.Visible = True
        DateTimePicker_AddStartDate.Focus()
    End Sub

    Private Sub DateTimePicker_AddStartDate_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker_AddStartDate.ValueChanged
        ' Update the TextBox with the selected date
        TextBox_AddStartDate.Text = DateTimePicker_AddStartDate.Value.ToString("MM/dd/yyyy")
        ' Hide the DateTimePicker after selection
        DateTimePicker_AddStartDate.Visible = False
    End Sub

    Private Sub click_Logo_Click(sender As Object, e As EventArgs) Handles click_Logo.Click
        System.Diagnostics.Process.Start("https://pineflatshandyservices.com")
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        MessageBox.Show("Sure you want to clear everything and cancel")
        ClearJobInformation()
    End Sub
    ' ************************************************
    ' *********** START TAB 2 ************************
    ' ************************************************
    Private Sub LoadInvoices()
        ' Check if excelPackage is initialized
        If excelPackage Is Nothing Then
            MessageBox.Show("Excel package is not initialized.")
            Return
        End If

        ' Access the "Jobs" worksheet
        Dim jobsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Jobs")
        If jobsWorksheet Is Nothing Then
            MessageBox.Show("Jobs worksheet not found.")
            Return
        End If

        ' Clear existing items in the list box
        ListBox_ListInvoices.Items.Clear()

        ' Iterate through the rows in the worksheet and add invoice numbers and client names to the list box
        Dim rowCount As Integer = jobsWorksheet.Dimension.End.Row
        For row As Integer = 2 To rowCount
            Dim invoiceNumber As String = jobsWorksheet.Cells(row, 2).Text ' Assuming invoice number is in Column B
            Dim clientName As String = jobsWorksheet.Cells(row, 3).Text ' Assuming client name is in Column C
            ListBox_ListInvoices.Items.Add(invoiceNumber & " - " & clientName)
        Next

        ' Optionally, select the first item in the list box
        If ListBox_ListInvoices.Items.Count > 0 Then
            ListBox_ListInvoices.SelectedIndex = 0
        End If
    End Sub

    Private Sub ListBox_ListInvoices_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox_ListInvoices.SelectedIndexChanged
        ' Check if an item is selected
        If ListBox_ListInvoices.SelectedIndex = -1 Then
            MessageBox.Show("Please select an invoice.")
            Return
        End If

        ' Get the selected invoice number from the list box
        Dim selectedText As String = ListBox_ListInvoices.SelectedItem.ToString()
        Dim invoiceNumber As String = selectedText.Split(" - ")(0).Trim()

        ' Access the "Jobs" worksheet
        Dim jobsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Jobs")
        If jobsWorksheet Is Nothing Then
            MessageBox.Show("Jobs worksheet not found.")
            Return
        End If

        ' Find the row with the selected invoice number
        Dim rowCount As Integer = jobsWorksheet.Dimension.End.Row
        For row As Integer = 2 To rowCount
            If jobsWorksheet.Cells(row, 2).Text = invoiceNumber Then
                ' Populate the form fields with the details from the selected invoice row
                'TextBox_AddInvoiceNumber.Text = jobsWorksheet.Cells(row, 2).Text ' Invoice number in Column B
                TextBox_EditClientName.Text = jobsWorksheet.Cells(row, 3).Text ' Client name in Column C
                TextBox_EditCompany.Text = jobsWorksheet.Cells(row, 4).Text ' Company in Column D
                TextBox_EditPhone1.Text = jobsWorksheet.Cells(row, 5).Text ' Phone1 in Column E
                TextBox_EditPhone2.Text = jobsWorksheet.Cells(row, 6).Text ' Phone2 in Column F
                TextBox_EditJobName.Text = jobsWorksheet.Cells(row, 7).Text ' JobName in Column G
                TextBox_EditStreet1.Text = jobsWorksheet.Cells(row, 8).Text ' Street1 in Column H
                TextBox_EditStreet2.Text = jobsWorksheet.Cells(row, 9).Text ' Street2 in Column I
                TextBox_EditCity.Text = jobsWorksheet.Cells(row, 10).Text ' City in Column J
                TextBox_EditState.Text = jobsWorksheet.Cells(row, 11).Text ' State in Column K
                TextBox_EditZip.Text = jobsWorksheet.Cells(row, 12).Text ' Zip in Column L
                TextBox_EditLaborHours.Text = jobsWorksheet.Cells(row, 13).Text ' Labor Hours in Column M
                TextBox_EditMaterialCost.Text = FormatCurrency(jobsWorksheet.Cells(row, 14).Text) ' Material Cost in Column N
                TextBox_EditRatePerHour.Text = FormatCurrency(jobsWorksheet.Cells(row, 15).Text) ' Rate Per Hour in Column O
                TextBox_EditEstimateTotal.Text = FormatCurrency(jobsWorksheet.Cells(row, 16).Text) ' Total Estimate in Column P
                TextBox_EditDownPayment.Text = FormatCurrency(jobsWorksheet.Cells(row, 17).Text) ' Down Payment in Column Q
                TextBox_EditStartDate.Text = jobsWorksheet.Cells(row, 18).Text ' Start Date in Column R
                chkBox_EditPaid.Checked = (jobsWorksheet.Cells(row, 19).Text = "Yes") ' Paid in Column S
                TextBox_EditJobDetails.Text = jobsWorksheet.Cells(row, 20).Text ' Job Details in Column T
                TextBox_EditJobNotes.Text = jobsWorksheet.Cells(row, 21).Text ' Job Notes in Column U
                TextBox_EditDateAdded.Text = jobsWorksheet.Cells(row, 22).Text ' Date Added in Column V
                TextBox_EditDate.Text = Today()
                TextBox_EditClientEmail.Text = jobsWorksheet.Cells(row, 24).Text ' Email in Column X

                ' Calculate and display the balance
                Dim estimateTotal As Double
                Dim downPayment As Double
                Dim balance As Double

                Double.TryParse(jobsWorksheet.Cells(row, 16).Text.Replace("$", "").Replace(",", ""), estimateTotal)
                Double.TryParse(jobsWorksheet.Cells(row, 17).Text.Replace("$", "").Replace(",", ""), downPayment)

                If chkBox_EditPaid.Checked Then
                    balance = 0.0
                Else
                    If downPayment = 0 Then
                        balance = estimateTotal
                    Else
                        balance = estimateTotal - downPayment
                    End If
                End If

                TextBox_EditBalance.Text = FormatCurrency(balance)
                Exit For
            End If
        Next
    End Sub



    Private Function FormatCurrency(value As String) As String
        Dim parsedValue As Double
        If Double.TryParse(value, parsedValue) Then
            Return parsedValue.ToString("C2")
        Else
            Return value
        End If
    End Function
    Private Sub TextBox_EditStartDate_Click(sender As Object, e As EventArgs) Handles TextBox_EditStartDate.Click
        ' Show the DateTimePicker near the TextBox
        DateTimePicker_EditStartDate.Location = New Point(TextBox_EditStartDate.Left, TextBox_EditStartDate.Top + TextBox_EditStartDate.Height)
        DateTimePicker_EditStartDate.Visible = True
        DateTimePicker_EditStartDate.Focus()
    End Sub

    Private Sub DateTimePicker_EditStartDate_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker_EditStartDate.ValueChanged
        ' Update the TextBox with the selected date
        TextBox_EditStartDate.Text = DateTimePicker_EditStartDate.Value.ToString("MM/dd/yyyy")
        ' Hide the DateTimePicker after selection
        DateTimePicker_EditStartDate.Visible = False
    End Sub

    Private Sub TextBox_EditFinishDate_Click(sender As Object, e As EventArgs) Handles TextBox_EditFinishDate.Click
        ' Show the DateTimePicker near the TextBox
        DateTimePicker_EditFinishDate.Location = New Point(TextBox_EditFinishDate.Left, TextBox_EditFinishDate.Top + TextBox_EditFinishDate.Height)
        DateTimePicker_EditFinishDate.Visible = True
        DateTimePicker_EditFinishDate.Focus()
    End Sub

    Private Sub DateTimePicker_EditEditFinishDate_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker_EditFinishDate.ValueChanged
        ' Update the TextBox with the selected date
        TextBox_EditFinishDate.Text = DateTimePicker_EditFinishDate.Value.ToString("MM/dd/yyyy")
        ' Hide the DateTimePicker after selection
        DateTimePicker_EditFinishDate.Visible = False
    End Sub

    Private Sub btn_UpdateJob_Click(sender As Object, e As EventArgs) Handles btn_UpdateJob.Click
        ' Check if an item is selected
        If ListBox_ListInvoices.SelectedIndex = -1 Then
            MessageBox.Show("Please select an invoice.")
            Return
        End If

        ' Get the selected invoice number from the list box
        Dim selectedText As String = ListBox_ListInvoices.SelectedItem.ToString()
        Dim invoiceNumber As String = selectedText.Split(" - ")(0).Trim()

        ' Access the "Jobs" worksheet
        Dim jobsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Jobs")
        If jobsWorksheet Is Nothing Then
            MessageBox.Show("Jobs worksheet not found.")
            Return
        End If

        ' Find the row with the selected invoice number
        Dim rowCount As Integer = jobsWorksheet.Dimension.End.Row
        For row As Integer = 2 To rowCount
            If jobsWorksheet.Cells(row, 2).Text = invoiceNumber Then
                ' Update the worksheet with the values from the form fields
                jobsWorksheet.Cells(row, 3).Value = TextBox_EditClientName.Text ' Client name in Column C
                jobsWorksheet.Cells(row, 4).Value = TextBox_EditCompany.Text ' Company in Column D
                jobsWorksheet.Cells(row, 5).Value = TextBox_EditPhone1.Text ' Phone1 in Column E
                jobsWorksheet.Cells(row, 6).Value = TextBox_EditPhone2.Text ' Phone2 in Column F
                jobsWorksheet.Cells(row, 7).Value = TextBox_EditJobName.Text ' JobName in Column G
                jobsWorksheet.Cells(row, 8).Value = TextBox_EditStreet1.Text ' Street1 in Column H
                jobsWorksheet.Cells(row, 9).Value = TextBox_EditStreet2.Text ' Street2 in Column I
                jobsWorksheet.Cells(row, 10).Value = TextBox_EditCity.Text ' City in Column J
                jobsWorksheet.Cells(row, 11).Value = TextBox_EditState.Text ' State in Column K
                jobsWorksheet.Cells(row, 12).Value = TextBox_EditZip.Text ' Zip in Column L
                jobsWorksheet.Cells(row, 13).Value = TextBox_EditLaborHours.Text ' Labor Hours in Column M

                ' Parse and format the currency fields
                Dim materialCost As Double
                If Double.TryParse(TextBox_EditMaterialCost.Text.Replace("$", "").Replace(",", ""), materialCost) Then
                    jobsWorksheet.Cells(row, 14).Value = materialCost ' Material Cost in Column N
                End If

                Dim ratePerHour As Double
                If Double.TryParse(TextBox_EditRatePerHour.Text.Replace("$", "").Replace(",", ""), ratePerHour) Then
                    jobsWorksheet.Cells(row, 15).Value = ratePerHour ' Rate Per Hour in Column O
                End If

                Dim estimateTotal As Double
                If Double.TryParse(TextBox_EditEstimateTotal.Text.Replace("$", "").Replace(",", ""), estimateTotal) Then
                    jobsWorksheet.Cells(row, 16).Value = estimateTotal ' Total Estimate in Column P
                End If

                Dim downPayment As Double
                If Double.TryParse(TextBox_EditDownPayment.Text.Replace("$", "").Replace(",", ""), downPayment) Then
                    jobsWorksheet.Cells(row, 17).Value = downPayment ' Down Payment in Column Q
                End If

                jobsWorksheet.Cells(row, 18).Value = TextBox_EditStartDate.Text ' Start Date in Column R
                jobsWorksheet.Cells(row, 19).Value = If(chkBox_EditPaid.Checked, "Yes", "No") ' Paid in Column S
                jobsWorksheet.Cells(row, 20).Value = TextBox_EditJobDetails.Text ' Job Details in Column T
                jobsWorksheet.Cells(row, 21).Value = TextBox_EditJobNotes.Text ' Job Notes in Column U
                jobsWorksheet.Cells(row, 22).Value = TextBox_EditDateAdded.Text ' Date Added in Column V
                jobsWorksheet.Cells(row, 23).Value = TextBox_EditDate.Text ' Date of Last Update in Column W
                jobsWorksheet.Cells(row, 24).Value = TextBox_EditClientEmail.Text ' Email in Column X

                ' Save the changes to the Excel file
                excelPackage.Save()

                MessageBox.Show("Job details updated successfully.")
                Exit For
            End If
        Next
        LoadInvoices()
    End Sub

    Private Sub btn_CalculateEstimate_Click(sender As Object, e As EventArgs) Handles btn_CalculateEstimate.Click
        ' Retrieve and validate labor hours
        Dim laborHours As Double
        If Not Double.TryParse(TextBox_EditLaborHours.Text, laborHours) OrElse laborHours = 0 Then
            MessageBox.Show("Labor hours cannot be blank or zero.")
            Return
        End If

        ' Retrieve and validate rate per hour
        Dim ratePerHour As Double
        If Not Double.TryParse(TextBox_EditRatePerHour.Text.Replace("$", "").Replace(",", ""), ratePerHour) OrElse ratePerHour = 0 Then
            MessageBox.Show("Rate per hour cannot be blank or zero.")
            Return
        End If

        ' Retrieve material cost
        Dim materialCost As Double
        If Not Double.TryParse(TextBox_EditMaterialCost.Text.Replace("$", "").Replace(",", ""), materialCost) Then
            materialCost = 0 ' Default to 0 if parsing fails
        End If

        ' Calculate the estimate
        Dim estimateTotal As Double = (laborHours * ratePerHour) + materialCost

        ' Display the result in the estimate total text box
        TextBox_EditEstimateTotal.Text = FormatCurrency(estimateTotal)
    End Sub
    Private Sub TextBox_EditDownPayment_Leave(sender As Object, e As EventArgs)
        ' Calculate and update the balance
        Dim estimateTotal As Double
        Dim downPayment As Double
        Dim balance As Double

        ' Parse the values from the textboxes
        Double.TryParse(TextBox_EditEstimateTotal.Text.Replace("$", "").Replace(",", ""), estimateTotal)
        Double.TryParse(TextBox_EditDownPayment.Text.Replace("$", "").Replace(",", ""), downPayment)

        ' Calculate the balance
        If downPayment = 0 Then
            balance = estimateTotal
        Else
            balance = estimateTotal - downPayment
        End If

        ' Update the balance textbox
        TextBox_EditBalance.Text = FormatCurrency(balance)

        ' Check the checkbox if the balance is zero
        If balance = 0 Then
            chkBox_EditPaid.Checked = True
        Else
            chkBox_EditPaid.Checked = False
        End If
    End Sub

    Private Sub chkBox_EditPaid_CheckedChanged(sender As Object, e As EventArgs) Handles chkBox_EditPaid.CheckedChanged
        ' Calculate and update the balance when the Paid checkbox state changes
        If chkBox_EditPaid.Checked Then
            TextBox_EditBalance.Text = "$0.00"
        Else
            ' Recalculate the balance if the checkbox is unchecked
            UpdateBalance()
        End If
    End Sub

    Private Sub UpdateBalance()
        ' Calculate the balance based on the total estimate and down payment
        Dim estimateTotal As Double
        Dim downPayment As Double
        Dim balance As Double

        ' Parse the values from the respective textboxes
        Double.TryParse(TextBox_EditEstimateTotal.Text.Replace("$", "").Replace(",", ""), estimateTotal)
        Double.TryParse(TextBox_EditDownPayment.Text.Replace("$", "").Replace(",", ""), downPayment)

        ' Calculate the balance
        If downPayment = 0 Then
            balance = estimateTotal
        Else
            balance = estimateTotal - downPayment
        End If

        ' Update the balance textbox with the formatted currency
        TextBox_EditBalance.Text = balance.ToString("C2")
    End Sub

    ' ***************************************
    ' ************ TAB 3 ********************

    Private Sub TextBox_ClientPHone1_TextChanged(sender As Object, e As EventArgs) Handles TextBox_ClientPhone1.TextChanged
        FormatPhoneNumberWhileTyping(TextBox_ClientPhone1)
    End Sub

    Private Sub TextBox_ClientPHone2_TextChanged(sender As Object, e As EventArgs) Handles TextBox_ClientPhone2.TextChanged
        FormatPhoneNumberWhileTyping(TextBox_ClientPhone2)
    End Sub

    Private Sub LoadClientInvoicesAndCalculateBalance()
        ' Get the client ID
        Dim clientID As String = TextBox_ClientID.Text
        If String.IsNullOrEmpty(clientID) Then
            Return
        End If

        ' Access the "Jobs" worksheet
        Dim jobsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Jobs")
        If jobsWorksheet Is Nothing Then
            MessageBox.Show("Jobs worksheet not found.")
            Return
        End If

        ' Initialize the balance and invoice list
        Dim totalBalance As Double = 0.0
        ListView_ClientInvoices.Items.Clear()

        ' Find all rows with the matching ClientID
        Dim jobRowCount As Integer = jobsWorksheet.Dimension.End.Row
        For jobRow As Integer = 2 To jobRowCount ' Assuming the data starts from row 2
            If jobsWorksheet.Cells(jobRow, 1).Text = clientID Then
                ' Load the invoice into the list view
                Dim invoiceNumber As String = jobsWorksheet.Cells(jobRow, 2).Text
                Dim paidStatus As String = jobsWorksheet.Cells(jobRow, 19).Text
                Dim listViewItem As New ListViewItem(invoiceNumber)
                listViewItem.SubItems.Add(paidStatus)
                ListView_ClientInvoices.Items.Add(listViewItem)

                ' Calculate the balance if the invoice is marked as "No" in the "Paid" column
                If paidStatus.Equals("No", StringComparison.OrdinalIgnoreCase) Then
                    Dim estimate As Double
                    Dim downPayment As Double

                    Double.TryParse(jobsWorksheet.Cells(jobRow, 16).Text.Replace("$", "").Replace(",", ""), estimate)
                    Double.TryParse(jobsWorksheet.Cells(jobRow, 17).Text.Replace("$", "").Replace(",", ""), downPayment)

                    totalBalance += (estimate - downPayment)
                End If
            End If
        Next

        ' Display the calculated balance
        TextBox_ClientBalance.Text = totalBalance.ToString("C")
    End Sub

    Private Sub ClearClientForm()
        ' Clear all form fields related to the client
        TextBox_ClientID.Clear()
        TextBox_ClientPhone1.Clear()
        TextBox_ClientCompany.Clear()
        TextBox_ClientEmail.Clear()
        TextBox_Street1.Clear()
        TextBox_Street2.Clear()
        TextBox_ClientCity.Clear()
        TextBox_ClientState.Clear()
        TextBox_ClientZip.Clear()
        TextBox_ClientComments.Clear()
        TextBox_DateAdded.Clear()
        TextBox_DateLastEdited.Clear()
        TextBox_ClientPhone2.Clear()
        ListView_ClientInvoices.Items.Clear()
        TextBox_ClientBalance.Clear()
        TextBox_TodaysDate.Text = Today()
        TextBox_Search.Text = "Search..."
    End Sub

    Private Sub btn_ClearAll_Click(sender As Object, e As EventArgs) Handles btn_ClearAll.Click
        TextBox_ClientName.Text = ""
        ClearClientForm()
    End Sub

    Private Sub TextBox_ClientName_Leave(sender As Object, e As EventArgs) Handles TextBox_ClientName.Leave
        ' Get the full name from the text box
        Dim fullName As String = TextBox_ClientName.Text.Trim()

        ' Check if the full name is empty
        If String.IsNullOrEmpty(fullName) Then
            ClearClientForm()
            Return
        End If

        ' Split the full name into first and last name
        Dim names() As String = fullName.Split(" "c)
        If names.Length < 2 Then
            MessageBox.Show("Please enter both first and last name.")
            Return
        End If

        Dim firstName As String = names(0).Trim()
        Dim lastName As String = names(1).Trim()

        ' Access the "Clients" worksheet
        Dim clientsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Clients")
        If clientsWorksheet Is Nothing Then
            MessageBox.Show("Clients worksheet not found.")
            Return
        End If

        ' Initialize a flag to indicate if a match is found
        Dim matchFound As Boolean = False
        Dim clientRow As Integer = -1

        ' Find the row with the matching first and last name
        Dim rowCount As Integer = clientsWorksheet.Dimension.End.Row
        For row As Integer = 3 To rowCount ' Assuming the data starts from row 3
            Dim clientFirstName As String = clientsWorksheet.Cells(row, 2).Text
            Dim clientLastName As String = clientsWorksheet.Cells(row, 3).Text

            If clientFirstName.Equals(firstName, StringComparison.OrdinalIgnoreCase) AndAlso
       clientLastName.Equals(lastName, StringComparison.OrdinalIgnoreCase) Then
                matchFound = True
                clientRow = row
                Exit For
            End If
        Next

        If matchFound Then
            ' Load the existing client details into the form
            TextBox_ClientID.Text = clientsWorksheet.Cells(clientRow, 1).Text
            TextBox_ClientPhone1.Text = clientsWorksheet.Cells(clientRow, 4).Text
            TextBox_ClientCompany.Text = clientsWorksheet.Cells(clientRow, 5).Text
            TextBox_ClientEmail.Text = clientsWorksheet.Cells(clientRow, 6).Text
            TextBox_Street1.Text = clientsWorksheet.Cells(clientRow, 7).Text
            TextBox_Street2.Text = clientsWorksheet.Cells(clientRow, 8).Text
            TextBox_ClientCity.Text = clientsWorksheet.Cells(clientRow, 9).Text
            TextBox_ClientState.Text = clientsWorksheet.Cells(clientRow, 10).Text
            TextBox_ClientZip.Text = clientsWorksheet.Cells(clientRow, 11).Text
            TextBox_ClientComments.Text = clientsWorksheet.Cells(clientRow, 14).Text
            TextBox_DateAdded.Text = clientsWorksheet.Cells(clientRow, 13).Text
            TextBox_DateLastEdited.Text = DateTime.Now.ToString("MM/dd/yyyy")
            TextBox_ClientPhone2.Text = clientsWorksheet.Cells(clientRow, 16).Text

            ' Load matching invoices and calculate balance
            LoadClientInvoicesAndCalculateBalance()

            MessageBox.Show("Client found and details loaded.")
        Else
            ClearClientForm()
            MessageBox.Show("Client not found.")
        End If
    End Sub

    Private Sub btn_AddClient_Click(sender As Object, e As EventArgs) Handles btn_AddClient.Click
        ' Ensure the client name is not empty
        Dim fullName As String = TextBox_ClientName.Text.Trim()

        If String.IsNullOrEmpty(fullName) Then
            MessageBox.Show("Client name cannot be empty.")
            Return
        End If

        ' Split the full name into first and last name
        Dim names() As String = fullName.Split(" "c)
        If names.Length < 2 Then
            MessageBox.Show("Please enter both first and last name.")
            Return
        End If

        Dim firstName As String = names(0).Trim()
        Dim lastName As String = names(1).Trim()

        ' Access the "Clients" worksheet
        Dim clientsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Clients")
        If clientsWorksheet Is Nothing Then
            MessageBox.Show("Clients worksheet not found.")
            Return
        End If

        ' Initialize a flag to indicate if a match is found
        Dim matchFound As Boolean = False
        Dim clientRow As Integer = -1

        ' Find the row with the matching first and last name
        Dim rowCount As Integer = clientsWorksheet.Dimension.End.Row
        For row As Integer = 3 To rowCount ' Assuming the data starts from row 3
            Dim clientFirstName As String = clientsWorksheet.Cells(row, 2).Text
            Dim clientLastName As String = clientsWorksheet.Cells(row, 3).Text

            If clientFirstName.Equals(firstName, StringComparison.OrdinalIgnoreCase) AndAlso
       clientLastName.Equals(lastName, StringComparison.OrdinalIgnoreCase) Then
                matchFound = True
                clientRow = row
                Exit For
            End If
        Next

        If matchFound Then
            ' Load the existing client details into the form
            TextBox_ClientID.Text = clientsWorksheet.Cells(clientRow, 1).Text
            TextBox_ClientPhone1.Text = clientsWorksheet.Cells(clientRow, 4).Text
            TextBox_ClientCompany.Text = clientsWorksheet.Cells(clientRow, 5).Text
            TextBox_ClientEmail.Text = clientsWorksheet.Cells(clientRow, 6).Text
            TextBox_Street1.Text = clientsWorksheet.Cells(clientRow, 7).Text
            TextBox_Street2.Text = clientsWorksheet.Cells(clientRow, 8).Text
            TextBox_ClientCity.Text = clientsWorksheet.Cells(clientRow, 9).Text
            TextBox_ClientState.Text = clientsWorksheet.Cells(clientRow, 10).Text
            TextBox_ClientZip.Text = clientsWorksheet.Cells(clientRow, 11).Text
            TextBox_ClientComments.Text = clientsWorksheet.Cells(clientRow, 14).Text
            TextBox_DateAdded.Text = clientsWorksheet.Cells(clientRow, 13).Text
            TextBox_DateLastEdited.Text = DateTime.Now.ToString("MM/dd/yyyy")
            TextBox_ClientPhone2.Text = clientsWorksheet.Cells(clientRow, 16).Text

            MessageBox.Show("Client details loaded.")
        Else
            ' Generate a new Client ID
            Dim newClientID As Integer = GenerateClientID(clientsWorksheet)

            ' Find the next empty row
            Dim newRow As Integer = rowCount + 1

            ' Add the new client details to the worksheet
            clientsWorksheet.Cells(newRow, 1).Value = newClientID
            clientsWorksheet.Cells(newRow, 2).Value = firstName
            clientsWorksheet.Cells(newRow, 3).Value = lastName
            clientsWorksheet.Cells(newRow, 4).Value = TextBox_ClientPhone1.Text
            clientsWorksheet.Cells(newRow, 5).Value = TextBox_ClientCompany.Text
            clientsWorksheet.Cells(newRow, 6).Value = TextBox_ClientEmail.Text
            clientsWorksheet.Cells(newRow, 7).Value = TextBox_Street1.Text
            clientsWorksheet.Cells(newRow, 8).Value = TextBox_Street2.Text
            clientsWorksheet.Cells(newRow, 9).Value = TextBox_ClientCity.Text
            clientsWorksheet.Cells(newRow, 10).Value = TextBox_ClientState.Text
            clientsWorksheet.Cells(newRow, 11).Value = TextBox_ClientZip.Text
            clientsWorksheet.Cells(newRow, 12).Value = TextBox_ClientBalance.Text
            clientsWorksheet.Cells(newRow, 13).Value = DateTime.Now.ToString("MM/dd/yyyy")
            clientsWorksheet.Cells(newRow, 14).Value = TextBox_ClientComments.Text
            clientsWorksheet.Cells(newRow, 15).Value = DateTime.Now.ToString("MM/dd/yyyy")
            clientsWorksheet.Cells(newRow, 16).Value = TextBox_ClientPhone2.Text

            ' Save the Excel package
            excelPackage.Save()

            ' Update the form with the new Client ID
            TextBox_ClientID.Text = newClientID.ToString()
            TextBox_DateAdded.Text = DateTime.Now.ToString("MM/dd/yyyy")
            TextBox_DateLastEdited.Text = DateTime.Now.ToString("MM/dd/yyyy")

            MessageBox.Show("New client added.")
        End If

        ' Load matching invoices and calculate balance
        LoadClientInvoicesAndCalculateBalance()
    End Sub

    Private Function GenerateClientID(clientsWorksheet As ExcelWorksheet) As Integer
        ' Get the maximum existing client ID and increment it
        Dim maxClientID As Integer = 0
        Dim rowCount As Integer = clientsWorksheet.Dimension.End.Row

        For row As Integer = 3 To rowCount ' Assuming the data starts from row 3
            Dim currentClientID As String = clientsWorksheet.Cells(row, 1).Text
            Dim clientIDNumber As Integer
            If Integer.TryParse(currentClientID, clientIDNumber) Then
                If clientIDNumber > maxClientID Then
                    maxClientID = clientIDNumber
                End If
            End If
        Next

        ' Return the new client ID
        Return maxClientID + 1
    End Function

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click
        ' Clear previous search results
        ListView_ClientInvoices.Items.Clear()
        TextBox_ClientBalance.Text = ""

        ' Get the search text
        Dim searchText As String = TextBox_Search.Text.Trim()

        ' Ensure the search text is not empty
        If String.IsNullOrEmpty(searchText) Then
            MessageBox.Show("Please enter a search term.")
            Return
        End If

        ' Split the search text into first and last name
        Dim names() As String = searchText.Split(" "c)
        Dim firstName As String = String.Empty
        Dim lastName As String = String.Empty

        If names.Length > 0 Then
            firstName = names(0).Trim()
        End If
        If names.Length > 1 Then
            lastName = names(1).Trim()
        End If

        ' Access the "Clients" worksheet
        Dim clientsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Clients")
        If clientsWorksheet Is Nothing Then
            MessageBox.Show("Clients worksheet not found.")
            Return
        End If

        ' Initialize a flag to indicate if a match is found
        Dim matchFound As Boolean = False

        ' Find the row with the matching client ID, first name, last name, or company
        Dim rowCount As Integer = clientsWorksheet.Dimension.End.Row
        For row As Integer = 3 To rowCount ' Assuming the data starts from row 3
            Dim clientID As String = clientsWorksheet.Cells(row, 1).Text
            Dim clientFirstName As String = clientsWorksheet.Cells(row, 2).Text
            Dim clientLastName As String = clientsWorksheet.Cells(row, 3).Text
            Dim clientCompany As String = clientsWorksheet.Cells(row, 5).Text

            ' Check if client ID, first name and last name, or company match the search text
            If clientID.Contains(searchText) OrElse
           (Not String.IsNullOrEmpty(firstName) AndAlso Not String.IsNullOrEmpty(lastName) AndAlso
            clientFirstName.Equals(firstName, StringComparison.OrdinalIgnoreCase) AndAlso
            clientLastName.Equals(lastName, StringComparison.OrdinalIgnoreCase)) OrElse
           clientCompany.Contains(searchText) Then

                ' Match found, set the flag
                matchFound = True

                ' Populate the form fields with the details from the matched row
                TextBox_ClientID.Text = clientID
                TextBox_ClientName.Text = $"{clientFirstName} {clientLastName}"
                TextBox_ClientPhone1.Text = clientsWorksheet.Cells(row, 4).Text ' Phone1 in Column D
                TextBox_ClientCompany.Text = clientsWorksheet.Cells(row, 5).Text ' Client Company Column E
                TextBox_ClientEmail.Text = clientsWorksheet.Cells(row, 6).Text ' Client Email Column F
                TextBox_Street1.Text = clientsWorksheet.Cells(row, 7).Text ' Street1 in Column F
                TextBox_Street2.Text = clientsWorksheet.Cells(row, 8).Text ' Street2 in Column G
                TextBox_ClientCity.Text = clientsWorksheet.Cells(row, 9).Text ' City in Column H
                TextBox_ClientState.Text = clientsWorksheet.Cells(row, 10).Text ' State in column J
                TextBox_ClientZip.Text = clientsWorksheet.Cells(row, 11).Text ' Zip in Column K
                TextBox_ClientBalance.Text = clientsWorksheet.Cells(row, 12).Text ' Client Balance Column L
                TextBox_DateAdded.Text = clientsWorksheet.Cells(row, 13).Text ' Clients Date Added Column M
                TextBox_ClientComments.Text = clientsWorksheet.Cells(row, 14).Text ' Client Comments Column N
                TextBox_DateLastEdited.Text = clientsWorksheet.Cells(row, 15).Text 'Todays Date in Column O
                TextBox_TodaysDate.Text = DateTime.Now.ToString("MM/dd/yyyy")
                TextBox_ClientPhone2.Text = clientsWorksheet.Cells(row, 16).Text ' Phone2 in Column P

                ' Search for the invoices in the Jobs worksheet for the matching ClientID
                Dim jobsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Jobs")
                If jobsWorksheet Is Nothing Then
                    MessageBox.Show("Jobs worksheet not found.")
                    Return
                End If

                rowCount = jobsWorksheet.Dimension.End.Row
                Dim totalBalance As Double = 0.0

                For jobRow As Integer = 2 To rowCount ' Assuming data starts from row 2 in Jobs worksheet
                    If jobsWorksheet.Cells(jobRow, 1).Text = clientID Then
                        Dim invoiceNumber As String = jobsWorksheet.Cells(jobRow, 2).Text
                        Dim paidStatus As String = jobsWorksheet.Cells(jobRow, 19).Text
                        Dim listViewItem As New ListViewItem(invoiceNumber)
                        listViewItem.SubItems.Add(paidStatus)
                        ListView_ClientInvoices.Items.Add(listViewItem)

                        Dim estimate As Double
                        Dim downPayment As Double
                        Dim balance As Double

                        Double.TryParse(jobsWorksheet.Cells(jobRow, 16).Text.Replace("$", "").Replace(",", ""), estimate)
                        Double.TryParse(jobsWorksheet.Cells(jobRow, 17).Text.Replace("$", "").Replace(",", ""), downPayment)
                        Dim paid As String = jobsWorksheet.Cells(jobRow, 19).Text

                        If paid = "No" Then
                            balance = estimate - downPayment
                        Else
                            balance = 0.0
                        End If

                        totalBalance += balance
                    End If
                Next

                TextBox_ClientBalance.Text = totalBalance.ToString("C2", CultureInfo.CurrentCulture)

                ' Display a message box indicating the client was found
                MessageBox.Show("Client found and invoices loaded.")
                Exit For
            End If
        Next

        ' If no match was found, display a message box
        If Not matchFound Then
            MessageBox.Show("Search found nothing.")
        End If
    End Sub

    Private Sub ListView_ClientInvoices_Click(sender As Object, e As EventArgs) Handles ListView_ClientInvoices.DoubleClick

        ' Get the selected invoice number
        Dim selectedItem As ListViewItem = ListView_ClientInvoices.SelectedItems(0)
        Dim invoiceNumber As String = selectedItem.Text

        ' Access the "Jobs" worksheet
        Dim jobsWorksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets("Jobs")
        If jobsWorksheet Is Nothing Then
            MessageBox.Show("Jobs worksheet not found.")
            Return
        End If

        ' Find the row with the selected invoice number
        Dim rowCount As Integer = jobsWorksheet.Dimension.End.Row
        For row As Integer = 2 To rowCount
            If jobsWorksheet.Cells(row, 2).Text = invoiceNumber Then
                ' Populate the form fields with the details from the selected invoice row
                TextBox_EditClientName.Text = jobsWorksheet.Cells(row, 3).Text ' Client name in Column C
                TextBox_EditCompany.Text = jobsWorksheet.Cells(row, 4).Text ' Company in Column D
                TextBox_EditPhone1.Text = jobsWorksheet.Cells(row, 5).Text ' Phone1 in Column E
                TextBox_EditPhone2.Text = jobsWorksheet.Cells(row, 6).Text ' Phone2 in Column F
                TextBox_EditJobName.Text = jobsWorksheet.Cells(row, 7).Text ' JobName in Column G
                TextBox_EditStreet1.Text = jobsWorksheet.Cells(row, 8).Text ' Street1 in Column H
                TextBox_EditStreet2.Text = jobsWorksheet.Cells(row, 9).Text ' Street2 in Column I
                TextBox_EditCity.Text = jobsWorksheet.Cells(row, 10).Text ' City in Column J
                TextBox_EditState.Text = jobsWorksheet.Cells(row, 11).Text ' State in Column K
                TextBox_EditZip.Text = jobsWorksheet.Cells(row, 12).Text ' Zip in Column L
                TextBox_EditLaborHours.Text = jobsWorksheet.Cells(row, 13).Text ' Labor Hours in Column M
                TextBox_EditMaterialCost.Text = FormatCurrency(jobsWorksheet.Cells(row, 14).Text) ' Material Cost in Column N
                TextBox_EditRatePerHour.Text = FormatCurrency(jobsWorksheet.Cells(row, 15).Text) ' Rate Per Hour in Column O
                TextBox_EditEstimateTotal.Text = FormatCurrency(jobsWorksheet.Cells(row, 16).Text) ' Total Estimate in Column P
                TextBox_EditDownPayment.Text = FormatCurrency(jobsWorksheet.Cells(row, 17).Text) ' Down Payment in Column Q
                TextBox_EditStartDate.Text = jobsWorksheet.Cells(row, 18).Text ' Start Date in Column R
                chkBox_EditPaid.Checked = (jobsWorksheet.Cells(row, 19).Text = "Yes") ' Paid in Column S
                TextBox_EditJobDetails.Text = jobsWorksheet.Cells(row, 20).Text ' Job Details in Column T
                TextBox_EditJobNotes.Text = jobsWorksheet.Cells(row, 21).Text ' Job Notes in Column U
                TextBox_EditDateAdded.Text = jobsWorksheet.Cells(row, 22).Text ' Date Added in Column V
                TextBox_EditDate.Text = Today()
                TextBox_EditClientEmail.Text = jobsWorksheet.Cells(row, 24).Text ' Email in Column X

                ' Switch to the TabPage_UpdateJob tab
                TabControl1.SelectedTab = TabPage_UpdateJob

                Exit For
            End If
        Next
    End Sub
End Class
