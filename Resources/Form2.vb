Public Class DebugInfoForm

    Public Sub New(lastInvoiceNumber As String, newInvoiceNumber As String)
        InitializeComponent()

        ' Display the last and new invoice numbers
        Label_LastInvoiceNumber.Text = "Last Invoice Number: " & lastInvoiceNumber
        Label_NewInvoiceNumber.Text = "New Invoice Number: " & newInvoiceNumber
    End Sub

    Private Sub btn_Confirm_Click(sender As Object, e As EventArgs) Handles btn_Confirm.Click
        ' Close this form and continue with the next step
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        ' Close this form without proceeding
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

End Class
