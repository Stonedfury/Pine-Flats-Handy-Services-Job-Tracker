<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DebugInfoForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label_LastInvoiceNumber = New System.Windows.Forms.Label()
        Me.Label_NewInvoiceNumber = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btn_Confirm = New System.Windows.Forms.Button()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label_LastInvoiceNumber
        '
        Me.Label_LastInvoiceNumber.AutoSize = True
        Me.Label_LastInvoiceNumber.Location = New System.Drawing.Point(267, 91)
        Me.Label_LastInvoiceNumber.Name = "Label_LastInvoiceNumber"
        Me.Label_LastInvoiceNumber.Size = New System.Drawing.Size(48, 16)
        Me.Label_LastInvoiceNumber.TabIndex = 0
        Me.Label_LastInvoiceNumber.Text = "Label1"
        '
        'Label_NewInvoiceNumber
        '
        Me.Label_NewInvoiceNumber.AutoSize = True
        Me.Label_NewInvoiceNumber.Location = New System.Drawing.Point(270, 186)
        Me.Label_NewInvoiceNumber.Name = "Label_NewInvoiceNumber"
        Me.Label_NewInvoiceNumber.Size = New System.Drawing.Size(48, 16)
        Me.Label_NewInvoiceNumber.TabIndex = 1
        Me.Label_NewInvoiceNumber.Text = "Label2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 91)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(129, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Last Invoice Number"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 186)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(131, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "New Invoice Number"
        '
        'btn_Confirm
        '
        Me.btn_Confirm.Location = New System.Drawing.Point(218, 297)
        Me.btn_Confirm.Name = "btn_Confirm"
        Me.btn_Confirm.Size = New System.Drawing.Size(75, 23)
        Me.btn_Confirm.TabIndex = 2
        Me.btn_Confirm.Text = "Confirm"
        Me.btn_Confirm.UseVisualStyleBackColor = True
        '
        'btn_Cancel
        '
        Me.btn_Cancel.Location = New System.Drawing.Point(417, 297)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(75, 23)
        Me.btn_Cancel.TabIndex = 2
        Me.btn_Cancel.Text = "Cancel"
        Me.btn_Cancel.UseVisualStyleBackColor = True
        '
        'DebugInfoForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.btn_Confirm)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label_NewInvoiceNumber)
        Me.Controls.Add(Me.Label_LastInvoiceNumber)
        Me.Name = "DebugInfoForm"
        Me.Text = "Debug Form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label_LastInvoiceNumber As Label
    Friend WithEvents Label_NewInvoiceNumber As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents btn_Confirm As Button
    Friend WithEvents btn_Cancel As Button
End Class
