<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHelp
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
        Me.lstHelp = New System.Windows.Forms.ListBox()
        Me.btnHelpClose = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lstHelp
        '
        Me.lstHelp.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstHelp.FormattingEnabled = True
        Me.lstHelp.ItemHeight = 15
        Me.lstHelp.Location = New System.Drawing.Point(8, 10)
        Me.lstHelp.Name = "lstHelp"
        Me.lstHelp.Size = New System.Drawing.Size(649, 274)
        Me.lstHelp.TabIndex = 0
        '
        'btnHelpClose
        '
        Me.btnHelpClose.Location = New System.Drawing.Point(598, 12)
        Me.btnHelpClose.Name = "btnHelpClose"
        Me.btnHelpClose.Size = New System.Drawing.Size(59, 21)
        Me.btnHelpClose.TabIndex = 1
        Me.btnHelpClose.Text = "Close"
        Me.btnHelpClose.UseVisualStyleBackColor = True
        '
        'frmHelp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(670, 295)
        Me.Controls.Add(Me.btnHelpClose)
        Me.Controls.Add(Me.lstHelp)
        Me.Name = "frmHelp"
        Me.Text = "Help"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lstHelp As ListBox
    Friend WithEvents btnHelpClose As Button
End Class
