Public Class CustomInputForm
    Inherits Form
    ' 2025-03-12
    Private WithEvents OKButton2 As Button
    Private WithEvents CancelButton2 As Button
    Public Property InputTextBox As TextBox

    ' Constructor with a default value for the TextBox
    Public Sub New(prompt As String, default_text As String)
        ' Initialize controls
        OKButton2 = New Button() With {.Text = "OK", .DialogResult = DialogResult.OK, .Top = 50, .Left = 30}
        CancelButton2 = New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Top = 50, .Left = 120}
        InputTextBox = New TextBox() With {.Top = 10, .Left = 30, .Width = 300, .Text = default_text} ' Set default value

        ' Add controls to the form
        Me.Controls.Add(OKButton2)
        Me.Controls.Add(CancelButton2)
        Me.Controls.Add(InputTextBox)

        ' Set up the form properties
        Me.Text = prompt
        InputTextBox.Text = default_text
        Me.ClientSize = New Size(360, 100) ' was 260,120

        ' Center the form in the parent window
        Me.StartPosition = FormStartPosition.CenterParent

        ' Set Accept and Cancel buttons for the form
        Me.AcceptButton = OKButton2
        Me.CancelButton = CancelButton2
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton2.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton2.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class
