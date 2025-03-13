Public Class frmHelp
    ' 2023-11-27
    ' Requires OO_LIB or OO_LIB_subset
    Private Help_File As String
    Private Top_Stored As Long = -1, Left_Stored As Long, Width_Stored As Long, Height_stored As Long
    Public Sub HelpDisplay(fn As String)
        ' Display the contents of file 'fn'
        ' If already loaded, just show (unhide) the form (defaulting to at the location used last)
        Me.Text = "Help"
        lstHelp.Sorted = False
        Me.Show()
        If fn <> Help_File Then
            HelpfileLoad(fn)
        End If
    End Sub
    Public Sub HelpDisplay(help_text() As String, n_lines As Integer)
        ' Display the text in help_text() beginning with element 1.
        Dim i As Integer
        Me.Text = "Help"
        lstHelp.Items.Clear()
        lstHelp.Sorted = False
        Me.Show()
        For i = 1 To n_lines
            If help_text(i) Is Nothing Then Exit For ' help_text contains fewer than n_lines lines
            lstHelp.Items.Add(help_text(i))
        Next i
    End Sub
    Public Sub HelpDisplaySortedListInitialize(window_title As String)
        ' Clear the lstHlp list; set the window text; show the window
        Me.Show()
        Me.Text = window_title
        lstHelp.Items.Clear()
        lstHelp.Sorted = True
    End Sub
    Public Sub HelpDisplaySortedListAdd(item As String)
        ' Clear the lstHlp list; set the window text; show the window
        lstHelp.Items.Add(item)
    End Sub

    Public Sub HelpWindowSetPos(my_left As Long, my_top As Long, my_width As Long, my_height As Long)
        Me.Show()
        HelpWindowPos(my_left, my_top, my_width, my_height)
        Me.Hide()
    End Sub
    Public Sub HelpWindowSetPosStored(my_left As Long, my_top As Long, my_width As Long, my_height As Long)
        ' Use the stored values, if any (i.e. if Left_Stored >=0). Otherwise use the parameters.
        If Left_Stored > 0 Then
            Me.Show()
            HelpWindowPos(Left_Stored, Top_Stored, Width_Stored, Height_stored)
            Me.Hide()
        Else
            HelpWindowSetPos(my_left, my_top, my_width, my_height)
        End If
    End Sub
    Public Sub HelpWindowSetPosStoredButWidth(my_left As Long, my_top As Long, my_width As Long, my_height As Long)
        ' Use the stored values except for width, if any (i.e. if Left_Stored >=0). Otherwise use the parameters.
        If Left_Stored > 0 Then
            Me.Show()
            HelpWindowPos(Left_Stored, Top_Stored, my_width, Height_stored)
            Me.Hide()
        Else
            HelpWindowSetPos(my_left, my_top, my_width, my_height)
        End If
    End Sub
    Public Sub HelpWindowSetSize(my_width As Long, my_height As Long)
        Me.Show()
        HelpWindowPos(Me.Left, Me.Top, my_width, my_height)
        Me.Hide()
    End Sub
    Public Sub HelpWindowClose()
        Me.Hide()
    End Sub
    Private Sub HelpWindowPos(my_left As Long, my_top As Long, my_width As Long, my_height As Long)
        Const w_min = 300
        Const h_min = 200
        If my_left >= 0 Then Me.Left = my_left
        If my_top >= 0 Then Me.Top = my_top
        If my_width >= w_min Then
            Me.Width = my_width
        Else
            Me.Width = w_min
        End If
        If my_height >= h_min Then
            Me.Height = my_height
        Else
            Me.Height = h_min
        End If
        lstHelp.Width = Me.Width - 4 * lstHelp.Left
        lstHelp.Height = Me.Height - 6 * lstHelp.Left
        btnHelpClose.Left = lstHelp.Left + lstHelp.Width - btnHelpClose.Width
    End Sub

    Private Sub HelpfileLoad(fn As String)
        Dim fno As Integer = FreeFile()
        lstHelp.Items.Clear()
        If FileExists(fn) Then
            Help_File = fn
            FileOpen(fno, fn, OpenMode.Input)
            While Not EOF(fno)
                lstHelp.Items.Add(LineInput(fno))
            End While
        Else
            lstHelp.Items.Add("Help file not found: " & fn)
        End If
    End Sub

    Private Sub btnHelpClose_Click(sender As Object, e As EventArgs) Handles btnHelpClose.Click
        Me.Hide()
    End Sub

    Private Sub frmHelp_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        HelpWindowPos(Me.Left, Me.Top, Me.Width, Me.Height)
        Left_Stored = Me.Left
        Top_Stored = Me.Top
        Width_Stored = Me.Width
        Height_stored = Me.Height
    End Sub
End Class