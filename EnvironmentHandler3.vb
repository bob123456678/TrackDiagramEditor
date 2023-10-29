Public Class EnvironmentHandler3
    Const Version = "2022-06-28" ' Added domaine function GetDom
    ' "2022-01-15" ' Added function Initialized and Sub SaveEnvAs
    Dim Env_Filename As String
    Dim ENV_INITIALIZED As Boolean
    Dim Read_Only As Boolean ' if config file is on a CD ROM
    Dim CHANGED As Boolean
    Structure ENV_VAR
        Public var_name As String
        Public var_value As String ' can be multiline, separated by CR-LF
    End Structure

    Const ENV_MAX = 1000
    Dim Env_Top As Integer
    Dim ENV(ENV_MAX) As ENV_VAR

    Public Function About() As String
        About = "Environment Handler 2: " + Version
    End Function
    Public Function Initialized() As Boolean
        Return ENV_INITIALIZED
    End Function

    Public Function InitEnv(fn As String) As Boolean
        Dim fn2 As String, fno As Integer, buffer As String

        fn2 = Trim(fn)
        If ENV_INITIALIZED Then
            If (UCase(Env_Filename) = UCase(fn2)) Then
                InitEnv = True
                Exit Function ' already ENV_INITIALIZED with this file name
            Else ' synchronize current env to file, close it and open a new one under the new name
                Call SaveEnv() ' write back the current environment file
                Call CloseEnv() ' close the current environment file
            End If
        End If

        Env_Filename = fn2
        Env_Top = 0
        ENV_INITIALIZED = True
        CHANGED = False
        Read_Only = False

        fno = FreeFile()
        If Dir(Env_Filename) <> "" Then ' file exists

            FileOpen(fno, Env_Filename, OpenMode.Input)

            Do While Not EOF(fno)
                buffer = LineInput(fno)
                buffer = Trim(buffer)
                If Left(buffer, 1) = "/" Then ' we have the name of an environment variable
                    If Env_Top = ENV_MAX Then
                        MsgBox("InitEnv: Capacity of EnvironmentHandler2 exceeded. Some environment variables may not have been loaded")
                        Exit Do
                    End If
                    Env_Top = Env_Top + 1
                    ENV(Env_Top).var_name = Mid(buffer, 2)
                    ENV(Env_Top).var_value = ""
                Else
                    If ENV(Env_Top).var_value = "" Then
                        ENV(Env_Top).var_value = buffer
                    Else
                        ENV(Env_Top).var_value = ENV(Env_Top).var_value + Chr(13) + Chr(10) + buffer
                    End If
                End If
            Loop
            FileClose(fno)
        Else
            ' File doesn't exist. Creating it is deferred until needed (in sub SaveEnv).
        End If
        InitEnv = True
    End Function

    Public Sub SaveEnv()
        ' Saves the environment values to file
        Dim fno As Integer, i As Integer, buffer As String
        If ENV_INITIALIZED Then
            If Read_Only Or Not CHANGED Then Exit Sub

            fno = FreeFile()
            On Error GoTo ROM ' if the file cannot be created or written to, set Read Only mode

            FileOpen(fno, Env_Filename, OpenMode.Output)
            For i = 1 To Env_Top
                buffer = "/" + ENV(i).var_name
                PrintLine(fno, buffer)
                buffer = ENV(i).var_value
                PrintLine(fno, buffer)
            Next i
            CHANGED = False
            On Error GoTo 0
            FileClose(fno)
        Else
            ' EnvironmentHnadler2.SaveEnv: Class not ENV_INITIALIZED
            MsgBox("Error in program logic: Attempt to save environment variables to an environment that hasn't been initialized")
        End If
        Exit Sub
ROM:
        On Error GoTo 0
        FileClose(fno)
        Read_Only = True
        MsgBox("Cannot save the environment variables. Read-only media?")
    End Sub
    Public Sub SaveEnvAs(cf_filename As String)
        ' Use this when wanting to save the environment variables under a new config file name
        If ENV_INITIALIZED Then
            Env_Filename = cf_filename
            Call SaveEnv()
        End If
    End Sub

    Public Sub CloseEnv()
        ' Does NOT write back the environment file, but closes the environment. Unsaved changes are lost
        If ENV_INITIALIZED Then
            Env_Filename = ""
            ENV_INITIALIZED = False
            Read_Only = False
            Env_Top = 0
        End If
    End Sub
    Public Function GetEnv(var_name As String) As String
        Dim i As Integer
        For i = 1 To Env_Top
            If ENV(i).var_name = var_name Then
                GetEnv = ENV(i).var_value
                Exit Function
            End If
        Next i
        GetEnv = ""
    End Function
    Public Sub SetEnv(var_name As String, value_string As String)
        ' Can handle multi-line contents of an environment variable.
        Dim i As Integer
        If Read_Only Or Not ENV_INITIALIZED Then Exit Sub

        i = Index(var_name)
        If i > 0 Then
            ENV(i).var_value = value_string
        Else ' var_name isn't in the current environment
            If Env_Top = ENV_MAX Then
                MsgBox("SetEnv: Capacity of EnvironmentHandler2 exceeded. Some environment variables will not be saved")
            Else
                Env_Top = Env_Top + 1
                ENV(Env_Top).var_name = var_name
                ENV(Env_Top).var_value = value_string
            End If
        End If
        CHANGED = True
    End Sub
    Public Sub DelEnv(var_name As String)
        Dim i As Integer
        If Read_Only Or Not ENV_INITIALIZED Then Exit Sub
        i = Index(var_name)
        If i > 0 Then
            ENV(i) = ENV(Env_Top)
            Env_Top = Env_Top - 1
            CHANGED = True
        End If
    End Sub
    Public Function GetDom() As String
        ' returns all environment variables in a TAB delimeted string
        Const TAB = Chr(9)
        Dim s As String = "", i As Integer
        For i = 1 To Env_Top
            s = s & TAB & ENV(i).var_name
        Next i
        If Env_Top >= 1 Then s = Mid(s, 2) ' remove leading TAB
        Return s
    End Function
    Private Function Index(var_name As String) As Integer
        Dim i As Integer
        For i = 1 To Env_Top
            If ENV(i).var_name = var_name Then
                Index = i
                Exit Function
            End If
        Next i
        Index = 0
    End Function

End Class
