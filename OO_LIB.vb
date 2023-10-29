Module OO_LIB
    ' 2023-10-16
    Private Log_File As String

    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) ' Sleep is not a built-in subroutine in vb.net
    'For 64 Bit Systems only

    ' String manipilation
    Public Function DSToday() As String
        ' Returns today as yyyy-mm-dd
        DSToday = StrFixR(DateAndTime.Year(Today), " ", 4) + "-" + StrFixR(DateAndTime.Month(Today), "0", 2) + "-" + StrFixR(DateAndTime.Day(Today), "0", 2)
    End Function

    Public Function ExtractField(s As String, N As Integer, delim As String) As String
        ' Returns field number 'N' (counting from 1) in a 'delim' separated string
        ' Returns "" if no such field.
        Dim pos_from, pos_to As Integer, count As Integer, res As String
        pos_from = 1 : count = 0
        Do
            While delim = " " And Mid(s, pos_from, 1) = " "
                pos_from = pos_from + 1
            End While
            pos_to = InStr(pos_from, s, delim, 1)
            If pos_to = 0 Then pos_to = Len(s) + 1
            'MsgBox( Str(pos_from) + " " + Str(pos_to))
            If pos_to >= pos_from Then
                count = count + 1
                If count = N Then Exit Do
                pos_from = pos_to + 1
            End If
        Loop Until pos_from > Len(s)
        'MsgBox ("Count = " + Str(N) + " result is...")
        If count = N Then
            res = Mid(s, pos_from, pos_to - pos_from)
            If res = delim Then res = "" ' 2010-06-20, fixing some bug
            ExtractField = res
        Else
            ExtractField = ""
        End If
    End Function
    Public Function ExtractField2(s As String, N As Integer, delim As String) As String
        ' Returns field number 'N' (counting from 1) in a 'delim' separated string.
        ' Returns the empty string if no such field.
        ' Fields can be enclosed between quote characters "...".
        ' Such surrounding quote characters are not included in the returned string.
        Dim cp As Integer, c As String, s2 As String
        Dim qc As String
        Dim count As Integer, l As Integer, in_quote As Boolean
        l = Len(s) : cp = 1 : qc = """" : count = 1 : s2 = ""
        Do
            c = Mid(s, cp, 1)
            Select Case c
                Case delim
                    If Not in_quote Then
                        count = count + 1
                    Else
                        If count = N Then s2 = s2 + c
                    End If
                Case qc
                    in_quote = Not in_quote
                    If count = N Then s2 = s2 + c
                Case Else
                    If count = N Then s2 = s2 + c
            End Select
            cp = cp + 1
        Loop Until cp > l Or count > N

        If Len(s2) >= 2 Then
            If Left(s2, 1) = qc And Right(s2, 1) = qc Then
                ' remove surrounding quote characters
                s2 = Mid(s2, 2, Len(s2) - 2)
            End If
        End If
        ExtractField2 = s2
    End Function

    Public Function ExtractFieldIndex(s As String, field_name As String, delim As String) As Integer
        ' Counting from 1, returns the index to the field, i.e. how far down the string 's' the field is. Case sensitive.
        ' Returns 0 if the field is not found in 's'
        Dim i As Integer = 0, f As String
        Do
            i = i + 1
            f = ExtractField2(s, i, delim)
            If f = field_name Then Return i
            If f = "" Then Return 0
        Loop
    End Function
    Public Function ExtractLastField(s As String, delim As String) As String
        ' Returns last field in a 'delim' separated string.
        ' If no DELIM character in s then s is returned
        Dim pos As Integer, pos_to As Integer, c As String
        pos_to = Len(s)
        If pos_to > 0 Then
            pos = pos_to
            Do
                c = Mid(s, pos, 1)
                If c = delim Then
                    pos = pos + 1
                    Return Mid(s, pos, pos_to - pos + 1)
                End If
                pos = pos - 1
            Loop Until pos = 0
        End If
        Return s
    End Function
    Public Function ExtractTail(s As String, N As Integer, delim As String) As String
        ' Returns field number 'N' (counting from 1) and all subsequent fields
        ' in a 'delim' separated string. Returns "" of no such field.
        Dim pos_from, pos_to As Integer, count As Integer
        pos_from = 1 : count = 0
        Do
            pos_to = InStr(pos_from, s, delim, 1)
            If pos_to = 0 Then pos_to = Len(s) + 1
            'MsgBox Str(pos_from) + " " + Str(pos_to)
            If pos_to >= pos_from Then
                count = count + 1
                If count = N Then Exit Do
                pos_from = pos_to + 1
            End If
        Loop Until pos_from > Len(s)
        'MsgBox ("Count = " + Str(N) + " result is...")
        If count = N Then
            ExtractTail = Mid(s, pos_from, Len(s) - pos_from + 1)
        Else
            ExtractTail = ""
        End If
    End Function
    Public Function ExtractHeadTail(ByRef head As String, ByRef tail As String,
                                ByVal s As String, delim_set As String) As String
        ' Extracts head and tail, splitting at the first occurance
        ' of a character from delim_set.
        ' Returns the delim character as function value.
        ' Head and tail can be enclosed between quote characters "...".
        ' Quote characters are not stripped from the returned strings.
        ' If ' ' is a delimiter, then s will be frontstripped before searching
        ' for head.
        Dim cp As Integer, c As String, delim As String = ""
        Dim qc As String
        Dim l As Integer, in_quote As Boolean
        l = Len(s) : cp = 1 : qc = """"
        If InStr(1, delim_set, " ", 1) > 0 Then s = Trim(s)
        head = s : tail = "" ' default if no split point found
        Do
            c = Mid(s, cp, 1)
            If InStr(1, delim_set, c, 1) > 0 Then
                If Not in_quote Then 'found split point
                    head = Left(s, cp - 1)
                    tail = Mid(s, cp + 1)
                    delim = c
                    Exit Do
                End If
            ElseIf c = qc Then
                in_quote = Not in_quote
            Else
            End If
            cp = cp + 1
        Loop Until cp > l
        Return delim
    End Function

    Public Function FieldCount(s As String, delim As String) As Integer
        ' Given delim = ",":
        ' "a,b,c"  returns 3
        ' "a,b,c," returns 3
        ' ",a,b,c" returns 4 (leading empty field)
        Dim n As Integer = 0, c As String, nc As Integer, i As Integer
        nc = Strings.Len(s)

        If nc = 0 Then Return 0

        For i = 1 To nc
            c = Mid(s, i, 1)
            If c = delim Then n = n + 1
        Next i
        If Strings.Right(s, 1) <> delim Then n = n + 1
        Return n
    End Function
    Public Function IsDigit(c As String) As Boolean
        ' True, if the single character 'c' is a digit
        IsDigit = ("0" <= c And c <= "9")
    End Function
    Public Function IsInList(list As Object, s As String) As Boolean
        ' Performs a case insensitive check if 's' is in the list.
        ' ListBox.Items.Contains is, unfortunately, case sensitive and og little use
        Dim li As Integer, ul As Integer, s2 As String
        ul = list.items.count - 1
        If ul >= 0 Then
            s2 = LCase(s)
            For li = 0 To ul
                If LCase(list.Items.item(li)) = s2 Then Return True
            Next li
        End If
        Return False
    End Function
    Public Function IsInteger(s As String) As Boolean
        Dim i As Integer, l As Integer, c As String, ok As Boolean
        l = Len(s)
        c = Left(s, 1)
        ok = (c = "-") Or IsDigit(c)
        For i = 2 To l
            ok = ok And IsDigit(Mid(s, i, 1))
        Next i
        IsInteger = ok
    End Function
    Public Function IsLetter(c As String) As Boolean
        IsLetter = ("a" <= c And c <= "z") Or ("A" <= c And c <= "Z") _
              Or c = "æ" Or c = "Æ" Or c = "ø" Or c = "Ø" Or c = "å" Or c = "Å" _
              Or c = "ü" Or c = "Ü" Or c = "ä" Or c = "Ä" Or c = "ö" Or c = "Ö"

    End Function

    Public Function StrBar(c As String, n As Integer) As String
        ' returns a string of 'n' 'c' characters
        Dim i As Integer, s As String = ""
        If n = 0 Then Return ""
        s = c
        For i = 2 To n
            s = s & c
        Next i
        Return s
    End Function
    Public Function StrFixL(s As String, c As String, l As Integer) As String
        ' Formerly StrAdjust
        ' Returns a string of lenght 'L', with Trim(s) left adjusted.
        ' If Trim(s) is shorter than L, appends c characters
        Dim r As String, ls2 As Integer, i As Integer, s2 As String
        If l < 1 Then
            StrFixL = ""
        Else
            s2 = Trim(s)
            ls2 = Len(s2)
            If ls2 > l Then ' truncate
                StrFixL = Left(s, l)
            Else
                r = ""
                For i = 1 To l
                    If i > ls2 Then
                        r = r + c
                    Else
                        r = r + Mid(s2, i, 1)
                    End If
                Next i
                StrFixL = r
            End If
        End If
    End Function

    Public Function StrFixLPlus(s As String, c As String, L As Integer) As String
        ' Returns a string of lenght 'L', with Trim(s) left adjusted, unless length of Trim('s') is > 'L' in which case 's' is returned
        ' If Trim(s) is shorter than L, appends c characters.
        Dim s2 As String
        s2 = Trim(s)
        If Len(s2) > L Then
            Return s2
        Else
            Return StrFixL(s2, c, L)
        End If
    End Function

    Public Function StrFixR(s As String, c As String, l As Integer) As String
        ' Formerly StrFix
        ' returns a string of lenght 'l', with Trim(s) right adjusted.
        ' If Trim(s) is shorter than l, prepends c characters
        Dim r As String, ls2 As Integer, i As Integer, N As Integer, s2 As String
        s2 = Trim(s)
        ls2 = Len(s2)
        If ls2 > l Then ' truncate
            StrFixR = Left(s, l)
        Else
            N = l - ls2
            r = ""
            For i = 1 To N
                r = r + c
            Next i
            StrFixR = r + s2
        End If
    End Function
    Public Function StrFixLong(ByVal v As Long, l As Integer) As String
        ' Returns a string of length 'l' with the long 'v' rightadjusted
        StrFixLong = StrFixR(Str(v), " ", l)
    End Function
    Public Function StrFixLong2(ByVal v As Long, l As Integer) As String
        ' Returns a string of length 'l' with the long 'v' rightadjusted and leading zeros
        StrFixLong2 = StrFixR(Format(v), "0", l)
    End Function

    Public Function StrLine(c As String, c2 As String, l As Integer) As String
        ' Returns a string of lenght 'L', with 'c' characters.
        ' In every five  positions the 'c2' character is inserted instead of 'c'
        Dim r As String, i As Integer
        r = ""
        For i = 1 To l
            If i Mod 5 = 0 Then
                r = r + c2
            Else
                r = r + c
            End If
        Next i
        Return r
    End Function
    Public Function StrRemoveExcessSpaces(s As String) As String
        ' "abc   def  gh" becomes abc def gh" etc.
        Dim l As Integer, i As Integer, c1 As String, c2 As String, s2 As String
        l = Len(s)
        If l <= 1 Then
            StrRemoveExcessSpaces = s
        Else
            c1 = Left(s, 1)
            s2 = c1
            For i = 2 To l
                c2 = Mid(s, i, 1)
                If Not (c2 = " " And c1 = " ") Then
                    s2 = s2 & c2
                End If
                c1 = c2
            Next i
            StrRemoveExcessSpaces = s2
        End If
    End Function
    Public Function StrRemoveSpaces(s) As String
        ' Removes all spaces
        Dim l As Integer, i As Integer, c As String, s2 As String = ""
        l = Len(s)
        For i = 1 To l
            c = Mid(s, i, 1)
            If c <> " " Then s2 = s2 & c
        Next i
        StrRemoveSpaces = s2
    End Function
    Public Function StrReplace(s As String, pattern As String, s_new As String) As String
        ' Substitutes pattern with s_new. -- eventually replace the body with the built-in function
        Dim l_pattern As Integer, pos As Integer, found As Integer, s2 As String = ""
        ' MsgBox("StrReplace>" + s + "< >" + pattern + "< >" + s_new + "<") 
        l_pattern = Strings.Len(pattern)
        If l_pattern = 0 Then
            StrReplace = s
            Exit Function
        End If

        found = InStr(s, pattern, CompareMethod.Text)
        pos = 1
        While found > 0
            s2 = s2 + Mid(s, pos, found - pos)
            s2 = s2 + s_new
            ' MsgBox("pos, found=" + Str(pos) + Str(found) + "/" + Mid(s, pos, found - pos) + "/" + s2)
            pos = found + l_pattern
            found = InStr(found + l_pattern, s, pattern)
        End While
        s2 = s2 + Mid(s, pos, Len(s) - pos + 1)
        StrReplace = s2
    End Function

    Public Function StrSubstituteCharSet(ByVal s As String, c_set As String, s_new As String) As String
        ' Substitutes all individual characters in c_set with s_new. s_new can have length >= 0.
        ' Case must match.
        Dim s_i As Long, s_c As String, s_l As Long
        Dim c_i As Integer, c_c As String, c_set_l As Integer
        Dim sres As String = ""

        s_l = Len(s)
        c_set_l = Len(c_set)
        For s_i = 1 To s_l
            s_c = Mid(s, s_i, 1)
            For c_i = 1 To c_set_l
                c_c = Mid(c_set, c_i, 1)
                If s_c = c_c Then
                    s_c = s_new
                    Exit For ' c_i = c_set_l ' Exit For
                End If
            Next c_i
            sres = sres + s_c
        Next s_i
        StrSubstituteCharSet = sres
    End Function

    Public Sub StringArraySort1(A() As String, ul As Integer)
        ' Insertion sort, case insensitive
        ' First element is at index 1, last element at index ul
        Dim i As Integer, j As Integer, s As String

        For i = 2 To ul
            s = A(i)
            j = i
            While j > 1 And UCase(A(j - 1)) > UCase(s)
                A(j) = A(j - 1)
                j = j - 1
            End While
            A(j) = s
        Next i
    End Sub


    ' File related

    Public Function DirectoryCreate(orig_path As String) As Boolean
        ' Creates the directory if it doesn't exist.
        ' Returns true if success, or if the directory already exists.
        ' 'orig_path' must include drive:  e:\duc1\duc2\
        ' The trailing "\" is optional
        Dim path As String, path2 As String, sub_path As String, next_path As String, N As Integer
        If Right(orig_path, 1) = "\" Then
            path = orig_path
        Else
            path = orig_path + "\"
        End If
        path2 = Left(path, Strings.Len(path) - 1)
        If Dir(path, vbDirectory) <> "" OrElse Dir(path2, vbDirectory) <> "" Then ' path already exists
            Return True
        End If
        If Mid(path, 2, 1) <> ":" Then
            MsgBox("Cannot create this directory, drive specification is missing: " + path)
            Return False
        End If
        If InStr(1, path, "/", 1) > 0 Then
            MsgBox("Cannot create this directory, name contains an illegal character ""/"": " + path)
            Return False
        End If
        N = 1
        sub_path = ExtractField(path, N, "\") + "\"
        Do
            If Dir(sub_path, vbDirectory) = "" Then
                MkDir(sub_path)
            End If
            N = N + 1
            next_path = ExtractField(path, N, "\")
            If next_path = "" Then Exit Do
            sub_path = sub_path + next_path + "\"
        Loop
        Return True
    End Function
    Public Function FilenameDriveAdd(ByVal fname As String, drive As String) As String
        ' Returns filename with the new drive "<drive>:" prepended unless the filename already had a drive.
        ' No "\" will be prepended.
        If Mid(fname, 2, 1) = ":" Then
            Return fname
        Else
            Return drive & ":" & fname
        End If
    End Function
    Public Function FilenameDriveChange(ByVal fname As String, drive As String) As String
        ' Returns filename with the new drive.
        ' If fname had no drive, "<drive>:" will be prepended. No "\" will be prepended.
        If Mid(fname, 2, 1) = ":" Then
            Return drive + Mid(fname, 2)
        Else
            Return drive + ":" + fname
        End If
    End Function

    Public Function FilenameChangeExtension(ByVal name As String, ext As String,
                                        must_change As Boolean) As String
        ' Returns filename with the new extension. If 'must_change' is true and the old and new
        ' extension are equal, returns "error.tmp" instead.
        Dim new_ext As String, new_name As String
        If Left(ext, 1) = "." Then new_ext = Mid(ext, 2) Else new_ext = ext
        new_name = FilenameNoExtension(name) + "." + new_ext
        If new_name = name And must_change Then
            MsgBox("Call of FileNameChangeExtension did not change extension. Returns error.tmp")
            FilenameChangeExtension = "error.tmp"
        Else
            FilenameChangeExtension = new_name
        End If
    End Function

    Public Function FileExists(name As String) As Boolean
        ' Returns true if file 'name' exists
        On Error GoTo some_error
        If Trim(name) <> "" Then
            FileExists = (Dir(name) <> "")
        Else
            FileExists = False
        End If
        On Error GoTo 0
        Exit Function
some_error:
        On Error GoTo 0
        MsgBox("Error in FileExists(" + name + ")")
    End Function
    Public Function FileExistsAs(filename As String, drivelist As String) As String
        ' Returns the drive letter, full path and filename where the file is first found.
        ' Example of returned string: "P:\foto\080706_77k.jpg"
        ' Returns "" if the file cannot be found on any of the drives.
        ' Example 'drive_list':  "CP" (spaces are not allowed).
        ' 'filename' must have the full path, except for the drive as in "\foto\2008\IMGP1234.jpg"
        ' 'filename' may have a drive designation. That drive is searched first, then the drive list
        Dim ul As Integer, pos As Integer, drive_letter As String, drives As String
        Dim fn As String, fn2 As String

        If filename = "" Then Return ""

        drives = drivelist
        If Mid(filename, 2, 1) = ":" Then ' drive letter included in filename
            fn = Mid(filename, 3)
            drive_letter = UCase(Left(filename, 1))
            If InStr(drivelist, drive_letter) = 0 Then
                drives = drive_letter & drivelist
            End If
        Else
            fn = filename
        End If
        ul = Len(drives)
        Do
            pos = pos + 1
            If pos > ul Then Exit Do
            drive_letter = Mid(drives, pos, 1)
            fn2 = drive_letter & ":" & fn
            If FileExists(fn2) Then
                Return fn2
            End If
        Loop
        Return ""
    End Function

    Public Function FilenameExtension(fn As String) As String
        ' Returns the file extension converted to lower case, "." not included.
        ' Returns "", if the file has no extension.
        ' The file name may include "." characters besides the one separating the extension.
        Dim decimal_point_pos As Long
        If fn = "" Then
            FilenameExtension = ""
        Else
            decimal_point_pos = InStrRev(fn, ".", Len(fn), CompareMethod.Text)
            If decimal_point_pos > 1 Then
                FilenameExtension = LCase(Mid(fn, decimal_point_pos + 1))
            Else
                FilenameExtension = ""
            End If
        End If
    End Function
    Public Function FilenameExtract(fn As String) As String
        ' Returns the file name excluding the path part
        ' c:\oo\f117.txt would return f117.txt
        Dim l As Integer, p, pos_backslash As Integer

        l = Len(fn)
        For p = l To 1 Step -1
            If Mid(fn, p, 1) = "\" Then
                If pos_backslash = 0 Then pos_backslash = p
                Exit For
            End If
        Next p
        FilenameExtract = Mid(fn, pos_backslash + 1)
    End Function
    Public Function FilenameNoDrive(fn As String)
        ' Returns the file name wothout drive designation, if any.
        ' c:\oo\f117.txt would return "\oo\f117.txt"
        Dim p, pos_colon As Integer

        pos_colon = 0
        For p = 1 To Len(fn)
            If Mid(fn, p, 1) = ":" Then
                pos_colon = p
                Exit For
            End If
        Next p
        FilenameNoDrive = Mid(fn, pos_colon + 1)
    End Function
    Public Function FilenameNoExtension(fn As String)
        ' Returns the file name without extension
        ' c:\oo\f117.txt would return "c:\oo\f117"
        ' c:\oo\f.117.txt would return "c:\oo\f.117"
        ' c:\oo\BR18.4 would return "c:\oo\BR18"
        ' c:\oo\BR18.4_030606 would return "c:\oo\BR18.4_030606"
        ' The filename may contain additional "." characters
        ' The extention must be 4 or less characters long to be considered an extension
        Dim ext As String, ext_l As Integer
        ext = FilenameExtension(fn)
        ext_l = Len(ext)
        If ext_l = 0 Then
            FilenameNoExtension = fn
        Else
            FilenameNoExtension = Left(fn, Len(fn) - ext_l - 1)
        End If
    End Function

    Public Function FilenamePath(fn As String) As String
        ' Returns the path part of a file name - includes the trailing "\"
        ' fn is assumed to be a file name, not just a path
        ' c:\oo\f117.txt would return c:\oo\
        ' C:f117.txt would return C:
        Dim pos_backslash As Integer, pos_colon As Integer
        If fn = "" Then Return ""
        pos_backslash = InStrRev(fn, "\", Len(fn), CompareMethod.Text)
        If pos_backslash > 0 Then
            FilenamePath = Left(fn, pos_backslash)
        Else
            pos_colon = InStrRev(fn, ":", Len(fn), CompareMethod.Text)
            If pos_colon > 0 Then
                FilenamePath = Left(fn, pos_colon)
            Else
                FilenamePath = ""
            End If
        End If
    End Function
    Public Function FilenameProper(fn As String) As String
        ' Returns the proper part of a file name
        ' (no path, no extension).
        ' c:\oo\f117.txt would return f117

        Dim l As Integer, p, pos_dot, pos_backslash As Integer

        l = Len(fn)
        For p = l To 1 Step -1
            Select Case Mid(fn, p, 1)
                Case "."
                    If pos_dot = 0 Then pos_dot = p
                Case "\"
                    If pos_backslash = 0 Then pos_backslash = p
                    Exit For
                Case Else
            End Select
        Next p
        If pos_dot = 0 Then pos_dot = l + 1
        FilenameProper = Mid(fn, pos_backslash + 1, pos_dot - pos_backslash - 1)
    End Function
    Public Function FilePathUp(path As String) As String
        ' NB: Different semantics than the VB4 version
        ' c:\aa\bb\  -> c:\aa\
        ' c:\aa\     -> c:\
        ' c:\        -> ""
        ' \a\b\      -> \a\
        ' \          -> ""
        Dim p As Integer, l As Integer, p_left As Integer = 1
        l = Len(path)
        If Mid(path, 2, 1) = ":" Then p_left = 3
        If l <= p_left Then
            Return "" ' cannot go further up
        Else
            For p = l - 1 To p_left Step -1
                If Mid(path, p, 1) = "\" Then
                    Return Left(path, p)
                End If
            Next p
            FilePathUp = ""
        End If
    End Function

    Public Function FileSearchFor(sdir As String, fname As String, ByRef err_msg As String) As String
        ' Starting in 'sdir' searches for the file here and in all sub and sub-sub etc. directories.
        ' fname may contain a path. This will be stripped before the search.
        ' If the file is found, the full pathname including drive letter is returned.
        ' sdir can optionally end with a "\".
        ' If 'sdir' has no drive letter, the letter of the current drive will be used
        ' Error messages are inserted into 'err_msg'.
        Dim fn As String ' fname with path, if any, stripped
        Dim try_name As String, subdir As String, answer As String
        Dim subdir_list(100) As String ' We can only handle 100 subdirectories - should be enough. If not, change to using REDIM
        Dim ul As Integer, i As Integer

        fn = FilenameExtract(fname)
        If Right(sdir, 1) <> "\" Then sdir = sdir & "\"
        If Mid(sdir, 2, 1) <> ":" Then ' add current drive letter
            If Left(sdir, 1) = "\" Then
                sdir = Left(CurDir(""), 2) & sdir
            Else
                sdir = Left(CurDir(""), 2) & "\" & sdir
            End If
        End If

        try_name = sdir & fn
        If FileExists(try_name) Then
            Return try_name
        End If
        ' File is not in sdir. look the subdirectories
        On Error GoTo skip_if ' this happens if directory contains Polish characters
        subdir = Dir(sdir + "*", vbDirectory)
        While subdir <> ""
            If subdir <> "." And subdir <> ".." Then
                ' GetAttr raises an exception if directory not found. Dir converts Polish characters to latin, this resulting in an exception here
                If (GetAttr(sdir & subdir) And vbDirectory) = vbDirectory Then
                    ' Dir cannot be called recursively so we have to first store the names of the subdirectories, then process them
                    If ul < 99 Then
                        ul = ul + 1
                        subdir_list(ul) = sdir & subdir
                    Else
                        err_msg = err_msg & "Too many subdirectories, not all were searched. "
                    End If
                End If
            End If
            subdir = Dir()
        End While
        On Error GoTo 0

        i = 1
        Do
            If i > ul Then Exit Do
            answer = FileSearchFor(subdir_list(i), fn, err_msg)
            If answer <> "" Then
                Return answer
            End If
            i = i + 1
        Loop
        Return ""
skip_if:
        On Error GoTo 0
        err_msg = err_msg & "Cannot proccess this directory due to foreign characters in sbudirectories: " & sdir & ". "
        Return ""
    End Function

    Public Function ValidateFileName(fn As String) As String
        ' Returns "" if fn is a valid filename. Otherwise returns the offending character
        Const bad_chars = ":\/.*?""´`|¨^~@"
        Dim ul As Integer, pos As Integer
        ul = Strings.Len(bad_chars)
        For pos = 1 To ul
            If InStr(fn, Mid(bad_chars, pos, 1)) > 0 Then
                Return Mid(bad_chars, pos, 1)
            End If
        Next
        Return ""
    End Function


    ' Log file
    ' The log file is closed after each write
    Public Sub LogFileInit(fn As String)
        Dim fno As Integer
        fno = FreeFile()
        FileOpen(fno, fn, OpenMode.Output)
        PrintLine(fno, DSToday)
        FileClose(fno)
        Log_File = fn
    End Sub

    Public Sub LogFileWrite(msg)
        Dim fno As Integer
        If FileExists(Log_File) Then
            fno = FreeFile()
            FileOpen(fno, Log_File, OpenMode.Append)
            PrintLine(fno, msg)
            FileClose(fno)
        Else
            MsgBox("LogFileWrite was called before initialization of the log file")
        End If
    End Sub

    ' Image related
    Public Function IsDisplayableImage(fn As String) As Boolean
        ' Vb .NET can also display .png's
        Dim ext As String, sl As Integer
        sl = fn.Length
        If sl < 5 Then Return False
        ext = LCase(Right(fn, 4))
        Select Case ext
            Case ".jpg", ".bmp", ".gif", ".png", ".tif", ".png"
                Return True
            Case "jpeg"
                If sl > 5 Then
                    Return LCase(Right(fn, 5)) = ".jpeg"
                Else
                    Return False
                End If
            Case Else
                Return False
        End Select
    End Function

    Public Function IsImageFile(fn As String) As Boolean
        Dim ext As String, sl As Integer
        sl = fn.Length
        If sl < 5 Then Return False
        ext = LCase(Right(fn, 4))
        Select Case ext
            Case ".jpg", ".bmp", ".gif", ".png", ".tif", ".dng", ".pef", ".psd", ".rwl"
                Return True
            Case "jpeg"
                If sl > 5 Then
                    Return LCase(Right(fn, 5)) = ".jpeg"
                Else
                    Return False
                End If
            Case Else
                Return False
        End Select
    End Function

    Public Function IsMusicFile(fn As String) As Boolean
        Dim ext As String, sl As Integer
        sl = fn.Length
        If sl < 5 Then Return False

        ext = LCase(Right(fn, 4))
        Select Case ext
            Case ".wav", ".mp3"
                Return True
            Case "flac"
                If sl > 5 Then
                    Return LCase(Right(fn, 5)) = ".flac"
                Else
                    Return False
                End If
            Case Else
                Return False
        End Select
    End Function
    Public Function PhotoDate(file_name As String) As String
        ' Returns the date of the photo as "yymmdd".
        ' Requires that the file_name follows Ole's standard:
        ' [<letter><letterdigit>*_]yymmdd[<any>*].<ext>
        ' Returns "" if the file_name doesn't conform.
        ' yyyymmdd is not accepted
        Dim fn As String, pd As String
        fn = PhotoNameCore(file_name)
        pd = Left(fn, 6)
        If Len(pd) = 6 And IsInteger(pd) Then
            PhotoDate = pd
        Else
            PhotoDate = ""
        End If
    End Function
    Public Function PhotoFilmName(file_name As String)
        ' Returns the negative name of the photo as "yymmL". Returns "" if file isn't from a negative.
        ' Requires that the file_name follows Ole's standard:
        ' [<letter><letterdigit>*_]yymm<letter>_nn[_<any>*].<ext>
        ' Returns "" if the file_name doesn't conform. This includes patterns of old negative names like Na38_F2.jpg,
        ' where we have no film name indication available.
        Dim fn As String, pd As String
        fn = PhotoNameCore(file_name)
        pd = Left(fn, 6)
        If Len(pd) = 6 And IsInteger(Left(pd, 4)) And IsLetter(Mid(pd, 5, 1)) And Mid(pd, 6, 1) = "_" Then
            PhotoFilmName = Left(pd, 5)
        Else
            PhotoFilmName = ""
        End If
    End Function

    Public Function PhotoLocateSimilar(pname As String) As String
        ' If pname is displayable (has a jpg, gif, png, or bmp extension) AND exists, then pname is returned.
        ' Otherwise: Locates a DISPLAYABLE photo with same core file name as pname.
        ' Only a .jpg file will be searched for and returned.
        ' Searches the current directory and one directory up.
        ' Returns "" if no such file.
        Dim pn_found As String, pn_found_no_path As String, pn_jpg_wildcard As String
        Dim pname_ext As String, pname_path As String, pname_path_up As String, pname_core As String

        If IsDisplayableImage(pname) And FileExists(pname) Then
            Return pname
        Else
            pname_ext = FilenameExtension(pname)
            pname_core = PhotoNameCore(pname)
            pname_path = FilenamePath(pname)

            pn_jpg_wildcard = pname_path & "*" & pname_core & "*.jpg"
            pn_found_no_path = Dir(pn_jpg_wildcard, vbNormal)
            If pn_found_no_path <> "" Then
                pn_found = pname_path & pn_found_no_path
                Return pn_found
            End If
            ' Look one level up
            pname_path_up = FilePathUp(pname_path)
            If pname_path_up <> "" Then
                pn_jpg_wildcard = pname_path_up & "*" & pname_core & "*.jpg"
                pn_found_no_path = Dir(pn_jpg_wildcard, vbNormal)
                If pn_found_no_path <> "" Then
                    pn_found = pname_path_up & pn_found_no_path
                    Return pn_found
                End If
            End If
            Return ""
        End If
    End Function

    Public Function PhotoNameCore(photofile_name As String) As String
        ' returns the core photo file name, e.g.:
        ' prefix_040708_33p_suffix.jpg -> 040708_33p
        Dim path As String = "", prefix As String = "", proper As String = "", seq As String = "", tail As String = "", ext As String = ""
        ' Debug.Print "-- BEGIN PhotoNameCore1: Called with= " + photofile_name
        Call PhotoNameSplit(photofile_name, path, prefix, proper, seq, tail, ext)
        ' Debug.Print "-- END PhotoNameCore1: proper | seq=" + proper + "|" + seq
        PhotoNameCore = proper + seq
    End Function
    Private Function PhotoNameKind(pn_any_case As String) As Integer
        Dim m_val As Integer, pn As String
        ' 'pn' must be the proper part of a photo name, void of path, prefix, tail, and extension
        ' Returns a value > 0 if pn is the beginning of a photo file name:
        ' 10: yymmdd...   e.g. 050312_19p, but not 197206 (digits 3^4 must be in range 01 .. 12)
        ' 11: IMGPnnnn, I2GPnnnn
        ' 12: IMG
        ' 13: CIMGnnnn
        ' 14: DSCnnnnn
        ' 15: _IMGnnnn, _I2Gnnnn
        ' 20: yymml[l]    e.g. 0504A_03   0504A_04k   0504B_04Ak   0504B_04Ak
        ' 21: snn_        e.g s40_03  (scanned slides)
        ' 22: ann_        e.g a40_03  (scanned slides from archive)
        ' 31: Nlnn_Ln     e.g. Na38_F2.jpg -- scanned b/w negative, album "a", page 28, row F, frame 2
        ' Extend later with other valid forms
        '        Call dpBegin("PhotoNameKind(" + pn_any_case + ")")

        pn = UCase(pn_any_case)
        If pn = "IMG" Then
            PhotoNameKind = 12
        ElseIf (Left(pn, 4) = "IMGP" Or Left(pn, 4) = "I2GP") And Len(pn) = 8 Then
            PhotoNameKind = 11
        ElseIf (Left(pn, 4) = "_IMG" Or Left(pn, 4) = "_IGP" Or Left(pn, 4) = "_I2G") And Len(pn) = 8 Then
            PhotoNameKind = 15
        ElseIf Left(pn, 4) = "CIMG" And Len(pn) = 8 Then
            PhotoNameKind = 13
        ElseIf Left(pn, 3) = "DSC" And Len(pn) = 8 Then
            PhotoNameKind = 14
        ElseIf Len(pn) >= 6 And IsInteger(Left(pn, 6)) Then ' yymmdd
            m_val = Val(Mid(pn, 3, 2))
            If 1 <= m_val And m_val <= 12 Then ' this extra check inserted 2009-01-02, fixed 2009-01-12
                PhotoNameKind = 10
                '          Call dpMsg("return value = 10")
            Else
                PhotoNameKind = 0
                '           Call dpMsg("return value = 0")
            End If
        ElseIf IsInteger(Left(pn, 4)) And IsLetter(Mid(pn, 5, 1)) Then ' yymml[l]
            PhotoNameKind = 20
            '        Call dpMsg("return value = 20")
        ElseIf IsInteger(Mid(pn, 2, 2)) And LCase(Left(pn, 1)) = "s" And Len(pn) = 3 Then
            PhotoNameKind = 21
            '         Call dpMsg("return value = 21")
        ElseIf IsInteger(Mid(pn, 2, 2)) And LCase(Left(pn, 1)) = "a" And Len(pn) = 3 Then
            PhotoNameKind = 22
            '          Call dpMsg("return value = 22")
        ElseIf IsInteger(Mid(pn, 3, 2)) And Left(pn, 1) = "N" And
       ("A" <= Mid(pn, 2, 1) And Mid(pn, 2, 1) <= "Z") And Len(pn) = 4 Then
            PhotoNameKind = 31 ' negative in old b/w negative albums
            '           Call dpMsg("return value = 31")
        Else
            PhotoNameKind = 0
        End If
        '        Call dpEnd("PhotoNameKind")
    End Function

    Public Function PhotoNameSplit(pf_name As String,
                          ByRef path As String,
                          ByRef prefix As String,
                          ByRef proper As String,
                          ByRef seq As String,
                          ByRef tail As String,
                          ByRef ext As String) As Integer
        ' <drive>:\<prefix>_<proper><seq>_<tail>.<ext>
        ' All parts except <proper><seq> and <ext> are optional
        ' <prefix> and <tail> may contain '_' characters
        ' <proper> must be yymmdd..., or one of the native formats IMGPnnnn etc, see below.
        ' <seq> can be "_nn", "nnn", or "-nn" followed by 0 to two letters,
        '                            or "-nnn" followed by 0 or one letter
        '
        ' If the filename doesn't follow the convention, then proper is set to the entire name excl. path and extension
        ' Return value describes the file name format:
        ' 10: yymmdd...   e.g. 050312_19p
        ' 11: IMGPnnnn, I2GPnnnn
        ' 12: IMG
        ' 13: CIMGnnnn
        ' 14: DSCnnnnn
        ' 15: _IMGnnnn, _IGPnnnn, _I2Gnnnn   ' ondskabsfuld p.g. af den der "_"
        ' 20: yymml[l]    e.g. 0504A_03   0504Ak_04
        ' 21: snn_        e.g. s40_01.jpg  -- scanned slide from mag. no. 40
        ' 22: ann_        e.g. a02_01.jpg  -- scanned slide from archive no. 02
        ' 31: Nlnn_Ln     e.g. Na38_F2.jpg -- scanned b/w negative, album "a", page 28, row F, frame 2
        '  0: Otherwise
        '
        ' Return parameters:
        ' seq:     Includes the leading '_' or '-' character if any
        Const ML = 10 ' max allowable number of fields separated by '_'
        Dim pn As String, pn_type As Integer, pn_type_tmp As Integer
        Dim field(ML) As String, p As Integer, p_basis_start As Integer, p_tail_start As Integer, i As Integer
        Dim photofile_name As String
        Static old_photofile_name As String, old_path As String, old_prefix As String, old_proper As String
        Static old_seq As String, old_tail As String, old_ext As String, old_return_value As Integer

        PhotoNameSplit = 0
        '       Call dpBegin("PhotoNameSplit " + pf_name)

        ' Temporarily map type 15 "_" into "*", a character which cannot occur in a file name
        photofile_name = pf_name
        If InStr(photofile_name, "I2GP") = 0 Then
            '            photofile_name = StrReplace(pf_name, "_I2G", "*I2G") ' Oles K20D, Adobe color space
            photofile_name = pf_name.Replace("_I2G", "*I2G") ' Oles K20D, Adobe color space
        End If
        If InStr(photofile_name, "IMGP") = 0 Then
            '            photofile_name = StrReplace(photofile_name, "_IMG", "*IMG") ' Pentax convemtion for Adobe color space
            photofile_name = photofile_name.Replace("_IMG", "*IMG") ' Pentax convemtion for Adobe color space
        End If

        '        photofile_name = StrReplace(photofile_name, "_IGP", "*IGP") ' Pentax convention for Adobe color space
        photofile_name = photofile_name.Replace("_IGP", "*IGP") ' Pentax convention for Adobe color space

        If photofile_name = old_photofile_name Then ' use cached values
            path = old_path
            prefix = old_prefix
            proper = old_proper
            seq = old_seq
            tail = old_tail
            ext = old_ext
            PhotoNameSplit = old_return_value
            ' Debug.Print "Used cache for " & old_photofile_name
            ' Debug.Print "-- END PhotoNameSplit ->" + str(old_return_value)
            Exit Function
        End If

        path = FilenamePath(photofile_name)
        prefix = ""
        proper = ""
        seq = ""
        tail = ""
        ext = FilenameExtension(photofile_name)

        pn = FilenameProper(photofile_name) ' no path, no extension
        ' Debug.Print "--- Proper= " + pn
        If pn = "" Then Exit Function

        p = 0 : pn_type = 0 : p_tail_start = 1
        Do
            p = p + 1
            If p > ML Then ' the name doesn't comply with the convention
                proper = pn
                PhotoNameSplit = 0
                Exit Function
            End If
            field(p) = ExtractField(pn, p, "_")
            If field(p) = "" Then
                p = p - 1
                Exit Do
            End If

            ' map the '*' back to '_'
            '            field(p) = StrSubstitute(field(p), "*", "_")
            field(p) = field(p).Replace("*", "_")

            pn_type_tmp = PhotoNameKind(field(p))
            ' Debug.Print "--- field(" + Format(p) + ")=" + field(p) + "  kind:" + str(pn_type_tmp)
            If pn_type_tmp > 0 And pn_type = 0 Then ' we located the core of the photoname
                p_basis_start = p
                pn_type = pn_type_tmp
            End If
        Loop

        ' Debug.Print "--- p_basis_start=" & str(p_basis_start) & ", pn_type=" & str(pn_type)
        Select Case pn_type
            Case 10           '030431_03p     030431-13p
                proper = Left(field(p_basis_start), 6)
                If Len(field(p_basis_start)) = 6 Then ' sequence number is in next field
                    seq = field(p_basis_start + 1)
                    If Not IsDigit(Mid(seq, 3)) Then seq = "_" + seq
                    p_tail_start = p_basis_start + 2
                Else
                    seq = Mid(field(p_basis_start), 7)
                    p_tail_start = p_basis_start + 1
                End If
            Case 11           'IMGP1234
                proper = Left(field(p_basis_start), 4)
                seq = Mid(field(p_basis_start), 5, 4) ' was , 3
                '      seq = seq + Chr(Asc("a") + val(Right(field(p_basis_start), 1)))
                p_tail_start = p_basis_start + 1
            Case 12           ' IMG   _nnnn
                proper = "IMG_" ' was proper = "IMGC" until 2018-02-09
                seq = Left(field(p_basis_start + 1), 4)
                '      seq = seq + Chr(Asc("a") + val(Right(field(p_basis_start + 1), 1)))
                p_tail_start = p_basis_start + 2
            Case 13           'CIMG1234
                proper = "CIMG"
                seq = Mid(field(p_basis_start), 5, 4)
                '      seq = seq + Chr(Asc("a") + val(Right(field(p_basis_start), 1)))
                p_tail_start = p_basis_start + 1
            Case 14           'CIMG1234
                proper = "DSC"
                seq = Mid(field(p_basis_start), 4, 5)
                '      seq = seq + Chr(Asc("a") + val(Right(field(p_basis_start), 1)))
                p_tail_start = p_basis_start + 1
            Case 15           '_IGP1234
                proper = Left(field(p_basis_start), 4)
                seq = Mid(field(p_basis_start), 5, 4)
                p_tail_start = p_basis_start + 1
            Case 20           '0304B_03  0304B_03A  0304B_03k  0304B_03Ak  0304C-13 etc.
                proper = Left(field(p_basis_start), 5)
                If Len(field(p_basis_start)) = 5 Then ' sequence number is in next field
                    seq = "_" + field(p_basis_start + 1)
                    p_tail_start = p_basis_start + 2
                Else
                    seq = Mid(field(p_basis_start), 6)
                    p_tail_start = p_basis_start + 1
                End If
            Case 21, 22         's40_03 s40_03_tail  prefix_s40_03_tail  a02_01 etc.
                proper = Left(field(p_basis_start), 3)
                seq = "_" + field(p_basis_start + 1)
                p_tail_start = p_basis_start + 2
'                Call dpMsg("Case 21, 22: proper=" + proper + " seq=" + seq + " p_tail_start=" + Str(p_tail_start))
            Case 31         'Na38_F2 Na38_F2_tail  prefix_Na38_F2_tail etc.
                proper = field(p_basis_start)
                seq = "_" + field(p_basis_start + 1)
                p_tail_start = p_basis_start + 2
                '                Call dpMsg("Case 31: proper=" + proper + " seq=" + seq + " p_tail_start=" + Str(p_tail_start))
        End Select

        For i = 1 To p_basis_start - 1
            prefix = prefix + "_" + field(i)
        Next i
        If Len(prefix) > 1 Then prefix = Mid(prefix, 2)

        For i = p_tail_start To p
            tail = tail + "_" + field(i)
        Next i
        If Len(tail) > 1 Then tail = Mid(tail, 2)
        If proper = "" Then
            ' non-conforming file name
            proper = prefix + seq + tail
            prefix = ""
            seq = ""
            tail = ""
            pn_type = 0
        End If
        PhotoNameSplit = pn_type
        ' Load cache
        old_photofile_name = photofile_name
        old_path = path
        old_prefix = prefix
        old_proper = proper
        old_seq = seq
        old_tail = tail
        old_ext = ext
        old_return_value = pn_type

        ' Debug.Print "--- >" & path & "< >" & prefix & "< >" & proper & "< >" & seq & "< >"; tail & "< >" & ext & "<" ' 170209
        '       Call dpEnd("PhotoNameSplit " + photofile_name) ' 170209
    End Function

End Module
