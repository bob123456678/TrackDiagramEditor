Option Compare Text ' Case insensitive string comparison
Imports System.ComponentModel

Public Class Form1
    ' CS2 Track Diagram Editor
    Const Version = "Version 2.2.2b" ' 2025-03-12
    ' 2.2.2:          Changed how deletion of text works
    ' 2.2.1:          Inproved display of text in connection with linksweiche og rechtsweiche
    ' 2.1.x -> 2.2.x: Two track diagrams can be combined into one. Report of digital addresses in sue can be provided.
    ' 2.0.1 -> 2.1.x: A track diagram can span 255 rows and 255 colums rather than just 17 x 30. The limit of 255 is imposed by Märklin's file format.
    '             .3: Display of addresses and text reworked to use event handler.
    ' 2.0   -> 2.0.1: pfeil address entry bug fix. Allow for leading spaces in texts. Doesn't place window as topmost.
    '  d Work-around fix for display of text and addresses after editing
    '  b Backup of deleted pages. Still to do: A mechanism for adding such pages back into the (or a) layout
    ' 1.x -> 2.0: Program can be launched from TC in addition to be run as a stand_alone track diagram editor
    ' 2023-10-15: Version 2: User interface reworked to be more intuitive (I hope)
    ' 2023-05-22: Corrected s88doppelbogen to have but one s88
    ' 2023-05-17: Removed ability to use alternative file type .txt for cs2 files
    ' From version 2023-04-05: Updates the index file "gleisbild.cs2" when required (but only if the CS2 directory convention is adhered to).
    ' This is handled by class instance MF (Master File).
    ' This is (mostly) done "lazily", i.e. accessing is postponed until absolutely required.
    '
    ' Märklin CS2's convention regarding folder structure:
    ' <my_path>\gleisbilder\<track diagram name>.cs2         ' one or more track diagram files, each assigned to a tab in CS2's layout display
    ' <my_path>\gleisbild.cs2                                ' this file constitues an index to the pages (tabs) with layout dagrams
    ' <my_path>\fahrstrassen.cs2                             ' this file holds all routes defined on the CS2 Memory page tabs
    '
    ' This layout editor handles the above structure. However, should the track diagrams not reside in a folder named "gleisbilder" we can handle
    ' that to a limited extent ("flat structure"): fahrstrassen.cs2 should reside in the same folder as the track diagrams and pages cannot be managed.

    'Installation (only relevant if running in Stand Alone mode is desired)
    '----------------------------------------------------------------------
    'Place TrackDiagramEditor.exe in some folder of your choice (the installation folder).
    'Optional create a short cut on the desktop.

    'Requires Windows 10 Or higher, a 64 bit CPU and the .NET framework 4.7.2 or higher.

    ' Form size
    Const ROW_MAX = 16 ' size of track diagram window (counting from zero)
    Const COL_MAX = 29
    Const W_MIN = 1155 ' minimum width allowed when resizing
    Const H_MIN = 737 '  minimum height allowed when resizing

    ' track diagram window position
    Dim Row_Offset As Integer, Col_Offset As Integer ' position of upper left corner of the grid

    Dim Test_Level As Integer  ' >=3: displays the input file, >=4 displays the output file as well

    Shared MODE As String = "View" ' "View" or "Edit" or "Routes" or "Join"
    Dim ROUTE_VIEW_ENABLED As Boolean = False ' This mode is only of interest to, and available for, layouts created by a cs2 and
    ' requires the route file ("fahrstrassen.cs2") to be availabe

    Dim Use_English_Terms As Boolean = True ' set or cleared in tools menu
    Dim Verbose As Boolean = False ' set in tools menu
    Dim Reveal_Text As Boolean = False
    Dim Reveal_Addr As Boolean = False
    Dim Reveal_Text_State As Boolean ' Records the Reveal_Text value upon entry to Route mode
    Dim Reveal_Addr_State As Boolean ' Records the Reveal_Addr value upon entry to Route mode

    Dim FILE_LOADED As Boolean

    Dim Current_Directory As String
    Dim TrackDiagram_Directory As String ' where the track diagram files reside. Per Märklin should be named ...\gleisbilder

    Shared TrackDiagram_Input_Filename As String ' most recent track diagram that was loaded
    Shared MasterFile_Directory As String ' where gleisbild.cs2 and fahrstrassen.cs2 resides, if those files are present
    Dim Flat_Structure As Boolean = False ' False if Märklin's folder structure is adhered to. Refer the comment at the top of this file.
    Shared STAND_ALONE As Boolean = True ' False if the editor is run from a command line. This is the case when the editor is activated from the train control program "TC"
    Dim CMDLINE_MODE As String = "" ' "Edit" if run from command line, otherwise "". Future versions may allow more options
    Dim Cmdline_Param1 As String
    Dim Cmdline_Param2 As String

    Dim Input_Drive As String ' The drive last read or imported from
    Dim Start_Directory As String
    Dim Active_Tool As String = ""
    ReadOnly tt As New ToolTip()
    Dim el_moving As ELEMENT
    Const HEADER_MAX = 10
    Dim Header_Line(HEADER_MAX) As String ' holds the header lines of the cs2 layout file up to and excluding the first line with ".element"
    Dim Header_Top As Integer

    Private Structure ELEMENT
        Public id As String ' hex number identifying layout page no. and cell in the grid. Ex: 0x30b
        ' The string holds: page_no, row_no, and col_no. Leading zeroes are omitted, however.
        Public typ_s As String ' type of element: prellbock, gerade, etc.
        Public drehung_i As Integer ' direction of the element 0, 1, 2, 3
        Public artikel_i As Integer ' encoding of the element's digital address
        Public deviceId_s As String ' New in CS2 version 4, used with chained s88. Hex string
        Public zustand_s As String ' Optional - state of the element when CS2 was closed down (red, green, etc.) - of no concern to us
        Public text As String
        ' Derived from the above
        Public id_normalized As String ' 6 characters, two for each of page_no, row_no, and col_no - in hex. Derived from the id.
        Public row_no As Integer ' derived form 'id' converted to decimal
        Public col_no As Integer ' derived form 'id' converted to decimal
        ' Example: ".id=0x2070a" mwans page no 2 hex (2 decimal), col no. 07 hex (7 decimal), row no 0a hex (10 decimal)
        ' Leading zeros are omitted

        Public my_index As Integer ' My place in the Elements array
    End Structure

    Const E_MAX = 1000 ' We can have at most E_MAX elements (gerade, bogen, ..., etc.). This an arbitrary limit. Increase as needed.
    Shared E_TOP As Integer ' Number of elements in track diagram (= index of last ELEMENT stored in Elements(*).
    Shared Elements(E_MAX) As ELEMENT
    Dim Empty_Element As ELEMENT = New ELEMENT

    ' Temporary storage of elements prior to editing - allows us to revert and cancel edits 
    Dim E_TOP2 As Integer ' We actually have E_TOP2 elements
    Dim Elements2(E_MAX) As ELEMENT

    Dim Page_No As Integer ' the page (tab) in CS2 or TC counting from 0 with the track diagram. This value is encoded in the ID of the elements

    Dim Developer_Environment As Boolean = False ' Is set to True in Form1_Load if we're running in the development environment
    Public Shared CHANGED As Boolean
    Public EDIT_CHANGES As Boolean

    Public env1 As New EnvironmentHandler3

    ' The below is for drawing addresses and other text on the track diagram
    ReadOnly Brush_black As Brush = New Drawing.SolidBrush(Color.Black)
    ReadOnly Brush_blue As Brush = New Drawing.SolidBrush(Color.Blue)
    ReadOnly Brush_red As Brush = New Drawing.SolidBrush(Color.Red)
    ReadOnly Brush_yellow As Brush = New Drawing.SolidBrush(Color.Yellow)
    ReadOnly Brush_white As Brush = New Drawing.SolidBrush(Color.White)

    ReadOnly myFont As Font = New System.Drawing.Font(“Verdana”, 8)
    Dim drawFormatVertical = New StringFormat()

    Class RouteHandler
        ' For analyzing routes in file fahrstrassen.cs2 if present
        Private Structure ROUTE_ELEMENT
            Dim param_id As String
            Dim param_value As String
            Dim stellung As String
            Dim sekunde As Single
        End Structure
        Dim R_MAX As Integer = 2000 ' Arbitrary constant. Wil be increased ig needed and data arrays ReDim'ed.
        Dim R_Top As Integer
        Dim Route_Data(R_MAX) As ROUTE_ELEMENT ' We only store what we need from the route file being read. 
        ' <parater_kind,parameter_value,stellung> stellung == "0" if not specified
        Dim RN_MAX As Integer = 320
        Dim RN_TOP As Integer
        Dim Route_Name(RN_MAX) As String
        Dim Route_Index(RN_MAX) As Integer ' index into Route_Data
        Public Function FahrStrassenLoad() As Boolean
            ' Load the route file "fahrstrassen.cs2" and return true if it exists.
            Dim fn As String, fno As Integer
            Dim buffer As String, param_kind As String, param_val As String
            Dim most_recent_id As String = ""
            Dim most_recent_route_element As String = ""

            Form1.cboRouteNames.Text = "Pick a route"
            Form1.lstRoutes.Items.Clear()
            Form1.cboRouteNames.Items.Clear()
            R_Top = 0
            RN_TOP = 0

            fn = Form1.MasterFile_Directory & "fahrstrassen.cs2"
            If Not FileExists(fn) Then
                MsgBox("Route file " & fn & " does not exist. Note that only routes created by a CS2 can be browsed")
                Form1.Message("Route file " & fn & " does not exist")
                Return False
            Else
                Form1.Message("Loading route file: " & fn)
            End If

            ' Read the route file and load the route names into the drop down

            fno = FreeFile()

            ' Process the route file
            FileOpen(fno, fn, OpenMode.Input)
            While Not (EOF(fno))
                If R_Top = R_MAX Then
                    R_MAX = R_MAX + 500
                    ReDim Preserve Route_Data(R_MAX)
                End If
                If RN_TOP = RN_MAX Then
                    RN_MAX = RN_MAX + 32
                    ReDim Preserve Route_Name(R_MAX)
                    ReDim Preserve Route_Index(R_MAX)
                End If

                buffer = Trim(LineInput(fno))
                param_kind = ExtractField(buffer, 1, "=")
                param_val = ExtractField(buffer, 2, "=")

                Select Case param_kind
                    Case ".id" ' address of the route being read
                        most_recent_id = param_val ' stored so as to be attached to next .name
                    Case ".name" : Form1.cboRouteNames.Items.Add(param_val)
                        R_Top = R_Top + 1
                        Route_Data(R_Top).param_id = param_val
                        Route_Data(R_Top).param_value = most_recent_id
                        If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                        RN_TOP = RN_TOP + 1
                        Route_Name(RN_TOP) = param_val
                        Route_Index(RN_TOP) = R_Top
                    Case ".extern" ' follows right after .name, if present. Indicates autoexecution of route is active
                        Route_Data(R_Top).stellung = param_val
                        If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                    Case ".s88"
                        R_Top = R_Top + 1
                        Route_Data(R_Top).param_id = param_kind
                        Route_Data(R_Top).param_value = param_val
                        Route_Data(R_Top).stellung = ""
                        If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                    Case ".s88Ein" ' follows right after .s88, if present. Indicates route fires when s88 goes from on to off
                        Route_Data(R_Top).stellung = param_val
                        If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                    Case ".s88Flag"
                        R_Top = R_Top + 1
                        Route_Data(R_Top).param_id = param_kind
                        Route_Data(R_Top).param_value = ""
                        Route_Data(R_Top).stellung = ""
                        If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                    Case "..kont"
                        Route_Data(R_Top).param_value = param_val
                        If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                    Case "..hi"
                        Route_Data(R_Top).stellung = param_val
                        If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                    Case "..typ"
                        ' only typ=mag will added to our route
                        most_recent_route_element = param_val
                        If param_val = "mag" Then
                            R_Top = R_Top + 1
                            Route_Data(R_Top).param_id = ""
                            Route_Data(R_Top).param_value = ""
                            Route_Data(R_Top).stellung = "0"
                            If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                        Else
                            If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add("(" & buffer & ")")
                        End If
                    Case "..magnetartikel"
                        Route_Data(R_Top).param_id = param_kind
                        Route_Data(R_Top).param_value = param_val
                        Route_Data(R_Top).stellung = "0"
                        If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                    Case "..sekunde"
                        If most_recent_route_element = "mag" Then
                            Route_Data(R_Top).sekunde = Val(param_val)
                            If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                        Else
                            If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add("(" & buffer & ")")
                        End If
                    Case "..stellung"
                        Route_Data(R_Top).stellung = param_val
                        If Form1.Test_Level > 0 Then Form1.lstRoutes.Items.Add(buffer)
                    Case Else ' discard this line
                        If Form1.Test_Level > 0 Then
                            If buffer = "fahrstrasse" Then Form1.lstRoutes.Items.Add("")
                            Form1.lstRoutes.Items.Add("(" & buffer & ")")
                        End If
                End Select
            End While
            On Error GoTo 0
            FileClose(fno)
            If Form1.Test_Level > 3 Then FahrstrassenDump()
            Return True
route_read_error:
            On Error GoTo 0
            FileClose(fno)
            MsgBox("Cannot read the route file " & fn)
            R_Top = 0
            Return False
        End Function

        Private Function StateExpand(element_type As String, state As String) As String
            Dim s As String = ""
            s = state & " "
            If InStr(element_type, "signal") > 0 Then
                Select Case state
                    Case "0"
                        Return s & "red"
                    Case Else
                        Return s & "green"
                End Select
            ElseIf InStr(element_type, "links") > 0 Or InStr(element_type, "rechts") > 0 Then
                Select Case state
                    Case "0"
                        Return s & "curve"
                    Case Else
                        Return s & "straight"
                End Select
            ElseIf InStr(element_type, "dreiweg") > 0 Then
                Select Case state
                    Case "0"
                        Return s & "left"
                    Case "1"
                        Return s & "straight"
                    Case Else
                        Return s & "right"
                End Select
            Else
                Return state
            End If
        End Function

        Public Sub RouteShow(r_name As String)
            Dim i As Integer, r_index As Integer, found As Boolean = False
            Dim ll As Integer, el As Form1.ELEMENT
            Dim cell As Object, elements_not_found As String = ""
            ' Find the route in question
            Form1.lstRoutes.Items.Clear()
            Form1.lstRoutes.Visible = True
            Form1.GleisbildGridSet(False)
            Form1.GleisBildDisplayClear()
            Form1.GleisbildDisplay()

            For i = 1 To RN_TOP
                If Route_Name(i) = r_name Then
                    r_index = Route_Index(i)
                    found = True
                    Exit For
                End If
            Next i

            If found Then
                'Form1.lstRoutes.Items.Clear()
                ll = r_index
                For i = r_index To R_Top Step 1

                    ' given the dig address, locate the element on the layout display
                    el = Form1.ElementGet3(Route_Data(i).param_id, Val(Route_Data(i).param_value))

                    If Strings.Left(Route_Data(i).param_id, 1) <> "." And i > ll Then Exit For
                    ' List the steps of the route - two lines per step
                    Form1.lstRoutes.Items.Add(Route_Data(i).param_id)
                    If Strings.Left(Route_Data(i).param_id, 1) <> "." Then
                        If Route_Data(i).stellung <> "" Then
                            Form1.lstRoutes.Items.Add("addr=" & StrFixR(Route_Data(i).param_value, " ", 3) & ", state=" & "Auto")
                        Else
                            Form1.lstRoutes.Items.Add("addr=" & StrFixR(Route_Data(i).param_value, " ", 3) & ", state=" & "Manual")
                        End If
                    ElseIf Route_Data(i).param_id = ".s88" Then
                        If Route_Data(i).stellung <> "" Then
                            Form1.lstRoutes.Items.Add("addr=" & StrFixR(Route_Data(i).param_value, " ", 3) & ", state=" & Route_Data(i).stellung)
                            Form1.lstRoutes.Items.Add(" fire if on->off")
                        Else
                            Form1.lstRoutes.Items.Add("addr=" & StrFixR(Route_Data(i).param_value, " ", 3) & ", state=" & Route_Data(i).stellung)
                            Form1.lstRoutes.Items.Add(" fire if off>on")
                        End If
                    ElseIf Route_Data(i).param_id = ".s88Flag" Then
                        If Route_Data(i).stellung <> "" Then
                            Form1.lstRoutes.Items.Add("addr=" & StrFixR(Route_Data(i).param_value, " ", 3) & ", state=" & Route_Data(i).stellung & "  free")
                        Else
                            Form1.lstRoutes.Items.Add("addr=" & StrFixR(Route_Data(i).param_value, " ", 3) & ", state=" & "  occupied")
                        End If
                    Else
                        Form1.lstRoutes.Items.Add(Form1.Translate(el.typ_s))
                        Form1.lstRoutes.Items.Add("ad=" & StrFixR(Route_Data(i).param_value, " ", 3) & ", st=" & StateExpand(el.typ_s, Route_Data(i).stellung))
                        If Route_Data(i).sekunde > 0 Then
                            Form1.lstRoutes.Items.Add("delay=" & Str(Route_Data(i).sekunde) & " seconds")
                        End If
                    End If

                    ' highlight the element on the layout display
                    Form1.Message(2, "RS1: " & el.typ_s & Str(el.row_no) & Str(el.col_no))
                    If el.typ_s <> "" Then
                        cell = Form1.PicGet(el.row_no, el.col_no)
                        If Not cell Is Nothing Then
                            cell.BorderStyle = 1
                            ' I have only included those elements which I use in routes. That is: The list is incomplete
                            If el.typ_s = "signal" And Val(Route_Data(i).stellung) = 0 Then
                                Form1.ElementDisplay(cell, el.drehung_i, Form1.picIconSignalRed.Image)
                            ElseIf Strings.Left(el.typ_s, 8) = "signal_f" And Val(Route_Data(i).stellung) = 0 Then
                                Form1.ElementDisplay(cell, el.drehung_i, Form1.picIconSignalFHP01red.Image)
                            ElseIf Strings.Left(el.typ_s, 8) = "signal_f" And Val(Route_Data(i).stellung) = 1 Then
                                Form1.ElementDisplay(cell, el.drehung_i, Form1.picIconSignalFHP01.Image)
                            ElseIf Strings.Left(el.typ_s, 8) = "signal_p" And Val(Route_Data(i).stellung) = 0 Then
                                Form1.ElementDisplay(cell, el.drehung_i, Form1.picIconSignalPHP012red.Image)
                            ElseIf Strings.Left(el.typ_s, 9) = "signal_sh" And Val(Route_Data(i).stellung) = 0 Then
                                Form1.ElementDisplay(cell, el.drehung_i, Form1.picIconSignalSH01red.Image)
                            ElseIf el.typ_s = "linksweiche" And Val(Route_Data(i).stellung) = 0 Then
                                Form1.ElementDisplay(cell, el.drehung_i, Form1.picIconSwitchLeftActive.Image)
                            ElseIf el.typ_s = "rechtsweiche" And Val(Route_Data(i).stellung) = 0 Then
                                Form1.ElementDisplay(cell, el.drehung_i, Form1.picIconSwitchRightActive.Image)
                            ElseIf el.typ_s = "dreiwegweiche" And Val(Route_Data(i).stellung) = 0 Then
                                Form1.ElementDisplay(cell, el.drehung_i, Form1.picIconThreeWayActive.Image)
                            ElseIf el.typ_s = "dreiwegweiche" And Val(Route_Data(i).stellung) = 2 Then
                                Form1.ElementDisplay(cell, el.drehung_i, Form1.picIconThreeWayActive2.Image)
                            ElseIf el.typ_s = "s88kontakt" AndAlso Route_Data(i).stellung = "" Then
                                If Route_Data(i).param_id <> ".s88Flag" Then
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88OffOn.Image)
                                Else
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88active.Image)
                                End If
                            ElseIf el.typ_s = "s88kontakt" AndAlso Route_Data(i).stellung <> "" Then
                                If Route_Data(i).param_id <> ".s88Flag" Then
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88OnOff.Image)
                                Else
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88.Image)
                                End If
                            ElseIf el.typ_s = "s88bogen" AndAlso Route_Data(i).stellung = "" Then
                                If Route_Data(i).param_id <> ".s88Flag" Then
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88CurveOffOn.Image)
                                Else
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88CurveActive.Image)
                                End If
                            ElseIf el.typ_s = "s88bogen" AndAlso Route_Data(i).stellung <> "" Then
                                If Route_Data(i).param_id <> ".s88Flag" Then
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88CurveOnOff.Image)
                                Else
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88Curve.Image)
                                End If
                            ElseIf el.typ_s = "s88doppelbogen" AndAlso Route_Data(i).stellung = "" Then
                                If Route_Data(i).param_id <> ".s88Flag" Then
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88CurveParallelOffOn.Image)
                                Else
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88CurveParallelActive.Image)
                                End If
                            ElseIf el.typ_s = "s88doppelbogen" AndAlso Route_Data(i).stellung <> "" Then
                                If Route_Data(i).param_id <> ".s88Flag" Then
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88CurveParallelOnOff.Image)
                                Else
                                    Form1.ElementDisplay(cell, el.drehung_i, Form1.picIcons88CurveParallel.Image)
                                End If
                            End If
                        ElseIf Route_Data(i).stellung <> "" AndAlso Strings.Left(Route_Data(i).param_id, 1) = "." Then
                            elements_not_found = elements_not_found & " " & Route_Data(i).param_value
                        End If ' Not cell Is Nothing
                    End If
                Next i
                If elements_not_found <> "" Then
                    Form1.Message("Route """ & Form1.cboRouteNames.Text & """: Elements with these addresses are not present on the current layout:")
                    Form1.Message(elements_not_found)
                End If
            End If
        End Sub
        Public Function RouteName(route_address As Integer) As String
            ' Find the name of the route with this address
            For i = 1 To R_Top
                If Strings.Left(Route_Data(i).param_id, 1) <> "." AndAlso Val(Route_Data(i).param_value) = route_address Then
                    Form1.Message(3, Route_Data(i).param_id & " " & Route_Data(i).param_value & " " & Route_Data(i).stellung)
                    Return Route_Data(i).param_id
                End If
            Next i
            Return ""
        End Function
        Private Sub FahrstrassenDump()
            ' dumps the data into the message list
            Dim i As Integer
            For i = 1 To RN_TOP
                Form1.Message(StrFixL(Route_Name(i), " ", 12) & Str(Route_Index(i)))
            Next
            For i = 1 To R_Top
                If Route_Data(i).stellung <> "" Then
                    Form1.Message(StrFixLong(i, 4) & ": " & StrFixL(Route_Data(i).param_id, " ", 17) & "addr=" & StrFixR(Route_Data(i).param_value, " ", 3) & ", stellung=" & Route_Data(i).stellung)
                Else
                    Form1.Message(StrFixLong(i, 4) & ": " & StrFixL(Route_Data(i).param_id, " ", 17) & "addr=" & StrFixR(Route_Data(i).param_value, " ", 3))
                End If
            Next
        End Sub
    End Class

    Class PageFileHandler
        Private PageFile_Name As String
        Private PL_MAX As Integer = 20
        Private PL_Top As Integer = -1 ' layout pages are counted from 0
        Private N_HeaderLines As Integer
        Private Structure PAGE
            Dim id As Integer
            Dim name As String
            Dim xOffSet As Integer, yOffSet As Integer
            ' --- additional fields not in CS2 file
            Dim file_name As String ' "" if the file doesn't exist (the gleisbild file is thus inconsistent/invalid)
        End Structure
        Private Page_List(PL_MAX) As PAGE
        Public Function PageFileExists(mf_dir As String) As Boolean
            Return PageFileLocate(mf_dir) <> ""
        End Function
        Private Function PageFileLocate(mf_dir As String) As String
            If FileExists(mf_dir & "gleisbild.cs2") Then
                Return mf_dir & "gleisbild.cs2"
            Else
                Return ""
            End If
        End Function
        Public Function PageFileLoad(mf_dir As String) As Boolean
            ' Loads the "gleisbild.cs2" master file, if it exists. Then returns True.
            ' Otherwise, returns false.
            Dim fn As String, fno As Integer, buffer As String = ""
            Dim param_kind As String, param_val As String, fn_tmp As String

            If mf_dir = "" Then Return False

            PageFile_Name = ""

            fn = PageFileLocate(mf_dir)

            PL_Top = -1
            N_HeaderLines = 0
            PageFile_Name = ""

            Form1.lstMasterFile.Items.Clear()
            If FileExists(fn) Then
                PageFile_Name = fn
                fno = FreeFile()
                On Error GoTo MF_Read_Error
                FileOpen(fno, fn, OpenMode.Input)

                ' Process the header lines
                While Not (EOF(fno))
                    buffer = LineInput(fno)
                    Form1.lstMasterFile.Items.Add(buffer)
                    If buffer = "seite" Then Exit While
                    N_HeaderLines = N_HeaderLines + 1
                    Select Case N_HeaderLines
                        Case 1 : If buffer <> "[gleisbild]" Then
                                Form1.Message(fn & " is not a valid CS2 layout map. You can continue working on the layout regardless")
                                On Error GoTo 0
                                FileClose(fno)
                                PageFile_Name = ""
                                Return False
                            End If
                        Case Else
                    End Select
                End While

                ' Process the layout pages. Buffer contains "seite"
                While Not EOF(fno)
                    ' buffer should contains "seite" at the beginning of each iteration.
                    If buffer <> "seite" Then
                        Form1.Message(fn & " is not a valid CS2 layout map file. You can continue working on the layout regardless.")
                        FileClose(fno)
                        On Error GoTo 0
                        FileClose(fno)
                        PageFile_Name = ""
                        Return False
                    End If
                    PL_Top = PL_Top + 1
                    If PL_Top > PL_MAX Then
                        PL_MAX = PL_MAX + 10
                        ReDim Preserve Page_List(PL_MAX)
                    End If

                    Page_List(PL_Top).id = 0 ' CS2 omits this value in the file when 0
                    Page_List(PL_Top).name = ""
                    Page_List(PL_Top).xOffSet = 0 ' CS2 omits this value in the file when 0
                    Page_List(PL_Top).yOffSet = 0 ' CS2 omits this value in the file when 0
                    Page_List(PL_Top).file_name = ""

                    While Not EOF(fno)
                        buffer = LineInput(fno)
                        Form1.lstMasterFile.Items.Add(buffer)
                        buffer = Trim(buffer)
                        If buffer = "seite" Then Exit While
                        param_kind = Trim(ExtractField(buffer, 1, "="))
                        param_val = Trim(ExtractField(buffer, 2, "="))
                        Select Case param_kind
                            Case ".id" : Page_List(PL_Top).id = Val(param_val)
                            Case ".name" : Page_List(PL_Top).name = param_val
                                fn_tmp = Form1.TrackDiagram_Directory & param_val & ".cs2"
                                If FileExists(fn_tmp) Then
                                    Page_List(PL_Top).file_name = fn_tmp
                                End If
                            Case ".xoffset" : Page_List(PL_Top).xOffSet = Val(param_val)
                            Case ".yoffset" : Page_List(PL_Top).yOffSet = Val(param_val)
                            Case Else
                                ' ignore empty lines
                        End Select
                    End While
                End While
                On Error GoTo 0
                FileClose(fno)
                Form1.Message(1, "Page file loaded: """ & fn & """")
                Form1.Message(2, "  No. of lines in page file" & Str(Form1.lstMasterFile.Items.Count))
            Else
                Form1.Message(1, "Could not locate CS2 layout map ""gleisbild"". (m21)")
                On Error GoTo 0
                Return False
            End If
            Return True
MF_Read_Error:
            On Error GoTo 0
            FileClose(fno)
            Form1.Message("Error reading CS2 layout map ""gleisbild"". You can continue working on the layout regardless. (m22)")
            PageFile_Name = ""
            PL_Top = -1
            Return False
        End Function

        Public Sub PageFileSave()
            Dim fno As Integer, i As Integer, ul As Integer
            If PL_Top = -1 Then ' no file to save
            Else
                fno = FreeFile()
                On Error GoTo MF_Write_Error
                FileOpen(fno, PageFile_Name, OpenMode.Output)
                ul = N_HeaderLines - 1
                For i = 0 To ul
                    If i = ul Then ' replace the last used page with page 0 since the contents of lstMasterFile could be outdated
                        If Form1.lstMasterFile.Items.Item(i - 1) = "zuletztBenutzt" Then
                            PrintLine(fno, " .name=" & Page_List(0).name)
                        Else
                            PrintLine(fno, Form1.lstMasterFile.Items.Item(i))
                        End If
                    Else
                        PrintLine(fno, Form1.lstMasterFile.Items.Item(i))
                    End If
                Next i

                For i = 0 To PL_Top
                    PrintLine(fno, "seite")
                    If i > 0 Then PrintLine(fno, " .id=" & Format(i))
                    PrintLine(fno, " .name=" & Page_List(i).name)
                    If Page_List(i).xOffSet > 0 Then PrintLine(fno, " .xoffset=" & Format(Page_List(i).xOffSet))
                    If Page_List(i).yOffSet > 0 Then PrintLine(fno, " .yoffset=" & Format(Page_List(i).yOffSet))
                Next
                On Error GoTo 0
                FileClose(fno)
            End If
            Exit Sub
MF_Write_Error:
            On Error GoTo 0
            FileClose(fno)
            Form1.Message("Error writing CS2 layout map """ & PageFile_Name & """ - write protected? Your layout file is not affected.")
        End Sub
        Public Sub PageReplaceName(pno As Integer, new_file_name As String)
            If 0 <= pno And pno <= PL_Top Then
                Page_List(pno).file_name = new_file_name
                Page_List(pno).name = FilenameNoExtension(FilenameExtract(new_file_name))
                PageFileSave()
            Else
                MsgBox("Internal error, PageReplaceName called with page number out of range")
            End If
        End Sub
        Public Function LayoutFileNameGet(pno As Integer) As String
            ' Returns the file name of the layout in page number pno. Returns "" if no such page or file
            If pno > PL_Top Then
                Return ""
            Else
                Return Page_List(pno).file_name
            End If
        End Function
        Public Function LayoutNameGet(pno As Integer) As String
            ' Returns the layout name of page number pno. Returns "" if no such page
            If pno > PL_Top Or pno < 0 Then
                Return ""
            Else
                Return Page_List(pno).name
            End If
        End Function
        Public Function LayoutPageNumberGet(track_diagram As String) As Integer
            ' returns the page number of the track diagram. Returns -1 if the diagram is not in Page_List.
            ' The 'track_diagram' parameter can be the full file name or the page name (aka core file name without extenbsion)
            Dim i As Integer
            For i = 0 To PL_Top Step 1
                If track_diagram = Page_List(i).file_name OrElse track_diagram = Page_List(i).name Then Return i
            Next i
            Return -1
        End Function
        Public Function NextAvailablePageNumber() As Integer
            Return PL_Top + 1
        End Function
        Public Function NextAvailablePageNumber(layout_name As String) As Integer
            PageFileLoad(Form1.MasterFile_Directory)
            Return PL_Top + 1
        End Function
        Public Function NumberOfPages() As Integer
            Return (PL_Top + 1)
        End Function

        Public Sub PageAdd(track_diagram_file As String)
            Dim page_name As String

            page_name = FilenameNoExtension(FilenameExtract(track_diagram_file))

            If (PL_Top + 1) = Form1.Page_No Then
                PL_Top = PL_Top + 1
                If PL_Top > PL_MAX Then
                    PL_MAX = PL_MAX + 10
                    ReDim Preserve Page_List(PL_MAX)
                End If

                Page_List(PL_Top).id = Format(PL_Top)
                Page_List(PL_Top).name = page_name
                Page_List(PL_Top).file_name = track_diagram_file
                PageFileSave()
            Else
                MsgBox("Error in program logic: Gap in page numbers not allowed. Layout map file was not created/updated (m61)")
            End If
        End Sub
        Public Function PageFileCreate(track_file As String) As Boolean
            ' Track file must have page 0.
            ' The page file will get extension .cs2
            Dim page_name As String

            page_name = FilenameNoExtension(FilenameExtract(track_file))

            If Form1.Page_No <> 0 Then
                MsgBox("Internal error: Track diagram """ & page_name & """ must be assigned to page 0. Page file not created (m65)")
                Return False
            End If

            ' create the header for the new file
            Form1.lstMasterFile.Items.Clear()
            Form1.lstMasterFile.Items.Add("[gleisbild]")
            Form1.lstMasterFile.Items.Add("Version")
            Form1.lstMasterFile.Items.Add(" .major=1")
            Form1.lstMasterFile.Items.Add("groesse")
            Form1.lstMasterFile.Items.Add("zuletztBenutzt")
            Form1.lstMasterFile.Items.Add(" .name=" & page_name)
            N_HeaderLines = 6
            PL_Top = 0
            Page_List(0).id = "0"
            Page_List(0).name = page_name
            Page_List(0).file_name = track_file
            PageFile_Name = Form1.MasterFile_Directory & "gleisbild.cs2"
            PageFileSave()
            Return True
        End Function
        Public Function PageFileNameGet() As String
            ' Returns the file name of the page map. Returns "" if no such page or file
            Return PageFile_Name
        End Function
        Public Function PageRemove(pno As Integer) As Boolean
            Dim p As Integer
            If PL_Top > 0 Then
                PL_Top = PL_Top - 1
                For p = pno To PL_Top
                    Page_List(p) = Page_List(p + 1)
                Next p
                PageFileSave()
                Return True
            Else
                ' the last remaining page cannot be removed
                Return False
            End If
        End Function
        Public Sub PageReplace(pno As Integer, fn As String)
            Page_List(pno).name = FilenameProper(fn)
            Page_List(pno).file_name = fn
            PageFileSave()
        End Sub
        Public Sub PagesPopulate()
            Dim i As Integer
            Form1.lstLayoutPages.Items.Clear()
            For i = 0 To PL_Top
                Form1.lstLayoutPages.Items.Add(Page_List(i).name)
                If Page_List(i).file_name = "" Then
                    Form1.Message("!! Track file missing for page " & Page_List(i).name)
                End If
            Next
            If Form1.Page_No >= 0 And Form1.Page_No <= (Form1.lstLayoutPages.Items.Count - 1) Then
                Form1.lstLayoutPages.SelectedIndex = Form1.Page_No
            End If
        End Sub
    End Class

    Class CombinePagesHandler
        Private Combine_Pages(2) As String ' name of the pages to combine
        Private Combine_Files(2) As String ' corresponding file names
        Private Combined_Page_Name As String = "" ' name of tje new page
        Private Combined_File_Name As String = "" ' name of tje new page
        Public Sub Initiate()
            Dim answer As DialogResult, current_page_name As String, new_page_name As String, c As String
            Combined_Page_Name = ""
            Combine_Pages(1) = ""
            Combine_Pages(2) = ""

            If CHANGED Then
                current_page_name = FilenameProper(TrackDiagram_Input_Filename)
                If STAND_ALONE Then answer = MsgBox("Track diagram """ & current_page_name & """ has been modified but not saved. Should changes be saved?", vbYesNo)
                If answer = vbYes Or Not STAND_ALONE Then
                    If Form1.GleisbildSave() Then
                        If STAND_ALONE Then Form1.Message("Track diagram saved in: " & TrackDiagram_Input_Filename)
                    Else
                        MsgBox("Page could not be saved: " & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)))
                        If STAND_ALONE Then Form1.Message("to folder " & FilenamePath(TrackDiagram_Input_Filename))
                    End If
                End If
            End If

            new_page_name = InputBox("Name of the new page", "Page name", "")

            If new_page_name <> "" Then
                If new_page_name = "gleisbild" Or new_page_name = "fahrstrassen" Then
                    MsgBox("The name """ & new_page_name & """ cannot be used as a name for a track diagram")
                    Exit Sub
                End If

                If PFH.LayoutPageNumberGet(new_page_name) <> -1 Then
                    MsgBox("Page name already in use. Operation cancelled")
                    Exit Sub
                End If

                c = ValidateFileName(new_page_name)
                If c <> "" Then
                    MsgBox("Unwanted character in page name: '" & c & "'. Operation cancelled")
                    Exit Sub
                End If

                ' We're good to continue
                MODE = "Join"
                Combined_Page_Name = new_page_name
                Combined_File_Name = MasterFile_Directory & "gleisbilder\" & new_page_name & ".cs2"

                Form1.cmdCancel.Text = "Cancel"
                Form1.AdjustPage()
                Form1.GleisBildDisplayClear()
                Form1.lstLayoutPages.SelectedIndex = -1
                Form1.txtCombine.Text = "Click in the box above on the two track diagrams to combine." & vbCrLf & vbCrLf &
                "Each of the two diagrams must have exactly one arrow element. The tracks will be joined as indicated by the arrow elements."
                Form1.txtCombine.Visible = True ' The instructions as to what to do next
            End If
        End Sub
        Public Sub PageSelected(page_name As String, file_name As String)
            If Combine_Pages(1) = "" Then
                Combine_Pages(1) = page_name
                Combine_Files(1) = file_name
            Else
                Combine_Pages(2) = page_name
                Combine_Files(2) = file_name
                Form1.txtCombine.Visible = False
                CombinePages()
            End If
        End Sub
        Private Sub CombinePages()
            Dim joining_point_index(2) As Integer ' index into elements of the pfeil joining points
            Dim r(3) As Integer, c(3) As Integer ' (1) and (2) : the location of pfeil. (3) location where they initially should be moved to to ensure
            Dim rm(2) As Integer, cm(2) As Integer ' vectors indicating how to reposition the two diagrams so that they will match
            Dim combined_page_no As Integer, el As ELEMENT, s As String
            ' that the combined diagram stays within bounds
            Dim direction(2) As Integer ' the direction of the joining point
            Dim i As Integer

            For i = 1 To 2
                joining_point_index(i) = LocateJoiningPoint(i) ' side effect: loads the track diagram
                Select Case joining_point_index(i)
                    Case -1 : OperationFailed("Page """ & Combine_Pages(i) & """ has no arrow element. Pages were not combined", 1)
                        Exit Sub
                    Case -2 : OperationFailed("Page """ & Combine_Pages(i) & """ has multiple arrow elements. Pages were not combined", 2)
                        Exit Sub
                End Select
                direction(i) = Elements(joining_point_index(i)).drehung_i
                r(i) = Elements(joining_point_index(i)).row_no
                c(i) = Elements(joining_point_index(i)).col_no
            Next i

            ' check that the joining points match, i.e. points at each other
            If direction(1) <> ((direction(2) + 2) Mod 4) Then
                OperationFailed("The joining points of the two track diagrams do not match", 1)
                Exit Sub
            End If

            ' calculate the new, common joining point
            r(3) = Math.Max(r(1), r(2))
            c(3) = Math.Max(c(1), c(2))

            ' calculate the move vectors
            rm(1) = r(3) - r(1) : cm(1) = c(3) - c(1)
            rm(2) = r(3) - r(2) : cm(2) = c(3) - c(2)
            ' adjust and close the gap where the pfeil elements are
            If direction(1) = 0 Then
                cm(1) = cm(1) + 1
            ElseIf direction(1) = 3 Then
                rm(1) = rm(1) + 1
            End If
            If direction(2) = 0 Then
                cm(2) = cm(2) + 1
            ElseIf direction(2) = 3 Then
                rm(2) = rm(2) + 1
            End If

            ' duplicate the first page
            combined_page_no = PFH.NextAvailablePageNumber
            Form1.GleisbildLoad(Combine_Files(1))
            'fix the page number
            Form1.ElementsRepaginate(combined_page_no)
            Form1.Page_No = combined_page_no
            TrackDiagram_Input_Filename = Combined_File_Name

            ' move the elements
            For i = 1 To E_TOP
                el = Elements(i)
                If el.typ_s = "pfeil" Then
                    If i = E_TOP Then
                        E_TOP = E_TOP - 1
                        Exit For
                    Else
                        Elements(i) = Elements(E_TOP)
                        E_TOP = E_TOP - 1
                        el = Elements(i)
                    End If
                End If
                    el.row_no = el.row_no + rm(1)
                el.col_no = el.col_no + cm(1)
                s = Form1.DecIntToHexStr(combined_page_no) & Form1.DecIntToHexStr(el.row_no) & Form1.DecIntToHexStr(el.col_no)
                el.id_normalized = s
                el.id = Form1.CS2_id_create(combined_page_no, el.row_no, el.col_no)
                Elements(i) = el
            Next

            ' save the new file
            If Not Form1.GleisbildSave() Then
                ' switch back to the original file and reload it to get the page number in order
                OperationFailed("New page was not created due to error when writing the copy to disk", 1)
            Else
                ' Update the page file
                PFH.PageAdd(Combined_File_Name)
                PFH.PageFileSave()
            End If

            ' process the second page
            Form1.GleisbildLoad(Combine_Files(2))
            'fix the page number
            Form1.ElementsRepaginate(combined_page_no)
            ' move the elements and remove the pfeil
            For i = 1 To E_TOP
                el = Elements(i)
                If el.typ_s = "pfeil" Then
                    If i = E_TOP Then
                        E_TOP = E_TOP - 1
                        Exit For
                    Else
                        Elements(i) = Elements(E_TOP)
                        E_TOP = E_TOP - 1
                        el = Elements(i)
                    End If
                End If
                el.row_no = el.row_no + rm(2)
                el.col_no = el.col_no + cm(2)
                s = Form1.DecIntToHexStr(combined_page_no) & Form1.DecIntToHexStr(el.row_no) & Form1.DecIntToHexStr(el.col_no)
                el.id_normalized = s
                el.id = Form1.CS2_id_create(combined_page_no, el.row_no, el.col_no)
                Elements(i) = el
            Next

            ' Append to the previous file
            If Not Form1.GleisbildWriteAppendToFile(Combined_File_Name) Then
                ' switch back to the original file and reload it to get the page number in order
                OperationFailed("New page was only created in part due to some error when writing to disk", 1)
            Else
                MODE = "View"
                Form1.GleisbildLoad(Combined_File_Name)
                Form1.ElementsSort()
                Form1.GleisbildSave() ' save the sorted file
                Form1.GleisbildDisplay()
                Form1.AdjustPage()
            End If
        End Sub
        Private Function LocateJoiningPoint(i As Integer) As Integer
            ' Checks that one and only one joining point (pfeil element) is present, and returns the index into Elements of said element. Returns -1 if 
            ' no joining point and -2 if multiple joining points
            Dim e As Integer, n As Integer, e_joining_point As Integer
            Form1.GleisbildLoad(Combine_Files(i))
            For e = 1 To E_TOP
                If Elements(e).typ_s = "pfeil" Then
                    e_joining_point = e
                    n = n + 1
                End If
            Next
            Select Case n
                Case 0 : Return -1
                Case 1 : Return e_joining_point
                Case Else : Return -2
            End Select
        End Function
        Private Sub OperationFailed(msg As String, indx As Integer)
            If msg <> "" Then MsgBox(msg)
            MODE = "Wiev"
            Form1.GleisbildLoad(Combine_Files(indx))
            Form1.GleisbildDisplay()
            Form1.AdjustPage()
        End Sub
    End Class
    ReadOnly RH As New RouteHandler
    Shared ReadOnly PFH As New PageFileHandler
    ReadOnly CPH As New CombinePagesHandler

    Private Sub AdjustPage()
        Static not_first_call As Boolean = False
        Dim delta As Integer
        Dim W As Integer, H As Integer, H_eff As Integer
        delta = 9

        W = Me.Width
        H = Me.Height

        If W = 0 And H = 0 Then ' Form has been minimized
            Exit Sub
        End If

        If Me.WindowState = 2 Then ' maximized, we do not allow since it reveals "secret" objects which may confuse the user
            Me.WindowState = 0
            W = W_MIN
            Me.Width = W
            H = H_MIN
            Me.Height = H
        Else
            If W < W_MIN Or Not Developer_Environment Then
                W = W_MIN
                Me.Width = W
            End If
            If H < H_MIN Then
                H = H_MIN
                Me.Height = H
            End If
        End If

        Message(1, "Adjust page, mode=" & MODE & ", E_Top=" & Format(E_TOP))
        ' If Developer_Environment Then Message("me.width=" & Str(Me.Width) & ", me.height=" & Str(Me.Height))

        H_eff = H - 3 * delta ' adjusting for the menu bar

        ' Only two UI object have to be resized when the window changes: lstMessage and lstMasterFile
        ' Therefore all other size adjustments and posiitoning are only done once.
        ' Object visibility is updated on each call depending on the valye of Mode

        If Not not_first_call Then
            not_first_call = True
            grpLayout.Left = delta
            grpLayout.Top = H - H_eff
            grpLayout.Height = (ROW_MAX + 1) * 30 + lblC00.Height + 2 * delta
            grpLayout.Width = (COL_MAX + 1) * 30 + 16 + delta

            lstMessage.Top = grpLayout.Top + grpLayout.Height + delta
            lstMessage.Width = grpLayout.Width - grpMoveWindow.Width - delta
            grpMoveWindow.Top = lstMessage.Top
            grpMoveWindow.Left = lstMessage.Left + lstMessage.Width + delta

            grpIcons.Top = grpLayout.Top

            lstProperties.Top = H - H_eff
            lstProperties.Left = grpLayout.Left + grpLayout.Width + delta
            lstProperties.Height = grpLayout.Height * 0.4
            lstProperties.Width = grpIcons.Width

            grpEditActions.Top = grpIcons.Top + grpIcons.Height
            grpEditActions.Left = grpIcons.Left
            grpEditActions.Width = grpIcons.Width

            cmdAddressEdit.Width = grpEditActions.Width - 2 * cmdAddressEdit.Left
            cmdTextEdit.Width = cmdAddressEdit.Width
            btnDisplayRefreshEdit.Width = cmdAddressEdit.Width

            grpActions.Top = grpEditActions.Top + grpEditActions.Height
            grpActions.Height = 3 * cmdCommit.Height + 2 * delta
            grpActions.Left = lstProperties.Left
            grpActions.Width = lstProperties.Width

            ' lstInput visibility is set in tools menu
            lstInput.Top = lstProperties.Top + lstProperties.Height + delta
            lstInput.Height = grpActions.Top - (lstProperties.Top + lstProperties.Height) - delta
            lstInput.Left = lstProperties.Left
            lstInput.Width = lstProperties.Width

            lstOutput.Top = lstInput.Top
            lstOutput.Height = lstInput.Height
            lstOutput.Left = lstInput.Left - lstOutput.Width - delta

            cmdCommit.Top = cmdEdit.Top
            cmdCommit.Left = cmdEdit.Left
            cmdCancel.Top = cmdCommit.Top + cmdCommit.Height
            cmdCancel.Left = cmdCommit.Left ' + cmdCommit.Width + delta

            If Not STAND_ALONE Then
                cmdCommit.Width = 135
                cmdCommit.Text = "Commit and Return to TC"
                cmdCancel.Width = 135
                cmdCancel.Text = "Cancel and Return to TC"
            End If

            cmdReturnToTC.Left = cmdCommit.Left
            cmdReturnToTC.Top = cmdEdit.Top + cmdEdit.Height  ' + delta

            cmdRoutesView.Left = cmdEdit.Left
            cmdRoutesView.Top = cmdReturnToTC.Top + cmdReturnToTC.Height ' + delta ' cmdEdit.Top

            cmdManagePages.Left = cmdCommit.Left
            cmdManagePages.Top = cmdReturnToTC.Top + cmdReturnToTC.Height
            cboRouteNames.Left = lstProperties.Left
            cboRouteNames.Top = lstProperties.Top
            cboRouteNames.Width = lstProperties.Width

            lstRoutes.Left = lstProperties.Left
            lstRoutes.Top = lstInput.Top
            lstRoutes.Height = lstInput.Height - delta
            lstRoutes.Width = lstProperties.Width

            lstMasterFile.Width = 1.5 * lstRoutes.Width ' only used for debugging. Not visible to user
            lstMasterFile.Width = 1.5 * lstRoutes.Width
            grpLayoutPages.Top = lstProperties.Top + lstProperties.Height
            grpLayoutPages.Left = lstProperties.Left
            grpLayoutPages.Width = lstProperties.Width
            grpLayoutPages.Height = grpActions.Top - grpLayoutPages.Top
            lstLayoutPages.Width = grpLayoutPages.Width - 2 * lstLayoutPages.Left
            lstLayoutPages.Height = 0.75 * lstProperties.Height
            lblLayoutPages.Top = lstLayoutPages.Top + lstLayoutPages.Height
            btnAddLayoutPage.Top = lblLayoutPages.Top + lblLayoutPages.Height
            btnAddLayoutPage.Width = grpLayoutPages.Width - 2 * btnAddLayoutPage.Left
            btnDuplicatePage.Top = btnAddLayoutPage.Top + btnAddLayoutPage.Height
            btnDuplicatePage.Width = btnAddLayoutPage.Width
            btnCombinePages.Top = btnDuplicatePage.Top + btnDuplicatePage.Height
            btnCombinePages.Width = btnAddLayoutPage.Width
            btnRenameLayoutPage.Top = btnCombinePages.Top + btnCombinePages.Height
            btnRenameLayoutPage.Width = btnAddLayoutPage.Width
            btnRemoveLayoutPage.Top = btnRenameLayoutPage.Top + btnRenameLayoutPage.Height
            btnRemoveLayoutPage.Width = btnAddLayoutPage.Width
            cmdRevealAddressesPage.Top = grpLayoutPages.Height - cmdRevealAddressesPage.Height - cmdRevealTextPage.Height - delta
            cmdRevealAddressesPage.Width = btnAddLayoutPage.Width
            cmdRevealTextPage.Top = cmdRevealAddressesPage.Top + cmdRevealAddressesPage.Height
            cmdRevealTextPage.Left = cmdRevealAddressesPage.Left
            cmdRevealTextPage.Width = cmdRevealAddressesPage.Width
            txtCombine.Top = btnAddLayoutPage.Top
            txtCombine.Left = btnAddLayoutPage.Left
            txtCombine.Width = btnAddLayoutPage.Width
            txtCombine.Height = cmdRevealTextPage.Top - btnAddLayoutPage.Top + cmdRevealTextPage.Height

        End If

        lstMessage.Height = H_eff - grpLayout.Height - 5 * delta
        lstMasterFile.Height = (lstMessage.Top + lstMessage.Height) - lstMasterFile.Top

        grpActions.Visible = FILE_LOADED
        grpMoveWindow.Visible = FILE_LOADED

        lstProperties.Visible = (MODE = "View")
        cmdEdit.Visible = (MODE = "View")
        cmdReturnToTC.Visible = (MODE = "View") And Not STAND_ALONE
        cmdRoutesView.Visible = (MODE = "View" And ROUTE_VIEW_ENABLED)
        grpLayoutPages.Visible = ((MODE = "View") Or (MODE = "Join")) And FILE_LOADED And Not lstInput.Visible

        grpIcons.Visible = (MODE = "Edit")
        grpEditActions.Visible = (MODE = "Edit")
        cmdCommit.Visible = (MODE = "Edit")
        cmdManagePages.Visible = (MODE = "Edit") And Not STAND_ALONE

        cmdCancel.Visible = (MODE = "Edit" Or MODE = "Routes" Or MODE = "Join")

        cboRouteNames.Visible = (MODE = "Routes")
        lstRoutes.Visible = (MODE = "Routes")

        lblLayoutPages.Visible = (MODE <> "Join")
        btnAddLayoutPage.Visible = (MODE <> "Join")
        btnDuplicatePage.Visible = (MODE <> "Join")
        btnCombinePages.Visible = (MODE <> "Join") And Not Flat_Structure  ' Join requires a layout page file
        btnRenameLayoutPage.Visible = (MODE <> "Join")
        btnRemoveLayoutPage.Visible = (MODE <> "Join")
        cmdRevealAddressesPage.Visible = (MODE <> "Join")
        cmdRevealTextPage.Visible = (MODE <> "Join")

        ' lstInput visibility is set in tools menu, not here
        lstOutput.Visible = False

        lstMasterFile.Visible = Developer_Environment
        ShowInputFileToolStripMenuItem.Visible = Developer_Environment Or Verbose
        ShowOutputFileToolStripMenuItem.Visible = Developer_Environment
        ConsistencyCheckToolStripMenuItem.Visible = Developer_Environment
        DumpElementsToolStripMenuItem.Visible = Developer_Environment
        SetTestLevelToolStripMenuItem.Visible = Developer_Environment
        StatusVariablesToolStripMenuItem.Visible = Developer_Environment
        TurnOffDeveloperModeToolStripMenuItem.Visible = Developer_Environment

        If Not STAND_ALONE Then VisibilitySetCommandlineMode(CMDLINE_MODE)
    End Sub
    Private Sub AxesDraw()
        lblC00.Text = Format(Col_Offset)
        lblC01.Text = Format(Col_Offset + 1)
        lblC02.Text = Format(Col_Offset + 2)
        lblC03.Text = Format(Col_Offset + 3)
        lblC04.Text = Format(Col_Offset + 4)
        lblC05.Text = Format(Col_Offset + 5)
        lblC06.Text = Format(Col_Offset + 6)
        lblC07.Text = Format(Col_Offset + 7)
        lblC08.Text = Format(Col_Offset + 8)
        lblC09.Text = Format(Col_Offset + 9)
        lblC10.Text = Format(Col_Offset + 10)
        lblC11.Text = Format(Col_Offset + 11)
        lblC12.Text = Format(Col_Offset + 12)
        lblC13.Text = Format(Col_Offset + 13)
        lblC14.Text = Format(Col_Offset + 14)
        lblC15.Text = Format(Col_Offset + 15)
        lblC16.Text = Format(Col_Offset + 16)
        lblC17.Text = Format(Col_Offset + 17)
        lblC18.Text = Format(Col_Offset + 18)
        lblC19.Text = Format(Col_Offset + 19)
        lblC20.Text = Format(Col_Offset + 20)
        lblC21.Text = Format(Col_Offset + 21)
        lblC22.Text = Format(Col_Offset + 22)
        lblC23.Text = Format(Col_Offset + 23)
        lblC24.Text = Format(Col_Offset + 24)
        lblC25.Text = Format(Col_Offset + 25)
        lblC26.Text = Format(Col_Offset + 26)
        lblC27.Text = Format(Col_Offset + 27)
        lblC28.Text = Format(Col_Offset + 28)
        lblC29.Text = Format(Col_Offset + COL_MAX)

        lblR00.Text = Format(Row_Offset)
        lblR01.Text = Format(Row_Offset + 1)
        lblR02.Text = Format(Row_Offset + 2)
        lblR03.Text = Format(Row_Offset + 3)
        lblR04.Text = Format(Row_Offset + 4)
        lblR05.Text = Format(Row_Offset + 5)
        lblR06.Text = Format(Row_Offset + 6)
        lblR07.Text = Format(Row_Offset + 7)
        lblR08.Text = Format(Row_Offset + 8)
        lblR09.Text = Format(Row_Offset + 9)
        lblR10.Text = Format(Row_Offset + 10)
        lblR11.Text = Format(Row_Offset + 11)
        lblR12.Text = Format(Row_Offset + 12)
        lblR13.Text = Format(Row_Offset + 13)
        lblR14.Text = Format(Row_Offset + 14)
        lblR15.Text = Format(Row_Offset + 15)
        lblR16.Text = Format(Row_Offset + ROW_MAX)
    End Sub

    Private Sub CloseShop()
        Dim ws As Integer, answer As Integer = vbNo
        frmHelp.Hide()
        frmHelp.Dispose()
        If CHANGED Then
            ' Save changes unconditionally unless we're in stand alone mode
            If STAND_ALONE Then
                answer = MsgBox("Track diagram has been modified but not saved. Save before exit?", vbYesNo)
            End If
            If answer = vbYes Or Not STAND_ALONE Then
                If Not GleisbildSave() Then
                    MsgBox("Page could not be saved: " & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)))
                    If STAND_ALONE Then Message("to folder " & FilenamePath(TrackDiagram_Input_Filename))
                End If
            End If
        End If

        ws = Me.WindowState
        env1.SetEnv("WindowState", Str(ws))
        If ws = 0 Then ' not maximized
            env1.SetEnv("Top", Str(Me.Top))
            env1.SetEnv("Left", Str(Me.Left))
            env1.SetEnv("Width", Str(Me.Width))
            env1.SetEnv("Height", Str(Me.Height))
        End If

        env1.SetEnv("Use_English_Terms", Str(UseEnglishTermsForElementsToolStripMenuItem.Checked))
        env1.SetEnv("Verbose", Str(Verbose))

        env1.SetEnv("Test_Level", Str(Test_Level))
        env1.SaveEnv()
        env1.CloseEnv()
        End
    End Sub
    Private Sub Initialize()
        drawFormatVertical.FormatFlags = StringFormatFlags.DirectionVertical

        Empty_Element.id = ""
        Empty_Element.typ_s = ""
        Empty_Element.drehung_i = -1 ' invalid value
        Empty_Element.artikel_i = 0
        Empty_Element.deviceId_s = ""
        Empty_Element.zustand_s = ""
        Empty_Element.text = ""
        Empty_Element.id_normalized = ""
        Empty_Element.row_no = 0
        Empty_Element.col_no = 0
        Empty_Element.my_index = 0
    End Sub
    Public Function InputPromptOK(ByRef result As String, prompt As String, default_text As String) As Boolean
        ' A better InputBox which destinguishes between entering empty text and hitting the Cancel button
        ' User hits OK: Inserts the new text string into 'result', can be empty. Trailing blanks will be removed. Function returns True
        ' User hits Cancel: Leaves 'result' unchanged. Function returns False

        ' Show the custom input dialog
        Dim inputForm As New CustomInputForm(prompt, default_text)
        Dim dialogResult As DialogResult = inputForm.ShowDialog()

        If dialogResult = DialogResult.Cancel Then
            ' User pressed Cancel
            Return False
        Else
            result = Strings.RTrim(inputForm.InputTextBox.Text)
            Return True
        End If
    End Function
    Private Sub GleisbildLoad(fn As String)
        ' Load the layout in 'fn' the input window. If 'fn' is blank (or doen't exist), prompt for the file name
        Dim answer As DialogResult = vbYes, fn2 As String, reopen As Boolean = False
        Dim buffer As String, fno As Integer, param_kind As String
        Dim param_val As String = "", param_val_normalized As String = ""
        Dim warnings As Boolean = False
        If CHANGED Then
            Dim layout_name As String
            layout_name = FilenameProper(TrackDiagram_Input_Filename)

            If STAND_ALONE Then answer = MsgBox("Track diagram """ & layout_name & """ has been modified but not saved. Should changes be saved?", vbYesNo)
            If answer = vbYes Or Not STAND_ALONE Then
                If GleisbildSave() Then
                    If STAND_ALONE Then Message("Track diagram saved in: " & TrackDiagram_Input_Filename)
                Else
                    MsgBox("Page could not be saved: " & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)))
                End If
            End If
        End If

        If MODE <> "Routes" And MODE <> "Join" Then
            MODE = "View" ' unless in Join or Routes mode, we switch to View mode
        End If

        ROUTE_VIEW_ENABLED = False
        cmdRoutesView.Visible = False

        Page_No = 0
        If Row_Offset <> 0 Or Col_Offset <> 0 Then
            WindowMove(-Row_Offset, -Col_Offset)
            Row_Offset = 0
            Col_Offset = 0
        End If

        GleisbildGridSet(False)
        GleisBildDisplayClear()

        lstProperties.Items.Clear()
        lstProperties.Items.Add("Click an element to")
        lstProperties.Items.Add("see its properties")
        lstProperties.Items.Add("here")
        lstInput.Items.Clear()
        lstOutput.Visible = False

        TrackDiagram_Input_Filename = ""

        Active_Tool = ""

        Header_Top = 0
        E_TOP = 0
        CHANGED = False
        FILE_LOADED = False

        Me.Text = "Track Diagram Editor  .NET" & "  [" & Version & "]"

        If fn = "" Then
            fn2 = "*.cs2"
        ElseIf FileExists(fn) Then
            fn2 = fn
            reopen = True
            answer = DialogResult.OK
        Else
            fn2 = "*.cs2"
        End If

        OpenFileDialog1.FileName = fn2
        If Not reopen Then
            OpenFileDialog1.InitialDirectory = TrackDiagram_Directory
            OpenFileDialog1.Filter = "CS2 layout files (*.cs2)|*.cs2"
            answer = OpenFileDialog1.ShowDialog()
        End If

        If answer = DialogResult.OK Then
            TrackDiagram_Input_Filename = OpenFileDialog1.FileName
            ' In Pages and Routes mode we're only allowed to open files in the current directory, i.e. the directory where the currently open track diagram resides
            If MODE = "Routes" Then
                If FilenamePath(TrackDiagram_Input_Filename) <> TrackDiagram_Directory Then
                    MsgBox("When managing pages or viewing routes you can only open files in the current folder: " & TrackDiagram_Directory)
                    TrackDiagram_Input_Filename = ""
                    Me.Text = "Track Diagram Editor .NET  [" & Version & "]"
                    Exit Sub
                End If
            End If

            env1.SetEnv("TrackDiagram_Input_Filename", TrackDiagram_Input_Filename)
            TrackDiagram_Directory = FilenamePath(TrackDiagram_Input_Filename)
            env1.SetEnv("TrackDiagram_Directory", TrackDiagram_Directory)
            Input_Drive = Mid(TrackDiagram_Directory, 1, 1)

            ' If the trackdiagram does not reside in a folder named "gleisbilder" we will operate in flat directory structure mode
            Message(3, "TrackDiagram_Directory=" & TrackDiagram_Directory)
            Flat_Structure = (InStr(TrackDiagram_Directory, "\gleisbilder") = 0)
            If Flat_Structure Then
                MasterFile_Directory = TrackDiagram_Directory
            Else
                MasterFile_Directory = FilePathUp(TrackDiagram_Directory)
                ' If we're already at the top we must switch to flat structure
                If MasterFile_Directory = "" Then
                    MasterFile_Directory = TrackDiagram_Directory
                    Flat_Structure = True
                End If
            End If
            Message(3, "MasterFile_Directory=  " & MasterFile_Directory)
            Message(3, "Flat_structure=" & Format(Flat_Structure))

            Me.Text = "Track Diagram Editor  .NET  [" & Version & "]   " & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)) &
                "   " & FilePathUp(MasterFile_Directory)

            ' Read the file and load the image file names.
            lstInput.Items.Clear()
            E_TOP = 0
            Header_Top = 0

            fno = FreeFile()

            ' Process the header lines
            FileOpen(fno, TrackDiagram_Input_Filename, OpenMode.Input)
            While Not (EOF(fno))
                If Header_Top = HEADER_MAX Then
                    MsgBox("Too many header lines in input file")
                    FileClose(fno)
                    Exit Sub
                End If
                buffer = LineInput(fno)
                lstInput.Items.Add(buffer)
                If buffer = "element" Then Exit While
                Header_Top = Header_Top + 1
                Header_Line(Header_Top) = buffer
                Select Case Header_Top
                    Case 1 : If Header_Line(1) <> "[gleisbildseite]" Then
                            MsgBox("Input file is not a valid track diagram file.")
                            FileClose(fno)
                            TrackDiagram_Input_Filename = ""
                            Exit Sub
                        End If
                    Case Else : If Strings.Left(buffer, 4) = "page" Then
                            Page_No = Val(ExtractField(buffer, 2, "="))
                        End If
                End Select
            End While

            If Flat_Structure Then
                Message("Loading track diagram: " & TrackDiagram_Input_Filename) ' page number not relevant information
            Else
                If Verbose Or Test_Level >= 1 Then
                    Message("Loading track diagram for page" & Str(Page_No) & ": """ & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)) & """")
                    Message("  from file: " & TrackDiagram_Input_Filename)
                End If
            End If

            ' Process the layout elements. Buffer contains "element"
            While Not EOF(fno)
                ' buffer should contains "element" at the beginning of each iteration.
                E_TOP = E_TOP + 1
                Elements(E_TOP).id = ""
                Elements(E_TOP).typ_s = ""
                Elements(E_TOP).drehung_i = 0 ' CS2 omits this value when 0
                Elements(E_TOP).artikel_i = 0 ' As far as I can see all elements must have a value for artikel and it must be <>0.
                ' -1 if the element has no digital address. We issue a warning if the artikel value is absent and then insert -1
                Elements(E_TOP).deviceId_s = "" ' CS2 omits this value when 0
                Elements(E_TOP).zustand_s = "" ' CS2 omits this value when 0
                Elements(E_TOP).text = ""
                Elements(E_TOP).id_normalized = "000000" ' CS2 omits the id for the element in cell (0,0) of page 0. A bit too clever for their own good.
                '                ELEMENTS(e_top).id = "0x0" ' We set this as default rather than empty, cf. above. We remove it when writing the file back
                Elements(E_TOP).row_no = 0
                Elements(E_TOP).col_no = 0
                Elements(E_TOP).my_index = E_TOP
                While Not EOF(fno)
                    buffer = LineInput(fno)
                    lstInput.Items.Add(buffer)
                    buffer = Trim(buffer)
                    If buffer = "element" Then Exit While
                    param_kind = Trim(ExtractField(buffer, 1, "="))
                    '                    param_val = Trim(ExtractField(buffer, 2, "=")) 
                    param_val = RTrim(ExtractField(buffer, 2, "=")) ' trim changed to rtrim 231022
                    Select Case param_kind
                        Case ".id" : Elements(E_TOP).id = param_val
                            param_val_normalized = IdNormalize(param_val)
                            Elements(E_TOP).id_normalized = param_val_normalized ' always 6 characters, 3 groups of two characters (hex numbers)
                            Elements(E_TOP).row_no = DecodeRowNo(param_val_normalized)
                            Elements(E_TOP).col_no = DecodeColNo(param_val_normalized)
                            'Message(param_val & " -> " & param_val_normalized & ": " & Str(ELEMENTS(e_top).page_no) & "," & ELEMENTS(e_top).col_no & "," & ELEMENTS(e_top).row_no)
                        Case ".typ" : Elements(E_TOP).typ_s = param_val
                        Case ".drehung" : Elements(E_TOP).drehung_i = Val(param_val)
                        Case ".artikel" : Elements(E_TOP).artikel_i = Val(param_val)
                        Case ".deviceId" : Elements(E_TOP).deviceId_s = param_val
                        Case ".zustand" : Elements(E_TOP).zustand_s = param_val
                        Case ".text" : Elements(E_TOP).text = param_val
                        Case Else
                            If buffer <> "" Then
                                Message("WARNING: Unknown parameter kind: """ & param_kind & """. This parameter will be discarded. Offending input line:")
                                Message("""" & buffer & """")
                                warnings = True
                            Else
                                ' ignore empty lines
                            End If
                    End Select
                End While
                If Elements(E_TOP).artikel_i = 0 And Elements(E_TOP).typ_s <> "pfeil" Then
                    ' artikel is missing. 0 is a valid address for pfeil, though, så we need the special case in the expression above.
                    ' We omit the warning unless verbose (test_level > 0) for elements that do not need an address per se.
                    If Test_Level > 0 Then
                        Message("WARNING: Element in cell " & Elements(E_TOP).id & " (" & Format(Elements(E_TOP).row_no) & "," & Format(Elements(E_TOP).col_no) & ") " & Elements(E_TOP).typ_s _
                            & " has no address. It has been set to -1. (m17)")
                        warnings = True
                    ElseIf Elements(E_TOP).typ_s <> "gerade" And Elements(E_TOP).typ_s <> "bogen" And Elements(E_TOP).typ_s <> "prellbock" _
                           And Strings.Left(Elements(E_TOP).typ_s, 6) <> "custom" Then
                        Message("WARNING: Element in cell " & Elements(E_TOP).id & " (" & Format(Elements(E_TOP).row_no) & "," & Format(Elements(E_TOP).col_no) & ") " & Elements(E_TOP).typ_s _
                            & " has no address. It has been set to -1. (m18)")
                        warnings = True
                    End If
                    Elements(E_TOP).artikel_i = -1
                    CHANGED = True
                End If
            End While
            FileClose(fno)
            Message(2, "  No. of lines in input file:" & Str(lstInput.Items.Count))
            If Verbose Or Test_Level >= 1 Then
                Message("  No. of elements: " & Str(E_TOP))
            End If

            If Not GleisbildConsistencyCheck() Then
                warnings = True
                If Not GleisbildSave() Then
                    Message("Could not save the layout to " & TrackDiagram_Input_Filename & ". (m19)")
                    CHANGED = True
                End If
            End If

            If warnings Then
                MsgBox("One or more warnings were issued. Check the message box.")
            End If

            ROUTE_VIEW_ENABLED = FileExists(MasterFile_Directory & "fahrstrassen.cs2")
            cmdRoutesView.Visible = (MODE = "View" And ROUTE_VIEW_ENABLED)

            FILE_LOADED = True
            AdjustPage()
            PFH.PageFileLoad(MasterFile_Directory)
            PFH.PagesPopulate()
        End If
    End Sub

    Private Function GleisBildImportExternalFile() As Boolean
        ' Creates a new page and imports an external file into said page.
        ' The name of the new page will be that of the file's name. If conflict with existing page names, a new name will be created.
        ' The page file is updated.
        ' Returns true iff all went well.
        Dim external_filename As String, new_pagename As String, new_pagename_rev As String, new_filename As String, new_pagenumber As Integer
        Dim counter As Integer = 0, answer As DialogResult
        If CHANGED Then
            Dim layout_name As String
            layout_name = FilenameProper(TrackDiagram_Input_Filename)

            If STAND_ALONE Then answer = MsgBox("Track diagram """ & layout_name & """ has been modified but not saved. Should changes be saved?", vbYesNo)
            If answer = vbYes Or Not STAND_ALONE Then
                If GleisbildSave() Then
                    If STAND_ALONE Then Message("Track diagram saved in: " & TrackDiagram_Input_Filename)
                Else
                    MsgBox("Page could not be saved: " & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)))
                    If STAND_ALONE Then Message("to folder " & FilenamePath(TrackDiagram_Input_Filename))
                End If
            End If
        End If

        MODE = "View"

        ROUTE_VIEW_ENABLED = False
        cmdRoutesView.Visible = False

        Row_Offset = 0
        Col_Offset = 0

        OpenFileDialog1.FileName = "*.cs2"
        OpenFileDialog1.InitialDirectory = TrackDiagram_Directory
        OpenFileDialog1.Filter = "CS2 layout files (*.cs2)|*.cs2"
        answer = OpenFileDialog1.ShowDialog()

        If answer = DialogResult.OK Then
            external_filename = OpenFileDialog1.FileName
            new_pagename = FilenameNoExtension(FilenameExtract(external_filename))
            If new_pagename = "gleisbild" Then
                MsgBox("A ""gleisbild.cs2"" file cannot be used as a track diagram")
                Return False
            End If
            If new_pagename = "fahrstrassen" Then
                MsgBox("A ""fahrstrassen.cs2"" file cannot be used as a track diagram")
                Return False
            End If

            new_pagename_rev = new_pagename
            While PFH.LayoutPageNumberGet(new_pagename_rev) >= 0  ' page name already taken
                counter = counter + 1
                new_pagename_rev = new_pagename & Format(counter)
            End While
            new_filename = MasterFile_Directory & "gleisbilder\" & new_pagename_rev & ".cs2"
            If external_filename <> new_filename Then
                Try
                    FileCopy(external_filename, new_filename)
                Catch
                    MsgBox("The file could not be copied to your layout. New page not created")
                    Return False
                End Try
            Else
                ' The file already resides in the correct directory
            End If
            new_pagenumber = PFH.NextAvailablePageNumber
            Page_No = new_pagenumber
            PFH.PageAdd(new_filename)

            ' the new file probably has the wrong page number. Load it and change th number
            GleisBildDisplayClear()
            GleisbildLoad(new_filename)
            ElementsRepaginate(new_pagenumber)
            Page_No = new_pagenumber
            PFH.PagesPopulate()
            GleisbildSave()
            GleisbildDisplay()
            Return True
        End If
        Return False
    End Function
    Private Function GleisbildSave() As Boolean
        ' Saves the current track diagram with its current name. Returns True unless the operation falied.
        Dim ok As Boolean

        If E_TOP = 0 Then Return True ' nothing to save

        If LCase(FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename))) = "gleisbild" Then
            ' Should never happen. If we get here there is an error in the program logic
            MsgBox("Saving a layout in the ""gleisbild"" layout page file is not allowed. Layout not saved.")
            Return False
        End If

        If Test_Level >= 3 Then lstOutput.Visible = True

        ok = GleisbildWriteFile(TrackDiagram_Input_Filename)
        If Not ok Then
            MsgBox("Cannot write to file " & TrackDiagram_Input_Filename)
        Else
            If STAND_ALONE Then Message("File saved in: " & TrackDiagram_Input_Filename)
        End If
        CHANGED = False
        Return ok
    End Function
    Private Function GleisbildSaveCopyAs() As Boolean
        ' Saves the current track diagram as a copy under a new name.
        ' Returns True unless the file couldn't be saved.
        ' The page file  "gleisbild.cs2" will NOT be updated. User must use the page management actions to add the layout to the page if so desired.

        Dim answer As DialogResult, ok As Boolean
        Dim trackDiagram_copy_filename As String

        If E_TOP = 0 Then Return True ' nothing to save

        SaveFileDialog1.InitialDirectory = TrackDiagram_Directory ' SaveFileDialog messes with me in case of an existing file
        trackDiagram_copy_filename = "*.cs2"
        SaveFileDialog1.FileName = trackDiagram_copy_filename
        SaveFileDialog1.Filter = "CS2 layout files (*.cs2)|*.cs2"

        answer = SaveFileDialog1.ShowDialog()

        If answer = DialogResult.OK Then
            trackDiagram_copy_filename = SaveFileDialog1.FileName

            If FilenameExtension(trackDiagram_copy_filename) = "" Then
                trackDiagram_copy_filename = trackDiagram_copy_filename & ".cs2"
            End If

            If LCase(FilenameNoExtension(FilenameExtract(trackDiagram_copy_filename))) = "gleisbild" Then
                MsgBox("Saving a track diagram in the ""gleisbild"" layout map is not allowed. File not saved.")
                Return False
            End If

            If FilenameNoExtension(trackDiagram_copy_filename) = FilenameNoExtension(TrackDiagram_Input_Filename) Then
                MsgBox("Please enter a name that is not in use. File not saved.")
                Return False
            End If

            If Not Flat_Structure AndAlso PFH.PageFileLoad(MasterFile_Directory) Then
                If PFH.LayoutPageNumberGet(trackDiagram_copy_filename) >= 0 Then
                    MsgBox("Please enter a name that is not in use. File not saved.")
                    Return False
                End If
            End If

            ok = GleisbildWriteFile(trackDiagram_copy_filename)
            If ok Then
                Message("Copy saved in: " & trackDiagram_copy_filename)
                If Not Flat_Structure Then Message("The page layout file was not updated. This must be done manually if so desired.")
            Else
                Message("Could not save copy to file " & trackDiagram_copy_filename)
            End If
            Return ok
        Else
            Return False
        End If
    End Function

    Private Function GleisbildWriteFile(filename As String) As Boolean
        ' Writes the currently loaded layout to file. The file is closed upon return.
        ' Returns True unless the operation failed.
        Dim buffer As String, fno As Integer

        lstOutput.Items.Clear()
        If Test_Level >= 3 Then lstOutput.Visible = True

        ElementsSort()

        On Error GoTo write_error
        fno = FreeFile()
        FileOpen(fno, filename, OpenMode.Output)
        For i = 1 To Header_Top
            ' Second line must be "page=", unless page number is 0. Since we may have repaginated we need to handle this explicitely
            If i = 2 Then
                If Page_No <> 0 Then
                    lstOutput.Items.Add("page=" & Format(Page_No))
                    PrintLine(fno, "page=" & Format(Page_No))
                End If
                If Strings.Left(Header_Line(2), 4) <> "page" Then
                    lstOutput.Items.Add(Header_Line(i))
                    PrintLine(fno, Header_Line(i))
                End If
            Else
                lstOutput.Items.Add(Header_Line(i))
                PrintLine(fno, Header_Line(i))
            End If
        Next i
        For i = 1 To E_TOP
            buffer = "element"
            lstOutput.Items.Add(buffer)
            PrintLine(fno, buffer)
            If Elements(i).id <> "" Then
                buffer = " .id=" & Elements(i).id
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            If Elements(i).typ_s <> "" Then
                buffer = " .typ=" & Elements(i).typ_s
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            If Elements(i).drehung_i > 0 Then
                buffer = " .drehung=" & Format(Elements(i).drehung_i)
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            buffer = " .artikel=" & Format(Elements(i).artikel_i) ' artikel is one of the few parameters which all elements have 
            lstOutput.Items.Add(buffer)
            PrintLine(fno, buffer)
            If Elements(i).deviceId_s <> "" Then
                buffer = " .deviceId=" & Format(Elements(i).deviceId_s)
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            If Elements(i).zustand_s <> "" Then
                buffer = " .zustand=" & Format(Elements(i).zustand_s)
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            If Elements(i).text <> "" Then
                buffer = " .text=" & Elements(i).text
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
        Next i
        On Error GoTo 0
        FileClose(fno)
        Return True
write_error:
        On Error GoTo 0
        FileClose(fno)
        Return False
    End Function
    Private Function GleisbildWriteAppendToFile(filename As String) As Boolean
        ' Appends the elements of the currently loaded layout to the file 'filename'. The file is closed upon return.
        ' Returns True unless the operation failed.
        Dim buffer As String, fno As Integer

        lstOutput.Items.Clear()

        On Error GoTo write_error
        fno = FreeFile()
        FileOpen(fno, filename, OpenMode.Append)
        For i = 1 To E_TOP
            buffer = "element"
            lstOutput.Items.Add(buffer)
            PrintLine(fno, buffer)
            If Elements(i).id <> "" Then
                buffer = " .id=" & Elements(i).id
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            If Elements(i).typ_s <> "" Then
                buffer = " .typ=" & Elements(i).typ_s
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            If Elements(i).drehung_i > 0 Then
                buffer = " .drehung=" & Format(Elements(i).drehung_i)
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            buffer = " .artikel=" & Format(Elements(i).artikel_i) ' artikel is one of the few parameters which all elements have 
            lstOutput.Items.Add(buffer)
            PrintLine(fno, buffer)
            If Elements(i).deviceId_s <> "" Then
                buffer = " .deviceId=" & Format(Elements(i).deviceId_s)
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            If Elements(i).zustand_s <> "" Then
                buffer = " .zustand=" & Format(Elements(i).zustand_s)
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
            If Elements(i).text <> "" Then
                buffer = " .text=" & Elements(i).text
                lstOutput.Items.Add(buffer)
                PrintLine(fno, buffer)
            End If
        Next i
        On Error GoTo 0
        FileClose(fno)
        Return True
write_error:
        On Error GoTo 0
        FileClose(fno)
        Return False
    End Function
    Private Sub GleisbildCreateBackup(file_name As String)
        ' Saves the contents of lstInput (i.e. the original input file) as a backup file in a \backup subdirectory
        ' and with the extension .qs2 overwriting an existing file with that name, If any.
        ' Currently (version 2.1) only called when user is deleting a page (i.e. track diagram)
        Dim fno As Integer, backup_file_name As String, ul As Integer, i As Integer
        backup_file_name = FilenamePath(file_name) & "Backup\" & FilenameProper(file_name) & ".cs2"

        If E_TOP = 0 OrElse file_name = "" OrElse backup_file_name = file_name Then Exit Sub ' nothing or no need to save
        DirectoryCreate(FilenamePath(file_name) & "Backup\")

        On Error GoTo write_error
        fno = FreeFile()
        FileOpen(fno, backup_file_name, OpenMode.Output)
        ul = lstInput.Items.Count - 1
        For i = 0 To ul
            PrintLine(fno, lstInput.Items.Item(i))
        Next i
        On Error GoTo 0
        FileClose(fno)
        Message(1, "Backup created: " & backup_file_name)
        Exit Sub
write_error:
        On Error GoTo 0
        Message("Could not create backup file " & backup_file_name)
        FileClose(fno)
    End Sub

    Private Function GleisbildConsistencyCheck() As Boolean
        ' Returns True if the layout had no double occupancy in a cell.
        ' Returns False if warnings were issued and elements were removed
        Dim i As Integer, found_double As Boolean = False
        Dim el As ELEMENT, el_double As ELEMENT
        Dim typ_s As String, typ_s_double As String

        For i = E_TOP To 1 Step -1
            el = Elements(i)
            ' see if there is another element in this cell.
            ' Note that ElementGet2 searcher from i=1 and upwards while we here search the other direction
            el_double = ElementGet2(el.row_no, el.col_no)
            If el.my_index <> el_double.my_index Then ' double found
                ' Which one to delete? I'd say the one with the lowest index because that's the one that is hidden from sight on the display
                ' We must special case for text elements since they have no type
                found_double = True
                typ_s = Translate(el.typ_s)
                If typ_s = "" And el.text <> "" Then typ_s = "text-" & el.text
                typ_s_double = Translate(el_double.typ_s)
                If typ_s_double = "" And el_double.text <> "" Then typ_s_double = "text-" & el_double.text
                Message("WARNING: Double occupancy in (" & Format(el.row_no) & "," & Format(el.col_no) & "): """ & typ_s_double & """ and """ & typ_s & """" _
                        & ". Removed """ & typ_s_double & """")
                ElementRemove(el_double.my_index)
            End If
        Next
        If found_double Then
            CHANGED = True
            GleisBildDisplayClear()
            GleisbildDisplay()
            Return False
        Else
            Return True
        End If
    End Function
    Private Sub GleisbildDisplay()
        Dim i As Integer
        For i = 1 To E_TOP
            ElementDisplay(Elements(i))
        Next i
    End Sub

    Private Sub GleisBildDisplayClear()
        Dim r As Integer, c As Integer
        For r = 0 To ROW_MAX
            For c = 0 To COL_MAX
                If PicGet(r, c).Image IsNot Nothing Then
                    PicGet(r, c).Image.Dispose()
                    PicGet(r, c).Image = picIconEmpty.Image.Clone
                End If
                PicGet(r, c).Visible = True
            Next c
        Next r
    End Sub
    Private Sub GleisbildGridSet(set_to_visible As Boolean)
        Dim r As Integer, c As Integer
        If set_to_visible Then
            For r = 0 To ROW_MAX
                For c = 0 To COL_MAX
                    PicGet(r, c).BorderStyle = 1
                Next c
            Next r
        Else
            For r = 0 To ROW_MAX
                For c = 0 To COL_MAX
                    PicGet(r, c).BorderStyle = 0
                Next c
            Next r
        End If
    End Sub

    Private Function ElementGet(coord_str As String) As ELEMENT
        ' Given a row/col coordinate string like "0305", returns the corresponding element. Returns Empty_Elelemnt of no match
        Dim row_no As Integer, col_no As Integer
        row_no = Val(Strings.Left(coord_str, 2))
        col_no = Val(Strings.Mid(coord_str, 3, 2))
        Return ElementGet2(row_no, col_no)
    End Function
    Private Function ElementGet2(row_no As Integer, col_no As Integer) As ELEMENT
        ' Given a row and col, returns the corresponding element. Returns Empty_Elelemnt of no match
        Dim i As Integer
        Dim e As ELEMENT
        For i = 1 To E_TOP
            e = Elements(i)
            If e.row_no = row_no And e.col_no = col_no Then
                Return e
            End If
        Next
        Return Empty_Element
    End Function

    Private Function ElementGet3(kind As String, addr As Integer) As ELEMENT
        ' given a digital address >0 and the kind of element, returns the element with that address, or EmptyElement, if not found
        Dim i As Integer, el As ELEMENT
        Message(3, "Locating element by kind and address; kind=" & kind & ", addr=" & Str(addr))
        Select Case kind
            Case "..magnetartikel"
                For i = 1 To E_TOP Step 1
                    el = Elements(i)
                    If Strings.Left(el.typ_s, 6) = "signal" OrElse (InStr(el.typ_s, "weiche") > 0) Then
                        If el.artikel_i = 2 * addr Then Return Elements(i)
                    End If
                Next i
            Case ".s88", ".s88Flag"
                For i = 1 To E_TOP Step 1
                    el = Elements(i)
                    If Strings.Left(el.typ_s, 3) = "s88" Then
                        If el.artikel_i = addr Then Return Elements(i)
                    End If
                Next i
            Case Else
        End Select
        Return Empty_Element
    End Function
    'Private Sub ElementDisplay(pic As Object, turn As Integer, fn As String)
    '    Dim img As Bitmap
    '    If pic.image IsNot Nothing Then pic.image.dispose
    '    If IsDisplayableImage(fn) And FileExists(fn) Then
    '        img = Image.FromFile(fn)
    '        Select Case turn
    '            Case 1 : img.RotateFlip(RotateFlipType.Rotate270FlipNone)
    '            Case 2 : img.RotateFlip(RotateFlipType.Rotate180FlipNone)
    '            Case 3 : img.RotateFlip(RotateFlipType.Rotate90FlipNone)
    '            Case Else
    '        End Select
    '        pic.Image = img
    '        pic.Visible = True
    '    Else
    '        pic.Visible = False
    '        Message("Icon not found: " & fn)
    '    End If
    'End Sub

    Private Sub ElementDisplay(pic As Object, turn As Integer, icon_image As Image)
        Dim img As Bitmap
        If pic IsNot Nothing Then
            If pic.image IsNot Nothing Then pic.image.dispose
            img = icon_image.Clone
            Select Case turn
                Case 1 : img.RotateFlip(RotateFlipType.Rotate270FlipNone)
                Case 2 : img.RotateFlip(RotateFlipType.Rotate180FlipNone)
                Case 3 : img.RotateFlip(RotateFlipType.Rotate90FlipNone)
                Case Else
            End Select
            pic.Image = img
            pic.Visible = True
        End If
    End Sub

    Private Sub ElementDisplay(el As ELEMENT)
        Dim img As Bitmap, pic As Object, turn As Integer
        If el.row_no < 0 Or el.row_no > ROW_MAX Or el.col_no < 0 Or el.col_no > COL_MAX Then Exit Sub ' no action, element is outside of the visible grid
        pic = PicGet(el.row_no, el.col_no) ' the cell where to display the element
        If pic.image IsNot Nothing Then pic.image.dispose ' to prevent automatic garbage collection
        turn = el.drehung_i
        img = ElementIconGet(el).Clone
        Select Case turn
            Case 1 : img.RotateFlip(RotateFlipType.Rotate270FlipNone)
            Case 2 : img.RotateFlip(RotateFlipType.Rotate180FlipNone)
            Case 3 : img.RotateFlip(RotateFlipType.Rotate90FlipNone)
            Case Else
        End Select
        pic.Image = img
        pic.Visible = True
    End Sub

    Private Function ElementIconGet(el As ELEMENT) As Image
        ' Returns the icon (not rotated) matching the element
        Dim typ_s As String
        typ_s = el.typ_s
        ' collapsing the many lamp types and the extra entkuppler type
        If Strings.Left(typ_s, 6) = "lampe_" Then
            typ_s = "lampe"
        ElseIf Strings.Left(typ_s, 11) = "entkuppler_" Then
            typ_s = "entkuppler"
        End If
        Select Case typ_s
            Case "andreaskreuz" : Return picIconAndreaskreuz.Image
            Case "bahnschranke" : Return picIconGate.Image
            Case "custom_perm_left" : Return picIconPermLeft.Image
            Case "custom_perm_right" : Return picIconPermRight.Image
            Case "custom_perm_y" : Return picIconPermY.Image
            Case "dkweiche" : Return picIconCrossSwitch.Image
            Case "drehscheibe" : Return picIconTurntable.Image
            Case "doppelbogen" : Return picIconCurveParallel.Image
            Case "bogen" : Return picIconCurve.Image
            Case "dreiwegweiche" : Return picIconThreeWay.Image
            Case "entkuppler" : Return picIconDecouple.Image
            Case "fahrstrasse" : Return picIconRoute.Image
            Case "gerade" : Return picIconStraight.Image
            Case "kreuzung" : Return picIconCross.Image
            Case "lampe" : Return picIconLamp.Image
            Case "linksweiche" : Return picIconSwitchLeft.Image
            Case "prellbock" : Return picIconEnd.Image
            Case "rechtsweiche" : Return picIconSwitchRight.Image
            Case "s88bogen" : Return picIcons88Curve.Image
            Case "s88doppelbogen" : Return picIcons88CurveParallel.Image
            Case "s88kontakt" : Return picIcons88.Image
            Case "signal" : Return picIconSignal.Image
            Case "signal_p_hp012s" : Return picIconSignal.Image
            Case "signal_f_hp01" : Return picIconSignalFHP01.Image
            Case "signal_f_hp02" : Return picIconSignalFHP01.Image
            Case "signal_f_hp012" : Return picIconSignalFHP012.Image
            Case "signal_f_hp012s" : Return picIconSignalFHP012.Image
            Case "signal_hp012" : Return picIconSignal.Image
            Case "signal_hp012s" : Return picIconSignal.Image
            Case "signal_hp02" : Return picIconSignal.Image
            Case "signal_p_hp012" : Return picIconSignalPHP012.Image
            Case "signal_sh01" : Return picIconSignalSH01.Image
            Case "std_rot" : Return picIconStdRed.Image
            Case "std_rot_gruen_0" : Return picIconStdRed.Image
            Case "std_rot_gruen_1" : Return picIconStdGreen.Image
            Case "pfeil" : Return picIconLayout.Image
            Case "tunnel" : Return picIconTunnel.Image
            Case "unterfuehrung" : Return picIconViaduct.Image
            Case "yweiche" : Return picIconSwitchY.Image
            Case Else
                If el.text <> "" And el.typ_s = "" Then
                    Return picIconText.Image
                Else
                    Message("Unknown element type: " & el.typ_s & " - ID=" & el.id & "  (" & Format(el.row_no) & "," & Format(el.col_no) & ")")
                    Return picIconUnknown.Image
                End If
        End Select
    End Function
    Private Function ElementInsert(picName As String, insertion_tool As String) As ELEMENT
        ' Inserts a new element into array Elements. 
        Dim el_new As ELEMENT = Empty_Element
        Dim s As String = "0x"
        el_new.row_no = Val(Mid(picName, 4, 2))
        el_new.col_no = Val(Mid(picName, 6, 2))
        s = DecIntToHexStr(Page_No) & DecIntToHexStr(el_new.row_no + Row_Offset) & DecIntToHexStr(el_new.col_no + Col_Offset)
        el_new.id_normalized = s
        el_new.id = CS2_id_create(Page_No, el_new.row_no + Row_Offset, el_new.col_no + Col_Offset)
        el_new.zustand_s = ""
        el_new.artikel_i = -1
        Select Case insertion_tool
            Case "andreaskreuz"
                el_new.typ_s = "andreaskreuz"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "cross"
                el_new.typ_s = "kreuzung"
                el_new.drehung_i = 0
            Case "cross_switch"
                el_new.typ_s = "dkweiche"
                el_new.drehung_i = 0
            Case "curve"
                el_new.typ_s = "bogen"
                el_new.drehung_i = DrehungGuessBogen(el_new.row_no, el_new.col_no)
            Case "curve_parallel"
                el_new.typ_s = "doppelbogen"
                el_new.drehung_i = 0
            Case "custom_perm_left"
                el_new.typ_s = "custom_perm_left"
                el_new.drehung_i = 0
            Case "custom_perm_right"
                el_new.typ_s = "custom_perm_right"
                el_new.drehung_i = 0
            Case "custom_perm_y"
                el_new.typ_s = "custom_perm_y"
                el_new.drehung_i = 0
            Case "decouple"
                el_new.typ_s = "entkuppler"
                el_new.drehung_i = DrehungGuessGerade(el_new.row_no, el_new.col_no)
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "end"
                el_new.typ_s = "prellbock"
                el_new.drehung_i = 0
            Case "gate"
                el_new.typ_s = "bahnschranke"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "lamp"
                el_new.typ_s = "lampe"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "pfeil"
                el_new.typ_s = "pfeil"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "route"
                el_new.typ_s = "fahrstrasse"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "signal"
                el_new.typ_s = "signal"
                el_new.drehung_i = DrehungGuessGerade(el_new.row_no, el_new.col_no)
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "signal_f_hp01"
                el_new.typ_s = "signal_f_hp01"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "signal_f_hp012"
                el_new.typ_s = "signal_f_hp012"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "signal_p_hp012"
                el_new.typ_s = "signal_p_hp012"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "signal_sh01"
                el_new.typ_s = "signal_sh01"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "s88"
                el_new.typ_s = "s88kontakt"
                el_new.drehung_i = DrehungGuessGerade(el_new.row_no, el_new.col_no)
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
                If el_new.artikel_i >= 1000 Then
                    el_new.deviceId_s = DeviceIdPrompt("0x...")
                Else
                    el_new.deviceId_s = ""
                End If
            Case "s88_curve"
                el_new.typ_s = "s88bogen"
                el_new.drehung_i = DrehungGuessBogen(el_new.row_no, el_new.col_no)
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
                If el_new.artikel_i >= 1000 Then
                    el_new.deviceId_s = DeviceIdPrompt("0x...")
                Else
                    el_new.deviceId_s = ""
                End If
            Case "s88_curve_parallel"
                el_new.typ_s = "s88doppelbogen"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
                If el_new.artikel_i >= 1000 Then
                    el_new.deviceId_s = DeviceIdPrompt("0x...")
                Else
                    el_new.deviceId_s = ""
                End If
            Case "std_red_green_0" ' is red
                el_new.typ_s = "std_rot_gruen_0"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "std_red_green_1" ' is green
                el_new.typ_s = "std_rot_gruen_1"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "straight"
                el_new.typ_s = "gerade"
                el_new.drehung_i = DrehungGuessGerade(el_new.row_no, el_new.col_no)
            Case "switch_left"
                el_new.typ_s = "linksweiche"
                el_new.drehung_i = 1 - DrehungGuessGerade(el_new.row_no, el_new.col_no)
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "switch_right"
                el_new.typ_s = "rechtsweiche"
                el_new.drehung_i = 1 - DrehungGuessGerade(el_new.row_no, el_new.col_no)
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "switch_y"
                el_new.typ_s = "yweiche"
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "text"
                el_new.typ_s = ""     ' CS2 doen's use a type field for text
                el_new.artikel_i = -1
                el_new.text = InputBox("Enter text", "Text", "Your text here")
                If el_new.text = "" Then el_new.text = "Enter text here" ' Text must be <>"" since we have no type field to identify text
            Case "threeway"
                el_new.typ_s = "dreiwegweiche"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "tunnel"
                el_new.typ_s = "tunnel"
                el_new.drehung_i = 0
            Case "turntable"
                el_new.typ_s = "drehscheibe"
                el_new.drehung_i = 0
                el_new.artikel_i = DigitalArtikelPrompt(el_new)
            Case "viaduct"
                el_new.typ_s = "unterfuehrung"
                el_new.drehung_i = 0
        End Select
        E_TOP = E_TOP + 1
        el_new.my_index = E_TOP
        Elements(E_TOP) = el_new
        Message(2, "ElementInsert(" & picName & ") " & insertion_tool & ", " & el_new.typ_s & " @ " & el_new.id_normalized _
                & " (" & Format(el_new.row_no) & "," & Format(el_new.col_no) & "), E_TOP=" & Format(E_TOP))
        EDIT_CHANGES = True
        Return el_new
    End Function
    Private Sub ElementRemove(element_index As Integer)
        Dim i As Integer
        If element_index > 0 And E_TOP > 0 Then
            E_TOP = E_TOP - 1
            Message(2, "ElementRemove(" & Format(element_index) & ") " & Elements(element_index).typ_s & " @ (" & Format(Elements(element_index).row_no) _
                    & "," & Format(Elements(element_index).col_no) & "), E_TOP hereafter=" & Format(E_TOP))
            For i = element_index To E_TOP
                Elements(i) = Elements(i + 1)
                Elements(i).my_index = i
            Next i
            EDIT_CHANGES = True
        End If
    End Sub

    Private Sub ElementsStore()
        ' store a copy of the current elements while we're editing
        For i = 1 To E_TOP
            Elements2(i) = Elements(i)
        Next
        E_TOP2 = E_TOP
    End Sub

    Private Sub ElementsRecall()
        ' recall the stored elements in case user cancelled out of the editing operation
        For i = 1 To E_TOP2
            Elements(i) = Elements2(i)
        Next
        E_TOP = E_TOP2
    End Sub

    Private Sub ElementsRepaginate(pno As Integer)
        ' changes id's to match the page number pno
        Dim i As Integer, el As ELEMENT
        For i = 1 To E_TOP
            el = Elements(i)
            el.id = CS2_id_create(pno, el.row_no, el.col_no)
            el.id_normalized = IdNormalize(el.id)
            Elements(i) = el
        Next i
    End Sub
    Private Sub ElementsSort()
        ' Insertion sort by id_normalized
        Dim i As Integer, j As Integer, s As String, el As ELEMENT
        For i = 2 To E_TOP
            el = Elements(i)
            s = el.id_normalized
            j = i
            While j > 1 And Elements(j - 1).id_normalized > s
                Elements(j) = Elements(j - 1)
                j = j - 1
            End While
            Elements(j) = el
        Next i
        For i = 1 To E_TOP
            Elements(i).my_index = i
        Next
    End Sub
    Private Function DigitalAddress(el As ELEMENT) As String
        Dim remainder As Integer, da As String
        ' Calculate the MM digital address that the user sees based on the artikel number
        If el.artikel_i > 0 OrElse (el.typ_s = "pfeil" And el.artikel_i = 0) Then
            If Strings.Left(el.typ_s, 3) = "s88" OrElse el.typ_s = "pfeil" OrElse el.typ_s = "fahrstrasse" Then
                da = Format(el.artikel_i)
            ElseIf el.typ_s = "entkuppler" OrElse el.typ_s = "entkuppler_1" OrElse Strings.Left(el.typ_s, 5) = "lampe" OrElse el.typ_s = "std_rot_gruen_1" _
                OrElse el.typ_s = "bahnschranke" OrElse el.typ_s = "andreaskreuz" Then
                da = Format(Math.DivRem(el.artikel_i, 2, remainder))
                If remainder = 1 Then da = da & "b"
            Else
                da = Format(el.artikel_i \ 2)
            End If
        Else
            da = ""
        End If
        Message(4, "DigitalAdress(" & el.typ_s & " ," & Format(el.artikel_i) & ") -> " & da)
        Return da
    End Function
    Private Function DigitalArtikelPrompt(el As ELEMENT) As Integer
        ' The address as represented in the CS2 is in most cases 2 * the value the user sees (called "artikel").
        Dim da_new_i As Integer, da_old_i As Integer, da_old_s As String, artikel As Integer
        da_old_s = DigitalAddress(el)
        da_old_i = Val(da_old_s) ' in case of e.g. 30b, da_old_i becomes 30

        da_new_i = Val(InputBox("Enter the address of this element (you can change it later). Enter -1 to clear the address.", "Address Prompt", Format(da_old_i)))

        If da_new_i < 0 Then ' And da_old_i <= 0 Then ' 231022 removed second expression
            artikel = -1
        ElseIf el.typ_s = "pfeil" Then
            artikel = da_new_i
        ElseIf da_new_i = 0 And da_old_i > 0 Then
            artikel = el.artikel_i
        ElseIf Strings.Left(el.typ_s, 3) = "s88" OrElse el.typ_s = "fahrstrasse" Then
            artikel = da_new_i
        ElseIf el.typ_s = "entkuppler" OrElse el.typ_s = "entkuppler_1" OrElse Strings.Left(el.typ_s, 5) = "lampe" OrElse el.typ_s = "bahnschranke" _
            OrElse el.typ_s = "andreaskreuz" Then
            If vbYes = MsgBox("Is this element controlled by the green button?", vbYesNo) Then
                artikel = 2 * da_new_i + 1
            Else
                artikel = 2 * da_new_i
            End If
        ElseIf el.typ_s = "std_rot_gruen_1" Then
            artikel = 2 * da_new_i + 1
        Else
            artikel = 2 * da_new_i
        End If
        Return artikel
    End Function

    Private Function DeviceIdPrompt(old_deviceId_s As String) As String
        Return (InputBox("Enter the device id (hex string)", "Device ID Prompt", old_deviceId_s))
    End Function
    Private Function DrehungGuessGerade(r As Integer, c As Integer) As Integer
        ' Given a grid cell, look north, south, east and west to guess the drehung for straight (and equivalent) track
        Dim el_n As ELEMENT, el_s As ELEMENT, el_w As ELEMENT, el_e As ELEMENT
        If r > 0 And r < ROW_MAX And c > 0 And c < COL_MAX Then
            el_n = ElementGet2(r - 1, c) ' returns the empty element object (having typ_s="") if the cell is blank
            el_s = ElementGet2(r + 1, c)
            el_w = ElementGet2(r, c - 1)
            el_e = ElementGet2(r, c + 1)
            Select Case el_n.typ_s
                Case "bogen", "s88bogen" : If el_n.drehung_i = 0 Or el_n.drehung_i = 3 Then Return 1
                Case "gerade", "entkupler", "s88kontakt", "signal" : If el_n.drehung_i = 1 Or el_n.drehung_i = 3 Then Return 1
                Case "linksweiche" : If el_n.drehung_i <> 3 Then Return 1
                Case "rechtsweiche" : If el_n.drehung_i <> 1 Then Return 1
            End Select

            Select Case el_s.typ_s
                Case "bogen", "s88bogen" : If el_s.drehung_i = 1 Or el_s.drehung_i = 2 Then Return 1
                Case "gerade", "entkupler", "s88kontakt", "signal" : If el_s.drehung_i = 1 Or el_s.drehung_i = 3 Then Return 1
                Case "linksweiche" : If el_s.drehung_i <> 1 Then Return 1
                Case "rechtsweiche" : If el_s.drehung_i <> 3 Then Return 1
            End Select

            Select Case el_e.typ_s
                Case "bogen", "s88bogen" : If el_e.drehung_i = 2 Or el_e.drehung_i = 3 Then Return 0
                Case "gerade", "entkupler", "s88kontakt", "signal" : If el_e.drehung_i = 1 Then Return 1
                Case "linksweiche" : If el_e.drehung_i = 2 Then Return 1
                Case "rechtsweiche" : If el_e.drehung_i = 0 Then Return 1
            End Select

            Select Case el_w.typ_s
                Case "bogen", "s88bogen" : If el_w.drehung_i = 0 Or el_w.drehung_i = 1 Then Return 0
                Case "gerade", "entkupler", "s88kontakt", "signal" : If el_w.drehung_i = 1 Then Return 1
                Case "linksweiche" : If el_w.drehung_i = 0 Then Return 1
                Case "rechtsweiche" : If el_w.drehung_i = 2 Then Return 1
            End Select
        ElseIf c = 0 Or c = COL_MAX Then
            Return 1
        End If
        Return 0
    End Function
    Private Function DrehungGuessBogen(r As Integer, c As Integer) As Integer
        ' Given a grid cell, look north, south, east and west to guess the drehung for straight (and equivalent) track
        Dim el_n As ELEMENT, el_s As ELEMENT, el_w As ELEMENT, el_e As ELEMENT
        If r > 0 And r < ROW_MAX And c > 0 And c < COL_MAX Then
            el_n = ElementGet2(r - 1, c) ' returns the empty element object (with typ_s="" and drehung_i = -1) if the cell is blank
            el_s = ElementGet2(r + 1, c)
            el_w = ElementGet2(r, c - 1)
            el_e = ElementGet2(r, c + 1)

            Select Case el_n.typ_s
                Case "gerade", "entkupler", "s88kontakt"
                    If el_n.drehung_i = 1 Then
                        If el_e.drehung_i = 1 Or el_w.drehung_i = 0 Then
                            Return 2
                        Else
                            Return 1
                        End If
                    End If
                Case "bogen", "s88bogen" : If el_n.drehung_i <> -1 Then Return (el_n.drehung_i + 2) Mod 4
            End Select

            Select Case el_s.typ_s
                Case "gerade", "entkupler", "s88kontakt"
                    If el_s.drehung_i = 1 Then
                        If el_w.drehung_i = 1 Or el_e.drehung_i = 0 Then
                            Return 0
                        Else
                            Return 3
                        End If
                    End If
                Case "bogen", "s88bogen" : If el_s.drehung_i <> -1 Then Return (el_s.drehung_i + 2) Mod 4
            End Select

            Select Case el_e.typ_s
                Case "gerade", "entkupler", "s88kontakt"
                    If el_e.drehung_i = 1 Then
                        If el_s.drehung_i = 1 Or el_n.drehung_i <> 0 Then
                            Return 3
                        Else
                            Return 0
                        End If
                    End If
                Case "bogen", "s88bogen" : If el_e.drehung_i <> -1 Then Return (el_e.drehung_i + 2) Mod 4
            End Select

            Select Case el_w.typ_s
                Case "gerade", "entkupler", "s88kontakt"
                    If el_w.drehung_i = 0 Then
                        If el_n.drehung_i = 1 Or el_s.drehung_i <> 0 Then
                            Return 2
                        Else
                            Return 1
                        End If
                    End If
                Case "bogen", "s88bogen" : If el_w.drehung_i <> -1 Then Return (el_w.drehung_i + 2) Mod 4
            End Select

        End If
        Return 0
    End Function
    Private Function HasDigitalAddress(element_type As String) As Boolean
        If Strings.Left(element_type, 5) = "entku" Then Return True
        If Strings.Left(element_type, 5) = "lampe" Then Return True
        If Strings.Left(element_type, 3) = "s88" Then Return True
        If Strings.Left(element_type, 6) = "signal" Then Return True
        If Strings.Left(element_type, 3) = "std" Then Return True
        If Strings.Right(element_type, 6) = "weiche" Then Return True
        Select Case element_type
            Case "andreaskreuz" : Return True
            Case "bahnschranke" : Return True
            Case "drehscheibe" : Return True
            Case "fahrstrasse" : Return True
            Case "pfeil" : Return True
            Case "sonstige_gbs" : Return True
            Case Else : Return False
        End Select
    End Function

    Private Function CS2_id_create(page As Integer, row As Integer, col As Integer) As String
        ' Converts the cell position (page, row, col) to a CS2-format id in hex form: 6 hex caracters (in groups of two: page, row, col), however with leading zeros omitted.
        ' Prepends "0x"
        Dim s As String, c As String
        s = DecIntToHexStr(page) & DecIntToHexStr(row) & DecIntToHexStr(col)
        ' remove leading zeros
        For i = 1 To 6
            c = Strings.Left(s, 1)
            If c = "0" Then
                s = Mid(s, 2)
            Else
                Exit For
            End If
        Next
        If s <> "" Then
            s = "0x" & s
        End If
        Return s
    End Function
    Private Function DecodePageNo(p_val As String) As Integer
        ' 02070a - page no 2
        DecodePageNo = HexStrToDecInt(Strings.Left(p_val, 2))
    End Function
    Private Function DecodeColNo(p_val As String) As Integer
        ' 02070a - col no 07 hex (7 decimal)
        DecodeColNo = HexStrToDecInt(Strings.Mid(p_val, 5, 2))
    End Function
    Private Function DecodeRowNo(p_val As String) As Integer
        DecodeRowNo = HexStrToDecInt(Strings.Mid(p_val, 3, 2))
    End Function

    Private Function DecIntToHexStr(n As Integer) As String
        ' converts n (must be <=99) to hex
        Dim v(2) As Integer, c(2) As String
        v(1) = n \ 16
        v(2) = n - 16 * v(1)
        For i = 1 To 2
            Select Case v(i)
                Case 10 : c(i) = "a"
                Case 11 : c(i) = "b"
                Case 12 : c(i) = "c"
                Case 13 : c(i) = "d"
                Case 14 : c(i) = "e"
                Case 15 : c(i) = "f"
                Case Else : c(i) = Format(v(i))
            End Select
        Next i
        Return c(1) & c(2)
    End Function
    Private Function HexStrToDecInt(s As String) As Integer
        ' s must contain two hex digits
        Dim d_str As String, i As Integer
        Dim d_val As Integer, res As Integer

        For i = 1 To 2
            d_str = Strings.Mid(s, i, 1)
            Select Case d_str
                Case "a" : d_val = 10
                Case "b" : d_val = 11
                Case "c" : d_val = 12
                Case "d" : d_val = 13
                Case "e" : d_val = 14
                Case "f" : d_val = 15
                Case Else : d_val = Val(d_str)
            End Select
            If i = 1 Then
                res = 16 * d_val
            Else
                res = res + d_val
            End If
        Next i
        ' Message("HexStrToDecInt(" & s & ") =" & Str(res))
        HexStrToDecInt = res
    End Function
    Private Function IdNormalize(id_val As String)
        ' 0xb05 - row no b hex (11 decimal)   (NB: leading zeroes omitted)
        ' 0x2070a - col no 07 hex (7 decimal)
        ' This function prepends missing zeros and removes "0x"
        Dim s As String = Mid(id_val, 3)
        IdNormalize = StrFixR(s, "0", 6)
    End Function

    Private Function LayoutCreate() As Boolean
        ' Creates a new layout including the directory structure and the page file gleisbild.cs2 and the first track diagram file <page name>.cs2 aas page 0
        ' Example:
        ' ...\Oles Kreds                                            -- layout name
        ' ...\Oles Kreds\config
        ' ...\Oles Kreds\config\gleisbild.cs2                       -- layout page file
        ' ...\Oles Kreds\config\gleisbilder
        ' ...\Oles Kreds\config\gleisbilder\my_track_diagram.cs2    -- track diagram assigned to pge 0
        '
        Dim answer As DialogResult, pno As Integer
        Dim fno As Integer
        Dim layout_root_dir As String
        Dim layout_dir As String, config_dir As String, trackdiagram_dir As String, layout_pages_fn As String
        Dim track_diagram_name As String, track_diagram_filename As String, c As String
        Dim sender As Object = Nothing, e As EventArgs = Nothing ' dummy objects

        If CHANGED Then
            answer = MsgBox("Track diagram " & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)) & " has been modified but not saved. Save before exit?", vbYesNo)
            If answer = vbYes Then
                If Not GleisbildSave() Then
                    answer = MsgBox("File could not be saved: " & TrackDiagram_Input_Filename & ". Continue regardless discarding the changes?", vbYesNo)
                    If answer = vbNo Then Return False
                End If
            End If
        End If

        CHANGED = False

        E_TOP = 0
        Header_Top = 0
        TrackDiagram_Input_Filename = ""
        Page_No = 0
        GleisBildDisplayClear()

        MsgBox("Navigate to the desired location and create a new subfolder with the name of your new layout")
        ' try and find the layout root dir
        If InStr(TrackDiagram_Directory, "\config\") > 0 Then
            layout_root_dir = FilePathUp(TrackDiagram_Directory)
            layout_root_dir = FilePathUp(layout_root_dir)
            layout_root_dir = FilePathUp(layout_root_dir)
            FolderBrowserDialog1.SelectedPath = layout_root_dir ' SaveFileDialog messes with me in case of an existing file
        Else
            FolderBrowserDialog1.SelectedPath = TrackDiagram_Directory ' SaveFileDialog messes with me in case of an existing file
        End If

        answer = FolderBrowserDialog1.ShowDialog()
        If answer = DialogResult.OK Then
            layout_dir = FolderBrowserDialog1.SelectedPath
            config_dir = layout_dir & "\config"

            ' Prompt for the name of the track diagram to go onto page 0. If no name or an invalide name is given, we abort with no changes made to the directory structure
            track_diagram_name = InputBox("Enter the desired name of the first track diagram (must be a valid file name)", "Track Diagram Name", "Main")
            If track_diagram_name = "" Then
                MsgBox("No layout was created")
                Return False
            End If
            track_diagram_filename = config_dir & "\gleisbilder\" & track_diagram_name & ".cs2"
            c = ValidateFileName(track_diagram_name)
            If c <> "" Then
                MsgBox("Illegal character in track diagram name: '" & c & "'. Try again")
                track_diagram_name = InputBox("Enter the desired name of the first track diagram (must be a valid file name)", "Track Diagram Name", "Main")
                If track_diagram_name = "" Then
                    MsgBox("No layout was created")
                    Return False
                End If
                track_diagram_filename = config_dir & "\gleisbilder\" & track_diagram_name & ".cs2"
                c = ValidateFileName(track_diagram_name)
                If c <> "" Then
                    MsgBox("Illegal character in track diagram name: '" & c & "'. No layout was created")
                    Return False
                End If
            End If

            If Dir(config_dir, vbDirectory) = "" Then
                On Error GoTo config_dir_create_error
                MkDir(config_dir)
                On Error GoTo 0
            End If

            trackdiagram_dir = config_dir & "\gleisbilder"

            If Dir(trackdiagram_dir, vbDirectory) = "" Then
                On Error GoTo track_diagram_dir_create_error
                MkDir(trackdiagram_dir)
                On Error GoTo 0
            End If

            layout_pages_fn = config_dir & "\gleisbild.cs2"

            If FileExists(layout_pages_fn) Then
                MsgBox("Layout exists already. You're not allowed to overwrite an existing layout")
                Return False
            End If

            ' Create the page file with page 0
            On Error GoTo write_error
            fno = FreeFile()
            FileOpen(fno, layout_pages_fn, OpenMode.Output)
            PrintLine(fno, "[gleisbild]")
            PrintLine(fno, "Version")
            PrintLine(fno, ".major = 1")
            PrintLine(fno, "groesse")
            PrintLine(fno, "zuletztBenutzt")
            PrintLine(fno, " .name = " & track_diagram_name)
            PrintLine(fno, "seite")
            PrintLine(fno, " .name = " & track_diagram_name)
            On Error GoTo 0
            FileClose(fno)

            ' Create the track diagram file
            On Error GoTo write_error
            fno = FreeFile()
            FileOpen(fno, track_diagram_filename, OpenMode.Output)
            PrintLine(fno, "[gleisbildseite]")
            PrintLine(fno, "version")
            PrintLine(fno, " .major=1")
            On Error GoTo 0
            FileClose(fno)

            MasterFile_Directory = layout_dir
            Message("Layout """ & ExtractLastField(layout_dir, "\") & """ was created here: " & layout_dir)
            Message("Track diagram """ & FilenameProper(track_diagram_filename) & """ was created and assigned to page 0")
            GleisbildLoad(track_diagram_filename)

            '            PFH.PageFileCreate(track_diagram_filename)
            '           PFH.PageFileLoad(MasterFile_Directory)
            '          PFH.PagesPopulate()

            cmdEdit_Click(sender, e)

            On Error GoTo 0
            FileClose(fno)
        End If

        Return True
config_dir_create_error:
        On Error GoTo 0
        MsgBox("Error creating configuration folder " & config_dir)
        Return False
track_diagram_dir_create_error:
        On Error GoTo 0
        MsgBox("Error creating track diagram folder " & trackdiagram_dir)
        Return False
write_error:
        On Error GoTo 0
        '  MsgBox("Error creating or writing file " & fn & " (m69)")
        FileClose(fno)
        Return False
    End Function
    Private Function SplitHorizontally(ByVal split_row As Integer) As Boolean
        ' Returns true if the split could be done.
        ' Note that row_no and col_no of the element can be displaced depending on the viewing window's position.
        ' Therefore we must rely on the ID for row/col information.
        Dim i As Integer, el As ELEMENT, s As String
        Dim rno As Integer, cno As Integer ' row/col numvers as represented in the ID
        split_row = split_row + Row_Offset
        If split_row > 0 And split_row < 255 And UpperLimitRow() < 255 Then
            For i = 1 To E_TOP Step 1
                el = Elements(i)
                cno = DecodeColNo(el.id_normalized)
                rno = DecodeRowNo(el.id_normalized)
                If rno >= split_row Then
                    el.row_no = el.row_no + 1
                    rno = rno + 1
                    s = DecIntToHexStr(Page_No) & DecIntToHexStr(rno) & DecIntToHexStr(cno)
                    el.id_normalized = s
                    el.id = CS2_id_create(Page_No, rno, cno)
                    Elements(i) = el
                End If
            Next i
            Return True
        Else
            Return False
        End If
    End Function
    Private Function SplitVertically(ByVal split_col As Integer) As Boolean
        ' Returns true if the split could be done
        ' Note that row_no and col_no of the element can be displaced depending on the viewing window's position.
        ' Therefore we must rely on the ID for row/col information.
        Dim i As Integer, el As ELEMENT, s As String
        Dim rno As Integer, cno As Integer ' row/col numvers as represented in the ID
        split_col = split_col + Col_Offset
        If split_col > 0 And split_col < 255 And UpperLimitCol() < 255 Then
            For i = 1 To E_TOP Step 1
                el = Elements(i)
                cno = DecodeColNo(el.id_normalized)
                rno = DecodeRowNo(el.id_normalized)
                If cno >= split_col Then
                    el.col_no = el.col_no + 1
                    cno = cno + 1
                    s = DecIntToHexStr(Page_No) & DecIntToHexStr(rno) & DecIntToHexStr(cno)
                    el.id_normalized = s
                    el.id = CS2_id_create(Page_No, rno, cno)
                    Elements(i) = el
                End If
            Next i
            Return True
        Else
            Return False
        End If
    End Function
    Public Sub Message(s As String)
        Message(0, s)
    End Sub
    Public Sub Message(level As Integer, s As String)
        If Test_Level >= level Then
            lstMessage.Items.Add(s)
            lstMessage.SelectedIndex = lstMessage.Items.Count - 1
        End If
    End Sub
    Private Sub PropertiesDisplay(el As ELEMENT)
        lstProperties.Items.Clear()
        If el.typ_s <> "" Or el.text <> "" Then
            lstProperties.Items.Add(Translate(el.typ_s))
            lstProperties.Items.Add("cell   = (" & Format(el.row_no) & "," & Format(el.col_no) & ")")
            lstProperties.Items.Add("page   = " & Format(Page_No))
            If Developer_Environment Then lstProperties.Items.Add("id_norm= " & el.id_normalized)
            If Developer_Environment Then lstProperties.Items.Add("row/col= (" & DecodeRowNo(el.id_normalized) & "," & DecodeColNo(el.id_normalized) & ")")
            lstProperties.Items.Add("address= " & DigitalAddress(el))
            If Developer_Environment Or Verbose Then lstProperties.Items.Add("element= " & Format(el.my_index) & " of " & Format(E_TOP))
            If el.text <> "" Then
                lstProperties.Items.Add("text   = " & el.text)
            Else
                lstProperties.Items.Add("")
            End If

            If Verbose Or Developer_Environment Then
                lstProperties.Items.Add("")
                lstProperties.Items.Add("In CS2 format:")
                lstProperties.Items.Add("element")
                If el.id <> "" Then lstProperties.Items.Add(" .id=" & el.id)
                If el.typ_s <> "" Then lstProperties.Items.Add(" .typ=" & el.typ_s)
                If el.drehung_i > 0 Then lstProperties.Items.Add(" .drehung=" & Format(el.drehung_i))
                lstProperties.Items.Add(" .artikel=" & Format(el.artikel_i))
                If el.deviceId_s <> "" Then lstProperties.Items.Add(" .deviceID=" & el.deviceId_s)
                If el.zustand_s <> "" Then lstProperties.Items.Add(" .zustand=" & Format(el.zustand_s))
                If el.text <> "" Then lstProperties.Items.Add(" .text=" & Format(el.text))
            End If
        End If
    End Sub

    Private Sub ElementDisplayAddress(el As ELEMENT, e As PaintEventArgs)
        ' Display the address of the element
        Dim turn As Integer, element_name As String, da As String
        If el.row_no > 255 Or el.col_no > 255 Then Exit Sub ' no action, element is outside of the visible grid
        element_name = el.typ_s
        turn = el.drehung_i
        If HasDigitalAddress(element_name) Then
            da = DigitalAddress(el)
            Select Case element_name
                Case "signal", "signal_sh01", "signal_hp02", "signal_hp012", "signal_hp012s", "signal_p_hp012"
                    If turn = 0 Then
                        StringDisplay(e, da, Brush_red, 5, 17)
                    ElseIf turn = 1 Then
                        StringDisplayVertical(e, da, Brush_red, 17, 5)
                    ElseIf turn = 2 Then
                        StringDisplay(e, da, Brush_red, 5, 0)
                    Else
                        StringDisplayVertical(e, da, Brush_red, -2, 5)
                    End If
                Case "andreaskreuz"
                    StringDisplay(e, da, Brush_black, 2, 13)
                Case "drehscheibe"
                    StringDisplay(e, da, Brush_yellow, 2, 15)
                Case "sonstige_gbs"
                    StringDisplayVertical(e, da, Brush_black, 17, 0)
                Case "entkuppler", "entkuppler_1", "fahrstrasse", "pfeil"
                    If turn = 0 Then
                        StringDisplay(e, da, Brush_black, 5, 17)
                    ElseIf turn = 1 Then
                        StringDisplayVertical(e, da, Brush_black, 17, 5)
                    ElseIf turn = 2 Then
                        StringDisplay(e, da, Brush_black, 5, 0)
                    Else
                        StringDisplayVertical(e, da, Brush_black, -2, 5)
                    End If
                Case "signal_f_hp01", "signal_f_hp012", "signal_f_hp012s", "signal_f_hp02"
                    If turn = 2 Then
                        StringDisplayVertical(e, da, Brush_red, 17, 5)
                    ElseIf turn = 3 Then
                        StringDisplay(e, da, Brush_red, 5, 0)
                    ElseIf turn = 0 Then
                        StringDisplayVertical(e, da, Brush_red, -2, 5)
                    Else
                        StringDisplay(e, da, Brush_red, 5, 17)
                    End If
                Case "std_rot_gruen_0", "std_rot"
                    StringDisplay(e, da, Brush_black, 0, 0)
                Case "std_rot_gruen_1"
                    StringDisplay(e, da, Brush_black, 0, 17)
                Case "s88kontakt", "s88bogen", "s88doppelbogen", "bahnschranke"
                    If turn = 0 Then
                        StringDisplay(e, da, Brush_black, 0, 18)
                    ElseIf turn = 2 Then
                        StringDisplay(e, da, Brush_black, 0, 0)
                    ElseIf turn = 1 Then
                        StringDisplayVertical(e, da, Brush_black, 17, 0)
                    Else
                        StringDisplayVertical(e, da, Brush_black, -2, 0)
                    End If
                Case "linksweiche"
                    If turn = 0 Then
                        StringDisplayVertical(e, da, Brush_blue, 17, 5)
                    ElseIf turn = 1 Then
                        StringDisplay(e, da, Brush_blue, 5, 0)
                    ElseIf turn = 2 Then
                        StringDisplayVertical(e, da, Brush_blue, -2, 5)
                    Else
                        StringDisplay(e, da, Brush_blue, 5, 17)
                    End If
                Case "rechtsweiche"
                    If turn = 2 Then
                        StringDisplayVertical(e, da, Brush_blue, 17, 5)
                    ElseIf turn = 3 Then
                        StringDisplay(e, da, Brush_blue, 5, 0)
                    ElseIf turn = 0 Then
                        StringDisplayVertical(e, da, Brush_blue, -2, 5)
                    Else
                        StringDisplay(e, da, Brush_blue, 5, 17)
                    End If
                Case "dreiwegweiche"
                    If turn = 3 Then
                        StringDisplay(e, da, Brush_red, 5, 17)
                    ElseIf turn = 0 Then
                        StringDisplayVertical(e, da, Brush_red, 17, 5)
                    ElseIf turn = 1 Then
                        StringDisplay(e, da, Brush_red, 5, -2)
                    Else
                        StringDisplayVertical(e, da, Brush_red, -2, 5)
                    End If
                Case "yweiche"
                    If turn = 2 Then
                        StringDisplay(e, da, Brush_red, 5, 17)
                    ElseIf turn = 3 Then
                        StringDisplayVertical(e, da, Brush_red, 17, 5)
                    ElseIf turn = 0 Then
                        StringDisplay(e, da, Brush_red, 5, 0)
                    Else
                        StringDisplayVertical(e, da, Brush_red, -2, 5)
                    End If
                Case "lampe", "lampe_ge", "lampe_rt", "lampe_gn", "lampe_bl"
                    StringDisplay(e, da, Brush_black, 5, 0)
                Case Else
                    If Developer_Environment Then Message("Address display not impelented for " & element_name)
            End Select
        End If
    End Sub
    Private Sub ElementDisplayText(el As ELEMENT, e As PaintEventArgs)
        ' overlays the el.text on the element's icon
        If el.row_no > 255 Or el.col_no > 255 Then Exit Sub ' no action, element is outside of the visible grid
        If el.text <> "" Then ' 231024
            Select Case el.typ_s
                Case "bogen"
                    If el.drehung_i = 0 Or el.drehung_i = 3 Then
                        StringDisplay(e, el.text, Brush_black, 0, 0)
                    Else
                        StringDisplay(e, el.text, Brush_black, 0, 15)
                    End If
                Case "linksweiche"
                    If el.drehung_i = 0 Then
                        StringDisplayVertical(e, el.text, Brush_black, 17, 5)
                    ElseIf el.drehung_i = 1 Then
                        StringDisplay(e, el.text, Brush_black, 5, -2)
                    ElseIf el.drehung_i = 2 Then
                        StringDisplayVertical(e, el.text, Brush_black, -2, 5)
                    Else
                        StringDisplay(e, el.text, Brush_black, 5, 17)
                    End If
                Case "rechtsweiche"
                    If el.drehung_i = 2 Then
                        StringDisplayVertical(e, el.text, Brush_black, 17, 5)
                    ElseIf el.drehung_i = 3 Then
                        StringDisplay(e, el.text, Brush_black, 5, -2)
                    ElseIf el.drehung_i = 0 Then
                        StringDisplayVertical(e, el.text, Brush_black, -2, 5)
                    Else
                        StringDisplay(e, el.text, Brush_black, 5, 17)
                    End If
                Case Else
                    If el.drehung_i = 0 Then
                        StringDisplay(e, el.text, Brush_black, 0, 18)
                    ElseIf el.drehung_i = 1 Then
                        StringDisplayVertical(e, el.text, Brush_black, 17, 5)
                    Else
                        StringDisplay(e, el.text, Brush_black, 0, 18) '231024
                    End If
            End Select
        End If
    End Sub

    Private Sub TextOrAddressDisplay(sender As Object, e As PaintEventArgs) Handles pic0000.Paint, pic0001.Paint, pic0002.Paint, pic0003.Paint, pic0004.Paint, pic0005.Paint, pic0006.Paint, pic0007.Paint, pic0008.Paint, pic0009.Paint,
                                                                                    pic0010.Paint, pic0011.Paint, pic0012.Paint, pic0013.Paint, pic0014.Paint, pic0015.Paint, pic0016.Paint, pic0017.Paint, pic0018.Paint, pic0019.Paint,
                                                                                    pic0020.Paint, pic0021.Paint, pic0022.Paint, pic0023.Paint, pic0024.Paint, pic0025.Paint, pic0026.Paint, pic0027.Paint, pic0028.Paint, pic0029.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay1(sender As Object, e As PaintEventArgs) Handles pic0100.Paint, pic0101.Paint, pic0102.Paint, pic0103.Paint, pic0104.Paint, pic0105.Paint, pic0106.Paint, pic0107.Paint, pic0108.Paint, pic0109.Paint,
                                                                                    pic0110.Paint, pic0111.Paint, pic0112.Paint, pic0113.Paint, pic0114.Paint, pic0115.Paint, pic0116.Paint, pic0117.Paint, pic0118.Paint, pic0119.Paint,
                                                                                    pic0120.Paint, pic0121.Paint, pic0122.Paint, pic0123.Paint, pic0124.Paint, pic0125.Paint, pic0126.Paint, pic0127.Paint, pic0128.Paint, pic0129.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay2(sender As Object, e As PaintEventArgs) Handles pic0200.Paint, pic0201.Paint, pic0202.Paint, pic0203.Paint, pic0204.Paint, pic0205.Paint, pic0206.Paint, pic0207.Paint, pic0208.Paint, pic0209.Paint,
                                                                                    pic0210.Paint, pic0211.Paint, pic0212.Paint, pic0213.Paint, pic0214.Paint, pic0215.Paint, pic0216.Paint, pic0217.Paint, pic0218.Paint, pic0219.Paint,
                                                                                    pic0220.Paint, pic0221.Paint, pic0222.Paint, pic0223.Paint, pic0224.Paint, pic0225.Paint, pic0226.Paint, pic0227.Paint, pic0228.Paint, pic0229.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay3(sender As Object, e As PaintEventArgs) Handles pic0300.Paint, pic0301.Paint, pic0302.Paint, pic0303.Paint, pic0304.Paint, pic0305.Paint, pic0306.Paint, pic0307.Paint, pic0308.Paint, pic0309.Paint,
                                                                                    pic0310.Paint, pic0311.Paint, pic0312.Paint, pic0313.Paint, pic0314.Paint, pic0315.Paint, pic0316.Paint, pic0317.Paint, pic0318.Paint, pic0319.Paint,
                                                                                    pic0320.Paint, pic0321.Paint, pic0322.Paint, pic0323.Paint, pic0324.Paint, pic0325.Paint, pic0326.Paint, pic0327.Paint, pic0328.Paint, pic0329.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay4(sender As Object, e As PaintEventArgs) Handles pic0400.Paint, pic0401.Paint, pic0402.Paint, pic0403.Paint, pic0404.Paint, pic0405.Paint, pic0406.Paint, pic0407.Paint, pic0408.Paint, pic0409.Paint,
                                                                                    pic0410.Paint, pic0411.Paint, pic0412.Paint, pic0413.Paint, pic0414.Paint, pic0415.Paint, pic0416.Paint, pic0417.Paint, pic0418.Paint, pic0419.Paint,
                                                                                    pic0420.Paint, pic0421.Paint, pic0422.Paint, pic0423.Paint, pic0424.Paint, pic0425.Paint, pic0426.Paint, pic0427.Paint, pic0428.Paint, pic0429.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay5(sender As Object, e As PaintEventArgs) Handles pic0500.Paint, pic0501.Paint, pic0502.Paint, pic0503.Paint, pic0504.Paint, pic0505.Paint, pic0506.Paint, pic0507.Paint, pic0508.Paint, pic0509.Paint,
                                                                                    pic0510.Paint, pic0511.Paint, pic0512.Paint, pic0513.Paint, pic0514.Paint, pic0515.Paint, pic0516.Paint, pic0517.Paint, pic0518.Paint, pic0519.Paint,
                                                                                    pic0520.Paint, pic0521.Paint, pic0522.Paint, pic0523.Paint, pic0524.Paint, pic0525.Paint, pic0526.Paint, pic0527.Paint, pic0528.Paint, pic0529.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay6(sender As Object, e As PaintEventArgs) Handles pic0600.Paint, pic0601.Paint, pic0602.Paint, pic0603.Paint, pic0604.Paint, pic0605.Paint, pic0606.Paint, pic0607.Paint, pic0608.Paint, pic0609.Paint,
                                                                                    pic0610.Paint, pic0611.Paint, pic0612.Paint, pic0613.Paint, pic0614.Paint, pic0615.Paint, pic0616.Paint, pic0617.Paint, pic0618.Paint, pic0619.Paint,
                                                                                    pic0620.Paint, pic0621.Paint, pic0622.Paint, pic0623.Paint, pic0624.Paint, pic0625.Paint, pic0626.Paint, pic0627.Paint, pic0628.Paint, pic0629.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay7(sender As Object, e As PaintEventArgs) Handles pic0700.Paint, pic0701.Paint, pic0702.Paint, pic0703.Paint, pic0704.Paint, pic0705.Paint, pic0706.Paint, pic0707.Paint, pic0708.Paint, pic0709.Paint,
                                                                                    pic0710.Paint, pic0711.Paint, pic0712.Paint, pic0713.Paint, pic0714.Paint, pic0715.Paint, pic0716.Paint, pic0717.Paint, pic0718.Paint, pic0719.Paint,
                                                                                    pic0720.Paint, pic0721.Paint, pic0722.Paint, pic0723.Paint, pic0724.Paint, pic0725.Paint, pic0726.Paint, pic0727.Paint, pic0728.Paint, pic0729.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay8(sender As Object, e As PaintEventArgs) Handles pic0800.Paint, pic0801.Paint, pic0802.Paint, pic0803.Paint, pic0804.Paint, pic0805.Paint, pic0806.Paint, pic0807.Paint, pic0808.Paint, pic0809.Paint,
                                                                                    pic0810.Paint, pic0811.Paint, pic0812.Paint, pic0813.Paint, pic0814.Paint, pic0815.Paint, pic0816.Paint, pic0817.Paint, pic0818.Paint, pic0819.Paint,
                                                                                    pic0820.Paint, pic0821.Paint, pic0822.Paint, pic0823.Paint, pic0824.Paint, pic0825.Paint, pic0826.Paint, pic0827.Paint, pic0828.Paint, pic0829.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay9(sender As Object, e As PaintEventArgs) Handles pic0900.Paint, pic0901.Paint, pic0902.Paint, pic0903.Paint, pic0904.Paint, pic0905.Paint, pic0906.Paint, pic0907.Paint, pic0908.Paint, pic0909.Paint,
                                                                                    pic0910.Paint, pic0911.Paint, pic0912.Paint, pic0913.Paint, pic0914.Paint, pic0915.Paint, pic0916.Paint, pic0917.Paint, pic0918.Paint, pic0919.Paint,
                                                                                    pic0920.Paint, pic0921.Paint, pic0922.Paint, pic0923.Paint, pic0924.Paint, pic0925.Paint, pic0926.Paint, pic0927.Paint, pic0928.Paint, pic0929.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay10(sender As Object, e As PaintEventArgs) Handles pic1000.Paint, pic1001.Paint, pic1002.Paint, pic1003.Paint, pic1004.Paint, pic1005.Paint, pic1006.Paint, pic1007.Paint, pic1008.Paint, pic1009.Paint,
                                                                                    pic1010.Paint, pic1011.Paint, pic1012.Paint, pic1013.Paint, pic1014.Paint, pic1015.Paint, pic1016.Paint, pic1017.Paint, pic1018.Paint, pic1019.Paint,
                                                                                    pic1020.Paint, pic1021.Paint, pic1022.Paint, pic1023.Paint, pic1024.Paint, pic1025.Paint, pic1026.Paint, pic1027.Paint, pic1028.Paint, pic1029.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay11(sender As Object, e As PaintEventArgs) Handles pic1100.Paint, pic1101.Paint, pic1102.Paint, pic1103.Paint, pic1104.Paint, pic1105.Paint, pic1106.Paint, pic1107.Paint, pic1108.Paint, pic1109.Paint,
                                                                                    pic1110.Paint, pic1111.Paint, pic1112.Paint, pic1113.Paint, pic1114.Paint, pic1115.Paint, pic1116.Paint, pic1117.Paint, pic1118.Paint, pic1119.Paint,
                                                                                    pic1120.Paint, pic1121.Paint, pic1122.Paint, pic1123.Paint, pic1124.Paint, pic1125.Paint, pic1126.Paint, pic1127.Paint, pic1128.Paint, pic1129.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay12(sender As Object, e As PaintEventArgs) Handles pic1200.Paint, pic1201.Paint, pic1202.Paint, pic1203.Paint, pic1204.Paint, pic1205.Paint, pic1206.Paint, pic1207.Paint, pic1208.Paint, pic1209.Paint,
                                                                                    pic1210.Paint, pic1211.Paint, pic1212.Paint, pic1213.Paint, pic1214.Paint, pic1215.Paint, pic1216.Paint, pic1217.Paint, pic1218.Paint, pic1219.Paint,
                                                                                    pic1220.Paint, pic1221.Paint, pic1222.Paint, pic1223.Paint, pic1224.Paint, pic1225.Paint, pic1226.Paint, pic1227.Paint, pic1228.Paint, pic1229.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay13(sender As Object, e As PaintEventArgs) Handles pic1300.Paint, pic1301.Paint, pic1302.Paint, pic1303.Paint, pic1304.Paint, pic1305.Paint, pic1306.Paint, pic1307.Paint, pic1308.Paint, pic1309.Paint,
                                                                                    pic1310.Paint, pic1311.Paint, pic1312.Paint, pic1313.Paint, pic1314.Paint, pic1315.Paint, pic1316.Paint, pic1317.Paint, pic1318.Paint, pic1319.Paint,
                                                                                    pic1320.Paint, pic1321.Paint, pic1322.Paint, pic1323.Paint, pic1324.Paint, pic1325.Paint, pic1326.Paint, pic1327.Paint, pic1328.Paint, pic1329.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay14(sender As Object, e As PaintEventArgs) Handles pic1400.Paint, pic1401.Paint, pic1402.Paint, pic1403.Paint, pic1404.Paint, pic1405.Paint, pic1406.Paint, pic1407.Paint, pic1408.Paint, pic1409.Paint,
                                                                                    pic1410.Paint, pic1411.Paint, pic1412.Paint, pic1413.Paint, pic1414.Paint, pic1415.Paint, pic1416.Paint, pic1417.Paint, pic1418.Paint, pic1419.Paint,
                                                                                    pic1420.Paint, pic1421.Paint, pic1422.Paint, pic1423.Paint, pic1424.Paint, pic1425.Paint, pic1426.Paint, pic1427.Paint, pic1428.Paint, pic1429.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay15(sender As Object, e As PaintEventArgs) Handles pic1500.Paint, pic1501.Paint, pic1502.Paint, pic1503.Paint, pic1504.Paint, pic1505.Paint, pic1506.Paint, pic1507.Paint, pic1508.Paint, pic1509.Paint,
                                                                                    pic1510.Paint, pic1511.Paint, pic1512.Paint, pic1513.Paint, pic1514.Paint, pic1515.Paint, pic1516.Paint, pic1517.Paint, pic1518.Paint, pic1519.Paint,
                                                                                    pic1520.Paint, pic1521.Paint, pic1522.Paint, pic1523.Paint, pic1524.Paint, pic1525.Paint, pic1526.Paint, pic1527.Paint, pic1528.Paint, pic1529.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub
    Private Sub TextOrAddressDisplay16(sender As Object, e As PaintEventArgs) Handles pic1600.Paint, pic1601.Paint, pic1602.Paint, pic1603.Paint, pic1604.Paint, pic1605.Paint, pic1606.Paint, pic1607.Paint, pic1608.Paint, pic1609.Paint,
                                                                                    pic1610.Paint, pic1611.Paint, pic1612.Paint, pic1613.Paint, pic1614.Paint, pic1615.Paint, pic1616.Paint, pic1617.Paint, pic1618.Paint, pic1619.Paint,
                                                                                    pic1620.Paint, pic1621.Paint, pic1622.Paint, pic1623.Paint, pic1624.Paint, pic1625.Paint, pic1626.Paint, pic1627.Paint, pic1628.Paint, pic1629.Paint
        TextOrAddressDisplayBody(sender, e)
    End Sub

    Private Sub TextOrAddressDisplayBody(sender As Object, e As PaintEventArgs)
        Dim picName As String = sender.name
        Dim el As ELEMENT
        If FILE_LOADED Then
            If Reveal_Addr Then
                el = ElementGet(Mid(picName, 4, 4))
                If el.id = "" Then Exit Sub ' empty element
                ElementDisplayAddress(el, e)
            End If
            If Reveal_Text Then
                el = ElementGet(Mid(picName, 4, 4))
                If el.text = "" Then Exit Sub ' no text (or empty element)
                ElementDisplayText(el, e)
            End If
        End If
    End Sub

    Public Sub StringDisplay(e As PaintEventArgs, msg As String, color As Brush, x_px As Integer, y_px As Integer)
        If e IsNot Nothing Then
            e.Graphics.DrawString(msg, myFont, color, x_px, y_px)
        End If
    End Sub

    Public Sub StringDisplayVertical(e As PaintEventArgs, msg As String, color As Brush, x_px As Integer, y_px As Integer)
        If e IsNot Nothing Then
            e.Graphics.DrawString(msg, myFont, color, x_px, y_px, drawFormatVertical)
        End If
    End Sub
    Private Function Translate(element_name As String) As String
        ' Translates the German names to English (solely for display purposes if selected in the tools menu)
        If Not Use_English_Terms Then
            Return element_name
        Else
            Select Case element_name
                Case "andreaskreuz" : Return "andreas cross"
                Case "bahnschranke" : Return "railroad crossing"
                Case "bogen" : Return "curve"
                Case "doppelbogen" : Return "double curve"
                Case "drehscheibe" : Return "turn table"
                Case "dreiwegweiche" : Return "three way switch"
                Case "dkweiche" : Return "cross switch"
                Case "entkuppler" : Return "decoupler"
                Case "entkuppler_1" : Return "decoupler"
                Case "fahrstrasse" : Return "route"
                Case "gerade" : Return "straight"
                Case "kreuzung" : Return "crossing"
                Case "lampe" : Return "lamp"
                Case "lampe_ge" : Return "lamp_ye"
                Case "lampe_rt" : Return "lamp_rd"
                Case "linksweiche" : Return "left hand switch"
                Case "pfeil" : Return "go to"
                Case "prellbock" : Return "track bumper"
                Case "rechtsweiche" : Return "right hand switch"
                Case "s88bogen" : Return "s88 contact curve"
                Case "s88doppelbogen" : Return "s88 contact dbl curve"
                Case "s88kontakt" : Return "s88 contact"
                Case "s88kontakt" : Return "s88 contact"
                Case "std_rot" : Return "std_red"
                Case "std_rot_gruen_0" : Return "std_red_green_0"
                Case "std_rot_gruen_1" : Return "std_red_green_1"
                Case "signal_f_hp012" : Return "signal"
                Case "signal_sh01" : Return "signal"
                Case "unterfuehrung" : Return "viaduct"
                Case "yweiche" : Return "Y switch"
                Case Else
                    Return element_name
            End Select
        End If

    End Function
    Private Function UpperLimitCol() As Integer
        ' returns the value of the rightmost column in use
        Dim col_ul As Integer = 0, i As Integer
        For i = 1 To E_TOP Step 1
            If Elements(i).col_no > col_ul Then col_ul = Elements(i).col_no
        Next i
        Return col_ul
    End Function
    Private Function UpperLimitRow() As Integer
        ' returns the value of the lower-most row in use
        Dim row_ul As Integer = 0, i As Integer
        For i = 1 To E_TOP Step 1
            If Elements(i).row_no > row_ul Then row_ul = Elements(i).row_no
        Next i
        Return row_ul
    End Function
    Private Sub VisibilitySetCommandlineMode(m As String)
        ' hides certain features and menus in command line mode
        If m = "Edit" Then
            NewLayoutToolStripMenuItem.Visible = False
            OpenToolStripMenuItem.Visible = False
            ReopenLastToolStripMenuItem.Visible = False
            SaveAsToolStripMenuItem.Visible = False
            HelpToolStripMenuItem.Visible = False
        End If
        QuitToolStripMenuItem.Visible = False ' menu item has been replaced with buttons
    End Sub
    Private Sub WindowMove(ByVal n_rows As Integer, ByVal n_cols As Integer)
        Dim ul_row As Integer, ul_col As Integer ' new coordinates of upper left cell after adjustment for sanity
        Dim i As Integer
        Dim delta As Integer
        If FILE_LOADED Then
            ul_row = Row_Offset + n_rows
            If n_rows < 0 And ul_row < 0 Then
                delta = ul_row
                n_rows = n_rows - delta
            ElseIf n_rows > 0 And ul_row >= (255 - ROW_MAX) Then
                delta = ul_row - (255 - ROW_MAX)
                n_rows = n_rows - delta
            End If

            Row_Offset = Row_Offset + n_rows

            ul_col = Col_Offset + n_cols
            If n_cols < 0 And ul_col < 0 Then
                delta = ul_col
                n_cols = n_cols - delta
            ElseIf n_cols > 0 And ul_col >= (255 - COL_MAX) Then
                delta = ul_col - (255 - COL_MAX)
                n_cols = n_cols - delta
            End If

            Col_Offset = Col_Offset + n_cols
            AxesDraw()

            If n_rows <> 0 Or n_cols <> 0 Then
                For i = 0 To E_TOP
                    Elements(i).row_no = Elements(i).row_no - n_rows
                    Elements(i).col_no = Elements(i).col_no - n_cols
                Next i
                If MODE <> "Routes" Or cboRouteNames.Text = "Pick a route" Then
                    GleisBildDisplayClear()
                    GleisbildDisplay()
                Else ' a route is being displayed. Reload it
                    RH.RouteShow(cboRouteNames.Text)
                End If
            End If
        End If
    End Sub
    Private Function PicGet(row_no As Integer, col_no As Integer) As PictureBox
        Dim e As ELEMENT
        Select Case row_no
            Case 0 : Select Case col_no
                    Case 0 : Return pic0000
                    Case 1 : Return pic0001
                    Case 2 : Return pic0002
                    Case 3 : Return pic0003
                    Case 4 : Return pic0004
                    Case 5 : Return pic0005
                    Case 6 : Return pic0006
                    Case 7 : Return pic0007
                    Case 8 : Return pic0008
                    Case 9 : Return pic0009
                    Case 10 : Return pic0010
                    Case 11 : Return pic0011
                    Case 12 : Return pic0012
                    Case 13 : Return pic0013
                    Case 14 : Return pic0014
                    Case 15 : Return pic0015
                    Case 16 : Return pic0016
                    Case 17 : Return pic0017
                    Case 18 : Return pic0018
                    Case 19 : Return pic0019
                    Case 20 : Return pic0020
                    Case 21 : Return pic0021
                    Case 22 : Return pic0022
                    Case 23 : Return pic0023
                    Case 24 : Return pic0024
                    Case 25 : Return pic0025
                    Case 26 : Return pic0026
                    Case 27 : Return pic0027
                    Case 28 : Return pic0028
                    Case 29 : Return pic0029
                End Select
            Case 1 : Select Case col_no
                    Case 0 : Return pic0100
                    Case 1 : Return pic0101
                    Case 2 : Return pic0102
                    Case 3 : Return pic0103
                    Case 4 : Return pic0104
                    Case 5 : Return pic0105
                    Case 6 : Return pic0106
                    Case 7 : Return pic0107
                    Case 8 : Return pic0108
                    Case 9 : Return pic0109
                    Case 10 : Return pic0110
                    Case 11 : Return pic0111
                    Case 12 : Return pic0112
                    Case 13 : Return pic0113
                    Case 14 : Return pic0114
                    Case 15 : Return pic0115
                    Case 16 : Return pic0116
                    Case 17 : Return pic0117
                    Case 18 : Return pic0118
                    Case 19 : Return pic0119
                    Case 20 : Return pic0120
                    Case 21 : Return pic0121
                    Case 22 : Return pic0122
                    Case 23 : Return pic0123
                    Case 24 : Return pic0124
                    Case 25 : Return pic0125
                    Case 26 : Return pic0126
                    Case 27 : Return pic0127
                    Case 28 : Return pic0128
                    Case 29 : Return pic0129
                End Select
            Case 2 : Select Case col_no
                    Case 0 : Return pic0200
                    Case 1 : Return pic0201
                    Case 2 : Return pic0202
                    Case 3 : Return pic0203
                    Case 4 : Return pic0204
                    Case 5 : Return pic0205
                    Case 6 : Return pic0206
                    Case 7 : Return pic0207
                    Case 8 : Return pic0208
                    Case 9 : Return pic0209
                    Case 10 : Return pic0210
                    Case 11 : Return pic0211
                    Case 12 : Return pic0212
                    Case 13 : Return pic0213
                    Case 14 : Return pic0214
                    Case 15 : Return pic0215
                    Case 16 : Return pic0216
                    Case 17 : Return pic0217
                    Case 18 : Return pic0218
                    Case 19 : Return pic0219
                    Case 20 : Return pic0220
                    Case 21 : Return pic0221
                    Case 22 : Return pic0222
                    Case 23 : Return pic0223
                    Case 24 : Return pic0224
                    Case 25 : Return pic0225
                    Case 26 : Return pic0226
                    Case 27 : Return pic0227
                    Case 28 : Return pic0228
                    Case 29 : Return pic0229
                End Select
            Case 3 : Select Case col_no
                    Case 0 : Return pic0300
                    Case 1 : Return pic0301
                    Case 2 : Return pic0302
                    Case 3 : Return pic0303
                    Case 4 : Return pic0304
                    Case 5 : Return pic0305
                    Case 6 : Return pic0306
                    Case 7 : Return pic0307
                    Case 8 : Return pic0308
                    Case 9 : Return pic0309
                    Case 10 : Return pic0310
                    Case 11 : Return pic0311
                    Case 12 : Return pic0312
                    Case 13 : Return pic0313
                    Case 14 : Return pic0314
                    Case 15 : Return pic0315
                    Case 16 : Return pic0316
                    Case 17 : Return pic0317
                    Case 18 : Return pic0318
                    Case 19 : Return pic0319
                    Case 20 : Return pic0320
                    Case 21 : Return pic0321
                    Case 22 : Return pic0322
                    Case 23 : Return pic0323
                    Case 24 : Return pic0324
                    Case 25 : Return pic0325
                    Case 26 : Return pic0326
                    Case 27 : Return pic0327
                    Case 28 : Return pic0328
                    Case 29 : Return pic0329
                End Select
            Case 4 : Select Case col_no
                    Case 0 : Return pic0400
                    Case 1 : Return pic0401
                    Case 2 : Return pic0402
                    Case 3 : Return pic0403
                    Case 4 : Return pic0404
                    Case 5 : Return pic0405
                    Case 6 : Return pic0406
                    Case 7 : Return pic0407
                    Case 8 : Return pic0408
                    Case 9 : Return pic0409
                    Case 10 : Return pic0410
                    Case 11 : Return pic0411
                    Case 12 : Return pic0412
                    Case 13 : Return pic0413
                    Case 14 : Return pic0414
                    Case 15 : Return pic0415
                    Case 16 : Return pic0416
                    Case 17 : Return pic0417
                    Case 18 : Return pic0418
                    Case 19 : Return pic0419
                    Case 20 : Return pic0420
                    Case 21 : Return pic0421
                    Case 22 : Return pic0422
                    Case 23 : Return pic0423
                    Case 24 : Return pic0424
                    Case 25 : Return pic0425
                    Case 26 : Return pic0426
                    Case 27 : Return pic0427
                    Case 28 : Return pic0428
                    Case 29 : Return pic0429
                End Select
            Case 5 : Select Case col_no
                    Case 0 : Return pic0500
                    Case 1 : Return pic0501
                    Case 2 : Return pic0502
                    Case 3 : Return pic0503
                    Case 4 : Return pic0504
                    Case 5 : Return pic0505
                    Case 6 : Return pic0506
                    Case 7 : Return pic0507
                    Case 8 : Return pic0508
                    Case 9 : Return pic0509
                    Case 10 : Return pic0510
                    Case 11 : Return pic0511
                    Case 12 : Return pic0512
                    Case 13 : Return pic0513
                    Case 14 : Return pic0514
                    Case 15 : Return pic0515
                    Case 16 : Return pic0516
                    Case 17 : Return pic0517
                    Case 18 : Return pic0518
                    Case 19 : Return pic0519
                    Case 20 : Return pic0520
                    Case 21 : Return pic0521
                    Case 22 : Return pic0522
                    Case 23 : Return pic0523
                    Case 24 : Return pic0524
                    Case 25 : Return pic0525
                    Case 26 : Return pic0526
                    Case 27 : Return pic0527
                    Case 28 : Return pic0528
                    Case 29 : Return pic0529

                End Select
            Case 6 : Select Case col_no
                    Case 0 : Return pic0600
                    Case 1 : Return pic0601
                    Case 2 : Return pic0602
                    Case 3 : Return pic0603
                    Case 4 : Return pic0604
                    Case 5 : Return pic0605
                    Case 6 : Return pic0606
                    Case 7 : Return pic0607
                    Case 8 : Return pic0608
                    Case 9 : Return pic0609
                    Case 10 : Return pic0610
                    Case 11 : Return pic0611
                    Case 12 : Return pic0612
                    Case 13 : Return pic0613
                    Case 14 : Return pic0614
                    Case 15 : Return pic0615
                    Case 16 : Return pic0616
                    Case 17 : Return pic0617
                    Case 18 : Return pic0618
                    Case 19 : Return pic0619
                    Case 20 : Return pic0620
                    Case 21 : Return pic0621
                    Case 22 : Return pic0622
                    Case 23 : Return pic0623
                    Case 24 : Return pic0624
                    Case 25 : Return pic0625
                    Case 26 : Return pic0626
                    Case 27 : Return pic0627
                    Case 28 : Return pic0628
                    Case 29 : Return pic0629

                End Select
            Case 7 : Select Case col_no
                    Case 0 : Return pic0700
                    Case 1 : Return pic0701
                    Case 2 : Return pic0702
                    Case 3 : Return pic0703
                    Case 4 : Return pic0704
                    Case 5 : Return pic0705
                    Case 6 : Return pic0706
                    Case 7 : Return pic0707
                    Case 8 : Return pic0708
                    Case 9 : Return pic0709
                    Case 10 : Return pic0710
                    Case 11 : Return pic0711
                    Case 12 : Return pic0712
                    Case 13 : Return pic0713
                    Case 14 : Return pic0714
                    Case 15 : Return pic0715
                    Case 16 : Return pic0716
                    Case 17 : Return pic0717
                    Case 18 : Return pic0718
                    Case 19 : Return pic0719
                    Case 20 : Return pic0720
                    Case 21 : Return pic0721
                    Case 22 : Return pic0722
                    Case 23 : Return pic0723
                    Case 24 : Return pic0724
                    Case 25 : Return pic0725
                    Case 26 : Return pic0726
                    Case 27 : Return pic0727
                    Case 28 : Return pic0728
                    Case 29 : Return pic0729

                End Select
            Case 8 : Select Case col_no
                    Case 0 : Return pic0800
                    Case 1 : Return pic0801
                    Case 2 : Return pic0802
                    Case 3 : Return pic0803
                    Case 4 : Return pic0804
                    Case 5 : Return pic0805
                    Case 6 : Return pic0806
                    Case 7 : Return pic0807
                    Case 8 : Return pic0808
                    Case 9 : Return pic0809
                    Case 10 : Return pic0810
                    Case 11 : Return pic0811
                    Case 12 : Return pic0812
                    Case 13 : Return pic0813
                    Case 14 : Return pic0814
                    Case 15 : Return pic0815
                    Case 16 : Return pic0816
                    Case 17 : Return pic0817
                    Case 18 : Return pic0818
                    Case 19 : Return pic0819
                    Case 20 : Return pic0820
                    Case 21 : Return pic0821
                    Case 22 : Return pic0822
                    Case 23 : Return pic0823
                    Case 24 : Return pic0824
                    Case 25 : Return pic0825
                    Case 26 : Return pic0826
                    Case 27 : Return pic0827
                    Case 28 : Return pic0828
                    Case 29 : Return pic0829

                End Select
            Case 9 : Select Case col_no
                    Case 0 : Return pic0900
                    Case 1 : Return pic0901
                    Case 2 : Return pic0902
                    Case 3 : Return pic0903
                    Case 4 : Return pic0904
                    Case 5 : Return pic0905
                    Case 6 : Return pic0906
                    Case 7 : Return pic0907
                    Case 8 : Return pic0908
                    Case 9 : Return pic0909
                    Case 10 : Return pic0910
                    Case 11 : Return pic0911
                    Case 12 : Return pic0912
                    Case 13 : Return pic0913
                    Case 14 : Return pic0914
                    Case 15 : Return pic0915
                    Case 16 : Return pic0916
                    Case 17 : Return pic0917
                    Case 18 : Return pic0918
                    Case 19 : Return pic0919
                    Case 20 : Return pic0920
                    Case 21 : Return pic0921
                    Case 22 : Return pic0922
                    Case 23 : Return pic0923
                    Case 24 : Return pic0924
                    Case 25 : Return pic0925
                    Case 26 : Return pic0926
                    Case 27 : Return pic0927
                    Case 28 : Return pic0928
                    Case 29 : Return pic0929

                End Select
            Case 10 : Select Case col_no
                    Case 0 : Return pic1000
                    Case 1 : Return pic1001
                    Case 2 : Return pic1002
                    Case 3 : Return pic1003
                    Case 4 : Return pic1004
                    Case 5 : Return pic1005
                    Case 6 : Return pic1006
                    Case 7 : Return pic1007
                    Case 8 : Return pic1008
                    Case 9 : Return pic1009
                    Case 10 : Return pic1010
                    Case 11 : Return pic1011
                    Case 12 : Return pic1012
                    Case 13 : Return pic1013
                    Case 14 : Return pic1014
                    Case 15 : Return pic1015
                    Case 16 : Return pic1016
                    Case 17 : Return pic1017
                    Case 18 : Return pic1018
                    Case 19 : Return pic1019
                    Case 20 : Return pic1020
                    Case 21 : Return pic1021
                    Case 22 : Return pic1022
                    Case 23 : Return pic1023
                    Case 24 : Return pic1024
                    Case 25 : Return pic1025
                    Case 26 : Return pic1026
                    Case 27 : Return pic1027
                    Case 28 : Return pic1028
                    Case 29 : Return pic1029

                End Select
            Case 11 : Select Case col_no
                    Case 0 : Return pic1100
                    Case 1 : Return pic1101
                    Case 2 : Return pic1102
                    Case 3 : Return pic1103
                    Case 4 : Return pic1104
                    Case 5 : Return pic1105
                    Case 6 : Return pic1106
                    Case 7 : Return pic1107
                    Case 8 : Return pic1108
                    Case 9 : Return pic1109
                    Case 10 : Return pic1110
                    Case 11 : Return pic1111
                    Case 12 : Return pic1112
                    Case 13 : Return pic1113
                    Case 14 : Return pic1114
                    Case 15 : Return pic1115
                    Case 16 : Return pic1116
                    Case 17 : Return pic1117
                    Case 18 : Return pic1118
                    Case 19 : Return pic1119
                    Case 20 : Return pic1120
                    Case 21 : Return pic1121
                    Case 22 : Return pic1122
                    Case 23 : Return pic1123
                    Case 24 : Return pic1124
                    Case 25 : Return pic1125
                    Case 26 : Return pic1126
                    Case 27 : Return pic1127
                    Case 28 : Return pic1128
                    Case 29 : Return pic1129

                End Select
            Case 12 : Select Case col_no
                    Case 0 : Return pic1200
                    Case 1 : Return pic1201
                    Case 2 : Return pic1202
                    Case 3 : Return pic1203
                    Case 4 : Return pic1204
                    Case 5 : Return pic1205
                    Case 6 : Return pic1206
                    Case 7 : Return pic1207
                    Case 8 : Return pic1208
                    Case 9 : Return pic1209
                    Case 10 : Return pic1210
                    Case 11 : Return pic1211
                    Case 12 : Return pic1212
                    Case 13 : Return pic1213
                    Case 14 : Return pic1214
                    Case 15 : Return pic1215
                    Case 16 : Return pic1216
                    Case 17 : Return pic1217
                    Case 18 : Return pic1218
                    Case 19 : Return pic1219
                    Case 20 : Return pic1220
                    Case 21 : Return pic1221
                    Case 22 : Return pic1222
                    Case 23 : Return pic1223
                    Case 24 : Return pic1224
                    Case 25 : Return pic1225
                    Case 26 : Return pic1226
                    Case 27 : Return pic1227
                    Case 28 : Return pic1228
                    Case 29 : Return pic1229

                End Select
            Case 13 : Select Case col_no
                    Case 0 : Return pic1300
                    Case 1 : Return pic1301
                    Case 2 : Return pic1302
                    Case 3 : Return pic1303
                    Case 4 : Return pic1304
                    Case 5 : Return pic1305
                    Case 6 : Return pic1306
                    Case 7 : Return pic1307
                    Case 8 : Return pic1308
                    Case 9 : Return pic1309
                    Case 10 : Return pic1310
                    Case 11 : Return pic1311
                    Case 12 : Return pic1312
                    Case 13 : Return pic1313
                    Case 14 : Return pic1314
                    Case 15 : Return pic1315
                    Case 16 : Return pic1316
                    Case 17 : Return pic1317
                    Case 18 : Return pic1318
                    Case 19 : Return pic1319
                    Case 20 : Return pic1320
                    Case 21 : Return pic1321
                    Case 22 : Return pic1322
                    Case 23 : Return pic1323
                    Case 24 : Return pic1324
                    Case 25 : Return pic1325
                    Case 26 : Return pic1326
                    Case 27 : Return pic1327
                    Case 28 : Return pic1328
                    Case 29 : Return pic1329

                End Select
            Case 14 : Select Case col_no
                    Case 0 : Return pic1400
                    Case 1 : Return pic1401
                    Case 2 : Return pic1402
                    Case 3 : Return pic1403
                    Case 4 : Return pic1404
                    Case 5 : Return pic1405
                    Case 6 : Return pic1406
                    Case 7 : Return pic1407
                    Case 8 : Return pic1408
                    Case 9 : Return pic1409
                    Case 10 : Return pic1410
                    Case 11 : Return pic1411
                    Case 12 : Return pic1412
                    Case 13 : Return pic1413
                    Case 14 : Return pic1414
                    Case 15 : Return pic1415
                    Case 16 : Return pic1416
                    Case 17 : Return pic1417
                    Case 18 : Return pic1418
                    Case 19 : Return pic1419
                    Case 20 : Return pic1420
                    Case 21 : Return pic1421
                    Case 22 : Return pic1422
                    Case 23 : Return pic1423
                    Case 24 : Return pic1424
                    Case 25 : Return pic1425
                    Case 26 : Return pic1426
                    Case 27 : Return pic1427
                    Case 28 : Return pic1428
                    Case 29 : Return pic1429

                End Select
            Case 15 : Select Case col_no
                    Case 0 : Return pic1500
                    Case 1 : Return pic1501
                    Case 2 : Return pic1502
                    Case 3 : Return pic1503
                    Case 4 : Return pic1504
                    Case 5 : Return pic1505
                    Case 6 : Return pic1506
                    Case 7 : Return pic1507
                    Case 8 : Return pic1508
                    Case 9 : Return pic1509
                    Case 10 : Return pic1510
                    Case 11 : Return pic1511
                    Case 12 : Return pic1512
                    Case 13 : Return pic1513
                    Case 14 : Return pic1514
                    Case 15 : Return pic1515
                    Case 16 : Return pic1516
                    Case 17 : Return pic1517
                    Case 18 : Return pic1518
                    Case 19 : Return pic1519
                    Case 20 : Return pic1520
                    Case 21 : Return pic1521
                    Case 22 : Return pic1522
                    Case 23 : Return pic1523
                    Case 24 : Return pic1524
                    Case 25 : Return pic1525
                    Case 26 : Return pic1526
                    Case 27 : Return pic1527
                    Case 28 : Return pic1528
                    Case 29 : Return pic1529
                End Select
            Case 16 : Select Case col_no
                    Case 0 : Return pic1600
                    Case 1 : Return pic1601
                    Case 2 : Return pic1602
                    Case 3 : Return pic1603
                    Case 4 : Return pic1604
                    Case 5 : Return pic1605
                    Case 6 : Return pic1606
                    Case 7 : Return pic1607
                    Case 8 : Return pic1608
                    Case 9 : Return pic1609
                    Case 10 : Return pic1610
                    Case 11 : Return pic1611
                    Case 12 : Return pic1612
                    Case 13 : Return pic1613
                    Case 14 : Return pic1614
                    Case 15 : Return pic1615
                    Case 16 : Return pic1616
                    Case 17 : Return pic1617
                    Case 18 : Return pic1618
                    Case 19 : Return pic1619
                    Case 20 : Return pic1620
                    Case 21 : Return pic1621
                    Case 22 : Return pic1622
                    Case 23 : Return pic1623
                    Case 24 : Return pic1624
                    Case 25 : Return pic1625
                    Case 26 : Return pic1626
                    Case 27 : Return pic1627
                    Case 28 : Return pic1628
                    Case 29 : Return pic1629
                End Select

        End Select
        e = ElementGet2(row_no, col_no)
        Return Nothing
    End Function
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        CloseShop()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim v As Integer
        Dim cmdline As String = Environment.CommandLine()
        Dim cmdline_params As String, exe_pos As Integer
        Message(1, "Test output level=" & Format(Test_Level))
        Message(1, "Command line=" & cmdline)

        STAND_ALONE = True
        CMDLINE_MODE = ""
        Cmdline_Param1 = ""
        Cmdline_Param2 = ""

        exe_pos = InStr(cmdline, ".exe")
        cmdline_params = Trim(Mid(cmdline, exe_pos + 5))
        Message(1, "params=" & cmdline_params)
        If cmdline_params <> "" Then
            Cmdline_Param1 = Trim(ExtractField2(cmdline_params, 1, " "))
            Cmdline_Param2 = Trim(ExtractField2(cmdline_params, 2, " "))
            Message(1, ">" & Cmdline_Param1 & "< ")
            Message(1, ">" & Cmdline_Param2 & "< ")
            Select Case Cmdline_Param1
                Case "Edit"
                    If FileExists(Cmdline_Param2) Then
                        CMDLINE_MODE = "Edit"
                        STAND_ALONE = False
                    Else
                        MsgBox("Cannot locate file on command line: " & Cmdline_Param2)
                        MsgBox("Activating stand-alone mode")
                    End If
                Case Else
                    MsgBox("Bad command line parameter: " & Cmdline_Param1)
                    MsgBox("Activating stand-alone mode")
            End Select
        End If
        If STAND_ALONE Then Message("Stand alone mode")

        Start_Directory = CurDir()
        Message(1, "Start directory=" & Start_Directory)
        env1.InitEnv(Start_Directory & "\" & "TrackDiagramEditor.cfg")

        Me.Text = "Track Diagram Editor  .NET" & "  [" & Version & "]"

        If STAND_ALONE Then
            Current_Directory = env1.GetEnv("Current_Directory")
            If Current_Directory = "" Then Current_Directory = Start_Directory
            TrackDiagram_Directory = env1.GetEnv("TrackDiagram_Directory")
            If TrackDiagram_Directory = "" Then TrackDiagram_Directory = Current_Directory
            TrackDiagram_Input_Filename = env1.GetEnv("TrackDiagram_Input_Filename")
            If TrackDiagram_Input_Filename = "" Then TrackDiagram_Input_Filename = "*.cs2"
        ElseIf CMDLINE_MODE = "Edit" Then
            TrackDiagram_Input_Filename = Cmdline_Param2
            TrackDiagram_Directory = FilenamePath(Cmdline_Param2)
            GleisBildDisplayClear()
            GleisbildLoad(TrackDiagram_Input_Filename)
            GleisbildDisplay()
        Else

        End If

        Use_English_Terms = (env1.GetEnv("Use_English_Terms") <> "False")
        UseEnglishTermsForElementsToolStripMenuItem.Checked = Use_English_Terms

        Verbose = (env1.GetEnv("Verbose") = "True")
        VerboseToolStripMenuItem.Checked = Verbose

        Developer_Environment = (InStr(Start_Directory, "\OleMain\") > 0)

        Initialize()

        Me.Top = Val(env1.GetEnv("Top"))
        Me.Left = Val(env1.GetEnv("Left"))
        v = Val(env1.GetEnv("Width"))
        If v >= W_MIN Then Me.Width = v
        v = Val(env1.GetEnv("Height"))
        If v > H_MIN Then
            Me.Height = v
        Else
            Me.Height = H_MIN
        End If
        Me.WindowState = Val(env1.GetEnv("WindowState"))

        If Developer_Environment Then
            Test_Level = Val(env1.GetEnv("Test_Level"))
        Else
            Test_Level = 0
        End If
        Message(1, "Test level=" & Format(Test_Level))
        AdjustPage()
        If Not STAND_ALONE Then Call cmdEdit_Click(sender, e)
    End Sub

    Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        AdjustPage()
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Static old_state As Integer ' 0: window, 1: minimized, 2: maximized
        If old_state <> Me.WindowState Then
            If Me.WindowState <> 1 Then AdjustPage()
            old_state = Me.WindowState
        End If
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        CloseShop()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        GleisBildDisplayClear()
        frmHelp.HelpWindowClose()
        GleisbildLoad("")
        If ROUTE_VIEW_ENABLED AndAlso cboRouteNames.Text <> "" AndAlso cboRouteNames.Text <> "Pick a route" Then
            RH.RouteShow(cboRouteNames.Text)
        Else
            GleisbildDisplay()
        End If
    End Sub

    Private Sub SetTestLevelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetTestLevelToolStripMenuItem.Click
        Test_Level = Val(InputBox("Test level (´0...5)", "Test Level", Format(Test_Level)))
        Message("Test level=" & Format(Test_Level))
        AdjustPage()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        GleisbildSave()
    End Sub

    Private Sub SaveCopyAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        GleisbildSaveCopyAs()
    End Sub

    Private Sub ReopenLastToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReopenLastToolStripMenuItem.Click
        GleisBildDisplayClear()
        frmHelp.HelpWindowClose()
        GleisbildLoad(TrackDiagram_Input_Filename)
        GleisbildDisplay()
    End Sub
    Private Sub picClick(picName As String)
        Dim el As ELEMENT, el_new As ELEMENT, dig_code As String, route_name As String = "", text_str As String = ""
        Dim row_clicked As Integer, col_clicked As Integer
        Dim s As String, user_input As Object

        grpMoveWindow.Enabled = True
        grpIcons.Enabled = True
        grpEditActions.Enabled = True

        row_clicked = Val(Mid(picName, 4, 2))
        col_clicked = Val(Mid(picName, 6, 2))

        el = ElementGet2(row_clicked, col_clicked)
        Message(2, "picClick(" & picName & ") " & Active_Tool & " el.id=" & el.id & " (" & Format(el.row_no) & "," &
                Format(el.col_no) & "), my_index=" & Format(el.my_index) & ", typ=" & el.typ_s)

        If MODE = "View" Then
            If Test_Level > 2 Then
                If el.typ_s <> "" Or el.text <> "" Then
                    If el.artikel_i >= 0 Then
                        dig_code = ", " & DigitalAddress(el)
                    Else
                        dig_code = ""
                    End If
                    If el.text <> "" Then text_str = ", text=" & el.text
                    Message(el.id & " (" & Format(el.row_no) & "," & Format(el.col_no) & "): " & el.typ_s & dig_code & ", drehung=" & Format(el.drehung_i) & text_str)
                End If
            End If
            PropertiesDisplay(el)
        ElseIf MODE = "Routes" Then ' the properties pane isn't available so we show the pertinent data as a message
            If el.typ_s <> "" Then
                If el.artikel_i >= 0 Then
                    dig_code = DigitalAddress(el)
                    If el.typ_s = "fahrstrasse" And el.artikel_i > 0 Then
                        route_name = RH.RouteName(el.artikel_i)
                        '      RH.RouteShow(route_name)
                        cboRouteNames.Text = route_name ' side effect: RH.RouteShow is called
                    End If
                    Message(1, Translate(el.typ_s) & " (" & Format(el.row_no) & "," & Format(el.col_no) & "): address=" & dig_code & " " & route_name)
                Else
                    Message(1, Translate(el.typ_s) & " (" & Format(el.row_no) & "," & Format(el.col_no) & ")")
                End If
            End If
        ElseIf Active_Tool <> "" Then ' MODE is "Edit"
            Select Case Active_Tool
                Case "edit_address" : If HasDigitalAddress(el.typ_s) Then
                        el.artikel_i = DigitalArtikelPrompt(el)
                        If Strings.Left(el.typ_s, 3) = "s88" And el.artikel_i >= 1000 Then
                            el.deviceId_s = DeviceIdPrompt(el.deviceId_s)
                        End If
                        Elements(el.my_index) = el
                        Reveal_Addr = True ' to counter that the user may have turned it off
                        Reveal_Text = False
                        ElementDisplay(el)
                        EDIT_CHANGES = True
                    End If
                Case "edit_text"
                    user_input = ""
                    If InputPromptOK(user_input, "Enter new text", el.text) Then
                        If el.typ_s = "" Then ' el is a CS2 text object 
                            If Trim(user_input) = "" Then ' text cannot be left blank in a text element due to Märklin's convention of leaving the type designation blank
                                MsgBox("Text cannot be blank in a text object. The text object was deleted")
                                ElementRemove(el.my_index)
                                ElementDisplay(PicGet(el.row_no, el.col_no), el.drehung_i, picIconEmpty.Image)
                                EDIT_CHANGES = True
                            Else
                                el.text = user_input
                                Elements(el.my_index) = el
                                Reveal_Text = True ' to counter that the user may have turned it off
                                Reveal_Addr = False
                                ElementDisplay(el)
                                EDIT_CHANGES = True
                            End If
                        Else ' allow for text in non-text objects - like CS2 allows for
                            el.text = user_input
                            Elements(el.my_index) = el
                            Reveal_Text = True ' to counter that the user may have turned it off
                            Reveal_Addr = False
                            ElementDisplay(el)
                            EDIT_CHANGES = True
                        End If
                    Else
                        ' User hit the cancel button or entered empty text. No action to be taken
                    End If
                Case "empty"
                    ElementRemove(el.my_index)
                    ElementDisplay(PicGet(el.row_no, el.col_no), el.drehung_i, picIconEmpty.Image)
                Case "rotate" ' NB, some elements should not be rotated, like cross. Some should only rotate once
                    If el.typ_s <> "kreuzung" And el.typ_s <> "fahrstrasse" And el.typ_s <> "drehscheibe" And el.typ_s <> "lampe" _
                        And Strings.Left(el.typ_s, 3) <> "std" And el.typ_s <> "" Then
                        el.drehung_i = el.drehung_i + 1
                        If el.drehung_i > 3 OrElse
                            ((el.typ_s = "gerade" Or el.typ_s = "dkweiche" Or el.typ_s = "doppelbogen" Or el.typ_s = "entkuppler" Or el.typ_s = "s88kontakt" Or el.typ_s = "unterfuehrung") _
                             And el.drehung_i > 1) _
                            Then
                            el.drehung_i = 0
                        End If
                        Elements(el.my_index) = el
                        EDIT_CHANGES = True
                        GleisbildDisplay()
                        If Test_Level > 1 Then Message("drehung now=" & Format(el.drehung_i))
                    End If
                Case "move_element"
                    If el.typ_s <> "" Or el.text <> "" Then ' i.e. no action if empty cell
                        el_moving = el
                        ElementDisplay(PicGet(el.row_no, el.col_no), el.drehung_i, picIconMoveElement.Image)
                        grpMoveWindow.Enabled = False
                        grpIcons.Enabled = False
                        grpEditActions.Enabled = False
                        Active_Tool = "move_element_step2"
                    End If
                Case "move_element_step2"
                    ' first remove the element being moved from the old location
                    ElementRemove(el_moving.my_index)
                    ' Then remove the element (if any) in the new location
                    ElementRemove(ElementGet2(row_clicked, col_clicked).my_index)
                    ' Update location and insert the new element
                    el_moving.row_no = row_clicked
                    el_moving.col_no = col_clicked
                    s = DecIntToHexStr(Page_No) & DecIntToHexStr(el_moving.row_no + Row_Offset) & DecIntToHexStr(el_moving.col_no + Col_Offset)
                    el_moving.id_normalized = s
                    el_moving.id = CS2_id_create(Page_No, el_moving.row_no + Row_Offset, el_moving.col_no + Col_Offset)
                    E_TOP = E_TOP + 1
                    el_moving.my_index = E_TOP
                    Elements(el_moving.my_index) = el_moving
                    Message(2, "Move inserted element: " & el_moving.typ_s & " @ " & el_moving.id_normalized _
                                 & " (" & Format(el_moving.row_no) & "," & Format(el_moving.col_no) & "), E_TOP=" & Format(E_TOP))
                    el_moving = Empty_Element
                    EDIT_CHANGES = True
                    GleisBildDisplayClear()
                    GleisbildDisplay()
                    Active_Tool = "move_element"
                Case "split_horizontally"
                    EDIT_CHANGES = EDIT_CHANGES Or SplitHorizontally(el.row_no)
                    GleisBildDisplayClear()
                    GleisbildDisplay()
                Case "split_vertically"
                    EDIT_CHANGES = EDIT_CHANGES Or SplitVertically(el.col_no)
                    GleisBildDisplayClear()
                    GleisbildDisplay()
                Case Else
                    ElementRemove(el.my_index) ' no action if el is empty_element (i.e. the cell clicked is empty)
                    el_new = ElementInsert(picName, Active_Tool)
                    If Test_Level = 1 Then
                        Message("New element in (row,col)=(" & Format(el_new.row_no) & "," & Format(el_new.col_no) & "): " & Translate(el_new.typ_s))
                    Else
                        Message(2, "New element in (row,col)=(" & Format(el_new.row_no) & "," & Format(el_new.col_no) & "): " & el_new.typ_s & ", artikel=" & Format(el_new.artikel_i) & ", drehung=" & Format(el_new.drehung_i) & ", E_Top=" & Format(E_TOP))
                    End If
                    Reveal_Addr = (el_new.text = "")
                    Reveal_Text = Not Reveal_Addr
                    ElementDisplay(el_new)
            End Select
        End If
    End Sub
    Private Sub picHover(picObject As Object)
        Dim picName As String = picObject.name
        '        Static picName_old As String
        Dim el As ELEMENT
        Dim row_clicked As Integer, col_clicked As Integer
        If FILE_LOADED And (MODE = "View" Or MODE = "Routes" Or MODE = "Edit") Then
            row_clicked = Val(Mid(picName, 4, 2))
            col_clicked = Val(Mid(picName, 6, 2))

            el = ElementGet2(row_clicked, col_clicked)

            If el.typ_s <> "" AndAlso el.artikel_i >= 0 Then
                tt.Show(Translate(el.typ_s) & " " & DigitalAddress(el), picObject)
            ElseIf el.text <> "" Then
                tt.Show(el.text, picObject)
            End If
        End If
    End Sub
    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        If FILE_LOADED Then
            ElementsStore()
            MODE = "Edit"
            Message(1, "Mode: Edit Track Diagram")
            cmdCancel.Text = "Cancel"
            Reveal_Addr_State = Reveal_Addr
            Reveal_Text_State = Reveal_Text
            FileToolStripMenuItem.Enabled = False
            ToolsToolStripMenuItem.Enabled = False
            AllPagesToolStripMenuItem2.Enabled = False
            AllPagesToolStripMenuItem3.Enabled = False
            QuitToolStripMenuItem.Enabled = False

            EDIT_CHANGES = False
            GleisbildGridSet(True)
            AdjustPage()
            Message("Click an icon or action in the righthand pane, then click a square in the layout to apply")
        Else
            MsgBox("No layout has been loaded")
        End If
    End Sub
    Private Sub cmdRoutesView_Click(sender As Object, e As EventArgs) Handles cmdRoutesView.Click
        If FILE_LOADED Then
            If RH.FahrStrassenLoad() Then
                MODE = "Routes"
                Message(1, "Mode: View Routes")
                cmdCancel.Text = "Back"
                Reveal_Addr_State = Reveal_Addr
                Reveal_Text_State = Reveal_Text
                Reveal_Addr = True
                Reveal_Text = False

                NewLayoutToolStripMenuItem.Enabled = False
                ReopenLastToolStripMenuItem.Enabled = False
                SaveToolStripMenuItem.Enabled = False
                SaveAsToolStripMenuItem.Enabled = False
                ToolsToolStripMenuItem.Enabled = False
                AllPagesToolStripMenuItem2.Enabled = False
                AllPagesToolStripMenuItem3.Enabled = False

                AdjustPage()
                GleisbildDisplay()
                Message("Select a route in the drop-down upper right corner or click a route in the track diagram")
            End If
        Else
            MsgBox("No layout has been loaded")
        End If
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        If MODE = "Edit" And EDIT_CHANGES Then
            ElementsRecall()
            EDIT_CHANGES = False
        End If

        If Not STAND_ALONE And Not (MODE = "Routes" Or MODE = "Join") Then CloseShop()
        grpMoveWindow.Enabled = True
        grpIcons.Enabled = True
        grpEditActions.Enabled = True


        If MODE = "Routes" Or MODE = "Edit" Then
            Reveal_Addr = Reveal_Addr_State
            Reveal_Text = Reveal_Text_State
        End If

        MODE = "View"
        Message(1, "Mode: View/Manage Pages")
        Message("Click an element to see its properties")
        el_moving = Empty_Element
        FileToolStripMenuItem.Enabled = True
        NewLayoutToolStripMenuItem.Enabled = True
        ReopenLastToolStripMenuItem.Enabled = True
        SaveToolStripMenuItem.Enabled = True
        SaveAsToolStripMenuItem.Enabled = True
        ToolsToolStripMenuItem.Enabled = True
        RefreshToolStripMenuItem.Enabled = True
        ShowOutputFileToolStripMenuItem.Enabled = True
        UseEnglishTermsForElementsToolStripMenuItem.Enabled = True
        ConsistencyCheckToolStripMenuItem.Enabled = True
        txtCombine.Visible = False

        cboRouteNames.Text = ""

        AllPagesToolStripMenuItem2.Enabled = True
        AllPagesToolStripMenuItem3.Enabled = True
        QuitToolStripMenuItem.Enabled = True
        Active_Tool = ""
        AdjustPage()
        GleisbildGridSet(False)
        GleisBildDisplayClear()
        GleisbildDisplay()
    End Sub

    Private Sub cmdCommit_Click(sender As Object, e As EventArgs) Handles cmdCommit.Click
        MODE = "View"
        Message(1, "Mode: View/Manage Pages")
        Message("Click an element to see its properties")

        grpMoveWindow.Enabled = True
        grpIcons.Enabled = True
        grpEditActions.Enabled = True

        el_moving = Empty_Element
        FileToolStripMenuItem.Enabled = True
        ToolsToolStripMenuItem.Enabled = True
        AllPagesToolStripMenuItem2.Enabled = True
        AllPagesToolStripMenuItem3.Enabled = True
        QuitToolStripMenuItem.Enabled = True
        CHANGED = CHANGED Or EDIT_CHANGES
        EDIT_CHANGES = False
        Active_Tool = ""
        If STAND_ALONE Then
            Reveal_Addr = Reveal_Addr_State
            Reveal_Text = Reveal_Text_State
            AdjustPage()
            GleisbildGridSet(False)
            GleisBildDisplayClear()
            GleisbildDisplay()
        Else
            CloseShop()
        End If
    End Sub
    Private Sub picIconStraight_Click(sender As Object, e As EventArgs) Handles picIconStraight.Click
        Active_Tool = "straight"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub picIconTunnel_Click(sender As Object, e As EventArgs) Handles picIconTunnel.Click
        Active_Tool = "tunnel"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconCrossSwitch_Click(sender As Object, e As EventArgs) Handles picIconCrossSwitch.Click
        Active_Tool = "cross_switch"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconEmpty_Click(sender As Object, e As EventArgs)
        Active_Tool = "empty"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        lstOutput.Visible = False
        GleisBildDisplayClear()
        GleisbildDisplay()
    End Sub

    Private Sub pic0000_Click(sender As Object, e As EventArgs) Handles pic0000.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0001_Click(sender As Object, e As EventArgs) Handles pic0001.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0002_Click(sender As Object, e As EventArgs) Handles pic0002.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0003_Click(sender As Object, e As EventArgs) Handles pic0003.Click
        picClick(sender.Name)
    End Sub
    Private Sub pic0004_Click(sender As Object, e As EventArgs) Handles pic0004.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0005_Click(sender As Object, e As EventArgs) Handles pic0005.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0006_Click(sender As Object, e As EventArgs) Handles pic0006.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0007_Click(sender As Object, e As EventArgs) Handles pic0007.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0008_Click(sender As Object, e As EventArgs) Handles pic0008.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0009_Click(sender As Object, e As EventArgs) Handles pic0009.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0010_Click(sender As Object, e As EventArgs) Handles pic0010.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0011_Click(sender As Object, e As EventArgs) Handles pic0011.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0012_Click(sender As Object, e As EventArgs) Handles pic0012.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0013_Click(sender As Object, e As EventArgs) Handles pic0013.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0014_Click(sender As Object, e As EventArgs) Handles pic0014.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0015_Click(sender As Object, e As EventArgs) Handles pic0015.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0016_Click(sender As Object, e As EventArgs) Handles pic0016.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0017_Click(sender As Object, e As EventArgs) Handles pic0017.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0018_Click(sender As Object, e As EventArgs) Handles pic0018.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0019_Click(sender As Object, e As EventArgs) Handles pic0019.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0020_Click(sender As Object, e As EventArgs) Handles pic0020.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0021_Click(sender As Object, e As EventArgs) Handles pic0021.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0022_Click(sender As Object, e As EventArgs) Handles pic0022.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0023_Click(sender As Object, e As EventArgs) Handles pic0023.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0024_Click(sender As Object, e As EventArgs) Handles pic0024.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0025_Click(sender As Object, e As EventArgs) Handles pic0025.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0026_Click(sender As Object, e As EventArgs) Handles pic0026.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0027_Click(sender As Object, e As EventArgs) Handles pic0027.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0028_Click(sender As Object, e As EventArgs) Handles pic0028.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0029_Click(sender As Object, e As EventArgs) Handles pic0029.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0100_Click(sender As Object, e As EventArgs) Handles pic0100.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0101_Click(sender As Object, e As EventArgs) Handles pic0101.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0102_Click(sender As Object, e As EventArgs) Handles pic0102.Click
        picClick(sender.Name)
    End Sub
    Private Sub pic0103_Click(sender As Object, e As EventArgs) Handles pic0103.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0104_Click(sender As Object, e As EventArgs) Handles pic0104.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0105_Click(sender As Object, e As EventArgs) Handles pic0105.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0106_Click(sender As Object, e As EventArgs) Handles pic0106.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0107_Click(sender As Object, e As EventArgs) Handles pic0107.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0108_Click(sender As Object, e As EventArgs) Handles pic0108.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0109_Click(sender As Object, e As EventArgs) Handles pic0109.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0110_Click(sender As Object, e As EventArgs) Handles pic0110.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0111_Click(sender As Object, e As EventArgs) Handles pic0111.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0112_Click(sender As Object, e As EventArgs) Handles pic0112.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0113_Click(sender As Object, e As EventArgs) Handles pic0113.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0114_Click(sender As Object, e As EventArgs) Handles pic0114.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0115_Click(sender As Object, e As EventArgs) Handles pic0115.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0116_Click(sender As Object, e As EventArgs) Handles pic0116.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0117_Click(sender As Object, e As EventArgs) Handles pic0117.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0118_Click(sender As Object, e As EventArgs) Handles pic0118.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0119_Click(sender As Object, e As EventArgs) Handles pic0119.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0120_Click(sender As Object, e As EventArgs) Handles pic0120.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0121_Click(sender As Object, e As EventArgs) Handles pic0121.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0122_Click(sender As Object, e As EventArgs) Handles pic0122.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0123_Click(sender As Object, e As EventArgs) Handles pic0123.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0124_Click(sender As Object, e As EventArgs) Handles pic0124.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0125_Click(sender As Object, e As EventArgs) Handles pic0125.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0126_Click(sender As Object, e As EventArgs) Handles pic0126.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0127_Click(sender As Object, e As EventArgs) Handles pic0127.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0128_Click(sender As Object, e As EventArgs) Handles pic0128.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0129_Click(sender As Object, e As EventArgs) Handles pic0129.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0200_Click(sender As Object, e As EventArgs) Handles pic0200.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0201_Click(sender As Object, e As EventArgs) Handles pic0201.Click
        picClick(sender.Name)
    End Sub
    Private Sub pic0202_Click(sender As Object, e As EventArgs) Handles pic0202.Click
        picClick(sender.Name)
    End Sub
    Private Sub pic0203_Click(sender As Object, e As EventArgs) Handles pic0203.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0204_Click(sender As Object, e As EventArgs) Handles pic0204.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0205_Click(sender As Object, e As EventArgs) Handles pic0205.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0206_Click(sender As Object, e As EventArgs) Handles pic0206.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0207_Click(sender As Object, e As EventArgs) Handles pic0207.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0208_Click(sender As Object, e As EventArgs) Handles pic0208.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0209_Click(sender As Object, e As EventArgs) Handles pic0209.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0210_Click(sender As Object, e As EventArgs) Handles pic0210.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0211_Click(sender As Object, e As EventArgs) Handles pic0211.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0212_Click(sender As Object, e As EventArgs) Handles pic0212.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0213_Click(sender As Object, e As EventArgs) Handles pic0213.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0214_Click(sender As Object, e As EventArgs) Handles pic0214.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0215_Click(sender As Object, e As EventArgs) Handles pic0215.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0216_Click(sender As Object, e As EventArgs) Handles pic0216.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0217_Click(sender As Object, e As EventArgs) Handles pic0217.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0218_Click(sender As Object, e As EventArgs) Handles pic0218.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0219_Click(sender As Object, e As EventArgs) Handles pic0219.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0220_Click(sender As Object, e As EventArgs) Handles pic0220.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0221_Click(sender As Object, e As EventArgs) Handles pic0221.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0222_Click(sender As Object, e As EventArgs) Handles pic0222.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0223_Click(sender As Object, e As EventArgs) Handles pic0223.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0224_Click(sender As Object, e As EventArgs) Handles pic0224.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0225_Click(sender As Object, e As EventArgs) Handles pic0225.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0226_Click(sender As Object, e As EventArgs) Handles pic0226.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0227_Click(sender As Object, e As EventArgs) Handles pic0227.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0228_Click(sender As Object, e As EventArgs) Handles pic0228.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0229_Click(sender As Object, e As EventArgs) Handles pic0229.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0300_Click(sender As Object, e As EventArgs) Handles pic0300.Click
        picClick(sender.Name)
    End Sub
    Private Sub pic0301_Click(sender As Object, e As EventArgs) Handles pic0301.Click
        picClick(sender.Name)
    End Sub
    Private Sub pic0302_Click(sender As Object, e As EventArgs) Handles pic0302.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0303_Click(sender As Object, e As EventArgs) Handles pic0303.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0304_Click(sender As Object, e As EventArgs) Handles pic0304.Click
        picClick(sender.Name)
    End Sub
    Private Sub pic0305_Click(sender As Object, e As EventArgs) Handles pic0305.Click
        picClick(pic0305.Name)
    End Sub

    Private Sub pic0306_Click(sender As Object, e As EventArgs) Handles pic0306.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0307_Click(sender As Object, e As EventArgs) Handles pic0307.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0308_Click(sender As Object, e As EventArgs) Handles pic0308.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0309_Click(sender As Object, e As EventArgs) Handles pic0309.Click
        picClick(sender.Name)
    End Sub
    Private Sub pic0310_Click(sender As Object, e As EventArgs) Handles pic0310.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0311_Click(sender As Object, e As EventArgs) Handles pic0311.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0312_Click(sender As Object, e As EventArgs) Handles pic0312.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0313_Click(sender As Object, e As EventArgs) Handles pic0313.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0314_Click(sender As Object, e As EventArgs) Handles pic0314.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0315_Click(sender As Object, e As EventArgs) Handles pic0315.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0316_Click(sender As Object, e As EventArgs) Handles pic0316.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0317_Click(sender As Object, e As EventArgs) Handles pic0317.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0318_Click(sender As Object, e As EventArgs) Handles pic0318.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0319_Click(sender As Object, e As EventArgs) Handles pic0319.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0320_Click(sender As Object, e As EventArgs) Handles pic0320.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0321_Click(sender As Object, e As EventArgs) Handles pic0321.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0322_Click(sender As Object, e As EventArgs) Handles pic0322.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0323_Click(sender As Object, e As EventArgs) Handles pic0323.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0324_Click(sender As Object, e As EventArgs) Handles pic0324.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0325_Click(sender As Object, e As EventArgs) Handles pic0325.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0326_Click(sender As Object, e As EventArgs) Handles pic0326.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0327_Click(sender As Object, e As EventArgs) Handles pic0327.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0328_Click(sender As Object, e As EventArgs) Handles pic0328.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0329_Click(sender As Object, e As EventArgs) Handles pic0329.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0400_Click(sender As Object, e As EventArgs) Handles pic0400.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0401_Click(sender As Object, e As EventArgs) Handles pic0401.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0402_Click(sender As Object, e As EventArgs) Handles pic0402.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0403_Click(sender As Object, e As EventArgs) Handles pic0403.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0404_Click(sender As Object, e As EventArgs) Handles pic0404.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0405_Click(sender As Object, e As EventArgs) Handles pic0405.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0406_Click(sender As Object, e As EventArgs) Handles pic0406.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0407_Click(sender As Object, e As EventArgs) Handles pic0407.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0408_Click(sender As Object, e As EventArgs) Handles pic0408.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0409_Click(sender As Object, e As EventArgs) Handles pic0409.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0410_Click(sender As Object, e As EventArgs) Handles pic0410.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0411_Click(sender As Object, e As EventArgs) Handles pic0411.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0412_Click(sender As Object, e As EventArgs) Handles pic0412.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0413_Click(sender As Object, e As EventArgs) Handles pic0413.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0414_Click(sender As Object, e As EventArgs) Handles pic0414.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0415_Click(sender As Object, e As EventArgs) Handles pic0415.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0416_Click(sender As Object, e As EventArgs) Handles pic0416.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0417_Click(sender As Object, e As EventArgs) Handles pic0417.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0418_Click(sender As Object, e As EventArgs) Handles pic0418.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0419_Click(sender As Object, e As EventArgs) Handles pic0419.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0420_Click(sender As Object, e As EventArgs) Handles pic0420.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0421_Click(sender As Object, e As EventArgs) Handles pic0421.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0422_Click(sender As Object, e As EventArgs) Handles pic0422.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0423_Click(sender As Object, e As EventArgs) Handles pic0423.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0424_Click(sender As Object, e As EventArgs) Handles pic0424.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0425_Click(sender As Object, e As EventArgs) Handles pic0425.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0426_Click(sender As Object, e As EventArgs) Handles pic0426.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0427_Click(sender As Object, e As EventArgs) Handles pic0427.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0428_Click(sender As Object, e As EventArgs) Handles pic0428.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic0429_Click(sender As Object, e As EventArgs) Handles pic0429.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic1118_Click(sender As Object, e As EventArgs) Handles pic1118.Click
        picClick(sender.Name)
    End Sub

    Private Sub pic1119_Click(sender As Object, e As EventArgs) Handles pic1119.Click
        picClick(sender.Name)
    End Sub


    Private Sub pic1120_Click(sender As Object, e As EventArgs) Handles pic1120.Click
        picClick(sender.Name)
    End Sub
    Private Sub pic1218_Click(sender As Object, e As EventArgs) Handles pic1218.Click
        picClick(sender.Name)
    End Sub

    Private Sub picIconRotate_Click(sender As Object, e As EventArgs)
        Active_Tool = "rotate"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconCross_Click(sender As Object, e As EventArgs) Handles picIconCross.Click
        Active_Tool = "cross"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconCurve_Click(sender As Object, e As EventArgs) Handles picIconCurve.Click
        Active_Tool = "curve"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconCurveParallel_Click(sender As Object, e As EventArgs) Handles picIconCurveParallel.Click
        Active_Tool = "curve_parallel"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub picIconEnd_Click(sender As Object, e As EventArgs) Handles picIconEnd.Click
        Active_Tool = "end"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub picIconLayout_Click(sender As Object, e As EventArgs) Handles picIconLayout.Click
        Active_Tool = "pfeil"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub picIconPermY_Click(sender As Object, e As EventArgs) Handles picIconPermY.Click
        Active_Tool = "custom_perm_y"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconPermLeft_Click(sender As Object, e As EventArgs) Handles picIconPermLeft.Click
        Active_Tool = "custom_perm_left"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconPermRight_Click(sender As Object, e As EventArgs) Handles picIconPermRight.Click
        Active_Tool = "custom_perm_right"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconRoute_Click(sender As Object, e As EventArgs) Handles picIconRoute.Click
        Active_Tool = "route"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub picIconSwitchLeft_Click(sender As Object, e As EventArgs) Handles picIconSwitchLeft.Click
        Active_Tool = "switch_left"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconSwitchRight_Click(sender As Object, e As EventArgs) Handles picIconSwitchRight.Click
        Active_Tool = "switch_right"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconThreeWay_Click(sender As Object, e As EventArgs) Handles picIconThreeWay.Click
        Active_Tool = "threeway"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconSwitchY_Click(sender As Object, e As EventArgs) Handles picIconSwitchY.Click
        Active_Tool = "switch_y"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIcons88_Click(sender As Object, e As EventArgs) Handles picIcons88.Click
        Active_Tool = "s88"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIcons88Curve_Click(sender As Object, e As EventArgs) Handles picIcons88Curve.Click
        Active_Tool = "s88_curve"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub picIconSignal_Click(sender As Object, e As EventArgs) Handles picIconSignal.Click
        Active_Tool = "signal"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub picIconText_Click(sender As Object, e As EventArgs) Handles picIconText.Click
        Active_Tool = "text"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconDecouple_Click(sender As Object, e As EventArgs) Handles picIconDecouple.Click
        Active_Tool = "decouple"
        Message("Active tool: " & Active_Tool)
    End Sub


    Private Sub pic0500_Click(sender As Object, e As EventArgs) Handles pic0500.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0501_Click(sender As Object, e As EventArgs) Handles pic0501.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0502_Click(sender As Object, e As EventArgs) Handles pic0502.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0503_Click(sender As Object, e As EventArgs) Handles pic0503.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0504_Click(sender As Object, e As EventArgs) Handles pic0504.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0505_Click(sender As Object, e As EventArgs) Handles pic0505.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0506_Click(sender As Object, e As EventArgs) Handles pic0506.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0507_Click(sender As Object, e As EventArgs) Handles pic0507.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0508_Click(sender As Object, e As EventArgs) Handles pic0508.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0509_Click(sender As Object, e As EventArgs) Handles pic0509.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0510_Click(sender As Object, e As EventArgs) Handles pic0510.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0511_Click(sender As Object, e As EventArgs) Handles pic0511.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0512_Click(sender As Object, e As EventArgs) Handles pic0512.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0513_Click(sender As Object, e As EventArgs) Handles pic0513.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0514_Click(sender As Object, e As EventArgs) Handles pic0514.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0515_Click(sender As Object, e As EventArgs) Handles pic0515.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0516_Click(sender As Object, e As EventArgs) Handles pic0516.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0517_Click(sender As Object, e As EventArgs) Handles pic0517.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0518_Click(sender As Object, e As EventArgs) Handles pic0518.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0519_Click(sender As Object, e As EventArgs) Handles pic0519.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0520_Click(sender As Object, e As EventArgs) Handles pic0520.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0521_Click(sender As Object, e As EventArgs) Handles pic0521.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0522_Click(sender As Object, e As EventArgs) Handles pic0522.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0523_Click(sender As Object, e As EventArgs) Handles pic0523.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0524_Click(sender As Object, e As EventArgs) Handles pic0524.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0525_Click(sender As Object, e As EventArgs) Handles pic0525.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0526_Click(sender As Object, e As EventArgs) Handles pic0526.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0527_Click(sender As Object, e As EventArgs) Handles pic0527.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0528_Click(sender As Object, e As EventArgs) Handles pic0528.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0529_Click(sender As Object, e As EventArgs) Handles pic0529.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0600_Click(sender As Object, e As EventArgs) Handles pic0600.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0601_Click(sender As Object, e As EventArgs) Handles pic0601.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0602_Click(sender As Object, e As EventArgs) Handles pic0602.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0603_Click(sender As Object, e As EventArgs) Handles pic0603.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0604_Click(sender As Object, e As EventArgs) Handles pic0604.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0605_Click(sender As Object, e As EventArgs) Handles pic0605.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0606_Click(sender As Object, e As EventArgs) Handles pic0606.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0607_Click(sender As Object, e As EventArgs) Handles pic0607.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0608_Click(sender As Object, e As EventArgs) Handles pic0608.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0609_Click(sender As Object, e As EventArgs) Handles pic0609.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0610_Click(sender As Object, e As EventArgs) Handles pic0610.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0611_Click(sender As Object, e As EventArgs) Handles pic0611.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0612_Click(sender As Object, e As EventArgs) Handles pic0612.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0613_Click(sender As Object, e As EventArgs) Handles pic0613.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0614_Click(sender As Object, e As EventArgs) Handles pic0614.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0615_Click(sender As Object, e As EventArgs) Handles pic0615.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0616_Click(sender As Object, e As EventArgs) Handles pic0616.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0617_Click(sender As Object, e As EventArgs) Handles pic0617.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0618_Click(sender As Object, e As EventArgs) Handles pic0618.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0619_Click(sender As Object, e As EventArgs) Handles pic0619.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0620_Click(sender As Object, e As EventArgs) Handles pic0620.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0621_Click(sender As Object, e As EventArgs) Handles pic0621.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0622_Click(sender As Object, e As EventArgs) Handles pic0622.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0623_Click(sender As Object, e As EventArgs) Handles pic0623.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0624_Click(sender As Object, e As EventArgs) Handles pic0624.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0625_Click(sender As Object, e As EventArgs) Handles pic0625.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0626_Click(sender As Object, e As EventArgs) Handles pic0626.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0627_Click(sender As Object, e As EventArgs) Handles pic0627.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0628_Click(sender As Object, e As EventArgs) Handles pic0628.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0629_Click(sender As Object, e As EventArgs) Handles pic0629.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0700_Click(sender As Object, e As EventArgs) Handles pic0700.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0701_Click(sender As Object, e As EventArgs) Handles pic0701.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0702_Click(sender As Object, e As EventArgs) Handles pic0702.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0703_Click(sender As Object, e As EventArgs) Handles pic0703.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0704_Click(sender As Object, e As EventArgs) Handles pic0704.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0705_Click(sender As Object, e As EventArgs) Handles pic0705.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0706_Click(sender As Object, e As EventArgs) Handles pic0706.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0707_Click(sender As Object, e As EventArgs) Handles pic0707.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0708_Click(sender As Object, e As EventArgs) Handles pic0708.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0709_Click(sender As Object, e As EventArgs) Handles pic0709.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0710_Click(sender As Object, e As EventArgs) Handles pic0710.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0711_Click(sender As Object, e As EventArgs) Handles pic0711.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0712_Click(sender As Object, e As EventArgs) Handles pic0712.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0713_Click(sender As Object, e As EventArgs) Handles pic0713.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0714_Click(sender As Object, e As EventArgs) Handles pic0714.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0715_Click(sender As Object, e As EventArgs) Handles pic0715.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0716_Click(sender As Object, e As EventArgs) Handles pic0716.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0717_Click(sender As Object, e As EventArgs) Handles pic0717.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0718_Click(sender As Object, e As EventArgs) Handles pic0718.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0719_Click(sender As Object, e As EventArgs) Handles pic0719.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0720_Click(sender As Object, e As EventArgs) Handles pic0720.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0721_Click(sender As Object, e As EventArgs) Handles pic0721.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0722_Click(sender As Object, e As EventArgs) Handles pic0722.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0723_Click(sender As Object, e As EventArgs) Handles pic0723.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0724_Click(sender As Object, e As EventArgs) Handles pic0724.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0725_Click(sender As Object, e As EventArgs) Handles pic0725.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0726_Click(sender As Object, e As EventArgs) Handles pic0726.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0727_Click(sender As Object, e As EventArgs) Handles pic0727.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0728_Click(sender As Object, e As EventArgs) Handles pic0728.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0729_Click(sender As Object, e As EventArgs) Handles pic0729.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0800_Click(sender As Object, e As EventArgs) Handles pic0800.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0801_Click(sender As Object, e As EventArgs) Handles pic0801.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0802_Click(sender As Object, e As EventArgs) Handles pic0802.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0803_Click(sender As Object, e As EventArgs) Handles pic0803.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0804_Click(sender As Object, e As EventArgs) Handles pic0804.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0805_Click(sender As Object, e As EventArgs) Handles pic0805.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0806_Click(sender As Object, e As EventArgs) Handles pic0806.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0807_Click(sender As Object, e As EventArgs) Handles pic0807.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0808_Click(sender As Object, e As EventArgs) Handles pic0808.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0809_Click(sender As Object, e As EventArgs) Handles pic0809.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0810_Click(sender As Object, e As EventArgs) Handles pic0810.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0811_Click(sender As Object, e As EventArgs) Handles pic0811.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0812_Click(sender As Object, e As EventArgs) Handles pic0812.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0813_Click(sender As Object, e As EventArgs) Handles pic0813.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0814_Click(sender As Object, e As EventArgs) Handles pic0814.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0815_Click(sender As Object, e As EventArgs) Handles pic0815.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0816_Click(sender As Object, e As EventArgs) Handles pic0816.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0817_Click(sender As Object, e As EventArgs) Handles pic0817.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0818_Click(sender As Object, e As EventArgs) Handles pic0818.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0819_Click(sender As Object, e As EventArgs) Handles pic0819.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0820_Click(sender As Object, e As EventArgs) Handles pic0820.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0821_Click(sender As Object, e As EventArgs) Handles pic0821.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0822_Click(sender As Object, e As EventArgs) Handles pic0822.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0823_Click(sender As Object, e As EventArgs) Handles pic0823.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0824_Click(sender As Object, e As EventArgs) Handles pic0824.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0825_Click(sender As Object, e As EventArgs) Handles pic0825.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0826_Click(sender As Object, e As EventArgs) Handles pic0826.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0827_Click(sender As Object, e As EventArgs) Handles pic0827.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0828_Click(sender As Object, e As EventArgs) Handles pic0828.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0829_Click(sender As Object, e As EventArgs) Handles pic0829.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0900_Click(sender As Object, e As EventArgs) Handles pic0900.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0901_Click(sender As Object, e As EventArgs) Handles pic0901.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0902_Click(sender As Object, e As EventArgs) Handles pic0902.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0903_Click(sender As Object, e As EventArgs) Handles pic0903.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0904_Click(sender As Object, e As EventArgs) Handles pic0904.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0905_Click(sender As Object, e As EventArgs) Handles pic0905.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0906_Click(sender As Object, e As EventArgs) Handles pic0906.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0907_Click(sender As Object, e As EventArgs) Handles pic0907.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0908_Click(sender As Object, e As EventArgs) Handles pic0908.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0909_Click(sender As Object, e As EventArgs) Handles pic0909.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0910_Click(sender As Object, e As EventArgs) Handles pic0910.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0911_Click(sender As Object, e As EventArgs) Handles pic0911.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0912_Click(sender As Object, e As EventArgs) Handles pic0912.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0913_Click(sender As Object, e As EventArgs) Handles pic0913.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0914_Click(sender As Object, e As EventArgs) Handles pic0914.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0915_Click(sender As Object, e As EventArgs) Handles pic0915.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0916_Click(sender As Object, e As EventArgs) Handles pic0916.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0917_Click(sender As Object, e As EventArgs) Handles pic0917.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0918_Click(sender As Object, e As EventArgs) Handles pic0918.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0919_Click(sender As Object, e As EventArgs) Handles pic0919.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0920_Click(sender As Object, e As EventArgs) Handles pic0920.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0921_Click(sender As Object, e As EventArgs) Handles pic0921.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0922_Click(sender As Object, e As EventArgs) Handles pic0922.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0923_Click(sender As Object, e As EventArgs) Handles pic0923.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0924_Click(sender As Object, e As EventArgs) Handles pic0924.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0925_Click(sender As Object, e As EventArgs) Handles pic0925.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0926_Click(sender As Object, e As EventArgs) Handles pic0926.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0927_Click(sender As Object, e As EventArgs) Handles pic0927.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0928_Click(sender As Object, e As EventArgs) Handles pic0928.Click
        picClick(sender.name)
    End Sub

    Private Sub pic0929_Click(sender As Object, e As EventArgs) Handles pic0929.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1000_Click(sender As Object, e As EventArgs) Handles pic1000.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1001_Click(sender As Object, e As EventArgs) Handles pic1001.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1002_Click(sender As Object, e As EventArgs) Handles pic1002.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1003_Click(sender As Object, e As EventArgs) Handles pic1003.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1004_Click(sender As Object, e As EventArgs) Handles pic1004.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1005_Click(sender As Object, e As EventArgs) Handles pic1005.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1006_Click(sender As Object, e As EventArgs) Handles pic1006.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1007_Click(sender As Object, e As EventArgs) Handles pic1007.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1008_Click(sender As Object, e As EventArgs) Handles pic1008.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1009_Click(sender As Object, e As EventArgs) Handles pic1009.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1010_Click(sender As Object, e As EventArgs) Handles pic1010.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1011_Click(sender As Object, e As EventArgs) Handles pic1011.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1012_Click(sender As Object, e As EventArgs) Handles pic1012.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1013_Click(sender As Object, e As EventArgs) Handles pic1013.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1014_Click(sender As Object, e As EventArgs) Handles pic1014.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1015_Click(sender As Object, e As EventArgs) Handles pic1015.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1016_Click(sender As Object, e As EventArgs) Handles pic1016.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1017_Click(sender As Object, e As EventArgs) Handles pic1017.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1018_Click(sender As Object, e As EventArgs) Handles pic1018.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1019_Click(sender As Object, e As EventArgs) Handles pic1019.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1020_Click(sender As Object, e As EventArgs) Handles pic1020.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1021_Click(sender As Object, e As EventArgs) Handles pic1021.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1022_Click(sender As Object, e As EventArgs) Handles pic1022.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1023_Click(sender As Object, e As EventArgs) Handles pic1023.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1024_Click(sender As Object, e As EventArgs) Handles pic1024.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1025_Click(sender As Object, e As EventArgs) Handles pic1025.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1026_Click(sender As Object, e As EventArgs) Handles pic1026.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1027_Click(sender As Object, e As EventArgs) Handles pic1027.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1028_Click(sender As Object, e As EventArgs) Handles pic1028.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1029_Click(sender As Object, e As EventArgs) Handles pic1029.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1100_Click(sender As Object, e As EventArgs) Handles pic1100.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1101_Click(sender As Object, e As EventArgs) Handles pic1101.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1102_Click(sender As Object, e As EventArgs) Handles pic1102.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1103_Click(sender As Object, e As EventArgs) Handles pic1103.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1104_Click(sender As Object, e As EventArgs) Handles pic1104.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1105_Click(sender As Object, e As EventArgs) Handles pic1105.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1106_Click(sender As Object, e As EventArgs) Handles pic1106.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1107_Click(sender As Object, e As EventArgs) Handles pic1107.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1108_Click(sender As Object, e As EventArgs) Handles pic1108.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1109_Click(sender As Object, e As EventArgs) Handles pic1109.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1110_Click(sender As Object, e As EventArgs) Handles pic1110.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1111_Click(sender As Object, e As EventArgs) Handles pic1111.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1112_Click(sender As Object, e As EventArgs) Handles pic1112.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1113_Click(sender As Object, e As EventArgs) Handles pic1113.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1114_Click(sender As Object, e As EventArgs) Handles pic1114.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1115_Click(sender As Object, e As EventArgs) Handles pic1115.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1116_Click(sender As Object, e As EventArgs) Handles pic1116.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1117_Click(sender As Object, e As EventArgs) Handles pic1117.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1121_Click(sender As Object, e As EventArgs) Handles pic1121.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1122_Click(sender As Object, e As EventArgs) Handles pic1122.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1123_Click(sender As Object, e As EventArgs) Handles pic1123.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1124_Click(sender As Object, e As EventArgs) Handles pic1124.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1125_Click(sender As Object, e As EventArgs) Handles pic1125.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1126_Click(sender As Object, e As EventArgs) Handles pic1126.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1127_Click(sender As Object, e As EventArgs) Handles pic1127.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1128_Click(sender As Object, e As EventArgs) Handles pic1128.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1129_Click(sender As Object, e As EventArgs) Handles pic1129.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1200_Click(sender As Object, e As EventArgs) Handles pic1200.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1201_Click(sender As Object, e As EventArgs) Handles pic1201.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1202_Click(sender As Object, e As EventArgs) Handles pic1202.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1203_Click(sender As Object, e As EventArgs) Handles pic1203.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1204_Click(sender As Object, e As EventArgs) Handles pic1204.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1205_Click(sender As Object, e As EventArgs) Handles pic1205.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1206_Click(sender As Object, e As EventArgs) Handles pic1206.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1207_Click(sender As Object, e As EventArgs) Handles pic1207.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1208_Click(sender As Object, e As EventArgs) Handles pic1208.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1209_Click(sender As Object, e As EventArgs) Handles pic1209.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1210_Click(sender As Object, e As EventArgs) Handles pic1210.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1211_Click(sender As Object, e As EventArgs) Handles pic1211.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1212_Click(sender As Object, e As EventArgs) Handles pic1212.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1213_Click(sender As Object, e As EventArgs) Handles pic1213.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1214_Click(sender As Object, e As EventArgs) Handles pic1214.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1215_Click(sender As Object, e As EventArgs) Handles pic1215.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1216_Click(sender As Object, e As EventArgs) Handles pic1216.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1217_Click(sender As Object, e As EventArgs) Handles pic1217.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1219_Click(sender As Object, e As EventArgs) Handles pic1219.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1220_Click(sender As Object, e As EventArgs) Handles pic1220.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1221_Click(sender As Object, e As EventArgs) Handles pic1221.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1222_Click(sender As Object, e As EventArgs) Handles pic1222.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1223_Click(sender As Object, e As EventArgs) Handles pic1223.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1224_Click(sender As Object, e As EventArgs) Handles pic1224.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1225_Click(sender As Object, e As EventArgs) Handles pic1225.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1226_Click(sender As Object, e As EventArgs) Handles pic1226.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1227_Click(sender As Object, e As EventArgs) Handles pic1227.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1228_Click(sender As Object, e As EventArgs) Handles pic1228.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1229_Click(sender As Object, e As EventArgs) Handles pic1229.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1300_Click(sender As Object, e As EventArgs) Handles pic1300.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1301_Click(sender As Object, e As EventArgs) Handles pic1301.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1302_Click(sender As Object, e As EventArgs) Handles pic1302.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1303_Click(sender As Object, e As EventArgs) Handles pic1303.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1304_Click(sender As Object, e As EventArgs) Handles pic1304.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1305_Click(sender As Object, e As EventArgs) Handles pic1305.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1306_Click(sender As Object, e As EventArgs) Handles pic1306.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1307_Click(sender As Object, e As EventArgs) Handles pic1307.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1308_Click(sender As Object, e As EventArgs) Handles pic1308.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1309_Click(sender As Object, e As EventArgs) Handles pic1309.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1310_Click(sender As Object, e As EventArgs) Handles pic1310.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1311_Click(sender As Object, e As EventArgs) Handles pic1311.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1312_Click(sender As Object, e As EventArgs) Handles pic1312.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1313_Click(sender As Object, e As EventArgs) Handles pic1313.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1314_Click(sender As Object, e As EventArgs) Handles pic1314.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1315_Click(sender As Object, e As EventArgs) Handles pic1315.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1316_Click(sender As Object, e As EventArgs) Handles pic1316.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1317_Click(sender As Object, e As EventArgs) Handles pic1317.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1318_Click(sender As Object, e As EventArgs) Handles pic1318.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1319_Click(sender As Object, e As EventArgs) Handles pic1319.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1320_Click(sender As Object, e As EventArgs) Handles pic1320.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1321_Click(sender As Object, e As EventArgs) Handles pic1321.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1322_Click(sender As Object, e As EventArgs) Handles pic1322.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1323_Click(sender As Object, e As EventArgs) Handles pic1323.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1324_Click(sender As Object, e As EventArgs) Handles pic1324.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1325_Click(sender As Object, e As EventArgs) Handles pic1325.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1326_Click(sender As Object, e As EventArgs) Handles pic1326.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1327_Click(sender As Object, e As EventArgs) Handles pic1327.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1328_Click(sender As Object, e As EventArgs) Handles pic1328.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1329_Click(sender As Object, e As EventArgs) Handles pic1329.Click
        picClick(sender.name)
    End Sub
    Private Sub pic1400_Click(sender As Object, e As EventArgs) Handles pic1400.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1401_Click(sender As Object, e As EventArgs) Handles pic1401.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1402_Click(sender As Object, e As EventArgs) Handles pic1402.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1403_Click(sender As Object, e As EventArgs) Handles pic1403.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1404_Click(sender As Object, e As EventArgs) Handles pic1404.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1405_Click(sender As Object, e As EventArgs) Handles pic1405.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1406_Click(sender As Object, e As EventArgs) Handles pic1406.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1407_Click(sender As Object, e As EventArgs) Handles pic1407.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1408_Click(sender As Object, e As EventArgs) Handles pic1408.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1409_Click(sender As Object, e As EventArgs) Handles pic1409.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1410_Click(sender As Object, e As EventArgs) Handles pic1410.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1411_Click(sender As Object, e As EventArgs) Handles pic1411.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1412_Click(sender As Object, e As EventArgs) Handles pic1412.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1413_Click(sender As Object, e As EventArgs) Handles pic1413.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1414_Click(sender As Object, e As EventArgs) Handles pic1414.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1415_Click(sender As Object, e As EventArgs) Handles pic1415.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1416_Click(sender As Object, e As EventArgs) Handles pic1416.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1417_Click(sender As Object, e As EventArgs) Handles pic1417.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1418_Click(sender As Object, e As EventArgs) Handles pic1418.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1419_Click(sender As Object, e As EventArgs) Handles pic1419.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1420_Click(sender As Object, e As EventArgs) Handles pic1420.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1421_Click(sender As Object, e As EventArgs) Handles pic1421.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1422_Click(sender As Object, e As EventArgs) Handles pic1422.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1423_Click(sender As Object, e As EventArgs) Handles pic1423.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1424_Click(sender As Object, e As EventArgs) Handles pic1424.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1425_Click(sender As Object, e As EventArgs) Handles pic1425.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1426_Click(sender As Object, e As EventArgs) Handles pic1426.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1427_Click(sender As Object, e As EventArgs) Handles pic1427.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1428_Click(sender As Object, e As EventArgs) Handles pic1428.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1429_Click(sender As Object, e As EventArgs) Handles pic1429.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1500_Click(sender As Object, e As EventArgs) Handles pic1500.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1501_Click(sender As Object, e As EventArgs) Handles pic1501.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1502_Click(sender As Object, e As EventArgs) Handles pic1502.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1503_Click(sender As Object, e As EventArgs) Handles pic1503.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1504_Click(sender As Object, e As EventArgs) Handles pic1504.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1505_Click(sender As Object, e As EventArgs) Handles pic1505.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1506_Click(sender As Object, e As EventArgs) Handles pic1506.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1507_Click(sender As Object, e As EventArgs) Handles pic1507.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1508_Click(sender As Object, e As EventArgs) Handles pic1508.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1509_Click(sender As Object, e As EventArgs) Handles pic1509.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1510_Click(sender As Object, e As EventArgs) Handles pic1510.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1511_Click(sender As Object, e As EventArgs) Handles pic1511.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1512_Click(sender As Object, e As EventArgs) Handles pic1512.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1513_Click(sender As Object, e As EventArgs) Handles pic1513.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1514_Click(sender As Object, e As EventArgs) Handles pic1514.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1515_Click(sender As Object, e As EventArgs) Handles pic1515.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1516_Click(sender As Object, e As EventArgs) Handles pic1516.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1517_Click(sender As Object, e As EventArgs) Handles pic1517.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1518_Click(sender As Object, e As EventArgs) Handles pic1518.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1519_Click(sender As Object, e As EventArgs) Handles pic1519.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1520_Click(sender As Object, e As EventArgs) Handles pic1520.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1521_Click(sender As Object, e As EventArgs) Handles pic1521.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1522_Click(sender As Object, e As EventArgs) Handles pic1522.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1523_Click(sender As Object, e As EventArgs) Handles pic1523.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1524_Click(sender As Object, e As EventArgs) Handles pic1524.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1525_Click(sender As Object, e As EventArgs) Handles pic1525.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1526_Click(sender As Object, e As EventArgs) Handles pic1526.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1527_Click(sender As Object, e As EventArgs) Handles pic1527.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1528_Click(sender As Object, e As EventArgs) Handles pic1528.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1529_Click(sender As Object, e As EventArgs) Handles pic1529.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1600_Click(sender As Object, e As EventArgs) Handles pic1600.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1601_Click(sender As Object, e As EventArgs) Handles pic1601.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1602_Click(sender As Object, e As EventArgs) Handles pic1602.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1603_Click(sender As Object, e As EventArgs) Handles pic1603.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1604_Click(sender As Object, e As EventArgs) Handles pic1604.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1605_Click(sender As Object, e As EventArgs) Handles pic1605.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1606_Click(sender As Object, e As EventArgs) Handles pic1606.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1607_Click(sender As Object, e As EventArgs) Handles pic1607.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1608_Click(sender As Object, e As EventArgs) Handles pic1608.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1609_Click(sender As Object, e As EventArgs) Handles pic1609.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1610_Click(sender As Object, e As EventArgs) Handles pic1610.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1611_Click(sender As Object, e As EventArgs) Handles pic1611.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1612_Click(sender As Object, e As EventArgs) Handles pic1612.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1613_Click(sender As Object, e As EventArgs) Handles pic1613.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1614_Click(sender As Object, e As EventArgs) Handles pic1614.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1615_Click(sender As Object, e As EventArgs) Handles pic1615.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1616_Click(sender As Object, e As EventArgs) Handles pic1616.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1617_Click(sender As Object, e As EventArgs) Handles pic1617.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1618_Click(sender As Object, e As EventArgs) Handles pic1618.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1619_Click(sender As Object, e As EventArgs) Handles pic1619.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1620_Click(sender As Object, e As EventArgs) Handles pic1620.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1621_Click(sender As Object, e As EventArgs) Handles pic1621.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1622_Click(sender As Object, e As EventArgs) Handles pic1622.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1623_Click(sender As Object, e As EventArgs) Handles pic1623.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1624_Click(sender As Object, e As EventArgs) Handles pic1624.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1625_Click(sender As Object, e As EventArgs) Handles pic1625.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1626_Click(sender As Object, e As EventArgs) Handles pic1626.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1627_Click(sender As Object, e As EventArgs) Handles pic1627.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1628_Click(sender As Object, e As EventArgs) Handles pic1628.Click
        picClick(sender.name)
    End Sub

    Private Sub pic1629_Click(sender As Object, e As EventArgs) Handles pic1629.Click
        picClick(sender.name)
    End Sub

    Private Sub picIconGarbageBin_Click_1(sender As Object, e As EventArgs) Handles picIconGarbageBin.Click
        Active_Tool = "empty"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconRotate_Click_1(sender As Object, e As EventArgs) Handles picIconRotate.Click
        Active_Tool = "rotate"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub cmdEleementMove_Click(sender As Object, e As EventArgs) Handles cmdEleementMove.Click
        Active_Tool = "move_element"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub cmdAddressEdit_Click(sender As Object, e As EventArgs) Handles cmdAddressEdit.Click
        Active_Tool = "edit_address"
        Message("Active tool: " & Active_Tool)
        Reveal_Addr = True
        Reveal_Text = False
        GleisbildDisplay()
    End Sub

    Private Sub cmdTextEdit_Click(sender As Object, e As EventArgs) Handles cmdTextEdit.Click
        Active_Tool = "edit_text"
        Message("Active tool: " & Active_Tool)
        Reveal_Addr = False
        Reveal_Text = True
        GleisbildDisplay()
    End Sub

    Private Sub cmdSplitHorizontally_Click_1(sender As Object, e As EventArgs) Handles cmdSplitHorizontally.Click
        Active_Tool = "split_horizontally"
        Message("Active tool: " & Active_Tool)
        Message("Click a non-empty cell just below of where to split")
    End Sub

    Private Sub cmdSplitVertically_Click_1(sender As Object, e As EventArgs) Handles cmdSplitVertically.Click
        Active_Tool = "split_vertically"
        Message("Active tool: " & Active_Tool)
        Message("Click a non-empty cell just to the right of where to split")
    End Sub

    Private Sub picIconLamp_Click(sender As Object, e As EventArgs) Handles picIconLamp.Click
        Active_Tool = "lamp"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub ShowOutputFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowOutputFileToolStripMenuItem.Click
        ShowOutputFileToolStripMenuItem.Checked = Not ShowOutputFileToolStripMenuItem.Checked
        lstOutput.Visible = ShowOutputFileToolStripMenuItem.Checked
    End Sub


    Private Sub picIconGate_Click(sender As Object, e As EventArgs) Handles picIconGate.Click
        Active_Tool = "gate"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconViaduct_Click(sender As Object, e As EventArgs) Handles picIconViaduct.Click
        Active_Tool = "viaduct"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIcons88CurveParallel_Click(sender As Object, e As EventArgs) Handles picIcons88CurveParallel.Click
        Active_Tool = "s88_curve_parallel"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconSignalFHP01_Click(sender As Object, e As EventArgs) Handles picIconSignalFHP01.Click
        Active_Tool = "signal_f_hp01"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconSignalFHP012_Click(sender As Object, e As EventArgs) Handles picIconSignalFHP012.Click
        Active_Tool = "signal_f_hp012"
        Message("Active tool: " & Active_Tool)
    End Sub
    Private Sub picIconSignalSH01_Click(sender As Object, e As EventArgs) Handles picIconSignalSH01.Click
        Active_Tool = "signal_sh01"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconTurntable_Click(sender As Object, e As EventArgs) Handles picIconTurntable.Click
        Active_Tool = "turntable"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconAndreaskreuz_Click(sender As Object, e As EventArgs) Handles picIconAndreaskreuz.Click
        Active_Tool = "andreaskreuz"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconSignalPHP012_Click(sender As Object, e As EventArgs) Handles picIconSignalPHP012.Click
        Active_Tool = "signal_p_hp012"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconStdRed_Click(sender As Object, e As EventArgs) Handles picIconStdRed.Click
        Active_Tool = "std_red_green_0"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub picIconStdGreen_Click(sender As Object, e As EventArgs) Handles picIconStdGreen.Click
        Active_Tool = "std_red_green_1"
        Message("Active tool: " & Active_Tool)
    End Sub

    Private Sub TurnOffDeveloperModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TurnOffDeveloperModeToolStripMenuItem.Click
        Developer_Environment = False
        AdjustPage()
    End Sub

    Private Sub cboRouteNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboRouteNames.SelectedIndexChanged
        RH.RouteShow(cboRouteNames.SelectedItem)
    End Sub

    Private Sub UseEnglishTermsForElementsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UseEnglishTermsForElementsToolStripMenuItem.Click
        Use_English_Terms = UseEnglishTermsForElementsToolStripMenuItem.Checked
    End Sub

    Private Sub ConsistencyCheckToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsistencyCheckToolStripMenuItem.Click
        If GleisbildConsistencyCheck() Then Message("No double occupancy found.")
    End Sub


    Private Sub DumpElementsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DumpElementsToolStripMenuItem.Click
        Dim el As ELEMENT
        lstOutput.Visible = True
        lstOutput.Items.Clear()
        For i = 1 To E_TOP
            el = Elements(i)
            lstOutput.Items.Add("element")
            lstOutput.Items.Add(" .id=" & el.id & " (" & Format(el.row_no) & "," & Format(el.col_no) & ")")
            lstOutput.Items.Add(" .id_norm=" & el.id_normalized)
            lstOutput.Items.Add(" .typ=" & el.typ_s)
            lstOutput.Items.Add(" .drehung=" & Format(el.drehung_i))
            lstOutput.Items.Add(" .artikel=" & Format(el.artikel_i))
            lstOutput.Items.Add(" .deviceid=" & Format(el.deviceId_s))
            lstOutput.Items.Add(" .zustand=" & Format(el.zustand_s))
            lstOutput.Items.Add(" .text=" & Format(el.text))
        Next
    End Sub
    Private Sub HowToUseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HowToUseToolStripMenuItem.Click
        ' To simplify installation the help file has been built into the source text
        Const s_top = 64
        Static been_here As Boolean
        Static s(s_top) As String
        frmHelp.HelpWindowSetPos(Me.Left, Me.Top + 30, grpLayout.Width, grpLayout.Height)
        '  frmHelp.HelpWindowSetSize(my_width As Long, my_height As Long)
        'frmHelp.HelpDisplay(Start_Directory & "\TrackDiagramEditor_hlp.txt")
        If Not been_here Then
            s(1) = "Track Diagram Editor " & Version
            s(2) = ""
            s(3) = "A track layout spans one or more pages with track diagrams."
            s(4) = "Each page (tab in CS2 or TC) corresponds to a unique track diagram file."
            s(5) = ""
            s(6) = "Track diagram files can be created by a CS2 or by this editor."
            s(7) = ""
            s(8) = "The mapping of track diagram files to pages of the track layout is created by the CS2 or by this "
            s(9) = "track diagram editor. Said mapping (""layout map"") Is stored in a file called gleisbild.cs2."
            s(10) = ""
            s(11) = "Märklin CS2's convention regarding folder structure (this structure is also required by the Track Control program ""TC""):"
            s(12) = "  <my_path>\config\gleisbilder\<track diagram name>.cs2   -- One or more track diagram files, each assigned to a tab"
            s(13) = "                                                             in CS2's layout display"
            s(14) = "  <my_path>\config\gleisbild.cs2                          -- This file constitues an index to the pages (CS2 tabs)"
            s(15) = "                                                             with track diagrams"
            s(16) = "  <my_path>\config\fahrstrassen.cs2                       -- this file holds all routes defined on the CS2 Memory "
            s(17) = "                                                             page tabs."
            s(18) = "                                                             This file Is not used by the Train Control program, but the"
            s(19) = "                                                             CS2 created routes can be analyzed by the Track Diagram Editor."
            s(20) = ""
            s(21) = "The Layout map is managed behind the scenes and user interaction with the lauout map Is not required."
            s(22) = "Nevertheless, we have included a few explicit operations on the map, refer 'Page Management Actions' below."
            s(23) = ""
            s(24) = "The Layout editor handles the above Structure. However, should the track diagrams not reside in a folder named"
            s(25) = """gleisbilder"" this can be handled to a limited extent (""flat Structure""): In this case fahrstrassen.cs2 should reside"
            s(26) = "in the same folder As the track diagrams and pages cannot be managed."
            s(27) = ""
            s(28) = "Ttack diagram files must have the file type .cs2."
            s(29) = ""
            s(30) = "Track diagram creation and editing"
            s(31) = "----------------------------------"
            s(32) = " - Menu ""File-New Layout..."" lets you create a new layout consisting of one or more pages of track diagram files."
            s(33) = ""
            s(34) = " - Menu ""File-Open..."" lets you open And edit a previously created track diagram or a track diagram copied over from a"
            s(35) = "   CS2. This option can handle track diagrams which are not assigned to a page of a layout (""flat structure"")"
            s(36) = ""
            s(37) = " - Menu ""File-Save"" saves your changes to the track diagram."
            s(38) = ""
            s(39) = " - Menu ""File-Save a Copy as..."" lets you save your track diagram under a new name. Note that this new track diagram will not"
            s(40) = "   be assigned a page and it will not automatically be included in the layout map. You can manually assign it "
            s(41) = "   to a page later, cf. ""Add Page"" below."
            s(42) = ""
            s(43) = ""
            s(44) = "Page Management Actions (refer the buttons in the right hand pane)"
            s(45) = "------------------------------------------------------------------"
            s(46) = " - Add Page"
            s(47) = "   This action has two options:"
            s(48) = "   1 Adds a new page so that you can create a track diagram from scratch."
            s(49) = "   2 Adds a new page and allows you to import an existing track diagram file into the new page."
            s(50) = ""
            s(51) = " - Duplicate Page"
            s(52) = "   Creates a new page with a copy of the current track diagram"
            s(53) = ""
            s(54) = " - Rename Page"
            s(55) = "   - Renames the page and the corresponding track diagram file"
            s(56) = ""
            s(57) = " - Remove Page"
            s(58) = "   - Removes the selected page from the layout. The corresponding track diagram file is not deleted from disk, but"
            s(59) = "     moved to the backup folder."
            s(60) = ""
            s(61) = "Installation"
            s(62) = "------------"
            s(63) = "The installation makes no changes to your Windows system or registry."
            s(64) = "No installation is necessary when the editor is activated from the train contral program ""TC""."
            been_here = True
        End If
        frmHelp.HelpDisplay(s, s_top)
    End Sub
    Private Sub LimitationsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimitationsToolStripMenuItem.Click
        ' To simplify installation the help file has been built into the source text
        Const s_top = 14
        Static been_here As Boolean
        Static s(s_top) As String
        frmHelp.HelpWindowSetPos(Me.Left, Me.Top + 30, grpLayout.Width, grpLayout.Height)
        If Not been_here Then
            s(1) = "Limitations in " & Version
            s(2) = ""
            s(3) = "All lamps are displayed in yellow."
            s(4) = ""
            s(5) = """entkuppler_1"" Is handled Like ""entkuppler""."
            s(6) = ""
            s(7) = "ELEMENT Type ""sonstige_gbs"" is displayed as a question mark and"
            s(8) = "new elements cannot be inserted. The address of existing elements can be changed,"
            s(9) = "however."
            s(10) = ""
            s(11) = "New instances of ""std_rot"" cannot be created. Use ""std_rot_gruen_0"" instead."
            s(12) = ""
            s(13) = "It is not understood how to derive the s88 deviceid from the artikel number."
            s(14) = "For now the user will be prompted for the appropriate hex string."
            been_here = True
        End If
        frmHelp.HelpDisplay(s, s_top)
    End Sub


    Private Sub VerboseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VerboseToolStripMenuItem.Click
        VerboseToolStripMenuItem.Checked = Not VerboseToolStripMenuItem.Checked
        Verbose = VerboseToolStripMenuItem.Checked
        ShowInputFileToolStripMenuItem.Visible = Developer_Environment Or Verbose
    End Sub

    Private Sub pic0000_MouseHover(sender As Object, e As EventArgs) Handles pic0000.MouseHover, pic0001.MouseHover, pic0002.MouseHover, pic0003.MouseHover, pic0004.MouseHover, pic0005.MouseHover, pic0006.MouseHover, pic0007.MouseHover, pic0008.MouseHover, pic0009.MouseHover,
            pic0010.MouseHover, pic0011.MouseHover, pic0012.MouseHover, pic0013.MouseHover, pic0014.MouseHover, pic0015.MouseHover, pic0016.MouseHover, pic0017.MouseHover, pic0018.MouseHover, pic0019.MouseHover,
            pic0020.MouseHover, pic0021.MouseHover, pic0022.MouseHover, pic0023.MouseHover, pic0024.MouseHover, pic0025.MouseHover, pic0026.MouseHover, pic0027.MouseHover, pic0028.MouseHover, pic0029.MouseHover
        picHover(sender)
    End Sub

    Private Sub pic0100_MouseHover(sender As Object, e As EventArgs) Handles pic0100.MouseHover, pic0101.MouseHover, pic0102.MouseHover, pic0103.MouseHover, pic0104.MouseHover, pic0105.MouseHover, pic0106.MouseHover, pic0107.MouseHover, pic0108.MouseHover, pic0109.MouseHover,
            pic0110.MouseHover, pic0111.MouseHover, pic0112.MouseHover, pic0113.MouseHover, pic0114.MouseHover, pic0115.MouseHover, pic0116.MouseHover, pic0117.MouseHover, pic0118.MouseHover, pic0119.MouseHover,
            pic0120.MouseHover, pic0121.MouseHover, pic0122.MouseHover, pic0123.MouseHover, pic0124.MouseHover, pic0125.MouseHover, pic0126.MouseHover, pic0127.MouseHover, pic0128.MouseHover, pic0129.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic0200_MouseHover(sender As Object, e As EventArgs) Handles pic0200.MouseHover, pic0201.MouseHover, pic0202.MouseHover, pic0203.MouseHover, pic0204.MouseHover, pic0205.MouseHover, pic0206.MouseHover, pic0207.MouseHover, pic0208.MouseHover, pic0209.MouseHover,
            pic0210.MouseHover, pic0211.MouseHover, pic0212.MouseHover, pic0213.MouseHover, pic0214.MouseHover, pic0215.MouseHover, pic0216.MouseHover, pic0217.MouseHover, pic0218.MouseHover, pic0219.MouseHover,
            pic0220.MouseHover, pic0221.MouseHover, pic0222.MouseHover, pic0223.MouseHover, pic0224.MouseHover, pic0225.MouseHover, pic0226.MouseHover, pic0227.MouseHover, pic0228.MouseHover, pic0229.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic0300_MouseHover(sender As Object, e As EventArgs) Handles pic0300.MouseHover, pic0301.MouseHover, pic0302.MouseHover, pic0303.MouseHover, pic0304.MouseHover, pic0305.MouseHover, pic0306.MouseHover, pic0307.MouseHover, pic0308.MouseHover, pic0309.MouseHover,
            pic0310.MouseHover, pic0311.MouseHover, pic0312.MouseHover, pic0313.MouseHover, pic0314.MouseHover, pic0315.MouseHover, pic0316.MouseHover, pic0317.MouseHover, pic0318.MouseHover, pic0319.MouseHover,
            pic0320.MouseHover, pic0321.MouseHover, pic0322.MouseHover, pic0323.MouseHover, pic0324.MouseHover, pic0325.MouseHover, pic0326.MouseHover, pic0327.MouseHover, pic0328.MouseHover, pic0329.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic0400_MouseHover(sender As Object, e As EventArgs) Handles pic0400.MouseHover, pic0401.MouseHover, pic0402.MouseHover, pic0403.MouseHover, pic0404.MouseHover, pic0405.MouseHover, pic0406.MouseHover, pic0407.MouseHover, pic0408.MouseHover, pic0409.MouseHover,
            pic0410.MouseHover, pic0411.MouseHover, pic0412.MouseHover, pic0413.MouseHover, pic0414.MouseHover, pic0415.MouseHover, pic0416.MouseHover, pic0417.MouseHover, pic0418.MouseHover, pic0419.MouseHover,
            pic0420.MouseHover, pic0421.MouseHover, pic0422.MouseHover, pic0423.MouseHover, pic0424.MouseHover, pic0425.MouseHover, pic0426.MouseHover, pic0427.MouseHover, pic0428.MouseHover, pic0429.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic0500_MouseHover(sender As Object, e As EventArgs) Handles pic0500.MouseHover, pic0501.MouseHover, pic0502.MouseHover, pic0503.MouseHover, pic0504.MouseHover, pic0505.MouseHover, pic0506.MouseHover, pic0507.MouseHover, pic0508.MouseHover, pic0509.MouseHover,
            pic0510.MouseHover, pic0511.MouseHover, pic0512.MouseHover, pic0513.MouseHover, pic0514.MouseHover, pic0515.MouseHover, pic0516.MouseHover, pic0517.MouseHover, pic0518.MouseHover, pic0519.MouseHover,
            pic0520.MouseHover, pic0521.MouseHover, pic0522.MouseHover, pic0523.MouseHover, pic0524.MouseHover, pic0525.MouseHover, pic0526.MouseHover, pic0527.MouseHover, pic0528.MouseHover, pic0529.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic0600_MouseHover(sender As Object, e As EventArgs) Handles pic0600.MouseHover, pic0601.MouseHover, pic0602.MouseHover, pic0603.MouseHover, pic0604.MouseHover, pic0605.MouseHover, pic0606.MouseHover, pic0607.MouseHover, pic0608.MouseHover, pic0609.MouseHover,
            pic0610.MouseHover, pic0611.MouseHover, pic0612.MouseHover, pic0613.MouseHover, pic0614.MouseHover, pic0615.MouseHover, pic0616.MouseHover, pic0617.MouseHover, pic0618.MouseHover, pic0619.MouseHover,
            pic0620.MouseHover, pic0621.MouseHover, pic0622.MouseHover, pic0623.MouseHover, pic0624.MouseHover, pic0625.MouseHover, pic0626.MouseHover, pic0627.MouseHover, pic0628.MouseHover, pic0629.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic0700_MouseHover(sender As Object, e As EventArgs) Handles pic0700.MouseHover, pic0701.MouseHover, pic0702.MouseHover, pic0703.MouseHover, pic0704.MouseHover, pic0705.MouseHover, pic0706.MouseHover, pic0707.MouseHover, pic0708.MouseHover, pic0709.MouseHover,
            pic0710.MouseHover, pic0711.MouseHover, pic0712.MouseHover, pic0713.MouseHover, pic0714.MouseHover, pic0715.MouseHover, pic0716.MouseHover, pic0717.MouseHover, pic0718.MouseHover, pic0719.MouseHover,
            pic0720.MouseHover, pic0721.MouseHover, pic0722.MouseHover, pic0723.MouseHover, pic0724.MouseHover, pic0725.MouseHover, pic0726.MouseHover, pic0727.MouseHover, pic0728.MouseHover, pic0729.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic0800_MouseHover(sender As Object, e As EventArgs) Handles pic0800.MouseHover, pic0801.MouseHover, pic0802.MouseHover, pic0803.MouseHover, pic0804.MouseHover, pic0805.MouseHover, pic0806.MouseHover, pic0807.MouseHover, pic0808.MouseHover, pic0809.MouseHover,
            pic0810.MouseHover, pic0811.MouseHover, pic0812.MouseHover, pic0813.MouseHover, pic0814.MouseHover, pic0815.MouseHover, pic0816.MouseHover, pic0817.MouseHover, pic0818.MouseHover, pic0819.MouseHover,
            pic0820.MouseHover, pic0821.MouseHover, pic0822.MouseHover, pic0823.MouseHover, pic0824.MouseHover, pic0825.MouseHover, pic0826.MouseHover, pic0827.MouseHover, pic0828.MouseHover, pic0829.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic0900_MouseHover(sender As Object, e As EventArgs) Handles pic0900.MouseHover, pic0901.MouseHover, pic0902.MouseHover, pic0903.MouseHover, pic0904.MouseHover, pic0905.MouseHover, pic0906.MouseHover, pic0907.MouseHover, pic0908.MouseHover, pic0909.MouseHover,
            pic0910.MouseHover, pic0911.MouseHover, pic0912.MouseHover, pic0913.MouseHover, pic0914.MouseHover, pic0915.MouseHover, pic0916.MouseHover, pic0917.MouseHover, pic0918.MouseHover, pic0919.MouseHover,
            pic0920.MouseHover, pic0921.MouseHover, pic0922.MouseHover, pic0923.MouseHover, pic0924.MouseHover, pic0925.MouseHover, pic0926.MouseHover, pic0927.MouseHover, pic0928.MouseHover, pic0929.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic1000_MouseHover(sender As Object, e As EventArgs) Handles pic1000.MouseHover, pic1001.MouseHover, pic1002.MouseHover, pic1003.MouseHover, pic1004.MouseHover, pic1005.MouseHover, pic1006.MouseHover, pic1007.MouseHover, pic1008.MouseHover, pic1009.MouseHover,
            pic1010.MouseHover, pic1011.MouseHover, pic1012.MouseHover, pic1013.MouseHover, pic1014.MouseHover, pic1015.MouseHover, pic1016.MouseHover, pic1017.MouseHover, pic1018.MouseHover, pic1019.MouseHover,
            pic1020.MouseHover, pic1021.MouseHover, pic1022.MouseHover, pic1023.MouseHover, pic1024.MouseHover, pic1025.MouseHover, pic1026.MouseHover, pic1027.MouseHover, pic1028.MouseHover, pic1029.MouseHover
        picHover(sender)
    End Sub

    Private Sub pic1100_MouseHover(sender As Object, e As EventArgs) Handles pic1100.MouseHover, pic1101.MouseHover, pic1102.MouseHover, pic1103.MouseHover, pic1104.MouseHover, pic1105.MouseHover, pic1106.MouseHover, pic1107.MouseHover, pic1108.MouseHover, pic1109.MouseHover,
            pic1110.MouseHover, pic1111.MouseHover, pic1112.MouseHover, pic1113.MouseHover, pic1114.MouseHover, pic1115.MouseHover, pic1116.MouseHover, pic1117.MouseHover, pic1118.MouseHover, pic1119.MouseHover,
            pic1120.MouseHover, pic1121.MouseHover, pic1122.MouseHover, pic1123.MouseHover, pic1124.MouseHover, pic1125.MouseHover, pic1126.MouseHover, pic1127.MouseHover, pic1128.MouseHover, pic1129.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic1200_MouseHover(sender As Object, e As EventArgs) Handles pic1200.MouseHover, pic1201.MouseHover, pic1202.MouseHover, pic1203.MouseHover, pic1204.MouseHover, pic1205.MouseHover, pic1206.MouseHover, pic1207.MouseHover, pic1208.MouseHover, pic1209.MouseHover,
            pic1210.MouseHover, pic1211.MouseHover, pic1212.MouseHover, pic1213.MouseHover, pic1214.MouseHover, pic1215.MouseHover, pic1216.MouseHover, pic1217.MouseHover, pic1218.MouseHover, pic1219.MouseHover,
            pic1220.MouseHover, pic1221.MouseHover, pic1222.MouseHover, pic1223.MouseHover, pic1224.MouseHover, pic1225.MouseHover, pic1226.MouseHover, pic1227.MouseHover, pic1228.MouseHover, pic1229.MouseHover
        picHover(sender)
    End Sub

    Private Sub pic1300_MouseHover(sender As Object, e As EventArgs) Handles pic1300.MouseHover, pic1301.MouseHover, pic1302.MouseHover, pic1303.MouseHover, pic1304.MouseHover, pic1305.MouseHover, pic1306.MouseHover, pic1307.MouseHover, pic1308.MouseHover, pic1309.MouseHover,
            pic1310.MouseHover, pic1311.MouseHover, pic1312.MouseHover, pic1313.MouseHover, pic1314.MouseHover, pic1315.MouseHover, pic1316.MouseHover, pic1317.MouseHover, pic1318.MouseHover, pic1319.MouseHover,
            pic1320.MouseHover, pic1321.MouseHover, pic1322.MouseHover, pic1323.MouseHover, pic1324.MouseHover, pic1325.MouseHover, pic1326.MouseHover, pic1327.MouseHover, pic1328.MouseHover, pic1329.MouseHover
        picHover(sender)
    End Sub

    Private Sub pic1400_MouseHover(sender As Object, e As EventArgs) Handles pic1400.MouseHover, pic1401.MouseHover, pic1402.MouseHover, pic1403.MouseHover, pic1404.MouseHover, pic1405.MouseHover, pic1406.MouseHover, pic1407.MouseHover, pic1408.MouseHover, pic1409.MouseHover,
            pic1410.MouseHover, pic1411.MouseHover, pic1412.MouseHover, pic1413.MouseHover, pic1414.MouseHover, pic1415.MouseHover, pic1416.MouseHover, pic1417.MouseHover, pic1418.MouseHover, pic1419.MouseHover,
            pic1420.MouseHover, pic1421.MouseHover, pic1422.MouseHover, pic1423.MouseHover, pic1424.MouseHover, pic1425.MouseHover, pic1426.MouseHover, pic1427.MouseHover, pic1428.MouseHover, pic1429.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic1500_MouseHover(sender As Object, e As EventArgs) Handles pic1500.MouseHover, pic1501.MouseHover, pic1502.MouseHover, pic1503.MouseHover, pic1504.MouseHover, pic1505.MouseHover, pic1506.MouseHover, pic1507.MouseHover, pic1508.MouseHover, pic1509.MouseHover,
            pic1510.MouseHover, pic1511.MouseHover, pic1512.MouseHover, pic1513.MouseHover, pic1514.MouseHover, pic1515.MouseHover, pic1516.MouseHover, pic1517.MouseHover, pic1518.MouseHover, pic1519.MouseHover,
            pic1520.MouseHover, pic1521.MouseHover, pic1522.MouseHover, pic1523.MouseHover, pic1524.MouseHover, pic1525.MouseHover, pic1526.MouseHover, pic1527.MouseHover, pic1528.MouseHover, pic1529.MouseHover
        picHover(sender)
    End Sub
    Private Sub pic1600_MouseHover(sender As Object, e As EventArgs) Handles pic1600.MouseHover, pic1601.MouseHover, pic1602.MouseHover, pic1603.MouseHover, pic1604.MouseHover, pic1605.MouseHover, pic1606.MouseHover, pic1607.MouseHover, pic1608.MouseHover, pic1609.MouseHover,
            pic1610.MouseHover, pic1611.MouseHover, pic1612.MouseHover, pic1613.MouseHover, pic1614.MouseHover, pic1615.MouseHover, pic1616.MouseHover, pic1617.MouseHover, pic1618.MouseHover, pic1619.MouseHover,
            pic1620.MouseHover, pic1621.MouseHover, pic1622.MouseHover, pic1623.MouseHover, pic1624.MouseHover, pic1625.MouseHover, pic1626.MouseHover, pic1627.MouseHover, pic1628.MouseHover, pic1629.MouseHover
        picHover(sender)
    End Sub
    Private Sub NewLayoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewLayoutToolStripMenuItem.Click
        ' File-New Layout
        frmHelp.HelpWindowClose()
        LayoutCreate()
    End Sub

    Private Sub StatusVariablesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatusVariablesToolStripMenuItem.Click
        Message("STAND_ALONE=      " & STAND_ALONE & ", MODE=" & MODE & ", ROUTE_VIEW_ENABLED=" & Str(ROUTE_VIEW_ENABLED))
        Message("FILE_LOADED=      " & Str(FILE_LOADED) & ", Page_No=" & Format(Page_No) & ", Test level=" & Format(Test_Level))
        Message("Start_Directory=  " & Start_Directory)
        Message("Current_Directory=" & Current_Directory)
        Message("TrackDiagram_Dir= " & TrackDiagram_Directory)
        Message("MasterFile_Direct=" & MasterFile_Directory) ' where gleisbild.cs2 and fahrstrassen.cs2 resides, if those files are present.
        Message("Flat_Structure=   " & Str(Flat_Structure) & ", Input_Drive=" & Input_Drive)
        Message("TrackDiagram_Input_Filename=" & TrackDiagram_Input_Filename)
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        CloseShop()
    End Sub

    Private Sub btnAddLayoutPage_Click(sender As Object, e As EventArgs) Handles btnAddLayoutPage.Click
        Dim track_diagram_filename As String, page_name As String, c As String, answer As Integer
        Dim fno As Integer
        If FILE_LOADED Then
            If CHANGED Then
                Dim layout_name As String
                layout_name = FilenameProper(TrackDiagram_Input_Filename)

                If STAND_ALONE Then answer = MsgBox("Track diagram """ & layout_name & """ has been modified but not saved. Should changes be saved?", vbYesNo)
                If answer = vbYes Or Not STAND_ALONE Then
                    If GleisbildSave() Then
                        If STAND_ALONE Then Message("Track diagram saved in: " & TrackDiagram_Input_Filename)
                    Else
                        MsgBox("Page could not be saved: " & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)))
                    End If
                End If
                CHANGED = False
            End If

            If PFH.PageFileExists(MasterFile_Directory) Then
                answer = Val(InputBox("Do you want to create a new page from scratch (1) or load an external file into your new page (2)", "Add page choice", "1"))
                Select Case answer
                    Case 1
                        page_name = InputBox("Name of the new page", "Page name", "")
                        If page_name = "gleisbild" Or page_name = "fahrstrassen" Then
                            MsgBox("The name """ & page_name & """ cannot be used as a name for a track diagram")
                            Exit Sub
                        End If

                        If page_name <> "" Then
                            c = ValidateFileName(page_name)
                            If c = "" Then
                                track_diagram_filename = MasterFile_Directory & "gleisbilder\" & page_name & ".cs2"

                                ' Create the track diagram file
                                Page_No = PFH.NextAvailablePageNumber

                                On Error GoTo write_error
                                fno = FreeFile()
                                FileOpen(fno, track_diagram_filename, OpenMode.Output)
                                PrintLine(fno, "[gleisbildseite]")
                                PrintLine(fno, "page=" & Format(Page_No))
                                PrintLine(fno, "version")
                                PrintLine(fno, " .major=1")
                                On Error GoTo 0
                                FileClose(fno)

                                PFH.PageAdd(track_diagram_filename)
                                PFH.PageFileSave()

                                GleisbildLoad(track_diagram_filename)
                                Call cmdEdit_Click(sender, e)
                            Else
                                MsgBox("New page was not created, unwanted character in page name: '" & c & "'")
                            End If
                        Else
                            MsgBox("New page was not created")
                        End If
                    Case 2
                        If Not GleisBildImportExternalFile() Then
                            MsgBox("New page wa not created")
                        End If
                    Case Else
                        MsgBox("New page was not created")
                End Select
            Else
                MsgBox("Warning: The layout page file ""gleisbild.cs2"" cannot be located")
            End If
        Else
            MsgBox("First load a track diagram from an existing layout")
        End If
        Exit Sub
write_error:
        On Error GoTo 0
        MsgBox("New page was not created due to error when writing the track diagram file: """ & track_diagram_filename & """")
    End Sub

    Private Sub btnDuplicatePage_Click(sender As Object, e As EventArgs) Handles btnDuplicatePage.Click
        Dim new_track_diagram_filename As String, new_page_name As String, new_page_no As Integer
        Dim current_track_diagram_filename As String, current_page_name As String, current_page_number As Integer
        Dim c As String, answer As DialogResult
        If FILE_LOADED Then
            If CHANGED Then
                current_page_name = FilenameProper(TrackDiagram_Input_Filename)
                If STAND_ALONE Then answer = MsgBox("Track diagram """ & current_page_name & """ has been modified but not saved. Should changes be saved?", vbYesNo)
                If answer = vbYes Or Not STAND_ALONE Then
                    If GleisbildSave() Then
                        If STAND_ALONE Then Message("Track diagram saved in: " & TrackDiagram_Input_Filename)
                    Else
                        MsgBox("Page could not be saved: " & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)))
                        If STAND_ALONE Then Message("to folder " & FilenamePath(TrackDiagram_Input_Filename))
                    End If
                End If
            End If

            If PFH.PageFileExists(MasterFile_Directory) Then
                ' Defensive programming. We make triple certain that the file to copy exists
                current_track_diagram_filename = TrackDiagram_Input_Filename
                current_page_number = PFH.LayoutPageNumberGet(TrackDiagram_Input_Filename)
                current_page_name = PFH.LayoutNameGet(current_page_number)

                new_page_name = InputBox("Name of the new page", "Page name", "")
                If new_page_name = "gleisbild" Or new_page_name = "fahrstrassen" Then
                    MsgBox("The name """ & new_page_name & """ cannot be used as a name for a track diagram")
                    Exit Sub
                End If

                If new_page_name <> "" And current_page_number <> -1 Then
                    c = ValidateFileName(new_page_name)
                    If c = "" Then
                        new_track_diagram_filename = MasterFile_Directory & "gleisbilder\" & new_page_name & ".cs2"
                        Message("Duplicating page """ & current_page_name & """ as """ & new_page_name & """")

                        ' Create the new track diagram file
                        new_page_no = PFH.NextAvailablePageNumber

                        'fix the page number
                        ElementsRepaginate(new_page_no)

                        ' Switch to the new name and save the file
                        Page_No = new_page_no
                        TrackDiagram_Input_Filename = new_track_diagram_filename

                        If Not GleisbildSave() Then
                            ' switch back to the original file and reload it to get the page number in order
                            TrackDiagram_Input_Filename = current_track_diagram_filename
                            Page_No = current_page_number
                            GleisbildLoad(TrackDiagram_Input_Filename)
                            MsgBox("New page was not created due to error when writing the copy to disk")
                        Else
                            ' Update the page file
                            PFH.PageAdd(new_track_diagram_filename)
                            PFH.PageFileSave()
                            GleisBildDisplayClear()
                            GleisbildLoad(new_track_diagram_filename)
                            GleisbildDisplay()
                        End If
                    Else
                        MsgBox("New page was not created, unwanted character in page name: '" & c & "'")
                    End If
                Else
                    MsgBox("New page was not created")
                End If
            Else
                MsgBox("Warning: The layout page file ""gleisbild.cs2"" cannot be located")
            End If
        Else
            MsgBox("First load a track diagram")
        End If
    End Sub

    Private Sub btnRemoveLayoutPage_Click(sender As Object, e As EventArgs) Handles btnRemoveLayoutPage.Click
        Dim delete_page_filename As String, delete_page_name As String, delete_page_no As Integer
        Dim answer As DialogResult
        If FILE_LOADED Then
            If PFH.PageFileExists(MasterFile_Directory) Then
                If PFH.NumberOfPages = 1 Then
                    MsgBox("You cannot delete the one and only page of the layout")
                Else
                    delete_page_no = Page_No
                    delete_page_name = PFH.LayoutNameGet(delete_page_no)
                    delete_page_filename = PFH.LayoutFileNameGet(delete_page_no)
                    GleisbildCreateBackup(delete_page_filename)

                    If delete_page_filename <> TrackDiagram_Input_Filename Then ' should never happen
                        MsgBox("Error in program logic. Cannot delete page")
                        Exit Sub
                    End If
                    answer = MsgBox("Remove page """ & delete_page_name & """ and the corresponding track diagram?", vbYesNo, vbNo)
                    If answer = vbYes Then
                        If Developer_Environment Then Message("Removing page number " & Str(Page_No) & ": " & delete_page_name)

                        ' pages which now get a new page number must be changed accordingly. We need to load them, modify them, and save
                        For pno = delete_page_no + 1 To (PFH.NumberOfPages - 1)
                            GleisbildLoad(PFH.LayoutFileNameGet(pno))
                            Page_No = Page_No - 1
                            ElementsRepaginate(pno - 1)
                            GleisbildSave()
                        Next

                        ' update the page file
                        PFH.PageRemove(delete_page_no)
                        PFH.PagesPopulate()

                        GleisBildDisplayClear()
                        GleisbildLoad(PFH.LayoutFileNameGet(0))
                        GleisbildDisplay()

                        Try
                            Kill(delete_page_filename)
                        Catch
                            MsgBox("Page was deleted, but not the corresponding track diagram file could not be deleted")
                            Exit Sub
                        End Try
                    End If
                End If
            Else
                MsgBox("Warning: The layout page file ""gleisbild.cs2"" cannot be located")
            End If
        Else
            MsgBox("First load a track diagram")
        End If

    End Sub

    Private Sub btnRenameLayoutPage_Click(sender As Object, e As EventArgs) Handles btnRenameLayoutPage.Click
        Dim old_track_diagram_filename As String, new_track_diagram_filename As String, old_page_name As String, new_page_name As String, c As String
        Dim fno As Integer
        If FILE_LOADED Then
            If PFH.PageFileExists(MasterFile_Directory) Then
                old_page_name = PFH.LayoutNameGet(Page_No)
                If old_page_name <> "" Then
                    new_page_name = InputBox("Enter the new name for page """ & old_page_name & """", "Page name", "")
                    If new_page_name = "gleisbild" Or new_page_name = "fahrstrassen" Then
                        MsgBox("The name """ & new_page_name & """ cannot be used as a name for a track diagram")
                        Exit Sub
                    End If

                    If old_page_name <> new_page_name And new_page_name <> "" Then
                        c = ValidateFileName(new_page_name)
                        If c = "" Then
                            old_track_diagram_filename = PFH.LayoutFileNameGet(Page_No)
                            new_track_diagram_filename = MasterFile_Directory & "gleisbilder\" & new_page_name & ".cs2"

                            On Error GoTo rename_error
                            Message(1, "Rename " & old_track_diagram_filename)
                            Message(1, "to     " & new_track_diagram_filename)
                            Rename(old_track_diagram_filename, new_track_diagram_filename)
                            On Error GoTo 0
                            PFH.PageReplaceName(Page_No, new_track_diagram_filename) ' side effect: saves the page file
                            PFH.PagesPopulate()
                            TrackDiagram_Input_Filename = new_track_diagram_filename
                            Me.Text = "Track Diagram Editor  .NET  [" & Version & "]   " & FilenameNoExtension(FilenameExtract(TrackDiagram_Input_Filename)) &
                                 "   " & FilePathUp(MasterFile_Directory)

                            env1.SetEnv("TrackDiagram_Input_Filename", TrackDiagram_Input_Filename)
                        Else
                            MsgBox("Page was not renamed: Unwanted character in page name: '" & c & "'")
                        End If
                    Else
                        MsgBox("Page was not renamed")
                    End If
                Else
                    MsgBox("No page was selected for renaming")
                End If
            Else
                MsgBox("Warning: The layout page file ""gleisbild.cs2"" cannot be located")
            End If
        Else
            MsgBox("First load a track diagram")
        End If
        Exit Sub
rename_error:
        On Error GoTo 0
        MsgBox("There was an error during renaming. Page was not renamed")
    End Sub

    Private Sub cmdManagePages_Click(sender As Object, e As EventArgs) Handles cmdManagePages.Click
        MODE = "View"
        Message(1, "Mode: View/Manage Pages")
        Message("Click an element to see its properties")

        grpMoveWindow.Enabled = True
        grpIcons.Enabled = True
        grpEditActions.Enabled = True

        el_moving = Empty_Element
        FileToolStripMenuItem.Enabled = True
        ToolsToolStripMenuItem.Enabled = True
        AllPagesToolStripMenuItem2.Enabled = True
        AllPagesToolStripMenuItem3.Enabled = True
        QuitToolStripMenuItem.Enabled = True
        CHANGED = CHANGED Or EDIT_CHANGES
        EDIT_CHANGES = False
        Active_Tool = ""
        AdjustPage()
        GleisbildGridSet(False)
        GleisBildDisplayClear()
        GleisbildDisplay()
    End Sub

    Private Sub ShowInputFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowInputFileToolStripMenuItem.Click
        lstInput.Visible = Not lstInput.Visible
        grpLayoutPages.Visible = Not lstInput.Visible
        ShowInputFileToolStripMenuItem.Checked = lstInput.Visible
    End Sub

    Private Sub cmdRevealAddressesPage_Click(sender As Object, e As EventArgs) Handles cmdRevealAddressesPage.Click
        Reveal_Addr = Not Reveal_Addr
        GleisbildDisplay() ' GleisbildDisplayAddresses()
    End Sub
    Private Sub cmdRevealTextPage_Click(sender As Object, e As EventArgs) Handles cmdRevealTextPage.Click
        Reveal_Text = Not Reveal_Text
        GleisbildDisplay() ' GleisbildDisplayTexts()
    End Sub

    Private Sub cmdReturnToTC_Click(sender As Object, e As EventArgs) Handles cmdReturnToTC.Click
        CloseShop()
    End Sub

    Private Sub cmdMoveLeft_Click(sender As Object, e As EventArgs) Handles cmdMoveLeft.Click
        ' Moves entire layout left by one column
        ' Note that row_no and col_no of the element can be displaced depending on the viewing window's position.
        ' Therefore we must rely on the ID for row/col information
        Dim i As Integer, el As ELEMENT, s As String
        Dim rno As Integer, cno As Integer ' row/col numvers as represented in the ID
        For i = 1 To E_TOP
            cno = DecodeColNo(Elements(i).id_normalized)
            If cno = 0 Then
                MsgBox("Layout cannot be moved further left")
                Exit Sub
            End If
        Next
        For i = 1 To E_TOP
            el = Elements(i)
            cno = DecodeColNo(el.id_normalized)
            rno = DecodeRowNo(el.id_normalized)
            el.col_no = el.col_no - 1
            cno = cno - 1
            s = DecIntToHexStr(Page_No) & DecIntToHexStr(rno) & DecIntToHexStr(cno)
            el.id_normalized = s
            el.id = CS2_id_create(Page_No, rno, cno)
            Elements(i) = el
        Next
        EDIT_CHANGES = True
        GleisBildDisplayClear()
        GleisbildDisplay()
    End Sub

    Private Sub cmdMoveUp_Click(sender As Object, e As EventArgs) Handles cmdMoveUp.Click
        ' Moves entire layout up by one row
        ' Note that row_no and col_no of the element can be displaced depending on the viewing window's position.
        ' Therefore we must rely on the ID for row/col information
        Dim i As Integer, el As ELEMENT, s As String
        Dim rno As Integer, cno As Integer ' row/col numvers as represented in the ID
        For i = 1 To E_TOP
            rno = DecodeRowNo(Elements(i).id_normalized)
            If rno = 0 Then
                MsgBox("Layout cannot be moved further up")
                Exit Sub
            End If
        Next
        For i = 1 To E_TOP
            el = Elements(i)
            cno = DecodeColNo(el.id_normalized)
            rno = DecodeRowNo(el.id_normalized)
            el.row_no = el.row_no - 1
            rno = rno - 1
            s = DecIntToHexStr(Page_No) & DecIntToHexStr(rno) & DecIntToHexStr(cno)
            el.id_normalized = s
            el.id = CS2_id_create(Page_No, rno, cno)
            Elements(i) = el
        Next
        EDIT_CHANGES = True
        GleisBildDisplayClear()
        GleisbildDisplay()
    End Sub

    Private Sub cmdMoveRight_Click(sender As Object, e As EventArgs) Handles cmdMoveRight.Click
        ' Moves entire layout right by one column.
        ' Note that row_no and col_no of the element can be displaced depending on the viewing window's position.
        ' Therefore we must rely on the ID for row/col information
        Dim i As Integer, el As ELEMENT, s As String
        Dim rno As Integer, cno As Integer ' row/col numvers as represented in the ID
        For i = 1 To E_TOP
            cno = DecodeColNo(Elements(i).id_normalized)
            If cno >= 255 Then
                MsgBox("Track diagram cannot be moved further right")
                Exit Sub
            End If
        Next
        For i = 1 To E_TOP
            el = Elements(i)
            cno = DecodeColNo(el.id_normalized)
            rno = DecodeRowNo(el.id_normalized)
            el.col_no = el.col_no + 1
            cno = cno + 1
            s = DecIntToHexStr(Page_No) & DecIntToHexStr(rno) & DecIntToHexStr(cno)
            el.id_normalized = s
            el.id = CS2_id_create(Page_No, rno, cno)
            Elements(i) = el
        Next
        EDIT_CHANGES = True
        GleisBildDisplayClear()
        GleisbildDisplay()
    End Sub

    Private Sub cmdMoveDown_Click(sender As Object, e As EventArgs) Handles cmdMoveDown.Click
        ' Moves entire layout down by one row space permitting
        ' Note that row_no and col_no of the element can be displaced depending on the viewing window's position.
        ' Therefore we must rely on the ID for row/col information
        Dim i As Integer, el As ELEMENT, s As String
        Dim rno As Integer, cno As Integer ' row/col numvers as represented in the ID
        For i = 1 To E_TOP
            rno = DecodeRowNo(Elements(i).id_normalized)
            If rno >= 255 Then
                MsgBox("Track diagram cannot be moved further down")
                Exit Sub
            End If
        Next
        For i = 1 To E_TOP
            el = Elements(i)
            cno = DecodeColNo(el.id_normalized)
            rno = DecodeRowNo(el.id_normalized)
            el.row_no = el.row_no + 1
            rno = rno + 1
            s = DecIntToHexStr(Page_No) & DecIntToHexStr(rno) & DecIntToHexStr(cno)
            el.id_normalized = s
            el.id = CS2_id_create(Page_No, rno, cno)
            Elements(i) = el
        Next
        EDIT_CHANGES = True
        GleisBildDisplayClear()
        GleisbildDisplay()
    End Sub

    Private Sub btnWindowMoveDown_Click(sender As Object, e As EventArgs) Handles btnWindowMoveDown.Click
        WindowMove(5, 0)
    End Sub

    Private Sub btnWindowMoveUp_Click(sender As Object, e As EventArgs) Handles btnWindowMoveUp.Click
        WindowMove(-5, 0)
    End Sub

    Private Sub btnWindowMoveLeft_Click(sender As Object, e As EventArgs) Handles btnWindowMoveLeft.Click
        WindowMove(0, -5)
    End Sub

    Private Sub btnWindowMoveRight_Click(sender As Object, e As EventArgs) Handles btnWindowMoveRight.Click
        WindowMove(0, 5)
    End Sub

    Private Sub btnWindowReset_Click(sender As Object, e As EventArgs) Handles btnWindowReset.Click
        WindowMove(-Row_Offset, -Col_Offset)
    End Sub

    Private Sub btnJoinPages_Click(sender As Object, e As EventArgs) Handles btnCombinePages.Click
        CPH.initiate
    End Sub

    Private Sub btnDisplayRefreshEdit_Click(sender As Object, e As EventArgs) Handles btnDisplayRefreshEdit.Click
        ' Toggle visibility of addresses and text. Since they could overwrite each other it is either or, or none.
        If Reveal_Addr Then
            Reveal_Addr = False
            Reveal_Text = True
        ElseIf Reveal_Text Then
            Reveal_Addr = False
            Reveal_Text = False
        Else
            Reveal_Addr = True
        End If
        GleisbildDisplay()
    End Sub

    Private Sub CurrentPageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CurrentPageToolStripMenuItem2.Click
        ' Shows addresses from the currently selected page only
        Const ww As Integer = 350
        If E_TOP > 0 Then
            If Me.Left <= ww Then
                frmHelp.HelpWindowSetPosStoredButWidth(Me.Left + 20, Me.Top + 50, ww, Me.Height - 50)
            Else
                frmHelp.HelpWindowSetPosStoredButWidth(Me.Left - ww, Me.Top, ww, Me.Height)
            End If
            frmHelp.HelpDisplaySortedListInitialize("Addresses in Use")
            frmHelp.HelpDisplaySortedListAdd("       page: " & FilenameProper(TrackDiagram_Input_Filename))
            frmHelp.HelpDisplaySortedListAdd("      element               row,col")
            ShowAddressesHelper("")
        End If
    End Sub

    Private Sub AllPagesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AllPagesToolStripMenuItem2.Click
        ' Shows addresses from all pages
        Const ww As Integer = 450
        Dim pno As Integer, fname As String, page_name As String, active_file As String
        If FILE_LOADED AndAlso PFH.PageFileExists(MasterFile_Directory) Then
            active_file = TrackDiagram_Input_Filename
            If Me.Left <= ww Then
                frmHelp.HelpWindowSetPosStoredButWidth(Me.Left + 20, Me.Top + 50, ww, Me.Height - 50)
            Else
                frmHelp.HelpWindowSetPosStoredButWidth(Me.Left - ww, Me.Top, ww, Me.Height)
            End If
            frmHelp.HelpDisplaySortedListInitialize("Addresses in Use")
            frmHelp.HelpDisplaySortedListAdd("                            row,col   page name")
            Do
                fname = PFH.LayoutFileNameGet(pno)
                If fname = "" Then Exit Do
                page_name = PFH.LayoutNameGet(pno)
                GleisbildLoad(fname)
                ShowAddressesHelper(page_name)
                pno = pno + 1
            Loop
            GleisbildLoad(active_file)
            GleisbildDisplay()
        Else
            MsgBox("First load a track diagram")
        End If
    End Sub

    Private Sub ShowAddressesHelper(page_name As String)
        Dim ei As Integer, el As ELEMENT, s As String
        For ei = 1 To E_TOP
            el = Elements(ei)
            If el.artikel_i > 0 AndAlso Strings.Left(el.typ_s, 3) <> "s88" AndAlso el.typ_s <> "fahrstrasse" AndAlso el.typ_s <> "pfeil" Then
                s = Strings.RSet(DigitalAddress(el), 4)
                If Strings.Right(s, 1) = "b" Then
                    s = " " & s
                Else
                    s = s & " "
                End If
                s = s & " " & Strings.LSet(Translate(el.typ_s), 20) &
                                    " (" & Strings.RSet(Format(DecodeRowNo(el.id_normalized)), 3) & "," & Strings.RSet(Format(DecodeColNo(el.id_normalized)), 3) & ")"
                If page_name <> "" Then s = s & "  " & page_name
                frmHelp.HelpDisplaySortedListAdd(s)
            End If
        Next ei
    End Sub
    Private Sub CurrentPageToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CurrentPageToolStripMenuItem3.Click
        ' Shows s88 from current page
        Const ww As Integer = 350
        If E_TOP > 0 Then
            If Me.Left <= ww Then
                frmHelp.HelpWindowSetPosStoredButWidth(Me.Left + 20, Me.Top + 50, ww, Me.Height - 50)
            Else
                frmHelp.HelpWindowSetPosStoredButWidth(Me.Left - ww, Me.Top, ww, Me.Height)
            End If
            frmHelp.HelpDisplaySortedListInitialize("s88 in Use")
            frmHelp.HelpDisplaySortedListAdd(" ")
            frmHelp.HelpDisplaySortedListAdd("            page: " & FilenameProper(TrackDiagram_Input_Filename))
            frmHelp.HelpDisplaySortedListAdd("           element                row,col")
            Shows88Helper("")
        End If
    End Sub

    Private Sub AllPagesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AllPagesToolStripMenuItem3.Click
        ' Shows s88 from all pages
        Const ww = 500
        Dim pno As Integer, fname As String, page_name As String, active_file As String
        If FILE_LOADED AndAlso PFH.PageFileExists(MasterFile_Directory) Then
            active_file = TrackDiagram_Input_Filename
            If Me.Left <= ww Then
                frmHelp.HelpWindowSetPosStoredButWidth(Me.Left + 20, Me.Top + 50, ww, Me.Height - 50)
            Else
                frmHelp.HelpWindowSetPosStoredButWidth(Me.Left - ww, Me.Top, ww, Me.Height)
            End If
            frmHelp.HelpDisplaySortedListInitialize("s88 in Use")
            frmHelp.HelpDisplaySortedListAdd("                                  row,col   page name")
            Do
                fname = PFH.LayoutFileNameGet(pno)
                If fname = "" Then Exit Do
                page_name = PFH.LayoutNameGet(pno)
                GleisbildLoad(fname)
                Shows88Helper(page_name)
                pno = pno + 1
            Loop
            GleisbildLoad(active_file)
            GleisbildDisplay()
        Else
            MsgBox("First load a track diagram")
        End If
    End Sub

    Private Sub Shows88Helper(page_name As String)
        Dim ei As Integer, el As ELEMENT, s As String
        For ei = 1 To E_TOP
            el = Elements(ei)
            If el.artikel_i > 0 AndAlso Strings.Left(el.typ_s, 3) = "s88" Then
                s = Strings.LSet(el.deviceId_s, 5) & " " & Strings.RSet(DigitalAddress(el), 4)
                s = s & " " & Strings.LSet(Translate(el.typ_s), 22) &
                    "(" & Strings.RSet(Format(DecodeRowNo(el.id_normalized)), 3) & "," & RSet(Format(DecodeColNo(el.id_normalized)), 3) & ")"
                If page_name <> "" Then s = s & "  " & page_name
                frmHelp.HelpDisplaySortedListAdd(s)
            End If
        Next ei
    End Sub


    Private Sub lstLayoutPages_Click(sender As Object, e As EventArgs) Handles lstLayoutPages.Click
        Dim pno As Integer, li As Integer, it As String, fn As String

        li = lstLayoutPages.SelectedIndex
        it = lstLayoutPages.SelectedItem
        pno = li
        If it <> "" Then
            fn = PFH.LayoutFileNameGet(pno)
            If fn <> "" Then
                If MODE <> "Join" Then
                    GleisBildDisplayClear()
                    GleisbildLoad(fn)
                    GleisbildDisplay()
                Else
                    CPH.PageSelected(it, fn)
                End If
            Else
                MsgBox("Track diagram file missing: " & fn) ' should never happen
            End If
        End If
    End Sub

End Class
