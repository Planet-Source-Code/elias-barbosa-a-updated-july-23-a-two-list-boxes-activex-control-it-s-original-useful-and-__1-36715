Attribute VB_Name = "modGeneral"
Option Explicit
'----------------------------------------------
'API necessary to put horizontal scroll bar
'on list boxes.
Private Declare Function SendMessage _
    Lib "user32" _
    Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByRef lParam As Any) As Long

Private Const LB_SETHORIZONTALEXTENT = &H194
'----------------------------------------------

Public intDB As Database
Public intRS As Recordset
Public intQD As QueryDef

Public m_DatabaseName As String
Public m_Password As String
Public m_RecordSource As String
Public m_IDFieldName As String
Public m_FieldName As String
Public m_QDField As String
Public m_QDFilter As String
Public m_SortBy As String
Public bolHasScrollBar1 As Boolean
Public bolHasScrollBar2 As Boolean
Public strFinalSelect As String

Private intScrollWidth1 As Integer
Private intScrollHeight1 As Integer
Private intScrollWidth2 As Integer
Private intScrollHeight2 As Integer
Private lngScrollLength As Long

'=============================================
'This Function is to find out short file names
'=============================================
Private Declare Function GetShortPathName _
    Lib "kernel32" _
    Alias "GetShortPathNameA" ( _
    ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

Public Function ShortFilename(sFIle As String) As String
    Dim stemp As String
    Dim l As Long
    
    stemp = Space(200)
    l = GetShortPathName(sFIle, stemp, Len(stemp))
    
    If l = 0 Then
        ShortFilename = sFIle
    Else
        ShortFilename = Left(stemp, l)
    End If
    
End Function

'=============================================
'=============================================

Public Function AddHorizScroll(m_List As ListBox, m_Form As Form) As Boolean
    Dim i As Integer
    Dim lngGreatestWidth As Long
    Dim txtGratestText As String
    Dim intTextHeight As Long
    Dim intLoopTo As Integer
    Dim myFont As String
    Dim myFontBold As Boolean
    Dim myFontItalic As Boolean
    Dim myFontSize As Long
    
    'Don't do anything if the ListBox has no item.
    If (m_List.ListCount = 0) Then
        'Use API to eliminate horizontal ScrollBar.
        Call SendMessage(m_List.hwnd, LB_SETHORIZONTALEXTENT, 1, 0)
        
        Exit Function
        
    End If
    
    With m_Form
        'Backup original font format...
        myFont = .Font
        myFontBold = .FontBold
        myFontItalic = .FontItalic
        myFontSize = .FontSize
        
        'Adapt font format to the ListBox...
        .Font = m_List.Font
        .FontBold = m_List.FontBold
        .FontItalic = m_List.FontItalic
        .FontSize = m_List.FontSize
        
        'Get text hight on the form...
        intTextHeight = .TextHeight("Anything")
    End With
    
    'Calculate how many items can be viewed at
    'once on the ListBox at any given time.
    intLoopTo = (m_List.TopIndex + Int(m_List.Height / intTextHeight))
    
    'Loop as many times as the number of
    'items that can be viewed at once.
    intLoopTo = IIf((intLoopTo > m_List.ListCount), m_List.ListCount, intLoopTo)
    
    'Start looping from the topmost item
    'that can be viewed on the ListBox.
    For i = m_List.TopIndex To intLoopTo - 1
        
        'Find Longest Text in ListBox...
        If Len(m_List.List(i)) > Len(txtGratestText) Then
            txtGratestText = m_List.List(i)
        End If
        
    Next i
    
    With m_Form
        
        'Get Twips...
        lngGreatestWidth = .TextWidth(txtGratestText & Space(1))
        'A space is added to prevent the
        'last Character from being cut off.
        
        On Error GoTo OverFlow
        
        'Determine which list
        'box is calling the sub.
        If (m_List.Name = "List1") Then
            'Determine whether there will be a
            'vertical scroll bar or not.
            If ((m_List.ListCount * intTextHeight) > (m_List.Height - intScrollHeight1)) Then
                intScrollWidth1 = 280
                
            Else
                intScrollWidth1 = 0
                
            End If
            
            If (lngGreatestWidth > (m_List.Width - intScrollWidth1)) Then
                intScrollHeight1 = 280
                
                bolHasScrollBar1 = True
                
            Else
                If (bolHasScrollBar1) Then
                    Exit Function
                End If
                
                intScrollHeight1 = 0
                
            End If
        Else
            'Determine whether there will be a
            'vertical scroll bar or not.
            If ((m_List.ListCount * intTextHeight) > (m_List.Height - intScrollHeight2)) Then
                intScrollWidth2 = 280
                
            Else
                intScrollWidth2 = 0
                
            End If
            
            If (lngGreatestWidth > (m_List.Width - intScrollWidth2)) Then
                intScrollHeight2 = 280
                
                bolHasScrollBar2 = True
                
            Else
                If (bolHasScrollBar2) Then
                    Exit Function
                End If
                
                intScrollHeight2 = 0
                
            End If
        End If
        
        'Restore original font format...
        .Font = myFont
        .FontBold = myFontBold
        .FontItalic = myFontItalic
        .FontSize = myFontSize
        
    End With
    
    'Convert to Pixels.
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    
    
    'Use API to add horizontal ScrollBar.
    Call SendMessage(m_List.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0)
    
    AddHorizScroll = True
    
    Exit Function
OverFlow:
    
    If (Err.Number <> 6) Then
        Debug.Print Err.Number
        Debug.Print Err.Description
        
    Else
        Debug.Print "Overflow"
        
    End If
    AddHorizScroll = True
    
End Function

Public Sub AddToolTip(m_List As ListBox, m_Form As Form, Y As Single)
    Dim myFont As String
    Dim myFontBold As Boolean
    Dim myFontItalic As Boolean
    Dim myFontSize As Long
    Dim intScrollHeight As Integer
    Dim intItemMouseFromTop As Integer
    Dim intItemMouseOnTop As Integer
    Dim intScrollWidth As Integer
    Dim intTextHeight As Integer
    Dim intTextWidth As Integer
    Dim intListWidth As Integer
    
    'Don't do anything if the ListBox has no item.
    If (m_List.ListCount < 1) Then
        'Erase ToolTip...
        m_List.ToolTipText = ""
        
        Exit Sub
        
    End If
    
    With m_Form
        'Backup original font format...
        myFont = .Font
        myFontBold = .FontBold
        myFontItalic = .FontItalic
        myFontSize = .FontSize
        
        'Adapt font format to the ListBox...
        .Font = m_List.Font
        .FontBold = m_List.FontBold
        .FontItalic = m_List.FontItalic
        .FontSize = m_List.FontSize
        
        'Get text hight on the form...
        intTextHeight = .TextHeight("Anything")
        
        'Restore original font format...
        .Font = myFont
        .FontBold = myFontBold
        .FontItalic = myFontItalic
        .FontSize = myFontSize
        
    End With
    
    'Determine which list
    'box is calling the sub.
    If (m_List.Name = "List1") Then
        intScrollWidth = intScrollWidth1
        intScrollHeight = intScrollHeight1
        
    Else
        intScrollWidth = intScrollWidth2
        intScrollHeight = intScrollHeight2
        
    End If
    
    intItemMouseFromTop = Int(Y / intTextHeight)
    intItemMouseOnTop = intItemMouseFromTop + m_List.TopIndex
    
    intTextWidth = m_Form.TextWidth(m_List.List(intItemMouseOnTop)) + 60
    intListWidth = m_List.Width - intScrollWidth
    
    If (intTextWidth > intListWidth - intScrollWidth) Then
        m_List.ToolTipText = m_List.List(intItemMouseOnTop)
        
    Else
        m_List.ToolTipText = ""
        
    End If
        
End Sub

Public Function fntFileExist(intFileName As String) As Boolean
    Dim MyFileSystem As FileSystemObject
    Dim intFullPath As String
    
    intFileName = fntFullOrRelative(intFileName)
    
    Set MyFileSystem = CreateObject("Scripting.FileSystemObject")
    
    If (MyFileSystem.FileExists(intFileName)) Then
        fntFileExist = True
    Else
        fntFileExist = False
    End If
    
    Set MyFileSystem = Nothing
    
End Function

'App.Path returns a string with the "\" character at the end
'if the path is the root drive (e.g., "C:\") but without that
'character if it isn't (e.g., "C:\Program Files"). Most of the
'time we need the "\" at the end, so this function saves you
'the inconvenience of adding it every time.
Public Function AppPath() As String
    Dim strAppP As String
    
    strAppP = App.Path
    
    If (Right(App.Path, 1) <> "\") Then
        strAppP = strAppP & "\"
        
    End If
    
    AppPath = strAppP
    
End Function

Public Function Create_File(myFileContent As String, myFilePath As String) As String
    Dim fso As FileSystemObject
    Dim txtfile As TextStream
    
    myFilePath = fntFullOrRelative(myFilePath)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo ReadOnlyFile
    
    'The following statement will create a
    'file and will overwrite any file that
    'was previously there. However, if the
    'file is marked as read-only, it will
    'generate an error message.
    Set txtfile = fso.CreateTextFile(myFilePath, True)
    txtfile.Write (myFileContent) ' Write a line.
    txtfile.Close
    
    Create_File = "OK"
    
    Set fso = Nothing
    
    Exit Function
    
ReadOnlyFile:
    If (Err.Number = 70) Then
        Create_File = "ReadOnly"
        
    ElseIf (Err.Number = 5) Then
        MsgBox "There was an error while trying" & Chr(10) & _
               "to save your file!" & Chr(10) & Chr(10) & _
               "It looks like the path provided" & Chr(10) & _
               "to save the file is ilegal. " & Chr(10) & _
               "Please, send the following error" & Chr(10) & _
               "description to CBFSI:" & Chr(10) & Chr(10) & _
               "Error # " & Err.Number & Chr(10) & _
               Err.Description
              
        Create_File = "Error"
        
    Else
        MsgBox "There was an error while trying" & Chr(10) & _
               "to save your file!" & Chr(10) & Chr(10) & _
               "Please, send the following error" & Chr(10) & _
               "description to CBFSI." & Chr(10) & Chr(10) & _
               "Error " & Err.Number & ":"
        Create_File = "Error"
    End If
    
End Function

'The user can provid the full path or
'the relative path for a file.
'This function will always return the
'full path.
Public Function fntFullOrRelative(FilePath As String) As String
    If Not (Mid(FilePath, 2, 2) = ":\") Then
        FilePath = AppPath & FilePath
        
    End If
    
    fntFullOrRelative = FilePath
    
End Function

Public Function fntGetTextFile(FilePath As String) As String
    Dim fso As FileSystemObject
    Dim txtfile As TextStream
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    
    FilePath = fntFullOrRelative(FilePath)
    
    If (fntFileExist(FilePath)) Then
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        On Error GoTo NowStop
        'The following statement will create a
        'file and will overwrite any file that
        'was previously there. However, if the
        'file is marked as read-only, it will
        'generate an error message.
        Set txtfile = fso.OpenTextFile(FilePath, ForReading)
        fntGetTextFile = txtfile.ReadAll ' Write a line.
        
        txtfile.Close
        
        Set txtfile = Nothing
        Set fso = Nothing
        
    Else
        fntGetTextFile = ""
        
    End If

    Exit Function
    
NowStop:
    fntGetTextFile = ""
    'MsgBox "Error Number: " & Err.Number & Chr(10) & "Error Description: " & Err.Description
'Stop
    
End Function

Public Function fntSpreadStr(m_String As String) As String
    Dim bolKeepGoing  As Boolean
    Dim strFirstHalf As String
    Dim strLastHalf As String
    Dim intStartLoop As Integer
    Dim intFinishLoop As Integer
    Dim intStartNumber As Integer
    Dim intFinishNumber As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    bolKeepGoing = True
    
    Do While bolKeepGoing = True
        For i = 1 To Len(m_String)
            If (Mid(m_String, i, 1) = ",") Then
                intStartNumber = i
            End If
            
            If (Mid(m_String, i, 1) = "-") Then
                strFirstHalf = Left(m_String, i - 1)
                strLastHalf = Right(m_String, Len(m_String) - i)
                
                intStartLoop = Val(Right(strFirstHalf, Len(strFirstHalf) - intStartNumber)) + 1
                
                intFinishNumber = InStr(1, strLastHalf, ",")
                
                If (intFinishNumber = 0) Then
                    intFinishNumber = Len(strLastHalf)
                    
                Else
                    intFinishNumber = intFinishNumber - 1
                    
                End If
                
                intFinishLoop = Val(Left(strLastHalf, intFinishNumber)) - 1
                'intFinishLoop = Val(Mid(m_String, i + 1, 1)) - 1
                
                
                For j = intStartLoop To intFinishLoop
                    strFirstHalf = strFirstHalf & "," & j
                    
                Next j
                
                m_String = strFirstHalf & "," & strLastHalf
                
                bolKeepGoing = True
                Exit For
            End If
            bolKeepGoing = False
        Next i
    Loop
    
    fntSpreadStr = m_String
    
    Debug.Print m_String
    
End Function

Public Function fntGenerSQL(m_SelectedID As String) As String
    Dim intArraySingle As Variant
    Dim intArrayMulti As Variant
    Dim strSQLString As String
    Dim i As Integer
    
    intArraySingle = Split(m_SelectedID, ",")
    
    For i = 0 To UBound(intArraySingle)
        If (InStr(1, intArraySingle(i), "-") = 0) Then
            strSQLString = strSQLString & "ID=" & intArraySingle(i) & " OR "
        Else
            intArrayMulti = Split(intArraySingle(i), "-")
            
            strSQLString = strSQLString & "(ID>=" & intArrayMulti(0) & " AND ID<=" & intArrayMulti(1) & ") OR "
        End If
    Next i
    strSQLString = Left(strSQLString, Len(strSQLString) - 3)
    fntGenerSQL = strSQLString
    
    Debug.Print strSQLString
    
End Function

