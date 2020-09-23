Attribute VB_Name = "modCommonDialog"
'**************************************
'Windows API/Global Declarations for:
'    Common Dialog without OCX
'**************************************

Private Declare Function GetSaveFileName _
                Lib "comdlg32.dll" _
                Alias "GetSaveFileNameA" ( _
                pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetOpenFileName _
                Lib "comdlg32.dll" _
                Alias "GetOpenFileNameA" ( _
                pOpenfilename As OPENFILENAME) As Long

Private strfileName As OPENFILENAME

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hwndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2 '&H2 Or &H40
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
    OFN_ENABLEINCLUDENOTIFY = &H400000          '// send include message to callback
    OFN_ENABLESIZING = &H800000
    OFN_NOREADONLYRETURN_C = &H8000&
End Enum

Public intGetFileNametoSave
Public intGetFileNametoOpen
Public intExtChoosen

Private Sub DialogFilter(WantedFilter As String)
    Dim intLoopCount As Integer
    With strfileName
        .lpstrFilter = ""
    
        For intLoopCount = 1 To Len(WantedFilter)
            If Mid(WantedFilter, intLoopCount, 1) = "|" Then .lpstrFilter = _
            .lpstrFilter + Chr(0) Else .lpstrFilter = _
            .lpstrFilter + Mid(WantedFilter, intLoopCount, 1)
        Next intLoopCount
        
        .lpstrFilter = .lpstrFilter + Chr(0)
    End With
End Sub

'This is The Function To get the File Name to Open.
'Even If you don't specify a Title or a Filter it is OK.
Public Function fncGetFileNametoOpen(Optional strDialogTitle As String = "Open", Optional strFilter As String = "All Files|*.*", Optional strDefaultExtention As String = "*.*", Optional strInitialDir As String) As Boolean
    Dim lngReturnValue As Long
    Dim intRest As Integer
    
    If (strInitialDir = "") Then
        strInitialDir = App.Path
    End If
    
    With strfileName
        .lpstrTitle = strDialogTitle
        .lpstrDefExt = strDefaultExtention
        .hInstance = App.hInstance
        .lpstrFile = Chr(0) & Space(259)
        .nMaxFile = 260
        .flags = &H4
        .lStructSize = Len(strfileName)
        .lpstrInitialDir = strInitialDir
        
        DialogFilter (strFilter)
    
        lngReturnValue = GetOpenFileName(strfileName)
        
        'I want to return the string without any
        'unnecessary character. Therefore, I am
        'using the Trim function to eliminate any
        'space before and after the string. I am
        'using, also, the Replace function to find
        'and delete an invalid character "Chr(0)"
        'that is, usually, added to the string.
        intGetFileNametoOpen = Replace(Trim(.lpstrFile), Chr(0), "")
        
    End With
    
    'Check if user cancelled saving.
    Select Case lngReturnValue
       Case 1
          fncGetFileNametoOpen = True
          
       Case 0
          'Cancelled:
          fncGetFileNametoOpen = False
          
       Case Else
          'Extended error:
          fncGetFileNametoOpen = False
          
    End Select
    
End Function

'This Function Returns the Save File Name.
'Remember, you have to specify a Filter and
'default Extention for this.
Public Function fncGetFileNameToSave( _
                strFilter As String, _
                strDefaultExtention As String, _
                strInitFolder As String, _
                Optional strDialogTitle As String = "Save", _
                Optional Filename As String = "File Name" _
                ) As Boolean
                
    Dim lngReturnValue As Long
    Dim intRest As Integer
    Dim s
    With strfileName
        .lpstrTitle = strDialogTitle
        .lpstrDefExt = strDefaultExtention
        
        .hInstance = App.hInstance
        .lpstrFile = Chr(0) & Space(259)
        .nMaxFile = 260
        '.flags = OFN_OVERWRITEPROMPT '&H2 Or &H40
        .lStructSize = Len(strfileName)
        .lpstrInitialDir = strInitFolder

        
        'If (App.hInstance > 0) Then
           '.flags = .flags Or OFN_ENABLETEMPLATE
           '.lpTemplateName = 1
        'End If
        s = Filename & String$(260 - Len(Filename), 0)
        .lpstrFile = s
        '.nMaxFile = 260

        DialogFilter (strFilter)
        
        lngReturnValue = GetSaveFileName(strfileName)
        
        'I want to return the string without any
        'unnecessary character. Therefore, I am
        'using the Trim function to eliminate any
        'space before and after the string. I am
        'using, also, the Replace function to find
        'and delete an invalid character "Chr(0)"
        'that is, usually, added to the string.
        intGetFileNametoSave = Replace(Trim(.lpstrFile), Chr(0), "")
        intExtChoosen = .nFilterIndex
        
    End With
    
    'Check if user cancelled saving.
    Select Case lngReturnValue
       Case 1
          fncGetFileNameToSave = True
          
       Case 0
          'Cancelled:
          fncGetFileNameToSave = False
          
       Case Else
          'Extended error:
          fncGetFileNameToSave = False
          
    End Select
    
End Function

