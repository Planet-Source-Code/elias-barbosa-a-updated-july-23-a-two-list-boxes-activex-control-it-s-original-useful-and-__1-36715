VERSION 5.00
Begin VB.UserControl TwoLstBxs 
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   PropertyPages   =   "TwoLstBxs.ctx":0000
   ScaleHeight     =   3375
   ScaleWidth      =   6135
   ToolboxBitmap   =   "TwoLstBxs.ctx":0044
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Timer Timer1 
         Interval        =   5
         Left            =   2880
         Top             =   2760
      End
      Begin VB.ListBox List2 
         Height          =   2985
         ItemData        =   "TwoLstBxs.ctx":0356
         Left            =   3120
         List            =   "TwoLstBxs.ctx":0358
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         ToolTipText     =   "testando"
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox List1 
         Height          =   2985
         ItemData        =   "TwoLstBxs.ctx":035A
         Left            =   120
         List            =   "TwoLstBxs.ctx":035C
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         ToolTipText     =   "Testando"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Title2"
         Height          =   195
         Left            =   3120
         TabIndex        =   4
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Title1"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   390
      End
   End
End
Attribute VB_Name = "TwoLstBxs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'===========================================
'===== EB8 Two ListBoxes Version 1.1.4 =====
'===========================================
'
'This ActiveX Control will connect to a database and list all
'the records on the first List Box (Left side).
'
'The users will be able to select items and move them to the
'second List Box (Right side). They can double-click or press
'Enter to move an item from one List Box to the other. They
'can, also, select various items from one List Box and move
'them to the other by using the buttons found below each
'List Box.
'
'This ActiveX Control does a task that, at first glance, looks
'very simple. However, it took me almost a month to finish it.
'If you think that this control can be of any use for you,
'please, give me your vote and post some comments.
'
'Regards,
'Elias Barbosa
'
'
'Author:
'   Elias Barbosa
'   Date: 07/08/2002
'   e-mail: elias@eb8.com
'   http://www.planet-source-code.com/vb/default.asp?lngCId=36715&lngWId=1
'
'Updated:
'   Elias Barbosa
'   Date: 07/22/2002
'
Option Explicit

'Default Property Values:
Const m_def_SaveLists = True
Const m_def_AutoConnect = True
Const m_def_SortBy = ""
Const m_def_Password = ""
Const m_def_IDFieldName = ""
Const m_def_FieldName = ""
Const m_def_Caption1 = "Title1"
Const m_def_Caption2 = "Title2"
Const m_def_RecordSource = ""
Const m_def_DatabaseName = ""
Const m_def_SQLString = ""

'Property Variables:
Dim m_SaveLists As Boolean
Dim m_AutoConnect As Boolean
Dim m_Caption1 As String
Dim m_Caption2 As String
Dim m_DatabaseName As String
Dim m_RecordSource As String
Dim m_FieldName As String
Dim m_IDFieldName As String
Dim m_SortBy As String
Dim m_SQLString As String

Dim intInitSQL As String
Dim intFinalSQL As String
Dim Array1() As String
Dim Array2() As String
Dim myAmbient As Boolean
Dim intFileName As String
Dim ResizeLoop As Boolean
Public RSFinal As Recordset

Public Enum BorderStyle
   None = 0
   [Fixed Single] = 1 'This Enum constant is within square brackets because, by doing this, I am allowed to use spaces.
End Enum

Public Enum Appearance
   [Style Flat] = 0
   [Style 3D] = 1
   
End Enum

Private Sub UserControl_Initialize()
    'I have to initialize the arrays to avoid
    'an error that would happen if I tried to
    'check the UBound of the array before
    'redimensioning it.
    ReDim Array1(0, 0)
    ReDim Array2(0, 0)
    
End Sub

'Resize the frames to fit the user control.
Private Sub UserControl_Resize()
    On Error Resume Next
    
    
    'First of all, calculate the
    'dimensions of the List Boxes.
    List1.Height = UserControl.Height - 320
    List2.Height = UserControl.Height - 320
    
    List1.Width = (UserControl.Width / 2) - 180
    List2.Width = (UserControl.Width / 2) - 180
    
    'Now, calculate the left position
    'of the second List Box.
    List2.Left = List1.Width + 200
    Label2.Left = List2.Left
    
    'The List Boxes cannot have their height
    'changed randomly. Depending on the height
    'that it was resized to, it will
    'automatically resize itself to a default
    'size. Because of this, the user control
    'has to be resized to adjust to the List
    'Box new size.
    UserControl.Height = List1.Height + 380
    
    'Now that everything is in its final
    'position, resize the Frame.
    Frame1.Width = UserControl.Width
    Frame1.Height = UserControl.Height
    
    'When the User Control is at Runtime,
    'verify if any of the List Boxes need to
    'have horizontal Scroll Bars.
    Call AddHorizScroll(List1, UserControl.Parent)
    Call AddHorizScroll(List2, UserControl.Parent)
    
End Sub

Private Sub UserControl_Terminate()
    'Prevent this action from firing while
    'the control is in design mode.
    'm_SaveLists is a property of the
    'control that determines whether the
    'control will memorize the selections
    'made on each listbox.
    If (m_SaveLists) _
    And (myAmbient) Then
        Call SaveQueryDefs
        
    End If
    
End Sub

'I was having a few problems when
'checking the UserMode. Because of
'this, I decided to create this Timer.
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    If (UserControl.Ambient.UserMode) Then
        intFileName = UserControl.Ambient.DisplayName & UserControl.Parent.Name & ".txt"
        If (m_AutoConnect) Then
            Call DataConnect
            
        End If
        myAmbient = True
        
    Else
        myAmbient = False
        
    End If
    
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'The following Sub will be called
    'to check if the text of the
    'ListBox item underneath the mouse
    'pointer is lengthier than the
    'ListBox itself. If so, this Sub
    'will copy the text of the respective
    'item to the ToolTip of the ListBox.
    Call AddToolTip(List1, UserControl.Parent, Y)
    
End Sub

Private Sub List1_Scroll()
    'The following Sub will be called
    'to check if any of the items
    'displayed on the ListBox window is
    'lengthier than the ListBox itself.
    'If so, the Sub will add a horizontal
    'ScrollBar to the ListBox.
    Call AddHorizScroll(List1, UserControl.Parent)
    
End Sub

Private Sub List1_DblClick()
    'The following Sub will be called to
    'move the selected item to ListBox2.
    Call MoveToList2
    
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    'This will allow the user to use the
    'keyboard to move items back and forth.
    If (KeyAscii = vbKeyReturn) _
    Or (KeyAscii = vbKeySeparator) Then
        'The following Sub will be called to
        'move the selected items to ListBox2.
        Call MoveToList2
        
    End If
    
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'The following Sub will be called
    'to check if the text of the
    'ListBox item underneath the mouse
    'pointer is lengthier than the
    'ListBox itself. If so, this Sub
    'will copy the text of the respective
    'item to the ToolTip of the ListBox.
    Call AddToolTip(List2, UserControl.Parent, Y)
    
End Sub

Private Sub List2_Scroll()
    'The following Sub will be called to
    'check if any of the items displayed
    'on the ListBox window is lengthier
    'than the ListBox itself.
    'If so, the Sub will add a horizontal
    'ScrollBar to the ListBox.
    Call AddHorizScroll(List2, UserControl.Parent)
    
End Sub

Private Sub List2_DblClick()
    'The following Sub will be called to
    'move the selected item to ListBox1.
    Call MoveToList1
    
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
    'This will allow the user to use the
    'keyboard to move items back and forth.
    If (KeyAscii = vbKeyReturn) _
    Or (KeyAscii = vbKeySeparator) Then
        'The following Sub will be called to
        'move the selected items to ListBox1.
        Call MoveToList1
        
    End If
    
End Sub

'This sub will check if there is any
'text assigned to Label1. If there
'isn't any, hide label from user.
Private Sub Label1_Change()
    If (Trim(Label1.Caption) = "") Then
        Label1.Visible = False
        
    Else
        Label1.Visible = True
        
    End If
    
End Sub

'This sub will check if there is any
'text assigned to Label2. If there
'isn't any, hide label from user.
Private Sub Label2_Change()
    If (Trim(Label2.Caption) = "") Then
        Label2.Visible = False
        
    Else
        Label2.Visible = True
        
    End If
    
End Sub

'=================================================================================
'======= Complementary Subs & Functions ==========================================
'=================================================================================

'This little function is the heart
'of this Control. It will get an
'array with numbers on ascending
'sequence and output a string with
'all the numbers from this array
'in an abbreviate mode.
'
'If, for example, this function
'receives an array with the numbers
'(1,5,6,7,8,9,10,11,15,16), it will
'generate the string ("1,5-11,15,16").
'With a string formatted this way,
'it will be much easier to create
'a SQL Query.
Private Function SaveQueryDefs() As Boolean
    Dim intUBnd As Integer
    Dim intPrevNumber As Integer
    Dim intCurrNumber As Integer
    Dim bolSequence As Boolean
    Dim strFinal As String
    Dim intLast As Integer
    Dim i As Integer
    Dim intLastComma As Integer
    Dim intLastDash As Integer
    Dim intLastStart As Integer
    Dim intArray() As String
    
    Dim test As String
    
    'Get the size of the Array.
    intUBnd = UBound(Array2, 2)
    
    'If array is not empty...
    If (intUBnd > 0) Then
        
        ReDim intArray(2, intUBnd)
        
        For i = 1 To intUBnd
            intArray(1, i) = Array2(1, i)
            test = test & "," & Array2(1, i)
            
        Next i
        
        Debug.Print test
        
        'Sort array by ID number.
        Call subSortArray(intArray)
        
        
        'Start creating the final string.
        strFinal = intArray(1, 1)
        test = strFinal
        
        'Get this number for future comparissons.
        intPrevNumber = Val(strFinal)
        
        'Loop through the remaining items of
        'the array, if any.
        For i = 2 To intUBnd
            'Get the current ID number from the array.
            intCurrNumber = Val(intArray(1, i))
            test = test & "," & intCurrNumber
            
            'If the current ID number is on a sequence
            'with the previous number, continue looping.
            If (intCurrNumber = intPrevNumber + 1) Then
                'Take note that a sequence of numbers
                'has started.
                bolSequence = True
                
            'If the current number is not on a sequence
            'with the previous number, take a moment to
            'record the information gathered so far to
            'the Final String.
            Else
                'If there was a sequence that was not
                'saved yet, save it and, then, save the
                'current number.
                If (bolSequence) Then
                    
                    intLastComma = InStrRev(strFinal, ",")
                    intLastDash = InStrRev(strFinal, "-")
                    
                    If (intLastComma > intLastDash) Then
                        intLastStart = intLastComma
                        
                    ElseIf (intLastComma < intLastDash) Then
                        intLastStart = intLastDash
                        
                    Else
                        intLastStart = 0
                        
                    End If
                    
                    intLast = Val(Right(strFinal, (Len(strFinal) - intLastStart)))
                    
                    'Check if the last character was in a
                    'sequence with the last character in
                    'that sequence. If so, put a comma
                    'between them.
                    If (intLast = (intPrevNumber - 1)) Then
                        strFinal = strFinal & "," & intPrevNumber & "," & intCurrNumber
                        
                    'If the numbers were not in a sequence,
                    'put a dash between them.
                    Else
                        strFinal = strFinal & "-" & intPrevNumber & "," & intCurrNumber
                        
                    End If
                    
                Else
                    strFinal = strFinal & "," & intCurrNumber
                    
                End If
                bolSequence = False
                
            End If
            
            If (i < intUBnd) Then
                intPrevNumber = intCurrNumber
            End If
        Next i
        
        intLastComma = InStrRev(strFinal, ",")
        intLastDash = InStrRev(strFinal, "-")
        
        If (intLastComma > intLastDash) Then
            intLastStart = intLastComma
            
        ElseIf (intLastComma < intLastDash) Then
            intLastStart = intLastDash
            
        Else
            intLastStart = 0
            
        End If
        
        'Get last number on the string...
        intLast = Val(Right(strFinal, (Len(strFinal) - intLastStart)))
        
        'If the end of the strFinal string is the
        'number right before the current number
        'followed by a dash, remove the dash and
        'add a comma. After that, add the current
        'number to the end of the string.
        If (Right(strFinal, Len(Trim(Str(intLast))) + 1) = "-" & (intCurrNumber - 1)) Then
            strFinal = Left(strFinal, (Len(strFinal) - Len(Str(intLast)))) & intCurrNumber
            
        'If the end of the strFinal string is the
        'number right before the current number
        'followed by a comma, just add the current
        'number to the end of the string.
        ElseIf (Right(strFinal, Len(Trim(Str(intLast))) + 1) = "," & (intCurrNumber - 1)) Then
            strFinal = strFinal & "," & intCurrNumber
            
        'If the last number was not in a
        'sequence with the current number...
        ElseIf (intLast < (intCurrNumber - 1)) Then
            'If a sequence was been carried
            'out, add a dash and, then, the
            'current number to the end of
            'the string.
            If (bolSequence) Then
                strFinal = strFinal & "-" & intCurrNumber
                
            'If there was no sequence been
            'carried out, add a comma and,
            'then, the current number to
            'the end of the string.
            Else
                strFinal = strFinal & "," & intCurrNumber
                
            End If
            
        'If there was only one number on
        'string, add a comma and,
        'then, the current number to
        'the end of the string.
        ElseIf (Len(Trim(strFinal)) = Len(Trim(intLast))) Then
            If (intCurrNumber > 0) Then
                strFinal = strFinal & "," & intCurrNumber
            End If
        End If
        
    End If
    
    Debug.Print test
    
    strFinalSelect = strFinal
    
    SaveQueryDefs = True
    
    'MsgBox strFinal & "*" & test
    Call Create_File(strFinal, intFileName)
    
End Function

'This sub will sort the database on numeric
'order using the ID field as reference.
Private Sub subSortArray(m_Array As Variant)
    Dim intFirstUBnd As Integer
    Dim intLastUBnd As Integer
    Dim intCurrItem As Integer
    Dim strCurrItem As String
    Dim intBiggest As Integer
    Dim intArray() As String
    Dim i As Integer
    Dim j As Integer
    
    intFirstUBnd = UBound(m_Array, 2)
    
    If (intFirstUBnd > 0) Then
        
        'Find out which item is the
        'biggest on Array2.
        For i = 1 To intFirstUBnd
            
            intCurrItem = Val(m_Array(1, i))
            
            If (intCurrItem > intBiggest) Then
                intBiggest = intCurrItem
            End If
            
        Next i
        
        'If there is an item ID that is
        'bigger than the original array size
        If (intBiggest > intFirstUBnd) Then
            ReDim Preserve m_Array(2, intBiggest)
            
        End If
        
        intLastUBnd = UBound(m_Array, 2)
        
        ReDim intArray(2, intLastUBnd)
        
        'Copy all itmes form Array2 to
        'intArray on proper order.
        For i = 1 To intFirstUBnd
            
            intCurrItem = Val(m_Array(1, i))
            
            intArray(1, intCurrItem) = m_Array(1, i)
            intArray(2, intCurrItem) = m_Array(2, i)
            
        Next i
        
        'Check if original array has been resized
        'before. If so, restore its original size.
        If (intBiggest > intFirstUBnd) Then
            ReDim m_Array(2, intFirstUBnd)
            
        End If
        
        'Copy all itmes back to Array2.
        'Just skip any possible gap on intArray.
        For i = 1 To intLastUBnd
            
            strCurrItem = intArray(1, i)
            
            'If is not the item that
            'has to be removed...
            If (strCurrItem <> "") Then
                j = j + 1
                m_Array(1, j) = intArray(1, i)
                m_Array(2, j) = intArray(2, i)
                
            End If
            
        Next i
        
    End If
    
End Sub

'The following Sub will move the
'provided item from Array2 to
'Array1.
Private Sub subMoveToArray1(m_item As Integer)
    Dim intArray() As String
    Dim j As Integer
    Dim List1Count As Integer
    Dim List2Count As Integer
    Dim i As Integer
    
    '-------------------------------------
    '---- Add item to the array ----------
    '-------------------------------------
    List1Count = List1.ListCount
    
    'If there is one or more items on
    'Array1, redimension it but preserve
    'the items that are already there.
    If (UBound(Array1, 2) > 0) Then
        ReDim Preserve Array1(2, List1Count)
        
    'If there is no item on Array1,
    'just redimesion it.
    Else
        ReDim Array1(2, List1Count)
    End If
    
    'Copy the requested item from Array2
    'to the last cell that has just been
    'added to Array1.
    Array1(1, List1Count) = Array2(1, m_item)
    Array1(2, List1Count) = Array2(2, m_item)
    
    '-------------------------------------
    '---- Delete item from previows array
    '-------------------------------------
    List2Count = List2.ListCount
    
    'Create a temporary array that will
    'be used to keep all the items from
    'Array2 while Array2 is been reduced
    'in size.
    ReDim intArray(2, List2Count)
    
    'Copy all the items from Array2 to
    'intArray except the item that is
    'been removed.
    For i = 1 To UBound(Array2, 2)
        
        'If is not the item that
        'has to be removed...
        If (i <> m_item) Then
            j = j + 1
            intArray(1, j) = Array2(1, i)
            intArray(2, j) = Array2(2, i)
            
        End If
    Next i
    
    'Reduce size of Array2...
    ReDim Array2(2, List2Count)
    
    'Put the items back to Array2,
    'now, reduced in size...
    For i = 1 To List2Count
        Array2(1, i) = intArray(1, i)
        Array2(2, i) = intArray(2, i)
        
    Next i
    
End Sub

'The following Sub will move the
'provided item from Array1 to
'Array2.
Private Sub subMoveToArray2(m_item As Integer)
    Dim intArray() As String
    Dim j As Integer
    Dim List1Count As Integer
    Dim List2Count As Integer
    Dim i As Integer
    
    '-------------------------------------
    '---- Add item to new array
    '-------------------------------------
    List2Count = List2.ListCount
    
    'If there is one or more items on
    'Array2, redimension it but preserve
    'the items that are already there.
    If (UBound(Array2, 2) > 0) Then
        ReDim Preserve Array2(2, List2Count)
        
    'If there is no item on Array2,
    'just redimesion it.
    Else
        ReDim Array2(2, List2Count)
    End If
    
    'Copy the requested item from Array1
    'to the last cell that has just been
    'added to Array2.
    Array2(1, List2Count) = Array1(1, m_item)
    Array2(2, List2Count) = Array1(2, m_item)
    
    '-------------------------------------
    '---- Delete item from previows array
    '-------------------------------------
    List1Count = List1.ListCount
    
    'Create a temporary array that will
    'be used to keep all the items from
    'Array1 while Array1 is been reduced
    'in size.
    ReDim intArray(2, List1Count)
    
    'Copy all the items from Array1 to
    'intArray except the item that is
    'been removed.
    For i = 1 To UBound(Array1, 2)
        
        'If is not the item that
        'has to be removed...
        If (i <> m_item) Then
            j = j + 1
            intArray(1, j) = Array1(1, i)
            intArray(2, j) = Array1(2, i)
            
        End If
    Next i
    
    'Reduce size of Array1...
    ReDim Array1(2, List1Count)
    
    'Put the items back to Array1,
    'now, reduced in size...
    For i = 1 To List1Count
        Array1(1, i) = intArray(1, i)
        Array1(2, i) = intArray(2, i)
        
    Next i
    
End Sub

'=================================================================================
'======= Public Subs =============================================================
'=================================================================================

'This Sub is called when the control is
'initializing to read the Database and
'populate both List Boxes with the items.
'This Sub will also populate the two
'arrays that will control the items on
'each ListBox.
Public Sub DataConnect()
    Dim intDataPath As String
    Dim intSQL
    Dim intRecCount As Integer
    Dim intFieldValue As String
    Dim myList1QD As Boolean
    Dim intQueryCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim strFileContent As String
    Dim bolFound As Boolean
    Dim intFieldID As Integer
    Dim intUBound As Integer
    Dim bolEmptyArray As Boolean
    Dim strArray As Variant
    Dim tmpArray1() As Variant
    Dim tmpArray2() As Variant
    Dim intList1Count As Integer
    Dim intList2Count As Integer
    Dim intQD As QueryDef
    Dim strNewSQL As String
    
    'Don't do anything if no Database
    'is provided by the user.
    If (m_DatabaseName <> "") _
    And (m_RecordSource <> "") _
    And (m_IDFieldName <> "") _
    And (m_FieldName <> "") Then
        
        'Check if the Database file really exists...
        If (fntFileExist(m_DatabaseName)) Then
            'The user can provid the full path or
            'a relative path for the Database.
            'Nevertheless, the control will be
            'able to find the Database file.
            intDataPath = fntFullOrRelative(m_DatabaseName)
            
            On Error GoTo DataBaseError
            Set intDB = Workspaces(0).OpenDatabase(intDataPath, True, False, ";pwd=" & m_Password)
            
            'If there is anything set as SQLString,
            'create a temporary Query Definition.
            If (m_SQLString <> "") Then
                'Create a temporary Query Definition...
                Set intQD = intDB.CreateQueryDef("")
                
                'Add the first part of the SQL string.
                'This part sets the Table name.
                strNewSQL = "SELECT * FROM " & m_RecordSource & " " & m_SQLString
                
                'An example of SQL statement would be:
                '"WHERE (ID >= 5) ORDER BY Principal1Name;"
                
                Debug.Print strNewSQL
                
                'Add the SQL statement to the new
                'Querry Definition.
                intQD.SQL = strNewSQL
                
                'Create a new Recordset based on the
                'Query Definition.
                Set intRS = intQD.OpenRecordset()
                
            Else
                Set intRS = intDB.OpenRecordset(m_RecordSource)
                
            End If
            
            'If the user chose to save the
            'items selected on each ListBox...
            If (m_SaveLists) Then
                'Get the text that is on the file.
                'This text has a list of the IDs
                'of all the items that were selected
                'when the 2ListBxs Control was
                'last opened.
                strFileContent = fntGetTextFile(intFileName)
                
            End If
            
            '===================================================
            '====== Populate both ListBoxes and Arrays =========
            '===================================================
            
            If (strFileContent = "") Then
                ReDim strArray(0)
                
                'When you split a string and it has only
                '1 item, the array generated will have
                'this item on Index 0. However, when
                'there is no item on the string, I ReDim
                'the array with 0 items. Therefore, it
                'could be a little confusing to determine
                'whether the string had an item or not.
                bolEmptyArray = True
                
            Else
                'Part of the SQL that connects to the Table...
                strFileContent = fntSpreadStr(strFileContent)
                strArray = Split(strFileContent, ",")
                
            End If
            
            intRecCount = intRS.RecordCount
            
            'Copy all items from the Recordset
            'to List1...
            If (intRecCount > 0) Then
                
                'Determine how many items there were on
                'the string.
                If (bolEmptyArray) Then
                    intUBound = 0
                    
                Else
                    intUBound = UBound(strArray) + 1
                    
                End If
                
                'Prepare the two temporary arrays to
                'collect the information.
                ReDim tmpArray1(2, intRecCount)
                ReDim tmpArray2(2, intUBound)
                
                intRS.MoveFirst
                
                'Clear the display on both ListBoxes.
                List1.Clear
                List2.Clear
                
                'If there was one or more items on
                'the string, add these items to the
                'Right ListBox.
                If (intUBound > 0) Then
                    'Loop though all the items on database
                    'and verify if any of them is on the
                    'list of items previously saved on the
                    'string.
                    For i = 1 To intRecCount
                        'Get the required information from database.
                        intFieldValue = intRS.Fields(m_FieldName).Value & ""
                        intFieldID = Val(intRS.Fields(m_IDFieldName).Value & "")
                        
                        'Now, loop through all the items found on the string.
                        For j = 0 To intUBound - 1
                            'Verify if any of the items on the string
                            'is equal to the current item of the database.
                            If (Val(strArray(j)) = intFieldID) Then
                                List2.AddItem intFieldValue
                                
                                intList2Count = List2.ListCount
                                
                                tmpArray2(1, intList2Count) = intRS.Fields(m_IDFieldName).Value & ""
                                tmpArray2(2, intList2Count) = intFieldValue
                                
                                bolFound = True
                                Exit For
                            End If
                        Next j
                        
                        If (bolFound) Then
                            bolFound = False
                            
                        Else
                            'n = n + 1
                            List1.AddItem intFieldValue
                            
                            intList1Count = List1.ListCount
                            
                            tmpArray1(1, intList1Count) = intRS.Fields(m_IDFieldName).Value & ""
                            tmpArray1(2, intList1Count) = intFieldValue
                            
                        End If
                        
                        'Keep track of every item added to the
                        'ListBox1 by using an array that will
                        'have an image of the items on the
                        'ListBox at all times.
                        
                        intRS.MoveNext
                        
                    Next i
                    
                    'Prepare the two permanent arrays to
                    'collect the information.
                    ReDim Array1(2, intList1Count)
                    ReDim Array2(2, intList2Count)
                    
                    'Copy the information from the temporary Arrays
                    'to the permanent arrays. These temporary arrays
                    'had to be created to circumvent a problem that
                    'would occur if I used the control with a database
                    'and, then, deleted items from the database. When
                    'reopening the control, it would get in to an
                    'error if one of the deleted items was one of the
                    'selected items that were saved on the string.
                    For i = 1 To intList1Count
                        Array1(1, i) = tmpArray1(1, i)
                        Array1(2, i) = tmpArray1(2, i)
                    Next i
                    
                    For i = 1 To intList2Count
                        Array2(1, i) = tmpArray2(1, i)
                        Array2(2, i) = tmpArray2(2, i)
                    Next i
                    
                Else
                    
                    'Prepare the two permanent arrays to
                    'collect the information.
                    ReDim Array1(2, intRecCount)
                    ReDim Array2(2, 0)
                    
                    For i = 1 To intRecCount
                        
                        intFieldValue = intRS.Fields(m_FieldName).Value & ""
                        
                        List1.AddItem intFieldValue
                        
                        Array1(1, i) = intRS.Fields(m_IDFieldName).Value & ""
                        Array1(2, i) = intFieldValue
                        
                        intRS.MoveNext
                        
                    Next i
                End If
            End If
            
            intRS.Close
            intDB.Close
            
            'Free up all resources...
            Set intRS = Nothing
            Set intDB = Nothing
            
        End If
    End If
    
    
    'Select the first item of the
    'first ListBox...
    If (List1.ListCount > 0) Then
        List1.Selected(0) = True
        
    End If
    
    'Select the first item of the
    'second ListBox...
    If (List2.ListCount > 0) Then
        List2.Selected(0) = True
        
    End If
    
    'The following called Subs will
    'check if any of the items
    'displayed on each ListBox window is
    'lengthier than the ListBox itself.
    'If so, it will add a horizontal
    'ScrollBar to the respective ListBox.
    Call AddHorizScroll(List1, UserControl.Parent)
    Call AddHorizScroll(List2, UserControl.Parent)
    
    Exit Sub
    
'Custom error handling...
DataBaseError:
    
    'The following error can occur if the
    'Database that is been accessed is
    'already open.
    If (Err.Number = 3356) Then
        List1.AddItem "Database is already open by other application."
        Call AddHorizScroll(List1, UserControl.Parent)
        
    Else
        MsgBox "There was an unespected error." & Chr(10) & _
               "Error number: " & Err.Number & Chr(10) & _
               "Error description:" & Chr(10) & _
               Err.Description
    End If
    
End Sub

'------------------------------------------------------
'---------- Subs related to ListBox1 ------------------
'------------------------------------------------------

'The following Public Sub will move the
'selected items from Array1 to
'Array2.
Public Sub MoveToList2()
    Dim i As Integer
    Dim List1Count As Integer
    Dim intItemToSelect As Integer
    
    List1Count = List1.ListCount
    
    'This For Next loop will go through each
    'item of the ListBox1 and move each item
    'that is selected to ListBox2.
    For i = 1 To List1Count
        If (i > List1Count) Then
            Exit For
        End If
        
        If (List1.Selected(i - 1)) Then
            List2.AddItem (List1.List(i - 1))
            List1.RemoveItem (i - 1)
            
            'This information will be used when
            'the moving is done.
            If (intItemToSelect < i - 1) Then
                intItemToSelect = i - 1
                
            End If
            
            'The called Sub will move the
            'current item from Array1 to
            'Array2.
            Call subMoveToArray2(i)
            
            'This conditional will reduce
            'the number of loops required
            'as long as it is not the last
            'item.
            If (i < List1Count) Then
                i = i - 1
                
            Else
                Exit For
                
            End If
            
            List1Count = List1Count - 1
            'Exit For
        End If
    Next i
    
    'Clear any selected item on Right List Box.
    For i = 0 To List2.ListCount - 1
        If (List2.Selected(i)) Then
            List2.Selected(i) = False
        End If
        
    Next i
    
    'Select the last item on Right List Box.
    List2.Selected(List2.ListCount - 1) = True
    
    'Select the item just after the
    'items that have just moved.
    If (intItemToSelect < List1.ListCount) Then
        List1.Selected(intItemToSelect) = True
        
    Else
        If (List1.ListCount > 0) Then
            List1.Selected(List1.ListCount - 1) = True
        End If
    End If
    
    bolHasScrollBar1 = False
    bolHasScrollBar2 = False
    
    'The following called Subs will
    'check if any of the items
    'displayed on each ListBox window is
    'lengthier than the ListBox itself.
    'If so, it will add a horizontal
    'ScrollBar to the respective ListBox.
    Call AddHorizScroll(List1, UserControl.Parent)
    Call AddHorizScroll(List2, UserControl.Parent)
    
End Sub

Public Sub SortList1()
    Dim intUBnd As Integer
    Dim i As Integer
    
    Call Sort_2D_Bubble(Array1, 2, 2)
    
    intUBnd = UBound(Array1, 2)
    
    If (intUBnd > 0) Then
        
        List1.Clear
        
        For i = 1 To intUBnd
            Call List1.AddItem(Array1(2, i))
            
        Next i
        
    End If
    
End Sub

'The following Public Sub will select
'all items displayed on ListBox1.
Public Sub SelectAllList1()
    Dim i As Integer
    
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = True
        
    Next i
    
End Sub

'The following Public Sub will clear
'any selections from ListBox1.
Public Sub ClearList1()
    Dim i As Integer
    
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = False
        
    Next i
    
End Sub

'------------------------------------------------------
'---------- Subs related to ListBox2 ------------------
'------------------------------------------------------

'The following Public Sub will move the
'selected items from Array2 to
'Array1.
Public Sub MoveToList1()
    Dim i As Integer
    Dim List2Count As Integer
    Dim intItemToSelect As Integer
    
    List2Count = List2.ListCount
    
    'This For Next loop will go through each
    'item of the ListBox2 and move each item
    'that is selected to ListBox1.
    For i = 1 To List2Count
        If (i > List2Count) Then
            Exit For
        End If
        
        If (List2.Selected(i - 1)) Then
            List1.AddItem (List2.List(i - 1))
            List2.RemoveItem (i - 1)
            
            'This information will be used when
            'the moving is done.
            If (intItemToSelect < i - 1) Then
                intItemToSelect = i - 1
                
            End If
            
            'The called Sub will move the
            'current item from Array2 to
            'Array1.
            Call subMoveToArray1(i)
            
            'This conditional will reduce
            'the number of loops require
            'as long as it is not the last
            'item.
            If (i < List2Count) Then
                i = i - 1
                
            Else
                Exit For
                
            End If
            
            List2Count = List2Count - 1
            
        End If
    Next i
    
    'Clear any selected item on Left List Box.
    For i = 0 To List1.ListCount - 1
        If (List1.Selected(i)) Then
            List1.Selected(i) = False
        End If
        
    Next i
    
    'Select the last item on Left List Box.
    List1.Selected(List1.ListCount - 1) = True
    
    'Select the item just after the
    'items that have just moved.
    If (intItemToSelect < List2.ListCount) Then
        List2.Selected(intItemToSelect) = True
        
    Else
        If (List2.ListCount > 0) Then
            List2.Selected(List2.ListCount - 1) = True
        End If
    End If
        
    bolHasScrollBar1 = False
    bolHasScrollBar2 = False
    
    'The following called Subs will
    'check if any of the items
    'displayed on each ListBox window is
    'lengthier than the ListBox itself.
    'If so, it will add a horizontal
    'ScrollBar to the respective ListBox.
    Call AddHorizScroll(List1, UserControl.Parent)
    Call AddHorizScroll(List2, UserControl.Parent)
    
End Sub

Public Sub SortList2()
    Dim intUBnd As Integer
    Dim i As Integer
    
    Call Sort_2D_Bubble(Array2, 2, 2)
    
    intUBnd = UBound(Array2, 2)
    
    If (intUBnd > 0) Then
        
        List2.Clear
        
        For i = 1 To intUBnd
            Call List2.AddItem(Array2(2, i))
            
        Next i
        
    End If
    
End Sub

'The following Public Sub will select
'all items displayed on ListBox2.
Public Sub SelectAllList2()
    Dim i As Integer
    
    For i = 0 To List2.ListCount - 1
        List2.Selected(i) = True
        
    Next i
    
End Sub

'The following Public Sub will clear
'any selections from ListBox2.
Public Sub ClearList2()
    Dim i As Integer
    
    For i = 0 To List2.ListCount - 1
        List2.Selected(i) = False
        
    Next i
    
End Sub

'--------------------------------------------
'--------- Getting the results --------------
'--------------------------------------------

'This Public Sub can be activated by the
'user to create a Recordset based on the
'list of all the records displayed on
'ListBox2.
Public Sub RSFinalConnect()
    Dim intDataPath As String
    Dim intSQL As String
    Dim strSortBy As String
    
    strSortBy = m_SortBy
    
    'Verify if user selected a
    'field to sort by...
    If (strSortBy = "") Then
        strSortBy = m_FieldName
        
    End If
    
    'Check if the minimum required
    'information is available.
    If (m_DatabaseName <> "") _
    And (m_RecordSource <> "") _
    And (m_IDFieldName <> "") _
    And (m_FieldName <> "") Then
        
        'Check if the Database file really exists...
        If (fntFileExist(m_DatabaseName)) Then
            
            'The user can provid the full path or
            'a relative path for the Database.
            'Nevertheless, the control will be
            'able to find the Database file.
            If (Mid(m_DatabaseName, 2, 2) = ":\") Then
                intDataPath = m_DatabaseName
                
            Else
                intDataPath = App.Path & "\" & m_DatabaseName
                
            End If
            
            On Error GoTo DataBaseError
            
            'Connect to Database...
            Set intDB = Workspaces(0).OpenDatabase(intDataPath, True, False, ";pwd=" & m_Password)
            
            'Create a temporary Query Definition...
            Set intQD = intDB.CreateQueryDef("")
            
            If (SaveQueryDefs) Then
                'Add the first part of the SQL string.
                'This part sets the Table name.
                intFinalSQL = "SELECT * FROM " & m_RecordSource & " WHERE ("
                
                If (strFinalSelect <> "") Then
                    
                    intFinalSQL = intFinalSQL & fntGenerSQL(strFinalSelect) & ") ORDER BY " & strSortBy & ";"
                    
                Else
                    intFinalSQL = intFinalSQL & "ID=0) ORDER BY " & strSortBy & ";"
                End If
            End If
            
            Debug.Print intFinalSQL
            
            'Add the SQL statement to the new
            'Querry Definition.
            intQD.SQL = intFinalSQL
            
            'Create a new Recordset based on the
            'Query Definition.
            Set RSFinal = intQD.OpenRecordset()
        End If
    End If
    
    Exit Sub
    
'Custom error handling...
DataBaseError:
    
    'The following error can occur if the
    'Database that is been accessed is
    'already open.
    If (Err.Number = 3356) Then
        List1.AddItem "Database is already open by other application."
        Call AddHorizScroll(List1, UserControl.Parent)
    Else
        MsgBox "There was an unespected error." & Chr(10) & _
               "Error number: " & Err.Number & Chr(10) & _
               "Error description:" & Chr(10) & _
               Err.Description
    End If
    
End Sub

'================================================
'======== Initialize and Save Properties ========
'================================================

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_DatabaseName = m_def_DatabaseName
    m_RecordSource = m_def_RecordSource
    m_Caption1 = m_def_Caption1
    m_Caption2 = m_def_Caption2
    m_FieldName = m_def_FieldName
    m_IDFieldName = m_def_IDFieldName
    m_Password = m_def_Password
    m_SortBy = m_def_SortBy
    m_AutoConnect = m_def_AutoConnect
    m_SaveLists = m_def_SaveLists
    m_SQLString = m_def_SQLString
    
End Sub

'Load property values from storage.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_DatabaseName = PropBag.ReadProperty("DatabaseName", m_def_DatabaseName)
    m_RecordSource = PropBag.ReadProperty("RecordSource", m_def_RecordSource)
    m_Caption1 = PropBag.ReadProperty("Caption1", m_def_Caption1)
    m_Caption2 = PropBag.ReadProperty("Caption2", m_def_Caption2)
    m_FieldName = PropBag.ReadProperty("FieldName", m_def_FieldName)
    m_IDFieldName = PropBag.ReadProperty("IDFieldName", m_def_IDFieldName)
    m_Password = PropBag.ReadProperty("Password", m_def_Password)
    m_SortBy = PropBag.ReadProperty("SortBy", m_def_SortBy)
    m_AutoConnect = PropBag.ReadProperty("AutoConnect", m_def_AutoConnect)
    m_SaveLists = PropBag.ReadProperty("SaveLists", m_def_SaveLists)
    
    Label1.Caption = m_Caption1
    Label2.Caption = m_Caption2
    
    List1.BackColor = PropBag.ReadProperty("L1BackColor", &H80000005)
    List2.BackColor = PropBag.ReadProperty("L2BackColor", &H80000005)
    List1.ForeColor = PropBag.ReadProperty("L1ForeColor", &H80000012)
    List2.ForeColor = PropBag.ReadProperty("L2ForeColor", &H80000012)
    
    Frame1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    
    If (Frame1.BorderStyle = 0) Then
        Label1.Visible = False
        Label2.Visible = False
        
    Else
        Label1.Visible = True
        Label2.Visible = True
        
    End If
    
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Label2.ForeColor = Label1.ForeColor
    
    Frame1.Appearance = PropBag.ReadProperty("Appearance", 1)
    
    Frame1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    
    Label1.BackColor = Frame1.BackColor
    Label2.BackColor = Frame1.BackColor
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Label1.FontBold = PropBag.ReadProperty("CaptionBold", 0)
    Label2.FontBold = Label1.FontBold
    
    Set List1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set List2.Font = List1.Font
    
    m_SQLString = PropBag.ReadProperty("SQLString", m_def_SQLString)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DatabaseName", m_DatabaseName, m_def_DatabaseName)
    Call PropBag.WriteProperty("RecordSource", m_RecordSource, m_def_RecordSource)
    Call PropBag.WriteProperty("Caption1", m_Caption1, m_def_Caption1)
    Call PropBag.WriteProperty("Caption2", m_Caption2, m_def_Caption2)
    Call PropBag.WriteProperty("FieldName", m_FieldName, m_def_FieldName)
    Call PropBag.WriteProperty("IDFieldName", m_IDFieldName, m_def_IDFieldName)
    Call PropBag.WriteProperty("Password", m_Password, m_def_Password)
    Call PropBag.WriteProperty("SortBy", m_SortBy, m_def_SortBy)
    Call PropBag.WriteProperty("AutoConnect", m_AutoConnect, m_def_AutoConnect)
    Call PropBag.WriteProperty("SaveLists", m_SaveLists, m_def_SaveLists)
    Call PropBag.WriteProperty("L1BackColor", List1.BackColor, &H80000005)
    Call PropBag.WriteProperty("L2BackColor", List2.BackColor, &H80000005)
    Call PropBag.WriteProperty("L1ForeColor", List1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("L2ForeColor", List2.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", Frame1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BorderStyle", Frame1.BorderStyle, 1)
    Call PropBag.WriteProperty("Appearance", Frame1.Appearance, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("CaptionBold", Label1.FontBold, 0)
    Call PropBag.WriteProperty("Font", List1.Font, Ambient.Font)
    Call PropBag.WriteProperty("SQLString", m_SQLString, m_def_SQLString)
    'm_SQLString
End Sub

'================================================
'============= Properties List ==================
'================================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get DataBaseName() As String
Attribute DataBaseName.VB_Description = "It is the path of the database to which the control will connect. The developer can provide the full path or the relative path for the database."
Attribute DataBaseName.VB_ProcData.VB_Invoke_Property = ";Data"
    DataBaseName = m_DatabaseName
    
End Property

Public Property Let DataBaseName(ByVal New_DatabaseName As String)
    m_DatabaseName = New_DatabaseName
    PropertyChanged "DatabaseName"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get RecordSource() As String
Attribute RecordSource.VB_Description = "RecordeSource is the name of the Table within the Database to which the control will connect."
Attribute RecordSource.VB_ProcData.VB_Invoke_Property = ";Data"
    RecordSource = m_RecordSource
    
End Property

Public Property Let RecordSource(ByVal New_RecordSource As String)
    m_RecordSource = New_RecordSource
    PropertyChanged "RecordSource"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Title1
Public Property Get Caption1() As String
Attribute Caption1.VB_Description = "The Caption1 property will determine the Title that will appear right on top of the Left ListBox."
Attribute Caption1.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption1.VB_UserMemId = -518
    Caption1 = m_Caption1
    
End Property

Public Property Let Caption1(ByVal New_Caption1 As String)
    m_Caption1 = New_Caption1
    Label1.Caption = New_Caption1
    PropertyChanged "Caption1"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Title2
Public Property Get Caption2() As String
Attribute Caption2.VB_Description = "The Caption2 property will determine the Title that will appear right on top of the Right ListBox."
Attribute Caption2.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption2.VB_UserMemId = -517
    Caption2 = m_Caption2
    
End Property

Public Property Let Caption2(ByVal New_Caption2 As String)
    m_Caption2 = New_Caption2
    Label2.Caption = New_Caption2
    PropertyChanged "Caption2"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FieldName() As String
Attribute FieldName.VB_Description = "FieldName is the name of the Field within the Table that will be listed on the ListBoxes."
Attribute FieldName.VB_ProcData.VB_Invoke_Property = ";Data"
    FieldName = m_FieldName
    
End Property

Public Property Let FieldName(ByVal New_FieldName As String)
    m_FieldName = New_FieldName
    PropertyChanged "FieldName"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get IDFieldName() As String
Attribute IDFieldName.VB_Description = "IDFieldName is the name of the Primary Key Field on the Table. This field must be numeric."
Attribute IDFieldName.VB_ProcData.VB_Invoke_Property = ";Data"
    IDFieldName = m_IDFieldName
    
End Property

Public Property Let IDFieldName(ByVal New_IDFieldName As String)
    m_IDFieldName = New_IDFieldName
    PropertyChanged "IDFieldName"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Password() As String
Attribute Password.VB_Description = "The Password property is only required if your database is password protected."
Attribute Password.VB_ProcData.VB_Invoke_Property = ";Data"
    Password = m_Password
    
End Property

Public Property Let Password(ByVal New_Password As String)
    m_Password = New_Password
    PropertyChanged "Password"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SortBy() As String
Attribute SortBy.VB_Description = "SortBy is the name of the Field within the Table to sort by in alphabetic order."
Attribute SortBy.VB_ProcData.VB_Invoke_Property = ";Data"
    SortBy = m_SortBy
    
End Property

Public Property Let SortBy(ByVal New_SortBy As String)
    m_SortBy = New_SortBy
    PropertyChanged "SortBy"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,0,True
Public Property Get AutoConnect() As Boolean
Attribute AutoConnect.VB_Description = "If the AutoConnect property is set to True, the control will connect to the specified database as soon as it is loaded."
Attribute AutoConnect.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoConnect = m_AutoConnect
    
End Property

Public Property Let AutoConnect(ByVal New_AutoConnect As Boolean)
    If Ambient.UserMode Then Err.Raise 382
    m_AutoConnect = New_AutoConnect
    PropertyChanged "AutoConnect"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get SaveLists() As Boolean
Attribute SaveLists.VB_Description = "If the SaveLists property is set to True, the control will memorize the items that were moved from the Left ListBox to the Right ListBox."
Attribute SaveLists.VB_ProcData.VB_Invoke_Property = ";Behavior"
    SaveLists = m_SaveLists
    
End Property

Public Property Let SaveLists(ByVal New_SaveLists As Boolean)
    m_SaveLists = New_SaveLists
    PropertyChanged "SaveLists"
    
End Property

'-----------------------------------------------------------------

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,BackColor
Public Property Get L1BackColor() As OLE_COLOR
Attribute L1BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute L1BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    L1BackColor = List1.BackColor
End Property

Public Property Let L1BackColor(ByVal New_L1BackColor As OLE_COLOR)
    List1.BackColor() = New_L1BackColor
    PropertyChanged "L1BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get L1ForeColor() As OLE_COLOR
Attribute L1ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute L1ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    L1ForeColor = List1.ForeColor
    
End Property

Public Property Let L1ForeColor(ByVal New_L1ForeColor As OLE_COLOR)
    List1.ForeColor() = New_L1ForeColor
    PropertyChanged "L1ForeColor"
    
End Property

'-----------------------------------------------------------------

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,BackColor
Public Property Get L2BackColor() As OLE_COLOR
Attribute L2BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute L2BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    L2BackColor = List2.BackColor
    
End Property

Public Property Let L2BackColor(ByVal New_L2BackColor As OLE_COLOR)
    List2.BackColor() = New_L2BackColor
    PropertyChanged "L2BackColor"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,ForeColor
Public Property Get L2ForeColor() As OLE_COLOR
Attribute L2ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute L2ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    L2ForeColor = List2.ForeColor
    
End Property

Public Property Let L2ForeColor(ByVal New_L2ForeColor As OLE_COLOR)
    List2.ForeColor() = New_L2ForeColor
    PropertyChanged "L2ForeColor"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Frame1,Frame1,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = Frame1.BorderStyle
    
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyle)
    Frame1.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    
    If (New_BorderStyle = 0) Then
        Label1.Visible = False
        Label2.Visible = False
        
    Else
        Label1.Visible = True
        Label2.Visible = True
        
    End If
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Frame1,Frame1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = Frame1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Frame1.BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    Label1.BackColor = Frame1.BackColor
    Label2.BackColor = Frame1.BackColor

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Label2.ForeColor = Label1.ForeColor
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Frame1,Frame1,-1,Appearance
Public Property Get Appearance() As Appearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = Frame1.Appearance
    
End Property

Public Property Let Appearance(ByVal New_Appearance As Appearance)
    Dim tmpBackColor As OLE_COLOR
    
    tmpBackColor = Frame1.BackColor
    
    Frame1.Appearance() = New_Appearance
    
    Frame1.BackColor = tmpBackColor
    
    PropertyChanged "Appearance"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
    
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontBold
Public Property Get CaptionBold() As Boolean
Attribute CaptionBold.VB_Description = "If the CaptionBold property is set to True, the Titles above the two ListBoxes will have their font set to Bold."
Attribute CaptionBold.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CaptionBold = Label1.FontBold
End Property

Public Property Let CaptionBold(ByVal New_CaptionBold As Boolean)
    Label1.FontBold = New_CaptionBold
    Label2.FontBold = Label1.FontBold
    PropertyChanged "CaptionBold"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = List1.Font
    
End Property

Public Property Set Font(ByVal New_Font As Font)
    'I decided to use the same Font
    'for both List Boxes for a reason.
    'When you changed the size of the
    'Font on a List Box, the List Box
    'will automatically change their
    'height property to accommodate to
    'this new Font size. As a result,
    'the two List Boxes would end up
    'been displayed with different
    'heights. This is an unacceptable
    'behavior. At list it is for me! :)
    Set List1.Font = New_Font
    Set List2.Font = New_Font
    
    Call UserControl_Resize
    
    PropertyChanged "Font"
    
End Property

Public Property Get SQLString() As Variant
Attribute SQLString.VB_Description = "If anything is specified, the control will create a temporary Query Definition and will connect to it."
Attribute SQLString.VB_ProcData.VB_Invoke_Property = ";Data"
    SQLString = m_SQLString
    
End Property

Public Property Let SQLString(ByVal vNewValue As Variant)
    m_SQLString = vNewValue
    
    PropertyChanged "SQLString"
    
End Property
