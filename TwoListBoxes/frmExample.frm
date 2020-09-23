VERSION 5.00
Object = "*\AvbpTwoListBox.vbp"
Begin VB.Form frmExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   6000
      Top             =   6480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go to my submission and vote"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   6135
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Previous"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin vbpEB82LstBxs.TwoLstBxs TwoLstBxs1 
      Height          =   2970
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5239
      DatabaseName    =   "Database.mdb"
      RecordSource    =   "ClientDatabase"
      Caption1        =   "Clients not Selected"
      Caption2        =   "Clients Selected"
      FieldName       =   "FullBusinessName"
      IDFieldName     =   "ID"
      Password        =   "password"
      SortBy          =   "Principal1Name"
      L1BackColor     =   16761024
      L2BackColor     =   16761024
      L1ForeColor     =   -2147483635
      L2ForeColor     =   -2147483635
      ForeColor       =   192
      CaptionBold     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SQLString       =   "WHERE (ID >= 125) ORDER BY FullBusinessName;"
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   5160
      Width           =   6135
      Begin VB.TextBox txtRecCount 
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   3720
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Full Business Name"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "ID"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Record Count"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   6135
      Begin VB.CommandButton cmdClearL2 
         Caption         =   "Clear"
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelAllL2 
         Caption         =   "Select All"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSortL2 
         Caption         =   "Sort"
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdMoveToL1 
         Caption         =   "<< Move"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearL1 
         Caption         =   "Clear"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelAllL1 
         Caption         =   "Select All"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSortL1 
         Caption         =   "Sort"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdMoveToL2 
         Caption         =   "Move >>"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
      Caption         =   "If you liked this control, please, vote and post some comments..."
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   6360
      Width           =   3495
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isConnected As Boolean
Dim strFieldName As String
Dim strIDField As String

Private Sub Form_Load()
    strFieldName = TwoLstBxs1.FieldName
    strIDField = TwoLstBxs1.IDFieldName
    
End Sub

'=================================================
'=========== Manage ListBox 1 ====================
'=================================================

Private Sub cmdMoveToL2_Click()
    TwoLstBxs1.MoveToList2
    
End Sub

Private Sub cmdSortL1_Click()
    TwoLstBxs1.SortList1
    
End Sub

Private Sub cmdSelAllL1_Click()
    TwoLstBxs1.SelectAllList1
    
End Sub

Private Sub cmdClearL1_Click()
    TwoLstBxs1.ClearList1
    
End Sub

'=================================================
'=========== Manage ListBox 2 ====================
'=================================================

Private Sub cmdMoveToL1_Click()
    TwoLstBxs1.MoveToList1
    
End Sub

Private Sub cmdSortL2_Click()
    TwoLstBxs1.SortList2
    
End Sub

Private Sub cmdSelAllL2_Click()
    TwoLstBxs1.SelectAllList2
    
End Sub

Private Sub cmdClearL2_Click()
    TwoLstBxs1.ClearList2
    
End Sub

'===========================================================
'=========== Connect to Database ===========================
'===========================================================

Private Sub cmdConnect_Click()
    Call TwoLstBxs1.RSFinalConnect
    
    txtRecCount.Text = TwoLstBxs1.RSFinal.RecordCount
    
    If (Val(txtRecCount.Text) > 0) Then
        isConnected = True
        Call DisplayItems
        
    Else
        isConnected = False
        
    End If
    
End Sub

Private Sub cmdFirst_Click()
    If (isConnected) Then
        TwoLstBxs1.RSFinal.MoveFirst
        Call DisplayItems
    End If
    
End Sub

Private Sub cmdPrevious_Click()
    If (isConnected) Then
        TwoLstBxs1.RSFinal.MovePrevious
        If (TwoLstBxs1.RSFinal.BOF) Then
            TwoLstBxs1.RSFinal.MoveFirst
        End If
        Call DisplayItems
    End If
    
End Sub

Private Sub cmdNext_Click()
    If (isConnected) Then
        TwoLstBxs1.RSFinal.MoveNext
        If (TwoLstBxs1.RSFinal.EOF) Then
            TwoLstBxs1.RSFinal.MoveLast
        End If
        Call DisplayItems
    End If
    
End Sub

Private Sub cmdLast_Click()
    If (isConnected) Then
        TwoLstBxs1.RSFinal.MoveLast
        Call DisplayItems
    End If
    
End Sub

Private Sub DisplayItems()
    txtID.Text = TwoLstBxs1.RSFinal.Fields(strIDField).Value & ""
    txtUserName.Text = TwoLstBxs1.RSFinal.Fields(strFieldName).Value & ""
    
    txtRecCount.Text = TwoLstBxs1.RSFinal.RecordCount
End Sub

'============================================================

Private Sub Command1_Click()
    'The API that allows me to open the browser
    'in shell mode is on the modShell Module.
    Call Shell("cmd /c start http://www.planet-source-code.com/vb/default.asp?lngCId=36715&lngWId=1")
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    'I had to call this API a second time
    'because, some times, PSC opens a default
    'page instead of the page with my submission...
    Call Shell("cmd /c start http://www.planet-source-code.com/vb/default.asp?lngCId=36715&lngWId=1")
    
End Sub
