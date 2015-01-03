VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmTimerEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Timer Edit"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Text            =   "Text1"
      ToolTipText     =   "Save this many active days of history for this timer"
      Top             =   3120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type"
      Height          =   540
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   4335
      Begin VB.OptionButton Option1 
         Caption         =   "Personal"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   210
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Business"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   2550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00C0C0C0&
      Height          =   690
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmTimerEdit.frx":0000
      ToolTipText     =   "Click to select a location. Even if you leave empty the timers themselves are still stored internally."
      Top             =   2430
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   3180
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   2
      Top             =   3180
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   630
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "Separating keywords with a space = AND. Separating with a comma = OR. A single ""*"" with no keywords =  global timer"
      Top             =   1470
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   4335
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Empty"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      ToolTipText     =   "Clear the timestamp path (Do not save any timestamp)"
      Top             =   2190
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Save a timestamp copy in:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   12
      Top             =   2190
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Days History Depth"
      Height          =   255
      Index           =   3
      Left            =   780
      TabIndex        =   11
      Top             =   3165
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Sniff Keywords"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1245
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Timer Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   975
   End
End
Attribute VB_Name = "frmTimerEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ThisTimerType As Integer

Option Explicit
Private Sub Form_Load()
    Dim CursorXY As CursorState
    
    frmMainWindow.Enabled = False
    If (frmMainWindow.Check1(0) = 1) Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 'On top of all windows Else SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3
        
    TimerEdited.IsNewHistoryDepth = False
    TimerEdited.TimerIndex = -1
    
    'The mouse coordinates are not in twips
    CursorXY = DoFindCursorState(CursorXY)
    Me.Left = CursorXY.X * Screen.TwipsPerPixelX - frmTimerEdit.Width / 10
    Me.Top = CursorXY.y * Screen.TwipsPerPixelX - frmTimerEdit.Height / 2.5
    Me.Tag = frmMainWindow.TimersListView.SelectedItem.Index 'The Me.tag property contains the originating selected Timer
    
    Text1(0).Text = frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).Text
    Text1(1).Text = frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).ListSubItems(3).Text
    Text1(2).Text = frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).ListSubItems(4).Text
    Text1(3).Text = frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).ListSubItems(4).Tag
    
    If (Text1(2).Text = "") Then Check1.Value = 1 Else Check1.Value = 0
    
    ThisTimerType = frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).ListSubItems(7).Tag
    Option1(ThisTimerType) = True
        
    'initialize the commondialog control
    CommonDialog1.Filter = "TextFiles|*.txt"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DefaultExt = ".txt"
    CommonDialog1.DialogTitle = "Timer will be saved in"
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMainWindow.Enabled = True
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim IsUpdated As Boolean
    Select Case Index
        Case 0 'cancel
            Unload Me
        Case 1 'apply
            'check for duplicate timer name
            '...
            
            'Update the timer
            IsUpdated = TestAndUpdate(frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)), Text1(0).Text)  'TimerName
            IsUpdated = TestAndUpdate(frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).ListSubItems(3), Text1(1).Text) 'Sniff keywords
            IsUpdated = TestAndUpdate(frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).ListSubItems(4), Text1(2).Text) 'Save in
            IsUpdated = TestAndUpdate(frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).ListSubItems(4), CInt(Abs(Val(Text1(3).Text))), True) 'History depth
            If (IsUpdated) Then TimerEdited.IsNewHistoryDepth = True
            IsUpdated = TestAndUpdate(frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).ListSubItems(7), DoGetTimerTypes(ThisTimerType)) 'Timer type text
            IsUpdated = TestAndUpdate(frmMainWindow.TimersListView.ListItems(CLng(Me.Tag)).ListSubItems(7), ThisTimerType, True) 'Timer type code
            
            Command1_Click (0)
    End Select
End Sub

Private Sub Check1_Click()
    If (Check1.Value = 1 And Text1(2).Text <> "") Then
        Text1(2).Text = ""
    ElseIf (Check1.Value = 0 And Text1(2).Text = "") Then
        Check1.Value = 1
        Call Text1_Click(2)
    End If
End Sub

Private Function TestAndUpdate(OldVar As Object, NewVar As Variant, Optional IsTag As Boolean = False) As Boolean
    TestAndUpdate = False
    
    Select Case IsTag
        Case False
            If (NewVar <> OldVar.Text) Then
                If (TimerEdited.TimerIndex <= 0) Then TimerEdited.TimerIndex = CLng(Me.Tag)
                OldVar.Text = NewVar
                TestAndUpdate = True
            End If
        Case True
            If (NewVar <> OldVar.Tag) Then
                If (TimerEdited.TimerIndex <= 0) Then TimerEdited.TimerIndex = CLng(Me.Tag)
                OldVar.Tag = NewVar
                TestAndUpdate = True
            End If
    End Select
    
End Function

Private Sub Option1_Click(Index As Integer)
    ThisTimerType = Index
End Sub

Private Sub Text1_Change(Index As Integer)
    Select Case Index
        Case 3 'History depth
            If (Val(Text1(3).Text) < 1 Or Val(Text1(3).Text) <> Abs(Int(Val(Text1(3))))) Then Text1(3).Text = Abs(Int(Val(Text1(3))))
    End Select
End Sub

Private Sub Text1_Click(Index As Integer)
    Dim Directories
    Dim Path As String
    On Error Resume Next
    Err.Clear
    Select Case Index
        Case 2 'select directory to save timer
            'Text1(2).Text = DoSelectDir 'alternate method for selecting directory only
            Path = Text1(2).Text
            If (Path <> "") Then
                Directories = Split(Path, "\")
                Path = Left(Path, Len(Path) - Len(Directories(UBound(Directories))))
            End If
            If ((GetAttr(Path) And vbDirectory) <> vbDirectory) Then Path = "C:\"
            CommonDialog1.FileName = Path & Text1(0).Text & "-ProjectTime"
            CommonDialog1.ShowSave
            If (Err.Number <> cdlCancel) Then Text1(2).Text = CommonDialog1.FileName
            Text1(1).SetFocus
            If (Text1(2).Text <> "" And Check1.Value = 1) Then Check1.Value = 0
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii = 13) Then Command1_Click (1)
End Sub
