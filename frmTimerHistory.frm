VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmTimerHistory 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Related History for:"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Today"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Running Total"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   1
      Left            =   1220
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   600
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   120
      X2              =   3720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   2625
   End
End
Attribute VB_Name = "frmTimerHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HistoryCount As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    frmMainWindow.Enabled = False
    If (frmMainWindow.Check1(0) = 1) Then SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 'On top of all windows Else SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3

    Me.Left = frmMainWindow.Left + (frmMainWindow.Width - Me.Width) / 2
    Me.Top = frmMainWindow.Top + (frmMainWindow.Height - Me.Height) / 2
    
    Me.Tag = frmMainWindow.TimersListView.SelectedItem.Index 'The Me.tag property contains the originating selected Timer

    Label1.Caption = frmMainWindow.TimersListView.SelectedItem.Text
    
    Call InitializeListViewHeaders
    HistoryCount = ShowRelatedHistoryFor(CLng(Me.Tag))
    
    Me.Caption = "Related history (" & HistoryCount & " entries) for:"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMainWindow.Enabled = True
End Sub

Private Sub InitializeListViewHeaders()
    ListView1.ColumnHeaders.Add 1, , "Date", 1050
    ListView1.ColumnHeaders.Add 2, , "Running Total", 1150, 2
    ListView1.ColumnHeaders.Add 3, , "Toatal Today", 1100, 2
End Sub

'Return count of items found
Private Function ShowRelatedHistoryFor(ThisTimer As Long) As Long
    Dim Count As Long
    Dim FoundCount As Long
    
    For Count = 1 To TmrDatabase.TimersHistory.Count
        If (CLng(TmrDatabase.TimersHistory(Count).LinkToTimerID) = CLng(TmrDatabase.MyTimers(ThisTimer).Tag)) Then
            FoundCount = FoundCount + 1
            ListView1.ListItems.Add FoundCount, , TmrDatabase.TimersHistory(Count).OnThisDate
            ListView1.ListItems(FoundCount).ListSubItems.Add 1, , DoShowTime(CDbl(TmrDatabase.TimersHistory(Count).RunningTotal))
            ListView1.ListItems(FoundCount).ListSubItems.Add 2, , TmrDatabase.TimersHistory(Count).TotalToday
            
        End If
    Next
    
    ShowRelatedHistoryFor = FoundCount
End Function

Private Sub Form_Resize()

    If (Me.Width <> 3990) Then
        Me.Enabled = False
        Me.Width = 3990
    ElseIf (Me.Height < 4035) Then
        Me.Enabled = False
        Me.Height = 4035
    End If
    
    Call DoAdjustLayout

End Sub

Private Sub DoAdjustLayout()

    Me.Enabled = True
    'Vertical changes
    ListView1.Height = Me.Height - ListView1.Top - 540

End Sub
