VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMainWindow 
   Caption         =   "Project Timers"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9495
   Icon            =   "frmMainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   9495
   Begin VB.Timer LightTimer 
      Left            =   8400
      Top             =   2280
   End
   Begin MSComctlLib.ListView StatsView 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   2805
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   556
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tray"
      Height          =   255
      Index           =   1
      Left            =   945
      TabIndex        =   9
      ToolTipText     =   "Send to tray when minimizing"
      Top             =   105
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pin"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Pin window on top of all others on the desktop"
      Top             =   105
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   2
      ToolTipText     =   "Threshold amount of idle mouse time (in sec) beyond which all timers stop"
      Top             =   80
      Width           =   495
   End
   Begin MSComctlLib.ListView ReportView 
      Height          =   480
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3120
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   847
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483631
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   8400
      MaxLength       =   4
      TabIndex        =   1
      ToolTipText     =   "Frequency of detection (1-15 sec)"
      Top             =   80
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   8850
      Top             =   2280
   End
   Begin MSComctlLib.ListView TimersListView 
      Height          =   2385
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Use mouse second button to reveal context menu."
      Top             =   420
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   4207
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Label Feedback 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   105
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   8
      Top             =   3620
      Width           =   3015
   End
   Begin VB.Label labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   1
      Left            =   165
      TabIndex        =   5
      Top             =   3620
      Width           =   3015
   End
   Begin VB.Label labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Sniff every | Threshold (sec)"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   3
      Top             =   110
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00A0A0A0&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   480
      Width           =   210
   End
   Begin VB.Menu ContextMenu 
      Caption         =   "ContextMenu"
      Visible         =   0   'False
      Begin VB.Menu ContextMenuItem 
         Caption         =   "Add Timer"
         Index           =   0
      End
      Begin VB.Menu ContextMenuItem 
         Caption         =   "Edit Timer"
         Index           =   1
      End
      Begin VB.Menu ContextMenuItem 
         Caption         =   "Save Stamp"
         Index           =   2
      End
      Begin VB.Menu ContextMenuItem 
         Caption         =   "Delete"
         Index           =   3
      End
      Begin VB.Menu ContextMenuItem 
         Caption         =   "History"
         Index           =   4
      End
      Begin VB.Menu ContextMenuItem 
         Caption         =   "Reset Trip"
         Index           =   5
      End
      Begin VB.Menu ContextMenuItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu ContextMenuItem 
         Caption         =   "Columns"
         Index           =   7
         Begin VB.Menu ColumnsItem 
            Caption         =   "Total Time"
            Checked         =   -1  'True
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "frmMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum Menu
    Add = 0
    Edit = 1
    Save = 2
    DELETE = 3
    showhistory = 4
    ResetTrip = 5
    Columns = 7
End Enum
Private Sub Form_Load()

    Call CheckForMultipleInstances
    
    'Preliminary housekeeping
    Call DoInitDefaultDir
    Call DoInitializeListViewHeaders 'It is important for the headers to be initialized very early
    Call GetDefaultWindowSizeAndPos
    
    TimerEdited.IsNewHistoryDepth = False
    TimerEdited.TimerIndex = -1
    TmrDatabase.IdleThreshold = 300 'This is in seconds
    TmrDatabase.SniffInterval = 2 'This is in seconds
    
    'Retreive settings from file
    Call DoRetreiveSettings(TmrDatabase)
    Call DoInitializeContextMenu
    
    'Load the options text fields and check them
    Text1(0).Text = TmrDatabase.SniffInterval
    Text1(1).Text = TmrDatabase.IdleThreshold
    
    labels(1).Caption = "version 20090702"
    labels(3).Caption = "Thomas Anagnostou - Rayflectar Graphics"
    Feedback.ForeColor = Color.red
    Feedback.Caption = "All timers paused"
    
    IsWindowInTray = 0
    TrayMotion.UserHoverInTray = False
    Timer1.Interval = TmrDatabase.SniffInterval * 1000
    LightTimer.Interval = 500 'half a second (This is only used for tray sensing and it is very light)
    
    'Setup Autosave
    Dim InitAutosave As AutoSave
    InitAutosave.SessionON = Now
    InitAutosave.LastSavedON = Now
    InitAutosave.IsTimestampSavePending = False
    InitAutosave.SaveEvery = 5 * Seconds.inMin 'This is in seconds
    InitAutosave.TimestampSaveDelay = 1 * Seconds.inMin
    InitAutosave.AreTimersChanged = False
    Call DoSetAutoSaveData(InitAutosave)
    
    'Load frmTrayView
    
    'Run the timer. Forced call, so there will be no delay indisplaying the stats and the sniffed windows
    Call Timer1_Timer
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case 0, 1, 3, 4, 5 'vbFormControlMenu, vbFormCode
            Dim Response As Variant
    
            'Popup warning/confirmation
            Response = MsgBox("Stop all timer monitoring and exit Project Timers ?", vbOKCancel, "Confirmation")
            If (Response = vbCancel) Then
                Cancel = vbCancel
                Exit Sub
            End If
            
            'also save timers to their respective locations
            Call DoSaveAllToUserDir(TmrDatabase)
        Case Else
            'also save timers to their respective locations
            Call DoSaveAllToUserDir(TmrDatabase, , True)
    End Select
    
    'Save timers into a catch-all ini file. Next time the program loads it should be able to continue
    Call DoStoreSettings(TmrDatabase)
    
    'unload forms
    Unload frmTimerEdit
    Unload frmMainWindow
    Unload frmTrayView
    Unload frmTimerHistory
End Sub

'This is a fast timer so be carefull it *has* to be extremely light ( do not load anything other than tray sensing
Private Sub LightTimer_Timer()
    Dim CurrentMouseXY As CursorState
    
    'Avoid unecessary probing
    If (TrayMotion.UserHoverInTray = False) Then Exit Sub
    
    'Handle TrayView
    GetCursorPos CurrentMouseXY
    If (Not (IsWindowInTray = 1 And CurrentMouseXY.X < (TrayMotion.EntryXY.X + 4) And CurrentMouseXY.X > (TrayMotion.EntryXY.X - 4) And CurrentMouseXY.y < (TrayMotion.EntryXY.y + 4) And CurrentMouseXY.y > (TrayMotion.EntryXY.y - 4))) Then
        frmTrayView.Hide
        SetWindowPos frmTrayView.hWnd, 1, 0, 0, 0, 0, 3 'On bottom of all windows
        Unload frmTrayView
        TrayMotion.UserHoverInTray = False
    End If
End Sub

'This is the heartbeat of this program (everything happens from here)
Private Sub Timer1_Timer()
    Dim TimersCondition As Integer

    'Detect active window captions
    ActiveWindow = DoSniffActiveWindow
    Call DoShowSniffedWindows
        
    'Detect mouse position and idle age
    MouseCondition = DoFindCursorState(MouseCondition)
    
    'Update mouse statistics
    Call UpdateMouseStatistics(TmrDatabase, Not CBool(MouseCondition.age))
    
    'Detect ProjectTimers current window size and position (only if window is normal)
    If (IsWindowInTray = 0 And frmMainWindow.WindowState = 0) Then Call GetCurrentMainWindowSizeAndPos
    
    'Update timers
    If (MouseCondition.age * Seconds.inday < TmrDatabase.IdleThreshold) Then
        Feedback.Visible = False
        
        'Test timers and return timer condition
        TimersCondition = DoUpdateTimers(TmrDatabase.MyTimers, TmrDatabase.TimersHistory, TmrDatabase.SniffInterval, ActiveWindow)
        
        'Update Timer Statistics
        If (TimersCondition > 0) Then Call UpdateTimerStatistics(TmrDatabase)
    Else
        'Tone down the color of the timers and keep checking the "new day condition"
        TimersCondition = CoolDown(TmrDatabase.MyTimers)
        
        'Show feedback
        Feedback.Visible = True
    End If
    
    'Refresh to reveal new color changes
    If (TimersCondition = 1) Then TimersListView.Refresh
    
    'Display statistics
    Call DisplayStatistics
    
    'Autosave
    Call DoAutoSave(TmrDatabase)
    
End Sub

Private Sub DisplayStatistics()

    Dim OutputString As String

    'All timers total
    OutputString = "All Timers Total: " & DoShowTime(CDate(TmrDatabase.AllTimeTotal), True, 2, True, True)
    StatsView.ListItems(1).ListSubItems(2).Text = OutputString
    
    'Today Timer Total
    OutputString = "Today Total: " & DoShowTime(CDbl(CDate(TmrDatabase.TodayTotal)), True, 2, True)
    StatsView.ListItems(1).ListSubItems(3).Text = OutputString
    
    'Update mouse idle and today idle tooltip
    OutputString = "Mouse Idle for: " & DoShowTime(MouseCondition.age, True, 2, True, True)
    StatsView.ListItems(1).ListSubItems(1).Text = OutputString
    OutputString = "Today: " & DoShowTime(CDate(TmrDatabase.MouseIdleToday), True, 2, True)
    OutputString = OutputString & " idle / " & DoShowTime(CDate(TmrDatabase.MouseBusyToday), True, 2, True) & " in motion"
    StatsView.ListItems(1).ListSubItems(1).Tag = OutputString 'place the tooltip in the tag (it will be shown from the mousemove event
    Select Case GetAbsColumnHit(Me, StatsView)
        Case 1
            StatsView.ToolTipText = "General stats view"
        Case 2
            StatsView.ToolTipText = StatsView.ListItems(1).ListSubItems(1).Tag
        Case 3
            StatsView.ToolTipText = "Sum of all history from all timers"
        Case 4
            StatsView.ToolTipText = "Sum of today's activity from all timers"
        Case Else
            StatsView.ToolTipText = ""
    End Select
    If (MouseCondition.age * Seconds.inday < 15) Then StatsView.ListItems(1).ListSubItems(1).ForeColor = Color.DisabledText Else StatsView.ListItems(1).ListSubItems(1).ForeColor = Color.red
    
End Sub

'When Form activates it is either freshly loaded or came back from a popup window (or from the tray)
Private Sub Form_Activate()

    If (TimerEdited.IsNewHistoryDepth) Then Call DeleteOldestHistoryEntries(TmrDatabase.TimersHistory, TmrDatabase.MyTimers(TimerEdited.TimerIndex).Tag, TmrDatabase.MyTimers(TimerEdited.TimerIndex).ListSubItems(4).Tag)
    If (TimerEdited.IsNewHistoryDepth Or TimerEdited.TimerIndex > 0) Then
        'Call DoAutoSave(TmrDatabase, True)
        Call DoStoreSettings(TmrDatabase)
        Call DoSaveOneTimer(TmrDatabase, TimerEdited.TimerIndex)
    End If
    
    TimerEdited.IsNewHistoryDepth = False
    TimerEdited.TimerIndex = -1
    Call UpdateTimersContextMenu
End Sub

'This is used for to detect click while minized to tray
'What happens is that when minimized, hovering the mouse over the tray icon returns the exact same pixel coordinate (the whole icon acts as a single pixel, in this case pixel 512)
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single) 'x and y are in tweeps
    Dim Xpixel As Long
    Dim TrayWindowLeft As Long
    Dim TrayWindowTop As Long
    
    If (IsWindowInTray = 0) Then Exit Sub
    Xpixel = X / Screen.TwipsPerPixelX
    Select Case Xpixel
        Case 512
            GetCursorPos TrayMotion.EntryXY
            TrayWindowLeft = TrayMotion.EntryXY.X * Screen.TwipsPerPixelX - frmTrayView.Width / 2
            TrayWindowTop = TrayMotion.EntryXY.y * Screen.TwipsPerPixelY - frmTrayView.Height - 100
            If (Screen.Width < (TrayWindowLeft + frmTrayView.Width)) Then TrayWindowLeft = Screen.Width - frmTrayView.Width
            If (TrayWindowLeft < 0) Then TrayWindowLeft = 0
            If (TrayWindowTop < 0) Then TrayWindowTop = TrayMotion.EntryXY.y + 500
            frmTrayView.Left = TrayWindowLeft
            frmTrayView.Top = TrayWindowTop
            If (TrayMotion.UserHoverInTray = False) Then
                Load frmTrayView
                Call RefreshTrayView
                frmTrayView.Show
                frmTrayView.Text1.SetFocus
                TrayMotion.UserHoverInTray = True
                SetWindowPos frmTrayView.hWnd, -1, 0, 0, 0, 0, 3 'On top of all windows
            End If
        Case WM_LButtonDown '513
            frmMainWindow.Show ' show form
            IsWindowInTray = 0
            Shell_NotifyIcon NIM_Delete, nid ' del tray icon
        Case WM_LButtonUp
        Case WM_LButtonDblClk
        Case WM_RButtonDown
        Case WM_RButtonUp '517
            frmMainWindow.Show
            IsWindowInTray = 0
            Shell_NotifyIcon NIM_Delete, nid
        Case WM_RButtonDblClk
    End Select
    
End Sub

Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 0 'Pin on top
            If (Check1(0).Value = 1) Then
                SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3 'On top of all windows
            Else
                SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3 'On top of all windows
            End If
        Case 1 'Use tray when minimizing
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    'User pressing enter
    If (KeyAscii = 13) Then TimersListView.SetFocus: Exit Sub
    'Block alpha characters
    If (KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57)) Then KeyAscii = 0
    
End Sub

Private Sub Text1_Change(Index As Integer)

    Select Case Index
        Case 0
            'Check user input for sniffInterval (it must be numeric and in a certain range
            If (TmrDatabase.SniffInterval = Val(Text1(Index).Text)) Then Exit Sub
            TmrDatabase.SniffInterval = Val(Text1(Index).Text)
            If (TmrDatabase.SniffInterval < 1) Then TmrDatabase.SniffInterval = 1
            If (TmrDatabase.SniffInterval > 15) Then TmrDatabase.SniffInterval = 15
            Timer1.Interval = (Val(TmrDatabase.SniffInterval) * 1000) 'This is in milliseconds
            Text1(Index).Text = Val(TmrDatabase.SniffInterval)
        Case 1
            'check user input for idleThreshold (it must be numeric and in a certain range)
            If (TmrDatabase.IdleThreshold = Val(Text1(Index).Text)) Then Exit Sub
            TmrDatabase.IdleThreshold = Val(Text1(Index).Text)
            If (TmrDatabase.IdleThreshold < 1) Then TmrDatabase.IdleThreshold = 1
            If (TmrDatabase.IdleThreshold > 999) Then TmrDatabase.IdleThreshold = 999
            Text1(Index).Text = Val(TmrDatabase.IdleThreshold)
    End Select
    
    'Issue a color warning if the threshold is less than the sniff interval
    If (TmrDatabase.SniffInterval > TmrDatabase.IdleThreshold) Then
        'Text1(1).ForeColor = Color.red
        'TmrDatabase.SniffInterval = TmrDatabase.IdleThreshold
    ElseIf (Text1(1).ForeColor = Color.red) Then
        'Text1(1).ForeColor = Color.ButtonText
    End If
    
End Sub

Private Sub Text1_lostfocus0(Index As Integer)
    '
End Sub

Private Sub DoShowSniffedWindows()
    'Display detected window captions using the bottom listview control
    ReportView.ListItems(1).ListSubItems(1).Text = ActiveWindow.ParentHandle
    ReportView.ListItems(1).ListSubItems(2).Text = ActiveWindow.ParentCaption
    ReportView.ListItems(2).ListSubItems(1).Text = ActiveWindow.ChildHandle
    ReportView.ListItems(2).ListSubItems(2).Text = ActiveWindow.ChildCaption
End Sub

'Sort by column header
Private Sub TimersListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If (frmMainWindow.TimersListView.Sorted = False) Then frmMainWindow.TimersListView.Sorted = True
    Select Case ColumnHeader.Index
        Case 1
            frmMainWindow.TimersListView.SortKey = 0
        Case 2
            frmMainWindow.TimersListView.SortKey = 1
        Case 3
            frmMainWindow.TimersListView.SortKey = 2
        Case 4
            frmMainWindow.TimersListView.SortKey = 3
        Case 5
            frmMainWindow.TimersListView.SortKey = 4
        Case 6
            frmMainWindow.TimersListView.SortKey = 5
        Case 7
            frmMainWindow.TimersListView.SortKey = 6
        Case 8
            frmMainWindow.TimersListView.SortKey = 7
        Case Else
            Exit Sub
    End Select
    If (frmMainWindow.TimersListView.SortOrder = lvwDescending) Then
        frmMainWindow.TimersListView.SortOrder = lvwAscending
    Else
        frmMainWindow.TimersListView.SortOrder = lvwDescending
    End If
End Sub

Private Sub TimersListView_DblClick()
    'This should cause the timer options window to come up
    ContextMenuItem_Click (Menu.Edit)
    'This should cause a label edit in the timersListView control
    'SendKeys (Chr(13))
End Sub

Private Sub TimersListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call UpdateTimersContextMenu
End Sub

Private Sub TimersListView_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 113, 13 'Pressing the F2 key or enter
            If (Not (TimersListView.SelectedItem Is Nothing)) Then TimersListView.StartLabelEdit
        Case 8, 46 'pressing the backspace or delete keys
            ContextMenuItem_Click (Menu.DELETE)
    End Select
End Sub

Private Sub TimersListView_AfterLabelEdit(Cancel As Integer, NewString As String)
    'Test for duplicate Timer names
    '...
    
    'Make sure the timer remains highlighted
    TmrDatabase.MyTimers(TimersListView.SelectedItem.Index).Selected = True
End Sub

Private Sub TimersListView_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'Check for RightMouseClick
    If Button = 2 Then PopupMenu ContextMenu
End Sub

'Handle context menu commands
Private Sub ContextMenuItem_Click(Index As Integer)
    Dim SelectedTimerIndex As Long
    Dim Message As String
    Dim UserDecision As Integer
    
    Select Case Index
        Case Menu.Add 'add item
            TimersListView.Sorted = False
            'add a timer at the position of the selected line in the listview control
            If (Not (TimersListView.SelectedItem Is Nothing)) Then
                SelectedTimerIndex = TimersListView.SelectedItem.Index + 1
            Else
                SelectedTimerIndex = 1
            End If
            
            'Add an empty timer
            TmrDatabase.MyTimers.Add SelectedTimerIndex
            Call SetTimerToDefault(TmrDatabase.MyTimers, CLng(SelectedTimerIndex))
            
            'Highlight what was just added
            TmrDatabase.MyTimers(SelectedTimerIndex).Selected = True
    
            'Custom Color
            Call ListItemColor(TmrDatabase.MyTimers(SelectedTimerIndex), , , , InterfaceOptions.UserColumnColor)
            
            'Enter into label edit mode by sending a key to the console (this will invoke the listvew_keydown event)
            SendKeys (Chr(13))
            
            'Enable appropriate contextmenu items
            Call UpdateTimersContextMenu
            
        Case Menu.Edit 'edit item
            If (Not (TimersListView.SelectedItem Is Nothing)) Then
                Load frmTimerEdit
                frmTimerEdit.Show
                frmTimerEdit.Text1(1).SetFocus
            End If
        Case Menu.Save 'save item
            Call DoSaveOneTimer(TmrDatabase, TimersListView.SelectedItem.Index, False)
        Case Menu.DELETE 'delete item
            If (TimersListView.SelectedItem Is Nothing) Then Exit Sub
            Message = "Delete the selected timer <" & TimersListView.SelectedItem.Text & "> ?"
            UserDecision = MsgBox(Message, vbOKCancel, "Confirm")
            If (UserDecision = 1) Then
                Call DeleteTimerHistory(TimersListView.SelectedItem.Tag, TmrDatabase.TimersHistory)
                TmrDatabase.MyTimers.Remove (TimersListView.SelectedItem.Index)
            End If
            
            If (TmrDatabase.MyTimers.Count > 0) Then
                TimersListView.SelectedItem.Selected = True
            Else
                'Disable appropriate contextmenu items
                Call UpdateTimersContextMenu
            End If
        Case Menu.ResetTrip 'Zero the TripCounter
            Call TripCounterReset(TimersListView.SelectedItem)
        Case Menu.showhistory 'Popup history display
            If (Not (TimersListView.SelectedItem Is Nothing)) Then
                Load frmTimerHistory
                frmTimerHistory.Show
                frmTimerHistory.Command1.SetFocus
            End If
    End Select
    
End Sub

Private Sub ColumnsItem_Click(Index As Integer)
    Dim Count As Integer
    
    'Store the previous state
    For Count = 1 To TimersListView.ColumnHeaders.Count
        If (CLng(TimersListView.ColumnHeaders(Count).Width) > 10 And InterfaceOptions.UserColumnVisible(Count) = 1 And InterfaceOptions.DefaultIsColWidthEditable(Index) = 1) Then
            InterfaceOptions.UserColWidth(Count) = CLng(TimersListView.ColumnHeaders(Count).Width)
        ElseIf (CLng(TimersListView.ColumnHeaders(Count).Width) <= 10 And InterfaceOptions.UserColumnVisible(Count) = 1 And InterfaceOptions.DefaultIsColWidthEditable(Index) = 1) Then
            InterfaceOptions.UserColWidth(Count) = InterfaceOptions.DefaultColWidth(Count)
        End If
    Next
    
   If (ColumnsItem(Index).Checked = True) Then ColumnsItem(Index).Checked = False Else ColumnsItem(Index).Checked = True
   InterfaceOptions.UserColumnVisible(Index) = Abs(CInt(ColumnsItem(Index).Checked))
   
   TimersListView.ColumnHeaders(Index).Width = InterfaceOptions.UserColWidth(Index) * InterfaceOptions.UserColumnVisible(Index)
End Sub

Private Sub DoInitializeContextMenu()
    Dim Count As Integer
    Dim IsEnabled As Boolean
    'Populate context menu with vaious commands
    'Context menu is a control, but you cannot add it via drag/drop into the form. You have to use the menu editor
    'By use of the menu editor the first menu item is already inserted. The rest can be added programatically (as for example in here below)
    
    'Load ContextMenuItem(Menu.Edit) 'Already loaded fomr the menu control (under Tools)
    'No need to load the main contextmenu items ... they are already in by the Menu editor
    
    'Load all columns in the columns submenu
    'The first column is the title (user cannot hide it)
    'The second column is already loaded by the menu editor)
    'So start from the third column)
    For Count = 2 To TimersListView.ColumnHeaders.Count
        If (Count > 2) Then Load ColumnsItem(Count)
        If (Count = 3) Then ColumnsItem(Count).Caption = "Indicator" Else ColumnsItem(Count).Caption = TimersListView.ColumnHeaders(Count).Text
        ColumnsItem(Count).Checked = CBool(InterfaceOptions.UserColumnVisible(Count))
        ColumnsItem(Count).Enabled = True
        ColumnsItem(Count).Visible = True
    Next
    
End Sub

'Add column headers to the listview controls
Private Sub DoInitializeListViewHeaders()
    Dim Count As Integer
    
    Set TmrDatabase.MyTimers = TimersListView.ListItems
    Set TmrDatabase.TimersHistory = New Collection
    'Set TmrDatabase.TimersHistory = New Dictionary
    
    'This is the upper (main) listview control, showing all user's timers
    'Never change the order of creation here (the entire program, all modules are hardwired to this particular order). Let the user rearrange columns at run time.
    TimersListView.ColumnHeaders.Add 1, , "Timer Name"
    TimersListView.ColumnHeaders.Add 2, , "Total Time", , 2
    TimersListView.ColumnHeaders.Add 3, , "", , 0
    TimersListView.ColumnHeaders.Add 4, , "Sniff Keywords"
    TimersListView.ColumnHeaders.Add 5, , "Save a Timestamp In"
    TimersListView.ColumnHeaders.Add 6, , "Created", , 2
    TimersListView.ColumnHeaders.Add 7, , "Today", , 2
    TimersListView.ColumnHeaders.Add 8, , "Type", , 2
    TimersListView.ColumnHeaders.Add 9, , "TripCount", , 2
    TimersListView.ColumnHeaders.Add 10, , "Last Trip Reset", , 2
    
    'Default Column Width
    ReDim InterfaceOptions.DefaultColWidth(1 To TimersListView.ColumnHeaders.Count)
    InterfaceOptions.DefaultColWidth(1) = 1776 'Name
    InterfaceOptions.DefaultColWidth(2) = 62 * 15  'total time
    InterfaceOptions.DefaultColWidth(3) = 16 * 15  'indicator
    InterfaceOptions.DefaultColWidth(4) = 1860 'Sniff Keywords
    InterfaceOptions.DefaultColWidth(5) = 1750 'Save a Timestamp In
    InterfaceOptions.DefaultColWidth(6) = 70 * 15  'created
    InterfaceOptions.DefaultColWidth(7) = 55 * 15  'today
    InterfaceOptions.DefaultColWidth(8) = 38 * 15  'type
    InterfaceOptions.DefaultColWidth(9) = 62 * 15  'Time Since Reset
    InterfaceOptions.DefaultColWidth(10) = 100 * 15  'Last Reset Date
    InterfaceOptions.UserColWidth = InterfaceOptions.DefaultColWidth
    
    'Apply column widths to the listview
    For Count = 1 To TimersListView.ColumnHeaders.Count
        TimersListView.ColumnHeaders(Count).Width = InterfaceOptions.UserColWidth(Count)
    Next
    
    'Default Editable column widths
    ReDim InterfaceOptions.DefaultIsColWidthEditable(1 To TimersListView.ColumnHeaders.Count)
    InterfaceOptions.DefaultIsColWidthEditable(1) = 1
    InterfaceOptions.DefaultIsColWidthEditable(2) = 0
    InterfaceOptions.DefaultIsColWidthEditable(3) = 0
    InterfaceOptions.DefaultIsColWidthEditable(4) = 1
    InterfaceOptions.DefaultIsColWidthEditable(5) = 1
    InterfaceOptions.DefaultIsColWidthEditable(6) = 0
    InterfaceOptions.DefaultIsColWidthEditable(7) = 0
    InterfaceOptions.DefaultIsColWidthEditable(8) = 0
    InterfaceOptions.DefaultIsColWidthEditable(9) = 0
    InterfaceOptions.DefaultIsColWidthEditable(10) = 0
    
    'Default Column Color
    ReDim InterfaceOptions.DefaultColumnColor(0 To TimersListView.ColumnHeaders.Count - 1)
    InterfaceOptions.DefaultColumnColor(0) = Color.ButtonText
    InterfaceOptions.DefaultColumnColor(1) = Color.ButtonText
    InterfaceOptions.DefaultColumnColor(2) = Color.ButtonText
    InterfaceOptions.DefaultColumnColor(3) = Color.ButtonText
    InterfaceOptions.DefaultColumnColor(4) = Color.DisabledText
    InterfaceOptions.DefaultColumnColor(5) = Color.DisabledText
    InterfaceOptions.DefaultColumnColor(6) = Color.ButtonText
    InterfaceOptions.DefaultColumnColor(7) = Color.DisabledText
    InterfaceOptions.DefaultColumnColor(8) = Color.ButtonText
    InterfaceOptions.DefaultColumnColor(9) = Color.DisabledText
    InterfaceOptions.UserColumnColor = InterfaceOptions.DefaultColumnColor
    
    'Default Visible Columns
    ReDim InterfaceOptions.DefaultColumnVisible(1 To TimersListView.ColumnHeaders.Count)
    InterfaceOptions.DefaultColumnVisible(1) = 1
    InterfaceOptions.DefaultColumnVisible(2) = 1
    InterfaceOptions.DefaultColumnVisible(3) = 1
    InterfaceOptions.DefaultColumnVisible(4) = 1
    InterfaceOptions.DefaultColumnVisible(5) = 1
    InterfaceOptions.DefaultColumnVisible(6) = 1
    InterfaceOptions.DefaultColumnVisible(7) = 1
    InterfaceOptions.DefaultColumnVisible(8) = 1
    InterfaceOptions.DefaultColumnVisible(9) = 1
    InterfaceOptions.DefaultColumnVisible(10) = 1
    InterfaceOptions.UserColumnVisible = InterfaceOptions.DefaultColumnVisible
    
    'Set Default Column order
    ReDim InterfaceOptions.DefaultColumnOrder(TimersListView.ColumnHeaders.Count - 1)
    ReDim InterfaceOptions.UserColumnOrder(TimersListView.ColumnHeaders.Count - 1)
    For Count = 1 To TimersListView.ColumnHeaders.Count
        InterfaceOptions.DefaultColumnOrder(Count - 1) = Count - 1
    Next
    InterfaceOptions.UserColumnOrder = InterfaceOptions.DefaultColumnOrder
    
    'This is the Stats Listview control
    StatsView.ColumnHeaders.Add 1, , "Type", 800
    StatsView.ColumnHeaders.Add 2, , "MouseIdle", 2730, 0
    StatsView.ColumnHeaders.Add 3, , "AllTimersTotal", 2800, 0
    StatsView.ColumnHeaders.Add 4, , "TodayTotal", 2655, 0
    StatsView.ListItems.Add 1, , "Stats:"
    StatsView.ListItems(1).ListSubItems.Add 1
    StatsView.ListItems(1).ListSubItems.Add 2
    StatsView.ListItems(1).ListSubItems.Add 3
    Call ListItemColor(StatsView.ListItems(1), , Color.DisabledText, -1)
    
    'This is the bottom listview control for showing the detected window captions
    ReportView.ColumnHeaders.Add 1, , "Detected"
    ReportView.ColumnHeaders.Add 2, , "Handle", , 1 'aligned to the right
    ReportView.ColumnHeaders.Add 3, , "Caption"
    ReportView.ListItems.Add 1, , "DetectParentWindow"
    ReportView.ListItems.Add 2, , "DetectChildWindow"
    ReportView.ListItems(1).ListSubItems.Add 1, , , , "Windows API Handle number"
    ReportView.ListItems(1).ListSubItems.Add 2
    ReportView.ListItems(2).ListSubItems.Add 1
    ReportView.ListItems(2).ListSubItems.Add 2
    
    ReportView.ColumnHeaders(1).Width = InterfaceOptions.DefaultColWidth(1)
    ReportView.ColumnHeaders(2).Width = 930 'handle
    ReportView.ColumnHeaders(3).Width = 6279 'caption
    
End Sub

Private Sub Form_Resize()
    
    If (frmMainWindow.WindowState = 1 And IsWindowInTray = 0 And Check1(1).Value = 1) Then
        IsWindowInTray = MinimizeToTray(frmMainWindow)
    ElseIf (frmMainWindow.WindowState = 1 And IsWindowInTray = 1 And Check1(1).Value = 1) Then
        frmMainWindow.WindowState = 0
    ElseIf (IsWindowInTray = 1 And frmMainWindow.WindowState = 0) Then 'The form became visible while the tray icon is still present
        Call Form_MouseMove(0, 0, 513 * Screen.TwipsPerPixelX, 0) 'Delete the tray icon
        Exit Sub
    End If
    
    If (frmMainWindow.WindowState = 1 Or frmMainWindow.WindowState = 2) Then Exit Sub
    
    'index interfaceoptions.userwindowsize(1) is desiredwidth
    'index interfaceoptions.userwindowsize(2) is desiredheight
    'index interfaceoptions.userwindowsize(0) is previous width
    
    If (frmMainWindow.Width < InterfaceOptions.DefaultWindowSize(1)) Then
        frmMainWindow.Enabled = False
        frmMainWindow.Width = InterfaceOptions.DefaultWindowSize(1)
    ElseIf (frmMainWindow.Height < InterfaceOptions.DefaultWindowSize(2)) Then
        frmMainWindow.Enabled = False
        frmMainWindow.Height = InterfaceOptions.DefaultWindowSize(2)
    End If
    
    Call DoAdjustLayout
    If (InterfaceOptions.UserWindowSize(0) <> frmMainWindow.Width) Then InterfaceOptions.UserWindowSize(0) = frmMainWindow.Width 'The previous width
End Sub

Private Sub DoAdjustLayout()
    Dim Counter As Integer
    
    frmMainWindow.Enabled = True
    
    'Vertical Changes
    labels(1).Top = frmMainWindow.Height - 685
    labels(3).Top = frmMainWindow.Height - 685
    ReportView.Top = frmMainWindow.Height - 1185
    TimersListView.Height = frmMainWindow.Height - TimersListView.Top - 1495
    StatsView.Top = frmMainWindow.Height - 1500
    
    'Horizontal changes (text boxes ... etc)
    If (frmMainWindow.Width = InterfaceOptions.UserWindowSize(0)) Then Exit Sub 'so far this is the only place you are using the zero element of the userwindowsize array
    Text1(0).Left = frmMainWindow.Width - Text1(0).Width - Text1(1).Width - 14 * 15
    Text1(1).Left = frmMainWindow.Width - Text1(1).Width - 14 * 15
    labels(0).Left = Text1(0).Left - labels(0).Width - 5 * 15
    labels(3).Left = frmMainWindow.Width - 3015 - 240
    
    'Horizontal changes (the list view windows)
    TimersListView.Width = frmMainWindow.Width - 22 * 15
    ReportView.Width = TimersListView.Width
    StatsView.Width = TimersListView.Width
    
    'Thecolumns of the stats view
    'StatsView.ColumnHeaders(2).Width = (StatsView.Width - 870) / 3
    'StatsView.ColumnHeaders(3).Width = (StatsView.Width - 870) / 3
    'StatsView.ColumnHeaders(4).Width = (StatsView.Width - 870) / 3
    
    'The columns of the ReportView
    'None adjustable
    
    'The columns of the TimersListView (if the non-editables have changed, then readjust them)
    For Counter = 1 To TimersListView.ColumnHeaders.Count
        If (CLng(TimersListView.ColumnHeaders(Counter).Width) <> InterfaceOptions.DefaultColWidth(Counter) * InterfaceOptions.UserColumnVisible(Counter) And InterfaceOptions.DefaultIsColWidthEditable(Counter) = 0) Then TimersListView.ColumnHeaders(Counter).Width = InterfaceOptions.DefaultColWidth(Counter) * InterfaceOptions.UserColumnVisible(Counter)
    Next
End Sub

Private Sub GetDefaultWindowSizeAndPos()

    ReDim InterfaceOptions.DefaultWindowSize(4)
    ReDim InterfaceOptions.UserWindowSize(4)
    InterfaceOptions.DefaultWindowSize(1) = 9615 'width
    InterfaceOptions.DefaultWindowSize(2) = 4440 'Height
    InterfaceOptions.DefaultWindowSize(3) = (Screen.Width - InterfaceOptions.DefaultWindowSize(1)) / 2 'Left
    InterfaceOptions.DefaultWindowSize(4) = (Screen.Height - InterfaceOptions.DefaultWindowSize(2)) / 2 'Top
    InterfaceOptions.UserWindowSize = InterfaceOptions.DefaultWindowSize

End Sub

Private Sub RefreshTrayView()
    Dim Count As Integer
    
    frmTrayView.ListView1.ListItems.Clear
    
    For Count = 1 To TimersListView.ListItems.Count
        frmTrayView.ListView1.ListItems.Add Count, , TimersListView.ListItems(Count).Text
        frmTrayView.ListView1.ListItems(Count).ListSubItems.Add 1, , TimersListView.ListItems(Count).ListSubItems(6).Text
    Next
    
    'frmTrayView.ListView1.ListItems(TimersListView.SelectedItem.Index).EnsureVisible
End Sub

Private Sub CheckForMultipleInstances()
    Dim hWnd As Long
    Dim OriginalCaption As String

    'Is there a previous instance?
    If (App.PrevInstance = False) Then Exit Sub
    
    'Store the current caption of this form
    OriginalCaption = Caption
    
    'Change the caption to something else
    frmMainWindow.Caption = "Project Timers (multiple instance)"
    
    'Find the window
    hWnd = FindWindow(0&, OriginalCaption)
    
    'If the window is not found ... then Something is wrong, so leave
    If hWnd = 0 Then Exit Sub
    
    'If you have not exited so far then this is indeed a multiple instance situation
    SetForegroundWindow (hWnd)              'Activate the program
    ShowWindow hWnd, SW_SHOWNORMAL          'And restore the window
    
    End 'Ends Program
End Sub
