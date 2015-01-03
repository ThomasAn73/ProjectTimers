Attribute VB_Name = "mdlTimersFileIO"
Option Explicit

Const Delimiter = "¶"

Public Type AutoSave
    SessionON As Date
    SaveEvery As Integer
    SavedCount As Integer
    LastSavedON As Date
    AreTimersChanged As Boolean
End Type

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LVM_FIRST = &H1000
Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Public Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)

Dim TimerTypes(1) As String
Dim TimersIOPath As String
Dim PeriodicSave As AutoSave

Public Sub DoInitializeTimerTypes()
    TimerTypes(0) = "BSN"
    TimerTypes(1) = "HME"
End Sub

Public Function DoGetTimerTypes(Index) As String
    DoGetTimerTypes = TimerTypes(Index)
End Function

Public Sub DoInitDefaultDir()
    TimersIOPath = App.Path & "\ProjectTimers.ini"
End Sub

Public Sub IsChangedSinceLastSave(Condition As Boolean)
    PeriodicSave.AreTimersChanged = Condition
End Sub

Public Function DoGetAutoSaveData() As AutoSave
    DoGetAutoSaveData = PeriodicSave
End Function

Public Sub DoSetAutoSaveData(NewAutoSaveData As AutoSave)
    PeriodicSave = NewAutoSaveData
End Sub

Public Sub DoAutoSave(TmrDatabase As ProjectTimers, Optional Forced As Boolean = False)
    Dim CurrentDateTime As Date
    Dim SecSinceStart As Single
    Dim SecSinceLastSave As Single
    
    CurrentDateTime = Now
    
    SecSinceStart = (Int(CurrentDateTime) * Seconds.inday + Hour(CurrentDateTime) * Seconds.inHour + Minute(CurrentDateTime) * Seconds.inMin + Second(CurrentDateTime)) - (Int(PeriodicSave.SessionON) * Seconds.inday + Hour(PeriodicSave.SessionON) * Seconds.inHour + Minute(PeriodicSave.SessionON) * Seconds.inMin + Second(PeriodicSave.SessionON))
    SecSinceLastSave = (Int(CurrentDateTime) * Seconds.inday + Hour(CurrentDateTime) * Seconds.inHour + Minute(CurrentDateTime) * Seconds.inMin + Second(CurrentDateTime)) - (Int(PeriodicSave.LastSavedON) * Seconds.inday + Hour(PeriodicSave.LastSavedON) * Seconds.inHour + Minute(PeriodicSave.LastSavedON) * Seconds.inMin + Second(PeriodicSave.LastSavedON))
    
    If ((SecSinceLastSave >= PeriodicSave.SaveEvery And PeriodicSave.AreTimersChanged = True) Or Forced = True) Then
        PeriodicSave.SavedCount = PeriodicSave.SavedCount + 1
        PeriodicSave.LastSavedON = CurrentDateTime
        Call DoStoreSettings(TmrDatabase)
        Call DoSaveAllToUserDir(TmrDatabase)
        PeriodicSave.AreTimersChanged = False
    End If
    
End Sub

Public Sub DoStoreSettings(TmrDatabase As ProjectTimers)
    Dim OneFileLine As String
    Dim FileNum As Integer
    Dim Counter As Integer
    Dim AboutColumnOrder As Long
    
    'Get the current column order (if success then 'sendMessage' returns non zero (actually it returns 1)
    Call SendMessage(frmMainWindow.TimersListView.hWnd, LVM_GETCOLUMNORDERARRAY, CLng(UBound(InterfaceOptions.UserColumnOrder) + 1), InterfaceOptions.UserColumnOrder(0))
    
    'Get the current column widths (provided they are editable)
    For Counter = 1 To frmMainWindow.TimersListView.ColumnHeaders.Count
        If (frmMainWindow.TimersListView.ColumnHeaders(Counter).Width > 10 And InterfaceOptions.DefaultIsColWidthEditable(Counter) = 1) Then InterfaceOptions.UserColWidth(Counter) = CLng(frmMainWindow.TimersListView.ColumnHeaders(Counter).Width)
    Next
    
    'Get window size and pos
    If (IsWindowInTray = 0 And frmMainWindow.WindowState = 0) Then Call GetCurrentMainWindowSizeAndPos
    
    FileNum = FreeFile
    Open TimersIOPath For Output As FileNum
    
    'Save the options
    OneFileLine = "Section=1" & vbCrLf & "#Program Options" & vbCrLf
    OneFileLine = OneFileLine & TmrDatabase.SniffInterval & Delimiter & TmrDatabase.IdleThreshold & Delimiter
    For Counter = 0 To UBound(InterfaceOptions.UserColumnOrder) 'Save the column order
        If (Counter > 0) Then OneFileLine = OneFileLine & ","
        OneFileLine = OneFileLine & InterfaceOptions.UserColumnOrder(Counter)
    Next
    OneFileLine = OneFileLine & Delimiter & frmMainWindow.Check1(0) & Delimiter & frmMainWindow.Check1(1) & Delimiter
    OneFileLine = OneFileLine & Replace(CStr(TmrDatabase.LastMouseActive), ",", ".") & "," & Replace(CStr(TmrDatabase.MouseIdleToday), ",", ".") & "," & Replace(CStr(TmrDatabase.MouseBusyToday), ",", ".") & Delimiter
    For Counter = 0 To UBound(InterfaceOptions.UserColumnColor) 'Save the column Colors
        If (Counter > 0) Then OneFileLine = OneFileLine & ","
        OneFileLine = OneFileLine & InterfaceOptions.UserColumnColor(Counter)
    Next
    OneFileLine = OneFileLine & Delimiter
    For Counter = 1 To UBound(InterfaceOptions.UserColWidth) 'Save the column widths
        If (Counter > 1) Then OneFileLine = OneFileLine & ","
        OneFileLine = OneFileLine & InterfaceOptions.UserColWidth(Counter)
    Next
    OneFileLine = OneFileLine & Delimiter
    For Counter = 1 To UBound(InterfaceOptions.UserColumnVisible) 'Save the visible column flags
        If (Counter > 1) Then OneFileLine = OneFileLine & ","
        OneFileLine = OneFileLine & InterfaceOptions.UserColumnVisible(Counter)
    Next
    OneFileLine = OneFileLine & Delimiter & InterfaceOptions.UserWindowSize(1) & "," & InterfaceOptions.UserWindowSize(2) & "," & InterfaceOptions.UserWindowSize(3) & "," & InterfaceOptions.UserWindowSize(4)
    OneFileLine = OneFileLine & vbCrLf
       
    'Save the timers
    OneFileLine = OneFileLine & vbCrLf & "Section=2" & vbCrLf & "#Project Timers" & vbCrLf
    For Counter = 1 To TmrDatabase.MyTimers.Count
        OneFileLine = OneFileLine & TmrDatabase.MyTimers(Counter).Tag & Delimiter 'Serial
        OneFileLine = OneFileLine & TmrDatabase.MyTimers(Counter).Text & Delimiter 'Name
        OneFileLine = OneFileLine & Replace(TmrDatabase.MyTimers(Counter).ListSubItems(1).Tag, ",", ".") & Delimiter 'Time
        OneFileLine = OneFileLine & TmrDatabase.MyTimers(Counter).ListSubItems(3).Text & Delimiter 'Keywords
        OneFileLine = OneFileLine & TmrDatabase.MyTimers(Counter).ListSubItems(4).Text & Delimiter 'SaveInfo
        OneFileLine = OneFileLine & Replace(TmrDatabase.MyTimers(Counter).ListSubItems(5).Tag, ",", ".") & Delimiter 'Creation Date
        OneFileLine = OneFileLine & Replace(CStr(CDbl(CDate(TmrDatabase.MyTimers(Counter).ListSubItems(6).Text))), ",", ".") & Delimiter 'Time Lapsed on the day of last update
        OneFileLine = OneFileLine & Replace(TmrDatabase.MyTimers(Counter).ListSubItems(6).Tag, ",", ".") & Delimiter 'Last Update
        OneFileLine = OneFileLine & TmrDatabase.MyTimers(Counter).ListSubItems(7).Tag & Delimiter 'Timer Type
        OneFileLine = OneFileLine & TmrDatabase.MyTimers(Counter).ListSubItems(4).Tag & Delimiter 'Timer History Depth
        OneFileLine = OneFileLine & Replace(TmrDatabase.MyTimers(Counter).ListSubItems(8).Tag, ",", ".") & Delimiter 'TripCount
        OneFileLine = OneFileLine & Replace(TmrDatabase.MyTimers(Counter).ListSubItems(9).Tag, ",", ".") 'Last Trip Reset date
        OneFileLine = OneFileLine & vbCrLf
    Next
    
    'Add Timer History Section Header
    OneFileLine = OneFileLine & vbCrLf & "Section=3" & vbCrLf & "#Project Timer History" & vbCrLf
    
    'Save timer history
    For Counter = 1 To TmrDatabase.TimersHistory.Count
        OneFileLine = OneFileLine & TmrDatabase.TimersHistory.Item(Counter).LinkToTimerID & Delimiter
        OneFileLine = OneFileLine & TmrDatabase.TimersHistory.Item(Counter).OnThisDate & Delimiter
        OneFileLine = OneFileLine & Replace(CStr(TmrDatabase.TimersHistory.Item(Counter).RunningTotal), ",", ".") & Delimiter
        OneFileLine = OneFileLine & TmrDatabase.TimersHistory.Item(Counter).TotalToday
        If (Counter < TmrDatabase.TimersHistory.Count) Then OneFileLine = OneFileLine & vbCrLf
    Next
    
    Print #FileNum, OneFileLine 'You just printed the whole file as a single line
    Close #FileNum
End Sub

Public Sub DoRetreiveSettings(TmrDatabase As ProjectTimers)
    On Error Resume Next
    Err.Clear
    
    Dim FileNum As Integer
    Dim OneFileLine As String
    Dim Fragment As Variant
    Dim Section(3) As Integer '(0) holds the active section, (1) counts section 1, (2) counts section 2
    Dim HistoryItem As HistoryEntry
    Dim EndOfFile As Boolean
    
    Call DoInitializeTimerTypes
    FileNum = FreeFile
    Open TimersIOPath For Input As FileNum
    
    'Retrieve the timers
    Do While (Err.Number = 0)
        If (EOF(FileNum) = True) Then Exit Do
        Line Input #FileNum, OneFileLine
        
        'Detect comment or empty lines (and skip them) otherwise continue
        If (Len(Trim(OneFileLine)) > 0 And Left(Trim(OneFileLine), 1) <> "#") Then
                        
            Fragment = Split(OneFileLine, Delimiter)
            
            If (LCase(Replace(OneFileLine, " ", "")) = "section=1") Then
                Section(0) = 1 'Active section
            ElseIf (LCase(Replace(OneFileLine, " ", "")) = "section=2") Then
                Section(0) = 2 'Active section
            ElseIf (LCase(Replace(OneFileLine, " ", "")) = "section=3") Then
                Section(0) = 3 'Active section
            ElseIf (Section(0) = 1) Then 'Retrieve options
                
                Section(1) = Section(1) + 1 'increment section line counter
                
                If (UBound(Fragment) >= 0) Then TmrDatabase.SniffInterval = Int(Val(Fragment(0)))
                If (UBound(Fragment) >= 1) Then TmrDatabase.IdleThreshold = Int(Val(Fragment(1)))
                If (UBound(Fragment) >= 2) Then Call GetColumnSequenceFrom(CStr(Fragment(2)), ",")
                If (UBound(Fragment) >= 3) Then frmMainWindow.Check1(0).Value = Int(Val(Fragment(3)))
                If (UBound(Fragment) >= 4) Then frmMainWindow.Check1(1).Value = Int(Val(Fragment(4)))
                If (UBound(Fragment) >= 5) Then Call RetrieveMouseStats(CStr(Fragment(5)), ",")
                If (UBound(Fragment) >= 6) Then Call RetrieveUserColumnColors(CStr(Fragment(6)), ",")
                If (UBound(Fragment) >= 7) Then Call RetrieveUserColumnWidths(CStr(Fragment(7)), ",")
                If (UBound(Fragment) >= 8) Then Call RetrieveColumnsVisible(CStr(Fragment(8)), ",")
                If (UBound(Fragment) >= 9) Then Call RetrieveWindowSizeAndPos(CStr(Fragment(9)), ",")
                Call ApplyUserInterfaceChanges
                
            ElseIf (Section(0) = 2) Then 'retrieve timer
                
                Section(2) = Section(2) + 1 'This also acts as a counter for the number of timers found
                
                If (UBound(Fragment) >= 0) Then ' retrieve timer serial
                    TmrDatabase.MyTimers.Add Section(2)
                    Call SetTimerToDefault(TmrDatabase.MyTimers, CLng(Section(2)))
                    TmrDatabase.MyTimers(Section(2)).Tag = CLng(Round(Val(Fragment(0)), 0))
                End If
                If (UBound(Fragment) >= 1) Then TmrDatabase.MyTimers(Section(2)).Text = Fragment(1) 'Timer Name
                If (UBound(Fragment) >= 2) Then 'Total Time
                    If (IsDate(Fragment(2))) Then TmrDatabase.MyTimers(Section(2)).ListSubItems(1).Tag = CDbl(CDate(Fragment(2))) Else TmrDatabase.MyTimers(Section(2)).ListSubItems(1).Tag = CDbl(Abs(Val(Fragment(2))))
                End If
                TmrDatabase.MyTimers(Section(2)).ListSubItems(1).Text = DoShowTime(TmrDatabase.MyTimers(Section(2)).ListSubItems(1).Tag)
                If (UBound(Fragment) >= 3) Then TmrDatabase.MyTimers(Section(2)).ListSubItems(3).Text = Fragment(3) 'Keywords
                If (UBound(Fragment) >= 4) Then TmrDatabase.MyTimers(Section(2)).ListSubItems(4).Text = Fragment(4) 'SaveIn directory
                If (UBound(Fragment) >= 5) Then 'Creation Date
                    If (IsDate(Fragment(5))) Then
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(5).Text = Format(Fragment(5), "yyyy/mm/dd")
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(5).Tag = CDbl(CDate(Fragment(5)))
                    Else
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(5).Text = Format(CDate(CDbl(Abs(Val(Fragment(5))))), "yyyy/mm/dd")
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(5).Tag = CDbl(Abs(Val(Fragment(5))))
                    End If
                End If
                If (UBound(Fragment) >= 6) Then 'Time lapsed today
                    If (IsDate(Fragment(6))) Then TmrDatabase.MyTimers(Section(2)).ListSubItems(6).Text = Format(Fragment(6), "hh:mm:ss") Else TmrDatabase.MyTimers(Section(2)).ListSubItems(6).Text = Format(CDate(CDbl(Abs(Val(Fragment(6))))), "hh:mm:ss")
                End If
                If (UBound(Fragment) >= 7) Then 'Date of last timer activity
                    If (IsDate(Fragment(7))) Then TmrDatabase.MyTimers(Section(2)).ListSubItems(6).Tag = CDbl(CDate(Fragment(7))) Else TmrDatabase.MyTimers(Section(2)).ListSubItems(6).Tag = CDbl(Abs(Val(Fragment(7))))
                End If
                If (UBound(Fragment) >= 8) Then 'Type of timer
                    If (Int(Val(Fragment(8)) <= 1)) Then
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(7).Text = TimerTypes(Int(Val(Fragment(8))))
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(7).Tag = Int(Val(Fragment(8)))
                    End If
                End If
                If (UBound(Fragment) >= 9) Then TmrDatabase.MyTimers(Section(2)).ListSubItems(4).Tag = Int(Abs(Val(Fragment(9)))) 'Timer History Depth
                If (UBound(Fragment) >= 10) Then 'Trip Counter
                    If (IsDate(Fragment(10))) Then TmrDatabase.MyTimers(Section(2)).ListSubItems(8).Tag = CDbl(CDate(Fragment(10))) Else TmrDatabase.MyTimers(Section(2)).ListSubItems(8).Tag = CDbl(Abs(Val(Fragment(10))))
                End If
                TmrDatabase.MyTimers(Section(2)).ListSubItems(8).Text = DoShowTime(TmrDatabase.MyTimers(Section(2)).ListSubItems(8).Tag)
                If (UBound(Fragment) >= 11) Then 'Last Trip Reset Date
                    If (IsDate(Fragment(11))) Then
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(9).Text = Format(Fragment(11), "yyyy/mm/dd HH:MM")
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(9).Tag = CDbl(CDate(Fragment(11)))
                    Else
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(9).Text = Format(CDate(CDbl(Abs(Val(Fragment(11))))), "yyyy/mm/dd HH:MM")
                        TmrDatabase.MyTimers(Section(2)).ListSubItems(9).Tag = CDbl(Abs(Val(Fragment(11))))
                    End If
                End If
                       
                'Highlight what was just added
                TmrDatabase.MyTimers(Section(2)).Selected = True
                
                'Custom color
                Call ListItemColor(TmrDatabase.MyTimers(Section(2)), , , , InterfaceOptions.UserColumnColor)
            
            ElseIf (Section(0) = 3) Then 'retrieve timer History
                Section(3) = Section(3) + 1
                Set HistoryItem = New HistoryEntry
                If (UBound(Fragment) >= 0) Then HistoryItem.LinkToTimerID = CLng(Val(Fragment(0)))
                If (UBound(Fragment) >= 1) Then HistoryItem.OnThisDate = Fragment(1)
                If (UBound(Fragment) >= 2) Then HistoryItem.RunningTotal = CDbl(Abs(Val(Fragment(2))))
                If (UBound(Fragment) >= 3) Then HistoryItem.TotalToday = Fragment(3)
                Call AddHistoryItem(TmrDatabase.TimersHistory, HistoryItem)
            End If
        End If
    Loop
    Close #FileNum
    
    Call UpdateTimerStatistics(TmrDatabase)
    
    Call UpdateTimersContextMenu
    
    If (Err.Number <> 0) Then MsgBox Err.Description & " while attempting to read the ini file" & vbCr & "ErrCode: " & Err.Number, vbOKOnly, "Error"
End Sub

Public Sub DoSaveAllToUserDir(TmrDatabase As ProjectTimers, Optional SurpressSuccDlgs As Boolean = True, Optional SurpressErrDlgs As Boolean = False)
    Dim Counter As Integer
    'Save the timers themselves
    For Counter = 1 To TmrDatabase.MyTimers.Count
        Call DoSaveOneTimer(TmrDatabase, CLng(Counter), SurpressSuccDlgs, SurpressErrDlgs)
    Next
End Sub

Public Sub DoSaveOneTimer(TmrDatabase As ProjectTimers, ThisTimer As Long, Optional SurpressDlg As Boolean = True, Optional SurpressErrDlg As Boolean = False)
    Dim OneFileLine As String
    Dim FileNum As Integer
    Dim OnePath As String
    Dim ProbeResult(1) As Variant
    Dim ToCheck As String
    Dim OneAttr As Integer
    Dim Directories As Variant
    Dim Counter As Long
    Dim ErrReport As String
    
    On Error Resume Next
    Err.Clear
    
    OnePath = TmrDatabase.MyTimers(ThisTimer).ListSubItems(4).Text
    Directories = Split(OnePath, "\")
    
    If (UBound(Directories) < 1) Then Exit Sub
    
    'Get the directory without the filename (basically add back up all elelments of the "Directories" array except the last one)
    ToCheck = Left(OnePath, Len(OnePath) - Len(Directories(UBound(Directories))))
    OneAttr = GetAttr(ToCheck)
    If (Err.Number = 53) Then Err.Clear
    ProbeResult(0) = Array(OneAttr, GetAttrDescription(OneAttr), Err.Number, Err.Description, ToCheck)
    Err.Clear
    
    'Check the path as a whole
    OneAttr = GetAttr(OnePath)
    ProbeResult(1) = Array(OneAttr, GetAttrDescription(OneAttr), Err.Number, Err.Description, OnePath)
    Err.Clear 'error code 76 is for path not found (53 is for file not found)
    
    'Debug.Print ProbeResult(0)(0) & " " & ProbeResult(0)(2) & " " & ProbeResult(1)(0) & " " & ProbeResult(1)(2)
    If (ProbeResult(0)(2) <> 76 And ProbeResult(1)(2) <> 76 And (CInt(ProbeResult(0)(0)) And (2 + 16 + 32 + 64 + 128 + 256 + 512 + 1024 + 2048 + 8192 + 16384)) = CInt(ProbeResult(0)(0)) And (CInt(ProbeResult(1)(0)) And (2 + 16 + 32 + 64 + 128 + 256 + 512 + 1024 + 2048 + 8192 + 16384)) = CInt(ProbeResult(1)(0))) Then
        FileNum = FreeFile
        Open OnePath For Output As FileNum
        
        'Main timer stats
        OneFileLine = "Timer Name = " & TmrDatabase.MyTimers(ThisTimer).Text & vbCrLf
        OneFileLine = OneFileLine & "Total time = " & TmrDatabase.MyTimers(ThisTimer).ListSubItems(1).Text & " <hhh:mm:ss>" & vbCrLf 'Time
        OneFileLine = OneFileLine & "Keywords used = " & TmrDatabase.MyTimers(ThisTimer).ListSubItems(3).Text & vbCrLf 'Keywords
        OneFileLine = OneFileLine & "Timer Creation Date = " & Format(CDate(TmrDatabase.MyTimers(ThisTimer).ListSubItems(5).Tag), "yyyy/mm/dd hh:mm AM/PM") & vbCrLf 'Creation Date
        OneFileLine = OneFileLine & "Last Updated = " & Format(CDate(TmrDatabase.MyTimers(ThisTimer).ListSubItems(6).Tag), "yyyy/mm/dd hh:mm AM/PM") & vbCrLf 'Date of last timer activity
        OneFileLine = OneFileLine & "Type = " & TmrDatabase.MyTimers(ThisTimer).ListSubItems(7).Text 'Timer Type
        OneFileLine = OneFileLine & vbCrLf
        
        'Related Timer History
        OneFileLine = OneFileLine & vbCrLf & "Timer History (max depth = " & TmrDatabase.MyTimers(ThisTimer).ListSubItems(4).Tag & " active days)" & vbCrLf
        OneFileLine = OneFileLine & "Format (comma separated): Date, Running Total <hhh:mm:ss>, Day Total <hh:mm:ss>" & vbCrLf
        For Counter = 1 To TmrDatabase.TimersHistory.Count
            If (CLng(TmrDatabase.TimersHistory.Item(Counter).LinkToTimerID) = CLng(TmrDatabase.MyTimers(ThisTimer).Tag)) Then
                OneFileLine = OneFileLine & Format(TmrDatabase.TimersHistory.Item(Counter).OnThisDate, "yyyy/mm/dd") & ", "
                OneFileLine = OneFileLine & DoShowTime(TmrDatabase.TimersHistory.Item(Counter).RunningTotal) & ", "
                OneFileLine = OneFileLine & TmrDatabase.TimersHistory.Item(Counter).TotalToday
                OneFileLine = OneFileLine & vbCrLf
            End If
        Next
        
        Print #FileNum, OneFileLine
        Close #FileNum
        
        If SurpressDlg = False Then MsgBox "Timer <" & TmrDatabase.MyTimers(ThisTimer).Text & "> has been saved", vbOKOnly, "Complete"
    Else
        ErrReport = "Check1 = Attr(" & ProbeResult(0)(0) & " " & ProbeResult(0)(1) & "), Err(" & ProbeResult(0)(2) & " " & ProbeResult(0)(3) & "). Dir =" & ProbeResult(0)(4) & vbCrLf
        ErrReport = ErrReport & "Check2 = Attr(" & ProbeResult(1)(0) & " " & ProbeResult(1)(1) & "), Err(" & ProbeResult(1)(2) & " " & ProbeResult(1)(3) & ")"
        If (SurpressErrDlg = False) Then MsgBox "Timestamp for timer <" & TmrDatabase.MyTimers(ThisTimer).Text & "> was not saved." & vbCrLf & vbCrLf & ErrReport, vbOKOnly, "Error while saving a timestamp"
        'consider using a visual indicator showing in realtime which paths are invalid
    End If
End Sub

'Receives a delimited string and checks if it is a valid column sequence (compared to a template default sequence)
Private Sub GetColumnSequenceFrom(ThisSequence As String, Separator As String)
    Dim Fragment2 As Variant
    Dim Parity() As Integer
    Dim Count As Integer

    If (ThisSequence = "" Or Separator = "") Then Exit Sub

    Fragment2 = Split(ThisSequence, Separator)
    If (UBound(Fragment2) > UBound(InterfaceOptions.DefaultColumnOrder)) Then Exit Sub
    
    ReDim Parity(UBound(Fragment2))
    
    'Use the array index to store how many times that digit appears in the list
    For Count = 0 To UBound(Fragment2)
        If (CLng(Abs(Val(Fragment2(Count)))) <= UBound(Fragment2) And CLng(Abs(Val(Fragment2(Count)))) >= 0) Then Parity(CLng(Abs(Val(Fragment2(Count))))) = Parity(CLng(Abs(Val(Fragment2(Count))))) + 1
    Next
    
    'Check the parity array. All digits from 0 to ubound(fragment2) must appear only once
    For Count = 0 To UBound(Fragment2)
        If (Parity(Count) <> 1) Then Exit Sub
    Next
    
    'Since the code did not exit earlier, we are clear to continue
    'Add the new data into the userarray
    For Count = 0 To UBound(Fragment2)
        InterfaceOptions.UserColumnOrder(Count) = CLng(Abs(Val(Fragment2(Count))))
    Next

End Sub

Private Sub RetrieveMouseStats(ThisString As String, Delim As String)
    Dim Fragment As Variant
    Dim LastActive As Double
    
    If (ThisString = "") Then Exit Sub
    
    Fragment = Split(ThisString, Delim)
    
    If (UBound(Fragment) >= 0) Then
        If (CStr(Abs(Val(Fragment(0)))) = Fragment(0)) Then LastActive = CDbl(Abs(Val(Fragment(0))))
        If (Year(LastActive) <> Year(Now) And Month(LastActive) <> Month(Now) And Day(LastActive) <> Day(Now)) Then Exit Sub
        TmrDatabase.LastMouseActive = LastActive
    End If
    If (UBound(Fragment) >= 1) Then
        If (CStr(Abs(Val(Fragment(1)))) = Fragment(1)) Then TmrDatabase.MouseIdleToday = CDbl(Abs(Val(Fragment(1))))
    End If
    If (UBound(Fragment) >= 2) Then
        If (CStr(Abs(Val(Fragment(2)))) = Fragment(2)) Then TmrDatabase.MouseBusyToday = CDbl(Abs(Val(Fragment(2))))
    End If
End Sub

Private Sub RetrieveUserColumnColors(ThisString As String, Delim As String)
    Dim Fragment As Variant
    Dim Counter As Integer
    
    If (ThisString = "") Then Exit Sub
    Fragment = Split(ThisString, Delim)
    
    For Counter = 0 To UBound(Fragment)
        If (Counter <= UBound(InterfaceOptions.UserColumnColor)) Then InterfaceOptions.UserColumnColor(Counter) = CLng(Val(Fragment(Counter)))
    Next
    
End Sub

Private Sub RetrieveUserColumnWidths(ThisString As String, Delim As String)
    Dim Fragment As Variant
    Dim Counter As Integer
    
    If (ThisString = "") Then Exit Sub
    Fragment = Split(ThisString, Delim)
    
    For Counter = 1 To UBound(Fragment) + 1
        If (Counter <= UBound(InterfaceOptions.UserColWidth) And Abs(CLng(Val(Fragment(Counter - 1)))) > 0 And InterfaceOptions.DefaultIsColWidthEditable(Counter) = 1) Then InterfaceOptions.UserColWidth(Counter) = Abs(CLng(Val(Fragment(Counter - 1))))
    Next
    
End Sub

Private Sub RetrieveColumnsVisible(ThisString As String, Delim As String)
    Dim Fragment As Variant
    Dim Counter As Integer
    
    If (ThisString = "") Then Exit Sub
    Fragment = Split(ThisString, Delim)
    
    For Counter = 1 To UBound(Fragment) + 1
        If (Counter <= UBound(InterfaceOptions.UserColumnVisible) And Cbool10(Fragment(Counter - 1)) = 0) Then InterfaceOptions.UserColumnVisible(Counter) = 0
    Next
    
End Sub

Private Sub RetrieveWindowSizeAndPos(ThisString As String, Delim As String)
    Dim Fragment As Variant
    Dim Counter As Integer
    Dim MalformedXY As Boolean
    
    If (ThisString = "") Then Exit Sub
    Fragment = Split(ThisString, Delim)
    
    If (UBound(Fragment) >= 0) Then
        'Read width
        If (Abs(Fix(Val(Fragment(0)))) > InterfaceOptions.DefaultWindowSize(1) And Abs(Fix(Val(Fragment(0)))) < (Screen.Width)) Then InterfaceOptions.UserWindowSize(1) = Abs(Fix(Val(Fragment(0))))
    End If
    If (UBound(Fragment) >= 1) Then
        'Read height
        If (Abs(Fix(Val(Fragment(1)))) > InterfaceOptions.DefaultWindowSize(2) And (Screen.Height)) Then InterfaceOptions.UserWindowSize(2) = Abs(Fix(Val(Fragment(1))))
    End If
    
    'caution: screen.witdh is also in twips (do not convert). The form.left property is in twips too
    If (UBound(Fragment) >= 2) Then
        'Read the X value for the window left
        If (Abs(Fix(Val(Fragment(2)))) >= 0 And Abs(Fix(Val(Fragment(2)))) < (Screen.Width - 100)) Then InterfaceOptions.UserWindowSize(3) = Abs(Fix(Val(Fragment(2)))) Else MalformedXY = True
    End If
    If (UBound(Fragment) >= 3) Then
        'Read the Y value for the window top
        If (Abs(Fix(Val(Fragment(3)))) >= 0 And Abs(Fix(Val(Fragment(3)))) < (Screen.Height - 100)) Then InterfaceOptions.UserWindowSize(4) = Abs(Fix(Val(Fragment(3)))) Else MalformedXY = True
    End If
    
    If (MalformedXY = True) Then
        InterfaceOptions.UserWindowSize(3) = (Screen.Width - InterfaceOptions.UserWindowSize(1)) / 2
        InterfaceOptions.UserWindowSize(4) = (Screen.Height - InterfaceOptions.UserWindowSize(2)) / 2
    End If
End Sub

Public Sub GetCurrentMainWindowSizeAndPos()
    InterfaceOptions.UserWindowSize(1) = frmMainWindow.Width
    InterfaceOptions.UserWindowSize(2) = frmMainWindow.Height
    InterfaceOptions.UserWindowSize(3) = frmMainWindow.Left
    InterfaceOptions.UserWindowSize(4) = frmMainWindow.Top
End Sub

Private Sub ApplyUserInterfaceChanges()
    Dim Count As Integer
    
    'RearrangeColumns
    Call SendMessage(frmMainWindow.TimersListView.hWnd, LVM_SETCOLUMNORDERARRAY, CLng(UBound(InterfaceOptions.UserColumnOrder) + 1), InterfaceOptions.UserColumnOrder(0))
    
    'Adjust column widths and hidden columns
    For Count = 1 To UBound(InterfaceOptions.UserColWidth)
        frmMainWindow.TimersListView.ColumnHeaders(Count).Width = InterfaceOptions.UserColWidth(Count) * InterfaceOptions.UserColumnVisible(Count)
    Next
    
    'Adjust window size
    frmMainWindow.Width = InterfaceOptions.UserWindowSize(1)
    frmMainWindow.Height = InterfaceOptions.UserWindowSize(2)
    frmMainWindow.Left = InterfaceOptions.UserWindowSize(3)
    frmMainWindow.Top = InterfaceOptions.UserWindowSize(4)
    
End Sub

'Converts a value to a 0 = false or 1 = true (anything other than zero = true)
Public Function Cbool10(thisval As Variant, Optional Invert As Boolean = False) As Integer
    If (Invert = False) Then
        If (Val(thisval) <> 0) Then Cbool10 = 1 Else Cbool10 = 0
    Else
        If (Val(thisval) <> 0) Then Cbool10 = 0 Else Cbool10 = 1
    End If
End Function

Public Function GetAttrDescription(ByVal AttrCode As Integer) As String
    Dim VBAttrDescr As Variant
    Dim Count As Integer
    Dim attrCount As Integer
    
    'The file attributes are represented by a 14 digit binary number: 11111111111111
    'Numerical values: 0,1,2,4,8,16,32,64,128,256,512,1024,2048,4096,8192,16384
    VBAttrDescr = Array("Normal", "Read-only", "Hidden", "System", "?Attr8", "Folder", "Archive-Ready", "Device", "Normal", "Temporary", "Sparse-file", "Reparse-Point", "Compressed", "Offline", "Not-Content-indexed", "Encrypted")
    
    If (AttrCode = 0) Then GetAttrDescription = VBAttrDescr(0)(1): Exit Function
    
    For Count = 14 To 0 Step -1
        If ((2 ^ Count And AttrCode) = 2 ^ Count) Then
            attrCount = attrCount + 1
            If (attrCount > 1) Then GetAttrDescription = GetAttrDescription & ", "
            GetAttrDescription = GetAttrDescription & VBAttrDescr(Count + 1)
        End If
    Next
End Function
