Attribute VB_Name = "mdlUpdateTimers"
Option Explicit

Enum Seconds
    inMin = 60
    inHour = 3600
    inday = 86400
End Enum

Public Enum Color
    red = &HDC&
    DisabledText = &H80000011
    ButtonText = &H80000012
End Enum

Public Type TimerChangeFlag
    TimerIndex As Long
    IsNewHistoryDepth As Boolean
End Type

Public Type Interface
    UserColumnOrder() As Long 'zero based array
    UserColWidth() As Long
    UserColumnColor() As Long 'zero based array
    UserColumnVisible() As Integer
    UserWindowSize() As Long 'zero based. (1)=width, (2)=height, (3)=Left, (4)=Top, (0) is left empty
    DefaultColumnOrder() As Long 'zero based array
    DefaultColWidth() As Long
    DefaultIsColWidthEditable() As Integer
    DefaultColumnColor() As Long 'zero based array
    DefaultColumnVisible() As Integer
    DefaultWindowSize() As Long 'zero based. (1)=width, (2)=height, (3)=Left, (4)=Top, (0) is left empty
End Type

Public Type ProjectTimers
    SniffInterval As Single '(in seconds) a user option that sets how often the timer is triggered
    IdleThreshold As Single '(in seconds) a user option to not update user timers if cursor is idle beyond this value
    TodayTotal As Double
    AllTimeTotal As Double
    MouseIdleToday As Double
    MouseBusyToday As Double
    LastMouseActive As Double
    MyTimers As ListItems 'a collection
    TimersHistory As Collection 'Use collections because when you delete a timer it will be easier to wipe its history entries than if you used an array
End Type

Public TmrDatabase As ProjectTimers 'This is the master container for all user timers
Public TimerEdited As TimerChangeFlag
Public InterfaceOptions As Interface

'For each timer get the user keywords
'search for the user keywords in the active window caption
'if keyword match is found, then increment that particular timer by "sniffInterval"
'Returns: (0 for no change), (1 for new timer activated), (2 for changed existing active timers)
Public Function DoUpdateTimers(Timers As ListItems, TimersHistory As Collection, Interval As Single, ActiveWindow As WindowData) As Integer
    Dim Counter As Integer
    Dim StatusChanges As Integer
    Dim Updates As Integer
    
    DoUpdateTimers = 0
    
    If Timers.Count = 0 Then Exit Function
    'Parse user timers one by one
    For Counter = 1 To Timers.Count
    
        'See if today is a different day than last time the timer was updated
        Call ResetIfNewDay(Timers(Counter))
        
        If (DoFindMatch(Timers(Counter).ListSubItems(3).Text, ActiveWindow)) Then
            'Update the "All time" counter for this timer
            Timers(Counter).ListSubItems(1).Tag = Timers(Counter).ListSubItems(1).Tag + Interval / Seconds.inday
            Timers(Counter).ListSubItems(1).Text = DoShowTime(Timers(Counter).ListSubItems(1).Tag)
            
            'Update the "Today" counter for this timer
            Timers(Counter).ListSubItems(6).Text = DoShowTime(CDbl(CDate(Timers(Counter).ListSubItems(6).Text) + Interval / Seconds.inday), , 2)
            Timers(Counter).ListSubItems(6).Tag = CDbl(Now)
            
            'Update the Trip Counter
            Timers(Counter).ListSubItems(8).Tag = Timers(Counter).ListSubItems(8).Tag + Interval / Seconds.inday
            Timers(Counter).ListSubItems(8).Text = DoShowTime(Timers(Counter).ListSubItems(8).Tag)
            
            'Update History
            Call UpdateTimerHistory(Timers(Counter), TimersHistory)
            
            'Update the activity indicator
            If (Timers(Counter).ListSubItems(2).Text <> "<" Or Timers(Counter).ListSubItems(2).ForeColor <> Color.red) Then
                Timers(Counter).ListSubItems(2).Text = "<" 'Activity indicator
                Call ListItemColor(Timers(Counter), , Color.red, 0)
                StatusChanges = StatusChanges + 1
            Else
                Updates = Updates + 1
            End If
            Call IsChangedSinceLastSave(True)
        Else
            If (Timers(Counter).ListSubItems(2).Text <> "") Then
                Timers(Counter).ListSubItems(2).Text = "" 'Activity indicator
                Call ListItemColor(Timers(Counter), , , , InterfaceOptions.UserColumnColor)
                StatusChanges = StatusChanges + 1
            End If
        End If
        
    Next
    
    If (StatusChanges > 0) Then
        DoUpdateTimers = 1
    ElseIf (Updates > 0) Then
        DoUpdateTimers = 2
    End If

End Function

Public Function CoolDown(Timers As ListItems) As Integer
    Dim Counter As Integer
    
    CoolDown = 0
    For Counter = 1 To Timers.Count
        Call ResetIfNewDay(Timers(Counter))
        If (Timers(Counter).ListSubItems(2).Text <> "" And Timers(Counter).ListSubItems(2).ForeColor = Color.red) Then
            Timers(Counter).ListSubItems(2).Text = "x"
            Call ListItemColor(Timers(Counter), , , , InterfaceOptions.UserColumnColor)
            CoolDown = 1
        End If
    Next
    
End Function

Public Sub ResetIfNewDay(ThisTimer As ListItem)
    If (ThisTimer.ListSubItems(6).Text = "00:00:00") Then Exit Sub
    If (Month(Now) <> Month(ThisTimer.ListSubItems(6).Tag) Or Year(Now) <> Year(ThisTimer.ListSubItems(6).Tag) Or Day(Now) <> Day(ThisTimer.ListSubItems(6).Tag)) Then
        ThisTimer.ListSubItems(6).Text = "00:00:00"
        ThisTimer.ListSubItems(6).Tag = CDbl(Now)
    End If
End Sub

'Takes the tag value (that holds the timer activity) and converts it to a visible "hhh:mm:ss" format in the listsubitem text field
Public Function DoShowTime(DecimalDays As Double, Optional Expanding As Boolean = False, Optional HourDigits As Integer = 3, Optional Descriptions As Boolean = False, Optional ShowDays As Boolean = False) As String
    Dim TotalTime, RemainingSecs As Double
    Dim TimeDays, TimeHrs, TimeMins As Double
    Dim DisplayTime As String
    Dim DescrText(3) As String
    
    TotalTime = DecimalDays
    RemainingSecs = Round(TotalTime * Seconds.inday, 0)
    TimeDays = Fix(RemainingSecs / Seconds.inday)
    RemainingSecs = RemainingSecs Mod Seconds.inday
    TimeHrs = Fix(RemainingSecs / Seconds.inHour)
    RemainingSecs = RemainingSecs Mod Seconds.inHour
    TimeMins = Fix(RemainingSecs / Seconds.inMin)
    RemainingSecs = RemainingSecs Mod Seconds.inMin
    
    If (Descriptions = True) Then
        DescrText(0) = "d"
        DescrText(1) = "h"
        DescrText(2) = "m"
        DescrText(3) = "sec"
    End If
    
    If ((ShowDays = True And Expanding = False) Or (ShowDays = True And TimeDays > 0 And HourDigits < 3)) Then DisplayTime = DisplayTime & Format(TimeDays, "00") & DescrText(0) & ":" Else TimeHrs = TimeHrs + TimeDays * 24
    If (TimeHrs > 0 Or Expanding = False) Then DisplayTime = DisplayTime & Format(TimeHrs, String(CLng(HourDigits), "0")) & DescrText(1) & ":"
    If (TimeMins > 0 Or Expanding = False) Then DisplayTime = DisplayTime & Format(TimeMins, "00") & DescrText(2) & ":"
    DisplayTime = DisplayTime & Format(RemainingSecs, "00") & DescrText(3)
    DoShowTime = DisplayTime
End Function

Private Function DoFindMatch(ToTest As String, TheTextIn As WindowData) As Boolean
    Dim Count As Integer
    Dim count2 As Integer
    Dim Fragment As Variant
    Dim Fragment2 As Variant
    Dim FoundCount As Integer
    
    DoFindMatch = False
    FoundCount = 0
    
    If ((Len(ToTest) = 0 Or IsNull(ToTest)) Or (Len(TheTextIn.ParentCaption) = 0 Or IsNull(TheTextIn.ParentCaption)) And (Len(TheTextIn.ChildCaption) = 0 Or IsNull(TheTextIn.ChildCaption))) Then Exit Function
    If (ToTest = "*") Then DoFindMatch = True: Exit Function
    
    Fragment = Split(ToTest, ",") 'First coarse split
    For Count = 0 To UBound(Fragment)
        Fragment2 = Split(Fragment(Count), " ") 'Finer split
        For count2 = 0 To UBound(Fragment2)
            If (InStr(1, TheTextIn.ParentCaption, Trim(Fragment2(count2)), vbTextCompare) Or InStr(1, TheTextIn.ChildCaption, Trim(Fragment2(count2)), vbTextCompare)) Then FoundCount = FoundCount + 1
        Next
        If (FoundCount = (UBound(Fragment2) + 1)) Then DoFindMatch = True Else FoundCount = 0
    Next
    
End Function

'Changes color of Listitem line of a listview control (affects the entire line, or just the headers, or just the subitems, depending on the Columns value)
Public Sub ListItemColor(ThisListitem As ListItem, Optional isBold As Boolean = False, Optional textForeColor As Long = Color.ButtonText, Optional Columns As Integer = 0, Optional TheseColors As Variant) 'As Variant
    Dim Count As Integer
    
    If (IsArray(TheseColors)) Then
        If (UBound(TheseColors) >= ThisListitem.ListSubItems.Count) Then
            ThisListitem.ForeColor = TheseColors(0)
            For Count = 1 To ThisListitem.ListSubItems.Count
                ThisListitem.ListSubItems(Count).ForeColor = TheseColors(Count)
            Next
        End If
    Else
        If (Columns = 0 Or Columns = 1) Then 'All, or Headers only
            ThisListitem.Bold = isBold
            ThisListitem.ForeColor = textForeColor
        End If
        
        If (Columns = 0 Or Columns = -1) Then 'All, or Subitems Only
            For Count = 1 To ThisListitem.ListSubItems.Count
                ThisListitem.ListSubItems(Count).Bold = isBold
                ThisListitem.ListSubItems(Count).ForeColor = textForeColor
            Next
        End If
        
        If (Columns > 1 And Columns < ThisListitem.ListSubItems.Count + 1) Then
            ThisListitem.ListSubItems(Columns).Bold = isBold
            ThisListitem.ListSubItems(Columns).ForeColor = textForeColor
        End If
    End If
    
End Sub

'Receives a listview object and finds the total of a specified column
'Returns an array. First element is the result. Second element is any related error codes
Public Function GetTotal(TheseItems As ListItems, ThisColumn As Integer, Optional UseTags As Boolean = False, Optional IsDate As Boolean = False) As Variant
    Dim Count As Integer
    Dim RunningTotal As Double
    Dim Value As Double
    
    GetTotal = Array(0, -1)
    
    If (TheseItems.Count = 0) Then Exit Function
    If (TheseItems(1).ListSubItems.Count < ThisColumn) Then Exit Function
    
    For Count = 1 To TheseItems.Count
        Select Case IsDate
            Case True 'it is a duration in the form of double
                If (ThisColumn = 1 And UseTags = False) Then Value = CDbl(CDate(TheseItems(Count).Text))
                If (ThisColumn = 1 And UseTags = True) Then Value = TheseItems(Count).Tag 'This holds a serial number in this case (doubt you will ever need to sum them, but it is here for the shake of completeness)
                If (ThisColumn > 1 And UseTags = False) Then Value = CDbl(CDate(TheseItems(Count).ListSubItems(ThisColumn - 1).Text))
                If (ThisColumn > 1 And UseTags = True) Then Value = TheseItems(Count).ListSubItems(ThisColumn - 1).Tag
                
                RunningTotal = RunningTotal + Value
                
            Case False 'it is just a number
        End Select
    Next
    GetTotal = Array(RunningTotal, 0)
End Function

'This is an older function (now obsolete)
Public Function ConvDurationToString(Duration As Double) As String
    Dim FormatedIdleTime As String
    'Date is a double that represents number of days (0.12345 represents 2:57:46 am). The fractional part of 'date'is the time. The integer part is the number of days
    If (Int(Duration) > 0) Then
        FormatedIdleTime = Format(Int(Duration), "00") & "d:" & Format(Duration, "hh:mm:ss")
    ElseIf (Round(Duration * Seconds.inday) < Seconds.inMin) Then
        FormatedIdleTime = Format(Second(Duration), "00") & "sec"
    ElseIf (Round(Duration * Seconds.inday) < Seconds.inHour) Then
        FormatedIdleTime = Format(Minute(Duration), "00") & "m:" & Format(Second(Duration), "00") & "sec"
    Else
        FormatedIdleTime = Format(Hour(Duration), "00") & "h:" & Format(Minute(Duration), "00") & "m:" & Format(Second(Duration), "00") & "sec"
    End If
    ConvDurationToString = FormatedIdleTime
End Function

Public Sub SetTimerToDefault(InTheseTimers As ListItems, ThisOne As Long)
    
    InTheseTimers(ThisOne).Text = "Untitled" 'Timer Title
    InTheseTimers(ThisOne).Tag = GetAvailableSerial(InTheseTimers)
    InTheseTimers(ThisOne).ListSubItems.Add 1, , "000:00:00", , "hhh:mm:ss" 'Total Time
    InTheseTimers(ThisOne).ListSubItems(1).Tag = CDbl(0)
    InTheseTimers(ThisOne).ListSubItems.Add 2, , "" 'Activity Indicator
    InTheseTimers(ThisOne).ListSubItems.Add 3, , "" 'Sniff Keywords
    InTheseTimers(ThisOne).ListSubItems.Add 4, , "" 'Save in directory
    InTheseTimers(ThisOne).ListSubItems(4).Tag = 100 'Default Timer History Depth
    InTheseTimers(ThisOne).ListSubItems.Add 5, , Format(Now, "yyyy/mm/dd") 'Creation date. This is a string (the "format" function converts the date into a string)
    InTheseTimers(ThisOne).ListSubItems(5).Tag = CDbl(Now) 'This is the full date
    InTheseTimers(ThisOne).ListSubItems.Add 6, , "00:00:00" 'Timer count for today only
    InTheseTimers(ThisOne).ListSubItems(6).Tag = CDbl(Now) 'This holds the date at the very last moment the timer was updated.
    InTheseTimers(ThisOne).ListSubItems.Add 7, , DoGetTimerTypes(0) 'Timer Type
    InTheseTimers(ThisOne).ListSubItems(7).Tag = 0 'This tag receives values from the enum TimerTypes found in the frmTimerEdit module
    InTheseTimers(ThisOne).ListSubItems.Add 8, , "000:00:00", , "hhh:mm:ss" 'TripCount
    InTheseTimers(ThisOne).ListSubItems(8).Tag = CDbl(0)
    InTheseTimers(ThisOne).ListSubItems.Add 9, , Format(Now, "yyyy/mm/dd HH:MM") 'Reset date
    InTheseTimers(ThisOne).ListSubItems(9).Tag = CDbl(Now) 'This is the full date as double
    
End Sub

Public Sub TripCounterReset(ThisListitem As ListItem)
    ThisListitem.ListSubItems(8).Tag = CDbl(0)
    ThisListitem.ListSubItems(9).Tag = CDbl(Now)
    
    ThisListitem.ListSubItems(8).Text = "000:00:00"
    ThisListitem.ListSubItems(9).Text = Format(Now, "yyyy/mm/dd HH:MM")
End Sub

Public Sub UpdateTimerStatistics(TmrDatabase As ProjectTimers)
    Dim OutputString As String
    Dim Fragment As Variant
    
    'Update the all time total
    TmrDatabase.AllTimeTotal = GetTotal(TmrDatabase.MyTimers, 2, True, True)(0)
    
    'Update the today total
    TmrDatabase.TodayTotal = GetTotal(TmrDatabase.MyTimers, 7, , True)(0)

End Sub

Public Sub UpdateMouseStatistics(TmrDatabase As ProjectTimers, MouseMoved As Boolean)
    Dim Fragment As Variant
    
    If (Month(Now) <> Month(TmrDatabase.LastMouseActive) Or Year(Now) <> Year(TmrDatabase.LastMouseActive) Or Day(Now) <> Day(TmrDatabase.LastMouseActive)) Then
        TmrDatabase.MouseBusyToday = 0
        TmrDatabase.MouseIdleToday = 0
    End If
    If (MouseMoved) Then 'This is true at first program launch (Mouse age always starts at zero)
        TmrDatabase.MouseBusyToday = TmrDatabase.MouseBusyToday + TmrDatabase.SniffInterval / Seconds.inday
        TmrDatabase.LastMouseActive = Now
    Else
        TmrDatabase.MouseIdleToday = TmrDatabase.MouseIdleToday + TmrDatabase.SniffInterval / Seconds.inday
    End If
    
    'If (TmrDatabase.TodayTotal > TmrDatabase.MouseBusyToday) Then TmrDatabase.MouseBusyToday = TmrDatabase.TodayTotal
    
End Sub

'Parses through listitems and returns the maximum serial found incremented by 10
Public Function GetAvailableSerial(FromTheseTimers As ListItems) As Long
    Dim Count As Integer
    Dim MaxFound As Long
    
    GetAvailableSerial = 0
        
    For Count = 1 To FromTheseTimers.Count
        If (MaxFound < Round(Val(FromTheseTimers(Count).Tag), 0)) Then MaxFound = Round(Val(FromTheseTimers(Count).Tag), 0)
    Next
    
    GetAvailableSerial = CLng(MaxFound + 10)
End Function

Private Sub UpdateTimerHistory(OfThisTimer As ListItem, InTimersHistory As Collection)
    Dim OneLineOfHistory As New HistoryEntry 'create the object (out of the class blueprint) so that it can be assigned to the "inTimersHistrory" collection
    Dim ThisTimerHistoryCount As Integer
    
    OneLineOfHistory.LinkToTimerID = CLng(OfThisTimer.Tag)
    OneLineOfHistory.OnThisDate = Year(OfThisTimer.ListSubItems(6).Tag) & "/" & Format(Month(OfThisTimer.ListSubItems(6).Tag), "00") & "/" & Format(Day(OfThisTimer.ListSubItems(6).Tag), "00")
    OneLineOfHistory.RunningTotal = CDbl(OfThisTimer.ListSubItems(1).Tag)
    OneLineOfHistory.TotalToday = OfThisTimer.ListSubItems(6).Text
            
    ThisTimerHistoryCount = AddHistoryItem(InTimersHistory, OneLineOfHistory)
    If (ThisTimerHistoryCount > CInt(OfThisTimer.ListSubItems(4).Tag)) Then Call DeleteOldestHistoryEntries(InTimersHistory, CLng(OfThisTimer.Tag), CInt(OfThisTimer.ListSubItems(4).Tag))
    
End Sub

Public Sub DeleteOldestHistoryEntries(InTimersHistory As Collection, ThisTimerID As Long, Optional DesiredHistoryDepth As Integer = -1, Optional ThisMany As Integer = 1)
    Dim Count As Integer
    Dim count2 As Long
    Dim CurrentHistoryDepth As Long
    Dim OldestIndex As Long
    Dim OldestDate As Double
        
    For Count = 1 To ThisMany
        OldestDate = Now
        OldestIndex = -1
        For count2 = 1 To InTimersHistory.Count
            If (InTimersHistory.Item(count2).LinkToTimerID = ThisTimerID) Then CurrentHistoryDepth = CurrentHistoryDepth + 1
            If (InTimersHistory.Item(count2).LinkToTimerID = ThisTimerID And CDbl(CDate(InTimersHistory.Item(count2).OnThisDate)) < OldestDate) Then
                OldestDate = CDbl(CDate(InTimersHistory.Item(count2).OnThisDate))
                OldestIndex = count2
            End If
        Next
        If (CurrentHistoryDepth > DesiredHistoryDepth And DesiredHistoryDepth > 0) Then ThisMany = CurrentHistoryDepth - DesiredHistoryDepth
        If (OldestIndex > 0) Then InTimersHistory.Remove OldestIndex
    Next
    
End Sub

'Adds a history item (or updates an existing one if a match is found)
'Returns the resulting history depth of the related counter
Public Function AddHistoryItem(InTimersHistory As Collection, OneLineOfHistory As HistoryEntry) As Long
    Dim Counter As Long
    Dim IsFound As Boolean
    Dim EntriesCount As Long
    
    IsFound = False
    For Counter = 1 To InTimersHistory.Count
        If (CLng(InTimersHistory.Item(Counter).LinkToTimerID) = OneLineOfHistory.LinkToTimerID) Then EntriesCount = EntriesCount + 1
        If (InTimersHistory.Item(Counter).GetKeyValue = OneLineOfHistory.GetKeyValue) Then
            OneLineOfHistory.CopyAllDataTo InTimersHistory.Item(Counter)
            IsFound = True
        End If
    Next
    If (IsFound = False) Then InTimersHistory.Add OneLineOfHistory, OneLineOfHistory.GetKeyValue
    AddHistoryItem = EntriesCount
End Function

Public Sub DeleteTimerHistory(ThisTimerID As String, AllTimersHistory As Collection)
    Dim Counter As Long
    Dim HistoryCount As Long
    
    If (AllTimersHistory.Count = 0) Then Exit Sub
    
    Counter = 1
    Do While (Counter <= AllTimersHistory.Count)
        If (AllTimersHistory.Item(Counter).LinkToTimerID = ThisTimerID) Then AllTimersHistory.Remove (Counter) Else Counter = Counter + 1
    Loop
    
End Sub

Public Sub UpdateTimersContextMenu()
    If (frmMainWindow.TimersListView.ListItems.Count >= 1) Then
        'Enable appropriate contextmenu items
        frmMainWindow.ContextMenuItem(Menu.Delete).Enabled = True
        frmMainWindow.ContextMenuItem(Menu.Edit).Enabled = True
        'frmMainWindow.ContextMenuItem(Menu.Save).Enabled = True
        If (frmMainWindow.TimersListView.SelectedItem.ListSubItems(4).Text = "") Then frmMainWindow.ContextMenuItem(Menu.Save).Enabled = False Else frmMainWindow.ContextMenuItem(Menu.Save).Enabled = True
        frmMainWindow.ContextMenuItem(Menu.showhistory).Enabled = True
        frmMainWindow.ContextMenuItem(Menu.ResetTrip).Enabled = True
    Else
        frmMainWindow.ContextMenuItem(Menu.Delete).Enabled = False
        frmMainWindow.ContextMenuItem(Menu.Edit).Enabled = False
        frmMainWindow.ContextMenuItem(Menu.Save).Enabled = False
        frmMainWindow.ContextMenuItem(Menu.showhistory).Enabled = False
        frmMainWindow.ContextMenuItem(Menu.ResetTrip).Enabled = False
        
    End If
End Sub
