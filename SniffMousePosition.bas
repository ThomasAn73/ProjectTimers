Attribute VB_Name = "SniffMousePosition"
Option Explicit

'Structure
Public Type CursorState
    X As Long
    y As Long
    CurrentXY_TimeStamp As Double 'The moment of probing
    PreviousX As Long
    PreviousY As Long
    PreviousXY_TimeStamp As Double 'The moment of last probing
    age As Double 'The length of time the cursor has been in this position
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As CursorState) As Long

Public MouseCondition As CursorState

Public Function DoFindCursorState(Snapshot As CursorState) As CursorState

    GetCursorPos Snapshot
    Snapshot.CurrentXY_TimeStamp = Now
    If (Snapshot.X = Snapshot.PreviousX And Snapshot.y = Snapshot.PreviousY) Then
        Snapshot.age = Snapshot.CurrentXY_TimeStamp - Snapshot.PreviousXY_TimeStamp
    Else
        Snapshot.PreviousX = Snapshot.X
        Snapshot.PreviousY = Snapshot.y
        Snapshot.PreviousXY_TimeStamp = Snapshot.CurrentXY_TimeStamp
        Snapshot.age = 0
    End If
    DoFindCursorState = Snapshot
End Function
