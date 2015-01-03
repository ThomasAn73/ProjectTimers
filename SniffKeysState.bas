Attribute VB_Name = "SniffKeysState"
'Do not use this it does not work

Option Explicit

Public Type KeyState
    CurrentKeyCode As Integer
    CurrentTimeStamp As Double
    PreviousKeyCode As Integer
    PreviousTimeStamp As Double
End Type

Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Public KeysCondition As KeyState

Public Function DoFindKeysState(Snapshot As KeyState) As KeyState
    Dim Count As Long
    Dim Result As Variant
    
    For Count = 0 To 255
        Result = GetAsyncKeyState(Count)
        If (Result <> 0) Then
            Snapshot.PreviousKeyCode = Snapshot.CurrentKeyCode
            Snapshot.PreviousTimeStamp = Snapshot.CurrentTimeStamp
            Snapshot.CurrentKeyCode = Result
            Snapshot.CurrentTimeStamp = CDbl(Now)
            Exit For
        End If
    Next
    
End Function
