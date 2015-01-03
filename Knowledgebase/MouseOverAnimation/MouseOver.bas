Attribute VB_Name = "MouseOver"
'**********************************************************
'* Written By: Rob S    Sep 2001, modified Mar 2003       *
'* You may only use this code if this message is included *
'* E-mail me with any feedback of tips to:                *
'* rjs9565@aol.com                                        *
'* Thanks for viewing my code!                            *
'**********************************************************

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Function GetMouseOver(hWnd As Long) As Boolean

    Dim wRect As RECT
    Dim Mouse As PointAPI
    
    GetCursorPos Mouse
    GetWindowRect hWnd, wRect
    
    If (Mouse.X <= wRect.Right And Mouse.X >= wRect.Left) And (Mouse.Y <= wRect.Bottom And Mouse.Y >= wRect.Top) Then
        GetMouseOver = True
    Else
        GetMouseOver = False
    End If

End Function
