Attribute VB_Name = "FlatenListViewHeadrers"
Option Explicit

Private Const GWL_STYLE        As Long = (-16)
Private Const LVM_FIRST        As Long = &H1000
Private Const LVM_GETHEADER    As Long = (LVM_FIRST + 31)
Private Const HDS_BUTTONS      As Long = 2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

' Give the ListView a flat header style.
Public Sub SetFlatHeaders(ByVal lvw As ListView)
    Dim lv_hwnd As Long
    Dim hHeader As Long
    Dim Style   As Long

    ' Get the handle to the listview header
    lv_hwnd = lvw.hwnd
    hHeader = SendMessage(lv_hwnd, LVM_GETHEADER, 0, ByVal 0&)

    ' Set the new style
    Style = GetWindowLong(hHeader, GWL_STYLE)
    Style = Style And Not HDS_BUTTONS
    SetWindowLong hHeader, GWL_STYLE, Style
End Sub


