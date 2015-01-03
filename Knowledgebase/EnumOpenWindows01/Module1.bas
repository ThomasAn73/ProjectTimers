Attribute VB_Name = "Module1"
Option Explicit

Public Const LB_SETTABSTOPS = &H192
Public Const MAX_PATH = 260

Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lpData As Long) As Long
Dim lResult    As Long
Dim lThreadId  As Long
Dim lProcessId As Long
Dim sWndName   As String
Dim sClassName As String
'
' This callback function is called by Windows (from the EnumWindows
' API call) for EVERY window that exists.  It populates the aWindowList
' array with a list of windows that we are interested in.
'
fEnumWindowsCallBack = 1
sClassName = Space$(MAX_PATH)
sWndName = Space$(MAX_PATH)

lResult = GetClassName(hwnd, sClassName, MAX_PATH)
sClassName = Left$(sClassName, lResult)
lResult = GetWindowText(hwnd, sWndName, MAX_PATH)
sWndName = Left$(sWndName, lResult)

lThreadId = GetWindowThreadProcessId(hwnd, lProcessId)

Form1.lstWindows.AddItem CStr(hwnd) & vbTab & sClassName & _
    vbTab & CStr(lProcessId) & vbTab & CStr(lThreadId) & _
    vbTab & sWndName
End Function

Public Function fEnumWindows() As Boolean
Dim hwnd As Long
'
' The EnumWindows function enumerates all top-level windows
' on the screen by passing the handle of each window, in turn,
' to an application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'
Call EnumWindows(AddressOf fEnumWindowsCallBack, hwnd)
End Function




