Attribute VB_Name = "mdlMinToTray"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NotifyIconData) As Boolean

Public Type NotifyIconData
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Type TrayTrack
    UserHoverInTray As Boolean
    EntryXY As CursorState
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_Add = &H0
Public Const NIM_Modify = &H1
Public Const NIM_Delete = &H2
Public Const NIF_Message = &H1
Public Const NIF_Icon = &H2
Public Const NIF_Tip = &H4
Public Const WM_Mousemove = &H200       'Mouse Move
Public Const WM_LButtonDown = &H201     'Button down
Public Const WM_LButtonUp = &H202       'Button up
Public Const WM_LButtonDblClk = &H203   'Double-click
Public Const WM_RButtonDown = &H204     'Button down
Public Const WM_RButtonUp = &H205       'Button up
Public Const WM_RButtonDblClk = &H206   'Double-click

Public nid As NotifyIconData
Public IsWindowInTray As Integer
Public TrayMotion As TrayTrack


Public Function MinimizeToTray(ThisForm As Form) As Integer
    
    ThisForm.Hide
    nid.cbSize = Len(nid)
    nid.hWnd = ThisForm.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_Icon Or NIF_Tip Or NIF_Message
    nid.uCallBackMessage = WM_Mousemove
    nid.hIcon = ThisForm.Icon ' the icon will be your thisform project icon
    nid.szTip = "Project Timers. Click to restore" & vbNullChar
    Shell_NotifyIcon NIM_Add, nid
    
    MinimizeToTray = 1
    
End Function
