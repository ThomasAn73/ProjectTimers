Attribute VB_Name = "SniffOpenWindows"
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long ''This is to pin the window on top (or not)

'Functions needed to prevent multiple Timers windows instances
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNORMAL = 1

Public Type WindowData
    ParentHandle As Long
    ParentCaption As String
    ChildHandle As Long
    ChildCaption As String
End Type

Public ActiveWindow As WindowData

Public Function DoSniffActiveWindow() As WindowData
    Dim hForeground As Long
    Dim hKeyFocus As Long
    Dim Caption As String * 512
    Dim CaptionLength As Long
    
    hKeyFocus = GetFocus 'Handle to the window that has the keyboard focus
    hForeground = GetForegroundWindow() 'Handle to the parent foreground active window (This is the application main window, not a child window)
    If hKeyFocus = 0 Then
        Call AttachThreadInput(GetWindowThreadProcessId(hForeground, 0&), GetCurrentThreadId, True)
        hKeyFocus = GetFocus
        Call AttachThreadInput(GetWindowThreadProcessId(hForeground, 0&), GetCurrentThreadId, False)
    End If
    CaptionLength = GetWindowTextLength(hForeground)
    Call GetWindowText(hForeground, Caption, CaptionLength + 1)
    DoSniffActiveWindow.ParentHandle = hForeground
    DoSniffActiveWindow.ParentCaption = Left(Caption, CaptionLength)
    CaptionLength = GetWindowTextLength(hKeyFocus)
    Call GetWindowText(hKeyFocus, Caption, CaptionLength + 1)
    DoSniffActiveWindow.ChildHandle = hKeyFocus
    DoSniffActiveWindow.ChildCaption = Left(Caption, CaptionLength)
End Function
