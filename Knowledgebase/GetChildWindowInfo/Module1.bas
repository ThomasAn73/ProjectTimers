Attribute VB_Name = "Module1"
Option Explicit

Public Type ProcData
    AppHwnd As Long
    title As String
    Placement As String
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public g_Contains As String
' Search this window and display information about its children.
Public Function GetWindowStructure(window_hwnd As Long, Optional ByVal Indent As String = "") As String
Dim txt As String
Dim buf As String
Dim buflen As Long
Dim child_hwnd As Long
Dim children() As Long
Dim num_children As Integer
Dim i As Integer

    ' Get the class name.
    buflen = 256
    buf = Space$(buflen - 1)
    buflen = GetClassName(window_hwnd, buf, buflen)
    buf = Left$(buf, buflen)

    ' Display the window's class.
    txt = Indent & buf

    ' See if we found an Edit object.
    If buf = "Edit" Then txt = txt & " (" & WindowText(window_hwnd) & ")"
    txt = txt & vbCrLf

    ' Search the children.
    ' Make a list of the child windows.
    num_children = 0
    child_hwnd = GetWindow(window_hwnd, GW_CHILD)
    Do While child_hwnd <> 0
        num_children = num_children + 1
        ReDim Preserve children(1 To num_children)
        children(num_children) = child_hwnd

        child_hwnd = GetWindow(child_hwnd, GW_HWNDNEXT)
    Loop

    ' Get information on the child windows.
    For i = 1 To num_children
        txt = txt & GetWindowStructure(children(i), Indent & "    ")
    Next i

    GetWindowStructure = txt
End Function
' ************************************************
' Return the text associated with the window.
' ************************************************
Public Function WindowText(window_hwnd As Long) As String
Dim txtlen As Long
Dim txt As String

    WindowText = ""
    If window_hwnd = 0 Then Exit Function
    
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
    If txtlen = 0 Then Exit Function
    
    txtlen = txtlen + 1
    txt = Space$(txtlen)
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, txtlen, ByVal txt)
    WindowText = Left$(txt, txtlen)
End Function


Public Function EnumProc(ByVal app_hwnd As Long, ByVal lParam As Long) As Boolean
Dim buf As String * 1024
Dim title As String
Dim length As Long
Dim txt As String

    ' Get the window's title.
    length = GetWindowText(app_hwnd, buf, Len(buf))
    title = Left$(buf, length)

    ' See if the title contains the target text.
    If InStr(title, g_Contains) > 0 Then
        ' Display the window's structure.
        frmWindowList.txtResults.Text = GetWindowStructure(app_hwnd)

        ' Stop searching.
        EnumProc = 0
    Else
        ' Continue searching til find it.
        EnumProc = 1
    End If
End Function
