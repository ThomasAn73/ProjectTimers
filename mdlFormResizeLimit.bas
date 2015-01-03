Attribute VB_Name = "mdlFormResizeLimit"
'This is not my code (taken from the microsoft website)

Option Explicit

      Private Const GWL_WNDPROC = -4
      Private Const WM_GETMINMAXINFO = &H24

      Private Type POINTAPI
          x As Long
          y As Long
      End Type

      Private Type MINMAXINFO
          ptReserved As POINTAPI
          ptMaxSize As POINTAPI
          ptMaxPosition As POINTAPI
          ptMinTrackSize As POINTAPI
          ptMaxTrackSize As POINTAPI
      End Type

      Global lpPrevWndProc As Long
      Global gHW As Long

    Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Sub CopyMemoryToMinMaxInfo Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
    Private Declare Sub CopyMemoryFromMinMaxInfo Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)

Public Sub Hook()
    'Start subclassing.
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim temp As Long

    'Cease subclassing.
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MinMax As MINMAXINFO

    'Check for request for min/max window sizes.
    If uMsg = WM_GETMINMAXINFO Then
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.x = 480
        MinMax.ptMinTrackSize.y = 200

        'Specify new maximum size for window.
        MinMax.ptMaxTrackSize.x = 1200
        MinMax.ptMaxTrackSize.y = 1000

        'Copy local structure back.
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Else
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    End If
End Function

