VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Sup?"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrGetFgWindow 
      Interval        =   1000
      Left            =   3240
      Top             =   120
   End
   Begin MSComctlLib.ListView lvwFGWindow 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1720
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu ctxPopup 
      Caption         =   "ctxPopup"
      Visible         =   0   'False
      Begin VB.Menu ctxClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu ctxSettingsSep 
         Caption         =   "-"
      End
      Begin VB.Menu ctxSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu ctxExitSep 
         Caption         =   "-"
      End
      Begin VB.Menu ctxExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const APP_NAME As String = "Sup"

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

' The hWnd of the most recently found window.
Private m_LastHwnd As Long
' Return the window's title.
Private Function GetWindowTitle(ByVal window_hwnd As Long) As String
Dim length As Long
Dim buf As String

    ' See how long the window's title is.
    length = GetWindowTextLength(window_hwnd) + 1
    If length <= 1 Then
        ' There's no title. Use the hWnd.
        GetWindowTitle = "<" & window_hwnd & ">"
    Else
        ' Get the title.
        buf = Space$(length)
        length = GetWindowText(window_hwnd, buf, length)
        GetWindowTitle = Left$(buf, length)
    End If
End Function

Private Sub ctxClear_Click()
    lvwFGWindow.ListItems.Clear
    m_LastHwnd = 0
End Sub

Private Sub ctxExit_Click()
    Unload Me
End Sub

Private Sub ctxSettings_Click()
    frmSettings.Show vbModal, Me
End Sub

Private Sub Form_Load()
    lvwFGWindow.View = lvwReport

    lvwFGWindow.ColumnHeaders.Clear
    lvwFGWindow.ColumnHeaders.Add Text:="Time"
    lvwFGWindow.ColumnHeaders.Add Text:="Window"

    ' Get settings.
    Me.Move _
        GetSetting(APP_NAME, "Settings", "Left", Me.Left), _
        GetSetting(APP_NAME, "Settings", "Top", Me.Top), _
        GetSetting(APP_NAME, "Settings", "Width", Me.Width), _
        GetSetting(APP_NAME, "Settings", "Height", Me.Height)
    tmrGetFgWindow.Interval = GetSetting(APP_NAME, "Settings", "Interval", 1000)
End Sub


Private Sub Form_Resize()
    lvwFGWindow.Move 0, 0, ScaleWidth, ScaleHeight
    lvwFGWindow.ColumnHeaders(2).Width = ScaleWidth - lvwFGWindow.ColumnHeaders(1).Width - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Save settings.
    SaveSetting APP_NAME, "Settings", "Left", Me.Left
    SaveSetting APP_NAME, "Settings", "Top", Me.Top
    SaveSetting APP_NAME, "Settings", "Width", Me.Width
    SaveSetting APP_NAME, "Settings", "Height", Me.Height
    SaveSetting APP_NAME, "Settings", "Interval", 1000
End Sub
Private Sub lvwFGWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu ctxPopup
    End If
End Sub


Private Sub tmrGetFgWindow_Timer()
Dim fg_hwnd As Long
Dim list_item As ListItem

    ' Get the window's handle.
    fg_hwnd = GetForegroundWindow()

    ' If this is the same as the previous foreground window,
    ' let that one remain the most recent entry.
    If m_LastHwnd = fg_hwnd Then Exit Sub
    m_LastHwnd = fg_hwnd

    ' Display the time and the window's title.
    Set list_item = lvwFGWindow.ListItems.Add(Text:=Format$(Now, "h:mm:ss"))
    list_item.SubItems(1) = GetWindowTitle(fg_hwnd)
    list_item.EnsureVisible
End Sub
