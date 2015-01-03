VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentState As Integer

'------------------------
'--- create tray icon ---
'------------------------
Sub minimize_to_tray()
    Form1.Hide
    currentState = 1
    nid.cbSize = Len(nid)
    nid.hwnd = Me.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = Me.Icon ' the icon will be your Form1 project icon
    nid.szTip = "Project Timers. Click to restore" & vbNullChar
    Shell_NotifyIcon NIM_ADD, nid
End Sub

'---------------------------------------------------
'-- Tray icon actions when mouse click on it, etc --
'---------------------------------------------------
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim msg As Long
    Dim sFilter As String
    msg = x / Screen.TwipsPerPixelX
    Label1.Caption = msg
    Select Case msg
        Case WM_LBUTTONDOWN
            Form1.Show ' show form
            currentState = 0
            Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
        Case WM_RBUTTONDOWN
        Case WM_RBUTTONUP
            Form1.Show
            currentState = 0
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub Form_Resize()
    Select Case Form1.WindowState
        Case 1
            If (currentState = 0) Then Call minimize_to_tray Else Form1.WindowState = 0
    End Select
End Sub

'------------------------------
'--- form Actions On unload ---
'------------------------------
Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
End Sub
