VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parag's Mouse Utility"
   ClientHeight    =   2790
   ClientLeft      =   3600
   ClientTop       =   3300
   ClientWidth     =   3855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog clrCmdg 
      Left            =   2430
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2310
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   4075
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Colours"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Picture1"
      Tab(1).Control(3)=   "Picture2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&About"
      TabPicture(2)   =   "Form1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(3)=   "Label6"
      Tab(2).Control(4)=   "Label7"
      Tab(2).Control(5)=   "Label8"
      Tab(2).ControlCount=   6
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   -72930
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   8
         Top             =   1305
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   -72930
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   810
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Settings:"
         Height          =   1635
         Left            =   270
         TabIndex        =   1
         Top             =   495
         Width           =   2355
         Begin VB.CheckBox Check1 
            Caption         =   "Show X Co-Ordinate"
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   4
            Top             =   450
            Value           =   1  'Checked
            Width           =   1770
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Show Y Co-Ordinate"
            Height          =   285
            Index           =   1
            Left            =   270
            TabIndex        =   3
            Top             =   810
            Value           =   1  'Checked
            Width           =   1770
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Show Colour"
            Height          =   285
            Index           =   2
            Left            =   270
            TabIndex        =   2
            Top             =   1170
            Value           =   1  'Checked
            Width           =   1770
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Do Visit:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   14
         Top             =   1470
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "parag_pp@hotmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74760
         MouseIcon       =   "Form1.frx":0496
         MousePointer    =   99  'Custom
         TabIndex        =   13
         ToolTipText     =   "mailto:parag_pp@hotmail.com"
         Top             =   1230
         Width           =   1995
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "(don't forget to sign the guestbook)"
         Height          =   195
         Left            =   -74370
         TabIndex        =   12
         Top             =   1950
         Width           =   2475
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "http://www.instantweb.com/p/paragpp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74850
         MouseIcon       =   "Form1.frx":07A0
         MousePointer    =   99  'Custom
         TabIndex        =   11
         ToolTipText     =   "http://www.instantweb.com/p/paragpp/Main/main.html"
         Top             =   1680
         Width           =   3345
      End
      Begin VB.Label Label4 
         Caption         =   "Please mail any comments or improvements to:"
         Height          =   435
         Left            =   -74760
         TabIndex        =   10
         Top             =   810
         Width           =   2475
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Parag's Mouse Utility Version 2.00"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   9
         Top             =   510
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Background Colour:"
         Height          =   195
         Left            =   -74505
         TabIndex        =   6
         Top             =   1350
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tip Colour:"
         Height          =   195
         Left            =   -74505
         TabIndex        =   5
         Top             =   855
         Width           =   765
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3060
      Top             =   2115
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3060
      Top             =   2520
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuChange 
         Caption         =   "&Disable"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "&Close Parag's Utility"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Form_Activate()
    If App.PrevInstance = True Then
        MsgBox "Parag's Utility Already Running!", vbExclamation, "Already Running"
        End
    End If

End Sub

Private Sub Form_Load()
    
    AddToTray Me, mnuTray
    
    SetTrayTip "Parag's Utility Working"
    Me.Hide
    Form2.Visible = True
    Enab = True
    first = 102
    SetWindowPos Form2.hwnd, _
            HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE + SWP_NOSIZE
End Sub

' Important! Remove the tray icon.
Private Sub Form_Unload(Cancel As Integer)
    If unloadType = 1 Then
        RemoveFromTray
        Unload Form2
    Else
        Cancel = 1
        Form1.Hide
    SetWindowPos Form2.hwnd, _
            HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE + SWP_NOSIZE
    End If
End Sub


Private Sub mnuFileExit_Click()
    End
    'Unload Form2
    'Unload Me
End Sub

Private Sub Label5_Click()
    Call ShellExecute(hwnd, "Open", "http://www.instantweb.com/p/paragpp/Main/main.html", 0, 0, 0)
End Sub

Private Sub Label7_Click()
    Call ShellExecute(hwnd, "Open", "mailto:parag_pp@hotmail.com", 0, 0, 0)
End Sub


Private Sub mnuChange_Click()
    If Enab = True Then
        SetTrayIcon LoadResPicture(101, vbResIcon)
        Enab = False
        mnuChange.Caption = "&Enable"
        SetTrayTip "Parag's Utility Stopped!"
        Exit Sub
    Else
        SetTrayIcon LoadResPicture(102, vbResIcon)
        Enab = True
        mnuChange.Caption = "&Disable"
        SetTrayTip "Parag's Utility Working"
        Exit Sub
    End If
End Sub

Private Sub mnuTrayClose_Click()
    unloadType = 1
    Unload Me
End Sub


Private Sub mnuTrayMaximize_Click()
    WindowState = vbMaximized
End Sub


Private Sub mnuTrayMinimize_Click()
    WindowState = vbMinimized
End Sub

Private Sub Picture1_Click()
    If Enab = True Then
        Form2.Visible = False
    End If
    
    clrCmdg.Color = Picture1.BackColor
    clrCmdg.ShowColor
    Picture1.BackColor = clrCmdg.Color
    Form2.Label1.ForeColor = Picture1.BackColor

    If Enab = True Then
        Form2.Visible = True
    End If
End Sub

Private Sub Picture2_Click()
    If Enab = True Then
        Form2.Visible = False
    End If
    
    clrCmdg.Color = Picture2.BackColor
    clrCmdg.ShowColor
    Picture2.BackColor = clrCmdg.Color
    Form2.BackColor = Picture2.BackColor

    If Enab = True Then
        Form2.Visible = True
    End If
End Sub

Private Sub Timer1_Timer()
    Dim dc, pnt As POINTAPI, str As String, colr
    Dim YesAll As Boolean
    Dim desktop_handle As Long
    
    If Enab = True Then
        desktop_handle = GetDesktopWindow()
        dc = GetWindowDC(desktop_handle)
        GetCursorPos pnt
        colr = GetPixel(dc, pnt.x, pnt.y)
        ReleaseDC desktop_handle, dc
        
        str = ""
        YesAll = True
        
        If Form1.Check1(1).Value = 1 Then
            str = "Ypos: " & Format(pnt.y, "0###")
            YesAll = False
        End If
        
        If Form1.Check1(0).Value = 1 Then
            If Form1.Check1(1).Value = 1 Then
                str = str & ", Xpos: " & Format(pnt.x, "0###")
            Else
                str = str & "Xpos: " & Format(pnt.x, "0###")
            End If
            
            YesAll = False
        End If
            
        If Form1.Check1(2).Value = 1 Then
            If Form1.Check1(1).Value = 1 Or Form1.Check1(0).Value = 1 Then
                str = str & ", Color: " & Format(Hex(colr), "0#####") & Space(5)
            Else
                str = str & "Color: " & Format(Hex(colr), "0#####") & Space(5)
            End If
            
            YesAll = False
        End If
        
        If YesAll = True Then
            Form2.Visible = False
            Exit Sub
        Else
            Form2.Visible = True
        End If
        
        Form2.Label1 = str
        Form2.Width = Form2.Label1.Width + 50
        Form2.Left = (pnt.x * Screen.TwipsPerPixelX) + 100
        Form2.Top = (pnt.y * Screen.TwipsPerPixelY) + 100
    Else
        Form2.Visible = False
    End If
End Sub

Private Sub Timer2_Timer()
    If Enab = True Then
        If first = 102 Then
            SetTrayIcon LoadResPicture(103, vbResIcon)
            Label3.ForeColor = &HFF0000 ' blue
            first = 103
        Else
            SetTrayIcon LoadResPicture(102, vbResIcon)
            Label3.ForeColor = &HFF& 'red
            first = 102
        End If
    End If
End Sub
