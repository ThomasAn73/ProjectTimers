VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enumerate All Top Level Windows"
   ClientHeight    =   3855
   ClientLeft      =   1650
   ClientTop       =   2100
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnum 
      Caption         =   "&Enumerate"
      Height          =   375
      Left            =   7980
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   9300
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ListBox lstWindows 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   10335
   End
   Begin VB.Label lblCount 
      Caption         =   "Windows Found: "
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Window Caption"
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Thread Id"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Process Id"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Class Name"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Handle"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEnum_Click()
lstWindows.Clear
Call fEnumWindows
lblCount = "Windows Found:   " & lstWindows.ListCount
End Sub
Private Sub cmdQuit_Click()
Unload Me
End Sub
Private Sub Form_Load()
ReDim aTabs(4) As Long
'
' Set up a listbox with TAB delimited columns.
' Add the desired tabstops to an array.
'
' NOTE: tabstops are defined in terms of "dialog units". While the
'       GetDialogBaseUnits function combined with a simple calculation
'       can be used to convert between dialog units and pixels, the
'       easiest way to set tabstops where you want is by trial and error.
'
aTabs(0) = 30
aTabs(1) = 165
aTabs(2) = 210
aTabs(3) = 255

'Clear any existing tabs.
Call SendMessageArray(lstWindows.hwnd, LB_SETTABSTOPS, 0&, 0&)
'Set the tabs.
Call SendMessageArray(lstWindows.hwnd, LB_SETTABSTOPS, 4&, aTabs(0))
End Sub


