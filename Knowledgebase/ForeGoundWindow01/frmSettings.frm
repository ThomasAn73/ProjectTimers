VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sup Settings"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Interval (ms)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo BadInterval
    frmMain.tmrGetFgWindow.Interval = CLng(txtInterval.Text)
    Unload Me
    Exit Sub

BadInterval:
    MsgBox "Interval must be an integer between 100 and 65535", _
        vbExclamation, "Invalid interval"
    Exit Sub
End Sub

Private Sub Form_Load()
    ' Display the current values.
    txtInterval.Text = frmMain.tmrGetFgWindow.Interval
End Sub


