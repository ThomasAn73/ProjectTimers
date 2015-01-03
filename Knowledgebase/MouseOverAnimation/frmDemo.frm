VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Demo"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Move this form and put your mouse over the button below:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End
End Sub

Private Sub Timer1_Timer()

    If GetMouseOver(Command1.hWnd) = True Then
        Command1.Caption = "MouseOver: True"
    Else
        Command1.Caption = "MouseOver: False"
    End If
    
End Sub
