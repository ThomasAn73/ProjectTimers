VERSION 5.00
Begin VB.Form frmWindowList 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   1620
   ClientTop       =   1545
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7140
   Begin VB.TextBox txtResults 
      Height          =   4815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   5655
   End
   Begin VB.TextBox txtContains 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   " - Microsoft Internet Explorer"
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdFindWindows 
      Caption         =   "Find Windows"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Contains"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmWindowList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal lParam As Long) As Long
' Start the enumeration.
Private Sub cmdFindWindows_Click()
    g_Contains = txtContains.Text
    EnumWindows AddressOf EnumProc, 0
End Sub

Private Sub Form_Resize()
Dim wid As Single
Dim hgt As Single

    wid = ScaleWidth - txtContains.Left
    If wid < 120 Then wid = 120
    txtContains.Width = wid

    hgt = ScaleHeight - txtResults.Top
    If hgt < 120 Then hgt = 120
    txtResults.Width = ScaleWidth
    txtResults.Height = hgt

    cmdFindWindows.Left = (ScaleWidth - cmdFindWindows.Width) / 2
End Sub


