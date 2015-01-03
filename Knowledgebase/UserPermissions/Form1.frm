VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Permissions As Variant
    
    Me.CommonDialog1.ShowOpen
    'Permissions = CheckPermissions(Me.CommonDialog1.Filename)
    Permissions = CheckPermissions("C:\Documents and Settings\ThomasAn\Desktop\Untitled-ProjectTime5.txt")
    
    MsgBox "Owner is " & Permissions(0)
    MsgBox "Write permission is " & Permissions(2)
    MsgBox "Read permission is " & Permissions(1)
End Sub

