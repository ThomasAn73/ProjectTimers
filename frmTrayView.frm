VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmTrayView 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   Icon            =   "frmTrayView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2400
      Width           =   255
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000011&
      Index           =   3
      Visible         =   0   'False
      X1              =   3255
      X2              =   3255
      Y1              =   60
      Y2              =   250
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000011&
      Index           =   2
      Visible         =   0   'False
      X1              =   3045
      X2              =   3045
      Y1              =   60
      Y2              =   250
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000011&
      Index           =   1
      Visible         =   0   'False
      X1              =   3045
      X2              =   3255
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000011&
      Index           =   0
      Visible         =   0   'False
      X1              =   3045
      X2              =   3255
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3030
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000011&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000011&
      X1              =   0
      X2              =   3360
      Y1              =   2810
      Y2              =   2810
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000011&
      X1              =   3350
      X2              =   3350
      Y1              =   2880
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   0
      X2              =   3360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Today"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   1
      Left            =   2265
      TabIndex        =   2
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Timer Name"
      ForeColor       =   &H80000011&
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmTrayView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call InitColumnHeaders
End Sub

Private Sub InitColumnHeaders()
    ListView1.ColumnHeaders.Add 1, , "Timer Name", 2000
    ListView1.ColumnHeaders.Add 2, , "Today", 825, 2
End Sub

