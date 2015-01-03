VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   LinkTopic       =   "Form2"
   ScaleHeight     =   3915
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView MyListView 
      Height          =   2730
      Left            =   375
      TabIndex        =   0
      Top             =   555
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   4815
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "First"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Second"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Third"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  With MyListView.ListItems.Add(, , "d")
    .ListSubItems.Add , , "d1"
    .ListSubItems.Add , , "d2"
  End With
  With MyListView.ListItems.Add(, , "e")
    .ListSubItems.Add , , "e1"
    .ListSubItems.Add , , "e2"
  End With
  With MyListView.ListItems.Add(, , "f")
    .ListSubItems.Add , , "f1"
    .ListSubItems.Add , , "f2"
  End With
    
  Call RegisterListView(MyListView)

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Call UnregisterListView(MyListView)

End Sub

Public Function MyListView_HeaderEvent(ByVal Action As lvHeaderActions, ByVal Column As Long) As Boolean

  Call ListViewHeaderEventDebugPrint(MyListView, Action, Column)
  
  ' Illustrates how to cancel all resizing actions.
  
  If Column = 2 Then
    MyListView_HeaderEvent = True
  End If
  
End Function

