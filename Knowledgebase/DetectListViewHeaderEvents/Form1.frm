VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer SynchColumns 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   285
      Top             =   3060
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2730
      Left            =   180
      TabIndex        =   0
      Top             =   210
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
   Begin MSComctlLib.ListView ListView2 
      Height          =   2730
      Left            =   225
      TabIndex        =   1
      Top             =   3555
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  With ListView1.ListItems.Add(, , "a")
    .ListSubItems.Add , , "a1"
    .ListSubItems.Add , , "a2"
  End With
  With ListView1.ListItems.Add(, , "b")
    .ListSubItems.Add , , "b1"
    .ListSubItems.Add , , "b2"
  End With
  With ListView1.ListItems.Add(, , "c")
    .ListSubItems.Add , , "c1"
    .ListSubItems.Add , , "c2"
  End With

  With ListView2.ListItems.Add(, , "aa")
    .ListSubItems.Add , , "aa1"
    .ListSubItems.Add , , "aa2"
  End With
  With ListView2.ListItems.Add(, , "bb")
    .ListSubItems.Add , , "bb1"
    .ListSubItems.Add , , "bb2"
  End With
  With ListView2.ListItems.Add(, , "cc")
    .ListSubItems.Add , , "cc1"
    .ListSubItems.Add , , "cc2"
  End With

  Call RegisterListView(ListView1)
  Call RegisterListView(ListView2)

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Call UnregisterListView(ListView2)
  Call UnregisterListView(ListView1)

End Sub

Public Function ListView1_HeaderEvent(ByVal Action As lvHeaderActions, ByVal Column As Long) As Boolean

  Call ListViewHeaderEventDebugPrint(ListView1, Action, Column)
  
  ' Illustrates one way to synchronize ListView column widths, with consideration for problem with some
  ' versions of MSCOMCTL.ocx that don't set the width property until after the resize end event.  Note that
  ' it is still not entirely reliable due to the Column = 2 condition.  For better reliability, remove the
  ' condition; however, this will result in more resizes than will be necessary.
  
  If Column = 2 Then
    If Action = lvHeaderActionResizeEnd Then
      SynchColumns.Enabled = True
    End If
  End If
  
  ' Here's an alternative technique that doesn't use a timer, but it doesn't work reliably when the resizing
  ' is too fast or triggered by a lvHeaderActionDividerDoubleClick.
  
  'If Column = 2 Then
  '  If Action = lvHeaderActionChanged Or Action = lvHeaderActionDividerDoubleClick Then
  '    ListView2.ColumnHeaders(2).Width = ListView1.ColumnHeaders(2).Width
  '  End If
  'End If

End Function

Public Function ListView2_HeaderEvent(ByVal Action As lvHeaderActions, ByVal Column As Long) As Boolean

  Call ListViewHeaderEventDebugPrint(ListView2, Action, Column)
  
End Function

Private Sub SynchColumns_Timer()

  SynchColumns.Enabled = False

  ListView2.ColumnHeaders(2).Width = ListView1.ColumnHeaders(2).Width

End Sub
