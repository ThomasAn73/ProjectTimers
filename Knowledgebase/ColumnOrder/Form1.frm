VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETCOLUMN As Long = (LVM_FIRST + 25)
Private Const LVM_GETCOLUMNORDERARRAY As Long = (LVM_FIRST + 59)
Private Const LVCF_TEXT As Long = &H4
Private Const LVCF_ORDER As Long = &H20

Private Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long



Private Sub Form_Load()

Dim itmx As ListItem

With ListView1
.ColumnHeaders.Add , , "Col 1"
.ColumnHeaders.Add , , "Col 2"
.ColumnHeaders.Add , , "Col 3"
.ColumnHeaders.Add , , "Col 4"

.View = lvwReport
.FullRowSelect = True
.AllowColumnReorder = True

Set itmx = .ListItems.Add(, , "Item 1")
itmx.SubItems(1) = "Item1 Sub1"
itmx.SubItems(2) = "Item1 Sub2"
itmx.SubItems(3) = "Item1 Sub3"

Set itmx = .ListItems.Add(, , "Item 2")
itmx.SubItems(1) = "Item2 Sub1"
itmx.SubItems(2) = "Item2 Sub2"
itmx.SubItems(3) = "Item2 Sub3"

Set itmx = .ListItems.Add(, , "Item 3")
itmx.SubItems(1) = "Item3 Sub1"
itmx.SubItems(2) = "Item3 Sub2"
itmx.SubItems(3) = "Item3 Sub3"

Set itmx = .ListItems.Add(, , "Item 4")
itmx.SubItems(1) = "Item4 Sub1"
itmx.SubItems(2) = "Item4 Sub2"
itmx.SubItems(3) = "Item4 Sub3"
End With

Command1.Caption = "Print column order (VB)"
Command2.Caption = "Print column order (API)"

End Sub

Private Sub Command1_Click()

Dim cnt As Long

Debug.Print ""
Debug.Print Command1.Caption

For cnt = 1 To ListView1.ColumnHeaders.Count
Debug.Print ListView1.ColumnHeaders(cnt).Text
Next

End Sub

Private Sub Command2_Click()

'working variables
Dim cnt As Long
Dim firstCol As Long
Dim lastCol As Long
Dim totalCols As Long

Dim msg As String
Dim tmp As String
Dim lvc As LVCOLUMN

'initialize the variables needed.
'totalCols is the 1-based
'total required for the API.
'lastCol is the 0-based
'number of columns in the listview
totalCols = ListView1.ColumnHeaders.Count
firstCol = 0
lastCol = totalCols - 1

'to get the column order, we have to pass
'an array to the API. On return, it will
'be filled with the index of the column in
'incrementing positions. For example,
'if column 2 from a 4-column header was
'moved to the first position the return
'array would hold 2, 1, 0, 3.
ReDim posarray(firstCol To lastCol) As Long

Call SendMessage(ListView1.hwnd, LVM_GETCOLUMNORDERARRAY, totalCols, posarray(firstCol))

Debug.Print ""
Debug.Print Command2.Caption
'with the array filled, loop through the
'array, and passing each item as the
'position (wParam). The LVCOLUMN type
'will be filled with the data for the
'passed index (LVCF_TEXT in this example).
For cnt = firstCol To lastCol

    'padded with sufficient room to
    'hold the columnheader string.
    tmp = Space$(256)
    
    With lvc
    .mask = LVCF_TEXT Or LVCF_ORDER
    .pszText = tmp
    .cchTextMax = Len(tmp)
    End With
    
    Call SendMessage(ListView1.hwnd, LVM_GETCOLUMN, posarray(cnt), lvc)
    
    'strip the trailing null
    Debug.Print Left$(lvc.pszText, InStr(lvc.pszText, Chr$(0)) - 1)

Next

End Sub

