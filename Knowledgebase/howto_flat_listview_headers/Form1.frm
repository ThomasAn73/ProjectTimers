VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Flat Headers"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Normal Headers"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_STYLE        As Long = (-16)
Private Const LVM_FIRST        As Long = &H1000
Private Const LVM_GETHEADER    As Long = (LVM_FIRST + 31)
Private Const HDS_BUTTONS      As Long = 2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' Load some data into this ListView.
Private Sub PrepareListView(ByVal lvw As ListView)
Const INCH As Single = 1440
Dim list_item As ListItem

    lvw.ColumnHeaders.Clear
    lvw.ColumnHeaders.Add Text:="Title", Width:=3.5 * INCH
    lvw.ColumnHeaders.Add Text:="URL", Width:=2.5 * INCH
    lvw.View = lvwReport

    Set list_item = lvw.ListItems.Add(Text:="Visual Basic Graphics Programming")
    list_item.ListSubItems.Add Text:="http://www.vb-helper.com/vbgp.htm"
    Set list_item = lvw.ListItems.Add(Text:="Visual Basic Algorithms")
    list_item.ListSubItems.Add Text:="http://www.vb-helper.com/vba.htm"
    Set list_item = lvw.ListItems.Add(Text:="Microsoft Office Programming: A Guide for Experienced Developers")
    list_item.ListSubItems.Add Text:="http://www.vb-helper.com/office.htm"
    Set list_item = lvw.ListItems.Add(Text:="Visual Basic .NET Database Programming")
    list_item.ListSubItems.Add Text:="http://www.vb-helper.com/vbdb.htm"
    Set list_item = lvw.ListItems.Add(Text:="Visual Basic .NET and XML")
    list_item.ListSubItems.Add Text:="http://www.vb-helper.com/xml.htm"
    Set list_item = lvw.ListItems.Add(Text:="Prototyping with Visual Basic")
    list_item.ListSubItems.Add Text:="http://www.vb-helper.com/proto.htm"
End Sub

' Give the ListView a flat header style.
Private Sub SetFlatHeaders(ByVal lvw As ListView)
Dim lv_hwnd As Long
Dim hHeader As Long
Dim Style   As Long

    ' Get the handle to the listview header
    lv_hwnd = lvw.hwnd
    hHeader = SendMessage(lv_hwnd, LVM_GETHEADER, 0, ByVal 0&)

    ' Set the new style
    Style = GetWindowLong(hHeader, GWL_STYLE)
    Style = Style And Not HDS_BUTTONS
    SetWindowLong hHeader, GWL_STYLE, Style
End Sub
Private Sub Form_Load()
    PrepareListView ListView1
    PrepareListView ListView2
    SetFlatHeaders ListView2
End Sub
Private Sub Form_Resize()
Dim hgt As Single

    hgt = (ScaleHeight - 2 * Label1.Height - 120) / 2
    If hgt < 120 Then hgt = 120

    ListView1.Move 0, Label1.Height, ScaleWidth, hgt

    Label2.Top = ListView1.Top + ListView1.Height + 120
    ListView2.Move 0, Label2.Top + Label2.Height, ScaleWidth, hgt
End Sub


