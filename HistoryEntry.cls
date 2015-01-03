VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This module basically provides a user defined type but in the form of a class

'local variable(s) to hold property value(s)
Private mvarLinkToTimerID As Long 'local copy
Private mvarOnThisDate As String 'local copy
Private mvarRunningTotal As Double 'local copy
Private mvarTotalToday As String 'local copy
Public Function GetKeyValue() As String
    GetKeyValue = CStr(mvarLinkToTimerID & Replace(mvarOnThisDate, "/", ""))
End Function


Public Property Let TotalToday(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalToday = 5
    mvarTotalToday = vData
End Property


Public Property Get TotalToday() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalToday
    TotalToday = mvarTotalToday
End Property



Public Property Let RunningTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RunningTotal = 5
    mvarRunningTotal = vData
End Property


Public Property Get RunningTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RunningTotal
    RunningTotal = mvarRunningTotal
End Property



Public Property Let OnThisDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OnThisDate = 5
    mvarOnThisDate = vData
End Property


Public Property Get OnThisDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OnThisDate
    OnThisDate = mvarOnThisDate
End Property



Public Property Let LinkToTimerID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LinkToTimerID = 5
    mvarLinkToTimerID = vData
End Property


Public Property Get LinkToTimerID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LinkToTimerID
    LinkToTimerID = mvarLinkToTimerID
End Property

Public Sub CopyAllDataTo(ThisObject As HistoryEntry)
    
    ThisObject.LinkToTimerID = mvarLinkToTimerID
    ThisObject.OnThisDate = mvarOnThisDate
    ThisObject.RunningTotal = mvarRunningTotal
    ThisObject.TotalToday = mvarTotalToday

End Sub
