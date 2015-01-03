Attribute VB_Name = "ListViewHeaders"
' Use at your own risk.  No warranties provided or liability assumed.  For demonstration purposes only.
'
' 1) Include this module in your project.
' 2) Include the declarations portion of the Declarations.bas module in your project.
' 3) Add RegisterListView() calls in the Form_Load event for each ListView control you want to monitor.
' 4) Add UnregisterListView() call in the Form_Unload event for each ListView control you monitored.
' 5) Add a xxx_HeaderEvent() function to each form for each monitored ListView control with this signature:
'
'     Public Function ListViewControlName_HeaderEvent(ByVal Action As lvHeaderActions, ByVal Column As Long) As Boolean
'
' 6) Set the function's result to True to cancel resize events (you can condition this on the column).
'
' WARNING: Never use the VB IDE's Stop button when running in the development environment.  This will cause the
'          the unregister calls to be skipped for any open forms, and will likely cause a GPF.

Option Explicit

'
' API declarations.
'

Private Const LVM_FIRST = &H1000
Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Private Type NMHDR
  hWndFrom As Long
  idfrom   As Long
  code     As Long
End Type

Private Type HD_HITTESTINFO
  pt    As POINTAPI
  flags As Long
  iItem As Long
End Type

Private Const HHT_ONHEADER = &H2
Private Const HHT_ONDIVIDER = &H4

Private Const HDM_HITTEST As Long = &H1206

Private Const HDN_FIRST            As Long = -300&
Private Const HDN_ITEMCLICK        As Long = (HDN_FIRST - 2)
Private Const HDN_DIVIDERDBLCLICK  As Long = (HDN_FIRST - 5)
Private Const HDN_BEGINTRACK       As Long = (HDN_FIRST - 6)
Private Const HDN_ENDTRACK         As Long = (HDN_FIRST - 7)
Private Const HDN_TRACK            As Long = (HDN_FIRST - 8)
Private Const HDN_GETDISPINFO      As Long = (HDN_FIRST - 9)
Private Const HDN_BEGINDRAG        As Long = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG          As Long = (HDN_FIRST - 11)
Private Const HDN_ITEMCHANGING     As Long = (HDN_FIRST - 0)
Private Const HDN_ITEMCHANGED      As Long = (HDN_FIRST - 1)
Private Const HDN_ITEMDBLCLICK     As Long = (HDN_FIRST - 3)
Private Const HDN_NM_RCLICK        As Long = -5

' Header event actions.

Public Enum lvHeaderActions
  lvHeaderActionClick = 1
  lvHeaderActionRightClick = 2
  lvHeaderActionDividerDoubleClick = 3
  lvHeaderActionResizeBegin = 4
  lvHeaderActionResizeEnd = 5
  lvHeaderActionChanging = 6
  lvHeaderActionChanged = 7
  lvHeaderActionDragBegin = 8
  lvHeaderActionDragEnd = 9
End Enum

'
' Private declarations.
'

Private RegisteredListViewControls As New Collection

Public Sub RegisterListView(ByVal ListViewControl As ListView)
   
  Call SetProp(ListViewControl.hWnd, "OrigWindowProc", GetWindowLong(ListViewControl.hWnd, GWL_WNDPROC))
  
  Call SetWindowLong(ListViewControl.hWnd, GWL_WNDPROC, AddressOf HandleListViewHeaderMsgs)
  
  Call RegisteredListViewControls.Add(ListViewControl, CStr(ListViewControl.hWnd))
   
End Sub

Public Sub UnregisterListView(ByVal ListViewControl As ListView)
   
  Dim OrigWindowProc As Long
  
  OrigWindowProc = GetProp(ListViewControl.hWnd, "OrigWindowProc")
  
  If (OrigWindowProc <> 0) Then
    Call SetWindowLong(ListViewControl.hWnd, GWL_WNDPROC, OrigWindowProc)
  End If
   
  Call RegisteredListViewControls.Remove(CStr(ListViewControl.hWnd))
   
End Sub

Public Function HandleListViewHeaderMsgs(ByVal ListViewhWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
   
  Const EVENT_SUFFIX As String = "_HeaderEvent"
   
  Dim ListViewControl As ListView
  Dim NmHdrMsg        As NMHDR
  Dim PointStruct     As POINTAPI
  Dim HitTestInfo     As HD_HITTESTINFO
  Dim HeaderhWnd      As Long
  Dim HeaderAction    As lvHeaderActions
  Dim CancelMsg       As Boolean
  
  If msg = WM_NOTIFY Then

    HandleListViewHeaderMsgs = CallWindowProc(GetProp(ListViewhWnd, "OrigWindowProc"), ListViewhWnd, msg, wp, lp)
  
    Call CopyMemory(NmHdrMsg, ByVal lp, Len(NmHdrMsg))
    
    HeaderhWnd = SendMessage(ListViewhWnd, LVM_GETHEADER, 0&, ByVal 0&)
    
    If (HeaderhWnd <> 0) Then

      Call GetCursorPos(PointStruct)
      Call ScreenToClient(HeaderhWnd, PointStruct)
      
      HitTestInfo.flags = HHT_ONHEADER Or HHT_ONDIVIDER
      HitTestInfo.pt = PointStruct
      
      Call SendMessage(HeaderhWnd, HDM_HITTEST, 0&, HitTestInfo)

      Select Case NmHdrMsg.code
        Case HDN_ITEMCLICK:       HeaderAction = lvHeaderActionClick
        Case HDN_NM_RCLICK:       HeaderAction = lvHeaderActionRightClick
        Case HDN_DIVIDERDBLCLICK: HeaderAction = lvHeaderActionDividerDoubleClick
        Case HDN_BEGINTRACK:      HeaderAction = lvHeaderActionResizeBegin
        Case HDN_ENDTRACK:        HeaderAction = lvHeaderActionResizeEnd
        Case HDN_ITEMCHANGING:    HeaderAction = lvHeaderActionChanging
        Case HDN_ITEMCHANGED:     HeaderAction = lvHeaderActionChanged
        Case HDN_BEGINDRAG:       HeaderAction = lvHeaderActionDragBegin
        Case HDN_ENDDRAG:         HeaderAction = lvHeaderActionDragEnd
      End Select
      
      If HeaderAction <> 0 Then
            
        On Error Resume Next
        
        Set ListViewControl = RegisteredListViewControls(CStr(ListViewhWnd))
        
        CancelMsg = CallByName(ListViewControl.Parent, ListViewControl.Name & EVENT_SUFFIX, VbCallType.VbMethod, HeaderAction, HitTestInfo.iItem + 1)
        
        On Error GoTo 0
        If CancelMsg Then
          HandleListViewHeaderMsgs = 1
          Exit Function
        End If
        
      End If
      
    End If
  
  End If
  
  HandleListViewHeaderMsgs = CallWindowProc(GetProp(ListViewhWnd, "OrigWindowProc"), ListViewhWnd, msg, wp, lp)
   
End Function

Public Sub ListViewHeaderEventDebugPrint(ByVal ListViewControl As ListView, ByVal Action As lvHeaderActions, ByVal Column As Long)

  Dim msg As String

  Select Case Action
    Case lvHeaderActionClick:               msg = "clicked"
    Case lvHeaderActionRightClick:          msg = "right-clicked"
    Case lvHeaderActionDividerDoubleClick:  msg = "divider dbl-clicked"
    Case lvHeaderActionResizeBegin:         msg = "resize begin"
    Case lvHeaderActionResizeEnd:           msg = "resize end"
    Case lvHeaderActionChanging:            msg = "changing"
    Case lvHeaderActionChanged:             msg = "changed"
    Case lvHeaderActionDragBegin:           msg = "drag begin"
    Case lvHeaderActionDragEnd:             msg = "drag end"
  End Select
    
  Debug.Print ListViewControl.Parent.Name & "." & ListViewControl.Name & ": " & msg & " (col=" & CStr(Column) & ")"

End Sub
