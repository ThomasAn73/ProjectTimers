Attribute VB_Name = "mdlGetDirectory"
'Alternative method of getting a directory
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Public Function DoSelectDir() As String
'Opens a Treeview control that displays the directories in a computer

   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo

   szTitle = "Timer will be saved in:"
   With tBrowseInfo
      .hWndOwner = frmTimerEdit.hWnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
   End With

   lpIDList = SHBrowseForFolder(tBrowseInfo)

   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      DoSelectDir = sBuffer
   End If
End Function

