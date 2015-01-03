Attribute VB_Name = "mdlGetUserDirPath"
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long ' Ret: 0=success
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Const MAX_PATH = 260
Public Const CSIDL_DESKTOP = &H0

'Returns a user directory path
Public Function GetShellFolderPath(ByVal CSIDL As Long) As String
    Dim pID As Long
    Dim sTmp As String
    
    If SHGetSpecialFolderLocation(0&, CSIDL, pID) = 0& Then
        sTmp = String(MAX_PATH + 2, 0)
        If SHGetPathFromIDList(ByVal pID, sTmp) <> 0& Then
            GetShellFolderPath = Left$(sTmp, InStr(1, sTmp, vbNullChar) - 1)
        End If
    End If
    If pID <> 0& Then GlobalFree pID
End Function

