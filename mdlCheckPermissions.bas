Attribute VB_Name = "mdlCheckPermissions"
Option Explicit

' Desired access rights constants
Private Const MAXIMUM_ALLOWED As Long = &H2000000
Private Const DELETE As Long = &H10000
Private Const READ_CONTROL As Long = &H20000
Private Const WRITE_DAC As Long = &H40000
Private Const WRITE_OWNER As Long = &H80000
Private Const SYNCHRONIZE As Long = &H100000

Private Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000

Private Const FILE_READ_DATA As Long = &H1 ' file & pipe
Private Const FILE_LIST_DIRECTORY As Long = &H1 ' directory
Private Const FILE_ADD_FILE As Long = &H2 ' directory
Private Const FILE_WRITE_DATA As Long = &H2 ' file & pipe
Private Const FILE_CREATE_PIPE_INSTANCE As Long = &H4 ' named pipe
Private Const FILE_ADD_SUBDIRECTORY As Long = &H4 ' directory
Private Const FILE_APPEND_DATA As Long = &H4 ' file
Private Const FILE_READ_EA As Long = &H8 ' file & directory
Private Const FILE_READ_PROPERTIES As Long = FILE_READ_EA
Private Const FILE_WRITE_EA As Long = &H10 ' file & directory
Private Const FILE_WRITE_PROPERTIES As Long = FILE_WRITE_EA
Private Const FILE_EXECUTE As Long = &H20 ' file
Private Const FILE_TRAVERSE As Long = &H20 ' directory
Private Const FILE_DELETE_CHILD As Long = &H40 ' directory
Private Const FILE_READ_ATTRIBUTES As Long = &H80 ' all
Private Const FILE_WRITE_ATTRIBUTES As Long = &H100 ' all

Private Const FILE_GENERIC_READ As Long = (STANDARD_RIGHTS_READ Or FILE_READ_DATA Or FILE_READ_ATTRIBUTES Or FILE_READ_EA Or SYNCHRONIZE)
Private Const FILE_GENERIC_WRITE As Long = (STANDARD_RIGHTS_WRITE Or FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or FILE_APPEND_DATA Or SYNCHRONIZE)
Private Const FILE_GENERIC_EXECUTE As Long = (STANDARD_RIGHTS_EXECUTE Or FILE_READ_ATTRIBUTES Or FILE_EXECUTE Or SYNCHRONIZE)
Private Const FILE_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1FF&)

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_EXECUTE As Long = &H20000000
Private Const GENERIC_ALL As Long = &H10000000

' Types, constants and functions
' to work with access rights
Private Const OWNER_SECURITY_INFORMATION As Long = &H1
Private Const GROUP_SECURITY_INFORMATION As Long = &H2
Private Const DACL_SECURITY_INFORMATION As Long = &H4
Private Const ERROR_INSUFFICIENT_BUFFER = 122&
Private Const MAX_PATH = 255
Private Const TOKEN_QUERY As Long = 8
Private Const SecurityImpersonation As Integer = 3
Private Const ANYSIZE_ARRAY = 1

Private Type GENERIC_MAPPING
    GenericRead As Long
    GenericWrite As Long
    GenericExecute As Long
    GenericAll As Long
End Type

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type PRIVILEGE_SET
    PrivilegeCount As Long
    Control As Long
    Privilege(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare Function AccessCheck Lib "advapi32.dll" (pSecurityDescriptor As Byte, ByVal ClientToken As Long, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, GrantedAccess As Long, Status As Long) As Long
Private Declare Function ImpersonateSelf Lib "advapi32.dll" (ByVal ImpersonationLevel As Integer) As Long
Private Declare Function RevertToSelf Lib "advapi32.dll" () As Long
Private Declare Sub MapGenericMask Lib "advapi32.dll" (AccessMask As Long, GenericMapping As GENERIC_MAPPING)
Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As Any, pOwner As Long, lpbOwnerDefaulted As Long) As Long
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' Types, constants and functions for OS version detection
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' Constant and function for detection of support
' of access rights by file system
Private Const FS_PERSISTENT_ACLS As Long = &H8

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Function CheckPermissions(ForThisFile As String) As Variant
Dim Result
Dim DesiredAccess

Result = Array("", vbTrue, vbTrue)
Result(0) = GetUserName(ForThisFile)

CheckFileAccess ForThisFile, DesiredAccess
Result(1) = CheckFileAccess(ForThisFile, FILE_GENERIC_READ) = FILE_GENERIC_READ
Result(2) = CheckFileAccess(ForThisFile, FILE_GENERIC_WRITE) = FILE_GENERIC_WRITE

CheckPermissions = Result
End Function

' CheckFileAccess function checks access rights to given file.
' DesiredAccess - bitmask of desired access rights.
' The function returns bitmask, which contains those bits of desired bitmask,
' which correspond with existing access rights.
Private Function CheckFileAccess(Filename As String, ByVal DesiredAccess As Long) As Long
                
    Dim r As Long, SecDesc() As Byte, SDSize As Long, hToken As Long
    Dim PrivSet As PRIVILEGE_SET, GenMap As GENERIC_MAPPING
    Dim Volume As String, FSFlags As Long

    ' Checking OS type
    ' Rights not supported. Returning -1.
    If Not IsNT() Then CheckFileAccess = -1: Exit Function

    ' Checking access rights support by file system
    If Left$(Filename, 2) = "\\" Then
        ' Path in UNC format. Extracting share name from it
        r = InStr(3, Filename, "\")
        If r = 0 Then Volume = Filename & "\" Else Volume = Left$(Filename, r)
    ElseIf Mid$(Filename, 2, 2) = ":\" Then
        ' Path begins with drive letter
        Volume = Left$(Filename, 3)
    'Else
    ' If path not set, we are leaving Volume blank.
    ' It retutns information about current drive.
    End If
    
    ' Getting information about drive
    GetVolumeInformation Volume, vbNullString, 0, ByVal 0&, ByVal 0&, FSFlags, vbNullString, 0
                    
    ' Rights not supported. Returning -1.
    If (FSFlags And FS_PERSISTENT_ACLS) = 0 Then CheckFileAccess = -1: Exit Function
    
    ' Determination of buffer size
    GetFileSecurity Filename, OWNER_SECURITY_INFORMATION Or GROUP_SECURITY_INFORMATION Or DACL_SECURITY_INFORMATION, 0, 0, SDSize
            
    ' Rights not supported. Returning -1.
    If Err.LastDllError <> 122 Then CheckFileAccess = -1: Exit Function
    If SDSize = 0 Then Exit Function
    ' Buffer allocation
    ReDim SecDesc(1 To SDSize)
    ' Once more call of function
    ' to obtain Security Descriptor
    
    ' Error. We must return no access rights.
    If GetFileSecurity(Filename, OWNER_SECURITY_INFORMATION Or GROUP_SECURITY_INFORMATION Or DACL_SECURITY_INFORMATION, SecDesc(1), SDSize, SDSize) = 0 Then Exit Function
    
    ' Adding Impersonation Token for thread
    ImpersonateSelf SecurityImpersonation
    
    ' Opening of Token of current thread
    OpenThreadToken GetCurrentThread(), TOKEN_QUERY, 0, hToken
    
    If hToken <> 0 Then
    ' Filling GenericMask type
        GenMap.GenericRead = FILE_GENERIC_READ
        GenMap.GenericWrite = FILE_GENERIC_WRITE
        GenMap.GenericExecute = FILE_GENERIC_EXECUTE
        GenMap.GenericAll = FILE_ALL_ACCESS
        ' Conversion of generic rights to specific file access rights
        MapGenericMask DesiredAccess, GenMap
        ' Checking access
        AccessCheck SecDesc(1), hToken, DesiredAccess, GenMap, PrivSet, Len(PrivSet), CheckFileAccess, r
        CloseHandle hToken
    End If
    
    ' Deleting Impersonation Token
    RevertToSelf
    
End Function


' IsNT() function returns True, if the program works
' in Windows NT/Windows 2000/Windows XP operating system, and False otherwise.
Private Function IsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = Len(OSVer)
    GetVersionEx OSVer
    IsNT = (OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Private Function GetUserName(ForThisFile As String) As String
    Dim szfilename As String   ' File name to retrieve the owner for
    Dim bSuccess As Long       ' Status variable
    Dim sizeSD As Long         ' Buffer size to store Owner's SID
    Dim pOwner As Long         ' Pointer to the Owner's SID
    Dim name As String         ' Name of the file owner
    Dim domain_name As String  ' Name of the first domain for the owner
    Dim name_len As Long       ' Required length for the owner name
    Dim domain_len As Long     ' Required length for the domain name
    Dim sdBuf() As Byte        ' Buffer for Security Descriptor
    Dim nLength As Long        ' Length of the Windows Directory
    Dim deUse As Long          ' Pointer to a SID_NAME_USE enumerated
    
    ' Initialize some required variables.
    bSuccess = 0
    name = ""
    domain_name = ""
    name_len = 0
    domain_len = 0
    pOwner = 0
    szfilename = ForThisFile
    GetUserName = "Unknown"
    
    ' Call GetFileSecurity the first time to obtain the size of the
    ' buffer required for the Security Descriptor.
    bSuccess = GetFileSecurity(szfilename, OWNER_SECURITY_INFORMATION, 0, 0&, sizeSD)
    If (bSuccess = 0) And (Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER) Then Exit Function 'MsgBox "GetLastError returned  : " & Err.LastDllError
    
    ' Create a buffer of the required size and call GetFileSecurity again.
    ReDim sdBuf(0 To sizeSD - 1) As Byte
    
    ' Fill the buffer with the security descriptor of the object specified
    ' by the szfilename parameter. The calling process must have the right
    ' to view the specified aspects of the object's security status.
    
    bSuccess = GetFileSecurity(szfilename, OWNER_SECURITY_INFORMATION, sdBuf(0), sizeSD, sizeSD)
    If (bSuccess <> 0) Then
        ' Obtain the owner's SID from the Security Descriptor.
        bSuccess = GetSecurityDescriptorOwner(sdBuf(0), pOwner, 0&)
        If (bSuccess = 0) Then Exit Function 'MsgBox "GetLastError returned : " & Err.LastDllError
        
        ' Retrieve the name of the account and the name of the first
        ' domain on which this SID is found.  Passes in the Owner's SID
        ' obtained previously.  Call LookupAccountSid twice, the first time
        ' to obtain the required size of the owner and domain names.
        bSuccess = LookupAccountSid(vbNullString, pOwner, name, name_len, domain_name, domain_len, deUse)
        If (bSuccess = 0) And (Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER) Then Exit Function 'MsgBox "GetLastError returned : " & Err.LastDllError
        
        '  Allocate the required space in the name and domain_name string
        '  variables. Allocate 1 byte less to avoid the appended NULL character.
        name = Space(name_len - 1)
        domain_name = Space(domain_len - 1)
        
        '  Call LookupAccountSid again to actually fill in the name of the owner
        '  and the first domain.
        bSuccess = LookupAccountSid(vbNullString, pOwner, name, name_len, domain_name, domain_len, deUse)
        If bSuccess = 0 Then Exit Function 'MsgBox "GetLastError returned : " & Err.LastDllError
        
        GetUserName = name
    End If

End Function


