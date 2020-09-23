Attribute VB_Name = "FileHandling"
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - September 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au
Option Explicit
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Function FileExists(sSource As String) As Boolean
    'Thorogh FileExists function
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists = False
        End If
    End If
End Function
Public Function SpecialFolder(ByVal CSIDL As Long) As String
    'Used in this demo to locate the user's Desktop
    Dim r As Long
    Dim sPath As String
    Dim IDL As ITEMIDLIST
    Const NOERROR = 0
    Const MAX_LENGTH = 260
    r = SHGetSpecialFolderLocation(GetDesktopWindow, CSIDL, IDL)
    If r = NOERROR Then
        sPath = Space$(MAX_LENGTH)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        If r Then
            SpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
        End If
    End If
 '0=Desktop
 '2=StartMenu\Programs
 '5=My Documents
 '6=Favorites
 '7=Startup
 '8=Recent
 '9=SendTo
 '11=StartMenu
 '16=Desktop
 '19=Nethood
 '20=Fonts
 '21=ShellNew
 '25=All users\desktop
 '26=Application Data
 '27=PrintHood
 '32=Temporary Internet Files
 '33=Cookies
 '34=History
'Default=0
'Programs=2
'Control Panel=3
'Printers=4
'My Documents=5
'Favourites=6
'StartUp=7
'Recent=8
'SendTo=9
'Recycle Bin=10
'Start Menu=11
'Desktop=16
'My Computer=17
'Network=18
'NetHood=19
'Fonts=20
'Templates=21
'All users \ desktop=25
'Application Data=26
'PrintHood=27
'Temporary Internet Files=32
'Cookies=33
'History=34

End Function
Public Function PathOnly(ByVal filepath As String) As String
    Dim temp As String
    temp = Mid$(filepath, 1, InStrRev(filepath, "\"))
    If Right(temp, 1) = "\" And Len(temp) > 3 Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function

Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function SafeSave(Path As String, Optional ByRef safesavename As String) As String
    'Simple parsing routine to return a unique filename
    Dim mPath As String, mname As String, mTemp As String, mfile As String, mExt As String, m As Integer
    On Error Resume Next
    mPath = Mid$(Path, 1, InStrRev(Path, "\"))
    mname = Mid$(Path, InStrRev(Path, "\") + 1)
    mfile = Left(Mid$(mname, 1, InStrRev(mname, ".")), Len(Mid$(mname, 1, InStrRev(mname, "."))) - 1)
    If mfile = "" Then mfile = mname
    mExt = Mid$(mname, InStrRev(mname, "."))
    mTemp = ""
    Do
        If Not FileExists(mPath + mfile + mTemp + mExt) Then
            SafeSave = mPath + mfile + mTemp + mExt
            safesavename = mfile + mTemp + mExt
            Exit Do
        End If
        m = m + 1
        mTemp = "(" & m & ")"
    Loop
End Function

