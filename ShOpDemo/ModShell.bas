Attribute VB_Name = "ModShell"
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - September 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au
Option Explicit
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
'Actions
Public Const FO_MOVE = &H1
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_RENAME = &H4&
'Flags
Public Const FOF_SILENT = &H4
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_ALLOWUNDO = &H40
Public Sub ShellAction(mSource As String, mDestination As String, mAction As Long, mFlags As Long)
    Dim SHFileOp As SHFILEOPSTRUCT
    mSource = mSource & Chr$(0) & Chr$(0)
    With SHFileOp
        .wFunc = mAction
        .pFrom = mSource
        .pTo = mDestination
        .fFlags = mFlags
    End With
    SHFileOperation SHFileOp
End Sub

