VERSION 5.00
Begin VB.Form frmSource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHFileOperation Demo"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   330
      Left            =   5040
      TabIndex        =   25
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Height          =   3255
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdUp 
         Height          =   330
         Left            =   2400
         Picture         =   "frmSource.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Up One Level"
         Top             =   240
         Width           =   330
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2895
      End
      Begin VB.FileListBox File1 
         Height          =   2820
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdDesktop 
         Height          =   330
         Left            =   2760
         Picture         =   "frmSource.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Desktop"
         Top             =   240
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   5040
      TabIndex        =   18
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   5040
      TabIndex        =   17
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Action"
      Height          =   1695
      Left            =   2400
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
      Begin VB.OptionButton OptAction 
         Caption         =   "Delete"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1000
         Width           =   1455
      End
      Begin VB.OptionButton OptAction 
         Caption         =   "Rename"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton OptAction 
         Caption         =   "Move"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptAction 
         Caption         =   "Copy"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   680
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Flags"
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
      Begin VB.CheckBox ChFlag 
         Caption         =   "No Confirmation"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Show Progress"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Silent"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Rename on Collision"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Allow Undo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   6255
      Begin VB.CommandButton cmdDest 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   5
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtDestination 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtSource 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "Destination:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Source:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - September 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au
Option Explicit
'Action to perform
Dim m_Action As Long

Private Sub cmdAbout_Click()
Dim temp As String
temp = "This is a simple demo of using the API to manage files. It is by no means"
temp = temp & " a 'File Manager' and Im sure if you fiddle with the paths enough, you'll"
temp = temp & " get around the simple error handling I have employed and cause some"
temp = temp & " errors. Choose some files that are dispensible in order to test this"
temp = temp & " demo. It was put together in order to answer questions from many coders"
temp = temp & " new to VB in the Discussion Forum. I hope it is of some help to those"
temp = temp & " folks."
MsgBox temp
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDesktop_Click()
    'Navigate to Desktop
    Dim temp As String
    temp = SpecialFolder(0)
    Drive1.Drive = Left(temp, 3)
    Dir1.Path = temp
    Dir1.SetFocus
End Sub

Private Sub cmdDest_Click()
    frmDest.Show vbModal, Me
End Sub

Private Sub cmdOK_Click()
    Dim z As Long, IsSpecial As Boolean, FolderDelete As Boolean
    'First some checks
    
    'Not deleting so need a destination
    If m_Action <> 3 Then
        If Len(txtDestination.Text) = 0 Then
            MsgBox "You need to specify a destination"
            cmdDest_Click
            Exit Sub
        End If
    Else
        If GetAttr(txtSource.Text) = vbDirectory Then FolderDelete = True
    End If
    'If we're performing an action on a folder remove traling backslash
    If Right(txtSource.Text, 1) = "\" Then txtSource.Text = Left(txtSource.Text, Len(txtSource.Text) - 1)
    'Does the file we're acting on exist?
    If Not FileExists(txtSource.Text) Then
        MsgBox "File not found"
        Exit Sub
    End If
    'Dont mess with drives!
    If Len(txtSource.Text) < 4 Then
        MsgBox "Not a good idea to perform actions on drives!"
        txtSource.SetFocus
        Exit Sub
    End If
    'Dont mess with Special folders
    For z = 0 To 40
        If LCase(SpecialFolder(z)) = LCase(txtSource.Text) Then
            IsSpecial = True
            Exit For
        End If
    Next
    If IsSpecial Then
        MsgBox "Not a good idea to perform actions on Special Folders!"
        txtSource.SetFocus
        Exit Sub
    End If
    'Make sure paths match when renaming
    If m_Action = 4 Then
        If LCase(PathOnly(txtSource.Text)) <> LCase(PathOnly(txtDestination.Text)) Then
            MsgBox "When renaming, the paths must be the same - only the name changes."
            txtDestination.SetFocus
            Exit Sub
        End If
    End If
    'Ok, do it!
    ShellAction txtSource.Text, txtDestination.Text, m_Action, GetFlags
    'If we deleted a folder go up one level
    If FolderDelete Then cmdUp_Click
    'Refresh the view
    Dir1.Refresh
    File1.Refresh
    'If we didn't delete then txtDestination.Text should now exist
    If m_Action <> 3 Then
        If Not FileExists(txtDestination.Text) Then GoTo woops
    Else
    'If we did delete then txtSource.Text should not exist
        If FileExists(txtSource.Text) Then GoTo woops
    End If
    Exit Sub
woops:
    MsgBox "Woops! Something went wrong!"
End Sub

Private Sub cmdUp_Click()
    Dir1.Path = PathOnly(Dir1.Path)
    Dir1.SetFocus
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    File1_Click
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    txtSource.Text = IIf(Right(File1.Path, 1) = "\", File1.Path, File1.Path & "\") & File1.FileName
End Sub

Private Sub Form_Load()
    m_Action = 1
    File1_Click
End Sub

Private Function GetFlags() As Long
    'Create SHFileOperation flags according to checkboxes
    Dim mFlags As Long, z As Long
    For z = 0 To ChFlag.Count - 1
        If ChFlag(z).Value = 1 Then
            Select Case z
                Case 0
                    mFlags = mFlags Or FOF_ALLOWUNDO
                Case 1
                    mFlags = mFlags Or FOF_RENAMEONCOLLISION
                Case 2
                    mFlags = mFlags Or FOF_SILENT
                Case 3
                    mFlags = mFlags Or FOF_SIMPLEPROGRESS
                Case 4
                    mFlags = mFlags Or FOF_NOCONFIRMATION
            End Select
        End If
    Next
    GetFlags = mFlags
End Function

Private Sub OptAction_Click(Index As Integer)
    'Adjust the SHFileOperation action variable according to the Option selected
    m_Action = Index
End Sub
