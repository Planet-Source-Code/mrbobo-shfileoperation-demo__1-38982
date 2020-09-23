VERSION 5.00
Begin VB.Form frmDest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Destination"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdNewFolder 
         Height          =   330
         Left            =   2760
         Picture         =   "frmDest.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "New Folder"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmdUp 
         Height          =   330
         Left            =   2040
         Picture         =   "frmDest.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Up One Level"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmdDesktop 
         Height          =   330
         Left            =   2400
         Picture         =   "frmDest.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Desktop"
         Top             =   240
         Width           =   330
      End
      Begin VB.DirListBox Dir1 
         Height          =   2790
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2895
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "frmDest"
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

Private Sub cmdNewFolder_Click()
    Dim temp As String, mSafe As String, mSafeName As String, mRoot As String
    Dir1.SetFocus
    mRoot = IIf(Right(Dir1.Path, 1) <> "\", Dir1.Path & "\", Dir1.Path)
    mSafe = SafeSave(mRoot & "New Folder", mSafeName)
    temp = InputBox("Enter a name for your new folder", , mSafeName)
    If temp = "" Then Exit Sub
    temp = Replace(temp, "\", "")
    MkDir SafeSave(mRoot & temp)
    Dir1.Path = mRoot & temp
End Sub

Private Sub cmdOK_Click()
    'Return the selected path to frmSource
    frmSource.txtDestination.Text = IIf(Right(Dir1.Path, 1) = "\", Dir1.Path, Dir1.Path & "\") & IIf(frmSource.File1.FileName <> "", frmSource.File1.FileName, FileOnly(Dir1.List(Dir1.ListIndex)))
    Unload Me
End Sub

Private Sub cmdUp_Click()
    Dir1.Path = PathOnly(Dir1.Path)
    Dir1.SetFocus
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
