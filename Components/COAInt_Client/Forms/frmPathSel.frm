VERSION 5.00
Begin VB.Form frmPathSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Photographs"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1050
   Icon            =   "frmPathSel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPath 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   150
      Locked          =   -1  'True
      MaxLength       =   240
      TabIndex        =   4
      Top             =   3925
      Width           =   3510
   End
   Begin VB.DirListBox dirDirs 
      Height          =   1440
      Left            =   150
      TabIndex        =   1
      Top             =   2035
      Width           =   3510
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   1270
      Width           =   3510
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1110
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   4390
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2460
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   4390
      Width           =   1200
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path Selected :"
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   3625
      Width           =   1305
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPathSel.frx":000C
      Height          =   780
      Left            =   150
      TabIndex        =   7
      Top             =   150
      Width           =   3555
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDirs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folders :"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   1720
      Width           =   750
   End
   Begin VB.Label lblDrives 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drives :"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   985
      Width           =   690
   End
End
Attribute VB_Name = "frmPathSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private msInitDir As String

Public Property Get SelectedFolder() As String
  SelectedFolder = txtPath.Text
End Property

Private Sub cmdCancel_Click()
  txtPath.Text = msInitDir
  Me.Hide
  DoEvents
End Sub

Private Sub cmdOK_Click()
  ' Store the selected path in the registry.
  If txtPath.Text <> vbNullString Then
    Me.Hide
    DoEvents
  End If

End Sub

Private Sub dirDirs_Change()
  
  On Error Resume Next

  ChDir dirDirs.Path
  If Err = 0 Then
    txtPath.Text = dirDirs.Path
    drvDrives.Drive = Left$(dirDirs.Path, 2)
  Else
    Err = 0
  End If

End Sub

Private Sub drvDrives_Change()
  
  On Error GoTo ErrTrap
  
  Static strolddrive As String
  
  dirDirs.Path = drvDrives.Drive
  strolddrive = dirDirs.Path
  
  Exit Sub

ErrTrap:
  Select Case Err.Number
    Case 68
      MsgBox "No disk in drive or drive not ready", vbExclamation + vbOKOnly, "Error"
      drvDrives.Drive = strolddrive
    Case Else
      MsgBox "Error : " & Err.Number & Chr(10) & "Descr : " & Err.Description
      drvDrives.Drive = strolddrive
  End Select

End Sub

Public Function Initialise(psInitPath As String) As Boolean

  Dim fOk As Boolean
  Dim fUNC As Boolean
  
  On Error GoTo ErrorTrap
  
  fOk = True
  
  msInitDir = psInitPath
  
  fUNC = (Left(psInitPath, 2) = "\\")
  
  If fUNC Then
    dirDirs.Path = psInitPath
    txtPath.Text = dirDirs.Path
  Else
    dirDirs.Path = psInitPath
    txtPath.Text = dirDirs.Path
    drvDrives.Drive = Left$(dirDirs.Path, 2)
    drvDrives_Change
  End If
  
TidyUpAndExit:
  Initialise = fOk
  Exit Function
  
ErrorTrap:
  fOk = False
  Resume TidyUpAndExit
  
End Function

Private Sub Form_Initialize()
  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub

