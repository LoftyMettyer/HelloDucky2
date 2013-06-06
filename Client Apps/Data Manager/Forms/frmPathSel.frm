VERSION 5.00
Begin VB.Form frmPathSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Photographs"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   3990
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
      Top             =   3960
      Width           =   3735
   End
   Begin VB.DirListBox dirDirs 
      Height          =   1440
      Left            =   150
      TabIndex        =   1
      Top             =   2070
      Width           =   3735
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   1305
      Width           =   3735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1425
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   4425
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2685
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   4425
      Width           =   1200
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path Selected :"
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   3660
      Width           =   1770
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPathSel.frx":000C
      Height          =   855
      Left            =   150
      TabIndex        =   7
      Top             =   150
      Width           =   3795
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDirs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folders :"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   1755
      Width           =   630
   End
   Begin VB.Label lblDrives 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drives :"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   1020
      Width           =   825
   End
End
Attribute VB_Name = "frmPathSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const giSELECTION_PHOTOPATH = 1
Const giSELECTION_OLEPATH = 2
Const giSELECTION_CRYSTALPATH = 4
Const giSELECTION_DOCUMENTSPATH = 8
Const giSELECTION_LOCALOLEPATH = 16

Private miSelectionType As Integer
Private mfQuietMode As Boolean

Private Sub cmdCancel_Click()
  
  If Not mfQuietMode Then
    If (COAMsgBox("Some features may not be available if a path is not entered." & vbCrLf & _
      "Are you sure you wish to cancel ?", vbQuestion + vbYesNo + vbDefaultButton2, "Path Selection") = vbNo) Then
      Exit Sub
    End If
  End If
  
  Unload Me
  Screen.MousePointer = vbHourglass
  
End Sub

Private Sub cmdOK_Click()
  ' Store the selected path in the registry.
  If txtPath.Text <> vbNullString Then
    Select Case miSelectionType
      
      Case giSELECTION_PHOTOPATH
        SavePCSetting "Datapaths", "photopath_" & gsDatabaseName, txtPath.Text
        gsPhotoPath = txtPath.Text
      
      Case giSELECTION_OLEPATH
        SavePCSetting "Datapaths", "olepath_" & gsDatabaseName, txtPath.Text
        gsOLEPath = txtPath.Text
        
      Case giSELECTION_CRYSTALPATH
        SavePCSetting "Datapaths", "crystalpath_" & gsDatabaseName, txtPath.Text
        gsCrystalPath = txtPath.Text
        
      Case giSELECTION_DOCUMENTSPATH
        SavePCSetting "Datapaths", "documentspath_" & gsDatabaseName, txtPath.Text
        gsDocumentsPath = txtPath.Text
        
      Case giSELECTION_LOCALOLEPATH
        SavePCSetting "Datapaths", "localolepath_" & gsDatabaseName, txtPath.Text
        gsLocalOLEPath = txtPath.Text
        
    End Select
    
    If Not mfQuietMode Then
      COAMsgBox "The path has been stored successfully.", vbInformation + vbOKOnly, "Path Selection"
    End If
    
    Unload Me
    DoEvents
    Screen.MousePointer = vbHourglass
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
  Screen.MousePointer = vbHourglass
  dirDirs.Path = drvDrives.Drive
  strolddrive = dirDirs.Path
  Screen.MousePointer = vbArrow
  Exit Sub
  
ErrTrap:
  Select Case Err.Number
    Case 68
      COAMsgBox "No disk in drive or drive not ready", vbExclamation + vbOKOnly, "Error"
      drvDrives.Drive = strolddrive
    Case Else
      COAMsgBox "Error : " & Err.Number & Chr(10) & "Descr : " & Err.Description
      drvDrives.Drive = strolddrive
  End Select
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  dirDirs.Path = Left$(App.Path, 3)
  txtPath.Text = dirDirs.Path
  drvDrives.Drive = Left$(dirDirs.Path, 2)
  drvDrives_Change
  Screen.MousePointer = vbArrow
  Err = 0
  
End Sub




Public Property Let SelectionType(ByVal piNewValue As Integer)
  On Error GoTo Err_Trap
  
  ' Set the path selection type flag.
  miSelectionType = piNewValue

  'TM20011107 Fault 3050 - Need to check that the directory exists.
  'before setting the path of the directory box.

  Select Case miSelectionType
    Case giSELECTION_PHOTOPATH
      Me.Caption = "Photographs"
      
      If Not mfQuietMode Then
        lblPrompt.Caption = "The folder where the OpenHR photographs (non-linked) are stored is invalid." & vbCrLf & _
                            "Please select a valid folder on the server."
      Else
        lblPrompt.Caption = "You have opted to change the folder where the OpenHR photographs (non-linked) are stored." & vbCrLf & _
                            "Please select a valid folder on the server."
        
        If Dir(frmConfiguration.txtPhotoPath.Text & IIf(Right(frmConfiguration.txtPhotoPath.Text, 1) = "\", "*.*", "\*.*"), vbDirectory) = vbNullString Then
          COAMsgBox "The folder where the OpenHR photographs (non-linked) are stored is invalid." & vbCrLf & _
                 "Please select a valid folder on the server.", _
                  vbExclamation + vbOKOnly, App.Title
        Else
          Me.dirDirs.Path = frmConfiguration.txtPhotoPath.Text
        End If

      End If
      
    Case giSELECTION_OLEPATH
      Me.Caption = "Server OLE Documents"
      
      If Not mfQuietMode Then
        lblPrompt.Caption = "The folder where the OpenHR OLE documents are stored is invalid." & vbCrLf & _
                            "Please select a valid folder on the server."
      Else
        lblPrompt.Caption = "You have opted to change the folder where the OpenHR OLE documents are stored." & vbCrLf & _
                            "Please select a valid folder on the server."
        
        If Dir(frmConfiguration.txtOLEPath.Text & IIf(Right(frmConfiguration.txtOLEPath.Text, 1) = "\", "*.*", "\*.*"), vbDirectory) = vbNullString Then
          COAMsgBox "The folder where the OpenHR OLE documents are stored is invalid." & vbCrLf & _
                            "Please select a valid folder on the server.", _
                            vbExclamation + vbOKOnly, App.Title
        Else
          Me.dirDirs.Path = frmConfiguration.txtOLEPath.Text
        End If
      
      End If
      
    Case giSELECTION_CRYSTALPATH
      Me.Caption = "Crystal Report Documents"
      
      If Not mfQuietMode Then
        lblPrompt.Caption = "The folder where the OpenHR Crystal Report documents are stored is invalid." & vbCrLf & _
                            "Please select a valid folder on the server."
      Else
        lblPrompt.Caption = "You have opted to change the folder where the OpenHR Crystal Report documents are stored." & vbCrLf & _
                            "Please select a valid folder on the server."
      
        If Dir(frmConfiguration.txtCrystalPath.Text & IIf(Right(frmConfiguration.txtCrystalPath.Text, 1) = "\", "*.*", "\*.*"), vbDirectory) = vbNullString Then
          COAMsgBox "The folder where the OpenHR Crystal Report documents are stored is invalid." & vbCrLf & _
                            "Please select a valid folder on the server.", _
                            vbExclamation + vbOKOnly, App.Title
        Else
          Me.dirDirs.Path = frmConfiguration.txtCrystalPath.Text
        End If
      
      End If
      
    Case giSELECTION_DOCUMENTSPATH
      Me.Caption = "Document default output"
      
      If Not mfQuietMode Then
        lblPrompt.Caption = "The folder where the OpenHR documents are stored is invalid." & vbCrLf & _
                            "Please select a valid folder on the server."
      Else
        lblPrompt.Caption = "You have opted to change the folder where the OpenHR documents are stored." & vbCrLf & _
                            "Please select a valid folder on the server."
        
        If Dir(frmConfiguration.txtDocumentsPath.Text & IIf(Right(frmConfiguration.txtDocumentsPath.Text, 1) = "\", "*.*", "\*.*"), vbDirectory) = vbNullString Then
          COAMsgBox "The folder where the OpenHR documents are stored is invalid." & vbCrLf & _
                            "Please select a valid folder on the server.", _
                            vbExclamation + vbOKOnly, App.Title
        Else
          Me.dirDirs.Path = frmConfiguration.txtDocumentsPath.Text
        End If
        
      End If
      
    Case giSELECTION_LOCALOLEPATH
      Me.Caption = "Local OLE Documents"
      
      If Not mfQuietMode Then
        lblPrompt.Caption = "The folder where the Local OpenHR OLE documents are stored is invalid." & vbCrLf & _
                            "Please select a valid folder."
      Else
        lblPrompt.Caption = "You have opted to change the folder where the Local OpenHR OLE documents are stored." & vbCrLf & _
                            "Please select a valid folder."
        
        If Dir(frmConfiguration.txtLocalOLEPath.Text & IIf(Right(frmConfiguration.txtLocalOLEPath.Text, 1) = "\", "*.*", "\*.*"), vbDirectory) = vbNullString Then
          COAMsgBox "The folder where the Local OpenHR OLE documents are stored is invalid." & vbCrLf & _
                            "Please select a valid folder.", _
                            vbExclamation + vbOKOnly, App.Title
        Else
          Me.dirDirs.Path = frmConfiguration.txtLocalOLEPath.Text
        End If
      
      End If
      
  End Select
  
  Exit Property

Err_Trap:
  Select Case Err.Number
    ' JPD20030116 Fault 4412
    Case 52, 53, 75
      Resume Next
    Case Else
  End Select
  
End Property

Public Property Let QuietMode(ByVal pfNewValue As Boolean)
  
  mfQuietMode = pfNewValue
  
End Property



Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub



