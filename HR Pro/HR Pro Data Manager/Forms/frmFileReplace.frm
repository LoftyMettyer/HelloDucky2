VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileReplace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirm File Replace"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileReplace.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picNew 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   500
      Left            =   1300
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2100
      Width           =   500
   End
   Begin VB.PictureBox picExist 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000F&
      Height          =   500
      Left            =   1300
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1100
      Width           =   500
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   240
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileReplace.frx":030A
            Key             =   ".JPG"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileReplace.frx":0978
            Key             =   ".GIF"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileReplace.frx":0FE6
            Key             =   ".BMP"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "&No"
      Height          =   400
      Left            =   4800
      TabIndex        =   1
      Top             =   2800
      Width           =   1200
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   400
      Left            =   3400
      TabIndex        =   0
      Top             =   2800
      Width           =   1200
   End
   Begin VB.Label lblNewDetails 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "modified on "
      Height          =   195
      Left            =   1995
      TabIndex        =   8
      Top             =   2355
      Width           =   1050
   End
   Begin VB.Label lblNewSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "263KB"
      Height          =   195
      Left            =   1995
      TabIndex        =   7
      Top             =   2100
      Width           =   675
   End
   Begin VB.Image imgNew 
      Height          =   495
      Left            =   1300
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label lblWith 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "with this one ?"
      Height          =   195
      Left            =   1000
      TabIndex        =   6
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label lblExistDetails 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "modified on "
      Height          =   195
      Left            =   2000
      TabIndex        =   5
      Top             =   1350
      Width           =   870
   End
   Begin VB.Label lblExistSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "263KB"
      Height          =   195
      Left            =   2000
      TabIndex        =   4
      Top             =   1100
      Width           =   450
   End
   Begin VB.Image imgExist 
      Height          =   495
      Left            =   1300
      Top             =   1100
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Would you like to replace the existing file"
      Height          =   195
      Left            =   1000
      TabIndex        =   3
      Top             =   800
      Width           =   2940
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This folder already contains a file called '"
      Height          =   195
      Left            =   1005
      TabIndex        =   2
      Top             =   195
      Width           =   2940
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   200
      Picture         =   "frmFileReplace.frx":1654
      Top             =   200
      Width           =   480
   End
End
Attribute VB_Name = "frmFileReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FileInfo As typSHFILEINFO
Private mfReplace As Boolean

Public Sub Initialise(sExistingFilePath As String, sNewFilePath As String, sName As String)
  Dim fImageTypeFile As Boolean
  Dim iIconIndex As Integer
  Dim lngIconHandle As Long
  Dim lngApplicationHandle As Long
  Dim sFileExtension As String
  Dim objFileSystem As New FileSystemObject
  Dim objFile As File
  Dim objIcon As Long
  
  ' JDM - 01/03/2005 - Some functions that call this routine are messy and pass in \\ instead of \
  sExistingFilePath = Replace(sExistingFilePath, "\\", "\")
  sNewFilePath = Replace(sNewFilePath, "\\", "\")
  
  ' Configure the screen message.
  lblTitle.Caption = "This folder already contains a file called '" & sName & "'."
   
  ' Get the existing file's size and date modified.
  Set objFile = objFileSystem.GetFile(sExistingFilePath)
  lblExistSize.Caption = CLng(objFile.Size / 1000) & "KB"
  lblExistDetails.Caption = "modified on " & objFile.DateLastModified

  ' Get the new file's size and date modified.
  Set objFile = objFileSystem.GetFile(sNewFilePath)
  lblNewSize.Caption = CLng(objFile.Size / 1000) & "KB"
  lblNewDetails.Caption = "modified on " & objFile.DateLastModified

  Set objFile = Nothing
  Set objFileSystem = Nothing
  
  ' Check if the file is a graphic or not.
  sFileExtension = UCase(Right(sName, 4))
  
  fImageTypeFile = (sFileExtension = ".JPG") Or _
    (sFileExtension = ".GIF") Or _
    (sFileExtension = ".BMP")
  
  imgExist.Visible = fImageTypeFile
  imgNew.Visible = fImageTypeFile
  picExist.Visible = Not fImageTypeFile
  picNew.Visible = Not fImageTypeFile
  
  ' Display the specified icon for the graphic file type.
  If fImageTypeFile Then
    Set imgExist.Picture = imlMain.ListImages(sFileExtension).Picture
    Set imgNew.Picture = imlMain.ListImages(sFileExtension).Picture
  Else
 
    objIcon = SHGetFileInfo(sExistingFilePath, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    If objIcon <> 0 Then
      With picExist
        .BackColor = vbButtonFace
        .Height = 15 * 32
        .Width = 15 * 32
        .ScaleHeight = 15 * 32
        .ScaleWidth = 15 * 32
        .Picture = LoadPicture("")
        .AutoRedraw = True
        objIcon = ImageList_Draw(objIcon, FileInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
    End If
  
    objIcon = SHGetFileInfo(sNewFilePath, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    If objIcon <> 0 Then
      With picNew
        .BackColor = vbButtonFace
        .Height = 15 * 32
        .Width = 15 * 32
        .ScaleHeight = 15 * 32
        .ScaleWidth = 15 * 32
        .Picture = LoadPicture("")
        .AutoRedraw = True
        objIcon = ImageList_Draw(objIcon, FileInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
    End If
  
  End If
  
End Sub

Public Property Get Replaced() As Boolean
  Replaced = mfReplace
End Property

Private Sub cmdNo_Click()
  mfReplace = False
  Me.Hide

End Sub

Private Sub cmdYes_Click()
  mfReplace = True
  Me.Hide

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub



