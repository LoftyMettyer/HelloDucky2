VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSelectEmbedded 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select..."
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1150
   Icon            =   "frmSelectEmbedded.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleMode       =   0  'User
   ScaleWidth      =   4235.572
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLinkEmbed 
      Caption         =   "    Linked / Embedded File"
      Height          =   1920
      Left            =   1620
      TabIndex        =   7
      Top             =   180
      Width           =   2370
      Begin VB.PictureBox picEmbedFile 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   105
         Picture         =   "frmSelectEmbedded.frx":000C
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   -15
         Width           =   225
      End
      Begin VB.PictureBox picDocumentShell 
         BackColor       =   &H00FFFFFF&
         Height          =   1410
         Left            =   195
         ScaleHeight     =   1350
         ScaleWidth      =   1920
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   255
         Width           =   1980
         Begin VB.PictureBox picViewIcon 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   675
            ScaleHeight     =   555
            ScaleWidth      =   585
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   75
            Width           =   585
         End
         Begin VB.Label lblEmbeddedFileName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   510
            Left            =   135
            TabIndex        =   11
            Top             =   690
            Width           =   1665
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picLinkedFile 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   105
         Picture         =   "frmSelectEmbedded.frx":0596
         ScaleHeight     =   225
         ScaleWidth      =   285
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   285
      End
      Begin MSComCtl2.Animation aniSearching 
         Height          =   765
         Left            =   750
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   945
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1349
         _Version        =   393216
         AutoPlay        =   -1  'True
         Center          =   -1  'True
         FullWidth       =   58
         FullHeight      =   51
      End
      Begin VB.Image imgPhoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2145
         Left            =   120
         Stretch         =   -1  'True
         Top             =   255
         Width           =   2130
      End
   End
   Begin VB.CommandButton cmdAddEmbed 
      Caption         =   "E&mbed..."
      Height          =   400
      Left            =   225
      TabIndex        =   6
      Top             =   780
      Width           =   1200
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Unlin&k"
      Height          =   400
      Left            =   225
      TabIndex        =   5
      Top             =   2475
      Width           =   1200
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   225
      TabIndex        =   4
      Top             =   1350
      Width           =   1200
   End
   Begin VB.CommandButton cmdAddLink 
      Caption         =   "&Link..."
      Height          =   400
      Left            =   225
      TabIndex        =   3
      Top             =   225
      Width           =   1200
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "&Properties"
      Height          =   400
      Left            =   225
      TabIndex        =   2
      Top             =   1905
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   400
      Left            =   1635
      TabIndex        =   0
      Top             =   3180
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2925
      TabIndex        =   1
      Top             =   3180
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   180
      Top             =   3165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open OLE Document"
   End
   Begin VB.Label lblInvalidFormat 
      Caption         =   "Image format unrecognised."
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1650
      TabIndex        =   15
      Top             =   2685
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Label lblLinkFileMessage 
      Caption         =   "WARNING : Linked file cannot be located."
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   1650
      TabIndex        =   14
      Top             =   2220
      Visible         =   0   'False
      Width           =   1950
   End
End
Attribute VB_Name = "frmSelectEmbedded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbDocumentEdited As Boolean
Private mbChanged As Boolean
Private mbLoading As Boolean

Private moptSelection As OptSelected
Private mobjStream As ADODB.Stream

' Column information
Private mbIsReadOnly As Boolean
Private mbIsPhoto As Boolean
Private miOLEType As DataMgr.OLEType
Private mbEmbeddedEnabled As Boolean
Private mlngMaxOLESize As Long

' Properties for embedded object
Private mstrFileName As String
Private mstrPath As String
Private mstrUNC As String
Private mstrDocumentSize As String
Private mstrFileVersion As String
Private mstrFileCreateDate As String
Private mstrFileModifyDate As String
Private mstrFileDateAccessed As String
Private miLinkFileStatus As LinkFileStatus
Private mstrErrorMessage As String

' General paths
Private mstrTempPath As String

' Recreate the file locally
Private mobjFileSystem As New FileSystemObject
Private mobjFileInfo As File

Private mstrEmptyFileName As String

Const PROCESS_QUERY_INFORMATION = &H400

Private Enum LinkFileStatus
  iLINKFILE_EXISTS = 1
  iLINKFILE_NOTEXIST = 2
  iLINKFILE_CHANGED = 3
  iLINKFILE_READONLY = 4
  iLINKFILE_NOACCESS = 5
End Enum

Dim FileInfo As typSHFILEINFO

Public Property Let IsReadOnly(pbNewValue As Boolean)
  mbIsReadOnly = pbNewValue
End Property

Public Property Let IsPhoto(pbNewValue As Boolean)
  mbIsPhoto = pbNewValue
End Property

Public Property Let OLEType(piNewValue As DataMgr.OLEType)
  miOLEType = piNewValue
End Property

Public Property Get OLEType() As DataMgr.OLEType
  OLEType = miOLEType
End Property

Public Property Let EmbeddedEnabled(pbNewValue As Boolean)
  mbEmbeddedEnabled = pbNewValue
End Property

Public Property Let MaxOLESize(plngNewValue As Long)
  mlngMaxOLESize = plngNewValue * 1000
End Property

Public Property Get Selection() As OptSelected
  Selection = moptSelection
End Property

Public Property Get EmbeddedFile() As ADODB.Stream
  
  If Not IsNull(mobjStream) Then
    Set EmbeddedFile = mobjStream
  End If
  
End Property

Public Sub Initialise(ByVal objStream As ADODB.Stream)

  Screen.MousePointer = vbHourglass

  mbLoading = True
  mbDocumentEdited = False

  Set mobjStream = objStream
  If mobjStream.State = adStateClosed Then
    mobjStream.Open
    mobjStream.Type = adTypeBinary
  End If

  If mobjStream.Size > 0 Then
    LoadPropertiesFromStream
  Else
    LoadBlankStream
  End If

  mbChanged = False
  
  RefreshButtons
  
  Screen.MousePointer = vbDefault

  mbLoading = False

End Sub

Public Function CreateDocumentStream(piOLEType As DataMgr.OLEType, pstrFileName As String, pbResetPath As Boolean) As Boolean

  Dim objPropertiesStream As TextStream
  Dim objFile As New ADODB.Stream
  Dim strTempFileName As String
  Dim strFileName As String
  Dim strUNC As String
  Dim strPath As String
  Dim strFileSize As String
  Dim strFileCreateDate As String
  Dim strFileModifyDate As String
  Dim strOLEType As String
  Dim bOK As Boolean
  Dim strDate As String

  bOK = True

  ' Save the document information to file to read in.
  strTempFileName = GetTmpFName 'Left(mstrTempPath, InStr(mstrTempPath, Chr(0)) - 1) & "ole.tmp"
  
  ' Create a textfile of properties from the passed in file
  Set mobjFileInfo = mobjFileSystem.GetFile(pstrFileName)
  
  strOLEType = Trim(Str(piOLEType))
  strUNC = IIf(pbResetPath, Trim(GetUNCOnly(pstrFileName)), mstrUNC)
  strPath = IIf(pbResetPath, GetPathOnly(pstrFileName, True), mstrPath)
  strFileName = mobjFileSystem.GetFileName(pstrFileName)
  strFileSize = Trim(Str(mobjFileInfo.Size))
  
  On Error Resume Next
  strFileCreateDate = mobjFileInfo.DateCreated
  strFileModifyDate = mobjFileInfo.DateLastModified
  On Error GoTo 0
   
  Set objPropertiesStream = mobjFileSystem.OpenTextFile(strTempFileName, ForAppending, True, TristateUseDefault)
  
  objPropertiesStream.Write "<<V002>>" ' Structure version info
  objPropertiesStream.Write strOLEType & Space(2 - Len(strOLEType))
  objPropertiesStream.Write strFileName & Space(70 - Len(strFileName))
  objPropertiesStream.Write strPath & Space(210 - Len(strPath))
  objPropertiesStream.Write strUNC & Space(60 - Len(strUNC))
  objPropertiesStream.Write strFileSize & Space(10 - Len(strFileSize))
  strDate = Replace(Format(strFileCreateDate, "dd/MM/yyyy HH:MM:SS"), UI.GetSystemDateSeparator, "/")
  objPropertiesStream.Write strDate & Space(20 - Len(strDate))
  strDate = Replace(Format(strFileModifyDate, "dd/MM/yyyy HH:MM:SS"), UI.GetSystemDateSeparator, "/")
  objPropertiesStream.Write strDate & Space(20 - Len(strDate))
  objPropertiesStream.Close

  ' Create the main document stream (header then file itself)

  If Not mobjStream.State = adStateOpen Then
    mobjStream.Open
    mobjStream.Type = adTypeBinary
  End If

  ' Load the properties header
  mobjStream.LoadFromFile strTempFileName
  mobjStream.Position = mobjStream.Size
  
  ' If embedded file tack onto the end of the stream
  If piOLEType = OLE_EMBEDDED Then
    mobjStream.Position = mobjStream.Size
    objFile.Open
    objFile.Type = adTypeBinary
    objFile.LoadFromFile pstrFileName
    If objFile.Size > 0 Then
      mobjStream.Write objFile.Read
    End If
    objFile.Close
  End If
   
  mobjFileSystem.DeleteFile strTempFileName, True

  CreateDocumentStream = True
  Set mobjFileInfo = Nothing
  Set objFile = Nothing
  Exit Function
  
End Function

Private Sub cmdAddEmbed_Click()
  AddDocument OLE_EMBEDDED
End Sub

Private Sub cmdAddLink_Click()
  AddDocument OLE_UNC
End Sub

Private Sub cmdCancel_Click()
  
  If CancelChanges Then
    moptSelection = optCancel
    Me.Hide
  End If
  
End Sub

Private Sub cmdOK_Click()

  If mbChanged Or miOLEType <> OLE_UNC Then
    moptSelection = optSelect
  Else
    moptSelection = optCancel
  End If
  
  Me.Hide
End Sub

Private Sub cmdProperties_Click()

  Dim strProperties As String
  Dim strCaption As String
    
  strCaption = IIf(miOLEType = OLE_UNC, "Linked File", "Embedded File") & " properties"
    
  strProperties = "File : " & mstrFileName & vbCrLf
  strProperties = strProperties & "Size : " & NiceSize(mstrDocumentSize) & vbCrLf
  
  If miOLEType = OLE_UNC Then
    strProperties = strProperties & "Location : " & mstrUNC & mstrPath & vbCrLf
  End If
  
  strProperties = strProperties & "Last Modified : " & mstrFileModifyDate

  COAMsgBox strProperties, vbInformation, strCaption


End Sub

Private Sub cmdRemove_Click()

  Dim strCaption As String
  Dim strDelType As String
  
  strCaption = IIf(miOLEType = OLE_EMBEDDED, "Embedded", "Linked") & " Object"
  strDelType = IIf(miOLEType = OLE_EMBEDDED, "delete", "unlink")

  If COAMsgBox("Are you sure you want to " & strDelType & " " & mstrFileName & "?", vbQuestion + vbYesNo, strCaption) = vbYes Then
    LoadBlankStream
    mbChanged = True
    mbDocumentEdited = False
    lblInvalidFormat.Visible = False
    'aniSearching.AutoPlay = False
  End If

  RefreshButtons

End Sub

Private Sub Form_Initialize()

  mstrTempPath = Space(1024)
  Call GetTempPath(1024, mstrTempPath)

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

  If mbIsPhoto Then
    Me.Caption = "Select Photo"
  Else
    Me.Caption = "Select Document"
  End If

  Me.Caption = Me.Caption + IIf(mbIsReadOnly, " (Read Only)", "")

  On Error GoTo ErrorNoAnimation
  'aniSearching.Open App.Path & "\Videos\search.avi"

ErrorNoAnimation:
  Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If UnloadMode = vbFormControlMenu Then
    If cmdCancel.Enabled Then
      If Not CancelChanges Then
        Cancel = True
      End If
    Else
      If mbChanged Or miOLEType <> OLE_UNC Then
        moptSelection = optSelect
      Else
        moptSelection = optCancel
      End If
    End If
  End If

End Sub

' Dump the heder section of the main binary stream out to a temporary file and read in the properties using a text stream
Private Sub LoadPropertiesFromStream()

  Dim objTextStream As TextStream

  Dim strTempFileName As String
  Dim strProperties As String
  Dim objStreamFileName As New ADODB.Stream
  Dim strFullFileName As String
  
  ' Save the document information to file to read in.
  strTempFileName = GetTmpFName
  
  If objStreamFileName.State = adStateClosed Then
    objStreamFileName.Open
    objStreamFileName.Type = adTypeBinary
  End If

  ' Load in the header information
  If mobjStream.Size > 0 Then
    mobjStream.Position = 0
    mobjStream.CopyTo objStreamFileName, 400
    objStreamFileName.SaveToFile strTempFileName, adSaveCreateOverWrite
    
    ' Read in the document information
    Set objTextStream = mobjFileSystem.OpenTextFile(strTempFileName, ForReading)
    strProperties = Trim(objTextStream.Read(400))
    
    miOLEType = Val(Mid(strProperties, 9, 2))
    mstrFileName = Trim(GetFileNameOnly(Mid(strProperties, 11, 70)))
    mstrPath = Trim(Mid(strProperties, 81, 210))
    mstrUNC = Trim(Mid(strProperties, 291, 60))
    mstrDocumentSize = Trim(Mid(strProperties, 351, 10))
    mstrFileCreateDate = Trim(Mid(strProperties, 361, 20))
    mstrFileModifyDate = Trim(Mid(strProperties, 381, 20))
    
    If miOLEType = OLE_UNC Then
      strFullFileName = mstrUNC & mstrPath & "\" & mstrFileName
      miLinkFileStatus = IIf(mobjFileSystem.FileExists(strFullFileName), iLINKFILE_EXISTS, iLINKFILE_NOTEXIST)
    End If
  
    objTextStream.Close
    mobjFileSystem.DeleteFile strTempFileName, True
  
  End If

  objStreamFileName.Close
  Set objTextStream = Nothing

End Sub

' Extracts the path from a given filename
Public Function GetPathOnly(pstrFilePath As String, pbStripDriveLetter As Boolean) As String
   
  Dim l As Integer
  Dim tempchar As String
  Dim strPath As String
  
  l = Len(pstrFilePath)
  
  While l > 0
    tempchar = Mid(pstrFilePath, l, 1)
    If tempchar = "\" Then
      strPath = Mid(pstrFilePath, 1, l - 1)
      
      ' Strip off drive letter
      If pbStripDriveLetter And Mid(strPath, 2, 1) = ":" Then
        strPath = Mid(strPath, 3, Len(strPath))
      End If
      
      GetPathOnly = strPath
      
      Exit Function
    End If
    l = l - 1
  Wend
  
End Function

' Extracts just the filename from a path
Function GetFileNameOnly(pstrFilePath As String) As String
  Dim astrPath() As String
  astrPath = Split(pstrFilePath, "\")
  GetFileNameOnly = astrPath(UBound(astrPath))
End Function

Private Sub cmdEdit_Click()
  EditDocument
End Sub

Private Sub picViewIcon_DblClick()
  EditDocument
End Sub

Private Sub lblEmbeddedFileName_DblClick()
  EditDocument
End Sub

' Edit the document
Private Sub EditDocument()

  Dim lngHandle As Long
  Dim strTempFileName As String
  Dim iPreviousAttributes As Integer
  Dim bResetFlags As Boolean

  ' Generate the document
  If miOLEType = OLE_EMBEDDED Then
    strTempFileName = GenerateDocumentFromStream
  Else
    strTempFileName = mstrUNC & mstrPath & "\" & mstrFileName
  End If
  
  ' Force file to be read only if necessary
  On Error GoTo ErrorTrap
  If mbIsReadOnly Then
    Set mobjFileInfo = mobjFileSystem.GetFile(strTempFileName)
    iPreviousAttributes = mobjFileInfo.Attributes
    bResetFlags = True
    mobjFileInfo.Attributes = ReadOnly
  End If
  On Error Resume Next
  
  ' Get a handle to the open document
  lngHandle = OpenDocument(strTempFileName, mbIsReadOnly)
  
  ' Pass the process ID to the system locked form
  If lngHandle > 0 Then
    frmSystemLocked.LockType = IIf(mbIsPhoto, giLOCKTYPE_PHOTO, giLOCKTYPE_OLE)
    frmSystemLocked.ProcessID = lngHandle
    
    ' Show the system locked form
    frmSystemLocked.Show vbModal
         
    ' Resave the document
    If frmSystemLocked.IsFileHandleOK Then
      CreateDocumentStream miOLEType, strTempFileName, False
      mbDocumentEdited = True
      
    Else
      COAMsgBox "OpenHR has been unable to establish a connection to " & mstrFileName & " because another instance of this application is open." & vbCrLf _
            & "Close the existing application and try again.", vbExclamation, "Error"
    End If
  End If

  ' Reset the previous state flag
  On Error GoTo ErrorTrap
  If mbIsReadOnly And bResetFlags Then
    Set mobjFileInfo = mobjFileSystem.GetFile(strTempFileName)
    mobjFileInfo.Attributes = iPreviousAttributes
  End If
  On Error Resume Next

  RefreshButtons

  Set mobjFileInfo = Nothing

  Exit Sub
  
ErrorTrap:

  If LCase(Err.Description) = "permission denied" Then
    bResetFlags = False
    Resume Next
  Else
    COAMsgBox "OpenHR has been unable to load " & mstrFileName & " because you do not have access." & vbCrLf, vbExclamation, "Error"
  End If


End Sub

Private Function GenerateDocumentFromStream() As String

  Dim objTextStream As TextStream

  Dim strTempFileName As String
  Dim strProperties As String
  Dim objDocumentStream As New ADODB.Stream
  
  ' Save the document information to file to read in.
  strTempFileName = Left(mstrTempPath, InStr(mstrTempPath, Chr(0)) - 1) & mstrFileName

  If mobjStream.State = adStateClosed Then
    mobjStream.Open
    mobjStream.Type = adTypeBinary
  End If
  
  ' Setup new document stream
  objDocumentStream.Type = adTypeBinary
  objDocumentStream.Open
  
  ' Copy out the document part of the stream
  mobjStream.Position = 400
  mobjStream.CopyTo objDocumentStream, mobjStream.Size - 400
  objDocumentStream.SaveToFile strTempFileName, adSaveCreateOverWrite

  GenerateDocumentFromStream = strTempFileName

  objDocumentStream.Close
  Set objDocumentStream = Nothing

End Function

Private Sub RefreshButtons()

  Dim bInvalidImage As Boolean

  lblLinkFileMessage.Visible = False

  ' Is there an object embedded
  If mobjStream.Size > 0 Then
    picDocumentShell.Enabled = True
    cmdAddEmbed.Enabled = False
    cmdAddLink.Enabled = False
    cmdEdit.Enabled = True '(Not mbIsReadOnly And miOLEType = OLE_UNC) Or miOLEType = OLE_EMBEDDED
    picViewIcon.Enabled = cmdEdit.Enabled
    cmdProperties.Enabled = True
    cmdRemove.Caption = IIf(miOLEType = OLE_EMBEDDED, "&Delete", "Unlin&k")
    cmdRemove.Enabled = Not mbIsReadOnly
  Else
    picDocumentShell.Enabled = False
    cmdAddEmbed.Enabled = mbEmbeddedEnabled And mlngMaxOLESize > 0 And Not mbIsReadOnly
    cmdAddLink.Enabled = Not mbIsReadOnly
    cmdEdit.Enabled = False
    cmdProperties.Enabled = False
    cmdRemove.Enabled = False
    
  End If

  cmdEdit.Caption = IIf(mbIsReadOnly, "Vi&ew", "&Edit")

  ' Display the image
  If Not DisplayFileImage Then
    bInvalidImage = True
    cmdOK.Enabled = False
    cmdEdit.Enabled = False
    lblInvalidFormat.Visible = True
  End If

  ' Status of linked file
  If miOLEType = OLE_UNC And miLinkFileStatus = iLINKFILE_NOTEXIST Then
    lblLinkFileMessage.Visible = True
    cmdEdit.Enabled = False
    cmdProperties.Enabled = False
  Else
    lblLinkFileMessage.Visible = False
  End If

  cmdOK.Enabled = (mbChanged Or mbDocumentEdited) And Not bInvalidImage And Not mbIsReadOnly
'  cmdCancel.Enabled = ((Not mbDocumentEdited And miOLEType = OLE_UNC) Or miOLEType <> OLE_UNC)
  cmdCancel.Enabled = Not (miOLEType = OLE_UNC And mbDocumentEdited) Or mbIsReadOnly


End Sub

Private Sub LoadBlankStream()

  Set mobjStream = New ADODB.Stream
  mobjStream.Open
  mobjStream.Type = adTypeBinary

  mstrUNC = ""
  mstrFileName = ""
  mstrPath = ""
  mstrDocumentSize = ""
  miOLEType = OLE_UNC
  miLinkFileStatus = iLINKFILE_EXISTS

End Sub

Private Function GetDriveOnly(ByVal pstrFileName As String) As String

  If Mid(pstrFileName, 2, 1) = ":" Then
    GetDriveOnly = Mid(pstrFileName, 1, 1) & ":"
  Else
    GetDriveOnly = ""
  End If

End Function


Private Function GetUNCOnly(ByVal pstrFileName As String) As String
    
  On Local Error GoTo GetUNCPath_Err
  Dim strMsg As String
  Dim lngReturn As Long
  Dim strLocalName As String
  Dim strRemoteName As String
  Dim lngRemoteName As Long
  Dim strUNCPath As String
  strLocalName = GetDriveOnly(pstrFileName)
  strRemoteName = String(255, Chr(32))
  lngRemoteName = Len(strRemoteName)
  
  'Attempt to grab UNC
  lngReturn = WNetGetConnection(strLocalName, strRemoteName, lngRemoteName)

  If lngReturn = 0 Then
    GetUNCOnly = Trim(Replace(strRemoteName, Chr(0), ""))
    
  ' Local path
  ElseIf lngReturn = 2250 Then
    'GetUNCOnly = CurrentMachineName & "\" & Mid(strLocalName, 1, 1) & "$"
    GetUNCOnly = GetDriveOnly(pstrFileName)
  Else
    GetUNCOnly = Trim(strLocalName)
  End If

GetUNCPath_End:
  Exit Function

GetUNCPath_Err:
  GetUNCOnly = Trim(strLocalName)
  Resume GetUNCPath_End
End Function
    
Private Function CurrentMachineName() As String
   Dim Buffer As String
   Dim yBuffer() As Byte
   Dim nRet As Long
   Dim nLen As Long
   Const NameLength = 16
   
   nLen = NameLength * 2
   ReDim yBuffer(0 To nLen - 1) As Byte
   If GetComputerNameW(yBuffer(0), nLen) Then
      Buffer = yBuffer
      CurrentMachineName = Left(Buffer, nLen)
   End If
End Function

Private Function NiceSize(pstrSize As String) As String

  Select Case Len(pstrSize)
    Case Is < 5
      NiceSize = pstrSize & " bytes"
    
    Case Is < 7
      NiceSize = Mid(pstrSize, 1, Len(pstrSize) - 3) & "KB"
    
    Case 7
      NiceSize = Mid(pstrSize, 1, 1) & "." & Mid(pstrSize, 2, 2) & "MB"
    
    Case Is < 10
      NiceSize = Mid(pstrSize, 1, Len(pstrSize) - 6) & "MB"
  
  End Select

End Function

Function OpenDocument(pstrFileName As String, pbReadOnly As Boolean) As Long
  
  Dim lngTemp As Long
  Dim fIsDLL As Boolean
  Dim strFileName As String
  Dim strExePath As String
  Dim strFilePath As String

  On Error GoTo Edit_Error
  
  ' Initialise an empty string to pass to the API call
  strExePath = Space(255)
  strFilePath = GetPathOnly(pstrFileName, False)
  strFileName = GetFileNameOnly(pstrFileName)
  
  ' Get the executables path for the path & document filename
  lngTemp = FindExecutable(strFileName, strFilePath, strExePath)
  
  ' If we have got a valid executable to run the document with then continue
  If Len(Trim(strExePath)) > 1 Then
    fIsDLL = False
    
    ' For some reason W95 adds /n or /e onto the end of the path, so lose anything
    ' after the xxx.exe
    If InStr(LCase(strExePath), ".exe") > 0 Then
      strExePath = Left(strExePath, InStr(LCase(strExePath), ".exe") + 3)
    Else
      ' JPD20030227 Fault 5090
      If InStr(LCase(strExePath), ".dll") > 0 Then
        fIsDLL = True
        strExePath = Left(strExePath, InStr(LCase(strExePath), ".dll") + 3)
      
        strExePath = "rundll32.exe " & Trim(strExePath)
        
        If UCase(Right(strExePath, 11)) = "SHIMGVW.DLL" Then
          strExePath = strExePath & ",ImageView_Fullscreen"
        End If
      End If
    End If
    
    ' Tidy up the path returned from the API. Trust me, its needed !
    strExePath = Replace(Trim(strExePath), Chr(0), "")
    
    ' If the executable path returned from the API is the same as the OLE
    ' documents path and filename, then we are running an exe, so empty
    ' the spath variable, otherwise add a space.
    If strExePath = strExePath & "\" & strFileName Then strExePath = "" Else strExePath = strExePath & " "
   
    ' JDM - 21/06/2004 - Fault 5314 - Botch to get Microsoft Word working
    If (LCase(Trim(Right(pstrFileName, 3))) = "doc") _
      Or (LCase(Trim(Right(pstrFileName, 4))) = "docx") Then
      strExePath = strExePath + "/x "
    End If

    ' Shell out the process and capture the ID
    OpenDocument = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(strExePath & IIf(fIsDLL, "", Chr(34)) & pstrFileName & IIf(fIsDLL, "", Chr(34)), vbNormalFocus))
  
    ' Pass the process ID to the system locked form
'    frmSystemLocked.LockType = IIf(mbIsPhoto, giLOCKTYPE_PHOTO, giLOCKTYPE_OLE)
    
'    frmSystemLocked.ProcessID = lngTemp
  
    ' Show the system locked form
'    frmSystemLocked.Show vbModal
    
    ' Reload the file into the current stream
'    CreateDocumentStream miOLEType, pstrFileName, False
    
  Else
    COAMsgBox "No application is associated with this file.", _
      vbExclamation + vbOKOnly, app.ProductName
  End If
  
  Exit Function
Edit_Error:

  
End Function


Public Sub dwAppTerminated(obj As Object)
  ' Pick up the signal that says that the graphic editing application
  ' has terminated.
  Dim lngCount As Long
    
  ' Get rid of the locking form.
  Unload frmSystemLocked
  
  cmdCancel.Enabled = True

  Screen.MousePointer = vbDefault
    
End Sub

Private Function MapUNCToDriveLetter(pstrUNC As String) As String

  MapUNCToDriveLetter = pstrUNC

End Function

Private Sub AddDocument(piOLEType As DataMgr.OLEType)

  Dim lngIconNo As Long
  Dim strFileName As String
  Dim objCheckFileSize As File
  Dim bOK As Boolean

  On Error GoTo ErrorTrap
 
  With cdlOpen
    
    bOK = True
    
    If mbIsPhoto Then
'      .InitDir = gsPhotoPath
      .Filter = "Picture Files|*.bmp;*.gif;*.jpg"
    Else
'      .InitDir = gsOLEPath
      .Filter = "All Files (*.*)|*.*"
    End If
    
    .Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
    .ShowOpen

    ' Check file size is below the maximum
    Set objCheckFileSize = mobjFileSystem.GetFile(.FileName)
    If mbEmbeddedEnabled And objCheckFileSize.Size >= mlngMaxOLESize And piOLEType = OLE_EMBEDDED Then
        
      ' File exceeds maximum size
      COAMsgBox "File is too large to embed." & vbCrLf & "Maximum for this column is " & (mlngMaxOLESize / 1000) & " Kb", vbInformation, Me.Caption
    
    Else
      
      ' Defined maximum filename length of 70
      If Len(GetFileNameOnly(.FileName)) > 70 Then
        COAMsgBox "File name is too long." & vbCrLf & "Maximum file length is 70 characters.", vbInformation, Me.Caption
        bOK = False
      End If

      ' Defined maximum filename length of 210
      If Len(GetPathOnly(.FileName, True)) > 210 And bOK Then
        COAMsgBox "Directory structure is too long." & vbCrLf & "Maximum length is 210 characters.", vbInformation, Me.Caption
        bOK = False
      End If

      ' Defined maximum filename length of 60
      If Len(Trim(GetUNCOnly(.FileName))) > 60 And bOK Then
        COAMsgBox "Network path is too long." & vbCrLf & "Maximum length is 60 characters.", vbInformation, Me.Caption
        bOK = False
      End If

      ' Make sure file is not zero length
      If objCheckFileSize.Size = 0 Then
        COAMsgBox "File is zero length and cannot be read.", vbInformation, Me.Caption
        bOK = False
      End If

      If Len(.FileName) > 0 And bOK Then
        CreateDocumentStream piOLEType, .FileName, True
        mbChanged = True
      End If
    
    End If

  End With

  ' Refresh document settings
  LoadPropertiesFromStream

  ' Refresh the buttons
  RefreshButtons
  
  Set objCheckFileSize = Nothing
  
Exit Sub

ErrorTrap:
  If Err.Number = 32755 Then
    Err.Clear
    Exit Sub
  End If

End Sub

Private Function DisplayFileImage() As Boolean

  On Error GoTo ErrorTrap

  Dim strIconFileName As String
  Dim objIcon As Long
  Dim objListItem As ListItem
  Dim bOK As Boolean
  
  bOK = True
  
  ' Hide all the photo stuff
  If (miOLEType = OLE_UNC And miLinkFileStatus = iLINKFILE_NOTEXIST) Then
  
    'lblLinkFileMessage.Caption = "WARNING : " & mstrUNC & "\" & mstrFileName & " cannot be found"
    fraLinkEmbed.BorderStyle = 0
    imgPhoto.Visible = False
    picDocumentShell.Visible = False
    aniSearching.Visible = True
    picLinkedFile.Visible = False
    picEmbedFile.Visible = False
    lblLinkFileMessage.Visible = True
  Else
  
    aniSearching.Visible = False
    fraLinkEmbed.BorderStyle = 1
  
    If Len(mstrFileName) = 0 Then
    
      ' No document here
      imgPhoto.Picture = Nothing
      imgPhoto.Visible = mbIsPhoto
      picDocumentShell.Visible = Not mbIsPhoto
      picLinkedFile.Visible = False
      picEmbedFile.Visible = False
      fraLinkEmbed.Caption = ""
    
      ' Set text file name
      imgPhoto.Enabled = False
      picDocumentShell.Enabled = False
      picDocumentShell.BackColor = vbButtonFace
      picViewIcon.BackColor = vbButtonFace
      picViewIcon.Picture = Nothing
      lblEmbeddedFileName.Caption = ""
  
    Else
    
      ' Generate the filename depending on how the object is linked/embedded etc...
      Select Case miOLEType
            
        ' Embedded files - file doesn't yet exist, so create a dummy file an extract the icon from it
        Case OLE_EMBEDDED
          
          ' If it's a photo generate the file, otherwise a dummy text file will do
          If mbIsPhoto Then
            strIconFileName = GenerateDocumentFromStream
          Else
            strIconFileName = Left(mstrTempPath, InStr(mstrTempPath, Chr(0)) - 1) & mstrFileName
            mobjFileSystem.CreateTextFile strIconFileName, True
          End If
          
          picLinkedFile.Visible = False
          picEmbedFile.Visible = True
          fraLinkEmbed.Caption = Space(6) & "Embedded " & IIf(mbIsPhoto, "Photo", "File")
        
        ' File is somewhere on the network
        Case OLE_UNC
          strIconFileName = mstrUNC & mstrPath & "\" & mstrFileName
      
          picLinkedFile.Visible = True
          picEmbedFile.Visible = False
          fraLinkEmbed.Caption = Space(7) & "Linked " & IIf(mbIsPhoto, "Photo", "File")
      
      End Select
  
      If mbIsPhoto Then
      
        ' Load photo
        aniSearching.Visible = False
        picDocumentShell.Visible = False
        imgPhoto.Visible = True
        imgPhoto.Picture = LoadPicture(strIconFileName)
      
      Else
      
        imgPhoto.Visible = False
        picDocumentShell.Visible = True
      
        ' Load document icon
        objIcon = SHGetFileInfo(strIconFileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
      
        If objIcon <> 0 Then
        
          picDocumentShell.BackColor = vbWhite
        
          With picViewIcon
            .BackColor = vbWhite
            .Height = 15 * 32
            .Width = 15 * 32
            .ScaleHeight = 15 * 32
            .ScaleWidth = 15 * 32
            .Picture = LoadPicture("")
            .AutoRedraw = True
            
            objIcon = ImageList_Draw(objIcon, FileInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
            .Refresh
          End With
          
          ' Set text file name
          'JPD 20060106 Fault 10397
          lblEmbeddedFileName.Caption = Replace(GetFileNameOnly(strIconFileName), "&", "&&")
          
          ' Tooltip text
          picDocumentShell.ToolTipText = mstrFileName
          picViewIcon.ToolTipText = picDocumentShell.ToolTipText
          lblEmbeddedFileName.ToolTipText = picDocumentShell.ToolTipText
         
        End If
      
        ' Get rid of the temporary file
        If miOLEType = OLE_EMBEDDED Then
          mobjFileSystem.DeleteFile strIconFileName, True
        End If
        
        Set mobjFileSystem = Nothing
      End If
        
    End If
    
    ' Format the frame nicely
    fraLinkEmbed.Height = IIf(mbIsPhoto, imgPhoto.Height, picDocumentShell.Height) + 400
    
  End If

TidyUpAndExit:
  
  DisplayFileImage = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

Private Function CancelChanges() As Boolean
  
  Dim pintAnswer As Integer
    
  If (mbChanged = True Or mbDocumentEdited) And Not mbIsReadOnly Then
    pintAnswer = COAMsgBox("All changes will be lost. Are you sure you want to cancel?", vbQuestion + vbYesNo, Me.Caption)
    CancelChanges = (pintAnswer = vbYes)
  Else
    CancelChanges = True
  End If
  
End Function

