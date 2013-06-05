VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmSelectPhoto 
   Caption         =   "Select Photo"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1059
   Icon            =   "frmSelectPhoto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6765
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   5145
      Top             =   2310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open Photo"
      Filter          =   "Picture Files|*.bmp;*.gif;*.jpg"
   End
   Begin VB.FileListBox filPhoto 
      Height          =   2625
      Left            =   150
      ReadOnly        =   0   'False
      TabIndex        =   6
      Top             =   150
      Width           =   2505
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5385
      TabIndex        =   5
      Top             =   2350
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   400
      Left            =   5385
      TabIndex        =   4
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "&None"
      Height          =   400
      Left            =   5385
      TabIndex        =   3
      Top             =   1250
      Width           =   1200
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit..."
      Height          =   400
      Left            =   5385
      TabIndex        =   2
      Top             =   700
      Width           =   1200
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add..."
      Height          =   400
      Left            =   5385
      TabIndex        =   1
      Top             =   150
      Width           =   1200
   End
   Begin VB.Frame fraPreview 
      Caption         =   "Preview :"
      Height          =   2700
      Left            =   2800
      TabIndex        =   0
      Top             =   75
      Width           =   2450
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         Caption         =   "Unsupported Image Format"
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image imgPhoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   200
         Stretch         =   -1  'True
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmSelectPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum OptSelected
    optSelect = 1
    optCancel = 2
    optNone
End Enum

Private moptPhoto As OptSelected
Private msPhoto As String
Private msPath As String
Private objProc As Object


Public Sub Initialise(sPhoto As String)
  ' Initialise the Photo Selection form.
  Dim lngCount As Long
  
  With filPhoto
    .Path = gsPhotoPath
    .Pattern = "*.bmp;*.gif;*.jpg"
    .Refresh
    
    If Len(sPhoto) > 1 Then
      For lngCount = 0 To .ListCount - 1
        If .List(lngCount) = sPhoto Then
          .ListIndex = lngCount
          Exit For
        End If
      Next
    Else
      If .ListCount > 0 Then
        .ListIndex = 0
        'TM20010910 Fault 2806
        'Changing the .ListIndex property calls the Click event of the file list box.
'        filPhoto_Click
      End If
    End If

    If .ListCount > 0 Then
      cmdEdit.Enabled = True
      cmdSelect.Enabled = True
      filPhoto.Enabled = True
    Else
      cmdEdit.Enabled = False
      cmdSelect.Enabled = False
      filPhoto.Enabled = False
    End If
  End With
  
  'JPD 20030828 Fault 5659
  Me.Top = GetPCSetting("SelectPhotoWindowCoOrdinates\" & gsDatabaseName, "Top", CLng((Screen.Height - Me.Height) / 2))
  Me.Left = GetPCSetting("SelectPhotoWindowCoOrdinates\" & gsDatabaseName, "Left", CLng((Screen.Width - Me.Width) / 2))
  Me.Width = GetPCSetting("SelectPhotoWindowCoOrdinates\" & gsDatabaseName, "Width", Me.Width)
  Me.Height = GetPCSetting("SelectPhotoWindowCoOrdinates\" & gsDatabaseName, "Height", Me.Height)
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub EditFile_Part2()
  ' Launch the default graphic editing package.
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim lngAPIResult As Long
  Dim lngKeyHandle As Long
  Dim lngTYPE As Long
  Dim sFileExtension As String
  Dim sRegValue As String
  Dim sApplication As String
  Dim pid&
  
  On Error GoTo Edit_Error

  Const iVALUESIZE = 1024

  fOK = False

  ' Get the extension of the file to open.
  sFileExtension = ""
  For iLoop = Len(filPhoto.FileName) To 1 Step -1
    sFileExtension = Mid(filPhoto.FileName, iLoop, 1) & sFileExtension

    If Mid(filPhoto.FileName, iLoop, 1) = "." Then
      fOK = True
      Exit For
    End If
  Next iLoop

  If Not fOK Then
    ' No extension on the file.
    COAMsgBox "Error opening the selected file." & vbCrLf & _
      "The file has no extension.", _
      vbExclamation + vbOKOnly, Application.Name
  Else
    ' Open the registry key for the given file extension's class.
    lngAPIResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, sFileExtension, 0, KEY_ALL_ACCESS, lngKeyHandle)
    fOK = (lngAPIResult = ERROR_SUCCESS)

    If Not fOK Then
      COAMsgBox "Error opening the selected file." & vbCrLf & _
        "Unable to open the file extension's key in the registry.", _
        vbExclamation + vbOKOnly, Application.Name
    End If
  End If

  If fOK Then
    ' Get the class name of the given file extension.
    sRegValue = Space$(iVALUESIZE)

    lngAPIResult = RegQueryValueEx(lngKeyHandle, "", 0, lngTYPE, ByVal sRegValue, iVALUESIZE)
    RegCloseKey (lngKeyHandle)

    fOK = ((lngAPIResult = ERROR_SUCCESS) And (lngTYPE = REG_SZ))

    If Not fOK Then
      COAMsgBox "Error opening the selected file." & vbCrLf & _
        "Unable to read the file extension's class from the registry.", _
        vbExclamation + vbOKOnly, Application.Name
    End If
  End If

  If fOK Then
    ' Extract the file extension's class from the value read from the registry.
    For iLoop = 1 To Len(sRegValue)
      If Mid(sRegValue, iLoop, 1) = vbNullChar Then
        sRegValue = Left(sRegValue, iLoop - 1)
        Exit For
      End If
    Next iLoop

    fOK = (Len(sRegValue) > 0)

    If Not fOK Then
      COAMsgBox "Error opening the selected file." & vbCrLf & _
        "Unable to read the file extension's class from the registry.", _
        vbExclamation + vbOKOnly, Application.Name
    End If
  End If

  If fOK Then
    ' Open the given file's class registry entry.
    lngAPIResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, sRegValue & "\shell\open\command", 0, KEY_ALL_ACCESS, lngKeyHandle)
    sRegValue = Space$(iVALUESIZE)

    fOK = (lngAPIResult = ERROR_SUCCESS)

    If Not fOK Then
      COAMsgBox "Error opening the selected file." & vbCrLf & _
        "Unable to open the given file's class key in the registry.", _
        vbExclamation + vbOKOnly, Application.Name
    End If
  End If

  If fOK Then
    ' Read the given file's default application from the registry.
    lngAPIResult = RegQueryValueEx(lngKeyHandle, "", 0, lngTYPE, ByVal sRegValue, iVALUESIZE)
    RegCloseKey (lngKeyHandle)

    fOK = ((lngAPIResult = ERROR_SUCCESS) And ((lngTYPE = REG_EXPAND_SZ) Or (lngTYPE = REG_SZ)))

    If Not fOK Then
      COAMsgBox "Error opening the selected file." & vbCrLf & _
        "Unable to read the default application for the given file's class from the registry.", _
        vbExclamation + vbOKOnly, Application.Name
    End If
  End If

  If fOK Then
    ' Extract the file extension's class from the value read from the registry.
    For iLoop = 1 To Len(sRegValue)
      If Mid(sRegValue, iLoop, 1) = vbNullChar Then
        sRegValue = Left(sRegValue, iLoop - 1)
        Exit For
      End If
    Next iLoop

    fOK = (Len(sRegValue) > 0)
    If Not fOK Then
      COAMsgBox "Error opening the selected file." & vbCrLf & _
        "Unable to read the default application for the given file's class from the registry.", _
        vbExclamation + vbOKOnly, Application.Name
    End If
  End If

  If fOK Then
    ' Remove any parameter definitions from the default application string.
    sApplication = sRegValue
    For iLoop = (Len(sRegValue) - 1) To 1 Step -1
      If (Mid(sRegValue, iLoop, 2) = " /") Or (Mid(sRegValue, iLoop, 2) = " """) Then
        sApplication = Left(sRegValue, iLoop - 1)
      End If
    Next iLoop

    sApplication = sApplication & " """ & filPhoto.Path & "\" & filPhoto.FileName & """"

    pid = Shell(sApplication, vbNormalFocus)

    fOK = (pid <> 0)

    If Not fOK Then
      COAMsgBox "Error opening the selected file." & vbCrLf & _
        "Unable to run the selected file's default application.", _
        vbExclamation + vbOKOnly, Application.Name
    End If
  End If

  If fOK Then
    ' Launch the default graphic editing package.
    cmdCancel.Enabled = False

    msPhoto = filPhoto.FileName
    Set objProc = CreateObject("procWatcher.dwAppWatch")
    objProc.SetAppWatch pid
    objProc.SetAppCallback Me

    ' Lock HR Pro until the photo edition has been done.
    frmSystemLocked.LockType = giLOCKTYPE_PHOTO
    frmSystemLocked.Show vbModal
  End If

  Exit Sub
  
Edit_Error:
  COAMsgBox "Error attempting to invoke the default editor for this file type." & vbCrLf & vbCrLf & _
         "(" & Err.Number & " - " & Err.Description & ")", _
         vbExclamation + vbOKOnly, App.Title

End Sub

Private Sub EditFile_Part1()
  Dim lngTemp As Long
  Dim sPath As String
  Dim fIsDLL As Boolean
  Dim sFilePath As String
  
  On Error GoTo Edit_Error
  
  ' Initialise an empty string to pass to the API call
  sPath = Space(255)
  
  sFilePath = filPhoto.Path
  If Right(sFilePath, 1) = "\" Then
    sFilePath = Left(sFilePath, Len(sFilePath) - 1)
  End If
  
  ' Get the executables path for the path & document filename
  lngTemp = FindExecutable(filPhoto.FileName, filPhoto.Path, sPath)
  
  ' If we have got a valid executable to run the document with then continue
  If Len(Trim(sPath)) > 1 Then
    fIsDLL = False
    
    ' For some reason W95 adds /n or /e onto the end of the path, so lose anything
    ' after the xxx.exe
    If InStr(LCase(sPath), ".exe") > 0 Then
      sPath = Left(sPath, InStr(LCase(sPath), ".exe") + 3)
    Else
      ' JPD20030227 Fault 5090
      If InStr(LCase(sPath), ".dll") > 0 Then
        fIsDLL = True
        sPath = Left(sPath, InStr(LCase(sPath), ".dll") + 3)
      
        sPath = "rundll32.exe " & Trim(sPath)
        
        If UCase(Right(sPath, 11)) = "SHIMGVW.DLL" Then
          sPath = sPath & ",ImageView_Fullscreen"
        End If
      End If
    End If
    
    ' Tidy up the path returned from the API. Trust me, its needed !
    sPath = Replace(Trim(sPath), Chr(0), "")
    
    ' If the executable path returned from the API is the same as the Photo
    ' documents path and filename, then we are running an exe, so empty
    ' the spath variable, otherwise add a space.
    If sPath = sFilePath & "\" & filPhoto.FileName Then sPath = "" Else sPath = sPath & " "
    
    ' Shell out the process and capture the ID
    lngTemp = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sPath & IIf(fIsDLL, "", Chr(34)) & sFilePath & "\" & filPhoto.FileName & IIf(fIsDLL, "", Chr(34)), vbNormalFocus))
  
    ' Pass the process ID to the system locked form
    frmSystemLocked.ProcessID = lngTemp
  
    ' Show the system locked form
    frmSystemLocked.LockType = giLOCKTYPE_PHOTO
    frmSystemLocked.Show vbModal
  Else
    COAMsgBox "No application is associated with this file.", _
      vbExclamation + vbOKOnly, App.ProductName
  End If
  
  Exit Sub
  
Edit_Error:
  
  EditFile_Part2

End Sub

Private Sub cmdAdd_Click()
  On Error GoTo Err_Trap
    
  With cdlOpen
    ' JDM - 20/01/2005 - Fault 8242 - Leave path as existing.
    '.InitDir = gsPhotoPath
    
    .ShowOpen
    If Len(.FileTitle) > 0 Then
      If Len(Dir(gsPhotoPath & "\" & .FileTitle)) = 0 Then
'        DoEvents
        FileCopy .FileName, gsPhotoPath & "\" & .FileTitle
        Initialise .FileTitle
        Exit Sub
      Else
        frmFileReplace.Initialise .FileName, gsPhotoPath & "\" & .FileTitle, .FileTitle
        frmFileReplace.Show vbModal
        If frmFileReplace.Replaced Then
'          DoEvents
          Screen.MousePointer = vbHourglass
          If UCase(Trim(.FileName)) <> UCase(Trim(gsPhotoPath & "\" & .FileTitle)) Then
            FileCopy .FileName, gsPhotoPath & "\" & .FileTitle
          End If
          Initialise .FileTitle
          Unload frmFileReplace
        End If
      End If
    End If
  End With
  
  Exit Sub
  
Err_Trap:
  If Err.Number = 32755 Then
    Err.Clear
    Exit Sub
  End If

End Sub

Private Sub cmdCancel_Click()

    optPhoto = optCancel
    Me.Hide

End Sub

Private Sub cmdEdit_Click()
  EditFile_Part1
  
End Sub


Private Sub cmdNone_Click()

    optPhoto = optNone
    Me.Hide

End Sub

Private Sub cmdSelect_Click()

    optPhoto = optSelect
    Photo = filPhoto.FileName
    Me.Hide

End Sub

Private Sub filPhoto_Click()

  On Error GoTo ErrorTrap

'  DoEvents
  Screen.MousePointer = vbHourglass

  'TM20010910 Fault 2806
  'Need to make sure the file still exists when it is selected.
  Dim sFilenameAndPath As String
  
  lblWarning.Visible = False
  sFilenameAndPath = filPhoto.Path & "\" & filPhoto.FileName
  
  If FileExists(sFilenameAndPath) Then
    imgPhoto.Picture = LoadPicture(sFilenameAndPath)
  Else
    COAMsgBox "The selected file has been moved or deleted." & vbCrLf & _
          "It will be removed from the list.", vbOKOnly + vbExclamation, App.Title
    filPhoto.Refresh
  End If
  
  msPath = filPhoto.Path

TidyUpAndExit:
  Screen.MousePointer = vbDefault
  Exit Sub

ErrorTrap:
  imgPhoto.Picture = Nothing
  lblWarning.Visible = True
  Resume TidyUpAndExit
  Exit Sub

End Sub

Private Function FileExists(sPath As String) As Boolean

  On Error GoTo ErrorTrap
  
  If Dir(sPath) <> vbNullString Then
    FileExists = True
  Else
    FileExists = False
  End If

TidyUpAndExit:
  Exit Function

ErrorTrap:
  FileExists = False
  Resume TidyUpAndExit
  
End Function

Public Property Get optPhoto() As OptSelected

    optPhoto = moptPhoto

End Property

Public Property Let optPhoto(ByVal opt As OptSelected)

    moptPhoto = opt

End Property

Public Property Get Photo() As String

    Photo = msPhoto

End Property

Public Property Let Photo(ByVal sPhoto As String)

    msPhoto = sPhoto

End Property

Private Sub Form_Load()
  RemoveIcon Me
  Hook Me.hWnd, 6900, 3500
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        optPhoto = optCancel
        Cancel = True
        Me.Hide
    End If

End Sub

Public Property Get Path() As String

    Path = msPath

End Property

Public Sub dwAppTerminated(obj As Object)
  ' Pick up the signal that says that the grpaphic editing application
  ' has terminated.
  Dim lngCount As Long
    
  ' Get rid of the locking form.
  Unload frmSystemLocked
  
  ' Highlight the selected photo file in the listbox.
  With filPhoto
    For lngCount = 0 To .ListCount - 1
      If .List(lngCount) = msPhoto Then
        .ListIndex = lngCount
        Exit For
      End If
    Next
  End With
    
  filPhoto_Click
  cmdCancel.Enabled = True

  Set objProc = Nothing
  Screen.MousePointer = vbDefault
    
End Sub


Private Sub Form_Resize()
  'JPD 20030828 Fault 5659
'  Const MINHEIGHT = 3500
'  Const MINWIDTH = 6900
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
'  If Me.Height < MINHEIGHT Then
'    Me.Height = MINHEIGHT
'  End If
'
'  If Me.Width < MINWIDTH Then
'    Me.Width = MINWIDTH
'  End If
  
  With filPhoto
    .Height = Me.ScaleHeight - 300
    .Width = Me.ScaleWidth - 4250
  
    fraPreview.Height = .Height + 75
    fraPreview.Left = .Left + .Width + 150
    
    cmdAdd.Left = fraPreview.Left + fraPreview.Width + 150
    cmdEdit.Left = fraPreview.Left + fraPreview.Width + 150
    
    cmdCancel.Left = fraPreview.Left + fraPreview.Width + 150
    cmdCancel.Top = .Height + .Top - cmdCancel.Height
    
    cmdSelect.Left = fraPreview.Left + fraPreview.Width + 150
    cmdSelect.Top = cmdCancel.Top - 550
  
    cmdNone.Left = fraPreview.Left + fraPreview.Width + 150
    cmdNone.Top = cmdSelect.Top - 550
  End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
  'JPD 20030828 Fault 5659
  If Me.WindowState = vbNormal Then
    SavePCSetting "SelectPhotoWindowCoOrdinates\" & gsDatabaseName, "Top", Me.Top
    SavePCSetting "SelectPhotoWindowCoOrdinates\" & gsDatabaseName, "Left", Me.Left
    SavePCSetting "SelectPhotoWindowCoOrdinates\" & gsDatabaseName, "Width", Me.Width
    SavePCSetting "SelectPhotoWindowCoOrdinates\" & gsDatabaseName, "Height", Me.Height
  End If

  Unhook Me.hWnd
End Sub



