VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmSelectOLE 
   Caption         =   "Select OLE"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1058
   Icon            =   "frmSelectOLE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4125
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add..."
      Height          =   400
      Left            =   2800
      TabIndex        =   1
      Top             =   150
      Width           =   1200
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit..."
      Height          =   400
      Left            =   2800
      TabIndex        =   2
      Top             =   700
      Width           =   1200
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "&None"
      Height          =   400
      Left            =   2800
      TabIndex        =   3
      Top             =   1250
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   400
      Left            =   2800
      TabIndex        =   4
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2800
      TabIndex        =   5
      Top             =   2350
      Width           =   1200
   End
   Begin VB.FileListBox filOLEs 
      Height          =   2625
      Left            =   150
      ReadOnly        =   0   'False
      TabIndex        =   0
      Top             =   150
      Width           =   2500
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   3675
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open OLE Document"
   End
   Begin VB.OLE OLE1 
      AutoActivate    =   0  'Manual
      AutoVerbMenu    =   0   'False
      Height          =   405
      Left            =   3720
      OLETypeAllowed  =   0  'Linked
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2010
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmSelectOLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbDocumentEdited As Boolean
Private moptSelection As OptSelected
Private msOLEFileName As String
Private msPath As String
Private mbIsReadOnly As Boolean
Private mfOleOnServer As Boolean
Private objProc As Object

Public Property Let IsReadOnly(pbNewValue As Boolean)
  mbIsReadOnly = pbNewValue
End Property

Public Sub dwAppTerminated(obj As Object)
  ' Pick up the signal that says that the grpaphic editing application
  ' has terminated.
  Dim lngCount As Long
    
  ' Get rid of the locking form.
  Unload frmSystemLocked
  
  ' Highlight the selected file in the listbox.
  With filOLEs
    For lngCount = 0 To .ListCount - 1
      If .List(lngCount) = msOLEFileName Then
        .ListIndex = lngCount
        Exit For
      End If
    Next
  End With
    
  filOLEs_Click
  cmdCancel.Enabled = True

  Set objProc = Nothing
  Screen.MousePointer = vbDefault
    
End Sub



Private Sub EditFile_Part1()
  Dim lngTemp As Long
  Dim sPath As String
  Dim fIsDLL As Boolean
  Dim sFilePath As String

  On Error GoTo Edit_Error
  
  ' Initialise an empty string to pass to the API call
  sPath = Space(255)
  
  sFilePath = filOLEs.Path
  If Right(sFilePath, 1) = "\" Then
    sFilePath = Left(sFilePath, Len(sFilePath) - 1)
  End If
  
  ' Get the executables path for the path & document filename
  lngTemp = FindExecutable(filOLEs.FileName, filOLEs.Path, sPath)
  
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
    
    ' If the executable path returned from the API is the same as the OLE
    ' documents path and filename, then we are running an exe, so empty
    ' the spath variable, otherwise add a space.
    If sPath = sFilePath & "\" & filOLEs.FileName Then sPath = "" Else sPath = sPath & " "
    
    ' JDM - 21/06/2004 - Fault 5314 - Botch to get Microsoft Word working
    If (LCase(Trim(Right(filOLEs.FileName, 3))) = "doc") _
      Or (LCase(Trim(Right(filOLEs.FileName, 4))) = "docx") Then
      sPath = sPath + "/x "
    End If
   
    ' Shell out the process and capture the ID
    lngTemp = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sPath & IIf(fIsDLL, "", Chr(34)) & sFilePath & "\" & filOLEs.FileName & IIf(fIsDLL, "", Chr(34)), vbNormalFocus))
  
    ' Pass the process ID to the system locked form
    frmSystemLocked.ProcessID = lngTemp
  
    ' Show the system locked form
    frmSystemLocked.LockType = giLOCKTYPE_OLE
    frmSystemLocked.Show vbModal
    
  Else
    COAMsgBox "No application is associated with this file.", _
      vbExclamation + vbOKOnly, App.ProductName
  End If
  
  Exit Sub
  
Edit_Error:
  
  EditFile_Part2

End Sub

Private Sub EditFile_Part2()
  ' Launch the default editing package.
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
  For iLoop = Len(filOLEs.FileName) To 1 Step -1
    sFileExtension = Mid(filOLEs.FileName, iLoop, 1) & sFileExtension

    If Mid(filOLEs.FileName, iLoop, 1) = "." Then
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

    sApplication = sApplication & " """ & filOLEs.Path & "\" & filOLEs.FileName & """"

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

    msOLEFileName = filOLEs.FileName
    Set objProc = CreateObject("procWatcher.dwAppWatch")
    objProc.SetAppWatch pid
    objProc.SetAppCallback Me

    ' Lock HR Pro until the edition has been done.
    frmSystemLocked.LockType = giLOCKTYPE_OLE
    frmSystemLocked.Show vbModal
  End If

  RefreshControls

  Exit Sub
  
Edit_Error:
  COAMsgBox "Error attempting to invoke the default editor for this file type." & vbCrLf & vbCrLf & _
         "(" & Err.Number & " - " & Err.Description & ")", _
         vbExclamation + vbOKOnly, App.Title
End Sub

Public Property Let OleOnServer(ByVal pfNewValue As Boolean)
  mfOleOnServer = pfNewValue

End Property

Public Property Get OleOnServer() As Boolean
  OleOnServer = mfOleOnServer

End Property

Public Property Let OLEFileName(ByVal psOLEFileName As String)
  msOLEFileName = psOLEFileName

End Property

Public Property Get optSelection() As OptSelected
  optSelection = moptSelection

End Property

Public Property Get Path() As String
  Path = msPath

End Property

Public Property Get OLEFileName() As String
  OLEFileName = msOLEFileName

End Property


Public Property Let optSelection(ByVal newOpt As OptSelected)
  moptSelection = newOpt

End Property
Public Sub Initialise(psOLEFile As String)
  Dim lngCount As Long
    
  mbDocumentEdited = False
    
  ' Initialise the file list box.
  With filOLEs
'
'    If mfOleOnServer = True Then
'      .Path = gsOLEPath
'    Else
'      .Path = gsLocalOLEPath
'    End If
'
    If mfOleOnServer = True Then
      If Dir(gsOLEPath & IIf(Right(gsOLEPath, 1) = "\", "*.*", "\*.*"), vbDirectory) <> vbNullString Then
        .Path = gsOLEPath
      End If
    Else
      If Dir(gsLocalOLEPath & IIf(Right(gsLocalOLEPath, 1) = "\", "*.*", "\*.*"), vbDirectory) <> vbNullString Then
        .Path = gsLocalOLEPath
      End If
    End If
    
    .Pattern = "*.*"
    .Refresh
       
    If Len(psOLEFile) > 0 Then
      ' Select the current OLE file in the listbox.
      For lngCount = 0 To .ListCount - 1
        If UCase(.List(lngCount)) = UCase(psOLEFile) Then
          .ListIndex = lngCount
          Exit For
        End If
      Next
    Else
      If .ListCount > 0 Then
        .ListIndex = 0
        filOLEs_Click
      End If
    End If
    
    RefreshControls
  End With
    
  'JPD 20030828 Fault 5659
  Me.Top = GetPCSetting("SelectOLEWindowCoOrdinates\" & gsDatabaseName, "Top", CLng((Screen.Height - Me.Height) / 2))
  Me.Left = GetPCSetting("SelectOLEWindowCoOrdinates\" & gsDatabaseName, "Left", CLng((Screen.Width - Me.Width) / 2))
  Me.Width = GetPCSetting("SelectOLEWindowCoOrdinates\" & gsDatabaseName, "Width", Me.Width)
  Me.Height = GetPCSetting("SelectOLEWindowCoOrdinates\" & gsDatabaseName, "Height", Me.Height)
    
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdAdd_Click()
  On Error GoTo ErrorTrap
  
  With cdlOpen
    
' JDM - 20/01/2005 - Fault 8243 - Leave path as existing.
'    If mfOleOnServer = True Then
'      .InitDir = gsOLEPath
'    Else
'      .InitDir = gsLocalOLEPath
'    End If
    
    ' RH 28/09/00 - BUG 1019
    .Filter = "All Files (*.*)|*.*"
    .Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
    
    .ShowOpen

    If Len(.FileTitle) > 0 Then
      
      If mfOleOnServer = True Then
      
        If Len(Dir(gsOLEPath & "\" & .FileTitle)) = 0 Then
          FileCopy .FileName, gsOLEPath & "\" & .FileTitle
          Initialise .FileTitle
          ' RH 21/08/00 - BUG 795
          'cmdEdit_Click
          Exit Sub
        Else
          
          ' RH 21/08/00 - BUG 797
          If LCase(.FileName) <> LCase(gsOLEPath & "\" & .FileTitle) Then
          
            frmFileReplace.Initialise .FileName, gsOLEPath & "\" & .FileTitle, .FileTitle
            frmFileReplace.Show vbModal
            
            If frmFileReplace.Replaced Then
              Screen.MousePointer = vbHourglass
              
              ' Copy the file to the HR Pro OLE document directory on the server
              ' if it is not already there.
              If UCase(Trim(.FileName)) <> UCase(Trim(gsOLEPath & "\" & .FileTitle)) Then
                FileCopy .FileName, gsOLEPath & "\" & .FileTitle
              End If
              Initialise .FileTitle
              Unload frmFileReplace
              Screen.MousePointer = vbDefault
            End If
          End If
          
        End If
      
      Else
      
        If Len(Dir(gsLocalOLEPath & "\" & .FileTitle)) = 0 Then
          FileCopy .FileName, gsLocalOLEPath & "\" & .FileTitle
          Initialise .FileTitle
          ' RH 21/08/00 - BUG 795
          'cmdEdit_Click
          Exit Sub
        Else
          
          ' RH 19/09/00 - BUG 796
          If LCase(.FileName) <> LCase(gsLocalOLEPath & IIf(Right(gsLocalOLEPath, 1) = "\", "", "\") & .FileTitle) Then
          
            frmFileReplace.Initialise .FileName, gsLocalOLEPath & "\" & .FileTitle, .FileTitle
            frmFileReplace.Show vbModal
            
            If frmFileReplace.Replaced Then
              Screen.MousePointer = vbHourglass
              
              ' Copy the file to the HR Pro OLE document directory on the local machine
              ' if it is not already there.
              If UCase(Trim(.FileName)) <> UCase(Trim(gsLocalOLEPath & "\" & .FileTitle)) Then
                FileCopy .FileName, gsLocalOLEPath & "\" & .FileTitle
              End If
              Initialise .FileTitle
              Unload frmFileReplace
              Screen.MousePointer = vbDefault
            End If
          End If
        End If
      
      End If
      
      ' RH 21/08/00 - BUG 795
      'cmdEdit_Click
    End If
  End With

  Exit Sub

ErrorTrap:
  If Err.Number = 32755 Then
    Err.Clear
    Exit Sub
  End If

End Sub


Private Sub cmdCancel_Click()
  optSelection = optCancel
  Me.Hide

End Sub


Private Sub cmdEdit_Click()
  EditFile_Part1
End Sub


Private Sub cmdNone_Click()
  optSelection = optNone
  Me.Hide

End Sub


Private Sub cmdSelect_Click()
  optSelection = optSelect
  OLEFileName = filOLEs.Path & "\" & filOLEs.FileName
  Me.Hide

End Sub


Private Sub filOLEs_Click()
  Dim sFilenameAndPath As String
  
  Screen.MousePointer = vbHourglass
  
  sFilenameAndPath = filOLEs.Path & "\" & filOLEs.FileName
  
  If Not FileExists(sFilenameAndPath) Then
    COAMsgBox "The selected file has been moved or deleted." & vbCrLf & _
          "It will be removed from the list.", vbOKOnly + vbExclamation, App.Title
    filOLEs.Refresh
  End If
  
  msPath = filOLEs.Path
  Screen.MousePointer = vbDefault

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


Private Sub filOLEs_DblClick()
  cmdSelect_Click
  
End Sub


Private Sub Form_Load()
  RemoveIcon Me
  Hook Me.hWnd, 4300, 3500
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    optSelection = optCancel
    Cancel = True
    Me.Hide
  End If

End Sub





Private Sub Form_Resize()
  'JPD 20030828 Fault 5659
'  Const MINHEIGHT = 3500
'  Const MINWIDTH = 4300
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
'  If Me.Height < MINHEIGHT Then
'    Me.Height = MINHEIGHT
'  End If
'
'  If Me.Width < MINWIDTH Then
'    Me.Width = MINWIDTH
'  End If
  
  With filOLEs
    .Height = Me.ScaleHeight - 300
    .Width = Me.ScaleWidth - 1600
  
    cmdAdd.Left = .Left + .Width + 150
    cmdEdit.Left = .Left + .Width + 150
    
    cmdCancel.Left = .Left + .Width + 150
    cmdCancel.Top = .Height + .Top - cmdCancel.Height
    
    cmdSelect.Left = .Left + .Width + 150
    cmdSelect.Top = cmdCancel.Top - 550
  
    cmdNone.Left = .Left + .Width + 150
    cmdNone.Top = cmdSelect.Top - 550
  End With
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
  'JPD 20030828 Fault 5659
  If Me.WindowState = vbNormal Then
    SavePCSetting "SelectOLEWindowCoOrdinates\" & gsDatabaseName, "Top", Me.Top
    SavePCSetting "SelectOLEWindowCoOrdinates\" & gsDatabaseName, "Left", Me.Left
    SavePCSetting "SelectOLEWindowCoOrdinates\" & gsDatabaseName, "Width", Me.Width
    SavePCSetting "SelectOLEWindowCoOrdinates\" & gsDatabaseName, "Height", Me.Height
  End If

  Unhook Me.hWnd
End Sub

Private Sub RefreshControls()

  cmdAdd.Enabled = Not mbIsReadOnly
  cmdEdit.Enabled = Not mbIsReadOnly And (filOLEs.ListCount > 0)
  cmdNone.Enabled = Not mbIsReadOnly
  cmdSelect.Enabled = Not mbIsReadOnly And (filOLEs.ListCount > 0)
  filOLEs.Enabled = Not mbIsReadOnly And (filOLEs.ListCount > 0)
  cmdCancel.Enabled = True

End Sub

