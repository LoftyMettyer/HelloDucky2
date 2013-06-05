Attribute VB_Name = "modHrProSecurity"
Option Explicit

' Public Objects
'Public rdoEnv As RDO.rdoEnvironment                     ' The Remote Data Object Environment
'Public rdoCon As RDO.rdoConnection                  ' The Remote Data Object Connection
Public gADOCon As ADODB.Connection

Public giSQLServerAuthenticationType As SecurityMgr.SQLServerAuthenticationType
Public gsSQLServerName As String
'Public glngSQLVersion As Long                          ' Database version
Public glngSQLVersion As Double     'changed from Long to Double for SQL2008 R2 (i.e. version 10.5)
Public gstrSQLFullVersion As String
Public gsConnectString As String
Public gsActualSQLLogin As String
Public gsUserName As String
Public gsPassword As String
Public gsDatabaseName As String
Public gsServerName As String
Public gsUserGroup As String
Public gbShiftSave As Boolean

Public gbCanUseWindowsAuthentication As Boolean
Public gbUseWindowsAuthentication As Boolean
Public gstrWindowsCurrentDomain As String
Public gstrWindowsCurrentUser As String
Public gstrServerDefaultDomain As String

'Public classes
Public Application As SecurityMgr.Application     ' Application Class
Public Database As SecurityMgr.Database           ' Database Class
Public UI As SecurityMgr.UI                       ' User interface class
'Public gobjProgress As COAProgress.COA_Progress
Public gobjProgress As clsProgress

'MH20060427
'''Public gobjCurrentUser As SecurityMgr.clsUser     ' Logged on user information
Public ASRDEVELOPMENT As Boolean                       ' Running in VB or as EXE
Public gobjNET As New SecurityMgr.Net             ' Useful network functions

' Net API constants (Used for building domain info)
Public Const FILTER_TEMP_DUPLICATE_ACCOUNT As Long = &H1&
Public Const FILTER_NORMAL_ACCOUNT As Long = &H2&
Public Const FILTER_PROXY_ACCOUNT As Long = &H4&
Public Const FILTER_INTERDOMAIN_TRUST_ACCOUNT As Long = &H8&
Public Const FILTER_WORKSTATION_TRUST_ACCOUNT As Long = &H10&
Public Const FILTER_SERVER_TRUST_ACCOUNT As Long = &H20&

'Windows API functions
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Sub Main()

  
  Dim plngPause As Long
  
  ASRDEVELOPMENT = Not vbCompiled

  ' If we get problems, just in case...
  gbDisableCodeJock = (InStr(LCase(Command$), "/skin=false") > 0)

  ' Default logged on user information
  gstrServerDefaultDomain = Environ("USERDOMAIN")
  gstrWindowsCurrentDomain = Environ("USERDOMAIN")
  gstrWindowsCurrentUser = Environ("USERNAME")

  ' Instantiate the Application class.
  Set Application = New SecurityMgr.Application

  ' Instantiate the Database class.
  Set Database = New SecurityMgr.Database
  
  'Instantiate Progress Bar class
  'Set gobjProgress = New COAProgress.COA_Progress
  Set gobjProgress = New clsProgress
  gobjProgress.StyleResource = CodeJockStylePath
  gobjProgress.StyleIni = CodeJockStyleIni
  
  ' Instantiate the User Interface class.
  Set UI = New SecurityMgr.UI
  
  'MH20060427
  ''' Create Current User class
  '''Set gobjCurrentUser = New SecurityMgr.clsUser
 
  If App.StartMode = vbSModeAutomation Then
    ' If started via OLE automation, return control back to client application.
    Exit Sub
  ElseIf App.StartMode = vbSModeStandalone Then
    ' Login to database
    If Application.Login Then
    
      ' Load the Desktop settings
      glngDesktopBitmapID = GetSystemSetting("DesktopSetting", "BitmapID", 0)
      glngDesktopBitmapLocation = GetSystemSetting("DesktopSetting", "BitmapLocation", 0)
      glngDeskTopColour = GetSystemSetting("DesktopSetting", "BackgroundColour", &H8000000C)
    
      ' Display the splash screen.
      frmSplash.Show
      frmSplash.Refresh
      
      
      
      'MH20011217 HHHhhhmmmm... this doesn't seem too good really..
      'Customers using Citrix say that the Security Manager is slow
      '(takes about 7 or 8 minutes to log in!) but the other modules
      'are fine.  Could it be due to this WELL DODGEY loop??
      
      '' Put a little pause in here because the splash screen hardly shows...
      '' looks really silly flashing up for half a second.
      'For plngPause = 0 To 200000
      '  DoEvents
      'Next plngPause
      
      'I'll change it to this to see if this helps...
      Dim dblEndLoop As Double
      dblEndLoop = Timer + 1
      Do While Timer < dblEndLoop
        DoEvents
      Loop
      
      
      
      ' Load the module settings
      SetupModuleParameters
      
      ' Activate the system.
      Application.Activate
  
      ' Unload the splash screen.
      Unload frmSplash
      
      Screen.MousePointer = vbNormal
      
    End If
  End If

End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@ Function  : FillCombo
'@@
'@@ Desc      : This function will fill a combo from a sql statement
'@@
'@@ Params    : psSql   (The sql string)
'@@           : pCombo  (The combo box)
'@@
'@@ Returns   : Boolean indicating success or failure
'@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@ Changes   :
'@@ 06/07/1998  RJB   Created

Function FillCombo(pCombo As ComboBox, Optional psSQL As String, Optional paArray) As Boolean
  Dim rsRecords As New ADODB.Recordset
  
  Dim iX As Integer
  
  On Error GoTo err_FillCombo
  
  ' Clear the contents of the supplied combo box
  pCombo.Clear
  
  ' Decide if it is a sql statement fill or an array fill
  If Not IsMissing(psSQL) And psSQL <> "" Then
    ' Open the required resultset
    rsRecords.Open psSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    
    ' Iterate through the resultset adding items to the combo box
    While Not rsRecords.EOF
      pCombo.AddItem rsRecords.Fields(0).Value
      rsRecords.MoveNext
    Wend
    
    ' Close and release the resultset
    rsRecords.Close
    Set rsRecords = Nothing
  End If
  
  If Not IsMissing(paArray) Then
    ' Add each element in the array to the combo box
    For iX = LBound(paArray) To UBound(paArray)
      pCombo.AddItem paArray(iX)
    Next iX
  End If
  
  ' Set the return value to true
  FillCombo = True
  
Exit Function

err_FillCombo:
  Dim lErrorNo As Long
  Dim sErrorMsg As String
  
  
  lErrorNo = Err.Number
  sErrorMsg = Err.Description
   
  If lErrorNo = 9 And pCombo.Name = "cboDomain" Then
    'NHRD20100413 This is a generic FillCombo Fucntion so put this specific checking in for Win Authent Error
    'paArray wasnt being populated when running local and there was no network to check for Domains
    MsgBox "No Domain values exist for this dropdown. You may be disconnected from the Network.", vbCritical + vbOKOnly, App.Title
  Else
    MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
  End If
  
  FillCombo = False
End Function


Public Function ValidNameChar(ByVal piAsciiCode As Integer, ByVal piPosition As Integer) As Integer
  ' Validate the characters used to create table and column names.
  On Error GoTo ErrorTrap
  
  ' Space character is valid in SQL Server 7.0 but not 6.5.
  If piAsciiCode = Asc(" ") And (Not IsVersion7) Then
    ' Substitute underscores for spaces.
    If piPosition <> 0 Then
      piAsciiCode = Asc("_")
    Else
      piAsciiCode = 0
    End If
  Else
    ' Allow only pure alpha-numerics and underscores.
    ' Do not allow numerics in the first chracter position.
    If Not (piAsciiCode = 8 Or _
      piAsciiCode = Asc("_") Or _
      (piAsciiCode >= Asc("0") And piAsciiCode <= Asc("9") And piPosition <> 0) Or _
      (piAsciiCode >= Asc("A") And piAsciiCode <= Asc("Z")) Or _
      (piAsciiCode >= Asc("a") And piAsciiCode <= Asc("z")) Or _
      (piAsciiCode = Asc(" ") And IsVersion7 And piPosition <> 0)) Then
      piAsciiCode = 0
    End If
  End If
  
  ValidNameChar = piAsciiCode
  Exit Function
  
ErrorTrap:
  ValidNameChar = 0
  Err = False
  
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@ Function  : IsVersion7()
'@@
'@@ Desc      : Sees if the sql version we are talking to is version 7
'@@
'@@ Returns   : True if it is, and false otherwise
'@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@ Changes   :
'@@ 28/07/98    RJB   Created
Public Function IsVersion7() As Boolean
  IsVersion7 = (glngSQLVersion >= 7)
  
End Function

Public Sub AuditPermission(ByVal sGroupName As String, ByVal sViewTableName As String, sAction As String, _
    sPermission As String, Optional ByVal sColumnName As String)
    
    On Error GoTo Err_Trap
    
    Dim sSQL As String
    
    sGroupName = RemoveBrackets(sGroupName)
    sViewTableName = RemoveBrackets(sViewTableName)
    sColumnName = RemoveBrackets(sColumnName)
        
    'JPD 20050812 Fault 10166
    sSQL = "INSERT INTO ASRSysAuditPermissions" & _
      " (userName, dateTimeStamp, groupName, viewTableName, columnName, action, permission)" & _
      " VALUES('" & Replace(gsUserName, "'", "''") & "', getdate(), '" & sGroupName & _
      "', '" & sViewTableName & "', '" & sColumnName & "', '" & sAction & "', '" & sPermission & "')"
    gADOCon.Execute sSQL, , adExecuteNoRecords
    
    Exit Sub
    
Err_Trap:
    MsgBox "Error auditing Table/Column Permissions.", vbExclamation, "Error"
    Exit Sub
        
End Sub

Public Sub AuditGroup(ByVal sGroupName As String, ByVal sAction As String, Optional ByVal sUserLogin As String)

    On Error GoTo Err_Trap
    
    Dim sSQL As String
    
    sGroupName = RemoveBrackets(sGroupName)
    sUserLogin = RemoveBrackets(sUserLogin)
    
    'JPD 20050812 Fault 10166
    sSQL = "INSERT INTO ASRSysAuditGroup" & _
      " (userName, dateTimeStamp, groupName, userLogin, action)" & _
      " VALUES('" & Replace(gsUserName, "'", "''") & "', getdate(), '" & sGroupName & _
      "', '" & Replace(sUserLogin, "'", "''") & "', '" & sAction & "')"
    gADOCon.Execute sSQL, , adExecuteNoRecords
    
    Exit Sub
    
Err_Trap:
    MsgBox "Error auditing User Group changes.", vbExclamation, "Error"
    Exit Sub

End Sub

Private Function RemoveBrackets(sString As String) As String

    If Left(sString, 1) = "[" Then
        sString = Mid$(sString, 2, Len(sString) - 2)
    End If
    RemoveBrackets = sString
    
End Function

'Public Sub LockUsers(pTool As ActiveBarLibraryCtl.Tool)
'  ' toggle the user lock.
'  Dim sSQL As String
'  Dim sMsg As String
'
'  If pTool.Checked = False Then
'    sSQL = "exec sp_ASRSetLock '" & gsUserName & "'"
'    gADOCon.Execute sSQL, , adExecuteNoRecords
'    sMsg = "All other users are now locked out of OpenHR."
'    pTool.Caption = "Un&lock Users"
'    pTool.Checked = True
'  Else
'    sSQL = "exec sp_ASRRemoveLock"
'    gADOCon.Execute sSQL, , adExecuteNoRecords
'    sMsg = "All users are now able to use OpenHR."
'    pTool.Caption = "&Lock Users"
'    pTool.Checked = False
'  End If
'
'  MsgBox sMsg, vbInformation
'  frmMain.RefreshMenu False
'
'End Sub
'
'Public Function GetUserLock() As Boolean
'  ' Return TRUE if the users are currentlty locked.
'  Dim sSQL As String
'
'  sSQL = "exec sp_ASRGetLockInfo"
'
'  Set rsTemp = rdoCon.OpenResultset(sSQL)
'  GetUserLock = (rsTemp(0) <> "No Lock")
'  rsTemp.Close
'  Set rsTemp = Nothing
'
'End Function

Private Function vbCompiled() As Boolean
  
  On Local Error Resume Next
  Err.Clear
  Debug.Print 1 / 0
  vbCompiled = (Err.Number = 0)

End Function
Public Function GetPictureFromDatabase(plngImageID As Long) As String

  Dim strTempName As String
  Dim intFileNo As Integer
  Dim lngColSize As Long
  Dim ChunkSize As Long
  Dim Fragment As Integer
  Dim i As Integer
  Dim Chunks As Integer
  Dim Chunk() As Byte
  Dim TempFile As Integer
  Dim recPictures As ADODB.Recordset

  ChunkSize = 2 ^ 14
  strTempName = ""

  Set recPictures = Database.GetPicture(plngImageID)

  If Not recPictures Is Nothing Then
  
    With recPictures
      lngColSize = !Picture.ActualSize
      If lngColSize > 0 Then
        strTempName = GetTmpFName
        intFileNo = FreeFile(1)
        Open strTempName For Binary Access Write As intFileNo
      
        Chunks = lngColSize \ ChunkSize
        Fragment = lngColSize Mod ChunkSize
        
        ReDim Chunk(Fragment)
        Chunk() = !Picture.GetChunk(Fragment)
        Put intFileNo, , Chunk()
        
        For i = 1 To Chunks
          ReDim Chunk(ChunkSize)
          Chunk() = !Picture.GetChunk(ChunkSize)
          Put intFileNo, , Chunk()
        Next i

        Close intFileNo

      End If
    End With
  End If

  recPictures.Close
  Set recPictures = Nothing

  GetPictureFromDatabase = strTempName
  
End Function

Public Function RemoveUnderScores(ByRef sName As String) As String

    Dim lPos As Long
    Dim sTableName As String
    
    sTableName = sName
    lPos = InStr(1, sTableName, "_")
    If lPos = 0 Then
        RemoveUnderScores = sTableName
        Exit Function
    End If
    
    Do While lPos > 0
        sTableName = Mid$(sTableName, 1, lPos - 1) & " " & Mid$(sTableName, lPos + 1, Len(sTableName))
        lPos = InStr(lPos, sTableName, "_")
    Loop
    RemoveUnderScores = sTableName
    
End Function


Public Function ValidateGTMaskDate(dtTemp As GTMaskDate.GTMaskDate) As Boolean

  Dim blnYearOkay As Boolean
  Dim sSysDateSeparator As String

  ValidateGTMaskDate = True

  sSysDateSeparator = UI.GetSystemDateSeparator
  
  With dtTemp
    If Trim(Replace(.Text, sSysDateSeparator, "")) <> vbNullString Then
  
      'MH20020423 Fault 3760 (Avoid changing 01/13/2002 to 13/01/2002)
      'If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Then
      'If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Or Left(.Text, 5) <> Left(.DateValue, 5) Then
      
      'MH20020423 Fault 3543 Also make sure that they enter a valid year
      blnYearOkay = (Val(Mid(.Text, InStrRev(.Text, sSysDateSeparator) + 1)) >= 1753)
      
      If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Or _
          Format(.DateValue, DateFormat) <> .Text Or Not blnYearOkay Then

        Clipboard.Clear
        Clipboard.SetText .Text
        .DateValue = Null
        .Paste
  
        .ForeColor = vbRed
        MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
        .ForeColor = vbWindowText
        .DateValue = Null
        If .Visible And .Enabled Then
          .SetFocus
        End If
        ValidateGTMaskDate = False
  
      End If
    End If
  End With

End Function

Public Function UDFFunctions(pastrUDFFunctions() As String, pbCreate As Boolean) As Boolean

  On Error GoTo UDFFunctions_ERROR

  Dim iCount As Integer
  Dim strDropCode As String
  Dim strFunctionName As String
  Dim sUDFCode As String
                       
  If gbEnableUDFFunctions Then
            
    For iCount = 1 To UBound(pastrUDFFunctions)
    
      strFunctionName = Mid(pastrUDFFunctions(iCount), 17, 15)
    
      'Drop existing function (could exist if the expression is used more than once in a report)
      strDropCode = "IF EXISTS" & _
        " (SELECT Name" & _
        "   FROM sysobjects" & _
        "   WHERE id = object_id('" & strFunctionName & "')" & _
        "     AND sysstat & 0xf = 0)" & _
        " DROP FUNCTION " & strFunctionName
      gADOCon.Execute strDropCode, , adExecuteNoRecords
    
      ' Create the new function
      If pbCreate Then
        sUDFCode = pastrUDFFunctions(iCount)
        gADOCon.Execute sUDFCode, , adExecuteNoRecords
      End If
    
    Next iCount
  End If

  UDFFunctions = True
  Exit Function
  
UDFFunctions_ERROR:
  UDFFunctions = False
  
End Function


