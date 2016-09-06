Attribute VB_Name = "modWorkflowSpecifics"
Option Explicit

Private mfInitTrue As Boolean
Private mabytArray() As Byte
Private mlngHiByte As Long
Private mlngHiBound As Long
Private mabytAddTable(255, 255) As Byte
Private mabytXTable(255, 255) As Byte

' WORKFLOW MODULE CONSTANTS
Public Const gsMODULEKEY_WORKFLOW = "MODULE_WORKFLOW"
' Parameter Type constants.
Public Const gsPARAMETERKEY_URL = "Param_URL"
Public Const gsPARAMETERKEY_WEBPARAM1 = "Param_Web1"
Public Const gsPARAMETERKEY_EMAILCOLUMN = "Param_EmailColumn"

Public Enum WorkflowStatus
  giWFSTATUS_ALL = -1
  giWFSTATUS_INPROGRESS = 0
  giWFSTATUS_CANCELLED = 1
  giWFSTATUS_ERROR = 2
  giWFSTATUS_COMPLETED = 3
  giWFSTATUS_SCHEDULED = 4
End Enum

Public Enum WorkflowStepStatus
  giWFSTEPSTATUS_ALL = -1
  giWFSTEPSTATUS_ONHOLD = 0
  giWFSTEPSTATUS_PENDINGENGINEACTION = 1
  giWFSTEPSTATUS_PENDINGUSERACTION = 2
  giWFSTEPSTATUS_COMPLETED = 3
  giWFSTEPSTATUS_FAILED = 4
  giWFSTEPSTATUS_INPROGRESS = 5
  giWFSTEPSTATUS_TIMEOUT = 6
  giWFSTEPSTATUS_PENDINGUSERCOMPLETION = 7
  giWFSTEPSTATUS_FAILEDACTION = 8
End Enum


Public Enum ElementType
  elem_Begin = 0
  elem_Terminator = 1
  elem_WebForm = 2
  elem_Email = 3
  elem_Decision = 4
  elem_StoredData = 5
  elem_SummingJunction = 6
  elem_Or = 7
  elem_Connector1 = 8
  elem_Connector2 = 9
End Enum


Public Sub CheckWorkflowOutOfOffice()
  ' Check if the user is flagged as being OutOfOffice for Workflows.
  ' If they are, ask them if they want to turn it off.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fOutOfOffice As Boolean
  Dim fTurnItOff As Boolean
  Dim sMsg As String
  Dim iRecordCount As Integer
  
  fOutOfOffice = False
  
  fOK = gbWorkflowOutOfOfficeEnabled
  
  If fOK Then
    ' Both of the required stored procedures exist, so check if the current user is OutOfOffice
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "spASRWorkflowOutOfOfficeCheck"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon

      Set pmADO = .CreateParameter("OutOfOffice", adBoolean, adParamOutput)
      .Parameters.Append pmADO

      Set pmADO = .CreateParameter("RecordCount", adInteger, adParamOutput)
      .Parameters.Append pmADO

      Set pmADO = Nothing

      .Execute
      
      fOutOfOffice = .Parameters(0).Value
      iRecordCount = .Parameters(1).Value
    End With
    Set cmADO = Nothing
  End If

  If fOutOfOffice Then
    ' If the current user IS OutOfOffice, ask them if they want to turn it off.
    sMsg = "Workflow Out of Office is currently on." & vbCrLf & "Would you like to turn it off"

    If iRecordCount > 1 Then
      If iRecordCount = 2 Then
        sMsg = sMsg & " for both"
      Else
        sMsg = sMsg & " for all " & CStr(iRecordCount)
      End If
      
      sMsg = sMsg & " of your identified personnel records"
    End If

    sMsg = sMsg & "?"

    fTurnItOff = (COAMsgBox(sMsg, vbQuestion + vbYesNo, "Workflow") = vbYes)

    If fTurnItOff Then
      sSQL = "EXEC spASRWorkflowOutOfOfficeSet 0"
      datGeneral.ExecuteSql sSQL, ""
    End If
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub


Public Sub CheckPendingWorkflowSteps(pfFromMenu As Boolean)
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim sEmailAddresses As String
  Dim rsTemp As ADODB.Recordset
  Dim rsTemp2 As ADODB.Recordset
  Dim avPendingForms() As Variant
  Dim alngSelectedSteps() As Variant
  Dim frmPrompt As frmDefSel
  Dim strPendingStepIDs As String
  Dim fExit As Boolean
  Dim lngCount As Long
  Dim lngCount2 As Long
  Dim sURL As String
  Dim strExePath As String
  Dim fIsDLL As Boolean
  Dim sGetEmailSQL As String
  Dim lngColumnID As Long
  Dim sTableName As String
  Dim sLoginTableName As String
  Dim sLoginColumnName As String
  Dim lngLastAction As Long
  
  sEmailAddresses = ""
  lngLastAction = 0
  
  sURL = WorkflowURL
  strExePath = GetDefaultBrowserApplication(fIsDLL)
    
  sSQL = "SELECT COUNT(*) AS objectCount" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('spASRCheckPendingWorkflowSteps')" & _
    "     AND sysstat & 0xf = 4"
  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If rsTemp!objectCount = 0 Then
    If pfFromMenu Then
      COAMsgBox "Unable to check for pending workflow steps. Contact your system administrator.", vbInformation + vbOKOnly, "Workflow"
    End If
  
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Sub
  End If
    
  rsTemp.Close
  Set rsTemp = Nothing
    
  If Len(sURL) = 0 Then
    If pfFromMenu Then
      COAMsgBox "No Workflow URL has been configured. Contact your system administrator.", vbInformation + vbOKOnly, "Workflow"
    End If
    Exit Sub
  ElseIf Len(Trim(strExePath)) = 0 Then
    If pfFromMenu Then
      COAMsgBox "Unable to open Workflow forms." & vbCrLf & vbCrLf & "No default browser application.", vbExclamation + vbOKOnly, "Workflow"
    End If
    Exit Sub
  End If
  
  ' Display the screen listing the pending steps.
  Set frmPrompt = New frmDefSel
  Screen.MousePointer = vbDefault

  Do
    strPendingStepIDs = ""
    
    ReDim avPendingForms(2, 0)
    'Column 0 = Instance ID
    'Column 1 = Element ID
    'Column 2 = Instance Step ID
    
    sSQL = "exec spASRCheckPendingWorkflowSteps"
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    
    With rsTemp
      Do While Not .EOF
        ReDim Preserve avPendingForms(2, UBound(avPendingForms, 2) + 1)
    
        avPendingForms(0, UBound(avPendingForms, 2)) = !instanceID
        avPendingForms(1, UBound(avPendingForms, 2)) = !elementID
        avPendingForms(2, UBound(avPendingForms, 2)) = !ID
        
        strPendingStepIDs = strPendingStepIDs & _
          IIf(Len(strPendingStepIDs) > 0, ",", "") & _
          CStr(!ID)
        
        .MoveNext
      Loop
      
      .Close
    End With
    Set rsTemp = Nothing
        
    fExit = True
    
    ' Display the list of workflow instance steps that are pending for the current user.
    If (UBound(avPendingForms, 2) > 0) Then
      frmPrompt.Options = edtSelect + edtRefresh
      frmPrompt.EnableRun = True
      frmPrompt.BatchPrompt = True
      frmPrompt.EventLogIDs = strPendingStepIDs
      frmPrompt.ShowList utlWorkflow, , True

      If frmPrompt.ListCount > 0 Then
        frmPrompt.Show vbModal
        
        If (frmPrompt.Action = edtSelect) Then
          ' Launch the selected workflow instance steps.
          alngSelectedSteps = frmPrompt.SelectedIDs
          
          For lngCount = 1 To UBound(alngSelectedSteps)
            For lngCount2 = 1 To UBound(avPendingForms, 2)
              If alngSelectedSteps(lngCount) = CLng(avPendingForms(2, lngCount2)) Then
                ' Launch IE with the required web form.
                OpenWebForm CLng(avPendingForms(0, lngCount2)), CLng(avPendingForms(1, lngCount2))
                
                ' NPG20120430 Fault HRPRO-2197
                Sleep 1000
                
                Exit For
              End If
            Next lngCount2
          Next lngCount
        
          'JPD 20090526 Fault 13679
          'If UBound(alngSelectedSteps) > 0 Then
          '  COAMsgBox "Workflow form" & IIf(UBound(alngSelectedSteps) > 1, "s", "") & " opened successfully.", vbInformation + vbOKOnly, "Workflow"
          'End If
        End If
        
        fExit = (frmPrompt.Action = edtCancel)
        lngLastAction = frmPrompt.Action
      End If
    ElseIf pfFromMenu Or (lngLastAction = edtRefresh) Then
      COAMsgBox "No workflow steps pending your action.", vbInformation + vbOKOnly, "Workflow"
    End If
  Loop While Not fExit
    
  Unload frmPrompt
  Set frmPrompt = Nothing
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub


Private Function CompactString(psSourceString As String) As String
  ' Compact the encrypted string.
  ' psSourceString is a string of the hexadecimal values of the Ascii codes for each character in the encrypted string.
  ' In this string each character in the encrypted string is represented as 2 hex digits.
  ' As it's a string of hex characters all characters are in the range 0-9, A-F
  ' Valid hypertext link characters are 0-9, A-Z, a-z and some others (we'll be using $ and @).
  ' Take advantage of this by implementing our own base64 encoding as follows:
  Dim sCompactedString As String
  Dim sSubString As String
  Dim sModifiedSourceString As String
  Dim iValue As Integer
  Dim iTemp As Integer
  Dim sNewString As String
  
  sCompactedString = ""
  sModifiedSourceString = psSourceString
  Do While Len(sModifiedSourceString) > 0
    ' Read the hex characters in chunks of 3 (ie. possible values 0 - 4095)
    ' This chunk of 3 Hex characters can then be translated into 2 base64 characters (ie. still have possible values 0 - 4095)
    ' Woohoo! We've reduced the length of the encrypted string by about one third!
    sNewString = ""
    sSubString = Left(sModifiedSourceString & "000", 3)
    sModifiedSourceString = Mid(sModifiedSourceString, 4)
    iValue = Val("&H" & sSubString)
    
    ' Use our own base64 digit set.
    ' Base64 digit values 0-9 are represented as 0-9
    ' Base64 digit values 10-35 are represented as A-Z
    ' Base64 digit values 36-61 are represented as a-z
    ' Base64 digit value 62 is represented as $
    ' Base64 digit value 63 is represented as @
    
    iTemp = iValue Mod 64
    If iTemp = 63 Then
      sNewString = "@"
    ElseIf iTemp = 62 Then
      sNewString = "$"
    ElseIf iTemp >= 36 Then
      sNewString = Chr(iTemp + 61)
    ElseIf iTemp >= 10 Then
      sNewString = Chr(iTemp + 55)
    Else
      sNewString = Chr(iTemp + 48)
    End If
    
    iTemp = (iValue - iTemp) / 64
    
    If iTemp = 63 Then
      sNewString = "@" & sNewString
    ElseIf iTemp = 62 Then
      sNewString = "$" & sNewString
    ElseIf iTemp >= 36 Then
      sNewString = Chr(iTemp + 61) & sNewString
    ElseIf iTemp >= 10 Then
      sNewString = Chr(iTemp + 55) & sNewString
    Else
      sNewString = Chr(iTemp + 48) & sNewString
    End If
    
    sCompactedString = sCompactedString & sNewString
  Loop
 
  ' Append the number of characters to ignore, to the compacted string
  CompactString = sCompactedString & CStr((3 - (Len(psSourceString) Mod 3)) Mod 3)
  
End Function

Public Function OutOfOfficeEnabled() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fWorkflowOutOfOfficeEnabled As Boolean
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  
  fWorkflowOutOfOfficeEnabled = False
  
  ' Check if the SP that checks if the current user is OutOfOffice exists
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRWorkflowOutOfOfficeConfigured"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("OutOfOffice", adBoolean, adParamOutput)
    .Parameters.Append pmADO

    Set pmADO = Nothing

    .Execute
    fWorkflowOutOfOfficeEnabled = .Parameters(0).Value
  End With
  Set cmADO = Nothing

TidyUpAndExit:
  OutOfOfficeEnabled = fWorkflowOutOfOfficeEnabled
  Exit Function
  
ErrorTrap:
  fWorkflowOutOfOfficeEnabled = False
  Resume TidyUpAndExit
    
End Function


Private Function ProcessEncryptString(psString) As String
  On Error GoTo ErrorTrap
  
  Dim sOutput As String
  Dim lngLoop As Long
  Dim iAscCode As Integer
  Dim sChar As String
  Dim sOutputPreProcess  As String
  Dim sSubTemp As String
  
  Const MARKERCHAR_1 = "J"
  Const MARKERCHAR_2 = "P"
  Const MARKERCHAR_3 = "D"
  Const DODGYCHARACTER_INCREMENT_1 = 174
  Const DODGYCHARACTER_INCREMENT_2 = 83
  Const DODGYCHARACTER_INCREMENT_3 = 1
  
  sOutputPreProcess = Replace(psString, MARKERCHAR_1, MARKERCHAR_1 & MARKERCHAR_1)
  sOutputPreProcess = Replace(sOutputPreProcess, MARKERCHAR_2, MARKERCHAR_2 & MARKERCHAR_2)
  sOutputPreProcess = Replace(sOutputPreProcess, MARKERCHAR_3, MARKERCHAR_3 & MARKERCHAR_3)

  sOutput = ""
  lngLoop = 1

  ' Loop through the output replacing dodgy characters with a MARKERCHAR and a safe character offset from the dodgy character.
  ' This is to avoid the dodgy characters messing up the querystring when used in a link to the Workflow website.
  Do While lngLoop <= Len(sOutputPreProcess)
    ' Process the next character.
    sChar = Mid(sOutputPreProcess, lngLoop, 1)
    iAscCode = Asc(sChar)
    
    If (iAscCode <= 32) _
      Or (iAscCode = 34) _
      Or (iAscCode = 35) _
      Or (iAscCode = 37) _
      Or (iAscCode = 60) _
      Or (iAscCode = 62) Then
      
      ' Dodgy character. Must replace with the MARKERCHAR_1 and a different character that we know is OK.
      ' Adding DODGYCHARACTER_INCREMENT_1 on the dodgy character's ASC value causes non-dodgy characters to be used.
      sOutput = sOutput & MARKERCHAR_1 & Chr(iAscCode + DODGYCHARACTER_INCREMENT_1)
      
    ElseIf (iAscCode = 91) _
      Or (iAscCode = 93) _
      Or (iAscCode = 94) _
      Or (iAscCode = 95) _
      Or (iAscCode = 96) _
      Or (iAscCode = 123) _
      Or (iAscCode = 125) _
      Or (iAscCode = 127) _
      Or (iAscCode = 129) _
      Or (iAscCode = 141) _
      Or (iAscCode = 143) _
      Or (iAscCode = 144) _
      Or (iAscCode = 157) _
      Or (iAscCode = 160) Then
      
      ' Dodgy character. Must replace with the MARKERCHAR_2 and a different character that we know is OK.
      ' Adding DODGYCHARACTER_INCREMENT_2 on the dodgy character's ASC value causes non-dodgy characters to be used.
      sOutput = sOutput & MARKERCHAR_2 & Chr(iAscCode + DODGYCHARACTER_INCREMENT_2)
      
    ElseIf (iAscCode = 173) Then
      
      ' Dodgy character. Must replace with the MARKERCHAR_3 and a different character that we know is OK.
      ' Adding DODGYCHARACTER_INCREMENT_3 on the dodgy character's ASC value causes non-dodgy characters to be used.
      sOutput = sOutput & MARKERCHAR_3 & Chr(iAscCode + DODGYCHARACTER_INCREMENT_3)
      
    Else
      ' NOT a dodgy character. Put it straight in the output string with out reprocessing.
      sOutput = sOutput & sChar
    End If

    lngLoop = lngLoop + 1
  Loop
  
  ' Always end with a decent character to avoid training code characters from being chopped.
  ' Use a random character between 65 and 90 (all safe)
  Randomize
  sChar = Chr(Int((90 - 65 + 1) * Rnd + 65))
  sOutput = sOutput & sChar
  
TidyUpAndExit:
  ProcessEncryptString = sOutput
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function

Public Function WorkflowElementTypeName(piElementType As Integer) As String
  ' Return the name of the given workflow element type.
  Dim sElementTypeName As String
  
  sElementTypeName = "<unknown>"
  
  Select Case piElementType
    Case elem_Begin
      sElementTypeName = "Begin"
    Case elem_Terminator
      sElementTypeName = "Terminator"
    Case elem_WebForm
      sElementTypeName = "Web Form"
    Case elem_Email
      sElementTypeName = "Email"
    Case elem_Decision
      sElementTypeName = "Decision"
    Case elem_StoredData
      sElementTypeName = "Stored Data"
    Case elem_SummingJunction
      sElementTypeName = "And"
    Case elem_Or
      sElementTypeName = "Or"
    Case elem_Connector1
      sElementTypeName = "Connector (part 1)"
    Case elem_Connector2
      sElementTypeName = "Connector (part 2)"
  End Select

  WorkflowElementTypeName = sElementTypeName
  
End Function

Public Function WorkflowElementTypeFromName(psElementTypeName As String) As Integer
  ' Return the enum value of the given workflow element type.
  Dim iElementType As Integer
  
  iElementType = -1
  
  Select Case UCase(psElementTypeName)
    Case UCase(WorkflowElementTypeName(elem_Begin))
      iElementType = elem_Begin
    Case UCase(WorkflowElementTypeName(elem_Terminator))
      iElementType = elem_Terminator
    Case UCase(WorkflowElementTypeName(elem_WebForm))
      iElementType = elem_WebForm
    Case UCase(WorkflowElementTypeName(elem_Email))
      iElementType = elem_Email
    Case UCase(WorkflowElementTypeName(elem_Decision))
      iElementType = elem_Decision
    Case UCase(WorkflowElementTypeName(elem_StoredData))
      iElementType = elem_StoredData
    Case UCase(WorkflowElementTypeName(elem_SummingJunction))
      iElementType = elem_SummingJunction
    Case UCase(WorkflowElementTypeName(elem_Or))
      iElementType = elem_Or
    Case UCase(WorkflowElementTypeName(elem_Connector1))
      iElementType = elem_Connector1
    Case UCase(WorkflowElementTypeName(elem_Connector2))
      iElementType = elem_Connector2
  End Select

  WorkflowElementTypeFromName = iElementType
  
End Function
Public Sub WorkflowOutOfOffice()
  ' Set/reset the user as being OutOfOffice for Workflows.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fOutOfOffice As Boolean
  Dim iRecordCount As Integer
  Dim fToggle As Boolean
  Dim sMsg As String
  
  fOutOfOffice = False
  fToggle = False
  iRecordCount = 0
  
  ' Both of the required stored procedures exist, so check if the current user is OutOfOffice
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRWorkflowOutOfOfficeCheck"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("OutOfOffice", adBoolean, adParamOutput)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("RecordCount", adInteger, adParamOutput)
    .Parameters.Append pmADO

    Set pmADO = Nothing

    .Execute
    
    fOutOfOffice = .Parameters(0).Value
    iRecordCount = .Parameters(1).Value
  End With
  Set cmADO = Nothing

  If iRecordCount = 0 Then
    COAMsgBox "Unable to set Workflow Out of Office." & vbCrLf & "You do not have an identifiable personnel record.", vbInformation + vbOKOnly, "Workflow"
  Else
    If fOutOfOffice Then
      ' If the current user IS OutOfOffice, ask them if they want to turn it off.
      sMsg = "Workflow Out of Office is currently on." & vbCrLf & _
        "Would you like to turn it off"
      
      If iRecordCount > 1 Then
        If iRecordCount = 2 Then
          sMsg = sMsg & _
            " for both"
        Else
          sMsg = sMsg & _
            " for all " & CStr(iRecordCount)
        End If
      
        sMsg = sMsg & _
          " of your identified personnel records"
      End If

      sMsg = sMsg & _
        "?"
    Else
      sMsg = "Workflow Out of Office is currently off." & vbCrLf & _
        "Would you like to turn it on"
      
      If iRecordCount > 1 Then
        If iRecordCount = 2 Then
          sMsg = sMsg & _
            " for both"
        Else
          sMsg = sMsg & _
            " for all " & CStr(iRecordCount)
        End If
      
        sMsg = sMsg & _
          " of your identified personnel records"
      End If

      sMsg = sMsg & _
        "?"
    End If
  
    fToggle = (COAMsgBox(sMsg, vbQuestion + vbYesNo, "Workflow") = vbYes)
  End If
  
  If fToggle Then
    sSQL = "EXEC spASRWorkflowOutOfOfficeSet " & IIf(fOutOfOffice, "0", "1")
    datGeneral.ExecuteSql sSQL, ""
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit

End Sub


Public Function WorkflowStatusDescription(piStatus As WorkflowStatus) As String
  ' Return the textual description of the given Workflow status
  Dim sDescription As String
  
  sDescription = "<unknown>"
  
  Select Case piStatus
    Case giWFSTATUS_ALL
      sDescription = "<All>"
    
    Case giWFSTATUS_INPROGRESS
      sDescription = "In progress"
    
    Case giWFSTATUS_CANCELLED
      sDescription = "Cancelled"
    
    Case giWFSTATUS_ERROR
      sDescription = "Error"
    
    Case giWFSTATUS_COMPLETED
      sDescription = "Completed"
    
    Case giWFSTATUS_SCHEDULED
      sDescription = "Scheduled"
        
  End Select
  
  WorkflowStatusDescription = sDescription

End Function

Public Function WorkflowStepStatusDescription(piStatus As WorkflowStepStatus) As String
  ' Return the textual description of the given Workflow Step status
  Dim sDescription As String
  
  sDescription = "<unknown>"
  
  Select Case piStatus
    Case giWFSTEPSTATUS_ALL
      sDescription = "<All>"
    
    Case giWFSTEPSTATUS_ONHOLD
      sDescription = "On hold"
    
    Case giWFSTEPSTATUS_PENDINGENGINEACTION
      sDescription = "Pending engine action"
    
    Case giWFSTEPSTATUS_PENDINGUSERACTION
      sDescription = "Pending user action"
    
    Case giWFSTEPSTATUS_COMPLETED
      sDescription = "Completed"
    
    Case giWFSTEPSTATUS_FAILED
      sDescription = "Failed"
    
    Case giWFSTEPSTATUS_INPROGRESS
      sDescription = "In progress"
    
    Case giWFSTEPSTATUS_TIMEOUT
      sDescription = "Timeout"
    
    Case giWFSTEPSTATUS_PENDINGUSERCOMPLETION
      sDescription = "Pending user completion"
       
    Case giWFSTEPSTATUS_FAILEDACTION
      sDescription = "Failed action"
       
  End Select
  
  WorkflowStepStatusDescription = sDescription

End Function


Public Function WorkflowURL() As String
  ' Read the Workflow parameter values from the database into local variables.
  Dim sURL As String
  
  sURL = GetModuleParameter(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_URL)
  sURL = Trim(sURL)

  If UCase(Right(sURL, 5)) <> ".ASPX" _
    And Right(sURL, 1) <> "/" _
    And Len(sURL) > 0 Then
    
    sURL = sURL + "/"
  End If

  WorkflowURL = sURL
  
End Function






Public Function EncryptQueryString(plngInstanceID As Long, _
  plngStepID As Long, _
  psUser As String, _
  psPassword As String) As String
  
  On Error GoTo ErrorTrap
  
  Dim sKey As String
  Dim sEncryptedString As String
  Dim sSourceString As String
  Dim sServerName As String
  Dim sDBName As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  Const ENCRYPTIONKEY = "jmltn"
  
  ' (NPG)TFS 23821: Fetch database and server name from shared stored proc
  ' to ensure that they remain consistent between apps.
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRGetSQLMetaData"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("ServerName", adVarChar, adParamOutput, 128)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("DBName", adVarChar, adParamOutput, 128)
    .Parameters.Append pmADO

    Set pmADO = Nothing

    .Execute
    
    sServerName = .Parameters(0).Value
    sDBName = .Parameters(1).Value
  End With
  Set cmADO = Nothing

  sKey = ENCRYPTIONKEY
  sSourceString = CStr(plngInstanceID) & _
    vbTab & CStr(plngStepID) & _
    vbTab & psUser & _
    vbTab & psPassword & _
    vbTab & sServerName & _
    vbTab & sDBName
    
  sEncryptedString = EncryptString(sSourceString, sKey, True)
  sEncryptedString = CompactString(sEncryptedString)

TidyUpAndExit:
  EncryptQueryString = sEncryptedString
  Exit Function
  
ErrorTrap:
  sEncryptedString = ""
  Resume TidyUpAndExit
  
End Function


Public Function EncryptString(psText As String, _
  Optional psKey As String, _
  Optional pbOutputInHex As Boolean) As String
  
  Dim abytArray() As Byte
  Dim abytKey() As Byte
  Dim abytOut() As Byte

  psText = psText & " "
  abytArray() = StrConv(psText, vbFromUnicode)
  abytKey() = StrConv(psKey, vbFromUnicode)
  abytOut() = EncryptByte(abytArray(), abytKey())
  EncryptString = StrConv(abytOut(), vbUnicode, 2057) ' 2057 is the LocaleID for English-UK

  If pbOutputInHex = True Then EncryptString = EnHex(EncryptString)
  
End Function


Public Function EncryptByte(pabytText() As Byte, pabytKey() As Byte)
  Dim abytTemp() As Byte
  Dim iTemp As Integer
  Dim iLoop As Long
  Dim iBound As Integer
  
  Call InitTbl
  
  ReDim abytTemp((UBound(pabytText)) + 4)
  Randomize
  abytTemp(0) = Int((Rnd * 254) + 1)
  abytTemp(1) = Int((Rnd * 254) + 1)
  abytTemp(2) = Int((Rnd * 254) + 1)
  abytTemp(3) = Int((Rnd * 254) + 1)
  abytTemp(4) = Int((Rnd * 254) + 1)
  
  Call CopyMemory(abytTemp(5), pabytText(0), UBound(pabytText))
  
  ReDim pabytText(UBound(abytTemp)) As Byte
  pabytText() = abytTemp()
  ReDim abytTemp(0)
  iBound = (UBound(pabytKey) - 1)
  iTemp = 0
  
  For iLoop = 0 To UBound(pabytText) - 1
    If iTemp = iBound Then iTemp = 0
    pabytText(iLoop) = mabytXTable(pabytText(iLoop), mabytAddTable(pabytText(iLoop + 1), pabytKey(iTemp)))
    pabytText(iLoop + 1) = mabytXTable(pabytText(iLoop), pabytText(iLoop + 1))
    pabytText(iLoop) = mabytXTable(pabytText(iLoop), mabytAddTable(pabytText(iLoop + 1), pabytKey(iTemp + 1)))
    iTemp = iTemp + 1
  Next iLoop
  
  EncryptByte = pabytText()

End Function

Public Function EnHex(psData As String) As String
  Dim dblCount As Double
  Dim sTemp As String
  
  Reset
  
  For dblCount = 1 To Len(psData)
    sTemp = Hex$(Asc(Mid$(psData, dblCount, 1)))
    If Len(sTemp) < 2 Then sTemp = "0" & sTemp
    Append sTemp
  Next
  
  EnHex = GData
  
  Reset
  
End Function

Private Sub InitTbl()
  Dim i As Integer
  Dim J As Integer
  Dim k As Integer
  
  If mfInitTrue = True Then Exit Sub
  
  For i = 0 To 255
    For J = 0 To 255
      mabytXTable(i, J) = CByte(i Xor J)
      mabytAddTable(i, J) = CByte((i + J) Mod 255)
    Next J
  Next i
  
  mfInitTrue = True
  
End Sub

Private Sub Reset()
  mlngHiByte = 0
  mlngHiBound = 1024
  ReDim mabytArray(mlngHiBound)
  
End Sub

Private Sub Append(ByRef psStringData As String, Optional plngLength As Long)
  Dim lngDataLength As Long
  
  If plngLength > 0 Then
    lngDataLength = plngLength
  Else
    lngDataLength = Len(psStringData)
  End If
  
  If lngDataLength + mlngHiByte > mlngHiBound Then
    mlngHiBound = mlngHiBound + 1024
    ReDim Preserve mabytArray(mlngHiBound)
  End If
  
  CopyMemory ByVal VarPtr(mabytArray(mlngHiByte)), ByVal psStringData, lngDataLength
  mlngHiByte = mlngHiByte + lngDataLength
    
End Sub

Private Property Get GData() As String
  Dim sStringData As String
  
  sStringData = Space(mlngHiByte)
  CopyMemory ByVal sStringData, ByVal VarPtr(mabytArray(0)), mlngHiByte
  GData = sStringData
  
End Property

Public Function GetDefaultBrowserApplication(ByRef pfIsDLL As Boolean) As String
  ' Get the default browser application.
  On Error GoTo ErrorTrap
  
  Dim strExePath As String
  Dim lngTemp As Long
  Dim fIsDLL As Boolean
  Dim strTmpPath As String
  Dim intFileNo As Integer
  
  strExePath = Space(255)
  strTmpPath = Space(1024)
  Call GetTempPath(1024, strTmpPath)
  intFileNo = FreeFile(1)
  strTmpPath = Left(Trim(strTmpPath), Len(Trim(strTmpPath)) - 1)
  Open strTmpPath & "dummy.htm" For Binary Access Write As intFileNo
  Close intFileNo
  
  ' Get the executables path for the path & document filename
  lngTemp = FindExecutable("dummy.htm", strTmpPath, strExePath)
    
  Kill strTmpPath & "dummy.htm"

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
  End If

  pfIsDLL = fIsDLL
  GetDefaultBrowserApplication = strExePath
  Exit Function
  
ErrorTrap:
  GetDefaultBrowserApplication = ""
  
End Function


Public Function OpenWebForm(plngInstanceID As Long, plngStepID As Long) As Boolean
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim strExePath As String
  Dim dblTaskID As Double
  Dim fIsDLL As Boolean
  Dim sEncryptedString As String
  Dim sURL As String
  Dim sUser As String
  Dim sPassword As String
  
  sURL = WorkflowURL
  ReadWebLogon sUser, sPassword

  strExePath = GetDefaultBrowserApplication(fIsDLL)
    
  fOK = Len(Trim(strExePath)) > 1
  
  If fOK Then
    sEncryptedString = EncryptQueryString(plngInstanceID, plngStepID, sUser, sPassword)
    
    fOK = (Len(sEncryptedString) > 0)
  End If
  
  'Launch IE with the queryString ?lngInstanceID&plngStepID for the required form.
  If fOK Then
    ' Shell out the process and capture the ID
    dblTaskID = Shell("""" & strExePath & IIf(fIsDLL, "", Chr(34)) & " """ & sURL & "?" & sEncryptedString & IIf(fIsDLL, "", Chr(34)), vbMaximizedFocus)
    OpenProcess PROCESS_QUERY_INFORMATION, False, dblTaskID
  End If

TidyUpAndExit:
  OpenWebForm = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Public Sub ReadWebLogon(strUserName As String, strPassword As String)
  
  Dim strInput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngStart As Long
  Dim lngFinish As Long

  strInput = GetModuleParameter(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_WEBPARAM1)

  If strInput = vbNullString Then
    Exit Sub
  End If

  lngStart = Len(strInput) - 12
  strEKey = Mid(strInput, lngStart + 1, 10)
  strLens = Right(strInput, 2)
  strInput = XOREncript(Left(strInput, lngStart), strEKey)

  lngStart = 1
  lngFinish = Asc(Mid(strLens, 1, 1)) - 127
  strUserName = Mid(strInput, lngStart, lngFinish)

  lngStart = lngStart + lngFinish
  lngFinish = Asc(Mid(strLens, 2, 1)) - 127
  strPassword = Mid(strInput, lngStart, lngFinish)

End Sub

