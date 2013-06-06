Attribute VB_Name = "basAuditMain"
Option Explicit

Private mfrmAuditTrail As Form

Public Enum audType
    audRecords = 1
    audPermissions = 2
    audGroups = 3
    audAccess = 4
End Enum

Public Enum FilterOperators
  giFILTEROP_UNDEFINED = 0
  giFILTEROP_EQUALS = 1
  giFILTEROP_NOTEQUALTO = 2
  giFILTEROP_ISATMOST = 3
  giFILTEROP_ISATLEAST = 4
  giFILTEROP_ISMORETHAN = 5
  giFILTEROP_ISLESSTHAN = 6
  giFILTEROP_ON = 7
  giFILTEROP_NOTON = 8
  giFILTEROP_AFTER = 9
  giFILTEROP_BEFORE = 10
  giFILTEROP_ONORAFTER = 11
  giFILTEROP_ONORBEFORE = 12
  giFILTEROP_CONTAINS = 13
  giFILTEROP_IS = 14
  giFILTEROP_DOESNOTCONTAIN = 15
  giFILTEROP_ISNOT = 16
End Enum

'SQL DatType
Public Enum SQLDataType
  sqlUnknown = 0      ' ?
  sqlTypeOle = -4     ' OLE columns
  sqlBoolean = -7     ' Logic columns
  sqlNumeric = 2      ' Numeric columns
  sqlInteger = 4      ' Integer columns
  sqlDate = 11        ' Date columns
  sqlVarchar = 12     ' Character columns
  sqlVarBinary = -3   ' Photo columns
  sqlLongVarChar = -1 ' Working Pattern columns
End Enum

Public Function DateFormat() As String
  ' Returns the date format.
  ' NB. Windows allows the user to configure totally stupid
  ' date formats (eg. d/M/yyMydy !). This function does not cater
  ' for such stupidity, and simply takes the first occurence of the
  ' 'd', 'M', 'y' characters.
  Dim sSysFormat As String
  Dim sSysDateSeparator As String
  Dim sDateFormat As String
  Dim iLoop As Integer
  Dim fDaysDone As Boolean
  Dim fMonthsDone As Boolean
  Dim fYearsDone As Boolean
  
  fDaysDone = False
  fMonthsDone = False
  fYearsDone = False
  sDateFormat = ""
    
  sSysFormat = UI.GetSystemDateFormat
  sSysDateSeparator = UI.GetSystemDateSeparator
    
  ' Loop through the string picking out the required characters.
  For iLoop = 1 To Len(sSysFormat)
      
    Select Case Mid(sSysFormat, iLoop, 1)
      Case "d"
        If Not fDaysDone Then
          ' Ensure we have two day characters.
          sDateFormat = sDateFormat & "dd"
          fDaysDone = True
        End If
          
      Case "M"
        If Not fMonthsDone Then
          ' Ensure we have two month characters.
          sDateFormat = sDateFormat & "mm"
          fMonthsDone = True
        End If
          
      Case "y"
        If Not fYearsDone Then
          ' Ensure we have four year characters.
          sDateFormat = sDateFormat & "yyyy"
          fYearsDone = True
        End If
          
      Case Else
        sDateFormat = sDateFormat & Mid(sSysFormat, iLoop, 1)
    End Select
      
  Next iLoop
    
  ' Ensure that all day, month and year parts of the date
  ' are present in the format.
  If Not fDaysDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "dd"
  End If
    
  If Not fMonthsDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "mm"
  End If
    
  If Not fYearsDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "yyyy"
  End If
    
  ' Return the date format.
  DateFormat = sDateFormat
  
End Function

'Public Sub CloseConnection()
'
'  If Not gADOCon Is Nothing Then
'    gADOCon.Close
'    gADOCon.Cancel
'    Set gADOCon = Nothing
'  End If
'
'End Sub

'Public Function Connect(sConnectString As String) As Boolean
'  ' Connect to the database.
'  On Error GoTo Err_Trap
'
'  Set gADOCon = New ADODB.Connection
'  With gADOCon
'    .ConnectionString = sConnectString
'    .Provider = "SQLOLEDB"
'    .CommandTimeout = 0
'    .CursorLocation = adUseClient
'    .Mode = adModeReadWrite
'    .Open
'  End With
'
'  Connect = True
'
'  Exit Function
'
'Err_Trap:
'  Connect = False
'  MsgBox Err.Description
'
'End Function

Public Function GetAllRecords(piAuditType As audType, Optional psOrder As String) As Recordset

  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "exec spstat_getaudittrail " & piAuditType & ", " & IIf(psOrder = "", "''", "'" & psOrder & "'")
  Set rsTemp = New Recordset
  gADOCon.CursorLocation = adUseClient
  rsTemp.Open sSQL, gADOCon, adOpenDynamic, adLockReadOnly
  gADOCon.CursorLocation = adUseServer
  Set GetAllRecords = rsTemp
    
End Function

Public Function RemoveBrackets(sString As String) As String

    If Left(sString, 1) = "[" Then
        sString = Mid$(sString, 2, Len(sString) - 2)
    End If
    
    RemoveBrackets = sString

End Function

Public Function GetCleardownData() As Recordset

    Dim sSQL As String
    Dim rsTemp As Recordset
    
    sSQL = "Select * From ASRSysAuditCleardown"
    Set rsTemp = New Recordset
    rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set GetCleardownData = rsTemp

End Function

Public Sub DeleteCleardowns()

    Dim sSQL As String

    sSQL = "Delete From ASRSysAuditCleardown"
    gADOCon.Execute sSQL, , adCmdText

End Sub

Public Sub InsertCleardown(sType As String, lFrequency As Long, sPeriod As String)

    Dim sSQL As String

    sSQL = "INSERT INTO ASRSysAuditCleardown " & _
           "(type, frequency, period) " & _
           "VALUES('" & sType & "', " & _
           lFrequency & ", " & _
           "'" & sPeriod & "')"
    
    gADOCon.Execute sSQL, , adCmdText

End Sub


