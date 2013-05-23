Option Strict Off
Option Explicit On

Imports System.Globalization

Module modSettings

  Public Const VARCHAR_MAX_Size As Integer = 2147483646 'Yup one below the actual max, needs to be otherwise things go so awfully wrong, you don't believe me, well go on then, change it, see if I care!!!)

  Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer
  Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer

  Public Function GetUniqueID(ByRef strSetting As String, ByRef strTable As String, ByRef strColumn As String) As Integer

    Dim lngNewMethodID As Integer 'From ASRSysSettings
    Dim lngOldMethodID As Integer 'SELECT MAX ID

    lngOldMethodID = UniqueColumnValue(strTable, strColumn)
    'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    lngNewMethodID = GetSystemSetting("AutoID", strSetting, 0) + 1

    GetUniqueID = IIf(lngOldMethodID > lngNewMethodID, lngOldMethodID, lngNewMethodID)
    SaveSystemSetting("AutoID", strSetting, GetUniqueID)

  End Function

  Public Function SaveUserSetting(ByRef strSection As String, ByRef strKey As String, ByRef varSetting As Object) As Boolean

    Dim datData As clsDataAccess
    Dim strSQL As String

    datData = New clsDataAccess

    strSQL = "DELETE FROM ASRSysUserSettings " & " WHERE Section = '" & Replace(LCase(strSection), "'", "''") & "'" & " AND SettingKey = '" & Replace(LCase(strKey), "'", "''") & "'" & " AND UserName = '" & Replace(LCase(gsUsername), "'", "''") & "'"
    datData.ExecuteSql(strSQL)

    'UPGRADE_WARNING: Couldn't resolve default property of object varSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    strSQL = "INSERT ASRSysUserSettings " & "(Section, SettingKey, SettingValue, UserName) " & "VALUES " & "('" & Replace(LCase(strSection), "'", "''") & "'," & " '" & Replace(LCase(strKey), "'", "''") & "'," & " '" & Replace(CStr(varSetting), "'", "''") & "'," & " '" & Replace(LCase(gsUsername), "'", "''") & "')"
    datData.ExecuteSql(strSQL)

    'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    datData = Nothing

  End Function

  Public Function GetUserSetting(ByRef strSection As String, ByRef strKey As String, ByRef varDefault As Object) As Object

    Dim datData As clsDataAccess
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    datData = New clsDataAccess

    strSQL = "SELECT SettingValue FROM ASRSysUserSettings " & " WHERE UserName = '" & Replace(LCase(gsUsername), "'", "''") & "'" & " AND Section = '" & Replace(LCase(strSection), "'", "''") & "'" & " AND SettingKey = '" & Replace(LCase(strKey), "'", "''") & "'"
    rsTemp = datData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    With rsTemp
      If Not .BOF And Not .EOF Then
        'UPGRADE_WARNING: Couldn't resolve default property of object GetUserSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetUserSetting = rsTemp.Fields("SettingValue").Value
      Else
        'UPGRADE_WARNING: Couldn't resolve default property of object varDefault. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object GetUserSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetUserSetting = varDefault
      End If
    End With

    rsTemp.Close()

    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing
    'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    datData = Nothing

  End Function

  Public Function SaveSystemSetting(ByRef strSection As String, ByRef strKey As String, ByRef varSetting As Object) As Boolean

    Dim datData As clsDataAccess
    Dim strSQL As String

    datData = New clsDataAccess

    strSQL = "DELETE FROM ASRSysSystemSettings " & " WHERE Section = '" & LCase(strSection) & "'" & " AND SettingKey = '" & LCase(strKey) & "'"
    datData.ExecuteSql(strSQL)

    'UPGRADE_WARNING: Couldn't resolve default property of object varSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    strSQL = "INSERT ASRSysSystemSettings " & "(Section, SettingKey, SettingValue) " & "VALUES " & "('" & LCase(strSection) & "'," & " '" & LCase(strKey) & "'," & " '" & CStr(varSetting) & "')"
    datData.ExecuteSql(strSQL)

    'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    datData = Nothing

  End Function

  Public Function GetSystemSetting(ByRef strSection As String, ByRef strKey As String, ByRef varDefault As Object) As Object

    Dim datData As clsDataAccess
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo LocalErr

    datData = New clsDataAccess

    strSQL = "SELECT SettingValue FROM ASRSysSystemSettings " & " WHERE Section = '" & LCase(strSection) & "'" & " AND SettingKey = '" & LCase(strKey) & "'"
    rsTemp = datData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    With rsTemp
      If Not .BOF And Not .EOF Then
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetSystemSetting = rsTemp.Fields("SettingValue").Value
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If IsDBNull(GetSystemSetting) Then GetSystemSetting = vbNullString
      Else
        'UPGRADE_WARNING: Couldn't resolve default property of object varDefault. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetSystemSetting = varDefault
      End If
    End With

    rsTemp.Close()

    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing
    'UPGRADE_NOTE: Object datData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    datData = Nothing

    Exit Function

LocalErr:
    'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    GetSystemSetting = vbNullString

  End Function

  Public Function GetModuleParameter(ByRef psModuleKey As String, ByRef psParameterKey As String) As String
    ' Return the value of the given parameter.
    GetModuleParameter = datGeneral.GetModuleParameter(psModuleKey, psParameterKey)

  End Function

  Public Function DateFormat() As String
    ' Returns the date format.
    ' NB. Windows allows the user to configure totally stupid
    ' date formats (eg. d/M/yyMydy !). This function does not cater
    ' for such stupidity, and simply takes the first occurence of the
    ' 'd', 'M', 'y' characters.
    Dim sSysFormat As String
    Dim sSysDateSeparator As String
    Dim sDateFormat As String
    Dim iLoop As Short
    Dim fDaysDone As Boolean
    Dim fMonthsDone As Boolean
    Dim fYearsDone As Boolean

    fDaysDone = False
    fMonthsDone = False
    fYearsDone = False
    sDateFormat = ""

    sSysFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
    sSysDateSeparator = CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator

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
            sDateFormat = sDateFormat & "MM"
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

  Public Function GetTmpFName() As String

    Dim strTmpPath, strTmpName As String

    strTmpPath = Space(1024)
    strTmpName = Space(1024)

    Call GetTempPath(1024, strTmpPath)
    Call GetTempFileName(strTmpPath, "_T", 0, strTmpName)

    strTmpName = Trim(strTmpName)
    If Len(strTmpName) > 0 Then
      strTmpName = Left(strTmpName, Len(strTmpName) - 1)

      'MH20021227 For some reason a zero byte file is created... ANNOYING!
      'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
      If Dir(strTmpName) <> vbNullString Then
        Kill(strTmpName)
      End If

    Else
      strTmpName = vbNullString
    End If

    GetTmpFName = Trim(strTmpName)

  End Function

  Public Sub ProgramError(ByVal strProcedureName As String, ByVal objErr As ErrObject, ByVal lngErrLine As Integer)

    On Error GoTo 0

    Dim strErrorText As String

    With objErr
      strErrorText = vbCrLf & vbCrLf & "Runtime error in COAInt_Server.DLL" & vbCrLf & "Error number: " & Err.Number & vbCrLf & "Error description: " & Err.Description & vbCrLf & vbCrLf & "Procedure: " & strProcedureName & vbCrLf & "Line: " & lngErrLine & vbCrLf & "Thread Id: " & System.Threading.Thread.CurrentThread.ManagedThreadId
      My.Application.Log.WriteEntry(strErrorText, System.Diagnostics.TraceEventType.Error)
    End With
  End Sub
End Module