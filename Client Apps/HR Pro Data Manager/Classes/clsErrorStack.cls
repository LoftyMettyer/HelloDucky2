VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrorStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private miStackLevel As Integer
Private mastrCallStack() As String
Private mbEnabled As Boolean

Public Function PopStack()

  'Remove the latest function/procedure from the stack array
  miStackLevel = miStackLevel - 1
  If miStackLevel <= 0 Then
    miStackLevel = 0
  End If
  
  ReDim Preserve mastrCallStack(miStackLevel)

End Function

Public Function PushStack(ByVal pstrMethodHeader As String, Optional ByVal parrArguments)

  'Put the current function/procedure onto the stack array
  miStackLevel = miStackLevel + 1
  
  ReDim Preserve mastrCallStack(miStackLevel)
  If IsArray(parrArguments) Then
    ' Turn array into string
    mastrCallStack(miStackLevel) = CreateNameValueList(pstrMethodHeader, parrArguments)
  Else
    mastrCallStack(miStackLevel) = pstrMethodHeader
  End If

End Function

Private Sub Class_Initialize()

  ' Clear the class
  miStackLevel = 0
  mbEnabled = True

End Sub

Public Property Get Description()

  ' Generate the error description
  Dim iCount As Integer
  Dim strMessage As String


  strMessage = _
      "Customer:  " & gsCustomerName & vbCrLf & _
      "User:      " & gsUserName & vbCrLf & _
      "Computer:  " & UCase(GetComputerName()) & vbCrLf & _
      "Database:  " & gsDatabaseName & " on " & gsServerName & vbCrLf & _
      "Date/Time: " & CStr(Now) & vbCrLf & _
      "Version:   " & GetAppVersion & vbCrLf & _
      "Message:   " & Err.Description & vbCrLf & vbCrLf & _
      "Details:   "

  For iCount = 1 To UBound(mastrCallStack)
    strMessage = strMessage & mastrCallStack(iCount) & vbCrLf
  Next iCount

  ' Bolt on line number if setup
  strMessage = strMessage & IIf(Erl > 0, "Line:= " & CStr(Erl), "")
  
  Description = strMessage

End Property

Private Function GetComputerName() As String
    Dim sBuffer As String * 255, lLen As Long
    lLen = Len(sBuffer)
    If CBool(GetComputerNameAPI(sBuffer, lLen)) Then GetComputerName = Left$(sBuffer, lLen)
End Function

Private Function CreateNameValueList(ByVal strNames As String, ByVal arrValues) As String
    
    Dim arrNames() As String, i As Long, strList As String, strValue As String, strNameValue As String
    Dim iStartPos As Integer
    Dim iEndPos As Integer
    Dim strProcName As String
    
    On Error GoTo ErrHandler
    'arrValues maybe array of values, or a single value.
    If Not IsArray(arrValues) Then arrValues = Array(arrValues)
    
    'Strip stuff before brackets and after)
    iStartPos = InStr(1, strNames, "(")
    iEndPos = InStr(1, strNames, ")")
    strProcName = Mid(strNames, 1, iStartPos - 1)
    strNames = Mid(strNames, iStartPos + 1, iEndPos - iStartPos - 1)
    
    arrNames = Split(strNames, ",")
    For i = 0 To UBound(arrValues)
      strValue = VarToString(arrValues(i))
      strNameValue = arrNames(i) & "=" & strValue
      If strList <> "" Then strList = strList & ", "
      strList = strList & strNameValue
    Next i
    CreateNameValueList = strProcName & "(" & strList & ")"
    Exit Function
ErrHandler:
    'Nothing to do: failed for whatever reason
End Function

Private Function GetAppVersion() As String
    GetAppVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Private Function VarToString(ByVal V) As String

  On Error GoTo ErrHandler
  If IsArray(V) Then
        VarToString = "{Array}"
    Else 'If IsArray(...
    Select Case VarType(V)
        Case vbInteger, vbLong, vbByte, _
             vbSingle, vbDouble, vbCurrency, _
             vbBoolean, vbDecimal:
                      VarToString = CStr(V)
        Case vbDate:      VarToString = "'" & CStr(V) & "'"
        Case vbError:     VarToString = "" 'Missing arg falls here
        Case vbEmpty:     VarToString = "{Empty}"
        Case vbNull:      VarToString = "{Null}"
        Case vbString:    VarToString = "'" & V & "'"
        Case vbObject:    VarToString = "{" & TypeName(V) & "}"
        Case Else:        VarToString = "{?}"
        End Select
    End If 'IsArray...
  Exit Function
ErrHandler:
  VarToString = "{?}"
  End Function

Public Sub HandleError(Optional ByVal EnvVarNames As String, _
                        Optional ByVal arrEnvVars)

  Dim oMSB As frmMessageBox
  Dim iButtonPress As Integer
  Dim strMessage As String
  
  If mbEnabled = True Then
    
    Set oMSB = New frmMessageBox
    strMessage = Me.Description
    
    ' Add any specified value to specifically monitor
    If IsArray(arrEnvVars) Then
      strMessage = strMessage & vbCrLf & "Current Procedure Variables" & vbCrLf & String(27, "-") & vbCrLf & CreateNameValueList(EnvVarNames, arrEnvVars)
    End If
    
    ' Add the error to the event log
    gobjEventLog.AddHeader eltSystemError, Err.Description
    gobjEventLog.ChangeHeaderStatus elsError
    gobjEventLog.AddDetailEntry strMessage
    
    ' Display the error
    iButtonPress = oMSB.ErrorBox(strMessage)
    
    'Quit application
    If iButtonPress = vbAbort Then
      End
    End If
    
    'Ignore error (you should be in break on all errors anyway???)
    If iButtonPress = vbIgnore Then
      Me.PopStack
      'Err.Raise 5, "", ""     'You can now edit code by looking in the call stack (not brilliant, but all I can do for the moment)
    End If
  Else
    Me.PopStack
  End If

End Sub

Public Sub Enable()

  mbEnabled = True

End Sub

Public Sub Disable()

  mbEnabled = False

End Sub
