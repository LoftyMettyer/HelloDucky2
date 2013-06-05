Attribute VB_Name = "modQAProErrorHandling"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
Public Const conHwndTopmost = -1
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40

Public mProcStack As New clsQAProProcedureStack

'*****************************************************
' Purpose:  Processes the Error object and displays the
'           error details to the user.
' Inputs:
'   strModuleName:  the name of the module in which the
'                   error has occurred.
'   mErr:           the Error object
'   strProcName:    the name of the procedure in which
'                   the error has occurred.
'*****************************************************

Public Function Process_Error(strModuleName As String, ByRef mErr As ErrObject, strProcName As String)

    Dim objProc As clsQAProProcedure
    'Dim nFRM As New frmQAProError
    Dim txtError As String
    
    Dim strMessage As String * 100
    Dim lngErrorMessageReturn
    Dim piErrResponse As Integer

    lngErrorMessageReturn = QA_ErrorMessage(mErr.Number, strMessage, 100)
    
    If InStr(strMessage, "Error ") = 1 Then
    
        strMessage = mErr.Description
    
    End If
    
    txtError = "Procedure Name:" & vbTab & strProcName & vbCrLf _
                    & "Module Name:" & vbTab & strModuleName & vbCrLf & vbCrLf _
                    & "Error Number:" & vbTab & mErr.Number & vbCrLf _
                    & "Error Description:" & vbTab & UnMakeCString(strMessage) & vbCrLf & vbCrLf _
                    & "Call Stack:" & vbCrLf
    
    '& "Error Source:" & vbTab & mErr.Source & vbCrLf _
    '& "Last DLL Error:" & vbTab & mErr.LastDllError & vbCrLf & vbCrLf _

    Set objProc = mProcStack.Top
    
    Do Until objProc Is Nothing
        txtError = txtError & vbTab & objProc.Name & " in " & objProc.Module & vbCrLf
        Set objProc = objProc.NextProc
    Loop
    
    piErrResponse = MsgBox(txtError, vbCritical + vbOKOnly, "Quick Address Error")

End Function
