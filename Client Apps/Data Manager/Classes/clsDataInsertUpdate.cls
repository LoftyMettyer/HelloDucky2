VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataInsertUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strArray() As String
Private mlngIndex As Long

Private Sub Class_Initialize()
  mlngIndex = -1
  ReDim strArray(1, 0)
End Sub

Public Sub AddColumn(strName As String, strValue As String, Optional blnText As Boolean)

  mlngIndex = mlngIndex + 1
  ReDim Preserve strArray(1, mlngIndex)

  strArray(0, mlngIndex) = "[" & strName & "]"
  If blnText Then
    strArray(1, mlngIndex) = "'" & Replace(strValue, "'", "''") & "'"
  Else
    strArray(1, mlngIndex) = strValue
  End If

End Sub

Public Function InsertUpdate(strTableName As String, strIDColumn As String, lngRecordID As Long) As Long

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsDataInsertUpdate.InsertUpdate", Array(strTableName, strIDColumn, lngRecordID)
  
  
  Dim rsTemp As Recordset
  Dim strSQL As String
  
  Dim strColumns As String
  Dim strValues As String
  Dim lngCount As Long

  If lngRecordID > 0 Then
    'Update
    For lngCount = 0 To mlngIndex
      strSQL = IIf(lngCount > 0, strSQL & ", ", vbNullString) & _
        strArray(0, lngCount) & " = " & strArray(1, lngCount)
    Next

    strSQL = "UPDATE " & strTableName & _
             " SET " & strSQL & _
             " WHERE [" & strIDColumn & "] = " & CStr(lngRecordID)

    gADOCon.Execute strSQL

  Else
    'Insert
    For lngCount = 0 To mlngIndex
      strColumns = IIf(lngCount > 0, strColumns & ", ", vbNullString) & _
          strArray(0, lngCount)
      strValues = IIf(lngCount > 0, strValues & ", ", vbNullString) & _
          strArray(1, lngCount)
    Next

    strSQL = "INSERT " & strTableName & _
             " (" & strColumns & ") " & _
             " VALUES" & _
             " (" & strValues & ")"

    gADOCon.BeginTrans
    gADOCon.Execute strSQL
    
    strSQL = "SELECT MAX([" & strIDColumn & "]) FROM " & strTableName
    Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
    
    lngRecordID = rsTemp.Fields(0).Value
    gADOCon.CommitTrans
    
    rsTemp.Close
    Set rsTemp = Nothing
  
  End If

  InsertUpdate = lngRecordID


TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function
