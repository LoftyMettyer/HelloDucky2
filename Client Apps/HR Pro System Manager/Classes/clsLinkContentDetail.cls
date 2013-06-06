VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinkContentDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngSequence As Long
Private mstrFixedText As String
Private mstrFieldCode As String
Private mlngFieldID As Long


Public Property Get Sequence() As Long
  Sequence = mlngSequence
End Property

Public Property Let Sequence(value As Long)
  mlngSequence = value
End Property


Public Property Get FixedText() As String
  FixedText = mstrFixedText
End Property

Public Property Let FixedText(value As String)
  mstrFixedText = value
End Property


Public Property Get FieldCode() As String
  FieldCode = mstrFieldCode
End Property

Public Property Let FieldCode(value As String)
  mstrFieldCode = value
End Property


Public Property Get FieldID() As Long
  FieldID = mlngFieldID
End Property

Public Property Let FieldID(value As Long)
  mlngFieldID = value
End Property


Public Function GetTableID() As Long

  Dim lngTableID As String
  
  lngTableID = 0
  
  Select Case mstrFieldCode
  Case "C"
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", mlngFieldID
      If Not .NoMatch Then
        lngTableID = !TableID
      End If
    End With
  
  Case "E"
    With recExprEdit
      .Index = "idxExprID"
      .Seek "=", mlngFieldID, False
      If Not .NoMatch Then
        lngTableID = !TableID
      End If
    End With
  
  End Select

  GetTableID = lngTableID

End Function


Public Function GetFieldForDisplay() As String
  
  Dim strField As String
  
  Select Case mstrFieldCode
  Case "C"
    strField = GetColumnName(mlngFieldID)
  Case "E"
    strField = "Calculation: " & GetExpressionName(mlngFieldID)
  Case "X"
    strField = "Function: Current User"
    'Select Case mintFieldID
    '  Case 0: strField = "Current User"
    '  Case 1: strField = "Old Column Value"
    '  End Select
    'End Select
  End Select

  If strField <> vbNullString Then
    GetFieldForDisplay = strDelimStart & strField & strDelimStop
  Else
    GetFieldForDisplay = vbNullString
  End If

End Function


Public Function GetFieldForSQL(lngBaseTableID As Long, strBaseTableName As String, strExprPrefix As String) As String
  
  Dim lngTableID As Long
  Dim strField As String

  
  Select Case mstrFieldCode
  Case "C"
  
    lngTableID = GetTableID
    If lngTableID = lngBaseTableID Then
      strField = strBaseTableName & "." & GetColumnName(mlngFieldID, True)
    Else
      strField = GetColumnName(mlngFieldID, False)
    End If
  
    Select Case recColEdit!DataType
    Case sqlBoolean
      strField = "case isnull(" & strField & ",0) when 0 then 'No' else 'Yes' end"
    Case sqlNumeric, sqlInteger
      'strField = "cast(cast(" & strField & " as decimal(" & CStr(!Size) & "," & CStr(!Decimal) & ")) as varchar(max))"
      'strField = "isnull(cast(" & strField & " as varchar(max)),'')"
      strField = "isnull(dbo.udfASRConvertNumeric(" & strField & "," & CStr(recColEdit!Decimals) & "," & IIf(recColEdit!Use1000Separator, "1", "0") & "),'')"
    Case sqlDate
      strField = "isnull(convert(varchar(max)," & strField & "," & CStr(glngEmailDateFormat) & "),'')"
    Case sqlVarChar
      strField = "isnull(" & strField & ",'')"
    End Select
  
  Case "E"
    strField = strExprPrefix & CStr(mlngFieldID)
  
  Case "X"
    'strField = "CASE WHEN UPPER(LEFT(APP_NAME(), " & Len(gsWORKFLOWAPPLICATIONPREFIX) & ")) = '" & UCase(gsWORKFLOWAPPLICATIONPREFIX) & "' THEN '" & gsWORKFLOWAPPLICATIONPREFIX & "' ELSE @UserName END"
    strField = "@UserName"
    'Select Case mintFieldID
    '  Case 0: strField = "Current User"
    '  Case 1: strField = "Old Column Value"
    '  End Select
    'End Select
  End Select

  If strField <> vbNullString Then
    GetFieldForSQL = strField
  Else
    GetFieldForSQL = vbNullString
  End If

End Function


Public Function WriteDetail(lngContentID As Long) As Boolean

  Dim lngID As Long

  With recLinkContentEdit

    lngID = Database.UniqueColumnValue("tmpLinkContent", "ID")

    .AddNew
    !id = lngID
    !ContentID = lngContentID
    !Sequence = mlngSequence
    !FixedText = mstrFixedText
    !FieldCode = mstrFieldCode
    !FieldID = mlngFieldID
    .Update
  End With
  
End Function

