VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinkContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolContentDetails As Collection
Private mlngContentID As Long

Private Sub Class_Initialize()
  Clear
End Sub

Public Sub Clear()
  Set mcolContentDetails = New Collection
End Sub


Public Function GetID() As Long
  GetID = mlngContentID
End Function


Public Function CloneDetail(pavCloneRegister As Variant, lngContentID As Long)
  
  Dim strFieldText As String
  Dim strFieldCode As String
  Dim lngFieldID As Long
  
  Clear

  With recLinkContentEdit
    .Index = "idxContentIDSequence"
    .Seek ">=", lngContentID, 0
  
    If Not .NoMatch Then
      Do While Not .EOF
        
        If !ContentID <> lngContentID Then
          Exit Do
        End If

        strFieldText = !FixedText
        strFieldCode = !FieldCode
        lngFieldID = !FieldID
        
        Select Case strFieldCode
        Case "C"
          lngFieldID = GetClone(pavCloneRegister, "COLUMN", lngFieldID)
        Case "E"
          lngFieldID = GetClone(pavCloneRegister, "EXPRESSION", lngFieldID)
        End Select
        
        AddDetail IIf(IsNull(strFieldText), "", strFieldText), _
                  IIf(IsNull(strFieldCode), "", strFieldCode), _
                  IIf(IsNull(lngFieldID), "", lngFieldID)

        .MoveNext
      Loop
    End If
  
  End With

End Function


Public Sub ReadDetail(lngContentID As Long)

  mlngContentID = lngContentID
  
  Clear

  With recLinkContentEdit
    .Index = "idxContentIDSequence"
    .Seek ">=", mlngContentID, 0
  
    If Not .NoMatch Then
      Do While Not .EOF
        
        If !ContentID <> mlngContentID Then
          Exit Do
        End If

        AddDetail IIf(IsNull(!FixedText), "", !FixedText), _
                  IIf(IsNull(!FieldCode), "", !FieldCode), _
                  IIf(IsNull(!FieldID), "", !FieldID)

        .MoveNext
      Loop
    End If
  
  End With

End Sub


Public Function AddDetailFromField(colAvailableComponents As Collection, strFixedText As String, strField As String, lngTableID As Long) As Boolean

  Dim strExprName As String

  On Local Error GoTo LocalErr

  strField = colAvailableComponents(strField)
  AddDetail strFixedText, Left(strField, 1), Val(Mid(strField, 2))

  AddDetailFromField = True

Exit Function

LocalErr:

  AddDetailFromField = False
  
  'MH20090804 Fault HRPRO-193
  'This is bad.
  
  'In column calcs, the duplicate name check is case sensitive
  '(so you can have one called 'AAA' and another called 'aaa'.
  
  'A VB Collection key is case INsensitive. This means the original method of
  'adding each component name to a collection (to look up the ID from the name)
  'failed with a duplicate key fault.
  
  'A recordset index is case INSensitive so when I then tried going back to the
  'recordset and getting the name, this also failed.
  
  'My only option was to do a loop to try and find the correct case version.
  
  'Rubbish.
  
  
  If Left(strField, 13) = "Calculation: " Then
    strExprName = Mid(strField, 14)
    With recExprEdit
      .Index = "idxExprName"
      .Seek "=", strExprName
      If Not .NoMatch Then
        Do
          'This is a case sensitive comparison
          'If strExprName = !Name Then
          '  If lngTableID = !TableID Then
          If strExprName = !Name And lngTableID = !TableID Then
            Select Case !Type
            Case giEXPR_COLUMNCALCULATION, giEXPR_RECORDDESCRIPTION, giEXPR_EMAIL
              AddDetail strFixedText, "E", !ExprID
              AddDetailFromField = True
              Exit Do
            End Select
          End If
          .MoveNext
        Loop While LCase(strExprName) = LCase(!Name)
      End If
    End With
  End If

End Function


Public Sub AddDetail(strFixedText As String, strFieldCode As String, lngFieldID As Long)

  Dim objDetail As clsLinkContentDetail
  
  Set objDetail = New clsLinkContentDetail
  
  If mcolContentDetails Is Nothing Then
    Set mcolContentDetails = New Collection
  End If
  
  objDetail.FixedText = strFixedText
  objDetail.FieldCode = strFieldCode
  objDetail.FieldID = lngFieldID
  objDetail.Sequence = mcolContentDetails.Count
  
  mcolContentDetails.Add objDetail
  
  
  Set objDetail = Nothing

End Sub


Public Function SetTextboxFromContent(txt As TextBox) As Boolean

  Dim objDetail As clsLinkContentDetail
  Dim strOutput As String
  
  strOutput = vbNullString
  For Each objDetail In mcolContentDetails
    strOutput = strOutput & objDetail.FixedText & objDetail.GetFieldForDisplay
  Next
  txt.Text = strOutput
  
  SetTextboxFromContent = True
  
Exit Function

LocalErr:
  MsgBox "Error Reading Merge Content (" & Err.Description & ")", vbCritical
  SetTextboxFromContent = False

End Function


Public Function SetContentFromTextbox(txt As TextBox, colAvailableComponents As Collection, lngTableID As Long) As String

  Dim strArray() As String
  Dim strArray2() As String
  Dim intIndex As Integer
  Dim strText As String
  Dim strMergeField As String
  Dim strHighlightText As String
  Dim blnLastField As Boolean
  
  SetContentFromTextbox = vbNullString
  
  Me.Clear
  strArray = Split(txt.Text, strDelimStop)
  For intIndex = 0 To UBound(strArray)
  
    blnLastField = (intIndex = UBound(strArray))
  
    strArray2 = Split(strArray(intIndex), strDelimStart)
    Select Case UBound(strArray2)
    Case 2
      SetContentFromTextbox = "Merge End Point Not Found."
      strHighlightText = strArray2(1)

    Case 1
      If Not AddDetailFromField(colAvailableComponents, strArray2(0), strArray2(1), lngTableID) Then
        SetContentFromTextbox = IIf(blnLastField, "Merge End Point Not Found.", "Invalid Merge Point.")
        strHighlightText = strArray2(1)
      End If

    Case 0
      If blnLastField Then
        AddDetail strArray2(0), "", 0
      Else
        SetContentFromTextbox = "Merge Start Point Not Found."
        strHighlightText = strArray(intIndex)
      End If

    End Select

  Next
  
  If strHighlightText <> vbNullString Then
    txt.SelStart = InStr(txt.Text, strHighlightText) - 1
    txt.SelLength = Len(strHighlightText)
  End If


Exit Function

LocalErr:
  SetContentFromTextbox = "Error Validating Merge Content (" & Err.Description & ")"

End Function



Public Function GetSQL(lngBaseTableID As Long, strBaseTableName As String, lngRecDescID As Long, strOutputVariable As String) As String

  Dim objExpr As CExpression
  Dim objDetail As clsLinkContentDetail
  Dim strOutput As String
  Dim col As New Collection
  'Dim strBaseTableName As String
  Dim strTableName As String
  Dim strType As String
  
  Dim strSQLJoin As String
  Dim strSQLDeclares As String
  Dim strSQLExpressions As String
  Dim strTemp As String
  
  Dim blnColumn As Boolean
  Dim blnDeclareChar As Boolean
  Dim blnDeclareFloat As Boolean
  Dim blnDeclareDate As Boolean
  Dim blnDeclareBit As Boolean
  
  Dim strTempVariable As String
  Dim strTempConvert As String


  strSQLJoin = vbNullString
  strSQLDeclares = vbNullString
  strSQLExpressions = vbNullString
  
  blnColumn = False
  blnDeclareChar = False
  blnDeclareFloat = False
  blnDeclareDate = False
  blnDeclareBit = False
  
  
  For Each objDetail In mcolContentDetails
    'strOutput = strOutput & _
        IIf(strOutput <> vbNullString, " + ", "") & _
        "'" & Replace(Replace(objDetail.FixedText, "'", "''"), vbCrLf, "' + char(13)+char(10) + " & vbCrLf & "      '") & "'"
    strOutput = strOutput & _
        IIf(strOutput <> vbNullString, " + ", "") & _
        "'" & Replace(objDetail.FixedText, "'", "''") & "'"

    strTemp = objDetail.GetFieldForSQL(lngBaseTableID, strBaseTableName, strOutputVariable)
    If strTemp <> vbNullString Then
      strOutput = strOutput & _
          " + " & strTemp
    End If

    If objDetail.FieldCode = "C" Then
      
      blnColumn = True
      With recColEdit
        .Index = "idxColumnID"
        .Seek "=", objDetail.FieldID

        If Not .NoMatch Then
          If !TableID <> lngBaseTableID Then
            strTableName = GetTableName(!TableID)
            If Not Exists(col, CStr(!TableID)) Then
              col.Add !TableID, CStr(!TableID)
              strSQLJoin = "     JOIN [" & strTableName & "]" & _
                        " ON [" & strTableName & "].ID" & _
                        " = [" & strBaseTableName & "].ID_" & CStr(!TableID) & vbCrLf
            End If
          End If
        End If
      
      End With
    
    ElseIf objDetail.FieldCode = "E" Then
      
      If objDetail.FieldID = lngRecDescID Then
        strSQLExpressions = strSQLExpressions & _
            "    SET " & strOutputVariable & CStr(objDetail.FieldID) & " = @recordDesc" & vbCrLf
      Else
      
        Set objExpr = New CExpression
        With objExpr
          .ExpressionID = objDetail.FieldID
          If .ConstructExpression Then

            strSQLExpressions = strSQLExpressions & _
              "    SET @id = @recordID" & vbNewLine


            Select Case objExpr.ReturnType
            Case 1  'char
              strTempVariable = "@tempChar"
              strTempConvert = "isnull(@tempChar,'')"
              blnDeclareChar = True
            Case 2  'numeric
              strTempVariable = "@tempFloat"
              strTempConvert = "isnull(convert(varchar(max),@tempFloat),'')"
              blnDeclareFloat = True
            Case 3  'bit
              strTempVariable = "@tempBit"
              strTempConvert = "case isnull(@tempBit,0) when 0 then 'No' else 'Yes' end"
              blnDeclareBit = True
            Case 4  'date
              strTempVariable = "@tempDate"
              strTempConvert = "isnull(convert(varchar(max),@tempDate,103),'')"
              blnDeclareDate = True
            End Select
  
  
            If strTableName = "deleted" Then
              strSQLExpressions = strSQLExpressions & _
                  .StoredProcedureCode(strTempVariable, strBaseTableName) & vbCrLf
            Else
              strSQLExpressions = strSQLExpressions & _
                  "    EXEC @hResult = dbo.sp_ASRExpr_" & CStr(objDetail.FieldID) & " " & strTempVariable & " OUTPUT, @recordID" & vbCrLf
            End If
            strSQLExpressions = strSQLExpressions & _
                "    SET " & strOutputVariable & CStr(objDetail.FieldID) & " = " & strTempConvert & vbCrLf
          
          End If
    
        End With
      
      End If

      strSQLDeclares = strSQLDeclares & _
          "    DECLARE " & strOutputVariable & CStr(objDetail.FieldID) & " varchar(max)" & vbCrLf

    End If
  
  Next

  If strOutput <> vbNullString Then
    strOutput = _
        IIf(blnDeclareChar, "    DECLARE @tempChar varchar(max)" & vbCrLf, "") & _
        IIf(blnDeclareFloat, "    DECLARE @tempFloat float" & vbCrLf, "") & _
        IIf(blnDeclareDate, "    DECLARE @tempDate datetime" & vbCrLf, "") & _
        IIf(blnDeclareBit, "    DECLARE @tempBit bit" & vbCrLf, "") & vbCrLf & _
        IIf(strSQLDeclares <> vbNullString, strSQLDeclares & vbCrLf, "") & _
        IIf(strSQLExpressions <> vbNullString, strSQLExpressions & vbCrLf, "") & _
        "    SET @id = @recordID" & vbCrLf & _
        "    SELECT " & strOutputVariable & " = " & strOutput

        
    If blnColumn Then
      strOutput = strOutput & vbCrLf & _
        "FROM [" & strBaseTableName & "]" & vbCrLf & _
        strSQLJoin & _
        "WHERE [" & strBaseTableName & "].ID = @RecordID"
    End If
  End If

  GetSQL = strOutput

End Function


Public Function WriteDetail() As Long

  Dim objDetail As clsLinkContentDetail
  Dim strSQL As String


  If mlngContentID = 0 Then
    mlngContentID = Database.UniqueColumnValue("tmpLinkContent", "ContentID")
  Else
    strSQL = "DELETE FROM tmpLinkContent WHERE ContentID = " & CStr(mlngContentID)
    daoDb.Execute strSQL, dbFailOnError
  End If
    
  For Each objDetail In mcolContentDetails
    objDetail.WriteDetail (mlngContentID)
  Next

  WriteDetail = mlngContentID

End Function


Private Function Exists(col As Collection, id As String) As Boolean
  On Local Error GoTo LocalErr
  Exists = (Val(col(id)) > 0)
Exit Function
LocalErr:
  Exists = False
End Function

