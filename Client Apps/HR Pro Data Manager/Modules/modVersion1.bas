Attribute VB_Name = "modVersion1"
Option Explicit

' Module Setup Constants
Public Const MODULEKEY_DOCMANAGEMENT = "MODULE_DOCUMENTMANAGEMENT"
Public Const PARAMETERKEY_DOCMAN_CATEGORYTABLE = "Param_DocmanCatageoryTable"
Public Const PARAMETERKEY_DOCMAN_CATEGORYCOLUMN = "Param_DocManCatageoryColumn"
Public Const PARAMETERKEY_DOCMAN_TYPETABLE = "Param_DocmanTypeTable"
Public Const PARAMETERKEY_DOCMAN_TYPECOLUMN = "Param_DocManTypeColumn"
Public Const PARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN = "Param_DocManTypeCategoryColumn"

Public Function IsV1ModuleSetupValid(ByVal WarnUser As Boolean) As Boolean
  
  Dim lngParameter As Long
  Dim sMessage As String
  Dim bOK As Boolean
  
  bOK = True
  
  If Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_CATEGORYTABLE)) = 0 Or _
    Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_CATEGORYCOLUMN)) = 0 Or _
    Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPETABLE)) = 0 Or _
    Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPECOLUMN)) = 0 Or _
    Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN)) = 0 Then
      bOK = False
  End If
     
  If WarnUser And Not bOK Then
    COAMsgBox "Document Management module setup is not completed.", vbExclamation, Application.Name
  End If
  
  IsV1ModuleSetupValid = bOK
  
  
End Function

Public Function GenerateV1MailMergeHeader(ByVal bManual As Boolean, ByVal sCategory As String, ByVal sType As String, ByVal lngKeyField As Long, _
                                      ByVal lngParentKeyfield1 As Long, ByVal lngParentKeyfield2 As Long, _
                                      ByVal sManualHeaderText As String) As String

  On Error GoTo ErrorTrap

  Dim objHeader As HRProDataMgr.clsStringBuilder
  Dim bOK As Boolean
  Dim sKeyField As String
  Dim sParentKeyfield1 As String
  Dim sParentKeyfield2 As String

  bOK = True
  sKeyField = datGeneral.GetColumnName(lngKeyField, True)
  sParentKeyfield1 = datGeneral.GetColumnName(lngParentKeyfield1, True)
  sParentKeyfield2 = datGeneral.GetColumnName(lngParentKeyfield2, True)
  
  Set objHeader = New HRProDataMgr.clsStringBuilder
  objHeader.TheString = vbNullString
  
  If Not bManual Then
  
    objHeader.Append "~~!:~~@[$TABLE:COMPLETE_FILES]~~@[DOC_SECTION:" & sCategory & "]~~@[DOC_TYPE:" & sType & "]"
    
    If Len(sParentKeyfield1) > 0 Then
      objHeader.Append "~~@[STAFF_NO:{MERGEFIELD""" & sParentKeyfield1 & """}]" & _
        "~~@[DOCUMENT_KEY:{MERGEFIELD""" & sKeyField & """}]"
    Else
      objHeader.Append "~~@[STAFF_NO:{MERGEFIELD""" & sKeyField & """}]"
    End If

    objHeader.Append "~~@[DOC_DATE:{DATE \@""dd/MM/yyyy""}]~~"
    
  Else
    objHeader.TheString = sManualHeaderText
  End If

TidyUpAndExit:
  GenerateV1MailMergeHeader = objHeader.ToString
  Set objHeader = Nothing
  Exit Function

ErrorTrap:
  objHeader.TheString = vbNullString
  bOK = False

End Function


