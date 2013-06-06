Attribute VB_Name = "modPostSpecifics"
Option Explicit

Public Const gsMODULEKEY_POST = "MODULE_POST"

Public Const gsPARAMETERKEY_POSTTABLE = "Param_PostTable"
Public Const gsPARAMETERKEY_POSTJOBTITLECOLUMN = "Param_PostJobTitleColumn"
Public Const gsPARAMETERKEY_POSTGRADECOLUMN = "Param_PostGradeColumn"
Public Const gsPARAMETERKEY_GRADETABLE = "Param_GradeTable"
Public Const gsPARAMETERKEY_GRADECOLUMN = "Param_GradeColumn"
Public Const gsPARAMETERKEY_NUMLEVELCOLUMN = "Param_NumLevelColumn"

Public Const gsPARAMETERKEY_SUCCESSIONDEF = "Param_SuccessionDef"
Public Const gsPARAMETERKEY_SUCCESSIONALLOWEQUAL = "Param_SuccessionAllowEqual"
Public Const gsPARAMETERKEY_SUCCESSIONRESTRICT = "Param_SuccessionRestrict"
Public Const gsPARAMETERKEY_SUCCESSIONLEVELS = "Param_SuccessionLevels"

Public Const gsPARAMETERKEY_CAREERDEF = "Param_CareerDef"
Public Const gsPARAMETERKEY_CAREERALLOWEQUAL = "Param_CareerAllowEqual"
Public Const gsPARAMETERKEY_CAREERRESTRICT = "Param_CareerRestrict"
Public Const gsPARAMETERKEY_CAREERLEVELS = "Param_CareerLevels"

Public glngPostTableID As Long
Public gstrPostTableName As String
Public glngJobTitleColumnID As Long
Public gstrJobTitleColumnName As String
Public glngPostGradeColumnID As Long
Public gstrPostGradeColumnName As String
Public glngGradeTableID As Long
Public gstrGradeTableName As String
Public glngGradeColumnID As Long
Public gstrGradeColumnName As String
Public glngNumLevelColumnID As Long
Public gstrNumLevelColumnName As String

Public glngSuccessionDef As Long
Public gblnSuccessionAllowEqual As Boolean
Public gblnSuccessionRestrict As Boolean
Public gblnSuccessionLevels As Boolean

Public glngCareerDef As Long
Public gblnCareerAllowEqual As Boolean
Public gblnCareerRestrict As Boolean
Public gblnCareerLevels As Boolean


Public Sub ReadPostParameters()

  glngPostTableID = Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_POSTTABLE))
  If glngPostTableID > 0 Then
    gstrPostTableName = datGeneral.GetTableName(glngPostTableID)
  Else
    gstrPostTableName = ""
  End If

  glngJobTitleColumnID = Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_POSTJOBTITLECOLUMN))
  If glngJobTitleColumnID > 0 Then
    gstrJobTitleColumnName = datGeneral.GetColumnName(glngJobTitleColumnID)
  Else
    gstrJobTitleColumnName = ""
  End If

  glngPostGradeColumnID = Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_POSTGRADECOLUMN))
  If glngPostGradeColumnID > 0 Then
    gstrPostGradeColumnName = datGeneral.GetColumnName(glngPostGradeColumnID)
  Else
    gstrPostGradeColumnName = ""
  End If

  glngGradeTableID = Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_GRADETABLE))
  If glngGradeTableID > 0 Then
    gstrGradeTableName = datGeneral.GetTableName(glngGradeTableID)
  Else
    gstrGradeTableName = ""
  End If

  glngGradeColumnID = Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_GRADECOLUMN))
  If glngGradeColumnID > 0 Then
    gstrGradeColumnName = datGeneral.GetColumnName(glngGradeColumnID)
  Else
    gstrGradeColumnName = ""
  End If

  glngNumLevelColumnID = Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_NUMLEVELCOLUMN))
  If glngNumLevelColumnID > 0 Then
    gstrNumLevelColumnName = datGeneral.GetColumnName(glngNumLevelColumnID)
  Else
    gstrNumLevelColumnName = ""
  End If


  glngSuccessionDef = Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_SUCCESSIONDEF))
  gblnSuccessionAllowEqual = (Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_SUCCESSIONALLOWEQUAL)) = -1)
  gblnSuccessionRestrict = (Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_SUCCESSIONRESTRICT)) = -1)
  gblnSuccessionLevels = (Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_SUCCESSIONLEVELS)) = -1)

  glngCareerDef = Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_CAREERDEF))
  gblnCareerAllowEqual = (Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_CAREERALLOWEQUAL)) = -1)
  gblnCareerRestrict = (Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_CAREERRESTRICT)) = -1)
  gblnCareerLevels = (Val(GetModuleParameter(gsMODULEKEY_POST, gsPARAMETERKEY_CAREERLEVELS)) = -1)

End Sub


Public Function ValidatePostParameters() As Boolean

  Dim strError As String

  strError = vbNullString
  
  If glngPersonnelTableID = 0 Then
    strError = strError & vbCrLf & _
      "Personnel table"
  End If

  If glngJobTitleColumnID = 0 Then
    strError = strError & vbCrLf & _
      "Job Title column"
  End If

  If glngPostGradeColumnID = 0 Then
    strError = strError & vbCrLf & _
      "Post Grade column"
  End If

  If glngGradeTableID = 0 Then
    strError = strError & vbCrLf & _
      "Grade table"
  End If

  If glngGradeColumnID = 0 Then
    strError = strError & vbCrLf & _
      "Grade column"
  End If

  If glngNumLevelColumnID = 0 Then
    strError = strError & vbCrLf & _
      "Hierarchy column"
  End If

  
  If strError <> vbNullString Then
    MsgBox "The following must be correctly configured in the Post module before proceeding:" & vbCrLf & _
      strError, vbExclamation + vbOKOnly, "Post Module"
  End If
  
  ValidatePostParameters = (strError = vbNullString)

End Function
