VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Instantiate internal classes
Private ODBC As New ODBC

'Local variables to hold property values
'Private mvarDatabaseName As String
'Private mvarKeyWords As String
'Private mvarMaxTableNameLength As Integer
'Private mvarMaxColumns As Integer
'Private mvarMaxColumnNameLength As Integer
'Private mvarName As String
'Private mvarVersion As String
'Private mvarUserName As String
'Private mvarServerName As String

Private mbValidation As Boolean

'Private datGeneral As ASRUserData.clsGeneral
Private datGeneral As HRProDataMgr.clsGeneral

Private Sub Class_Initialize()

'    Set datGeneral = New ASRUserData.clsGeneral
  Set datGeneral = New HRProDataMgr.clsGeneral

End Sub

Private Sub Class_Terminate()
  'Release internal classes
  Set ODBC = Nothing
End Sub

'MH20020122 Not Used ?
'Public Function GetColumnDefault(lTableID As Long, sColumnName As String) As Variant
'
'    Dim rsTemp As New ADODB.Recordset
'    Dim vResult As Variant
'
'    Set rsTemp = datGeneral.GetColumnDefault(sColumnName, lTableID)
'    If Not rsTemp.EOF Then
'        If Not IsNull(rsTemp!COLUMN_DEF) Then
'            Select Case rsTemp!TYPE_NAME
'                Case "varchar", "char"
'                    vResult = Mid$(rsTemp!COLUMN_DEF, 3, Len(rsTemp!COLUMN_DEF) - 4)
'                Case "bit"
'                    vResult = Mid$(rsTemp!COLUMN_DEF, 2, Len(rsTemp!COLUMN_DEF) - 2)
'            End Select
'
'            GetColumnDefault = vResult
'        End If
'    Else
'        GetColumnDefault = vbNullString
'    End If
'    rsTemp.Close
'    Set rsTemp = Nothing
'
'End Function

Public Property Get Validation() As Boolean

    Validation = mbValidation

End Property

Public Property Let Validation(ByVal bValidate As Boolean)

    mbValidation = bValidate

End Property
