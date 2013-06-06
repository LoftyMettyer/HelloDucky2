VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystemPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variables to hold property values.
Private mvar_lngItemID As Long
Private mvar_fAllowed As Boolean
Private mvar_sItemKey As String
Private mvar_sCategoryKey As String

Private mlngPageNum As Long
Private mlngBottom As Long

Private asCategoryDescriptions() As String
Private asItemDescriptions() As String




Public Property Let Allowed(ByVal pfData As Boolean)
  mvar_fAllowed = pfData

End Property


Public Function PrintSecurity(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
'
  Dim fOK As Boolean
  Dim sPermission As String
  Dim intSameLine As Integer

  fOK = True
  
  'NHRD28082003 Fault 6391
  If mvar_sCategoryKey = "CMG" Then
    If IsModuleEnabled(modCMG) = False Then GoTo TidyUpAndExit
  End If
  
  If mvar_sCategoryKey = "ACCORD" Then
    If IsModuleEnabled(modACCORD) = False Then GoTo TidyUpAndExit
  End If
  
  If mvar_sCategoryKey = "WORKFLOW" Then
    If IsModuleEnabled(modWorkflow) = False Then GoTo TidyUpAndExit
  End If
  
  sPermission = IIf(mvar_fAllowed, "TRUE", "FALSE")
  Dim intOldFontSize As Integer

    PopulateDescriptionArray

    With Printer
        intOldFontSize = .FontSize
        .FontSize = 6
        
        intSameLine = .CurrentY + (giPRINT_YSPACE * 1)

'MH20050812 Fault 10267
'        'This will force a printing of a blank-checkbox
'        If gasPrintOptions(1).PrintBlankVersion = True Then mvar_fAllowed = False
'
'        If mvar_fAllowed Then
'           Printer.PaintPicture frmGroupMaint1.PicTickedCheckBox.Image, giPRINT_XINDENT + (giPRINT_XCOLUMN * (piLevel - 1)), (intSameLine - giPRINT_YSPACE)
'        Else
'          Printer.PaintPicture frmGroupMaint1.PicBlankCheckBox.Image, giPRINT_XINDENT + (giPRINT_XCOLUMN * (piLevel - 1)), (intSameLine - giPRINT_YSPACE)
'        End If
        If gasPrintOptions(1).PrintBlankVersion = True Or Not mvar_fAllowed Then
          Printer.PaintPicture frmGroupMaint1.PicBlankCheckBox.Image, giPRINT_XINDENT + (giPRINT_XCOLUMN * (piLevel - 1)), (intSameLine - giPRINT_YSPACE)
        Else
          Printer.PaintPicture frmGroupMaint1.PicTickedCheckBox.Image, giPRINT_XINDENT + (giPRINT_XCOLUMN * (piLevel - 1)), (intSameLine - giPRINT_YSPACE)
        End If
       
       
        .CurrentX = giPRINT_XINDENT + (giPRINT_XCOLUMN * (piLevel - 1)) + 280
          
        .CurrentY = intSameLine
        Printer.Print GetItemDescription(mvar_sItemKey)
        
        'Restore font size
        .FontSize = intOldFontSize
    End With

TidyUpAndExit:
  PrintSecurity = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Property Let ItemID(ByVal plngNewValue As Long)
  mvar_lngItemID = plngNewValue
  
End Property



Public Property Let ItemKey(ByVal psNewValue As String)
  mvar_sItemKey = psNewValue
  
End Property
Public Property Let CategoryKey(ByVal psNewValue As String)
  mvar_sCategoryKey = psNewValue
  
End Property


Public Property Get ItemID() As Long
  
  ItemID = mvar_lngItemID
  
End Property



Public Property Get ItemKey() As String
  
  ItemKey = mvar_sItemKey
  
End Property

Public Property Get CategoryKey() As String
  
  CategoryKey = mvar_sCategoryKey
  
End Property


Public Property Get Allowed() As Boolean
  Allowed = mvar_fAllowed

End Property

Public Sub PopulateDescriptionArray()
'NHRD08082003 Fault 6205
Dim rsPermissionCategory As New ADODB.Recordset
Dim rsPermissionItem As New ADODB.Recordset
Dim sSQL As String
Dim i As Integer

'Select the full names if of the Category items into a temporary table
sSQL = "SELECT DISTINCT CategoryKey, Description FROM ASRSysPermissionCategories"

rsPermissionCategory.Open sSQL, gADOCon, adOpenKeyset, adLockReadOnly
With rsPermissionCategory
  i = 1
  ReDim asCategoryDescriptions(2, .RecordCount + 1)
  Do Until .EOF
    asCategoryDescriptions(1, i) = !CategoryKey
    asCategoryDescriptions(2, i) = !Description
    i = i + 1
    .MoveNext
  Loop
End With

Set rsPermissionCategory = Nothing

'Select the full names if of the Items into a temporary table
sSQL = "SELECT DISTINCT ItemKey, Description FROM ASRSysPermissionItems"
rsPermissionItem.Open sSQL, gADOCon, adOpenKeyset, adLockReadOnly
With rsPermissionItem
  i = 1
  ReDim asItemDescriptions(2, .RecordCount + 1)
  Do Until .EOF
    asItemDescriptions(1, i) = !ItemKey
    asItemDescriptions(2, i) = !Description
    i = i + 1
    .MoveNext
  Loop
End With

Set rsPermissionItem = Nothing

End Sub
Public Function GetCategoryDescription(pstrCategory) As String
'NHRD08082003 Fault 6205
Dim i As Integer
  For i = i To UBound(asCategoryDescriptions, 2)
    If asCategoryDescriptions(1, i) = pstrCategory Then
      GetCategoryDescription = asCategoryDescriptions(2, i)
      Exit For
    End If
  Next
End Function

Public Function GetItemDescription(pstrCategory) As String
'NHRD08082003 Fault 6205
Dim i As Integer
  For i = i To UBound(asItemDescriptions, 2)
    If asItemDescriptions(1, i) = pstrCategory Then
      GetItemDescription = asItemDescriptions(2, i)
      Exit For
    End If
  Next
End Function
