VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msKey As String
Private miPageNo As Integer
Private mlngTableID As Long
Private mlngColumnID As Long
Private miControlType As ControlTypes
Private mlngTopCoord As Long
Private mlngLeftCoord As Long
Private mlngHeight As Long
Private mlngWidth As Long
Private msCaption As String
Private mclrBackColor As OLE_COLOR
Private mclrForeColor As OLE_COLOR
Private msFontName As String
Private miFontSize As Integer
Private mfFontBold As Boolean
Private mfFontItalic As Boolean
Private mfFontStrikethru As Boolean
Private mfFontUnderline As Boolean
Private mlngPictureID As Long
Private miDisplayType As Integer
Private mstrNavigateTo As String
Private miNavigateIn As enum_NavigateIn
Private mbNavigateOnSave As Boolean
Private miTabIndex As Integer
Private miBorderStyle As Integer
Private miAlignment As Integer
Private msColumnName As String
Private miColumnType As Integer
Private miDataType As SQLDataType
Private msDefaultValue As String
Private mlngSize As Long
Private miDecimals As Integer
Private mlngLookupTableID As Long
Private mlngLookupColumnID As Long
Private mlngLookupFilterColumnID As Long
Private miLookupFilterOperator As FilterOperators
Private mlngLookupFilterValueID As Long
Private mlngSpinnerMinimum As Long
Private mlngSpinnerMaximum As Long
Private mlngSpinnerIncrement As Long
Private mfMandatory As Boolean
Private mfUniqueCheck As Boolean
Private miConvertCase As Integer
Private msMask As String
Private mfBlankIfZero As Boolean
Private mfMultiline As Boolean
Private miColumnAlignment As Integer
Private mlngGotFocusExprID As Long
Private mfReadOnly As Boolean
Private msStatusBarMessage As String
Private msErrorMessage As String
Private mlngLinkTableID As Long
Private mlngLinkViewID As Long
Private mfAFDEnabled As Boolean
Private msTableName As String
Private mvControlValues As Variant
Private mlngDfltValueExprID As Long
Private mlngLinkOrderID As Long
'Private mfOleOnServer As Boolean
Private mbUse1000Separator As Boolean
Private miTrimmingType As TrimmingTypes
Private mfQAddressEnabled As Boolean
Private miOLEType As DataMgr.OLEType
Private miMaxOLESize As Integer
Private mbEmbeddedEnabled As Boolean
Private mbAutoUpdateLookupValues As Boolean


' This needs to be removed once the 1000 separator has been rolled into the application
'Public Property Get DigitGrouping() As Integer
'  DigitGrouping = 3
'End Property

'Public Property Let DigitGrouping(ByVal piNewValue As Integer)
'End Property

'' This needs to be removed once the 1000 separator has been rolled into the application
'Public Property Get DigitSeparator() As String
'  DigitSeparator = ","
'End Property

'Public Property Let DigitSeparator(ByVal pstrNewValue As String)
'End Property

'Public Property Get OleOnServer() As Boolean
'  OleOnServer = mfOleOnServer
'End Property
'
'Public Property Let OleOnServer(ByVal pfNewValue As Boolean)
'  mfOleOnServer = pfNewValue
'End Property

Public Property Get FontName() As String
  FontName = msFontName
End Property

Public Property Let FontName(ByVal psNewValue As String)
  msFontName = psNewValue
End Property

Public Property Get FontSize() As Integer
  FontSize = miFontSize
End Property

Public Property Let FontSize(ByVal piNewValue As Integer)
  miFontSize = piNewValue
End Property

Public Property Get FontBold() As Boolean
  FontBold = mfFontBold
End Property

Public Property Let FontBold(ByVal pfNewValue As Boolean)
  mfFontBold = pfNewValue
End Property

Public Property Get FontItalic() As Boolean
  FontItalic = mfFontItalic
End Property

Public Property Let FontItalic(ByVal pfNewValue As Boolean)
  mfFontItalic = pfNewValue
  
End Property

Public Property Get FontStrikethru() As Boolean
  FontStrikethru = mfFontStrikethru
  
End Property

Public Property Let FontStrikethru(ByVal pfNewValue As Boolean)
  mfFontStrikethru = pfNewValue
  
End Property

Public Property Get FontUnderline() As Boolean
  FontUnderline = mfFontUnderline
  
End Property

Public Property Let FontUnderline(ByVal pfNewValue As Boolean)
  mfFontUnderline = pfNewValue
  
End Property

Public Property Get ControlType() As ControlTypes
  ControlType = miControlType
  
End Property

Public Property Let ControlType(ByVal piNewValue As ControlTypes)
  miControlType = piNewValue
  
End Property

Public Property Get DataType() As SQLDataType
  DataType = miDataType
  
End Property

Public Property Let DataType(ByVal piNewValue As SQLDataType)
  miDataType = piNewValue
  
End Property

Public Property Get PageNo() As Integer
  PageNo = miPageNo
  
End Property

Public Property Let PageNo(ByVal piNewValue As Integer)
  miPageNo = piNewValue
  
End Property

Public Property Get TableID() As Long
  TableID = mlngTableID
  
End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
  
End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
  
End Property

Public Property Let ColumnID(ByVal plngNewValue As Long)
  mlngColumnID = plngNewValue
  
End Property

Public Property Get TopCoord() As Long
  TopCoord = mlngTopCoord
  
End Property

Public Property Let TopCoord(ByVal plngNewValue As Long)
  mlngTopCoord = plngNewValue
  
End Property

Public Property Get LeftCoord() As Long
  LeftCoord = mlngLeftCoord
  
End Property

Public Property Let LeftCoord(ByVal plngNewValue As Long)
  mlngLeftCoord = plngNewValue
  
End Property

Public Property Get Height() As Long
  Height = mlngHeight
  
End Property

Public Property Let Height(ByVal plngNewValue As Long)
  mlngHeight = plngNewValue
  
End Property

Public Property Get Width() As Long
  Width = mlngWidth
  
End Property

Public Property Let Width(ByVal plngNewValue As Long)
  mlngWidth = plngNewValue
  
End Property

Public Property Get Caption() As String
  Caption = msCaption
  
End Property

Public Property Let Caption(ByVal psNewValue As String)
  msCaption = psNewValue
  
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = mclrBackColor
  
End Property

Public Property Let BackColor(ByVal pclrNewValue As OLE_COLOR)
  mclrBackColor = pclrNewValue
  
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = mclrForeColor
  
End Property

Public Property Let ForeColor(ByVal pclrNewValue As OLE_COLOR)
  mclrForeColor = pclrNewValue
  
End Property

Public Property Get PictureID() As Long
  PictureID = mlngPictureID
  
End Property

Public Property Let PictureID(ByVal plngNewValue As Long)
  mlngPictureID = plngNewValue
  
End Property

Public Property Get DisplayType() As Integer
  DisplayType = miDisplayType
  
End Property

Public Property Let DisplayType(ByVal piNewValue As Integer)
  miDisplayType = piNewValue
  
End Property

Public Property Get NavigateTo() As String
  NavigateTo = mstrNavigateTo
End Property

Public Property Let NavigateTo(ByVal NewValue As String)
  mstrNavigateTo = NewValue
End Property

Public Property Get NavigateIn() As Integer
  NavigateIn = miNavigateIn
End Property

Public Property Let NavigateIn(ByVal NewValue As Integer)
  miNavigateIn = NewValue
End Property

Public Property Get NavigateOnSave() As Boolean
  NavigateOnSave = mbNavigateOnSave
End Property

Public Property Let NavigateOnSave(ByVal NewValue As Boolean)
  mbNavigateOnSave = NewValue
End Property

Public Property Get TabIndex() As Integer
  TabIndex = miTabIndex
End Property

Public Property Let TabIndex(ByVal piNewValue As Integer)
  miTabIndex = piNewValue
  
End Property

Public Property Get BorderStyle() As Integer
  BorderStyle = miBorderStyle
  
End Property

Public Property Let BorderStyle(ByVal piNewValue As Integer)
  miBorderStyle = piNewValue
  
End Property

Public Property Get Alignment() As Integer
  Alignment = miAlignment
  
End Property

Public Property Let Alignment(ByVal piNewValue As Integer)
  miAlignment = piNewValue
  
End Property

Public Property Get ColumnName() As String
  ColumnName = msColumnName
  
End Property

Public Property Let ColumnName(ByVal psNewValue As String)
  msColumnName = psNewValue
  
End Property

Public Property Get ColumnType() As Integer
  ColumnType = miColumnType
  
End Property

Public Property Let ColumnType(ByVal piNewValue As Integer)
  miColumnType = piNewValue
  
End Property

Public Property Get DefaultValue() As String
  DefaultValue = msDefaultValue
  
End Property

Public Property Let DefaultValue(ByVal psNewValue As String)
  msDefaultValue = psNewValue
  
End Property

Public Property Get Size() As Long
  Size = mlngSize
End Property

Public Property Let Size(ByVal plngNewValue As Long)
  mlngSize = plngNewValue
End Property

Public Property Get Decimals() As Integer
  Decimals = miDecimals
  
End Property

Public Property Let Decimals(ByVal piNewValue As Integer)
  miDecimals = piNewValue
  
End Property

Public Property Get LookupTableID() As Long
  LookupTableID = mlngLookupTableID
  
End Property

Public Property Let LookupTableID(ByVal plngNewValue As Long)
  mlngLookupTableID = plngNewValue
  
End Property

Public Property Get LookupColumnID() As Long
  LookupColumnID = mlngLookupColumnID
  
End Property

Public Property Let LookupColumnID(ByVal plngNewValue As Long)
  mlngLookupColumnID = plngNewValue
  
End Property

Public Property Get SpinnerMinimum() As Long
  SpinnerMinimum = mlngSpinnerMinimum
  
End Property

Public Property Let SpinnerMinimum(ByVal plngNewValue As Long)
  mlngSpinnerMinimum = plngNewValue
  
End Property

Public Property Get SpinnerMaximum() As Long
  SpinnerMaximum = mlngSpinnerMaximum
  
End Property

Public Property Let SpinnerMaximum(ByVal plngNewValue As Long)
  mlngSpinnerMaximum = plngNewValue
  
End Property

Public Property Get SpinnerIncrement() As Long
  SpinnerIncrement = mlngSpinnerIncrement
  
End Property

Public Property Let SpinnerIncrement(ByVal plngNewValue As Long)
  mlngSpinnerIncrement = plngNewValue
  
End Property

Public Property Get Mandatory() As Boolean
  Mandatory = mfMandatory
  
End Property

Public Property Let Mandatory(ByVal pfNewValue As Boolean)
  mfMandatory = pfNewValue
  
End Property


Public Property Get UniqueCheck() As Boolean
  UniqueCheck = mfUniqueCheck
  
End Property

Public Property Let UniqueCheck(ByVal pfNewValue As Boolean)
  mfUniqueCheck = pfNewValue
  
End Property

Public Property Get ConvertCase() As Integer
  ConvertCase = miConvertCase
  
End Property

Public Property Let ConvertCase(ByVal piNewValue As Integer)
  miConvertCase = piNewValue
  
End Property

Public Property Get Mask() As String
  Mask = msMask
  
End Property

Public Property Let Mask(ByVal psNewValue As String)
  msMask = psNewValue
  
End Property

Public Property Get BlankIfZero() As Boolean
  BlankIfZero = mfBlankIfZero
  
End Property

Public Property Let BlankIfZero(ByVal pfNewValue As Boolean)
  mfBlankIfZero = pfNewValue
  
End Property

Public Property Get Multiline() As Boolean
  Multiline = mfMultiline
  
End Property

Public Property Let Multiline(ByVal pfNewValue As Boolean)
  mfMultiline = pfNewValue
  
End Property

Public Property Get ColumnAlignment() As Integer
  ColumnAlignment = miColumnAlignment
  
End Property

Public Property Let ColumnAlignment(ByVal piNewValue As Integer)
  miColumnAlignment = piNewValue
  
End Property

Public Property Get GotFocusExprID() As Long
  GotFocusExprID = mlngGotFocusExprID
  
End Property

Public Property Get DfltValueExprID() As Long
  DfltValueExprID = mlngDfltValueExprID
  
End Property


Public Property Let GotFocusExprID(ByVal plngNewValue As Long)
  mlngGotFocusExprID = plngNewValue
  
End Property

Public Property Let DfltValueExprID(ByVal plngNewValue As Long)
  mlngDfltValueExprID = plngNewValue
  
End Property


Public Property Get ReadOnly() As Boolean
  ReadOnly = mfReadOnly
  
End Property

Public Property Let ReadOnly(ByVal pfNewValue As Boolean)
  mfReadOnly = pfNewValue
  
End Property

Public Property Get StatusBarMessage() As String
  StatusBarMessage = msStatusBarMessage
  
End Property

Public Property Let StatusBarMessage(ByVal psNewValue As String)
  msStatusBarMessage = psNewValue
  
End Property

Public Property Get ErrorMessage() As String
  ErrorMessage = msErrorMessage
  
End Property

Public Property Let ErrorMessage(ByVal psNewValue As String)
  msErrorMessage = psNewValue
  
End Property

Public Property Get LinkTableID() As Long
  LinkTableID = mlngLinkTableID
  
End Property

Public Property Get LinkViewID() As Long
  LinkViewID = mlngLinkViewID
  
End Property

Public Property Get LinkOrderID() As Long
  LinkOrderID = mlngLinkOrderID
  
End Property


Public Property Let LinkTableID(ByVal plngNewValue As Long)
  mlngLinkTableID = plngNewValue
  
End Property

Public Property Let LinkViewID(ByVal plngNewValue As Long)
  mlngLinkViewID = plngNewValue
  
End Property

Public Property Let LinkOrderID(ByVal plngNewValue As Long)
  mlngLinkOrderID = plngNewValue
  
End Property


Public Property Get AFDEnabled() As Boolean
  AFDEnabled = mfAFDEnabled
  
End Property

Public Property Let AFDEnabled(ByVal pfNewValue As Boolean)
  mfAFDEnabled = pfNewValue
  
End Property

Public Property Get TableName() As String
  TableName = msTableName
  
End Property

Public Property Let TableName(ByVal psNewValue As String)
  msTableName = psNewValue
  
End Property

Public Property Get ControlValues() As Variant
  ControlValues = mvControlValues
  
End Property

Public Property Let ControlValues(ByVal pvNewValue As Variant)
  mvControlValues = pvNewValue
  
End Property

Public Sub ClearControlValues()
  ReDim masControlValues(0)
  
End Sub

Public Property Get Key() As String
  Key = msKey
  
End Property

Public Property Let Key(ByVal psNewValue As String)
  msKey = psNewValue
  
End Property

Public Property Get Use1000Separator() As Boolean
  Use1000Separator = mbUse1000Separator
End Property

Public Property Let Use1000Separator(ByVal pbNewValue As Boolean)
  mbUse1000Separator = pbNewValue
End Property

Public Property Get TrimmingType() As TrimmingTypes
  TrimmingType = miTrimmingType
End Property

Public Property Let TrimmingType(ByVal piNewValue As TrimmingTypes)
  miTrimmingType = piNewValue
End Property

Public Property Get LookupFilterColumnID() As Long
  LookupFilterColumnID = mlngLookupFilterColumnID
End Property

Public Property Get LookupFilterOperator() As FilterOperators
  LookupFilterOperator = miLookupFilterOperator
End Property

Public Property Let LookupFilterColumnID(ByVal plngNewValue As Long)
  mlngLookupFilterColumnID = plngNewValue
End Property

Public Property Let LookupFilterOperator(ByVal piNewValue As FilterOperators)
  miLookupFilterOperator = piNewValue
End Property


Public Property Get LookupFilterValueID() As Long
  LookupFilterValueID = mlngLookupFilterValueID
End Property

Public Property Let LookupFilterValueID(ByVal plngNewValue As Long)
  mlngLookupFilterValueID = plngNewValue
End Property

Public Property Get QAddressEnabled() As Boolean
  QAddressEnabled = mfQAddressEnabled
End Property

Public Property Let QAddressEnabled(ByVal pfNewValue As Boolean)
  mfQAddressEnabled = pfNewValue
End Property

Public Property Get OLEType() As DataMgr.OLEType
  OLEType = miOLEType
End Property

Public Property Let OLEType(ByVal piNewValue As DataMgr.OLEType)
  miOLEType = piNewValue
End Property

Public Property Get MaxOLESize() As Integer
  MaxOLESize = miMaxOLESize
End Property

Public Property Let MaxOLESize(ByVal piNewValue As Integer)
  miMaxOLESize = piNewValue
End Property

Public Property Get EmbeddedEnabled() As Boolean
  EmbeddedEnabled = mbEmbeddedEnabled
End Property

Public Property Let EmbeddedEnabled(ByVal pbNewValue As Boolean)
  mbEmbeddedEnabled = pbNewValue
End Property

Public Property Let AutoUpdateLookupValues(ByVal pbNewValue As Boolean)
  mbAutoUpdateLookupValues = pbNewValue
  
End Property

Public Property Get AutoUpdateLookupValues() As Boolean
  AutoUpdateLookupValues = mbAutoUpdateLookupValues
  
End Property



