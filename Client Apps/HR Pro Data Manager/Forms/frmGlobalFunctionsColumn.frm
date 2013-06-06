VERSION 5.00
Object = "{604A59D5-2409-101D-97D5-46626B63EF2D}#1.0#0"; "TDBNumbr.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmGlobalFunctionsColumn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Global Update Column"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGlobalFunctionsColumn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Value :"
      Height          =   2100
      Left            =   105
      TabIndex        =   2
      Top             =   700
      Width           =   5300
      Begin VB.Frame fraLogicValues 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2250
         TabIndex        =   8
         Top             =   165
         Visible         =   0   'False
         Width           =   2000
         Begin VB.OptionButton optLogicValue 
            Caption         =   "&False"
            Height          =   315
            Index           =   1
            Left            =   1000
            TabIndex        =   10
            Top             =   0
            Width           =   750
         End
         Begin VB.OptionButton optLogicValue 
            Caption         =   "&True"
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   700
         End
      End
      Begin TDBNumberCtrl.TDBNumber tdbNumberValue 
         Height          =   315
         Left            =   2220
         TabIndex        =   13
         Top             =   255
         Visible         =   0   'False
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   999999999
         MinValue        =   -999999999
         Value           =   0
         SelStart        =   0
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   ""
         Format          =   "###############"
         DisplayFormat   =   "###############"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "&Caption"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   0
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmGlobalFunctionsColumn.frx":000C
         MousePointer    =   0
      End
      Begin VB.TextBox txtTextValue 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         TabIndex        =   7
         Top             =   300
         Width           =   2820
      End
      Begin COASpinner.COA_Spinner asrSpinnerValue 
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaximumValue    =   999999999
         MinimumValue    =   -999999999
         Text            =   "999"
      End
      Begin VB.CommandButton cmdTable 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   4725
         TabIndex        =   16
         Top             =   700
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox txtTable 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   700
         Width           =   2475
      End
      Begin VB.ComboBox cboField 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1100
         Width           =   2775
      End
      Begin VB.CommandButton cmdExpr 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   4725
         TabIndex        =   19
         Top             =   1500
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox txtExpr 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1500
         Width           =   2475
      End
      Begin VB.OptionButton optExpr 
         Caption         =   "C&alculation"
         Height          =   315
         Left            =   200
         TabIndex        =   6
         Top             =   1500
         Width           =   1365
      End
      Begin VB.OptionButton optField 
         Caption         =   "Colu&mn"
         Height          =   315
         Left            =   200
         TabIndex        =   5
         Top             =   1100
         Width           =   1125
      End
      Begin VB.OptionButton optTable 
         Caption         =   "&Lookup Table Value"
         Height          =   315
         Left            =   200
         TabIndex        =   4
         Top             =   700
         Width           =   2055
      End
      Begin VB.OptionButton optValue 
         Caption         =   "&Value"
         Height          =   315
         Left            =   200
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   840
      End
      Begin GTMaskDate.GTMaskDate ASRDateValue 
         Height          =   315
         Left            =   2250
         TabIndex        =   14
         Top             =   1710
         Visible         =   0   'False
         Width           =   1530
         _Version        =   65537
         _ExtentX        =   2699
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         NullText        =   "__/__/____"
         BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSelect      =   -1  'True
         MaskCentury     =   2
         SpinButtonEnabled=   0   'False
         BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTips        =   0   'False
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboOptions 
         Height          =   315
         Left            =   2235
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2940
      TabIndex        =   20
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4200
      TabIndex        =   21
      Top             =   3000
      Width           =   1200
   End
   Begin VB.ComboBox cboColumns 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   200
      Width           =   3090
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Column :"
      Height          =   195
      Index           =   0
      Left            =   200
      TabIndex        =   0
      Top             =   260
      Width           =   630
   End
End
Attribute VB_Name = "frmGlobalFunctionsColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlTableID As Long
'Private datGlobal As clsGlobal
Private mtypGlobal As GlobalType

Private mlLookupTableID As Long
Private mlLookupColumnID As Long
Private mlValueType As Long
Private msValue As String
Private mlValueID As Long
Private blnMandatory As Boolean

Private mlParentTableID As Long
Private mfrmParent As Form
Private mbCancelled As Boolean

Private Const sDATA_TYPE = "Invalid data type."
Private miColumnDataType As SQLDataType
Private miControlType As ControlTypes
Private mstrColumnsAlreadySelected As String

Private alngColumnInfo() As Variant
Private mblnOptionGroup As Boolean

Public Property Get LookupTableID() As Long
  LookupTableID = mlLookupTableID
End Property

Public Property Get LookupColumnID() As Long
  LookupColumnID = mlLookupColumnID
End Property

Public Property Get ParentForm() As Form
  Set ParentForm = mfrmParent
End Property

Public Property Let ParentForm(ByVal frmNewValue As Form)
  Set mfrmParent = frmNewValue
  Me.HelpContextID = frmNewValue.HelpContextID
End Property

Public Sub Initialise(bNew As Boolean, lTableID As Long, typGlobal As GlobalType, _
  Optional lColumnID As Long, Optional lValueTypeID As Long, Optional lValueID As Long, _
  Optional sValue As String, Optional lLookupTableID As Long, Optional lLookupColumnID As Long, _
  Optional lParentTableID As Long)

  Dim objExpression As clsExprExpression
  
  'Set datGlobal = New HRProDataMgr.clsGlobal
  
  Call CheckWhichColumnsAreAlreadyUsed(bNew)
  
  ' Clear the controls.
  ClearValueControls
  
  If lValueTypeID = globfuncvaltyp_STRAIGHTVALUE Then   ' Straight value
    Call CheckIfOptionGroup(lColumnID)
  End If
  
  mlTableID = lTableID
  mlParentTableID = lParentTableID
  mtypGlobal = typGlobal
    
  If mtypGlobal = glAdd Then
    Me.Caption = "Global Add Column"
  Else
    Me.Caption = "Global Update Column"
  End If
    
  ' Populate the Columns combo.
  GetColumns

  If Not bNew Then
    ' Put the existing details in the controls if this isn't a new column definition.
    
    'miColumnDataType = GetDataType(lColumnID)
    Call GetDataType(lColumnID, miColumnDataType, miControlType)
    
    SetComboText cboColumns, datGeneral.GetColumnName(lColumnID)
    
    Select Case lValueTypeID
    Case globfuncvaltyp_STRAIGHTVALUE ' Straight value
      If mblnOptionGroup Then
        SetComboText cboOptions, sValue
      
      Else
        optValue.Value = True
        optValue_Click
        
        Select Case miColumnDataType
          Case sqlBoolean
            optLogicValue(0).Value = (sValue = "True")
            optLogicValue(1).Value = (sValue = "False")
          Case sqlNumeric
            tdbNumberValue.Value = Val(datGeneral.ConvertNumberForSQL(Replace(sValue, ",", "")))
          Case sqlDate
            If IsDate(sValue) Then
              ASRDateValue.Text = sValue
            End If
          Case sqlInteger
            If miControlType = ctlSpin Then
              asrSpinnerValue.Text = sValue
            Else
              tdbNumberValue.Value = Val(datGeneral.ConvertNumberForSQL(Replace(sValue, ",", "")))
            End If
          Case sqlVarChar, sqlLongVarChar
            txtTextValue.Text = sValue
        End Select
        
      End If
    
    Case globfuncvaltyp_LOOKUPTABLE ' Lookup table value.
      optTable.Value = True
      optTable_Click
      
      txtTable.Text = sValue
      'cmdTable.Tag = lValueID
      mlLookupTableID = lLookupTableID
      mlLookupColumnID = lLookupColumnID


    Case globfuncvaltyp_FIELD ' Field value.
      optField.Value = True
      optField_Click

      GetParentColumns
      SetComboText cboField, Mid$(sValue, 2, Len(sValue) - 2)
    
    Case globfuncvaltyp_CALCULATION ' Calculated value.
      optExpr.Value = True
      optExpr_Click
      
      cmdExpr.Tag = lValueID
      Set objExpression = New clsExprExpression
      With objExpression
        ' Initialise the expression object.
        .ExpressionID = lValueID
        If Trim$(.Name) = vbNullString Then
          COAMsgBox "This calculation has been deleted by another user", vbExclamation
          cmdExpr.Tag = vbNullString
        Else
          txtExpr.Text = .Name
        End If
      End With
      Set objExpression = Nothing
      
    End Select
  
  Else
    ' If this is a new column definition then select the 'straight value' value type.
    optValue.Value = True
    optValue_Click
  End If

End Sub

Private Sub GetColumns()
  ' Populate the columns combo with a list of non-system, non-link columns.
  ' NB. only character, integer, numeric, logic and date columns are permitted.
  Dim sSQL As String
  Dim rsColumns As Recordset
  Dim clsData As clsDataAccess
  Dim iNextIndex As Integer
  
'  sSQL = "SELECT columnID, columnName, size, decimals, spinnerIncrement, spinnerMaximum, spinnerminimum" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & mlTableID & _
    " AND (columnType = " & Trim(Str(ColData)) & _
    " OR columnType = " & Trim(Str(colLookup)) & _
    " OR columnType = " & Trim(Str(colCalc)) & _
    " OR columnType = " & Trim(Str(colWorkingPattern)) & ")" & _
    " AND (dataType = " & Trim(Str(sqlBoolean)) & _
    " OR dataType = " & Trim(Str(sqlNumeric)) & _
    " OR dataType = " & Trim(Str(sqlInteger)) & _
    " OR dataType = " & Trim(Str(sqlDate)) & _
    " OR dataType = " & Trim(Str(sqlVarChar)) & ")"
  sSQL = "SELECT * FROM ASRSysColumns" & _
    " WHERE tableID = " & mlTableID & _
    " AND ReadOnly = 0 " & _
    " AND (columnType = " & CStr(ColData) & _
    " OR columnType = " & CStr(colLookup) & _
    " OR columnType = " & CStr(colCalc) & ")" & _
    " AND (dataType = " & CStr(sqlBoolean) & _
    " OR dataType = " & CStr(sqlNumeric) & _
    " OR dataType = " & CStr(sqlInteger) & _
    " OR dataType = " & CStr(sqlDate) & _
    " OR dataType = " & CStr(sqlVarChar) & _
    " OR dataType = " & CStr(sqlLongVarChar) & ")"
  
  If mstrColumnsAlreadySelected <> vbNullString Then
    sSQL = sSQL & " AND ColumnID NOT IN (" & mstrColumnsAlreadySelected & ")"
  End If


  Set clsData = New clsDataAccess
  Set rsColumns = clsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set clsData = Nothing
      
  ReDim alngColumnInfo(10, 0)
  
  With cboColumns
    .Clear
    Do While Not rsColumns.EOF
      .AddItem rsColumns!ColumnName
      .ItemData(.NewIndex) = rsColumns!ColumnID
            
      ' Save extra column info in an array.
      iNextIndex = UBound(alngColumnInfo, 2) + 1
      ReDim Preserve alngColumnInfo(10, iNextIndex)
      alngColumnInfo(1, iNextIndex) = rsColumns!ColumnID
      alngColumnInfo(2, iNextIndex) = IIf(IsNull(rsColumns!Size), 0, rsColumns!Size)
      alngColumnInfo(3, iNextIndex) = IIf(IsNull(rsColumns!Decimals), 0, rsColumns!Decimals)
      alngColumnInfo(4, iNextIndex) = IIf(IsNull(rsColumns!SpinnerMaximum), 0, rsColumns!SpinnerMaximum)
      alngColumnInfo(5, iNextIndex) = IIf(IsNull(rsColumns!SpinnerMinimum), 0, rsColumns!SpinnerMinimum)
      alngColumnInfo(6, iNextIndex) = IIf(IsNull(rsColumns!SpinnerIncrement), 0, rsColumns!SpinnerIncrement)
      alngColumnInfo(7, iNextIndex) = IIf(IsNull(rsColumns!ControlType), 0, rsColumns!ControlType)
      alngColumnInfo(8, iNextIndex) = IIf(IsNull(rsColumns!Mandatory), 0, rsColumns!Mandatory)
      alngColumnInfo(9, iNextIndex) = IIf(IsNull(rsColumns!Use1000Separator), 0, rsColumns!Use1000Separator)
      alngColumnInfo(10, iNextIndex) = IIf(IsNull(rsColumns!Multiline), 0, rsColumns!Multiline)
      
      rsColumns.MoveNext
    Loop
    
    If .ListCount > 0 Then
      .ListIndex = 0
      .Enabled = (.ListCount > 1)
      .BackColor = IIf(.ListCount > 1, vbWindowBackground, vbButtonFace)
    End If
  End With

End Sub


Private Sub ASRDateValue_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    ASRDateValue.DateValue = Date
  End If

End Sub

Private Sub ASRDateValue_LostFocus()

'  If IsNull(ASRDateValue.DateValue) And Not _
'     IsDate(ASRDateValue.DateValue) And _
'     ASRDateValue.Text <> "  /  /" Then
'
'     COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     ASRDateValue.DateValue = Null
'     ASRDateValue.SetFocus
'     Exit Sub
'  End If

  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate ASRDateValue

End Sub

Private Sub ASRDateValue_Validate(Cancel As Boolean)

    ' Poxy new date control lets you lose focus for some invalid dates.
'      If ASRDateValue.Text = "  /  /" Then
'        Exit Sub
'      End If
'      If Not IsDate(ASRDateValue.Text) Then
'        COAMsgBox "You have entered an invalid date.", vbExclamation + vbOKOnly, App.Title
'        Cancel = True
'        Exit Sub
'      ElseIf CDate(ASRDateValue.Text) < "01/01/1800" Then
'        COAMsgBox "You have entered an invalid date." & vbCrLf & "Date must be after 01/01/1800.", vbExclamation + vbOKOnly, App.Title
'        Cancel = True
'      End If
  
End Sub

Private Sub cboColumns_Click()
  
'  ' Clear existing parametes if the clumn data type has changed.
'  If miColumnDataType <> datGlobal.GetDataType(cboColumns.ItemData(cboColumns.ListIndex)) Then
'    ' Remember the new data type.
'    miColumnDataType = datGlobal.GetDataType(cboColumns.ItemData(cboColumns.ListIndex))
'
'    ClearValueControls
'
'    If optField.Value Then
'      ' Clear the Field parameters.
'      GetParentColumns
'
'    End If
'  End If
'
'  If optValue.Value Then
'    ' Clear the Value parameters.
'    FormatStraightValueControls
'  End If

  If cboColumns.ListIndex <> -1 Then
    'miColumnDataType = GetDataType(cboColumns.ItemData(cboColumns.ListIndex))
    Call GetDataType(cboColumns.ItemData(cboColumns.ListIndex), miColumnDataType, miControlType)
    
    CheckIfOptionGroup (cboColumns.ItemData(cboColumns.ListIndex))
    FormatStraightValueControls
  
    If optValue.Value = True Then
      FormatValueControls
    Else
      optValue.Value = True
    End If

  End If
  
End Sub


Private Sub cmdCancel_Click()

    mbCancelled = True
    Me.Hide

End Sub

Private Sub cmdExpr_Click()
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  
  fOK = True
  
  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression
  
  With objExpression
    ' Initialise the expression object.
    
    ' Get the data type of the selected column.
    'Select Case miColumnDataType
    '  Case sqlBoolean
    '    fOK = .Initialise(mlTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_LOGIC)
    '  Case sqlNumeric
    '    fOK = .Initialise(mlTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_NUMERIC)
    '  Case sqlInteger
    '    fOK = .Initialise(mlTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_NUMERIC)
    '  Case sqlDate
    '    fOK = .Initialise(mlTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_DATE)
    '  Case sqlVarChar
    '    fOK = .Initialise(mlTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER)
    '  Case Else
    '    fOK = False
    'End Select
    
    
    Select Case miColumnDataType
    Case sqlBoolean
      fOK = .Initialise(mlParentTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_LOGIC)
    Case sqlNumeric
      fOK = .Initialise(mlParentTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_NUMERIC)
    Case sqlInteger
      fOK = .Initialise(mlParentTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_NUMERIC)
    Case sqlDate
      fOK = .Initialise(mlParentTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_DATE)
    Case sqlVarChar, sqlLongVarChar
      fOK = .Initialise(mlParentTableID, Val(cmdExpr.Tag), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_CHARACTER)
    Case Else
      fOK = False
    End Select
    
    
    If fOK Then
      ' Instruct the expression object to display the expression selection/creation/modification form.
      .SelectExpression True
    
      If .Access = "HD" Then
        If mfrmParent.DefinitionCreator = False Then
          'JPD 20030903 Fault 6459
          COAMsgBox "Unable to select this calculation as it is a hidden calculation and you are not the owner of this definition.", vbExclamation
          fOK = False
        End If
      End If

      If fOK Then
        ' Read the selected expression info.
        txtExpr.Text = .Name
        cmdExpr.Tag = .ExpressionID
      End If

    End If
  End With
  
  Set objExpression = Nothing

End Sub

Private Sub cmdOK_Click()
  ' Validate the column update parameters.
  Dim fError As Boolean
  Dim sMsg As String
  Dim strSQL As String
  Dim rsTemp As Recordset

  fError = False
  
  
  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  ' Write the displayed control values to the component
' If ASRDateValue.Visible Then
    If ValidateGTMaskDate(ASRDateValue) = False Then
      fError = True
      Exit Sub
    End If
'    With ASRDateValue      'NHRD09092004 Fault 8895 - - Commented out on 8/12/04
'      If Len(Trim(Replace(.Text, UI.GetSystemDateSeparator, ""))) = 0 Then
'        Clipboard.Clear
'        Clipboard.SetText .Text
'        .DateValue = Null
'        .Paste
'        .ForeColor = vbRed
'
'        DoEvents
'          Dim fSelected As Boolean
'          fSelected = (COAMsgBox("No date has been entered." + vbCrLf + vbCrLf + "Do you want to keep it blank?", vbYesNo + vbExclamation, App.Title) = vbNo)
'
'        If fSelected = False Then
'          .ForeColor = vbWindowText
'          .DateValue = Null
'          If .Visible And .Enabled Then
'            .SetFocus
'          End If
'        Else
'          Exit Sub
'        End If
'      End If
'    End With
'  End If

  If optValue Then
    ' The column is to be updated with a straight value.
    mlValueType = globfuncvaltyp_STRAIGHTVALUE
    If mblnOptionGroup Then
      msValue = cboOptions.Text
    Else
      Select Case miColumnDataType
      Case sqlBoolean
        msValue = IIf(optLogicValue(0).Value, "True", "False")
      Case sqlNumeric, sqlInteger
        If miControlType = ctlSpin Then
          msValue = Trim(asrSpinnerValue.Text)
        Else
          msValue = tdbNumberValue.Text
          'msValue = Val(txtTextValue.Text)
        End If
      Case sqlDate
        If IsDate(ASRDateValue.Text) Then
          msValue = ASRDateValue.Text
        Else
          msValue = "Null"
        End If
      Case sqlVarChar, sqlLongVarChar

        'MH20090825 HRPRO-277 & HRPRO-278
        '''TM20020910 Fault 4395 - Don't RTrim the value.
        '''msValue = RTrim(txtTextValue.Text)
        ''msValue = txtTextValue.Text
        msValue = Replace(txtTextValue.Text, vbTab, " ")

      End Select
    End If
    mlValueID = 0
      
  ElseIf optTable Then
    ' The column is to be updated with a value pulled from a lookup table.
    'mlLookupTableID = 0 Or
    If Trim(txtTable.Text) = vbNullString Then
      COAMsgBox "No lookup table value selected.", vbExclamation + vbOKOnly, Me.Caption
      fError = True
      Exit Sub
    Else
      If Not Validate(globfuncvaltyp_LOOKUPTABLE) Then
        sMsg = sDATA_TYPE
        fError = True
      End If
    End If
    mlValueType = globfuncvaltyp_LOOKUPTABLE
    msValue = Trim$(txtTable.Text)
    'mlValueID = cmdTable.Tag
    mlValueID = 0
  ElseIf optField Then
    ' The column is to be updated with a value pulled from a lookup table.
    If cboField.ListIndex = -1 Then
      COAMsgBox "No field selected.", vbExclamation, Me.Caption
      fError = True
      Exit Sub
    End If

    'MH20010820 Fault 2703 Remove this check.
    'If cboColumns.ItemData(cboColumns.ListIndex) = cboField.ItemData(cboField.ListIndex) Then
    '  sMsg = "Cannot update a column with a value from the same column."
    '  fError = True
    'Else
      If Not Validate(globfuncvaltyp_FIELD) Then
        sMsg = sDATA_TYPE
        fError = True
      End If
    'End If
    mlValueType = globfuncvaltyp_FIELD
    msValue = "<" & cboField.Text & ">"
    mlValueID = cboField.ItemData(cboField.ListIndex)
  Else
    ' The column is to be updated with a value pulled from a lookup table.
    If Val(cmdExpr.Tag) = 0 Then
      sMsg = "No Calculation selected."
      fError = True
    Else
'      strSQL = "SELECT COUNT(*) FROM AsrSysExpressions " & _
'               "WHERE ExprID = " & CStr(cmdExpr.Tag)
'      Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
'      If rsTemp.Fields(0) = 0 Then
'        sMsg = "This calculation has been deleted by another user."
'        txtExpr = vbNullString
'        cmdExpr.Tag = 0
'        fError = True
'      Else
'        mlValueType = globfuncvaltyp_CALCULATION
'        msValue = "<" & txtExpr & ">"
'        mlValueID = Val(cmdExpr.Tag)
'      End If
'
'      rsTemp.Close
'      Set rsTemp = Nothing
      sMsg = IsCalcValid(cmdExpr.Tag)
      If sMsg <> vbNullString Then
        txtExpr = vbNullString
        cmdExpr.Tag = 0
        fError = True
      Else
        mlValueType = globfuncvaltyp_CALCULATION
        msValue = "<" & txtExpr & ">"
        mlValueID = Val(cmdExpr.Tag)
      End If

    End If
  End If
  
  If fError Then
    COAMsgBox sMsg, vbExclamation, Me.Caption
  Else
    'If mlValueType <> globfuncvaltyp_LOOKUPTABLE And Len(cmdTable.Tag) > 0 Then
    '  datGlobal.DeleteTableValue CLng(cmdTable.Tag)
    'End If
      
    mbCancelled = False
    Me.Hide
  End If

End Sub

Private Sub cmdTable_Click()

  'Dim bNew As Boolean
  'Dim lTableID As Long
  'Dim lColumnID As Long
  
  If mlLookupTableID = 0 Then
    Call GetDefaultLookupIDs
  End If
  
  With frmGlobalFunctionsTableValue
    If .Initialise(mlLookupTableID, mlLookupColumnID, miColumnDataType, txtTable.Text) Then
      .HelpContextID = Me.HelpContextID
      .Show vbModal

      If Not .Cancelled Then
        mlLookupTableID = .LookupTableID
        mlLookupColumnID = .LookupColumnID
        msValue = .lstRecords.Text
        txtTable = msValue
      End If
    End If

    'If Not .Cancelled Then
    '  SaveTableValue bNew, .cboTable.ItemData(.cboTable.ListIndex), .cboColumn.ItemData(.cboColumn.ListIndex), _
    '    .lstRecords.ItemData(.lstRecords.ListIndex), .lstRecords.Text
    'End If
  End With
  
  Unload frmGlobalFunctionsTableValue

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub Form_Load()
  'JPD 20041118 Fault 8231
  UI.FormatGTDateControl ASRDateValue
  
  'SetDateComboFormat Me.ASRDateValue
  FormatFormControls
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode = vbFormControlMenu Then
    Cancel = True
    mbCancelled = True
    Me.Hide
  End If

End Sub

'Private Sub Form_Unload(Cancel As Integer)

    'Set datGlobal = Nothing

'End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled

End Property

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optExpr_Click()
  FormatValueControls
    
End Sub

Private Sub optField_Click()
  FormatValueControls

End Sub

Private Sub optTable_Click()
  FormatValueControls

End Sub

Private Sub optValue_Click()
  FormatValueControls
  
End Sub

'Private Sub SaveTableValue(bNew As Boolean, lTableID As Long, lColumnID As Long, lRecordID As Long, _
'  vValue As Variant)
'
'  If bNew Then
'    cmdTable.Tag = datGlobal.NewTableValue(lTableID, lColumnID, lRecordID)
'    txtTable.Text = vValue
'  Else
'    datGlobal.UpdateTableValue CLng(cmdTable.Tag), lTableID, lColumnID, lRecordID
'    txtTable.Text = vValue
'  End If
'
'End Sub

Public Property Get ValueType() As Long
  ValueType = mlValueType

End Property

Public Property Get Value() As String
  Value = msValue

End Property

Public Property Get ValueID() As Long
  ValueID = mlValueID

End Property

Private Sub GetParentColumns()
  ' Populate the columns combo with a list of non-system, non-link columns from the parent table.
  ' NB. only character, integer, numeric, logic and date columns are permitted.
  Dim sSQL As String
  Dim rsColumns As Recordset
  Dim clsData As clsDataAccess
    
'  sSQL = "SELECT columnID, columnName" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & mlParentTableID & _
    " AND (columnType = " & Trim(Str(ColData)) & _
    " OR columnType = " & Trim(Str(colLookup)) & _
    " OR columnType = " & Trim(Str(colCalc)) & _
    " OR columnType = " & Trim(Str(colWorkingPattern)) & ")" & _
    " AND dataType = " & Trim(Str(miColumnDataType))
  sSQL = "SELECT columnID, columnName" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & mlParentTableID & _
    " AND (columnType = " & Trim(Str(ColData)) & _
    " OR columnType = " & Trim(Str(colLookup)) & _
    " OR columnType = " & Trim(Str(colCalc)) & ")" & _
    " AND dataType = " & Trim(Str(miColumnDataType))
  Set clsData = New clsDataAccess
  Set rsColumns = clsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set clsData = Nothing
    
  With cboField
    .Enabled = True
    .Clear
    
    Do While Not rsColumns.EOF
      .AddItem rsColumns!ColumnName
      .ItemData(.NewIndex) = rsColumns!ColumnID
      rsColumns.MoveNext
    Loop
    If .ListCount > 0 Then
      .ListIndex = 0
    Else
      .Enabled = False
    End If
  End With

End Sub

Private Function Validate(lValueType As Long) As Boolean
  ' Validate the entered value matches the given column type.
  Dim vValue As Variant
  Dim lColumnID As Long
  Dim lNewDataType As SQLDataType
    
  Validate = True

  Select Case lValueType
    Case globfuncvaltyp_LOOKUPTABLE      'Value from table
      'vValue = datGlobal.GetGlobalTableValue(CLng(cmdTable.Tag))
      vValue = txtTable.Text
      Validate = ValidateValue(vValue)

    Case globfuncvaltyp_FIELD      'Ref to column
      lColumnID = cboField.ItemData(cboField.ListIndex)
      'lNewDataType = GetDataType(lColumnID)
      Call GetDataType(lColumnID, lNewDataType)
      
      Select Case miColumnDataType
        Case sqlDate
          Validate = (lNewDataType = sqlDate)
            
        Case sqlNumeric, sqlInteger
          Validate = (lNewDataType = sqlInteger Or lNewDataType = sqlNumeric)
                
        Case sqlBoolean
          Validate = (lNewDataType = sqlBoolean)
      End Select
  End Select

End Function

Private Function ValidateValue(vValue As Variant) As Boolean
  ' Validate the entered value matches the required column type.
  ValidateValue = True
  
  Select Case miColumnDataType
    Case sqlDate
      ValidateValue = IsDate(vValue)
    
    Case sqlNumeric, sqlInteger
      ValidateValue = IsNumeric(vValue)
        
    Case sqlBoolean
      ValidateValue = (vValue = 0 Or vValue = 1)
  End Select

End Function


Private Sub FormatValueControls()
  ' Clear the existing values.
  ClearValueControls
  
  If optValue.Value Then
    ' Display only the required Straight value controls.
    If cboColumns.ListIndex <> -1 Then
      Call CheckIfOptionGroup(cboColumns.ItemData(cboColumns.ListIndex))
      FormatStraightValueControls
    End If
  ElseIf optField.Value Then
    ' Populate the combo with the required fields.
    GetParentColumns
  End If
  
  ' Enable/disable the straight value controls.
  txtTextValue.Enabled = optValue.Value
  txtTextValue.BackColor = IIf(optValue.Value, vbWindowBackground, vbButtonFace)
  optLogicValue(0).Enabled = optValue.Value
  optLogicValue(1).Enabled = optValue.Value
  asrSpinnerValue.Enabled = optValue.Value
  asrSpinnerValue.BackColor = IIf(optValue.Value, vbWindowBackground, vbButtonFace)
  ASRDateValue.Enabled = optValue.Value
  ASRDateValue.BackColor = IIf(optValue.Value, vbWindowBackground, vbButtonFace)
  tdbNumberValue.Enabled = optValue.Value
  tdbNumberValue.BackColor = IIf(optValue.Value, vbWindowBackground, vbButtonFace)
  cboOptions.Enabled = optValue.Value
  cboOptions.BackColor = IIf(optValue.Value, vbWindowBackground, vbButtonFace)

  optTable.Enabled = Not (mblnOptionGroup)
  optField.Enabled = Not (mblnOptionGroup)

  ' Enable/disable the Lookup Table value controls.
  cmdTable.Enabled = optTable.Value
  If cmdTable.Enabled = False Then
    mlLookupTableID = 0
    mlLookupColumnID = 0
  End If
  
  ' Enable/disable the Field value controls.
  cboField.Enabled = optField.Value
  cboField.BackColor = IIf(optField.Value, &H80000005, &H8000000F)
    
  ' Enable/disable the Calculated value controls.
  cmdExpr.Enabled = optExpr.Value

End Sub
Private Sub FormatStraightValueControls()
  ' Display the Value controls that match the selcted column's data type.
  Dim iLoop As Integer
  Dim lngCount As Long
  Dim lngSize As Long
  Dim lngDecimals As Long
  Dim lngSpinnerMaximum As Long
  Dim lngSpinnerMinimum As Long
  Dim lngSpinnerIncrement As Long
  Dim lngControlType As Long
  Dim bThousandSeparators As Boolean
  Dim sFormat As String
  Dim bIsUnlimitedSize As Boolean
  
  Dim blnSpinner As Boolean
  Dim blnTextBox As Boolean
  
  
  ' Read the column's details from the array.
  If cboColumns.ListIndex <> -1 Then
    For iLoop = 1 To UBound(alngColumnInfo, 2)
      If alngColumnInfo(1, iLoop) = cboColumns.ItemData(cboColumns.ListIndex) Then
        
        If miColumnDataType = sqlInteger Then
          'Integers always size 10
          '(don't know why but size is stored as 1) !!!!
          lngSize = 10
        Else
          lngSize = alngColumnInfo(2, iLoop)
        End If
        
        lngDecimals = alngColumnInfo(3, iLoop)
        lngSpinnerMaximum = alngColumnInfo(4, iLoop)
        lngSpinnerMinimum = alngColumnInfo(5, iLoop)
        lngSpinnerIncrement = alngColumnInfo(6, iLoop)
        lngControlType = alngColumnInfo(7, iLoop)
        blnMandatory = alngColumnInfo(8, iLoop)
        bThousandSeparators = alngColumnInfo(9, iLoop)
        bIsUnlimitedSize = alngColumnInfo(10, iLoop)
      End If
    Next iLoop
  End If
  
  
  cboOptions.Visible = mblnOptionGroup
  If mblnOptionGroup Then
    fraLogicValues.Visible = False
    tdbNumberValue.Visible = False
    asrSpinnerValue.Visible = False
    ASRDateValue.Visible = False
    txtTextValue.Visible = False
    
    If Not blnMandatory Then
      cboOptions.AddItem ""
    End If
    Exit Sub
  End If


  fraLogicValues.Visible = (miColumnDataType = sqlBoolean)
 
  tdbNumberValue.Visible = (miColumnDataType = sqlNumeric Or ((miColumnDataType = sqlInteger) And (miControlType <> ctlSpin)))
  If (miColumnDataType = sqlNumeric Or miColumnDataType = sqlInteger) Then
    
    ' Loop and create the format mask
    sFormat = "0"
    For lngCount = 2 To (lngSize - lngDecimals)
      If bThousandSeparators = True Then
        sFormat = IIf(lngCount Mod 3 = 0 And (lngCount <> (lngSize - lngDecimals)), ",#", "#") & sFormat
      Else
        sFormat = "#" & sFormat
      End If
    Next lngCount

    If lngDecimals > 0 Then
      sFormat = sFormat & "."
      For lngCount = 1 To lngDecimals
        sFormat = sFormat & "0"
      Next lngCount
    End If
    
    With tdbNumberValue
      .Format = sFormat
      .DisplayFormat = sFormat
    End With
  End If
  
  
  ASRDateValue.Visible = (miColumnDataType = sqlDate)
  If miColumnDataType = sqlDate Then
    'MH20001003 Fault 1048
    'After making the greentree control visible you will
    'be unable to use the arrow keys to scroll though the
    'combo box items.  Setting focus to the form seems to
    'fixed this !
    If Me.Visible Then Me.SetFocus
  End If
  
  
  blnSpinner = (miColumnDataType = sqlInteger And lngControlType = ctlSpin)
  
  asrSpinnerValue.Visible = blnSpinner
  If (miColumnDataType = sqlInteger) Then
    With asrSpinnerValue
      .MinimumValue = lngSpinnerMinimum
      .MaximumValue = lngSpinnerMaximum
      .Increment = lngSpinnerIncrement
    End With
  End If


  blnTextBox = (miColumnDataType = sqlVarChar) Or _
               (miColumnDataType = sqlLongVarChar)

  txtTextValue.Visible = blnTextBox
  If blnTextBox Then
    txtTextValue.MaxLength = IIf(bIsUnlimitedSize, 0, lngSize)
  End If
  
End Sub

Private Sub ClearStraightValueControls()
  ' Reset the Value controls .
  optLogicValue(0).Value = True
  tdbNumberValue.Value = 0
  asrSpinnerValue.Text = "0"
  ASRDateValue.Text = ""
  txtTextValue.Text = ""
  cboOptions.Clear
End Sub


Private Sub ClearValueControls()
  ' Clear all value controls.
  
  ' Clear straight value controls.
  ClearStraightValueControls
  
  ' Clear the Lookup Table value control.
  cmdTable.Tag = ""
  txtTable.Text = ""
  
  ' Clear the Field value control.
  cboField.Clear

  ' Clear the Calculation value controls.
  cmdExpr.Tag = ""
  txtExpr.Text = ""
  
End Sub

Private Sub FormatFormControls()
  ' Position controls that aren't correctly positioned at design time.
  ' ie. the Straight value controls which share the same position.
  Dim dblLeftCoord As Double
  Dim dblTopCoord As Double
  
  dblLeftCoord = txtTextValue.Left
  dblTopCoord = txtTextValue.Top
  
  fraLogicValues.Left = dblLeftCoord
  fraLogicValues.Top = dblTopCoord
  fraLogicValues.BackColor = Me.BackColor
  
  tdbNumberValue.Left = dblLeftCoord
  tdbNumberValue.Top = dblTopCoord
  datGeneral.FormatTDBNumberControl Me.tdbNumberValue

  asrSpinnerValue.Left = dblLeftCoord
  asrSpinnerValue.Top = dblTopCoord
  
  ASRDateValue.Left = dblLeftCoord
  ASRDateValue.Top = dblTopCoord
    
  cboOptions.Left = dblLeftCoord
  cboOptions.Top = dblTopCoord

End Sub

Private Sub txtTextValue_GotFocus()
  With txtTextValue
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  
End Sub


Private Sub CheckWhichColumnsAreAlreadyUsed(blnNew As Boolean)

  Dim lngRow As Long
  Dim pvarbookmark As Variant
  Dim lngColumnID As Long
  Dim lngColumnCurrentlyEditting As Long

  mstrColumnsAlreadySelected = vbNullString
  
  With mfrmParent.grdColumns
    
    If blnNew = False Then
      'Store the ID of the row which you are currently
      'editting and do not include this from the list
      'of columns already selected
      lngColumnCurrentlyEditting = Val(.Columns(2).Text)
    Else
      lngColumnCurrentlyEditting = 0
    End If
    
    
    'MH20001109 Fault 1331
    '.Row = 0
    .MoveFirst
    For lngRow = 0 To .Rows - 1
      pvarbookmark = .GetBookmark(lngRow)
      lngColumnID = Val(.Columns(2).CellText(pvarbookmark))
      
      'If columnID is not column currently editting
      'and columnID is greater than zero then don't
      'allow the user to select the column again
      If lngColumnID > 0 And lngColumnID <> lngColumnCurrentlyEditting Then
        mstrColumnsAlreadySelected = mstrColumnsAlreadySelected & _
          IIf(mstrColumnsAlreadySelected <> "", ", ", "") & CStr(lngColumnID)
      End If
    Next
  End With

End Sub


Private Sub GetDefaultLookupIDs()

  Dim datData As clsDataAccess
  Dim sSQL As String
  Dim rsTemp As Recordset
  
  If cboColumns.ListIndex <> -1 Then
  
    Set datData = New clsDataAccess
  
    'Get default for column
    sSQL = "SELECT LookupTableID as TableID, " & _
           "       LookupColumnID as ColumnID " & _
           "FROM ASRSysColumns " & _
           "WHERE ColumnID = " & CStr(cboColumns.ItemData(cboColumns.ListIndex))
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
    If Not rsTemp.BOF And Not rsTemp.EOF Then
      mlLookupTableID = rsTemp!TableID
      mlLookupColumnID = rsTemp!ColumnID
    End If
  
    rsTemp.Close
    Set rsTemp = Nothing
    Set datData = Nothing

  End If

End Sub


Private Sub GetDataType(lColumnID As Long, lDataType As SQLDataType, Optional lControlType As ControlTypes)

    Dim datData As clsDataAccess
    Dim rsTemp As Recordset
    Dim sSQL As String
  
    Set datData = New clsDataAccess
    
    sSQL = "Select DataType, ControlType From ASRSysColumns Where ColumnID = " & lColumnID
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    'GetDataType = rsTemp(0)
    lDataType = rsTemp(0)
    lControlType = rsTemp(1)

    rsTemp.Close
    Set rsTemp = Nothing

    Set datData = Nothing

End Sub


Private Sub CheckIfOptionGroup(lngColumnID As Long)

  Dim datData As clsDataAccess
  Dim rsControlValues As Recordset
  Dim sSQL As String
  
  Set datData = New clsDataAccess
  
  'JPD 20051101 Fault 10521
  sSQL = "SELECT ASRSysColumnControlValues.value" & _
      " FROM ASRSysColumnControlValues" & _
      " INNER JOIN ASRSysColumns ON ASRSysColumnControlValues.columnID = ASRSysColumns.columnID" & _
      " WHERE ASRSysColumnControlValues.columnID = " & CStr(lngColumnID) & _
      "   AND ASRSysColumns.columnType = 0" & _
      "   AND ASRSysColumns.controlType IN (2,16)" & _
      " ORDER BY ASRSysColumnControlValues.sequence"
  Set rsControlValues = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  
  mblnOptionGroup = (Not rsControlValues.EOF)
  
  If mblnOptionGroup Then
    With cboOptions
      .Clear
      Do While Not rsControlValues.EOF
        .AddItem rsControlValues!Value
        rsControlValues.MoveNext
      Loop
      .ListIndex = 0
    End With
  End If
  
  Set datData = Nothing

End Sub

Private Sub txtTextValue_KeyPress(KeyAscii As Integer)

  'Only allow numbers for integer column

  If miColumnDataType = sqlInteger Then
    If KeyAscii > 31 Then
      If InStr("1234567890.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
        Beep
      End If
    End If
  End If

End Sub

