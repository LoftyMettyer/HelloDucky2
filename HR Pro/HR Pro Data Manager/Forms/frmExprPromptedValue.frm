VERSION 5.00
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form frmExprPromptedValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prompted Value"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1039
   Icon            =   "frmExprPromptedValue.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TDBNumber6Ctl.TDBNumber TDBNumericValue 
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   1305
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   556
      Calculator      =   "frmExprPromptedValue.frx":000C
      Caption         =   "frmExprPromptedValue.frx":002C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmExprPromptedValue.frx":127E
      Keys            =   "frmExprPromptedValue.frx":129C
      Spin            =   "frmExprPromptedValue.frx":12E6
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.#######; -##,###,##0.#######"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.#######; -##,###,##0.#######"
      HighlightText   =   -1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBMask6Ctl.TDBMask TDBCharacterValue 
      Height          =   315
      Left            =   2070
      TabIndex        =   2
      Top             =   795
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "frmExprPromptedValue.frx":130E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmExprPromptedValue.frx":1374
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   ""
      HighlightText   =   2
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   ""
      Value           =   ""
   End
   Begin VB.OptionButton optLogicValue 
      Caption         =   "&No"
      Height          =   315
      Index           =   1
      Left            =   1300
      TabIndex        =   6
      Top             =   2300
      Width           =   645
   End
   Begin VB.OptionButton optLogicValue 
      Caption         =   "&Yes"
      Height          =   315
      Index           =   0
      Left            =   300
      TabIndex        =   5
      Top             =   2300
      Value           =   -1  'True
      Width           =   705
   End
   Begin VB.ComboBox cboTableValue 
      Height          =   315
      Left            =   300
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2800
      Width           =   1725
   End
   Begin VB.TextBox txtCharacterValue 
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   800
      Width           =   1560
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2300
      TabIndex        =   8
      Top             =   2800
      Width           =   1200
   End
   Begin GTMaskDate.GTMaskDate ASRDateValue 
      Height          =   315
      Left            =   315
      TabIndex        =   4
      Top             =   1770
      Width           =   1560
      _Version        =   65537
      _ExtentX        =   2752
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
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prompt :"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "frmExprPromptedValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Component variables.
Private mobjComponent As clsExprPromptedValue

' Prompt value variables.
Private mvValue As Variant

' Lookup table prompt variables.
Private mDataType As SQLDataType

Private fPointerWasHourglass As Boolean


Private Sub cboTableValue_Refresh()
  ' Populate the Table Value combo, and then select the default value.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsLookupValues As Recordset
  Dim rsColumnInfo As Recordset
  Dim sDfltValue As String
  Dim iIndex As Integer
  Dim sTableName As String
  Dim sColumnName As String
  Dim lngMaxWidth As Long
  Dim lngWidth As Long
  
  iIndex = 0
  lngMaxWidth = cboTableValue.Width
  
  sDfltValue = mobjComponent.DefaultValue
  
  ' Clear the current contents of the combo.
  cboTableValue.Clear

  ' Get the column and table names.
  sSQL = "SELECT ASRSysColumns.columnName, ASRSysColumns.dataType, ASRSysTables.tableName" & _
    " FROM ASRSysColumns" & _
    " JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysColumns.columnID = " & Trim(Str(mobjComponent.LookupColumn))
  Set rsColumnInfo = datGeneral.GetRecords(sSQL)
  With rsColumnInfo
    fOK = Not (.EOF And .BOF)
    
    If fOK Then
      sTableName = !TableName
      sColumnName = !ColumnName
      mDataType = !DataType
    End If
    
    .Close
  End With
  Set rsColumnInfo = Nothing
  
  If fOK Then
    ' Get the values from the lookup table.
    sSQL = "SELECT DISTINCT " & sColumnName & " AS lookUpValue" & _
      " FROM " & sTableName & _
      " ORDER BY lookUpValue"
    Set rsLookupValues = datGeneral.GetRecords(sSQL)
    With rsLookupValues
      Do While Not .EOF
        Select Case mDataType
          Case sqlNumeric, sqlInteger
            cboTableValue.AddItem Trim(Str(!LookupValue))
            If !LookupValue = Val(sDfltValue) Then
              iIndex = cboTableValue.NewIndex
            End If
            
          Case sqlDate
            If IsDate(!LookupValue) Then
              'JPD 20041118 Fault 8231
              'cboTableValue.AddItem Format(!LookupValue, "long date")
              'If Format(!LookupValue, "mm/dd/yyyy") = sDfltValue Then
              cboTableValue.AddItem Format(!LookupValue, DateFormat)
              If Replace(Format(!LookupValue, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") = sDfltValue Then
                iIndex = cboTableValue.NewIndex
              End If
            End If
            
          Case Else
            cboTableValue.AddItem Trim(!LookupValue)
            
            lngWidth = Me.TextWidth(Trim(!LookupValue)) + 500 'Added a bit for the dropdown arrow
            If lngMaxWidth < lngWidth Then
              lngMaxWidth = lngWidth
            End If
            
            'MH20010823 Fault 2057
            'If !LookupValue = sDfltValue Then
            If Trim(!LookupValue) = Trim(sDfltValue) Then
              iIndex = cboTableValue.NewIndex
            End If
        End Select
            
        .MoveNext
      Loop
          
      .Close
    End With
    Set rsLookupValues = Nothing
  End If
        
  ' Select a list item, and enable the combo, if there are items in the list.
  With cboTableValue
    If .ListCount > 0 Then
      .Enabled = True
      .ListIndex = iIndex
    Else
      .Enabled = False
    End If
  End With

  cboTableValue.Width = lngMaxWidth

  Exit Sub
  
ErrorTrap:
  Set rsLookupValues = Nothing
  cboTableValue.Enabled = False
  Err = False

End Sub



Public Property Set Component(pobjNewValue As clsExprPromptedValue)
  ' Set the component property.
  Set mobjComponent = pobjNewValue
  
  ' Initialize the requiredcontrols with the component values.
  InitializeControls
  
End Property




Private Sub FormatControls()
  ' Format the controls required for the current Prompted Value component.
  Dim iType As ExpressionValueTypes
  
  iType = mobjComponent.ValueType
  
  ' Display the required controls.
  txtCharacterValue.Visible = (iType = giEXPRVALUE_CHARACTER) And _
    (Len(mobjComponent.ValueFormat) = 0)
  TDBCharacterValue.Visible = (iType = giEXPRVALUE_CHARACTER) And _
    (Len(mobjComponent.ValueFormat) > 0)
    
  TDBNumericValue.Visible = (iType = giEXPRVALUE_NUMERIC)
  optLogicValue(0).Visible = (iType = giEXPRVALUE_LOGIC)
  optLogicValue(1).Visible = optLogicValue(0).Visible
  ASRDateValue.Visible = (iType = giEXPRVALUE_DATE)
  cboTableValue.Visible = (iType = giEXPRVALUE_TABLEVALUE)

End Sub


Private Sub InitializeControls()
  ' Initialize the controls required for the current Prompted Value component.
  Dim fHasPrompt As Boolean
  Dim lngXGAP As Long
  Dim lngYGAP As Long
  Dim lngXExtent As Long
  Dim lngCtrlXCoord As Long
  Dim lngMinFormWidth As Long
  Dim lngMaxFormWidth As Long
  Dim iColumnSize As Integer
  Dim dtDateValue As Date
  
  Const XSTART = 200
  Const YSTART = 200
  Const XGAP = 100
  Const YGAP = 200
  Const MAXFORMWIDTH = 10000
  Const MINFORMWIDTH = 3135
    
  ' Initialise and position the prompt.
  fHasPrompt = Len(Trim(mobjComponent.Prompt)) > 0
  With lblPrompt
    .Visible = fHasPrompt
    If fHasPrompt Then
      'TM20010905 Fault 2785
      'Replaces & with && to avoid the & being interpreted as _
      'If other cases of this occur with different characters then a generic function should be created,
      'rather than having a string of Replace statments.
      .Caption = Trim(Replace(mobjComponent.Prompt, "&", "&&")) & IIf(fHasPrompt, " :", "")
      .Left = XSTART
      .Top = YSTART + ((txtCharacterValue.Height - .Height) / 2)
      lngCtrlXCoord = .Left + .Width + XGAP
    Else
      lngCtrlXCoord = XSTART
    End If
  End With
  
  ' Initialise the default return value.
  Select Case mobjComponent.ReturnType
    Case giEXPRVALUE_CHARACTER
      mvValue = ""
    Case giEXPRVALUE_NUMERIC
      mvValue = 0
    Case giEXPRVALUE_LOGIC
      mvValue = True
    Case giEXPRVALUE_DATE
      mvValue = dtDateValue
  End Select

  ' Initialise and position the required controls.
  Select Case mobjComponent.ValueType
    Case giEXPRVALUE_CHARACTER
      If Len(mobjComponent.ValueFormat) > 0 Then
        With TDBCharacterValue
          .Text = mobjComponent.DefaultValue
          .Left = lngCtrlXCoord
          .Top = YSTART
          .Width = Me.TextWidth(String(mobjComponent.ReturnSize, "W")) + _
            (2 * UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
          If (.Left + .Width + XSTART) > MAXFORMWIDTH Then
            .Width = MAXFORMWIDTH - .Left - XSTART
          End If
  
          .Format = mobjComponent.ValueFormat
  
          lngXExtent = .Left + .Width
          cmdOK.Top = .Top + .Height + YGAP
        End With
      Else
        With txtCharacterValue
          .Text = mobjComponent.DefaultValue
          .Left = lngCtrlXCoord
          .Top = YSTART
          .MaxLength = mobjComponent.ReturnSize
          .Width = Me.TextWidth(String(mobjComponent.ReturnSize, "W")) + _
            (2 * UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
          If (.Left + .Width + XSTART) > MAXFORMWIDTH Then
            .Width = MAXFORMWIDTH - .Left - XSTART
          End If
  
          lngXExtent = .Left + .Width
          cmdOK.Top = .Top + .Height + YGAP
        End With
      End If
      
    Case giEXPRVALUE_NUMERIC
      With TDBNumericValue
        'MH20010130 Fault 1610
        datGeneral.FormatTDBNumberControl TDBNumericValue
        
        .Value = mobjComponent.DefaultValue
        .Left = lngCtrlXCoord
        .Top = YSTART
        
        'MH20010130 Size was not taking into account decimal places
        '.Width = Me.TextWidth(String(mobjComponent.ReturnSize, "W")) + _
          (2 * UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
        .Width = Me.TextWidth(String(mobjComponent.ReturnSize + 1 + mobjComponent.ReturnDecimals, "W")) + _
          (2 * UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
        If (.Left + .Width + XSTART) > MAXFORMWIDTH Then
          .Width = MAXFORMWIDTH - .Left - XSTART
        End If
        lngXExtent = .Left + .Width
        cmdOK.Top = .Top + .Height + YGAP
      End With

    Case giEXPRVALUE_LOGIC
      optLogicValue(0).Left = lngCtrlXCoord
      optLogicValue(1).Left = optLogicValue(0).Left + optLogicValue(0).Width + 200
      optLogicValue(0).Top = YSTART + ((txtCharacterValue.Height - optLogicValue(0).Height) / 2)
      optLogicValue(1).Top = optLogicValue(0).Top
      optLogicValue(0).Value = mobjComponent.DefaultValue
      optLogicValue(1).Value = Not optLogicValue(0).Value
      lngXExtent = optLogicValue(1).Left + optLogicValue(1).Width
      cmdOK.Top = optLogicValue(1).Top + optLogicValue(1).Height + YGAP
      
    Case giEXPRVALUE_DATE
      With ASRDateValue
        ' Set the mask date control formats and values.
        If mobjComponent.DefaultValue <> 0 Then
          .Text = mobjComponent.DefaultValue
        End If
        
        .Left = lngCtrlXCoord
        .Top = YSTART
        
        .Width = 1560
        If (.Left + .Width + XSTART) > MAXFORMWIDTH Then
          .Width = MAXFORMWIDTH - .Left - XSTART
        End If
        lngXExtent = .Left + .Width
        cmdOK.Top = .Top + .Height + YGAP
      End With
      
    Case giEXPRVALUE_TABLEVALUE
      
      cboTableValue_Refresh
      
      With cboTableValue
        .Left = lngCtrlXCoord
        .Top = YSTART
        '.Width = 2500
        If (.Left + .Width + XSTART) > MAXFORMWIDTH Then
          .Width = MAXFORMWIDTH - .Left - XSTART
        End If
        lngXExtent = .Left + .Width
        cmdOK.Top = .Top + .Height + YGAP
      End With
        
  End Select
  
  ' Size the form.
  With Me
    .Width = lngXExtent + XSTART + (UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX)
    If .Width < MINFORMWIDTH Then
      .Width = MINFORMWIDTH
    End If

    .Height = cmdOK.Top + cmdOK.Height + YSTART + _
      (Screen.TwipsPerPixelY * (UI.GetSystemMetrics(SM_CYCAPTION) + UI.GetSystemMetrics(SM_CYFRAME)))
  
    ' Position the OK command control.
    cmdOK.Left = (.Width - cmdOK.Width) / 2
    
    'TM20030328 Fault 4173 - define in caption whether the expression is filter or calculation.
    Select Case mobjComponent.BaseComponent.ParentExpression.ExpressionType
    Case giEXPR_RUNTIMEFILTER
      .Caption = "Prompted Filter Value"
    Case giEXPR_RUNTIMECALCULATION
      .Caption = "Prompted Calculation Value"
    Case Else
      .Caption = "Prompted Value"
    End Select
    
  End With
  
End Sub
Public Property Get Value() As Variant
  ' Return the prompted value.
  Value = mvValue

End Property




Private Sub ASRDateValue_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    ASRDateValue.DateValue = Date
  End If

End Sub

Private Sub ASRDateValue_LostFocus()

  'If Not (IsDate(ASRDateValue.DateValue) And ASRDateValue.Text = Format(ASRDateValue.DateValue, DateFormat)) Then
  '  COAMsgBox "Please enter a valid date", vbExclamation, Me.Caption
  '  ASRDateValue.Text = vbNullString
  '  ASRDateValue.SetFocus
  'End If
'  If ASRDateValue.Visible Then
'    If IsNull(ASRDateValue.DateValue) And Not _
'       IsDate(ASRDateValue.DateValue) And _
'       ASRDateValue.Text <> "  /  /" Then
'
'       COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'       ASRDateValue.DateValue = Null
'       ASRDateValue.SetFocus
'       Exit Sub
'    End If
'  End If

  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate ASRDateValue

End Sub

Private Sub cmdOK_Click()
  

  
  ' Validate the entered value and return to the calling method.
  Select Case mobjComponent.ValueType
    Case giEXPRVALUE_CHARACTER
      If Len(mobjComponent.ValueFormat) > 0 Then
        mvValue = TDBCharacterValue.Text
      Else
        mvValue = txtCharacterValue.Text
      End If
      
    Case giEXPRVALUE_NUMERIC
      mvValue = TDBNumericValue.Value
  
    Case giEXPRVALUE_DATE
  
      'MH20020424 Fault 3760
      '(Avoid changing 01/13/2002 to 13/01/2002)
      
      ''' JDM - 13/08/01 - Fault 2628 - Couldn't enter two digit years
      ''If IsDate(ASRDateValue.DateValue) Then 'And ASRDateValue.Text = Format(ASRDateValue.DateValue, DateFormat) Then
      If ValidateGTMaskDate(ASRDateValue) Then
        'TM20020610 Fault 3842 & 3855 - don't allow empty dates in a prompted date value.
        '                               use the User Inteface System Date Separator.
        If Trim(Replace(ASRDateValue.Text, UI.GetSystemDateSeparator, "")) <> vbNullString Then
          mvValue = ASRDateValue.Text
        Else
          mvValue = 0
          COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
          ASRDateValue.SetFocus
          Exit Sub
        End If
      Else
        mvValue = 0
        'COAMsgBox "Please enter a valid date", vbExclamation, Me.Caption
        ASRDateValue.SetFocus
        Exit Sub
      End If

    Case giEXPRVALUE_LOGIC
      mvValue = optLogicValue(0).Value
  
    Case giEXPRVALUE_TABLEVALUE
      Select Case mDataType
        Case sqlVarChar, sqlLongVarChar
          mvValue = cboTableValue.List(cboTableValue.ListIndex)
        Case sqlDate
          If IsDate(cboTableValue.List(cboTableValue.ListIndex)) Then
            mvValue = CDate(cboTableValue.List(cboTableValue.ListIndex))
          Else
            mvValue = 0
          End If
        Case sqlInteger
          mvValue = Val(cboTableValue.List(cboTableValue.ListIndex))
        Case sqlNumeric
          mvValue = Val(cboTableValue.List(cboTableValue.ListIndex))
      End Select
  End Select

  ' Unload the form.
  Unload Me

End Sub

Private Sub Form_Activate()
  ' Format the required prompt controls.
  FormatControls
  
End Sub



Private Sub Form_Load()
  'SetDateComboFormat Me.ASRDateValue
  
  'JPD 20041118 Fault 8231
  UI.FormatGTDateControl ASRDateValue
  
  'RH 23/01/01 - Disable the X button - it does nothing
  EnableCloseButton Me.hWnd, False
  
  If Screen.MousePointer = vbHourglass Then fPointerWasHourglass = True
  Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    Cancel = True
  End If

  If Cancel = False And fPointerWasHourglass Then Screen.MousePointer = vbHourglass
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
  
End Sub

Private Sub txtCharacterValue_GotFocus()
  UI.txtSelText

End Sub



