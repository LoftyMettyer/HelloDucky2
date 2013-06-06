VERSION 5.00
Object = "{604A59D5-2409-101D-97D5-46626B63EF2D}#1.0#0"; "TDBNumbr.ocx"
Begin VB.Form frmQuickFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QuickFind"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1053
   Icon            =   "frmQuickFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Close"
      Height          =   400
      Index           =   1
      Left            =   2150
      TabIndex        =   3
      Top             =   1500
      Width           =   1200
   End
   Begin VB.Frame fraQuickFind 
      Caption         =   "Criteria :"
      Height          =   1250
      Left            =   150
      TabIndex        =   4
      Top             =   100
      Width           =   3200
      Begin TDBNumberCtrl.TDBNumber tdbNumberValue 
         Height          =   315
         Left            =   795
         TabIndex        =   0
         Top             =   795
         Visible         =   0   'False
         Width           =   2205
         _ExtentX        =   3889
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
         MouseIcon       =   "frmQuickFind.frx":000C
         MousePointer    =   0
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Left            =   800
         TabIndex        =   1
         Top             =   700
         Width           =   2200
      End
      Begin VB.ComboBox cboField 
         Height          =   315
         Left            =   800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2200
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         Height          =   195
         Left            =   200
         TabIndex        =   7
         Top             =   765
         Width           =   495
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Field :"
         Height          =   195
         Left            =   200
         TabIndex        =   6
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Find Record"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   1500
      Width           =   1200
   End
End
Attribute VB_Name = "frmQuickFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form that we shall be navigating
Private mfrmParentForm As Form
Private mavColumnInfo() As Variant

Private mfCancelled As Boolean
Public Property Get Cancelled() As Boolean
  ' Return the cancelled flag.
  Cancelled = mfCancelled

End Property

Public Sub Initialise(pfrmParentForm As Form)
  ' Recordset of all the fields for the relevant table
  Dim objColumn As CColumnPrivilege
  Dim sSQL As String
  Dim prstSize As Recordset
  Dim intTemp As Integer
  
  Set mfrmParentForm = pfrmParentForm
  ReDim mavColumnInfo(5, 0)

  ' Clear the existing entries in the the field combo.
  cboField.Clear
  intTemp = 0

  If pfrmParentForm.Name = "frmRecEdit4" Then
    For Each objColumn In pfrmParentForm.ColumnSelectPrivileges
      ' Add any readable, unique columns to the combo.
      ' NB. they must be a reasonable datatype.
      If objColumn.AllowSelect And _
        objColumn.UniqueCheck And _
        (objColumn.DataType = sqlNumeric Or _
        objColumn.DataType = sqlInteger Or _
        objColumn.DataType = sqlDate Or _
        objColumn.DataType = sqlVarChar) Then

        cboField.AddItem RemoveUnderScores(objColumn.ColumnName)
        cboField.ItemData(cboField.NewIndex) = objColumn.ColumnID
              
        mavColumnInfo(0, intTemp) = objColumn.ColumnID
        mavColumnInfo(1, intTemp) = objColumn.Size
        mavColumnInfo(2, intTemp) = objColumn.Decimals
        mavColumnInfo(3, intTemp) = objColumn.DataType
        mavColumnInfo(4, intTemp) = objColumn.UseThousandSeparator
                            
        intTemp = UBound(mavColumnInfo, 2) + 1
        ReDim Preserve mavColumnInfo(5, intTemp)
          
      End If
    
    Next objColumn
    Set objColumn = Nothing
  
  End If

  'If there are no unique fields in the underlying table, inform user
  If cboField.ListCount = 0 Then
    COAMsgBox "Quick Find can only be used on tables with columns defined as unique." & vbCrLf & _
      "The current table has no unique columns.", vbInformation + vbOKOnly, Me.Caption
    Exit Sub
  End If
  
  'If there is only 1 unique field in the table, then select it
  'If cboField.ListCount = 1 Then cboField.ListIndex = 0
  
  cboField.ListIndex = 0
  If cboField.ListCount = 1 Then
    txtValue.TabIndex = 0
    cboField.TabIndex = 4
  End If

  'Display the form
  Me.Show vbModal

End Sub



Private Sub cboField_Click()

Dim strFormat As String
Dim iSize As Integer
Dim iDecimals As Integer
Dim iCount As Integer

'Print cboField.Index
  txtValue.Text = ""
  tdbNumberValue.Text = 0

  Select Case mavColumnInfo(3, cboField.ListIndex)

    Case sqlNumeric
    
      strFormat = IIf(mavColumnInfo(4, cboField.ListIndex), "#0", "0")
      iDecimals = mavColumnInfo(2, cboField.ListIndex)
      iSize = mavColumnInfo(1, cboField.ListIndex)
    
      strFormat = "0"
      For iCount = 2 To (iSize - iDecimals)
        If mavColumnInfo(4, cboField.ListIndex) = True Then
          strFormat = IIf(iCount Mod 3 = 0 And (iCount <> (iSize - iDecimals)), ",#", "#") & strFormat
        Else
          strFormat = "#" & strFormat
        End If
      Next iCount
  
      If iDecimals > 0 Then
        strFormat = strFormat & "."
        For iCount = 1 To iDecimals
          strFormat = strFormat & "0"
        Next iCount
      End If
    
      tdbNumberValue.DisplayFormat = strFormat
      tdbNumberValue.Format = strFormat
      tdbNumberValue.Visible = True
      txtValue.Visible = False

    Case Else
      tdbNumberValue.Visible = False
      txtValue.Visible = True

  End Select

End Sub

Private Sub cmdButton_Click(Index As Integer)
  
  Dim fOK As Boolean
  Dim fFound As Boolean
  Dim iDataType As Integer
  Dim sSQL As String
  Dim sColumnName As String
  Dim sFindString As String
  Dim rsTemp As Recordset
  Dim objColumn As CColumnPrivilege
  Dim lColumnID As Long
  Dim iCount As Integer
  Dim lCurrentID As Long
  Dim strSearchString As String
  Dim strValue As String
  
  On Error GoTo ErrTrap
  
  ' Only do the 'quick find' if the 'Find' button is pressed.
  If Index = 0 Then
    ' Do nothing if the required crieria have not been entered.
    If Not ValidateIt Then Exit Sub
        
    ' ID of the current record before we attempt the quick find
    lCurrentID = mfrmParentForm.RecordID
    
    ' Get the selected column's datatype.
    fFound = False
    For Each objColumn In mfrmParentForm.ColumnSelectPrivileges
      ' Add any readable, unique columns to the combo.
      ' NB. they must be a reasonable datatype.
      If objColumn.ColumnID = cboField.ItemData(cboField.ListIndex) Then
        fFound = True
        iDataType = objColumn.DataType
        sColumnName = objColumn.ColumnName
        Exit For
      End If
    Next objColumn
    Set objColumn = Nothing
    
    If fFound Then
      fOK = True
      
      If iDataType = sqlDate Then
        ' Check that the entered value is a date.
        If Not IsDate(txtValue.Text) Then
          COAMsgBox "You must enter a valid date.", vbInformation + vbOKOnly, Me.Caption
          Exit Sub
        Else
          sSQL = "SELECT ID" & _
            " FROM " & mfrmParentForm.TableName & _
            " WHERE " & sColumnName & " = '" & Replace(Format(ConvertData(txtValue.Text, sqlDate), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'"
        End If
      ElseIf (iDataType = sqlNumeric) Then
'        If Not IsNumeric(txtValue.Text) Then
'          COAMsgBox "You must enter a valid numeric value.", vbInformation + vbOKOnly, Me.Caption
'          Exit Sub
'        Else
        
          ' RH 11/01/01 - Ensure we havent typed too much.
          ' This caused overflow problems in SQL...better solution may be to
          ' set the maxlen property of the text box.
          
          ' RH 16/02/01 - Bug 1873
          
'          For iCount = 0 To UBound(mavColumnInfo, 2)
'            If mavColumnInfo(0, iCount) = cboField.ItemData(cboField.ListIndex) Then
'              If Len(txtValue.Text) > mavColumnInfo(1, iCount) Then
'                COAMsgBox "You have entered " & Len(txtValue.Text) & " characters." & vbCrLf & _
'                       "The " & cboField.Text & " field accepts a maximum of " & mavColumnInfo(1, iCount) & " characters.", vbExclamation + vbOKOnly, App.Title
'                Exit Sub
'              End If
'            End If
'          Next iCount

          ' RH 11/01/01 - Regional Settings - Fix numerics in quickfind.
        
          '        sSQL = "SELECT ID" & _
          '          " FROM " & mfrmParentForm.TableName & _
          '          " WHERE " & sColumnName & " = " & Trim(txtValue.Text)
        
          sSQL = "SELECT ID" & _
            " FROM " & mfrmParentForm.TableName & _
            " WHERE " & sColumnName & " = " & Trim(Replace(Replace(tdbNumberValue.Text, ",", ""), UI.GetSystemDecimalSeparator, "."))
       
          'Dim TempSQL As String
          'TempSQL = sColumnName & " = " & Trim(Replace(txtValue.Text, UI.GetSystemDecimalSeparator, "."))

'        End If
      ElseIf iDataType = sqlVarChar Then
        sSQL = "SELECT ID" & _
          " FROM " & mfrmParentForm.TableName & _
          " WHERE " & sColumnName & " = '" & Replace(txtValue.Text, "'", "''") & "'"
      Else
        fOK = False
      End If
        
      If fOK Then
        
        'mfrmParentForm.Recordset.Seek TempSQL
        
        Set rsTemp = datGeneral.GetRecords(sSQL)
        
        If Not (rsTemp.BOF And rsTemp.EOF) Then
          mfrmParentForm.LocateRecord rsTemp!ID
          If mfrmParentForm.RecordID <> rsTemp!ID Then
            Me.Visible = False
            mfrmParentForm.LocateRecord lCurrentID
            Screen.MousePointer = vbDefault
            
            'JPD 20030905 Fault 6358
            COAMsgBox "No record can be found matching the following criteria:" & vbCrLf & vbCrLf & _
                    cboField.Text & " = " & _
                    IIf((iDataType = sqlNumeric), tdbNumberValue.Text, txtValue.Text) & "." & vbCrLf & vbCrLf & _
                    IIf(mfrmParentForm.Filtered = True, "The current recordset is filtered - try removing the filter.", ""), vbInformation + vbOKOnly, Me.Caption
          End If
          mfrmParentForm.UpdateControls
          mfrmParentForm.UpdateChildren
          
          frmMain.RefreshMainForm frmMain.ActiveForm
          ' JPD20030211 Fault 5043
          mfCancelled = False
        Else
          COAMsgBox "No record can be found matching the following criteria:" & vbCrLf & vbCrLf & cboField.Text & " = " & IIf(iDataType = sqlNumeric, tdbNumberValue.Text, txtValue.Text) & ".", vbInformation + vbOKOnly, Me.Caption
          Screen.MousePointer = vbDefault
          Exit Sub
        End If
        
        rsTemp.Close
        Set rsTemp = Nothing
      End If
    End If
  End If
  
  Unload Me

  Exit Sub
  
ErrTrap:
  
  COAMsgBox "Error whilst attempting to validate quickfind parameters." & vbCrLf & _
         "If this problem persists, please contact support stating :" & vbCrLf & vbCrLf & _
         Err.Number & " - " & Err.Description, vbExclamation + vbOKOnly, App.Title
         
End Sub

Private Function ValidateIt() As Boolean
  ' Has a field been selected in the field combo box?
  If cboField.ListIndex = -1 Then
    COAMsgBox "You must select a field.", vbInformation + vbOKOnly, Me.Caption
    ValidateIt = False
    cboField.SetFocus
    Exit Function
  End If
  
  ' Has anything been entered in the value text box?
  'NHRD16042004 Fault 8775 Changed the And to Or in the line below
   If txtValue.Visible And txtValue.Text = "" Then
    COAMsgBox "You must enter a value.", vbInformation + vbOKOnly, Me.Caption
    ValidateIt = False
    Exit Function
  End If

  If tdbNumberValue.Visible And tdbNumberValue.Text = "0.00" Then
    COAMsgBox "You must enter a value.", vbInformation + vbOKOnly, Me.Caption
    ValidateIt = False
    Exit Function
  End If
    
'  If txtValue.Text = "" Or tdbNumberValue.Text = "0.00" Then
'    COAMsgBox "You must enter a value.", vbInformation + vbOKOnly, Me.Caption
'    ValidateIt = False
'
'    'txtValue.SetFocus
'
'    Exit Function
'  End If
  
  ValidateIt = True
  
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub Form_Load()
  mfCancelled = True

  tdbNumberValue.Top = txtValue.Top

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub



