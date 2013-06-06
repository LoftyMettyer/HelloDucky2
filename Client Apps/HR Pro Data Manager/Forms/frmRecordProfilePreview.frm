VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmRecordProfilePreview 
   Caption         =   "Record Profile"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1065
   Icon            =   "frmRecordProfilePreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8145
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picOutput 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   5640
      ScaleHeight     =   735
      ScaleWidth      =   975
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   800
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   5535
      Begin VB.CheckBox chkShowTableRelationshipTitles 
         Caption         =   "&Show Table Relationship Titles"
         Height          =   200
         Left            =   100
         TabIndex        =   12
         Top             =   500
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkSuppressEmptyTableTitles 
         Caption         =   "Suppress Empty Related Table &Titles"
         Height          =   200
         Left            =   100
         TabIndex        =   11
         Top             =   250
         Width           =   3450
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   400
         Left            =   4095
         TabIndex        =   14
         Top             =   300
         Width           =   1200
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "&Output..."
         Height          =   400
         Left            =   2760
         TabIndex        =   13
         Top             =   300
         Width           =   1200
      End
      Begin VB.CheckBox chkIndent 
         Caption         =   "I&ndent Related Tables"
         Height          =   200
         Left            =   100
         TabIndex        =   10
         Top             =   0
         Value           =   1  'Checked
         Width           =   2355
      End
   End
   Begin MSComCtl2.FlatScrollBar scrollVertical 
      Height          =   1455
      Left            =   5160
      TabIndex        =   5
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1179648
   End
   Begin MSComCtl2.FlatScrollBar scrollHorizontal 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1179649
   End
   Begin VB.PictureBox picContainer 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4755
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Frame fraRecord 
      Caption         =   "Record :"
      Height          =   825
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   5150
      Begin VB.ComboBox cboPage 
         Height          =   315
         ItemData        =   "frmRecordProfilePreview.frx":000C
         Left            =   1500
         List            =   "frmRecordProfilePreview.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   3500
      End
      Begin COASpinner.COA_Spinner asrPage 
         Height          =   315
         Left            =   200
         TabIndex        =   1
         Top             =   300
         Width           =   1000
         _ExtentX        =   1773
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
         MaximumValue    =   99999
         Text            =   "1"
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdOutput 
      Height          =   1620
      Index           =   0
      Left            =   5520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2385
      ScrollBars      =   0
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      ColumnHeaders   =   0   'False
      Col.Count       =   0
      stylesets.count =   2
      stylesets(0).Name=   "Separator"
      stylesets(0).BackColor=   -2147483633
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmRecordProfilePreview.frx":0010
      stylesets(1).Name=   "Heading"
      stylesets(1).BackColor=   -2147483633
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frmRecordProfilePreview.frx":002C
      BevelColorFrame =   0
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      BalloonHelp     =   0   'False
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   4207
      _ExtentY        =   2857
      _StockProps     =   79
      BackColor       =   -2147483643
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblColumnSizingLabel 
      AutoSize        =   -1  'True
      Caption         =   "Do NOT delete me !"
      Height          =   195
      Left            =   6240
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgTemp 
      Height          =   735
      Left            =   6840
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Table"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "frmRecordProfilePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FORM_MINHEIGHT = 6200
Private Const FORM_STARTHEIGHT = 7020
Private Const FORM_STARTWIDTH = 8790
Private Const CBOPAGE_MINWIDTH = 4100

Private mblnLoading As Boolean
'Private gblnBatchMode As Boolean
Private mblnUserCancelled As Boolean
Private mstrErrorMessage As String
Private msRecordProfileName As String
Private mrsResults As ADODB.Recordset
Private malngEmptyTableCaptions() As Long

Private mlngOutputFormat As Long
Private mblnOutputScreen As Boolean
Private mblnOutputPrinter As Boolean
Private mstrOutputPrinterName As String
Private mblnOutputSave As Boolean
Private mlngOutputSaveExisting As Long
'Private mlngOutputSaveFormat As Long
Private mblnOutputEmail As Boolean
Private mlngOutputEmailAddr As Long
Private mstrOutputEmailSubject As String
Private mstrOutputEmailAttachAs As String
'Private mlngOutputEmailFileFormat As Long
Private mstrOutputFileName As String

Private mblnIndentRelatedTables As Boolean
Private mblnSuppressEmptyRelatedTableTitles As Boolean
Private mblnSuppressTableRelationshipTitles As Boolean

Private mlngBC_Data As Long
Private mlngFC_Data As Long
Private mlngBC_Heading As Long
Private mlngFC_Heading As Long

Private mobjDefinition As clsRecordProfileTabDtls

Private Const sTYPECODE_HEADING = "H"
Private Const sTYPECODE_SEPARATOR = "S"
Private Const sTYPECODE_COLUMN = "C"
Private Const sTYPECODE_ID = "I"

Private Const sPHOTOSTYLESET = "PhotoSS_"
  
Private Const sCOLUMN_VALUE = "Value"
Private Const sCOLUMN_HEADING = "Heading"
Private Const sCOLUMN_ISHEADING = "IsHeading"
Private Const sCOLUMN_DECPLACES = "DecPlaces"
Private Const sCOLUMN_ISPHOTO = "IsPhoto"
Private Const sCOLUMN_ID = "ID"
Private Const sCOLUMN_BLANKIFZERO = "BlankIfZero"
Private Const sCOLUMN_USE1000SEPARATOR = "ThousandSeparator"

Private Type PROFILEINFO
  TableID As Long
  IsGrid As Boolean
  Index As Long
  FollowsGrid As Long
  RecordIDs As String
  AssociatedControlIndex As Long
End Type

Private matypInfo() As PROFILEINFO

Private malngMultiRecordGrids() As Long

Private Const STANDARDINDENT = 100
Private Const DYNAMICINDENT = 500
Private Const GAPABOVELABEL = 100
Private Const GAPABOVEGRID = 100
Private Const XGAP = 200
Private Const SEPARATORWIDTH = 200
Private Const PHOTOFRAME = 100

Private Const STANDARDROWHEIGHT = 250
Private Const STANDARDCOLUMNWIDTH = 2000

Private Const SCROLLMAX = 32767

Private mdblVerticalScrollRatio As Double
Private mdblHorizontalScrollRatio As Double

Private Const RECPROFFOLLOWONCORRECTION = 10

Private malngControlOrder() As Long

Public Property Let OutputFormat(plngOutputFormat As Long)
  mlngOutputFormat = plngOutputFormat
End Property

Public Property Let OutputScreen(pblnOutputScreen As Boolean)
  mblnOutputScreen = pblnOutputScreen
End Property


Public Property Let OutputPrinter(pblnOutputPrinter As Boolean)
  mblnOutputPrinter = pblnOutputPrinter
End Property


Public Property Let OutputPrinterName(pstrOutputPrinterName As String)
  mstrOutputPrinterName = pstrOutputPrinterName
End Property


Public Property Let OutputSave(pblnOutputSave As Boolean)
  mblnOutputSave = pblnOutputSave
End Property


Public Property Let OutputSaveExisting(plngOutputSaveExisting As Long)
  mlngOutputSaveExisting = plngOutputSaveExisting
End Property

'Public Property Let OutputSaveFormat(plngOutputSaveFormat As Long)
'  mlngOutputSaveFormat = plngOutputSaveFormat
'End Property



Public Function IsEmptyTableCaption(plngIndex As Long) As Boolean
  Dim lngLoop As Long
  Dim fIsEmpty As Boolean
  
  fIsEmpty = False
  
  For lngLoop = 1 To UBound(malngEmptyTableCaptions)
    If plngIndex = malngEmptyTableCaptions(lngLoop) Then
      fIsEmpty = True
      Exit For
    End If
  Next lngLoop
  
  IsEmptyTableCaption = fIsEmpty
  
End Function

Public Function AssociatedControlIndex(pfIsGrid As Boolean, plngIndex As Long) As Long
  Dim lngLoop As Long
  Dim lngAssociatedControlIndex As Long
  
  lngAssociatedControlIndex = 0
  
  For lngLoop = 1 To UBound(matypInfo)
    If (matypInfo(lngLoop).IsGrid = pfIsGrid) And _
      (matypInfo(lngLoop).Index = plngIndex) Then
      lngAssociatedControlIndex = matypInfo(lngLoop).AssociatedControlIndex
      Exit For
    End If
  Next lngLoop
  
  AssociatedControlIndex = lngAssociatedControlIndex
  
End Function


Public Property Let SuppressTableRelationshipTitles(pblnSuppressTableRelationshipTitles As Boolean)
  mblnSuppressTableRelationshipTitles = pblnSuppressTableRelationshipTitles

  chkShowTableRelationshipTitles.Value = IIf(mblnSuppressTableRelationshipTitles, vbUnchecked, vbChecked)
End Property


Public Property Let IndentRelatedTables(pblnIndentRelatedTables As Boolean)
  mblnIndentRelatedTables = pblnIndentRelatedTables
  
  chkIndent.Value = IIf(mblnIndentRelatedTables, vbChecked, vbUnchecked)
  
End Property

Public Property Let SuppressEmptyRelatedTableTitles(pblnSuppressEmptyRelatedTableTitles As Boolean)
  mblnSuppressEmptyRelatedTableTitles = pblnSuppressEmptyRelatedTableTitles

  chkSuppressEmptyTableTitles.Value = IIf(mblnSuppressEmptyRelatedTableTitles, vbChecked, vbUnchecked)
End Property

Public Property Let OutputEmail(pblnOutputEmail As Boolean)
  mblnOutputEmail = pblnOutputEmail
End Property


Public Property Let OutputEmailAddr(plngOutputEmailAddr As Long)
  mlngOutputEmailAddr = plngOutputEmailAddr
End Property


Public Property Let OutputEmailSubject(pstrOutputEmailSubject As String)
  mstrOutputEmailSubject = pstrOutputEmailSubject
End Property

Public Property Let OutputEmailAttachAs(pstrOutputEmailAttachAs As String)
  mstrOutputEmailAttachAs = pstrOutputEmailAttachAs
End Property

'Public Property Let OutputEmailFileFormat(plngOutputEmailFileFormat As Long)
'  mlngOutputEmailFileFormat = plngOutputEmailFileFormat
'End Property


Public Property Let OutputFilename(pstrOutputFilename As String)
  mstrOutputFileName = pstrOutputFilename
End Property


Private Function FormatGrid_Horizontal(pgrdGrid As SSDBGrid, _
  pobjRecProfTable As clsRecordProfileTabDtl) As Boolean
  
  ' Horizontal grid display
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim lngColWidth As Long
  Dim objRecProfColumn As clsRecordProfileColDtl
  Dim fLastColumnWasHeader As Boolean
  
  fLastColumnWasHeader = False
  
  With pgrdGrid
    .ColumnHeaders = True
    .GroupHeaders = pobjRecProfTable.HasHeadings
    .Columns.RemoveAll
    .Groups.RemoveAll

    For Each objRecProfColumn In pobjRecProfTable.Columns
      objRecProfColumn.GridColumnWidth = 0

      Select Case objRecProfColumn.ColType
        Case sTYPECODE_HEADING
          If fLastColumnWasHeader Then
            ' Two headers in a row ! Put a separator after the preceding header to avoid dodgy gaps in the grid.
            pobjRecProfTable.Columns.Add CStr(pobjRecProfTable.Columns.Count + 1), _
              sTYPECODE_SEPARATOR, _
              0, _
              "", _
              0, _
              0, _
              0, _
              pobjRecProfTable.TableID, _
              "", _
              objRecProfColumn.Sequence, _
              objRecProfColumn.Key

            .Groups(.Groups.Count - 1).Columns.Add .Groups(.Groups.Count - 1).Columns.Count
            .Groups(.Groups.Count - 1).Width = .Groups(.Groups.Count - 1).Width + SEPARATORWIDTH
            .Columns(.Columns.Count - 1).StyleSet = "Separator"
            .Columns(.Columns.Count - 1).Caption = ""
            .Columns(.Columns.Count - 1).Locked = True
            .Columns(.Columns.Count - 1).Width = SEPARATORWIDTH
            .Columns(.Columns.Count - 1).TagVariant = "-1"
          End If
          
          .Groups.Add .Groups.Count
          .Groups(.Groups.Count - 1).Caption = objRecProfColumn.Heading
          .Groups(.Groups.Count - 1).CaptionAlignment = ssCaptionAlignmentCenter
          .Groups(.Groups.Count - 1).Width = 0

          fLastColumnWasHeader = True
          
        Case sTYPECODE_COLUMN, sTYPECODE_ID
          If (pobjRecProfTable.HasHeadings) Then
            If (.Groups.Count = 0) Then
              .Groups.Add .Groups.Count
              .Groups(.Groups.Count - 1).Width = 0
            End If

            .Groups(.Groups.Count - 1).Columns.Add .Groups(.Groups.Count - 1).Columns.Count
          Else
            .Columns.Add .Columns.Count
          End If

          .Columns(.Columns.Count - 1).Caption = objRecProfColumn.Heading
          .Columns(.Columns.Count - 1).Locked = True
          .Columns(.Columns.Count - 1).Visible = (objRecProfColumn.ColType = sTYPECODE_COLUMN)
          .Columns(.Columns.Count - 1).Name = CStr(objRecProfColumn.PositionInRecordset)
          .Columns(.Columns.Count - 1).CaptionAlignment = ssColCapAlignCenter
          
          ' Have to use properties to memorise values (ugly - but it works)
          .Columns(.Columns.Count - 1).TagVariant = CStr(IIf(objRecProfColumn.IsNumeric, objRecProfColumn.DecPlaces, -1))
          .Columns(.Columns.Count - 1).AllowSizing = objRecProfColumn.BlankIfZero
          .Columns(.Columns.Count - 1).ButtonsAlways = objRecProfColumn.ThousandSeparator

          If objRecProfColumn.IsNumeric Then
            .Columns(.Columns.Count - 1).Alignment = ssCaptionAlignmentRight
          ElseIf objRecProfColumn.IsLogic Then
            .Columns(.Columns.Count - 1).Alignment = ssCaptionAlignmentCenter
          ElseIf objRecProfColumn.IsPhoto Then
            .Columns(.Columns.Count - 1).TagVariant = sCOLUMN_ISPHOTO
            .Columns(.Columns.Count - 1).Style = ssStyleButton
          End If

          lngColWidth = WidthOfText(objRecProfColumn.Heading)
          objRecProfColumn.GridColumnWidth = IIf(objRecProfColumn.ColType = sTYPECODE_COLUMN, _
            XGAP + lngColWidth, 0)

          If (pobjRecProfTable.HasHeadings) Then
            .Groups(.Groups.Count - 1).Width = .Groups(.Groups.Count - 1).Width + objRecProfColumn.GridColumnWidth
          End If

          .Columns(.Columns.Count - 1).Width = objRecProfColumn.GridColumnWidth

          fLastColumnWasHeader = False
        
        Case Else
          objRecProfColumn.GridColumnWidth = SEPARATORWIDTH

          If (pobjRecProfTable.HasHeadings) Then
            If (.Groups.Count = 0) Then
              .Groups.Add .Groups.Count
              .Groups(.Groups.Count - 1).Width = 0
            End If

            .Groups(.Groups.Count - 1).Columns.Add .Groups(.Groups.Count - 1).Columns.Count
            .Groups(.Groups.Count - 1).Width = .Groups(.Groups.Count - 1).Width + objRecProfColumn.GridColumnWidth
          Else
            .Columns.Add .Columns.Count
          End If

          .Columns(.Columns.Count - 1).StyleSet = "Separator"
          .Columns(.Columns.Count - 1).Caption = ""
          .Columns(.Columns.Count - 1).Locked = True
          .Columns(.Columns.Count - 1).Width = objRecProfColumn.GridColumnWidth
          .Columns(.Columns.Count - 1).TagVariant = "-1"
      
          fLastColumnWasHeader = False
      End Select
    Next objRecProfColumn
    Set objRecProfColumn = Nothing

    If pobjRecProfTable.HasHeadings Then
      iCount = -1
      iCount2 = 0

      For Each objRecProfColumn In pobjRecProfTable.Columns
        If objRecProfColumn.Displayed Then
          If objRecProfColumn.ColType = sTYPECODE_HEADING Then
            iCount = iCount + 1
            iCount2 = 0
          Else
            If iCount < 0 Then
              iCount = 0
            End If

            If iCount2 = (.Groups(iCount).Columns.Count - 1) Then
              .Groups(iCount).Columns(iCount2).Width = .Groups(iCount).Width - (.Groups(iCount).Columns(iCount2).Left - .Groups(iCount).Left)
            Else
              .Groups(iCount).Columns(iCount2).Width = objRecProfColumn.GridColumnWidth + 5
            End If

            If .Groups(iCount).Columns(iCount2).Width < objRecProfColumn.GridColumnWidth + 5 Then
              .Groups(iCount).Width = .Groups(iCount).Width + (objRecProfColumn.GridColumnWidth + 5 - .Groups(iCount).Columns(iCount2).Width)
              .Groups(iCount).Columns(iCount2).Width = objRecProfColumn.GridColumnWidth + 5
            End If

            iCount2 = iCount2 + 1
          End If
        End If
      Next objRecProfColumn
      Set objRecProfColumn = Nothing
    End If
    
    .Height = .RowHeight + IIf(pobjRecProfTable.HasHeadings, .RowHeight + 25, 0) + 50

    If (pobjRecProfTable.HasHeadings) Then
      .Groups(.Groups.Count - 1).Width = .Groups(.Groups.Count - 1).Columns(.Groups(.Groups.Count - 1).Columns.Count - 1).Left _
        - .Groups(.Groups.Count - 1).Left _
        + .Groups(.Groups.Count - 1).Columns(.Groups(.Groups.Count - 1).Columns.Count - 1).Width
      
      .Width = .Groups(.Groups.Count - 1).Left + .Groups(.Groups.Count - 1).Width
    Else
      .Width = .Columns(.Columns.Count - 1).Left + .Columns(.Columns.Count - 1).Width
    End If
  End With
  
End Function

Private Function FormatIndents(pfCheckGroups As Boolean) As Boolean
  Dim fIndent As Boolean
  Dim lngLoop As Long
  Dim ctlLabel As Label
  Dim ctlGrid As SSDBGrid
  Dim objRecProfTable As clsRecordProfileTabDtl
  Dim alngWidths() As Long
  Dim iIndents As Integer
  Dim objGroup As Group
  Dim lngColWidth As Long
  
  fIndent = (chkIndent.Value = vbChecked)
  
  ReDim alngWidths(picOutput.Count - 1)
  
  For Each ctlLabel In lblCaption
    With ctlLabel
      If .Index > 0 Then
        If fIndent Then
          Set objRecProfTable = mobjDefinition.Item(.Tag)
          iIndents = IIf(objRecProfTable.Generation = 0, 0, objRecProfTable.Generation)
          .Left = STANDARDINDENT + (DYNAMICINDENT * iIndents)
          Set objRecProfTable = Nothing
        Else
          .Left = STANDARDINDENT
        End If
        
        If (.Left + .Width) > alngWidths(.Container.Index) Then
          alngWidths(.Container.Index) = (.Left + .Width)
        End If
      End If
    End With
  Next ctlLabel
  Set ctlLabel = Nothing
  
  For Each ctlGrid In grdOutput
    With ctlGrid
      If .Index > 0 Then
        Set objRecProfTable = mobjDefinition.Item(.Tag)
        
        If pfCheckGroups And _
          (objRecProfTable.Orientation = giHORIZONTAL) And _
          (objRecProfTable.HasHeadings) Then
          
          For Each objGroup In .Groups
            lngColWidth = WidthOfText(objGroup.Caption) + XGAP
            
            If lngColWidth > objGroup.Width Then
              objGroup.Width = lngColWidth
            End If
          Next objGroup
          Set objGroup = Nothing
        
          .Width = .Groups(.Groups.Count - 1).Left + .Groups(.Groups.Count - 1).Width
        End If
        
        If fIndent Then
          iIndents = IIf(objRecProfTable.Generation = 0, 0, objRecProfTable.Generation)
          .Left = STANDARDINDENT + (DYNAMICINDENT * iIndents)
        Else
          .Left = STANDARDINDENT
        End If
        
        If (.Left + .Width) > alngWidths(.Container.Index) Then
          alngWidths(.Container.Index) = (.Left + .Width)
        End If
      
        Set objRecProfTable = Nothing
      End If
    End With
  Next ctlGrid
  Set ctlGrid = Nothing
  
  For lngLoop = 1 To (picOutput.Count - 1)
    picOutput(lngLoop).Width = alngWidths(lngLoop) + STANDARDINDENT + 200
  Next lngLoop
  
End Function


Private Function FormatGrid_Vertical2(pgrdGrid As SSDBGrid, _
  pobjRecProfTable As clsRecordProfileTabDtl, _
  piStartColumn As Integer, _
  pfJustAddedPhotoGrid As Boolean, _
  pfFollowingOn As Boolean, _
  plngPicBoxIndex As Long) As Boolean
  
  Dim objRecProfColumn As clsRecordProfileColDtl
  Dim lngColWidth As Long
  Dim sAddString As String
  Dim iColumnCount As Integer
  
  iColumnCount = 0
  
  With pgrdGrid
    ' Add 2 columns.
    .Columns.RemoveAll
    
    ' Column 1 : headings
    .Columns.Add .Columns.Count
    .Columns(.Columns.Count - 1).ButtonsAlways = True
    .Columns(.Columns.Count - 1).Style = ssStyleButton
    .Columns(.Columns.Count - 1).Locked = True
    .Columns(.Columns.Count - 1).Name = sCOLUMN_HEADING
    .Columns(.Columns.Count - 1).Width = 0
  
    ' Column 2 : heading styleset flag (hidden)
    .Columns.Add .Columns.Count
    .Columns(.Columns.Count - 1).Visible = False
    .Columns(.Columns.Count - 1).Locked = True
    .Columns(.Columns.Count - 1).Name = sCOLUMN_ISHEADING
    .Columns(.Columns.Count - 1).Width = 0
    
    ' Column 3 : decimal places (hidden) - required when ouputing to Excel
    .Columns.Add .Columns.Count
    .Columns(.Columns.Count - 1).Visible = False
    .Columns(.Columns.Count - 1).Locked = True
    .Columns(.Columns.Count - 1).Name = sCOLUMN_DECPLACES
    .Columns(.Columns.Count - 1).Width = 0
    
    ' Column 4 : Blank if zero (hidden) - required when ouputing to Excel
    .Columns.Add .Columns.Count
    .Columns(.Columns.Count - 1).Visible = False
    .Columns(.Columns.Count - 1).Locked = True
    .Columns(.Columns.Count - 1).Name = sCOLUMN_BLANKIFZERO
    .Columns(.Columns.Count - 1).Width = 0
    
    ' Column 5 : Use 1000 Separator (hidden) - required when ouputing to Excel
    .Columns.Add .Columns.Count
    .Columns(.Columns.Count - 1).Visible = False
    .Columns(.Columns.Count - 1).Locked = True
    .Columns(.Columns.Count - 1).Name = sCOLUMN_USE1000SEPARATOR
    .Columns(.Columns.Count - 1).Width = 0
    
    For Each objRecProfColumn In pobjRecProfTable.Columns
      iColumnCount = iColumnCount + 1
      
      If iColumnCount >= piStartColumn Then
        If objRecProfColumn.Displayed Then
          If pfJustAddedPhotoGrid And _
            (iColumnCount > piStartColumn) Then
            FormatOutput_AddGrid2 pobjRecProfTable, iColumnCount, pgrdGrid.Index, objRecProfColumn.IsPhoto, plngPicBoxIndex, False, 0, 0
            Exit For
          End If
          
          Select Case objRecProfColumn.ColType
            Case sTYPECODE_COLUMN
              If objRecProfColumn.IsPhoto Then
                If iColumnCount > piStartColumn Then
                  FormatOutput_AddGrid2 pobjRecProfTable, iColumnCount, pgrdGrid.Index, True, plngPicBoxIndex, False, 0, 0
                  Exit For
                End If
              End If
  
              sAddString = objRecProfColumn.Heading & _
                vbTab & "0" & _
                vbTab & CStr(IIf(objRecProfColumn.IsNumeric, objRecProfColumn.DecPlaces, -1))
  
              lngColWidth = XGAP + WidthOfText(objRecProfColumn.Heading)
              If (lngColWidth > .Columns(sCOLUMN_HEADING).Width) Then
                .Columns(sCOLUMN_HEADING).Width = XGAP + lngColWidth
              End If
  
              If pfFollowingOn Then
                If .Columns(sCOLUMN_HEADING).Width < grdOutput(pgrdGrid.Index - 1).Columns(sCOLUMN_HEADING).Width Then
                  .Columns(sCOLUMN_HEADING).Width = grdOutput(pgrdGrid.Index - 1).Columns(sCOLUMN_HEADING).Width
                ElseIf .Columns(sCOLUMN_HEADING).Width > grdOutput(pgrdGrid.Index - 1).Columns(sCOLUMN_HEADING).Width Then
                  FormatPrecedingGridHeadingWidth pgrdGrid.Index - 1, .Columns(sCOLUMN_HEADING).Width
                End If
              End If
    
            Case sTYPECODE_HEADING
              sAddString = objRecProfColumn.Heading & _
                vbTab & "1" & _
                vbTab & "0"
              'JPD 20030728 Fault 6225
              Me.FontBold = True
              lngColWidth = XGAP + WidthOfText(objRecProfColumn.Heading)
              Me.FontBold = False
              
              If (lngColWidth > .Columns(sCOLUMN_HEADING).Width) Then
                .Columns(sCOLUMN_HEADING).Width = XGAP + lngColWidth
              End If
  
            Case Else
              sAddString = "" & _
                vbTab & "1" & _
                vbTab & "0"
          End Select
  
          .AddItem sAddString
        End If
      End If
    Next objRecProfColumn
    Set objRecProfColumn = Nothing

    .Height = (.Rows * .RowHeight) + 15
    .Width = .Columns(.Columns.Count - 1).Left + .Columns(.Columns.Count - 1).Width
  End With
  
End Function

  


Private Function FormatOutput_AddDataToVerticalGrid(pobjRecProfTable As clsRecordProfileTabDtl, _
  pgrdGrid As SSDBGrid, _
  piStartColumn As Integer, _
  pfAddingPhotoData As Boolean, _
  pfFollowingOn As Boolean)
  
  Dim iCount As Integer
  Dim objRecProfColumn As clsRecordProfileColDtl
  Dim sTemp As String
  Dim lngColWidth As Long
  Dim iColumnCount As Integer
  Dim sPhotoSS As String
  Dim iLoop As Integer
  Dim objOLEStream As New ADODB.Stream
  
  iColumnCount = 0
  
  With pgrdGrid
    ' Add a new column for the data.
    .Columns.Add .Columns.Count
    .Columns(.Columns.Count - 1).Locked = True
    .Columns(.Columns.Count - 1).Width = 0
  
    iCount = 0
    For Each objRecProfColumn In pobjRecProfTable.Columns
      iColumnCount = iColumnCount + 1
      
      If iColumnCount >= piStartColumn Then
        If objRecProfColumn.Displayed Then
          iCount = iCount + 1
        
          If pfAddingPhotoData And _
            (iColumnCount > piStartColumn) Then
            'FormatOutput_AddDataToVerticalGrid pobjRecProfTable, _
              grdOutput(pgrdGrid.Index + 1), _
              iColumnCount, _
              False, _
              True
            FormatOutput_AddDataToVerticalGrid pobjRecProfTable, _
              grdOutput(pgrdGrid.Index + 1), _
              iColumnCount, _
              objRecProfColumn.IsPhoto, _
              True
            
            Exit For
          End If
        
          If objRecProfColumn.ColType = sTYPECODE_COLUMN Then
            
            ' Photo columns
            If objRecProfColumn.IsPhoto Then
              If iColumnCount = piStartColumn Then
              
                ' Is the photo linked or embedded
                If objRecProfColumn.OLEType = OLE_EMBEDDED Then
                  
                  ' Get the value into a stream
                  If Not IsNull(mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value) Then
                    objOLEStream.Open
                    objOLEStream.Type = adTypeBinary
                    objOLEStream.Write mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value
                                                        
                    sPhotoSS = sPHOTOSTYLESET & CStr(pgrdGrid.Columns.Count)
                    pgrdGrid.StyleSets.Add sPhotoSS
                    pgrdGrid.StyleSets(sPhotoSS).Picture = LoadPictureFromStream(objOLEStream)
                    pgrdGrid.StyleSets(sPhotoSS).AlignmentPicture = 0
                    pgrdGrid.StyleSets(sPhotoSS).AlignmentText = 1
  
                    Set imgTemp.Picture = pgrdGrid.StyleSets(sPhotoSS).Picture
                    If (imgTemp.Height + PHOTOFRAME) > pgrdGrid.RowHeight Then
                      pgrdGrid.RowHeight = imgTemp.Height + PHOTOFRAME
                    End If
                    lngColWidth = imgTemp.Width + PHOTOFRAME
  
                    If objOLEStream.State = adStateOpen Then
                      objOLEStream.Close
                    End If
                    
                    sTemp = ""
                  End If

                Else
                  sTemp = FormatData(objRecProfColumn, mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value)
                
                  If sTemp <> "" Then
                    If (Dir(gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "", "\") & sTemp, vbDirectory) <> vbNullString) Then
                      sPhotoSS = sPHOTOSTYLESET & CStr(pgrdGrid.Columns.Count)
                      pgrdGrid.StyleSets.Add sPhotoSS
                      pgrdGrid.StyleSets(sPhotoSS).Picture = gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "", "\") & sTemp
                      pgrdGrid.StyleSets(sPhotoSS).AlignmentPicture = 0
                      pgrdGrid.StyleSets(sPhotoSS).AlignmentText = 1
  
                      Set imgTemp.Picture = LoadPicture(gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "", "\") & sTemp)
                      If (imgTemp.Height + PHOTOFRAME) > pgrdGrid.RowHeight Then
                        pgrdGrid.RowHeight = imgTemp.Height + PHOTOFRAME
                      End If
                      lngColWidth = imgTemp.Width + PHOTOFRAME
  
                      sTemp = ""
                    Else
                      lngColWidth = XGAP + WidthOfText(sTemp)
                      
                      .Bookmark = (.AddItemBookmark(iCount - 1))
                      .Columns(.Columns.Count - 1).Value = sTemp
                    End If
                  End If
                End If
              Else
                FormatOutput_AddDataToVerticalGrid pobjRecProfTable, _
                  grdOutput(pgrdGrid.Index + 1), _
                  iColumnCount, _
                  True, _
                  True
                Exit For
              End If
              
            ' OLE columns
            ElseIf objRecProfColumn.IsOLE Then
            
              ' Is the OLE linked or embedded
              If objRecProfColumn.OLEType = OLE_EMBEDDED Then
                
                ' Get the value into a stream
                If Not IsNull(mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value) Then
                  objOLEStream.Open
                  objOLEStream.Type = adTypeBinary
                  objOLEStream.Write mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value
                End If

                sTemp = LoadFileNameFromStream(objOLEStream, True, True)
                'sTemp = IIf(Len(sTemp) = 0, "Empty", sTemp)
                'NHRD16072004 Fault 8726 Suppress 'Empty' string
                sTemp = IIf(Len(sTemp) = 0, "", sTemp)
                
                If objOLEStream.State = adStateOpen Then
                  objOLEStream.Close
                End If
                
              Else
                sTemp = FormatData(objRecProfColumn, mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value)
              End If
          
              lngColWidth = XGAP + WidthOfText(sTemp)
            
              .Bookmark = (.AddItemBookmark(iCount - 1))
              .Columns(.Columns.Count - 1).Value = sTemp
            
            
            ' Any other type of column
            Else
              sTemp = FormatData(objRecProfColumn, mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value)
        
              lngColWidth = XGAP + WidthOfText(sTemp)
            
              .Bookmark = (.AddItemBookmark(iCount - 1))
              .Columns(.Columns.Count - 1).Value = sTemp
            End If
            
            If (lngColWidth > .Columns(.Columns.Count - 1).Width) Then
              .Columns(.Columns.Count - 1).Width = lngColWidth
            End If
            
            If pfFollowingOn Then
              If .Columns(.Columns.Count - 1).Width < grdOutput(pgrdGrid.Index - 1).Columns(grdOutput(pgrdGrid.Index - 1).Columns.Count - 1).Width Then
                .Columns(.Columns.Count - 1).Width = grdOutput(pgrdGrid.Index - 1).Columns(grdOutput(pgrdGrid.Index - 1).Columns.Count - 1).Width
              ElseIf .Columns(.Columns.Count - 1).Width > grdOutput(pgrdGrid.Index - 1).Columns(grdOutput(pgrdGrid.Index - 1).Columns.Count - 1).Width Then
                FormatPrecedingGridColumnWidth pgrdGrid.Index - 1, .Columns(.Columns.Count - 1).Width
              End If
            End If
            
            .Columns(sCOLUMN_BLANKIFZERO).Value = objRecProfColumn.BlankIfZero
            .Columns(sCOLUMN_USE1000SEPARATOR).Value = objRecProfColumn.ThousandSeparator
            .Columns(sCOLUMN_DECPLACES).Value = objRecProfColumn.DecPlaces
            
          End If
        End If
      End If
    Next objRecProfColumn
    Set objRecProfColumn = Nothing
  
    .Bookmark = (.AddItemBookmark(0))
    
    .Width = .Columns(.Columns.Count - 1).Left + .Columns(.Columns.Count - 1).Width
    .Height = (.RowHeight * .Rows - 1) + 15
  End With
  
  Set objOLEStream = Nothing
  
End Function

Private Function FormatOutput_AddDataToHorizontalGrid(pobjRecProfTable As clsRecordProfileTabDtl, _
  pgrdGrid As SSDBGrid)
  
  Dim sAddString As String
  Dim iHeadingCount As Integer
  Dim iColumnCount As Integer
  Dim objRecProfColumn As clsRecordProfileColDtl
  Dim sTemp As String
  Dim lngColWidth As Long
  Dim sPhotoSS As String
  Dim lngWidthDiff As Long
  Dim iCount As Integer
  Dim iCount2 As Integer
  Dim objOLEStream As New ADODB.Stream
  
  sAddString = ""
    
  iHeadingCount = 0
  iColumnCount = 0
  For Each objRecProfColumn In pobjRecProfTable.Columns
    lngColWidth = 0
    
    Select Case objRecProfColumn.ColType
      Case sTYPECODE_HEADING
        'JPD 20040930 Fault 9254
        If (iColumnCount > 0) And (iHeadingCount = 0) Then
          iHeadingCount = iHeadingCount + 1
        End If
        
        iHeadingCount = iHeadingCount + 1
        
      Case sTYPECODE_COLUMN, sTYPECODE_ID
        iColumnCount = iColumnCount + 1
                   
        ' Photo objects
        If objRecProfColumn.IsPhoto Then
          
          ' Is the photo linked or embedded
          If objRecProfColumn.OLEType = OLE_EMBEDDED Then
                  
            ' Get the value into a stream
            If Not IsNull(mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value) Then
              objOLEStream.Open
              objOLEStream.Type = adTypeBinary
              objOLEStream.Write mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value
                                                                                        
              sPhotoSS = sPHOTOSTYLESET & CStr(iColumnCount) & "_" & mrsResults.Fields(pobjRecProfTable.IDPosition).Value
              pgrdGrid.StyleSets.Add sPhotoSS
              pgrdGrid.StyleSets(sPhotoSS).Picture = LoadPictureFromStream(objOLEStream)
              pgrdGrid.StyleSets(sPhotoSS).AlignmentPicture = 0
              pgrdGrid.StyleSets(sPhotoSS).AlignmentText = 1
                
              Set imgTemp.Picture = pgrdGrid.StyleSets(sPhotoSS).Picture
              If (imgTemp.Height + PHOTOFRAME) > pgrdGrid.RowHeight Then
                pgrdGrid.RowHeight = imgTemp.Height + PHOTOFRAME
              End If
              lngColWidth = imgTemp.Width + PHOTOFRAME
              
              If objOLEStream.State = adStateOpen Then
                objOLEStream.Close
              End If
              
              sTemp = ""
          
            End If
         
          Else
          
            sTemp = FormatData(objRecProfColumn, mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value)
            lngColWidth = WidthOfText(sTemp) + XGAP
          
            If sTemp <> "" Then
              If (Dir(gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "", "\") & sTemp, vbDirectory) <> vbNullString) Then
                sPhotoSS = sPHOTOSTYLESET & CStr(iColumnCount) & "_" & mrsResults.Fields(pobjRecProfTable.IDPosition).Value
                pgrdGrid.StyleSets.Add sPhotoSS
                pgrdGrid.StyleSets(sPhotoSS).Picture = gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "", "\") & sTemp
                pgrdGrid.StyleSets(sPhotoSS).AlignmentPicture = 0
                pgrdGrid.StyleSets(sPhotoSS).AlignmentText = 1
                  
                Set imgTemp.Picture = LoadPicture(gsPhotoPath & IIf(Right(gsPhotoPath, 1) = "\", "", "\") & sTemp)
                If (imgTemp.Height + PHOTOFRAME) > pgrdGrid.RowHeight Then
                  pgrdGrid.RowHeight = imgTemp.Height + PHOTOFRAME
                End If
                lngColWidth = imgTemp.Width + PHOTOFRAME
                  
                sTemp = ""
              End If
            End If
          End If
        
        ' OLE objects
        ElseIf objRecProfColumn.IsOLE Then
        
          ' Is the photo linked or embedded
          If objRecProfColumn.OLEType = OLE_EMBEDDED Then
                  
            ' Get the value into a stream
            If Not IsNull(mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value) Then
              objOLEStream.Open
              objOLEStream.Type = adTypeBinary
              objOLEStream.Write mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value
            End If
                
            sTemp = LoadFileNameFromStream(objOLEStream, True, True)
            'sTemp = IIf(Len(sTemp) = 0, "Empty", sTemp)
            'NHRD16072004 Fault 8726 Suppress 'Empty' string
            sTemp = IIf(Len(sTemp) = 0, "", sTemp)
            If objOLEStream.State = adStateOpen Then
              objOLEStream.Close
            End If
                   
            lngColWidth = WidthOfText(sTemp) + XGAP
                   
          End If
        Else
        
          sTemp = FormatData(objRecProfColumn, mrsResults.Fields(objRecProfColumn.PositionInRecordset).Value)
          lngColWidth = WidthOfText(sTemp) + XGAP
        
        End If
          
        If objRecProfColumn.Displayed Then
          If lngColWidth > objRecProfColumn.GridColumnWidth Then
            lngWidthDiff = lngColWidth - objRecProfColumn.GridColumnWidth
            objRecProfColumn.GridColumnWidth = lngColWidth

            If (pobjRecProfTable.HasHeadings) Then
              pgrdGrid.Groups(IIf(iHeadingCount > 0, iHeadingCount - 1, iHeadingCount)).Width = pgrdGrid.Groups(IIf(iHeadingCount > 0, iHeadingCount - 1, iHeadingCount)).Width + lngWidthDiff
            End If

            pgrdGrid.Columns(iColumnCount - 1).Width = objRecProfColumn.GridColumnWidth
          End If
        End If
        
        sAddString = sAddString & sTemp & vbTab

      Case Else
        iColumnCount = iColumnCount + 1
        sAddString = sAddString & "" & vbTab
    End Select
  Next objRecProfColumn
  Set objRecProfColumn = Nothing
    
  Set objOLEStream = Nothing

  With pgrdGrid
    If pobjRecProfTable.HasHeadings Then
      iCount = -1
      iCount2 = 0
        
      For Each objRecProfColumn In pobjRecProfTable.Columns
        If objRecProfColumn.Displayed Then
          If objRecProfColumn.ColType = sTYPECODE_HEADING Then
            iCount = iCount + 1
            iCount2 = 0
          Else
            If iCount < 0 Then
              iCount = 0
            End If
            
            If iCount2 = (.Groups(iCount).Columns.Count - 1) Then
              .Groups(iCount).Columns(iCount2).Width = .Groups(iCount).Width - _
                (.Groups(iCount).Columns(iCount2).Left - .Groups(iCount).Left)
            Else
              .Groups(iCount).Columns(iCount2).Width = objRecProfColumn.GridColumnWidth + 5
            End If
              
            If .Groups(iCount).Columns(iCount2).Width < objRecProfColumn.GridColumnWidth + 5 Then
              .Groups(iCount).Width = .Groups(iCount).Width + (objRecProfColumn.GridColumnWidth + 5 - .Groups(iCount).Columns(iCount2).Width)
              .Groups(iCount).Columns(iCount2).Width = objRecProfColumn.GridColumnWidth + 5
            End If
              
            iCount2 = iCount2 + 1
          End If
        End If
      Next objRecProfColumn
      Set objRecProfColumn = Nothing
    End If
  
    .AddItem sAddString
  
    .Height = (.RowHeight * .Rows) + _
      STANDARDROWHEIGHT + _
      IIf(pobjRecProfTable.HasHeadings, STANDARDROWHEIGHT + 25, 5) + 30
      
    If (pobjRecProfTable.HasHeadings) Then
      .Groups(.Groups.Count - 1).Width = .Groups(.Groups.Count - 1).Columns(.Groups(.Groups.Count - 1).Columns.Count - 1).Left _
        - .Groups(.Groups.Count - 1).Left _
        + .Groups(.Groups.Count - 1).Columns(.Groups(.Groups.Count - 1).Columns.Count - 1).Width
      
      .Width = .Groups(.Groups.Count - 1).Left + .Groups(.Groups.Count - 1).Width
    Else
      .Width = .Columns(.Columns.Count - 1).Left + .Columns(.Columns.Count - 1).Width
    End If
  End With
  
End Function

  

Private Function FormatOutput_AddGrid2(pobjRecProfTable As clsRecordProfileTabDtl, _
  piStartColumn As Integer, _
  plngFollowsGrid As Long, _
  pfAddingPhotoGrid As Boolean, _
  plngPicBoxIndex As Long, _
  pfNoRecords As Boolean, _
  plngRecordID As Long, _
  plngParentRecordID As Long) As Boolean

  Dim grdTemp As SSDBGrid
  Dim sCaption As String
  Dim lngIndex As Long
  Dim lngLoop As Long
  
  ' Locate the correct position for the new grid (and label)
  lngIndex = 0
  
  If plngFollowsGrid > 0 Then
    For lngLoop = 1 To UBound(matypInfo)
      If (matypInfo(lngLoop).IsGrid) And _
        (matypInfo(lngLoop).Index = plngFollowsGrid) Then
        lngIndex = lngLoop
        Exit For
      End If
    Next lngLoop
  Else
    If (pobjRecProfTable.Relationship = "CHILD") And _
      (pobjRecProfTable.Generation > 1) Then
  
      For lngLoop = 1 To UBound(matypInfo)
        If (matypInfo(lngLoop).IsGrid) And _
          (matypInfo(lngLoop).TableID = pobjRecProfTable.RelatedTableID) And _
          (InStr(matypInfo(lngLoop).RecordIDs, "," & CStr(plngParentRecordID) & ",") > 0) Then
          lngIndex = lngLoop
        End If
      Next lngLoop
  
      If (lngIndex > 0) Then
        For lngLoop = lngIndex + 1 To UBound(matypInfo)
          If ((Not matypInfo(lngLoop).IsGrid) Or (matypInfo(lngLoop).FollowsGrid = 0)) Then
            Exit For
          End If
  
          lngIndex = lngLoop
        Next lngLoop
      End If
    End If
  End If

  If piStartColumn = 0 Then
    ' Load the label for the table grid.
   
    Load lblCaption(lblCaption.Count)
    With lblCaption(lblCaption.Count - 1)
      Set .Container = picOutput(plngPicBoxIndex)
      .Visible = True
      
      sCaption = Replace(pobjRecProfTable.TableName, "_", " ")
      If (pobjRecProfTable.Generation > 0) And _
        (Not mblnSuppressTableRelationshipTitles) Then
        
        sCaption = sCaption & _
          " (" & LCase(pobjRecProfTable.Relationship) & _
          " of '" & Replace(mobjDefinition.Item(pobjRecProfTable.RelatedTableID).TableName, "_", " ") & "')"
      End If
      
      If pfNoRecords Then
        sCaption = sCaption & " - no records"
        pobjRecProfTable.GridIndex = 0
        ReDim Preserve malngEmptyTableCaptions(UBound(malngEmptyTableCaptions) + 1)
        malngEmptyTableCaptions(UBound(malngEmptyTableCaptions)) = .Index
      End If
      
      .Caption = sCaption
      .Tag = CStr(pobjRecProfTable.TableID)
      
      .Visible = (Not pfNoRecords) Or _
        (Not mblnSuppressEmptyRelatedTableTitles)
    End With
  
    ReDim Preserve matypInfo(UBound(matypInfo) + 1)
    If lngIndex > 0 Then
      lngIndex = lngIndex + 1
      
      For lngLoop = UBound(matypInfo) To lngIndex Step -1
        With matypInfo(lngLoop)
          .IsGrid = matypInfo(lngLoop - 1).IsGrid
          .Index = matypInfo(lngLoop - 1).Index
          .RecordIDs = matypInfo(lngLoop - 1).RecordIDs
          .TableID = matypInfo(lngLoop - 1).TableID
          .FollowsGrid = matypInfo(lngLoop - 1).FollowsGrid
          .AssociatedControlIndex = matypInfo(lngLoop - 1).AssociatedControlIndex
        End With
      Next lngLoop
    Else
      lngIndex = UBound(matypInfo)
    End If

    With matypInfo(lngIndex)
      .IsGrid = False
      .Index = lblCaption.Count - 1
      .RecordIDs = "," & CStr(plngRecordID) & ","
      .TableID = pobjRecProfTable.TableID
      .FollowsGrid = 0
      .AssociatedControlIndex = IIf(pfNoRecords, 0, grdOutput.Count)
    End With
  End If
                
  ' Load the grid for the table (if the table has records)
  If Not pfNoRecords Then
    Load grdOutput(grdOutput.Count)
    Set grdTemp = grdOutput(grdOutput.Count - 1)

    Set grdTemp.Container = picOutput(plngPicBoxIndex)
    grdTemp.Visible = True
    grdTemp.RowHeight = STANDARDROWHEIGHT
    grdTemp.Tag = CStr(pobjRecProfTable.TableID)
    
    ReDim Preserve matypInfo(UBound(matypInfo) + 1)
    If lngIndex > 0 Then
      lngIndex = lngIndex + 1
      For lngLoop = UBound(matypInfo) To lngIndex Step -1
        With matypInfo(lngLoop)
          .IsGrid = matypInfo(lngLoop - 1).IsGrid
          .Index = matypInfo(lngLoop - 1).Index
          .RecordIDs = matypInfo(lngLoop - 1).RecordIDs
          .TableID = matypInfo(lngLoop - 1).TableID
          .FollowsGrid = matypInfo(lngLoop - 1).FollowsGrid
          .AssociatedControlIndex = matypInfo(lngLoop - 1).AssociatedControlIndex
        End With
      Next lngLoop
    Else
      lngIndex = UBound(matypInfo)
    End If
    
    With matypInfo(lngIndex)
      .IsGrid = True
      .Index = grdOutput.Count - 1
      .RecordIDs = "," & CStr(plngRecordID) & ","
      .TableID = pobjRecProfTable.TableID
      .FollowsGrid = plngFollowsGrid
      .AssociatedControlIndex = lblCaption.Count - 1
    End With
                  
    If piStartColumn = 0 Then
      pobjRecProfTable.GridIndex = grdOutput.Count - 1
    End If
    
    If pobjRecProfTable.Orientation = giVERTICAL Then
      If pfAddingPhotoGrid Then
        grdTemp.TagVariant = sCOLUMN_ISPHOTO
      End If
      
      FormatGrid_Vertical2 grdTemp, pobjRecProfTable, piStartColumn, pfAddingPhotoGrid, (plngFollowsGrid > 0), plngPicBoxIndex
    Else
      FormatGrid_Horizontal grdTemp, pobjRecProfTable
    End If
  
    Set grdTemp = Nothing
  End If

End Function

Private Function FormatPictureBox2(plngIndex As Long) As Boolean
  
  'MH20040625 Fault 8782
  'Need to make sure that the tabindex is correct (even though tabstop = false)
  'as the tabindex will be used in output options.
  
  
  ' Position the controls contained in the given picturebox.
  ' Size the picture box to fit its contents.
  Dim lngIndex As Long
  Dim lngCurrentY As Long
  Dim typInfo As PROFILEINFO
  Dim lngTabIndex As Long       'MH20040625
  
  lngCurrentY = 0
  lngTabIndex = 0               'MH20040625
  
  ' Create an array of the controls and the order they appear in the picturebox.
  ' Column 1 = page no.
  ' Column 2 = control type (0 = label, 1 = grid)
  ' Column 3 = control index

  For lngIndex = 1 To UBound(matypInfo)
    typInfo = matypInfo(lngIndex)
    
    If typInfo.IsGrid Then
      With grdOutput(typInfo.Index)
        If .Container.Index = plngIndex Then
          .Top = lngCurrentY + IIf(typInfo.FollowsGrid > 0, 0, GAPABOVEGRID) - RECPROFFOLLOWONCORRECTION
          lngCurrentY = .Top + .Height
          
          'MH20040625
          If lngTabIndex = 0 Then
            lngTabIndex = .TabIndex
          Else
            lngTabIndex = lngTabIndex + 1
            .TabIndex = lngTabIndex
          End If
          
          malngControlOrder(1, lngTabIndex) = plngIndex
          malngControlOrder(2, lngTabIndex) = 1
          malngControlOrder(3, lngTabIndex) = .Index
        End If
      End With
    Else
      With lblCaption(typInfo.Index)
        If .Container.Index = plngIndex Then
          If .Visible Then
            .Top = lngCurrentY + GAPABOVELABEL
            lngCurrentY = .Top + .Height
          
            'MH20040625
            If lngTabIndex = 0 Then
              lngTabIndex = .TabIndex
            Else
              lngTabIndex = lngTabIndex + 1
              .TabIndex = lngTabIndex
            End If
          
            malngControlOrder(1, lngTabIndex) = plngIndex
            malngControlOrder(2, lngTabIndex) = 0
            malngControlOrder(3, lngTabIndex) = .Index
          End If
        End If
      End With
    End If
  Next lngIndex
  
  picOutput(plngIndex).Height = lngCurrentY + GAPABOVEGRID + 200

End Function



Public Function OutputReport(pfPrompt As Boolean) As Boolean
  
  Dim objOutput As clsOutputRun
  Dim ctlPictureBox As PictureBox
  Dim fOK As Boolean
  
  fOK = True
  Set objOutput = New clsOutputRun

  objOutput.HeaderRows = 2
  objOutput.HeaderCols = 1

  objOutput.ShowFormats True, False, True, True, True, False, False

  If objOutput.SetOptions _
    (pfPrompt, mlngOutputFormat, mblnOutputScreen, _
    mblnOutputPrinter, mstrOutputPrinterName, _
    mblnOutputSave, mlngOutputSaveExisting, _
    mblnOutputEmail, mlngOutputEmailAddr, mstrOutputEmailSubject, _
    mstrOutputEmailAttachAs, mstrOutputFileName) Then

'''    objOutput.SizeColumnsIndependently = False
    
    If objOutput.GetFile Then

      If Not gblnBatchMode Then
        objOutput.OpenProgress "Record Profile", Me.Caption, (picOutput.Count + 1)
      End If

      For Each ctlPictureBox In picOutput
        If fOK And (ctlPictureBox.Index > 0) Then
          
          'objOutput.AddPage Me.Caption, IIf(cboPage.Enabled, cboPage.List(ctlPictureBox.Index), "")
          If cboPage.Enabled Then
            objOutput.AddPage Me.Caption, cboPage.List(ctlPictureBox.Index - 1)
          Else
            objOutput.AddPage Me.Caption, msRecordProfileName & "(" & ctlPictureBox.Index & ")"
          End If
          fOK = objOutput.RecordProfilePage(Me, ctlPictureBox.Index)
    
          If Not fOK Then
            mstrErrorMessage = objOutput.ErrorMessage
            If Len(mstrErrorMessage) = 0 Then mstrErrorMessage = "Error"
            fOK = False
            Exit For
          End If

          If gobjProgress.Cancelled Then
            mstrErrorMessage = "Cancelled by user."
            fOK = False
            Exit For
          End If
          
          If fOK And Not gblnBatchMode Then
            gobjProgress.UpdateProgress gblnBatchMode
          End If
        End If
      Next ctlPictureBox
      Set ctlPictureBox = Nothing

      If fOK Then
        objOutput.ResetColumns
        objOutput.ResetStyles
        objOutput.ResetMerges
      End If
      If Not gobjProgress.Cancelled Then
        objOutput.Complete
      End If

    End If

    'JPD 20050610 Fault 8957 & Fault 9593
    If fOK Then
      mblnUserCancelled = objOutput.UserCancelled
      mstrErrorMessage = objOutput.ErrorMessage
      fOK = (mstrErrorMessage = vbNullString)
    End If

  Else
    'mblnUserCancelled = objOutput.UserCancelled
    pfPrompt = (pfPrompt And Not objOutput.UserCancelled)
    mstrErrorMessage = objOutput.ErrorMessage
    fOK = (mstrErrorMessage = vbNullString)

  End If

  If pfPrompt Then
    gobjProgress.CloseProgress
    If fOK Then
      COAMsgBox "Record Profile: '" & msRecordProfileName & "' output complete.", _
          vbInformation, "Record Profile"
    Else
      COAMsgBox "Record Profile: '" & msRecordProfileName & "' output failed." & vbCrLf & vbCrLf & mstrErrorMessage, _
          vbExclamation, "Record Profile"
    End If
  End If

  If Not objOutput Is Nothing Then objOutput.ClearUp
  Set objOutput = Nothing

  OutputReport = fOK

End Function

Public Property Get ControlOrder() As Variant
  ControlOrder = malngControlOrder
  
End Property

Public Property Get ErrorMessage() As String
  ErrorMessage = mstrErrorMessage
  
End Property


Private Function FormatOutput2() As Boolean
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim sAddString As String
  Dim objRecProfTable As clsRecordProfileTabDtl
  Dim objLastRecProfTable As clsRecordProfileTabDtl
  Dim grdTemp As SSDBGrid
  Dim typInfo As PROFILEINFO
  Dim iArrayIndex As Integer
  Dim fAdded As Boolean
  Dim lngRecordID As Long
  Dim fHasRecDesc As Boolean
  Dim lngLoop As Long
  Dim lngLoop2 As Long
  Dim iLoop As Integer
  Dim lngIndex As Long
  Dim fFound As Boolean
  Dim lngGridIndex As Long
  Dim lngLastTableDone As Long
  Dim avRecordsDone() As Variant
  Dim sRelatedRecordsDone As String
  Dim lngLastTableOrder As Long
  Dim objColumn As clsRecordProfileColDtl
  Dim lngRelatedRecordID As Long
  Dim alngChildRecords() As Long
  Dim fUnwantedRecord As Boolean
  
  fOK = True
  fHasRecDesc = (mobjDefinition.BaseTable.RecordDescriptionID > 0)
  
  ReDim avRecordsDone(4, 0)
  ' Column 1 = tableID
  ' Column 2 = record ID
  ' Column 3 = parent tables with no matching records
  ' Column 4 = base tables record ID
  
  ReDim alngChildRecords(3, 0)
  ' Column 1 = table ID
  ' Column 2 = record count
  ' Column 3 = parent record ID
  
  ' Initialise the progress bar
  With gobjProgress
    If Not gblnBatchMode Then
      'MH20030708 If you select no preview (ie straight to output) then the cancel button
      'becomes disabled.  Add one to the maxvalue to keep it enabled.
      '.Bar1MaxValue = mrsResults.RecordCount + 1
      .Bar1MaxValue = mrsResults.RecordCount + 1
    Else
      .ResetBar2
      .Bar2MaxValue = mrsResults.RecordCount + 1
      .Bar2Caption = "Record Profile : " & msRecordProfileName
    End If
  End With
  
  lngLastTableOrder = 1
  sRelatedRecordsDone = ","

  mrsResults.MoveFirst
  Do While (Not mrsResults.EOF) And (Not gobjProgress.Cancelled)
    ' Update the progress bar
    gobjProgress.UpdateProgress gblnBatchMode

    fUnwantedRecord = False

    Set objRecProfTable = mobjDefinition.ItemByPosition(mrsResults!ASRSysTableOrder + 1)
    lngRecordID = IIf(IsNull(mrsResults.Fields(objRecProfTable.IDPosition).Value), 0, mrsResults.Fields(objRecProfTable.IDPosition).Value)
    
    ReDim Preserve avRecordsDone(4, UBound(avRecordsDone, 2) + 1)
    avRecordsDone(1, UBound(avRecordsDone, 2)) = objRecProfTable.TableID
    avRecordsDone(2, UBound(avRecordsDone, 2)) = lngRecordID
    avRecordsDone(3, UBound(avRecordsDone, 2)) = ","
    avRecordsDone(4, UBound(avRecordsDone, 2)) = mrsResults.Fields(mobjDefinition.BaseTable.IDPosition).Value
    For Each objColumn In objRecProfTable.Columns
      If (objColumn.ColType = sTYPECODE_ID) And (objColumn.ColumnName <> "ID") Then
        If IsNull(mrsResults.Fields(objColumn.PositionInRecordset).Value) Then
          avRecordsDone(3, UBound(avRecordsDone, 2)) = avRecordsDone(3, UBound(avRecordsDone, 2)) & _
            Mid(objColumn.ColumnName, 4) & ","
        End If
      End If
    Next objColumn
    Set objColumn = Nothing
    
    If mrsResults!ASRSysTableOrder = 0 Then
      ' New base table record.

      DoEvents

      ' Create a new picturebox page.
      Load picOutput(picOutput.Count)
      With picOutput(picOutput.Count - 1)
        .Visible = False
        Set .Container = picContainer
        .Top = 0
        .Left = 0
        .Tag = CStr(lngRecordID)
      End With
      
      FormatOutput_AddGrid2 objRecProfTable, 0, 0, False, picOutput.Count - 1, False, 0, 0

      If objRecProfTable.GridIndex > 0 Then
        ' Load the data into the grid.
        Set grdTemp = grdOutput(objRecProfTable.GridIndex)

        If objRecProfTable.Orientation = giVERTICAL Then
          FormatOutput_AddDataToVerticalGrid objRecProfTable, grdTemp, 0, False, False
        Else
          FormatOutput_AddDataToHorizontalGrid objRecProfTable, grdTemp
        End If
      
        Set grdTemp = Nothing
      End If

      ' Add the record's record description to the dropdown list.
      If fHasRecDesc Then
        sAddString = EvaluateRecordDescription(mrsResults.Fields(objRecProfTable.IDPosition), objRecProfTable.RecordDescriptionID)
        If Len(Trim(sAddString)) = 0 Then
          sAddString = "<empty record description>"
        End If
        cboPage.AddItem sAddString
      End If
    Else
      ' Not the base table.
      For iLoop = lngLastTableOrder + 1 To mrsResults!ASRSysTableOrder
        Set objLastRecProfTable = mobjDefinition.ItemByPosition(iLoop)
        If objLastRecProfTable.Relationship = "PARENT" Then
          For lngLoop = 1 To UBound(avRecordsDone, 2)
            If InStr(avRecordsDone(3, lngLoop), "," & CStr(objLastRecProfTable.TableID) & ",") > 0 Then
              ' No related parent record.
              
              For lngLoop2 = 1 To (picOutput.Count - 1)
                If picOutput(lngLoop2).Tag = CStr(avRecordsDone(4, lngLoop)) Then
                  lngIndex = lngLoop2
                  Exit For
                End If
              Next lngLoop2
              
              FormatOutput_AddGrid2 objLastRecProfTable, 0, 0, False, lngIndex, True, 0, 0
            End If
          Next lngLoop
        Else
          For lngLoop = 1 To UBound(avRecordsDone, 2)
            If (avRecordsDone(1, lngLoop) = objLastRecProfTable.RelatedTableID) And _
              (InStr(sRelatedRecordsDone, "," & CStr(avRecordsDone(2, lngLoop)) & ",") = 0) Then
              ' No records in the current table for the related parent.
      
              For lngLoop2 = 1 To (picOutput.Count - 1)
                If picOutput(lngLoop2).Tag = CStr(avRecordsDone(4, lngLoop)) Then
                  lngIndex = lngLoop2
                  Exit For
                End If
              Next lngLoop2
              
              FormatOutput_AddGrid2 objLastRecProfTable, 0, 0, False, lngIndex, True, 0, CLng(avRecordsDone(2, lngLoop))
            End If
          Next lngLoop
        End If
        
        Set objLastRecProfTable = Nothing
        
        sRelatedRecordsDone = ","
      Next iLoop
        
      If lngLastTableOrder <> mrsResults!ASRSysTableOrder Then
        sRelatedRecordsDone = ","
        lngLastTableOrder = mrsResults!ASRSysTableOrder
      End If
      
      If objRecProfTable.Relationship = "CHILD" Then
        sRelatedRecordsDone = sRelatedRecordsDone & CStr(mrsResults.Fields(objRecProfTable.RelatedTableIDPosition)) & ","
      
        fFound = False
        For lngLoop = 1 To UBound(alngChildRecords, 2)
          If (alngChildRecords(1, lngLoop) = objRecProfTable.TableID) And _
            (alngChildRecords(2, lngLoop) = mrsResults.Fields(objRecProfTable.RelatedTableIDPosition).Value) Then
            
            fFound = True
            alngChildRecords(3, lngLoop) = alngChildRecords(3, lngLoop) + 1
            
            If (objRecProfTable.MaxRecords > 0) And _
              (alngChildRecords(3, lngLoop) > objRecProfTable.MaxRecords) Then
              fUnwantedRecord = True
            End If
            
            Exit For
          End If
        Next lngLoop
        
        If Not fFound Then
          ReDim Preserve alngChildRecords(3, UBound(alngChildRecords, 2) + 1)
          alngChildRecords(1, UBound(alngChildRecords, 2)) = objRecProfTable.TableID
          alngChildRecords(2, UBound(alngChildRecords, 2)) = mrsResults.Fields(objRecProfTable.RelatedTableIDPosition)
          alngChildRecords(3, UBound(alngChildRecords, 2)) = 1
        End If
      End If
        
      If Not fUnwantedRecord Then
        ' Determine which picturebox this record is to go in.
        For lngLoop = 1 To (picOutput.Count - 1)
          If picOutput(lngLoop).Tag = CStr(mrsResults.Fields(mobjDefinition.BaseTable.IDPosition)) Then
            lngIndex = lngLoop
            Exit For
          End If
        Next lngLoop
    
        If objRecProfTable.HasChildren Or _
          (objRecProfTable.Relationship = "PARENT") Then
          ' Table is parent OR has children so put all records in individual grids.
          FormatOutput_AddGrid2 objRecProfTable, 0, 0, False, lngIndex, False, _
            IIf(objRecProfTable.Relationship = "PARENT", 0, lngRecordID), _
            IIf(objRecProfTable.Relationship = "PARENT", 0, mrsResults.Fields(objRecProfTable.RelatedTableIDPosition))
        
          If objRecProfTable.GridIndex > 0 Then
            ' Load the data into the grid.
            Set grdTemp = grdOutput(objRecProfTable.GridIndex)
            
            ' Don't add a record if its already been added !
            If objRecProfTable.Orientation = giVERTICAL Then
              FormatOutput_AddDataToVerticalGrid objRecProfTable, grdTemp, 0, False, False
            Else
              FormatOutput_AddDataToHorizontalGrid objRecProfTable, grdTemp
            End If
          
            Set grdTemp = Nothing
          End If
        Else
          ' Table has no children so put all records in a single grid.
          
          ' Determine which grid this record is to go in.
          fFound = False
          For lngLoop = 1 To UBound(malngMultiRecordGrids, 2)
            If (malngMultiRecordGrids(1, lngLoop) = objRecProfTable.TableID) And _
              (malngMultiRecordGrids(2, lngLoop) = mrsResults.Fields(objRecProfTable.RelatedTableIDPosition).Value) Then
              
              fFound = True
              lngGridIndex = malngMultiRecordGrids(3, lngLoop)
              Exit For
            End If
          Next lngLoop
          
          If fFound Then
            objRecProfTable.GridIndex = lngGridIndex
          Else
            FormatOutput_AddGrid2 objRecProfTable, 0, 0, False, lngIndex, False, 0, _
              IIf(objRecProfTable.Relationship = "PARENT", 0, mrsResults.Fields(objRecProfTable.RelatedTableIDPosition))
  
            
            ReDim Preserve malngMultiRecordGrids(3, UBound(malngMultiRecordGrids, 2) + 1)
            malngMultiRecordGrids(1, UBound(malngMultiRecordGrids, 2)) = objRecProfTable.TableID
            malngMultiRecordGrids(2, UBound(malngMultiRecordGrids, 2)) = mrsResults.Fields(objRecProfTable.RelatedTableIDPosition).Value
            malngMultiRecordGrids(3, UBound(malngMultiRecordGrids, 2)) = objRecProfTable.GridIndex
          End If
        
          If objRecProfTable.GridIndex > 0 Then
            ' Load the data into the grid.
            Set grdTemp = grdOutput(objRecProfTable.GridIndex)
        
            If objRecProfTable.Orientation = giVERTICAL Then
              FormatOutput_AddDataToVerticalGrid objRecProfTable, grdTemp, 0, False, False
            Else
              FormatOutput_AddDataToHorizontalGrid objRecProfTable, grdTemp
            End If
          
            Set grdTemp = Nothing
          End If
        End If
      End If
    End If
    
    Set objRecProfTable = Nothing
    
    mrsResults.MoveNext
  Loop

  For iLoop = lngLastTableOrder + 1 To mobjDefinition.Count
    Set objLastRecProfTable = mobjDefinition.ItemByPosition(iLoop)
    If objLastRecProfTable.Relationship = "PARENT" Then
      For lngLoop = 1 To UBound(avRecordsDone, 2)
        If InStr(avRecordsDone(3, lngLoop), "," & CStr(objLastRecProfTable.TableID) & ",") > 0 Then
          ' No related parent record.
          
          For lngLoop2 = 1 To (picOutput.Count - 1)
            If picOutput(lngLoop2).Tag = CStr(avRecordsDone(4, lngLoop)) Then
              lngIndex = lngLoop2
              Exit For
            End If
          Next lngLoop2
          
          FormatOutput_AddGrid2 objLastRecProfTable, 0, 0, False, lngIndex, True, 0, 0
        End If
      Next lngLoop
    Else
      For lngLoop = 1 To UBound(avRecordsDone, 2)
        If (avRecordsDone(1, lngLoop) = objLastRecProfTable.RelatedTableID) And _
          (InStr(sRelatedRecordsDone, "," & CStr(avRecordsDone(2, lngLoop)) & ",") = 0) Then
          ' No records in the current table for the related parent.
  
          For lngLoop2 = 1 To (picOutput.Count - 1)
            If picOutput(lngLoop2).Tag = CStr(avRecordsDone(4, lngLoop)) Then
              lngIndex = lngLoop2
              Exit For
            End If
          Next lngLoop2
          
          FormatOutput_AddGrid2 objLastRecProfTable, 0, 0, False, lngIndex, True, 0, CLng(avRecordsDone(2, lngLoop))
        End If
      Next lngLoop
    End If
  
    Set objLastRecProfTable = Nothing
    
    sRelatedRecordsDone = ","
  Next iLoop

  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    fOK = False
  Else

    ReDim malngControlOrder(3, Me.Controls.Count)
    
    For lngLoop = 1 To (picOutput.Count - 1)
      FormatPictureBox2 (lngLoop)
    Next lngLoop

    FormatIndents True

    ' Update the page spinner.
    With asrPage
      .MinimumValue = IIf(picOutput.Count = 1, 0, 1)
      .MaximumValue = picOutput.Count - 1
      .Value = .MinimumValue
    End With

    With cboPage
      .Enabled = fHasRecDesc And (.ListCount > 0)

      If Not .Enabled Then
        .BackColor = vbButtonFace
        .AddItem "<no record description>"
      End If

      .ListIndex = 0
    End With
    asrPage_Change
  End If
  
TidyUpAndExit:
  FormatOutput2 = fOK
  Exit Function
  
ErrorTrap:
  mstrErrorMessage = Err.Description
  If Err.Number = 7 Then
    mstrErrorMessage = mstrErrorMessage & vbCrLf & "Try reducing the number of base table records in the Record Profile."
  End If
  fOK = False
  GoTo TidyUpAndExit
  
End Function





Private Function FormatData(pobjColumn As clsRecordProfileColDtl, pvData As Variant) As String
  Dim iCount As Integer
  Dim iDigitCount As Integer
  Dim sTemp As String
  Dim sTemp2 As String
  Dim sFormat As String
  Dim bBlankIfZero As Boolean
  Dim iDecimals As Integer
  Dim bThousandSeparator As Boolean
  Dim iSize As Integer
  
  ' Do the DP thing
  If pobjColumn.IsNumeric Then
    
    pvData = IIf(IsNull(pvData), 0, pvData)
    
    bBlankIfZero = pobjColumn.BlankIfZero
    iDecimals = pobjColumn.DecPlaces
    bThousandSeparator = pobjColumn.ThousandSeparator
    iSize = pobjColumn.Size
    sFormat = ""
    
    ' Format positive section of value
    For iCount = 1 To (iSize - iDecimals)
      If bThousandSeparator Then
        sFormat = IIf(iCount Mod 3 = 0 And (iCount <> (iSize - iDecimals)), ",#", "#") & sFormat
      Else
        sFormat = "#" & sFormat
      End If
    Next iCount
   
    ' Single digit format
    If Not bBlankIfZero Then
      If Len(sFormat) > 0 Then
        sFormat = Left(sFormat, Len(sFormat) - 1) & "0"
      End If
    End If
    
    ' Decimal places
    If pvData <> 0 Then
      sFormat = sFormat & IIf(iDecimals > 0, ".", "")
      For iCount = 1 To iDecimals
        sFormat = sFormat & IIf(bBlankIfZero, "#", "0")
      Next iCount
    End If

    ' Format the number
    pvData = Format(pvData, sFormat)

  End If

  
  ' Is it a boolean calculation ? If so, change to Y or N
  If pobjColumn.IsLogic Then
    If pvData = "True" Then pvData = "Y"
    If pvData = "False" Then pvData = "N"
  End If
  
  ' If its a date column, format it as dateformat
  If pobjColumn.IsDate Then
    pvData = Format(pvData, DateFormat)
  End If

  If Not IsNull(pvData) Then pvData = Replace(pvData, vbCrLf, " ")
  If Not IsNull(pvData) Then pvData = Replace(pvData, vbTab, " ")
  If IsNull(pvData) Then pvData = ""
  
  FormatData = CStr(pvData)
  
End Function

Private Sub FormatPrecedingGridColumnWidth(plngGridIndex As Long, _
  plngWidth As Long)

  Dim iLoop As Integer
  
  With grdOutput(plngGridIndex)
    .Columns(.Columns.Count - 1).Width = plngWidth
  End With
  
  For iLoop = 1 To UBound(matypInfo)
   If matypInfo(iLoop).Index = plngGridIndex Then
     If matypInfo(iLoop).FollowsGrid > 0 Then
        FormatPrecedingGridColumnWidth matypInfo(iLoop).FollowsGrid, plngWidth
     End If
     
     Exit For
   End If
  Next iLoop

End Sub

Private Sub FormatPrecedingGridHeadingWidth(plngGridIndex As Long, _
  plngWidth As Long)

  Dim iLoop As Integer
  
  With grdOutput(plngGridIndex)
    .Columns(sCOLUMN_HEADING).Width = plngWidth
  End With
  
  For iLoop = 1 To UBound(matypInfo)
   If matypInfo(iLoop).Index = plngGridIndex Then
     If matypInfo(iLoop).FollowsGrid > 0 Then
        FormatPrecedingGridHeadingWidth matypInfo(iLoop).FollowsGrid, plngWidth
     End If
     
     Exit For
   End If
  Next iLoop

End Sub


Public Function ShowResults() As Boolean

  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  fOK = True

  ' Set the loading flag
  mblnLoading = True

  'JPD 20030728 Fault 6408
  chkIndent.Enabled = (mobjDefinition.Count > 1)
  chkSuppressEmptyTableTitles.Enabled = (mobjDefinition.Count > 1)
  chkShowTableRelationshipTitles.Enabled = (mobjDefinition.Count > 1)

  If fOK Then fOK = FormatOutput2

  If fOK Then Form_Resize

  Screen.MousePointer = vbDefault

  mblnLoading = False

TidyUpAndExit:
  ShowResults = fOK
  
  Exit Function

ErrorTrap:
  mblnLoading = False
  fOK = False
  GoTo TidyUpAndExit

End Function


Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

'Public Property Let BatchMode(pgblnBatchMode As Boolean)
'  gblnBatchMode = pgblnBatchMode
'End Property

Private Sub asrPage_Change()
  Dim objTemp As PictureBox
  
  If cboPage.Enabled Then
    cboPage.ListIndex = (asrPage.Value - 1)
  End If
  
  For Each objTemp In picOutput
    objTemp.Visible = (objTemp.Index = asrPage.Value)
  Next objTemp
  Set objTemp = Nothing
  
  picOutput(asrPage.Value).Top = 0
  picOutput(asrPage.Value).Left = 0
  scrollHorizontal.Value = 0
  scrollVertical.Value = 0
  
  SetScrollBarValues
  
End Sub

Private Sub cboPage_Click()
  asrPage.Value = cboPage.ListIndex + 1

End Sub


Private Sub chkIndent_Click()
  FormatIndents False
  SetScrollBarValues
  
End Sub

Private Sub chkShowTableRelationshipTitles_Click()
  Dim lngLoop As Long
  Dim objRecProfTable As clsRecordProfileTabDtl
  Dim sCaption As String
  Dim lngLoop2 As Long
  
  For lngLoop = 1 To UBound(matypInfo)
    mblnSuppressTableRelationshipTitles = (chkShowTableRelationshipTitles.Value = vbUnchecked)
    
    If (Not matypInfo(lngLoop).IsGrid) Then
      Set objRecProfTable = mobjDefinition.Item(CStr(matypInfo(lngLoop).TableID))
      
      sCaption = Replace(objRecProfTable.TableName, "_", " ")
      If (objRecProfTable.Generation > 0) And _
        (Not mblnSuppressTableRelationshipTitles) Then
        
        sCaption = sCaption & _
          " (" & LCase(objRecProfTable.Relationship) & _
          " of '" & Replace(mobjDefinition.Item(objRecProfTable.RelatedTableID).TableName, "_", " ") & "')"
      End If
      
      Set objRecProfTable = Nothing
      
      If IsEmptyTableCaption(matypInfo(lngLoop).Index) Then
        sCaption = sCaption & " - no records"
      End If
      
      lblCaption(matypInfo(lngLoop).Index).Caption = sCaption
    End If
  Next lngLoop

  FormatIndents False
  
End Sub

Private Sub chkSuppressEmptyTableTitles_Click()
  Dim lngLoop As Long
  Dim ctlPictureBox As PictureBox
  
  If picOutput.Count > 1 Then
    mblnSuppressEmptyRelatedTableTitles = (chkSuppressEmptyTableTitles.Value = vbChecked)
    
    For lngLoop = 1 To UBound(malngEmptyTableCaptions)
      lblCaption(malngEmptyTableCaptions(lngLoop)).Visible = Not mblnSuppressEmptyRelatedTableTitles
    Next lngLoop
    
    ReDim malngControlOrder(3, Me.Controls.Count)
    
    For Each ctlPictureBox In picOutput
      If ctlPictureBox.Index > 0 Then
        FormatPictureBox2 ctlPictureBox.Index
      End If
    Next ctlPictureBox
    Set ctlPictureBox = Nothing
  
    picOutput(asrPage.Value).Top = 0
    picOutput(asrPage.Value).Left = 0
    scrollHorizontal.Value = 0
    scrollVertical.Value = 0
    
    SetScrollBarValues
  End If
  
End Sub




Private Sub cmdClose_Click()
  Unload Me

End Sub

Private Sub cmdOutput_Click()
  OutputReport True
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
  mblnLoading = True
  Me.Width = FORM_STARTWIDTH
  Me.Height = FORM_STARTHEIGHT
  mblnLoading = False

  scrollVertical.SmallChange = STANDARDROWHEIGHT
  scrollHorizontal.SmallChange = STANDARDCOLUMNWIDTH

  ReDim matypInfo(0)
  ReDim malngEmptyTableCaptions(0)
  ReDim malngMultiRecordGrids(3, 0)
  
  Dim lngMinWidth As Long
  lngMinWidth = (2 * fraRecord.Left) + cboPage.Left + CBOPAGE_MINWIDTH + asrPage.Left + 450
  Hook Me.hWnd, lngMinWidth, FORM_STARTHEIGHT
  
  ' Get rid of the icon off the form
  RemoveIcon Me
  
End Sub

Public Property Let RecordProfileName(psName As String)
  msRecordProfileName = psName
  Me.Caption = "Record Profile - " & psName
  ' Get rid of the icon off the form
  RemoveIcon Me
End Property



Public Property Let Results(prsResults As ADODB.Recordset)
  Set mrsResults = prsResults
  
End Property


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

  Dim lngMinWidth As Long
  
  If (Me.WindowState = vbMinimized) Or (mblnLoading) Then
    Exit Sub
  End If

'  lngMinWidth = (2 * fraRecord.Left) + cboPage.Left + CBOPAGE_MINWIDTH + asrPage.Left
'  If (Me.Width < lngMinWidth) Then
'    mblnLoading = True
'    Me.Width = lngMinWidth
'    mblnLoading = False
'  End If
'
'  If (Me.Height < FORM_MINHEIGHT) Then
'    mblnLoading = True
'    Me.Height = FORM_MINHEIGHT
'    mblnLoading = False
'  End If

  With fraRecord
    .Width = Me.ScaleWidth - (2 * .Left)
    cboPage.Width = .Width - cboPage.Left - asrPage.Left
  End With
  
  With fraButtons
    .Width = fraRecord.Width
    .Top = Me.ScaleHeight - .Height
    cmdClose.Left = .Width - cmdClose.Width
    cmdOutput.Left = cmdClose.Left - cmdOutput.Width - chkIndent.Left
  End With
  
  With picContainer
    .Width = fraRecord.Width - scrollVertical.Width
    .Height = fraButtons.Top - .Top - scrollHorizontal.Height - (.Top - (fraRecord.Top + fraRecord.Height))
    
    scrollVertical.Left = .Left + .Width
    scrollVertical.Height = .Height

    scrollHorizontal.Top = .Top + .Height
    scrollHorizontal.Width = .Width
  End With

  SetScrollBarValues

  If (scrollVertical.Value = scrollVertical.Max) Then
    scrollVertical_Change
  End If

  If (scrollHorizontal.Value = scrollHorizontal.Max) Then
    scrollHorizontal_Change
  End If

End Sub


Private Function SetScrollBarValues() As Boolean
  On Error GoTo ErrorTrap
  
  Dim picTemp As PictureBox
  Dim lngTemp As Long
  Dim lngMax As Long
  
  Set picTemp = picOutput(asrPage.Value)
  
  With scrollVertical
    If (picTemp.Height <= picContainer.Height) Then
      .Value = 0
      .Enabled = False
    Else
      .Enabled = True
  
      If (picTemp.Height - picContainer.Height) > SCROLLMAX Then
        lngMax = SCROLLMAX
        mdblVerticalScrollRatio = (picTemp.Height - picContainer.Height) / SCROLLMAX
      Else
        lngMax = picTemp.Height - picContainer.Height
        mdblVerticalScrollRatio = 1
      End If
    
      .Max = lngMax
      
      If lngMax > (picContainer.Height / mdblVerticalScrollRatio) Then
        .LargeChange = CInt(picContainer.Height / mdblVerticalScrollRatio)
      Else
        .LargeChange = CInt(lngMax * 9 / 10)
      End If
    End If
  End With
  
  With scrollHorizontal
    If (picTemp.Width <= picContainer.Width) Then
      .Value = 0
      .Enabled = False
    Else
      .Enabled = True
  
      If (picTemp.Width - picContainer.Width) > SCROLLMAX Then
        lngMax = SCROLLMAX
        mdblHorizontalScrollRatio = (picTemp.Width - picContainer.Width) / SCROLLMAX
      Else
        lngMax = picTemp.Width - picContainer.Width
        mdblHorizontalScrollRatio = 1
      End If
  
      .Max = lngMax
      
      If lngMax > (picContainer.Width / mdblHorizontalScrollRatio) Then
        .LargeChange = CInt(picContainer.Width / mdblHorizontalScrollRatio)
      Else
        .LargeChange = CInt(lngMax * 9 / 10)
      End If
      
      lngTemp = (scrollHorizontal.Max * 0.01)

    End If
  End With
  
  Set picTemp = Nothing
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function


Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

Private Sub grdOutput_RowLoaded(Index As Integer, ByVal Bookmark As Variant)
  ' Apply a styleset as required
  Dim sTag As String
  Dim sTableID As String
  Dim objTable As clsRecordProfileTabDtl
  Dim sTemp As String
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  
  sTableID = grdOutput(Index).Tag
  
  Set objTable = mobjDefinition.Item(sTableID)
  
  If objTable.Orientation = giVERTICAL Then
    If grdOutput(Index).Columns(sCOLUMN_ISHEADING).CellText(Bookmark) = "1" Then
      For iLoop = 0 To grdOutput(Index).Columns.Count - 1
        grdOutput(Index).Columns(iLoop).CellStyleSet "Heading"
      Next iLoop
    End If
    
    If grdOutput(Index).TagVariant = sCOLUMN_ISPHOTO Then
      For iLoop = 0 To grdOutput(Index).Columns.Count - 1
        If grdOutput(Index).Columns(iLoop).Visible Then
          sTemp = sPHOTOSTYLESET & CStr(iLoop + 1)
          
          For iLoop2 = 0 To grdOutput(Index).StyleSets.Count - 1
            If grdOutput(Index).StyleSets(iLoop2).Name = sTemp Then
              grdOutput(Index).Columns(iLoop).CellStyleSet sTemp
              Exit For
            End If
          Next iLoop2
        End If
      Next iLoop
    End If
  Else
    For iLoop = 0 To grdOutput(Index).Columns.Count - 1
      If grdOutput(Index).Columns(iLoop).TagVariant = sCOLUMN_ISPHOTO Then
        sTemp = sPHOTOSTYLESET & CStr(iLoop + 1) & "_" & grdOutput(Index).Columns(CStr(objTable.IDPosition)).CellText(Bookmark)
        
        For iLoop2 = 0 To grdOutput(Index).StyleSets.Count - 1
          If grdOutput(Index).StyleSets(iLoop2).Name = sTemp Then
            grdOutput(Index).Columns(iLoop).CellStyleSet sTemp
            Exit For
          End If
        Next iLoop2
      End If
    Next iLoop
  End If
  
  Set objTable = Nothing
  
End Sub


Private Sub lblColumnSizingLabel_Click()
  ' This control is used when the record profile results are
  ' output to Excel.
  
End Sub

Private Sub scrollHorizontal_Change()
  picOutput(asrPage.Value).Left = (CLng(scrollHorizontal.Value) * -1) * mdblHorizontalScrollRatio

End Sub

Private Sub scrollVertical_Change()
  picOutput(asrPage.Value).Top = (CLng(scrollVertical.Value) * -1) * mdblVerticalScrollRatio

End Sub



Public Property Get Definition() As clsRecordProfileTabDtls
  Set Definition = mobjDefinition
  
End Property

Public Property Set Definition(ByVal pobjNewValue As clsRecordProfileTabDtls)
  Set mobjDefinition = pobjNewValue
  
End Property

Private Function WidthOfText(psText As String) As Long
  ' Returns the length of the given text when displayed on the form.
  ' NB. We split the string into blocks of size BLOCKLENGTH to avoid
  ' the over flow error described in MSDN article Q298825.
  
  ' NB (2). Capped the return size otherwise grid hangs if data is very large (multiline columns for example)
  '         data is still stored in its entirty so output still work.
  
  On Error GoTo ErrorTrap
  
  Dim sText As String
  Dim lngTextWidth As Long
  Const BLOCKLENGTH = 500

  lngTextWidth = 0
  sText = psText
  
  Do
    lngTextWidth = lngTextWidth + _
      Me.TextWidth(Left(sText, BLOCKLENGTH))
    
    sText = Mid(sText, BLOCKLENGTH + 1)
  Loop While Len(sText) > 0
  
TidyUpAndExit:
  WidthOfText = Minimum(CSng(lngTextWidth), 40000)
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function

