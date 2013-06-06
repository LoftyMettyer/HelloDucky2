VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{051CE3FC-5250-4486-9533-4E0723733DFA}#1.0#0"; "coa_colourpicker.ocx"
Begin VB.Form frmWorkflowWFItemProps 
   Caption         =   "Control Properties"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5072
   Icon            =   "frmWorkflowWFItemProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   4590
   Visible         =   0   'False
   Begin COAColourPicker.COA_ColourPicker colPickDlg 
      Left            =   720
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   4785
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   ""
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid ssGridProperties 
      Height          =   3720
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   4590
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   4
      stylesets.count =   15
      stylesets(0).Name=   "ssetActiveRowBold"
      stylesets(0).ForeColor=   -2147483634
      stylesets(0).BackColor=   -2147483635
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmWorkflowWFItemProps.frx":000C
      stylesets(1).Name=   "ssetBackColorEven"
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "frmWorkflowWFItemProps.frx":0028
      stylesets(2).Name=   "ssetBackColorOdd"
      stylesets(2).HasFont=   -1  'True
      BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(2).Picture=   "frmWorkflowWFItemProps.frx":0044
      stylesets(3).Name=   "ssetDormantRowBold"
      stylesets(3).ForeColor=   -2147483630
      stylesets(3).BackColor=   -2147483643
      stylesets(3).HasFont=   -1  'True
      BeginProperty stylesets(3).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(3).Picture=   "frmWorkflowWFItemProps.frx":0060
      stylesets(4).Name=   "ssetEnabled"
      stylesets(4).ForeColor=   -2147483640
      stylesets(4).BackColor=   -2147483643
      stylesets(4).HasFont=   -1  'True
      BeginProperty stylesets(4).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(4).Picture=   "frmWorkflowWFItemProps.frx":007C
      stylesets(5).Name=   "ssetForeColorHighlight"
      stylesets(5).HasFont=   -1  'True
      BeginProperty stylesets(5).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(5).Picture=   "frmWorkflowWFItemProps.frx":0098
      stylesets(6).Name=   "ssetBackColorValue"
      stylesets(6).HasFont=   -1  'True
      BeginProperty stylesets(6).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(6).Picture=   "frmWorkflowWFItemProps.frx":00B4
      stylesets(7).Name=   "ssetForeColorOdd"
      stylesets(7).HasFont=   -1  'True
      BeginProperty stylesets(7).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(7).Picture=   "frmWorkflowWFItemProps.frx":00D0
      stylesets(8).Name=   "ssetForeColorValue"
      stylesets(8).HasFont=   -1  'True
      BeginProperty stylesets(8).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(8).Picture=   "frmWorkflowWFItemProps.frx":00EC
      stylesets(9).Name=   "ssetHeaderBackColor"
      stylesets(9).HasFont=   -1  'True
      BeginProperty stylesets(9).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(9).Picture=   "frmWorkflowWFItemProps.frx":0108
      stylesets(10).Name=   "ssetBackColorHighlight"
      stylesets(10).HasFont=   -1  'True
      BeginProperty stylesets(10).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(10).Picture=   "frmWorkflowWFItemProps.frx":0124
      stylesets(11).Name=   "ssetActiveRow"
      stylesets(11).ForeColor=   -2147483634
      stylesets(11).BackColor=   -2147483635
      stylesets(11).HasFont=   -1  'True
      BeginProperty stylesets(11).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(11).Picture=   "frmWorkflowWFItemProps.frx":0140
      stylesets(12).Name=   "ssetDisabled"
      stylesets(12).ForeColor=   -2147483631
      stylesets(12).BackColor=   -2147483633
      stylesets(12).HasFont=   -1  'True
      BeginProperty stylesets(12).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(12).Picture=   "frmWorkflowWFItemProps.frx":015C
      stylesets(13).Name=   "ssetDormantRow"
      stylesets(13).ForeColor=   -2147483630
      stylesets(13).BackColor=   -2147483643
      stylesets(13).HasFont=   -1  'True
      BeginProperty stylesets(13).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(13).Picture=   "frmWorkflowWFItemProps.frx":0178
      stylesets(14).Name=   "ssetForeColorEven"
      stylesets(14).HasFont=   -1  'True
      BeginProperty stylesets(14).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(14).Picture=   "frmWorkflowWFItemProps.frx":0194
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      RowNavigation   =   3
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   4683
      Columns(0).Caption=   "Property"
      Columns(0).Name =   "colProperties"
      Columns(0).AllowSizing=   0   'False
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(0).HasBackColor=   -1  'True
      Columns(0).BackColor=   -2147483643
      Columns(1).Width=   3200
      Columns(1).Caption=   "Value"
      Columns(1).Name =   "colValues"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(1).Style=   3
      Columns(1).BackColor=   16777215
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Tag"
      Columns(2).Name =   "colTag"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "colDisabled"
      Columns(3).Name =   "colDisabled"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      UseDefaults     =   0   'False
      _ExtentX        =   8096
      _ExtentY        =   6562
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog comDlgBox 
      Left            =   130
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Verdana"
   End
   Begin VB.Label lblSizeTester 
      Caption         =   "<Size Tester>"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   1665
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmWorkflowWFItemProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Properties.
Private mfrmWebForm As frmWorkflowWFDesigner

Private miCXFrame As Integer
Private miCYFrame As Integer
Private miCXBorder As Integer
Private miCYBorder As Integer
Private miAlignment As Integer
Private miBackStyle As ASRBackStyleConstants

' Modular level variables
Private msWFIdentifier As String
Private msCaption As String
Private mlngDescriptionExprID As Long
Private mfDescriptionHasWorkflowName As Boolean
Private mfDescriptionHasElementCaption As Boolean
Private msTabCaption As String

Private mlngWFDatabaseRecord As WorkflowRecordSelectorTypes
Private msWFWebForm As String
Private msWFRecordSelector As String
Private mlngSize As Long
Private mlngDecimals As Long
Private mlngMaxSize As Long

Private msDefault_CHAR As String
Private mdtDefault_DATE As Date
Private msDefault_DATESTRING As String
Private mfDefault_LOGIC As Boolean
Private mdblDefault_NUMERIC As Double
Private msDefault_LISTVALUE As String
Private msDefault_LookupValue As String

Private mlngLeft As Long
Private mlngTop As Long
Private mlngWidth As Long
Private mlngHeight As Long
Private miOrientation As Integer
Private mlngVerticalOffset() As Long
Private mlngVerticalOffsetBehaviour As Integer
Private mlngHorizontalOffset() As Long
Private mlngHorizontalOffsetBehaviour As Long
Private mlngHeightBehaviour As Long
Private mlngWidthBehaviour As Long
Private mColBackColor As OLE_COLOR
Private mColBackColorEven As OLE_COLOR
Private mColBackColorOdd As OLE_COLOR
Private mColBackColorHighlight As OLE_COLOR
Private mColForeColor As OLE_COLOR
Private mColForeColorEven As OLE_COLOR
Private mColForeColorOdd As OLE_COLOR
Private mColForeColorHighlight As OLE_COLOR
Private mColHeaderBackColor As OLE_COLOR
Private mlngHeadlines As Long
Private mlngPictureID As Long
Private mlngTableID As Long
Private miBorderStyle As Integer
Private mfColumnHeaders As Boolean
Private mObjFont As StdFont
Private mObjHeadFont As StdFont
Private mlngTimeoutFrequency As Long
Private miTimeoutPeriod As TimeoutPeriod
Private msControlValueList_TAB As String
Private mlngLookupTableID As Long
Private mlngLookupColumnID As Long
Private mlngRecordTableID As Long
Private mfPasswordType As Boolean
Private mlngTabNumber As Long

' mlngPictureLocation is used as main web form property i.e. not a WebFormItem
Private mlngPictureLocation As Long

' BackStyle text constants.
Private Const msBACKSTYLEOPAQUETEXT = "Opaque"
Private Const msBACKSTYLETRANSPARENTTEXT = "Transparent"
' Alignment text constants.
Private Const msALIGNMENTLEFTTEXT = "Left Alignment"
Private Const msALIGNMENTRIGHTTEXT = "Right Alignment"
' Orientation text constants.
Private Const msORIENTATIONVERTICAL = "Vertical"
Private Const msORIENTATIONHORIZONTAL = "Horizontal"
Private Const msVERTICALFROMTOP = "Top"
Private Const msVERTICALFROMBOTTOM = "Bottom"
Private Const msHORIZONTALFROMLEFT = "Left"
Private Const msHORIZONTALFROMRIGHT = "Right"
Private Const msBEHAVEFIXED = "Fixed"
Private Const msBEHAVEFULL = "Full"
' BorderStyle text constants.
Private Const msBORDERSTYLENONETEXT = "No"
Private Const msBORDERSTYLEFIXEDSINGLETEXT = "Yes"
' Display Type text constants.
Private Const gsDISPLAYTYPECONTENTSTEXT = "Contents"
Private Const msDISPLAYTYPEICONTEXT = "Icon"
' Picture text constants.
Private Const msPICTURENONETEXT = "(None)"
Private Const msPICTURESELECTEDTEXT = "(Picture)"

'Boolean True/False text string constants
Private Const msBOOLEAN_TRUE = "True"
Private Const msBOOLEAN_FALSE = "False"

Private Const PICLOC_TOPLEFT = "Top Left"
Private Const PICLOC_TOPRIGHT = "Top Right"
Private Const PICLOC_CENTRE = "Centre"
Private Const PICLOC_LEFTTILE = "Left Tile"
Private Const PICLOC_RIGHTTILE = "Right Tile"
Private Const PICLOC_TOPTILE = "Top Tile"
Private Const PICLOC_BOTTOMTILE = "Bottom Tile"
Private Const PICLOC_TILE = "Tile"

Private Const MIN_FORM_HEIGHT = 2850
Private Const MIN_FORM_WIDTH = 2850

Private mactlSelectedControls() As VB.Control

Private miCurrentRowFormat As Integer

' Functions to display/tile the background image
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal lDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal lDC As Long, ByVal hObject As Long) As Long

Private mlngPersonnelTableID As Long
Private miInitiationType As WorkflowInitiationTypes
Private mlngBaseTableID As Long



Public Sub ApplyCurrentProperty()
  ssGridProperties_BeforeUpdate 0

End Sub


Public Property Get CurrentWebForm() As frmWorkflowWFDesigner
  ' Return the current web form.
  Set CurrentWebForm = mfrmWebForm
End Property
Public Property Set CurrentWebForm(pFrmNewValue As frmWorkflowWFDesigner)
  ' Set the current Web form property of the form.
  Set mfrmWebForm = pFrmNewValue
  mlngBaseTableID = pFrmNewValue.BaseTable
  miInitiationType = pFrmNewValue.InitiationType
End Property

Private Sub SetDBRecordSelectorType(piWFDatabaseRecord As WorkflowRecordSelectorTypes)
  Dim aWFPrecedingElements() As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sWebForm As String
  Dim sDfltWebForm As String
  Dim fWebFormOK As Boolean
  Dim lngTableID As Long
  Dim ctlControl As Control
  Dim iControlType As Integer
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngLoop As Long

  ReDim aWFPrecedingElements(0)
  
  sWebForm = msWFWebForm
  lngTableID = 0
  
  If piWFDatabaseRecord <> mlngWFDatabaseRecord Then
    mlngWFDatabaseRecord = piWFDatabaseRecord
    UpdateControls (WFITEMPROP_DBRECORD)
  End If

  If mlngWFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD Then
    For Each ctlControl In mfrmWebForm.Controls
      If mfrmWebForm.IsWebFormControl(ctlControl) Then
        iControlType = mfrmWebForm.WebFormControl_Type(ctlControl)
        
        With ctlControl
          If (.Selected) Then
            If ((iControlType = giWFFORMITEM_DBVALUE) _
              Or (iControlType = giWFFORMITEM_DBFILE)) Then
              
              ' Check if the selected database values are for different tables
              ' lngTableID = 0, not yet determined the selected database value's table
              ' lngTableID = -1, selected database value's have different tables
              If lngTableID = 0 Then
                lngTableID = GetTableIDFromColumnID(.ColumnID)
              ElseIf lngTableID > 0 And lngTableID <> GetTableIDFromColumnID(.ColumnID) Then
                lngTableID = -1
              End If
            End If
          End If
        End With
      End If
    Next ctlControl
    Set ctlControl = Nothing

    If lngTableID > 0 Then
      mfrmWebForm.PrecedingElements aWFPrecedingElements

      If UBound(aWFPrecedingElements) > 1 Then
        fWebFormOK = False
        sDfltWebForm = vbNullString

        For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
          Set wfTemp = aWFPrecedingElements(iLoop)
          
          If wfTemp.ElementType = elem_WebForm Then
            asItems = wfTemp.Items

            For iLoop2 = 1 To UBound(asItems, 2)
              If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                ReDim alngValidTables(0)
                TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                
                fFound = False
                For lngLoop = 1 To UBound(alngValidTables)
                  If alngValidTables(lngLoop) = lngTableID Then
                    fFound = True
                    Exit For
                  End If
                Next lngLoop
                If fFound Then
                  If UCase(wfTemp.Identifier) = UCase(sWebForm) Then
                    fWebFormOK = True
                  Else
                    If Len(sDfltWebForm) = 0 Or (UCase(wfTemp.Identifier) < UCase(sDfltWebForm)) Then
                      sDfltWebForm = wfTemp.Identifier
                    End If
                  End If

                  Exit For
                End If
              End If
            Next iLoop2

            Set wfTemp = Nothing
          ElseIf wfTemp.ElementType = elem_StoredData Then
            ReDim alngValidTables(0)
            TableAscendants wfTemp.DataTableID, alngValidTables
            
            'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
            'If wfTemp.DataAction = DATAACTION_DELETE Then
            '  ' Cannot do anything with a Deleted record, but can use its ascendants.
            '  ' Remove the table itself from the array of valid tables.
            '  alngValidTables(1) = 0
            'End If
            
            fFound = False
            For lngLoop = 1 To UBound(alngValidTables)
              If alngValidTables(lngLoop) = lngTableID Then
                fFound = True
                Exit For
              End If
            Next lngLoop
            If fFound Then
              If UCase(wfTemp.Identifier) = UCase(sWebForm) Then
                fWebFormOK = True
              Else
                If Len(sDfltWebForm) = 0 Or (UCase(wfTemp.Identifier) < UCase(sDfltWebForm)) Then
                  sDfltWebForm = wfTemp.Identifier
                End If
              End If
            End If
          End If
        Next iLoop

        If Not fWebFormOK Then
          sWebForm = sDfltWebForm
        End If
      End If
    End If
  Else
    sWebForm = vbNullString
  End If

  SetElementSelection sWebForm, WFITEMPROP_DBRECORD, lngTableID

End Sub


Private Sub ConfigureGridRow()
  ' Configure the grid for the currently selected row.
  On Error GoTo ErrorTrap

  Dim iLoop As Integer
  Dim iRow As Integer
  Dim iPropertyTag As Integer
  Dim fFixedCaption As Boolean
  Dim fCalcCaption As Boolean
  
  With ssGridProperties
    ' Set the styleSet of the rows to show which is selected.
    For iLoop = 0 To ssGridProperties.Rows
      If iLoop = .Row Then
        If (val(.Columns(2).CellValue(.AddItemBookmark(iLoop))) = WFITEMPROP_NONE) _
          Or (val(.Columns(2).CellValue(.AddItemBookmark(iLoop))) = WFITEMPROP_UNKNOWN) Then
          
          .Columns(0).CellStyleSet "ssetActiveRowBold", iLoop
        Else
          .Columns(0).CellStyleSet "ssetActiveRow", iLoop
        End If
      Else
        If (val(.Columns(2).CellValue(.AddItemBookmark(iLoop))) = WFITEMPROP_NONE) _
          Or (val(.Columns(2).CellValue(.AddItemBookmark(iLoop))) = WFITEMPROP_UNKNOWN) Then
          
          .Columns(0).CellStyleSet "ssetDormantRowBold", iLoop
        Else
          .Columns(0).CellStyleSet "ssetDormantRow", iLoop
        End If
      End If
    Next iLoop

    .AllowUpdate = True

    miCurrentRowFormat = val(.Columns(2).CellText(.Bookmark))

    Select Case val(.Columns(2).CellText(.Bookmark))
      ' --------------------
      ' CATEGORY MARKERS
      ' --------------------
      Case WFITEMPROP_NONE
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEdit

      ' --------------------
      ' PROPERTIES SCREEN
      ' --------------------
      Case WFITEMPROP_UNKNOWN
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEditButton

      ' --------------------
      ' GENERAL PROPERTIES
      ' --------------------
      Case WFITEMPROP_WFIDENTIFIER
        .Columns(1).Locked = False
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEdit

      Case WFITEMPROP_CAPTION
        .Columns(1).Locked = False
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEdit

      Case WFITEMPROP_VERTICALOFFSET
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit
      
      Case WFITEMPROP_VERTICALOFFSETBEHAVIOUR
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msVERTICALFROMTOP, 0
        .Columns(1).AddItem msVERTICALFROMBOTTOM, 1
        
      Case WFITEMPROP_TOP
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      Case WFITEMPROP_ORIENTATION
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msORIENTATIONHORIZONTAL, 0
        .Columns(1).AddItem msORIENTATIONVERTICAL, 1
        
      Case WFITEMPROP_HORIZONTALOFFSET
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      Case WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msHORIZONTALFROMLEFT, 0
        .Columns(1).AddItem msHORIZONTALFROMRIGHT, 1
        
      Case WFITEMPROP_LEFT
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      Case WFITEMPROP_HEIGHTBEHAVIOUR
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msBEHAVEFIXED, 0
        .Columns(1).AddItem msBEHAVEFULL, 1
        
      Case WFITEMPROP_HEIGHT
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      Case WFITEMPROP_WIDTHBEHAVIOUR
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msBEHAVEFIXED, 0
        .Columns(1).AddItem msBEHAVEFULL, 1

      Case WFITEMPROP_WIDTH
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      ' --------------------
      ' APPEARANCE PROPERTIES
      ' --------------------
      Case WFITEMPROP_BORDERSTYLE
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msBORDERSTYLENONETEXT, 0
        .Columns(1).AddItem msBORDERSTYLEFIXEDSINGLETEXT, 1

      Case WFITEMPROP_ALIGNMENT
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msALIGNMENTLEFTTEXT, 0
        .Columns(1).AddItem msALIGNMENTRIGHTTEXT, 1

      Case WFITEMPROP_PASSWORDTYPE
        .Columns(1).Locked = True
        .Columns(1).DataType = vbBoolean
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msBOOLEAN_TRUE, 0
        .Columns(1).AddItem msBOOLEAN_FALSE, 1

      Case WFITEMPROP_COLUMNHEADERS
        .Columns(1).Locked = True
        .Columns(1).DataType = vbBoolean
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msBOOLEAN_TRUE, 0
        .Columns(1).AddItem msBOOLEAN_FALSE, 1

      Case WFITEMPROP_HEADLINES
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      Case WFITEMPROP_HEADFONT
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEditButton

      Case WFITEMPROP_HEADERBACKCOLOR
        .Columns(1).Locked = True
        .Columns(1).Style = ssStyleEditButton

      Case WFITEMPROP_FONT
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEditButton

      Case WFITEMPROP_FORECOLOR
        .Columns(1).Locked = True
        .Columns(1).Style = ssStyleEditButton

      Case WFITEMPROP_FORECOLOREVEN
        .Columns(1).Locked = True
        .Columns(1).Style = ssStyleEditButton

      Case WFITEMPROP_FORECOLORODD
        .Columns(1).Locked = True
        .Columns(1).Style = ssStyleEditButton

      Case WFITEMPROP_FORECOLORHIGHLIGHT
        .Columns(1).Locked = True
        .Columns(1).Style = ssStyleEditButton

      Case WFITEMPROP_BACKSTYLE
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem msBACKSTYLEOPAQUETEXT, 0
        .Columns(1).AddItem msBACKSTYLETRANSPARENTTEXT, 1

      Case WFITEMPROP_BACKCOLOR
        .Columns(1).Style = ssStyleEditButton
        .Columns(1).Locked = True

      Case WFITEMPROP_BACKCOLOREVEN
        .Columns(1).Style = ssStyleEditButton
        .Columns(1).Locked = True

      Case WFITEMPROP_BACKCOLORODD
        .Columns(1).Style = ssStyleEditButton
        .Columns(1).Locked = True

      Case WFITEMPROP_BACKCOLORHIGHLIGHT
        .Columns(1).Locked = True
        .Columns(1).Style = ssStyleEditButton

      Case WFITEMPROP_PICTURE
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEditButton

      Case WFITEMPROP_PICTURELOCATION
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem PICLOC_TOPLEFT, 0
        .Columns(1).AddItem PICLOC_TOPRIGHT, 1
        .Columns(1).AddItem PICLOC_CENTRE, 2
        .Columns(1).AddItem PICLOC_LEFTTILE, 3
        .Columns(1).AddItem PICLOC_RIGHTTILE, 4
        .Columns(1).AddItem PICLOC_TOPTILE, 5
        .Columns(1).AddItem PICLOC_BOTTOMTILE, 6
        .Columns(1).AddItem PICLOC_TILE, 7
        
      Case WFITEMPROP_TABNUMBER
        .Columns(1).Locked = False
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEdit

      Case WFITEMPROP_TABCAPTION
        .Columns(1).Locked = False
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEdit

      Case Else
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEdit
    End Select

    ' Activate the 'values' column.
    If .col <> 1 Then
      .col = 1
    End If

  End With

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit

End Sub

Private Sub SetLookupDefaultValue(psValue As String)
  
  If psValue <> msDefault_LookupValue Then
    msDefault_LookupValue = psValue
    UpdateControls WFITEMPROP_DEFAULTVALUE_LOOKUP
  End If

End Sub

Private Sub SetRecordSelectorType(piWFDatabaseRecord As WorkflowRecordSelectorTypes)
  Dim aWFPrecedingElements() As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim aLngTableIds() As Long
  Dim sWebForm As String
  Dim sDfltWebForm As String
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim fWebFormOK As Boolean
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngLoop As Long
  
  sWebForm = msWFWebForm
  ReDim aWFPrecedingElements(0)
  
  If piWFDatabaseRecord <> mlngWFDatabaseRecord Then
    mlngWFDatabaseRecord = piWFDatabaseRecord
    UpdateControls (WFITEMPROP_RECSELTYPE)
  End If
  
  If mlngWFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD Then
    mfrmWebForm.PrecedingElements aWFPrecedingElements

    If UBound(aWFPrecedingElements) > 1 Then
      ReDim aLngTableIds(0)
  
      sSQL = "SELECT tmpRelations.parentID" & _
        " FROM tmpRelations" & _
        " WHERE tmpRelations.childID = " & CStr(mlngTableID)
      Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
      Do While Not (rsTables.BOF Or rsTables.EOF)
        ReDim Preserve aLngTableIds(UBound(aLngTableIds) + 1)
        aLngTableIds(UBound(aLngTableIds)) = rsTables!parentID
        
        rsTables.MoveNext
      Loop
      rsTables.Close
      Set rsTables = Nothing

      fWebFormOK = False
      sDfltWebForm = vbNullString
      
      For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
        If aWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
          Set wfTemp = aWFPrecedingElements(iLoop)
          asItems = wfTemp.Items
          
          For iLoop2 = 1 To UBound(asItems, 2)
            If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
              ReDim alngValidTables(0)
              TableAscendants CLng(asItems(44, iLoop2)), alngValidTables

              For iLoop3 = 1 To UBound(aLngTableIds)
                fFound = False
                For lngLoop = 1 To UBound(alngValidTables)
                  If alngValidTables(lngLoop) = aLngTableIds(iLoop3) Then
                    fFound = True
                    Exit For
                  End If
                Next lngLoop
                If fFound Then
                  If UCase(wfTemp.Identifier) = UCase(sWebForm) Then
                    fWebFormOK = True
                  Else
                    If Len(sDfltWebForm) = 0 Or (UCase(wfTemp.Identifier) < UCase(sDfltWebForm)) Then
                      sDfltWebForm = wfTemp.Identifier
                    End If
                  End If
                  
                  Exit For
                End If
              Next iLoop3
            End If
          Next iLoop2
          
          Set wfTemp = Nothing
        End If
      Next iLoop
    
      If Not fWebFormOK Then
        sWebForm = sDfltWebForm
      End If
    End If
  Else
    sWebForm = vbNullString
  End If
  
  SetElementSelection sWebForm, WFITEMPROP_RECSELTYPE, 0

End Sub



Private Sub SetRecordTableID(plngRecordTableID As Long)
  If plngRecordTableID <> mlngRecordTableID Then
    mlngRecordTableID = plngRecordTableID
    UpdateControls WFITEMPROP_RECORDTABLEID
  End If

End Sub



Private Sub SetTableID(plngTableID As Long)
  Dim aWFPrecedingElements() As VB.Control
  Dim lngLoop As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim sTableIDs As String
  Dim iRecordSelectorType As WorkflowRecordSelectorTypes
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  
  ReDim aWFPrecedingElements(0)
  
  iRecordSelectorType = mlngWFDatabaseRecord

  If plngTableID <> mlngTableID Then
    mlngTableID = plngTableID
    UpdateControls WFITEMPROP_TABLEID
  End If
  
  ' Ensure the record selector options are correct.
  Select Case mlngWFDatabaseRecord
    Case giWFRECSEL_INITIATOR:
      ' If the table is not a child of the defined Personnel table then RecSel cannot be Initiator
      ReDim alngValidTables(0)
      TableAscendants mlngPersonnelTableID, alngValidTables
      
      fFound = False
      For lngLoop = 1 To UBound(alngValidTables)
        If IsChildOfTable(alngValidTables(lngLoop), mlngTableID) Then
          fFound = True
          Exit For
        End If
      Next lngLoop
      If Not fFound Then
        iRecordSelectorType = giWFRECSEL_ALL
      End If
      
    Case giWFRECSEL_TRIGGEREDRECORD:
      ' If the table is not a child of the defined base table then RecSel cannot be triggered
      ReDim alngValidTables(0)
      TableAscendants mlngBaseTableID, alngValidTables
      
      fFound = False
      For lngLoop = 1 To UBound(alngValidTables)
        If IsChildOfTable(alngValidTables(lngLoop), mlngTableID) Then
          fFound = True
          Exit For
        End If
      Next lngLoop
      If Not fFound Then
        iRecordSelectorType = giWFRECSEL_ALL
      End If
          
    Case giWFRECSEL_IDENTIFIEDRECORD:
      ' If the table is not a child of a preceding webform's recordSelector then RecSel cannot be RecSel
      sTableIDs = "0"

      mfrmWebForm.PrecedingElements aWFPrecedingElements
        
      If UBound(aWFPrecedingElements) > 1 Then
        ' Add  an item to the combo for each preceding web form.
        For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
          Set wfTemp = aWFPrecedingElements(iLoop)
          
          If wfTemp.ElementType = elem_WebForm Then
            asItems = wfTemp.Items
            For iLoop2 = 1 To UBound(asItems, 2)
              If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                ReDim alngValidTables(0)
                TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                
                For lngLoop = 1 To UBound(alngValidTables)
                  sTableIDs = sTableIDs & "," & CStr(alngValidTables(lngLoop))
                Next lngLoop
              End If
            Next iLoop2
          ElseIf wfTemp.ElementType = elem_StoredData Then
            ReDim alngValidTables(0)
            TableAscendants wfTemp.DataTableID, alngValidTables
                          
            'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
            If wfTemp.DataAction = DATAACTION_DELETE Then
              ' Cannot do anything with a Deleted record, but can use its ascendants.
              ' Remove the table itself from the array of valid tables.
              alngValidTables(1) = 0
            End If
                          
            For lngLoop = 1 To UBound(alngValidTables)
              sTableIDs = sTableIDs & "," & CStr(alngValidTables(lngLoop))
            Next lngLoop
          End If
        
          Set wfTemp = Nothing
        Next iLoop
      End If

      sSQL = "SELECT COUNT(*) AS [result]" & _
        " FROM tmpRelations" & _
        " WHERE tmpRelations.parentID IN(" & sTableIDs & ")" & _
        " AND tmpRelations.childID = " & CStr(mlngTableID)
      Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          
      If rsTables!result = 0 Then
        iRecordSelectorType = giWFRECSEL_ALL
      End If
          
      rsTables.Close
      Set rsTables = Nothing

    Case giWFRECSEL_ALL:
  End Select
        
  SetRecordSelectorType iRecordSelectorType

End Sub




Private Sub SetLookupColumnID(plngColumnID As Long)
  Dim sSQL As String
  Dim rsLookupValues As New ADODB.Recordset
  Dim sDefault_LookupValue As String
  Dim iDataType As Integer
  Dim iOldDataType As Integer
  
  sDefault_LookupValue = msDefault_LookupValue
  iOldDataType = GetColumnDataType(mlngLookupColumnID)

  If plngColumnID <> mlngLookupColumnID Then
    mlngLookupColumnID = plngColumnID
    UpdateControls WFITEMPROP_LOOKUPCOLUMNID
  End If

  iDataType = GetColumnDataType(mlngLookupColumnID)
  
  ' Ensure the Lookup Default value is valid.
  If (mlngLookupColumnID > 0) And (iDataType = iOldDataType) Then
    
    sSQL = "SELECT COUNT(*) AS [result]" & _
      " FROM " & GetTableName(GetTableIDFromColumnID(mlngLookupColumnID)) & _
      " WHERE " & GetColumnName(mlngLookupColumnID)
      
    Select Case iDataType
      Case dtNUMERIC, dtINTEGER
        If Len(msDefault_LookupValue) = 0 Then
          sSQL = sSQL & _
            " = 0"
        Else
          sSQL = sSQL & _
            " = " & UI.ConvertNumberForSQL(msDefault_LookupValue)
        End If
      Case dtTIMESTAMP
        If Len(msDefault_LookupValue) = 0 Then
          sSQL = sSQL & _
            " IS null"
        Else
          sSQL = sSQL & _
            " = '" & UI.ConvertDateLocaleToSQL(msDefault_LookupValue) & "'"
        End If
      Case Else
        sSQL = sSQL & _
          " = " & "'" & msDefault_LookupValue & "'"
    End Select
      
    rsLookupValues.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    If rsLookupValues!result = 0 Then
      sDefault_LookupValue = vbNullString
    End If
  
    rsLookupValues.Close
    Set rsLookupValues = Nothing
  Else
    sDefault_LookupValue = vbNullString
  End If
  
  SetLookupDefaultValue sDefault_LookupValue

End Sub



Private Sub SetLookupTableID(plngTableID As Long)
  Dim sSQL As String
  Dim rsCheck As DAO.Recordset
  Dim rsColumns As DAO.Recordset
  Dim lngLookupColumnID As Long

  lngLookupColumnID = mlngLookupColumnID

  If plngTableID <> mlngLookupTableID Then
    mlngLookupTableID = plngTableID
    UpdateControls WFITEMPROP_LOOKUPTABLEID
  End If

  If mlngLookupTableID > 0 Then
    ' Ensure the Lookup Column is correct.
    sSQL = "SELECT COUNT(*) AS [result]" & _
      " FROM tmpColumns" & _
      " WHERE tmpColumns.columnID  = " & CStr(mlngLookupColumnID) & _
      " AND tmpColumns.tableID = " & CStr(mlngLookupTableID) & _
      " AND tmpColumns.columnType <> " & CStr(giCOLUMNTYPE_SYSTEM) & _
      " AND tmpColumns.columnType <> " & CStr(giCOLUMNTYPE_LINK) & _
      " AND tmpColumns.deleted = FALSE " & _
      " AND tmpColumns.dataType <> " & CStr(dtLONGVARBINARY) & _
      " AND tmpColumns.dataType <> " & CStr(dtVARBINARY) & _
      " AND tmpColumns.dataType <> " & CStr(dtBIT)
    Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
    If rsCheck!result = 0 Then
      sSQL = "SELECT TOP 1 tmpColumns.columnID" & _
        " FROM tmpColumns" & _
        " WHERE tmpColumns.tableID = " & CStr(mlngLookupTableID) & _
        " AND tmpColumns.columnType <> " & CStr(giCOLUMNTYPE_SYSTEM) & _
        " AND tmpColumns.columnType <> " & CStr(giCOLUMNTYPE_LINK) & _
        " AND tmpColumns.deleted = FALSE " & _
        " AND tmpColumns.dataType <> " & CStr(dtLONGVARBINARY) & _
        " AND tmpColumns.dataType <> " & CStr(dtVARBINARY) & _
        " AND tmpColumns.dataType <> " & CStr(dtBIT) & _
        " ORDER BY tmpColumns.columnName"
        
      Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      If Not (rsColumns.BOF And rsColumns.EOF) Then
        lngLookupColumnID = rsColumns!ColumnID
      Else
        lngLookupColumnID = 0
      End If
    
      rsColumns.Close
      Set rsColumns = Nothing
    End If
  
    rsCheck.Close
    Set rsCheck = Nothing
  Else
    lngLookupColumnID = 0
  End If
  
  SetLookupColumnID lngLookupColumnID

End Sub





Private Sub SetElementSelection(psWebFormIdentifier As String, _
  piProperty As WFItemProperty, _
  plngTableID As Long)
  
  Dim sRecSelIdentifier As String
  Dim aWFPrecedingElements() As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim aLngTableIds() As Long
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim fRecSelIdentifierOK As Boolean
  Dim sDfltRecSelIdentifier As String
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngLoop As Long
  
  ReDim aWFPrecedingElements(0)
  
  sRecSelIdentifier = msWFRecordSelector
  
  If msWFWebForm <> psWebFormIdentifier Then
    msWFWebForm = psWebFormIdentifier
    UpdateControls (WFITEMPROP_ELEMENTIDENTIFIER)
  End If
  
  If Len(msWFWebForm) > 0 Then
    mfrmWebForm.PrecedingElements aWFPrecedingElements

    If UBound(aWFPrecedingElements) > 1 Then
      ReDim aLngTableIds(0)

      If piProperty = WFITEMPROP_RECSELTYPE Then
        ' Element selection for RecSel control.
        sSQL = "SELECT tmpRelations.parentID" & _
          " FROM tmpRelations" & _
          " WHERE tmpRelations.childID = " & CStr(mlngTableID)
        Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
        Do While Not (rsTables.BOF Or rsTables.EOF)
          ReDim Preserve aLngTableIds(UBound(aLngTableIds) + 1)
          aLngTableIds(UBound(aLngTableIds)) = rsTables!parentID
  
          rsTables.MoveNext
        Loop
        rsTables.Close
        Set rsTables = Nothing
      Else
        ' Element selection for DBValue control.
        ReDim Preserve aLngTableIds(UBound(aLngTableIds) + 1)
        aLngTableIds(UBound(aLngTableIds)) = plngTableID
      End If
      
      fRecSelIdentifierOK = False
      sDfltRecSelIdentifier = vbNullString
      
      For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
        If aWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
          Set wfTemp = aWFPrecedingElements(iLoop)
          
          If wfTemp.Identifier = msWFWebForm Then
            If wfTemp.ElementType = elem_WebForm Then
              asItems = wfTemp.Items
              For iLoop2 = 1 To UBound(asItems, 2)
                If asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID Then
                  ReDim alngValidTables(0)
                  TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                  
                  For iLoop3 = 1 To UBound(aLngTableIds)
                    fFound = False
                    For lngLoop = 1 To UBound(alngValidTables)
                      If alngValidTables(lngLoop) = aLngTableIds(iLoop3) Then
                        fFound = True
                        Exit For
                      End If
                    Next lngLoop
                    If fFound Then
                      'JPD 20061010 Fault 11355
                      If UCase(asItems(9, iLoop2)) = UCase(sRecSelIdentifier) Then
                        fRecSelIdentifierOK = True
                      Else
                        If Len(sDfltRecSelIdentifier) = 0 Or (UCase(asItems(9, iLoop2)) < UCase(sDfltRecSelIdentifier)) Then
                          sDfltRecSelIdentifier = asItems(9, iLoop2)
                        End If
                      End If
                      
                      Exit For
                    End If
                  Next iLoop3
                End If
              Next iLoop2
            End If
            
            Exit For
          End If
          
          Set wfTemp = Nothing
        End If
      Next iLoop
    
      If Not fRecSelIdentifierOK Then
        sRecSelIdentifier = sDfltRecSelIdentifier
      End If
    End If
  Else
    sRecSelIdentifier = vbNullString
  End If
  
  SetRecordSelector sRecSelIdentifier, piProperty
  
End Sub

Private Sub SetRecordSelector(psRecSelIdentifier As String, _
  piProperty As WFItemProperty)
  
  Dim lngTemp As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim aWFPrecedingElements() As VB.Control
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim fDone As Boolean
  Dim alngValidTables() As Long
  Dim sValidTableIDs As String
  Dim lngRecordTableID As Long
  Dim lngDfltRecordTableID As Long
  Dim fRecordTableOK As Boolean
  Dim lngLoop As Long
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  
  fRecordTableOK = False
  
  If msWFRecordSelector <> psRecSelIdentifier Then
    msWFRecordSelector = psRecSelIdentifier
    UpdateControls (WFITEMPROP_RECORDSELECTOR)
  End If

  lngRecordTableID = mlngRecordTableID
  ReDim alngValidTables(0)

  If mlngWFDatabaseRecord = giWFRECSEL_INITIATOR Then
    lngTemp = mlngPersonnelTableID
    TableAscendants lngTemp, alngValidTables
  
  ElseIf mlngWFDatabaseRecord = giWFRECSEL_TRIGGEREDRECORD Then
    lngTemp = mlngBaseTableID
    TableAscendants lngTemp, alngValidTables
  
  ElseIf mlngWFDatabaseRecord = giWFRECSEL_IDENTIFIEDRECORD Then
    ReDim aWFPrecedingElements(0)
    mfrmWebForm.PrecedingElements aWFPrecedingElements
    
    fDone = False
    For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore the first item, as it will be the current web form.
      Set wfTemp = aWFPrecedingElements(iLoop)

      If UCase(wfTemp.Identifier) = UCase(msWFWebForm) Then
        If aWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
          asItems = wfTemp.Items

          For iLoop2 = 1 To UBound(asItems, 2)
            If (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID) Then
              If (asItems(9, iLoop2) = msWFRecordSelector) Then
                lngTemp = CLng(asItems(44, iLoop2))
                TableAscendants lngTemp, alngValidTables
              End If
            End If
          Next iLoop2

        ElseIf aWFPrecedingElements(iLoop).ElementType = elem_StoredData Then
          lngTemp = aWFPrecedingElements(iLoop).DataTableID
          TableAscendants lngTemp, alngValidTables
        
          'JPD 20061227 DBValues can now be from DELETE StoredData elements, but NOT RecSels
          If piProperty = WFITEMPROP_RECSELTYPE Then
            If aWFPrecedingElements(iLoop).DataAction = DATAACTION_DELETE Then
              ' Cannot do anything with a Deleted record, but can use its ascendants.
              ' Remove the table itself from the array of valid tables.
              alngValidTables(1) = 0
            End If
          End If
        End If
        
        fDone = True
        Exit For
      End If

      If fDone Then
        Exit For
      End If
    Next iLoop
  End If
  
  sValidTableIDs = "0"
  For lngLoop = 1 To UBound(alngValidTables)
    sValidTableIDs = sValidTableIDs & "," & CStr(alngValidTables(lngLoop))
  Next lngLoop
  
  sSQL = "SELECT tmpRelations.parentID, tmpTables.tableName" & _
    " FROM tmpRelations, tmpTables" & _
    " WHERE tmpRelations.parentID IN (" & sValidTableIDs & ")" & _
    " AND tmpRelations.childID = " & CStr(mlngTableID) & _
    " AND tmpRelations.parentID = tmpTables.tableID" & _
    " AND tmpTables.deleted = FALSE" & _
    " ORDER BY tmpTables.tableName"
  Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  While Not rsTables.EOF
    If lngRecordTableID = rsTables!parentID Then
      fRecordTableOK = True
    Else
      If lngDfltRecordTableID = 0 Then
        lngDfltRecordTableID = rsTables!parentID
      End If
    End If
    rsTables.MoveNext
  Wend
  rsTables.Close
  Set rsTables = Nothing
  
 If Not fRecordTableOK Then
    lngRecordTableID = lngDfltRecordTableID
 End If
  
  SetRecordTableID lngRecordTableID
  
End Sub



Private Function ValidSize(psString As String) As Boolean
  ValidSize = ValidIntegerString(psString)
  
  If ValidSize Then
    ValidSize = (val(psString) <= mlngMaxSize)
  End If
  
End Function

Private Sub Form_Initialize()
  ssGridProperties.RowHeight = 239
  ReDim mactlSelectedControls(0)
  ReDim mlngVerticalOffset(0)
  ReDim mlngHorizontalOffset(0)
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

  On Error GoTo ErrorTrap
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
  
  'Get dimensions of windows borders
  miCXFrame = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
  miCYFrame = UI.GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY
  miCXBorder = UI.GetSystemMetrics(SM_CXBORDER) * Screen.TwipsPerPixelX
  miCYBorder = UI.GetSystemMetrics(SM_CYBORDER) * Screen.TwipsPerPixelY
  
  'Position and size form
  Me.Move 0, 0, 4650, (Forms(0).ScaleHeight * 2 / 3)

  ' Format properties grid.
  With ssGridProperties
    .Top = 0
    .Left = 0
  End With
  
  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
   
  ' Position the form at the right hand side of the screen.
  Me.Left = Screen.Width - Me.Width - miCXFrame
  
TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_Resize()

  On Error GoTo ErrorTrap

  Dim lngColumnWidth  As Long
  
  If Me.WindowState <> vbMinimized Then
    
    With ssGridProperties
      .Scroll -1, 0
      .Width = Me.ScaleWidth
      .Height = Me.ScaleHeight - .Top - StatusBar1.Height
      
      lngColumnWidth = .Width - .Columns(0).Width - (miCXBorder * 2)
      ' Need to make it a little narrower if the vertical scrollbar is displayed.
      If .VisibleRows = .Rows Then
        If CLng(.RowTop(.VisibleRows)) = CLng(.Height) Then
          lngColumnWidth = lngColumnWidth - 235
        End If
      ElseIf .VisibleRows < .Rows Then
        lngColumnWidth = lngColumnWidth - 235
      End If

      .Columns(1).Width = lngColumnWidth
    End With
    
    RefreshProperties True
  End If

  ' Get rid of the icon off the form
  RemoveIcon Me

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Disassociate object variables.
  Set mObjFont = Nothing
  Set mObjHeadFont = Nothing
  Set frmWorkflowWFItemProps = Nothing
  
  Unhook Me.hWnd
  
End Sub

Private Sub ssGridProperties_BeforeUpdate(Cancel As Integer)
  ' Update the selected controls with the new value.
  On Error GoTo ErrorTrap

  Dim fUpdateControls As Boolean
  Dim iPropertyTag As WFItemProperty
  Dim sNewValue As String
  Dim ctlCurrentControl As VB.Control
  Dim iarrLoop As Integer
  
  ' Read the new property value from the grid.
  iPropertyTag = val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  sNewValue = ssGridProperties.ActiveCell.Text
  fUpdateControls = True

  ' Update the required global variable with the new value if required.
  Select Case iPropertyTag
    ' --------------------
    ' GENERAL PROPERTIES
    ' --------------------
    Case WFITEMPROP_WFIDENTIFIER
      If Len(sNewValue) > 200 Then
        sNewValue = Left(sNewValue, 200)
        ssGridProperties.ActiveCell.Text = sNewValue
      End If

      If UBound(mactlSelectedControls) = 1 Then
        Set ctlCurrentControl = mactlSelectedControls(1)

        If ValidateWFIdentifier(sNewValue, ctlCurrentControl) Then
          msWFIdentifier = sNewValue
        End If
      ElseIf UBound(mactlSelectedControls) = 0 Then
        If ValidateWFElementIdentifier(sNewValue) Then
          msWFIdentifier = sNewValue
        End If
      End If

    Case WFITEMPROP_CAPTION
      If Len(sNewValue) > 200 Then
        sNewValue = Left(sNewValue, 200)
        ssGridProperties.ActiveCell.Text = sNewValue
      End If
      msCaption = sNewValue
    
    Case WFITEMPROP_VERTICALOFFSET
      If ValidIntegerString(sNewValue) Then
        If val(sNewValue) > 2000 Then
          sNewValue = "2000"
        ElseIf val(sNewValue) < 0 Then
          sNewValue = "0"
        End If
        
        For iarrLoop = 1 To UBound(mlngVerticalOffset)
          mlngVerticalOffset(iarrLoop) = val(sNewValue)
        Next iarrLoop
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(mlngVerticalOffset(1)))
        fUpdateControls = False
        Cancel = True
      End If
      
    Case WFITEMPROP_TOP
      If ValidIntegerString(sNewValue) Then
        If val(sNewValue) > 2000 Then
          sNewValue = "2000"
        ElseIf val(sNewValue) < 0 Then
          sNewValue = "0"
        End If
        
        mlngTop = val(sNewValue)
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(mlngTop))
        fUpdateControls = False
        Cancel = True
      End If

    Case WFITEMPROP_HORIZONTALOFFSET
      If ValidIntegerString(sNewValue) Then
        If val(sNewValue) > 2000 Then
          sNewValue = "2000"
        ElseIf val(sNewValue) < 0 Then
          sNewValue = "0"
        End If
        
        For iarrLoop = 1 To UBound(mlngHorizontalOffset)
          mlngHorizontalOffset(iarrLoop) = val(sNewValue)
        Next iarrLoop
        
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(mlngHorizontalOffset(1)))
        fUpdateControls = False
        Cancel = True
      End If
      
    Case WFITEMPROP_LEFT
      If ValidIntegerString(sNewValue) Then
        If val(sNewValue) > 2000 Then
          sNewValue = "2000"
        ElseIf val(sNewValue) < 0 Then
          sNewValue = "0"
        End If
        
        mlngLeft = val(sNewValue)
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(mlngLeft))
        fUpdateControls = False
        Cancel = True
      End If

    Case WFITEMPROP_HEIGHT
      If ValidIntegerString(sNewValue) Then
        If val(sNewValue) > 2000 Then
          sNewValue = "2000"
        ElseIf val(sNewValue) < 1 Then
          sNewValue = "1"
        End If
        
        mlngHeight = val(sNewValue)
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(mlngHeight))
        fUpdateControls = False
        Cancel = True
      End If

    Case WFITEMPROP_WIDTH
      If ValidIntegerString(sNewValue) Then
        If val(sNewValue) > 2000 Then
          sNewValue = "2000"
        ElseIf val(sNewValue) < 1 Then
          sNewValue = "1"
        End If
        
        mlngWidth = val(sNewValue)
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(mlngWidth))
        fUpdateControls = False
        Cancel = True
      End If
    
    Case WFITEMPROP_TABNUMBER
      If ValidIntegerString(sNewValue) Then
        mlngTabNumber = val(sNewValue)
        
        'MH20060915 Fault 11422
        'UpdateControls WFITEMPROP_TABNUMBER
      Else
        fUpdateControls = False
      End If
      
    Case WFITEMPROP_TABCAPTION
      If Len(sNewValue) > 200 Then
        sNewValue = Left(sNewValue, 200)
        ssGridProperties.ActiveCell.Text = sNewValue
      End If
      
    ' --------------------
    ' APPEARANCE PROPERTIES
    ' --------------------
    Case WFITEMPROP_HEADLINES
      If ValidIntegerString(sNewValue) Then
        mlngHeadlines = val(sNewValue)
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(mlngHeadlines))
        fUpdateControls = False
        Cancel = True
      End If
  
  End Select

  ' Update the selected controls with the new property values.
  If fUpdateControls Then
    UpdateControls iPropertyTag
  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ssGridProperties_BtnClick()
  ' Process the button click in the properties grid.
  On Error GoTo ErrorTrap

  Dim iPropertyTag As WFItemProperty
  Dim iItemType As WorkflowWebFormItemTypes
  Dim iLoop As Integer
  
  ' AE20080306 Fault #12970
  If ssGridProperties.Columns(3).value Then
    Exit Sub
  End If
  
  iPropertyTag = val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  
  ' Determine the type of the selected item.
  iItemType = giWFFORMITEM_UNKNOWN
  If Not (mfrmWebForm Is Nothing) Then
    ' Process each selected control.
    If UBound(mactlSelectedControls) > 0 Then
      For iLoop = 1 To UBound(mactlSelectedControls)
        iItemType = mfrmWebForm.WebFormControl_Type(mactlSelectedControls(iLoop))
      Next iLoop
    Else
      iItemType = giWFFORMITEM_FORM
    End If
  End If
  
  If iItemType <> giWFFORMITEM_UNKNOWN Then
    ' Display the form required for changing the current property.
    ' Read the new property value from the form into a global variable.
    Select Case iPropertyTag
      ' --------------------
      ' PROPERTIES SCREEN
      ' --------------------
      Case WFITEMPROP_UNKNOWN
        ' Display the general properties form.
        If Not (mfrmWebForm Is Nothing) Then
          mfrmWebForm.ShowPropertiesForm
        End If
        
      ' --------------------
      ' APPEARANCE PROPERTIES
      ' --------------------
      Case WFITEMPROP_HEADFONT
        ' Display the Font dialogue box.
        With comDlgBox
          .FontName = mObjHeadFont.Name
          .FontSize = mObjHeadFont.Size
          .FontBold = mObjHeadFont.Bold
          .FontItalic = mObjHeadFont.Italic
          .FontUnderline = mObjHeadFont.Underline
          .FontStrikethru = mObjHeadFont.Strikethrough
          .Flags = cdlCFScreenFonts Or cdlCFEffects
          .ShowFont
          mObjHeadFont.Name = .FontName
          mObjHeadFont.Size = .FontSize
          mObjHeadFont.Bold = .FontBold
          mObjHeadFont.Italic = .FontItalic
          mObjHeadFont.Underline = .FontUnderline
          mObjHeadFont.Strikethrough = .FontStrikethru
        End With
        ' Update the grid display.
        ssGridProperties.Columns(1).Text = GetHeadFontDescription
  
      Case WFITEMPROP_HEADERBACKCOLOR
        ' Display the Background Colour dialogue box.
'        With comDlgBox
'          .Flags = cdlCCRGBInit
'          .Color = mColHeaderBackColor
'          .ShowColor
'          mColHeaderBackColor = .Color
'        End With

        With colPickDlg
          .Color = mColHeaderBackColor
          .ShowPalette
          mColHeaderBackColor = .Color
        End With
        
        ' Update the grid display.
        With ssGridProperties
          .StyleSets("ssetHeaderBackColor").BackColor = mColHeaderBackColor
          .Columns(1).CellStyleSet "ssetHeaderBackColor", .Row
        End With
  
      Case WFITEMPROP_FONT
        ' Display the Font dialogue box.
        With comDlgBox
          .FontName = mObjFont.Name
          .FontSize = mObjFont.Size
          .FontBold = mObjFont.Bold
          .FontItalic = mObjFont.Italic
          .FontUnderline = mObjFont.Underline
          .FontStrikethru = mObjFont.Strikethrough
          .Flags = cdlCFScreenFonts Or cdlCFEffects
            .ShowFont
          mObjFont.Name = .FontName
          mObjFont.Size = .FontSize
          mObjFont.Bold = .FontBold
          mObjFont.Italic = .FontItalic
          mObjFont.Underline = .FontUnderline
          mObjFont.Strikethrough = .FontStrikethru
        End With
        ' Update the grid display.
        ssGridProperties.Columns(1).Text = GetFontDescription
  
      Case WFITEMPROP_FORECOLOR
        ' Display the Foreground Colour dialogue box.
'        With comDlgBox
'          .Flags = cdlCCRGBInit
'          .Color = mColForeColor
'          .ShowColor
'          mColForeColor = .Color
'        End With
        
        With colPickDlg
          .Color = mColForeColor
          .ShowPalette
          mColForeColor = .Color
        End With
        
        ' Update the grid display.
        With ssGridProperties
          .StyleSets("ssetForeColorValue").BackColor = mColForeColor
          .Columns(1).CellStyleSet "ssetForeColorValue", .Row
        End With
  
      Case WFITEMPROP_FORECOLOREVEN
        ' Display the Foreground Colour dialogue box.
'        With comDlgBox
'          .Flags = cdlCCRGBInit
'          .Color = mColForeColorEven
'          .ShowColor
'          mColForeColorEven = .Color
'        End With
        
        With colPickDlg
          .Color = mColForeColorEven
          .ShowPalette
          mColForeColorEven = .Color
        End With
        
        With ssGridProperties
          .StyleSets("ssetForeColorEven").BackColor = mColForeColorEven
          .Columns(1).CellStyleSet "ssetForeColorEven", .Row
        End With
  
      Case WFITEMPROP_FORECOLORODD
        ' Display the Foreground Colour dialogue box.
'        With comDlgBox
'          .Flags = cdlCCRGBInit
'          .Color = mColForeColorOdd
'          .ShowColor
'          mColForeColorOdd = .Color
'        End With

        With colPickDlg
          .Color = mColForeColorOdd
          .ShowPalette
          mColForeColorOdd = .Color
        End With
        
        ' Update the grid display.
        With ssGridProperties
          .StyleSets("ssetForeColorOdd").BackColor = mColForeColorOdd
          .Columns(1).CellStyleSet "ssetForeColorOdd", .Row
        End With
  
      Case WFITEMPROP_FORECOLORHIGHLIGHT
        ' Display the Colour dialogue box.
'        With comDlgBox
'          .Flags = cdlCCRGBInit
'          .Color = mColForeColorHighlight
'          .ShowColor
'          mColForeColorHighlight = .Color
'        End With

        With colPickDlg
          .Color = mColForeColorHighlight
          .ShowPalette
          mColForeColorHighlight = .Color
        End With
        
        ' Update the grid display.
        With ssGridProperties
          .StyleSets("ssetForeColorHighlight").BackColor = mColForeColorHighlight
          .Columns(1).CellStyleSet "ssetForeColorHighlight", .Row
        End With
  
      Case WFITEMPROP_BACKCOLOR
        ' Display the Background Colour dialogue box.
'        With comDlgBox
'          .Flags = cdlCCRGBInit
'          .Color = mColBackColor
'          .ShowColor
'          mColBackColor = .Color
'        End With
        
        With colPickDlg
          .Color = mColBackColor
          .ShowPalette
          mColBackColor = .Color
        End With
        
        ' Update the grid display.
        With ssGridProperties
          .StyleSets("ssetBackColorValue").BackColor = mColBackColor
          .Columns(1).CellStyleSet "ssetBackColorValue", .Row
        End With
  
      Case WFITEMPROP_BACKCOLOREVEN
        ' Display the Background Colour dialogue box.
'        With comDlgBox
'          .Flags = cdlCCRGBInit
'          .Color = mColBackColorEven
'          .ShowColor
'          mColBackColorEven = .Color
'        End With

        With colPickDlg
          .Color = mColBackColorEven
          .ShowPalette
          mColBackColorEven = .Color
        End With
        
        ' Update the grid display.
        With ssGridProperties
          .StyleSets("ssetBackColorEven").BackColor = mColBackColorEven
          .Columns(1).CellStyleSet "ssetBackColorEven", .Row
        End With
  
      Case WFITEMPROP_BACKCOLORODD
        ' Display the Background Colour dialogue box.
'        With comDlgBox
'          .Flags = cdlCCRGBInit
'          .Color = mColBackColorOdd
'          .ShowColor
'          mColBackColorOdd = .Color
'        End With

        With colPickDlg
          .Color = mColBackColorOdd
          .ShowPalette
          mColBackColorOdd = .Color
        End With
        
        ' Update the grid display.
        With ssGridProperties
          .StyleSets("ssetBackColorOdd").BackColor = mColBackColorOdd
          .Columns(1).CellStyleSet "ssetBackColorOdd", .Row
        End With
  
      Case WFITEMPROP_BACKCOLORHIGHLIGHT
        ' Display the Colour dialogue box.
'        With comDlgBox
'          .Flags = cdlCCRGBInit
'          .Color = mColBackColorHighlight
'          .ShowColor
'          mColBackColorHighlight = .Color
'        End With
        
        With colPickDlg
          .Color = mColBackColorHighlight
          .ShowPalette
          mColBackColorHighlight = .Color
        End With
        
        ' Update the grid display.
        With ssGridProperties
          .StyleSets("ssetBackColorHighlight").BackColor = mColBackColorHighlight
          .Columns(1).CellStyleSet "ssetBackColorHighlight", .Row
        End With

      Case WFITEMPROP_PICTURE
        ' Display the Picture selection form.
        frmPictSel.SelectedPicture = mlngPictureID
        frmPictSel.ExcludedExtensions = ""
        frmPictSel.Show vbModal
        If frmPictSel.SelectedPicture > 0 Then
          With recPictEdit
            .Index = "idxID"
            .Seek "=", frmPictSel.SelectedPicture
            If Not .NoMatch Then
              mlngPictureID = !PictureID
              ssGridProperties.Columns(1).Text = !Name
            End If
          End With
        Else
          ssGridProperties.Columns(1).Text = "<None>"
        End If
  
        Set frmPictSel = Nothing
    End Select
  
    ' Update the selected controls with the new property value.
    If iPropertyTag <> WFITEMPROP_UNKNOWN Then
      UpdateControls iPropertyTag
    End If
  End If
  
TidyUpAndExit:
  Exit Sub

ErrorTrap:
  ' User pressed cancel.
  Err = False
  Resume TidyUpAndExit
  
End Sub

Private Sub ssGridProperties_Click()
  frmWorkflowWFItemProps.SetFocus
  frmSysMgr.RefreshMenu
  
End Sub

Private Sub ssGridProperties_ComboCloseUp()
  ' Process the combo selection in the properties grid.
  On Error GoTo ErrorTrap

  Dim iPropertyTag As WFItemProperty
  Dim sNewValue As String

  iPropertyTag = val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  sNewValue = ssGridProperties.ActiveCell.Text

  ' Read the new property value from the grid.
  Select Case iPropertyTag
    ' --------------------
    ' APPEARANCE PROPERTIES
    ' --------------------
    Case WFITEMPROP_BORDERSTYLE
      If sNewValue = msBORDERSTYLENONETEXT Then
        miBorderStyle = vbBSNone
      Else
        miBorderStyle = vbFixedSingle
      End If

    Case WFITEMPROP_ORIENTATION
      If sNewValue = msORIENTATIONHORIZONTAL Then
        miOrientation = wfItemPropertyOrientation_Horizontal
      ElseIf sNewValue = msORIENTATIONVERTICAL Then
        miOrientation = wfItemPropertyOrientation_Vertical
      End If
    
    Case WFITEMPROP_VERTICALOFFSETBEHAVIOUR
      If sNewValue = msVERTICALFROMTOP And mlngVerticalOffsetBehaviour <> offsetTop Then
        mlngVerticalOffsetBehaviour = offsetTop
      ElseIf sNewValue = msVERTICALFROMBOTTOM And mlngVerticalOffsetBehaviour <> offsetBottom Then
        mlngVerticalOffsetBehaviour = offsetBottom
      End If
        
    Case WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR
      If sNewValue = msHORIZONTALFROMLEFT And mlngHorizontalOffsetBehaviour <> offsetLeft Then
        mlngHorizontalOffsetBehaviour = offsetLeft
      ElseIf sNewValue = msHORIZONTALFROMRIGHT And mlngHorizontalOffsetBehaviour <> offsetRight Then
        mlngHorizontalOffsetBehaviour = offsetRight
      End If
    
    Case WFITEMPROP_HEIGHTBEHAVIOUR
      If sNewValue = msBEHAVEFIXED And Not (mlngHeightBehaviour = behaveFixed) Then
        mlngHeightBehaviour = behaveFixed
      ElseIf sNewValue = msBEHAVEFULL And Not (mlngHeightBehaviour = behaveFull) Then
        mlngHeightBehaviour = behaveFull
      End If
    
    Case WFITEMPROP_WIDTHBEHAVIOUR
      If sNewValue = msBEHAVEFIXED And Not (mlngWidthBehaviour = behaveFixed) Then
        mlngWidthBehaviour = behaveFixed
      ElseIf sNewValue = msBEHAVEFULL And Not (mlngWidthBehaviour = behaveFull) Then
        mlngWidthBehaviour = behaveFull
      End If
    
    Case WFITEMPROP_ALIGNMENT
      If sNewValue = msALIGNMENTLEFTTEXT Then
        miAlignment = vbLeftJustify
      Else
        miAlignment = vbRightJustify
      End If

    Case WFITEMPROP_PASSWORDTYPE ' Hide text
      If sNewValue = msBOOLEAN_TRUE Then
        mfPasswordType = True
      Else
        mfPasswordType = False
      End If

    Case WFITEMPROP_COLUMNHEADERS
      If sNewValue = msBOOLEAN_TRUE Then
        mfColumnHeaders = True
      Else
        mfColumnHeaders = False
      End If

    Case WFITEMPROP_BACKSTYLE
      If sNewValue = msBACKSTYLEOPAQUETEXT Then
        miBackStyle = BACKSTYLE_OPAQUE
      Else
        miBackStyle = BACKSTYLE_TRANSPARENT
      End If

    Case WFITEMPROP_PICTURELOCATION
      Select Case sNewValue
        Case PICLOC_TOPLEFT
          mlngPictureLocation = 0
        Case PICLOC_TOPRIGHT
          mlngPictureLocation = 1
        Case PICLOC_CENTRE
          mlngPictureLocation = 2
        Case PICLOC_LEFTTILE
          mlngPictureLocation = 3
        Case PICLOC_RIGHTTILE
          mlngPictureLocation = 4
        Case PICLOC_TOPTILE
          mlngPictureLocation = 5
        Case PICLOC_BOTTOMTILE
          mlngPictureLocation = 6
        Case PICLOC_TILE
          mlngPictureLocation = 7
      End Select
      
  End Select

  ' Update the selected controls with the new property value.
  UpdateControls iPropertyTag

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit

End Sub
Private Sub ssGridProperties_DblClick()
  ' Process the double click in the properties grid.
  On Error GoTo ErrorTrap

  With ssGridProperties
    Select Case .Columns(1).Style
      Case ssStyleEditButton
        ssGridProperties_BtnClick
      
      Case ssStyleComboBox
        If .Columns(1).ListCount > 1 Then
          .Columns(1).ListIndex = ((.Columns(1).ListIndex + 1) Mod .Columns(1).ListCount)
          ssGridProperties_ComboCloseUp
        End If
      
      Case Else
    
    End Select
  End With
  
TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ssGridProperties_KeyPress(KeyAscii As Integer)
  ' Process the ENTER key press.
  On Error GoTo ErrorTrap

  Dim iProperty As Integer
  Dim sNewValue As String
  Dim iIndex As Integer
  
  iProperty = val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  sNewValue = ssGridProperties.ActiveCell.Text

  If KeyAscii = vbKeyReturn Then
    If ssGridProperties.Columns(1).Style = ssStyleEditButton Then
      ssGridProperties_BtnClick
    Else
      ssGridProperties_BeforeUpdate 0
    End If

  ElseIf (KeyAscii = vbKeyDelete) Or (KeyAscii = vbKeyBack) Then
    If val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark)) = WFITEMPROP_PICTURE Then
      mlngPictureID = 0
      UpdateControls WFITEMPROP_PICTURE
    End If
  End If


TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ssGridProperties_KeyUp(KeyCode As Integer, Shift As Integer)
  ' Process the ENTER key press.
  On Error GoTo ErrorTrap

  Dim iProperty As Integer
  Dim sNewValue As String
  Dim iIndex As Integer
  
  iProperty = val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  sNewValue = ssGridProperties.ActiveCell.Text
  
  If KeyCode <> vbKeyReturn Then
    ' --------------------
    ' GENERAL PROPERTIES
    ' --------------------
    If iProperty = WFITEMPROP_WFIDENTIFIER Then
      If Len(sNewValue) > 200 Then
        iIndex = ssGridProperties.ActiveCell.SelStart

        sNewValue = Left(sNewValue, 200)
        ssGridProperties.ActiveCell.Text = sNewValue
        ssGridProperties.ActiveCell.SelStart = iIndex
      End If
    End If
    
    If iProperty = WFITEMPROP_CAPTION Then
      If Len(sNewValue) > 200 Then
        iIndex = ssGridProperties.ActiveCell.SelStart

        sNewValue = Left(sNewValue, 200)
        ssGridProperties.ActiveCell.Text = sNewValue
        ssGridProperties.ActiveCell.SelStart = iIndex
      End If
    End If
  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit

End Sub

Private Sub ssGridProperties_LostFocus()
  ApplyCurrentProperty
  
End Sub

Private Sub ssGridProperties_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  ConfigureGridRow
End Sub

Private Function GetElementByIdentifier(psIdentifier As String) As VB.Control
  ' Return the element with the given identifier.
  Dim lngLoop As Long
  Dim wfTemp As VB.Control
  Dim aWFPrecedingElements() As VB.Control
  
  If Len(Trim(psIdentifier)) = 0 Then
    Exit Function
  End If
    
  ReDim aWFPrecedingElements(0)
  mfrmWebForm.PrecedingElements aWFPrecedingElements
  
  For lngLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
    Set wfTemp = aWFPrecedingElements(lngLoop)

    If (UCase(Trim(wfTemp.Identifier)) = UCase(Trim(psIdentifier))) Then
      Set GetElementByIdentifier = wfTemp
      Exit For
    End If
    
    Set wfTemp = Nothing
  Next lngLoop

End Function
Public Function EditMenu(psMenuOption As String) As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Pass all menu evants onto the active screen designer.
  mfrmWebForm.EditMenu psMenuOption
  
TidyUpAndExit:
  EditMenu = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function RefreshProperties(Optional pfStayOnSameLine As Boolean) As Boolean

  ' Refresh the properties grid for the selected controls in the screen designer.

  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fFixedCaption As Boolean
  Dim fPassedOnce As Boolean
  Dim iLoop As Integer
  Dim iControlType As WorkflowWebFormItemTypes
  Dim iLastControlType As WorkflowWebFormItemTypes
  Dim sDescription As String
  Dim objCtlFont As StdFont
  Dim ctlControl As VB.Control
  Dim avProperties(WFITEMPROPERTYCOUNT, 2) As Variant
  Dim iSelectedControlCount As Integer
  Dim iCurrentRow As Integer
  Dim bHorizontalLine As Boolean
  Dim bVerticalLine As Boolean
  Dim lngTabNumber As Long

  fOK = True

  ' Initialise the property array.
  ' NB. This array has a row for each property, and 3 columns.
  ' Column 1 - True if the selected controls all have the property.
  ' Column 2 - True if selected controls have different values for the property.
  For iLoop = 0 To UBound(avProperties)
    avProperties(iLoop, 1) = False
    avProperties(iLoop, 2) = False
  Next iLoop

  ' Ensure the global font object is instantiated.
  If mObjFont Is Nothing Then
    Set mObjFont = New StdFont
  End If
  If mObjHeadFont Is Nothing Then
    Set mObjHeadFont = New StdFont
  End If

  ReDim mactlSelectedControls(0)
  ReDim mlngVerticalOffset(0)
  ReDim mlngHorizontalOffset(0)
  Dim iCurrentIndex As Integer
  
  ' Determine which properties should be displayed in the grid, and their values.
  If Not (mfrmWebForm Is Nothing) Then
    ' --------------------
    ' At least one webForm item selected
    ' --------------------
    iSelectedControlCount = mfrmWebForm.SelectedControlsCount

    ' Process each selected control.
    If iSelectedControlCount > 0 Then
      ' Initialise the property array.
      For iLoop = 0 To UBound(avProperties)
        avProperties(iLoop, 1) = True
      Next iLoop

      fPassedOnce = False

      ' Get the properties of each screen control type.
      For Each ctlControl In mfrmWebForm.Controls
        If mfrmWebForm.IsWebFormControl(ctlControl) Then
          ' Get the screen control type.
          iControlType = mfrmWebForm.WebFormControl_Type(ctlControl)

          With ctlControl
            If (.Selected) Then
              ReDim Preserve mactlSelectedControls(UBound(mactlSelectedControls) + 1)
              Set mactlSelectedControls(UBound(mactlSelectedControls)) = ctlControl
              
              If WebFormItemHasProperty(ctlControl.WFItemType, WFITEMPROP_VERTICALOFFSET) _
                Or WebFormItemHasProperty(ctlControl.WFItemType, WFITEMPROP_HORIZONTALOFFSET) Then
                
                ' As we'll possibly be changing multiple controls offsets we need to keep them in an array
                ReDim Preserve mlngVerticalOffset(UBound(mlngVerticalOffset) + 1)
                ReDim Preserve mlngHorizontalOffset(UBound(mlngHorizontalOffset) + 1)
                
                mlngVerticalOffset(UBound(mlngVerticalOffset)) = ctlControl.VerticalOffset
                mlngHorizontalOffset(UBound(mlngHorizontalOffset)) = ctlControl.HorizontalOffset
                
                iCurrentIndex = UBound(mlngHorizontalOffset)
              End If
              
              If iControlType = giWFFORMITEM_LINE Then
                bVerticalLine = IIf(.Alignment = wfItemPropertyOrientation_Vertical, True, False)
                bHorizontalLine = IIf(.Alignment = wfItemPropertyOrientation_Horizontal, True, False)
              End If
              
              If iSelectedControlCount = 1 Then
                If ((iControlType = giWFFORMITEM_DBVALUE) _
                  Or (iControlType = giWFFORMITEM_DBFILE)) Then
                  
                  StatusBar1.SimpleText = "" & GetTableColumnName(.ColumnID)
                  Me.Caption = Me.StatusBar1.SimpleText
                ElseIf ((iControlType = giWFFORMITEM_WFVALUE) _
                  Or (iControlType = giWFFORMITEM_WFFILE)) Then
                 
                  StatusBar1.SimpleText = .WFWorkflowForm & "." & .WFWorkflowValue
                  Me.Caption = Me.StatusBar1.SimpleText
                Else
                  If WebFormItemHasProperty(iControlType, WFITEMPROP_WFIDENTIFIER) Then
                    StatusBar1.SimpleText = .WFIdentifier
                    Me.Caption = .WFIdentifier & " - Properties"
                  Else
                    StatusBar1.SimpleText = "Control Properties"
                    Me.Caption = "Control Properties"
                  End If
                End If
              ElseIf iSelectedControlCount > 1 Then
                StatusBar1.SimpleText = "<Multiple Controls>"
                Me.Caption = "<Multiple Controls>"
              End If

              ' --------------------
              ' PROPERTIES SCREEN
              ' --------------------
              avProperties(WFITEMPROP_UNKNOWN, 1) = (iSelectedControlCount = 1)
              
              ' --------------------
              ' GENERAL PROPERTIES
              ' --------------------
              ' Read the Web Form Identifier property from the control if required.
              If avProperties(WFITEMPROP_WFIDENTIFIER, 1) Then
                ' Do not allow Identifiers to be changed en-masse, as identifiers must be unique.
                avProperties(WFITEMPROP_WFIDENTIFIER, 1) = (iSelectedControlCount = 1)

                If WebFormItemHasProperty(iControlType, WFITEMPROP_WFIDENTIFIER) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_WFIDENTIFIER, 2)) Then
                    If msWFIdentifier <> .WFIdentifier Then
                      avProperties(WFITEMPROP_WFIDENTIFIER, 2) = True
                    End If
                  Else
                    msWFIdentifier = .WFIdentifier
                  End If
                Else
                  avProperties(WFITEMPROP_WFIDENTIFIER, 1) = False
                End If
              End If

              ' Read the Caption property from the control if required.
              If avProperties(WFITEMPROP_CAPTION, 1) Then
                
                fFixedCaption = WebFormItemHasProperty(iControlType, WFITEMPROP_CAPTION)
                If fFixedCaption Then
                  If WebFormItemHasProperty(iControlType, WFITEMPROP_CAPTIONTYPE) Then
                    fFixedCaption = (.CaptionType = giWFDATAVALUE_FIXED)
                  End If
                End If
                
                If fFixedCaption Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_CAPTION, 2)) Then
                    If msCaption <> .Caption Then
                      avProperties(WFITEMPROP_CAPTION, 2) = True
                    End If
                  Else
                    msCaption = Replace(.Caption, "&&", "&")
                  End If
                Else
                  avProperties(WFITEMPROP_CAPTION, 1) = False
                End If
              End If

              ' Read the Vertical Offset Behaviour property from the control.
              If avProperties(WFITEMPROP_VERTICALOFFSETBEHAVIOUR, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_VERTICALOFFSETBEHAVIOUR) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_VERTICALOFFSETBEHAVIOUR, 2)) Then
                    If mlngVerticalOffsetBehaviour <> .VerticalOffsetBehaviour Then
                      avProperties(WFITEMPROP_VERTICALOFFSETBEHAVIOUR, 2) = True
                    End If
                  Else
                    mlngVerticalOffsetBehaviour = .VerticalOffsetBehaviour
                  End If
                Else
                  avProperties(WFITEMPROP_VERTICALOFFSETBEHAVIOUR, 1) = False
                End If
              End If
              
              ' Read the Vertical Offset property from the control.
              If avProperties(WFITEMPROP_VERTICALOFFSET, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_VERTICALOFFSET) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_VERTICALOFFSET, 2)) Then
                      If mlngVerticalOffset(iCurrentIndex - 1) <> TwipsToPixels(.VerticalOffset) Then
                        avProperties(WFITEMPROP_VERTICALOFFSET, 2) = True
                      End If
                    End If
                    
                  If mlngVerticalOffsetBehaviour = offsetTop Then
                    mlngVerticalOffset(iCurrentIndex) = TwipsToPixels(.Top)
                    .VerticalOffset = PixelsToTwips(CDbl(mlngVerticalOffset(iCurrentIndex)))
                  Else
                    .VerticalOffset = mfrmWebForm.ScaleHeight - (.Height + .Top)
                    mlngVerticalOffset(iCurrentIndex) = TwipsToPixels(.VerticalOffset)
                  End If
                  
                Else
                  avProperties(WFITEMPROP_VERTICALOFFSET, 1) = False
                End If
              End If
              
              ' Read the Top property from the control.
              If avProperties(WFITEMPROP_TOP, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_TOP) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_TOP, 2)) Then
                    If mlngTop <> TwipsToPixels(.Top) Then
                      avProperties(WFITEMPROP_TOP, 2) = True
                    End If
                  Else
                    mlngTop = TwipsToPixels(.Top)
                    
                    If WebFormItemHasProperty(iControlType, WFITEMPROP_HEIGHTBEHAVIOUR) _
                      And mlngTop <> .Top Then
                      
                      mlngHeightBehaviour = behaveFixed
                      .HeightBehaviour = mlngHeightBehaviour
                    End If
                  End If
                Else
                  avProperties(WFITEMPROP_TOP, 1) = False
                End If
              End If

              ' Read the Horizontal Offset Behaviour property from the control.
              If avProperties(WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR, 2)) Then
                    If mlngHorizontalOffsetBehaviour <> .HorizontalOffsetBehaviour Then
                      avProperties(WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR, 2) = True
                    End If
                  Else
                    mlngHorizontalOffsetBehaviour = .HorizontalOffsetBehaviour
                  End If
                Else
                  avProperties(WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR, 1) = False
                End If
              End If
              
              ' Read the Orientation property from the control if required.
              If avProperties(WFITEMPROP_ORIENTATION, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_ORIENTATION) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_ORIENTATION, 2)) Then
                    If miOrientation <> .Alignment Then
                      avProperties(WFITEMPROP_ORIENTATION, 2) = True
                    End If
                  Else
                    miOrientation = .Alignment
                  End If
                Else
                  avProperties(WFITEMPROP_ORIENTATION, 1) = False
                End If
              End If
              
              ' Read the Horizontal Offset property from the control.
              If avProperties(WFITEMPROP_HORIZONTALOFFSET, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_HORIZONTALOFFSET) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_HORIZONTALOFFSET, 2)) Then
                    If mlngHorizontalOffset(iCurrentIndex - 1) <> TwipsToPixels(.HorizontalOffset) Then
                      avProperties(WFITEMPROP_HORIZONTALOFFSET, 2) = True
                    End If
                  End If
                  
                  If mlngHorizontalOffsetBehaviour = offsetLeft Then
                    mlngHorizontalOffset(iCurrentIndex) = TwipsToPixels(.Left)
                    .HorizontalOffset = PixelsToTwips(CDbl(mlngHorizontalOffset(iCurrentIndex)))
                  Else
                    .HorizontalOffset = mfrmWebForm.ScaleWidth - (.Width + .Left)
                    mlngHorizontalOffset(iCurrentIndex) = TwipsToPixels(.HorizontalOffset)
                  End If
                  
                  mlngHorizontalOffsetBehaviour = .HorizontalOffsetBehaviour

                Else
                  avProperties(WFITEMPROP_HORIZONTALOFFSET, 1) = False
                End If
              End If
              
              ' Read the Left property from the control.
              If avProperties(WFITEMPROP_LEFT, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_LEFT) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_LEFT, 2)) Then
                    If mlngLeft <> TwipsToPixels(.Left) Then
                      avProperties(WFITEMPROP_LEFT, 2) = True
                    End If
                  Else
                    mlngLeft = TwipsToPixels(.Left)
                    
                    If WebFormItemHasProperty(iControlType, WFITEMPROP_WIDTHBEHAVIOUR) _
                      And mlngLeft <> .Left Then
                      
                      mlngWidthBehaviour = behaveFixed
                      .WidthBehaviour = mlngWidthBehaviour
                    End If
                  End If
                Else
                  avProperties(WFITEMPROP_LEFT, 1) = False
                End If
              End If

              ' Read the Height Behaviour property from the control.
              If avProperties(WFITEMPROP_HEIGHTBEHAVIOUR, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_HEIGHTBEHAVIOUR, 2)) Then
                    If mlngHeightBehaviour <> .HeightBehaviour Then
                      avProperties(WFITEMPROP_HEIGHTBEHAVIOUR, 2) = True
                    End If
                  Else
                    mlngHeightBehaviour = .HeightBehaviour
                  End If
                Else
                  avProperties(WFITEMPROP_HEIGHTBEHAVIOUR, 1) = False
                End If
              End If
              
              ' Read the Height property from the control if required.
              If avProperties(WFITEMPROP_HEIGHT, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_HEIGHT) _
                  And (iControlType <> giWFFORMITEM_INPUTVALUE_DATE) _
                  And (iControlType <> giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                  And (iControlType <> giWFFORMITEM_INPUTVALUE_LOOKUP) _
                  And (iControlType <> giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                  And ((iControlType <> giWFFORMITEM_LINE) Or bVerticalLine) Then
                
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_HEIGHT, 2)) Then
                    If mlngHeight <> TwipsToPixels(.Height) Then
                      avProperties(WFITEMPROP_HEIGHT, 2) = True
                    End If
                  Else
                    mlngHeight = TwipsToPixels(.Height)
                  End If
                Else
                  avProperties(WFITEMPROP_HEIGHT, 1) = False
                End If
              End If

              ' Read the Width Behaviour property from the control.
              If avProperties(WFITEMPROP_WIDTHBEHAVIOUR, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_WIDTHBEHAVIOUR) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_WIDTHBEHAVIOUR, 2)) Then
                    If mlngWidthBehaviour <> .WidthBehaviour Then
                      avProperties(WFITEMPROP_WIDTHBEHAVIOUR, 2) = True
                    End If
                  Else
                    mlngWidthBehaviour = .WidthBehaviour
                  End If
                Else
                  avProperties(WFITEMPROP_WIDTHBEHAVIOUR, 1) = False
                End If
              End If
              
              ' Read the Width property from the control if required.
              If avProperties(WFITEMPROP_WIDTH, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_WIDTH) _
                  And (iControlType <> giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                  And ((iControlType <> giWFFORMITEM_LINE) Or bHorizontalLine) Then
                
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_WIDTH, 2)) Then
                    If mlngWidth <> TwipsToPixels(.Width) Then
                      avProperties(WFITEMPROP_WIDTH, 2) = True
                    End If
                  Else
                    mlngWidth = TwipsToPixels(.Width)
                  End If
                Else
                  avProperties(WFITEMPROP_WIDTH, 1) = False
                End If
              End If

              ' --------------------
              ' APPEARANCE PROPERTIES
              ' --------------------
              ' Read the BorderStyle property from the control if required.
              If avProperties(WFITEMPROP_BORDERSTYLE, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_BORDERSTYLE) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_BORDERSTYLE, 2)) Then
                    If miBorderStyle <> .BorderStyle Then
                      avProperties(WFITEMPROP_BORDERSTYLE, 2) = True
                    End If
                  Else
                    miBorderStyle = .BorderStyle
                  End If
                Else
                  avProperties(WFITEMPROP_BORDERSTYLE, 1) = False
                End If
              End If

              ' Read the Alignment property from the control if required.
              If avProperties(WFITEMPROP_ALIGNMENT, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_ALIGNMENT) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_ALIGNMENT, 2)) Then
                    If miAlignment <> .Alignment Then
                      avProperties(WFITEMPROP_ALIGNMENT, 2) = True
                    End If
                  Else
                    miAlignment = .Alignment
                  End If
                Else
                  avProperties(WFITEMPROP_ALIGNMENT, 1) = False
                End If
              End If

              ' Read the PasswordType property from the control if required.
              If avProperties(WFITEMPROP_PASSWORDTYPE, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_PASSWORDTYPE) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_PASSWORDTYPE, 2)) Then
                    If mfPasswordType <> .PasswordType Then
                      avProperties(WFITEMPROP_PASSWORDTYPE, 2) = True
                    End If
                  Else
                    mfPasswordType = .PasswordType
                  End If
                Else
                  avProperties(WFITEMPROP_PASSWORDTYPE, 1) = False
                End If
              End If

              ' Read the ColumnHeaders property from the control if required.
              If avProperties(WFITEMPROP_COLUMNHEADERS, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_COLUMNHEADERS) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_COLUMNHEADERS, 2)) Then
                    If mfColumnHeaders <> .ColumnHeaders Then
                      avProperties(WFITEMPROP_COLUMNHEADERS, 2) = True
                    End If
                  Else
                    mfColumnHeaders = .ColumnHeaders
                  End If
                Else
                  avProperties(WFITEMPROP_COLUMNHEADERS, 1) = False
                End If
              End If

              ' Read the Headlines property from the control.
              If avProperties(WFITEMPROP_HEADLINES, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_HEADLINES) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_HEADLINES, 2)) Then
                    If mlngHeadlines <> .Headlines Then
                      avProperties(WFITEMPROP_HEADLINES, 2) = True
                    End If
                  Else
                    mlngHeadlines = .Headlines
                  End If
                Else
                  avProperties(WFITEMPROP_HEADLINES, 1) = False
                End If
              End If

              ' Read the HeadFont property from the control if required.
              If avProperties(WFITEMPROP_HEADFONT, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_HEADFONT) Then
                  Set objCtlFont = .HeadFont

                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_HEADFONT, 2)) Then
                    If (objCtlFont.Name <> mObjHeadFont.Name) Or _
                      (objCtlFont.Size <> mObjHeadFont.Size) Or _
                      (objCtlFont.Bold <> mObjHeadFont.Bold) Or _
                      (objCtlFont.Italic <> mObjHeadFont.Italic) Or _
                      (objCtlFont.Strikethrough <> mObjHeadFont.Strikethrough) Or _
                      (objCtlFont.Underline <> mObjHeadFont.Underline) Then
                      
                      avProperties(WFITEMPROP_HEADFONT, 2) = True
                    End If
                  Else
                    mObjHeadFont.Name = objCtlFont.Name
                    mObjHeadFont.Size = objCtlFont.Size
                    mObjHeadFont.Bold = objCtlFont.Bold
                    mObjHeadFont.Italic = objCtlFont.Italic
                    mObjHeadFont.Strikethrough = objCtlFont.Strikethrough
                    mObjHeadFont.Underline = objCtlFont.Underline
                  End If

                  Set objCtlFont = Nothing
                Else
                  avProperties(WFITEMPROP_HEADFONT, 1) = False
                End If
              End If

              ' Read the HeaderBackColor property from the control if required.
              If avProperties(WFITEMPROP_HEADERBACKCOLOR, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_HEADERBACKCOLOR) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_HEADERBACKCOLOR, 2)) Then
                    If mColHeaderBackColor <> .HeaderBackColor Then
                      avProperties(WFITEMPROP_HEADERBACKCOLOR, 2) = True
                    End If
                  Else
                    mColHeaderBackColor = .HeaderBackColor
                  End If
                Else
                  avProperties(WFITEMPROP_HEADERBACKCOLOR, 1) = False
                End If
              End If

              ' Read the Font property from the control if required.
              If avProperties(WFITEMPROP_FONT, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_FONT) Then
                  Set objCtlFont = .Font

                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_FONT, 2)) Then
                    If (objCtlFont.Name <> mObjFont.Name) Or _
                      (objCtlFont.Size <> mObjFont.Size) Or _
                      (objCtlFont.Bold <> mObjFont.Bold) Or _
                      (objCtlFont.Italic <> mObjFont.Italic) Or _
                      (objCtlFont.Strikethrough <> mObjFont.Strikethrough) Or _
                      (objCtlFont.Underline <> mObjFont.Underline) Then
                      avProperties(WFITEMPROP_FONT, 2) = True
                    End If
                  Else

                    mObjFont.Name = objCtlFont.Name
                    mObjFont.Size = objCtlFont.Size
                    mObjFont.Bold = objCtlFont.Bold
                    mObjFont.Italic = objCtlFont.Italic
                    mObjFont.Strikethrough = objCtlFont.Strikethrough
                    mObjFont.Underline = objCtlFont.Underline
                  End If

                  Set objCtlFont = Nothing
                Else
                  avProperties(WFITEMPROP_FONT, 1) = False
                End If
              End If
              
              ' Read the ForeColor property from the control if required.
              If avProperties(WFITEMPROP_FORECOLOR, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLOR) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_FORECOLOR, 2)) Then
                    If mColForeColor <> .ForeColor Then
                      avProperties(WFITEMPROP_FORECOLOR, 2) = True
                    End If
                  Else
                    mColForeColor = .ForeColor
                  End If
                Else
                  avProperties(WFITEMPROP_FORECOLOR, 1) = False
                End If
              End If

              ' Read the ForeColorEven property from the control if required.
              If avProperties(WFITEMPROP_FORECOLOREVEN, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLOREVEN) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_FORECOLOREVEN, 2)) Then
                    If mColForeColorEven <> .ForeColorEven Then
                      avProperties(WFITEMPROP_FORECOLOREVEN, 2) = True
                    End If
                  Else
                    mColForeColorEven = .ForeColorEven
                  End If
                Else
                  avProperties(WFITEMPROP_FORECOLOREVEN, 1) = False
                End If
              End If

              ' Read the ForeColorOdd property from the control if required.
              If avProperties(WFITEMPROP_FORECOLORODD, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLORODD) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_FORECOLORODD, 2)) Then
                    If mColForeColorOdd <> .ForeColorOdd Then
                      avProperties(WFITEMPROP_FORECOLORODD, 2) = True
                    End If
                  Else
                    mColForeColorOdd = .ForeColorOdd
                  End If
                Else
                  avProperties(WFITEMPROP_FORECOLORODD, 1) = False
                End If
              End If
              
              ' Read the ForeColorHighlight property from the control if required.
              If avProperties(WFITEMPROP_FORECOLORHIGHLIGHT, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_FORECOLORHIGHLIGHT) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_FORECOLORHIGHLIGHT, 2)) Then
                    If mColForeColorHighlight <> .ForeColorHighlight Then
                      avProperties(WFITEMPROP_FORECOLORHIGHLIGHT, 2) = True
                    End If
                  Else
                    mColForeColorHighlight = .ForeColorHighlight
                  End If
                Else
                  avProperties(WFITEMPROP_FORECOLORHIGHLIGHT, 1) = False
                End If
              End If

              ' Read the BackStyle property from the control if required.
              If avProperties(WFITEMPROP_BACKSTYLE, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKSTYLE) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_BACKSTYLE, 2)) Then
                    If miBackStyle <> .BackStyle Then
                      avProperties(WFITEMPROP_BACKSTYLE, 2) = True
                    End If
                  Else
                    miBackStyle = .BackStyle
                  End If
                Else
                  avProperties(WFITEMPROP_BACKSTYLE, 1) = False
                End If
              End If

              ' Read the BackColor property from the control if required.
              If avProperties(WFITEMPROP_BACKCOLOR, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLOR) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_BACKCOLOR, 2)) Then
                    If mColBackColor <> .BackColor Then
                      avProperties(WFITEMPROP_BACKCOLOR, 2) = True
                    End If
                  Else
                    mColBackColor = .BackColor
                  End If
                Else
                  avProperties(WFITEMPROP_BACKCOLOR, 1) = False
                End If
              End If

              ' Read the BackColorEven property from the control if required.
              If avProperties(WFITEMPROP_BACKCOLOREVEN, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLOREVEN) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_BACKCOLOREVEN, 2)) Then
                    If mColBackColorEven <> .BackColorEven Then
                      avProperties(WFITEMPROP_BACKCOLOREVEN, 2) = True
                    End If
                  Else
                    mColBackColorEven = .BackColorEven
                  End If
                Else
                  avProperties(WFITEMPROP_BACKCOLOREVEN, 1) = False
                End If
              End If

              ' Read the BackColorOdd property from the control if required.
              If avProperties(WFITEMPROP_BACKCOLORODD, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLORODD) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_BACKCOLORODD, 2)) Then
                    If mColBackColorOdd <> .BackColorOdd Then
                      avProperties(WFITEMPROP_BACKCOLORODD, 2) = True
                    End If
                  Else
                    mColBackColorOdd = .BackColorOdd
                  End If
                Else
                  avProperties(WFITEMPROP_BACKCOLORODD, 1) = False
                End If
              End If

              ' Read the BackColorHighlight property from the control if required.
              If avProperties(WFITEMPROP_BACKCOLORHIGHLIGHT, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_BACKCOLORHIGHLIGHT) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_BACKCOLORHIGHLIGHT, 2)) Then
                    If mColBackColorHighlight <> .BackColorHighlight Then
                      avProperties(WFITEMPROP_BACKCOLORHIGHLIGHT, 2) = True
                    End If
                  Else
                    mColBackColorHighlight = .BackColorHighlight
                  End If
                Else
                  avProperties(WFITEMPROP_BACKCOLORHIGHLIGHT, 1) = False
                End If
              End If

              ' Read the Picture property from the control if required.
              If avProperties(WFITEMPROP_PICTURE, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_PICTURE) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_PICTURE, 2)) Then
                    If mlngPictureID <> .PictureID Then
                      avProperties(WFITEMPROP_PICTURE, 2) = True
                    End If
                  Else
                    mlngPictureID = .PictureID
                  End If
                Else
                  avProperties(WFITEMPROP_PICTURE, 1) = False
                End If
              End If

              ' Read the Picture property from the control if required.
              If avProperties(WFITEMPROP_PICTURELOCATION, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_PICTURELOCATION) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_PICTURELOCATION, 2)) Then
                    If mlngPictureLocation <> .PictureLocation Then
                      avProperties(WFITEMPROP_PICTURELOCATION, 2) = True
                    End If
                  Else
                    mlngPictureLocation = .PictureLocation
                  End If
                Else
                  avProperties(WFITEMPROP_PICTURELOCATION, 1) = False
                End If
              End If
              
              ' Read the Page Tab Number property from the control if required.
              If avProperties(WFITEMPROP_TABNUMBER, 1) Then
                If WebFormItemHasProperty(iControlType, WFITEMPROP_TABNUMBER) Then
                  If (fPassedOnce) And (Not avProperties(WFITEMPROP_TABNUMBER, 2)) Then
                    If mlngTabNumber <> .SelectedItem.Index Then
                      avProperties(WFITEMPROP_TABNUMBER, 2) = True
                    End If
                  Else
                    mlngTabNumber = .SelectedItem.Index
                  End If
                Else
                  avProperties(WFITEMPROP_TABNUMBER, 1) = False
                End If
              End If

              ' Flag that we have read a control's properties.
              fPassedOnce = True
              iLastControlType = iControlType
            End If
          End With
        End If
      Next ctlControl
      ' Disassociate object variables.
      Set ctlControl = Nothing
    Else
      If CurrentWebForm.ActiveControl.Name = "TabPages" Then
      ' The screen's tab strip is the active control.
        Me.StatusBar1.SimpleText = "Tab"
        If CurrentWebForm.tabPages.Tabs.Count > 0 Then
          
          With CurrentWebForm.tabPages
            ' Read the Caption property from the tab strip.
            avProperties(WFITEMPROP_TABCAPTION, 1) = True
            msTabCaption = Replace(.SelectedItem.Caption, "&&", "&")
            ' Read the Font property from the tab strip.
            avProperties(WFITEMPROP_FONT, 1) = True
            
            mObjFont.Name = Me.ActiveControl.Font.Name
            mObjFont.Size = Me.ActiveControl.Font.Size
            mObjFont.Bold = Me.ActiveControl.Font.Bold
            mObjFont.Italic = Me.ActiveControl.Font.Italic
            mObjFont.Strikethrough = Me.ActiveControl.Font.Strikethrough
            mObjFont.Underline = Me.ActiveControl.Font.Underline
            ' Read the tab number from the tab strip
            avProperties(WFITEMPROP_TABNUMBER, 1) = True
            avProperties(WFITEMPROP_TABNUMBER, 2) = False
            mlngTabNumber = .SelectedItem.Index
          End With
        End If
      Else
          ' --------------------
          ' The webform is the active control.
          ' --------------------
          StatusBar1.SimpleText = mfrmWebForm.Caption
          Me.Caption = "Web Form Properties"
    
          ' Initialise the property array.
          For iLoop = 0 To UBound(avProperties)
            avProperties(iLoop, 1) = WebFormItemHasProperty(giWFFORMITEM_FORM, iLoop)
          Next iLoop
       
          ' Display the appropriate properties for the
          With mfrmWebForm
            ' GENERAL
            msWFIdentifier = .WFIdentifier
            msCaption = .Caption
            mlngDescriptionExprID = .DescriptionExprID
            mfDescriptionHasWorkflowName = .DescriptionHasWorkflowName
            mfDescriptionHasElementCaption = .DescriptionHasElementCaption
            
            mlngWidth = TwipsToPixels(.Width)
            mlngHeight = TwipsToPixels(.Height)
            
            mlngTimeoutFrequency = .TimeoutFrequency
            miTimeoutPeriod = .TimeoutPeriod
            
            ' APPEARANCE
            mObjFont.Name = .Font.Name
            mObjFont.Size = .Font.Size
            mObjFont.Bold = .Font.Bold
            mObjFont.Italic = .Font.Italic
            mObjFont.Strikethrough = .Font.Strikethrough
            mObjFont.Underline = .Font.Underline
            mColForeColor = .ForeColor
    
            mColBackColor = .BackColor
            mlngPictureID = .PictureID
            mlngPictureLocation = .PictureLocation
    
            ' DATA
            '''jpd Validations - future development
          End With
        End If
      End If
      End If

  ' ++++++++++++++++++++
  ' Add the required rows to the properties grid as required.
  ' ++++++++++++++++++++
  iCurrentRow = ssGridProperties.Row

  With ssGridProperties
    ' Clear the properties grid.
    .Redraw = True
    .RemoveAll
    ' --------------------
    ' PROPERTIES SCREEN
    ' --------------------
    If avProperties(WFITEMPROP_UNKNOWN, 1) Then
      .AddItem "(Properties)" & vbTab & "" & vbTab & Str(WFITEMPROP_UNKNOWN)
    End If

    ' --------------------
    ' GENERAL PROPERTIES
    ' --------------------
    If avProperties(WFITEMPROP_WFIDENTIFIER, 1) _
      Or avProperties(WFITEMPROP_CAPTION, 1) _
      Or avProperties(WFITEMPROP_DESCRIPTION, 1) _
      Or avProperties(WFITEMPROP_ORIENTATION, 1) _
      Or avProperties(WFITEMPROP_TOP, 1) _
      Or avProperties(WFITEMPROP_WIDTH, 1) _
      Or avProperties(WFITEMPROP_LEFT, 1) _
      Or avProperties(WFITEMPROP_HEIGHT, 1) _
      Or avProperties(WFITEMPROP_TIMEOUT, 1) _
      Or avProperties(WFITEMPROP_DESCRIPTION_WORKFLOWNAME, 1) _
      Or avProperties(WFITEMPROP_DESCRIPTION_ELEMENTCAPTION, 1) Then
    
      .AddItem "General" & vbTab & "" & vbTab & Str(WFITEMPROP_NONE)
    End If

    ' Add the Workflow Identifier property row to the properties grid if required.
    If avProperties(WFITEMPROP_WFIDENTIFIER, 1) Then
      If avProperties(WFITEMPROP_WFIDENTIFIER, 2) Then
        sDescription = ""
      Else
        sDescription = msWFIdentifier
      End If
  
      .AddItem "Identifier" & vbTab & sDescription & vbTab & Str(WFITEMPROP_WFIDENTIFIER)
    End If

    ' Add the Caption property row to the properties grid if required.
    If avProperties(WFITEMPROP_CAPTION, 1) Then
      If avProperties(WFITEMPROP_CAPTION, 2) Then
        sDescription = ""
      Else
        sDescription = msCaption
      End If

      .AddItem "Caption" & vbTab & sDescription & vbTab & Str(WFITEMPROP_CAPTION)
    End If

    ' Add the Vertical Offset property row to the properties grid if required.
    If avProperties(WFITEMPROP_VERTICALOFFSET, 1) Then
      If avProperties(WFITEMPROP_VERTICALOFFSET, 2) Then
        sDescription = ""
      Else
        sDescription = Trim(Str(mlngVerticalOffset(iCurrentIndex)))
      End If
  
      .AddItem "Vertical Offset" & vbTab & sDescription & vbTab & Str(WFITEMPROP_VERTICALOFFSET)
    ElseIf avProperties(WFITEMPROP_TOP, 1) Then
      ' Add the Top property row to the properties grid if required.
      If avProperties(WFITEMPROP_TOP, 2) Then
        sDescription = ""
      Else
        sDescription = Trim(Str(mlngTop))
      End If
  
      .AddItem "Top" & vbTab & sDescription & vbTab & Str(WFITEMPROP_TOP)
    End If

    ' Add the Vertical Offset Behaviour property row to the properties grid if required.
    If avProperties(WFITEMPROP_VERTICALOFFSETBEHAVIOUR, 1) Then
      If avProperties(WFITEMPROP_VERTICALOFFSETBEHAVIOUR, 2) Then
        sDescription = ""
      Else
        If mlngVerticalOffsetBehaviour = offsetTop Then
          sDescription = msVERTICALFROMTOP
        Else
          sDescription = msVERTICALFROMBOTTOM
        End If
      End If
      
      .AddItem "Vertical From" & vbTab & sDescription & vbTab & Str(WFITEMPROP_VERTICALOFFSETBEHAVIOUR)
    End If
    
    ' Add the Orientation property row to the properties grid if required.
    If avProperties(WFITEMPROP_ORIENTATION, 1) Then
      If avProperties(WFITEMPROP_ORIENTATION, 2) Then
        sDescription = ""
      Else
        If miOrientation = 0 Then
          sDescription = msORIENTATIONVERTICAL
        Else
          sDescription = msORIENTATIONHORIZONTAL
        End If
      End If
      
      .AddItem "Orientation" & vbTab & sDescription & vbTab & Str(WFITEMPROP_ORIENTATION)
    End If
  
    ' Add the Left property row to the properties grid if required.
    If avProperties(WFITEMPROP_HORIZONTALOFFSET, 1) Then
      If avProperties(WFITEMPROP_HORIZONTALOFFSET, 2) Then
        sDescription = ""
      Else
        sDescription = Trim(Str(mlngHorizontalOffset(iCurrentIndex)))
      End If
  
      .AddItem "Horizontal Offset" & vbTab & sDescription & vbTab & Str(WFITEMPROP_HORIZONTALOFFSET)
    ElseIf avProperties(WFITEMPROP_LEFT, 1) Then
    ' Add the Left property row to the properties grid if required.
      If avProperties(WFITEMPROP_LEFT, 2) Then
        sDescription = ""
      Else
        sDescription = Trim(Str(mlngLeft))
      End If
  
      .AddItem "Left" & vbTab & sDescription & vbTab & Str(WFITEMPROP_LEFT)
    End If
  
    ' Add the Horizontal Offset Behaviour property row to the properties grid if required.
    If avProperties(WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR, 1) Then
      If avProperties(WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR, 2) Then
        sDescription = ""
      Else
        If mlngHorizontalOffsetBehaviour = offsetLeft Then
          sDescription = msHORIZONTALFROMLEFT
        Else
          sDescription = msHORIZONTALFROMRIGHT
        End If
      End If
      
      .AddItem "Horizontal From" & vbTab & sDescription & vbTab & Str(WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR)
    End If
    
    ' Add the Height Behaviour property row to the properties grid if required.
    If avProperties(WFITEMPROP_HEIGHTBEHAVIOUR, 1) Then
      If avProperties(WFITEMPROP_HEIGHTBEHAVIOUR, 2) Then
        sDescription = ""
      Else
        If mlngHeightBehaviour = behaveFixed Then
          sDescription = msBEHAVEFIXED
        Else
          sDescription = msBEHAVEFULL
        End If
      End If
  
      .AddItem "Height Behaviour" & vbTab & sDescription & vbTab & Str(WFITEMPROP_HEIGHTBEHAVIOUR)
    End If
    
    ' Add the Height property row to the properties grid if required.
    If avProperties(WFITEMPROP_HEIGHT, 1) Then
      If avProperties(WFITEMPROP_HEIGHT, 2) Then
        sDescription = ""
      Else
        sDescription = Trim(Str(mlngHeight))
      End If
  
      .AddItem "Height" & vbTab & sDescription & vbTab & Str(WFITEMPROP_HEIGHT)
    End If

    ' Add the Width Behaviour property row to the properties grid if required.
    If avProperties(WFITEMPROP_WIDTHBEHAVIOUR, 1) Then
      If avProperties(WFITEMPROP_WIDTHBEHAVIOUR, 2) Then
        sDescription = ""
      Else
        If mlngWidthBehaviour = behaveFixed Then
          sDescription = msBEHAVEFIXED
        Else
          sDescription = msBEHAVEFULL
        End If
      End If
  
      .AddItem "Width Behaviour" & vbTab & sDescription & vbTab & Str(WFITEMPROP_WIDTHBEHAVIOUR)
    End If
    
    ' Add the Width property row to the properties grid if required.
    If avProperties(WFITEMPROP_WIDTH, 1) Then
      If avProperties(WFITEMPROP_WIDTH, 2) Then
        sDescription = ""
      Else
        sDescription = Trim(Str(mlngWidth))
      End If
  
      .AddItem "Width" & vbTab & sDescription & vbTab & Str(WFITEMPROP_WIDTH)
    End If
    
    ' Add the Page Number property row to the properties grid if required.
    If avProperties(WFITEMPROP_TABNUMBER, 1) Then
      If avProperties(WFITEMPROP_TABNUMBER, 2) Then
        sDescription = ""
      Else
        sDescription = Trim(Str(mlngTabNumber))
      End If
      
      ssGridProperties.AddItem "Page Tab Order" & vbTab & sDescription & vbTab & Str(WFITEMPROP_TABNUMBER)
      avProperties(WFITEMPROP_TABNUMBER, 2) = ssGridProperties.Rows - 1
    End If
    
    
    ' Add the Caption property row to the properties grid if required.
    If avProperties(WFITEMPROP_TABCAPTION, 1) Then
      If avProperties(WFITEMPROP_TABCAPTION, 2) Then
        sDescription = ""
      Else
        sDescription = msTabCaption
      End If
      .AddItem "Caption" & vbTab & sDescription & vbTab & Str(WFITEMPROP_TABCAPTION)
    End If
    
    
    ' --------------------
    ' APPEARANCE PROPERTIES
    ' --------------------
    If avProperties(WFITEMPROP_BORDERSTYLE, 1) _
      Or avProperties(WFITEMPROP_ALIGNMENT, 1) _
      Or avProperties(WFITEMPROP_PASSWORDTYPE, 1) _
      Or avProperties(WFITEMPROP_COLUMNHEADERS, 1) _
      Or avProperties(WFITEMPROP_HEADLINES, 1) _
      Or avProperties(WFITEMPROP_HEADFONT, 1) _
      Or avProperties(WFITEMPROP_HEADERBACKCOLOR, 1) _
      Or avProperties(WFITEMPROP_FONT, 1) _
      Or avProperties(WFITEMPROP_FORECOLOR, 1) _
      Or avProperties(WFITEMPROP_FORECOLOREVEN, 1) _
      Or avProperties(WFITEMPROP_FORECOLORODD, 1) _
      Or avProperties(WFITEMPROP_FORECOLORHIGHLIGHT, 1) _
      Or avProperties(WFITEMPROP_BACKSTYLE, 1) _
      Or avProperties(WFITEMPROP_BACKCOLOR, 1) _
      Or avProperties(WFITEMPROP_BACKCOLOREVEN, 1) _
      Or avProperties(WFITEMPROP_BACKCOLORODD, 1) _
      Or avProperties(WFITEMPROP_BACKCOLORHIGHLIGHT, 1) _
      Or avProperties(WFITEMPROP_PICTURE, 1) _
      Or avProperties(WFITEMPROP_PICTURELOCATION, 1) Then
    
      .AddItem "Appearance" & vbTab & "" & vbTab & Str(WFITEMPROP_NONE)
    End If

    ' Add the BorderStyle property row to the properties grid if required.
    If avProperties(WFITEMPROP_BORDERSTYLE, 1) Then
      If avProperties(WFITEMPROP_BORDERSTYLE, 2) Then
        sDescription = ""
      Else
        If miBorderStyle = vbBSNone Then
          sDescription = msBORDERSTYLENONETEXT
        Else
          sDescription = msBORDERSTYLEFIXEDSINGLETEXT
        End If
      End If
  
      .AddItem "Border" & vbTab & sDescription & vbTab & Str(WFITEMPROP_BORDERSTYLE)
    End If
  
    ' Add the Alignment property row to the properties grid if required.
    If avProperties(WFITEMPROP_ALIGNMENT, 1) Then
      If avProperties(WFITEMPROP_ALIGNMENT, 2) Then
        sDescription = ""
      Else
        If miAlignment = vbLeftJustify Then
          sDescription = msALIGNMENTLEFTTEXT
        Else
          sDescription = msALIGNMENTRIGHTTEXT
        End If
      End If
  
      .AddItem "Alignment" & vbTab & sDescription & vbTab & Str(WFITEMPROP_ALIGNMENT)
    End If
  
    ' Add the PASSWORDTYPE property row to the properties grid if required.
    If avProperties(WFITEMPROP_PASSWORDTYPE, 1) Then
      If avProperties(WFITEMPROP_PASSWORDTYPE, 2) Then
        sDescription = ""
      Else
        If mfPasswordType = True Then
          sDescription = msBOOLEAN_TRUE
        Else
          sDescription = msBOOLEAN_FALSE
        End If
      End If
      .AddItem "Hide Text" & vbTab & sDescription & vbTab & Str(WFITEMPROP_PASSWORDTYPE)
    End If
    
    ' Add the ColumnHeaders property row to the properties grid if required.
    If avProperties(WFITEMPROP_COLUMNHEADERS, 1) Then
      If avProperties(WFITEMPROP_COLUMNHEADERS, 2) Then
        sDescription = ""
      Else
        If mfColumnHeaders = True Then
          sDescription = msBOOLEAN_TRUE
        Else
          sDescription = msBOOLEAN_FALSE
        End If
      End If
      .AddItem "Column Headers" & vbTab & sDescription & vbTab & Str(WFITEMPROP_COLUMNHEADERS)
    End If
    
    ' Add the Headlines property row to the properties grid if required.
    If avProperties(WFITEMPROP_HEADLINES, 1) Then
      If avProperties(WFITEMPROP_HEADLINES, 2) Then
        sDescription = ""
      Else
        sDescription = CStr(mlngHeadlines)
      End If
      .AddItem "Header Lines" & vbTab & sDescription & vbTab & Str(WFITEMPROP_HEADLINES)
    End If
  
    ' Add the HeadFont property row to the properties grid if required.
    If avProperties(WFITEMPROP_HEADFONT, 1) Then
      If avProperties(WFITEMPROP_HEADFONT, 2) Then
        sDescription = ""
      Else
        sDescription = GetHeadFontDescription
      End If
  
      .AddItem "Header Font" & vbTab & sDescription & vbTab & Str(WFITEMPROP_HEADFONT)
    End If
  
    ' Add the HeaderBackColor property row to the properties grid if required.
    If avProperties(WFITEMPROP_HEADERBACKCOLOR, 1) Then
      If avProperties(WFITEMPROP_HEADERBACKCOLOR, 2) Then
        ssGridProperties.StyleSets("ssetHeaderBackColor").BackColor = vbWhite
      Else
        ssGridProperties.StyleSets("ssetHeaderBackColor").BackColor = mColHeaderBackColor
      End If
  
      .AddItem "Header Background Colour" & vbTab & "" & vbTab & Str(WFITEMPROP_HEADERBACKCOLOR)
    End If
  
    ' Add the Font property row to the properties grid if required.
    If avProperties(WFITEMPROP_FONT, 1) Then
      If avProperties(WFITEMPROP_FONT, 2) Then
        sDescription = ""
      Else
        sDescription = GetFontDescription
      End If
  
      .AddItem "Font" & vbTab & sDescription & vbTab & Str(WFITEMPROP_FONT)
    End If
  
    ' Add the ForeColor property row to the properties grid if required.
    If avProperties(WFITEMPROP_FORECOLOR, 1) Then
      If avProperties(WFITEMPROP_FORECOLOR, 2) Then
        .StyleSets("ssetForeColorValue").BackColor = vbWhite
      Else
        .StyleSets("ssetForeColorValue").BackColor = mColForeColor
      End If

      .AddItem "Foreground Colour" & vbTab & "" & vbTab & Str(WFITEMPROP_FORECOLOR)
    End If

    ' Add the ForeColorEven property row to the properties grid if required.
    If avProperties(WFITEMPROP_FORECOLOREVEN, 1) Then
      If avProperties(WFITEMPROP_FORECOLOREVEN, 2) Then
        .StyleSets("ssetForeColorEven").BackColor = vbWhite
      Else
        .StyleSets("ssetForeColorEven").BackColor = mColForeColorEven
      End If

      .AddItem "Foreground Colour Even" & vbTab & "" & vbTab & Str(WFITEMPROP_FORECOLOREVEN)
    End If

    ' Add the ForeColorOdd property row to the properties grid if required.
    If avProperties(WFITEMPROP_FORECOLORODD, 1) Then
      If avProperties(WFITEMPROP_FORECOLORODD, 2) Then
        .StyleSets("ssetForeColorOdd").BackColor = vbWhite
      Else
        .StyleSets("ssetForeColorOdd").BackColor = mColForeColorOdd
      End If

      .AddItem "Foreground Colour Odd" & vbTab & "" & vbTab & Str(WFITEMPROP_FORECOLORODD)
    End If

    ' Add the ForeColorHighlight property row to the properties grid if required.
    If avProperties(WFITEMPROP_FORECOLORHIGHLIGHT, 1) Then
      If avProperties(WFITEMPROP_FORECOLORHIGHLIGHT, 2) Then
        .StyleSets("ssetForeColorHighlight").BackColor = vbWhite
      Else
        .StyleSets("ssetForeColorHighlight").BackColor = mColForeColorHighlight
      End If
  
      .AddItem "Foreground Colour Highlight" & vbTab & "" & vbTab & Str(WFITEMPROP_FORECOLORHIGHLIGHT)
    End If
  
    ' Add the BackStyle property row to the properties grid if required.
    If avProperties(WFITEMPROP_BACKSTYLE, 1) Then
      If avProperties(WFITEMPROP_BACKSTYLE, 2) Then
        sDescription = ""
      Else
        If miBackStyle = BACKSTYLE_OPAQUE Then
          sDescription = msBACKSTYLEOPAQUETEXT
        Else
          sDescription = msBACKSTYLETRANSPARENTTEXT
        End If
      End If
      ' AE20080306 Fault #12971
      '.AddItem "BackStyle" & vbTab & sDescription & vbTab & Str(WFITEMPROP_BACKSTYLE)
      .AddItem "Background Style" & vbTab & sDescription & vbTab & Str(WFITEMPROP_BACKSTYLE)
    End If

    ' Add the BackColor property row to the properties grid if required.
    If avProperties(WFITEMPROP_BACKCOLOR, 1) Then
      If avProperties(WFITEMPROP_BACKCOLOR, 2) Then
        .StyleSets("ssetBackColorValue").BackColor = vbWhite
      Else
        .StyleSets("ssetBackColorValue").BackColor = mColBackColor
      End If
  
      .AddItem "Background Colour" & vbTab & "" & vbTab & Str(WFITEMPROP_BACKCOLOR)
    End If
  
    ' Add the BackColorEven property row to the properties grid if required.
    If avProperties(WFITEMPROP_BACKCOLOREVEN, 1) Then
      If avProperties(WFITEMPROP_BACKCOLOREVEN, 2) Then
        .StyleSets("ssetBackColorEven").BackColor = vbWhite
      Else
        .StyleSets("ssetBackColorEven").BackColor = mColBackColorEven
      End If
  
      .AddItem "Background Colour Even" & vbTab & "" & vbTab & Str(WFITEMPROP_BACKCOLOREVEN)
    End If
  
    ' Add the BackColorOdd property row to the properties grid if required.
    If avProperties(WFITEMPROP_BACKCOLORODD, 1) Then
      If avProperties(WFITEMPROP_BACKCOLORODD, 2) Then
        .StyleSets("ssetBackColorOdd").BackColor = vbWhite
      Else
        .StyleSets("ssetBackColorOdd").BackColor = mColBackColorOdd
      End If
  
      .AddItem "Background Colour Odd" & vbTab & "" & vbTab & Str(WFITEMPROP_BACKCOLORODD)
    End If
  
    ' Add the BackColorHighlight property row to the properties grid if required.
    If avProperties(WFITEMPROP_BACKCOLORHIGHLIGHT, 1) Then
      If avProperties(WFITEMPROP_BACKCOLORHIGHLIGHT, 2) Then
        .StyleSets("ssetBackColorHighlight").BackColor = vbWhite
      Else
        .StyleSets("ssetBackColorHighlight").BackColor = mColBackColorHighlight
      End If
  
      .AddItem "Background Colour Highlight" & vbTab & "" & vbTab & Str(WFITEMPROP_BACKCOLORHIGHLIGHT)
    End If
  
    ' Add the Picture property row to the properties grid if required.
    If avProperties(WFITEMPROP_PICTURE, 1) Then
      If avProperties(WFITEMPROP_PICTURE, 2) Then
        sDescription = ""
      Else
        If mlngPictureID = 0 Then
          sDescription = msPICTURENONETEXT
        Else
          With recPictEdit
            .Index = "idxID"
            .Seek "=", mlngPictureID
            If Not .NoMatch Then
              sDescription = !Name
            Else
              sDescription = msPICTURESELECTEDTEXT
            End If
          End With
        End If
      End If
  
      .AddItem "Picture" & vbTab & sDescription & vbTab & Str(WFITEMPROP_PICTURE)
    End If
  
    ' Add the Location property row to the properties grid if required.
    If avProperties(WFITEMPROP_PICTURELOCATION, 1) Then
      If avProperties(WFITEMPROP_PICTURELOCATION, 2) Then
        sDescription = ""
      Else
        Select Case mlngPictureLocation
          Case 0
            sDescription = PICLOC_TOPLEFT
          Case 1
            sDescription = PICLOC_TOPRIGHT
          Case 2
            sDescription = PICLOC_CENTRE
          Case 3
            sDescription = PICLOC_LEFTTILE
          Case 4
            sDescription = PICLOC_RIGHTTILE
          Case 5
            sDescription = PICLOC_TOPTILE
          Case 6
            sDescription = PICLOC_BOTTOMTILE
          Case 7
            sDescription = PICLOC_TILE
          Case Else
            sDescription = PICLOC_TOPLEFT
        End Select
      End If
  
      .AddItem "Location" & vbTab & sDescription & vbTab & Str(WFITEMPROP_PICTURELOCATION)
    End If
    
    .Redraw = True
  End With

  'Set the row number back to what it was before.
  If pfStayOnSameLine Then
    ssGridProperties.Row = iCurrentRow
  Else
    'JPD 20060704 Fault 11251
    ' Reset the row style to avoid dropdowns or buttons being shown on the wrong rows.
    ssGridProperties.Columns(1).Style = ssStyleEdit
  End If

  If Not (mfrmWebForm Is Nothing) Then
    ssGridProperties.Enabled = (Not mfrmWebForm.ReadOnly)
  End If

  ' Ensure the grid has the correct column properties applied for the selected property.
  ' Note that we do this here even if the selected row/col has not changed, as the property
  ' displayed in the selected row may have.
  If frmSysMgr.ActiveForm Is Me Then
    ssGridProperties_RowColChange 0, 0
  End If
  
  ' Get rid of the icon off the form
  Me.Icon = Nothing
  SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME
  
TidyUpAndExit:
  Set objCtlFont = Nothing
  Set ctlControl = Nothing
  RefreshProperties = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function UpdateControls(piProperty As WFItemProperty) As Boolean
  
  ' Update the screen controls with the given property value from the grid.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim sFileName As String
  Dim objCurrentFont As StdFont
  Dim objFont As StdFont
  Dim ctlControl As VB.Control
  Dim blnDontRefresh As Boolean
  Dim iCount As Integer
  Dim fChangeMade As Boolean
  Dim sOriginalIdentifier As String
  Dim lngTargetPageNumber As Long
  Dim sCaption As String
  
  fChangeMade = False
  fOK = Not (mfrmWebForm Is Nothing)

  If fOK Then
    If UBound(mactlSelectedControls) > 0 Then
      For iIndex = 1 To UBound(mactlSelectedControls)
        Set ctlControl = mactlSelectedControls(iIndex)
        With ctlControl
          ' Update the control with the new property value(s).
          Select Case piProperty
            ' --------------------
            ' GENERAL PROPERTIES
            ' --------------------
            Case WFITEMPROP_WFIDENTIFIER
              sOriginalIdentifier = .WFIdentifier
              msWFIdentifier = CStr(Left(msWFIdentifier, 200))
              If (.WFIdentifier <> msWFIdentifier) Then
                fChangeMade = True
              End If
              .WFIdentifier = msWFIdentifier
              mfrmWebForm.UpdateIdentifiers False, _
                sOriginalIdentifier, _
                msWFIdentifier, _
                0, _
                0

            Case WFITEMPROP_CAPTION
              msCaption = CStr(Left(msCaption, 200))
              If (.Caption <> msCaption) Then
                fChangeMade = True
              End If
              .Caption = Replace(msCaption, "&", "&&")

              If AutoResizeControl(ctlControl) Then
                fChangeMade = True
              End If

            Case WFITEMPROP_ORIENTATION
              If miOrientation <> .Alignment Then
                .Alignment = miOrientation
              End If
              
            Case WFITEMPROP_VERTICALOFFSET
              If (.VerticalOffset <> PixelsToTwips(CDbl(mlngVerticalOffset(iIndex)))) Then
                fChangeMade = True
                
                .VerticalOffset = PixelsToTwips(CDbl(mlngVerticalOffset(iIndex)))
                
                If .VerticalOffsetBehaviour = offsetTop Then
                  .Top = PixelsToTwips(CDbl(mlngVerticalOffset(iIndex)))
                Else
                  .Top = mfrmWebForm.ScaleHeight - (PixelsToTwips(CDbl(mlngVerticalOffset(iIndex))) + .Height)
                End If
              End If
              
            Case WFITEMPROP_VERTICALOFFSETBEHAVIOUR
              If (.VerticalOffsetBehaviour <> mlngVerticalOffsetBehaviour) Then
                fChangeMade = True
                            
                If mlngVerticalOffsetBehaviour = offsetTop Then
                  mlngVerticalOffset(iIndex) = TwipsToPixels(mfrmWebForm.ScaleHeight) - (TwipsToPixels(.Height) + mlngVerticalOffset(iIndex))
                  .VerticalOffset = PixelsToTwips(CDbl(mlngVerticalOffset(iIndex)))
                  .Top = .VerticalOffset
                Else
                  .VerticalOffset = mfrmWebForm.ScaleHeight - (.Height + PixelsToTwips(CDbl(mlngVerticalOffset(iIndex))))
                  .Top = PixelsToTwips(CDbl(mlngVerticalOffset(iIndex)))
                  mlngVerticalOffset(iIndex) = TwipsToPixels(.VerticalOffset)
                End If
                
                .VerticalOffsetBehaviour = mlngVerticalOffsetBehaviour
              End If
              
            Case WFITEMPROP_TOP
              If (.Top <> PixelsToTwips(CDbl(mlngTop))) Then
                fChangeMade = True
              End If
              .Top = PixelsToTwips(CDbl(mlngTop))

            Case WFITEMPROP_HORIZONTALOFFSET
              If (.HorizontalOffset <> PixelsToTwips(CDbl(mlngHorizontalOffset(iIndex)))) Then
                fChangeMade = True
              
                .HorizontalOffset = PixelsToTwips(CDbl(mlngHorizontalOffset(iIndex)))
                
                If .HorizontalOffsetBehaviour = offsetLeft Then
                  .Left = .HorizontalOffset
                Else
                  .Left = mfrmWebForm.ScaleWidth - .Width - .HorizontalOffset
                End If
              End If
            
            Case WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR
              If (.HorizontalOffsetBehaviour <> mlngHorizontalOffsetBehaviour) Then
                fChangeMade = True
                            
                If mlngHorizontalOffsetBehaviour = offsetLeft Then
                  mlngHorizontalOffset(iIndex) = TwipsToPixels(mfrmWebForm.ScaleWidth) - (TwipsToPixels(.Width) + mlngHorizontalOffset(iIndex))
                  .HorizontalOffset = PixelsToTwips(CDbl(mlngHorizontalOffset(iIndex)))
                  .Left = .HorizontalOffset
                Else
                  .HorizontalOffset = mfrmWebForm.ScaleWidth - (.Width + PixelsToTwips(CDbl(mlngHorizontalOffset(iIndex))))
                  .Left = PixelsToTwips(CDbl(mlngHorizontalOffset(iIndex)))
                  mlngHorizontalOffset(iIndex) = TwipsToPixels(.HorizontalOffset)
                End If
                
                .HorizontalOffsetBehaviour = mlngHorizontalOffsetBehaviour
              End If
            
            Case WFITEMPROP_LEFT
              If (.Left <> PixelsToTwips(CDbl(mlngLeft))) Then
                fChangeMade = True
              End If
              .Left = PixelsToTwips(CDbl(mlngLeft))

            Case WFITEMPROP_HEIGHTBEHAVIOUR
              If (.HeightBehaviour <> mlngHeightBehaviour) Then
                fChangeMade = True
                            
                If Not (mlngHeightBehaviour = behaveFixed) Then
                  mlngTop = 0
                  mlngHeight = TwipsToPixels(mfrmWebForm.ScaleHeight)
                                    
                  .Top = mlngTop
                  .Height = PixelsToTwips(CDbl(mlngHeight))
                End If
                
                .HeightBehaviour = mlngHeightBehaviour
              End If

            Case WFITEMPROP_HEIGHT
              If (.Height <> PixelsToTwips(CDbl(mlngHeight))) Then
                If WebFormItemHasProperty(.WFItemType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
                  .HeightBehaviour = behaveFixed
                End If
                
                fChangeMade = True
              End If
              .Height = PixelsToTwips(CDbl(mlngHeight))
              
            Case WFITEMPROP_WIDTHBEHAVIOUR
              If (.WidthBehaviour <> mlngWidthBehaviour) Then
                fChangeMade = True
                            
                If Not (mlngWidthBehaviour = behaveFixed) Then
                  mlngLeft = 0
                  mlngWidth = TwipsToPixels(mfrmWebForm.ScaleWidth)
                                    
                  .Left = mlngLeft
                  .Width = PixelsToTwips(CDbl(mlngWidth))
                End If
                
                .WidthBehaviour = mlngWidthBehaviour
              End If
              
            Case WFITEMPROP_WIDTH
              If (.Width <> PixelsToTwips(CDbl(mlngWidth))) Then
                If WebFormItemHasProperty(.WFItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then
                  .WidthBehaviour = behaveFixed
                End If
                
                fChangeMade = True
              End If
              .Width = PixelsToTwips(CDbl(mlngWidth))
              
            ' --------------------
            ' APPEARANCE PROPERTIES
            ' --------------------
            Case WFITEMPROP_BORDERSTYLE
              If (.BorderStyle <> miBorderStyle) Then
                fChangeMade = True
              End If
              .BorderStyle = miBorderStyle
'''              blnDontRefresh = True
              If AutoResizeControl(ctlControl) Then
                fChangeMade = True
              End If
              
            Case WFITEMPROP_ALIGNMENT
              If (.Alignment <> miAlignment) Then
                fChangeMade = True
              End If
              .Alignment = miAlignment
'''              blnDontRefresh = True

            Case WFITEMPROP_PASSWORDTYPE ' Hide text
              If (.PasswordType <> mfPasswordType) Then
                fChangeMade = True
              End If
              .PasswordType = mfPasswordType

            Case WFITEMPROP_COLUMNHEADERS
              If (.ColumnHeaders <> mfColumnHeaders) Then
                fChangeMade = True
              End If
              .ColumnHeaders = mfColumnHeaders

            Case WFITEMPROP_HEADLINES
              If (.Headlines <> mlngHeadlines) Then
                fChangeMade = True
              End If
              .Headlines = mlngHeadlines

            Case WFITEMPROP_HEADFONT
              Set objFont = New StdFont
              objFont.Name = mObjHeadFont.Name
              objFont.Size = mObjHeadFont.Size
              objFont.Bold = mObjHeadFont.Bold
              objFont.Italic = mObjHeadFont.Italic
              objFont.Strikethrough = mObjHeadFont.Strikethrough
              objFont.Underline = mObjHeadFont.Underline
              
              Set objCurrentFont = .HeadFont
              If (objCurrentFont.Name <> objFont.Name) _
                Or (objCurrentFont.Size <> objFont.Size) _
                Or (objCurrentFont.Bold <> objFont.Bold) _
                Or (objCurrentFont.Italic <> objFont.Italic) _
                Or (objCurrentFont.Strikethrough <> objFont.Strikethrough) _
                Or (objCurrentFont.Underline <> objFont.Underline) Then
              
                fChangeMade = True
              End If
              Set objCurrentFont = Nothing
              
              Set .HeadFont = mObjHeadFont
              Set objFont = Nothing

            Case WFITEMPROP_HEADERBACKCOLOR
              If (.HeaderBackColor <> mColHeaderBackColor) Then
                fChangeMade = True
              End If
              .HeaderBackColor = mColHeaderBackColor

            Case WFITEMPROP_FONT
              Set objFont = New StdFont
              objFont.Name = mObjFont.Name
              objFont.Size = mObjFont.Size
              objFont.Bold = mObjFont.Bold
              objFont.Italic = mObjFont.Italic
              objFont.Strikethrough = mObjFont.Strikethrough
              objFont.Underline = mObjFont.Underline
              
              Set objCurrentFont = .Font
              If (objCurrentFont.Name <> objFont.Name) _
                Or (objCurrentFont.Size <> objFont.Size) _
                Or (objCurrentFont.Bold <> objFont.Bold) _
                Or (objCurrentFont.Italic <> objFont.Italic) _
                Or (objCurrentFont.Strikethrough <> objFont.Strikethrough) _
                Or (objCurrentFont.Underline <> objFont.Underline) Then
              
                fChangeMade = True
              End If
              Set objCurrentFont = Nothing
              
              Set .Font = objFont
              Set objFont = Nothing

              If AutoResizeControl(ctlControl) Then
                fChangeMade = True
              End If

            Case WFITEMPROP_FORECOLOR
              If (.ForeColor <> mColForeColor) Then
                fChangeMade = True
              End If
              .ForeColor = mColForeColor

            Case WFITEMPROP_FORECOLOREVEN
              If (.ForeColorEven <> mColForeColorEven) Then
                fChangeMade = True
              End If
              .ForeColorEven = mColForeColorEven

            Case WFITEMPROP_FORECOLORODD
              If (.ForeColorOdd <> mColForeColorOdd) Then
                fChangeMade = True
              End If
              .ForeColorOdd = mColForeColorOdd

            Case WFITEMPROP_FORECOLORHIGHLIGHT
              If (.ForeColorHighlight <> mColForeColorHighlight) Then
                fChangeMade = True
              End If
              .ForeColorHighlight = mColForeColorHighlight

            Case WFITEMPROP_BACKSTYLE
              If (.BackStyle <> miBackStyle) Then
                fChangeMade = True
              End If
              .BackStyle = miBackStyle
'''              blnDontRefresh = True

            Case WFITEMPROP_BACKCOLOR
              If (.BackColor <> mColBackColor) Then
                fChangeMade = True
              End If
              .BackColor = mColBackColor

            Case WFITEMPROP_BACKCOLOREVEN
              If (.BackColorEven <> mColBackColorEven) Then
                fChangeMade = True
              End If
              .BackColorEven = mColBackColorEven

            Case WFITEMPROP_BACKCOLORODD
              If (.BackColorOdd <> mColBackColorOdd) Then
                fChangeMade = True
              End If
              .BackColorOdd = mColBackColorOdd

            Case WFITEMPROP_BACKCOLORHIGHLIGHT
              If (.BackColorHighlight <> mColBackColorHighlight) Then
                fChangeMade = True
              End If
              .BackColorHighlight = mColBackColorHighlight

            Case WFITEMPROP_PICTURE
              sFileName = ""
              If mlngPictureID > 0 Then
                recPictEdit.Index = "idxID"
                recPictEdit.Seek "=", mlngPictureID

                If recPictEdit.NoMatch Then
                  mlngPictureID = 0
                Else
                  sFileName = ReadPicture
                  .PictureID = mlngPictureID
                  .Picture = sFileName
                  Kill sFileName
                End If
              End If
              
              If (.PictureID <> mlngPictureID) Then
                fChangeMade = True
              End If
              
              If mlngPictureID = 0 Then
                .Picture = sFileName
                .PictureID = mlngPictureID
              End If

            Case WFITEMPROP_PICTURELOCATION
              ' Only available for forms (see below).

          End Select
        End With

        ' Refresh the selection markers
        For iCount = 1 To mfrmWebForm.ASRSelectionMarkers.Count - 1
          With mfrmWebForm.ASRSelectionMarkers(iCount)
            If .Visible Then
              .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize, .AttachedObject.Width + (.MarkerSize * 2), .AttachedObject.Height + (.MarkerSize * 2)
              .RefreshSelectionMarkers True
            End If
          End With
        Next iCount

        ' If tab page is selected then force a refresh.
        If mfrmWebForm.tabPages.Selected Then
          mfrmWebForm.DockPagesToTabStrip
        End If

        If Not fOK Then
          Exit For
        End If
      Next iIndex
    Else
      ' Update the form control with the new properties
      With mfrmWebForm
        Select Case piProperty
          ' --------------------
          ' GENERAL PROPERTIES
          ' --------------------
          Case WFITEMPROP_WFIDENTIFIER
            sOriginalIdentifier = .WFIdentifier
            If (.WFIdentifier <> msWFIdentifier) Then
              fChangeMade = True
            End If
            .WFIdentifier = msWFIdentifier
            mfrmWebForm.UpdateIdentifiers True, _
              sOriginalIdentifier, _
              msWFIdentifier, _
              0, _
              0
            
          Case WFITEMPROP_CAPTION
            If (.Caption <> msCaption) Then
              fChangeMade = True
            End If
            SetFormCaption mfrmWebForm, msCaption
            
          Case WFITEMPROP_DESCRIPTION
            If (.DescriptionExprID <> mlngDescriptionExprID) Then
              fChangeMade = True
            End If
            .DescriptionExprID = mlngDescriptionExprID
            
          Case WFITEMPROP_DESCRIPTION_WORKFLOWNAME
            If (.DescriptionHasWorkflowName <> mfDescriptionHasWorkflowName) Then
              fChangeMade = True
            End If
            .DescriptionHasWorkflowName = mfDescriptionHasWorkflowName
        
          Case WFITEMPROP_DESCRIPTION_ELEMENTCAPTION
            If (.DescriptionHasElementCaption <> mfDescriptionHasElementCaption) Then
              fChangeMade = True
            End If
            .DescriptionHasElementCaption = mfDescriptionHasElementCaption
          
          Case WFITEMPROP_WIDTH
            If (.Width <> PixelsToTwips(CDbl(mlngWidth))) Then
              fChangeMade = True
            End If
            .Width = PixelsToTwips(CDbl(mlngWidth))
            
          Case WFITEMPROP_HEIGHT
            If (.Height <> PixelsToTwips(CDbl(mlngHeight))) Then
              fChangeMade = True
            End If
            .Height = PixelsToTwips(CDbl(mlngHeight))
            
          Case WFITEMPROP_TIMEOUT
            If (.TimeoutFrequency <> mlngTimeoutFrequency) _
              Or (.TimeoutPeriod <> miTimeoutPeriod) Then
              fChangeMade = True
            End If
            .TimeoutFrequency = mlngTimeoutFrequency
            .TimeoutPeriod = miTimeoutPeriod
        
          ' --------------------
          ' APPEARANCE PROPERTIES
          ' --------------------
          Case WFITEMPROP_FONT
            Set objFont = New StdFont
            objFont.Name = mObjFont.Name
            objFont.Size = mObjFont.Size
            objFont.Bold = mObjFont.Bold
            objFont.Italic = mObjFont.Italic
            objFont.Strikethrough = mObjFont.Strikethrough
            objFont.Underline = mObjFont.Underline
            
            Set objCurrentFont = .Font
            If (objCurrentFont.Name <> objFont.Name) _
              Or (objCurrentFont.Size <> objFont.Size) _
              Or (objCurrentFont.Bold <> objFont.Bold) _
              Or (objCurrentFont.Italic <> objFont.Italic) _
              Or (objCurrentFont.Strikethrough <> objFont.Strikethrough) _
              Or (objCurrentFont.Underline <> objFont.Underline) Then
            
              fChangeMade = True
            End If
            Set objCurrentFont = Nothing
            
            Set .Font = objFont
            Set objFont = Nothing
            
          Case WFITEMPROP_FORECOLOR
            If (.ForeColor <> mColForeColor) Then
              fChangeMade = True
            End If
            .ForeColor = mColForeColor
            
          Case WFITEMPROP_BACKCOLOR
            If (.BackColor <> mColBackColor) Then
              fChangeMade = True
            End If
            .BackColor = mColBackColor
            
          Case WFITEMPROP_PICTURE
            sFileName = ""
            If mlngPictureID > 0 Then
              recPictEdit.Index = "idxID"
              recPictEdit.Seek "=", mlngPictureID
  
              If recPictEdit.NoMatch Then
                mlngPictureID = 0
              Else
                sFileName = ReadPicture
                .PictureID = mlngPictureID
                .Picture = LoadPicture(sFileName)
                Kill sFileName
              End If
            End If
            
            If (.PictureID <> mlngPictureID) Then
              fChangeMade = True
            End If
              
            If mlngPictureID = 0 Then
              .Picture = LoadPicture(sFileName)
              .PictureID = mlngPictureID
            End If
            
          Case WFITEMPROP_PICTURELOCATION
            If (.PictureLocation <> mlngPictureLocation) Then
              fChangeMade = True
            End If
            .PictureLocation = mlngPictureLocation
'''            blnDontRefresh = True
            
          Case WFITEMPROP_VALIDATION
            'jpd future development
        
          Case WFITEMPROP_TABNUMBER
            
            If mlngTabNumber <> mfrmWebForm.tabPages.SelectedItem.Tag Then
              lngTargetPageNumber = mfrmWebForm.tabPages.SelectedItem.Tag
              sCaption = mfrmWebForm.tabPages.SelectedItem.Caption
              
              mfrmWebForm.tabPages.SelectedItem.Tag = mfrmWebForm.tabPages.Tabs.Item(mlngTabNumber).Tag
              mfrmWebForm.tabPages.Tabs.Item(mlngTabNumber).Tag = lngTargetPageNumber
              
              ' Select the new page
              mfrmWebForm.tabPages.SelectedItem.Caption = mfrmWebForm.tabPages.Tabs.Item(mlngTabNumber).Caption
              mfrmWebForm.tabPages.Tabs.Item(mlngTabNumber).Caption = sCaption
              mfrmWebForm.PageNo = mlngTabNumber
              mfrmWebForm.IsChanged = True
              blnDontRefresh = True
            End If
            
          Case WFITEMPROP_TABCAPTION
            mfrmWebForm.tabPages.Tabs.Item(mlngTabNumber).Caption = ssGridProperties.ActiveCell.Text
            mfrmWebForm.IsChanged = True
            blnDontRefresh = True
            
        End Select
      End With
    End If

    If fChangeMade Then
      ' Update save button
      mfrmWebForm.IsChanged = True
    End If
  End If

TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  Set objFont = Nothing
  If Not blnDontRefresh Then RefreshProperties True
  UpdateControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function AutoResizeControl(pctlControl As VB.Control) As Boolean
  ' Adjust dimensions of labels if the font/caption change.
  ' Return TRUE of changes were incurred
  Dim fFixedCaption As Boolean
  Dim fChangesMade As Boolean
  
  fChangesMade = False
  
  fFixedCaption = True
  If WebFormItemHasProperty(pctlControl.WFItemType, WFITEMPROP_CAPTIONTYPE) Then
    fFixedCaption = (pctlControl.CaptionType = giWFDATAVALUE_FIXED)
  End If

  If fFixedCaption Then
    Select Case pctlControl.WFItemType
      Case giWFFORMITEM_LABEL
        lblSizeTester.Width = pctlControl.Width
        lblSizeTester.Height = pctlControl.Height
        lblSizeTester.Caption = ""
        Set lblSizeTester.Font = pctlControl.Font
        lblSizeTester.WordWrap = True
        lblSizeTester.BorderStyle = miBorderStyle
      
        lblSizeTester.AutoSize = True
        lblSizeTester.Caption = msCaption
        lblSizeTester.AutoSize = False
      
        If (pctlControl.Width <> lblSizeTester.Width) _
          Or (pctlControl.Height <> lblSizeTester.Height) Then
          fChangesMade = True
        End If
        pctlControl.Width = lblSizeTester.Width
        pctlControl.Height = lblSizeTester.Height
        
      Case giWFFORMITEM_DBVALUE, _
        giWFFORMITEM_WFVALUE, _
        giWFFORMITEM_DBFILE, _
        giWFFORMITEM_WFFILE
        
        lblSizeTester.Caption = ""
        Set lblSizeTester.Font = pctlControl.Font
        lblSizeTester.WordWrap = False
        lblSizeTester.BorderStyle = miBorderStyle
  
        lblSizeTester.AutoSize = True
        lblSizeTester.Caption = "A"
        lblSizeTester.AutoSize = False
  
        If pctlControl.Height < lblSizeTester.Height Then
          pctlControl.Height = lblSizeTester.Height
        End If
    End Select
  End If
  
  AutoResizeControl = fChangesMade
  
End Function


Private Function GetFontDescription() As String

  ' Return the test description of the current font for display in the grid.
  On Error GoTo ErrorTrap
  
  Dim sFontDescription As String
  
  If Not mObjFont Is Nothing Then
    With mObjFont
    
      sFontDescription = .Name
          
      If .Bold Then
        If .Italic Then
          sFontDescription = sFontDescription & ", Bold Italic"
        Else
          sFontDescription = sFontDescription & ", Bold"
        End If
      Else
        If .Italic Then
          sFontDescription = sFontDescription & ", Italic"
        Else
          sFontDescription = sFontDescription & ", Regular"
        End If
      End If
      
      sFontDescription = sFontDescription & IIf(.Strikethrough, ", Strikethrough", "")
      sFontDescription = sFontDescription & IIf(.Underline, ", Underline", "")
    
    End With
  Else
    sFontDescription = ""
  End If
  
TidyUpAndExit:
  GetFontDescription = sFontDescription
  Exit Function
  
ErrorTrap:
  sFontDescription = "<unknown>"
  Resume TidyUpAndExit
  
End Function

Private Function GetHeadFontDescription() As String

  ' Return the test description of the current font for display in the grid.
  On Error GoTo ErrorTrap
  
  Dim sHeadFontDescription As String
  
  If Not mObjHeadFont Is Nothing Then
    With mObjHeadFont
    
      sHeadFontDescription = .Name
          
      If .Bold Then
        If .Italic Then
          sHeadFontDescription = sHeadFontDescription & ", Bold Italic"
        Else
          sHeadFontDescription = sHeadFontDescription & ", Bold"
        End If
      Else
        If .Italic Then
          sHeadFontDescription = sHeadFontDescription & ", Italic"
        Else
          sHeadFontDescription = sHeadFontDescription & ", Regular"
        End If
      End If
      
      sHeadFontDescription = sHeadFontDescription & IIf(.Strikethrough, ", Strikethrough", "")
      sHeadFontDescription = sHeadFontDescription & IIf(.Underline, ", Underline", "")
    
    End With
  Else
    sHeadFontDescription = ""
  End If
  
TidyUpAndExit:
  GetHeadFontDescription = sHeadFontDescription
  Exit Function
  
ErrorTrap:
  sHeadFontDescription = "<unknown>"
  Resume TidyUpAndExit
  
End Function

Private Function ValidIntegerString(psString As String) As Boolean
  
  ' Return true if the given string is a string.
  On Error GoTo ErrorTrap
  
  Dim fValid As Boolean
  Dim lngValueOfString As Long
  Dim sStringOfValue As String
  
  psString = Trim(psString)
  lngValueOfString = val(psString)
  sStringOfValue = Trim(Str(lngValueOfString))
  
  fValid = (psString = sStringOfValue)
  
TidyUpAndExit:
  ValidIntegerString = fValid
  Exit Function

ErrorTrap:
  fValid = False
  Resume TidyUpAndExit

End Function

Private Function ValidDecimalString(psString As String) As Boolean
  
  ' Return true if the given string is a string.
  On Error GoTo ErrorTrap
  
  Dim OrdinalPart As Long
  Dim DecimalPart As Long
  
  Dim fValid As Boolean
  Dim dblValueOfString As Double
  Dim sStringOfValue As String
 
  psString = Trim(psString)
  dblValueOfString = CDbl(psString)
    
  fValid = True
  
TidyUpAndExit:
  ValidDecimalString = fValid
  Exit Function

ErrorTrap:
  fValid = False
  Resume TidyUpAndExit

End Function

Private Function ValidateWFElementIdentifier(psString As String) As Boolean

  ' Return true if the given string is a string.
  On Error GoTo ErrorTrap
  
  Dim fValid As Boolean
  
  psString = Trim(psString)
  
  fValid = (Len(psString) > 0)
    'And (mfrmWebForm.IsUniqueElementIdentifier(psString))
  
TidyUpAndExit:
  ValidateWFElementIdentifier = fValid
  Exit Function

ErrorTrap:
  fValid = False
  Resume TidyUpAndExit

End Function

Private Function ValidateWFIdentifier(psString As String, _
  pctlControlToIgnore As VB.Control) As Boolean

  ' Return true if the given string is a string.
  On Error GoTo ErrorTrap
  
  Dim fValid As Boolean
  
  psString = Trim(psString)
  
  fValid = (Len(psString) > 0)
    'And (mfrmWebForm.IsUniqueIdentifier(psString, pctlControlToIgnore))
  
TidyUpAndExit:
  ValidateWFIdentifier = fValid
  Exit Function

ErrorTrap:
  fValid = False
  Resume TidyUpAndExit

End Function

Private Sub ssGridProperties_RowLoaded(ByVal Bookmark As Variant)
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim ctlControl As VB.Control

  fOK = Not (mfrmWebForm Is Nothing)

  If fOK Then
    With ssGridProperties
      If UBound(mactlSelectedControls) > 0 Then
        For iIndex = 1 To UBound(mactlSelectedControls)
          Set ctlControl = mactlSelectedControls(iIndex)
            Select Case val(.Columns(2).CellValue(Bookmark))
            
            Case WFITEMPROP_HEIGHT
              If WebFormItemHasProperty(ctlControl.WFItemType, WFITEMPROP_HEIGHTBEHAVIOUR) Then
                If UBound(mactlSelectedControls) = 1 Then
                  If Not (mlngHeightBehaviour = behaveFixed) Then
                    .Columns(0).CellStyleSet ("ssetDisabled")
                    .Columns(1).CellStyleSet ("ssetDisabled")
                    .Columns(3).value = True  ' Disabled
                  End If
                End If
              Else
                .Columns(0).CellStyleSet ("ssetEnabled")
                .Columns(1).CellStyleSet ("ssetEnabled")
                .Columns(3).value = False  ' Enabled
              End If
    
            Case WFITEMPROP_WIDTH
              If WebFormItemHasProperty(ctlControl.WFItemType, WFITEMPROP_WIDTHBEHAVIOUR) Then
                If UBound(mactlSelectedControls) = 1 Then
                  If Not (mlngWidthBehaviour = behaveFixed) Then
                    .Columns(0).CellStyleSet ("ssetDisabled")
                    .Columns(1).CellStyleSet ("ssetDisabled")
                    .Columns(3).value = True  ' Disabled
                  End If
                End If
              Else
                .Columns(0).CellStyleSet ("ssetEnabled")
                .Columns(1).CellStyleSet ("ssetEnabled")
                .Columns(3).value = False  ' Enabled
              End If
          
            Case WFITEMPROP_BACKCOLOR
              If WebFormItemHasProperty(ctlControl.WFItemType, WFITEMPROP_BACKSTYLE) Then
                If UBound(mactlSelectedControls) = 1 Then
                  If Not (miBackStyle = BACKSTYLE_OPAQUE) Then
                    .Columns(0).CellStyleSet ("ssetDisabled")
                    .Columns(1).CellStyleSet ("ssetDisabled")
                    .Columns(3).value = True  ' Disabled
                  Else
                    .Columns(1).CellStyleSet ("ssetBackColorValue")
                    .Columns(3).value = False  ' Enabled
                  End If
                End If
              Else
                .Columns(1).CellStyleSet ("ssetBackColorValue")
                .Columns(3).value = False  ' Enabled
              End If
            
            End Select
        Next
      End If
    
      ' AE20080220 Fault #12925
      Select Case val(.Columns(2).CellValue(Bookmark))
      Case WFITEMPROP_NONE
        .Columns(0).CellStyleSet ("ssetDormantRowBold")

      Case WFITEMPROP_UNKNOWN
        .Columns(0).CellStyleSet ("ssetDormantRowBold")

      Case WFITEMPROP_FORECOLOR
        .Columns(1).CellStyleSet ("ssetForeColorValue")
      
      ' AE20080306 Fault #12970,#12925
      Case WFITEMPROP_BACKCOLOR
        If UBound(mactlSelectedControls) = 0 Then
          .Columns(1).CellStyleSet ("ssetBackColorValue")
        End If
        
      Case WFITEMPROP_BACKCOLOREVEN
        .Columns(1).CellStyleSet ("ssetBackColorEven")

      Case WFITEMPROP_BACKCOLORODD
        .Columns(1).CellStyleSet ("ssetBackColorOdd")

      Case WFITEMPROP_FORECOLOREVEN
        .Columns(1).CellStyleSet ("ssetForeColorEven")

      Case WFITEMPROP_FORECOLORODD
        .Columns(1).CellStyleSet ("ssetForeColorOdd")

      Case WFITEMPROP_HEADERBACKCOLOR
        .Columns(1).CellStyleSet ("ssetHeaderBackColor")

      Case WFITEMPROP_BACKCOLORHIGHLIGHT
        .Columns(1).CellStyleSet ("ssetBackColorHighlight")

      Case WFITEMPROP_FORECOLORHIGHLIGHT
        .Columns(1).CellStyleSet ("ssetForeColorHighlight")

      End Select
    End With
  End If
        
End Sub
