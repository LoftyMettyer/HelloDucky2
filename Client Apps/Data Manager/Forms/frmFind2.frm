VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "actbar.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmFind2 
   Caption         =   "Find"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1040
   Icon            =   "frmFind2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   6360
   Visible         =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3690
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10689
            Text            =   "x Records"
            TextSave        =   "x Records"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraSummary 
      Caption         =   "History Summary :"
      Height          =   1000
      Left            =   150
      TabIndex        =   3
      Top             =   2200
      Visible         =   0   'False
      Width           =   6000
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3200
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   195
      End
      Begin TDBMask6Ctl.TDBMask Text1 
         Height          =   300
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   315
         Visible         =   0   'False
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   529
         Caption         =   "frmFind2.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmFind2.frx":0061
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   -2147483633
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   ""
         HighlightText   =   0
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
         PromptChar      =   " "
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   ""
         Value           =   ""
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   300
         Index           =   0
         Left            =   1395
         TabIndex        =   8
         Top             =   315
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   529
         Calculator      =   "frmFind2.frx":00A3
         Caption         =   "frmFind2.frx":00C3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFind2.frx":0129
         Keys            =   "frmFind2.frx":0147
         Spin            =   "frmFind2.frx":0191
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483633
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "########0; -########0"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "########; -########"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1329790981
         MinValueVT      =   1668218885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   2205
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin VB.Frame fraOrders 
      Caption         =   "Order :"
      Height          =   800
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   6000
      Begin VB.ComboBox cmbOrders 
         Height          =   315
         Left            =   195
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   5595
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssOleDBGridFindColumns 
      Height          =   1000
      Left            =   150
      TabIndex        =   0
      Top             =   1100
      Visible         =   0   'False
      Width           =   6000
      ScrollBars      =   0
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      GroupHeadLines  =   0
      AllowUpdate     =   0   'False
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
      UseExactRowCount=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   3
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   10583
      _ExtentY        =   1764
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   165
      Tag             =   "BAND_FIND"
      Top             =   3270
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmFind2.frx":01B9
   End
End
Attribute VB_Name = "frmFind2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'MH20061213 Fault 11785
Private mintResizeCount As Integer

' Form variables.
Private mfFindForm As Boolean
Private mvarScreenType As ScreenType

' Parent form variables.
Private mfrmParent As Form
Private mlngParentFormID As Long
Private mrsParentData As ADODB.Recordset

' Table/view variables.
Private mobjTableView As CTablePrivilege
Private msCurrentTableViewName As String
Private mfFirstColumnsMatch As Boolean
Private mfFirstColumnAscending As Boolean
Private miFirstColumnDataType As Integer
Private mlngFirstSortColumnID As Long
Private mfFirstOrderColumnIsFindColumn As Boolean

' Recordset variables.
Private mrsFindRecords As ADODB.Recordset
Private mlngOrderID As Long
Private miOldOrderIndex As Integer

' General form handling variables.
Private mfSizing As Boolean
Private mfCancelled As Boolean
Private mfFormattingGrid As Boolean
Private mdblMinSummaryWidth As Double
Private mlngMinFormWidth As Long
Private mlngMinFormHeight As Long
Private mfMessageDisplayed As Boolean
Private mlngCurrentBookMark As Long

' Array to hold details of the indexes of the different controls.
Private aryControlArray()

' Boolean to indicate if the History details have been loaded.
Private blnSummaryDetailsLoaded As Boolean

Private mlngCurrentRecordID As Long       ' ID of currently selected record
Private mbIsLoading As Boolean            ' Is the form in the process of loading
Private mbIsUnloading As Boolean          ' Is the form unloading

Private mblnPrintCancelled As Boolean

Private mfCanBookCourse As Boolean
Private mfCanAddFromWaitingList As Boolean
Private mfCanCancelBooking As Boolean
Private mfCanTransferBooking As Boolean
Private mfCanBulkBooking As Boolean

Private mfBookCourseVisible As Boolean
Private mfAddFromWaitingListVisible As Boolean
Private mfCancelBookingVisible As Boolean
Private mfTransferVisible As Boolean
Private mfBulkBookingVisible As Boolean

Private mfCustomReportExists As Boolean
Private mfCalendarReportExists As Boolean
Private mfGlobalUpdateExists As Boolean
Private mfDataTransferExists As Boolean
Private mfMailMergeExists As Boolean

Private mfBusy As Boolean

Private mavFindColumns() As Variant        ' Find columns details

Private Const dblFINDFORM_MINWIDTH = 5000
Private Const dblFINDFORM_MINHEIGHT = 6000
Private Const dblHISTORYFORM_MINWIDTH = 4000
Private Const dblHISTORYFORM_MINHEIGHT = 5000
Private Const dblCOORD_XGAP = 150
Private Const dblCOORD_YGAP = 150

Private mstrSelectedRecords As String
Private mblnRefreshing As Boolean


  
Public Property Get AddFromWaitingListVisible() As Boolean
  AddFromWaitingListVisible = mfAddFromWaitingListVisible
  
End Property
Public Property Get CancelBookingVisible() As Boolean
  CancelBookingVisible = mfCancelBookingVisible
  
End Property
Public Property Get BulkBookingVisible() As Boolean
  BulkBookingVisible = mfBulkBookingVisible
  
End Property
Public Property Get BookCourseVisible() As Boolean
  BookCourseVisible = mfBookCourseVisible

End Property
Public Property Get TransferVisible() As Boolean
  TransferVisible = mfTransferVisible
  
End Property

Public Sub ResizeFindColumns()

  Dim dblCurrentSize As Double
  Dim dblNewSize As Double
  Dim iCount As Integer
  Dim dblResizeFactor As Double
  Dim bNeedScrollBars As Boolean

  DoEvents
  
  With ssOleDBGridFindColumns
    
    .Redraw = False
    
    If .VisibleRows > .Rows Then
      .MoveFirst
      SetCurrentRecord
    Else
      UpdateStatusBar
    End If
    
    'Check if the number of rows in the recordset is more than can be displayed in the grid.
    'OR
    'Check if the first 'visible' row is the first in the recordset.
    bNeedScrollBars = IIf((.Rows > .VisibleRows) Or (.FirstRow > 1), True, False)

    ' Calculate the existing size of the find grid
    dblCurrentSize = 0
    For iCount = 0 To (.Cols - 1)
      If .Columns(iCount).Visible Then
        dblCurrentSize = dblCurrentSize + .Columns(iCount).Width
      End If
    Next iCount

    ' Calculate size of resized grid
    dblNewSize = .Width
    
    If bNeedScrollBars Then
      dblNewSize = dblNewSize - (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX) '+ 20
      .ScrollBars = ssScrollBarsVertical
    Else
      .ScrollBars = ssScrollBarsNone
    End If
    
    dblNewSize = dblNewSize - (UI.GetSystemMetrics(SM_CXFRAME) * 2)
    dblNewSize = dblNewSize - (UI.GetSystemMetrics(SM_CXBORDER) * Screen.TwipsPerPixelX)

'    'TM20020327 Fault 2568
    If dblCurrentSize < 1 Or .Cols < 1 Then Exit Sub
    
    ' Calculate the ratio that the grid needs to be resized to
    dblResizeFactor = Round(dblNewSize / dblCurrentSize, 2)

  ' Scroll through adjusting each column according to the resize factor
    For iCount = 0 To (.Cols - 1)
      If .Columns(iCount).Visible Then

        'Make the last column nice & snug
        If iCount = (.Cols - 2) Then
          .Columns(iCount).Width = dblNewSize
        Else
          .Columns(iCount).Width = (.Columns(iCount).Width * dblResizeFactor)
          dblNewSize = dblNewSize - .Columns(iCount).Width
        End If

        ' SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(iCount).Name, .Columns(iCount).Width
        SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(iCount).Caption, .Columns(iCount).Width
      End If
    Next iCount

    ' JPD20021025 Fault 4647
    .Redraw = True
  End With

End Sub

Private Sub ReformatHistoryElements()
  
  ' Format the summary field controls.
  Dim lLeftColumns As Long
  Dim objNewControl As Object
  Dim objControlArray As Object
  Dim iControlType As ControlTypes
  Dim iControlDataType As SQLDataType
  Dim objSummaryField As clsSummaryField
  Dim objSummaryFields As clsSummaryFields
  
  Dim lngControlSize As Long
  Dim lngLeftPos As Long
  Dim lngControlLeft As Long
  
  Const lngFRAMEOFFSET = 260
  Const lngXOFFSET_LEFT = 200
  Const lngXOFFSET_RIGHT = 200
  Const lngXCONTROLOFFSET = 300
  Const lngXLABELOFFSET = 60
  Const lngMinControlSize = 500
  
  If blnSummaryDetailsLoaded Then
    
    ' Get the summary field definitions.
    Set objSummaryFields = datGeneral.GetHistorySummaryFields(mobjTableView.TableID, mfrmParent.ParentTableID)
  
    'Calculate the amount of columns on the left
    If objSummaryFields.ManualColumnBreak = True Then
      lLeftColumns = objSummaryFields.ColumnBreakPoint
    Else
      lLeftColumns = (objSummaryFields.Count + 1) / 2
    End If
  
    If Label1.Count > 1 Then
      ' Loop through the collection changing the left property of the labels .
      For Each objSummaryField In objSummaryFields.Collection
        If objSummaryField.Sequence <= lLeftColumns Then
          
          With Label1(objSummaryField.Sequence)
          
            .Left = lngXOFFSET_LEFT
            
            If .Width + .Left > lngLeftPos Then
              lngLeftPos = .Width + .Left
            End If
            
          End With
  
        End If
      Next objSummaryField
      Set objSummaryField = Nothing
   
      lngControlLeft = lngLeftPos + lngXLABELOFFSET
    
      If (fraSummary.Width / 2) - lngControlLeft - lngXOFFSET_RIGHT > lngMinControlSize Then
        lngControlSize = (fraSummary.Width / 2) - lngControlLeft - lngXOFFSET_RIGHT
      Else
        lngControlSize = lngMinControlSize
      End If
    End If
    
    ' Loop through the collection changing the left property and the width of the non-checkbox controls.
    For Each objSummaryField In objSummaryFields.Collection
      If objSummaryField.Sequence <= lLeftColumns Then
        iControlType = objSummaryField.ControlType
        iControlDataType = objSummaryField.DataType
        
        Set objControlArray = GetControlArray(iControlType, iControlDataType)
  
        If Not objControlArray Is Nothing Then
  
          Set objNewControl = objControlArray(aryControlArray(1, objSummaryField.Sequence - 1))
          If objControlArray.Count > 1 Then
            With objNewControl
            
              .Left = lngControlLeft
              
              If iControlType <> ctlCheck Then
                .Width = lngControlSize
              End If
      
              'JPD 20030819 Fault 6782
              If .Name = "TDBNumber1" Then
                .Caption.Size = 0
              End If
              
              If .Width + .Left > lngLeftPos Then
                lngLeftPos = .Width + .Left
              End If
      
            End With
          End If
        End If
      End If
    Next objSummaryField
    Set objSummaryField = Nothing

    lngControlLeft = lngLeftPos + lngXOFFSET_RIGHT + lngXOFFSET_LEFT
        
    ' Loop through the collection changing the left property of the labels .
    If Label1.Count > 1 Then
      For Each objSummaryField In objSummaryFields.Collection
        If objSummaryField.Sequence > lLeftColumns Then
          
          With Label1(objSummaryField.Sequence)
  
            .Left = lngControlLeft
          
              If .Width + .Left > lngLeftPos Then
                lngLeftPos = .Width + .Left
              End If
          
          End With
  
        End If
      Next objSummaryField
      Set objSummaryField = Nothing
  
      lngControlLeft = lngLeftPos + lngXLABELOFFSET
    End If
    
    If fraSummary.Width - lngControlLeft - lngXOFFSET_RIGHT > lngMinControlSize Then
      lngControlSize = fraSummary.Width - lngControlLeft - lngXOFFSET_RIGHT
    Else
      lngControlSize = lngMinControlSize
    End If
    
    ' Loop through the collection changing the left property and the width of the non-checkbox controls.
    For Each objSummaryField In objSummaryFields.Collection
      If objSummaryField.Sequence > lLeftColumns Then
        iControlType = objSummaryField.ControlType
        iControlDataType = objSummaryField.DataType
        
        Set objControlArray = GetControlArray(iControlType, iControlDataType)
        If objControlArray.Count > 1 Then
          If Not objControlArray Is Nothing Then
   
            Set objNewControl = objControlArray(aryControlArray(1, objSummaryField.Sequence - 1))
            
            With objNewControl
              .Left = lngControlLeft
  
              If iControlType <> ctlCheck Then
                .Width = lngControlSize
                
                'JPD 20030819 Fault 6782
                If .Name = "TDBNumber1" Then
                  .Caption.Size = 0
                End If
              End If
            End With
            
          End If
        End If
      End If
    Next objSummaryField
    Set objSummaryField = Nothing
  End If
  
End Sub

Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  ' RH 20/07/00 - To deactivate the right-click menu totally (preventing customisation)
  '             - NB : The Allow Customisation = False property seems not to work!
  Cancel = True
  
  
End Sub

Public Property Get Recordset() As ADODB.Recordset
  ' Return the recordset for the form.
  Set Recordset = mfrmParent.Recordset

End Property

Public Property Get AllowInsert() As Boolean
  ' Get's the parent's allow insert.
  AllowInsert = mfrmParent.AllowInsert
  
End Property

Public Property Get AllowUpdate() As Boolean
  AllowUpdate = False
  
End Property


Public Property Get Busy() As Boolean
  Busy = mfBusy
  
End Property



Public Property Get AllowDelete() As Boolean
  ' Get's the parent's allow delete property.
  AllowDelete = mfrmParent.AllowDelete
  
End Property


Public Property Get ParentID() As Long
  ParentID = mfrmParent.ParentID
End Property


Public Function FindStartFromPrimary(ByVal pobjTableView As CTablePrivilege, _
  ByVal plngOrderID As Long, _
  pfrmParentForm As Form, _
  pfFindForm As Boolean) As Boolean
  ' Initialise the form to be called from a primary screen.
  Dim fOK As Boolean

  ' JPD20020924 Fault 4414
  Screen.MousePointer = vbHourglass

  fOK = True
  mfMessageDisplayed = False

  ' Store the parameters in local form variables
  mfFindForm = pfFindForm
  Set mobjTableView = pobjTableView
  mlngOrderID = plngOrderID
  
  ' Get extra information from the parent form.
  Set mfrmParent = pfrmParentForm
  With mfrmParent
    mlngParentFormID = .FormID
    
    SetFormCaption Me, mfrmParent.FindCaption
    
    If pfFindForm Then
      mvarScreenType = screenFind
    End If
  End With
    
  If mobjTableView.IsTable Then
    msCurrentTableViewName = mobjTableView.TableName
  Else
    msCurrentTableViewName = mobjTableView.ViewName
  End If
  
  Dim blnShowSummary As Boolean
  blnShowSummary = (Not mfFindForm Or mfrmParent.ParentTableID > 0)
  
  ' Get rid of the pesky icon
  RemoveIcon Me
  
  ' Populate the Orders combo.
  fOK = ConfigureOrdersCombo

  ' JPD20020924 Fault 4414
  Screen.MousePointer = vbHourglass

  If fOK Then
    ' Get the find records.
    Set mrsFindRecords = New ADODB.Recordset
    fOK = GetFindRecords
  End If
  
  If fOK Then

    'MH20010824 Fault 2075
    'If Not mfFindForm Then
    If Not mfFindForm Or mfrmParent.ParentTableID > 0 Then
      
      ' Format the summary controls.
      SetupSummary
    
      With ssOleDBGridFindColumns
        If .Rows > 0 Then
          .SelBookmarks.RemoveAll
          .MoveFirst
          .SelBookmarks.Add .Bookmark
        End If
      End With
    End If
    
    ' Format the required controls.
    FormatControls
    frmMain.RefreshRecordMenu Me

    ' Populate the grid.
    If fOK Then
      If pfFindForm Then
        ' Load previous size setting
        'TM23112004 Fault 9388 - Restore the form if it is minimized
        If Me.WindowState = vbMinimized Then
          Me.WindowState = vbNormal
        Else
          Me.Top = GetPCSetting("FindWindowCoOrdinates\" & gsDatabaseName & "\" & mfrmParent.ScreenID, "Top", Me.Top)
          Me.Left = GetPCSetting("FindWindowCoOrdinates\" & gsDatabaseName & "\" & mfrmParent.ScreenID, "Left", Me.Left)
          Me.Width = GetPCSetting("FindWindowCoOrdinates\" & gsDatabaseName & "\" & mfrmParent.ScreenID, "Width", Me.Width)
          Me.Height = GetPCSetting("FindWindowCoOrdinates\" & gsDatabaseName & "\" & mfrmParent.ScreenID, "Height", Me.Height)
        End If
        
        'AE20080103 Fault #12480
'        SetCurrentRecord
'        Me.Visible = True
        UI.LockWindow Me.hWnd
        Me.Visible = True
        ResizeFindColumns
        SetCurrentRecord
        UI.UnlockWindow
      End If
    End If
  End If
  
ExitFindStartFromPrimary:
  ' JPD20020920 Fault 4414
  Screen.MousePointer = vbDefault
  mbIsLoading = False
  
  FindStartFromPrimary = fOK
  Exit Function
  
End Function

Private Function GetFindRecords() As Boolean
  ' Read the required find and summary information.
  Dim fOK As Boolean
  Dim fColumnOK As Boolean
  Dim fFound As Boolean
  Dim fNoSelect As Boolean
  Dim iIndex As Long ' Integer
  Dim iNextIndex As Integer
  Dim lngFirstFindColumnID As Long
  Dim sSQL As String
  Dim sSource As String
  Dim sRealSource As String
  Dim sColumnCode As String
  Dim sColumnList As String
  Dim sJoinCode As String
  Dim sWhereSQL As String
  Dim sOrderString As String
  Dim objTableView As CTablePrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim rsInfo As Recordset
  Dim alngTableViews() As Long
  Dim asViews() As String

  fOK = True
  fNoSelect = False

  sOrderString = ""
  sJoinCode = ""
  sColumnList = ""
  sWhereSQL = ""

  ' Dimension an array of tables/views joined to the base table/view.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ReDim alngTableViews(2, 0)

  mfFirstColumnsMatch = False
  lngFirstFindColumnID = 0
  mlngFirstSortColumnID = 0
  mfFirstColumnAscending = True
  miFirstColumnDataType = 0
  mfFirstOrderColumnIsFindColumn = False
  
  ' NPG20081128 Fault 13248
  ' ReDim mavFindColumns(3, 0)
  ReDim mavFindColumns(4, 0)

  ' Get the default order items from the database.
  Set rsInfo = datGeneral.GetOrderDefinition(mlngOrderID)

  If rsInfo.EOF And rsInfo.BOF Then
    COAMsgBox "No order defined for this " & IIf(mobjTableView.IsTable, "table.", "view.") & _
      vbCrLf & "Unable to display records.", vbExclamation, "Security"
    fOK = False
  Else
    ' Check the user's privilieges on the order columns.
    Do While Not rsInfo.EOF
      ' Get the column privileges collection for the given table.
      If rsInfo!TableID = mobjTableView.TableID Then
        sSource = msCurrentTableViewName
      Else
        sSource = rsInfo!TableName
      End If
      Set objColumnPrivileges = GetColumnPrivileges(sSource)
      sRealSource = gcoTablePrivileges.Item(sSource).RealSource
      
      fColumnOK = objColumnPrivileges.IsValid(rsInfo!ColumnName)

      If fColumnOK Then
        fColumnOK = objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect
      End If
      Set objColumnPrivileges = Nothing

      If fColumnOK Then
        ' The column can be read from the base table/view, or directly from a parent table.
        If rsInfo!Type = "F" Then
          ' Add the column to the column list.
          sColumnList = sColumnList & _
            IIf(Len(sColumnList) > 0, ", ", "") & _
            sRealSource & "." & Trim(rsInfo!ColumnName)
          
          mavFindColumns(0, UBound(mavFindColumns, 2)) = rsInfo!ColumnID
          mavFindColumns(1, UBound(mavFindColumns, 2)) = rsInfo!Size
          mavFindColumns(2, UBound(mavFindColumns, 2)) = rsInfo!Decimals
          mavFindColumns(3, UBound(mavFindColumns, 2)) = rsInfo!Use1000Separator
          mavFindColumns(4, UBound(mavFindColumns, 2)) = rsInfo!BlankIfZero
          ReDim Preserve mavFindColumns(4, UBound(mavFindColumns, 2) + 1)
          
          ' Remember the first Find column.
          If lngFirstFindColumnID = 0 Then
            lngFirstFindColumnID = rsInfo!ColumnID
          End If
        Else
          ' Add the column to the order string.
          sOrderString = sOrderString & _
            IIf(Len(sOrderString) > 0, ", ", "") & _
            sRealSource & "." & Trim(rsInfo!ColumnName) & _
            IIf(rsInfo!Ascending, "", " DESC")
          
          ' Remember the first Order column.
          If mlngFirstSortColumnID = 0 Then
            mlngFirstSortColumnID = rsInfo!ColumnID
            mfFirstColumnAscending = rsInfo!Ascending
            miFirstColumnDataType = rsInfo!DataType
          End If
        End If

        ' If the column comes from a parent table, then add the table to the Join code.
        If rsInfo!TableID <> mobjTableView.TableID Then
          ' Check if the table has already been added to the join code.
          fFound = False
          For iNextIndex = 1 To UBound(alngTableViews, 2)
            If alngTableViews(1, iNextIndex) = 0 And _
              alngTableViews(2, iNextIndex) = rsInfo!TableID Then
              fFound = True
              Exit For
            End If
          Next iNextIndex
          
          If Not fFound Then
            ' The table has not yet been added to the join code, so add it to the array and the join code.
            iNextIndex = UBound(alngTableViews, 2) + 1
            ReDim Preserve alngTableViews(2, iNextIndex)
            alngTableViews(1, iNextIndex) = 0
            alngTableViews(2, iNextIndex) = rsInfo!TableID
            
            sJoinCode = sJoinCode & _
              " LEFT OUTER JOIN " & sRealSource & _
              " ON " & mobjTableView.RealSource & ".ID_" & Trim(Str(rsInfo!TableID)) & _
              " = " & sRealSource & ".ID"
          End If
        End If
      Else
        ' The column cannot be read from the base table/view, or directly from a parent table.
        ' If it is a column from a prent table, then try to read it from the views on the parent table.
        If rsInfo!TableID <> mobjTableView.TableID Then
          ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
          ReDim asViews(0)
          For Each objTableView In gcoTablePrivileges.Collection
            If (Not objTableView.IsTable) And _
              (objTableView.TableID = rsInfo!TableID) And _
              (objTableView.AllowSelect) Then
              
              sSource = objTableView.ViewName
              sRealSource = gcoTablePrivileges.Item(sSource).RealSource

              ' Get the column permission for the view.
              Set objColumnPrivileges = GetColumnPrivileges(sSource)

              If objColumnPrivileges.IsValid(rsInfo!ColumnName) Then
                If objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect Then
                  ' Add the view info to an array to be put into the column list or order code below.
                  iNextIndex = UBound(asViews) + 1
                  ReDim Preserve asViews(iNextIndex)
                  asViews(iNextIndex) = objTableView.ViewName
                  
                  ' Add the view to the Join code.
                  ' Check if the view has already been added to the join code.
                  fFound = False
                  For iNextIndex = 1 To UBound(alngTableViews, 2)
                    If alngTableViews(1, iNextIndex) = 1 And _
                      alngTableViews(2, iNextIndex) = objTableView.ViewID Then
                      fFound = True
                      Exit For
                    End If
                  Next iNextIndex
          
                  If Not fFound Then
                    ' The view has not yet been added to the join code, so add it to the array and the join code.
                    iNextIndex = UBound(alngTableViews, 2) + 1
                    ReDim Preserve alngTableViews(2, iNextIndex)
                    alngTableViews(1, iNextIndex) = 1
                    alngTableViews(2, iNextIndex) = objTableView.ViewID
          
                    sJoinCode = sJoinCode & _
                      " LEFT OUTER JOIN " & sRealSource & _
                      " ON " & mobjTableView.RealSource & ".ID_" & Trim(Str(objTableView.TableID)) & _
                      " = " & sRealSource & ".ID"
                  End If
                End If
              End If
              Set objColumnPrivileges = Nothing

            End If
          Next objTableView
          Set objTableView = Nothing
        
          ' The current user does have permission to 'read' the column through a/some view(s) on the
          ' table.
          If UBound(asViews) = 0 Then
            fNoSelect = True
          Else
            ' Add the column to the column list.
            sColumnCode = ""
            For iNextIndex = 1 To UBound(asViews)
              If iNextIndex = 1 Then
                sColumnCode = "CASE "
              End If
                
              sColumnCode = sColumnCode & _
                " WHEN NOT " & asViews(iNextIndex) & "." & rsInfo!ColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & rsInfo!ColumnName
            Next iNextIndex
              
            If Len(sColumnCode) > 0 Then
              sColumnCode = sColumnCode & _
                " ELSE NULL" & _
                " END AS " & _
                IIf(rsInfo!Type = "F", "", "'?") & _
                rsInfo!ColumnName & _
                IIf(rsInfo!Type = "F", "", "'")
                
              sColumnList = sColumnList & _
                IIf(Len(sColumnList) > 0, ", ", "") & _
                sColumnCode

              If rsInfo!Type = "F" Then
                ' Remember the first Find column.
                If lngFirstFindColumnID = 0 Then
                  lngFirstFindColumnID = rsInfo!ColumnID
                End If
              
                ' NPG20081128 Fault 13248
                ReDim Preserve mavFindColumns(4, UBound(mavFindColumns, 2) + 1)
                mavFindColumns(0, UBound(mavFindColumns, 2)) = rsInfo!ColumnID
                mavFindColumns(1, UBound(mavFindColumns, 2)) = rsInfo!Size
                mavFindColumns(2, UBound(mavFindColumns, 2)) = rsInfo!Decimals
                mavFindColumns(3, UBound(mavFindColumns, 2)) = rsInfo!Use1000Separator
                mavFindColumns(4, UBound(mavFindColumns, 2)) = rsInfo!BlankIfZero
              
              Else
                ' Add the column to the order string.
                sOrderString = sOrderString & _
                  IIf(Len(sOrderString) > 0, ", ", "") & _
                  "'?" & Trim(rsInfo!ColumnName) & "'" & _
                  IIf(rsInfo!Ascending, "", " DESC")

                ' Remember the first Order column.
                If mlngFirstSortColumnID = 0 Then
                  mlngFirstSortColumnID = rsInfo!ColumnID
                  mfFirstColumnAscending = rsInfo!Ascending
                  miFirstColumnDataType = rsInfo!DataType
                End If
              End If
            End If
          End If
        End If
      End If
      rsInfo.MoveNext
    Loop

    ' Inform the user if they do not have permission to see the data.
    If fNoSelect And (Not mfMessageDisplayed) Then
      COAMsgBox "You do not have 'read' permission on all of the columns in the selected order." & _
        vbCrLf & "Only permitted columns will be shown.", vbExclamation, "Security"
      mfMessageDisplayed = True
    End If
    
    mfFirstColumnsMatch = (lngFirstFindColumnID = mlngFirstSortColumnID)
  
    ' JPD20030205 Fault 5020
    For iIndex = 1 To UBound(mavFindColumns, 2)
      If mavFindColumns(0, iIndex) = mlngFirstSortColumnID Then
        mfFirstOrderColumnIsFindColumn = True
        Exit For
      End If
    Next iIndex
    
    ' Create the string for creating the items that will appear in the picklist definition listbox.
    If LenB(sColumnList) > 0 Then
      ' Get the 'Where' clause code from the parent table's recordset.
      sWhereSQL = UCase(mfrmParent.Recordset.Source)
      iIndex = InStr(sWhereSQL, " WHERE ")
      If iIndex > 0 Then
        sWhereSQL = Mid(sWhereSQL, iIndex + 1)
        
        iIndex = InStr(sWhereSQL, " ORDER BY ")
        If iIndex > 0 Then
          sWhereSQL = Left(sWhereSQL, iIndex - 1)
        End If
      Else
        sWhereSQL = ""
      End If
      
      sSQL = "SELECT " & sColumnList & ", " & mobjTableView.RealSource & ".id" & _
        " FROM " & mobjTableView.RealSource & _
        " " & sJoinCode & _
        " " & sWhereSQL & _
        IIf(LenB(sOrderString) > 0, " ORDER BY " & sOrderString, "")

      ' Get the required recordset.
      If mfrmParent.RequiresLocalCursor Then gADOCon.CursorLocation = adUseClient
      Set mrsFindRecords = datGeneral.GetPersistentRecords(sSQL, adOpenStatic, adLockReadOnly)
      If mfrmParent.RequiresLocalCursor Then gADOCon.CursorLocation = adUseServer
      
      ConfigureGrid
      
    Else
      COAMsgBox "You do not have permission to read any of the columns in the selected order for this " & _
        IIf(mobjTableView.IsTable, "table.", "view.") & _
        vbCrLf & "Unable to display records.", vbExclamation, "Security"
      fOK = False
    End If
  End If

  'mrsFindRecords.MoveLast

  rsInfo.Close
  Set rsInfo = Nothing

  GetFindRecords = fOK
  
End Function
    



Public Property Get Filtered() As Boolean
  Filtered = mfrmParent.Filtered

End Property


Private Sub SetupSummary()
  ' Format the summary field controls.
  Dim fColumnOK As Boolean
  Dim lCount As Long
  Dim lParent As Long
  Dim lLeftWidth As Long
  Dim lLeftColumns As Long
  Dim lLabelTop As Long
  Dim lLabelCount As Long
  Dim lngCurrentYPosition As Long
  Dim frmForm As Form
  Dim objNewControl As Object
  Dim objControlArray As Object
  Dim iControlType As ControlTypes
  Dim iControlDataType As SQLDataType
  Dim objSummaryField As clsSummaryField
  Dim objSummaryFields As clsSummaryFields
  Dim objTableView As CTablePrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim lngControlSize As Long
  Dim lngLeftPos As Long
  Dim lngControlLeft As Long
  Dim lngLeftLabelSize As Long
  Dim lngRightLabelSize As Long
  Dim sFormat As String
  Dim iDigitCount As Integer
  Dim iCount As Integer
  
  Const lngYOffset = 300
  Const lngYLABELOFFSET = 60
  Const lngYCONTROLOFFSET = 400
  Const lngYSTARTGROUPOFFSET = 300
  Const lngFRAMEOFFSET = 260
  Const lngXOFFSET_LEFT = 200
  Const lngXOFFSET_RIGHT = 200
  Const lngXCONTROLOFFSET = 300
  Const lngXLABELOFFSET = 60
  Const lngMinControlSize = 500
  
  fraSummary.Visible = True
  fraSummary.Caption = "History Summary :"
  
  'JM - 21/11/01 - Fault 3148 - Only run if not already loaded
  If Label1.Count > 1 Then
    LoadSummaryDetails
    Exit Sub
  End If
  
  ' Get the summary field definitions.
  Set objSummaryFields = datGeneral.GetHistorySummaryFields(mobjTableView.TableID, mfrmParent.ParentTableID)
  If objSummaryFields.Count = 0 Then
    fraSummary.Height = 0
    fraSummary.Visible = False
    Exit Sub
  End If

  'Calculate the amount of columns on the left
  If objSummaryFields.ManualColumnBreak = True Then
    lLeftColumns = objSummaryFields.ColumnBreakPoint
  Else
    lLeftColumns = (objSummaryFields.Count + 1) / 2
  End If

  lngCurrentYPosition = lngYOffset + lngYLABELOFFSET

  ' Redimension array to the number of control elements.
  ReDim aryControlArray(2, objSummaryFields.Count)
      
  mdblMinSummaryWidth = 0
  
  ' Loop through the recordset adding the required labels to the left column on the form.
  For Each objSummaryField In objSummaryFields.Collection
    If objSummaryField.Sequence <= lLeftColumns Then
  
      ' Load a new label for the summary control.
      lLabelCount = lLabelCount + 1
      Load Label1(lLabelCount)
      
      With Label1(lLabelCount)
        .Visible = True
        .Left = lngXOFFSET_LEFT
        .Caption = RemoveUnderScores(objSummaryField.ColumnName) & " :"

        If (lLabelCount > 1) And objSummaryField.StartOfGroup Then
          lngCurrentYPosition = lngCurrentYPosition + lngYSTARTGROUPOFFSET
        End If
        .Top = lngCurrentYPosition
  
        If .Top > lLabelTop Then
          lLabelTop = .Top
        End If
        
        If .Width + .Left > lngLeftPos Then
          lngLeftPos = .Width + .Left
        End If
        
        If .Width > lngLeftLabelSize Then lngLeftLabelSize = .Width
        
        lngCurrentYPosition = lngCurrentYPosition + lngYCONTROLOFFSET
  
      End With
    End If
  Next objSummaryField
  Set objSummaryField = Nothing
  
  lCount = 1

  lngControlLeft = lngLeftPos + lngXLABELOFFSET
 
  If (fraSummary.Width / 2) - lngControlLeft - lngXOFFSET_RIGHT > lngMinControlSize Then
    lngControlSize = (fraSummary.Width / 2) - lngControlLeft - lngXOFFSET_RIGHT
  Else
    lngControlSize = lngMinControlSize
  End If
    
  ' Loop through the recordset adding the required controls to the left column on the form.
  For Each objSummaryField In objSummaryFields.Collection
    If objSummaryField.Sequence <= lLeftColumns Then
      iControlType = objSummaryField.ControlType
      iControlDataType = objSummaryField.DataType
      
      Set objControlArray = GetControlArray(iControlType, iControlDataType)

      If Not objControlArray Is Nothing Then
        Load objControlArray(objControlArray.Count)

        ' Populate array with Sequence and Index of the Control.
        aryControlArray(0, objSummaryField.Sequence - 1) = objSummaryField.Sequence
        aryControlArray(1, objSummaryField.Sequence - 1) = objControlArray.Count - 1

        Set objNewControl = objControlArray(objControlArray.Count - 1)
  
        With objNewControl
          .Visible = True
  
          .Top = Label1(lCount).Top - lngYLABELOFFSET
          .Tag = objSummaryField.ColumnName
  
          ' Grey out controls if they are for columns thaqt are not in the parent view.
          If mfrmParent.ParentViewID > 0 Then
            ' Check if the column is in the parent view.
            Set objTableView = gcoTablePrivileges.FindViewID(mfrmParent.ParentViewID)
            Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)
            
            fColumnOK = objColumnPrivileges.IsValid(objSummaryField.ColumnName)
            If fColumnOK Then
              fColumnOK = objColumnPrivileges.Item(objSummaryField.ColumnName).AllowSelect
            End If
            
            If Not fColumnOK Then
              .BackColor = COL_GREY
              .Enabled = False
              .Tag = "0"
            End If
          End If
          
          If .Name = "Text1" Then
            .AlignHorizontal = objSummaryField.Alignment
          End If
            
          If .Name = "TDBNumber1" Then
            'JPD 20041109 Fault 8230
            datGeneral.FormatTDBNumberControl objNewControl
            
            sFormat = ""
            iDigitCount = 1
            
            ' Loop and create the format mask
            For iCount = 1 To (objSummaryField.Size - objSummaryField.Decimals)
              If objSummaryField.Use1000Separator Then
                sFormat = IIf(iCount Mod 3 = 0 And (iCount <> (objSummaryField.Size - objSummaryField.Decimals)), ",#", "#") & sFormat
              Else
                sFormat = "#" & sFormat
              End If
            Next iCount

            If Not objSummaryField.BlankIfZero And Len(sFormat) > 0 Then
              sFormat = Left(sFormat, Len(sFormat) - 1) & "0"
            End If

            If objSummaryField.Decimals > 0 Then
              sFormat = sFormat & "."
              For iCount = 1 To objSummaryField.Decimals
                If objSummaryField.BlankIfZero Then
                  sFormat = sFormat & "#"
                Else
                  sFormat = sFormat & "0"
                End If
              Next iCount
            End If
            
            .DecimalPoint = UI.GetSystemDecimalSeparator
            .Separator = UI.GetSystemThousandSeparator
            .DisplayFormat = sFormat
            .Format = sFormat

            'JPD 20041109 Fault 9429
            .MaxValue = Val(String(objSummaryField.Size - objSummaryField.Decimals, "9") & "." & String(objSummaryField.Decimals, "9"))
            .MinValue = (.MaxValue * -1)
          End If

          .Left = lngControlLeft
          
          If iControlType <> ctlCheck Then
            .Width = lngControlSize
          End If
  
          If .Width + .Left > lngLeftPos Then
            lngLeftPos = .Width + .Left
          End If

        End With
      
        Set objControlArray = Nothing
      End If
  
      lCount = lCount + 1
    End If
  Next objSummaryField
  Set objSummaryField = Nothing

  lngControlLeft = lngLeftPos + lngXOFFSET_RIGHT + lngXOFFSET_LEFT

  lngCurrentYPosition = lngYOffset + lngYLABELOFFSET

  ' Loop through the recordset adding the required labels to the right column on the form.
  For Each objSummaryField In objSummaryFields.Collection
    If objSummaryField.Sequence > lLeftColumns Then
      lLabelCount = lLabelCount + 1
      
      Load Label1(lLabelCount)
  
      With Label1(lLabelCount)
        .Visible = True
         .Left = lngControlLeft
         
        .Caption = RemoveUnderScores(objSummaryField.ColumnName) & " :"
  
        If ((lLabelCount - lLeftColumns) > 1) And objSummaryField.StartOfGroup Then
          lngCurrentYPosition = lngCurrentYPosition + lngYSTARTGROUPOFFSET
        End If
        .Top = lngCurrentYPosition
  
        If .Top > lLabelTop Then
          lLabelTop = .Top
        End If
        
        If .Width + .Left > lngLeftPos Then
          lngLeftPos = .Width + .Left
        End If
        
        If .Width > lngRightLabelSize Then lngRightLabelSize = .Width

        lngCurrentYPosition = lngCurrentYPosition + lngYCONTROLOFFSET
  
      End With
    End If
  Next objSummaryField
  Set objSummaryField = Nothing

  lngControlLeft = lngLeftPos + lngXLABELOFFSET

  If fraSummary.Width - lngControlLeft - lngXOFFSET_RIGHT > lngMinControlSize Then
    lngControlSize = fraSummary.Width - lngControlLeft - lngXOFFSET_RIGHT
  Else
    lngControlSize = lngMinControlSize
  End If

  lCount = lLeftColumns + 1
  
  ' Loop through the recordset adding the required controls to the right column on the form.
  For Each objSummaryField In objSummaryFields.Collection
    If objSummaryField.Sequence > lLeftColumns Then
      iControlType = objSummaryField.ControlType
      iControlDataType = objSummaryField.DataType
    
      Set objControlArray = GetControlArray(iControlType, iControlDataType)

      If Not objControlArray Is Nothing Then
        Load objControlArray(objControlArray.Count)

        ' Populate array with Sequence and Index of the Control.
        aryControlArray(0, objSummaryField.Sequence - 1) = objSummaryField.Sequence
        aryControlArray(1, objSummaryField.Sequence - 1) = objControlArray.Count - 1

        Set objNewControl = objControlArray(objControlArray.Count - 1)

        With objNewControl
          .Visible = True
          .Left = lLeftWidth
          .Top = Label1(lCount).Top - lngYLABELOFFSET
          .Tag = objSummaryField.ColumnName
  
          If mfrmParent.ParentViewID > 0 Then
            ' Check if the column is in the parent view.
            Set objTableView = gcoTablePrivileges.FindViewID(mfrmParent.ParentViewID)
            Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)
            
            fColumnOK = objColumnPrivileges.IsValid(objSummaryField.ColumnName)
            If fColumnOK Then
              fColumnOK = objColumnPrivileges.Item(objSummaryField.ColumnName).AllowSelect
            End If
            
            If Not fColumnOK Then
              .BackColor = COL_GREY
              .Enabled = False
              .Tag = "0"
            End If
          End If
  
          If .Name = "Text1" Then
            .AlignHorizontal = objSummaryField.Alignment
          End If
            
          If .Name = "TDBNumber1" Then
            'JPD 20041109 Fault 8230
            datGeneral.FormatTDBNumberControl objNewControl
                   
            sFormat = ""
            iDigitCount = 1
            
            ' Loop and create the format mask
            For iCount = 1 To (objSummaryField.Size - objSummaryField.Decimals)
              If objSummaryField.Use1000Separator Then
                sFormat = IIf(iCount Mod 3 = 0 And (iCount <> (objSummaryField.Size - objSummaryField.Decimals)), ",#", "#") & sFormat
              Else
                sFormat = "#" & sFormat
              End If
            Next iCount

            If Not objSummaryField.BlankIfZero And Len(sFormat) > 0 Then
              sFormat = Left(sFormat, Len(sFormat) - 1) & "0"
            End If

            If objSummaryField.Decimals > 0 Then
              sFormat = sFormat & "."
              For iCount = 1 To objSummaryField.Decimals
                If objSummaryField.BlankIfZero Then
                  sFormat = sFormat & "#"
                Else
                  sFormat = sFormat & "0"
                End If
              Next iCount
            End If
            
            .DisplayFormat = sFormat
            .Format = sFormat
      
            'JPD 20041109 Fault 9429
            .MaxValue = Val(String(objSummaryField.Size - objSummaryField.Decimals, "9") & "." & String(objSummaryField.Decimals, "9"))
            .MinValue = (.MaxValue * -1)
          End If

          .Left = lngControlLeft

          If iControlType <> ctlCheck Then
            .Width = lngControlSize
          End If

          lCount = lCount + 1
        End With
      
        Set objControlArray = Nothing
      End If
    End If
  Next objSummaryField
  Set objSummaryField = Nothing

  mdblMinSummaryWidth = lngLeftLabelSize + lngRightLabelSize _
                        + (2 * lngXLABELOFFSET) + (2 * (lngXOFFSET_LEFT + lngXOFFSET_RIGHT)) _
                        + (2 * lngMinControlSize)
 
  With fraSummary
    .Height = lLabelTop + 500
    .Width = mdblMinSummaryWidth
  End With

  mlngMinFormWidth = mdblMinSummaryWidth + (2 * fraSummary.Left)
  mlngMinFormHeight = dblFINDFORM_MINHEIGHT

  lParent = mfrmParent.ParentFormID
  For Each frmForm In Forms
    If (frmForm.Name = "frmRecEdit4") Then
      If frmForm.FormID = lParent Then
        Set mrsParentData = frmForm.Recordset
        Exit For
      End If
    End If
  Next

  'Begin subclassing (set minimum form size)
  Unhook Me.hWnd
  Hook Me.hWnd, mlngMinFormWidth, mlngMinFormHeight

  ' the required controls are now loaded, so populate them.
  LoadSummaryDetails
  blnSummaryDetailsLoaded = True

End Sub

Public Sub UpdateSummaryWindow()
  ' The parent data has changed, so update the find form.
  Dim lParent As Long
  Dim frmForm As Form

  
  SetFormCaption Me, mfrmParent.FindCaption

  ' Get an updated recordset.
  GetFindRecords
  
  If mblnRefreshing Then
    Exit Sub
  End If
  
  
  ' JPD20021206 Fault 4843 - Get a handle on the parent recordset again
  ' in case its been lost
  lParent = mfrmParent.ParentFormID
  For Each frmForm In Forms
    If (frmForm.Name = "frmRecEdit4") Then
      If frmForm.FormID = lParent Then
        Set mrsParentData = frmForm.Recordset
        Exit For
      End If
    End If
  Next

  ' JPD20021126 Fault 4676
  If Recordset.State = adStateClosed Then
    Exit Sub
  End If
  
  With ssOleDBGridFindColumns
    If .Rows > 0 Then
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  'TM20011219 Fault 3253 - so scrollbars are refreshed.
  ResizeFindColumns

  ' JPD20021126 Fault 4676
  If Recordset.State = adStateClosed Then
    Exit Sub
  End If

  'MH20001124 Fault 1335
  'Edit and delete still enabled when shouldn't be !
  FormatControls
  frmMain.RefreshRecordMenu Me
    
  ' Refresh the summary controls.
  LoadSummaryDetails

End Sub


Private Function GetControlArray(ByVal piControlType As Integer, ByVal piDataType As Integer) As Object
  ' Returns a control array of the approriate type
  Dim objControlArray As Object
 
  Set objControlArray = Nothing
  
  ' Decide on the control type that is being requested
  Select Case piControlType
    Case ctlCheck
      Set objControlArray = Me.Check1
    Case Else
      If piDataType = sqlNumeric Then
        Set objControlArray = Me.TDBNumber1
      Else
        Set objControlArray = Me.Text1
      End If
  End Select
  
  Set GetControlArray = objControlArray
  
End Function


Public Property Get ScreenType() As ScreenType
  ' Get's the Screen Type value.
  ScreenType = mvarScreenType
  
End Property

Public Property Let ScreenType(vData As ScreenType)
  ' Set's the Screen Type Value.
  mvarScreenType = vData
  
End Property

Public Property Let Cancelling(pfNewValue As Boolean)
  ' JPD20020926 Fault 4440
  mfCancelled = pfNewValue
  
End Property


Private Function LoadSummaryDetails()
  ' Load the summary controls with the summary values.
  Dim sField As String
  Dim ctlTemp As Control
  Dim sTag As String
    
  On Error GoTo Err_Trap
        
  For Each ctlTemp In Me.Controls
    With ctlTemp
      sTag = .Tag
      
      'JPD 20030610
      If TypeOf ctlTemp Is ActiveBar Then
        sTag = ""
      End If
      
      If Len(sTag) > 0 Then
        sField = .Tag
        
        If sField <> "0" Then
          
          If TypeOf ctlTemp Is CheckBox Then
            If IsNull(mrsParentData.Fields(sField)) Then
              .Value = 0
            Else
              .Value = IIf(mrsParentData.Fields(sField), 1, 0)
            End If
          
          ElseIf TypeOf ctlTemp Is TDBNumber6Ctl.TDBNumber Then
            .Value = IsNullCheck(mrsParentData.Fields(sField), 0)
          
          Else
            .Text = IsNullCheck(mrsParentData.Fields(sField), "")
          
          End If
        Else
          .BackColor = COL_GREY
          .Enabled = False
        End If
      End If
    End With
  Next ctlTemp
  Set ctlTemp = Nothing
  
  Exit Function
  
Err_Trap:
  If Err.Number = 3265 Then
    ctlTemp.BackColor = COL_GREY
    ctlTemp.Enabled = False
    Resume Next
  End If
    
End Function


Public Sub Rebind()
  
  GetFindRecords

  If Not mfFindForm Then
    LoadSummaryDetails
  End If
  
End Sub
Public Property Let ParentFormID(plngData As Long)
  ' Set's the Parents Form ID.
  mlngParentFormID = plngData
  
End Property


Public Property Get ParentFormID() As Long
  ' Get's the Parents Form ID.
  ParentFormID = mlngParentFormID
  
End Property

Private Sub FormatControls()
  
  On Error GoTo ErrorTrap
  
  ' Format the command controls in the History frame.
  Dim fFindFromEmployeeRecords As Boolean
  Dim fFindFromCourseRecords As Boolean
  Dim fFindWaitingList As Boolean
  Dim fFindTrainingBooking As Boolean
  
  Dim objTableView As CTablePrivilege
  Dim colColumnPrivileges As CColumnPrivileges
  Dim blnIsCancelledCourse As Boolean
  
  Dim frmTemp As Form
  Dim strSQL As String
  Dim rsTemp As Recordset
  
  If gfTrainingBookingEnabled Then
    fFindFromEmployeeRecords = (mfrmParent.ParentTableID = glngEmployeeTableID)
    fFindFromCourseRecords = (mfrmParent.ParentTableID = glngCourseTableID)
    
'    'Check if training booking form has been opened from course records
'    'and if the course is cancelled
    blnIsCancelledCourse = False
    If Not (mfrmParent Is Nothing) Then
      For Each frmTemp In Forms
        With frmTemp
          If .Name = "frmRecEdit4" Then
            If (.FormID = mfrmParent.ParentFormID) Then
              If .TableID = glngCourseTableID Then

                Set objTableView = .TableView
                Set colColumnPrivileges = .ColumnSelectPrivileges

                If colColumnPrivileges.IsValid(gsCourseCancelDateColumnName) Then
                  If colColumnPrivileges.Item(gsCourseCancelDateColumnName).AllowSelect Then
                    strSQL = "SELECT " & gsCourseCancelDateColumnName & _
                      " FROM " & objTableView.RealSource & _
                      " WHERE id = " & CStr(.RecordID)
                    Set rsTemp = datGeneral.GetRecords(strSQL)
                    
                    If Not IsNull(rsTemp.Fields(gsCourseCancelDateColumnName).Value) Then
                      blnIsCancelledCourse = True
                    End If
                    
                    rsTemp.Close
                    Set rsTemp = Nothing
                  End If
                End If
                
                Exit For
              End If
            End If
          End If
        End With
      Next frmTemp
    End If
    Set frmTemp = Nothing
    
    fFindWaitingList = (mobjTableView.TableID = glngWaitListTableID)
    fFindTrainingBooking = (mobjTableView.TableID = glngTrainBookTableID)
  Else
    fFindFromEmployeeRecords = False
    fFindFromCourseRecords = False
  
    fFindWaitingList = False
    fFindTrainingBooking = False
  End If
  
  ' Display the module menu options as required.
  With Me.ActiveBar1.Bands(0)
    'JPD 20050113 Fault 8787
    mfBookCourseVisible = (fFindFromEmployeeRecords And fFindWaitingList)
    mfAddFromWaitingListVisible = (fFindFromCourseRecords And fFindTrainingBooking)
    mfCancelBookingVisible = ((fFindFromEmployeeRecords Or fFindFromCourseRecords) And fFindTrainingBooking)
    mfTransferVisible = ((fFindFromEmployeeRecords Or fFindFromCourseRecords) And fFindTrainingBooking)
    mfBulkBookingVisible = (fFindFromCourseRecords And fFindTrainingBooking)
    
    .Tools("BookCourseFind").Visible = mfBookCourseVisible And .Tools("BookCourseFind").Visible
    .Tools("AddFromWaitingListFind").Visible = mfAddFromWaitingListVisible And .Tools("AddFromWaitingListFind").Visible
    .Tools("CancelBookingFind").Visible = mfCancelBookingVisible And .Tools("CancelBookingFind").Visible
    .Tools("TransferFind").Visible = mfTransferVisible And .Tools("TransferFind").Visible
    .Tools("TransferFind").BeginGroup = Not .Tools("AddFromWaitingListFind").Visible
    .Tools("BulkBookingFind").Visible = mfBulkBookingVisible And .Tools("BulkBookingFind").Visible
  End With
    
  ' AddFromWaitingListFind option requires the user to have 'new' permission on the
  ' Training Bookings table, and delete permission on the Waiting List table.
  mfCanBookCourse = False
  mfCanAddFromWaitingList = False
  mfCanCancelBooking = False
  mfCanTransferBooking = False
  mfCanBulkBooking = False

  ' JPD20030203 Fault 4995
  If gfTrainingBookingEnabled Then
    Set objTableView = gcoTablePrivileges.FindTableID(glngTrainBookTableID)
    If Not objTableView Is Nothing Then
      mfCanAddFromWaitingList = (fFindFromCourseRecords And fFindTrainingBooking) And objTableView.AllowInsert And Not blnIsCancelledCourse
      mfCanCancelBooking = ((fFindFromEmployeeRecords Or fFindFromCourseRecords) And fFindTrainingBooking) And objTableView.AllowUpdate
      mfCanTransferBooking = ((fFindFromEmployeeRecords Or fFindFromCourseRecords) And fFindTrainingBooking) And (objTableView.AllowUpdate And objTableView.AllowInsert)
      mfCanBulkBooking = (fFindFromCourseRecords And fFindTrainingBooking) And objTableView.AllowInsert And Not blnIsCancelledCourse
      mfCanBookCourse = (fFindFromEmployeeRecords And fFindWaitingList) And objTableView.AllowInsert
    End If
    Set objTableView = Nothing
    
    If (mfCanAddFromWaitingList Or mfCanBookCourse) Then
      Set objTableView = gcoTablePrivileges.FindTableID(glngWaitListTableID)
      If objTableView Is Nothing Then
        mfCanAddFromWaitingList = False
        mfCanBookCourse = False
      Else
        mfCanAddFromWaitingList = (mfCanAddFromWaitingList And objTableView.AllowDelete)
        mfCanBookCourse = (mfCanBookCourse And objTableView.AllowDelete)
      End If
      Set objTableView = Nothing
    End If
  End If
  
  strSQL = "SELECT COUNT(*) FROM ASRSysCustomReportsName " & _
           "WHERE baseTable = " & CStr(mobjTableView.TableID)
  mfCustomReportExists = (GetRecCount(strSQL) > 0)

  strSQL = "SELECT COUNT(*) FROM ASRSysCalendarReports " & _
           "WHERE baseTable = " & CStr(mobjTableView.TableID)
  mfCalendarReportExists = (GetRecCount(strSQL) > 0)

  strSQL = "SELECT COUNT(*) FROM ASRSysGlobalFunctions " & _
           "WHERE TableID = " & CStr(mobjTableView.TableID) & " AND Type = 'U'"
  mfGlobalUpdateExists = (GetRecCount(strSQL) > 0)
  
  strSQL = "SELECT COUNT(*) FROM ASRSysDataTransferName " & _
           "WHERE FromTableID = " & CStr(mobjTableView.TableID)
  mfDataTransferExists = (GetRecCount(strSQL) > 0)
  
  strSQL = "SELECT COUNT(*) FROM ASRSysMailMergeName " & _
           "WHERE TableID = " & CStr(mobjTableView.TableID) & " AND IsLabel = 0"
  mfMailMergeExists = (GetRecCount(strSQL) > 0)
  
  Exit Sub

ErrorTrap:
  'fOK = False
  Select Case Err.Number
    Case 3021 'either BOF or EOF
      'When you COPY the parent record, the frmTemp.RecordID hasn't been
      'set up yet - the cancelation date will not be found in the resulting recordset.
      'At this point there is no harm in ignoring this error
      Err.Clear
      Resume Next
  End Select
  
End Sub

Public Property Get CustomReportExists() As Boolean
  CustomReportExists = mfCustomReportExists
End Property

Public Property Get CalendarReportExists() As Boolean
  CalendarReportExists = mfCalendarReportExists
End Property

Public Property Get GlobalUpdateExists() As Boolean
  GlobalUpdateExists = mfGlobalUpdateExists
End Property

Public Property Get DataTransferExists() As Boolean
  DataTransferExists = mfDataTransferExists
End Property

Public Property Get MailMergeExists() As Boolean
  MailMergeExists = mfMailMergeExists
End Property


Private Function ConfigureOrdersCombo() As Boolean
  ' Initialise the form to be called from a primary screen.
  Dim fOK As Boolean
  Dim fOrderFound As Boolean
  Dim iIndex As Integer
  Dim rsOrder As Recordset

  fOK = True
  
  ' Get the set of orders for the current table/view.
  If mobjTableView.IsTable Then
    Set rsOrder = datGeneral.GetRuntimeOrders(mobjTableView.TableID)
  Else
    Set rsOrder = datGeneral.GetRuntimeViewOrders(mobjTableView.ViewID, mobjTableView.TableID)
  End If

  ' Populate the Orders combo.
  With cmbOrders
    .Clear
  
    Do While Not rsOrder.EOF
      .AddItem RemoveUnderScores(Trim(rsOrder!Name))
      .ItemData(.NewIndex) = rsOrder!OrderID
      rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
        
    If .ListCount > 0 Then
      fOrderFound = False
      For iIndex = 0 To (.ListCount - 1)
        If (.ItemData(iIndex) = mlngOrderID) Then
          fOrderFound = True
          .ListIndex = iIndex
          Exit For
        End If
      Next iIndex
      
      If Not fOrderFound Then
        .ListIndex = 0
      End If
      
      .Enabled = True
    Else
      COAMsgBox "No orders defined for this " & IIf(mobjTableView.IsTable, "table.", "view.") & _
        vbCrLf & "Unable to display records.", vbExclamation, "Security"
      fOK = False
    End If
  End With
  
  ConfigureOrdersCombo = fOK
  
End Function


Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  
  'MH20061213 Fault 11785
  If mintResizeCount > 0 Then
    Exit Sub
  End If


  ' Perform the given toolbar function.
  Select Case Tool.Name
    Case "NewRecordFind":
      NewRecord
      
    Case "CopyRecord"
      AddNewCopyOf

    Case "EditFind":
      EditRecord
      
    Case "DeleteFind":
      DeleteRecord
      
    Case "BookCourseFind":
      BookCourse
      
    Case "AddFromWaitingListFind":
      AddFromWaitingList
      
    Case "CancelBookingFind":
      CancelBooking
      
    Case "TransferFind":
      TransferBooking
      
      'TM09092003 Fault 5682
      frmMain.RefreshMainForm Screen.ActiveForm
      
    Case "BulkBookingFind":
      BulkBooking
      
    Case "Filter"
      SelectFilter
      ' JDM - Fault 1910 - Was not disabling the history menu when "doin' stuff"
      ' JPD20030116 Fault 4942
      'frmMain.RefreshMainForm Me
      frmMain.RefreshMainForm Screen.ActiveForm

    Case "FilterClear"
      ClearFilter
      ' JDM - Fault 1910 - Was not disabling the history menu when "doin' stuff"
      ' JPD20030116 Fault 4942
      'frmMain.RefreshMainForm Me
      frmMain.RefreshMainForm Screen.ActiveForm
      
    Case "ID_Print"
      PrintGrid
      
    Case "CustomReports"
      UtilityClick utlCustomReport
      
    Case "CalendarReports"
      UtilityClick utlCalendarReport
      
    Case "GlobalUpdate"
      UtilityClick utlGlobalUpdate
    
    Case "MailMerge"
      UtilityClick utlMailMerge
    
    Case "DataTransfer"
      UtilityClick utlDataTransfer
  
  End Select

  UpdateStatusBar

End Sub

Public Sub SelectFilter()
  
  mfrmParent.SelectFilter
  Rebind

  With ssOleDBGridFindColumns
    If .Rows > 0 Then
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  ' NPG20100528 Fault 702
  ' ensure any scrollbars are displayed as required
  ResizeFindColumns
  
  ' JPD20030116 Fault 4942
  'frmMain.RefreshMainForm Me
  frmMain.RefreshMainForm Screen.ActiveForm

  Me.Width = Me.Width + 1

End Sub

Public Sub ClearFilter()

  mfrmParent.ClearFilter
  Rebind
  
  With ssOleDBGridFindColumns
    If .Rows > 0 Then
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  ' NPG20100528 Fault 702
  ' ensure any scrollbars are displayed as required
  ResizeFindColumns
  
  ' JPD20030116 Fault 4942
  'frmMain.RefreshMainForm Me
  frmMain.RefreshMainForm Screen.ActiveForm

End Sub

Private Sub cmbOrders_Click()
  Dim fOK As Boolean
  
  Screen.MousePointer = vbHourglass
  
  ' Do nothing if the form is not visible.
  fOK = Me.Visible

  ' Do nothing if there are no orders defined.
  If fOK Then
    fOK = (cmbOrders.ItemData(cmbOrders.ListIndex) > 0)
  End If

  If fOK Then
    ' Set the order ID variable.
    mlngOrderID = cmbOrders.ItemData(cmbOrders.ListIndex)
    
    ' Requery the recordset with the new order.
    If Not mrsFindRecords Is Nothing Then
      mrsFindRecords.Close
      Set mrsFindRecords = Nothing
    End If
    
    If GetFindRecords Then
      If Not (mrsFindRecords.BOF And mrsFindRecords.EOF) Then
        mrsFindRecords.MoveFirst
      End If
    Else
      cmbOrders.ListIndex = miOldOrderIndex
    End If
  
  End If
  
  ResizeFindColumns
  
  Screen.MousePointer = vbDefault
  
End Sub


Public Sub AddNew()
  'MH20010523
  ' Add a new record.
  'If mfFindForm Then
  '  With mfrmParent
  '    .AddNew
  '    .Visible = True
  '    .SetFocus
  '  End With
  'Else
  '  NewRecord
  'End If
  NewRecord

End Sub

Public Sub AddFromWaitingList()
  ' Training Booking specific function.
  '
  ' This function is available when the find window is displaying the Training Booking records
  ' for a Course.
  ' It displays the list of delegates that are waiting to go on the current course,
  ' and allows the user to book the delegates onto the course.
  Dim lngSelectedCourseID As Long
  Dim frmDelegateSelection As frmAddFromWaitingList
  Dim objCourseTableView As CTablePrivilege
  
  ' If the Training Booking module is enabled ...
  If gfTrainingBookingEnabled Then
    ' Get the current Course record ID.
    lngSelectedCourseID = mfrmParent.ParentID
  
    If lngSelectedCourseID > 0 Then
      ' Get the current Course table/view object.
      If mfrmParent.ParentViewID > 0 Then
        Set objCourseTableView = gcoTablePrivileges.FindViewID(mfrmParent.ParentViewID)
      Else
        Set objCourseTableView = gcoTablePrivileges.FindTableID(mfrmParent.ParentTableID)
      End If
      
      If Not objCourseTableView Is Nothing Then
        Set frmDelegateSelection = New frmAddFromWaitingList
        With frmDelegateSelection
          ' Initialise the 'Add From Waiting List' form to show all
          ' personnel waiting for the current course.
          If .Initialise(lngSelectedCourseID, objCourseTableView) Then
            .Show vbModal
      
            If Not .Cancelled Then
              ' Refresh the grid.
              mrsFindRecords.Requery
              mfrmParent.Requery False
              Requery False
            
              ' JPD20021209 Fault 4866
              With ssOleDBGridFindColumns
                If .Rows > 0 Then
                  .MoveFirst
                  .SelBookmarks.Add .Bookmark
                End If
              End With
              frmMain.RefreshMainForm Me
            End If
          Else
            With ssOleDBGridFindColumns
                If .SelBookmarks.Count = 0 Then
                  If .Rows > 0 Then
                    .MoveFirst
                    .SelBookmarks.Add .Bookmark
                  End If
                End If
            End With
          End If
        End With
        Unload frmDelegateSelection
        Set frmDelegateSelection = Nothing
      End If
      
      Set objCourseTableView = Nothing
    Else
      ' No selected course record.
      COAMsgBox "No course record selected.", vbOKOnly + vbInformation, app.ProductName
    End If
  End If
  
End Sub

Public Sub Requery(pfReset As Boolean)
  ' This function requeries the database
  With ssOleDBGridFindColumns
    .SelBookmarks.RemoveAll
    .Rebind
  End With
  
End Sub


Public Sub BookCourse()
  ' Training Booking specific command button.
  
  ' This function is available when the find window is displaying the Waiting List records
  ' for an Employee.
  ' It displays the list of Courses that are match the selected Waiting List course,
  ' and allows the user to book the current employee on the course.
  Dim fOK As Boolean
  Dim lngSelectedEmployeeID As Long
  Dim lngWaitingListID As Long
  Dim sSQL As String
  Dim sSelectedCourseTitle As String
  Dim vBookMark As Variant
  Dim frmCourseSelection As frmBookCourse
  Dim objColumns As CColumnPrivileges
  Dim rsInfo As Recordset
  
  fOK = gfTrainingBookingEnabled
  
  If Not fOK Then
    COAMsgBox "Training Booking functionality is not enabled.", vbOKOnly + vbInformation, app.ProductName
  Else
    ' Get the ID of the selected Waiting List record.
    lngWaitingListID = SelectedRecordID
  
    fOK = (lngWaitingListID > 0)
    If Not fOK Then
      COAMsgBox "No Waiting List record selected.", vbOKOnly + vbInformation, app.ProductName
    End If
  End If
  
  If fOK Then
    ' Get the Course Title from the Waiting List if the user has permission to read it.
    Set objColumns = GetColumnPrivileges(msCurrentTableViewName)
    fOK = objColumns.IsValid(gsWaitListCourseTitleColumnName)
    If Not fOK Then
      COAMsgBox "The '" & gsWaitListCourseTitleColumnName & "' column is not the current view.", vbOKOnly + vbInformation, app.ProductName
    End If
  End If
  
  If fOK Then
    fOK = objColumns.Item(gsWaitListCourseTitleColumnName).AllowSelect
    If Not fOK Then
      COAMsgBox "You do not have 'read' permission on the '" & gsWaitListCourseTitleColumnName & "'.", vbOKOnly + vbInformation, app.ProductName
    End If
  End If

  If fOK Then
    ' Get the course title and the employee id of the selected Waiting List record.
    sSQL = "SELECT " & gsWaitListCourseTitleColumnName & ", " & _
      "ID_" & Trim(Str(glngEmployeeTableID)) & _
      " FROM " & mobjTableView.RealSource & _
      " WHERE id = " & Trim(Str(lngWaitingListID))
    Set rsInfo = datGeneral.GetRecords(sSQL)
    fOK = Not (rsInfo.EOF And rsInfo.BOF)
    If Not fOK Then
      COAMsgBox "Unable to read the Course Title from the Waiting List record.", vbOKOnly + vbInformation, app.ProductName
    Else
      sSelectedCourseTitle = IIf(IsNull(rsInfo.Fields(gsWaitListCourseTitleColumnName)), "", rsInfo.Fields(gsWaitListCourseTitleColumnName))
      lngSelectedEmployeeID = IIf(IsNull(rsInfo.Fields("ID_" & glngEmployeeTableID)), 0, rsInfo.Fields("ID_" & glngEmployeeTableID))
    End If
    rsInfo.Close
    Set rsInfo = Nothing
  End If
  
  If fOK Then
    ' Initialise the 'Book Course' form to show all
    ' courses that match the required one.
    Set frmCourseSelection = New frmBookCourse
    With frmCourseSelection
      If .Initialise(sSelectedCourseTitle, lngSelectedEmployeeID, lngWaitingListID) Then
        .Show vbModal

        If Not .Cancelled Then
          ' Refresh the grid.
          mrsFindRecords.Requery
          mfrmParent.Requery True
          vBookMark = ssOleDBGridFindColumns.SelBookmarks.Item(0)
          
          Requery False
          'NHRD24012007 Fault 10081
          Screen.MousePointer = vbDefault
          ssOleDBGridFindColumns.SelBookmarks.Add vBookMark
          With ssOleDBGridFindColumns
            If .SelBookmarks.Count = 0 Then
              If .Rows > 0 Then
                .SelBookmarks.Add .Bookmark
                If vBookMark = 0 Then
                  .MoveFirst
                Else
                  .Bookmark = vBookMark
                End If
              End If
            Else
              If vBookMark = 0 Then
                .MoveFirst
              Else
                .Bookmark = vBookMark
              End If
            End If
          End With
  
        End If
      End If
    End With
    Unload frmCourseSelection
    Set frmCourseSelection = Nothing
  End If
  frmMain.RefreshRecordMenu Me
  
End Sub



Private Function ConfigureGrid() As Boolean
  ' Configure the grid to display the required columns.
  Dim iLoop As Integer
  Dim lngWidth As Long
  Dim dblPreviousColumnWidth As Double
  
  UI.LockWindow Me.hWnd
  
  'Setting the form to disabled here stops the find window getting focus
  'when it shouldn't (i.e. when scrolling though RecEdit records!)
  Me.Enabled = False
  
  
  lngWidth = 0
  mfFormattingGrid = True
  
  With ssOleDBGridFindColumns
    .Redraw = False
    .Columns.RemoveAll

    For iLoop = 0 To (mrsFindRecords.Fields.Count - 1)
      .Columns.Add iLoop
      'TM20090707 HRPRO-52
      'Right, this might be a bigger change than I expect, but here goes...
      'Don't think we need to use the name of the db column as the Name attribute
      'of the find grid; this is set by the 'Caption' attribute. If it all goes
      'wrong, blame Nick, Nick, Nick!!
      
      ' JDM - 26/08/2009 - HRPRO-299 - Looks like Tim was right - this was a bigger change than expected. It messed up
      ' quick access find windows. Changed the below to set the column name to be just the loop number, guaranteeing
      ' uniqueness. Hopefully this will work and no one will ever see these comments again. If you do blame Nick, and
      ' if you are Nick then shame on you!!!
      If (UCase(mrsFindRecords.Fields(iLoop).Name) <> "ID") Then
        .Columns(iLoop).Name = iLoop ' mavFindColumns(0, iLoop) 'mrsFindRecords.Fields(iLoop).Name
      Else
        .Columns(iLoop).Name = mrsFindRecords.Fields(iLoop).Name
      End If
      .Columns(iLoop).Visible = (UCase(mrsFindRecords.Fields(iLoop).Name) <> "ID") And _
        (Left(mrsFindRecords.Fields(iLoop).Name, 1) <> "?")
      .Columns(iLoop).Caption = RemoveUnderScores(mrsFindRecords.Fields(iLoop).Name)
      .Columns(iLoop).Alignment = ssCaptionAlignmentLeft
      .Columns(iLoop).CaptionAlignment = ssColCapAlignUseColumnAlignment

      ' If the find column is a logic column then set the grid column style to be 'checkbox'.
      If mrsFindRecords.Fields.Item(iLoop).Type = adBoolean Then
        .Columns(iLoop).Style = ssStyleCheckBox
      End If

      'Has the user changed the width of this column
      ' dblPreviousColumnWidth = GetUserSetting("FindOrder" + LTrim$(Str(mlngOrderID)), mrsFindRecords.Fields(iLoop).Name, 0)
      dblPreviousColumnWidth = GetUserSetting("FindOrder" + LTrim$(Str(mlngOrderID)), Replace(mrsFindRecords.Fields(iLoop).Name, "_", " "), 0)
      If dblPreviousColumnWidth > 0 Then .Columns(iLoop).Width = dblPreviousColumnWidth

      ' Total the size of the grid columns
      If .Columns(iLoop).Visible Then
        lngWidth = lngWidth + .Columns(iLoop).Width
      End If
    Next iLoop

    ' Add the extra width to handle the scroll bar
    '.Width = lngWidth + (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX) + 20

'    'Update the find window
'    ResizeFindColumns

    mfFormattingGrid = False
    '.Rebind
    .Rows = RecordCount
    
    'MH20110121 HRPRO-1093 Put the resizefindcolumns bit after setting the rows so it know if we need a scrollbar or not
    ResizeFindColumns

    .Redraw = True

  End With
    
  'Adjust size of find window to fit the grid
  lngWidth = lngWidth + (fraOrders.Left * 2) + _
    (((UI.GetSystemMetrics(SM_CXFRAME) * 2) + _
    UI.GetSystemMetrics(SM_CXBORDER)) * Screen.TwipsPerPixelX)
      
  If ssOleDBGridFindColumns.Rows > ssOleDBGridFindColumns.VisibleRows Then
    lngWidth = lngWidth + (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX) + 20
  End If

  ' RH 13/10/00 - BUG 1121 - Only resize the window if its not min/max.
  'StatusBar1.Panels(1).Text = ssOleDBGridFindColumns.Rows & " Record" & IIf(ssOleDBGridFindColumns.Rows = 1, "", "s") & _
        IIf(ssOleDBGridFindColumns.SelBookmarks.Count > 1, " - " & ssOleDBGridFindColumns.SelBookmarks.Count & " Selected", "")
  'UpdateStatusBar
  
  
  'Setting the form to disabled here stops the find window getting focus
  'when it shouldn't (i.e. when scrolling though RecEdit records!)
  'Here is where we enable the form again.
  DoEvents
  Me.Enabled = True
  
  UI.UnlockWindow
  
  ConfigureGrid = True
  
End Function

Public Sub UpdateFindWindow()
  ' The parent data has changed, so update the find form.
  SetFormCaption Me, mfrmParent.FindCaption

  ' Get an updated recordset.
  GetFindRecords
  
  With ssOleDBGridFindColumns
    If .Rows > 0 Then
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With

  'MH20001124 Fault 1335
  'Edit and delete still enabled when shouldn't be !
  FormatControls
  ResizeFindColumns
  frmMain.RefreshRecordMenu Me
  
End Sub

Private Sub LocateRecordIDWithoutSelect(plngID As Long)
  ' Locate the given record in the recordset.
  Dim fFound As Boolean
  Dim iColumnDataType As Integer
  Dim sColumnName As String
  Dim vOrderValue As Variant
  Dim iComparisonResult As Integer
  Dim lngUpper As Long
  Dim lngLower As Long
  Dim lngJump As Long
  Dim varFoundBookmark As Variant
  Dim objColumn As CColumnPrivilege
  Dim varbkCurrentRecord As Variant
  
  With mrsFindRecords
    ' Check if we can determine the order column.
    If mlngFirstSortColumnID = 0 Then
      ' Move to the first record.
      If (Not .BOF) Then
        .MoveFirst
        
        'JPD 20030916 Fault 6979
        Do While Not !ID = plngID
          If (Not .EOF) Then .MoveNext
        
          If .EOF Then
            If (Not .BOF) Then .MoveFirst
            Exit Do
          End If
        Loop
      End If
    Else
      ' Check if the first order column is in the current table/view.
      fFound = False
      For Each objColumn In mfrmParent.ColumnSelectPrivileges
        If objColumn.ColumnID = mlngFirstSortColumnID Then
          fFound = True
          sColumnName = objColumn.ColumnName
          iColumnDataType = objColumn.DataType
        End If
      Next objColumn
      Set objColumn = Nothing

      ' JPD20030205 Fault 5020
      If (Not fFound) Or _
        (Not mfFirstOrderColumnIsFindColumn) Or _
        ((iColumnDataType <> sqlVarChar) And _
        (iColumnDataType <> sqlVarBinary) And _
        (iColumnDataType <> sqlNumeric) And _
        (iColumnDataType <> sqlInteger)) Then

        .MoveFirst
        .Find "ID = " & plngID
        If .EOF Then
          If (Not .BOF) Then .MoveFirst
        End If
      Else
        ' Binary search the recordset for the required record.
        vOrderValue = datGeneral.GetOrderValue(plngID, sColumnName, mobjTableView.RealSource)

        If IsEmpty(vOrderValue) Or IsNull(vOrderValue) Then
          .MoveFirst
          .Find "ID = " & plngID

          If .EOF Then
            ' Dodgy bit of recordset handling. I encountered errors when trying to moveFirst
            ' even though BOF was false. Doing the movePrevious and then the moveFirst sorted it out.
            If (Not .BOF) Then .MovePrevious
            If (Not .BOF) Then .MoveFirst
          End If
        Else
          fFound = False
          lngLower = 1
          lngUpper = RecordCount

          .MoveFirst

          Do
            Select Case iColumnDataType
              Case sqlVarChar, sqlVarBinary
                ' JPD String comparison changed from using VB's strComp function to
                ' using our own DictionaryCompareStrings function. VB's strComp
                ' function does not use the same order as that used when SQL orders
                ' by a character column. The DictionaryCompareStrings does.
                'iComparisonResult = StrComp(UCase(Left(IIf(IsNull(.Fields(sColumnName).Value), "", .Fields(sColumnName).Value), _
                  Len(vOrderValue))), UCase(vOrderValue), vbBinaryCompare)
                iComparisonResult = datGeneral.DictionaryCompareStrings(.Fields(sColumnName).Value, vOrderValue)

              Case sqlNumeric, sqlInteger
                If IsNull(.Fields(sColumnName).Value) Then
                  iComparisonResult = -1
                Else
                  If Val(.Fields(sColumnName).Value) = Val(vOrderValue) Then
                    iComparisonResult = 0
                  ElseIf Val(.Fields(sColumnName).Value) < Val(vOrderValue) Then
                    iComparisonResult = -1
                  Else
                    iComparisonResult = 1
                  End If
                End If
            End Select

            If Not mfFirstColumnAscending Then
              iComparisonResult = iComparisonResult * -1
            End If

            Select Case iComparisonResult
              Case 0    ' String found.
                fFound = True
                varFoundBookmark = .Bookmark
                lngUpper = .Bookmark - 1
                lngJump = -((.Bookmark - lngLower) \ 2) - 1
                If lngLower > lngUpper Then Exit Do

              Case -1   ' Current record is before the required record.
                lngLower = .Bookmark + 1
                lngJump = ((lngUpper - .Bookmark) \ 2)
                If lngLower > lngUpper Then Exit Do

              Case 1    ' Current record is after the required record.
                lngUpper = .Bookmark - 1
                lngJump = -((.Bookmark - lngLower) \ 2) - 1
                If lngLower > lngUpper Then Exit Do
            End Select

            If lngLower = lngUpper Then
              lngJump = lngUpper - .Bookmark
            End If

            ' Move to the middle record of the remaining records to search.
            ' Only move forward if we're not on the EOF marker already.
            ' Only move back if we're not on the BOF marker already.
            If ((lngJump > 0) And (Not .EOF)) Or _
              ((lngJump < 0) And (Not .BOF)) Then
              .Move lngJump

              ' Check if we're now BOF or EOF.
              If .BOF Or .EOF Then
                Exit Do
              End If
            Else
              Exit Do
            End If
          Loop

          If fFound Then
            ' Find the record that has the same ID as the required one.
            .Bookmark = varFoundBookmark
            Do While Not !ID = plngID
              If (Not .EOF) Then .MoveNext

              If .EOF Then
                .Bookmark = varFoundBookmark
                Exit Do
              End If
            Loop
          Else
            ' Move to the first record.
            ' Dodgy bit of recordset handling. I encountered errors when trying to moveFirst
            ' even though BOF was false. Doing the movePrevious and then the moveFirst sorted it out.
            If (Not .BOF) Then .MovePrevious
            If (Not .BOF) Then .MoveFirst

          End If
        End If
      End If
    End If
  End With

  varbkCurrentRecord = mrsFindRecords.Bookmark
  ssOleDBGridFindColumns.MoveRecords varbkCurrentRecord
  ssOleDBGridFindColumns.Bookmark = varbkCurrentRecord
  CurrentBookMark = ssOleDBGridFindColumns.Bookmark
  
End Sub


Private Sub LocateRecordID(plngID As Long)
  
  LocateRecordIDWithoutSelect plngID
  
  'MH20020430 Fault 3809 RemoveAll bookmarks before setting current record...
  ssOleDBGridFindColumns.SelBookmarks.RemoveAll
  ssOleDBGridFindColumns.SelBookmarks.Add ssOleDBGridFindColumns.Bookmark

End Sub





Public Property Get RecordCount() As Long
  'JPD 20050810 Fault 8844
  'RecordCount = mfrmParent.RecordCount
  RecordCount = mrsFindRecords.RecordCount
  
End Property




Public Sub CancelBooking()
  ' Training Booking specific function.
  '
  ' This function is available when the find window is displaying the Training Booking records
  ' for a Course or Delegate.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fFound As Boolean
  Dim fInTransaction As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim iUserChoice As Integer
  Dim lngSelectedBookingID As Long
  Dim lngSelectedEmployeeID As Long
  Dim lngSelectedCourseID As Long
  Dim sSQL As String
  Dim sColumnList As String
  Dim sValueList As String
  Dim sSelectedBookingStatus As String
  Dim objTBColumn As CColumnPrivilege
  Dim objWLColumn As CColumnPrivilege
  Dim objTBColumnPrivileges As CColumnPrivileges
  Dim objWLColumnPrivileges As CColumnPrivileges
  Dim objTrainingBookingTable As CTablePrivilege
  Dim objWaitingListTable As CTablePrivilege
  Dim rsInfo As Recordset
  Dim alngRelatedColumns() As Long
  Dim asAddedColumns() As String
  Dim vBookMark As Variant
  Dim sErrorMsg As String
  Dim sCurrentCourseTitle As String
  Dim fColumnOK As Boolean
  Dim sColumnCode As String
  Dim sRealSource As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim objTableView As CTablePrivilege
  Dim asViews() As String
  Dim alngTableViews() As Long
  Dim fCourseNameFound As Boolean

  fOK = gfTrainingBookingEnabled
  fInTransaction = False
  
  If ssOleDBGridFindColumns.SelBookmarks.Count = 0 Then
    COAMsgBox "No record selected.", vbOKOnly + vbInformation, app.ProductName
    fOK = False
  End If
  
  If fOK Then
    ' Validate the required Training Bookings table parameters.
    ' Get the column privileges for the Training Bookings table.
    Set objTrainingBookingTable = gcoTablePrivileges.Item(gsTrainBookTableName)
    Set objTBColumnPrivileges = GetColumnPrivileges(gsTrainBookTableName)

    If fOK Then
      ' Check that the user has permission to update the Training Bookings Status column.
      fOK = objTBColumnPrivileges.Item(gsTrainBookStatusColumnName).AllowUpdate
      If Not fOK Then
        COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookStatusColumnName & "' column.", vbOKOnly + vbInformation, app.ProductName
        GoTo QuitWithoutRequery
      End If
    End If
  
    ' If the training booking cancellation date is defined, check that the current user can update it.
    If Len(gsTrainBookCancelDateColumnName) > 0 Then
      fOK = objTBColumnPrivileges.Item(gsTrainBookCancelDateColumnName).AllowUpdate
      If Not fOK Then
        COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookCancelDateColumnName & "' column.", vbOKOnly + vbInformation, app.ProductName
        GoTo QuitWithoutRequery
      End If
    End If
  End If

  If fOK Then
    mrsFindRecords.Bookmark = ssOleDBGridFindColumns.SelBookmarks.Item(0)
    lngSelectedBookingID = mrsFindRecords("ID").Value
    vBookMark = ssOleDBGridFindColumns.SelBookmarks.Item(0)
    
    ' Get the Booking details if the user has permission to read them.
    sSQL = "SELECT " & gsTrainBookStatusColumnName & ", " & _
      "ID_" & Trim(Str(glngEmployeeTableID)) & ", " & _
      "ID_" & Trim(Str(glngCourseTableID)) & _
      " FROM " & mobjTableView.RealSource & _
      " WHERE id = " & Trim(Str(lngSelectedBookingID))
    Set rsInfo = datGeneral.GetRecords(sSQL)
    If Not (rsInfo.EOF And rsInfo.BOF) Then
      lngSelectedEmployeeID = IIf(IsNull(rsInfo("ID_" & Trim(Str(glngEmployeeTableID))).Value), 0, rsInfo("ID_" & Trim(Str(glngEmployeeTableID))).Value)
      lngSelectedCourseID = IIf(IsNull(rsInfo("ID_" & Trim(Str(glngCourseTableID))).Value), 0, rsInfo("ID_" & Trim(Str(glngCourseTableID))).Value)
      sSelectedBookingStatus = IIf(IsNull(rsInfo(gsTrainBookStatusColumnName).Value), "", rsInfo(gsTrainBookStatusColumnName).Value)
    End If
    rsInfo.Close
    Set rsInfo = Nothing

    fOK = (lngSelectedEmployeeID > 0)
    If Not fOK Then
      COAMsgBox "The selected Training Booking record has no associated Employee record.", vbOKOnly + vbInformation, app.ProductName
      GoTo QuitWithoutRequery
    End If
    
    If fOK Then
      fOK = (lngSelectedCourseID > 0)
      If Not fOK Then
        COAMsgBox "The selected Training Booking record has no associated Course record.", vbOKOnly + vbInformation, app.ProductName
        GoTo QuitWithoutRequery
      End If
    End If
  End If

  If fOK Then
    ' Check that the booking status is either 'B' or 'P'.
    fOK = (UCase(Left(sSelectedBookingStatus, 1)) = "B") Or _
      (UCase(Left(sSelectedBookingStatus, 1)) = "P")

    If Not fOK Then
      COAMsgBox "Bookings can only be cancelled if they have 'Booked'" & IIf(gfTrainBookStatus_P, " or 'Provisional'", "") & " status.", vbOKOnly + vbInformation, app.ProductName
      GoTo QuitWithoutRequery:
    End If
  End If

  If fOK Then
    ' Get the name of the course.
    fCourseNameFound = True
    
    ' Dimension an array of tables/views joined to the base table/view.
    ' Column 1 = view ID.
    ReDim alngTableViews(0)
    
    sRealSource = gcoTablePrivileges.Item(gsCourseTableName).RealSource
    Set objColumnPrivileges = GetColumnPrivileges(gsCourseTableName)
    fColumnOK = objColumnPrivileges.Item(gsCourseTitleColumnName).AllowSelect
    Set objColumnPrivileges = Nothing
    If fColumnOK Then
      sSQL = "SELECT " & sRealSource & "." & Trim(gsCourseTitleColumnName) & _
        " FROM " & sRealSource & _
        " WHERE id = " & Trim(Str(lngSelectedCourseID))
    Else
      ' The column CANNOT be read from the Course table.
      ' Try to read it from the views on the table.
            
      ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
      ReDim asViews(0)
      For Each objTableView In gcoTablePrivileges.Collection
        If (Not objTableView.IsTable) And _
          (objTableView.TableID = glngCourseTableID) And _
          (objTableView.AllowSelect) Then
                    
          sRealSource = gcoTablePrivileges.Item(objTableView.ViewName).RealSource
        
          ' Get the column permission for the view.
          Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)
        
          fColumnOK = objColumnPrivileges.IsValid(gsCourseTitleColumnName)
          If fColumnOK Then
            fColumnOK = objColumnPrivileges.Item(gsCourseTitleColumnName).AllowSelect
          End If
          
          If fColumnOK Then
            ' Add the view info to an array to be put into the column list or order code below.
            iNextIndex = UBound(asViews) + 1
            ReDim Preserve asViews(iNextIndex)
            asViews(iNextIndex) = objTableView.ViewName
                            
            ' Add the view to the Join code.
            ' Check if the view has already been added to the join code.
            fFound = False
            For iNextIndex = 1 To UBound(alngTableViews)
              If alngTableViews(iNextIndex) = objTableView.ViewID Then
                fFound = True
                Exit For
              End If
            Next iNextIndex
                          
            If Not fFound Then
              ' The view has not yet been added to the join code, so add it to the array and the join code.
              iNextIndex = UBound(alngTableViews) + 1
              ReDim Preserve alngTableViews(iNextIndex)
              alngTableViews(iNextIndex) = objTableView.ViewID
            End If
          End If
          Set objColumnPrivileges = Nothing
        End If
      Next objTableView
      Set objTableView = Nothing
         
      ' The current user does have permission to 'read' the column through a/some view(s) on the
      ' table.
      If UBound(asViews) = 0 Then
        fCourseNameFound = False
        COAMsgBox "You do not have 'read' permission on the '" & gsCourseTitleColumnName & "'.", vbOKOnly + vbInformation, app.ProductName
      Else
        ' Add the column to the column list.
        sColumnCode = ""
        For iNextIndex = 1 To UBound(asViews)
          If iNextIndex = 1 Then
            sColumnCode = "CASE "
          End If
                  
          sColumnCode = sColumnCode & _
            " WHEN NOT " & asViews(iNextIndex) & "." & gsCourseTitleColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & gsCourseTitleColumnName
        Next iNextIndex
                  
        sSQL = "SELECT " & sColumnCode & _
          " ELSE NULL" & _
          " END AS " & gsCourseTitleColumnName & _
          " FROM " & gcoTablePrivileges.Item(gsCourseTableName).RealSource
          
        For iNextIndex = 1 To UBound(alngTableViews)
          Set objTableView = gcoTablePrivileges.FindViewID(alngTableViews(iNextIndex))
            
          ' Join a view of the Course table.
          sSQL = sSQL & _
            " LEFT OUTER JOIN " & objTableView.RealSource & _
            " ON " & gcoTablePrivileges.Item(gsCourseTableName).RealSource & ".ID = " & objTableView.RealSource & ".ID"
          
          Set objTableView = Nothing
        Next iNextIndex
        
        sSQL = sSQL & _
          " WHERE " & gcoTablePrivileges.Item(gsCourseTableName).RealSource & ".id = " & Trim(Str(lngSelectedCourseID))
      End If
    End If
    
    If fCourseNameFound Then
      Set rsInfo = datGeneral.GetRecords(sSQL)
      With rsInfo
        fCourseNameFound = Not (.EOF And .BOF)
        If fCourseNameFound Then
          fCourseNameFound = Not IsNull(.Fields(gsCourseTitleColumnName))
        End If
        
        If fCourseNameFound Then
          sCurrentCourseTitle = .Fields(gsCourseTitleColumnName)
        End If
        
        .Close
      End With
      Set rsInfo = Nothing
    End If
  End If
  
  If fOK Then
    If fCourseNameFound Then
      iUserChoice = COAMsgBox("Transfer the booking to the employee's waiting list ?", vbYesNoCancel + vbQuestion, app.ProductName)
    Else
      iUserChoice = vbNo
    End If
    
    If iUserChoice <> vbCancel Then
      Screen.MousePointer = vbHourglass

      gADOCon.BeginTrans
      fInTransaction = True
      
      ' Change the selected booking's status to 'C'.
      sSQL = "UPDATE " & mobjTableView.RealSource & _
        " SET " & gsTrainBookStatusColumnName & " = 'C'"

      If Len(gsTrainBookCancelDateColumnName) > 0 Then
        sSQL = sSQL & _
          ", " & gsTrainBookCancelDateColumnName & " = '" & Replace(Format(Date, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'"
      End If
        
      sSQL = sSQL & _
        " WHERE id = " & Trim(Str(lngSelectedBookingID))
      
      sErrorMsg = ""
      fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)

      If Not fOK Then
        Screen.MousePointer = vbDefault
        COAMsgBox "Unable to cancel the booking." & vbCrLf & vbCrLf & sErrorMsg, vbExclamation + vbOKOnly, app.ProductName
        
        Screen.MousePointer = vbHourglass
        
        gADOCon.RollbackTrans
        fInTransaction = False
      Else
        If iUserChoice = vbYes Then
          ' Check the user has permission to read all of the related columns.
          Set objWaitingListTable = gcoTablePrivileges.Item(gsWaitListTableName)
          Set objWLColumnPrivileges = GetColumnPrivileges(gsWaitListTableName)
            
          ' Initialise the string for transfering info from the Training Booking
          ' table back to the Waiting List table.
          ReDim asAddedColumns(1)
          asAddedColumns(1) = UCase(Trim(gsWaitListCourseTitleColumnName))
          sColumnList = gsWaitListCourseTitleColumnName & _
            ", id_" & Trim(Str(glngEmployeeTableID))
          sValueList = "'" & Replace(sCurrentCourseTitle, "'", "''") & "'" & _
            ", id_" & Trim(Str(glngEmployeeTableID))
          
          alngRelatedColumns = RelatedColumns
          
          ' Check that the user has 'read' permission on the related Training Booking columns,
          ' and 'edit' permission on the realted Waiting List columns.
          For iLoop = 1 To UBound(alngRelatedColumns, 2)
            Set objTBColumn = objTBColumnPrivileges.FindColumnID(alngRelatedColumns(1, iLoop))
            Set objWLColumn = objWLColumnPrivileges.FindColumnID(alngRelatedColumns(2, iLoop))
            
            fOK = Not objTBColumn Is Nothing
            If Not fOK Then
              COAMsgBox "Unable to find all related columns in the '" & gsTrainBookTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
              Exit For
            Else
              fOK = objTBColumn.AllowSelect
              If Not fOK Then
                COAMsgBox "You do not have 'read' permission on the '" & objTBColumn.ColumnName & "' column in the '" & gsTrainBookTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
                Exit For
              End If
            End If
            
            If fOK Then
              fOK = Not objWLColumn Is Nothing
              If Not fOK Then
                COAMsgBox "Unable to find all related columns in the '" & gsWaitListTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
                Exit For
              Else
                fOK = objWLColumn.AllowUpdate
                If Not fOK Then
                  COAMsgBox "You do not have 'edit' permission on the '" & objWLColumn.ColumnName & "' column in the '" & gsWaitListTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
                  Exit For
                End If
              End If
            End If
            
            If fOK Then
              ' Check that the Training Booking column has not already been added to the 'insert' string.
              fFound = False
              For iNextIndex = 1 To UBound(asAddedColumns)
                If UCase(Trim(objWLColumn.ColumnName)) = asAddedColumns(iNextIndex) Then
                  fFound = True
                  Exit For
                End If
              Next iNextIndex
            
              If Not fFound Then
                ' The current WL column is not in the 'insert' string so add it now,
                ' and add it to the array of added columns.
                sColumnList = sColumnList & _
                  ", " & objWLColumn.ColumnName
              
                iNextIndex = UBound(asAddedColumns) + 1
                ReDim Preserve asAddedColumns(iNextIndex)
                asAddedColumns(iNextIndex) = UCase(Trim(objWLColumn.ColumnName))
                
                sValueList = sValueList & _
                  ", " & objTBColumn.ColumnName
              End If
            End If
            
            Set objTBColumn = Nothing
            Set objWLColumn = Nothing
          Next iLoop
          
          If fOK Then
            ' Validate the required Waiting List table parameters.
            ' Check that the user has permission to insert records from the Waiting List table.
            fOK = objWaitingListTable.AllowInsert
            If Not fOK Then
              COAMsgBox "You do not have 'new' permission on the '" & gsWaitListTableName & "' table.", vbOKOnly + vbInformation, app.ProductName
            End If
          End If
          
          If fOK Then
            ' Check that the user has permission to see the Waiting List Course Title column.
            fOK = objWLColumnPrivileges.Item(gsWaitListCourseTitleColumnName).AllowUpdate
            If Not fOK Then
              COAMsgBox "You do not have 'edit' permission on the '" & gsWaitListCourseTitleColumnName & "' column.", vbOKOnly + vbInformation, app.ProductName
            End If
          End If
          
          If fOK Then
            ' Do not add the Waiting List records if the course title is already in the Waiting List.
            sSQL = "SELECT COUNT(id) AS recCount" & _
              " FROM " & objWaitingListTable.RealSource & _
              " WHERE id_" & Trim(Str(glngEmployeeTableID)) & " = " & Trim(Str(lngSelectedEmployeeID)) & _
              " AND " & gsWaitListCourseTitleColumnName & " = '" & Replace(sCurrentCourseTitle, "'", "''") & "'"
            Set rsInfo = datGeneral.GetRecords(sSQL)
            If rsInfo!recCount = 0 Then
  
              ' Transfer the booking to the employee's waiting list.
              sSQL = "INSERT INTO " & objWaitingListTable.RealSource & _
                " (" & sColumnList & ")" & _
                " SELECT " & sValueList & _
                " FROM " & objTrainingBookingTable.RealSource & _
                " WHERE id = " & Trim(Str(lngSelectedBookingID))
            
              sErrorMsg = ""
              fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)
              If Not fOK Then
                Screen.MousePointer = vbDefault
                COAMsgBox "Unable to create waiting list records." & vbCrLf & vbCrLf & sErrorMsg, vbOKOnly + vbInformation, app.ProductName
                Screen.MousePointer = vbHourglass
                  
                gADOCon.RollbackTrans
                fInTransaction = False
              End If
            End If
            rsInfo.Close
            Set rsInfo = Nothing
          End If
        End If
      End If
    End If
  End If
      
TidyUpAndExit:
  If fInTransaction Then
    If fOK Then
      gADOCon.CommitTrans
      ' JPD20011101 Fault 3076
      COAMsgBox "Booking cancelled.", vbOKOnly & vbInformation, app.ProductName
    Else
      gADOCon.RollbackTrans
    End If
    fInTransaction = False
  End If
  
  ' Refresh the grid.
  mrsFindRecords.Requery
  mfrmParent.Requery False
  Requery False
            
  Screen.MousePointer = vbDefault
      
  ssOleDBGridFindColumns.SelBookmarks.Add vBookMark
  
  With ssOleDBGridFindColumns
    If .SelBookmarks.Count = 0 Then
      If .Rows > 0 Then
        .SelBookmarks.Add .Bookmark
        If vBookMark = 0 Then
          .MoveFirst
        Else
          .Bookmark = vBookMark
        End If
      End If
    Else
      If vBookMark = 0 Then
        .MoveFirst
      Else
        .Bookmark = vBookMark
      End If
    End If
  End With
  
QuitWithoutRequery:
  'JPD 20030519 Fault 5698
  frmMain.RefreshMainForm Me

  Set rsInfo = Nothing
  Set objTBColumn = Nothing
  Set objWLColumn = Nothing
  Set objTBColumnPrivileges = Nothing
  Set objWLColumnPrivileges = Nothing
  Set objTrainingBookingTable = Nothing
  Set objWaitingListTable = Nothing
  
  Exit Sub
  
ErrorTrap:
  fOK = False
  COAMsgBox Err.Description, vbExclamation + vbOKOnly, Application.Name
  Resume TidyUpAndExit
  
End Sub
Public Property Get StatusCaption() As String
  StatusCaption = mfrmParent.FindStatusCaption

End Property


Public Sub DeleteRecord()
  ' Delete the selected record.
  'NHRD 22042002 Fault 3362 Enables multi-select deletions.
  'Added a lot of code to this sub
  'Old sub directly below
  Dim fDeleteOK As Boolean
  Dim lngSelectedRecordID As Long
  Dim lngOriginalRecordID As Long
  Dim nTotalSelRows As Variant
  Dim intCount As Integer
  
  'Declare Array to hold Bookmarks and another for the record IDs we want to delete
  Dim arrayBookmarks() As Variant
  Dim arrayDelRecordIDs() As Variant
  
  fDeleteOK = True
  
  'Workout how many records have been selected
  nTotalSelRows = ssOleDBGridFindColumns.SelBookmarks.Count
  
  If (nTotalSelRows > 1) Then
    If COAMsgBox("Delete the selected records, are you sure ?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Records") <> vbYes Then
      Exit Sub
    End If
  End If
  
  'Redimension the arrays to the count of the bookmarks
  ReDim arrayBookmarks(nTotalSelRows)
  ReDim arrayDelRecordIDs(nTotalSelRows)

  'Populate the array with bookmark indeces
  'These will have their bookmarks stored in .SelBookmarks.item(iIndex)
  For intCount = 1 To nTotalSelRows
    arrayBookmarks(intCount) = ssOleDBGridFindColumns.SelBookmarks.Item(intCount - 1)
  Next intCount
  
  'Now need to populate an array with all the IDs that we are going to delete
  For intCount = 1 To nTotalSelRows Step 1
    ssOleDBGridFindColumns.Bookmark = arrayBookmarks(intCount)
    CurrentBookMark = ssOleDBGridFindColumns.Bookmark
    arrayDelRecordIDs(intCount) = SelectedRecordID
  Next intCount
  ssOleDBGridFindColumns.SelBookmarks.RemoveAll

  If nTotalSelRows > 0 Then
    If mfrmParent.SaveChanges(False) Then
      If Not Database.Validation Then
        Exit Sub
      End If
  
      ' JPD20030306 Fault 5118
      mfBusy = True

      'JPD 20030905 Fault 5184
      frmMain.DisableMenu
      UI.LockWindow Me.hWnd

      With mfrmParent
        For intCount = 1 To nTotalSelRows
          lngOriginalRecordID = .RecordID
          lngSelectedRecordID = arrayDelRecordIDs(intCount)
          .LocateRecord lngSelectedRecordID
          
          ' If the located record differs from the selected
          ' record then the record must already be deleted.
          fDeleteOK = (lngSelectedRecordID = .RecordID)
          
          If fDeleteOK Then
            .DeleteRecord (nTotalSelRows = 1), True
          End If
          
'JPD 20030905 Fault 5184
'''          ' JPD 26/9/00 If the recordset is now in 'add mode' then
'''          ' the record must have been deleted okay.
'''          If .Recordset.EditMode <> adEditAdd Then
'''            .LocateRecord lngOriginalRecordID
'''
'''            If (.RecordID <> lngOriginalRecordID) Or (Not fDeleteOK) Then
'''              .Requery False
'''            Else
'''              .UpdateControls
'''              'NHRD 04032002  Fault3363 commenting out the following line
'''              'should keep the focus on the previously highlighted row,
'''              'wherever the deleteRecord sub is used.
'''              ''''.UpdateChildren
'''            End If
'''          End If
        Next intCount
        
        'JPD 20030905 Fault 5184
        .UpdateAll
      End With

      'JPD 20030905 Fault 5184
      UI.UnlockWindow
      frmMain.EnableMenu Me
    End If
  End If
  
  ssOleDBGridFindColumns.SelBookmarks.RemoveAll
  ssOleDBGridFindColumns.SelBookmarks.Add ssOleDBGridFindColumns.Bookmark

  'JPD 20030730 Fault 5221
  With ssOleDBGridFindColumns
    If .Rows > 0 Then
      'JPD 20050216 Fault 8858
      If .Enabled And .Visible Then
        .SetFocus
      End If
    Else
      'MH20070615 Fault 12333
      'If cmbOrders.Enabled Then
      '  cmbOrders.SetFocus
      'Else
      '  Me.SetFocus
      'End If
      cmbOrdersSetFocus
    End If
  End With

  If Me.Visible Then
    ' JPD20020925 Fault 4444
    frmMain.RefreshMainForm Me
  End If

  ' JPD20030306 Fault 5118
  mfBusy = False

End Sub


'MH20070615 Fault 12333
Private Function cmbOrdersSetFocus() As Boolean
  
  On Local Error Resume Next
  Err.Clear

  If cmbOrders.Enabled Then
    cmbOrders.SetFocus
  Else
    Me.SetFocus
  End If

  cmbOrdersSetFocus = (Err.Number = 0)

End Function


Public Sub EditRecord()
  ssOleDBGridFindColumns_DblClick

End Sub


Public Sub AddNewCopyOf()
  ssOleDBGridFindColumns_DblClick

  ' Makes this a new record.
  'Me.Visible = False
  
  ' Instruct the parent record editing form to create a new record.
  With mfrmParent
    .OriginalRecordID = .RecordID
    .AddNewCopyOf
    .Visible = True
    
    '#RH 16/11
    .Enabled = True
    .SetFocus
  End With

End Sub
Private Sub NewRecord()
  ' Add a new record.
  Me.Visible = False
  
  ' Instruct the parent record editing form to create a new record.
  With mfrmParent
    .AddNew
    

    .Visible = True
    
    '#RH 16/11
    .Enabled = True
    .SetFocus
  End With

End Sub

Public Sub TransferBooking()
  ' Training Booking specific function.
  '
  ' This function is available when the find window is displaying the Training Booking records
  ' for a Course or Delegate.
  ' It display a list of other courses with the same title, and allows the user
  ' to transfer the booking to another course.
  Dim fOK As Boolean
  Dim lngSelectedBookingID As Long
  Dim sSQL As String
  Dim sSelectedBookingStatus As String
  Dim frmCourseSelection As frmTransferBooking
  Dim objColumnPrivileges As CColumnPrivileges
  Dim rsInfo As Recordset
  
  fOK = gfTrainingBookingEnabled

  If fOK Then
    ' Get the selected Booking record ID.
    lngSelectedBookingID = SelectedRecordID
  
    ' Get the Booking Status if the user has permission to read it.
    Set objColumnPrivileges = GetColumnPrivileges(gsTrainBookTableName)
    fOK = objColumnPrivileges.Item(gsTrainBookStatusColumnName).AllowSelect
    If Not fOK Then
      COAMsgBox "You do not have 'read' permission on the '" & gsTrainBookStatusColumnName & "' column.", vbOKOnly + vbInformation, app.ProductName
    Else
      fOK = objColumnPrivileges.Item(gsTrainBookStatusColumnName).AllowUpdate
      If Not fOK Then
        COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookStatusColumnName & "' column.", vbOKOnly + vbInformation, app.ProductName
      End If
    End If
    
    ' If the training booking cancellation date is defined, check that the current user can update it.
    If Len(gsTrainBookCancelDateColumnName) > 0 Then
      fOK = objColumnPrivileges.Item(gsTrainBookCancelDateColumnName).AllowUpdate
      If Not fOK Then
        COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookCancelDateColumnName & "' column.", vbOKOnly + vbInformation, app.ProductName
      End If
    End If
    
    Set objColumnPrivileges = Nothing
  End If
  
  If fOK Then
    sSQL = "SELECT " & gsTrainBookStatusColumnName & _
      " FROM " & mobjTableView.RealSource & _
      " WHERE id = " & Trim(Str(lngSelectedBookingID))
    Set rsInfo = datGeneral.GetRecords(sSQL)
    fOK = Not (rsInfo.EOF And rsInfo.BOF)
    If fOK Then
      fOK = Not IsNull(rsInfo.Fields(gsTrainBookStatusColumnName))
    End If
    
    If fOK Then
      sSelectedBookingStatus = rsInfo.Fields(gsTrainBookStatusColumnName)
    Else
      COAMsgBox "Error reading the selected booking's status.", vbOKOnly + vbExclamation, app.ProductName
    End If
  
    rsInfo.Close
    Set rsInfo = Nothing
  End If
  
  If fOK Then
    fOK = (UCase(Left(sSelectedBookingStatus, 1)) = "B") Or _
      (UCase(Left(sSelectedBookingStatus, 1)) = "P")
    If Not fOK Then
      ' JPD20030206 Fault 5013
      COAMsgBox "Training bookings can only be transferred if they have 'Booked'" & _
        IIf(gfTrainBookStatus_P, " or 'Provisional'", "") & " status.", vbOKOnly + vbInformation, app.ProductName
    End If
  End If
  
  If fOK Then
    Set frmCourseSelection = New frmTransferBooking
    
    With frmCourseSelection
      If .Initialise(lngSelectedBookingID) Then
        .Show vbModal

        If Not .Cancelled Then
          ' Refresh the grid.
          mrsFindRecords.Requery
          mfrmParent.Requery False
          Requery False
          
          ' JPD20021025 Fault 4660
          LocateRecordID lngSelectedBookingID
        End If
      End If
    End With
    Unload frmCourseSelection
    Set frmCourseSelection = Nothing
  End If

End Sub


Private Sub cmbOrders_GotFocus()
  miOldOrderIndex = cmbOrders.ListIndex
  
End Sub

Private Sub Form_Activate()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmFind2.Form_Activate()"
  
  ' JPD20020924 Fault 4414
  Screen.MousePointer = vbHourglass

  ' JPD20021209 Fault 4863
  If Not mfrmParent.SaveAscendants Then
    Exit Sub
  End If

  'ssOleDBGridFindColumns.SetFocus
  'ssOleDBGridFindColumns.DoClick
  
  If mfFindForm Then
    If mfrmParent.SaveChanges(False, False) Then
      DoEvents
      If Not Database.Validation Then
        mfrmParent.SetFocus
        GoTo TidyUpAndExit
      End If
    End If
    
    DoEvents
    If mfrmParent.Cancelled Then
      mfrmParent.SetFocus
      GoTo TidyUpAndExit
    End If
  End If

  ' JDM - 26/10/01 - Fault 2449/2933 - Stop messing up when closing/minimising form
  'Highlight the current row in the grid
  If Not WindowState = vbMinimized Then
    If Not mbIsUnloading Then
      
      ' Show the find records
      ssOleDBGridFindColumns.Visible = True
      
      'MH2000305 The fix for fault 3248 caused a further problem (fault 3589).
      'Users want the grid to have focus so that they are starting typing as
      'soon as the find window opens to locate the correct record...
      '
      ''JDM - 05/12/01 - Fault 3248 - Crashes on faster PCs (never development though)
      ''ssOleDBGridFindColumns.SetFocus
      ''SendKeys vbKeySpace, 0
      GridSetFocus
      
      ' JPD20020920 Fault 4414
      'TM20020403 Fault 3364
      'TM20020718 Fault 4167 - can only set the current record if one exists.
      'If mrsFindRecords.RecordCount > 0 Then
      '  SetCurrentRecord
      'End If
    
      FormatControls
    End If

    ' JPD20030203 Fault 4995
    With ssOleDBGridFindColumns
      If (.SelBookmarks.Count = 0) And (.Rows > 0) Then
        .MoveFirst
        .SelBookmarks.Add .Bookmark
      End If
    End With
    
    frmMain.RefreshMainForm Me
  End If
  
'  If ssOleDBGridFindColumns.SelBookmarks.Count = 0 Then
'    With ssOleDBGridFindColumns
'      If .Rows > 0 Then
'        .MoveFirst
'        .SelBookmarks.Add .Bookmark
'      End If
'    End With
'  End If

TidyUpAndExit:
  ' JPD20020924 Fault 4414
  Screen.MousePointer = vbDefault
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub Form_Deactivate()

  DebugOutput "frmFind2", "Form_Deactivate"

  DoEvents
  
  'TM20020612 Fault 2302 - disable the toolbar controls.
  EnableActiveBar Me.ActiveBar1, False
  
End Sub



Private Sub Form_Initialize()
  mbIsLoading = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  '# RH 26/08/99. To pass shortcut keys thru to the activebar control
  Dim bHandled As Boolean

  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
    
    Case vbKeyF3
      frmMain.LogOff
    
    Case vbKeyN
      If Shift = vbCtrlMask Then
        AddNew
      End If
    
    Case vbKeyE
      If Shift = vbCtrlMask Then
        EditRecord
      End If
    
    Case Else
      bHandled = frmMain.abMain.OnKeyDown(KeyCode, Shift)
      If bHandled Then
        KeyCode = 0
        Shift = 0
      End If
  
  End Select
End Sub

Private Sub Form_Load()
   
  'Begin subclassing
  Hook Me.hWnd, dblFINDFORM_MINWIDTH, dblFINDFORM_MINWIDTH
  
  ' Set the user defined activebar
  OrganiseToolbarControls ActiveBar1
  'NHRD15092006 Fault 11493
  ActiveBar1.Bands(0).Flags = 1 + 256 + 512
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim frmTemp As Form
  
  DebugOutput "frmFind2", "Form_QueryUnload"
  
  ' JPD20030306 Fault 5118
  If mfBusy Then
    Cancel = True
    Exit Sub
  End If
  
  DoEvents
  
  'MH20030908 Fault 6885 Check that there is a parent!
  If Not (mfrmParent Is Nothing) Then
    'JPD 20030820 Fault 3048
    For Each frmTemp In Forms
      With frmTemp
        If .Name = "frmRecEdit4" Then
          If (.FormID = mfrmParent.ParentFormID) And (.Changed) Then
            Cancel = True
            Exit Sub
          End If
        End If
      End With
    Next frmTemp
  End If
  Set frmTemp = Nothing
  
  If UnloadMode = vbFormControlMenu Then
    mfCancelled = True
  End If
  
  Screen.MousePointer = vbDefault
  mbIsUnloading = True

End Sub


Private Sub Form_Resize()
  
  ' Resize the form's controls as the form is itself resized.
  Dim lngScaleHeight As Long
  Dim blnShowSummary As Boolean

  ' Performance - No need to run if form isn't visible. Is there? I guess if you're reading this there probably is! :-(
  If Not Me.Visible Then
    Exit Sub
  End If

  'MH20061213 Fault 11785
  mintResizeCount = mintResizeCount + 1
  
  'MH20010824 Fault 2075
  blnShowSummary = (Not mfFindForm Or mfrmParent.ParentTableID > 0)

  ' Do nothing if the form is minimised.
  ' NB. It cannot be maximised.
  If Me.WindowState = vbNormal Then
  
    ' Size the Order frame and the controls therein.
    fraOrders.Width = Me.ScaleWidth - (dblCOORD_XGAP * 2)
    cmbOrders.Width = fraOrders.Width - (dblCOORD_XGAP * 2)
    
    ' Size the Summary frame.
    If blnShowSummary Then

      If Me.Width >= mlngMinFormWidth Then
  
        fraSummary.Width = fraOrders.Width
        
        If blnSummaryDetailsLoaded Then
          ' Reformat all controls to fit new form size.
          ReformatHistoryElements
        End If
        
      End If
      
    End If
        
    ' Size the Find grid.
    With ssOleDBGridFindColumns

      .Width = fraOrders.Width
      lngScaleHeight = Me.ScaleHeight
      
      If Not blnShowSummary Then
        .Height = Me.ScaleHeight - .Top - (dblCOORD_YGAP * 2.5)
      Else
      
        ' TM - Fix for 2544, checks the size to be assigned is greater than 0.
        lngScaleHeight = Me.ScaleHeight - .Top - dblCOORD_YGAP - (IIf(fraSummary.Visible, fraSummary.Height + (dblCOORD_YGAP * 2.5), dblCOORD_YGAP * 1.5))
        If lngScaleHeight > 0 Then
          .Height = lngScaleHeight
        End If
        
        fraSummary.Top = .Top + .Height + dblCOORD_YGAP
      End If
      
      ' Resize the find columns
      ResizeFindColumns
    
    End With
  End If
  
  'JPD 20041109 Fault 9387
  'Refresh form to sort out minimise maximise thing - Fault 8257
  'frmMain.RefreshMainForm mfrmParent
  frmMain.RefreshMainForm Screen.ActiveForm
  Me.Refresh
  
  'MH20061213 Fault 11785
  mintResizeCount = mintResizeCount - 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  DebugOutput "frmFind2", "Form_Unload"
  
  ' Dont do the unload event if form has been loaded from configuration screen
  If mbIsLoading Then
    Exit Sub
  End If
  
  DebugOutput "frmFind2", "Form_Unload 1"
  
  If Not mfrmParent Is Nothing Then
    If mfFindForm Then
      mfrmParent.ReleaseFindWindow
    End If
  End If
  
  DebugOutput "frmFind2", "Form_Unload 2"
  
  ' RH 16/05/00 - Save the find window coordinates
  If Me.WindowState = vbNormal Then
    SavePCSetting "FindWindowCoOrdinates\" & gsDatabaseName & "\" & mfrmParent.ScreenID, "Top", Me.Top
    SavePCSetting "FindWindowCoOrdinates\" & gsDatabaseName & "\" & mfrmParent.ScreenID, "Left", Me.Left
    SavePCSetting "FindWindowCoOrdinates\" & gsDatabaseName & "\" & mfrmParent.ScreenID, "Width", Me.Width
    SavePCSetting "FindWindowCoOrdinates\" & gsDatabaseName & "\" & mfrmParent.ScreenID, "Height", Me.Height
  End If
  
  DebugOutput "frmFind2", "Form_Unload 3"
  
  ' Check to see if we should discard the history window as well
  If Not mfrmParent Is Nothing Then
    If (mfrmParent.ScreenType = screenHistoryTable) Or _
      (mfrmParent.ScreenType = screenHistoryView) Then
  
  DebugOutput "frmFind2", "Form_Unload 4"
      
      ' RH 07/03/01
      ' If find window has bn cancelled and the recedit is visible, then dont
      ' unload the find window, just make it invisible, and dont unload recedit
      If mfCancelled And mfrmParent.Visible = True Then
        Me.Visible = False
        If mfrmParent.Visible And mfrmParent.Enabled Then
  DebugOutput "frmFind2", "Form_Unload 5"
          mfrmParent.SetFocus
        End If
        Cancel = True
  DebugOutput "frmFind2", "Form_Unload 6"
        Exit Sub
      ElseIf mfCancelled Then
  DebugOutput "frmFind2", "Form_Unload 7"
        mfrmParent.ParentUnload = True
        Unload mfrmParent
      End If
  
      'mfrmParent.ParentUnload = True
      'Unload mfrmParent
      
    Else
    
      'Its a Primary/Lookup/Quick Access
  DebugOutput "frmFind2", "Form_Unload 8"
      
      If mfrmParent.Visible = False And mfCancelled Then
        'Recedit is invisible and user cancelled the findform so unload recedit too
  DebugOutput "frmFind2", "Form_Unload 9"
        Unload mfrmParent
      Else
        'JPD 20031009 Fault 7080
        If Not mfrmParent.Recordset Is Nothing Then
          If (mfrmParent.Recordset.State <> adStateClosed) Then
            'If Not mfrmParent.Visible Then
              'Recedit is invisible and user selects somebody so make recedit visible and give focus
    DebugOutput "frmFind2", "Form_Unload 10"
              mfrmParent.Visible = True
              mfrmParent.Enabled = True
              frmMain.Enabled = True
              If mfrmParent.Visible And mfrmParent.Enabled Then
    DebugOutput "frmFind2", "Form_Unload 11"
                mfrmParent.SetFocus
              End If
              frmMain.RefreshMainForm mfrmParent
    DebugOutput "frmFind2", "Form_Unload 12"
              Exit Sub
            'End If
          End If
        End If
      End If
    End If
  End If
    
  DebugOutput "frmFind2", "Form_Unload 13"
  
  'Stop subclassing.
  Unhook Me.hWnd
      
  If Not frmMain Is Nothing Then
  DebugOutput "frmFind2", "Form_Unload 14"
    frmMain.RefreshMainForm Me, True
  End If

  DebugOutput "frmFind2", "Form_Unload 15"
  
End Sub

Private Sub ssOleDBGridFindColumns_Click()
  ' JPD20021126 Fault 4800 - Commented out. Not required ?
  '  frmMain.RefreshRecordMenu Me
  '
  '  'NHRD 08032002 Fault 3366
  '  ssOleDBGridFindColumns.SelBookmarks.Add ssOleDBGridFindColumns.Bookmark
  '
  '  If Not IsNull(ssOleDBGridFindColumns.Bookmark) Then
  '    CurrentBookMark = ssOleDBGridFindColumns.Bookmark
  '  End If
  
  ' RH 13/10/00 - BUG 1121 - Only resize the window if its not min/max.
  'StatusBar1.Panels(1).Text = ssOleDBGridFindColumns.Rows & " Record" & IIf(ssOleDBGridFindColumns.Rows = 1, "", "s") & _
        IIf(ssOleDBGridFindColumns.SelBookmarks.Count > 1, " - " & ssOleDBGridFindColumns.SelBookmarks.Count & " Selected", "")
  UpdateStatusBar
  
End Sub

Private Sub UpdateStatusBar()
  'StatusBar1.Panels(1).Text = ssOleDBGridFindColumns.Rows & " Record" & IIf(ssOleDBGridFindColumns.Rows = 1, "", "s") & _
        IIf(ssOleDBGridFindColumns.SelBookmarks.Count > 1, " - " & ssOleDBGridFindColumns.SelBookmarks.Count & " Selected", "")
  StatusBar1.Panels(1).Text = ssOleDBGridFindColumns.Rows & " Record" & IIf(ssOleDBGridFindColumns.Rows = 1, "", "s") & _
        IIf(ssOleDBGridFindColumns.SelBookmarks.Count > 0, " - " & ssOleDBGridFindColumns.SelBookmarks.Count & " Selected", "")
  
End Sub




Private Sub ssOleDBGridFindColumns_PrintError(ByVal PrintError As Long, Response As Integer)
    
  'TM20020930 Fault 4461 - if the user cancelled the print standard dialog box then set
  'print cancelled flag.
  mblnPrintCancelled = False
  If PrintError = 30457 Then 'User cancelled print
    mblnPrintCancelled = True
  End If

  'Set to 0 to prevent a default error message from being displayed
  Response = 0

End Sub

Public Property Get IsLoading() As Boolean
  IsLoading = mbIsLoading
End Property

Private Sub ssOleDBGridFindColumns_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

  Dim dblSizeAmendment As Double

  With ssOleDBGridFindColumns

    .Redraw = False

    'Find last visible column
    If .Columns(ColIndex + 1).Visible Then
      dblSizeAmendment = .Columns(ColIndex).Width - .ResizeWidth
      .Columns(ColIndex + 1).Width = .Columns(ColIndex + 1).Width + dblSizeAmendment
      ' SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(ColIndex + 1).Name, .Columns(ColIndex + 1).Width
      SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(ColIndex + 1).Caption, .Columns(ColIndex + 1).Width
    End If

    'Save the resized column width
    ' SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(ColIndex).Name, .ResizeWidth
    SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(ColIndex).Caption, .ResizeWidth
    
    .Redraw = True

  End With

End Sub

Private Sub ssOleDBGridFindColumns_DblClick()
  Dim lngSelectedRecord As Long
  
  'MH20061213 Fault 11785
  If mintResizeCount > 0 Then
    Exit Sub
  End If
  
  ' JPD20021126 Fault 4800
  If Not Screen.ActiveForm Is Me Then Exit Sub
  
  ' JPD20020926 Fault 4414
  If mbIsLoading Then Exit Sub

  If ssOleDBGridFindColumns.SelBookmarks.Count > 0 Then
    
    lngSelectedRecord = SelectedRecordID
    
    ' RH 09/10/00 - Workaround for when the user doubleclicks on the empty row which
    '               sometimes appears when the user manouvers around the grid using the keyboard.
    '               SELECTEDRECORDID will return 0 if line is blank, and if thats the case, we
    '               exit this sub
    If lngSelectedRecord = 0 Then Exit Sub

    If mfrmParent.SaveChanges(False) Then
      
      If Not Database.Validation Then
        Exit Sub
      End If
    
      If mfFindForm Then
        ' Find form.
        With mfrmParent
          ' JPD20021126 Fault 4804 - Only do the Requery if the record is not located.
          ' JPD20021113 Fault 4749
          '.Requery False
          .LocateRecord lngSelectedRecord
          .UpdateControls
          
          If .RecordID <> lngSelectedRecord Then
            .Requery False
            .LocateRecord lngSelectedRecord
            .UpdateControls
          End If
          
          If .RecordID <> lngSelectedRecord Then
            COAMsgBox "The selected record has been deleted by another user.", vbExclamation, app.ProductName
            .Requery False
          End If
  
          .UpdateChildren
        End With
    
        mfCancelled = False
        
        Unload Me
        
        ' JPD20021126 Fault 4803
        If Not Screen.ActiveForm Is Nothing Then
          frmMain.RefreshMainForm Screen.ActiveForm
        End If
      Else
        ' History summary form.
        Me.Visible = False
            
        With mfrmParent
          .LocateRecord lngSelectedRecord
          .UpdateControls

          ' JPD20021113 Fault 4749
          If .RecordID <> lngSelectedRecord Then
            .Requery False
            .LocateRecord lngSelectedRecord
            .UpdateControls
          End If
          
          If .RecordID <> lngSelectedRecord Then
            COAMsgBox "The selected record has been deleted by another user.", vbExclamation, app.ProductName
            .Requery False
          End If
          
          .UpdateChildren
  
          .Visible = True
          '#RH 8/11 - bug fix preventing form locking
          .Enabled = True
          .SetFocus
        
          ' RH 18/09/00 - BUG 826 - Ensure control with first tabindex has the focus
          '                         when we are re-showing the history recedit screen
          '                         after a user has finished with the find window
          On Error Resume Next

          Dim objControl As Control
          For Each objControl In .Controls
            If Not TypeOf objControl Is ActiveBar And Not TypeOf objControl Is Menu Then
              If objControl.Visible Then
                If objControl.TabIndex = 0 Then
                  objControl.SetFocus
                  Exit For
                End If
              End If
            End If
          Next objControl
        
        End With
      End If
    End If
  End If

End Sub


Private Function SelectedRecordID() As Long
  ' Return the ID of the selected reocrd in the grid.
  Dim iIndex As Integer
  Dim iIDColumnIndex As Integer
  
  SelectedRecordID = 0
  
  If ssOleDBGridFindColumns.SelBookmarks.Count > 0 Then
    ' Find the index of the ID column.
    iIDColumnIndex = 0
    For iIndex = 0 To (ssOleDBGridFindColumns.Cols - 1)
      If UCase(ssOleDBGridFindColumns.Columns(iIndex).Name) = "ID" Then
        iIDColumnIndex = iIndex
        Exit For
      End If
    Next iIndex
    
    If ssOleDBGridFindColumns.Columns(iIDColumnIndex).Value <> "" Then
      SelectedRecordID = ssOleDBGridFindColumns.Columns(iIDColumnIndex).Value
    Else
      SelectedRecordID = 0
    End If
    
  End If
  
End Function

Private Sub ssOleDBGridFindColumns_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    If ActiveBar1.Bands(0).Tools("DeleteFind").Enabled Then
      DeleteRecord
    End If
  End If
End Sub

Private Sub ssOleDBGridFindColumns_KeyPress(KeyAscii As Integer)
  Dim lngThistime As Long
  Static sFind As String
  Static lngLastTime As Long
  
  Select Case KeyAscii
    Case vbKeyReturn
      ssOleDBGridFindColumns_DblClick
    
    ' Otherwise find the record
    Case Else
      ' Only search for alphanumeric characters.
      If (KeyAscii >= 32) And (KeyAscii <= 255) Then
        lngThistime = GetTickCount
        If lngLastTime + 1500 < lngThistime Then
          sFind = Chr(KeyAscii)
        Else
          sFind = sFind & Chr(KeyAscii)
        End If
        lngLastTime = lngThistime
        LocateRecord sFind
      End If
  End Select

End Sub

Private Sub LocateRecord(psSearchString As String)
  Dim fFound As Boolean
  Dim fUseBinarySearch As Boolean
  Dim iComparisonResult As Integer
  Dim lngLoop As Long
  Dim lngUpper As Long
  Dim lngLower As Long
  Dim lngJump As Long
  Dim varFoundBookmark As Variant
  Dim varOriginalBookmark As Variant
  
  
  'Avoid crash when no rows !!    MH20000713
  If ssOleDBGridFindColumns.Rows = 0 Then
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  
  fUseBinarySearch = mfFirstColumnsMatch
  
  If fUseBinarySearch Then
    If (miFirstColumnDataType <> sqlVarChar) And _
     (miFirstColumnDataType <> sqlVarBinary) And _
     (miFirstColumnDataType <> sqlNumeric) And _
     (miFirstColumnDataType <> sqlInteger) Then
    
      fUseBinarySearch = False
    End If
  End If
  
  ' Search the grid for the required string.
  fFound = False
  
  lngLower = 1
  lngUpper = RecordCount
  
  With ssOleDBGridFindColumns
    .Redraw = False
    
    varOriginalBookmark = .Bookmark
    
    If fUseBinarySearch Then
      ' Binary search the grid for the required string.
      Do
        Select Case miFirstColumnDataType
          Case sqlVarChar, sqlVarBinary
            ' JPD String comparison changed from using VB's strComp function to
            ' using our own DictionaryCompareStrings function. VB's strComp
            ' function does not use the same order as that used when SQL orders
            ' by a character column. The DictionaryCompareStrings does.
            'iComparisonResult = StrComp(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString), vbTextCompare)
            iComparisonResult = datGeneral.DictionaryCompareStrings(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString))
          
          Case sqlNumeric, sqlInteger
            If Val(ssOleDBGridFindColumns.Columns(0).Text) = Val(psSearchString) Then
              iComparisonResult = 0
            ElseIf Val(ssOleDBGridFindColumns.Columns(0).Text) < Val(psSearchString) Then
              iComparisonResult = -1
            Else
              iComparisonResult = 1
            End If
        End Select
        
        If Not mfFirstColumnAscending Then
          iComparisonResult = iComparisonResult * -1
        End If
        
        Select Case iComparisonResult
          Case 0    ' String found.
            fFound = True
            varFoundBookmark = .Bookmark
            lngUpper = .Bookmark - 1
            lngJump = -((.Bookmark - lngLower) \ 2) - 1
            If lngLower > lngUpper Then Exit Do
  
          Case -1   ' Current record is before the required record.
            lngLower = .Bookmark + 1
            lngJump = ((lngUpper - .Bookmark) \ 2)
            If lngLower > lngUpper Then Exit Do
                   
          Case 1    ' Current record is after the required record.
            lngUpper = .Bookmark - 1
            lngJump = -((.Bookmark - lngLower) \ 2) - 1
            If lngLower > lngUpper Then Exit Do
        End Select
        
        If lngLower = lngUpper Then
          lngJump = lngUpper - .Bookmark
        End If
        
        ' Move to the middle record of the recmaining records to search.
        .MoveRecords lngJump
      Loop
  
      If fFound Then
        .Bookmark = varFoundBookmark
      Else
        .MoveRecords varOriginalBookmark - .Bookmark
      End If
    Else
      ' Sequential search the grid for the required string.
      .MoveFirst
      For lngLoop = lngLower To lngUpper
        ' JPD String comparison changed from using VB's strComp function to
        ' using our own DictionaryCompareStrings function. VB's strComp
        ' function does not use the same order as that used when SQL orders
        ' by a character column. The DictionaryCompareStrings does.
        'If StrComp(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString), vbTextCompare) = 0 Then
        If datGeneral.DictionaryCompareStrings(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString)) = 0 Then
          Exit For
        End If
        
        If lngLoop < lngUpper Then
          .MoveNext
        Else
          .Bookmark = varOriginalBookmark
        End If
      Next lngLoop
    End If
    
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark
  
    .Redraw = True
  End With
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub ssOleDBGridFindColumns_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B_OLEDB.ssPrintInfo)
    'NHRD 16042002 Fault 3462 Added this sub to populate the ssPrintInfo object with data.
    'Define a page header that includes Table and View Name
    
    If mvarScreenType = screenHistorySummary Then
        'Put the History owner's Forename and Surname into the header as the view instead.
        'ssPrintInfo.PageHeader = "TABLE: " & mobjTableView.TableName & "     VIEW: " & Me.Caption
        ssPrintInfo.PageHeader = mfrmParent.FindPrintHeader & _
          IIf(Filtered, " (Filtered)", "")
    Else
        If mobjTableView.TableName = msCurrentTableViewName Then
            'Omit the VIEW part of the header
            ssPrintInfo.PageHeader = mobjTableView.TableName & _
              IIf(Filtered, " (Filtered)", "")
        Else
            'Include Table name and View name
             ssPrintInfo.PageHeader = mobjTableView.TableName & " (" & msCurrentTableViewName & " view)" & _
              IIf(Filtered, " (Filtered)", "")
        End If
    End If
        
    'Define a page footer that specifies when the grid was printed.
    'vbTAb will centre the text, two vbTab's will right justify
    'More info in Data Widgets 3.0 Help.
     ssPrintInfo.PageFooter = "Printed on <date> at <Time> by " + gsUserName + "         Page <page number> "

    'Specify that we want each row's height to expand so that all data is displayed,
    'but up to a maximum of 10 lines.
    ssPrintInfo.RowAutoSize = True      'So rows are expanded in height as necessary
    ssPrintInfo.MaxLinesPerRow = 10     'but up to a maximum of 10 lines.

    'Print column and group headers at the top of each page.
    '(Use ssTopOfReport if you want the headers to appear on the first page.)
    ssPrintInfo.PrintHeaders = ssTopOfPage
End Sub

Private Sub ssOleDBGridFindColumns_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  With ssOleDBGridFindColumns
    If .Rows > 0 And .SelBookmarks.Count = 0 Then
      .SelBookmarks.Add .Bookmark
    End If
  End With
  frmMain.RefreshRecordMenu Me
End Sub


Private Sub ssOleDBGridFindColumns_Scroll(Cancel As Integer)
  'TM20020319 Fault 3135 - Mouse scroll working to fast.
  'Not sure why this works but probably slows things down a little,
  'and gives the system chance to catch up.
  DoEvents
End Sub

Private Sub ssOleDBGridFindColumns_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrsFindRecords.MoveLast
    Else
      mrsFindRecords.MoveFirst
    End If
  Else
    mrsFindRecords.Bookmark = StartLocation
  End If
  
  'JPD 20040803 Fault 9013
  If StartLocation + NumberOfRowsToMove <= 0 Then
    NumberOfRowsToMove = 0
  End If

  mrsFindRecords.Move NumberOfRowsToMove
  NewLocation = mrsFindRecords.Bookmark
  
End Sub


Private Sub ssOleDBGridFindColumns_UnboundReadData(ByVal RowBuf As SSDataWidgets_B_OLEDB.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim strFormat As String
  
  iRowsRead = 0
  
  'MH20001026 Fault 780
  'Setting focus to the grid stops the order combo box changing
  'when using the fancy (tarty) mouse wheel.  This stops the
  'invalid bookmark error
  With ssOleDBGridFindColumns
    If .Visible And .Enabled And Not mbIsLoading Then
      'MH20040820 Fault 9084
      '.SetFocus
      GridSetFocus
    End If
  End With
    
  
  ' Do nothing if we a re just formatting the grid,
  ' ot if there a re no records to display.
  If (mfFormattingGrid) Or (RecordCount = 0) Then Exit Sub
  
  'If IsNull(StartLocation) Or (StartLocation = 0) Then
  If IsNull(StartLocation) Then
    If ReadPriorRows Then
      If Not mrsFindRecords.EOF Then
        mrsFindRecords.MoveLast
      End If
    Else
      If Not mrsFindRecords.BOF Then
        mrsFindRecords.MoveFirst
      End If
    End If
  Else
    mrsFindRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsFindRecords.MovePrevious
    Else
      mrsFindRecords.MoveNext
    End If
  End If

  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsFindRecords.BOF Or mrsFindRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsFindRecords.Fields.Count - 1)
        
          Select Case mrsFindRecords.Fields(iFieldIndex).Type
            Case adDBTimeStamp
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsFindRecords(iFieldIndex), DateFormat)
            
            Case adNumeric
            
              ' Are thousand separators used
              strFormat = "0"
              If mavFindColumns(3, iFieldIndex) Then
                strFormat = "#,0"
              End If
              
              
              ' NPG20081128 Fault 13248
              ' Is BlankIfZero option set
              If mavFindColumns(4, iFieldIndex) Then
                strFormat = "#"
              End If
                           
              If (mavFindColumns(2, iFieldIndex) > 0) And (mrsFindRecords(iFieldIndex) > 0) Then
                  strFormat = strFormat & "." & String(mavFindColumns(2, iFieldIndex), "0")
              'TM20090706 Fault HRPRO-48
              ElseIf (mavFindColumns(2, iFieldIndex) > 0) And (mrsFindRecords(iFieldIndex) < 0) Then
                  strFormat = strFormat & "." & String(mavFindColumns(2, iFieldIndex), "0")
              End If
              
              
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsFindRecords(iFieldIndex), strFormat)
            
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrsFindRecords(iFieldIndex)
          
          End Select
        
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsFindRecords.Bookmark
  
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsFindRecords.Bookmark
    End Select
    
    If ReadPriorRows Then
      mrsFindRecords.MovePrevious
    Else
      mrsFindRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead

End Sub



Public Sub BulkBooking()
  ' Training Booking specific function.
  '
  ' This function is available when the find window is displaying the Training Booking records
  ' for a Course.
  Dim lngSelectedCourseID  As Long
  Dim frmBulkPick As frmBulkBooking
  Dim objCourseTableView As CTablePrivilege

  ' If the Training Booking module is enabled ...
  If gfTrainingBookingEnabled Then
    lngSelectedCourseID = mfrmParent.ParentID

    ' Get the current Course table/view object.
    If mfrmParent.ParentViewID > 0 Then
      Set objCourseTableView = gcoTablePrivileges.FindViewID(mfrmParent.ParentViewID)
    Else
      Set objCourseTableView = gcoTablePrivileges.FindTableID(mfrmParent.ParentTableID)
    End If
    
    Set frmBulkPick = New frmBulkBooking
    With frmBulkPick
      If .Initialise(lngSelectedCourseID, objCourseTableView) Then
        .Show vbModal

        If Not .Cancelled Then
          ' Refresh the grid.
          mrsFindRecords.Requery
          mfrmParent.Requery False
          Requery False
        
          ' JPD20021112 Fault 4729
          With ssOleDBGridFindColumns
            If .Rows > 0 Then
              .MoveFirst
              .SelBookmarks.Add .Bookmark
            End If
          End With
          frmMain.RefreshMainForm Me
        Else
          With ssOleDBGridFindColumns
              If .SelBookmarks.Count = 0 Then
                If .Rows > 0 Then
                  .MoveFirst
                  .SelBookmarks.Add .Bookmark
                End If
              End If
          End With
        End If
      End If
    End With
    Unload frmBulkPick
    Set frmBulkPick = Nothing
  
    Set objCourseTableView = Nothing
  End If
    
End Sub
Public Property Get CanAddFromWaitingList() As Boolean
  CanAddFromWaitingList = mfCanAddFromWaitingList
End Property

Public Property Get CanBookCourse() As Boolean
  CanBookCourse = mfCanBookCourse
End Property

Public Property Get CanBulkBooking() As Boolean
  CanBulkBooking = mfCanBulkBooking
End Property

Public Property Get CanCancelBooking() As Boolean
  CanCancelBooking = mfCanCancelBooking
End Property

Public Property Get CanTransferBooking() As Boolean
  CanTransferBooking = mfCanTransferBooking
End Property

Public Property Let CurrentRecordID(plngData As Long)
  ' Set's the Parents Form ID.
  mlngCurrentRecordID = plngData
  
End Property


Public Property Get CurrentRecordID() As Long
  ' Get's the Parents Form ID.
  CurrentRecordID = mlngCurrentRecordID
  
End Property

Public Property Let CurrentBookMark(plngData As Long)
  ' Records the current bookmark
  mlngCurrentBookMark = plngData
  
End Property


Public Property Get CurrentBookMark() As Long
  ' Records the current bookmark
  CurrentBookMark = mlngCurrentBookMark
  
End Property

Public Sub SetCurrentRecord()

  Dim mvarbkCurrentRecord As Variant
  Dim fFirstRecord As Boolean

  ssOleDBGridFindColumns.SelBookmarks.RemoveAll

  fFirstRecord = False
  
  'Only search for existing records.
  ' JPD20030225 Fault 5079
  'If mfrmParent.RecordID > 0 Then
  If (mfrmParent.RecordID > 0) And (ssOleDBGridFindColumns.Rows > 0) Then
'    If Not Filtered Then
    
    With ssOleDBGridFindColumns
      mvarbkCurrentRecord = .RowBookmark(0)
      
      If .Columns("ID").CellValue(mvarbkCurrentRecord) <> "" Then
        If CLng(.Columns("ID").CellValue(mvarbkCurrentRecord)) = CLng(mfrmParent.RecordID) Then
          .MoveRecords mvarbkCurrentRecord
          .Bookmark = mvarbkCurrentRecord
          .SelBookmarks.Add mvarbkCurrentRecord
          fFirstRecord = True
        End If
      End If
    End With
        
    If Not fFirstRecord Then
      LocateRecordID mfrmParent.RecordID
    End If
    
'      ' Only search if the first column is from this table/view
'      Set rsInfo = datGeneral.GetOrderDefinition(mlngOrderID)
'      rsInfo.MoveFirst
'
'      If (rsInfo.Fields("TableName").Value = mobjTableView.TableName) And _
'        (Not Left(mobjTableView.RealSource, 8) = "ASRSysCV") Then
'
'        ' This puts us near where we need to be, but not exactly as we could have first columns with the same value
'        vData = datGeneral.GetOrderValue(mlngCurrentRecordID, mrsFindRecords.Fields(0).Name, mobjTableView.RealSource)
'        'vData = mrsFindRecords.Fields(0).Value
'        lngRecordNumber = datGeneral.GetRecordOrderNumber(vData, mrsFindRecords.Fields(0).Name, mobjTableView.RealSource, mfFirstColumnAscending)
'
'        ' If record number overshoots, maybe there's a filter applied
'        If lngRecordNumber > mrsFindRecords.RecordCount Then
'          Exit Sub
'        End If
'
'        mrsFindRecords.Move lngRecordNumber - 1, 1
'      Else
'        mrsFindRecords.MoveFirst
'      End If
'
'      mrsFindRecords.MoveFirst
'    ' Position on current record
'      Do Until mrsFindRecords.EOF Or mrsFindRecords.Fields("ID") = mfrmParent.RecordID
'        mrsFindRecords.MoveNext
'      Loop
'
'      ' Find the bookmark for the desired record
'      mvarbkCurrentRecord = mrsFindRecords.Bookmark
'      ssOleDBGridFindColumns.MoveRecords mvarbkCurrentRecord
'      ssOleDBGridFindColumns.Bookmark = mvarbkCurrentRecord
'
'      'Highlight the current selection
'      'ssOleDBGridFindColumns.FirstRow = ssOleDBGridFindColumns.Bookmark
'
'      'MH20020430 Fault 3809 RemoveAll bookmarks before setting current record...
'      ssOleDBGridFindColumns.SelBookmarks.RemoveAll
'      ssOleDBGridFindColumns.SelBookmarks.Add ssOleDBGridFindColumns.Bookmark
'  End If
'    Else
'      'TM20020403 Fault 3364 - if the records are a filtered set of records then search
'      'through the rows in the grid until the record id is found.
'      With ssOleDBGridFindColumns
'        For i = 0 To .Rows - 1 Step 1
'          mvarbkCurrentRecord = .RowBookmark(i)
'
'          If CLng(.Columns("ID").CellValue(mvarbkCurrentRecord)) = CLng(mfrmParent.RecordID) Then
'            .MoveRecords mvarbkCurrentRecord
'            .Bookmark = mvarbkCurrentRecord
'            .SelBookmarks.Add mvarbkCurrentRecord
'          End If
'        Next i
'      End With
'    End If
'
'    CurrentBookMark = ssOleDBGridFindColumns.Bookmark
  Else
    ' JPD20021127 Fault 4218
    With ssOleDBGridFindColumns
      If .Rows > 0 Then
        .SelBookmarks.RemoveAll
        .MoveFirst
        .SelBookmarks.Add .Bookmark
      End If
    End With
  End If
 
  UpdateStatusBar

End Sub

Public Sub PrintGrid()

  Dim pstrError As String
  Dim intResponse As Integer

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsPrintGrid.PrintGrid()"
  
  'NHRD16072004 Fault 8740
  If Printers.Count = 0 Then
    pstrError = "Unable to print as no printers are installed."
    COAMsgBox pstrError, vbExclamation + vbOKOnly, Application.Name
    GoTo TidyUpAndExit
  End If
  
  With ssOleDBGridFindColumns
    If .Rows < 1 Then
      COAMsgBox "There is no data in the current view to print", vbInformation + vbOKOnly, app.title
      Exit Sub
    End If
    .Redraw = False
    Screen.MousePointer = vbHourglass

    .HeadFont.Underline = True

    .PageHeaderFont.Name = "Tahoma"
    .PageHeaderFont.Size = 8
    .PageHeaderFont.Bold = True
    .PageHeaderFont.Underline = True

    .PageFooterFont.Name = "Tahoma"
    .PageFooterFont.Size = 8
    .PageFooterFont.Bold = False
    .PageFooterFont.Underline = False
    
    'intResponse = vbNo
    'If .SelBookmarks.Count > 1 Then
    '  intResponse = MsgBox("Would you just like to print the selected rows?", vbQuestion + vbYesNoCancel)
    'End If
    
    'Select Case intResponse
    'Case vbNo
      .PrintData ssPrintAllRows + ssPrintFieldOrder, False, gbPrinterPrompt
    'Case vbYes
    '  .PrintData ssPrintSelectedRows + ssPrintFieldOrder, False, gbPrinterPrompt
    'Case Else
    '  Exit Sub
    'End Select

    .HeadFont.Underline = False

    .Redraw = True
    Screen.MousePointer = vbDefault

  End With
  
  'TM20011219 Fault 3154 - Show print confirm message if required.
  'NB. Would have used the clsPrintGrid class but the grid on this form is of a
  'different type.
  ' Display a printing complete prompt
  Dim strMBText As String
  Dim msb As frmMessageBox
  Dim iShowMeAgain As Integer
  
  If gbPrinterConfirm And Not (mblnPrintCancelled) Then

'TM20020924 Fault 4356 - Ideally would have used the clsPrintGrid class but this requires a different grid type.
' For the short term have removed the printer device name from the Print Confirm message box.
'    strMBText = "Printing complete." _
'      & vbCrLf & vbCrLf & "(" & Printer.DeviceName & ")"
    strMBText = "Printing complete."
  
    iShowMeAgain = IIf(gbPrinterConfirm, 1, 0)
    If iShowMeAgain = 1 Then
      Set msb = New frmMessageBox
      'TM20020930 Fault 4462 - the checkbox should not be checked.
      iShowMeAgain = 0
      msb.MessageBox strMBText, vbInformation, app.ProductName, iShowMeAgain, "Don't show me this confirmation again."
      gbPrinterConfirm = IIf(iShowMeAgain = 0, True, False)
      SavePCSetting "Printer", "Confirm", gbPrinterConfirm
      Set msb = Nothing
    End If
  
  End If

  
TidyUpAndExit:
  mblnPrintCancelled = False
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub GridSetFocus()

  'MH 2000305 Fault 3589 - Keep trying to set focus to the grid
  '(limit to a maximum of 5 seconds so as not to hang the system too long!)

  Dim lngStartTime As Long

  On Local Error Resume Next
  
  'JPD 20041112 Fault 9456
  If Me.ActiveControl Is ssOleDBGridFindColumns Then
  'MH20041027 Fault 9348
  'If Me.ActiveControl = ssOleDBGridFindColumns Then
    Exit Sub
  End If
  
  ' JPD20021025 Fault 4659
  If Me.Enabled = False Then Exit Sub

  lngStartTime = Timer
  Do
    DoEvents
    Err.Clear
    ssOleDBGridFindColumns.SetFocus
    
  'JPD 20031024 Fault 7383
  'Loop While Err.Number > 0 And (Timer - lngStartTime) < 5
  Loop While (Err.Number > 0 And (Timer - lngStartTime) < 5) And Me.Visible

End Sub


Public Function GetSelectedIDs() As String

  Dim strSelectedRecords As String
  Dim intCount As Integer
  Dim objTemp As Variant
  
  
  strSelectedRecords = vbNullString
  With ssOleDBGridFindColumns
    objTemp = .Bookmark
    .Redraw = False
    For intCount = 0 To .SelBookmarks.Count - 1
      'strSelectedRecords = strSelectedRecords & _
          IIf(strSelectedRecords <> vbNullString, ",", "") & _
          .Columns("ID").CellValue(.SelBookmarks(intCount))
      If .SelBookmarks(intCount) = 0 Then
        .MoveFirst
      Else
        .Bookmark = .SelBookmarks(intCount)
      End If
      strSelectedRecords = strSelectedRecords & _
          IIf(strSelectedRecords <> vbNullString, ",", "") & _
          .Columns("ID").Text
    Next
    .Bookmark = objTemp
    .Redraw = True
  End With
  
  GetSelectedIDs = strSelectedRecords

End Function


Public Sub UtilityClick(lngUtilType As UtilityType)
  
  Dim objCustomReportRun As clsCustomReportsRUN
  Dim objCalendarReport As clsCalendarReportsRUN
  Dim objGlobalRun As clsGlobalRun
  Dim objMailMergeRun As clsMailMergeRun
  Dim objDataTransferRun As clsDataTransferRun
  
  Dim frmSelection As frmDefSel
  Dim blnExit As Boolean
  Dim blnOK As Boolean
  Dim iFirstRow As Integer
  
  'Dim varBookmark As Variant

  mstrSelectedRecords = GetSelectedIDs
  ' iFirstRow = ssOleDBGridFindColumns.FirstRow
  'varBookmark = ssOleDBGridFindColumns.Bookmark
  
  If mstrSelectedRecords <> vbNullString Then
    If mfrmParent.SaveChanges(False) Then
      If Not Database.Validation Then
        Exit Sub
      End If
  
      mfBusy = True
      frmMain.DisableMenu

      With mfrmParent
        
        Set frmSelection = New frmDefSel
        blnExit = False
        
        With frmSelection
          Do While Not blnExit
            
            .TableComboEnabled = False
            .TableComboVisible = True
            .TableID = mobjTableView.TableID
            .Options = edtSelect
            .EnableRun = True

            If .ShowList(lngUtilType) Then

              .CustomShow vbModal

              Select Case .Action
                Case edtSelect

                  Select Case lngUtilType
                  Case utlCustomReport
                    Set objCustomReportRun = New clsCustomReportsRUN
                    objCustomReportRun.CustomReportID = .SelectedID
                    objCustomReportRun.RunCustomReport mstrSelectedRecords
                    Set objCustomReportRun = Nothing

                  Case utlCalendarReport
                    Set objCalendarReport = New clsCalendarReportsRUN
                    objCalendarReport.CalendarReportID = .SelectedID
                    objCalendarReport.RunCalendarReport mstrSelectedRecords
                    Set objCalendarReport = Nothing

                  Case utlGlobalUpdate
                    Set objGlobalRun = New clsGlobalRun
                    objGlobalRun.RunGlobal .SelectedID, glUpdate, mstrSelectedRecords
                    Set objGlobalRun = Nothing

                    'Need to refresh the window as data may have changed
                    UI.LockWindow Me.hWnd
                    mblnRefreshing = True
                    frmMain.RefreshRecordEditScreens
                    mblnRefreshing = False
                    ReinstateSelectedRows mstrSelectedRecords
                    UI.UnlockWindow

                  Case utlMailMerge
                    Set objMailMergeRun = New clsMailMergeRun
                    objMailMergeRun.ExecuteMailMerge .SelectedID, mstrSelectedRecords
                    Set objMailMergeRun = Nothing

                  Case utlDataTransfer
                    Set objDataTransferRun = New clsDataTransferRun
                    objDataTransferRun.ExecuteDataTransfer .SelectedID, mstrSelectedRecords
                    Set objDataTransferRun = Nothing

                  End Select

                  blnExit = gbCloseDefSelAfterRun
                Case edtCancel
                  blnExit = True  'cancel
              End Select
            
            End If
          
          Loop
        End With
      
      End With

      frmMain.EnableMenu Me
    End If

  End If


  With ssOleDBGridFindColumns
    If .Rows > 0 Then
      If .Enabled And .Visible Then
        .SetFocus
      End If
    Else
      cmbOrdersSetFocus
    End If
  End With

  If Me.Visible Then
    frmMain.RefreshMainForm Me
  End If

  mfBusy = False

End Sub

Public Function ReinstateSelectedRows(pstrSelectedRecords As String)
  Dim arrayBookmarks() As String
  Dim nTotalSelRows As Variant
  Dim intCount As Integer
  Dim iGridLoop As Integer
  
  Dim strIDs() As String
  Dim intIndex As Integer
    
  If pstrSelectedRecords = vbNullString Then Exit Function
  
  'Avoid crash when no rows !!    MH20000713
  If ssOleDBGridFindColumns.Rows = 0 Then
    Exit Function
  End If
  
  With ssOleDBGridFindColumns
  
    .Redraw = False
    .SelBookmarks.RemoveAll
    
    strIDs = Split(pstrSelectedRecords, ",")
    For intIndex = 0 To UBound(strIDs)
      mrsFindRecords.MoveFirst
      mrsFindRecords.Find "ID = " & strIDs(intIndex)
      If Not mrsFindRecords.EOF Then
        .SelBookmarks.Add mrsFindRecords.Bookmark
      End If
    Next
    
    '.MoveFirst
    'pstrSelectedRecords = "," & pstrSelectedRecords & ","
    'For iGridLoop = 1 To RecordCount
    '  nTotalSelRows = .Bookmark
    '  If InStr(pstrSelectedRecords, "," & CStr(.Columns("ID").Value) & ",") > 0 Then
    '    .SelBookmarks.Add nTotalSelRows
    '  End If
    '  .MoveNext
    'Next
  
    If .SelBookmarks.Count > 0 Then
      .Bookmark = .SelBookmarks(0)
    End If
      
    .Redraw = True
  
  End With
  
End Function
Private Function GetRecCount(strSQL As String) As Long

  Dim rsTemp As ADODB.Recordset

  GetRecCount = 0

  Set rsTemp = New ADODB.Recordset
  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    GetRecCount = Val(rsTemp(0).Value)
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing

End Function

' Wrapper for the isnulls
Private Function IsNullCheck(ByRef objValue As ADODB.Field, ByRef objDefault As Variant) As Variant

  If IsNull(objValue) Then
    IsNullCheck = objDefault
  Else
    IsNullCheck = objValue
  End If

End Function

