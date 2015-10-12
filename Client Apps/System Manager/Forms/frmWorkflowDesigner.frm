VERSION 5.0
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "actbar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{714061F3-25A6-4821-B196-7D15DCCDE00E}#1.0#0"; "coasd_selectionbox.ocx"
Object = "{F3C5146D-8FDA-4D29-8E41-0C27C803C808}#1.0#0"; "coawf_beginend.ocx"
Object = "{08EDC6C1-0A62-485F-8917-8D9FB93DB156}#1.0#0"; "coawf_decision.ocx"
Object = "{FA64823C-ABCB-45AC-ADF2-640EA91D7B88}#1.0#0"; "coawf_email.ocx"
Object = "{9833D366-F890-48E4-BB54-43ACC99E8E7C}#1.0#0"; "coawf_junction.ocx"
Object = "{853234F9-0AB0-42A6-8030-F601CCDCEDBB}#1.0#0"; "COAWF_Link.ocx"
Object = "{63212438-5384-4CC0-B836-A2C015CCBF9B}#1.0#0"; "coawf_webform.ocx"
Begin VB.Form frmWorkflowDesigner 
   AutoRedraw      =   -1  'True
   Caption         =   "Workflow Designer"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5053
   Icon            =   "frmWorkflowDesigner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picOffPage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   600
      Picture         =   "frmWorkflowDesigner.frx":000C
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   400
      Left            =   4800
      TabIndex        =   3
      Top             =   4320
      Width           =   4000
      Begin VB.CommandButton cmdValidate 
         Caption         =   "&Validate"
         Height          =   400
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   400
         Left            =   1400
         TabIndex        =   5
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   2760
         TabIndex        =   6
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H80000005&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   8595
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   8655
      Begin VB.PictureBox picDefinition 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   0
         ScaleHeight     =   3075
         ScaleWidth      =   8355
         TabIndex        =   7
         Top             =   0
         Width           =   8355
         Begin SystemMgr.COAWF_StoredData ASRWFStoredData1 
            Height          =   780
            Index           =   0
            Left            =   6240
            TabIndex        =   15
            Top             =   120
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   1376
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Stored Data"
         End
         Begin COAWFDecision.COAWF_Decision ASRWFDecision1 
            Height          =   1230
            Index           =   0
            Left            =   6240
            TabIndex        =   16
            Top             =   1200
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   2170
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Decision"
         End
         Begin COAWFWebForm.COAWF_Webform ASRWFWebform1 
            Height          =   795
            Index           =   0
            Left            =   4440
            TabIndex        =   14
            Top             =   120
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   1402
            Caption         =   "Web Form"
         End
         Begin COAWFJunction.COAWF_Junction ASRWFJunctionElement1 
            Height          =   525
            Index           =   0
            Left            =   3600
            TabIndex        =   13
            Top             =   120
            Visible         =   0   'False
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   926
            Caption         =   "A"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin COAWFEmail.COAWF_Email ASRWFEmail1 
            Height          =   780
            Index           =   0
            Left            =   1800
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   1376
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Email"
         End
         Begin COAWFBeginEnd.COAWF_BeginEnd ASRWFBeginEnd1 
            Height          =   540
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   953
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "BEGIN"
         End
         Begin COAWFLink.COAWF_Link ASRWFLink1 
            Height          =   120
            Index           =   0
            Left            =   240
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   120
            _ExtentX        =   370
            _ExtentY        =   370
         End
         Begin COASDSelectionBox.COASD_SelectionBox asrboxMultiSelection 
            Height          =   570
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Visible         =   0   'False
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   1005
            BorderColor     =   -2147483640
            BorderStyle     =   3
         End
      End
   End
   Begin MSComCtl2.FlatScrollBar scrollHorizontal 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1572865
   End
   Begin MSComCtl2.FlatScrollBar scrollVertical 
      Height          =   3375
      Left            =   8880
      TabIndex        =   2
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   5953
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   0
      Orientation     =   1572864
   End
   Begin ActiveBarLibraryCtl.ActiveBar abMenu 
      Left            =   120
      Top             =   120
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
      Bands           =   "frmWorkflowDesigner.frx":04C2
   End
End
Attribute VB_Name = "frmWorkflowDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
Private mblnLoading As Boolean
Private mblnStarted As Boolean

Private Const MIN_FORM_HEIGHT = 9000
Private Const MIN_FORM_WIDTH = 9000

Private Const SCROLLMAX = 32767
Private Const SMALLSCROLL = 250

Private Const giSTANDARDMOVEMENT = 15

Private mblnDragging As Boolean
Private mblnMouseDownFired As Boolean

Private mdblVerticalScrollRatio As Double
Private mdblHorizontalScrollRatio As Double

Private miControlIndex As Integer
Private mlngWorkflowID As Long
Private mfChanged As Boolean
Private mfPerge As Boolean
Private mfAppChanged As Boolean
Private mfNewWorkflow As Boolean
Private msWorkflowName As String
Private msWorkflowDescription As String
Private mlngWorkflowPictureID As Long
Private mfWorkflowEnabled As Boolean
Private mfExitToWorkflow As Boolean
Private mfReadOnly As Boolean
Private miInitiationType As WorkflowInitiationTypes
Private miRecSelType As WorkflowRecordSelectorTypes
Private miUsageChoice As Integer
Private miWFUsageSelection As WorkflowFindUsageOption
Private msWFUsageElement As String
Private msWFUsageItem As String
Private msExternalInitiationQueryString As String

Private mlngBaseTableID As Long
Private mcolwfElements As Collection
Private mcolwfSelectedElements As Collection
Private mcolwfSelectedLinks As Collection
Private malngIndexDirectory() As Long

Private mlngXDrop As Long
Private mlngYDrop As Long

Private mfMultiSelecting As Boolean
Private mlngMultiSelectionXStart As Long
Private mlngMultiSelectionYStart As Long

Private miSelectionOrder() As Integer
Private malngElementsDone() As Long

Private mobjPrinter As clsPrintDef
Private mlngBottom As Long

Private Declare Function ClipCursor Lib "user32" (lpRect As Rect) As Long
Private Declare Function ClipCursorByNum Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private startPointSingle As POINTAPI
Private startPointMulti As POINTAPI
Private currPoint As POINTAPI
Private WindowRect As Rect

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const c_DTDefFmt = DT_NOPREFIX 'Or DT_SINGLELINE Or DT_VCENTER

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type Page
  x As Long
  y As Long
End Type

Private msngXOffset As Single
Private msngYOffset As Single
Private mlngXOffset As Long
Private mlngYOffset As Long
Private mlngLastX As Long
Private mlngLastY As Long

Private mlngMarginTop_Twips As Long
Private mlngMarginBottom_Twips As Long
Private mlngMarginLeft_Twips As Long
Private mlngMarginRight_Twips As Long
Private msngTopGap As Single

Private Const TWIPSPERMM = 56.7
Private Const MARGINCORRECTION = 5 * TWIPSPERMM
Private Const iGAPOFFPAGE = 800

Private msOffPageCharacter As String

Private mavOffPageLinks() As Variant
Private mlngRealBottom As Long

Private mactlClipboardControls() As VB.Control

Private miLastActionFlag As UndoActionFlags

Private mactlUndoControls() As VB.Control

Private mavValidationMessages() As Variant
Private mfFixableValidationFailures As Boolean

Private mlngPersonnelTableID As Long
Private maobjOriginalExpressions() As CExpression
Private mfExpressionsChanged As Boolean

Private mbLocked As Boolean

Private Function ArrowSelect(piKeyCode As Integer) As Boolean
  Dim ctlCurrentItem As VB.Control
  Dim ctlNextItem As VB.Control
  Dim ctlTopMostItem As VB.Control
  Dim ctlBottomMostItem As VB.Control
  Dim ctlLeftMostItem As VB.Control
  Dim ctlRightMostItem As VB.Control
  Dim ctlTemp As VB.Control
  Dim fGoodItem As Boolean
  
  If (piKeyCode <> vbKeyLeft) _
    And (piKeyCode <> vbKeyRight) _
    And (piKeyCode <> vbKeyUp) _
    And (piKeyCode <> vbKeyDown) Then
    
    ArrowSelect = False
    Exit Function
  End If
          
  ' Determine the currently selected element/link in the extreme of the direction of the arrow key pressed.
  Set ctlCurrentItem = Nothing
  Set ctlNextItem = Nothing
  Set ctlTopMostItem = Nothing
  Set ctlBottomMostItem = Nothing
  Set ctlLeftMostItem = Nothing
  Set ctlRightMostItem = Nothing

  For Each ctlTemp In Me.Controls
    If IsWorkflowControl(ctlTemp) Then
    
      If (ctlTemp.Visible) Then
        If (ctlTemp.Highlighted) Then
          If ctlCurrentItem Is Nothing Then
            Set ctlCurrentItem = ctlTemp
          Else
            Select Case piKeyCode
              Case vbKeyLeft
                If (ctlCurrentItem.Left > ctlTemp.Left) Then
                  Set ctlCurrentItem = ctlTemp
                ElseIf (ctlCurrentItem.Left = ctlTemp.Left) And (ctlCurrentItem.Top > ctlTemp.Top) Then
                  Set ctlCurrentItem = ctlTemp
                End If

              Case vbKeyRight
                If (ctlCurrentItem.Left < ctlTemp.Left) Then
                  Set ctlCurrentItem = ctlTemp
                ElseIf (ctlCurrentItem.Left = ctlTemp.Left) And (ctlCurrentItem.Top < ctlTemp.Top) Then
                  Set ctlCurrentItem = ctlTemp
                End If

              Case vbKeyUp
                If (ctlCurrentItem.Top > ctlTemp.Top) Then
                  Set ctlCurrentItem = ctlTemp
                ElseIf (ctlCurrentItem.Top = ctlTemp.Top) And (ctlCurrentItem.Left > ctlTemp.Left) Then
                  Set ctlCurrentItem = ctlTemp
                End If
              
              Case vbKeyDown
                If (ctlCurrentItem.Top < ctlTemp.Top) Then
                  Set ctlCurrentItem = ctlTemp
                ElseIf (ctlCurrentItem.Top = ctlTemp.Top) And (ctlCurrentItem.Left < ctlTemp.Left) Then
                  Set ctlCurrentItem = ctlTemp
                End If
            End Select
          End If
        End If
      
        If ctlTopMostItem Is Nothing Then
          Set ctlTopMostItem = ctlTemp
        Else
          If (ctlTopMostItem.Top > ctlTemp.Top) Then
            Set ctlTopMostItem = ctlTemp
          ElseIf (ctlTopMostItem.Top = ctlTemp.Top) And (ctlTopMostItem.Left > ctlTemp.Left) Then
            Set ctlTopMostItem = ctlTemp
          End If
        End If
      
        If ctlBottomMostItem Is Nothing Then
          Set ctlBottomMostItem = ctlTemp
        Else
          If (ctlBottomMostItem.Top < ctlTemp.Top) Then
            Set ctlBottomMostItem = ctlTemp
          ElseIf (ctlBottomMostItem.Top = ctlTemp.Top) And (ctlBottomMostItem.Left < ctlTemp.Left) Then
            Set ctlBottomMostItem = ctlTemp
          End If
        End If

        If ctlLeftMostItem Is Nothing Then
          Set ctlLeftMostItem = ctlTemp
        Else
          If (ctlLeftMostItem.Left > ctlTemp.Left) Then
            Set ctlLeftMostItem = ctlTemp
          ElseIf (ctlLeftMostItem.Left = ctlTemp.Left) And (ctlLeftMostItem.Top > ctlTemp.Top) Then
            Set ctlLeftMostItem = ctlTemp
          End If
        End If
        
        If ctlRightMostItem Is Nothing Then
          Set ctlRightMostItem = ctlTemp
        Else
          If (ctlRightMostItem.Left < ctlTemp.Left) Then
            Set ctlRightMostItem = ctlTemp
          ElseIf (ctlRightMostItem.Left = ctlTemp.Left) And (ctlRightMostItem.Top < ctlTemp.Top) Then
            Set ctlRightMostItem = ctlTemp
          End If
        End If
      End If
    End If
  Next ctlTemp
  Set ctlTemp = Nothing

  If ctlCurrentItem Is Nothing Then
    ' No selected item, so select the topmost item if there is one.
    If Not ctlTopMostItem Is Nothing Then
      Set ctlNextItem = ctlTopMostItem
    End If
  Else
    ' Determine the next item (element/link) to select.
    For Each ctlTemp In Me.Controls
      If IsWorkflowControl(ctlTemp) Then

        If (ctlTemp.Visible) Then
          fGoodItem = False
          
          Select Case piKeyCode
            Case vbKeyLeft
              fGoodItem = (ctlTemp.Left < ctlCurrentItem.Left) _
                Or ((ctlTemp.Left = ctlCurrentItem.Left) And (ctlTemp.Top < ctlCurrentItem.Top))
            Case vbKeyRight
              fGoodItem = (ctlTemp.Left > ctlCurrentItem.Left) _
                Or ((ctlTemp.Left = ctlCurrentItem.Left) And (ctlTemp.Top > ctlCurrentItem.Top))
            Case vbKeyUp
              fGoodItem = (ctlTemp.Top < ctlCurrentItem.Top) _
                Or ((ctlTemp.Top = ctlCurrentItem.Top) And (ctlTemp.Left < ctlCurrentItem.Left))
            Case vbKeyDown
              fGoodItem = (ctlTemp.Top > ctlCurrentItem.Top) _
                Or ((ctlTemp.Top = ctlCurrentItem.Top) And (ctlTemp.Left > ctlCurrentItem.Left))
          End Select
        
          If fGoodItem Then
            If ctlNextItem Is Nothing Then
              Set ctlNextItem = ctlTemp
            Else
              Select Case piKeyCode
                Case vbKeyLeft
                  If (ctlNextItem.Left < ctlTemp.Left) Then
                    Set ctlNextItem = ctlTemp
                  ElseIf (ctlNextItem.Left = ctlTemp.Left) And (ctlNextItem.Top < ctlTemp.Top) Then
                    Set ctlNextItem = ctlTemp
                  End If
  
                Case vbKeyRight
                  If (ctlNextItem.Left > ctlTemp.Left) Then
                    Set ctlNextItem = ctlTemp
                  ElseIf (ctlNextItem.Left = ctlTemp.Left) And (ctlNextItem.Top > ctlTemp.Top) Then
                    Set ctlNextItem = ctlTemp
                  End If
  
                Case vbKeyUp
                  If (ctlNextItem.Top < ctlTemp.Top) Then
                    Set ctlNextItem = ctlTemp
                  ElseIf (ctlNextItem.Top = ctlTemp.Top) And (ctlNextItem.Left < ctlTemp.Left) Then
                    Set ctlNextItem = ctlTemp
                  End If
  
                Case vbKeyDown
                  If (ctlNextItem.Top > ctlTemp.Top) Then
                    Set ctlNextItem = ctlTemp
                  ElseIf (ctlNextItem.Top = ctlTemp.Top) And (ctlNextItem.Left > ctlTemp.Left) Then
                    Set ctlNextItem = ctlTemp
                  End If
              End Select
            End If
          End If
        End If
      End If
    Next ctlTemp
    Set ctlTemp = Nothing
  
    ' No items in the direction we want
    If ctlNextItem Is Nothing Then
      Select Case piKeyCode
        Case vbKeyLeft
          Set ctlNextItem = ctlRightMostItem
        Case vbKeyRight
          Set ctlNextItem = ctlLeftMostItem
        Case vbKeyUp
          Set ctlNextItem = ctlBottomMostItem
        Case vbKeyDown
          Set ctlNextItem = ctlTopMostItem
      End Select
    End If
  End If
  
  If Not ctlNextItem Is Nothing Then
    DeselectAllElements
    
'    ctlNextItem.HighLighted = True
    SelectElement ctlNextItem
    ctlNextItem.ZOrder 0

    'JPD 20061129 Fault 11533 - Ensure the selected element is visible.
    MoveToItem ctlNextItem
    
    If IsWorkflowElement(ctlNextItem) Then
      ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
      miSelectionOrder(UBound(miSelectionOrder)) = ctlNextItem.ControlIndex
    End If
    
    RefreshMenu
  End If
  
  ArrowSelect = True
  
End Function

Private Function ClipboardControlsCount() As Integer
  ' Return a count of the number of controls in the clipboard control.
  On Error GoTo ErrorTrap
  
  ClipboardControlsCount = UBound(mactlClipboardControls)
  Exit Function
  
ErrorTrap:
  ClipboardControlsCount = 0
  
End Function

Private Function AddElement(piElementType As ElementType) As VB.Control
  ' Load the required element.
  Dim wfTemp As VB.Control
    
  Set wfTemp = LoadNewElementOfType(piElementType)
        
  With wfTemp
    If piElementType <> elem_Begin Then
      .Top = 1000 - picDefinition.Top
      .Left = 1000 - picDefinition.Left
    End If
    
    ' AE20080609 Fault #13202
    '.Visible = True
    If .ElementType = elem_Begin Then
      .Visible = True
    End If
    
'    If Not mblnLoading Then
'      SelectElement wfTemp
'
'      ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
'      miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
'    End If
    
      .ZOrder 0
  End With

  ' Remember what we've just done, so that we can undo it.
  If Not mblnLoading Then
    SetLastActionFlag giACTION_DROPCONTROL
    ReDim mactlUndoControls(1)
    Set mactlUndoControls(1) = wfTemp
  End If
  
  Set AddElement = wfTemp

  IsChanged = True

End Function

Private Sub AddLinks()
  ' Add links between the selected elements.
  Dim iLoop As Integer
  Dim wfTempLink As COAWF_Link
  
  ' Check that at least two elements have been selected.
  If SelectedElementCount < 2 Then
    MsgBox "Fewer than two elements have been selected. Select two or more elements to define precedence.", _
      vbExclamation + vbOKOnly, App.ProductName
    Exit Sub
  End If
  
  ' Remember what we've just done, so that we can undo it.
  SetLastActionFlag giACTION_DROPCONTROL
  
  ' Go ahead and create the links
  For iLoop = 2 To UBound(miSelectionOrder)
    Set wfTempLink = CreateLink(miSelectionOrder(1), miSelectionOrder(iLoop))
    
    ' Remember what we've just done, so that we can undo it.
    ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
    Set mactlUndoControls(UBound(mactlUndoControls)) = wfTempLink
    
    Set wfTempLink = Nothing
  Next iLoop
  
  IsChanged = True
End Sub
Private Sub AutoFormat()
  ' AutoFormat the workflow.
  On Error GoTo ErrorTrap
  
  Dim wfElement As VB.Control
  Dim wfLink As COAWF_Link
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim sngMaxElementWidth As Single
  Dim sngMaxElementHeight As Single
  Dim aWFImmediatelyPrecedingElements() As VB.Control
  Dim iMaxColumn As Integer

  Const XSTARTOFFSET = 500
  Const YSTARTOFFSET = 500
  Const XCOLUMNOFFSET = 2000
  Dim YCOLUMNOFFSET As Single

  sngMaxElementWidth = 0
  sngMaxElementHeight = 0
  
  Screen.MousePointer = vbHourglass
  
  UI.LockWindow Me.hWnd
  
  ' Create an array of elements that have been autoFormatted.
  ' Column 1 = element index
  ' Column 2 = column
  ' Column 3 = row
  ReDim malngElementsDone(3, 0)
  
  For Each wfElement In mcolwfElements
    With wfElement
      If .Visible Then
        If (.ElementType = elem_Begin) Then
      
          ReDim Preserve malngElementsDone(3, UBound(malngElementsDone, 2) + 1)
          malngElementsDone(1, UBound(malngElementsDone, 2)) = .ControlIndex
          malngElementsDone(2, UBound(malngElementsDone, 2)) = 0
          malngElementsDone(3, UBound(malngElementsDone, 2)) = 0
          
          AutoFormatElement wfElement
        End If
        
        sngMaxElementWidth = IIf(sngMaxElementWidth < .Width, .Width, sngMaxElementWidth)
        sngMaxElementHeight = IIf(sngMaxElementHeight < .Height, .Height, sngMaxElementHeight)
        
        Exit For
      End If
    End With
  Next wfElement
  Set wfElement = Nothing

  For Each wfElement In mcolwfElements
    With wfElement
      If .Visible Then
        If (.ElementType <> elem_Begin) Then
      
          ReDim aWFImmediatelyPrecedingElements(1)
          Set aWFImmediatelyPrecedingElements(UBound(aWFImmediatelyPrecedingElements)) = wfElement
          ImmediatelyPrecedingElements wfElement, aWFImmediatelyPrecedingElements
      
          If UBound(aWFImmediatelyPrecedingElements) <= 1 Then
            iMaxColumn = -1
            For iLoop = 1 To UBound(malngElementsDone, 2)
              If iMaxColumn < malngElementsDone(2, iLoop) Then
                iMaxColumn = malngElementsDone(2, iLoop)
              End If
            Next iLoop
  
            ReDim Preserve malngElementsDone(3, UBound(malngElementsDone, 2) + 1)
            malngElementsDone(1, UBound(malngElementsDone, 2)) = .ControlIndex
            malngElementsDone(2, UBound(malngElementsDone, 2)) = iMaxColumn + 1
            malngElementsDone(3, UBound(malngElementsDone, 2)) = 0
            
            AutoFormatElement wfElement
          End If
        End If
        
        sngMaxElementWidth = IIf(sngMaxElementWidth < .Width, .Width, sngMaxElementWidth)
        sngMaxElementHeight = IIf(sngMaxElementHeight < .Height, .Height, sngMaxElementHeight)
      End If
    End With
  Next wfElement
  Set wfElement = Nothing

  YCOLUMNOFFSET = sngMaxElementHeight + 700
  
  For iLoop = 1 To UBound(malngElementsDone, 2)
    ' Format the controls
    mcolwfElements(CStr(malngElementsDone(1, iLoop))).Left = XSTARTOFFSET + (XCOLUMNOFFSET * malngElementsDone(2, iLoop)) + ((sngMaxElementWidth - mcolwfElements(CStr(malngElementsDone(1, iLoop))).Width) / 2)
    mcolwfElements(CStr(malngElementsDone(1, iLoop))).Top = YSTARTOFFSET + (YCOLUMNOFFSET * malngElementsDone(3, iLoop))
  Next iLoop
  
  For Each wfLink In ASRWFLink1
    If wfLink.Visible Then
      FormatLink wfLink
    End If
  Next wfLink
  Set wfLink = Nothing
  
  ResizeCanvas

  IsChanged = True

TidyUpAndExit:
  UI.UnlockWindow
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
End Sub

Private Sub RememberOriginalExpressions()
  ' Read all of the Workflows original expressions.
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim objExpression As CExpression
  
  ReDim maobjOriginalExpressions(0)
  
  sSQL = "SELECT tmpExpressions.exprID" & _
    " FROM tmpExpressions" & _
    " WHERE tmpExpressions.utilityID = " & CStr(mlngWorkflowID) & _
    "   AND tmpExpressions.deleted = FALSE" & _
    "   AND tmpExpressions.parentComponentID = 0"
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  Do While Not rsTemp.EOF
    Set objExpression = New CExpression
    objExpression.ExpressionID = rsTemp!ExprID
    objExpression.ConstructExpression
    
    ReDim Preserve maobjOriginalExpressions(UBound(maobjOriginalExpressions) + 1)
    Set maobjOriginalExpressions(UBound(maobjOriginalExpressions)) = objExpression
    Set objExpression = Nothing
    
    rsTemp.MoveNext
  Loop
  rsTemp.Close
  Set rsTemp = Nothing
    
End Sub

Private Sub RestoreOriginalExpressions()
  ' Restore the Workflows original expressions.
  Dim sSQL As String
  Dim objExpression As CExpression
  Dim iLoop As Integer
  Dim sOriginalExprIDs As String
  Dim rsTemp As DAO.Recordset
  Dim aWFAllElements() As VB.Control
  
  ReDim aWFAllElements(0)
  AllElements aWFAllElements
  
  sOriginalExprIDs = "0"
  
  For iLoop = 1 To UBound(maobjOriginalExpressions)
    Set objExpression = maobjOriginalExpressions(iLoop)
    
    sSQL = "UPDATE tmpExpressions" & _
      " SET deleted = FALSE" & _
      " WHERE exprID = " & CStr(objExpression.ExpressionID)
    daoDb.Execute sSQL, dbFailOnError
        
    sOriginalExprIDs = sOriginalExprIDs & "," & CStr(objExpression.ExpressionID)
    objExpression.EvaluatedReturnType = objExpression.ReturnType
    
    objExpression.WriteExpression_Transaction
    
    Set objExpression = Nothing
  Next iLoop
  
  'JPD 20070615 Fault 12335
  '' Mark any 'live' expressions that were newly created as deleted.
  'sSQL = "UPDATE tmpExpressions" & _
  '  " SET deleted = TRUE" & _
  '  " WHERE exprID NOT IN (" & sOriginalExprIDs & ")" & _
  '  "   AND utilityID = " & CStr(mlngWorkflowID)
  'daoDb.Execute sSQL, dbFailOnError
  sSQL = "SELECT tmpExpressions.exprID" & _
    " FROM tmpExpressions" & _
    " WHERE tmpExpressions.utilityID = " & CStr(mlngWorkflowID) & _
    "   AND exprID NOT IN (" & sOriginalExprIDs & ")" & _
    "   AND tmpExpressions.deleted = FALSE" & _
    "   AND tmpExpressions.parentComponentID = 0"
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  Do While Not rsTemp.EOF
    Set objExpression = New CExpression
    objExpression.ExpressionID = rsTemp!ExprID
    objExpression.AllWorkflowElements = aWFAllElements
    objExpression.DeleteExpression
    Set objExpression = Nothing
    
    rsTemp.MoveNext
  Loop
  rsTemp.Close
  Set rsTemp = Nothing
    
End Sub

Public Function WorkflowExpressionsChanged() As Boolean
  ' Return TRUE if any of the Workflow expressions have been modified or created.
  Dim fExpressionsChanged As Boolean
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim fFound As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim objExpression As CExpression
  Dim alngCurrentExpressions() As Long
  
  If Not mfExpressionsChanged Then
    ReDim alngCurrentExpressions(0)
    
    ' Check each 'live' expression (ie. the one currently to be saved)
    sSQL = "SELECT tmpExpressions.exprID, tmpExpressions.lastSave" & _
      " FROM tmpExpressions" & _
      " WHERE tmpExpressions.utilityID = " & CStr(mlngWorkflowID) & _
      "   AND tmpExpressions.deleted = FALSE" & _
      "   AND tmpExpressions.parentComponentID = 0"
      
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
    Do While Not rsTemp.EOF
      ' Check if the lastSave value of the 'live' expression is more recent than the one read when this definition was loaded.
      fFound = False
      
      ReDim Preserve alngCurrentExpressions(UBound(alngCurrentExpressions) + 1)
      alngCurrentExpressions(UBound(alngCurrentExpressions)) = rsTemp!ExprID
      
      For iLoop = 1 To UBound(maobjOriginalExpressions)
        Set objExpression = maobjOriginalExpressions(iLoop)

        If rsTemp!ExprID = objExpression.ExpressionID Then
          fExpressionsChanged = (rsTemp!LastSave <> objExpression.LastSave)
          fFound = True
          Set objExpression = Nothing
          Exit For
        End If
      
        Set objExpression = Nothing
      Next iLoop
      
      If Not fFound Then
        fExpressionsChanged = True
      End If
      
      If fExpressionsChanged Then
        Exit Do
      End If
      
      rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    If Not fExpressionsChanged Then
      ' All 'live' expressions are unchanged from the original values, but have any original ones been deleted?
      For iLoop = 1 To UBound(maobjOriginalExpressions)
        Set objExpression = maobjOriginalExpressions(iLoop)
        fFound = False
        
        For iLoop2 = 1 To UBound(alngCurrentExpressions)
          If objExpression.ExpressionID = alngCurrentExpressions(iLoop2) Then
            fFound = True
            Exit For
          End If
        Next iLoop2
        
        Set objExpression = Nothing
        
        If Not fFound Then
          ' Original expression no longer 'live', so must have been deleted.
          fExpressionsChanged = True
          Exit For
        End If
      Next iLoop
    End If
  
    mfExpressionsChanged = fExpressionsChanged
  End If

  WorkflowExpressionsChanged = mfExpressionsChanged
  
End Function

Private Sub ManualResizeCanvas()
  ' Resize the workflow canvas.
  On Error GoTo ErrorTrap
  
  Dim frmResize As frmWorkflowEditOptions
  
  Set frmResize = New frmWorkflowEditOptions
  
  frmResize.CanvasHeight = picDefinition.Height
  frmResize.CanvasWidth = picDefinition.Width
      
  frmResize.Show vbModal

  If Not frmResize.Cancelled Then
    picDefinition.Height = frmResize.CanvasHeight
    picDefinition.Width = frmResize.CanvasWidth
  
    ' Ensure the canvas is still big enough for all elements.
    ResizeCanvas
  
    IsChanged = True
    Form_Resize
  End If
  
  UnLoad frmResize
  Set frmResize = Nothing

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Stop
  Resume TidyUpAndExit
End Sub


Private Sub AutoFormatElement(pwfElement As VB.Control)
  ' AutoFormat the elements that the given element is linked to (if they haven't already been autoFormatted).
  On Error GoTo ErrorTrap
  
  Dim wfLink As COAWF_Link
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim fElementDone As Boolean
  Dim iColumn As Integer
  Dim iCurrentColumn As Integer
  Dim iCurrentRow As Integer
  Dim iMaxColumn As Integer
  Dim iMaxColumnRow As Integer
  Dim avOutboundFlowInfo() As Variant
  Dim avOutboundFlowOrder() As Variant
  Dim iFlowCode As Integer
  Dim iLastOrder As Integer
  
  iColumn = 0
  iMaxColumn = -1
  
  ' Get the array of outbound flow information from the start element.
  ' Column 1 = Tag (see enums, or -1 if there's only a single outbound flow)
  ' Column 2 = Direction
  ' Column 3 = XOffset
  ' Column 4 = YOffset
  ' Column 5 = Maximum
  ' Column 6 = Minimum
  ' Column 7 = Description
  avOutboundFlowInfo = pwfElement.OutboundFlows_Information
  
  For iLoop = 1 To UBound(malngElementsDone, 2)
    If pwfElement.ControlIndex = malngElementsDone(1, iLoop) Then
      iCurrentColumn = malngElementsDone(2, iLoop)
      iCurrentRow = malngElementsDone(3, iLoop)
    End If
    
    If malngElementsDone(2, iLoop) > iMaxColumn Then
      iMaxColumn = malngElementsDone(2, iLoop)
      iMaxColumnRow = malngElementsDone(3, iLoop)
    End If
  Next iLoop
  
  ' Work out what order to format the outbound flows.
  iLastOrder = 1
  ReDim avOutboundFlowOrder(UBound(avOutboundFlowInfo, 2))
    
  For iLoop = 1 To UBound(avOutboundFlowInfo, 2)
    If avOutboundFlowInfo(2, iLoop) = lineDirection_Up Then
      avOutboundFlowOrder(iLastOrder) = iLoop
      iLastOrder = iLastOrder + 1
    End If
  Next iLoop
    
  For iLoop = 1 To UBound(avOutboundFlowInfo, 2)
    If avOutboundFlowInfo(2, iLoop) = lineDirection_Left Then
      avOutboundFlowOrder(iLastOrder) = iLoop
      iLastOrder = iLastOrder + 1
    End If
  Next iLoop
    
  For iLoop = 1 To UBound(avOutboundFlowInfo, 2)
    If avOutboundFlowInfo(2, iLoop) = lineDirection_Down Then
      avOutboundFlowOrder(iLastOrder) = iLoop
      iLastOrder = iLastOrder + 1
    End If
  Next iLoop
    
  For iLoop = 1 To UBound(avOutboundFlowInfo, 2)
    If avOutboundFlowInfo(2, iLoop) = lineDirection_Right Then
      avOutboundFlowOrder(iLastOrder) = iLoop
      iLastOrder = iLastOrder + 1
    End If
  Next iLoop
      
  For iLoop2 = 1 To UBound(avOutboundFlowOrder)
    iFlowCode = avOutboundFlowInfo(1, avOutboundFlowOrder(iLoop2))
    
    For Each wfLink In ASRWFLink1
      If wfLink.Visible Then
        If (wfLink.StartElementIndex = pwfElement.ControlIndex) And _
          ((wfLink.StartOutboundFlowCode < 0) Or (wfLink.StartOutboundFlowCode = iFlowCode)) Then
          
          fElementDone = False
          For iLoop = 1 To UBound(malngElementsDone, 2)
            If wfLink.EndElementIndex = malngElementsDone(1, iLoop) Then
              fElementDone = True
              
              ' Linked element HAS been formatted. Ensure that the current element is not in the same
              ' column as the linked element.
              If iCurrentColumn = malngElementsDone(2, iLoop) Then
                
              End If
              
              Exit For
            End If
          Next iLoop

          If Not fElementDone Then
            ReDim Preserve malngElementsDone(3, UBound(malngElementsDone, 2) + 1)
            malngElementsDone(1, UBound(malngElementsDone, 2)) = wfLink.EndElementIndex
            malngElementsDone(2, UBound(malngElementsDone, 2)) = IIf(iMaxColumnRow > iCurrentRow, iMaxColumn + 1, iCurrentColumn)
            malngElementsDone(3, UBound(malngElementsDone, 2)) = iCurrentRow + 1
                    
            AutoFormatElement mcolwfElements(CStr(wfLink.EndElementIndex))
          
            iCurrentColumn = iCurrentColumn + 1
            For iLoop = 1 To UBound(malngElementsDone, 2)
              If malngElementsDone(2, iLoop) >= iCurrentColumn Then
                iCurrentColumn = malngElementsDone(2, iLoop) + 1
              End If
            Next iLoop
          End If
        End If
      End If
    Next wfLink
    Set wfLink = Nothing
  Next iLoop2
  
  
''' STILL TO DO
''' 1 - Connector2 elements, etc.
''' 2 - Always do Decision-False flows to the next column
''' 3 - If links go upwards to elements in the same column, force the current element into the next column.

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Stop
  Resume TidyUpAndExit
End Sub
Private Sub CancelElementAddMode()
  Dim objTool As ActiveBarLibraryCtl.Tool
  
  For Each objTool In abMenu.Tools
    objTool.Checked = False
  Next objTool
  Set objTool = Nothing
  
  Me.MousePointer = vbNormal

End Sub

Private Function CanDeleteElementsAndLinks() As Boolean
  ' Return true if the selected elements can be deleted.
  Dim avIdentifierLog() As Variant
  Dim wfElement As VB.Control
  Dim wfTempElement As VB.Control
  Dim wfLink As COAWF_Link
  Dim fOK As Boolean
  Dim asMessages() As String
  Dim frmUsage As frmUsage
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iElementCount As Integer
  Dim iLinkCount As Integer
  Dim avarElementsToValidate() As Variant
  Dim avarDisconnectedElements() As Variant
  Dim fFound As Boolean
  Dim sMsg As String
  Dim aWFSucceedingElements() As VB.Control
  Dim aWFPrecedingElements() As VB.Control

  fOK = True
  iElementCount = 0
  iLinkCount = 0

  ' Clear the array of validation messages
  ' Column 0 = The message
  ReDim asMessages(0)

  ' Column 1 = the element object itself
  ' Column 2 = TRUE if object is being deleted
  ReDim avarElementsToValidate(2, 0)
  
  ' Column 1 = the element object itself
  ' Column 2 = TRUE if object is being deleted
  ReDim avarDisconnectedElements(2, 0)

  For Each wfElement In mcolwfElements
    If wfElement.Highlighted And _
      (wfElement.Visible) Then

      iElementCount = iElementCount + 1

      ' Determine which elements will need revalidating.
      fFound = False
      For iLoop2 = 1 To UBound(avarElementsToValidate, 2)
        If avarElementsToValidate(1, iLoop2) Is wfElement Then
          avarElementsToValidate(2, iLoop2) = True
          fFound = True
          Exit For
        End If
      Next iLoop2

      If Not fFound Then
        ReDim Preserve avarElementsToValidate(2, UBound(avarElementsToValidate, 2) + 1)
        Set avarElementsToValidate(1, UBound(avarElementsToValidate, 2)) = wfElement
        avarElementsToValidate(2, UBound(avarElementsToValidate, 2)) = True
      End If

      ' Need to validate succeeding elements to as the associated link deletion may also now cause validation exceptions.
      ReDim aWFSucceedingElements(1)
      Set aWFSucceedingElements(UBound(aWFSucceedingElements)) = wfElement
      SucceedingElements wfElement, aWFSucceedingElements

      For iLoop = 2 To UBound(aWFSucceedingElements) ' Ignore index 1 as it is the current element, already checked.
        Set wfTempElement = aWFSucceedingElements(iLoop)

        fFound = False
        For iLoop2 = 1 To UBound(avarElementsToValidate, 2)
          If avarElementsToValidate(1, iLoop2) Is wfTempElement Then
            fFound = True
            Exit For
          End If
        Next iLoop2

        If Not fFound Then
          ReDim Preserve avarElementsToValidate(2, UBound(avarElementsToValidate, 2) + 1)
          Set avarElementsToValidate(1, UBound(avarElementsToValidate, 2)) = wfTempElement
          avarElementsToValidate(2, UBound(avarElementsToValidate, 2)) = False
        End If

        Set wfTempElement = Nothing
      Next iLoop
      
      ' Determine which elements will need to be treated as deleted/disconnected when revalidating.
      fFound = False
      For iLoop2 = 1 To UBound(avarDisconnectedElements, 2)
        If avarDisconnectedElements(1, iLoop2) Is wfElement Then
          avarDisconnectedElements(2, iLoop2) = True
          fFound = True
          Exit For
        End If
      Next iLoop2

      If Not fFound Then
        ReDim Preserve avarDisconnectedElements(2, UBound(avarDisconnectedElements, 2) + 1)
        Set avarDisconnectedElements(1, UBound(avarDisconnectedElements, 2)) = wfElement
        avarDisconnectedElements(2, UBound(avarDisconnectedElements, 2)) = True
      End If

      ' Need to validate preceding elements to as the associated link deletion may also now cause validation exceptions.
      ReDim aWFPrecedingElements(1)
      Set aWFPrecedingElements(UBound(aWFPrecedingElements)) = wfElement
      PrecedingElements wfElement, aWFPrecedingElements

      For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as it is the current element, already checked.
        Set wfTempElement = aWFPrecedingElements(iLoop)

        fFound = False
        For iLoop2 = 1 To UBound(avarDisconnectedElements, 2)
          If avarDisconnectedElements(1, iLoop2) Is wfTempElement Then
            fFound = True
            Exit For
          End If
        Next iLoop2

        If Not fFound Then
          ReDim Preserve avarDisconnectedElements(2, UBound(avarDisconnectedElements, 2) + 1)
          Set avarDisconnectedElements(1, UBound(avarDisconnectedElements, 2)) = wfTempElement
          avarDisconnectedElements(2, UBound(avarDisconnectedElements, 2)) = False
        End If

        Set wfTempElement = Nothing
      Next iLoop
    End If
  Next wfElement
  Set wfElement = Nothing

  For Each wfLink In ASRWFLink1
    If wfLink.Highlighted And (wfLink.Visible) Then
      iLinkCount = iLinkCount + 1

      Set wfElement = mcolwfElements(CStr(wfLink.EndElementIndex))

      ' Determine which elements will need revalidating.
      fFound = False
      For iLoop2 = 1 To UBound(avarElementsToValidate, 2)
        If avarElementsToValidate(1, iLoop2) Is wfElement Then
          fFound = True
          Exit For
        End If
      Next iLoop2

      If Not fFound Then
        ReDim Preserve avarElementsToValidate(2, UBound(avarElementsToValidate, 2) + 1)
        Set avarElementsToValidate(1, UBound(avarElementsToValidate, 2)) = wfElement
        avarElementsToValidate(2, UBound(avarElementsToValidate, 2)) = False
      End If

      ' Need to validate succeeding elements to as the associated link deletion may also now cause validation exceptions.
      ReDim aWFSucceedingElements(1)
      Set aWFSucceedingElements(UBound(aWFSucceedingElements)) = wfElement
      SucceedingElements wfElement, aWFSucceedingElements

      For iLoop = 1 To UBound(aWFSucceedingElements)
        Set wfTempElement = aWFSucceedingElements(iLoop)

        fFound = False
        For iLoop2 = 1 To UBound(avarElementsToValidate, 2)
          If avarElementsToValidate(1, iLoop2) Is wfTempElement Then
            fFound = True
            Exit For
          End If
        Next iLoop2

        If Not fFound Then
          ReDim Preserve avarElementsToValidate(2, UBound(avarElementsToValidate, 2) + 1)
          Set avarElementsToValidate(1, UBound(avarElementsToValidate, 2)) = wfTempElement
          avarElementsToValidate(2, UBound(avarElementsToValidate, 2)) = False
        End If

        Set wfTempElement = Nothing
      Next iLoop

      Set wfElement = Nothing
      
      ' Need to validate precediting elements to as the associated link deletion may also now cause validation exceptions.
      Set wfElement = mcolwfElements(CStr(wfLink.StartElementIndex))
      
      fFound = False
      For iLoop2 = 1 To UBound(avarDisconnectedElements, 2)
        If avarDisconnectedElements(1, iLoop2) Is wfElement Then
          fFound = True
          Exit For
        End If
      Next iLoop2

      If Not fFound Then
        ReDim Preserve avarDisconnectedElements(2, UBound(avarDisconnectedElements, 2) + 1)
        Set avarDisconnectedElements(1, UBound(avarDisconnectedElements, 2)) = wfElement
        avarDisconnectedElements(2, UBound(avarDisconnectedElements, 2)) = False
      End If
      
      ReDim aWFPrecedingElements(1)
      Set aWFPrecedingElements(UBound(aWFPrecedingElements)) = wfElement
      PrecedingElements wfElement, aWFPrecedingElements

      For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as it is the link's end element, no check required.
        Set wfTempElement = aWFPrecedingElements(iLoop)

        fFound = False
        For iLoop2 = 1 To UBound(avarDisconnectedElements, 2)
          If avarDisconnectedElements(1, iLoop2) Is wfTempElement Then
            fFound = True
            Exit For
          End If
        Next iLoop2

        If Not fFound Then
          ReDim Preserve avarDisconnectedElements(2, UBound(avarDisconnectedElements, 2) + 1)
          Set avarDisconnectedElements(1, UBound(avarDisconnectedElements, 2)) = wfTempElement
          avarDisconnectedElements(2, UBound(avarDisconnectedElements, 2)) = False
        End If

        Set wfTempElement = Nothing
      Next iLoop
    
      Set wfElement = Nothing
    End If
  Next wfLink
  Set wfLink = Nothing

  ' Clear the array of validation messages
  'Column 0 = The message
  'Column 1 = Associated element index
  ReDim mavValidationMessages(1, 0)
  mfFixableValidationFailures = False

  For iLoop = 1 To UBound(avarElementsToValidate, 2)
    If Not CBool(avarElementsToValidate(2, iLoop)) Then
      ' Perform the element specific checks.
      Set wfElement = avarElementsToValidate(1, iLoop)
      ValidateElement wfElement, False, avarDisconnectedElements
    End If
  Next iLoop

  If UBound(mavValidationMessages, 2) > 0 Then
    Set frmUsage = New frmUsage
    frmUsage.ResetList
      
    For iLoop = 1 To UBound(mavValidationMessages, 2)
      frmUsage.AddToList CStr(mavValidationMessages(0, iLoop)), mavValidationMessages(1, iLoop)
    Next iLoop

    Screen.MousePointer = vbDefault

    frmUsage.Width = (3 * Screen.Width / 4)

    sMsg = ""
    If iElementCount > 0 Then
      sMsg = "element" & IIf(iElementCount > 1, "s", "")
    End If
    If iLinkCount > 0 Then
      sMsg = sMsg & _
        IIf(Len(sMsg) > 0, " and", "") _
        & "link" & IIf(iLinkCount > 1, "s", "")
    End If
    sMsg = sMsg & _
      IIf(iElementCount + iLinkCount = 1, " is", " are")

    frmUsage.ShowMessage "Workflow '" & Trim(msWorkflowName) & "'", _
      "The following validation exceptions will occur if the selected " & sMsg & " deleted:" & _
      vbCrLf & "Do you wish to continue?", UsageCheckObject.Workflow, _
      USAGEBUTTONS_PRINT + USAGEBUTTONS_YES + USAGEBUTTONS_NO, "validation"

    fOK = (frmUsage.Choice = vbYes)

    UnLoad frmUsage
    Set frmUsage = Nothing
  End If

  CanDeleteElementsAndLinks = fOK
  
End Function

Private Sub CopyElementProperties(pwfSourceElement As VB.Control, pwfDestElement As VB.Control)
  ' Copy the properties from one element to another.
  With pwfDestElement
  
    .Left = pwfSourceElement.Left
    .Top = pwfSourceElement.Top

    Select Case pwfSourceElement.ElementType
    Case elem_WebForm
      .Identifier = pwfSourceElement.Identifier
      .DescriptionExprID = pwfSourceElement.DescriptionExprID
      .DescriptionHasWorkflowName = pwfSourceElement.DescriptionHasWorkflowName
      .DescriptionHasElementCaption = pwfSourceElement.DescriptionHasElementCaption
      .WebFormFGColor = pwfSourceElement.WebFormFGColor
      .WebFormBGColor = pwfSourceElement.WebFormBGColor
      .WebFormBGImageID = pwfSourceElement.WebFormBGImageID
      .WebFormBGImageLocation = pwfSourceElement.WebFormBGImageLocation
      .WebFormWidth = pwfSourceElement.WebFormWidth
      .WebFormHeight = pwfSourceElement.WebFormHeight
      .WebFormTimeoutFrequency = pwfSourceElement.WebFormTimeoutFrequency
      .WebFormTimeoutPeriod = pwfSourceElement.WebFormTimeoutPeriod
      .WebFormTimeoutExcludeWeekend = pwfSourceElement.WebFormTimeoutExcludeWeekend
      .Items = pwfSourceElement.Items
      .Validations = pwfSourceElement.Validations
      Set .WebFormDefaultFont = pwfSourceElement.WebFormDefaultFont
'      .WebFormDefaultFont.Name = pwfSourceElement.Font.Name
'      .WebFormDefaultFont.Size = pwfSourceElement.Font.Size
'      .WebFormDefaultFont.Bold = pwfSourceElement.Font.Bold
'      .WebFormDefaultFont.Italic = pwfSourceElement.Font.Italic
'      .WebFormDefaultFont.Strikethrough = pwfSourceElement.Font.Strikethrough
'      .WebFormDefaultFont.Underline = pwfSourceElement.Font.Underline
      .WFCompletionMessageType = pwfSourceElement.WFCompletionMessageType
      .WFCompletionMessage = pwfSourceElement.WFCompletionMessage
      .WFSavedForLaterMessageType = pwfSourceElement.WFSavedForLaterMessageType
      .WFSavedForLaterMessage = pwfSourceElement.WFSavedForLaterMessage
      .WFFollowOnFormsMessageType = pwfSourceElement.WFFollowOnFormsMessageType
      .WFFollowOnFormsMessage = pwfSourceElement.WFFollowOnFormsMessage

    Case elem_Email
      .Identifier = pwfSourceElement.Identifier
      .EmailID = pwfSourceElement.EmailID
      .EmailCCID = pwfSourceElement.EmailCCID
      .EmailRecord = pwfSourceElement.EmailRecord
      .EMailSubject = pwfSourceElement.EMailSubject
      
      .Attachment_Type = pwfSourceElement.Attachment_Type
      .Attachment_File = pwfSourceElement.Attachment_File
      .Attachment_WFElementIdentifier = pwfSourceElement.Attachment_WFElementIdentifier
      .Attachment_WFValueIdentifier = pwfSourceElement.Attachment_WFValueIdentifier
      .Attachment_DBColumnID = pwfSourceElement.Attachment_DBColumnID
      .Attachment_DBRecord = pwfSourceElement.Attachment_DBRecord
      .Attachment_DBElement = pwfSourceElement.Attachment_DBElement
      .Attachment_DBValue = pwfSourceElement.Attachment_DBValue
      
      .RecordSelectorWebFormIdentifier = pwfSourceElement.RecordSelectorWebFormIdentifier
      .RecordSelectorIdentifier = pwfSourceElement.RecordSelectorIdentifier
      .Items = pwfSourceElement.Items

    Case elem_Decision
      .Identifier = pwfSourceElement.Identifier
      .DecisionCaptionType = pwfSourceElement.DecisionCaptionType
      .DecisionFlowType = pwfSourceElement.DecisionFlowType
      .TrueFlowIdentifier = pwfSourceElement.TrueFlowIdentifier
      .DecisionFlowExpressionID = pwfSourceElement.DecisionFlowExpressionID

    Case elem_StoredData
      .Identifier = pwfSourceElement.Identifier
      .DataAction = pwfSourceElement.DataAction
      .DataColumns = pwfSourceElement.DataColumns
      .DataRecord = pwfSourceElement.DataRecord
      .DataTableID = pwfSourceElement.DataTableID
      .RecordSelectorWebFormIdentifier = pwfSourceElement.RecordSelectorWebFormIdentifier
      .RecordSelectorIdentifier = pwfSourceElement.RecordSelectorIdentifier
      .DataRecordTableID = pwfSourceElement.DataRecordTableID
      .SecondaryDataRecord = pwfSourceElement.SecondaryDataRecord
      .SecondaryRecordSelectorWebFormIdentifier = pwfSourceElement.SecondaryRecordSelectorWebFormIdentifier
      .SecondaryRecordSelectorIdentifier = pwfSourceElement.SecondaryRecordSelectorIdentifier
      .SecondaryDataRecordTableID = pwfSourceElement.SecondaryDataRecordTableID
      .UseAsTargetIdentifier = pwfSourceElement.UseAsTargetIdentifier
    
    Case Else
      .ElementType = pwfSourceElement.ElementType
      
    End Select
    
    .Caption = pwfSourceElement.Caption
    .ConnectorPairIndex = pwfSourceElement.ConnectorPairIndex
    
    .Highlighted = False
    .Visible = False
  End With

End Sub

Private Sub CopyLinkProperties(pwfSourceLink As COAWF_Link, pwfDestLink As COAWF_Link)
  ' Copy the properties from one link to another.
  With pwfDestLink
    .StartElementIndex = pwfSourceLink.StartElementIndex
    .EndElementIndex = pwfSourceLink.EndElementIndex
    .StartOutboundFlowCode = pwfSourceLink.StartOutboundFlowCode

    .Highlighted = False
    .Visible = False
  End With

End Sub


Private Function CreateLink(piStartElementIndex As Integer, _
  piEndElementIndex As Integer, _
  Optional piOutboundFlowCode As Variant) As COAWF_Link
  
  ' Create the link between the given elements.
  Dim wfStartElement As VB.Control
  Dim wfEndElement As VB.Control
  Dim wfNewLink As COAWF_Link
  Dim wfTempLink As COAWF_Link
  Dim iLinkCount As Integer
  Dim sngStartXOffset As Single
  Dim sngStartYOffset As Single
  Dim sngEndXOffset As Single
  Dim sngEndYOffset As Single
  Dim avOutboundFlowInfo() As Variant
  Dim iOutboundFlowCode As Integer
  Dim iOutboundFlowIndex As Integer
  Dim frmWorkflowPrompt As frmWorkflowPrompt
  Dim iLoop As Integer
  
  iOutboundFlowCode = -1
  
  Set wfStartElement = mcolwfElements(CStr(piStartElementIndex))
  Set wfEndElement = mcolwfElements(CStr(piEndElementIndex))
  
  ' Get the array of outbound flow information from the start element.
  ' Column 1 = Tag (see enums, or -1 if there's only a single outbound flow)
  ' Column 2 = Direction
  ' Column 3 = XOffset
  ' Column 4 = YOffset
  ' Column 5 = Maximum
  ' Column 6 = Minimum
  ' Column 7 = Description
  avOutboundFlowInfo = wfStartElement.OutboundFlows_Information

  ' How many outbound flows?
  If UBound(avOutboundFlowInfo, 2) = 0 Then
    ' No outbound flows from the element. Do nothing.
    MsgBox "The selected element has no outbound flows.", _
      vbExclamation + vbOKOnly, App.ProductName
    picDefinition.BackColor = vbInactiveTitleBar
    Set frmWorkflowPrompt = Nothing
    Exit Function
  Else
    If UBound(avOutboundFlowInfo, 2) > 1 Then
      ' Find out which outbound flow the link is for.
      Set frmWorkflowPrompt = New frmWorkflowPrompt
      With frmWorkflowPrompt
        Set .Element = wfStartElement
        
        If Not IsMissing(piOutboundFlowCode) Then
          frmWorkflowPrompt.OutboundFlowCode = piOutboundFlowCode
        End If
        
        .Show vbModal
      End With
      
      If frmWorkflowPrompt.Cancelled Then
        Exit Function
      End If
      
      iOutboundFlowCode = frmWorkflowPrompt.OutboundFlowCode

      Set frmWorkflowPrompt = Nothing
    End If
  End If
  
  If iOutboundFlowCode < 0 Then
    iOutboundFlowIndex = 1
  Else
    For iLoop = 1 To UBound(avOutboundFlowInfo, 2)
      If avOutboundFlowInfo(1, iLoop) = iOutboundFlowCode Then
        iOutboundFlowIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If

  ' Check that the link is valid.
  ' The startElement cannot have exceeded its maximum number of outbound flows.
  If avOutboundFlowInfo(5, iOutboundFlowIndex) >= 0 Then
    iLinkCount = 0
    For Each wfTempLink In ASRWFLink1
      If wfTempLink.Visible Then
        If wfTempLink.StartElementIndex = piStartElementIndex Then
          iLinkCount = iLinkCount + 1
        End If
      End If
    Next wfTempLink
    Set wfTempLink = Nothing
    
    If (iLinkCount + 1) > avOutboundFlowInfo(5, iOutboundFlowIndex) Then
      MsgBox "The selected " & wfStartElement.ElementTypeDescription & " element already has the maximum number of outbound flows.", _
        vbExclamation + vbOKOnly, App.ProductName
      picDefinition.BackColor = vbInactiveTitleBar
      Set frmWorkflowPrompt = Nothing
      Exit Function
    End If
  End If
  
  ' The endElement cannot have exceeded its maximum number of inbound flows.
  If wfEndElement.InboundFlows_Maximum >= 0 Then
    iLinkCount = 0
    For Each wfTempLink In ASRWFLink1
      If wfTempLink.Visible Then
        If wfTempLink.EndElementIndex = piEndElementIndex Then
          iLinkCount = iLinkCount + 1
        End If
      End If
    Next wfTempLink
    Set wfTempLink = Nothing
    
    If (iLinkCount + 1) > wfEndElement.InboundFlows_Maximum Then
      MsgBox "The selected " & wfEndElement.ElementTypeDescription & " element already has the maximum number of inbound flows.", _
        vbExclamation + vbOKOnly, App.ProductName
      picDefinition.BackColor = vbInactiveTitleBar
      Set frmWorkflowPrompt = Nothing
      Exit Function
    End If
  End If
  
  ' Check that the link doesn't already exist?
  For Each wfTempLink In ASRWFLink1
    If wfTempLink.Visible Then
      If (wfTempLink.StartElementIndex = piStartElementIndex) _
        And (wfTempLink.EndElementIndex = piEndElementIndex) _
        And ((iOutboundFlowCode = wfTempLink.StartOutboundFlowCode) Or (wfEndElement.ElementType = elem_SummingJunction)) Then

        MsgBox "A link already exists between the selected elements.", _
          vbExclamation + vbOKOnly, App.ProductName
        picDefinition.BackColor = vbInactiveTitleBar
        Set frmWorkflowPrompt = Nothing
        Exit Function
      End If
    End If
  Next wfTempLink
  Set wfTempLink = Nothing
  
  ' Create the link.
  Load ASRWFLink1(ASRWFLink1.UBound + 1)
  Set wfNewLink = ASRWFLink1(ASRWFLink1.UBound)

  With wfNewLink
    .StartElementIndex = piStartElementIndex
    .EndElementIndex = piEndElementIndex
    .StartOutboundFlowCode = iOutboundFlowCode
    
    .StartDirection = avOutboundFlowInfo(2, iOutboundFlowIndex)
    .EndDirection = wfEndElement.InboundFlow_Direction
    
    FormatLink wfNewLink
    
    .Highlighted = False
    .Visible = True
    .ZOrder 0
  End With
  
  picDefinition.BackColor = vbInactiveTitleBar
  
  Set frmWorkflowPrompt = Nothing
  Set CreateLink = wfNewLink

End Function

Private Sub ClearFlowchart(pfClearBeginElement As Boolean)
  ' Clear all elements from the flowchart (except) the 'Begin' element
  Dim wfTemp As VB.Control
  Dim wfTempLink As COAWF_Link
  
  ' Remember what we've just done, so that we can undo it.
  SetLastActionFlag giACTION_DELETECONTROLS
  
  For Each wfTemp In mcolwfElements
    If wfTemp.Visible Then
      If pfClearBeginElement Or (wfTemp.ElementType <> elem_Begin) Then
        ' Hide the element (do not unload it as we need to keep it in case we 'undo' the deletion.
        wfTemp.Visible = False
        ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
        Set mactlUndoControls(UBound(mactlUndoControls)) = wfTemp
      End If
    End If
  Next wfTemp
  Set wfTemp = Nothing
  
  For Each wfTempLink In ASRWFLink1
    If wfTempLink.Visible Then
      ' Hide the link (do not unload it as we need to keep it in case we 'undo' the deletion.
      wfTempLink.Visible = False
      ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
      Set mactlUndoControls(UBound(mactlUndoControls)) = wfTempLink
    End If
  Next wfTempLink
  Set wfTempLink = Nothing
  
End Sub

Private Sub DeleteElementsAndLinks()
  ' Delete the selected elements.
  Dim iBeginElementIndex As Integer
  Dim wfElement As VB.Control
  Dim wfTempElement As VB.Control
  Dim wfLink As COAWF_Link
  
  If Not CanDeleteElementsAndLinks Then
    Exit Sub
  End If
  
  iBeginElementIndex = -1
  
  ' Remember what we've just done, so that we can undo it.
  SetLastActionFlag giACTION_DELETECONTROLS
    
  For Each wfElement In mcolwfElements
    If wfElement.Highlighted And (wfElement.Visible) Then
      If wfElement.ElementType = elem_Begin Then
        iBeginElementIndex = wfElement.ControlIndex
      Else
        ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
        Set mactlUndoControls(UBound(mactlUndoControls)) = wfElement
                
        ' Don't forget to remember the links that are automatically deleted when the elements are deleted!!!
        For Each wfLink In ASRWFLink1
          If ((wfLink.StartElementIndex = wfElement.ControlIndex) Or _
            (wfLink.EndElementIndex = wfElement.ControlIndex)) And _
            (wfLink.Visible) Then
            
            ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
            Set mactlUndoControls(UBound(mactlUndoControls)) = wfLink
          End If
        Next wfLink
        Set wfLink = Nothing

        ' Don't forget to remember the connectorPairs that are automatically deleted when the elements are deleted!!!
        If (wfElement.ElementType = elem_Connector1) Or (wfElement.ElementType = elem_Connector2) Then
          For Each wfTempElement In mcolwfElements
            If (wfTempElement.ConnectorPairIndex = wfElement.ControlIndex) Then
              ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
              Set mactlUndoControls(UBound(mactlUndoControls)) = wfTempElement
            
              'JPD 20060629 Fault 11201
              ' Don't forget to remember the links that are automatically deleted when the elements are deleted!!!
              For Each wfLink In ASRWFLink1
                If ((wfLink.StartElementIndex = wfTempElement.ControlIndex) Or _
                  (wfLink.EndElementIndex = wfTempElement.ControlIndex)) And _
                  (wfLink.Visible) Then
                  
                  ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
                  Set mactlUndoControls(UBound(mactlUndoControls)) = wfLink
                End If
              Next wfLink
              Set wfLink = Nothing
            End If
          Next wfTempElement
          Set wfTempElement = Nothing
        End If
        
        ' Now delete the element itself.
        DeleteElement wfElement, False
      End If
    End If
  Next wfElement
  Set wfElement = Nothing

  ReDim miSelectionOrder(0)

  If iBeginElementIndex > 0 Then
    MsgBox "The 'Begin' element cannot be deleted.", _
      vbExclamation + vbOKOnly, App.ProductName
      
    ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
    miSelectionOrder(UBound(miSelectionOrder)) = iBeginElementIndex
  End If

  ' Delete the selected links.
  For Each wfLink In ASRWFLink1
    If wfLink.Highlighted And (wfLink.Visible) Then
      ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
      Set mactlUndoControls(UBound(mactlUndoControls)) = wfLink

      wfLink.Visible = False
    End If
  Next wfLink
  Set wfLink = Nothing

  IsChanged = True

End Sub

Private Sub DeleteElement(pwfElement As VB.Control, pfFromDeletedConnector As Boolean)
  ' Delete the given element.
  Dim iIndex As Integer
  Dim iElementType As ElementType
  Dim wfLink As COAWF_Link
  Dim wfTempElement As VB.Control
  
  iIndex = pwfElement.ControlIndex
  iElementType = pwfElement.ElementType
  
  ' Hide the element (do not unload it as we need to keep it in case we 'undo' the deletion.
  pwfElement.Visible = False
  
  ' AE20080502 Fault #13149
  If IsWorkflowElement(pwfElement) Then
    pwfElement.Highlighted = False
    
    If mcolwfSelectedElements.Count > 0 Then
      'JPD 20080729 Fault 13302
      If IsValidCollectionItem(mcolwfSelectedElements, CStr(pwfElement.ControlIndex)) Then
        mcolwfSelectedElements.Remove CStr(pwfElement.ControlIndex)
      End If
    End If
  End If
    
  ' Delete any links to/from the deleted element.
  For Each wfLink In ASRWFLink1
    If ((wfLink.StartElementIndex = iIndex) Or (wfLink.EndElementIndex = iIndex)) _
      And (wfLink.Visible) Then
      
      ' Hide the link (do not unload it as we need to keep it in case we 'undo' the deletion.
      wfLink.Visible = False
    End If
  Next wfLink
  Set wfLink = Nothing

  ' Delete any paired connector elements.
  If Not pfFromDeletedConnector Then
    If (iElementType = elem_Connector1) Or (iElementType = elem_Connector2) Then
      For Each wfTempElement In mcolwfElements
        If (wfTempElement.ConnectorPairIndex = iIndex) Then
          DeleteElement wfTempElement, True
        End If
      Next wfTempElement
      Set wfTempElement = Nothing
    End If
  End If
  
End Sub

Private Sub SelectAllElements()
On Error GoTo Select_Err

  Dim ctrl As VB.Control
  Dim wfLink As COAWF_Link
  
  UI.LockWindow Me.hWnd
  
  Call DeselectAllElements
    
  For Each ctrl In mcolwfElements
    ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
    miSelectionOrder(UBound(miSelectionOrder)) = ctrl.ControlIndex
    
    SelectElement ctrl
  Next
  Set ctrl = Nothing
  
Select_Err_Exit:
  UI.UnlockWindow
  
  Exit Sub
  
Select_Err:
  Resume Select_Err_Exit
End Sub

Private Sub DeselectAllElements()
On Error GoTo Deselect_Err

  Dim ctrl As VB.Control
  Dim wfLink As COAWF_Link
  
  UI.LockWindow Me.hWnd
  
'  For Each ctrl In mcolwfElements
'    If IsWorkflowControl(ctrl) Then
'      'AE20080207
'      'Only deselect if already selected..... ever so slighlty better speed with lots of controls
'      If ctrl.HighLighted Then ctrl.HighLighted = False
'    End If
'    Set ctrl = Nothing
'  Next ctrl
  
  For Each ctrl In mcolwfSelectedElements
    ctrl.Highlighted = False
    mcolwfSelectedElements.Remove CStr(ctrl.ControlIndex)
  Next
  Set ctrl = Nothing
  Set mcolwfSelectedElements = Nothing
  Set mcolwfSelectedElements = New Collection

  ReDim miSelectionOrder(0)

'  For Each wfLink In ASRWFLink1
'    'AE20080207
'    'Only deselect if already selected..... ever so slighlty better speed with lots of controls
'    If wfLink.HighLighted Then wfLink.HighLighted = False
'    wfLink.ZOrder 1
'  Next wfLink
'  Set wfLink = Nothing

  For Each wfLink In mcolwfSelectedLinks
    wfLink.Highlighted = False
  Next
  Set wfLink = Nothing
  Set mcolwfSelectedLinks = Nothing
  Set mcolwfSelectedLinks = New Collection
  
Deselect_Err_Exit:
  UI.UnlockWindow
  
  Exit Sub
  
Deselect_Err:
  Resume Deselect_Err_Exit
End Sub

Private Sub SelectElement(pwfElement As VB.Control)
  
  ' Add the element to the selected collection
  If (IsWorkflowElement(pwfElement) _
    And pwfElement.Visible And (Not pwfElement.Highlighted)) Then
    
    pwfElement.Highlighted = True
    mcolwfSelectedElements.Add pwfElement, CStr(pwfElement.ControlIndex)
  End If

End Sub

Private Sub SelectLink(pwfElement As COAWF_Link)

  ' Add the link to the selected collection
  If (pwfElement.Visible And (Not pwfElement.Highlighted)) Then
    pwfElement.Highlighted = True
    
    mcolwfSelectedLinks.Add pwfElement, CStr(pwfElement.Index)
  End If
  
End Sub

Private Sub ElementEdit(pwfElement As VB.Control)
  
  Dim frmWorkflowElementEdit As frmWorkflowElementEdit
  Dim frmDes As frmWorkflowWFDesigner

  CancelElementAddMode
  
  If pwfElement.ElementType = elem_WebForm Then

    Me.Visible = False

    'Show the Web form designer.
    Set frmDes = New frmWorkflowWFDesigner
    With frmDes
      .Loading = True
      Set .CallingForm = Me
      Set .Element = pwfElement
      .Show
      .Loading = False
    End With

    ' Display the toolbox form.
    If Not mfReadOnly Then
      With frmWorkflowWFToolbox
        Set .CurrentWebForm = frmDes
        .Show
      End With
    End If
    
    ' Display the screen object properties form.
    With frmWorkflowWFItemProps
      Set .CurrentWebForm = frmDes
      .Show
    End With
    Set frmDes = Nothing
  Else

    Set frmWorkflowElementEdit = New frmWorkflowElementEdit
    With frmWorkflowElementEdit
      Set .CallingForm = Me
      Set .Element = pwfElement
      
      If Not .CanBeEdited Then
        Set frmWorkflowElementEdit = Nothing
        MsgBox "The selected element has no editable properties.", vbInformation + vbOKOnly, App.ProductName
        Exit Sub
      End If
      
      .Show vbModal
    
      If Not .Cancelled Then
        If (pwfElement.ElementType = elem_Connector1) Or _
          (pwfElement.ElementType = elem_Connector2) Then
          
          mcolwfElements(CStr(pwfElement.ConnectorPairIndex)).Caption = pwfElement.Caption
        End If
      
        IsChanged = True
      End If
    End With
  
    Set frmWorkflowElementEdit = Nothing
  
    If Not IsChanged Then
      IsChanged = WorkflowExpressionsChanged
    End If
  End If
  
End Sub

Public Sub Start()
  Dim wfElement As VB.Control
  
  DeselectAllElements
  
  For Each wfElement In mcolwfElements
    If wfElement.ElementType = elem_Begin _
      And wfElement.ControlIndex > 0 Then
    
      If Not wfElement.Highlighted Then
        ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
        miSelectionOrder(UBound(miSelectionOrder)) = wfElement.ControlIndex
      End If
             
'      wfElement.HighLighted = True
      SelectElement wfElement
      wfElement.ZOrder 0

      RefreshMenu
      
      Exit For
    End If
  Next wfElement
  Set wfElement = Nothing

  mblnStarted = True

  picDefinition.SetFocus
  
End Sub

Private Sub ValidateElement_Expression(pwfElement As VB.Control, _
  plngExprID As Long, _
  psBaseMsg As String, _
  Optional pavarDisconnectedElements As Variant)
  
  Dim aWFPrecedingElements() As VB.Control
  Dim aWFAllElements() As VB.Control
  Dim objExpression As CExpression
  Dim fValid1 As Boolean  ' Expression can be constructed
  Dim fValid2 As Boolean
  Dim iValidityCode As ExprValidationCodes
  Dim sMessagePrefix As String
  Dim sMessageSuffix As String
  Dim alngExpressions() As Long
  Dim lngLoop As Long
  Dim fDisconnectedElement As Boolean
  Dim fDoingDeleteCheck As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iIndex As Integer
  Dim fFound As Boolean
  
  fDoingDeleteCheck = Not IsMissing(pavarDisconnectedElements)
  
  ' Do nothing if there's no expression to validate.
  If plngExprID = 0 Then Exit Sub
  
  ' Get the elements that precede the given element
  ReDim aWFPrecedingElements(1)
  Set aWFPrecedingElements(1) = pwfElement
  PrecedingElements pwfElement, aWFPrecedingElements

  ReDim aWFAllElements(0)
  AllElements aWFAllElements

  If fDoingDeleteCheck Then
    ' Remove the element we're trying to delete from the array of preceding elements, if we're doing the deletion check.
    iIndex = 1
    For iLoop = 1 To UBound(aWFPrecedingElements)
      fFound = False
      For iLoop2 = 1 To UBound(pavarDisconnectedElements, 2)
        If aWFPrecedingElements(iLoop) Is pavarDisconnectedElements(1, iLoop2) Then
          fFound = True
          Exit For
        End If
      Next iLoop2
      
      If Not fFound Then
        Set aWFPrecedingElements(iIndex) = aWFPrecedingElements(iLoop)
        iIndex = iIndex + 1
      End If
    Next iLoop
    ReDim Preserve aWFPrecedingElements(iIndex - 1)
    
    ' Remove the element we're trying to delete from the array of preceding elements, if we're doing the deletion check.
    iIndex = 1
    For iLoop = 1 To UBound(aWFAllElements)
      fFound = False
      For iLoop2 = 1 To UBound(pavarDisconnectedElements, 2)
        If aWFAllElements(iLoop) Is pavarDisconnectedElements(1, iLoop2) Then
          fFound = True
          Exit For
        End If
      Next iLoop2
      
      If Not fFound Then
        Set aWFAllElements(iIndex) = aWFAllElements(iLoop)
        iIndex = iIndex + 1
      End If
    Next iLoop
    ReDim Preserve aWFAllElements(iIndex - 1)
  End If
  
  ' Get the message prefix of the given element
  sMessagePrefix = ValidateElement_MessagePrefix(pwfElement)
  sMessageSuffix = ""

  ' Get an array of all expressions used in this expression
  ' First item is this expression.
  ReDim alngExpressions(1)
  alngExpressions(1) = plngExprID
  
  Set objExpression = New CExpression
  objExpression.ExpressionID = plngExprID
  objExpression.ConstructExpression
  objExpression.ExpressionsUsedInThisExpression alngExpressions
  Set objExpression = Nothing

  ' Validate each expression
  For lngLoop = 1 To UBound(alngExpressions)
    ' Construct the required expression.
    Set objExpression = New CExpression
    objExpression.ExpressionID = alngExpressions(lngLoop)
    objExpression.PrecedingWorkflowElements = aWFPrecedingElements
    objExpression.AllWorkflowElements = aWFAllElements
    
    fValid1 = objExpression.ConstructExpression
    fValid2 = True
    
    If Not fValid1 Then
      sMessageSuffix = IIf(lngLoop = 1, "", " - sub-component")
    Else
      iValidityCode = objExpression.ValidateExpression(True)
  
      fValid2 = (iValidityCode = giEXPRVALIDATION_NOERRORS)
      If Not fValid2 Then
        sMessageSuffix = IIf(lngLoop = 1, "", " - sub-component") & " (" & objExpression.Name & ") - " & objExpression.ValidityMessage(iValidityCode)
      End If
    End If
      
    Set objExpression = Nothing
  
    If (Not fValid1) Or (Not fValid2) Then
      ValidateWorkflow_AddMessage _
        sMessagePrefix & psBaseMsg & sMessageSuffix, _
        pwfElement.ControlIndex
    End If
  Next lngLoop
  
End Sub

Private Sub ValidateElement(pwfElement As VB.Control, _
  pfFix As Boolean, _
  Optional pavarDisconnectedElements As Variant)
  
  ' Validate the individual element types.
  Dim iLoop As Integer
  Dim fValid1 As Boolean
  Dim fValid2 As Boolean
  Dim iTerminatorCount As Integer
  Dim iOtherElementCount As Integer
  Dim wfElement As VB.Control
  Dim wfSucceedingElement As VB.Control
  Dim aWFPrecedingElements() As VB.Control
  Dim aWFImmediatelySucceedingElements() As VB.Control
  Dim aWFImmediatelySucceedingElements_true() As VB.Control
  Dim aWFImmediatelySucceedingElements_false() As VB.Control
  Dim sMessagePrefix As String
  
  sMessagePrefix = ValidateElement_MessagePrefix(pwfElement)
  
  Select Case pwfElement.ElementType
    Case elem_Decision
      ValidateElement_Decision pwfElement, pavarDisconnectedElements
    Case elem_Email
      ValidateElement_Email pwfElement, pavarDisconnectedElements
    Case elem_StoredData
      ValidateElement_StoredData pwfElement, pfFix, pavarDisconnectedElements
    Case elem_WebForm
      ValidateElement_WebForm pwfElement, pavarDisconnectedElements
  End Select
  
  ' pavarDisconnectedElements is used when checking what will go wrong if you delete
  ' elements and links (see method CanDeleteElementAndLinks). Do NOT do the following
  ' standard checks on elements if we're doing this kind of validation.
  If IsMissing(pavarDisconnectedElements) Then
    '------------------------------------------------------------
    ' 1. Cannot be linked to Terminator and another element.
    '------------------------------------------------------------
    fValid1 = True
    iTerminatorCount = 0
    iOtherElementCount = 0
    If pwfElement.ElementType = elem_Decision Then
      ' Get the elements that immediately succeed the given element from the True flow.
      ReDim aWFImmediatelySucceedingElements_true(1)
      Set aWFImmediatelySucceedingElements_true(UBound(aWFImmediatelySucceedingElements_true)) = pwfElement
      ImmediatelySucceedingElements pwfElement, _
        aWFImmediatelySucceedingElements_true, _
        False, _
        decisionOutFlow_True
    
      For iLoop = 2 To UBound(aWFImmediatelySucceedingElements_true) ' Ignore index 1 as that is the current element
        Set wfSucceedingElement = aWFImmediatelySucceedingElements_true(iLoop)
        
        If wfSucceedingElement.ElementType = elem_Terminator Then
          iTerminatorCount = iTerminatorCount + 1
        ElseIf ((wfSucceedingElement.ElementType <> elem_Connector1) _
          And (wfSucceedingElement.ElementType <> elem_Connector2) _
          And (wfSucceedingElement.ElementType <> elem_Or)) Then
          
          iOtherElementCount = iOtherElementCount + 1
        End If
  
        Set wfSucceedingElement = Nothing
      Next iLoop
      
      fValid1 = (iTerminatorCount = 0) _
        Or ((iTerminatorCount = 1) And (iOtherElementCount = 0))
  
      If fValid1 Then
        iTerminatorCount = 0
        iOtherElementCount = 0
        
        ' Get the elements that immediately succeed the given element from the False flow.
        ReDim aWFImmediatelySucceedingElements_false(1)
        Set aWFImmediatelySucceedingElements_false(UBound(aWFImmediatelySucceedingElements_false)) = pwfElement
        ImmediatelySucceedingElements pwfElement, _
          aWFImmediatelySucceedingElements_false, _
          False, _
          decisionOutFlow_False
                
        For iLoop = 2 To UBound(aWFImmediatelySucceedingElements_false) ' Ignore index 1 as that is the current element
          Set wfSucceedingElement = aWFImmediatelySucceedingElements_false(iLoop)
          
          If wfSucceedingElement.ElementType = elem_Terminator Then
            iTerminatorCount = iTerminatorCount + 1
          ElseIf ((wfSucceedingElement.ElementType <> elem_Connector1) _
            And (wfSucceedingElement.ElementType <> elem_Connector2) _
            And (wfSucceedingElement.ElementType <> elem_Or)) Then
            
            iOtherElementCount = iOtherElementCount + 1
          End If
    
          Set wfSucceedingElement = Nothing
        Next iLoop
        
        fValid1 = (iTerminatorCount = 0) _
          Or ((iTerminatorCount = 1) And (iOtherElementCount = 0))
      End If
    ElseIf pwfElement.ElementType = elem_WebForm Then
      ' Get the elements that immediately succeed the given element from the Normal flow.
      ReDim aWFImmediatelySucceedingElements_true(1)
      Set aWFImmediatelySucceedingElements_true(UBound(aWFImmediatelySucceedingElements_true)) = pwfElement
      ImmediatelySucceedingElements pwfElement, _
        aWFImmediatelySucceedingElements_true, _
        False, _
        webFormOutFlow_Normal
  
      For iLoop = 2 To UBound(aWFImmediatelySucceedingElements_true) ' Ignore index 1 as that is the current element
        Set wfSucceedingElement = aWFImmediatelySucceedingElements_true(iLoop)
  
        If wfSucceedingElement.ElementType = elem_Terminator Then
          iTerminatorCount = iTerminatorCount + 1
        ElseIf ((wfSucceedingElement.ElementType <> elem_Connector1) _
          And (wfSucceedingElement.ElementType <> elem_Connector2) _
          And (wfSucceedingElement.ElementType <> elem_Or)) Then
  
          iOtherElementCount = iOtherElementCount + 1
        End If
  
        Set wfSucceedingElement = Nothing
      Next iLoop
  
      fValid1 = (iTerminatorCount = 0) _
        Or ((iTerminatorCount = 1) And (iOtherElementCount = 0))
  
      If fValid1 Then
        iTerminatorCount = 0
        iOtherElementCount = 0
  
        ' Get the elements that immediately succeed the given element from the Timeout flow.
        ReDim aWFImmediatelySucceedingElements_false(1)
        Set aWFImmediatelySucceedingElements_false(UBound(aWFImmediatelySucceedingElements_false)) = pwfElement
        ImmediatelySucceedingElements pwfElement, _
          aWFImmediatelySucceedingElements_false, _
          False, _
          webFormOutFlow_Timeout
  
        For iLoop = 2 To UBound(aWFImmediatelySucceedingElements_false) ' Ignore index 1 as that is the current element
          Set wfSucceedingElement = aWFImmediatelySucceedingElements_false(iLoop)
  
          If wfSucceedingElement.ElementType = elem_Terminator Then
            iTerminatorCount = iTerminatorCount + 1
          ElseIf ((wfSucceedingElement.ElementType <> elem_Connector1) _
            And (wfSucceedingElement.ElementType <> elem_Connector2) _
            And (wfSucceedingElement.ElementType <> elem_Or)) Then
  
            iOtherElementCount = iOtherElementCount + 1
          End If
  
          Set wfSucceedingElement = Nothing
        Next iLoop
  
        fValid1 = (iTerminatorCount = 0) _
          Or ((iTerminatorCount = 1) And (iOtherElementCount = 0))
      End If
      
    ElseIf pwfElement.ElementType = elem_StoredData Then
      ' Get the elements that immediately succeed the given element from the Success flow.
      ReDim aWFImmediatelySucceedingElements_true(1)
      Set aWFImmediatelySucceedingElements_true(UBound(aWFImmediatelySucceedingElements_true)) = pwfElement
      ImmediatelySucceedingElements pwfElement, _
        aWFImmediatelySucceedingElements_true, _
        False, _
        storedDataOutFlow_Success
      
      For iLoop = 2 To UBound(aWFImmediatelySucceedingElements_true) ' Ignore index 1 as that is the current element
        Set wfSucceedingElement = aWFImmediatelySucceedingElements_true(iLoop)

        If wfSucceedingElement.ElementType = elem_Terminator Then
          iTerminatorCount = iTerminatorCount + 1
        ElseIf ((wfSucceedingElement.ElementType <> elem_Connector1) _
          And (wfSucceedingElement.ElementType <> elem_Connector2) _
          And (wfSucceedingElement.ElementType <> elem_Or)) Then

          iOtherElementCount = iOtherElementCount + 1
        End If

        Set wfSucceedingElement = Nothing
      Next iLoop

      fValid1 = (iTerminatorCount = 0) _
        Or ((iTerminatorCount = 1) And (iOtherElementCount = 0))

      If fValid1 Then
        iTerminatorCount = 0
        iOtherElementCount = 0

        ' Get the elements that immediately succeed the given element from the Failure flow.
        ReDim aWFImmediatelySucceedingElements_false(1)
        Set aWFImmediatelySucceedingElements_false(UBound(aWFImmediatelySucceedingElements_false)) = pwfElement
        ImmediatelySucceedingElements pwfElement, _
          aWFImmediatelySucceedingElements_false, _
          False, _
          storedDataOutFlow_Failure

        For iLoop = 2 To UBound(aWFImmediatelySucceedingElements_false) ' Ignore index 1 as that is the current element
          Set wfSucceedingElement = aWFImmediatelySucceedingElements_false(iLoop)

          If wfSucceedingElement.ElementType = elem_Terminator Then
            iTerminatorCount = iTerminatorCount + 1
          ElseIf ((wfSucceedingElement.ElementType <> elem_Connector1) _
            And (wfSucceedingElement.ElementType <> elem_Connector2) _
            And (wfSucceedingElement.ElementType <> elem_Or)) Then

            iOtherElementCount = iOtherElementCount + 1
          End If

          Set wfSucceedingElement = Nothing
        Next iLoop

        fValid1 = (iTerminatorCount = 0) _
          Or ((iTerminatorCount = 1) And (iOtherElementCount = 0))
      End If
    Else
      ' Get the elements that immediately succeed the given element.
      ReDim aWFImmediatelySucceedingElements(1)
      Set aWFImmediatelySucceedingElements(UBound(aWFImmediatelySucceedingElements)) = pwfElement
      ImmediatelySucceedingElements pwfElement, _
        aWFImmediatelySucceedingElements, _
        False

      For iLoop = 2 To UBound(aWFImmediatelySucceedingElements) ' Ignore index 1 as that is the current element
        Set wfSucceedingElement = aWFImmediatelySucceedingElements(iLoop)
        
        If wfSucceedingElement.ElementType = elem_Terminator Then
          iTerminatorCount = iTerminatorCount + 1
        ElseIf ((wfSucceedingElement.ElementType <> elem_Connector1) _
          And (wfSucceedingElement.ElementType <> elem_Connector2) _
          And (wfSucceedingElement.ElementType <> elem_Or)) Then
          
          iOtherElementCount = iOtherElementCount + 1
        End If
  
        Set wfSucceedingElement = Nothing
      Next iLoop
      
      fValid1 = (iTerminatorCount = 0) _
        Or ((iTerminatorCount = 1) And (iOtherElementCount = 0))
    End If
    
    '------------------------------------------------------------
    ' Add the required validation messages to the array.
    '------------------------------------------------------------
    ' 1. Cannot be linked to Terminator and another element.
    If Not fValid1 Then
      ValidateWorkflow_AddMessage _
        sMessagePrefix & "Cannot be succeeded by other elements as well as a terminator element", _
        pwfElement.ControlIndex
    End If
    
    '------------------------------------------------------------
    ' 2. Cannot have any cyclic flows without a web form (ie. user action) in them.
    '------------------------------------------------------------
    ' Get the elements that immediately precede the given element.
    If pwfElement.ElementType = elem_Email Or pwfElement.ElementType = elem_Decision _
      Or pwfElement.ElementType = elem_StoredData Or pwfElement.ElementType = elem_SummingJunction _
      Or pwfElement.ElementType = elem_Or Or pwfElement.ElementType = elem_Connector1 _
      Or pwfElement.ElementType = elem_Connector2 Then
        ReDim aWFPrecedingElements(1)
        Set aWFPrecedingElements(UBound(aWFPrecedingElements)) = pwfElement
        fValid2 = CycliclyValid(pwfElement, aWFPrecedingElements)
    Else
      fValid2 = True
    End If
    
    '------------------------------------------------------------
    ' Add the required validation messages to the array.
    '------------------------------------------------------------
    ' 2. Cannot have any cyclic flows without a web form (ie. user action) in them.
    If Not fValid2 Then
      ValidateWorkflow_AddMessage _
        sMessagePrefix & "Cannot have cyclic flows without user actions (ie. Web Form elements)", _
        pwfElement.ControlIndex
    End If
  End If

End Sub


Private Sub ValidateElement_Decision(pwfElement As VB.Control, _
  Optional pavarDisconnectedElements As Variant)
  
  On Error GoTo ErrorTrap
  
  Dim wfPrecedingElement As VB.Control
  Dim wfSucceedingElement1 As VB.Control
  Dim wfSucceedingElement2 As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim aWFImmediatelyPrecedingElements() As VB.Control
  'Dim aWFImmediatelySucceedingElements_true() As vb.Control
  'Dim aWFImmediatelySucceedingElements_false() As VB.Control
  Dim aWFImmediatelySucceedingElements_timeout() As VB.Control
  Dim fValid1 As Boolean
  Dim fValid2 As Boolean
  Dim fValid3 As Boolean
  'Dim fValid4 As Boolean
  Dim fValid5 As Boolean
  Dim fValid6 As Boolean
  Dim sTemp As String
  Dim asItems() As String
  Dim sMessagePrefix As String
  Dim sMessageSuffix As String
  Dim fSkipBack As Boolean
  Dim objExpression As CExpression
  Dim iValidityCode As ExprValidationCodes
  Dim aWFPrecedingElements() As VB.Control
  Dim fDisconnectedElement As Boolean
  Dim fDoingDeleteCheck As Boolean
  
  fDoingDeleteCheck = Not IsMissing(pavarDisconnectedElements)
  
  sMessagePrefix = ValidateElement_MessagePrefix(pwfElement)
  sMessageSuffix = ""
  
  ' Get the element that immediately preceeds the given element.
  ReDim aWFImmediatelyPrecedingElements(1)
  Set aWFImmediatelyPrecedingElements(UBound(aWFImmediatelyPrecedingElements)) = pwfElement
  ImmediatelyPrecedingElements pwfElement, aWFImmediatelyPrecedingElements


  'JPD 20060719 Fault 11334 - Ignore Connectors and Decision elements to
  ' determine the preceding WebForm
  fSkipBack = True

  Do While fSkipBack
    fSkipBack = False

    If UBound(aWFImmediatelyPrecedingElements) > 1 Then
      Set wfPrecedingElement = aWFImmediatelyPrecedingElements(2)

      If (wfPrecedingElement.ElementType = elem_Decision) Then
        fSkipBack = True

        ReDim aWFImmediatelyPrecedingElements(1)
        Set aWFImmediatelyPrecedingElements(UBound(aWFImmediatelyPrecedingElements)) = wfPrecedingElement
        ImmediatelyPrecedingElements wfPrecedingElement, aWFImmediatelyPrecedingElements

      ElseIf (wfPrecedingElement.ElementType = elem_Connector2) Then
        fSkipBack = True

        ReDim aWFImmediatelyPrecedingElements(1)
        Set aWFImmediatelyPrecedingElements(UBound(aWFImmediatelyPrecedingElements)) = mcolwfElements(CStr(wfPrecedingElement.ConnectorPairIndex))
        ImmediatelyPrecedingElements mcolwfElements(CStr(wfPrecedingElement.ConnectorPairIndex)), aWFImmediatelyPrecedingElements
      End If

      Set wfPrecedingElement = Nothing
    End If
  Loop
  
  '' Get the elements that immediately succeed the given element from the True flow.
  'ReDim aWFImmediatelySucceedingElements_true(1)
  'Set aWFImmediatelySucceedingElements_true(UBound(aWFImmediatelySucceedingElements_true)) = pwfElement
  'ImmediatelySucceedingElements pwfElement, _
  '  aWFImmediatelySucceedingElements_true, _
  '  True, _
  '  decisionOutFlow_True

  '' Get the elements that immediately succeed the given element from the False flow.
  'ReDim aWFImmediatelySucceedingElements_false(1)
  'Set aWFImmediatelySucceedingElements_false(UBound(aWFImmediatelySucceedingElements_false)) = pwfElement
  'ImmediatelySucceedingElements pwfElement, _
  '  aWFImmediatelySucceedingElements_false, _
  '  True, _
  '  decisionOutFlow_False
            
  '------------------------------------------------------------
  ' 1. Decision element must immediately follow a web form (if the True flow type is Button).
  ' 2. Decision element must have a TRUE FLOW identifier selected (if the True flow type is Button).
  ' 3. Decision element must have a TRUE FLOW identifier that is a button in the preceding web form (if the True flow type is Button).
  ' 4. Decision outbound flows cannot flow to the same element. - JPD YES THEY CAN!!!
  ' 5. Decision element must immediately follow a web form (via the 'normal' flow, not the 'timeout' flow).
  ' 6. Decision element must have a TRUE FLOW calculation selected (if the True flow type is Calculation).
  '------------------------------------------------------------
  fValid1 = (pwfElement.DecisionFlowType = decisionFlowType_Expression)
  fValid2 = (Len(Trim(pwfElement.TrueFlowIdentifier)) > 0) _
    Or (pwfElement.DecisionFlowType = decisionFlowType_Expression)
  fValid3 = (pwfElement.DecisionFlowType = decisionFlowType_Expression)
  'fValid4 = True
  fValid5 = True
  fValid6 = (pwfElement.DecisionFlowExpressionID > 0) _
    Or (pwfElement.DecisionFlowType = decisionFlowType_Button)
            
  For iLoop = 2 To UBound(aWFImmediatelyPrecedingElements) ' Ignore index 1 as that is the current element
    Set wfPrecedingElement = aWFImmediatelyPrecedingElements(iLoop)
              
    fDisconnectedElement = False
    If fDoingDeleteCheck Then
      For iLoop2 = 1 To UBound(pavarDisconnectedElements, 2)
        If wfPrecedingElement Is pavarDisconnectedElements(1, iLoop2) Then
          fDisconnectedElement = True
        End If
      Next iLoop2
    End If

    If Not fDisconnectedElement Then
      If wfPrecedingElement.ElementType = elem_WebForm Then
        fValid1 = True
  
        If fValid2 And (pwfElement.DecisionFlowType = decisionFlowType_Button) Then
          asItems = wfPrecedingElement.Items
  
          For iLoop2 = 1 To UBound(asItems, 2)
            If (asItems(2, iLoop2) = WorkflowWebFormItemTypes.giWFFORMITEM_BUTTON) _
              And (UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(pwfElement.TrueFlowIdentifier))) Then
  
              fValid3 = True
              Exit For
            End If
          Next iLoop2
        End If
  
        fValid5 = True
        If (pwfElement.DecisionFlowType = decisionFlowType_Button) Then
          ReDim aWFImmediatelySucceedingElements_timeout(1)
          Set aWFImmediatelySucceedingElements_timeout(UBound(aWFImmediatelySucceedingElements_timeout)) = wfPrecedingElement
          ImmediatelySucceedingElements wfPrecedingElement, _
            aWFImmediatelySucceedingElements_timeout, _
            True, _
            webFormOutFlow_Timeout
    
          For iLoop3 = 2 To UBound(aWFImmediatelySucceedingElements_timeout) ' Ignore index 1
            If aWFImmediatelySucceedingElements_timeout(iLoop3) Is pwfElement Then
              fValid5 = False
              Exit For
            End If
          Next iLoop3
        End If
        
        Exit For
      End If
    End If

    Set wfPrecedingElement = Nothing
  Next iLoop
  
  'JPD 20070829 - I reckon they should be able to go to the same element now.
  '' Check if any of the TRUE outbound flows go to the same element as the FALSE outbound flows.
  'For iLoop = 2 To UBound(aWFImmediatelySucceedingElements_true) ' Ignore index 1 as that is the current element
  '  Set wfSucceedingElement1 = aWFImmediatelySucceedingElements_true(iLoop)
  '
  '  For iLoop2 = 2 To UBound(aWFImmediatelySucceedingElements_false) ' Ignore index 1 as that is the current element
  '    Set wfSucceedingElement2 = aWFImmediatelySucceedingElements_false(iLoop2)
  '
  '    If wfSucceedingElement1.Index = wfSucceedingElement2.Index Then
  '      fValid4 = False
  '      Exit For
  '    End If
  '
  '    Set wfSucceedingElement2 = Nothing
  '  Next iLoop2
  '
  '  Set wfSucceedingElement1 = Nothing
  '
  '  If Not fValid4 Then
  '    Exit For
  '  End If
  'Next iLoop
  
  '------------------------------------------------------------
  ' Validate the true flow calculation (if required)
  '------------------------------------------------------------
  If (pwfElement.DecisionFlowExpressionID > 0) _
    And (pwfElement.DecisionFlowType = decisionFlowType_Expression) Then
    
    sTemp = GetDecisionCaptionDescription(pwfElement.DecisionCaptionType, True)
    ValidateElement_Expression _
      pwfElement, _
      pwfElement.DecisionFlowExpressionID, _
      "Invalid '" & sTemp & "' flow calculation", _
      pavarDisconnectedElements
  End If
  
  '------------------------------------------------------------
  ' Add the required validation messages to the array.
  '------------------------------------------------------------
  ' 1. Decision element must immediately follow a web form.
  ' 4. Decision outbound flows cannot flow to the same element.
  ' 5. Decision element must immediately follow a web form (via the 'normal' flow, not the 'timeout' flow).
  If Not fValid1 Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Must succeed a web form element", _
      pwfElement.ControlIndex
  End If
  'If (Not fValid4) And (Not fDoingDeleteCheck) Then
  '  ValidateWorkflow_AddMessage _
  '    sMessagePrefix & "Linked to the same element through both outbound flows", _
  '    pwfElement.Index
  'End If
  If (Not fValid5) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Cannot succeed a web form element via the 'Timeout' outbound flow", _
      pwfElement.ControlIndex
  End If

  ' 2. Decision element must have a TRUE FLOW identifier selected.
  ' 3. Decision element must have a TRUE FLOW identifier that is a button in the preceding web form.
  If ((Not fValid2) Or (Not fValid3)) Then
    'JPD 20070615 Fault 12248
    'pwfElement.TrueFlowIdentifier = ""
    
    sTemp = GetDecisionCaptionDescription(pwfElement.DecisionCaptionType, True)
    
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid '" & sTemp & "' flow button selected", _
      pwfElement.ControlIndex
  End If
                        
  ' 6. Decision element must have a TRUE FLOW calculation selected (if the True flow type is Calculation).
  If (Not fValid6) And (Not fDoingDeleteCheck) Then
    sTemp = GetDecisionCaptionDescription(pwfElement.DecisionCaptionType, True)

    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid '" & sTemp & "' flow calculation selected" & sMessageSuffix, _
      pwfElement.ControlIndex
  End If
                        
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ValidateElement_Email(pwfElement As VB.Control, _
  Optional pavarDisconnectedElements As Variant)
  
  On Error GoTo ErrorTrap
  
  Dim wfTempElement As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iLoop4 As Integer
  Dim aWFPrecedingElements() As VB.Control
  Dim fValid1 As Boolean
  Dim fValid2 As Boolean
  Dim fValid3 As Boolean
  Dim fValid4 As Boolean
  Dim fValid5 As Boolean
  Dim fValid6 As Boolean
  Dim fValid7 As Boolean
  Dim fValid8 As Boolean
  Dim fValid9 As Boolean
  Dim fValid10 As Boolean
  Dim fValid11 As Boolean
  Dim asItems() As String
  Dim asElementItems() As String
  Dim iEmailType As Integer
  Dim lngTableID As Long
  Dim sSubMessage1 As String
  Dim sMessagePrefix As String
  Dim fTableOK As Boolean
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngExcludedTableID As Long
  Dim fDisconnectedElement As Boolean
  Dim fDoingDeleteCheck As Boolean
  Dim iIndex As Integer
  Dim sTemp As String
  
  fDoingDeleteCheck = Not IsMissing(pavarDisconnectedElements)

  sMessagePrefix = ValidateElement_MessagePrefix(pwfElement)
  lngExcludedTableID = 0

  ' Get the elements that precede the given element.
  ReDim aWFPrecedingElements(1)
  Set aWFPrecedingElements(UBound(aWFPrecedingElements)) = pwfElement
  PrecedingElements pwfElement, aWFPrecedingElements

  If fDoingDeleteCheck Then
    ' Remove the element we're trying to delete from the array of preceding elements, if we're doing the deletion check.
    iIndex = 1
    For iLoop = 1 To UBound(aWFPrecedingElements)
      fFound = False
      For iLoop2 = 1 To UBound(pavarDisconnectedElements, 2)
        If aWFPrecedingElements(iLoop) Is pavarDisconnectedElements(1, iLoop2) Then
          fFound = True
          Exit For
        End If
      Next iLoop2
      
      If Not fFound Then
        Set aWFPrecedingElements(iIndex) = aWFPrecedingElements(iLoop)
        iIndex = iIndex + 1
      End If
    Next iLoop
    ReDim Preserve aWFPrecedingElements(iIndex - 1)
  End If
  
  '------------------------------------------------------------
  ' 1. Email element must have a valid email To defined.
  ' 2. Email element must have have a valid email record.
  ' 3. Email element must have have a valid email record element identifier (where required).
  ' 4. Email element must have have a valid email record selector identifier (where required).
  
  ' 5. Email element items (DBValue) must have valid record.
  ' 6. Email element items (DBValue) must have valid record element identifier (where required).
  ' 7. Email element items (DBValue) must have valid record selector identifier (where required).
  ' 10. Email element items (DBValue) must have valid column.
  
  ' 8. Email element items (WFValue) must have valid WebForm identifier.
  ' 9. Email element items (WFValue) must have valid WebForm InputValue identifier.
  
  ' 11. Email element must have a valid email CC if one is defined.
  '------------------------------------------------------------
  fValid1 = (pwfElement.EmailID > 0)
  fValid2 = True
  fValid3 = True
  fValid4 = True
  fValid11 = True
    
  If fValid1 Then
    ' Email defined - does it still exist?
    With recEmailAddrEdit
      .Index = "idxID"
      .Seek "=", pwfElement.EmailID

      If .NoMatch Then
        fValid1 = False
      Else
        If !Deleted Then
          fValid1 = False
        Else
          lngTableID = !TableID
          iEmailType = !Type
        End If
      End If
    End With
  End If
              
  If fValid1 Then
    Select Case pwfElement.EmailRecord
      Case giWFRECSEL_UNKNOWN           ' Oops, something's wrong
        fValid2 = False
      
      Case giWFRECSEL_INITIATOR           ' Initiator's personnel table record
        ' Email must be 'fixed' or based on the Personnel table.
        fTableOK = (iEmailType = 0)
        If Not fTableOK Then
          ' Get an array of the valid table IDs (base table and it's ascendants)
          ReDim alngValidTables(0)
          TableAscendants mlngPersonnelTableID, alngValidTables
          
          For iLoop4 = 1 To UBound(alngValidTables)
            If alngValidTables(iLoop4) = lngTableID Then
              fTableOK = True
              Exit For
            End If
          Next iLoop4
        End If
        
        fValid1 = (miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL) _
          And fTableOK
        
      Case giWFRECSEL_TRIGGEREDRECORD
        fTableOK = (iEmailType = 0)
        If Not fTableOK Then
          ' Get an array of the valid table IDs (base table and it's ascendants)
          ReDim alngValidTables(0)
          TableAscendants mlngBaseTableID, alngValidTables
          
          For iLoop4 = 1 To UBound(alngValidTables)
            If alngValidTables(iLoop4) = lngTableID Then
              fTableOK = True
              Exit For
            End If
          Next iLoop4
        End If
        
        fValid1 = (miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) _
          And fTableOK

      Case giWFRECSEL_IDENTIFIEDRECORD    ' Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
        ' Check identification is valid.
        fValid3 = (Len(Trim(pwfElement.RecordSelectorWebFormIdentifier)) > 0)

        If fValid3 Then
          fValid3 = False
          
          For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
            Set wfTempElement = aWFPrecedingElements(iLoop)

            If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(pwfElement.RecordSelectorWebFormIdentifier)) Then
              
              If wfTempElement.ElementType = elem_WebForm Then
                fValid3 = True
                
                fValid4 = (Len(Trim(pwfElement.RecordSelectorIdentifier)) > 0)
                If fValid4 Then
                  fValid4 = False
                
                  asItems = wfTempElement.Items

                  For iLoop2 = 1 To UBound(asItems, 2)
                    If (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID) _
                      And UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(pwfElement.RecordSelectorIdentifier)) Then
                      
                      fTableOK = (iEmailType = 0)
                      If Not fTableOK Then
                        ' Get an array of the valid table IDs (base table and it's ascendants)
                        ReDim alngValidTables(0)
                        TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
                        
                        For iLoop4 = 1 To UBound(alngValidTables)
                          If alngValidTables(iLoop4) = lngTableID Then
                            fTableOK = True
                            Exit For
                          End If
                        Next iLoop4
                      End If
                      
                      fValid1 = fTableOK
                      fValid4 = True
                      Exit For
                    End If
                  Next iLoop2
                End If
                Exit For
              ElseIf wfTempElement.ElementType = elem_StoredData Then
                fValid3 = True
                
                fTableOK = (iEmailType = 0)
                If Not fTableOK Then
                  ' Get an array of the valid table IDs (base table and it's ascendants)
                  ReDim alngValidTables(0)
                  TableAscendants wfTempElement.DataTableID, alngValidTables
                  
                  'JPD 20061227
                  'If (wfTempElement.DataAction = DATAACTION_DELETE) Then
                  '  lngExcludedTableID = wfTempElement.DataTableID
                  'End If
                  
                  For iLoop4 = 1 To UBound(alngValidTables)
                    If (alngValidTables(iLoop4) = lngTableID) _
                      And (lngExcludedTableID <> lngTableID) Then
                      fTableOK = True
                      Exit For
                    End If
                  Next iLoop4
                End If
                
                fValid1 = fTableOK
                Exit For
              End If
            End If
            Set wfTempElement = Nothing
          Next iLoop
        End If
        
      Case giWFRECSEL_ALL                 ' Show all records from the table in a WebForm RecordSelector
        ' Not used for Email Record selection.
        fValid2 = False

      Case giWFRECSEL_UNIDENTIFIED        ' Used when StoredData Inserts into a top-level table
        ' Email must be 'fixed'.
        fValid1 = (iEmailType = 0)
        
    End Select
  End If

  If pwfElement.EmailCCID > 0 Then
    ' Email defined - does it still exist?
    With recEmailAddrEdit
      .Index = "idxID"
      .Seek "=", pwfElement.EmailCCID
  
      If .NoMatch Then
        fValid11 = False
      Else
        If !Deleted Then
          fValid11 = False
        Else
          lngTableID = !TableID
          iEmailType = !Type
        End If
      End If
    End With
  
    If fValid11 Then
      Select Case pwfElement.EmailRecord
        Case giWFRECSEL_UNKNOWN           ' Oops, something's wrong
          fValid11 = False
  
        Case giWFRECSEL_INITIATOR           ' Initiator's personnel table record
          ' Email must be 'fixed' or based on the Personnel table.
          fTableOK = (iEmailType = 0)
          If Not fTableOK Then
            ' Get an array of the valid table IDs (base table and it's ascendants)
            ReDim alngValidTables(0)
            TableAscendants mlngPersonnelTableID, alngValidTables
  
            For iLoop4 = 1 To UBound(alngValidTables)
              If alngValidTables(iLoop4) = lngTableID Then
                fTableOK = True
                Exit For
              End If
            Next iLoop4
          End If
  
          fValid11 = (miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL) _
            And fTableOK
  
        Case giWFRECSEL_TRIGGEREDRECORD
          fTableOK = (iEmailType = 0)
          If Not fTableOK Then
            ' Get an array of the valid table IDs (base table and it's ascendants)
            ReDim alngValidTables(0)
            TableAscendants mlngBaseTableID, alngValidTables
  
            For iLoop4 = 1 To UBound(alngValidTables)
              If alngValidTables(iLoop4) = lngTableID Then
                fTableOK = True
                Exit For
              End If
            Next iLoop4
          End If
  
          fValid11 = (miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) _
            And fTableOK
  
        Case giWFRECSEL_IDENTIFIEDRECORD    ' Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
          ' Check identification is valid.
          fValid11 = (Len(Trim(pwfElement.RecordSelectorWebFormIdentifier)) > 0)
  
          If fValid11 Then
            fValid11 = False
  
            For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
              Set wfTempElement = aWFPrecedingElements(iLoop)
  
              If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(pwfElement.RecordSelectorWebFormIdentifier)) Then
  
                If wfTempElement.ElementType = elem_WebForm Then
                  fValid11 = (Len(Trim(pwfElement.RecordSelectorIdentifier)) > 0)
                  If fValid11 Then
                    fValid11 = False
  
                    asItems = wfTempElement.Items
  
                    For iLoop2 = 1 To UBound(asItems, 2)
                      If (asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID) _
                        And UCase(Trim(asItems(9, iLoop2))) = UCase(Trim(pwfElement.RecordSelectorIdentifier)) Then
  
                        fTableOK = (iEmailType = 0)
                        If Not fTableOK Then
                          ' Get an array of the valid table IDs (base table and it's ascendants)
                          ReDim alngValidTables(0)
                          TableAscendants CLng(asItems(44, iLoop2)), alngValidTables
  
                          For iLoop4 = 1 To UBound(alngValidTables)
                            If alngValidTables(iLoop4) = lngTableID Then
                              fTableOK = True
                              Exit For
                            End If
                          Next iLoop4
                        End If
  
                        fValid11 = fTableOK
                        Exit For
                      End If
                    Next iLoop2
                  End If
                  Exit For
                ElseIf wfTempElement.ElementType = elem_StoredData Then
                  fValid11 = True
  
                  fTableOK = (iEmailType = 0)
                  If Not fTableOK Then
                    ' Get an array of the valid table IDs (base table and it's ascendants)
                    ReDim alngValidTables(0)
                    TableAscendants wfTempElement.DataTableID, alngValidTables
  
                    For iLoop4 = 1 To UBound(alngValidTables)
                      If (alngValidTables(iLoop4) = lngTableID) _
                        And (lngExcludedTableID <> lngTableID) Then
                        fTableOK = True
                        Exit For
                      End If
                    Next iLoop4
                  End If
  
                  fValid11 = fTableOK
                  Exit For
                End If
              End If
              Set wfTempElement = Nothing
            Next iLoop
          End If
  
        Case giWFRECSEL_ALL                 ' Show all records from the table in a WebForm RecordSelector
          ' Not used for Email Record selection.
          fValid11 = False
  
        Case giWFRECSEL_UNIDENTIFIED        ' Used when StoredData Inserts into a top-level table
          ' Email must be 'fixed'.
          fValid11 = (iEmailType = 0)
          
      End Select
    End If
  End If

  ' ------------------------------------
  ' Validate the email items.
  ' ------------------------------------
  ' 5. Email element items (DBValue) must have valid record.
  ' 6. Email element items (DBValue) must have valid record element identifier (where required).
  ' 7. Email element items (DBValue) must have valid record selector identifier (where required).
  ' 10. Email element items (DBValue) must have valid column.
  
  ' 8. Email element items (WFValue) must have valid WebForm identifier.
  ' 9. Email element items (WFValue) must have valid WebForm InputValue identifier.
  asElementItems = pwfElement.Items
  For iLoop2 = 1 To UBound(asElementItems, 2)
    sSubMessage1 = ""
    
    If (asElementItems(2, iLoop2) = giWFEMAILITEM_DBVALUE) Then
      ' 5. Email element items (DBValue) must have valid record.
      ' 6. Email element items (DBValue) must have valid record element identifier (where required).
      ' 7. Email element items (DBValue) must have valid record selector identifier (where required).
      ' 10. Email element items (DBValue) must have valid column.
      fValid5 = True
      fValid6 = True
      fValid7 = True
      fValid10 = (CLng(asElementItems(4, iLoop2)) > 0)
      If fValid10 Then
        ' Column defined - does it still exist?
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", CLng(asElementItems(4, iLoop2))
  
          If .NoMatch Then
            fValid10 = False
          Else
            If !Deleted Then
              fValid10 = False
            Else
              lngTableID = !TableID
            End If
          End If
        End With
      End If
            
      If fValid10 Then
        sSubMessage1 = " (" & GetColumnName(CLng(asElementItems(4, iLoop2))) & ")"
        
        Select Case asElementItems(5, iLoop2)
          Case giWFRECSEL_UNKNOWN           ' Oops, something's wrong
            fValid5 = False
    
          Case giWFRECSEL_INITIATOR           ' Initiator's personnel table record
            ' DBValue must be based on the Personnel table.
            ReDim alngValidTables(0)
            TableAscendants mlngPersonnelTableID, alngValidTables
            
            fFound = False
            For iLoop4 = 1 To UBound(alngValidTables)
              If alngValidTables(iLoop4) = lngTableID Then
                fFound = True
                Exit For
              End If
            Next iLoop4
      
            fValid5 = fFound
    
          Case giWFRECSEL_TRIGGEREDRECORD     ' Base table record
            ' DBValue must be based on the Base table.
            ReDim alngValidTables(0)
            TableAscendants mlngBaseTableID, alngValidTables
            
            fFound = False
            For iLoop4 = 1 To UBound(alngValidTables)
              If alngValidTables(iLoop4) = lngTableID Then
                fFound = True
                Exit For
              End If
            Next iLoop4
            
            fValid5 = fFound
    
          Case giWFRECSEL_IDENTIFIEDRECORD    ' Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
            ' Check identification is valid.
            fValid6 = (Len(Trim(asElementItems(13, iLoop2))) > 0)
            
            If fValid6 Then
              fValid6 = False
              
              For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
                Set wfTempElement = aWFPrecedingElements(iLoop)
    
                If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(asElementItems(13, iLoop2))) Then
                  If wfTempElement.ElementType = elem_WebForm Then
                    fValid6 = True
                    fValid7 = (Len(Trim(asElementItems(14, iLoop2))) > 0)
                    If fValid7 Then
                      fValid7 = False
    
                      asItems = wfTempElement.Items
    
                      For iLoop3 = 1 To UBound(asItems, 2)
                        If (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_GRID) _
                          And UCase(Trim(asItems(9, iLoop3))) = UCase(Trim(asElementItems(14, iLoop2))) Then
                                  
                          ReDim alngValidTables(0)
                          TableAscendants CLng(asItems(44, iLoop3)), alngValidTables
                          
                          fFound = False
                          For iLoop4 = 1 To UBound(alngValidTables)
                            If alngValidTables(iLoop4) = lngTableID Then
                              fFound = True
                              Exit For
                            End If
                          Next iLoop4

                          fValid7 = fFound
                          Exit For
                        End If
                      Next iLoop3
                    End If
                    Exit For
                  ElseIf wfTempElement.ElementType = elem_StoredData Then
                    ReDim alngValidTables(0)
                    TableAscendants wfTempElement.DataTableID, alngValidTables
                    
                    'JPD 20061227
                    'If wfTempElement.DataAction = DATAACTION_DELETE Then
                    '  ' Cannot do anything with a Deleted record, but can use its ascendants.
                    '  ' Remove the table itself from the array of valid tables.
                    '  alngValidTables(1) = 0
                    'End If
                    
                    fFound = False
                    For iLoop4 = 1 To UBound(alngValidTables)
                      If alngValidTables(iLoop4) = lngTableID Then
                        fFound = True
                        Exit For
                      End If
                    Next iLoop4
            
                    fValid6 = fFound
                    Exit For
                  End If
                End If
                Set wfTempElement = Nothing
              Next iLoop
            End If
    
          Case giWFRECSEL_ALL                 ' Show all records from the table in a WebForm RecordSelector
            ' Not used for DBValue selection.
            fValid5 = False
    
          Case giWFRECSEL_UNIDENTIFIED        ' Used when StoredData Inserts into a top-level table
            ' Not used for DBValue selection.
            fValid5 = False
        End Select
      End If
      
      '------------------------------------------------------------
      ' Add the required validation messages to the array.
      '------------------------------------------------------------
      ' 5. Email element items (DBValue) must have valid record.
      ' 6. Email element items (DBValue) must have valid record element identifier (where required).
      ' 7. Email element items (DBValue) must have valid record selector identifier (where required).
      ' 10. Email element items (DBValue) must have valid column.
      If (Not fValid10) And (Not fDoingDeleteCheck) Then
        asElementItems(4, iLoop2) = ""
  
        ValidateWorkflow_AddMessage _
          sMessagePrefix & "Database Value - Invalid column", _
          pwfElement.ControlIndex
      End If
      If Not fValid5 Then
        ValidateWorkflow_AddMessage _
          sMessagePrefix & "Database Value" & sSubMessage1 & " - Invalid record", _
          pwfElement.ControlIndex
      End If
      If Not fValid6 Then
        asElementItems(11, iLoop2) = ""
        asElementItems(12, iLoop2) = ""

        ValidateWorkflow_AddMessage _
          sMessagePrefix & "Database Value" & sSubMessage1 & " - Invalid element identifier", _
          pwfElement.ControlIndex
      End If
      If Not fValid7 Then
        asElementItems(12, iLoop2) = ""
        
        ValidateWorkflow_AddMessage _
          sMessagePrefix & "Database Value" & sSubMessage1 & " - Invalid record selector", _
          pwfElement.ControlIndex
      End If
    
    ElseIf (asElementItems(2, iLoop2) = giWFEMAILITEM_WFVALUE) Then
      ' 8. Email element items (WFValue) must have valid WebForm identifier.
      ' 9. Email element items (WFValue) must have valid WebForm InputValue identifier.
      fValid8 = (Len(Trim(asElementItems(11, iLoop2))) > 0)
      fValid9 = True
      
      If fValid8 Then
        sSubMessage1 = " (" & asElementItems(11, iLoop2) & ")"
        
        fValid9 = (Len(Trim(asElementItems(12, iLoop2))) > 0)
      
        If fValid9 Then
          sSubMessage1 = " (" & asElementItems(11, iLoop2) & "." & asElementItems(12, iLoop2) & ")"
          
          fValid9 = False
          
          For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
            Set wfTempElement = aWFPrecedingElements(iLoop)
    
            If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(asElementItems(11, iLoop2))) Then
              If wfTempElement.ElementType = elem_WebForm Then
              
                asItems = wfTempElement.Items
  
                For iLoop3 = 1 To UBound(asItems, 2)
                  If ((asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_CHAR) _
                      Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                      Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                      Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_DATE) _
                      Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                      Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                      Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                      Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
                    And UCase(Trim(asItems(9, iLoop3))) = UCase(Trim(asElementItems(12, iLoop2))) Then
  
                    fValid9 = True
                    Exit For
                  End If
                Next iLoop3
              End If
              
              Exit For
            End If
            Set wfTempElement = Nothing
          Next iLoop
        End If
      End If
      
      '------------------------------------------------------------
      ' Add the required validation messages to the array.
      '------------------------------------------------------------
      ' 8. Email element items (WFValue) must have valid WebForm identifier.
      ' 9. Email element items (WFValue) must have valid WebForm InputValue identifier.
      If Not fValid8 Then
        asElementItems(11, iLoop2) = ""
        asElementItems(12, iLoop2) = ""
        
        ValidateWorkflow_AddMessage _
          sMessagePrefix & "Workflow Value" & sSubMessage1 & " - Invalid web form identifier", _
          pwfElement.ControlIndex
      End If
      If Not fValid9 Then
        asElementItems(12, iLoop2) = ""
        
        ValidateWorkflow_AddMessage _
          sMessagePrefix & "Workflow Value" & sSubMessage1 & " - Invalid value identifier", _
          pwfElement.ControlIndex
      End If
    ElseIf (asElementItems(2, iLoop2) = giWFEMAILITEM_CALCULATION) Then
      If (asElementItems(56, iLoop2) > 0) Then
        sTemp = GetExpressionName(CLng(asElementItems(56, iLoop2)))
        If Len(Trim(sTemp)) = 0 Then
          sTemp = "<unknown>"
        Else
          sTemp = "<" & sTemp & ">"
        End If
        sTemp = "Calculation - " & sTemp
        
        ValidateElement_Expression _
          pwfElement, _
          CLng(asElementItems(56, iLoop2)), _
          "'" & sTemp & "' - invalid calculation", _
          pavarDisconnectedElements
      Else
        
        ValidateWorkflow_AddMessage _
          sMessagePrefix & "No calculation selected", _
          pwfElement.ControlIndex
      End If
    End If
  Next iLoop2

  '------------------------------------------------------------
  ' Add the required validation messages to the array.
  '------------------------------------------------------------
  ' 1. Email element must have a valid email defined.
  ' 2. Email element must have have a valid email record.
  ' 3. Email element must have have a valid email record element identifier (where required).
  ' 11. Email element must have a valid email CC if one is defined.
  If (Not fValid1) And (Not fDoingDeleteCheck) Then
    pwfElement.EmailID = 0
    
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "No Email To address selected", _
      pwfElement.ControlIndex
  End If
  If (Not fValid11) And (Not fDoingDeleteCheck) Then
    pwfElement.EmailCCID = 0
    
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid Email Copy address selected", _
      pwfElement.ControlIndex
  End If
  If (Not fValid2) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid email record", _
      pwfElement.ControlIndex
  End If
  If Not fValid3 Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid email record element", _
      pwfElement.ControlIndex
  End If
  If Not fValid4 Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid email record selector", _
      pwfElement.ControlIndex
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
End Sub


Private Sub ValidateElement_WebForm(pwfElement As VB.Control, _
  Optional pavarDisconnectedElements As Variant)
  
  On Error GoTo ErrorTrap
  
  Dim wfTempElement As VB.Control
  Dim wfPrecedingElement As VB.Control
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim aWFImmediatelyPrecedingElements() As VB.Control
  Dim aWFPrecedingElements() As VB.Control
  Dim aWFImmediatelySucceedingElements_timeout() As VB.Control
  Dim fValid1 As Boolean
  Dim fValid2 As Boolean
  Dim fValid3 As Boolean
  Dim fValid4 As Boolean
  Dim fValid5 As Boolean
  Dim fValid6 As Boolean
  Dim fValid7 As Boolean
  Dim fValid8 As Boolean
  Dim fValid9 As Boolean
  Dim fValid10 As Boolean
  Dim fValid11 As Boolean
  Dim fValid12 As Boolean
  Dim fValid13 As Boolean
  Dim fValid15 As Boolean
  Dim fValid16 As Boolean
  Dim fValid17 As Boolean
  Dim fValid18 As Boolean
  Dim fValid19 As Boolean
  Dim fValid20 As Boolean
  Dim fValid21 As Boolean
  Dim fValid22 As Boolean
  Dim fValid23 As Boolean
  Dim fValid24 As Boolean
  Dim fValid25 As Boolean
  Dim asItems() As String
  Dim asElementItems() As String
  Dim asValidations() As String
  Dim lngTableID As Long
  Dim sSubMessage1 As String
  Dim sMessagePrefix As String
  Dim wfLink As COAWF_Link
  Dim fLinkOK As Boolean
  Dim iTimeoutLinkCount As Integer
  Dim objMisc As Misc
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngLoop As Long
  Dim iValidPrecedingElements As Integer
  Dim fDisconnectedElement As Boolean
  Dim fDoingDeleteCheck As Boolean
  
  fDoingDeleteCheck = Not IsMissing(pavarDisconnectedElements)
  
  sMessagePrefix = ValidateElement_MessagePrefix(pwfElement)

  ' Get the elements that precede the given element.
  ReDim aWFPrecedingElements(1)
  Set aWFPrecedingElements(UBound(aWFPrecedingElements)) = pwfElement
  PrecedingElements pwfElement, aWFPrecedingElements

  ' Get the element that immediately preceeds the given element.
  ReDim aWFImmediatelyPrecedingElements(1)
  Set aWFImmediatelyPrecedingElements(UBound(aWFImmediatelyPrecedingElements)) = pwfElement
  ImmediatelyPrecedingElements pwfElement, aWFImmediatelyPrecedingElements

  If fDoingDeleteCheck Then
    ' Remove the element we're trying to delete from the array of preceding elements, if we're doing the deletion check.
    iIndex = 1
    For iLoop = 1 To UBound(aWFPrecedingElements)
      fFound = False
      For iLoop2 = 1 To UBound(pavarDisconnectedElements, 2)
        If aWFPrecedingElements(iLoop) Is pavarDisconnectedElements(1, iLoop2) Then
          fFound = True
          Exit For
        End If
      Next iLoop2
      
      If Not fFound Then
        Set aWFPrecedingElements(iIndex) = aWFPrecedingElements(iLoop)
        iIndex = iIndex + 1
      End If
    Next iLoop
    ReDim Preserve aWFPrecedingElements(iIndex - 1)
  
    ' Remove the element we're trying to delete from the array of immediately preceding elements, if we're doing the deletion check.
    iIndex = 1
    For iLoop = 1 To UBound(aWFImmediatelyPrecedingElements)
      fFound = False
      For iLoop2 = 1 To UBound(pavarDisconnectedElements, 2)
        If aWFImmediatelyPrecedingElements(iLoop) Is pavarDisconnectedElements(1, iLoop2) Then
          fFound = True
          Exit For
        End If
      Next iLoop2
      
      If Not fFound Then
        Set aWFImmediatelyPrecedingElements(iIndex) = aWFImmediatelyPrecedingElements(iLoop)
        iIndex = iIndex + 1
      End If
    Next iLoop
    ReDim Preserve aWFImmediatelyPrecedingElements(iIndex - 1)
  End If

  ' 1. WebForm element must have an identifier.
  ' 2. WebForm element must have a unique identifier.
  ' 3. WebForm element must be actionable to someone.
  '    ie. immediately follow a Begin, Email, or other WebForm element.
  ' 4. WebForm element must have at least 1 Submit button.
  ' 5. WebForm element items (DBValue/RecSel) must have valid record.
  ' 6. WebForm element items (DBValue/RecSel) must have valid record element identifier (where required).
  ' 7. WebForm element items (DBValue/RecSel) must have valid record selector identifier (where required).
  ' 8. WebForm element items (WFValue) must have valid WebForm identifier.
  ' 9. WebForm element items (WFValue) must have valid WebForm InputValue identifier.
  ' 10. WebForm element items (DBValue/RecSel/Lookup) must have valid table/column.
  ' 11. WebForm element items (Input/RecSel/Button) must have identifier.
  ' 12. WebForm element items (Input/RecSel/Button) must have unique identifier.
  ' 13. WebForm element items (Image) must have a valid picture.
  
  'JPD 20060719 Fault 11334 - Check 14 no longer required
  ' 14. WebForm element must have at most 2 buttons.
  
  ' 15. WebForm element must have a Timeout link defined if a Timeout Frequency has been defined.
  ' 16. WebForm element must have a Timeout Frequency defined if a Timeout link has been defined.
  ' 17. WebForm element items (Dropdown/OptionGroup) must have ControValues defined.
  
  ' 19. WebForm element items (Input) must have valid default defined wrt size or basic validity
  ' 20. WebForm element items (Input - numeric) must have valid default defined wrt decimals.
  ' 21. WebForm element cannot follow another WebForm element from the 'timeout' flow.
  ' 22. WebForm element items (Input - numeric/character) must have valid size.
  ' 23. WebForm element must have a hypertext section of a custom Completion message.
  ' 24. WebForm element must have a hypertext section of a custom SavedForLater message.
  ' 25. WebForm element must have a hypertext section of a custom FollowOnForms message.
  fValid1 = (Len(Trim(pwfElement.Identifier)) > 0)
  fValid2 = True
  fValid3 = False
  fValid4 = False
  fValid21 = True
  
  If fValid1 Then
    fValid2 = UniqueIdentifier(pwfElement.Identifier, pwfElement.ControlIndex)
  End If

  fValid3 = True
  iValidPrecedingElements = 0
  For iLoop = 2 To UBound(aWFImmediatelyPrecedingElements) ' Ignore index 1 as that is the current element
    Set wfPrecedingElement = aWFImmediatelyPrecedingElements(iLoop)

    If (wfPrecedingElement.ElementType <> elem_Connector1) _
      And (wfPrecedingElement.ElementType <> elem_Connector2) _
      And (wfPrecedingElement.ElementType <> elem_Or) _
      And (wfPrecedingElement.ElementType <> elem_Decision) Then

      If (wfPrecedingElement.ElementType = elem_WebForm) _
        Or ((wfPrecedingElement.ElementType = elem_Begin) _
          And ((miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL) Or (miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL))) _
        Or (wfPrecedingElement.ElementType = elem_Email) Then

        iValidPrecedingElements = 1
        
        If wfPrecedingElement.ElementType = elem_WebForm Then

          ReDim aWFImmediatelySucceedingElements_timeout(1)
          Set aWFImmediatelySucceedingElements_timeout(UBound(aWFImmediatelySucceedingElements_timeout)) = wfPrecedingElement
          ImmediatelySucceedingElements wfPrecedingElement, _
            aWFImmediatelySucceedingElements_timeout, _
            True, _
            webFormOutFlow_Timeout

          For iLoop3 = 2 To UBound(aWFImmediatelySucceedingElements_timeout) ' Ignore index 1
            If aWFImmediatelySucceedingElements_timeout(iLoop3) Is pwfElement Then
              fValid21 = False
              Exit For
            End If
          Next iLoop3
        End If
      
      ElseIf (wfPrecedingElement.ElementType = elem_StoredData) Then
        If (glngSQLVersion <= 8) Then
          fValid3 = False
        End If
      
      Else
        fValid3 = False
      End If
    End If
    
    Set wfPrecedingElement = Nothing
  Next iLoop
  
  If iValidPrecedingElements = 0 Then
    fValid3 = False
  End If
   
  ' 15. WebForm element must have a Timeout link defined if a Timeout Frequency has been defined.
  ' 16. WebForm element must have a Timeout Frequency defined if a Timeout link has been defined.
  iTimeoutLinkCount = 0

  For Each wfLink In ASRWFLink1
    fLinkOK = wfLink.Visible
    If (Not fLinkOK) Then
      ' Link might not be .visible but still valid
      ' if this method is called from the Workflow properties screen.
      fLinkOK = True

      If (miLastActionFlag = giACTION_DELETECONTROLS) Then
        For iLoop = 1 To UBound(mactlUndoControls)
          If TypeOf mactlUndoControls(iLoop) Is COAWF_Link Then
            If mactlUndoControls(iLoop).Index = wfLink.Index Then
              fLinkOK = False
              Exit For
            End If
          End If
        Next iLoop
      End If

      If fLinkOK Then
        If (miLastActionFlag = giACTION_SWAPCONTROL) Then
          If UBound(mactlUndoControls) >= 1 Then
            If TypeOf mactlUndoControls(1) Is COAWF_Link Then
              If mactlUndoControls(1).Index = wfLink.Index Then
                fLinkOK = False
              End If
            End If
          End If
        End If
      End If
      
      ' JPD 20060719 Fault 11339
      If fLinkOK Then
        For iLoop = 1 To UBound(mactlClipboardControls)
          If TypeOf mactlClipboardControls(iLoop) Is COAWF_Link Then
            If mactlClipboardControls(iLoop).Index = wfLink.Index Then
              fLinkOK = False
              Exit For
            End If
          End If
        Next iLoop
      End If
    End If

    If fLinkOK Then
      If wfLink.StartElementIndex = pwfElement.ControlIndex Then
        If (wfLink.StartOutboundFlowCode = 1) Then

          iTimeoutLinkCount = iTimeoutLinkCount + 1
          Exit For
        End If
      End If
    End If
  Next wfLink
  Set wfLink = Nothing

  fValid15 = (pwfElement.WebFormTimeoutFrequency = 0) Or (iTimeoutLinkCount > 0)
  fValid16 = (pwfElement.WebFormTimeoutFrequency > 0) Or (iTimeoutLinkCount = 0)

  ' Validate the description calculation (if required)
  If (pwfElement.DescriptionExprID > 0) Then
    ValidateElement_Expression _
      pwfElement, _
      pwfElement.DescriptionExprID, _
      "Invalid description calculation", _
      pavarDisconnectedElements
  End If

  ' 23. WebForm element must have a hypertext section of a custom Completion message.
  ' 24. WebForm element must have a hypertext section of a custom SavedForLater message.
  ' 25. WebForm element must have a hypertext section of a custom FollowOnForms message.
  fValid23 = True
  If (pwfElement.WFCompletionMessageType = MESSAGE_CUSTOM) Then
    fValid23 = (InStr(Replace(pwfElement.WFCompletionMessage, "\\", ""), "\ul ") > 0)
  End If
  fValid24 = True
  If (pwfElement.WFSavedForLaterMessageType = MESSAGE_CUSTOM) Then
    fValid24 = (InStr(Replace(pwfElement.WFSavedForLaterMessage, "\\", ""), "\ul ") > 0)
  End If
  fValid25 = True
  If (pwfElement.WFFollowOnFormsMessageType = MESSAGE_CUSTOM) Then
    fValid25 = (InStr(Replace(pwfElement.WFFollowOnFormsMessage, "\\", ""), "\ul ") > 0)
  End If
  
  '------------------------------------------------------------
  ' Add the required validation messages to the array.
  '------------------------------------------------------------
  ' 1. WebForm element must have an identifier.
  ' 2. WebForm element must have a unique identifier.
  ' 3. WebForm element must be actionable to someone.
  '    ie. immediately follow a Begin, Email, or other WebForm element.
  ' 15. WebForm element must have a Timeout link defined if a Timeout Frequency has been defined.
  ' 16. WebForm element must have a Timeout Frequency defined if a Timeout link has been defined.
  ' 21. WebForm element cannot follow another WebForm element from the 'timeout' flow.
  ' 23. WebForm element must have a hypertext section of a custom Completion message.
  ' 24. WebForm element must have a hypertext section of a custom SavedForLater message.
  ' 25. WebForm element must have a hypertext section of a custom FollowOnForms message.
  If (Not fValid1) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "No identifier", _
      pwfElement.ControlIndex
  End If
  If (Not fValid2) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Non-unique identifier", _
      pwfElement.ControlIndex
  End If
  If Not fValid3 Then
    If (miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL) _
      Or (miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL) Then
      ValidateWorkflow_AddMessage _
        sMessagePrefix & "Not actionable to anyone (ie. does not follow a Begin, Email or Web Form element)", _
        pwfElement.ControlIndex
    Else
      ValidateWorkflow_AddMessage _
        sMessagePrefix & "Not actionable to anyone (ie. does not follow an Email or Web Form element)", _
        pwfElement.ControlIndex
    End If
  End If
  If (Not fValid15) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Timeout Frequency defined, but no 'Timeout' outbound flow", _
      pwfElement.ControlIndex
  End If
  If (Not fValid16) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "'Timeout' outbound flow defined, but no Timeout Frequency", _
      pwfElement.ControlIndex
  End If
  If (Not fValid21) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Cannot succeed a web form element via the 'Timeout' outbound flow", _
      pwfElement.ControlIndex
  End If
  If (Not fValid23) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Custom 'Completion' message must contain a hypertext section.", _
      pwfElement.ControlIndex
  End If
  If (Not fValid24) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Custom 'Saved For Later' message must contain a hypertext section.", _
      pwfElement.ControlIndex
  End If
  If (Not fValid25) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Custom 'Follow On Forms' message must contain a hypertext section.", _
      pwfElement.ControlIndex
  End If
  
  ' ------------------------------------
  ' Validate the WebForm items
  ' ------------------------------------
  asElementItems = pwfElement.Items
  
  For iLoop2 = 1 To UBound(asElementItems, 2)
    sSubMessage1 = ""

    Select Case CInt(asElementItems(2, iLoop2))
      '------------------------------------------------------------
      Case giWFFORMITEM_UNKNOWN
        ' No validation required.
        
      '------------------------------------------------------------
      Case giWFFORMITEM_BUTTON
        ' There is at least 1 button. Good.
        ' No validation required on the button itself.
        If (CInt(asElementItems(54, iLoop2)) = WORKFLOWBUTTONACTION_SUBMIT) _
          Or (CInt(asElementItems(54, iLoop2)) = WORKFLOWBUTTONACTION_CANCEL) Then
          
          fValid4 = True
        End If

        ' Item must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then
              
              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If

        If Not fValid11 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Button - no identifier", _
            pwfElement.ControlIndex
        End If
        If Not fValid12 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Button (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If

      '------------------------------------------------------------
      Case giWFFORMITEM_DBVALUE
        ' 5. WebForm element items (DBValue) must have valid record.
        ' 6. WebForm element items (DBValue) must have valid record element identifier (where required).
        ' 7. WebForm element items (DBValue) must have valid record selector identifier (where required).
        ' 10. WebForm element items (DBValue) must have valid column.
        fValid5 = True
        fValid6 = True
        fValid7 = True
        fValid10 = (CLng(asElementItems(4, iLoop2)) > 0)
        If fValid10 Then
          ' Column defined - does it still exist?
          With recColEdit
            .Index = "idxColumnID"
            .Seek "=", CLng(asElementItems(4, iLoop2))
  
            If .NoMatch Then
              fValid10 = False
            Else
              If !Deleted Then
                fValid10 = False
              Else
                lngTableID = !TableID
              End If
            End If
          End With
        End If

        If fValid10 Then
          sSubMessage1 = " (" & GetColumnName(CLng(asElementItems(4, iLoop2))) & ")"
  
          Select Case asElementItems(5, iLoop2)
            Case giWFRECSEL_UNKNOWN           ' Oops, something's wrong
              fValid5 = False

            Case giWFRECSEL_INITIATOR           ' Initiator's personnel table record
              ' DBValue must be based on the Personnel table.
              ReDim alngValidTables(0)
              TableAscendants mlngPersonnelTableID, alngValidTables
                            
              fFound = False
              For lngLoop = 1 To UBound(alngValidTables)
                If lngTableID = alngValidTables(lngLoop) Then
                  fFound = True
                  Exit For
                End If
              Next lngLoop
              
              fValid5 = fFound
    
            Case giWFRECSEL_TRIGGEREDRECORD     ' Base table record
              ' DBValue must be based on the base table.
              ReDim alngValidTables(0)
              TableAscendants mlngBaseTableID, alngValidTables
                            
              fFound = False
              For lngLoop = 1 To UBound(alngValidTables)
                If lngTableID = alngValidTables(lngLoop) Then
                  fFound = True
                  Exit For
                End If
              Next lngLoop
              
              fValid5 = fFound
  
            Case giWFRECSEL_IDENTIFIEDRECORD    ' Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
              ' Check identification is valid.
              fValid6 = (Len(Trim(asElementItems(11, iLoop2))) > 0)
  
              If fValid6 Then
                fValid6 = False
                
                For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
                  Set wfTempElement = aWFPrecedingElements(iLoop)

                  If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(asElementItems(11, iLoop2))) Then
                    If wfTempElement.ElementType = elem_WebForm Then
                      fValid6 = True
  
                      fValid7 = (Len(Trim(asElementItems(12, iLoop2))) > 0)
                      If fValid7 Then
                        fValid7 = False
  
                        asItems = wfTempElement.Items
  
                        For iLoop3 = 1 To UBound(asItems, 2)
                          If (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_GRID) _
                            And UCase(Trim(asItems(9, iLoop3))) = UCase(Trim(asElementItems(12, iLoop2))) Then
  
                            ReDim alngValidTables(0)
                            TableAscendants CLng(asItems(44, iLoop3)), alngValidTables
                                          
                            fFound = False
                            For lngLoop = 1 To UBound(alngValidTables)
                              If lngTableID = alngValidTables(lngLoop) Then
                                fFound = True
                                Exit For
                              End If
                            Next lngLoop
                            fValid7 = fFound
                            Exit For
                          End If
                        Next iLoop3
                      End If
                      Exit For
                    ElseIf wfTempElement.ElementType = elem_StoredData Then
                      ReDim alngValidTables(0)
                      TableAscendants wfTempElement.DataTableID, alngValidTables
                                    
                      'JPD 20061227
                      'If wfTempElement.DataAction = DATAACTION_DELETE Then
                      '  ' Cannot do anything with a Deleted record, but can use its ascendants.
                      '  ' Remove the table itself from the array of valid tables.
                      '  alngValidTables(1) = 0
                      'End If
                                    
                      fFound = False
                      For lngLoop = 1 To UBound(alngValidTables)
                        If lngTableID = alngValidTables(lngLoop) Then
                          fFound = True
                          Exit For
                        End If
                      Next lngLoop
                      fValid6 = fFound
                      Exit For
                    End If
                  End If
                  Set wfTempElement = Nothing
                Next iLoop
              End If
  
            Case giWFRECSEL_ALL                 ' Show all records from the table in a WebForm RecordSelector
              ' Not used for DBValue selection.
              fValid5 = False
  
            Case giWFRECSEL_UNIDENTIFIED        ' Used when StoredData Inserts into a top-level table
              ' Not used for DBValue selection.
              fValid5 = False
          End Select
        End If

        '------------------------------------------------------------
        ' Add the required validation messages to the array.
        '------------------------------------------------------------
        ' 5. WebForm element items (DBValue) must have valid record.
        ' 6. WebForm element items (DBValue) must have valid record element identifier (where required).
        ' 7. WebForm element items (DBValue) must have valid record selector identifier (where required).
        ' 10. WebForm element items (DBValue) must have valid column.
        If (Not fValid10) And (Not fDoingDeleteCheck) Then
          asElementItems(4, iLoop2) = ""
  
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Database Value - Invalid column", _
            pwfElement.ControlIndex
        End If
        If (Not fValid5) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Database Value" & sSubMessage1 & " - Invalid record", _
            pwfElement.ControlIndex
        End If
        If Not fValid6 Then
          asElementItems(11, iLoop2) = ""
          asElementItems(12, iLoop2) = ""
  
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Database Value" & sSubMessage1 & " - Invalid element identifier", _
            pwfElement.ControlIndex
        End If
        If Not fValid7 Then
          asElementItems(12, iLoop2) = ""
  
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Database Value" & sSubMessage1 & " - Invalid record selector", _
            pwfElement.ControlIndex
        End If
      
      '------------------------------------------------------------
      Case giWFFORMITEM_LABEL
        ' Validate the label calculation (if required)
        If CInt(asElementItems(57, iLoop2)) = giWFDATAVALUE_CALC Then
          If (CLng(asElementItems(56, iLoop2)) > 0) Then
            ValidateElement_Expression _
              pwfElement, _
              CLng(asElementItems(56, iLoop2)), _
              "Invalid label calculation", _
              pavarDisconnectedElements
          Else
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Label - <Calculated> - No calculation selected", _
              pwfElement.ControlIndex
          End If
        End If
        
      '------------------------------------------------------------
      Case giWFFORMITEM_INPUTVALUE_CHAR
        ' InputChar must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        fValid19 = (CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC) _
          Or (Len(asElementItems(10, iLoop2)) <= CInt(asElementItems(7, iLoop2)))
        
        If CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC Then
          If (CLng(asElementItems(56, iLoop2)) > 0) Then
            ValidateElement_Expression _
              pwfElement, _
              CLng(asElementItems(56, iLoop2)), _
              "Character Input (" & asElementItems(9, iLoop2) & " - Invalid default value calculation", _
              pavarDisconnectedElements
          Else
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Character Input (" & asElementItems(9, iLoop2) & ") - No default value calculation selected", _
              pwfElement.ControlIndex
          End If
        End If
        
        fValid22 = (CInt(asElementItems(7, iLoop2)) <= WORKFLOWWEBFORM_MAXSIZE_CHARINPUT)

        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then

              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If

        If (Not fValid11) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Character Input - no identifier", _
            pwfElement.ControlIndex
        End If
        If (Not fValid12) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Character Input (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If
        If (Not fValid19) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Character Input (" & asElementItems(9, iLoop2) & ") - default value too big", _
            pwfElement.ControlIndex
        End If
        If (Not fValid22) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Character Input (" & asElementItems(9, iLoop2) & ") - size exceeds maximum (" & CStr(WORKFLOWWEBFORM_MAXSIZE_CHARINPUT) & ")", _
            pwfElement.ControlIndex
        End If
      
      '------------------------------------------------------------
      Case giWFFORMITEM_WFVALUE, _
        giWFFORMITEM_WFFILE
        
        ' 8. WebForm element items (WFValue) must have valid WebForm identifier.
        ' 9. WebForm element items (WFValue) must have valid WebForm InputValue identifier.
        fValid8 = (Len(Trim(asElementItems(11, iLoop2))) > 0)
        fValid9 = True

        If fValid8 Then
          fValid8 = False
          sSubMessage1 = " (" & asElementItems(11, iLoop2) & ")"
  
          For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
            Set wfTempElement = aWFPrecedingElements(iLoop)

            If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(asElementItems(11, iLoop2))) Then
              If wfTempElement.ElementType = elem_WebForm Then
                fValid8 = True

                fValid9 = (Len(Trim(asElementItems(12, iLoop2))) > 0)

                If fValid9 Then
                  sSubMessage1 = " (" & asElementItems(11, iLoop2) & "." & asElementItems(12, iLoop2) & ")"

                  fValid9 = False
                
                  asItems = wfTempElement.Items
  
                  For iLoop3 = 1 To UBound(asItems, 2)
                    If ((asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_CHAR) _
                        Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                        Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                        Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_DATE) _
                        Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                        Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                        Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                        Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
                      And UCase(Trim(asItems(9, iLoop3))) = UCase(Trim(asElementItems(12, iLoop2))) Then
  
                      fValid9 = True
                      Exit For
                    End If
                  Next iLoop3
                End If
              End If

              Exit For
            End If
            Set wfTempElement = Nothing
          Next iLoop
        End If
  
        '------------------------------------------------------------
        ' Add the required validation messages to the array.
        '------------------------------------------------------------
        ' 8. WebForm element items (WFValue) must have valid WebForm identifier.
        ' 9. WebForm element items (WFValue) must have valid WebForm InputValue identifier.
        If Not fValid8 Then
          asElementItems(11, iLoop2) = ""
          asElementItems(12, iLoop2) = ""
  
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Workflow Value" & sSubMessage1 & " - Invalid web form identifier", _
            pwfElement.ControlIndex
        End If
        If Not fValid9 Then
          asElementItems(12, iLoop2) = ""
  
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Workflow Value" & sSubMessage1 & " - Invalid value identifier", _
            pwfElement.ControlIndex
        End If
      
      '------------------------------------------------------------
      Case giWFFORMITEM_INPUTVALUE_NUMERIC
        ' InputNum must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        ' AE20090409 Fault #13655
'        fValid19 = (CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC) _
'          Or (Len(Replace(asElementItems(10, iLoop2), ".", "")) <= CInt(asElementItems(7, iLoop2)))
        fValid19 = (CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC) _
          Or (Len(Replace(Replace(asElementItems(10, iLoop2), ".", ""), "-", "")) <= CInt(asElementItems(7, iLoop2)))

        If CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC Then
          If (CLng(asElementItems(56, iLoop2)) > 0) Then
            ValidateElement_Expression _
              pwfElement, _
              CLng(asElementItems(56, iLoop2)), _
              "Numeric Input (" & asElementItems(9, iLoop2) & " - Invalid default value calculation", _
              pavarDisconnectedElements
          Else
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Numeric Input (" & asElementItems(9, iLoop2) & ") - No default value calculation selected", _
              pwfElement.ControlIndex
          End If
        End If

        fValid20 = True
        fValid22 = (CInt(asElementItems(7, iLoop2)) <= WORKFLOWWEBFORM_MAXSIZE_NUMINPUT)
        
        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then

              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If

        If fValid19 _
          And (CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_FIXED) _
          And (Len(asElementItems(10, iLoop2)) > 0) Then
          ' Size was ok, now check decimals
          iIndex = InStr(asElementItems(10, iLoop2), ".")
          If iIndex > 0 Then
            fValid20 = (Len(Mid(asElementItems(10, iLoop2), iIndex + 1)) <= CInt(asElementItems(8, iLoop2)))
          End If
        End If
        
        If (Not fValid11) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Numeric Input - no identifier", _
            pwfElement.ControlIndex
        End If
        If (Not fValid12) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Numeric Input (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If
        If (Not fValid19) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Numeric Input (" & asElementItems(9, iLoop2) & ") - default value too big", _
            pwfElement.ControlIndex
        End If
        If (Not fValid20) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Numeric Input (" & asElementItems(9, iLoop2) & ") - default value has too many decimals", _
            pwfElement.ControlIndex
        End If
        If (Not fValid22) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Numeric Input (" & asElementItems(9, iLoop2) & ") - size exceeds maximum (" & CStr(WORKFLOWWEBFORM_MAXSIZE_NUMINPUT) & ")", _
            pwfElement.ControlIndex
        End If
        
      '------------------------------------------------------------
      Case giWFFORMITEM_INPUTVALUE_LOGIC
        ' InputLogic must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        
        fValid19 = (CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC) _
          Or (UCase(asElementItems(10, iLoop2)) = "TRUE") _
          Or (UCase(asElementItems(10, iLoop2)) = "FALSE")
        
        If CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC Then
          If (CLng(asElementItems(56, iLoop2)) > 0) Then
            ValidateElement_Expression _
              pwfElement, _
              CLng(asElementItems(56, iLoop2)), _
              "Logic Input (" & asElementItems(9, iLoop2) & " - Invalid default value calculation", _
              pavarDisconnectedElements
          Else
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Logic Input (" & asElementItems(9, iLoop2) & ") - No default value calculation selected", _
              pwfElement.ControlIndex
          End If
        End If
        
        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then

              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If

        If (Not fValid11) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Logic Input - no identifier", _
            pwfElement.ControlIndex
        End If
        If (Not fValid12) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Logic Input (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If
        If (Not fValid19) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Logic Input (" & asElementItems(9, iLoop2) & ") - invalid default value", _
            pwfElement.ControlIndex
        End If
      
      '------------------------------------------------------------
      Case giWFFORMITEM_INPUTVALUE_DATE
        ' InputDate must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        fValid19 = True
        If CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC Then
          If (CLng(asElementItems(56, iLoop2)) > 0) Then
            ValidateElement_Expression _
              pwfElement, _
              CLng(asElementItems(56, iLoop2)), _
              "Date Input (" & asElementItems(9, iLoop2) & " - Invalid default value calculation", _
              pavarDisconnectedElements
          Else
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Date Input (" & asElementItems(9, iLoop2) & ") - No default value calculation selected", _
              pwfElement.ControlIndex
          End If
        End If
        
        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then

              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If

        If (Len(asElementItems(10, iLoop2)) > 0) _
          And (CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_FIXED) Then
          Set objMisc = New Misc
          fValid19 = IsDate(objMisc.ConvertSQLDateToLocale(asElementItems(10, iLoop2)))
          Set objMisc = Nothing
        End If
        
        If (Not fValid11) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Date Input - no identifier", _
            pwfElement.ControlIndex
        End If
        If (Not fValid12) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Date Input (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If
        If (Not fValid19) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Date Input (" & asElementItems(9, iLoop2) & ") - default value invalid", _
            pwfElement.ControlIndex
        End If
      
      '------------------------------------------------------------
      Case giWFFORMITEM_FRAME
        ' No validation required.
      
      '------------------------------------------------------------
      Case giWFFORMITEM_LINE
        ' No validation required.

      '------------------------------------------------------------
      Case giWFFORMITEM_IMAGE
        ' 13. WebForm element items (Image) must have a valid picture.
        fValid13 = (val(asElementItems(25, iLoop2)) > 0)
        If fValid13 Then
          ' Picture defined - does it still exist?
          With recPictEdit
            .Index = "idxID"
            .Seek "=", CLng(asElementItems(25, iLoop2))

            If .NoMatch Then
              fValid13 = False
            Else
              If !Deleted Then
                fValid13 = False
              End If
            End If
          End With
        End If

        If (Not fValid13) And (Not fDoingDeleteCheck) Then
          asElementItems(25, iLoop2) = "0"
          
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Image - invalid picture", _
            pwfElement.ControlIndex
        End If
      
      '------------------------------------------------------------
      Case giWFFORMITEM_INPUTVALUE_GRID
        ' InputRecSel must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then

              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If

        If (Not fValid11) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Record Selector Input - no identifier", _
            pwfElement.ControlIndex
        End If
        If (Not fValid12) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Record Selector Input (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If
        
        ' 5. WebForm element items (RecSel) must have valid record.
        ' 6. WebForm element items (RecSel) must have valid record element identifier (where required).
        ' 7. WebForm element items (RecSel) must have valid record selector identifier (where required).
        ' 10. WebForm element items (RecSel) must have valid table.
        fValid5 = True
        fValid6 = True
        fValid7 = True
        fValid10 = (CLng(asElementItems(44, iLoop2)) > 0)
        If fValid10 Then
          ' Table defined - does it still exist?
          With recTabEdit
            .Index = "idxTableID"
            .Seek "=", CLng(asElementItems(44, iLoop2))

            If .NoMatch Then
              fValid10 = False
            Else
              If !Deleted Then
                fValid10 = False
              Else
                lngTableID = !TableID
              End If
            End If
          End With
        End If

        If fValid10 Then
          Select Case asElementItems(5, iLoop2)
            Case giWFRECSEL_UNKNOWN           ' Oops, something's wrong
              fValid5 = False

            Case giWFRECSEL_INITIATOR           ' Initiator's personnel table record
              ' RecSel must be based on a child of the Personnel table.
              ReDim alngValidTables(0)
              TableAscendants mlngPersonnelTableID, alngValidTables
                            
              fFound = False
              For lngLoop = 1 To UBound(alngValidTables)
                If IsChildOfTable(alngValidTables(lngLoop), lngTableID) Then
                  fFound = True
                  Exit For
                End If
              Next lngLoop
              fValid5 = fFound

            Case giWFRECSEL_TRIGGEREDRECORD     ' Base table record
              ' RecSel must be based on a child of the Base table.
              ReDim alngValidTables(0)
              TableAscendants mlngBaseTableID, alngValidTables
                            
              fFound = False
              For lngLoop = 1 To UBound(alngValidTables)
                If IsChildOfTable(alngValidTables(lngLoop), lngTableID) Then
                  fFound = True
                  Exit For
                End If
              Next lngLoop
              fValid5 = fFound

            Case giWFRECSEL_IDENTIFIEDRECORD    ' Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
              ' Check identification is valid.
              ' RecSel must be based on a child of the identified record's table.
              fValid6 = (Len(Trim(asElementItems(11, iLoop2))) > 0)

              If fValid6 Then
                fValid6 = False
                
                For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
                  Set wfTempElement = aWFPrecedingElements(iLoop)

                  If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(asElementItems(11, iLoop2))) Then
                    If wfTempElement.ElementType = elem_WebForm Then
                      fValid6 = True

                      fValid7 = (Len(Trim(asElementItems(12, iLoop2))) > 0)
                      If fValid7 Then
                        fValid7 = False

                        asItems = wfTempElement.Items

                        For iLoop3 = 1 To UBound(asItems, 2)
                          If (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_GRID) _
                            And UCase(Trim(asItems(9, iLoop3))) = UCase(Trim(asElementItems(12, iLoop2))) Then

                            ReDim alngValidTables(0)
                            TableAscendants CLng(asItems(44, iLoop3)), alngValidTables
                                          
                            fFound = False
                            For lngLoop = 1 To UBound(alngValidTables)
                              If IsChildOfTable(alngValidTables(lngLoop), lngTableID) Then
                                fFound = True
                                Exit For
                              End If
                            Next lngLoop
                            fValid7 = fFound
                            Exit For
                          End If
                        Next iLoop3
                      End If
                      Exit For
                    ElseIf wfTempElement.ElementType = elem_StoredData Then
                      ReDim alngValidTables(0)
                      TableAscendants wfTempElement.DataTableID, alngValidTables
                                    
                      'JPD 20061227
                      If wfTempElement.DataAction = DATAACTION_DELETE Then
                        ' Cannot do anything with a Deleted record, but can use its ascendants.
                        ' Remove the table itself from the array of valid tables.
                        alngValidTables(1) = 0
                      End If
                                    
                      fFound = False
                      For lngLoop = 1 To UBound(alngValidTables)
                        If IsChildOfTable(alngValidTables(lngLoop), lngTableID) Then
                          fFound = True
                          Exit For
                        End If
                      Next lngLoop
                      fValid6 = fFound
                      Exit For
                    End If
                  End If
                  Set wfTempElement = Nothing
                Next iLoop
              End If

            Case giWFRECSEL_ALL                 ' Show all records from the table in a WebForm RecordSelector
              ' No validation required.

            Case giWFRECSEL_UNIDENTIFIED        ' Used when StoredData Inserts into a top-level table
              ' Not used for RecSel selection.
              fValid5 = False
          End Select
        End If

        'JPD 20070424 Fault 12164
        ' Validate the recordSelector filter (if required)
        If (CLng(asElementItems(53, iLoop2)) > 0) Then
          ValidateElement_Expression _
            pwfElement, _
            CLng(asElementItems(53, iLoop2)), _
            "Invalid record filter", _
            pavarDisconnectedElements
        End If

        '------------------------------------------------------------
        ' Add the required validation messages to the array.
        '------------------------------------------------------------
        ' 5. WebForm element items (RecSel) must have valid record.
        ' 6. WebForm element items (RecSel) must have valid record element identifier (where required).
        ' 7. WebForm element items (RecSel) must have valid record selector identifier (where required).
        ' 10. WebForm element items (RecSel) must have valid table.
        If (Not fValid10) And (Not fDoingDeleteCheck) Then
          asElementItems(4, iLoop2) = ""

          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Record Selector Input (" & asElementItems(9, iLoop2) & ") - Invalid table", _
            pwfElement.ControlIndex
        End If
        If (Not fValid5) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Record Selector Input (" & asElementItems(9, iLoop2) & ") - Invalid record", _
            pwfElement.ControlIndex
        End If
        If Not fValid6 Then
          asElementItems(11, iLoop2) = ""
          asElementItems(12, iLoop2) = ""

          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Record Selector Input (" & asElementItems(9, iLoop2) & ") - Invalid element identifier", _
            pwfElement.ControlIndex
        End If
        If Not fValid7 Then
          asElementItems(12, iLoop2) = ""

          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Record Selector Input (" & asElementItems(9, iLoop2) & ") - Invalid record selector", _
            pwfElement.ControlIndex
        End If
  
      '------------------------------------------------------------
      Case giWFFORMITEM_FORMATCODE  ' NB. Only used in emails.
        ' No validation required.
    
      '------------------------------------------------------------
      Case giWFFORMITEM_INPUTVALUE_DROPDOWN
        ' Button must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then
              
              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If

        ' 17. WebForm element items (Dropdown/OptionGroup) must have ControValues defined.
        fValid17 = (Len(Trim(asElementItems(47, iLoop2))) > 0)
        fValid19 = True
      
        If fValid17 Then
          If CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC Then
            If (CLng(asElementItems(56, iLoop2)) > 0) Then
              ValidateElement_Expression _
                pwfElement, _
                CLng(asElementItems(56, iLoop2)), _
                "Dropdown Input (" & asElementItems(9, iLoop2) & " - Invalid default value calculation", _
                pavarDisconnectedElements
            Else
              ValidateWorkflow_AddMessage _
                sMessagePrefix & "Dropdown Input (" & asElementItems(9, iLoop2) & ") - No default value calculation selected", _
                pwfElement.ControlIndex
            End If
          End If
        
          If (Len(asElementItems(10, iLoop2)) > 0) _
            And (CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_FIXED) Then
            
            fValid19 = (InStr(vbTab & asElementItems(47, iLoop2) & vbTab, vbTab & asElementItems(10, iLoop2) & vbTab) > 0)
          End If
        End If
      
        '------------------------------------------------------------
        ' Add the required validation messages to the array.
        '------------------------------------------------------------
        If Not fValid11 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Dropdown Input - no identifier", _
            pwfElement.ControlIndex
        End If
        If Not fValid12 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Dropdown Input (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If
        
        ' 17. WebForm element items (Dropdown/OptionGroup) must have ControValues defined.
        If (Not fValid17) And (Not fDoingDeleteCheck) Then
          asElementItems(47, iLoop2) = ""
  
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Dropdown Input (" & asElementItems(9, iLoop2) & ") - no control values", _
            pwfElement.ControlIndex
        End If
        If (Not fValid19) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Dropdown Input (" & asElementItems(9, iLoop2) & ") - default value invalid", _
            pwfElement.ControlIndex
        End If
      
      '------------------------------------------------------------
      Case giWFFORMITEM_INPUTVALUE_LOOKUP
        ' Item must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then
              
              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If
      
        ' 10. WebForm element items (Lookup) must have valid table/column.
        fValid10 = (CLng(asElementItems(49, iLoop2)) > 0)
        If fValid10 Then
          ' Column defined - does it still exist?
          With recColEdit
            .Index = "idxColumnID"
            .Seek "=", CLng(asElementItems(49, iLoop2))
  
            If .NoMatch Then
              fValid10 = False
            Else
              If !Deleted Then
                fValid10 = False
              End If
            End If
          End With
        End If
  
        If CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC Then
          If (CLng(asElementItems(56, iLoop2)) > 0) Then
            ValidateElement_Expression _
              pwfElement, _
              CLng(asElementItems(56, iLoop2)), _
              "Lookup Input (" & asElementItems(9, iLoop2) & " - Invalid default value calculation", _
              pavarDisconnectedElements
          Else
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Lookup Input (" & asElementItems(9, iLoop2) & ") - No default value calculation selected", _
              pwfElement.ControlIndex
          End If
        End If

        '------------------------------------------------------------
        ' Add the required validation messages to the array.
        '------------------------------------------------------------
        If Not fValid11 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Lookup Input - no identifier", _
            pwfElement.ControlIndex
        End If
        If Not fValid12 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Lookup Input (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If
        ' 10. WebForm element items (Lookup) must have valid table/column.
        If (Not fValid10) And (Not fDoingDeleteCheck) Then
          asElementItems(49, iLoop2) = ""
  
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Lookup Input (" & asElementItems(9, iLoop2) & ") - invalid column", _
            pwfElement.ControlIndex
        End If
      
      '------------------------------------------------------------
      Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
        ' Item must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then
              
              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If

        ' 17. WebForm element items (Dropdown/OptionGroup) must have ControValues defined.
        fValid17 = (Len(Trim(asElementItems(47, iLoop2))) > 0)
        fValid19 = True
        
        If fValid17 Then
          If CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_CALC Then
            If (CLng(asElementItems(56, iLoop2)) > 0) Then
              ValidateElement_Expression _
                pwfElement, _
                CLng(asElementItems(56, iLoop2)), _
                "Option Group Input (" & asElementItems(9, iLoop2) & " - Invalid default value calculation", _
                pavarDisconnectedElements
            Else
              ValidateWorkflow_AddMessage _
                sMessagePrefix & "Option Group Input (" & asElementItems(9, iLoop2) & ") - No default value calculation selected", _
                pwfElement.ControlIndex
            End If
          End If
        
          If (Len(asElementItems(10, iLoop2)) > 0) _
            And (CInt(asElementItems(58, iLoop2)) = giWFDATAVALUE_FIXED) Then
            fValid19 = (InStr(vbTab & asElementItems(47, iLoop2) & vbTab, vbTab & asElementItems(10, iLoop2) & vbTab) > 0)
          End If
        End If
      
        '------------------------------------------------------------
        ' Add the required validation messages to the array.
        '------------------------------------------------------------
        If Not fValid11 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Option Group Input - no identifier", _
            pwfElement.ControlIndex
        End If
        If Not fValid12 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Option Group Input (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If
      
        ' 17. WebForm element items (Dropdown/OptionGroup) must have ControlValues defined.
        If (Not fValid17) And (Not fDoingDeleteCheck) Then
          asElementItems(47, iLoop2) = ""
  
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Option Group Input (" & asElementItems(9, iLoop2) & ") - no control values", _
            pwfElement.ControlIndex
        End If
        
        If (Not fValid19) And (Not fDoingDeleteCheck) Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Option Group Input (" & asElementItems(9, iLoop2) & ") - default value invalid", _
            pwfElement.ControlIndex
        End If

      '------------------------------------------------------------
      Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD
        ' Item must have an identifier
        fValid11 = (Len(Trim(asElementItems(9, iLoop2))) > 0)
        fValid12 = True
        If fValid11 Then
          ' Identifier must be unique within this WebForm.
          For iLoop = 1 To UBound(asElementItems, 2)
            If (iLoop <> iLoop2) _
              And ((CInt(asElementItems(2, iLoop)) = giWFFORMITEM_BUTTON) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_CHAR) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DATE) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_GRID) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                Or (CInt(asElementItems(2, iLoop)) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
              And (UCase(Trim(asElementItems(9, iLoop2))) = UCase(Trim(asElementItems(9, iLoop)))) Then
              
              fValid12 = False
              Exit For
            End If
          Next iLoop
        End If

        If Not fValid11 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "File Upload - no identifier", _
            pwfElement.ControlIndex
        End If
        If Not fValid12 Then
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "File Upload (" & asElementItems(9, iLoop2) & ") - non-unique identifier", _
            pwfElement.ControlIndex
        End If

    End Select
  Next iLoop2

  '------------------------------------------------------------
  ' Add the required validation messages to the array.
  '------------------------------------------------------------
  ' 4. WebForm element must have at least 1 button.
  If (Not fValid4) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Must have at least 1 Submit button", _
      pwfElement.ControlIndex
  End If

  ' ------------------------------------
  ' Validate the WebForm validations
  ' ------------------------------------
  asValidations = pwfElement.Validations
  
  For iLoop2 = 1 To UBound(asValidations, 2)
    If (CLng(asValidations(1, iLoop2)) > 0) Then
      ValidateElement_Expression _
        pwfElement, _
        CLng(asValidations(1, iLoop2)), _
        "Invalid validation calculation", _
        pavarDisconnectedElements
    Else
      ValidateWorkflow_AddMessage _
        sMessagePrefix & "Validation - No calculation selected", _
        pwfElement.ControlIndex
    End If
  
    If (Len(Trim(asValidations(3, iLoop2))) = 0) Then
      ValidateWorkflow_AddMessage _
        sMessagePrefix & "Validation - <" & GetExpressionName(CLng(asValidations(1, iLoop2))) & "> - No message", _
        pwfElement.ControlIndex
    End If
  Next iLoop2

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
End Sub

Private Sub ValidateElement_StoredData(pwfElement As VB.Control, _
  pfFix As Boolean, _
  Optional pavarDisconnectedElements As Variant)
  
  On Error GoTo ErrorTrap
  
  Dim wfTempElement As VB.Control
  Dim lngTempElementIndex As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iLoop4 As Integer
  Dim iLoop5 As Integer
  Dim iLoop6 As Integer
  Dim lngLoop As Long
  Dim aWFPrecedingElements() As VB.Control
  Dim fValid1 As Boolean
  Dim fValid2 As Boolean
  Dim fValid3 As Boolean
  Dim fValid4 As Boolean
  Dim fValid5 As Boolean
  Dim fValid6 As Boolean
  Dim fValid7 As Boolean
  Dim fValid8 As Boolean
  Dim fValid9 As Boolean
  Dim fValid10 As Boolean
  Dim fValid11 As Boolean
  Dim fValid12 As Boolean
  Dim fValid13 As Boolean
  Dim fValid14 As Boolean
  Dim fValid15 As Boolean
  Dim fValid16 As Boolean
  Dim fValid17 As Boolean
  Dim fValid18 As Boolean
  Dim fValid19 As Boolean
  Dim fValid20 As Boolean
  Dim asItems() As String
  Dim avColumns() As Variant
  Dim iTableType As Integer
  Dim sSQL As String
  Dim rsInfo As DAO.Recordset
  Dim iParentCount As Integer
  Dim sMessagePrefix As String
  Dim sMessagePrefix2 As String
  Dim iColumnDataType As DataTypes
  Dim iColumnSize As Integer
  Dim iColumnDecimals As Integer
  Dim lngSpinnerMin As Long
  Dim lngSpinnerMax As Long
  Dim sValue As String
  Dim sSubMessage1 As String
  Dim sSubMessage2 As String
  Dim sSubMessage3 As String
  Dim sSubMessage4 As String
  Dim objMisc As Misc
  Dim sTemp As String
  Dim iDBColumnDataType As DataTypes
  Dim iDBColumnSize As Integer
  Dim iDBColumnDecimals As Integer
  Dim lngDBColumnTableID As Long
  Dim iMaxLength As Integer
  Dim asItemValues() As String
  Dim lngPrimaryParentTable As Long
  Dim lngSecondaryParentTable As Long
  Dim fTableOK As Boolean
  Dim alngValidTables() As Long
  Dim fFound As Boolean
  Dim lngExcludedTableID As Long
  Dim fDisconnectedElement As Boolean
  Dim fDoingDeleteCheck As Boolean
  Dim iIndex As Integer
  Dim sFormatString As String
  
  fDoingDeleteCheck = Not IsMissing(pavarDisconnectedElements)
  
  sMessagePrefix = ValidateElement_MessagePrefix(pwfElement)
  lngExcludedTableID = 0

  ' Get the elements that precede the given element.
  ReDim aWFPrecedingElements(1)
  Set aWFPrecedingElements(UBound(aWFPrecedingElements)) = pwfElement
  PrecedingElements pwfElement, aWFPrecedingElements
        
  If fDoingDeleteCheck Then
    ' Remove the element we're trying to delete from the array of preceding elements, if we're doing the deletion check.
    iIndex = 1
    For iLoop = 1 To UBound(aWFPrecedingElements)
      fFound = False
      For iLoop2 = 1 To UBound(pavarDisconnectedElements, 2)
        If aWFPrecedingElements(iLoop) Is pavarDisconnectedElements(1, iLoop2) Then
          fFound = True
          Exit For
        End If
      Next iLoop2
      
      If Not fFound Then
        Set aWFPrecedingElements(iIndex) = aWFPrecedingElements(iLoop)
        iIndex = iIndex + 1
      End If
    Next iLoop
    ReDim Preserve aWFPrecedingElements(iIndex - 1)
  End If
        
  '------------------------------------------------------------
  ' 1. StoredData element must have an identifier.
  ' 2. StoredData element must have a unique identifier.
  ' 3. StoredData element must have valid table.
  ' 4. StoredData element must have at least one column.
  ' 5. StoredData elements for Updates/Deletes or Inserts into histories must have valid primary record.
  ' 6. StoredData element with Identified primary record must have valid Element.
  ' 7. StoredData element with RecSel Identified primary record must have valid RecSel.

  ' 8. StoredData elements for Inserts into shared histories must have valid secondary record (or unidentified).
  ' 9. StoredData element with Identified secondary record must have valid Element.
  ' 10. StoredData element with RecSel Identified secondary record must have valid RecSel.

  ' 11. StoredData element columns must be the correct type, size, etc.

  ' 12. StoredData element columns (DBValue) must have valid column.
  ' 13. StoredData element columns (DBValue) must have valid record.
  ' 14. StoredData element columns (DBValue) must have valid record element identifier (where required).
  ' 15. StoredData element columns (DBValue) must have valid record selector identifier (where required).

  ' 16. StoredData element columns (WFValue) must have valid WebForm identifier.
  ' 17. StoredData element columns (WFValue) must have valid WebForm InputValue identifier.
  
  ' 20. StoredData element for shared history (ie. with Secondary record) must have secondary parent table different to primary parent table.
  '------------------------------------------------------------

  fValid1 = (Len(Trim(pwfElement.Identifier)) > 0)
  fValid2 = True
  If fValid1 Then
    fValid2 = UniqueIdentifier(pwfElement.Identifier, pwfElement.ControlIndex)
  End If
  
  fValid3 = (pwfElement.DataTableID > 0)
  If fValid3 Then
    With recTabEdit
      .Index = "idxTableID"
      .Seek "=", pwfElement.DataTableID

      If .NoMatch Then
        fValid3 = False
      Else
        If !Deleted Then
          fValid3 = False
        Else
          iTableType = !TableType
        
          sSQL = "SELECT COUNT(*) AS recCount" & _
            " FROM tmpRelations" & _
            " WHERE childID = " & CStr(pwfElement.DataTableID)
          Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          iParentCount = rsInfo!reccount
          rsInfo.Close
          Set rsInfo = Nothing
        End If
      End If
    End With
  End If
            
  lngPrimaryParentTable = 0
  lngSecondaryParentTable = 0
  
  fValid5 = True
  fValid6 = True
  fValid7 = True
  If fValid3 Then
    If (pwfElement.DataAction <> DATAACTION_INSERT) _
      Or (iTableType = iTabChild) Then
      ' 5. StoredData elements for Updates/Deletes or Inserts into histories must have valid primary record.
      Select Case pwfElement.DataRecord
        Case giWFRECSEL_UNKNOWN           ' Oops, something's wrong
          fValid5 = False

        Case giWFRECSEL_INITIATOR           ' Initiator's personnel table record
          ' Must be Updating/Deleting a record in the Personnel table.
          ' Or Inserting into a history of the Personnel table.
          ' Get an array of the valid table IDs (base table and it's ascendants)
          ReDim alngValidTables(0)
          TableAscendants mlngPersonnelTableID, alngValidTables
                    
          If (pwfElement.DataAction = DATAACTION_INSERT) Then
            fFound = False
            
            For lngLoop = 1 To UBound(alngValidTables)
              If IsChildOfTable(alngValidTables(lngLoop), pwfElement.DataTableID) _
                And (alngValidTables(lngLoop) = pwfElement.DataRecordTableID) Then
                
                lngPrimaryParentTable = alngValidTables(lngLoop)
                fFound = True
                Exit For
              End If
            Next lngLoop
            
            fValid5 = fFound
          Else
            fFound = False
            
            For lngLoop = 1 To UBound(alngValidTables)
              If alngValidTables(lngLoop) = pwfElement.DataTableID Then
                fFound = True
                Exit For
              End If
            Next lngLoop
            
            fValid5 = fFound
          End If
          
        Case giWFRECSEL_TRIGGEREDRECORD      ' Base table record
          ' Must be Updating/Deleting a record in the base table.
          ' Or Inserting into a history of the base table.
          ' Get an array of the valid table IDs (base table and it's ascendants)
          ReDim alngValidTables(0)
          TableAscendants mlngBaseTableID, alngValidTables
                    
          If (pwfElement.DataAction = DATAACTION_INSERT) Then
            fFound = False
            
            For lngLoop = 1 To UBound(alngValidTables)
              If IsChildOfTable(alngValidTables(lngLoop), pwfElement.DataTableID) _
                And (alngValidTables(lngLoop) = pwfElement.DataRecordTableID) Then
                
                lngPrimaryParentTable = alngValidTables(lngLoop)
                fFound = True
                Exit For
              End If
            Next lngLoop
            
            fValid5 = fFound
          Else
            fFound = False
            
            For lngLoop = 1 To UBound(alngValidTables)
              If alngValidTables(lngLoop) = pwfElement.DataTableID Then
                fFound = True
                Exit For
              End If
            Next lngLoop
            
            fValid5 = fFound
          End If
          
        Case giWFRECSEL_IDENTIFIEDRECORD    ' Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
          ' Check identification is valid.
          fValid6 = (Len(Trim(pwfElement.RecordSelectorWebFormIdentifier)) > 0)

          If fValid6 Then
            fValid6 = False
            
            For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
              Set wfTempElement = aWFPrecedingElements(iLoop)
              
              If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(pwfElement.RecordSelectorWebFormIdentifier)) Then
                If wfTempElement.ElementType = elem_WebForm Then
                  fValid6 = True

                  fValid7 = (Len(Trim(pwfElement.RecordSelectorIdentifier)) > 0)
                  If fValid7 Then
                    fValid7 = False

                    asItems = wfTempElement.Items

                    For iLoop3 = 1 To UBound(asItems, 2)
                      If (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_GRID) _
                        And UCase(Trim(asItems(9, iLoop3))) = UCase(Trim(pwfElement.RecordSelectorIdentifier)) Then

                        ' Get an array of the valid table IDs (base table and it's ascendants)
                        ReDim alngValidTables(0)
                        TableAscendants CLng(asItems(44, iLoop3)), alngValidTables
                                                
                        If (pwfElement.DataAction = DATAACTION_INSERT) Then
                          fFound = False
                          
                          For lngLoop = 1 To UBound(alngValidTables)
                            If IsChildOfTable(alngValidTables(lngLoop), pwfElement.DataTableID) _
                              And (alngValidTables(lngLoop) = pwfElement.DataRecordTableID) Then
                              
                              lngPrimaryParentTable = alngValidTables(lngLoop)
                              fFound = True
                              Exit For
                            End If
                          Next lngLoop
            
                          fValid7 = fFound
                        Else
                          fFound = False
                          
                          For lngLoop = 1 To UBound(alngValidTables)
                            If alngValidTables(lngLoop) = pwfElement.DataTableID Then
                              fFound = True
                              Exit For
                            End If
                          Next lngLoop
            
                          fValid7 = fFound
                        End If
                        Exit For
                      End If
                    Next iLoop3
                  End If
                  Exit For
                ElseIf wfTempElement.ElementType = elem_StoredData Then
                  ' Get an array of the valid table IDs (base table and it's ascendants)
                  ReDim alngValidTables(0)
                  TableAscendants wfTempElement.DataTableID, alngValidTables

                  'JPD 20061227
                  If wfTempElement.DataAction = DATAACTION_DELETE Then
                    alngValidTables(1) = 0
                  End If

                  If (pwfElement.DataAction = DATAACTION_INSERT) Then
                    fFound = False
                    
                    For lngLoop = 1 To UBound(alngValidTables)
                      If IsChildOfTable(alngValidTables(lngLoop), pwfElement.DataTableID) _
                        And (alngValidTables(lngLoop) = pwfElement.DataRecordTableID) Then
                        
                        lngPrimaryParentTable = alngValidTables(lngLoop)
                        fFound = True
                        Exit For
                      End If
                    Next lngLoop
            
                    fValid6 = fFound
                  Else
                    fFound = False
                    
                    For lngLoop = 1 To UBound(alngValidTables)
                      If alngValidTables(lngLoop) = pwfElement.DataTableID Then
                        fFound = True
                        Exit For
                      End If
                    Next lngLoop
            
                    fValid6 = fFound
                  End If
                  Exit For
                End If
              End If
              Set wfTempElement = Nothing
            Next iLoop
          End If

        Case giWFRECSEL_ALL                 ' Show all records from the table in a WebForm RecordSelector
          ' Not used for Primary Record selection.
          fValid5 = False

        Case giWFRECSEL_UNIDENTIFIED        ' Used when StoredData Inserts into a top-level table
          ' Not used for Primary Record selection for Updates/Deletes or Inserts into histories.
          fValid5 = False
      End Select
    End If
  End If
  
  fValid8 = True
  fValid9 = True
  fValid10 = True
  fValid20 = True
  If fValid3 Then
    If (pwfElement.DataAction = DATAACTION_INSERT) _
      And (iParentCount > 1) Then
      
      ' 8. StoredData elements for Inserts into shared histories must have valid secondary record.
      ' NB. 'Unidentified' is a valid selection
      Select Case pwfElement.SecondaryDataRecord
        Case giWFRECSEL_UNKNOWN           ' Oops, something's wrong
          fValid8 = False

        Case giWFRECSEL_INITIATOR           ' Initiator's personnel table record
          ' Must be Inserting into a history of the Personnel table.
          ' Get an array of the valid table IDs (base table and it's ascendants)
          ReDim alngValidTables(0)
          TableAscendants mlngPersonnelTableID, alngValidTables
          
          fFound = False
          
          For lngLoop = 1 To UBound(alngValidTables)
            If IsChildOfTable(alngValidTables(lngLoop), pwfElement.DataTableID) _
              And (alngValidTables(lngLoop) = pwfElement.SecondaryDataRecordTableID) Then
              
              lngSecondaryParentTable = alngValidTables(lngLoop)
              fFound = True
              Exit For
            End If
          Next lngLoop
            
          fValid8 = fFound

        Case giWFRECSEL_TRIGGEREDRECORD     ' Base table record
          ' Must be Inserting into a history of the base table.
          ' Get an array of the valid table IDs (base table and it's ascendants)
          ReDim alngValidTables(0)
          TableAscendants mlngBaseTableID, alngValidTables
          
          fFound = False
          
          For lngLoop = 1 To UBound(alngValidTables)
            If IsChildOfTable(alngValidTables(lngLoop), pwfElement.DataTableID) _
              And (alngValidTables(lngLoop) = pwfElement.SecondaryDataRecordTableID) Then
              
              lngSecondaryParentTable = alngValidTables(lngLoop)
              fFound = True
              Exit For
            End If
          Next lngLoop
            
          fValid8 = fFound

        Case giWFRECSEL_IDENTIFIEDRECORD    ' Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
          ' Check identification is valid.
          fValid9 = (Len(Trim(pwfElement.SecondaryRecordSelectorWebFormIdentifier)) > 0)

          If fValid9 Then
            fValid9 = False
            
            For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
              Set wfTempElement = aWFPrecedingElements(iLoop)

              If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(pwfElement.SecondaryRecordSelectorWebFormIdentifier)) Then
                If wfTempElement.ElementType = elem_WebForm Then
                  fValid9 = True

                  fValid10 = (Len(Trim(pwfElement.SecondaryRecordSelectorIdentifier)) > 0)
                  If fValid10 Then
                    fValid10 = False

                    asItems = wfTempElement.Items

                    For iLoop3 = 1 To UBound(asItems, 2)
                      If (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_GRID) _
                        And UCase(Trim(asItems(9, iLoop3))) = UCase(Trim(pwfElement.SecondaryRecordSelectorIdentifier)) Then

                        ' Get an array of the valid table IDs (base table and it's ascendants)
                        ReDim alngValidTables(0)
                        TableAscendants CLng(asItems(44, iLoop3)), alngValidTables
                                                
                        fFound = False
                        
                        For lngLoop = 1 To UBound(alngValidTables)
                          If IsChildOfTable(alngValidTables(lngLoop), pwfElement.DataTableID) _
                            And (alngValidTables(lngLoop) = pwfElement.SecondaryDataRecordTableID) Then
                            
                            lngSecondaryParentTable = alngValidTables(lngLoop)
                            fFound = True
                            Exit For
                          End If
                        Next lngLoop
            
                        fValid10 = fFound
                        Exit For
                      End If
                    Next iLoop3
                  End If
                  Exit For
                ElseIf wfTempElement.ElementType = elem_StoredData Then
                  ' Get an array of the valid table IDs (base table and it's ascendants)
                  ReDim alngValidTables(0)
                  TableAscendants wfTempElement.DataTableID, alngValidTables
                  
                  'JPD 20061227
                  If wfTempElement.DataAction = DATAACTION_DELETE Then
                    ' Cannot do anything with a Deleted record, but can use its ascendants.
                    ' Remove the table itself from the array of valid tables.
                    alngValidTables(1) = 0
                  End If
                                    
                  fFound = False
                  
                  For lngLoop = 1 To UBound(alngValidTables)
                    If IsChildOfTable(alngValidTables(lngLoop), pwfElement.DataTableID) _
                      And (alngValidTables(lngLoop) = pwfElement.SecondaryDataRecordTableID) Then
                      
                      lngSecondaryParentTable = alngValidTables(lngLoop)
                      fFound = True
                      Exit For
                    End If
                  Next lngLoop
            
                  fValid9 = fFound
                  Exit For
                End If
              End If
              Set wfTempElement = Nothing
            Next iLoop
          End If

        Case giWFRECSEL_ALL                 ' Show all records from the table in a WebForm RecordSelector
          ' Not used for Secondary Record selection.
          fValid8 = False

        Case giWFRECSEL_UNIDENTIFIED        ' Used when StoredData Inserts into a top-level table
          ' Valid for Secondary Record selection for Inserts into shared histories.
          fValid8 = True
      End Select
      
      fValid20 = (lngPrimaryParentTable <> lngSecondaryParentTable) _
        And (lngPrimaryParentTable > 0)
    End If
  End If
            
  '------------------------------------------------------------
  ' Add the required validation messages to the array.
  '------------------------------------------------------------
  ' 1. StoredData element must have an identifier.
  ' 2. StoredData element must have a unique identifier.
  ' 3. StoredData element must have valid table.
  ' 4. StoredData element must have at least one column.
  ' 5. StoredData elements for Updates/Deletes or Inserts into histories must have valid primary record.
  ' 6. StoredData element with Identified primary record must have valid Element.
  ' 7. StoredData element with RecSel Identified primary record must have valid RecSel.
  ' 8. StoredData elements for Inserts into shared histories must have valid secondary record (or unidentified).
  ' 9. StoredData element with Identified secondary record must have valid Element.
  ' 10. StoredData element with RecSel Identified secondary record must have valid RecSel.
  ' 20. StoredData element for shared history (ie. with Secondary record) must have secondary parent table different to primary parent table.
  If (Not fValid1) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "No identifier", _
      pwfElement.ControlIndex
  End If
  If (Not fValid2) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Non-unique identifier", _
      pwfElement.ControlIndex
  End If
  If (Not fValid3) And (Not fDoingDeleteCheck) Then
    pwfElement.DataTableID = 0
    ReDim avColumns(10, 0)
    pwfElement.DataColumns = avColumns
    
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid table", _
      pwfElement.ControlIndex
  End If
  If (Not fValid5) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid primary record", _
      pwfElement.ControlIndex
  End If
  If Not fValid6 Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid primary record element", _
      pwfElement.ControlIndex
  End If
  If Not fValid7 Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid primary record selector", _
      pwfElement.ControlIndex
  End If
  If (Not fValid8) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid secondary record", _
      pwfElement.ControlIndex
  End If
  If Not fValid9 Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid secondary record element", _
      pwfElement.ControlIndex
  End If
  If Not fValid10 Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid secondary record selector", _
      pwfElement.ControlIndex
  End If
  If (Not fValid20) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "Invalid secondary record selector", _
      pwfElement.ControlIndex
      
    pwfElement.SecondaryDataRecord = giWFRECSEL_UNIDENTIFIED
  End If
  
  ' Validate the StoredData columns.
  avColumns = pwfElement.DataColumns
  fValid4 = (UBound(avColumns, 2) > 0) Or (pwfElement.DataAction = DATAACTION_DELETE)
  
  If (Not fValid4) And (Not fDoingDeleteCheck) Then
    ValidateWorkflow_AddMessage _
      sMessagePrefix & "No column values have been defined", _
      pwfElement.ControlIndex
  End If
  
  ' ------------------------------------
  ' Validate the StoredData columns if the table is valid.
  ' ------------------------------------
  ' 11. StoredData element columns must be for a valid column in the table.
  ' 12. StoredData element columns must be the correct type, size, etc.
  ' 13. StoredData element columns (DBValue) must have valid column.
  ' 14. StoredData element columns (DBValue) must have valid record.
  ' 15. StoredData element columns (DBValue) must have valid record element identifier (where required).
  ' 16. StoredData element columns (DBValue) must have valid record selector identifier (where required).
  ' 17. StoredData element columns (WFValue) must have valid WebForm identifier.
  ' 18. StoredData element columns (WFValue) must have valid WebForm InputValue identifier.
  
  ' DataColumns:
  ' Col 1 = Column Description
  ' Col 2 = Value Description
  ' Col 3 = Column ID
  ' Col 4 = Value Type
  ' Col 5 = Value
  ' Col 6 = WF Form Identifier
  ' Col 7 = WF Value Identifier
  ' Col 8 = DB Value Column ID
  ' Col 9 = DB Value Record
  ' Col 10 = CalcID
            
  For iLoop2 = 1 To UBound(avColumns, 2)
    fValid11 = True
    fValid12 = True
    fValid13 = True
    fValid14 = True
    fValid15 = True
    fValid16 = True
    fValid17 = True
    fValid18 = True
    fValid19 = True
    
    sSubMessage1 = ""
    sSubMessage2 = ""
    sSubMessage3 = ""
    sSubMessage4 = ""
    
    ' Check the column is still valid.
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", CLng(avColumns(3, iLoop2))

      If .NoMatch Then
        fValid11 = False
      Else
        If !Deleted _
          Or (!TableID <> pwfElement.DataTableID) _
          Or (!columntype = giCOLUMNTYPE_LINK) _
          Or (!columntype = giCOLUMNTYPE_SYSTEM) Then
          fValid11 = False
        Else
          iColumnDataType = !DataType
          iColumnSize = !Size
          iColumnDecimals = !Decimals
          'lngSpinnerMin = IIf(IsNull(.Fields("SpinnerMinimum")), 0, .Fields("SpinnerMinimum"))
          'lngSpinnerMax = IIf(IsNull(.Fields("SpinnerMaximum")), 0, .Fields("SpinnerMaximum"))
        End If
      End If
    End With
    
    If fValid11 Then
      sSubMessage1 = GetColumnName(CLng(avColumns(3, iLoop2)), True)
      sValue = UCase(CStr(avColumns(5, iLoop2)))
      
      fValid12 = True
      
      Select Case CInt(avColumns(4, iLoop2))
        Case giWFDATAVALUE_FIXED
          ' 12. StoredData element columns must be the correct type, size, etc.
          Select Case iColumnDataType
            Case sqlBoolean 'Logic
              ' Just check the value is a valid boolean term.
              fValid12 = (Trim(sValue) = "TRUE") _
                Or (Trim(sValue) = "FALSE")
            
            Case sqlLongVarChar 'Working Pattern
              fValid12 = (Len(sValue) <= 14) _
                And (Len(Trim(Replace(Mid(sValue, 1, 1), "S", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 2, 1), "S", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 3, 1), "M", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 4, 1), "M", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 5, 1), "T", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 6, 1), "T", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 7, 1), "W", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 8, 1), "W", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 9, 1), "T", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 10, 1), "T", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 11, 1), "F", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 12, 1), "F", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 13, 1), "S", ""))) = 0) _
                And (Len(Trim(Replace(Mid(sValue, 14, 1), "S", ""))) = 0)
            
              If Not fValid12 Then
                sSubMessage2 = " (" & sValue & " - invalid working pattern)"
              End If
            
            Case sqlNumeric 'Numeric
              sFormatString = "#0"
              If iColumnDecimals > 0 Then
                sFormatString = sFormatString & "."
              
                For iLoop6 = 1 To iColumnDecimals
                  sFormatString = sFormatString & "0"
                Next iLoop6
              End If
              fValid12 = (sValue = Format(CStr(val(sValue)), sFormatString))
              'fValid12 = (sValue = CStr(Val(sValue)))
              If fValid12 Then
                ' Check the size is valid.
                fValid12 = (Len(Replace(CStr(CInt(sValue)), "0", "")) <= (iColumnSize - iColumnDecimals))
              
                If Not fValid12 Then
                  sSubMessage2 = " (" & sValue & " - invalid size & decimals)"
                  sSubMessage3 = ", size = " & CStr(iColumnSize) & ", decimals = " & CStr(iColumnDecimals)
                End If
              End If
              If fValid12 Then
                ' Check the decimals is valid.
                fValid12 = (Len(Mid(sValue, Len(CStr(CInt(sValue))) + 2)) <= iColumnDecimals)
              
                If Not fValid12 Then
                  sSubMessage2 = " (" & sValue & "  - invalid decimals)"
                  sSubMessage3 = ", size = " & CStr(iColumnSize) & ", decimals = " & CStr(iColumnDecimals)
                End If
              End If
              
            Case sqlInteger 'Integer
              fValid12 = (sValue = CStr(val(sValue)))
              If fValid12 Then
                ' Check it's a valid integer
                fValid12 = (sValue = CStr(CInt(sValue)))
              End If
              
            Case sqlDate 'Date
              If (Len(sValue) > 0) And (Trim(sValue) <> "NULL") Then
                Set objMisc = New Misc
                sTemp = objMisc.ConvertSQLDateToLocale(sValue)
                fValid12 = IsDate(sTemp)
                Set objMisc = Nothing
              End If
            
            Case sqlVarChar 'Character
              ' Just check the value fits into the column.
              fValid12 = (Len(sValue) <= iColumnSize)
              
              If Not fValid12 Then
                sSubMessage2 = " (" & sValue & " - invalid size)"
                sSubMessage3 = ", size = " & CStr(iColumnSize)
              End If
          End Select
          
          '------------------------------------------------------------
          ' Add the required validation messages to the array.
          '------------------------------------------------------------
          ' 12. StoredData element columns must be the correct type, size, etc.
          If Not fValid12 Then
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Column (" & sSubMessage1 & sSubMessage3 & ") - Invalid fixed value" & sSubMessage2, _
              pwfElement.ControlIndex
          End If

        Case giWFDATAVALUE_DBVALUE
          ' 12. StoredData element columns must be the correct type, size, etc.
          ' 13. StoredData element columns (DBValue) must have valid column.
          ' 14. StoredData element columns (DBValue) must have valid record.
          ' 15. StoredData element columns (DBValue) must have valid record element identifier (where required).
          ' 16. StoredData element columns (DBValue) must have valid record selector identifier (where required).
          fValid12 = True
          fValid13 = (CLng(avColumns(8, iLoop2)) > 0)
          fValid14 = True
          fValid15 = True
          fValid16 = True
          
          If fValid13 Then
            ' Column defined - does it still exist?
            With recColEdit
              .Index = "idxColumnID"
              .Seek "=", CLng(avColumns(8, iLoop2))

              If .NoMatch Then
                fValid13 = False
              Else
                If !Deleted _
                  Or (!columntype = giCOLUMNTYPE_LINK) _
                  Or (!columntype = giCOLUMNTYPE_SYSTEM) Then
                  
                  fValid13 = False
                Else
                  iDBColumnDataType = !DataType
                  iDBColumnSize = !Size
                  iDBColumnDecimals = !Decimals
                  lngDBColumnTableID = !TableID
                  'lngSpinnerMin = IIf(IsNull(.Fields("SpinnerMinimum")), 0, .Fields("SpinnerMinimum"))
                  'lngSpinnerMax = IIf(IsNull(.Fields("SpinnerMaximum")), 0, .Fields("SpinnerMaximum"))
                
                 End If
              End If
            End With
          End If

          If fValid13 Then
            ' Check StoredData Column and DBValue Column match in terms of type, size, decs.
            Select Case iDBColumnDataType
              Case sqlBoolean 'Logic
                fValid12 = (iColumnDataType = sqlBoolean)
                If Not fValid12 Then
                  sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                    ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                    " - invalid data type)"
                  sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                End If
                
              Case sqlLongVarChar 'Working Pattern
                fValid12 = (iColumnDataType = sqlLongVarChar)
                If Not fValid12 Then
                  sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                    ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                    " - invalid data type)"
                  sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                End If
              
              Case sqlNumeric 'Numeric
                fValid12 = (iColumnDataType = sqlNumeric) _
                  Or (iColumnDataType = sqlInteger)
                  
                If Not fValid12 Then
                  sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                    ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                    " - invalid data type)"
                  sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                Else
                  ' Check size, decs
                  If iColumnDataType = sqlNumeric Then
                    ' Column is Numeric
                    fValid12 = (iDBColumnDecimals <= iColumnDecimals) _
                      And ((iDBColumnSize - iDBColumnDecimals) <= (iColumnSize - iColumnDecimals))
                  
                    If Not fValid12 Then
                      sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                        ", size = " & CStr(iDBColumnSize) & ", decimals = " & CStr(iDBColumnDecimals) & _
                        " - invalid size & decimals)"
                      sSubMessage3 = ", size = " & CStr(iColumnSize) & ", decimals = " & CStr(iColumnDecimals)
                    End If
                  Else
                    ' Column is Integer
                    fValid12 = (iDBColumnDecimals = 0) _
                      And (iDBColumnSize <= 9)
                  
                    If Not fValid12 Then
                      sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                        ", size = " & CStr(iDBColumnSize) & ", decimals = " & CStr(iDBColumnDecimals) & _
                        " - invalid size & decimals)"
                      sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                    End If
                  End If
                End If

              Case sqlInteger 'Integer
                fValid12 = (iColumnDataType = sqlNumeric) _
                  Or (iColumnDataType = sqlInteger)

                If Not fValid12 Then
                  sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                    ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                    " - invalid data type)"
                  sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                ElseIf (iColumnDataType = sqlNumeric) Then
                  ' Check size, decs
                  fValid12 = (iColumnSize - iColumnDecimals <= 9)
                
                  If Not fValid12 Then
                    sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                      ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                      " - invalid size)"
                    sSubMessage3 = ", size = " & CStr(iColumnSize) & ", decimals = " & CStr(iColumnDecimals)
                  End If
                End If

              Case sqlDate 'Date
                fValid12 = (iColumnDataType = sqlDate)
                If Not fValid12 Then
                  sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                    ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                    " - invalid data type)"
                  sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                End If
              
              Case sqlVarChar 'Character
                fValid12 = (iColumnDataType = sqlVarChar) _
                  Or (iColumnDataType = sqlLongVarChar)
                  
                If Not fValid12 Then
                  sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                    ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                    " - invalid data type)"
                  sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                Else
                  ' Check size
                  If iColumnDataType = sqlVarChar Then
                    ' Column is Character
                    fValid12 = (iColumnSize >= iDBColumnSize)
                  Else
                    ' Column is WorkingPattern
                    fValid12 = (iDBColumnSize <= 14)
                  End If
                
                  If Not fValid12 Then
                    sSubMessage2 = " (" & GetColumnName(CLng(avColumns(8, iLoop2)), False) & _
                      ", size = " & CStr(iDBColumnSize) & _
                      " - invalid size)"
                    sSubMessage3 = ", size = " & CStr(iColumnSize)
                  End If
                End If
            
            End Select
            
            ' Validate the DBValue record identifier
            Select Case avColumns(9, iLoop2)
              Case giWFRECSEL_UNKNOWN           ' Oops, something's wrong
                fValid14 = False

              Case giWFRECSEL_INITIATOR           ' Initiator's personnel table record
                ' DBValue must be based on the Personnel table.
                ReDim alngValidTables(0)
                TableAscendants mlngPersonnelTableID, alngValidTables
                
                fFound = False
                For iLoop5 = 1 To UBound(alngValidTables)
                  If alngValidTables(iLoop5) = lngDBColumnTableID Then
                    fFound = True
                    Exit For
                  End If
                Next iLoop5
          
                fValid14 = fFound

              Case giWFRECSEL_TRIGGEREDRECORD     ' Base table record
                ' DBValue must be based on the base table.
                ReDim alngValidTables(0)
                TableAscendants mlngBaseTableID, alngValidTables
                
                fFound = False
                For iLoop5 = 1 To UBound(alngValidTables)
                  If alngValidTables(iLoop5) = lngDBColumnTableID Then
                    fFound = True
                    Exit For
                  End If
                Next iLoop5
          
                fValid14 = fFound

              Case giWFRECSEL_IDENTIFIEDRECORD    ' Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
                ' Check identification is valid.
                fValid15 = (Len(Trim(avColumns(6, iLoop2))) > 0)

                If fValid15 Then
                  fValid15 = False
                  
                  For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
                    Set wfTempElement = aWFPrecedingElements(iLoop)

                    If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(avColumns(6, iLoop2))) Then
                      If wfTempElement.ElementType = elem_WebForm Then
                        fValid15 = True

                        fValid16 = (Len(Trim(avColumns(7, iLoop2))) > 0)
                        If fValid16 Then
                          fValid16 = False

                          asItems = wfTempElement.Items

                          For iLoop3 = 1 To UBound(asItems, 2)
                            If (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_GRID) _
                              And UCase(Trim(asItems(9, iLoop3))) = UCase(Trim(avColumns(7, iLoop2))) Then

                              ReDim alngValidTables(0)
                              TableAscendants CLng(asItems(44, iLoop3)), alngValidTables
                              
                              fFound = False
                              For iLoop5 = 1 To UBound(alngValidTables)
                                If alngValidTables(iLoop5) = lngDBColumnTableID Then
                                  fFound = True
                                  Exit For
                                End If
                              Next iLoop5
                        
                              fValid16 = fFound
                              Exit For
                            End If
                          Next iLoop3
                        End If
                        Exit For
                      ElseIf wfTempElement.ElementType = elem_StoredData Then
                        ReDim alngValidTables(0)
                        TableAscendants wfTempElement.DataTableID, alngValidTables
                        
                        'JPD 20061227
                        'If wfTempElement.DataAction = DATAACTION_DELETE Then
                        '  ' Cannot do anything with a Deleted record, but can use its ascendants.
                        '  ' Remove the table itself from the array of valid tables.
                        '  alngValidTables(1) = 0
                        'End If
                        
                        fFound = False
                        For iLoop5 = 1 To UBound(alngValidTables)
                          If alngValidTables(iLoop5) = lngDBColumnTableID Then
                            fFound = True
                            Exit For
                          End If
                        Next iLoop5
                  
                        fValid15 = fFound
                        Exit For
                      End If
                    End If
                    Set wfTempElement = Nothing
                  Next iLoop
                End If

              Case giWFRECSEL_ALL                 ' Show all records from the table in a WebForm RecordSelector
                ' Not used for DBValue selection.
                fValid14 = False

              Case giWFRECSEL_UNIDENTIFIED        ' Used when StoredData Inserts into a top-level table
                ' Not used for DBValue selection.
                fValid14 = False
            End Select
          End If

          '------------------------------------------------------------
          ' Add the required validation messages to the array.
          '------------------------------------------------------------
          ' 12. StoredData element columns (DBValue) must be the correct type, size, etc.
          ' 13. StoredData element columns (DBValue) must have valid column.
          ' 14. StoredData element columns (DBValue) must have valid record.
          ' 15. StoredData element columns (DBValue) must have valid record element identifier (where required).
          ' 16. StoredData element columns (DBValue) must have valid record selector identifier (where required).
          If (Not fValid12) And (Not fDoingDeleteCheck) Then
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Column (" & sSubMessage1 & sSubMessage3 & ") - Invalid database value" & sSubMessage2, _
              pwfElement.ControlIndex
          End If
          If (Not fValid13) And (Not fDoingDeleteCheck) Then
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Column (" & sSubMessage1 & sSubMessage3 & ") - Invalid database value", _
              pwfElement.ControlIndex
          End If
          If (Not fValid14) And (Not fDoingDeleteCheck) Then
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Column (" & sSubMessage1 & sSubMessage3 & ") - Invalid database value record", _
              pwfElement.ControlIndex
          End If
          If Not fValid15 Then
            avColumns(6, iLoop2) = ""
            avColumns(7, iLoop2) = ""

            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Column (" & sSubMessage1 & sSubMessage3 & ") - Invalid database value element identifier", _
              pwfElement.ControlIndex
          End If
          If Not fValid16 Then
            avColumns(7, iLoop2) = ""
  
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Column (" & sSubMessage1 & sSubMessage3 & ") - Invalid database value record selector", _
              pwfElement.ControlIndex
          End If

        Case giWFDATAVALUE_WFVALUE
          ' 12. StoredData element columns (WFValue - lookups) must be the correct type, size, etc.
          ' 17. StoredData element columns (WFValue) must have valid WebForm identifier.
          ' 18. StoredData element columns (WFValue) must have valid WebForm InputValue identifier.
          ' 19. StoredData element columns (WFValue) must have valid WebForm InputValue identifier with respect to size and decimals.
          fValid12 = True
          fValid17 = (Len(Trim(avColumns(6, iLoop2))) > 0)
          fValid18 = True
          fValid19 = True

          If fValid17 Then
            fValid17 = False
            
            For iLoop = 2 To UBound(aWFPrecedingElements) ' Ignore index 1 as that is the current element
              Set wfTempElement = aWFPrecedingElements(iLoop)

              If UCase(Trim(wfTempElement.Identifier)) = UCase(Trim(avColumns(6, iLoop2))) Then
                If wfTempElement.ElementType = elem_WebForm Then
                  fValid17 = True
  
                  sMessagePrefix2 = ValidateElement_MessagePrefix(wfTempElement)
                  lngTempElementIndex = wfTempElement.ControlIndex
                  
                  fValid18 = (Len(Trim(avColumns(7, iLoop2))) > 0)
                  If fValid18 Then
                    fValid18 = False

                    asItems = wfTempElement.Items

                    For iLoop3 = 1 To UBound(asItems, 2)
                      If ((asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_CHAR) _
                          Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_NUMERIC) _
                          Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_LOGIC) _
                          Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_DATE) _
                          Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_DROPDOWN) _
                          Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_LOOKUP) _
                          Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) _
                          Or (asItems(2, iLoop3) = giWFFORMITEM_INPUTVALUE_FILEUPLOAD)) _
                        And UCase(Trim(asItems(9, iLoop3))) = UCase(Trim(avColumns(7, iLoop2))) Then
                      
                        fValid18 = True

                        ' Check type, size, etc. are okay.
                        Select Case asItems(2, iLoop3)
                          Case giWFFORMITEM_INPUTVALUE_LOGIC 'Logic
                            fValid18 = (iColumnDataType = sqlBoolean)
                            If Not fValid18 Then
                              sSubMessage2 = " (" & wfTempElement.Caption & "." & asItems(9, iLoop3) & _
                                ", data type = " & GetWebFormItemTypeName(CInt(asItems(2, iLoop3))) & _
                                " - invalid data type)"
                              sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                            End If
    
                          Case giWFFORMITEM_INPUTVALUE_NUMERIC 'Numeric
                            fValid18 = (iColumnDataType = sqlNumeric) _
                              Or (iColumnDataType = sqlInteger)
                            If Not fValid18 Then
                              sSubMessage2 = " (" & wfTempElement.Caption & "." & asItems(9, iLoop3) & _
                                ", data type = " & GetWebFormItemTypeName(CInt(asItems(2, iLoop3))) & _
                                " - invalid data type)"
                              sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                            Else
                              ' Check size, decs
                              If iColumnDataType = sqlNumeric Then
                                ' Column is Numeric
                                fValid19 = (asItems(8, iLoop3) <= iColumnDecimals) _
                                  And ((asItems(7, iLoop3) - asItems(8, iLoop3)) <= (iColumnSize - iColumnDecimals))
                              
                                If Not fValid19 Then
                                  If pfFix Then
                                    fValid19 = True
                                    
                                    If (asItems(8, iLoop3) > iColumnDecimals) Then
                                      asItems(8, iLoop3) = iColumnDecimals
                                    End If
                                  
                                    If ((asItems(7, iLoop3) - asItems(8, iLoop3)) > (iColumnSize - iColumnDecimals)) Then
                                      asItems(7, iLoop3) = iColumnSize
                                    
                                      If (asItems(8, iLoop3) < iColumnDecimals) Then
                                        asItems(8, iLoop3) = iColumnDecimals
                                      End If
                                    End If
                                    
                                    wfTempElement.Items = asItems
                                  Else
                                    sSubMessage4 = "Numeric Input (" & asItems(9, iLoop3) & _
                                      ", size = " & asItems(7, iLoop3) & ", decimals = " & asItems(8, iLoop3) & ") - invalid size & decimals for " & _
                                      pwfElement.ElementTypeDescription & " (" & pwfElement.Caption & ") column (" & GetColumnName(CLng(avColumns(3, iLoop2)), False) & ", size = " & CStr(iColumnSize) & ", decimals = " & CStr(iColumnDecimals) & ")"
                                    mfFixableValidationFailures = True
                                  End If
                                End If
                              Else
                                ' Column is Integer
                                fValid19 = (asItems(8, iLoop3) = 0) _
                                  And (asItems(7, iLoop3) <= 9)
                              
                                If Not fValid19 Then
                                  If pfFix Then
                                    fValid19 = True
                                    
                                    If (asItems(8, iLoop3) > 0) Then
                                      asItems(8, iLoop3) = 0
                                    End If
                                    
                                    If (asItems(7, iLoop3) > 9) Then
                                      asItems(7, iLoop3) = 9
                                    End If
                                    
                                    wfTempElement.Items = asItems
                                  Else
                                    sSubMessage4 = "Numeric Input (" & asItems(9, iLoop3) & _
                                      ", size = " & asItems(7, iLoop3) & ", decimals = " & asItems(8, iLoop3) & ") - invalid size & decimals for " & _
                                      pwfElement.ElementTypeDescription & " (" & pwfElement.Caption & ") column (" & GetColumnName(CLng(avColumns(3, iLoop2)), False) & ", data type = " & GetDataTypeName(iColumnDataType) & ")"
                                    mfFixableValidationFailures = True
                                  End If
                                End If
                              End If
                            End If
                        
                          Case giWFFORMITEM_INPUTVALUE_DATE 'Date
                            fValid18 = (iColumnDataType = sqlDate)
                            If Not fValid18 Then
                              sSubMessage2 = " (" & wfTempElement.Caption & "." & asItems(9, iLoop3) & _
                                ", data type = " & GetWebFormItemTypeName(CInt(asItems(2, iLoop3))) & _
                                " - invalid data type)"
                              sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                            End If
  
                          Case giWFFORMITEM_INPUTVALUE_CHAR 'Character
                            fValid18 = (iColumnDataType = sqlVarChar) _
                              Or (iColumnDataType = sqlLongVarChar)

                            If Not fValid18 Then
                              sSubMessage2 = " (" & wfTempElement.Caption & "." & asItems(9, iLoop3) & _
                                ", data type = " & GetWebFormItemTypeName(CInt(asItems(2, iLoop3))) & _
                                " - invalid data type)"
                              sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                            Else
                              ' Check size
                              If iColumnDataType = sqlVarChar Then
                                ' Column is Character
                                fValid19 = (iColumnSize >= asItems(7, iLoop3))
                              Else
                                ' Column is WorkingPattern
                                fValid19 = (asItems(7, iLoop3) <= 14)
                              End If
        
                              If Not fValid19 Then
                                If pfFix Then
                                  fValid19 = True
                                  
                                  If iColumnDataType = sqlVarChar Then
                                    ' Column is Character
                                    If (asItems(7, iLoop3) > iColumnSize) Then
                                      asItems(7, iLoop3) = iColumnSize
                                    End If
                                  Else
                                    ' Column is WorkingPattern
                                    If (asItems(7, iLoop3) > 14) Then
                                      asItems(7, iLoop3) = 14
                                    End If
                                  End If
                                  wfTempElement.Items = asItems
                                Else
                                  sSubMessage4 = "Character Input (" & asItems(9, iLoop3) & _
                                    ", size = " & asItems(7, iLoop3) & ") - invalid size for " & _
                                    pwfElement.ElementTypeDescription & " (" & pwfElement.Caption & ") column (" & GetColumnName(CLng(avColumns(3, iLoop2)), False) & ", size = " & CStr(iColumnSize) & ")"
                                  mfFixableValidationFailures = True
                                End If
                              End If
                            End If
                        
                          Case giWFFORMITEM_INPUTVALUE_DROPDOWN
                            fValid18 = (iColumnDataType = sqlVarChar) _
                              Or (iColumnDataType = sqlLongVarChar)

                            If Not fValid18 Then
                              sSubMessage2 = " (" & wfTempElement.Caption & "." & asItems(9, iLoop3) & _
                                ", data type = " & GetWebFormItemTypeName(CInt(asItems(2, iLoop3))) & _
                                " - invalid data type)"
                              sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                            Else
                              ' Check size
                              iMaxLength = 0
                              asItemValues = Split(asItems(47, iLoop3), vbTab)
                              
                              For iLoop4 = 0 To UBound(asItemValues)
                                If Len(asItemValues(iLoop4)) > iMaxLength Then
                                  iMaxLength = Len(asItemValues(iLoop4))
                                End If
                              Next iLoop4
                              
                              'asItems(7, iLoop3)
                              If iColumnDataType = sqlVarChar Then
                                ' Column is Character
                                fValid19 = (iColumnSize >= iMaxLength)
                              Else
                                ' Column is WorkingPattern
                                fValid19 = (iMaxLength <= 14)
                              End If

                              If Not fValid19 Then
                                sSubMessage4 = "Dropdown Input (" & asItems(9, iLoop3) & _
                                  ", size = " & iMaxLength & ") - invalid size for " & _
                                  pwfElement.ElementTypeDescription & " (" & pwfElement.Caption & ") column (" & GetColumnName(CLng(avColumns(3, iLoop2)), False) & ", size = " & CStr(iColumnSize) & ")"
                              End If
                            End If
                          
                          Case giWFFORMITEM_INPUTVALUE_LOOKUP
                            With recColEdit
                              .Index = "idxColumnID"
                              .Seek "=", CLng(asItems(49, iLoop3))
                
                              If Not .NoMatch Then
                                If Not (!Deleted _
                                  Or (!DataType = dtLONGVARBINARY) _
                                  Or (!DataType = dtVARBINARY) _
                                  Or (!columntype = giCOLUMNTYPE_LINK) _
                                  Or (!columntype = giCOLUMNTYPE_SYSTEM)) Then
                                  
                                  iDBColumnDataType = !DataType
                                  iDBColumnSize = !Size
                                  iDBColumnDecimals = !Decimals
                                  lngDBColumnTableID = !TableID
                                  'lngSpinnerMin = IIf(IsNull(.Fields("SpinnerMinimum")), 0, .Fields("SpinnerMinimum"))
                                  'lngSpinnerMax = IIf(IsNull(.Fields("SpinnerMaximum")), 0, .Fields("SpinnerMaximum"))
                                
                                  ' Check StoredData Column and DBValue Column match in terms of type, size, decs.
                                  Select Case iDBColumnDataType
                                    Case sqlBoolean 'Logic
                                      fValid12 = (iColumnDataType = sqlBoolean)
                                      If Not fValid12 Then
                                        sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                          ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                                          " - invalid data type)"
                                        sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                                      End If
    
                                    Case sqlLongVarChar 'Working Pattern
                                      fValid12 = (iColumnDataType = sqlLongVarChar)
                                      If Not fValid12 Then
                                        sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                          ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                                          " - invalid data type)"
                                        sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                                      End If
  
                                    Case sqlNumeric 'Numeric
                                      fValid12 = (iColumnDataType = sqlNumeric) _
                                        Or (iColumnDataType = sqlInteger)
                                        
                                      If Not fValid12 Then
                                        sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                          ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                                          " - invalid data type)"
                                        sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                                      Else
                                        ' Check size, decs
                                        If iColumnDataType = sqlNumeric Then
                                          ' Column is Numeric
                                          fValid12 = (iDBColumnDecimals <= iColumnDecimals) _
                                            And ((iDBColumnSize - iDBColumnDecimals) <= (iColumnSize - iColumnDecimals))
                                        
                                          If Not fValid12 Then
                                            sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                              ", size = " & CStr(iDBColumnSize) & ", decimals = " & CStr(iDBColumnDecimals) & _
                                              " - invalid size & decimals)"
                                            sSubMessage3 = ", size = " & CStr(iColumnSize) & ", decimals = " & CStr(iColumnDecimals)
                                          End If
                                        Else
                                          ' Column is Integer
                                          fValid12 = (iDBColumnDecimals = 0) _
                                            And (iDBColumnSize <= 9)
                                        
                                          If Not fValid12 Then
                                            sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                              ", size = " & CStr(iDBColumnSize) & ", decimals = " & CStr(iDBColumnDecimals) & _
                                              " - invalid size & decimals)"
                                            sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                                          End If
                                        End If
                                      End If

                                    Case sqlInteger 'Integer
                                      fValid12 = (iColumnDataType = sqlNumeric) _
                                        Or (iColumnDataType = sqlInteger)
                                  
                                      If Not fValid12 Then
                                        sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                          ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                                          " - invalid data type)"
                                        sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                                      ElseIf (iColumnDataType = sqlNumeric) Then
                                        ' Check size, decs
                                        fValid12 = (iColumnSize - iColumnDecimals <= 9)
    
                                        If Not fValid12 Then
                                          sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                            ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                                            " - invalid size)"
                                          sSubMessage3 = ", size = " & CStr(iColumnSize) & ", decimals = " & CStr(iColumnDecimals)
                                        End If
                                      End If

                                    Case sqlDate 'Date
                                      fValid12 = (iColumnDataType = sqlDate)
                                      If Not fValid12 Then
                                        sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                          ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                                          " - invalid data type)"
                                        sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                                      End If
                                    
                                    Case sqlVarChar 'Character
                                      fValid12 = (iColumnDataType = sqlVarChar) _
                                        Or (iColumnDataType = sqlLongVarChar)
                                        
                                      If Not fValid12 Then
                                        sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                          ", data type = " & GetDataTypeName(iDBColumnDataType) & _
                                          " - invalid data type)"
                                        sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                                      Else
                                        ' Check size
                                        If iColumnDataType = sqlVarChar Then
                                          ' Column is Character
                                          fValid12 = (iColumnSize >= iDBColumnSize)
                                        Else
                                          ' Column is WorkingPattern
                                          fValid12 = (iDBColumnSize <= 14)
                                        End If
                                      
                                        If Not fValid12 Then
                                          sSubMessage2 = " (" & GetColumnName(CLng(asItems(49, iLoop3)), False) & _
                                            ", size = " & CStr(iDBColumnSize) & _
                                            " - invalid size)"
                                          sSubMessage3 = ", size = " & CStr(iColumnSize)
                                        End If
                                      End If
                                  End Select
                                End If
                              End If
                            End With

                          Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP
                            fValid18 = (iColumnDataType = sqlVarChar) _
                              Or (iColumnDataType = sqlLongVarChar)

                            If Not fValid18 Then
                              sSubMessage2 = " (" & wfTempElement.Caption & "." & asItems(9, iLoop3) & _
                                ", data type = " & GetWebFormItemTypeName(CInt(asItems(2, iLoop3))) & _
                                " - invalid data type)"
                              sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                            Else
                              ' Check size
                              iMaxLength = 0
                              asItemValues = Split(asItems(47, iLoop3), vbTab)
                              
                              For iLoop4 = 0 To UBound(asItemValues)
                                If Len(asItemValues(iLoop4)) > iMaxLength Then
                                  iMaxLength = Len(asItemValues(iLoop4))
                                End If
                              Next iLoop4
                              
                              'asItems(7, iLoop3)
                              If iColumnDataType = sqlVarChar Then
                                ' Column is Character
                                fValid19 = (iColumnSize >= iMaxLength)
                              Else
                                ' Column is WorkingPattern
                                fValid19 = (iMaxLength <= 14)
                              End If

                              If Not fValid19 Then
                                sSubMessage4 = "Option Group Input (" & asItems(9, iLoop3) & _
                                  ", size = " & iMaxLength & ") - invalid size for " & _
                                  pwfElement.ElementTypeDescription & " (" & pwfElement.Caption & ") column (" & GetColumnName(CLng(avColumns(3, iLoop2)), False) & ", size = " & CStr(iColumnSize) & ")"
                              End If
                            End If
                        
                          Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD 'FileUpload
                            fValid18 = (iColumnDataType = sqlOle) _
                              Or (iColumnDataType = sqlVarBinary)
                            
                            If Not fValid18 Then
                              sSubMessage2 = " (" & wfTempElement.Caption & "." & asItems(9, iLoop3) & _
                                ", data type = " & GetWebFormItemTypeName(CInt(asItems(2, iLoop3))) & _
                                " - invalid data type)"
                              sSubMessage3 = ", data type = " & GetDataTypeName(iColumnDataType)
                            End If
                        End Select

                        Exit For
                      End If
                    Next iLoop3
                  End If
                  
                  Exit For
                End If
              End If
              
              Set wfTempElement = Nothing
            Next iLoop
          End If

          '------------------------------------------------------------
          ' Add the required validation messages to the array.
          '------------------------------------------------------------
          ' 12. StoredData element columns (WFValue - lookups) must be the correct type, size, etc.
          ' 17. StoredData element columns (WFValue) must have valid WebForm identifier.
          ' 18. StoredData element columns (WFValue) must have valid WebForm InputValue identifier.
          ' 19. StoredData element columns (WFValue) must have valid WebForm InputValue identifier with respect to size and decimals.
          If (Not fValid12) And (Not fDoingDeleteCheck) Then
            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Column (" & sSubMessage1 & sSubMessage3 & ") - Invalid workflow value lookup column" & sSubMessage2, _
              pwfElement.ControlIndex
          End If
          If Not fValid17 Then
            avColumns(6, iLoop2) = ""
            avColumns(7, iLoop2) = ""

            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Column (" & sSubMessage1 & sSubMessage3 & ") - Invalid workflow value web form identifier", _
              pwfElement.ControlIndex
          End If
          If Not fValid18 Then
            avColumns(7, iLoop2) = ""

            ValidateWorkflow_AddMessage _
              sMessagePrefix & "Column (" & sSubMessage1 & sSubMessage3 & ") -  Invalid workflow value identifier" & sSubMessage2, _
              pwfElement.ControlIndex
          End If
          If (Not fValid19) And (Not fDoingDeleteCheck) Then
            ValidateWorkflow_AddMessage _
              sMessagePrefix2 & sSubMessage4, _
              lngTempElementIndex
          End If
      
        Case giWFDATAVALUE_CALC
          If (CLng(avColumns(10, iLoop2)) > 0) Then
            ValidateElement_Expression _
              pwfElement, _
              CLng(avColumns(10, iLoop2)), _
              "Column (" & sSubMessage1 & sSubMessage3 & ") - Invalid column calculation" & sSubMessage2, _
              pavarDisconnectedElements
          Else
            ValidateWorkflow_AddMessage _
              "Column (" & sSubMessage1 & sSubMessage3 & ") - No calculation selected", _
              pwfElement.ControlIndex
          End If
        
      End Select
    End If
  
    '------------------------------------------------------------
    ' Add the required validation messages to the array.
    '------------------------------------------------------------
    ' 11. StoredData element columns must be for a valid column in the table.
    If (Not fValid11) And (Not fDoingDeleteCheck) Then
      ValidateWorkflow_AddMessage _
        sMessagePrefix & " Invalid column", _
        pwfElement.ControlIndex
    End If
  Next iLoop2

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
End Sub


Private Function ValidateElement_MessagePrefix(pwfElement As VB.Control) As String
  Dim sPrefix As String

  sPrefix = pwfElement.ElementTypeDescription & " (" & pwfElement.Caption & ") : "
  
  ValidateElement_MessagePrefix = sPrefix
  
End Function

Private Function ElementExists(plngIndex As Long) As Boolean
  ' Return tru if there exists a Workflow element with the given index.
  On Error GoTo ErrorTrap
  
  Dim wfTemp As VB.Control
  
  wfTemp = mcolwfElements(CStr(plngIndex))
  ElementExists = True
  Exit Function
  
ErrorTrap:
  ElementExists = False
  
End Function

Private Function ElementHasIdentifier(pwfElement As VB.Control) As Boolean
  ElementHasIdentifier = (pwfElement.ElementType = elem_StoredData) _
    Or (pwfElement.ElementType = elem_WebForm)
  
End Function

Private Sub FormatLink(pwfLink As COAWF_Link)
  ' Format and position the link to fit the 'linked' elements
  Dim wfStartElement As VB.Control
  Dim wfEndElement As VB.Control
  Dim sngStartXOffset As Single
  Dim sngStartYOffset As Single
  Dim sngEndXOffset As Single
  Dim sngEndYOffset As Single
  Dim avOutboundFlowInfo() As Variant
  Dim iOutboundFlowIndex As Integer
  Dim iLoop As Integer
  
  Set wfStartElement = mcolwfElements(CStr(pwfLink.StartElementIndex))
  Set wfEndElement = mcolwfElements(CStr(pwfLink.EndElementIndex))

  If (wfStartElement Is Nothing) Or (wfEndElement Is Nothing) Then
  
    If pwfLink.Highlighted Then
      mcolwfSelectedLinks.Remove CStr(pwfLink.Index)
    End If
        
    UnLoad pwfLink
    Exit Sub
  End If
  
  ' Get the array of outbound flow information from the start element.
  ' Column 1 = Tag (see enums, or -1 if there's only a single outbound flow)
  ' Column 2 = Direction
  ' Column 3 = XOffset
  ' Column 4 = YOffset
  ' Column 5 = Maximum
  ' Column 6 = Minimum
  ' Column 7 = Description
  avOutboundFlowInfo = wfStartElement.OutboundFlows_Information
  
  If pwfLink.StartOutboundFlowCode < 0 Then
    iOutboundFlowIndex = 1
  Else
    For iLoop = 1 To UBound(avOutboundFlowInfo, 2)
      If avOutboundFlowInfo(1, iLoop) = pwfLink.StartOutboundFlowCode Then
        iOutboundFlowIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
  sngStartXOffset = wfStartElement.Left + avOutboundFlowInfo(3, iOutboundFlowIndex)
  sngStartYOffset = wfStartElement.Top + avOutboundFlowInfo(4, iOutboundFlowIndex)
  sngEndXOffset = wfEndElement.Left + wfEndElement.InboundFlow_XOffset
  sngEndYOffset = wfEndElement.Top + wfEndElement.InboundFlow_YOffset

  With pwfLink
    .CurvedLinks = True
    
    .XOffset = sngEndXOffset - sngStartXOffset
    .YOffset = sngEndYOffset - sngStartYOffset
        
    .Left = sngStartXOffset - .StartXOffset
    .Top = sngStartYOffset - .StartYOffset
  End With
  
End Sub

Private Function GetElementByIdentifier(psIdentifier As Variant) As VB.Control
  ' Return the element with the given identifier.
  Dim wfTemp As VB.Control
  Dim iLoop As Integer
  Dim fElementOK As Boolean
  
  If Len(Trim(psIdentifier)) = 0 Then
    Exit Function
  End If
  
  For Each wfTemp In mcolwfElements
    With wfTemp
      If (UCase(Trim(.Identifier)) = UCase(Trim(psIdentifier))) Then
        fElementOK = True
        
        If (miLastActionFlag = giACTION_DELETECONTROLS) Then
          For iLoop = 1 To UBound(mactlUndoControls)
            If IsWorkflowElement(mactlUndoControls(iLoop)) Then
              If mactlUndoControls(iLoop).ControlIndex = .ControlIndex Then
                fElementOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
        
        'JPD 20060719 Fault 11339
        If fElementOK Then
          For iLoop = 1 To UBound(mactlClipboardControls)
            If IsWorkflowElement(mactlClipboardControls(iLoop)) Then
              If mactlClipboardControls(iLoop).ControlIndex = .ControlIndex Then
                fElementOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
        
        If fElementOK Then
          Set GetElementByIdentifier = wfTemp
          Exit For
        End If
      End If
    End With
  Next wfTemp
  
  Set wfTemp = Nothing
  
End Function

Private Function GetUniqueIdentifier(pwfElement As VB.Control) As String
  Dim wfElement As VB.Control
  Dim sIdentifier As String
  Dim iMaxSuffix As Integer
  Dim iSuffix As Integer
  Dim fRootIdentifierFound As Boolean
  Dim sIdentifierRoot As String
  
  Select Case pwfElement.ElementType
    Case elem_WebForm
      sIdentifierRoot = "Web Form"
    Case elem_StoredData
      sIdentifierRoot = "Stored Data"
  End Select

  iMaxSuffix = 0
  fRootIdentifierFound = False
  
  For Each wfElement In mcolwfElements
    If TypeOf wfElement Is COAWF_Webform _
      Or TypeOf wfElement Is COAWF_StoredData Then
      
      If wfElement.ControlIndex <> pwfElement.ControlIndex _
        And ElementHasIdentifier(wfElement) _
        And Left(wfElement.Identifier, Len(sIdentifierRoot)) = sIdentifierRoot Then
  
        If UCase(wfElement.Identifier) = UCase(sIdentifierRoot) Then
          fRootIdentifierFound = True
        End If
        
        iSuffix = val(Mid(wfElement.Identifier, Len(sIdentifierRoot) + 1))
        
        If iSuffix > iMaxSuffix Then
          iMaxSuffix = iSuffix
        End If
      End If
      
    End If
  Next wfElement
  Set wfElement = Nothing

  If Not fRootIdentifierFound Then
    GetUniqueIdentifier = sIdentifierRoot
  Else
    GetUniqueIdentifier = sIdentifierRoot & " " & CStr(iMaxSuffix + 1)
  End If
  
End Function

Private Function NextConnectorCaption() As String
  Dim iTemp As Integer
  Dim wfElement As VB.Control
  Dim fDoCheck As Boolean
  Dim sMaxConnectorCaption As String
  
  sMaxConnectorCaption = "A"
  
  ' Ensure the caption has not already been used.
  fDoCheck = True
  Do While fDoCheck
    fDoCheck = False
    
    For Each wfElement In mcolwfElements
      If (wfElement.Visible) And (wfElement.ElementType = elem_Connector1) Then
        If wfElement.Caption = sMaxConnectorCaption Then
          iTemp = Asc(sMaxConnectorCaption) + 1
          sMaxConnectorCaption = Chr(iTemp)
          
          fDoCheck = True
        End If
      End If
    Next wfElement
    Set wfElement = Nothing
  Loop
  
  NextConnectorCaption = sMaxConnectorCaption
  
End Function


Private Sub PrintOffPageConnector(piLinkIndex As Integer, _
  pSngX As Single, _
  pSngY As Single, _
  pfNeedsArrow As Boolean, _
  pArrowDirection As LineDirection)
  
  Dim iLoop As Integer
  Dim fFound As Boolean
  Dim sOffPageCharacter As String
  Dim sngOffPageXOffset As Single
  Dim sngOffPageYOffset As Single
  Dim rctDraw As Rect
  Dim sngXArrowOffset As Single
  Dim sngYArrowOffset As Single
  
  Const CAPTIONOFFSET = 20
  
  sngOffPageXOffset = IIf(pArrowDirection = lineDirection_Left Or pArrowDirection = lineDirection_Right, 0, (picOffPage.Width / 2))
  sngOffPageYOffset = IIf(pArrowDirection = lineDirection_Down Or pArrowDirection = lineDirection_Up, 0, (picOffPage.Height / 2))
            
  ' Print the offPage connector image
  Printer.PaintPicture picOffPage.Picture, pSngX - sngOffPageXOffset, pSngY - sngOffPageYOffset
          
  fFound = False
  For iLoop = 1 To UBound(mavOffPageLinks, 2)
    If mavOffPageLinks(0, iLoop) = piLinkIndex Then
      fFound = True
      sOffPageCharacter = mavOffPageLinks(1, iLoop)
      Exit For
    End If
  Next iLoop
            
  If Not fFound Then
    ReDim Preserve mavOffPageLinks(1, UBound(mavOffPageLinks, 2) + 1)
  
    mavOffPageLinks(0, UBound(mavOffPageLinks, 2)) = piLinkIndex
    mavOffPageLinks(1, UBound(mavOffPageLinks, 2)) = msOffPageCharacter
  
    sOffPageCharacter = msOffPageCharacter
    
    msOffPageCharacter = Chr(Asc(msOffPageCharacter) + 1)
  End If
  
  ' AE20080529 Fault #13182
  'Printer.Font.Name = mcolwfElements(CStr(0)).Font.Name
  'Printer.Font.Size = mcolwfElements(CStr(0)).Font.Size
  Printer.Font.Name = mcolwfElements(CStr(1)).Font.Name
  Printer.Font.Size = mcolwfElements(CStr(1)).Font.Size
  Printer.Font.Bold = False
  Printer.Font.Underline = False
  Printer.Font.Strikethrough = False
  Printer.Font.Italic = False
  
  ' Print the link arrow (if required)
'  If pfNeedsArrow Then
'    ASRWFLink1(0).EndDirection = pArrowDirection
'
'    sngXArrowOffset = 0
'    sngYArrowOffset = 0
'
'    Select Case pArrowDirection
'      Case lineDirection_Down
'        sngXArrowOffset = -(ScaleX(ASRWFLink1(0).ArrowPicture.Width, vbHimetric, vbTwips) / 2)
'        sngYArrowOffset = picOffPage.Height
'      Case lineDirection_Left
'        sngXArrowOffset = -ScaleX(ASRWFLink1(0).ArrowPicture.Width, vbHimetric, vbTwips)
'        sngYArrowOffset = -(ScaleY(ASRWFLink1(0).ArrowPicture.Height, vbHimetric, vbTwips) / 2)
'      Case lineDirection_Right
'        sngXArrowOffset = picOffPage.Width
'        sngYArrowOffset = -(ScaleY(ASRWFLink1(0).ArrowPicture.Height, vbHimetric, vbTwips) / 2)
'      Case Else
'        sngXArrowOffset = -(ScaleX(ASRWFLink1(0).ArrowPicture.Width, vbHimetric, vbTwips) / 2)
'        sngYArrowOffset = -ScaleY(ASRWFLink1(0).ArrowPicture.Height, vbHimetric, vbTwips)
'    End Select
'
'    Printer.PaintPicture ASRWFLink1(0).ArrowPicture, _
'      pSngX + sngXArrowOffset, _
'      pSngY + sngYArrowOffset
'  End If

  ' Print the offpage caption.
  rctDraw.Left = pSngX - sngOffPageXOffset + ((picOffPage.Width - Printer.TextWidth(sOffPageCharacter)) / 2)
  rctDraw.Right = rctDraw.Left + picOffPage.Width
  rctDraw.Top = pSngY - sngOffPageYOffset + ((picOffPage.Height - Printer.TextHeight(sOffPageCharacter)) / 2) - CAPTIONOFFSET
  rctDraw.Bottom = rctDraw.Top + picOffPage.Height
  
  ' Scale the print coordinates for use with the DrawText function.
  rctDraw.Left = Printer.ScaleY(rctDraw.Left, Printer.ScaleMode, vbPixels)
  rctDraw.Right = Printer.ScaleY(rctDraw.Right, Printer.ScaleMode, vbPixels)
  rctDraw.Top = Printer.ScaleY(rctDraw.Top, Printer.ScaleMode, vbPixels)
  rctDraw.Bottom = Printer.ScaleY(rctDraw.Bottom, Printer.ScaleMode, vbPixels)
  
  DrawText Printer.hDC, sOffPageCharacter, -1, rctDraw, c_DTDefFmt Or DT_LEFT

End Sub

Public Sub PrintWorkflow()
  ' Print the workflow definition.
  On Error GoTo ErrorTrap
  
  Dim objDefPrinter As cSetDfltPrinter
  Dim ctlTemp As Control
  Dim frmPrompt As frmWorkflowPrintOptions
  Dim fPrintDetails As Boolean
  Dim fPrintOverview As Boolean
  Dim fPrintCancelled As Boolean
  Dim awfElements() As VB.Control
  Dim wfElement As VB.Control
  Dim iLoop As Integer
  Dim iCount As Integer
  Dim sngMinX As Single
  Dim sngMaxX As Single
  Dim sngMinY As Single
  Dim sngMaxY As Single
  Dim fOnCurrentPage As Boolean
  Dim fOnPageRight As Boolean
  Dim fOnPageBelow As Boolean
  Dim fNeedPageRight As Boolean
  Dim fNeedPageBelow As Boolean
  Dim pgCurrentPage As Page
  Dim pgNextPage As Page
  Dim intMarginTop As Integer
  Dim intMarginBottom As Integer
  Dim intMarginLeft As Integer
  Dim intMarginRight As Integer
  Dim fPagePrintedOn As Boolean
  Dim sngFirstPageXOffset As Single
  Dim sngPageRightMinX As Single
  Dim sngPageBelowMinY As Single
  Dim fConnected As Boolean
  Dim wfLink As COAWF_Link
  Dim sURL As String
  Dim fOK As Boolean
  
  Const iGAPUNDERTITLE = 5 * TWIPSPERMM
    
  glngPageNum = 0
  sngMinX = -1
  sngMaxX = -1
  sngMinY = -1
  sngMaxY = -1
  fOK = True
  
  ' Find out if the user want to print an overview, details, or everything.
  Set frmPrompt = New frmWorkflowPrintOptions
  With frmPrompt
    .Show vbModal
  
    fPrintDetails = .PrintDetails
    fPrintOverview = .PrintOverview
    fPrintCancelled = .Cancelled
  End With
  Set frmPrompt = Nothing

  If fPrintCancelled Then
    Exit Sub
  End If
  
  ' Load the printer object
  Set mobjPrinter = New clsPrintDef
  With mobjPrinter
    fOK = .IsOK

    If fOK Then
      If .PrintStart(True) Then
        intMarginBottom = .MarginBottom
        intMarginTop = .MarginTop
        intMarginLeft = .MarginLeft
        intMarginRight = .MarginRight
      
        mlngMarginBottom_Twips = (intMarginBottom * TWIPSPERMM)
        mlngMarginTop_Twips = (intMarginTop * TWIPSPERMM) - MARGINCORRECTION
        mlngMarginLeft_Twips = (intMarginLeft * TWIPSPERMM) - MARGINCORRECTION
        mlngMarginRight_Twips = (intMarginRight * TWIPSPERMM) - MARGINCORRECTION
      
        If fPrintOverview Then
          msOffPageCharacter = "A"
          mlngRealBottom = CalculateBottomOfPage
          
          ' Index 0 = link index
          ' Index 1 = character
          ReDim mavOffPageLinks(1, 0)
          
          ' Print the header
          .PrintHeader "Workflow Definition : " & Trim(msWorkflowName)
  
          msngYOffset = Printer.CurrentY + iGAPUNDERTITLE + iGAPOFFPAGE
          msngTopGap = msngYOffset
  
          ' Get an array of the workflow elements in order of vertical, then horizontal position.
          ReDim awfElements(0)
          iCount = -1
          For Each ctlTemp In mcolwfElements
            If ctlTemp.ControlIndex > 0 Then
              iCount = iCount + 1
              ReDim Preserve awfElements(iCount)
              Set awfElements(iCount) = ctlTemp
            
              sngMinX = IIf(sngMinX < 0, ctlTemp.Left, IIf(sngMinX <= ctlTemp.Left, sngMinX, ctlTemp.Left))
              sngMaxX = IIf(sngMaxX < 0, ctlTemp.Left + ctlTemp.Width, IIf(sngMaxX >= ctlTemp.Left + ctlTemp.Width, sngMaxX, ctlTemp.Left + ctlTemp.Width))
              sngMinY = IIf(sngMinY < 0, ctlTemp.Top, IIf(sngMinY <= ctlTemp.Top, sngMinY, ctlTemp.Top))
              sngMaxY = IIf(sngMaxY < 0, ctlTemp.Top + ctlTemp.Height, IIf(sngMaxY >= ctlTemp.Top + ctlTemp.Height, sngMaxY, ctlTemp.Top + ctlTemp.Height))
            End If
          Next ctlTemp
          Set ctlTemp = Nothing
  
          ' Sort the array of elements into order by vertical position, then horizontal position
  ''' needed?
          ShellSortElements awfElements
  
          ' Centre the print out horizontally if possible.
          ' ie. if the image will not overflow horizontally onto another page.
          If (sngMaxX - sngMinX) >= (Printer.Width - mlngMarginLeft_Twips - mlngMarginRight_Twips - (2 * MARGINCORRECTION) - (2 * iGAPOFFPAGE)) Then
            ' Does not fit onto a single page horizontally, so push over to the left.
            msngXOffset = mlngMarginLeft_Twips + MARGINCORRECTION + iGAPOFFPAGE - sngMinX
          Else
            ' Does fit onto a single page horizontally, so centre.
            msngXOffset = mlngMarginLeft_Twips - sngMinX + _
              (((Printer.Width - (2 * MARGINCORRECTION) - (2 * iGAPOFFPAGE) - mlngMarginRight_Twips) - (sngMaxX - sngMinX)) / 2)
          End If
          msngYOffset = msngYOffset - sngMinY
          
          sngFirstPageXOffset = msngXOffset
          
          ' Print each element.
          pgNextPage.x = 0
          pgNextPage.y = 0
  
          Do While (pgNextPage.x >= 0) And (pgNextPage.y >= 0)
            sngPageRightMinX = -1
            sngPageBelowMinY = -1
    
            pgCurrentPage.x = pgNextPage.x
            pgCurrentPage.y = pgNextPage.y
          
            fNeedPageRight = False
            fNeedPageBelow = False
          
            fPagePrintedOn = False
    
            For iLoop = 0 To iCount
              fOnPageRight = False
              fOnPageBelow = False
                                  
              ' Should the element be printed on the next page to the right?
              If (awfElements(iLoop).Left + awfElements(iLoop).Width + msngXOffset) > _
                (Printer.Width - MARGINCORRECTION - iGAPOFFPAGE - mlngMarginRight_Twips) _
                And (msngYOffset + awfElements(iLoop).Top >= msngTopGap) Then
    
                sngPageRightMinX = IIf(sngPageRightMinX < 0, awfElements(iLoop).Left, IIf(sngPageRightMinX <= awfElements(iLoop).Left, sngPageRightMinX, awfElements(iLoop).Left))
                fOnPageRight = True
                fNeedPageRight = True
              End If
            
              ' Should the element be printed on the next page below?
              If (awfElements(iLoop).Top + awfElements(iLoop).Height + msngYOffset) > _
                (mlngRealBottom - iGAPOFFPAGE) Then
                
                sngPageBelowMinY = IIf(sngPageBelowMinY < 0, awfElements(iLoop).Top, IIf(sngPageBelowMinY <= awfElements(iLoop).Top, sngPageBelowMinY, awfElements(iLoop).Top))
    
                fOnPageBelow = True
                fNeedPageBelow = True
              End If
  
              ' Should the element be printed on this page?
              fOnCurrentPage = (Not fOnPageRight) _
                And (Not fOnPageBelow) _
                And (msngXOffset + awfElements(iLoop).Left >= (MARGINCORRECTION + mlngMarginLeft_Twips + iGAPOFFPAGE)) _
                And (msngYOffset + awfElements(iLoop).Top >= msngTopGap)
                
              If fOnCurrentPage Then
                fPagePrintedOn = True
                PrintWorkflowElementOverview awfElements(iLoop)
              End If
            Next iLoop
  
            pgNextPage.x = -1
            pgNextPage.y = -1
            
            If fNeedPageRight Then
              pgNextPage.x = pgCurrentPage.x + 1
              pgNextPage.y = pgCurrentPage.y
              
              msngXOffset = mlngMarginLeft_Twips + MARGINCORRECTION + iGAPOFFPAGE - sngPageRightMinX
            ElseIf fNeedPageBelow Then
              pgNextPage.x = 0
              pgNextPage.y = pgCurrentPage.y + 1
  
              msngXOffset = sngFirstPageXOffset
              msngYOffset = msngTopGap - sngPageBelowMinY
            End If
          
            If (fNeedPageRight Or fNeedPageBelow) And fPagePrintedOn Then
              mlngBottom = CalculateBottomOfPage
              Printer.CurrentY = mlngBottom + 1
              CheckEndOfPage2 mlngBottom, False
          
              ' Print the header
              .PrintHeader "Workflow Definition : " & Trim(msWorkflowName) & " (continued)"
            End If
          Loop
  
          If fPrintDetails Then
            mlngBottom = CalculateBottomOfPage
            Printer.CurrentY = mlngBottom + 1
            CheckEndOfPage2 mlngBottom, False
          End If
        
          .PageNumber = glngPageNum
        End If
      
        If fPrintDetails Then
          .PrintHeader "Workflow Definition : " & Trim(msWorkflowName)
          .PrintNormal "Name : " & Trim(msWorkflowName)
          .PrintNormal "Description : " & Trim(msWorkflowDescription)
          
          If mlngWorkflowPictureID > 0 Then
              recPictEdit.Index = "idxID"
              recPictEdit.Seek "=", mlngWorkflowPictureID
              If Not recPictEdit.NoMatch Then
                .PrintNormal "Picture : " & recPictEdit!Name
              End If
          End If
                    
          .PrintNormal "Initiation Type : " & WorkflowInitiationTypeDescription(miInitiationType)
          If miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
            .PrintNormal "Base Table : " & GetTableName(mlngBaseTableID)
          End If
          If miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL Then
            sURL = GetWorkflowURL
            .PrintNormal "External Initiation URL : " & IIf((Len(sURL) > 0) And (Len(msExternalInitiationQueryString) > 0), sURL & "?" & msExternalInitiationQueryString, "<undefined>")
          End If
          .PrintNormal "Enabled : " & IIf(mfWorkflowEnabled, "Yes", "No")
          .PrintNormal
  
          ' Print the elements.
          Printer.FontBold = False
          
          ReDim awfElements(0)
          For Each wfElement In mcolwfElements
            If wfElement.ControlIndex > 0 Then
              If wfElement.ElementType = elem_Begin Then
                ReadDefinitionIntoArray awfElements, wfElement
                Exit For
              End If
            End If
          Next wfElement
          Set wfElement = Nothing
          
          ' Remember to add any 'disconnected' elements to the array.
          For Each wfElement In mcolwfElements
            If wfElement.ControlIndex > 0 Then
              If (wfElement.ElementType <> elem_Begin) Then
                fConnected = False
                For Each wfLink In ASRWFLink1
                  If wfLink.EndElementIndex = wfElement.ControlIndex Then
                    fConnected = True
                    Exit For
                  End If
                Next wfLink
                Set wfLink = Nothing
              
                If Not fConnected Then
                  ReadDefinitionIntoArray awfElements, wfElement
                End If
              End If
            End If
          Next wfElement
          Set wfElement = Nothing
          
          For iLoop = 1 To UBound(awfElements)
            PrintWorkflowElementDetails iLoop, awfElements
          Next iLoop
        End If
      
        .PrintEnd
      End If
    End If
  End With
  Set mobjPrinter = Nothing

  If fOK Then
    Set objDefPrinter = New cSetDfltPrinter
    Do
      objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
    Loop While Printer.DeviceName <> gstrDefaultPrinterName
    Set objDefPrinter = Nothing
  End If
  
TidyUpAndExit:
  Exit Sub

ErrorTrap:
  MsgBox "Unable to print the workflow." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Sub

Public Property Get ReadOnly() As Boolean
  ReadOnly = mfReadOnly
  
End Property

Private Sub RefreshMenu()
  Dim iSelectedElementCount As Integer
  Dim iSelectedLinkCount As Integer
  Dim iOrCount As Integer
  Dim iSummingJunctionCount As Integer
  
  ' Get the count of select elements and links
  iSelectedElementCount = SelectedElementCount
  iSelectedLinkCount = SelectedLinkCount
  
  iOrCount = SelectedElementTypeCount(elem_Or)
  iSummingJunctionCount = SelectedElementTypeCount(elem_SummingJunction)
  
  With abMenu
    .Tools("ID_WFElement_Connector").Enabled = (Not mfReadOnly)
    .Tools("ID_WFElement_Decision").Enabled = (Not mfReadOnly)
    .Tools("ID_WFElement_Email").Enabled = (Not mfReadOnly)
    .Tools("ID_WFElement_Or").Enabled = (Not mfReadOnly)
    .Tools("ID_WFElement_StoredData").Enabled = (Not mfReadOnly)
    .Tools("ID_WFElement_SummingJunction").Enabled = (Not mfReadOnly)
    .Tools("ID_WFElement_Terminator").Enabled = (Not mfReadOnly)
    .Tools("ID_WFElement_WebForm").Enabled = (Not mfReadOnly)
    .Tools("ID_WFElement_Link").Enabled = (Not mfReadOnly)
    
    .Tools("ID_WorkflowCut").Enabled = (Not mfReadOnly) _
      And ((iSelectedElementCount > 0) Or (iSelectedLinkCount > 0))

    .Tools("ID_WorkflowCopy").Enabled = (Not mfReadOnly) _
      And (iSelectedElementCount > 0)

    .Tools("ID_WorkflowPaste").Enabled = (Not mfReadOnly) _
      And (ClipboardControlsCount > 0)

    .Tools("ID_WorkflowDelete").Enabled = (Not mfReadOnly) _
      And ((iSelectedElementCount > 0) Or (iSelectedLinkCount > 0))
    
    .Tools("ID_WorkflowClear").Enabled = (Not mfReadOnly) _
      And (mcolwfElements.Count > 0)
    
    .Tools("ID_WorkflowUndo").Enabled = (Not mfReadOnly) _
      And (miLastActionFlag <> giACTION_NOACTION)
  
    .Tools("ID_WorkflowElementProperties").Enabled = (iSelectedElementCount = 1) _
      And (iOrCount + iSummingJunctionCount = 0)
      
    .Tools("ID_WorkflowAutoLayout").Enabled = (Application.AccessMode = accFull Or _
      Application.AccessMode = accSupportMode)
      
    .Tools("ID_WorkflowResizeCanvas").Enabled = (Application.AccessMode = accFull Or _
      Application.AccessMode = accSupportMode)
    
  End With
  
End Sub

Private Sub SetElementAddPointer()
  Dim sMenuOption As String
  Dim objTool As ActiveBarLibraryCtl.Tool
  Dim fFound As Boolean
  
  fFound = False
  
  For Each objTool In abMenu.Tools
    If objTool.Checked Then
      fFound = True

      Me.MouseIcon = LoadResPicture(objTool.Name, vbResCursor)
      Me.MousePointer = vbCustom
      Exit For
    End If
  Next objTool
  Set objTool = Nothing

  If Not fFound Then
    Me.MousePointer = vbNormal
  End If
  
End Sub


Private Sub ShellSortElements(vArray As Variant)
  Dim lLoop1 As Long
  Dim lHold As Long
  Dim lHValue As Long
  Dim lTemp As Long
  Dim wfElement1 As VB.Control
  Dim wfElement2 As VB.Control
  Dim wfTmpElement As VB.Control

  lHValue = LBound(vArray)
  
  Do
    lHValue = 3 * lHValue + 1
  Loop Until lHValue > UBound(vArray)
  
  Do
    lHValue = lHValue / 3
    
    For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
      Set wfTmpElement = vArray(lLoop1)
      lHold = lLoop1
      
      Do While (vArray(lHold - lHValue).Top > wfTmpElement.Top) _
        Or ((vArray(lHold - lHValue).Top = wfTmpElement.Top) And (vArray(lHold - lHValue).Left > wfTmpElement.Left))

        Set vArray(lHold) = vArray(lHold - lHValue)
        lHold = lHold - lHValue
        If lHold < lHValue Then Exit Do
      Loop
      
      Set vArray(lHold) = wfTmpElement
    Next lLoop1
  Loop Until lHValue = LBound(vArray)
  
End Sub

Private Sub PrintWorkflowElementOverview(pwfElement As VB.Control)
  ' Print the overview (graphic) of the given workflow element.
  Dim rctDraw As Rect
  Dim iSM As Integer
  Dim wfLink As COAWF_Link
  Dim wfElement As VB.Control
  Dim asngLineCoordinates() As Single
  Dim iLoop As Integer
  Dim sngXOffset As Single
  Dim sngYOffset As Single
  Dim iDrawWidth As Integer
  Dim sngCaptionXOffset As Single
  Dim sngMinX As Single
  Dim sngMinY As Single
  Dim sngMaxX As Single
  Dim sngMaxY As Single
  Dim fOnCurrentPage As Boolean
  Dim sngX1 As Single
  Dim sngY1 As Single
  Dim sngX2 As Single
  Dim sngY2 As Single
  Dim fChoppedStart As Boolean
  Dim fChoppedEnd As Boolean
  Dim fPrintArrow As Boolean
  Dim linDirection As LineDirection
  Dim fLastArrowDone As Boolean
  
  iSM = Printer.ScaleMode
  sngCaptionXOffset = IIf(pwfElement.ElementType = elem_Connector1 Or pwfElement.ElementType = elem_Connector2, 0, 25)
  
  ' Print the element image
  Printer.PaintPicture pwfElement.ElementPicture, _
    msngXOffset + pwfElement.Left, _
    msngYOffset + pwfElement.Top

  If Len(pwfElement.Caption) > 0 Then
    ' Print the caption.
    rctDraw.Left = msngXOffset + pwfElement.Left + pwfElement.CaptionHorizontalPosition - sngCaptionXOffset
    rctDraw.Right = rctDraw.Left + pwfElement.CaptionWidth + (2 * sngCaptionXOffset)
    rctDraw.Top = msngYOffset + pwfElement.Top + pwfElement.CaptionVerticalPosition
    rctDraw.Bottom = rctDraw.Top + pwfElement.CaptionHeight
    
    ' Scale the print coordinates for use with the DrawText function.
    rctDraw.Left = Printer.ScaleY(rctDraw.Left, iSM, vbPixels)
    rctDraw.Right = Printer.ScaleY(rctDraw.Right, iSM, vbPixels)
    rctDraw.Top = Printer.ScaleY(rctDraw.Top, iSM, vbPixels)
    rctDraw.Bottom = Printer.ScaleY(rctDraw.Bottom, iSM, vbPixels)
    
    Printer.Font.Name = pwfElement.Font.Name
    Printer.Font.Size = pwfElement.Font.Size
    Printer.Font.Bold = pwfElement.Font.Bold
    Printer.Font.Underline = pwfElement.Font.Underline
    Printer.Font.Strikethrough = pwfElement.Font.Strikethrough
    Printer.Font.Italic = pwfElement.Font.Italic

    DrawText Printer.hDC, pwfElement.Caption, -1, rctDraw, c_DTDefFmt Or DT_LEFT
  End If
  
  ' Print the mini captions (the little captions for optional outbound flows)
  Select Case pwfElement.ElementType
    Case elem_WebForm
      ' Print the 'timeout' label
      If pwfElement.WebFormTimeoutFrequency > 0 Then
        rctDraw.Left = msngXOffset + pwfElement.Left + pwfElement.MiniCaptionHorizontalPosition(0)
        rctDraw.Right = rctDraw.Left + pwfElement.MiniCaptionWidth(0)
        rctDraw.Top = msngYOffset + pwfElement.Top + pwfElement.MiniCaptionVerticalPosition(0)
        rctDraw.Bottom = rctDraw.Top + pwfElement.MiniCaptionHeight(0)
    
        ' Scale the print coordinates for use with the DrawText function.
        rctDraw.Left = Printer.ScaleY(rctDraw.Left, iSM, vbPixels)
        rctDraw.Right = Printer.ScaleY(rctDraw.Right, iSM, vbPixels)
        rctDraw.Top = Printer.ScaleY(rctDraw.Top, iSM, vbPixels)
        rctDraw.Bottom = Printer.ScaleY(rctDraw.Bottom, iSM, vbPixels)
    
        Printer.Font.Name = pwfElement.MiniCaptionFont.Name
        Printer.Font.Size = pwfElement.MiniCaptionFont.Size
        Printer.Font.Bold = pwfElement.MiniCaptionFont.Bold
        Printer.Font.Underline = pwfElement.MiniCaptionFont.Underline
        Printer.Font.Strikethrough = pwfElement.MiniCaptionFont.Strikethrough
        Printer.Font.Italic = pwfElement.MiniCaptionFont.Italic
    
        DrawText Printer.hDC, pwfElement.MiniCaption(0), -1, rctDraw, c_DTDefFmt Or DT_LEFT
      End If
  
    Case elem_Decision
      ' Print the 'false' label
      rctDraw.Left = msngXOffset + pwfElement.Left + pwfElement.MiniCaptionHorizontalPosition(1)
      rctDraw.Right = rctDraw.Left + pwfElement.MiniCaptionWidth(1)
      rctDraw.Top = msngYOffset + pwfElement.Top + pwfElement.MiniCaptionVerticalPosition(1)
      rctDraw.Bottom = rctDraw.Top + pwfElement.MiniCaptionHeight(1)
      
      ' Scale the print coordinates for use with the DrawText function.
      rctDraw.Left = Printer.ScaleY(rctDraw.Left, iSM, vbPixels)
      rctDraw.Right = Printer.ScaleY(rctDraw.Right, iSM, vbPixels)
      rctDraw.Top = Printer.ScaleY(rctDraw.Top, iSM, vbPixels)
      rctDraw.Bottom = Printer.ScaleY(rctDraw.Bottom, iSM, vbPixels)
      
      Printer.Font.Name = pwfElement.MiniCaptionFont.Name
      Printer.Font.Size = pwfElement.MiniCaptionFont.Size
      Printer.Font.Bold = pwfElement.MiniCaptionFont.Bold
      Printer.Font.Underline = pwfElement.MiniCaptionFont.Underline
      Printer.Font.Strikethrough = pwfElement.MiniCaptionFont.Strikethrough
      Printer.Font.Italic = pwfElement.MiniCaptionFont.Italic
  
      DrawText Printer.hDC, pwfElement.MiniCaption(1), -1, rctDraw, c_DTDefFmt Or DT_LEFT
    
      ' Print the 'true' label
      rctDraw.Left = msngXOffset + pwfElement.Left + pwfElement.MiniCaptionHorizontalPosition(0)
      rctDraw.Right = rctDraw.Left + pwfElement.MiniCaptionWidth(0)
      rctDraw.Top = msngYOffset + pwfElement.Top + pwfElement.MiniCaptionVerticalPosition(0)
      rctDraw.Bottom = rctDraw.Top + pwfElement.MiniCaptionHeight(0)
      
      ' Scale the print coordinates for use with the DrawText function.
      rctDraw.Left = Printer.ScaleY(rctDraw.Left, iSM, vbPixels)
      rctDraw.Right = Printer.ScaleY(rctDraw.Right, iSM, vbPixels)
      rctDraw.Top = Printer.ScaleY(rctDraw.Top, iSM, vbPixels)
      rctDraw.Bottom = Printer.ScaleY(rctDraw.Bottom, iSM, vbPixels)
      
      Printer.Font.Name = pwfElement.DecisionCaptionFont.Name
      Printer.Font.Size = pwfElement.DecisionCaptionFont.Size
      Printer.Font.Bold = pwfElement.DecisionCaptionFont.Bold
      Printer.Font.Underline = pwfElement.DecisionCaptionFont.Underline
      Printer.Font.Strikethrough = pwfElement.DecisionCaptionFont.Strikethrough
      Printer.Font.Italic = pwfElement.DecisionCaptionFont.Italic
  
      DrawText Printer.hDC, pwfElement.MiniCaption(0), -1, rctDraw, c_DTDefFmt Or DT_LEFT
      
    Case elem_StoredData
      ' Print the 'failure' label
      rctDraw.Left = msngXOffset + pwfElement.Left + pwfElement.MiniCaptionHorizontalPosition(1)
      rctDraw.Right = rctDraw.Left + pwfElement.MiniCaptionWidth(1)
      rctDraw.Top = msngYOffset + pwfElement.Top + pwfElement.MiniCaptionVerticalPosition(1)
      rctDraw.Bottom = rctDraw.Top + pwfElement.MiniCaptionHeight(1)

      ' Scale the print coordinates for use with the DrawText function.
      rctDraw.Left = Printer.ScaleY(rctDraw.Left, iSM, vbPixels)
      rctDraw.Right = Printer.ScaleY(rctDraw.Right, iSM, vbPixels)
      rctDraw.Top = Printer.ScaleY(rctDraw.Top, iSM, vbPixels)
      rctDraw.Bottom = Printer.ScaleY(rctDraw.Bottom, iSM, vbPixels)

      Printer.Font.Name = pwfElement.MiniCaptionFont.Name
      Printer.Font.Size = pwfElement.MiniCaptionFont.Size
      Printer.Font.Bold = pwfElement.MiniCaptionFont.Bold
      Printer.Font.Underline = pwfElement.MiniCaptionFont.Underline
      Printer.Font.Strikethrough = pwfElement.MiniCaptionFont.Strikethrough
      Printer.Font.Italic = pwfElement.MiniCaptionFont.Italic

      DrawText Printer.hDC, pwfElement.MiniCaption(1), -1, rctDraw, c_DTDefFmt Or DT_LEFT

      ' Print the 'success' label
      rctDraw.Left = msngXOffset + pwfElement.Left + pwfElement.MiniCaptionHorizontalPosition(0)
      rctDraw.Right = rctDraw.Left + pwfElement.MiniCaptionWidth(0)
      rctDraw.Top = msngYOffset + pwfElement.Top + pwfElement.MiniCaptionVerticalPosition(0)
      rctDraw.Bottom = rctDraw.Top + pwfElement.MiniCaptionHeight(0)

      ' Scale the print coordinates for use with the DrawText function.
      rctDraw.Left = Printer.ScaleY(rctDraw.Left, iSM, vbPixels)
      rctDraw.Right = Printer.ScaleY(rctDraw.Right, iSM, vbPixels)
      rctDraw.Top = Printer.ScaleY(rctDraw.Top, iSM, vbPixels)
      rctDraw.Bottom = Printer.ScaleY(rctDraw.Bottom, iSM, vbPixels)

      Printer.Font.Name = pwfElement.Font.Name
      Printer.Font.Size = pwfElement.Font.Size
      Printer.Font.Bold = pwfElement.Font.Bold
      Printer.Font.Underline = pwfElement.Font.Underline
      Printer.Font.Strikethrough = pwfElement.Font.Strikethrough
      Printer.Font.Italic = pwfElement.Font.Italic

      DrawText Printer.hDC, pwfElement.MiniCaption(0), -1, rctDraw, c_DTDefFmt Or DT_LEFT
  End Select
  
  ' Print the links for the given element.
  For Each wfLink In ASRWFLink1
    If (wfLink.StartElementIndex = pwfElement.ControlIndex Or _
      wfLink.EndElementIndex = pwfElement.ControlIndex) Then
      
      ' Print the link lines.
      asngLineCoordinates = wfLink.LineCoordinates
      
      sngXOffset = msngXOffset + wfLink.Left
      sngYOffset = msngYOffset + wfLink.Top
      
      fLastArrowDone = False
      
      For iLoop = 0 To UBound(asngLineCoordinates, 2)
        ' Is the link line on this page?
        sngMinX = sngXOffset + asngLineCoordinates(0, iLoop)
        sngMinX = IIf(sngMinX <= sngXOffset + asngLineCoordinates(1, iLoop), sngMinX, sngXOffset + asngLineCoordinates(1, iLoop))
        
        sngMinY = sngYOffset + asngLineCoordinates(2, iLoop)
        sngMinY = IIf(sngMinY <= sngYOffset + asngLineCoordinates(3, iLoop), sngMinY, sngYOffset + asngLineCoordinates(3, iLoop))
        
        sngMaxX = sngXOffset + asngLineCoordinates(0, iLoop)
        sngMaxX = IIf(sngMaxX >= sngXOffset + asngLineCoordinates(1, iLoop), sngMaxX, sngXOffset + asngLineCoordinates(1, iLoop))
        
        sngMaxY = sngYOffset + asngLineCoordinates(2, iLoop)
        sngMaxY = IIf(sngMaxY >= sngYOffset + asngLineCoordinates(3, iLoop), sngMaxY, sngYOffset + asngLineCoordinates(3, iLoop))
        
        fOnCurrentPage = (sngMaxX > (MARGINCORRECTION + mlngMarginLeft_Twips)) _
            And (sngMinX < (Printer.Width - MARGINCORRECTION - mlngMarginLeft_Twips - mlngMarginRight_Twips)) _
            And (sngMaxY > (msngTopGap - iGAPOFFPAGE)) _
            And (sngMinY < (mlngRealBottom - iGAPOFFPAGE))
        
        If fOnCurrentPage Then
          ' Line should be printed on this page.
          sngX1 = sngXOffset + asngLineCoordinates(0, iLoop)
          sngX2 = sngXOffset + asngLineCoordinates(1, iLoop)
          sngY1 = sngYOffset + asngLineCoordinates(2, iLoop)
          sngY2 = sngYOffset + asngLineCoordinates(3, iLoop)
            
          fChoppedStart = False
          fChoppedEnd = False
          fPrintArrow = False

          If sngX1 < (MARGINCORRECTION + mlngMarginLeft_Twips) Then
            sngX1 = (MARGINCORRECTION + mlngMarginLeft_Twips)
            fChoppedStart = True
            linDirection = lineDirection_Left
            fPrintArrow = False
          End If
          If sngX1 > (Printer.Width - MARGINCORRECTION - mlngMarginLeft_Twips - mlngMarginRight_Twips) Then
            sngX1 = (Printer.Width - MARGINCORRECTION - mlngMarginLeft_Twips - mlngMarginRight_Twips)
            fChoppedStart = True
            linDirection = lineDirection_Right
            fPrintArrow = False
          End If
          If sngX2 < (MARGINCORRECTION + mlngMarginLeft_Twips) Then
            sngX2 = (MARGINCORRECTION + mlngMarginLeft_Twips)
            fChoppedEnd = True
            linDirection = lineDirection_Right
            fPrintArrow = True
          End If
          If sngX2 > (Printer.Width - MARGINCORRECTION - mlngMarginLeft_Twips - mlngMarginRight_Twips) Then
            sngX2 = (Printer.Width - MARGINCORRECTION - mlngMarginLeft_Twips - mlngMarginRight_Twips)
            fChoppedEnd = True
            linDirection = lineDirection_Left
            fPrintArrow = True
          End If

          If sngY1 < (msngTopGap - iGAPOFFPAGE) Then
            sngY1 = (msngTopGap - iGAPOFFPAGE)
            fChoppedStart = True
            linDirection = lineDirection_Up
            fPrintArrow = False
          End If
          If sngY1 > (mlngRealBottom - iGAPOFFPAGE) Then
            sngY1 = (mlngRealBottom - iGAPOFFPAGE)
            fChoppedStart = True
            linDirection = lineDirection_Down
            fPrintArrow = False
          End If
          If sngY2 < (msngTopGap - iGAPOFFPAGE) Then
            sngY2 = (msngTopGap - iGAPOFFPAGE)
            fChoppedEnd = True
            linDirection = lineDirection_Down
            fPrintArrow = True
          End If
          If sngY2 > (mlngRealBottom - iGAPOFFPAGE) Then
            sngY2 = (mlngRealBottom - iGAPOFFPAGE)
            fChoppedEnd = True
            linDirection = lineDirection_Up
            fPrintArrow = True
          End If

          Printer.Line (sngX1, sngY1)-(sngX2, sngY2)

          If Not fChoppedEnd And (iLoop = UBound(asngLineCoordinates, 2)) Then
            ' Check if the link ends on the page, but the associated control will not fit.
            Set wfElement = mcolwfElements(CStr(wfLink.EndElementIndex))
            fOnCurrentPage = True
  
            ' Should the element be printed on another page?
            If ((wfElement.Left + wfElement.Width + msngXOffset) > (Printer.Width - MARGINCORRECTION - iGAPOFFPAGE - mlngMarginRight_Twips) _
                And (msngYOffset + wfElement.Top >= msngTopGap)) _
              Or (msngXOffset + wfElement.Left < (MARGINCORRECTION + mlngMarginLeft_Twips + iGAPOFFPAGE)) _
              Or ((wfElement.Top + wfElement.Height + msngYOffset) > (mlngRealBottom - iGAPOFFPAGE)) _
              Or (msngYOffset + wfElement.Top < msngTopGap) Then
              
              fOnCurrentPage = False
            End If

            ' Should the element be printed on this page?
            If Not fOnCurrentPage Then
              fChoppedEnd = True
              linDirection = wfLink.EndDirection
            End If
          End If
          
          If Not fChoppedStart And (iLoop = 0) Then
            ' Check if the link starts on the page, but the associated control will not fit.
            Set wfElement = mcolwfElements(CStr(wfLink.StartElementIndex))
            fOnCurrentPage = True
            
            ' Should the element be printed on another page?
'''            If ((wfElement.Left + wfElement.Width + msngXOffset) > (Printer.Width - MARGINCORRECTION - iGAPOFFPAGE - mlngMarginLeft_Twips - mlngMarginRight_Twips)
            If ((wfElement.Left + wfElement.Width + msngXOffset) > (Printer.Width - MARGINCORRECTION - iGAPOFFPAGE - mlngMarginRight_Twips) _
                And (msngYOffset + wfElement.Top >= msngTopGap)) _
              Or (msngXOffset + wfElement.Left < (MARGINCORRECTION + mlngMarginLeft_Twips + iGAPOFFPAGE)) _
              Or ((wfElement.Top + wfElement.Height + msngYOffset) > (mlngRealBottom - iGAPOFFPAGE)) _
              Or (msngYOffset + wfElement.Top < msngTopGap) Then
              
              fOnCurrentPage = False
            End If

            If Not fOnCurrentPage Then
              fChoppedStart = True
              linDirection = wfLink.StartDirection
            End If
          End If
          
          If fChoppedStart Then
            ' Print the offPage connector image
            PrintOffPageConnector wfLink.Index, sngX1, sngY1, False, linDirection
          End If
            
          If fChoppedEnd Then
            ' Print the offPage connector image
            PrintOffPageConnector wfLink.Index, sngX2, sngY2, True, linDirection
          End If
        Else
          ' Line not printed on this page
          If iLoop = UBound(asngLineCoordinates, 2) Then
            fLastArrowDone = True
          End If
        End If
      Next iLoop
    
'      If Not fLastArrowDone Then
'        ' Print the link arrow.
'        Printer.PaintPicture wfLink.ArrowPicture, _
'          sngXOffset + wfLink.ArrowHorizontalPosition, _
'          sngYOffset + wfLink.ArrowVerticalPosition
'      End If
    End If
  Next wfLink
  Set wfLink = Nothing
  
End Sub


Private Sub PrintWorkflowElementDetails(piIndex As Integer, pawfElements As Variant)
  ' Print the definition of the given workflow element.
  Dim sTitle As String
  Dim sTemp As String
  Dim asTemp() As String
  Dim wfElement As VB.Control
  Dim wfLink As COAWF_Link
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim avOutboundFlowInfo() As Variant
  Dim fFound As Boolean
  
  Set wfElement = pawfElements(piIndex)
  
  With mobjPrinter
    sTitle = CStr(piIndex) & " : " & wfElement.ElementTypeDescription
    
    If wfElement.ElementType <> elem_Or And _
      wfElement.ElementType <> elem_SummingJunction Then
      
      sTitle = sTitle & " - '" & wfElement.Caption & "'"
    End If
    
    .PrintTitle sTitle
    
    Printer.FontBold = False
    
    .PrintNormal "Type : " & wfElement.ElementTypeDescription
    
    Select Case wfElement.ElementType
      Case elem_Begin
        PrintElementDetails_Begin wfElement
      Case elem_Terminator
        PrintElementDetails_Terminator wfElement
      Case elem_WebForm
        PrintElementDetails_WebForm wfElement
      Case elem_Email
        PrintElementDetails_Email wfElement
      Case elem_Decision
        PrintElementDetails_Decision wfElement
      Case elem_StoredData
        PrintElementDetails_StoredData wfElement
      Case elem_SummingJunction
        PrintElementDetails_SummingJunction wfElement
      Case elem_Or
        PrintElementDetails_Or wfElement
      Case elem_Connector1
        PrintElementDetails_Connector1 wfElement
      Case elem_Connector2
        PrintElementDetails_Connector2 wfElement, pawfElements
    End Select
    
    ' Print preceding elements
    sTemp = ""
    For iLoop = 1 To UBound(pawfElements)
      For Each wfLink In ASRWFLink1
        If wfLink.StartElementIndex = pawfElements(iLoop).ControlIndex _
          And wfLink.EndElementIndex = wfElement.ControlIndex Then

          sTemp = sTemp & IIf(Len(sTemp) > 0, ", ", "") & CStr(iLoop)
        End If
      Next wfLink
      Set wfLink = Nothing
    Next iLoop
    If Len(sTemp) > 0 Then
      .PrintNormal "Preceding Element" & IIf(InStr(sTemp, ",") > 0, "s", "") & " : " & sTemp
    End If
        
    ' Print succeeding connector element details
    If (wfElement.ElementType = elem_Connector1) Then
      For iLoop = 1 To UBound(pawfElements)
        If pawfElements(iLoop).ControlIndex = wfElement.ConnectorPairIndex Then
          .PrintNormal "Succeeding Connector : " & CStr(iLoop)
          Exit For
        End If
      Next iLoop
    End If
    
    ' Print succeeding elements
    
    ' Get the array of outbound flow information from the start element.
    ' Column 1 = Tag (see enums, or -1 if there's only a single outbound flow)
    ' Column 2 = Direction
    ' Column 3 = XOffset
    ' Column 4 = YOffset
    ' Column 5 = Maximum
    ' Column 6 = Minimum
    ' Column 7 = Description
    avOutboundFlowInfo = wfElement.OutboundFlows_Information
    
    ReDim asTemp(UBound(avOutboundFlowInfo, 2))
    ' Succeeding element indexes
    
    For iLoop = 1 To UBound(pawfElements)
      For Each wfLink In ASRWFLink1
        If wfLink.EndElementIndex = pawfElements(iLoop).ControlIndex _
          And wfLink.StartElementIndex = wfElement.ControlIndex Then
          
          fFound = False
          For iLoop2 = 1 To UBound(asTemp)
            If avOutboundFlowInfo(1, iLoop2) = wfLink.StartOutboundFlowCode Then
              fFound = True
              asTemp(iLoop2) = asTemp(iLoop2) & IIf(Len(asTemp(iLoop2)) > 0, ", ", "") & CStr(iLoop)
              Exit For
            End If
          Next iLoop2
          
          If Not fFound Then
            asTemp(1) = asTemp(1) & IIf(Len(asTemp(1)) > 0, ", ", "") & CStr(iLoop)
          End If
        End If
      Next wfLink
      Set wfLink = Nothing
    Next iLoop
    
    For iLoop = 1 To UBound(asTemp)
      If UBound(asTemp) = 1 Then
        sTemp = "Succeeding Element" & IIf(InStr(sTemp, ",") > 0, "s", "") & " : " & asTemp(iLoop)
      Else
        sTemp = "Succeeding '" & avOutboundFlowInfo(7, iLoop) & "' Element" & IIf(InStr(sTemp, ",") > 0, "s", "") & " : " & asTemp(iLoop)
      End If
      
      .PrintNormal sTemp
    Next iLoop
  End With
  
End Sub



Private Sub PrintElementDetails_Decision(pwfElement As VB.Control)
  ' Print the definition of the given workflow element.
  Dim sTemp As String

  With mobjPrinter
    ' Caption
    .PrintNormal "Caption : " & pwfElement.Caption

    ' True flow criteria
    If (pwfElement.DecisionFlowType = decisionFlowType_Expression) Then
      ' Expression
      If pwfElement.DecisionFlowExpressionID = 0 Then
        sTemp = "No calculation"
      Else
        sTemp = "'" & GetExpressionName(pwfElement.DecisionFlowExpressionID) & "' calculation"
      End If
    Else
      ' Button
      If Len(Trim(pwfElement.TrueFlowIdentifier)) = 0 Then
        sTemp = "Unidentified button"
      Else
        sTemp = "'" & pwfElement.TrueFlowIdentifier & "' button"
      End If
    End If
    
    .PrintNormal GetDecisionCaptionDescription(pwfElement.DecisionCaptionType, True) & " Flow Criteria : " & sTemp
  End With
  
End Sub




Private Sub PrintElementDetails_SummingJunction(pwfElement As VB.Control)
  ' Print the definition of the given workflow element.
  
  ' Not required for SummingJunction elements
  
End Sub





Private Sub PrintElementDetails_Begin(pwfElement As VB.Control)
  ' Print the definition of the given workflow element.
  With mobjPrinter
    .PrintNormal "Caption : " & pwfElement.Caption
  End With
  
End Sub

Private Sub PrintElementDetails_Terminator(pwfElement As VB.Control)
  ' Print the definition of the given workflow element.
  With mobjPrinter
    .PrintNormal "Caption : " & pwfElement.Caption
  End With
  
End Sub





Private Sub PrintElementDetails_StoredData(pwfElement As VB.Control)
  ' Print the definition of the given workflow element.
  Dim wfElement2 As VB.Control
  Dim iLoop As Integer
  Dim avColumns() As Variant
  Dim objMisc As Misc
  
  Set objMisc = New Misc
  
  With mobjPrinter
    ' Caption
    .PrintNormal "Caption : " & pwfElement.Caption

    ' Identifier
    .PrintNormal "Identifier : " & pwfElement.Identifier

    ' Use as workflow traget identifier
    .PrintNormal "Use As Target Identifier : " & pwfElement.UseAsTargetIdentifier

    ' Data Action
    .PrintNormal "Data Action : " & IIf(pwfElement.DataAction = DATAACTION_DELETE, "Delete", IIf(pwfElement.DataAction = DATAACTION_INSERT, "Insert", "Update"))

    ' Data Table
    .PrintNormal "Data Table : " & GetTableName(pwfElement.DataTableID)

    ' Primary Data Record
    .PrintNormal "Data Record : " & GetRecordSelectionDescription(pwfElement.DataRecord)

    If pwfElement.DataRecord = giWFRECSEL_IDENTIFIEDRECORD Then
      ' Primary Data Record - Element (if required)
      .PrintNormal "Element : " & pwfElement.RecordSelectorWebFormIdentifier

      Set wfElement2 = GetElementByIdentifier(pwfElement.RecordSelectorWebFormIdentifier)
      If Not wfElement2 Is Nothing Then
        If wfElement2.ElementType = elem_WebForm Then
          ' Primary Data Record - Record Selector (if required)
          .PrintNormal "Record Selector : " & pwfElement.RecordSelectorIdentifier
        End If
      End If
    End If

    ' Primary Data Record Table
    If pwfElement.DataRecordTableID > 0 Then
      .PrintNormal "Data Record Table : " & GetTableName(pwfElement.DataRecordTableID)
    End If

    If Len(pwfElement.SecondaryRecordSelectorWebFormIdentifier) > 0 Then
      ' Secondary Data Record (if required)
      .PrintNormal "Secondary Data Record : " & GetRecordSelectionDescription(pwfElement.SecondaryDataRecord)

      If pwfElement.SecondaryDataRecord = giWFRECSEL_IDENTIFIEDRECORD Then
        ' Secondary Data Record - Element (if required)
        .PrintNormal "Secondary Element : " & pwfElement.SecondaryRecordSelectorWebFormIdentifier

        Set wfElement2 = GetElementByIdentifier(pwfElement.SecondaryRecordSelectorWebFormIdentifier)
        If Not wfElement2 Is Nothing Then
          If wfElement2.ElementType = elem_WebForm Then
            ' Secondary Data Record - Record Selector (if required)
            .PrintNormal "Secondary Record Selector : " & pwfElement.SecondaryRecordSelectorIdentifier
          End If
        End If
      End If
    
      ' Secondary Data Record Table (if required)
      If pwfElement.SecondaryDataRecordTableID > 0 Then
        .PrintNormal "Secondary Data Record Table : " & GetTableName(pwfElement.SecondaryDataRecordTableID)
      End If
    End If

    ' Stored Data Columns
    avColumns = pwfElement.DataColumns
    
    If UBound(avColumns, 2) > 0 Then
      .PrintNormal "Columns :"

      For iLoop = 1 To UBound(avColumns, 2)
        ' Column Name
        .PrintNormal "     Column : " & GetColumnName(CLng(avColumns(3, iLoop)), True)

        Select Case CInt(avColumns(4, iLoop))
          Case giWFDATAVALUE_FIXED
            If (GetColumnDataType(CLng(avColumns(3, iLoop))) = dtTIMESTAMP) _
              And UCase(avColumns(5, iLoop)) <> "NULL" Then
              
              .PrintNormal "          Fixed value : " & objMisc.ConvertSQLDateToLocale(CStr(avColumns(5, iLoop)))
            Else
              .PrintNormal "          Fixed value : " & avColumns(5, iLoop)
            End If

          Case giWFDATAVALUE_WFVALUE
            .PrintNormal "          Workflow value : " & avColumns(6, iLoop) & "." & avColumns(7, iLoop)

          Case giWFDATAVALUE_DBVALUE
            .PrintNormal "          Database value : " & GetColumnName(CLng(avColumns(8, iLoop)))
            .PrintNormal "               Record : " & GetRecordSelectionDescription(CInt(avColumns(9, iLoop)))
            If CInt(avColumns(9, iLoop)) = giWFRECSEL_IDENTIFIEDRECORD Then
              .PrintNormal "               Element : " & avColumns(6, iLoop)

              Set wfElement2 = GetElementByIdentifier(CStr(avColumns(6, iLoop)))
              If Not wfElement2 Is Nothing Then
                If wfElement2.ElementType = elem_WebForm Then
                  .PrintNormal "               Record Selector : " & avColumns(7, iLoop)
                End If
              End If
            End If
            
          Case giWFDATAVALUE_CALC
            .PrintNormal "          Calculated value : " & GetColumnName(CLng(avColumns(3, iLoop)))
            If CLng(avColumns(10, iLoop)) > 0 Then
              .PrintNormal "               Calculation : " & GetExpressionName(CLng(avColumns(10, iLoop)))
            Else
              .PrintNormal "               Calculation : None"
            End If
        End Select
      Next iLoop
    End If
  End With
  
  Set objMisc = Nothing
  
End Sub





Private Sub PrintElementDetails_WebForm(pwfElement As VB.Control)
  ' Print the definition of the given workflow element.
  Dim sTemp As String
  Dim sTemp2 As String
  Dim sTemp3 As String
  Dim sTemp4 As String
  Dim sTemp5 As String
  Dim sTemp6 As String
  Dim sTemp_pt1 As String
  Dim sTemp_pt2 As String
  Dim sTemp_pt3 As String
  Dim wfElement2 As VB.Control
  Dim iLoop As Integer
  Dim iLoop3 As Integer
  Dim asItems() As String
  Dim asItemValues() As String
  Dim asValidations() As String
  Dim objOperatorDef As clsOperatorDef
  Dim sOperatorName As String
  
  With mobjPrinter
    ' Caption
    .PrintNormal "Caption : " & pwfElement.Caption

    ' Identifier
    .PrintNormal "Identifier : " & pwfElement.Identifier

    ' Description
    If pwfElement.DescriptionExprID = 0 Then
      sTemp = "No description"
    Else
      sTemp = "'" & GetExpressionName(pwfElement.DescriptionExprID) & "' calculation"
    End If
    .PrintNormal "Description : " & sTemp
    
    If pwfElement.DescriptionExprID > 0 Then
      .PrintNormal "     Description prefixed with Workflow Name : " & IIf(pwfElement.DescriptionHasWorkflowName, "True", "False")
      .PrintNormal "     Description prefixed with Element Caption : " & IIf(pwfElement.DescriptionHasElementCaption, "True", "False")
    End If
    
    ' Timeout
    If pwfElement.WebFormTimeoutFrequency > 0 Then
      .PrintNormal "Timeout : " & pwfElement.TimeoutPeriodDescription(pwfElement.WebFormTimeoutFrequency, pwfElement.WebFormTimeoutPeriod)
      .PrintNormal "Exclude Weekends : " & IIf(pwfElement.WebFormTimeoutExcludeWeekend, "True", "False")
    End If

    If (pwfElement.WFCompletionMessageType = MESSAGE_CUSTOM) Then
      ' Trim out "\ul " and "\ulnone "
      ParseWebFormMessage pwfElement.WFCompletionMessage, _
        sTemp_pt1, _
        sTemp_pt2, _
        sTemp_pt3
      sTemp = Replace(Replace(sTemp_pt1 & sTemp_pt2 & sTemp_pt3, vbCr, ""), vbLf, "")
    ElseIf (pwfElement.WFCompletionMessageType = MESSAGE_NONE) Then
      sTemp = "<none>"
    Else
      'MESSAGE_SYSTEMDEFAULT
      sTemp = "<system default>"
    End If
    .PrintNormal "Completion Message : " & sTemp

    If (pwfElement.WFSavedForLaterMessageType = MESSAGE_CUSTOM) Then
      ' Trim out "\ul " and "\ulnone "
      ParseWebFormMessage pwfElement.WFSavedForLaterMessage, _
        sTemp_pt1, _
        sTemp_pt2, _
        sTemp_pt3
      sTemp = Replace(Replace(sTemp_pt1 & sTemp_pt2 & sTemp_pt3, vbCr, ""), vbLf, "")
    ElseIf (pwfElement.WFSavedForLaterMessageType = MESSAGE_NONE) Then
      sTemp = "<none>"
    Else
      'MESSAGE_SYSTEMDEFAULT
      sTemp = "<system default>"
    End If
    .PrintNormal "Saved For Later Message : " & sTemp
    
    If (pwfElement.WFFollowOnFormsMessageType = MESSAGE_CUSTOM) Then
      ' Trim out "\ul " and "\ulnone "
      ParseWebFormMessage pwfElement.WFFollowOnFormsMessage, _
        sTemp_pt1, _
        sTemp_pt2, _
        sTemp_pt3
      sTemp = Replace(Replace(sTemp_pt1 & sTemp_pt2 & sTemp_pt3, vbCr, ""), vbLf, "")
    ElseIf (pwfElement.WFFollowOnFormsMessageType = MESSAGE_NONE) Then
      sTemp = "<none>"
    Else
      'MESSAGE_SYSTEMDEFAULT
      sTemp = "<system default>"
    End If
    .PrintNormal "Follow On Forms Message : " & sTemp
    
    ' WebForm items
    asItems = pwfElement.Items

    If UBound(asItems, 2) > 0 Then
      .PrintNormal "Items :"

      For iLoop = 1 To UBound(asItems, 2)
        Select Case CInt(asItems(2, iLoop))

          Case giWFFORMITEM_BUTTON
            .PrintNormal "     Button : " & asItems(9, iLoop)
            .PrintNormal "          Caption : " & asItems(3, iLoop)
            .PrintNormal "          Action : " & _
              IIf(CInt(asItems(54, iLoop)) = WORKFLOWBUTTONACTION_SAVEFORLATER, _
                "Save For Later", _
                IIf(CInt(asItems(54, iLoop)) = WORKFLOWBUTTONACTION_CANCEL, _
                  "Cancel", "Submit"))

          Case giWFFORMITEM_LABEL
            .PrintNormal "     Label"
            If CInt(asItems(57, iLoop)) = giWFDATAVALUE_CALC Then
              .PrintNormal "          Caption Type : Calculation"
              If CLng(asItems(56, iLoop)) > 0 Then
                .PrintNormal "               Calculation : " & GetExpressionName(CLng(asItems(56, iLoop)))
              Else
                .PrintNormal "               Calculation : None"
              End If
            Else
              .PrintNormal "          Caption Type : Fixed Value"
              .PrintNormal "               Fixed Value : " & asItems(3, iLoop)
            End If
            
          Case giWFFORMITEM_INPUTVALUE_CHAR, _
            giWFFORMITEM_INPUTVALUE_LOGIC, _
            giWFFORMITEM_INPUTVALUE_DATE, _
            giWFFORMITEM_INPUTVALUE_NUMERIC

            .PrintNormal "     Input value : " & asItems(9, iLoop)

            sTemp2 = ""
            sTemp3 = ""
            sTemp4 = ""
            sTemp5 = ""
            sTemp6 = ""
            
            Select Case asItems(6, iLoop)
              Case giEXPRVALUE_CHARACTER
                sTemp = "Character"
                sTemp2 = "Size : " & asItems(7, iLoop)
                
                If CInt(asItems(58, iLoop)) = giWFDATAVALUE_CALC Then
                  sTemp3 = "Default Value Type : Calculation"
                  If CLng(asItems(56, iLoop)) > 0 Then
                    sTemp4 = "     Calculation : " & GetExpressionName(CLng(asItems(56, iLoop)))
                  Else
                    sTemp4 = "     Calculation : None"
                  End If
                Else
                  sTemp3 = "Default Value Type : Fixed Value"
                  sTemp4 = "     Fixed Value : " & asItems(10, iLoop)
                End If
                
                sTemp5 = "Mandatory : " & IIf(CBool(asItems(55, iLoop)), "True", "False")
                sTemp6 = "Hide Text : " & IIf(CBool(asItems(65, iLoop)), "True", "False")
                
              Case giEXPRVALUE_NUMERIC
                sTemp = "Numeric"
                sTemp2 = "Size : " & asItems(7, iLoop)
                sTemp3 = "Decimals : " & asItems(8, iLoop)
                
                If CInt(asItems(58, iLoop)) = giWFDATAVALUE_CALC Then
                  sTemp4 = "Default Value Type : Calculation"
                  If CLng(asItems(56, iLoop)) > 0 Then
                    sTemp5 = "     Calculation : " & GetExpressionName(CLng(asItems(56, iLoop)))
                  Else
                    sTemp5 = "     Calculation : None"
                  End If
                Else
                  sTemp4 = "Default Value Type : Fixed Value"
                  sTemp5 = "     Fixed Value : " & asItems(10, iLoop)
                End If
                
                sTemp6 = "Mandatory : " & IIf(CBool(asItems(55, iLoop)), "True", "False")
              
              Case giEXPRVALUE_LOGIC
                sTemp = "Logic"
                sTemp2 = "Caption : " & asItems(3, iLoop)
                
                If CInt(asItems(58, iLoop)) = giWFDATAVALUE_CALC Then
                  sTemp3 = "Default Value Type : Calculation"
                  If CLng(asItems(56, iLoop)) > 0 Then
                    sTemp4 = "     Calculation : " & GetExpressionName(CLng(asItems(56, iLoop)))
                  Else
                    sTemp4 = "     Calculation : None"
                  End If
                Else
                  sTemp3 = "Default Value Type : Fixed Value"
                  sTemp4 = "     Fixed Value : " & asItems(10, iLoop)
                End If
                
              Case giEXPRVALUE_DATE
                sTemp = "Date"
                If CInt(asItems(58, iLoop)) = giWFDATAVALUE_CALC Then
                  sTemp2 = "Default Value Type : Calculation"
                  If CLng(asItems(56, iLoop)) > 0 Then
                    sTemp3 = "     Calculation : " & GetExpressionName(CLng(asItems(56, iLoop)))
                  Else
                    sTemp3 = "     Calculation : None"
                  End If
                Else
                  sTemp2 = "Default Value Type : Fixed Value"
                  sTemp3 = "     Fixed Value : " & asItems(10, iLoop)
                End If
                
                sTemp4 = "Mandatory : " & IIf(CBool(asItems(55, iLoop)), "True", "False")
                
              Case Else
                sTemp = "Unknown"
            End Select
            .PrintNormal "          Input Type : " & sTemp

            If Len(sTemp2) > 0 Then
              .PrintNormal "          " & sTemp2
            End If
            If Len(sTemp3) > 0 Then
              .PrintNormal "          " & sTemp3
            End If
            If Len(sTemp4) > 0 Then
              .PrintNormal "          " & sTemp4
            End If
            If Len(sTemp5) > 0 Then
              .PrintNormal "          " & sTemp5
            End If
            If Len(sTemp6) > 0 Then
              .PrintNormal "          " & sTemp6
            End If

          Case giWFFORMITEM_INPUTVALUE_DROPDOWN, _
            giWFFORMITEM_INPUTVALUE_OPTIONGROUP

            sTemp = ""
            sTemp5 = ""

            .PrintNormal "     Input value : " & asItems(9, iLoop)

            Select Case CInt(asItems(2, iLoop))
              Case giWFFORMITEM_INPUTVALUE_DROPDOWN:
                sTemp = "Dropdown"
                sTemp5 = "Mandatory : " & IIf(CBool(asItems(55, iLoop)), "True", "False")
              Case giWFFORMITEM_INPUTVALUE_OPTIONGROUP:
                sTemp = "Option Group"
                .PrintNormal "          Caption : " & asItems(3, iLoop)
            End Select

            .PrintNormal "          Input Type : " & sTemp
            .PrintNormal "          Control Values : "

            asItemValues = Split(asItems(47, iLoop), vbTab)
            For iLoop3 = 0 To UBound(asItemValues)
              .PrintNonBold Space$(20) & asItemValues(iLoop3)
            Next iLoop3
            
            If CInt(asItems(58, iLoop)) = giWFDATAVALUE_CALC Then
              .PrintNormal "          Default Value Type : Calculation"
              If CLng(asItems(56, iLoop)) > 0 Then
                .PrintNormal "               Calculation : " & GetExpressionName(CLng(asItems(56, iLoop)))
              Else
                .PrintNormal "               Calculation : None"
              End If
            Else
              .PrintNormal "          Default Value Type : Fixed Value"
              .PrintNormal "               Fixed Value : " & IIf(Len(asItems(10, iLoop)) > 0, asItems(10, iLoop), "<None>")
            End If
            
            If Len(sTemp5) > 0 Then
              .PrintNormal "          " & sTemp5
            End If

          Case giWFFORMITEM_INPUTVALUE_LOOKUP
            .PrintNormal "     Input value : " & asItems(9, iLoop)

            .PrintNormal "          Input Type : Lookup"
            .PrintNormal "          Column : " & GetColumnName(CLng(asItems(49, iLoop)))

            .PrintNormal "          Filter Lookup Values : " & IIf((CLng(asItems(67, iLoop)) > 0) And (Len(asItems(69, iLoop)) > 0), "Yes", "No")
            If (CLng(asItems(67, iLoop)) > 0) And (Len(asItems(69, iLoop)) > 0) Then
              .PrintNormal "               Filter Column : " & GetColumnName(CLng(asItems(67, iLoop)))
              
              gobjOperatorDefs.Initialise
              If gobjOperatorDefs.IsValidID(CLng(asItems(68, iLoop))) Then
                Set objOperatorDef = gobjOperatorDefs.Item("O" & Trim$(Str(CLng(asItems(68, iLoop)))))
                sOperatorName = objOperatorDef.Name
                Set objOperatorDef = Nothing
              Else
                sOperatorName = "<Unknown>"
              End If
              
              .PrintNormal "               Filter Operator : " & sOperatorName
              .PrintNormal "               Filter Value : " & asItems(69, iLoop)
            End If

            If CInt(asItems(58, iLoop)) = giWFDATAVALUE_CALC Then
              .PrintNormal "          Default Value Type : Calculation"
              If CLng(asItems(56, iLoop)) > 0 Then
                .PrintNormal "               Calculation : " & GetExpressionName(CLng(asItems(56, iLoop)))
              Else
                .PrintNormal "               Calculation : None"
              End If
            Else
              .PrintNormal "          Default Value Type : Fixed Value"
              .PrintNormal "               Fixed Value : " & IIf(Len(asItems(10, iLoop)) > 0, asItems(10, iLoop), "<None>")
            End If

            .PrintNormal "          Mandatory : " & IIf(CBool(asItems(55, iLoop)), "True", "False")

          Case giWFFORMITEM_FRAME
          
            .PrintNormal "     Frame"
            .PrintNormal "          Caption : " & asItems(3, iLoop)
            If Len(asItems(81, iLoop)) > 0 Then
              .PrintNormal "          Hotspot Identifier : " & asItems(81, iLoop)
            End If

          Case giWFFORMITEM_INPUTVALUE_GRID
            .PrintNormal "     Record Selector : " & asItems(9, iLoop)
            .PrintNormal "          Type : Record Selector"
            .PrintNormal "          Identifier : " & asItems(9, iLoop)
            .PrintNormal "          Use As Target Identifier : " & IIf(CBool(asItems(82, iLoop)), "True", "False")
            .PrintNormal "          Table : " & GetTableName(CLng(asItems(44, iLoop)))
            .PrintNormal "          Record : " & GetRecordSelectionDescription(CInt(asItems(5, iLoop)))

            If asItems(5, iLoop) = giWFRECSEL_IDENTIFIEDRECORD Then
              .PrintNormal "          Element : " & asItems(11, iLoop)
              Set wfElement2 = GetElementByIdentifier(asItems(11, iLoop))
              If Not wfElement2 Is Nothing Then
                If wfElement2.ElementType = elem_WebForm Then
                  .PrintNormal "          Record Selector : " & asItems(12, iLoop)
                End If
              End If
            End If
            
            If CLng(asItems(50, iLoop)) > 0 Then
              .PrintNormal "          Record Table : " & GetTableName(CLng(asItems(50, iLoop)))
            End If

            If CLng(asItems(52, iLoop)) > 0 Then
              .PrintNormal "          Record Order : " & GetOrderName(CLng(asItems(52, iLoop)))
            End If

            If CLng(asItems(53, iLoop)) > 0 Then
              .PrintNormal "          Record Filter : " & GetExpressionName(CLng(asItems(53, iLoop)))
            End If

            .PrintNormal "          Mandatory : " & IIf(CBool(asItems(55, iLoop)), "True", "False")
        
          Case giWFFORMITEM_INPUTVALUE_FILEUPLOAD
            .PrintNormal "     File Upload : " & asItems(9, iLoop)
            .PrintNormal "          Size : " & asItems(7, iLoop)
            .PrintNormal "          Mandatory : " & IIf(CBool(asItems(55, iLoop)), "True", "False")

            asItemValues = Split(asItems(66, iLoop), vbTab)
            If UBound(asItemValues) < 0 Then
              .PrintNormal "          File Extensions : None"
            Else
              .PrintNormal "          File Extensions : "
              For iLoop3 = 0 To UBound(asItemValues)
                .PrintNonBold Space$(20) & asItemValues(iLoop3)
              Next iLoop3
            End If
        
          Case giWFFORMITEM_WFVALUE, _
            giWFFORMITEM_WFFILE
            
            .PrintNormal "     Workflow value : " & asItems(11, iLoop) & "." & asItems(12, iLoop)
          
          Case giWFFORMITEM_DBVALUE, _
            giWFFORMITEM_DBFILE
            
            .PrintNormal "     Database value : " & GetColumnName(CLng(asItems(4, iLoop)))
            .PrintNormal "          Record : " & GetRecordSelectionDescription(CInt(asItems(5, iLoop)))
            If CInt(asItems(5, iLoop)) = giWFRECSEL_IDENTIFIEDRECORD Then
              .PrintNormal "          Element : " & asItems(11, iLoop)

              Set wfElement2 = GetElementByIdentifier(CStr(asItems(11, iLoop)))
              If Not wfElement2 Is Nothing Then
                If wfElement2.ElementType = elem_WebForm Then
                  .PrintNormal "          Record Selector : " & asItems(12, iLoop)
                End If
              End If
            End If

        End Select
      Next iLoop
    End If
  
    ' Web Form Validations
    asValidations = pwfElement.Validations

    If UBound(asValidations, 2) > 0 Then
      .PrintNormal "Validations :"

      For iLoop = 1 To UBound(asValidations, 2)
        ' Expression Name
        .PrintNormal "     Calculation : " & GetExpressionName(CLng(asValidations(1, iLoop)))
        .PrintNormal "          Type : " & WorkflowWebFormValidationTypeDescription(CInt(asValidations(2, iLoop)))
        .PrintNormal "          Message : " & asValidations(3, iLoop)
      Next iLoop
    End If
  End With
  
End Sub





Private Sub PrintElementDetails_Connector1(pwfElement As VB.Control)
  ' Print the definition of the given workflow element.
  With mobjPrinter
    .PrintNormal "Caption : " & pwfElement.Caption
  End With
  
End Sub





Private Sub PrintElementDetails_Or(pwfElement As VB.Control)
  ' Print the definition of the given workflow element.
  
  ' Not required for SummingJunction elements

End Sub

Private Sub PrintElementDetails_Connector2(pwfElement As VB.Control, pawfElements As Variant)
  ' Print the definition of the given workflow element.
  Dim iLoop As Integer

  With mobjPrinter
    .PrintNormal "Caption : " & pwfElement.Caption

    For iLoop = 1 To UBound(pawfElements)
      If pawfElements(iLoop).ControlIndex = pwfElement.ConnectorPairIndex Then
        .PrintNormal "Preceding Connector : " & CStr(iLoop)
        Exit For
      End If
    Next iLoop
  End With
  
End Sub

Private Sub PrintElementDetails_Email(pwfElement As VB.Control)
  ' Print the definition of the given workflow element.
  Dim sTemp As String
  Dim wfElement2 As VB.Control
  Dim iLoop As Integer
  Dim asItems() As String

  With mobjPrinter
    ' Caption
    .PrintNormal "Caption : " & pwfElement.Caption
    
    ' Email Record
    .PrintNormal "Email Record : " & GetRecordSelectionDescription(pwfElement.EmailRecord)
    
    If pwfElement.EmailRecord = giWFRECSEL_IDENTIFIEDRECORD Then
      ' Email Record - Element (if required)
      .PrintNormal "Element : " & pwfElement.RecordSelectorWebFormIdentifier

      Set wfElement2 = GetElementByIdentifier(pwfElement.RecordSelectorWebFormIdentifier)
      If Not wfElement2 Is Nothing Then
        If wfElement2.ElementType = elem_WebForm Then
          ' Email Record - Record Selector (if required)
          .PrintNormal "Record Selector : " & pwfElement.RecordSelectorIdentifier
        End If
      End If
    End If
    
    ' Email To
    sTemp = "Unknown"
    With recEmailAddrEdit
      .Index = "idxID"
      .Seek "=", pwfElement.EmailID

      If Not .NoMatch Then
        sTemp = recEmailAddrEdit!Name
      End If
    End With
    .PrintNormal "Email To : " & sTemp

    ' Email CC
    If pwfElement.EmailCCID > 0 Then
      sTemp = "Unknown"
      With recEmailAddrEdit
        .Index = "idxID"
        .Seek "=", pwfElement.EmailCCID
      
        If Not .NoMatch Then
          sTemp = recEmailAddrEdit!Name
        End If
      End With
      .PrintNormal "Email Copy : " & sTemp
    End If
    
    ' Email Subject
    .PrintNormal "Email Subject : " & pwfElement.EMailSubject
      
    ' EMail attachment
    sTemp = "<none>"

    Select Case pwfElement.Attachment_Type
      Case giWFEMAILITEM_DBVALUE
        sTemp = "Database value - " & GetColumnName(pwfElement.Attachment_DBColumnID)
      Case giWFEMAILITEM_WFVALUE
        sTemp = "Workflow value - " & pwfElement.Attachment_WFElementIdentifier & "." & pwfElement.Attachment_WFValueIdentifier
      Case giWFEMAILITEM_FILEATTACHMENT
        sTemp = "File - '" & pwfElement.Attachment_File & "'"
    End Select

    .PrintNormal "Email Attachment : " & sTemp

    ' Print Email items
    asItems = pwfElement.Items

    If UBound(asItems, 2) > 0 Then
      .PrintNormal "Items :"
      
      For iLoop = 1 To UBound(asItems, 2)
        
        Select Case CInt(asItems(2, iLoop))
          Case giWFFORMITEM_DBVALUE
            .PrintNormal "     Database value - " & GetColumnName(CLng(asItems(4, iLoop)))
            .PrintNormal "          Record : " & GetRecordSelectionDescription(CInt(asItems(5, iLoop)))
            If CInt(asItems(5, iLoop)) = giWFRECSEL_IDENTIFIEDRECORD Then
              .PrintNormal "          Element : " & asItems(13, iLoop)

              Set wfElement2 = GetElementByIdentifier(asItems(13, iLoop))
              If Not wfElement2 Is Nothing Then
                If wfElement2.ElementType = elem_WebForm Then
                  .PrintNormal "          Record Selector : " & asItems(14, iLoop)
                End If
              End If
            End If

          Case giWFFORMITEM_LABEL
            .PrintNormal "     Label - '" & asItems(3, iLoop) & "'"

          Case giWFFORMITEM_WFVALUE
            .PrintNormal "     Workflow value - " & asItems(11, iLoop) & "." & asItems(12, iLoop)

          Case giWFFORMITEM_FORMATCODE
            .PrintNormal "     Formatting - " & FormatDescription(asItems(3, iLoop))
          
          Case giWFFORMITEM_CALC
            .PrintNormal "     Calculated value - " & GetColumnName(CLng(asItems(4, iLoop)))
            If CLng(asItems(56, iLoop)) > 0 Then
              .PrintNormal "          Calculation : " & GetExpressionName(CLng(asItems(56, iLoop)))
            Else
              .PrintNormal "          Calculation : None"
            End If
            
        End Select
      Next iLoop
    End If
  End With
  
End Sub

Private Sub ReadDefinitionIntoArray(pawfElements As Variant, pwfElement As VB.Control)
  Dim wfLink As COAWF_Link
  Dim fFound As Boolean
  Dim iLoop As Integer

  fFound = False
  For iLoop = 1 To UBound(pawfElements)
    If pawfElements(iLoop).ControlIndex = pwfElement.ControlIndex Then
      fFound = True
      Exit For
    End If
  Next
  
  If Not fFound Then
    ReDim Preserve pawfElements(UBound(pawfElements) + 1)
    Set pawfElements(UBound(pawfElements)) = pwfElement
  
    For Each wfLink In ASRWFLink1
      If wfLink.StartElementIndex = pwfElement.ControlIndex Then
        ReadDefinitionIntoArray pawfElements, mcolwfElements(CStr(wfLink.EndElementIndex))
      End If
    Next wfLink
    Set wfLink = Nothing
    
    If pwfElement.ElementType = elem_Connector1 Then
      ReadDefinitionIntoArray pawfElements, mcolwfElements(CStr(pwfElement.ConnectorPairIndex))
    End If
  End If
  
End Sub

Private Function SelectedLinkCount() As Integer
  ' Count the selected links.
'  Dim wfLink As COAWF_Link
'  Dim iSelectedCount As Integer
'
'  ' Check that at least two elements have been selected.
'  iSelectedCount = 0
'  For Each wfLink In ASRWFLink1
'    If (wfLink.HighLighted) And (wfLink.Visible) Then iSelectedCount = iSelectedCount + 1
'  Next wfLink
'  Set wfLink = Nothing

'  SelectedLinkCount = iSelectedCount

  SelectedLinkCount = mcolwfSelectedLinks.Count
  
End Function

Private Function SelectedElementCount() As Integer
  ' Count the selected elements.
'  Dim wfElement As VB.Control
'  Dim iSelectedCount As Integer
  
'  iSelectedCount = 0
'  For Each wfElement In mcolwfElements
'    If (wfElement.HighLighted) And (wfElement.Visible) Then iSelectedCount = iSelectedCount + 1
'  Next wfElement
'  Set wfElement = Nothing
'
'  SelectedElementCount = iSelectedCount
  
  SelectedElementCount = mcolwfSelectedElements.Count
End Function


Private Function SelectedElementTypeCount(piElementType As ElementType) As Integer
  ' Count the selected elements of the given type.
  Dim wfElement As VB.Control
  Dim iSelectedCount As Integer
  
  iSelectedCount = 0
'  For Each wfElement In mcolwfElements
'    If (wfElement.HighLighted) And _
'      (wfElement.Visible) And _
'      (wfElement.ElementType = piElementType) Then iSelectedCount = iSelectedCount + 1
'  Next wfElement
'  Set wfElement = Nothing

  For Each wfElement In mcolwfSelectedElements
    If wfElement.ElementType = piElementType Then iSelectedCount = iSelectedCount + 1
  Next
  
  SelectedElementTypeCount = iSelectedCount
  
End Function

Private Sub ToggleElementAddMode(ByVal psMenuOption As String)
  Dim fStartingAddMode As Boolean
  
  fStartingAddMode = Not abMenu.Tools(psMenuOption).Checked
  
  CancelElementAddMode
  
  abMenu.Tools(psMenuOption).Checked = fStartingAddMode
  
  If fStartingAddMode Then
    Me.MouseIcon = LoadResPicture(psMenuOption, vbResCursor)
    Me.MousePointer = vbCustom
  End If

End Sub

Public Function UniqueIdentifier(psIdentifier As String, piIgnoreIndex As Integer) As Boolean
  ' Check if the given identifier is unique.
  ' Ignore the given element index.
  Dim wfElement As VB.Control
  Dim fUnique As Boolean
  Dim fElementOK As Boolean
  Dim iLoop As Integer
  
  fUnique = True
  
  For Each wfElement In mcolwfElements
    If wfElement.ControlIndex <> piIgnoreIndex _
      And UCase(Trim(wfElement.Identifier)) = UCase(Trim(psIdentifier)) Then
    
      fElementOK = wfElement.Visible
    
      If (Not fElementOK) Then
        ' Element might not be .visible but still valid
        ' if this method is called from the the Workflow properties screen.
        fElementOK = (wfElement.ControlIndex > 0)

        If (miLastActionFlag = giACTION_DELETECONTROLS) Then
          For iLoop = 1 To UBound(mactlUndoControls)
            If IsWorkflowElement(mactlUndoControls(iLoop)) Then
              If mactlUndoControls(iLoop).ControlIndex = wfElement.ControlIndex Then
                fElementOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
  
        'JPD 20060719 Fault 11339
        If fElementOK Then
          For iLoop = 1 To UBound(mactlClipboardControls)
            If IsWorkflowElement(mactlClipboardControls(iLoop)) Then
              If mactlClipboardControls(iLoop).ControlIndex = wfElement.ControlIndex Then
                fElementOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
      End If
    
      If fElementOK Then
        fUnique = False
        Exit For
      End If
    End If
  Next wfElement
  Set wfElement = Nothing
  
  UniqueIdentifier = fUnique
  
End Function

Public Sub UpdateIdentifiers(pwfBaseElement As VB.Control, _
  pawfIgnoreElements As Variant, _
  pavIdentifierLog As Variant, _
  Optional pasMessages As Variant)
  
  ' Loop through the array of original and new identifiers.
  ' Update any references to the original identifiers to refer to the new identifiers.
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iLoop4 As Integer
  Dim fElementIdentifierChanged As Boolean
  Dim fElementTableChanged As Boolean
  Dim fItemsChanged As Boolean
  Dim fItemTablesChanged As Boolean
  Dim wfTemp As VB.Control
  Dim fCheckElement As Boolean
  Dim wfLink As COAWF_Link
  Dim fElementOK As Boolean
  Dim asItems() As String
  Dim avColumns() As Variant
  Dim fFound As Boolean
  Dim frmUsage As frmUsage
  Dim asMessages() As String
  Dim sTemp As String
  Dim sSubMessage1 As String
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim objExpr As CExpression
  Dim objComp As CExprComponent
  Dim lngExprID As Long
  Dim sExprType As String
  Dim sExprName As String
  Dim sComponentType As String
  Dim fElementsNeedReviewing As Boolean
  Dim fExprsNeedReviewing As Boolean
  Dim fInvalidElement As Boolean
  Dim fInvalidItem As Boolean
  Dim alngValidTables() As Long
  Dim lngDBTableID As Long
  
  fItemsChanged = False
  fElementsNeedReviewing = False
  fExprsNeedReviewing = False
  
  ' Clear the array of validation messages
  ' Column 0 = The message
  ReDim asMessages(0)
  
  ' Check if ther elements own identifier was changed.
  ' Column 2 = original identifier
  ' Column 3 = new identifier
  fElementIdentifierChanged = (UCase(Trim(pavIdentifierLog(2, 0))) <> UCase(Trim(pavIdentifierLog(3, 0))))
  
  If pwfBaseElement.ElementType = elem_StoredData Then
    fElementTableChanged = (CLng(pavIdentifierLog(5, 0)) <> CLng(pavIdentifierLog(6, 0)))
  End If
  
  ' Check if any element item identifiers were changed.
  For iLoop = 1 To UBound(pavIdentifierLog, 2)
    If UCase(Trim(pavIdentifierLog(2, iLoop))) <> UCase(Trim(pavIdentifierLog(3, iLoop))) Then
      fItemsChanged = True
    End If
  
    If CLng(pavIdentifierLog(5, iLoop)) <> CLng(pavIdentifierLog(6, iLoop)) Then
      fItemTablesChanged = True
    End If
  
    If fItemsChanged _
      Or fItemTablesChanged Then
      
      Exit For
    End If
  Next iLoop
  
  If fElementIdentifierChanged _
    Or fElementTableChanged _
    Or fItemsChanged _
    Or fItemTablesChanged Then
    ' Either the element identifier has changed
    ' or the element (Stored Data) table has changed
    ' or an element item (eg. web form control) identifier has changed.
    ' Loop through the other elements looking for references.
    For Each wfTemp In mcolwfElements
      fCheckElement = True
      
      For iLoop = 1 To UBound(pawfIgnoreElements)
        If wfTemp Is pawfIgnoreElements(iLoop) Then
          fCheckElement = False
          Exit For
        End If
      Next iLoop
      
      If fCheckElement Then
        With wfTemp
          Select Case .ElementType
            '--------------------------------------------------------
            Case elem_Decision
              ' Only need to check Decision elements that immediately follow the given element.
              ' Update the element's TrueFlowIdentifier if it changed.
              If (Not pwfBaseElement Is Nothing) _
                And (fItemsChanged Or fElementIdentifierChanged) Then
                
                If pwfBaseElement.ElementType = elem_WebForm Then
                  fElementOK = False
                  
                  For Each wfLink In ASRWFLink1
                    If (wfLink.StartElementIndex = pwfBaseElement.ControlIndex) _
                      And (wfLink.EndElementIndex = .ControlIndex) Then
                      
                      fElementOK = wfLink.Visible
                      If (Not fElementOK) Then
                        ' Link might not be .visible but still valid
                        ' if this method is called from the web form designer.
                        fElementOK = True
                        
                        If (miLastActionFlag = giACTION_DELETECONTROLS) Then
                          For iLoop = 1 To UBound(mactlUndoControls)
                            If TypeOf mactlUndoControls(iLoop) Is COAWF_Link Then
                              If mactlUndoControls(iLoop).Index = wfLink.Index Then
                                fElementOK = False
                                Exit For
                              End If
                            End If
                          Next iLoop
                        End If
                        
                        'JPD 20060719 Fault 11339
                        If fElementOK Then
                          For iLoop = 1 To UBound(mactlClipboardControls)
                            If TypeOf mactlClipboardControls(iLoop) Is COAWF_Link Then
                              If mactlClipboardControls(iLoop).Index = wfLink.Index Then
                                fElementOK = False
                                Exit For
                              End If
                            End If
                          Next iLoop
                        End If
                      End If
                    End If
                  Next wfLink
                  Set wfLink = Nothing
                  
                  If fElementOK Then
                    If Len(pavIdentifierLog(3, 0)) = 0 Then
                      ' Base element deleted
                      ' Identifier object deleted, not just renamed. Need to inform the user.
                      sTemp = GetDecisionCaptionDescription(wfTemp.DecisionCaptionType, True)
                      
                      ReDim Preserve asMessages(UBound(asMessages) + 1)
                      asMessages(UBound(asMessages)) = _
                        ValidateElement_MessagePrefix(wfTemp) & _
                        "Invalid '" & sTemp & "' flow button selected"
                    
                      fElementsNeedReviewing = True
                    Else
                      ' Base element still exists, maybe button deleted/renamed.
                      For iLoop = 1 To UBound(pavIdentifierLog, 2)
                        If UCase(Trim(.TrueFlowIdentifier)) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then

                          If Len(pavIdentifierLog(3, iLoop)) = 0 Then
                            ' Identifier object deleted, not just renamed. Inform the user.
                            sTemp = GetDecisionCaptionDescription(wfTemp.DecisionCaptionType, True)
                            
                            ReDim Preserve asMessages(UBound(asMessages) + 1)
                            asMessages(UBound(asMessages)) = _
                              ValidateElement_MessagePrefix(wfTemp) & _
                              "Invalid '" & sTemp & "' flow button selected"

                            fElementsNeedReviewing = True
                          Else
                            .TrueFlowIdentifier = pavIdentifierLog(3, iLoop)
                          End If

                          Exit For
                        End If
                      Next iLoop
                    End If
                  End If
                End If
              End If
              
            '--------------------------------------------------------
            Case elem_Email
              If .EmailRecord = giWFRECSEL_IDENTIFIEDRECORD _
                And UCase(Trim(.RecordSelectorWebFormIdentifier)) = UCase(Trim(pavIdentifierLog(2, 0))) Then
              
                lngDBTableID = 0

                If .EmailID > 0 Then
                  With recEmailAddrEdit
                    .Index = "idxID"
                    .Seek "=", wfTemp.EmailID
              
                    ' Read the expression's tableID from the recordset.
                    If Not .NoMatch Then
                      lngDBTableID = !TableID
                    End If
                  End With
                End If
                If (lngDBTableID = 0) And (.EmailCCID > 0) Then
                  With recEmailAddrEdit
                    .Index = "idxID"
                    .Seek "=", wfTemp.EmailCCID
              
                    ' Read the expression's tableID from the recordset.
                    If Not .NoMatch Then
                      lngDBTableID = !TableID
                    End If
                  End With
                End If
              
                fInvalidElement = (Len(pavIdentifierLog(3, 0)) = 0)
                If (Not fInvalidElement) Then
                  .RecordSelectorWebFormIdentifier = pavIdentifierLog(3, 0)
                
                  If (Not pwfBaseElement Is Nothing) Then
                    If pwfBaseElement.ElementType = elem_StoredData Then
                      ' Check if the StoredData table has changed.
                      
                      If (CLng(pavIdentifierLog(5, 0)) <> CLng(pavIdentifierLog(6, 0))) _
                        And (lngDBTableID > 0) Then

                        ReDim alngValidTables(0)
                        TableAscendants CLng(pavIdentifierLog(6, 0)), alngValidTables

                        fInvalidElement = True

                        For iLoop4 = 1 To UBound(alngValidTables)
                          If (alngValidTables(iLoop4) = lngDBTableID) Then
                            fInvalidElement = False
                            Exit For
                          End If
                        Next iLoop4
                      End If
                    End If
                  End If
                End If
                
                If fInvalidElement Then
                  ' Identifier object deleted, not just renamed. Inform the user.
                  ReDim Preserve asMessages(UBound(asMessages) + 1)
                  asMessages(UBound(asMessages)) = _
                    ValidateElement_MessagePrefix(wfTemp) & _
                    "Invalid email record element"
                
                  fElementsNeedReviewing = True
                End If
                
                If Not pwfBaseElement Is Nothing Then
                  If pwfBaseElement.ElementType = elem_WebForm Then
                    For iLoop = 1 To UBound(pavIdentifierLog, 2)
                      If UCase(Trim(.RecordSelectorIdentifier)) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then
                        
                        fInvalidItem = (Len(pavIdentifierLog(3, iLoop)) = 0)
                        
                        If (Not fInvalidItem) Then
                          .RecordSelectorIdentifier = pavIdentifierLog(3, iLoop)
                        
                          ' Check if the recordSelector table has changed.
                          If (CLng(pavIdentifierLog(5, iLoop)) <> CLng(pavIdentifierLog(6, iLoop))) _
                            And (lngDBTableID > 0) Then

                            ReDim alngValidTables(0)
                            TableAscendants CLng(pavIdentifierLog(6, iLoop)), alngValidTables

                            fInvalidItem = True

                            For iLoop4 = 1 To UBound(alngValidTables)
                              If (alngValidTables(iLoop4) = lngDBTableID) Then
                                fInvalidItem = False
                                Exit For
                              End If
                            Next iLoop4
                          End If
                        End If
                        
                        If fInvalidItem Then
                          ' Identifier object deleted, not just renamed. Inform the user.
                          ReDim Preserve asMessages(UBound(asMessages) + 1)
                          asMessages(UBound(asMessages)) = _
                            ValidateElement_MessagePrefix(wfTemp) & _
                            "Invalid email record selector"
                        
                          fElementsNeedReviewing = True
                        End If
                        
                        Exit For
                      End If
                    Next iLoop
                  End If
                End If
              End If
              
              ' Check for changes to identifiers used in DBValue and WFValue items.
              asItems = wfTemp.Items
              
              For iLoop2 = 1 To UBound(asItems, 2)
                If CInt(asItems(2, iLoop2)) = giWFFORMITEM_DBVALUE Then
                  If CInt(asItems(5, iLoop2)) = giWFRECSEL_IDENTIFIEDRECORD _
                    And UCase(Trim(asItems(13, iLoop2))) = UCase(Trim(pavIdentifierLog(2, 0))) Then
                  
                    lngDBTableID = GetTableIDFromColumnID(CLng(asItems(4, iLoop2)))

                    fInvalidElement = (Len(pavIdentifierLog(3, 0)) = 0)
                    If (Not fInvalidElement) Then
                      asItems(13, iLoop2) = pavIdentifierLog(3, 0)
                    
                      If (Not pwfBaseElement Is Nothing) Then
                        If pwfBaseElement.ElementType = elem_StoredData Then
                          ' Check if the StoredData table has changed.
                          
                          If (CLng(pavIdentifierLog(5, 0)) <> CLng(pavIdentifierLog(6, 0))) _
                            And (lngDBTableID > 0) Then

                            ReDim alngValidTables(0)
                            TableAscendants CLng(pavIdentifierLog(6, 0)), alngValidTables

                            fInvalidElement = True

                            For iLoop4 = 1 To UBound(alngValidTables)
                              If (alngValidTables(iLoop4) = lngDBTableID) Then
                                fInvalidElement = False
                                Exit For
                              End If
                            Next iLoop4
                          End If
                        End If
                      End If
                    End If
    
                    If fInvalidElement Then
                      ' Identifier object deleted, not just renamed. Inform the user.
                      sSubMessage1 = " (" & GetColumnName(CLng(asItems(4, iLoop2))) & ")"
  
                      ReDim Preserve asMessages(UBound(asMessages) + 1)
                      asMessages(UBound(asMessages)) = _
                        ValidateElement_MessagePrefix(wfTemp) & _
                        "Database Value" & sSubMessage1 & " - Invalid element identifier"
                    
                      fElementsNeedReviewing = True
                    End If
                    
                    If Not pwfBaseElement Is Nothing Then
                      If pwfBaseElement.ElementType = elem_WebForm Then
                        ' Update the element recSel identifier.
                        For iLoop = 1 To UBound(pavIdentifierLog, 2)
                          If UCase(Trim(asItems(14, iLoop2))) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then
                            
                            fInvalidItem = Len(pavIdentifierLog(3, iLoop)) = 0
                            
                            If (Not fInvalidItem) Then
                              asItems(14, iLoop2) = pavIdentifierLog(3, iLoop)
                        
                              ' Check if the recordSelector table has changed.
                             If (CLng(pavIdentifierLog(5, iLoop)) <> CLng(pavIdentifierLog(6, iLoop))) _
                               And (lngDBTableID > 0) Then

                                ReDim alngValidTables(0)
                                TableAscendants CLng(pavIdentifierLog(6, iLoop)), alngValidTables

                                fInvalidItem = True

                                For iLoop4 = 1 To UBound(alngValidTables)
                                  If (alngValidTables(iLoop4) = lngDBTableID) Then
                                    fInvalidItem = False
                                    Exit For
                                  End If
                                Next iLoop4
                              End If
                            End If
                            
                            If fInvalidItem Then
                              ' Identifier object deleted, not just renamed. Inform the user.
                              sSubMessage1 = " (" & GetColumnName(CLng(asItems(4, iLoop2))) & ")"
                            
                              ReDim Preserve asMessages(UBound(asMessages) + 1)
                              asMessages(UBound(asMessages)) = _
                                ValidateElement_MessagePrefix(wfTemp) & _
                                "Database Value" & sSubMessage1 & " - Invalid record selector"
                          
                              fElementsNeedReviewing = True
                            End If
                            
                            Exit For
                          End If
                        Next iLoop
                      End If
                    End If
                  End If
                ElseIf (CInt(asItems(2, iLoop2)) = giWFFORMITEM_WFVALUE) _
                  Or (CInt(asItems(2, iLoop2)) = giWFFORMITEM_WFFILE) Then
                
                  If UCase(Trim(asItems(11, iLoop2))) = UCase(Trim(pavIdentifierLog(2, 0))) Then
                  
                    ' Update the element identifier.
                    If Len(pavIdentifierLog(3, 0)) = 0 Then
                      ' Identifier object deleted, not just renamed. Inform the user.
                      sSubMessage1 = " (" & asItems(11, iLoop2) & "." & asItems(12, iLoop2) & ")"
  
                      ReDim Preserve asMessages(UBound(asMessages) + 1)
                      asMessages(UBound(asMessages)) = _
                        ValidateElement_MessagePrefix(wfTemp) & _
                        "Workflow Value" & sSubMessage1 & " - Invalid web form identifier"
                    
                      fElementsNeedReviewing = True
                    Else
                      asItems(11, iLoop2) = pavIdentifierLog(3, 0)
                      asItems(1, iLoop2) = "Workflow value - " & asItems(11, iLoop2) & "." & asItems(12, iLoop2)
                    End If
                  
                    ' Update the element item identifier.
                    For iLoop = 1 To UBound(pavIdentifierLog, 2)
                      If UCase(Trim(asItems(12, iLoop2))) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then
                        If Len(pavIdentifierLog(3, iLoop)) = 0 Then
                          ' Identifier object deleted, not just renamed. Inform the user.
                          sSubMessage1 = " (" & asItems(11, iLoop2) & "." & asItems(12, iLoop2) & ")"
                        
                          ReDim Preserve asMessages(UBound(asMessages) + 1)
                          asMessages(UBound(asMessages)) = _
                            ValidateElement_MessagePrefix(wfTemp) & _
                            "Workflow Value" & sSubMessage1 & " - Invalid value identifier"
                        
                          fElementsNeedReviewing = True
                        Else
                          asItems(12, iLoop2) = pavIdentifierLog(3, iLoop)
                          asItems(1, iLoop2) = "Workflow value - " & asItems(11, iLoop2) & "." & asItems(12, iLoop2)
                        End If
                        Exit For
                      End If
                    Next iLoop
                    
                  End If
                End If
              Next iLoop2

              wfTemp.Items = asItems
              
            '--------------------------------------------------------
            Case elem_StoredData
              If .DataRecord = giWFRECSEL_IDENTIFIEDRECORD _
                And UCase(Trim(.RecordSelectorWebFormIdentifier)) = UCase(Trim(pavIdentifierLog(2, 0))) Then

                lngDBTableID = .DataRecordTableID

                fInvalidElement = (Len(pavIdentifierLog(3, 0)) = 0)
                If (Not fInvalidElement) Then
                  .RecordSelectorWebFormIdentifier = pavIdentifierLog(3, 0)
                
                  If (Not pwfBaseElement Is Nothing) Then
                    If pwfBaseElement.ElementType = elem_StoredData Then
                      ' Check if the StoredData table has changed.
                      If (CLng(pavIdentifierLog(5, 0)) <> CLng(pavIdentifierLog(6, 0))) _
                        And (lngDBTableID > 0) Then

                        ReDim alngValidTables(0)
                        TableAscendants CLng(pavIdentifierLog(6, 0)), alngValidTables

                        fInvalidElement = True

                        For iLoop4 = 1 To UBound(alngValidTables)
                          If (alngValidTables(iLoop4) = lngDBTableID) Then
                            fInvalidElement = False
                            Exit For
                          End If
                        Next iLoop4
                      End If
                    End If
                  End If
                End If
                
                If fInvalidElement Then
                  ' Identifier object deleted, not just renamed. Inform the user.
                  ReDim Preserve asMessages(UBound(asMessages) + 1)
                  asMessages(UBound(asMessages)) = _
                    ValidateElement_MessagePrefix(wfTemp) & _
                    "Invalid primary record element"
                
                  fElementsNeedReviewing = True
                End If

                If Not pwfBaseElement Is Nothing Then
                  If pwfBaseElement.ElementType = elem_WebForm Then
                    For iLoop = 1 To UBound(pavIdentifierLog, 2)
                      If UCase(Trim(.RecordSelectorIdentifier)) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then
                        
                        fInvalidItem = Len(pavIdentifierLog(3, iLoop)) = 0
                        
                        If (Not fInvalidItem) Then
                          .RecordSelectorIdentifier = pavIdentifierLog(3, iLoop)
                    
                          ' Check if the recordSelector table has changed.
                          If (CLng(pavIdentifierLog(5, iLoop)) <> CLng(pavIdentifierLog(6, iLoop))) _
                            And (lngDBTableID > 0) Then

                            ReDim alngValidTables(0)
                            TableAscendants CLng(pavIdentifierLog(6, iLoop)), alngValidTables

                            fInvalidItem = True

                            For iLoop4 = 1 To UBound(alngValidTables)
                              If (alngValidTables(iLoop4) = lngDBTableID) Then
                                fInvalidItem = False
                                Exit For
                              End If
                            Next iLoop4
                          End If
                        End If
                        
                        If fInvalidItem Then
                          ' Identifier object deleted, not just renamed. Inform the user.
                          ReDim Preserve asMessages(UBound(asMessages) + 1)
                          asMessages(UBound(asMessages)) = _
                            ValidateElement_MessagePrefix(wfTemp) & _
                            "Invalid primary record selector"
                        
                          fElementsNeedReviewing = True
                        End If
                        
                        Exit For
                      End If
                    Next iLoop
                  End If
                End If
              End If

              If .SecondaryDataRecord = giWFRECSEL_IDENTIFIEDRECORD _
                And UCase(Trim(.SecondaryRecordSelectorWebFormIdentifier)) = UCase(Trim(pavIdentifierLog(2, 0))) Then

                lngDBTableID = .SecondaryDataRecordTableID
                
                fInvalidElement = (Len(pavIdentifierLog(3, 0)) = 0)
                If (Not fInvalidElement) Then
                  .SecondaryRecordSelectorWebFormIdentifier = pavIdentifierLog(3, 0)
                
                  If (Not pwfBaseElement Is Nothing) Then
                    If pwfBaseElement.ElementType = elem_StoredData Then
                      ' Check if the StoredData table has changed.
                      If (CLng(pavIdentifierLog(5, 0)) <> CLng(pavIdentifierLog(6, 0))) _
                        And (lngDBTableID > 0) Then

                        ReDim alngValidTables(0)
                        TableAscendants CLng(pavIdentifierLog(6, 0)), alngValidTables

                        fInvalidElement = True

                        For iLoop4 = 1 To UBound(alngValidTables)
                          If (alngValidTables(iLoop4) = lngDBTableID) Then
                            fInvalidElement = False
                            Exit For
                          End If
                        Next iLoop4
                      End If
                    End If
                  End If
                End If

                If fInvalidElement Then
                  ' Identifier object deleted, not just renamed. Inform the user.
                  ReDim Preserve asMessages(UBound(asMessages) + 1)
                  asMessages(UBound(asMessages)) = _
                    ValidateElement_MessagePrefix(wfTemp) & _
                    "Invalid secondary record element"
                
                  fElementsNeedReviewing = True
                End If

                If Not pwfBaseElement Is Nothing Then
                  If pwfBaseElement.ElementType = elem_WebForm Then
                    For iLoop = 1 To UBound(pavIdentifierLog, 2)
                      If UCase(Trim(.SecondaryRecordSelectorIdentifier)) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then
                        
                        fInvalidItem = Len(pavIdentifierLog(3, iLoop)) = 0
                        
                        If (Not fInvalidItem) Then
                          .SecondaryRecordSelectorIdentifier = pavIdentifierLog(3, iLoop)
                    
                          ' Check if the recordSelector table has changed.
                          If (CLng(pavIdentifierLog(5, iLoop)) <> CLng(pavIdentifierLog(6, iLoop))) _
                            And (lngDBTableID > 0) Then

                            ReDim alngValidTables(0)
                            TableAscendants CLng(pavIdentifierLog(6, iLoop)), alngValidTables

                            fInvalidItem = True

                            For iLoop4 = 1 To UBound(alngValidTables)
                              If (alngValidTables(iLoop4) = lngDBTableID) Then
                                fInvalidItem = False
                                Exit For
                              End If
                            Next iLoop4
                          End If
                        End If
                        
                        If fInvalidItem Then
                          ' Identifier object deleted, not just renamed. Inform the user.
                          ReDim Preserve asMessages(UBound(asMessages) + 1)
                          asMessages(UBound(asMessages)) = _
                            ValidateElement_MessagePrefix(wfTemp) & _
                            "Invalid secondary record selector"
                        
                          fElementsNeedReviewing = True
                        End If
                        
                        Exit For
                      End If
                    Next iLoop
                  End If
                End If
              End If

              ' Check for changes to identifiers used in WFValue items.
              avColumns = wfTemp.DataColumns

              For iLoop2 = 1 To UBound(avColumns, 2)
                If CInt(avColumns(4, iLoop2)) = giWFDATAVALUE_WFVALUE Then
                  If UCase(Trim(avColumns(6, iLoop2))) = UCase(Trim(pavIdentifierLog(2, 0))) Then

                    ' Update the element identifier.
                    If Len(pavIdentifierLog(3, 0)) = 0 Then
                      ' Identifier object deleted, not just renamed. Inform the user.
                      sSubMessage1 = GetColumnName(CLng(avColumns(3, iLoop2)), True)
  
                      ReDim Preserve asMessages(UBound(asMessages) + 1)
                      asMessages(UBound(asMessages)) = _
                        ValidateElement_MessagePrefix(wfTemp) & _
                        "Column (" & sSubMessage1 & ") - Invalid workflow value web form identifier"
                    
                      fElementsNeedReviewing = True
                    Else
                      avColumns(6, iLoop2) = pavIdentifierLog(3, 0)
                    End If

                    ' Update the element recSel identifier.
                    For iLoop = 1 To UBound(pavIdentifierLog, 2)
                      If UCase(Trim(avColumns(7, iLoop2))) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then
                        
                        If Len(pavIdentifierLog(3, iLoop)) = 0 Then
                          ' Identifier object deleted, not just renamed. Inform the user.
                          sSubMessage1 = GetColumnName(CLng(avColumns(3, iLoop2)), True)
  
                          ReDim Preserve asMessages(UBound(asMessages) + 1)
                          asMessages(UBound(asMessages)) = _
                            ValidateElement_MessagePrefix(wfTemp) & _
                            "Column (" & sSubMessage1 & ") -  Invalid workflow value identifier"
                        
                          fElementsNeedReviewing = True
                        Else
                          avColumns(7, iLoop2) = pavIdentifierLog(3, iLoop)
                        End If
                        
                        Exit For
                      End If
                    Next iLoop

                    avColumns(2, iLoop2) = "Workflow value - " & avColumns(6, iLoop2) & "." & avColumns(7, iLoop2)
                  End If
                ElseIf CInt(avColumns(4, iLoop2)) = giWFDATAVALUE_DBVALUE Then
                  lngDBTableID = GetTableIDFromColumnID(CLng(avColumns(8, iLoop2)))
                  
                  If UCase(Trim(avColumns(6, iLoop2))) = UCase(Trim(pavIdentifierLog(2, 0))) Then
                  
                    fInvalidElement = (Len(pavIdentifierLog(3, 0)) = 0)
                    If (Not fInvalidElement) Then
                      avColumns(6, iLoop2) = pavIdentifierLog(3, 0)
                    
                      If (Not pwfBaseElement Is Nothing) Then
                        If pwfBaseElement.ElementType = elem_StoredData Then
                          ' Check if the StoredData table has changed.
                          If (CLng(pavIdentifierLog(5, 0)) <> CLng(pavIdentifierLog(6, 0))) _
                            And (lngDBTableID > 0) Then

                            ReDim alngValidTables(0)
                            TableAscendants CLng(pavIdentifierLog(6, 0)), alngValidTables

                            fInvalidElement = True

                            For iLoop4 = 1 To UBound(alngValidTables)
                              If (alngValidTables(iLoop4) = lngDBTableID) Then
                                fInvalidElement = False
                                Exit For
                              End If
                            Next iLoop4
                          End If
                        End If
                      End If
                    End If
                    
                    If fInvalidElement Then
                      ' Identifier object deleted, not just renamed. Inform the user.
                      sSubMessage1 = GetColumnName(CLng(avColumns(3, iLoop2)), True)

                      ReDim Preserve asMessages(UBound(asMessages) + 1)
                      asMessages(UBound(asMessages)) = _
                        ValidateElement_MessagePrefix(wfTemp) & _
                        "Column (" & sSubMessage1 & ") - Invalid database value element identifier"
                    
                      fElementsNeedReviewing = True
                    End If
                  
                    ' Update the element recSel identifier.
                    If Not pwfBaseElement Is Nothing Then
                      If pwfBaseElement.ElementType = elem_WebForm Then
                        For iLoop = 1 To UBound(pavIdentifierLog, 2)
                          If UCase(Trim(avColumns(7, iLoop2))) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then
                            
                            fInvalidItem = Len(pavIdentifierLog(3, iLoop)) = 0
                            
                            If (Not fInvalidItem) Then
                              avColumns(7, iLoop2) = pavIdentifierLog(3, iLoop)
                        
                              ' Check if the recordSelector table has changed.
                             If (CLng(pavIdentifierLog(5, iLoop)) <> CLng(pavIdentifierLog(6, iLoop))) _
                               And (lngDBTableID > 0) Then

                                ReDim alngValidTables(0)
                                TableAscendants CLng(pavIdentifierLog(6, iLoop)), alngValidTables

                                fInvalidItem = True

                                For iLoop4 = 1 To UBound(alngValidTables)
                                  If (alngValidTables(iLoop4) = lngDBTableID) Then
                                    fInvalidItem = False
                                    Exit For
                                  End If
                                Next iLoop4
                              End If
                            End If
                            
                            If fInvalidItem Then
                              ' Identifier object deleted, not just renamed. Inform the user.
                              sSubMessage1 = GetColumnName(CLng(avColumns(3, iLoop2)), True)
                            
                              ReDim Preserve asMessages(UBound(asMessages) + 1)
                              asMessages(UBound(asMessages)) = _
                                ValidateElement_MessagePrefix(wfTemp) & _
                                "Column (" & sSubMessage1 & ") - Invalid database value record selector"
                              
                              fElementsNeedReviewing = True
                            End If
                        
                            Exit For
                          End If
                        Next iLoop
                      End If
                    End If
                  End If
                End If
              Next iLoop2

              wfTemp.DataColumns = avColumns
            
            '--------------------------------------------------------
            Case elem_WebForm
              ' Check for changes to identifiers used in DBValue, RecSel and WFValue items.
              asItems = wfTemp.Items
              
              For iLoop2 = 1 To UBound(asItems, 2)
                If (CInt(asItems(2, iLoop2)) = giWFFORMITEM_DBVALUE) _
                  Or (CInt(asItems(2, iLoop2)) = giWFFORMITEM_INPUTVALUE_GRID) Then
                  
                  If CInt(asItems(5, iLoop2)) = giWFRECSEL_IDENTIFIEDRECORD _
                    And UCase(Trim(asItems(11, iLoop2))) = UCase(Trim(pavIdentifierLog(2, 0))) Then

                    lngDBTableID = GetTableIDFromColumnID(CLng(asItems(4, iLoop2)))
                    
                    fInvalidElement = (Len(pavIdentifierLog(3, 0)) = 0)
                    If (Not fInvalidElement) Then
                      asItems(11, iLoop2) = pavIdentifierLog(3, 0)
                    
                      If (Not pwfBaseElement Is Nothing) Then
                        If pwfBaseElement.ElementType = elem_StoredData Then
                          ' Check if the StoredData table has changed.
                          If (CLng(pavIdentifierLog(5, 0)) <> CLng(pavIdentifierLog(6, 0))) _
                            And (lngDBTableID > 0) Then

                            ReDim alngValidTables(0)
                            TableAscendants CLng(pavIdentifierLog(6, 0)), alngValidTables

                            fInvalidElement = True

                            For iLoop4 = 1 To UBound(alngValidTables)
                              If (alngValidTables(iLoop4) = lngDBTableID) Then
                                fInvalidElement = False
                                Exit For
                              End If
                            Next iLoop4
                          End If
                        End If
                      End If
                    End If
                    
                    If fInvalidElement Then
                      ' Identifier object deleted, not just renamed. Inform the user.
                      fElementsNeedReviewing = True
                      ReDim Preserve asMessages(UBound(asMessages) + 1)
                      
                      If (CInt(asItems(2, iLoop2)) = giWFFORMITEM_DBVALUE) Then
                        sSubMessage1 = " (" & GetColumnName(CLng(asItems(4, iLoop2))) & ")"
                        asMessages(UBound(asMessages)) = _
                          ValidateElement_MessagePrefix(wfTemp) & _
                          "Database Value " & sSubMessage1 & " - Invalid element identifier"
                      Else
                        asMessages(UBound(asMessages)) = _
                          ValidateElement_MessagePrefix(wfTemp) & _
                          "Record Selector Input (" & asItems(9, iLoop2) & ") - Invalid element identifier"
                      End If
                    End If

                    If Not pwfBaseElement Is Nothing Then
                      If pwfBaseElement.ElementType = elem_WebForm Then
                        ' Update the element recSel identifier.
                        For iLoop = 1 To UBound(pavIdentifierLog, 2)
                          If UCase(Trim(asItems(12, iLoop2))) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then
                            
                            fInvalidItem = Len(pavIdentifierLog(3, iLoop)) = 0
                            
                            If (Not fInvalidItem) Then
                              asItems(12, iLoop2) = pavIdentifierLog(3, iLoop)
                        
                              ' Check if the recordSelector table has changed.
                             If (CLng(pavIdentifierLog(5, iLoop)) <> CLng(pavIdentifierLog(6, iLoop))) _
                               And (lngDBTableID > 0) Then

                                ReDim alngValidTables(0)
                                TableAscendants CLng(pavIdentifierLog(6, iLoop)), alngValidTables

                                fInvalidItem = True

                                For iLoop4 = 1 To UBound(alngValidTables)
                                  If (alngValidTables(iLoop4) = lngDBTableID) Then
                                    fInvalidItem = False
                                    Exit For
                                  End If
                                Next iLoop4
                              End If
                            End If
                            
                            If fInvalidItem Then
                              ' Identifier object deleted, not just renamed. Inform the user.
                              ReDim Preserve asMessages(UBound(asMessages) + 1)
                              fElementsNeedReviewing = True

                              If (CInt(asItems(2, iLoop2)) = giWFFORMITEM_DBVALUE) Then
                                sSubMessage1 = " (" & GetColumnName(CLng(asItems(4, iLoop2))) & ")"

                                asMessages(UBound(asMessages)) = _
                                  ValidateElement_MessagePrefix(wfTemp) & _
                                  "Database Value" & sSubMessage1 & " - Invalid record selector"
                              Else
                                asMessages(UBound(asMessages)) = _
                                  ValidateElement_MessagePrefix(wfTemp) & _
                                  "Record Selector Input (" & asItems(9, iLoop2) & ") - Invalid record selector"
                              End If
                            End If
                            
                            Exit For
                          End If
                        Next iLoop
                      End If
                    End If
                  End If
                ElseIf (CInt(asItems(2, iLoop2)) = giWFFORMITEM_WFVALUE) _
                  Or (CInt(asItems(2, iLoop2)) = giWFFORMITEM_WFFILE) Then
                  
                  If UCase(Trim(asItems(11, iLoop2))) = UCase(Trim(pavIdentifierLog(2, 0))) Then

                    ' Update the element identifier.
                    If Len(pavIdentifierLog(3, 0)) = 0 Then
                      ' Identifier object deleted, not just renamed. Inform the user.
                      sSubMessage1 = " (" & asItems(11, iLoop2) & ")"
  
                      ReDim Preserve asMessages(UBound(asMessages) + 1)
                      asMessages(UBound(asMessages)) = _
                        ValidateElement_MessagePrefix(wfTemp) & _
                        "Workflow Value" & sSubMessage1 & " - Invalid web form identifier"
                      fElementsNeedReviewing = True
                    Else
                      asItems(11, iLoop2) = pavIdentifierLog(3, 0)
                    End If

                    ' Update the element recSel identifier.
                    For iLoop = 1 To UBound(pavIdentifierLog, 2)
                      If UCase(Trim(asItems(12, iLoop2))) = UCase(Trim(pavIdentifierLog(2, iLoop))) Then
                        If Len(pavIdentifierLog(3, iLoop)) = 0 Then
                          ' Identifier object deleted, not just renamed. Inform the user.
                          sSubMessage1 = " (" & asItems(11, iLoop2) & "." & asItems(12, iLoop2) & ")"
  
                          ReDim Preserve asMessages(UBound(asMessages) + 1)
                          asMessages(UBound(asMessages)) = _
                            ValidateElement_MessagePrefix(wfTemp) & _
                            "Workflow Value" & sSubMessage1 & " - Invalid value identifier"
                          fElementsNeedReviewing = True
                        Else
                          asItems(12, iLoop2) = pavIdentifierLog(3, iLoop)
                        End If
                        
                        Exit For
                      End If
                    Next iLoop
                  End If
                End If
              Next iLoop2

              wfTemp.Items = asItems
           
           End Select
        End With
      End If
    Next wfTemp
    Set wfTemp = Nothing
    
    ' Update the identifiers in any of this Workflow's expressions
    sSQL = "SELECT tmpComponents.componentID, tmpComponents.type, tmpComponents.workflowItem, tmpComponents.workflowRecordTableID" & _
      " FROM tmpComponents" & _
      " INNER JOIN tmpExpressions ON tmpComponents.exprID = tmpExpressions.exprID" & _
      " WHERE tmpExpressions.utilityID = " & CStr(mlngWorkflowID) & _
      "   AND (tmpExpressions.type = " & CStr(giEXPR_WORKFLOWCALCULATION) & _
      "     OR tmpExpressions.type = " & CStr(giEXPR_WORKFLOWSTATICFILTER) & _
      "     OR tmpExpressions.type = " & CStr(giEXPR_WORKFLOWRUNTIMEFILTER) & ")" & _
      "   AND ucase(ltrim(rtrim(tmpComponents.workflowElement))) = '" & Replace(UCase(Trim(pavIdentifierLog(2, 0))), "'", "''") & "'"
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    Do While Not (rsTemp.BOF Or rsTemp.EOF)
      sComponentType = ComponentTypeName(rsTemp!Type)
            
      If fElementIdentifierChanged Or fElementTableChanged Then
        Set objComp = New CExprComponent
        objComp.ComponentID = rsTemp!ComponentID
        lngExprID = objComp.RootExpressionID
        Set objComp = Nothing
        
        fInvalidElement = (Len(pavIdentifierLog(3, 0)) = 0)
        If (Not fInvalidElement) Then
          sSQL = "UPDATE tmpExpressions" & _
            " SET tmpExpressions.changed = TRUE" & _
            " WHERE tmpExpressions.exprID = " & CStr(lngExprID)
          daoDb.Execute sSQL, dbFailOnError
        
          sSQL = "UPDATE tmpComponents" & _
            " SET tmpComponents.workflowElement = '" & Replace(pavIdentifierLog(3, 0), "'", "''") & "'" & _
            " WHERE tmpComponents.componentID = " & CStr(rsTemp!ComponentID)
          daoDb.Execute sSQL, dbFailOnError
        
          If (Not pwfBaseElement Is Nothing) Then
            If pwfBaseElement.ElementType = elem_StoredData Then
              ' Check if the StoredData table has changed.
              If (CLng(pavIdentifierLog(5, 0)) <> CLng(pavIdentifierLog(6, 0))) _
                And (rsTemp!WorkflowRecordTableID > 0) Then

                ReDim alngValidTables(0)
                TableAscendants CLng(pavIdentifierLog(6, 0)), alngValidTables

                fInvalidElement = True

                For iLoop4 = 1 To UBound(alngValidTables)
                  If (alngValidTables(iLoop4) = rsTemp!WorkflowRecordTableID) Then
                    fInvalidElement = False
                    Exit For
                  End If
                Next iLoop4
              End If
            End If
          End If
        End If

        If fInvalidElement Then
          ' Get the expression name and type description.
          Set objExpr = New CExpression
          objExpr.ExpressionID = lngExprID

          If objExpr.ReadExpressionDetails Then
            sExprName = objExpr.Name
            sExprType = objExpr.ExpressionTypeName

            ReDim Preserve asMessages(UBound(asMessages) + 1)
            asMessages(UBound(asMessages)) = _
              sExprType & " (" & sExprName & ") : " & _
              "Invalid " & sComponentType & " element selected"
          
            fExprsNeedReviewing = True
          End If
        End If
      End If
      
      For iLoop = 1 To UBound(pavIdentifierLog, 2)
        If ((UCase(Trim(pavIdentifierLog(2, iLoop))) <> UCase(Trim(pavIdentifierLog(3, iLoop)))) _
          Or (UCase(Trim(pavIdentifierLog(5, iLoop))) <> UCase(Trim(pavIdentifierLog(6, iLoop))))) _
          And (UCase(Trim(rsTemp!WorkflowItem)) = UCase(Trim(pavIdentifierLog(2, iLoop)))) Then
          
          Set objComp = New CExprComponent
          objComp.ComponentID = rsTemp!ComponentID
          lngExprID = objComp.RootExpressionID
          Set objComp = Nothing
          
          fInvalidItem = Len(pavIdentifierLog(3, iLoop)) = 0
          
          If (Not fInvalidItem) Then
            sSQL = "UPDATE tmpExpressions" & _
              " SET tmpExpressions.changed = TRUE" & _
              " WHERE tmpExpressions.exprID = " & CStr(lngExprID)
            daoDb.Execute sSQL, dbFailOnError
            
            sSQL = "UPDATE tmpComponents" & _
              " SET tmpComponents.WorkflowItem = '" & Replace(pavIdentifierLog(3, iLoop), "'", "''") & "'" & _
              " WHERE tmpComponents.componentID = " & CStr(rsTemp!ComponentID)
            daoDb.Execute sSQL, dbFailOnError
      
            ' Check if the recordSelector table has changed.
            If (CLng(pavIdentifierLog(5, iLoop)) <> CLng(pavIdentifierLog(6, iLoop))) _
              And (rsTemp!WorkflowRecordTableID > 0) Then

              ReDim alngValidTables(0)
              TableAscendants CLng(pavIdentifierLog(6, iLoop)), alngValidTables

              fInvalidItem = True

              For iLoop4 = 1 To UBound(alngValidTables)
                If (alngValidTables(iLoop4) = rsTemp!WorkflowRecordTableID) Then
                  fInvalidItem = False
                  Exit For
                End If
              Next iLoop4
            End If
          End If
          
          If fInvalidItem Then
            ' Identifier object deleted, not just renamed. Need to inform the user.
  
            ' Get the expression name and type description.
            Set objExpr = New CExpression
            objExpr.ExpressionID = lngExprID
  
            If objExpr.ReadExpressionDetails Then
              sExprName = objExpr.Name
              sExprType = objExpr.ExpressionTypeName
  
              ReDim Preserve asMessages(UBound(asMessages) + 1)
              asMessages(UBound(asMessages)) = _
                sExprType & " (" & sExprName & ") : " & _
                "Invalid " & sComponentType & IIf(rsTemp!Type = giCOMPONENT_WORKFLOWVALUE, " value", " record selector") & " selected"
            
              fExprsNeedReviewing = True
            End If
          End If
           
          Exit For
        End If
      Next iLoop

      rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
  End If
  
  If UBound(asMessages) > 0 Then
    If IsMissing(pasMessages) Then
      Set frmUsage = New frmUsage
      frmUsage.ResetList
  
      For iLoop = 1 To UBound(asMessages)
        frmUsage.AddToList asMessages(iLoop)
      Next iLoop
  
      Screen.MousePointer = vbDefault
  
      frmUsage.Width = (3 * Screen.Width / 4)
  
      sTemp = "Elements"
      If fExprsNeedReviewing Then
        If fElementsNeedReviewing Then
          sTemp = "Elements/Expressions"
        Else
          sTemp = "Expressions"
        End If
      End If

      frmUsage.ShowMessage "Workflow '" & Trim(msWorkflowName) & "'", "The following " & sTemp & " made reference to deleted web form items, and will need reviewing:", _
        UsageCheckObject.Workflow, _
        USAGEBUTTONS_PRINT + USAGEBUTTONS_OK, "validation"
  
      UnLoad frmUsage
      Set frmUsage = Nothing
    Else
      For iLoop = 1 To UBound(asMessages)
        ReDim Preserve pasMessages(UBound(pasMessages) + 1)
        pasMessages(UBound(pasMessages)) = asMessages(iLoop)
      Next iLoop
    End If
  End If
  
End Sub

Private Sub ValidateWorkflow_AddMessage(psMessage As String, plngIndex As Long)
  ' Add the given item to the array of validation messages.
  ReDim Preserve mavValidationMessages(1, UBound(mavValidationMessages, 2) + 1)
  mavValidationMessages(0, UBound(mavValidationMessages, 2)) = psMessage
  mavValidationMessages(1, UBound(mavValidationMessages, 2)) = plngIndex

End Sub

Private Sub abMenu_Click(ByVal pTool As ActiveBarLibraryCtl.Tool)
  
  If abMenu.ActiveBand.Name <> "ElementBand" Then
    mlngXDrop = -1
    mlngYDrop = -1
  End If
  
  EditMenu pTool.Name

End Sub


Public Sub EditMenu(ByVal psMenuOption As String)
  ' Process the menu options.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iCount As Integer
  Dim iElementIndex As Integer
  Dim wfTempElement1 As VB.Control
  Dim wfTempElement2 As VB.Control
  Dim sName As String
  Dim sDescription As String
  Dim lPictureID As Long
  Dim fEnabled As Boolean
  
  Select Case psMenuOption
    ' Cancel out of ElementAdd mode
    Case "ID_WFElement_Selector"
      CancelElementAddMode
    
    ' Add a new element to the flowchart.
    Case "ID_WFElement_Terminator"
      If (mlngXDrop >= 0) Or (mlngYDrop >= 0) Then
        DeselectAllElements
  
        Set wfTempElement1 = AddElement(elem_Terminator)
        
        wfTempElement1.Left = mlngXDrop - wfTempElement1.InboundFlow_XOffset
        wfTempElement1.Top = mlngYDrop - wfTempElement1.InboundFlow_YOffset
        
        ' AE20080627 Fault #13244
        wfTempElement1.Visible = True
        
        ' AE20080722 Fault #13283, #13284, #13285
          If Not mblnLoading Then
            SelectElement wfTempElement1
      
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
          End If
        
        'JPD 20070321 Fault 11936
        ResizeCanvas
      Else
        ToggleElementAddMode psMenuOption
      End If
  
    Case "ID_WFElement_WebForm"
      If (mlngXDrop >= 0) Or (mlngYDrop >= 0) Then
        DeselectAllElements
        
        Set wfTempElement1 = AddElement(elem_WebForm)
        
        wfTempElement1.Left = mlngXDrop - wfTempElement1.InboundFlow_XOffset
        wfTempElement1.Top = mlngYDrop - wfTempElement1.InboundFlow_YOffset
        
        ' AE20080627 Fault #13244
        wfTempElement1.Visible = True
        
        ' AE20080722 Fault #13283, #13284, #13285
          If Not mblnLoading Then
            SelectElement wfTempElement1
      
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
          End If
        
        'JPD 20070321 Fault 11936
        ResizeCanvas
      Else
        ToggleElementAddMode psMenuOption
      End If
  
    Case "ID_WFElement_Email"
      If (mlngXDrop >= 0) Or (mlngYDrop >= 0) Then
        DeselectAllElements
        
        Set wfTempElement1 = AddElement(elem_Email)
        
        wfTempElement1.Left = mlngXDrop - wfTempElement1.InboundFlow_XOffset
        wfTempElement1.Top = mlngYDrop - wfTempElement1.InboundFlow_YOffset
        
        ' AE20080627 Fault #13244
        wfTempElement1.Visible = True
        
        ' AE20080722 Fault #13283, #13284, #13285
          If Not mblnLoading Then
            SelectElement wfTempElement1
      
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
          End If
        
        'JPD 20070321 Fault 11936
        ResizeCanvas
      Else
        ToggleElementAddMode psMenuOption
      End If
  
    Case "ID_WFElement_Decision"
      If (mlngXDrop >= 0) Or (mlngYDrop >= 0) Then
        DeselectAllElements
        
        Set wfTempElement1 = AddElement(elem_Decision)
        
        wfTempElement1.Left = mlngXDrop - wfTempElement1.InboundFlow_XOffset
        wfTempElement1.Top = mlngYDrop - wfTempElement1.InboundFlow_YOffset
        
        ' AE20080627 Fault #13244
        wfTempElement1.Visible = True
        
        ' AE20080722 Fault #13283, #13284, #13285
          If Not mblnLoading Then
            SelectElement wfTempElement1
      
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
          End If
        
        'JPD 20070321 Fault 11936
        ResizeCanvas
      Else
        ToggleElementAddMode psMenuOption
      End If
      
    Case "ID_WFElement_StoredData"
      If (mlngXDrop >= 0) Or (mlngYDrop >= 0) Then
        DeselectAllElements
        
        Set wfTempElement1 = AddElement(elem_StoredData)
        
        wfTempElement1.Left = mlngXDrop - wfTempElement1.InboundFlow_XOffset
        wfTempElement1.Top = mlngYDrop - wfTempElement1.InboundFlow_YOffset
        
        ' AE20080627 Fault #13244
        wfTempElement1.Visible = True
        
        ' AE20080722 Fault #13283, #13284, #13285
          If Not mblnLoading Then
            SelectElement wfTempElement1
      
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
          End If
        
        'JPD 20070321 Fault 11936
        ResizeCanvas
      Else
        ToggleElementAddMode psMenuOption
      End If
  
    Case "ID_WFElement_SummingJunction"
      If (mlngXDrop >= 0) Or (mlngYDrop >= 0) Then
        DeselectAllElements
        
        Set wfTempElement1 = AddElement(elem_SummingJunction)
        
        wfTempElement1.Left = mlngXDrop - wfTempElement1.InboundFlow_XOffset
        wfTempElement1.Top = mlngYDrop - wfTempElement1.InboundFlow_YOffset
        
        ' AE20080627 Fault #13244
        wfTempElement1.Visible = True
            
        ' AE20080722 Fault #13283, #13284, #13285
          If Not mblnLoading Then
            SelectElement wfTempElement1
      
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
          End If
        
        'JPD 20070321 Fault 11936
        ResizeCanvas
      Else
        ToggleElementAddMode psMenuOption
      End If
  
    Case "ID_WFElement_Or"
      If (mlngXDrop >= 0) Or (mlngYDrop >= 0) Then
        DeselectAllElements
        
        Set wfTempElement1 = AddElement(elem_Or)
        
        wfTempElement1.Left = mlngXDrop - wfTempElement1.InboundFlow_XOffset
        wfTempElement1.Top = mlngYDrop - wfTempElement1.InboundFlow_YOffset
        
        ' AE20080627 Fault #13244
        wfTempElement1.Visible = True
        
        ' AE20080722 Fault #13283, #13284, #13285
          If Not mblnLoading Then
            SelectElement wfTempElement1
      
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
          End If
        
        'JPD 20070321 Fault 11936
        ResizeCanvas
      Else
        ToggleElementAddMode psMenuOption
      End If

    Case "ID_WFElement_Connector"
      If (mlngXDrop >= 0) Or (mlngYDrop >= 0) Then
        DeselectAllElements
        
        Set wfTempElement1 = AddElement(elem_Connector1)
        Set wfTempElement2 = AddElement(elem_Connector2)
        
        'JPD 20070713 Fault 12255
        ' Add the first connector to the Undo array as it will have been removed when the
        ' second connector was added
        ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
        Set mactlUndoControls(UBound(mactlUndoControls)) = wfTempElement1
        
        wfTempElement1.Left = mlngXDrop - wfTempElement1.InboundFlow_XOffset
        wfTempElement1.Top = mlngYDrop - wfTempElement1.InboundFlow_YOffset
        wfTempElement1.Visible = True
      
        With wfTempElement1
          wfTempElement2.Left = wfTempElement1.Left + wfTempElement1.Width + 200
          wfTempElement2.Top = wfTempElement1.Top
          
          .ConnectorPairIndex = wfTempElement2.ControlIndex
          wfTempElement2.ConnectorPairIndex = .ControlIndex
          
          .Caption = NextConnectorCaption
          wfTempElement2.Caption = .Caption
          
          ' AE20080627 Fault #13244
          wfTempElement2.Visible = True
        
'          .HighLighted = True
          SelectElement wfTempElement1
        End With
        'JPD 20070321 Fault 11936
        ResizeCanvas
      Else
        ToggleElementAddMode psMenuOption
      End If
  
    Case "ID_WFElement_Link"
      ToggleElementAddMode psMenuOption
      
    Case "ID_WorkflowCut"
      If Not CutSelectedElements Then
        MsgBox "Unable to cut workflow elements." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    Case "ID_WorkflowCopy"
      If Not CopySelectedElements Then
        MsgBox "Unable to copy workflow elements." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    Case "ID_WorkflowPaste"
      If Not PasteElements Then
        MsgBox "Unable to paste workflow elements." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    Case "ID_WorkflowDelete"
      DeleteElementsAndLinks
    
    Case "ID_WorkflowClear"
      'JPD 20070615 Fault 12331
      If MsgBox("This will remove all elements except the 'Begin' element." & vbCrLf & vbCrLf & "Do you wish to continue?", _
        vbQuestion + vbYesNo, Me.Caption) = vbYes Then
      
        ClearFlowchart False
      End If
    
    Case "ID_WorkflowUndo"
      If Not UndoLastAction Then
        MsgBox "Unable to undo the last action." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    Case "ID_WorkflowProperties"
      CancelElementAddMode
      
      Dim perge As Boolean
      
      frmWorkflowEdit.WorkflowID = mlngWorkflowID
      Set frmWorkflowEdit.CallingForm = Me
      frmWorkflowEdit.WorkflowName = msWorkflowName
      frmWorkflowEdit.InitiationType = miInitiationType
      frmWorkflowEdit.ExternalInitiationQueryString = msExternalInitiationQueryString
      frmWorkflowEdit.WorkflowDescription = msWorkflowDescription
      frmWorkflowEdit.WorkflowPictureID = mlngWorkflowPictureID
      frmWorkflowEdit.WorkflowEnabled = mfWorkflowEnabled
      frmWorkflowEdit.MustSaveChanges = False
      frmWorkflowEdit.Show vbModal
      fOK = Not frmWorkflowEdit.Cancelled
      If fOK Then
        sName = frmWorkflowEdit.WorkflowName
        sDescription = frmWorkflowEdit.WorkflowDescription
        lPictureID = frmWorkflowEdit.WorkflowPictureID
        fEnabled = frmWorkflowEdit.WorkflowEnabled
      
        'JPD 20060919 Fault 11283
        fOK = (Trim(msWorkflowName) <> Trim(sName)) _
          Or (Trim(msWorkflowDescription) <> Trim(sDescription)) _
          Or (mlngWorkflowPictureID <> lPictureID) _
          Or (mfWorkflowEnabled <> fEnabled)
          
        perge = (Trim(msWorkflowName) <> Trim(sName)) _
          Or (Trim(msWorkflowDescription) <> Trim(sDescription)) _
          Or (mfWorkflowEnabled <> fEnabled)
      End If
       
      Set frmWorkflowEdit = Nothing

      If fOK Then
        ' The WORKFLOW name may have been changed so update the form caption, and
        ' also the frmWorkflowOpen screen list if it is loaded.
        msWorkflowName = sName
        msWorkflowDescription = sDescription
        mlngWorkflowPictureID = lPictureID
        mfWorkflowEnabled = fEnabled
        
        Me.Caption = "Workflow Designer - " & sName
        
        SetChanged perge
      End If
      
    Case "ID_WorkflowElementProperties"
      iCount = 0
      For Each wfTempElement1 In mcolwfElements
        If (wfTempElement1.Highlighted) And (wfTempElement1.Visible) Then
          iCount = iCount + 1
          iElementIndex = wfTempElement1.ControlIndex
        End If
      Next wfTempElement1
      Set wfTempElement1 = Nothing
      
      If iCount = 1 Then
        ' Edit the single selected element.
        ElementEdit mcolwfElements(CStr(iElementIndex))
      ElseIf iCount > 1 Then
        ' More than one element selected. Don't know which one to edit, so tell the user.
        MsgBox "More than one element selected.", _
          vbExclamation + vbOKOnly, App.ProductName
      Else
        ' No elements selected. So tell the user.
        MsgBox "No elements selected.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If

    Case "ID_WorkflowAutoLayout"
      AutoFormat
  
    Case "ID_WorkflowResizeCanvas"
      ManualResizeCanvas
    
    Case "ID_WorkflowUsageCheck"
      FindUsage
    
    Case "ID_SelectAll"
      SelectAllElements
      
  End Select
  
  RefreshMenu
  
  Exit Sub
  
ErrorTrap:

End Sub

Private Function UndoLastAction() As Boolean
  ' Undo the last action.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  Screen.MousePointer = vbHourglass
  UI.LockWindow Me.hWnd
  
  Select Case miLastActionFlag
    ' Undo the previous element/link creation.
    Case giACTION_DROPCONTROL
      UndoDropControl

    ' Undo the previous control Delete.
    Case giACTION_DELETECONTROLS
      UndoDeleteControls
      
    ' Undo the previous control Delete.
    Case giACTION_SWAPCONTROL
      UndoSwapControl
  End Select

  ' Clear the last action flag.
  SetLastActionFlag giACTION_NOACTION

  
  fOK = True
  
TidyUpAndExit:
  Screen.MousePointer = vbDefault
  UI.UnlockWindow
  
  UndoLastAction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function UndoDeleteControls() As Boolean
  ' Paste the deleted controls onto their original page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim ctlNewControl As VB.Control
 
  ' Restore the deleted controls to their original positions.
  For iIndex = 1 To UBound(mactlUndoControls)
    Set ctlNewControl = mactlUndoControls(iIndex)
    
    If TypeOf ctlNewControl Is COAWF_Link Then
      FormatLink ctlNewControl
    End If
    
    ctlNewControl.Visible = True
    
    Set mactlUndoControls(iIndex) = Nothing
  Next iIndex
  
  ResizeCanvas

TidyUpAndExit:
  UndoDeleteControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function UndoDropControl() As Boolean
  ' Delete the last element/links that were created.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim wfElement As VB.Control
  Dim wfLink As COAWF_Link
  
  fOK = True
  
  For iLoop = 1 To UBound(mactlUndoControls)
    If IsWorkflowElement(mactlUndoControls(iLoop)) Then
      ' Delete the element.
      Set wfElement = mactlUndoControls(iLoop)
      DeleteElement wfElement, False
      Set wfElement = Nothing
    End If

    If TypeOf mactlUndoControls(iLoop) Is COAWF_Link Then
      ' Delete the link.
      Set wfLink = mactlUndoControls(iLoop)
      wfLink.Visible = False
      Set wfLink = Nothing
    End If
  Next iLoop
  
  DeselectAllElements
  
  IsChanged = True

TidyUpAndExit:
  UndoDropControl = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function UndoSwapControl() As Boolean
  ' Swapped a link (ie. deleted an old one, and created a new one).
  ' To undo, simply deelte the new one, and restore the old one.
  ' We assume that the first item in the Undo array is the old link,
  ' and the second item in the new link.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim wfOldLink As COAWF_Link
  Dim wfNewLink As COAWF_Link

  fOK = True

  If UBound(mactlUndoControls) = 2 Then
    If (TypeOf mactlUndoControls(1) Is COAWF_Link) _
      And (TypeOf mactlUndoControls(2) Is COAWF_Link) Then

      Set wfOldLink = mactlUndoControls(1)
      Set wfNewLink = mactlUndoControls(2)

      wfOldLink.Visible = True
      Set mactlUndoControls(1) = Nothing

      wfNewLink.Visible = False
      Set wfNewLink = Nothing
    End If
  End If

  DeselectAllElements
  
  IsChanged = True

TidyUpAndExit:
  UndoSwapControl = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function PasteElements() As Boolean
  ' Paste the elements and links from the clipboard.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iIndex2 As Integer
  Dim iIndex3 As Integer
  Dim lngXOffset As Long
  Dim lngYOffset As Long
  Dim wfElement As VB.Control
  Dim wfNewElement As VB.Control
  Dim wfLink As COAWF_Link
  Dim wfNewLink As COAWF_Link
  Dim aiCopiedElements() As Integer
  Dim iOutboundFlowIndex As Integer
  Dim avOutboundFlowInfo() As Variant
  Dim asFormIdentifiers() As String
  Dim sBaseIdentifier As String
  Dim fFound As Boolean
  Dim sIdentifier As String
  Dim wfElement2 As VB.Control
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim wfTemp As VB.Control
  Dim asItems() As String
  Dim avColumns() As Variant
  Dim asValidations() As String

  fOK = True
    
  ' Do nothing if there's nothing in the clipboard.
  If UBound(mactlClipboardControls) = 0 Then
    PasteElements = True
    Exit Function
  End If

  Screen.MousePointer = vbHourglass
  UI.LockWindow Me.hWnd
  
  ' Deselect all existing controls.
  DeselectAllElements

  ReDim aiCopiedElements(1, 0)
  ' Column 0 = source element index
  ' Column 1 = copied element index

  ReDim asFormIdentifiers(1, 0)
  ' Column 0 = original wf form identifier
  ' Column 1 = copied wf form identifier

  SetLastActionFlag giACTION_DROPCONTROL

  lngXOffset = picDefinition.Width
  lngYOffset = picDefinition.Height
  
  For iIndex = 1 To UBound(mactlClipboardControls)
    If IsWorkflowElement(mactlClipboardControls(iIndex)) Then
      Set wfElement = mactlClipboardControls(iIndex)
      
      With wfElement
        If .Left < lngXOffset Then
          lngXOffset = .Left
        End If
        If .Top < lngYOffset Then
          lngYOffset = .Top
        End If
      End With
    
      Set wfElement = Nothing
    End If
  Next iIndex
  
  'JPD 20060911 Fault 11486
  'lngXOffset = lngXOffset - 1000
  'lngYOffset = lngYOffset - 1000
  lngXOffset = lngXOffset - 1000 + picDefinition.Left
  lngYOffset = lngYOffset - 1000 + picDefinition.Top
  
  ' Drop each element/link from the clipboard.
  For iIndex = 1 To UBound(mactlClipboardControls)
    If IsWorkflowElement(mactlClipboardControls(iIndex)) Then
      ' Paste an element.
      Set wfElement = mactlClipboardControls(iIndex)
      Set wfNewElement = LoadNewElementOfType(wfElement.ElementType)
      'Set wfNewElement = AddElement(wfElement.ElementType)
    
      fOK = Not (wfNewElement Is Nothing)
      If fOK Then
        
        CopyElementProperties wfElement, wfNewElement

        If Len(wfNewElement.Identifier) > 0 Then
          fFound = True
          sBaseIdentifier = "Copy of " & wfNewElement.Identifier
          sIdentifier = sBaseIdentifier
          iIndex3 = 2
        
          Do While fFound
            fFound = False
        
            For Each wfElement2 In mcolwfElements
              If wfElement2.Identifier = sIdentifier Then
                fFound = True
                sIdentifier = sBaseIdentifier & " (" & CStr(iIndex3) & ")"
                iIndex3 = iIndex3 + 1
                Exit For
              End If
            Next wfElement2
            Set wfElement2 = Nothing
          Loop
        
          wfNewElement.Identifier = sIdentifier
        
          ReDim Preserve asFormIdentifiers(1, UBound(asFormIdentifiers, 2) + 1)
          asFormIdentifiers(0, UBound(asFormIdentifiers, 2)) = wfElement.Identifier
          asFormIdentifiers(1, UBound(asFormIdentifiers, 2)) = sIdentifier
        End If

        With wfNewElement
          If fOK Then
            .Top = .Top - lngYOffset
            .Left = .Left - lngXOffset
            
            .Visible = True
            SelectElement wfNewElement
            .ZOrder 0

            ReDim Preserve aiCopiedElements(1, UBound(aiCopiedElements, 2) + 1)
            aiCopiedElements(0, UBound(aiCopiedElements, 2)) = wfElement.ControlIndex
            aiCopiedElements(1, UBound(aiCopiedElements, 2)) = wfNewElement.ControlIndex

            ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
            Set mactlUndoControls(UBound(mactlUndoControls)) = wfNewElement
            
          End If
        End With
      End If
    Else
      ' Paste a link.
      Set wfLink = mactlClipboardControls(iIndex)
      
      Load ASRWFLink1(ASRWFLink1.UBound + 1)
      Set wfNewLink = ASRWFLink1(ASRWFLink1.UBound)

      fOK = Not (wfNewLink Is Nothing)
      If fOK Then
        CopyLinkProperties wfLink, wfNewLink
    
        With wfNewLink
          For iIndex2 = 1 To UBound(aiCopiedElements, 2)
            If aiCopiedElements(0, iIndex2) = .StartElementIndex Then
              .StartElementIndex = aiCopiedElements(1, iIndex2)
            End If
            If aiCopiedElements(0, iIndex2) = .EndElementIndex Then
              .EndElementIndex = aiCopiedElements(1, iIndex2)
            End If
          Next iIndex2
          
          avOutboundFlowInfo = mcolwfElements(CStr(.StartElementIndex)).OutboundFlows_Information
          If .StartOutboundFlowCode < 0 Then
            iOutboundFlowIndex = 1
          Else
            For iIndex2 = 1 To UBound(avOutboundFlowInfo, 2)
              If avOutboundFlowInfo(1, iIndex2) = .StartOutboundFlowCode Then
                iOutboundFlowIndex = iIndex2
                Exit For
              End If
            Next iIndex2
          End If

          .StartDirection = avOutboundFlowInfo(2, iOutboundFlowIndex)
          .EndDirection = mcolwfElements(CStr(.EndElementIndex)).InboundFlow_Direction
          
          FormatLink wfNewLink
          
          .Visible = True
          .ZOrder 1
                
          ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
          Set mactlUndoControls(UBound(mactlUndoControls)) = wfNewLink
        End With
      End If
    End If

    If Not fOK Then
      Exit For
    End If
  Next iIndex

  ' Update the web form identifiers that have been copied.
  For iLoop = 1 To UBound(mactlUndoControls)
    If IsWorkflowElement(mactlUndoControls(iLoop)) Then
      Set wfTemp = mactlUndoControls(iLoop)
  
      Select Case wfTemp.ElementType
        Case elem_Email
          If wfTemp.EmailRecord = giWFRECSEL_IDENTIFIEDRECORD Then
            For iLoop3 = 1 To UBound(asFormIdentifiers, 2)
              If UCase(Trim(wfTemp.RecordSelectorWebFormIdentifier)) = UCase(Trim(asFormIdentifiers(0, iLoop3))) Then
                wfTemp.RecordSelectorWebFormIdentifier = asFormIdentifiers(1, iLoop3)
                Exit For
              End If
            Next iLoop3
          End If
  
          asItems = wfTemp.Items
    
          For iLoop2 = 1 To UBound(asItems, 2)
            If CInt(asItems(2, iLoop2)) = giWFFORMITEM_DBVALUE Then
              If CInt(asItems(5, iLoop2)) = giWFRECSEL_IDENTIFIEDRECORD Then
                For iLoop3 = 1 To UBound(asFormIdentifiers, 2)
                  If UCase(Trim(asItems(13, iLoop2))) = UCase(Trim(asFormIdentifiers(0, iLoop3))) Then
                    asItems(13, iLoop2) = asFormIdentifiers(1, iLoop3)
                    Exit For
                  End If
                Next iLoop3
              End If
            ElseIf (CInt(asItems(2, iLoop2)) = giWFFORMITEM_WFVALUE) _
              Or (CInt(asItems(2, iLoop2)) = giWFFORMITEM_WFFILE) Then
              
              For iLoop3 = 1 To UBound(asFormIdentifiers, 2)
                If UCase(Trim(asItems(11, iLoop2))) = UCase(Trim(asFormIdentifiers(0, iLoop3))) Then
                  asItems(11, iLoop2) = asFormIdentifiers(1, iLoop3)
                  Exit For
                End If
              Next iLoop3
            End If
          Next iLoop2
    
          wfTemp.Items = asItems
        
        Case elem_WebForm
          asItems = wfTemp.Items
    
          For iLoop2 = 1 To UBound(asItems, 2)
            If (((CInt(asItems(2, iLoop2)) = giWFFORMITEM_DBVALUE) _
              Or (CInt(asItems(2, iLoop2)) = giWFFORMITEM_INPUTVALUE_GRID)) _
              And (CInt(asItems(5, iLoop2)) = giWFRECSEL_IDENTIFIEDRECORD)) _
              Or (CInt(asItems(2, iLoop2)) = giWFFORMITEM_WFVALUE) _
              Or (CInt(asItems(2, iLoop2)) = giWFFORMITEM_WFFILE) Then
              
              For iLoop3 = 1 To UBound(asFormIdentifiers, 2)
                If UCase(Trim(asItems(11, iLoop2))) = UCase(Trim(asFormIdentifiers(0, iLoop3))) Then
                  asItems(11, iLoop2) = asFormIdentifiers(1, iLoop3)
                  Exit For
                End If
              Next iLoop3
            End If
          Next iLoop2
    
          wfTemp.Items = asItems
          
          asValidations = wfTemp.Validations
          wfTemp.Validations = asValidations
          
        Case elem_StoredData
          If wfTemp.DataRecord = giWFRECSEL_IDENTIFIEDRECORD Then
            For iLoop3 = 1 To UBound(asFormIdentifiers, 2)
              If UCase(Trim(wfTemp.RecordSelectorWebFormIdentifier)) = UCase(Trim(asFormIdentifiers(0, iLoop3))) Then
                wfTemp.RecordSelectorWebFormIdentifier = asFormIdentifiers(1, iLoop3)
                Exit For
              End If
            Next iLoop3
          End If
          
          If wfTemp.SecondaryDataRecord = giWFRECSEL_IDENTIFIEDRECORD Then
            For iLoop3 = 1 To UBound(asFormIdentifiers, 2)
              If UCase(Trim(wfTemp.SecondaryRecordSelectorWebFormIdentifier)) = UCase(Trim(asFormIdentifiers(0, iLoop3))) Then
                wfTemp.SecondaryRecordSelectorWebFormIdentifier = asFormIdentifiers(1, iLoop3)
                Exit For
              End If
            Next iLoop3
          End If
          
          avColumns = wfTemp.DataColumns
    
          For iLoop2 = 1 To UBound(avColumns, 2)
            If CInt(avColumns(4, iLoop2)) = giWFDATAVALUE_WFVALUE _
              Or CInt(avColumns(4, iLoop2)) = giWFDATAVALUE_DBVALUE Then
              
              For iLoop3 = 1 To UBound(asFormIdentifiers, 2)
                If UCase(Trim(avColumns(6, iLoop2))) = UCase(Trim(asFormIdentifiers(0, iLoop3))) Then
                  avColumns(6, iLoop2) = asFormIdentifiers(1, iLoop3)
                  Exit For
                End If
              Next iLoop3
            End If
          Next iLoop2
    
          wfTemp.DataColumns = avColumns
  
      End Select
  
      Set wfTemp = Nothing
    End If
  Next iLoop

  If fOK Then
    ' Update the connector pair indexes.
    For iIndex = 1 To UBound(aiCopiedElements, 2)
      If (mcolwfElements(CStr(aiCopiedElements(1, iIndex))).ElementType = elem_Connector1) _
        Or (mcolwfElements(CStr(aiCopiedElements(1, iIndex))).ElementType = elem_Connector2) Then
        
        For iIndex2 = 1 To UBound(aiCopiedElements, 2)
          If mcolwfElements(CStr(aiCopiedElements(1, iIndex))).ConnectorPairIndex = aiCopiedElements(0, iIndex2) Then
            mcolwfElements(CStr(aiCopiedElements(1, iIndex))).ConnectorPairIndex = aiCopiedElements(1, iIndex2)
            Exit For
          End If
        Next iIndex2
      End If
    Next iIndex
    
    For iIndex = 1 To UBound(aiCopiedElements, 2)
      If (mcolwfElements(CStr(aiCopiedElements(1, iIndex))).ElementType = elem_Connector1) Then
        mcolwfElements(CStr(aiCopiedElements(1, iIndex))).Caption = NextConnectorCaption
        mcolwfElements(CStr(mcolwfElements(CStr(aiCopiedElements(1, iIndex))).ConnectorPairIndex)).Caption = _
          mcolwfElements(CStr(aiCopiedElements(1, iIndex))).Caption
      End If
    Next iIndex
    
    IsChanged = True

    ReDim miSelectionOrder(0)
  End If

TidyUpAndExit:

  UI.UnlockWindow
  Screen.MousePointer = vbDefault
  
  PasteElements = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function CutSelectedElements() As Boolean
  ' Cut the selected controls.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Copy the selected Elements to the clipboard.
  fOK = CopySelectedElements
  
  If fOK Then
    ' Delete the selected Elements.
    DeleteElementsAndLinks
  End If
  
  If fOK Then
    IsChanged = True
  End If

TidyUpAndExit:
  CutSelectedElements = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function CopySelectedElements() As Boolean
  ' Copy the selected Elements to the clipboard array.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim ctlSourceControl As VB.Control
  Dim ctlCopiedControl As VB.Control
  Dim wfElement As VB.Control
  Dim wfElement2 As VB.Control
  Dim wfTemp As VB.Control
  Dim wfTemp2 As VB.Control
  Dim wfLink As COAWF_Link
  Dim wfNewLink As COAWF_Link
  Dim aiCopiedElements() As Integer
  Dim sCopiedElementIndexes As String

  sCopiedElementIndexes = ","
  ReDim aiCopiedElements(1, 0)
  ' Column 0 = source element index
  ' Column 1 = copied element index
  
  ' Do nothing if no elements are selected.
  If SelectedElementCount = 0 Then
    CopySelectedElements = True
    Exit Function
  End If

  Screen.MousePointer = vbHourglass
  UI.LockWindow Me.hWnd
  
  ' Clear the array of copied elements.
  For iIndex = 1 To UBound(mactlClipboardControls)
    Set ctlCopiedControl = mactlClipboardControls(iIndex)
    
    If IsWorkflowElement(ctlCopiedControl) Then
      mcolwfElements.Remove CStr(ctlCopiedControl.ControlIndex)
    End If
    
    UnLoad ctlCopiedControl
  Next iIndex
  
  ReDim mactlClipboardControls(0)
  
  ' Disassociate object variables.
  Set ctlCopiedControl = Nothing

  ' Create a copy of each selected element (and the links) in the array.
  For Each wfElement In mcolwfElements

    If (wfElement.Highlighted) _
      And (wfElement.Visible) _
      And (wfElement.ElementType <> elem_Begin) Then
      
      ' Create a new instance of the element.
      'Load asrwfElements(mcolwfElements.UBound + 1)
      'Set wfTemp = mcolwfElements(CStr(mcolwfElements.UBound)
      'Set wfTemp = AddElement(wfElement.ElementType)
      Set wfTemp = LoadNewElementOfType(wfElement.ElementType)

      fOK = Not (wfTemp Is Nothing)

      If fOK Then

        ' Copy the properties from the selected element to the new element.
        CopyElementProperties wfElement, wfTemp
        
        With wfTemp
          ' Copy connector pair elements (if not already selected to be copied)
          If ((wfTemp.ElementType = elem_Connector1) Or (wfTemp.ElementType = elem_Connector2)) Then
            If (Not mcolwfElements(CStr(wfTemp.ConnectorPairIndex)).Highlighted) Then
          
              Set wfElement2 = mcolwfElements(CStr(wfTemp.ConnectorPairIndex))
              
              ' Create a new instance of the element.
              Set wfTemp2 = LoadNewElementOfType(wfElement2.ElementType)
              'Set wfTemp2 = AddElement(wfElement2.ElementType)
        
              fOK = Not (wfTemp2 Is Nothing)
        
              If fOK Then
                ' Copy the properties from the selected element to the new element.
                CopyElementProperties wfElement2, wfTemp2
              
                iIndex = UBound(mactlClipboardControls) + 1
                ReDim Preserve mactlClipboardControls(iIndex)
                Set mactlClipboardControls(iIndex) = wfTemp2
              
                sCopiedElementIndexes = sCopiedElementIndexes & CStr(wfElement2.ControlIndex) & ","
                ReDim Preserve aiCopiedElements(1, UBound(aiCopiedElements, 2) + 1)
                aiCopiedElements(0, UBound(aiCopiedElements, 2)) = wfElement2.ControlIndex
                aiCopiedElements(1, UBound(aiCopiedElements, 2)) = wfTemp2.ControlIndex
                
                'miControlIndex = miControlIndex + 1
                'mcolwfElements.Add wfTemp2, CStr(miControlIndex)
              End If
            
              Set wfTemp2 = Nothing
              Set wfElement2 = Nothing

            End If
          End If
        End With

        iIndex = UBound(mactlClipboardControls) + 1
        ReDim Preserve mactlClipboardControls(iIndex)
        Set mactlClipboardControls(iIndex) = wfTemp
      
        sCopiedElementIndexes = sCopiedElementIndexes & CStr(wfElement.ControlIndex) & ","
        
        ReDim Preserve aiCopiedElements(1, UBound(aiCopiedElements, 2) + 1)
        aiCopiedElements(0, UBound(aiCopiedElements, 2)) = wfElement.ControlIndex
        aiCopiedElements(1, UBound(aiCopiedElements, 2)) = wfTemp.ControlIndex
        
      Else
        Exit For
      End If

      Set wfTemp = Nothing
    End If
  Next wfElement
  Set wfElement = Nothing

  ' Update the web form identifiers that have been copied.
  For iLoop = 1 To UBound(mactlClipboardControls)
    Set wfTemp = mactlClipboardControls(iLoop)
      If wfTemp.ElementType = elem_Connector1 _
        Or wfTemp.ElementType = elem_Connector2 Then

      For iIndex = 1 To UBound(aiCopiedElements, 2)
        If aiCopiedElements(0, iIndex) = wfTemp.ConnectorPairIndex Then
          wfTemp.ConnectorPairIndex = aiCopiedElements(1, iIndex)
        End If
      Next iIndex
    End If

    Set wfTemp = Nothing
  Next iLoop
  
  ' Copy the links between the copied elements
  For Each wfLink In ASRWFLink1
    If (InStr(sCopiedElementIndexes, "," & CStr(wfLink.StartElementIndex) & ",") > 0) _
      And (InStr(sCopiedElementIndexes, "," & CStr(wfLink.EndElementIndex) & ",") > 0) Then
      ' The link IS between two copied elements.
      
      ' Create a new instance of the link.
      Load ASRWFLink1(ASRWFLink1.UBound + 1)
      Set wfNewLink = ASRWFLink1(ASRWFLink1.UBound)

      fOK = Not (wfNewLink Is Nothing)

      If fOK Then

        CopyLinkProperties wfLink, wfNewLink
        
        For iIndex = 1 To UBound(aiCopiedElements, 2)
          If aiCopiedElements(0, iIndex) = wfNewLink.StartElementIndex Then
            wfNewLink.StartElementIndex = aiCopiedElements(1, iIndex)
          End If
          
          If aiCopiedElements(0, iIndex) = wfNewLink.EndElementIndex Then
            wfNewLink.EndElementIndex = aiCopiedElements(1, iIndex)
          End If
        Next iIndex
        
        iIndex = UBound(mactlClipboardControls) + 1
        ReDim Preserve mactlClipboardControls(iIndex)
        Set mactlClipboardControls(iIndex) = wfNewLink
      Else
        Exit For
      End If
      
    End If
  Next wfLink
  Set wfLink = Nothing

TidyUpAndExit:
  If Not fOK Then
    ' Clear the array of copied elements.
    For iIndex = 1 To UBound(mactlClipboardControls)
      Set ctlCopiedControl = mactlClipboardControls(iIndex)
      Set mcolwfElements(CStr(ctlCopiedControl.ControlIndex)) = Nothing
      
      If IsWorkflowElement(ctlCopiedControl) Then
        mcolwfElements.Remove CStr(ctlCopiedControl.ControlIndex)
      End If
      UnLoad ctlCopiedControl
    Next iIndex
    ReDim mactlClipboardControls(0)
  End If

  UI.UnlockWindow
  Screen.MousePointer = vbDefault
  
  ' Disassociate object variables.
  Set ctlSourceControl = Nothing
  Set ctlCopiedControl = Nothing
  
  CopySelectedElements = fOK
  frmSysMgr.RefreshMenu
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub abMenu_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  ' Do not let the user modify the layout.
  Cancel = True
End Sub

Private Sub ASRWFBeginEnd1_DblClick(Index As Integer)
  ElementEdit ASRWFBeginEnd1(Index)
End Sub

Private Sub ASRWFBeginEnd1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRWFBeginEnd1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseDown ASRWFBeginEnd1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFBeginEnd1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseMove ASRWFBeginEnd1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFBeginEnd1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseUp ASRWFBeginEnd1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFEmail1_DblClick(Index As Integer)
  ElementEdit ASRWFEmail1(Index)
End Sub

Private Sub ASRWFEmail1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRWFEmail1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseDown ASRWFEmail1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFEmail1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseMove ASRWFEmail1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFEmail1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseUp ASRWFEmail1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFJunctionElement1_DblClick(Index As Integer)
  ElementEdit ASRWFJunctionElement1(Index)
End Sub

Private Sub ASRWFJunctionElement1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRWFJunctionElement1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseDown ASRWFJunctionElement1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFJunctionElement1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseMove ASRWFJunctionElement1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFJunctionElement1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseUp ASRWFJunctionElement1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFStoredData1_DblClick(Index As Integer)
  ElementEdit ASRWFStoredData1(Index)
End Sub

Private Sub ASRWFStoredData1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRWFStoredData1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseDown ASRWFStoredData1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFStoredData1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseMove ASRWFStoredData1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFStoredData1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseUp ASRWFStoredData1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFWebform1_DblClick(Index As Integer)
  ElementEdit ASRWFWebform1(Index)
End Sub

Private Sub ASRWFWebform1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRWFWebform1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseDown ASRWFWebform1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFWebform1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseMove ASRWFWebform1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFWebform1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseUp ASRWFWebform1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFDecision1_DblClick(Index As Integer)
  ElementEdit ASRWFDecision1(Index)
End Sub

Private Sub ASRWFDecision1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRWFDecision1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseDown ASRWFDecision1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFDecision1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseMove ASRWFDecision1(Index).ControlIndex, Button, Shift, x, y
End Sub

Private Sub ASRWFDecision1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  WorkflowElement_MouseUp ASRWFDecision1(Index).ControlIndex, Button, Shift, x, y
End Sub

'Private Sub ASRWFElement1_DblClick(Index As Integer)
'  ElementEdit ASRWFElement1(Index)
'End Sub
'
'Private Sub ASRWFElement1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'  Form_KeyDown KeyCode, Shift
'End Sub
'
'Private Sub ASRWFElement1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'  WorkflowElement_MouseDown ASRWFElement1(Index).ControlIndex, Button, Shift, X, Y
'End Sub
'
'Private Sub ASRWFElement1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'  ' Move the control.
'  WorkflowElement_MouseMove ASRWFElement1(Index).ControlIndex, Button, Shift, X, Y
'End Sub
'
'Private Sub ASRWFElement1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'  WorkflowElement_MouseUp ASRWFElement1(Index).ControlIndex, Button, Shift, X, Y
'End Sub

Private Sub ASRWFLink1_DblClick(Index As Integer)
  Dim wfStartElement As VB.Control
  Dim wfEndElement As VB.Control
  Dim avOutboundFlowInfo() As Variant
  Dim wfTempLink As COAWF_Link
  
  On Error GoTo ErrorTrap
  
  Set wfStartElement = mcolwfElements(CStr(ASRWFLink1(Index).StartElementIndex))
  Set wfEndElement = mcolwfElements(CStr(ASRWFLink1(Index).EndElementIndex))

  If (Not mfReadOnly) _
    And (Not wfStartElement Is Nothing) _
    And (Not wfEndElement Is Nothing) Then
    ' Get the array of outbound flow information from the start element.
    ' Column 1 = Tag (see enums, or -1 if there's only a single outbound flow)
    ' Column 2 = Direction
    ' Column 3 = XOffset
    ' Column 4 = YOffset
    ' Column 5 = Maximum
    ' Column 6 = Minimum
    ' Column 7 = Description
    avOutboundFlowInfo = wfStartElement.OutboundFlows_Information
    
    If UBound(avOutboundFlowInfo, 2) > 1 Then
      SetLastActionFlag giACTION_SWAPCONTROL
      
      ' Get rid of the original link.
      ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
      Set mactlUndoControls(UBound(mactlUndoControls)) = ASRWFLink1(Index)
      ASRWFLink1(Index).Visible = False
      
      ' Create the new link.
      Set wfTempLink = CreateLink(wfStartElement.ControlIndex, wfEndElement.ControlIndex, ASRWFLink1(Index).StartOutboundFlowCode)
      
      If wfTempLink Is Nothing Then
        ' Change cancelled, restore the original link.
        ASRWFLink1(Index).Visible = True
        Set mactlUndoControls(UBound(mactlUndoControls)) = Nothing
    
        SetLastActionFlag giACTION_NOACTION
      Else
        ' Remember what we've just done, so that we can undo it.
        ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
        Set mactlUndoControls(UBound(mactlUndoControls)) = wfTempLink
        Set wfTempLink = Nothing
      
        IsChanged = True
      End If
    End If
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ASRWFLink1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift

End Sub

Private Sub ASRWFLink1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

  If Button = vbLeftButton Then
    If (Shift <> vbShiftMask) And (Shift <> vbCtrlMask) And (Not ASRWFLink1(Index).Highlighted) Then
      DeselectAllElements
    End If
  
    'ASRWFLink1(Index).HighLighted = True
    
    SelectLink ASRWFLink1(Index)
    ASRWFLink1(Index).ZOrder 0
    
    RefreshMenu
  End If

End Sub

Private Sub cmdCancel_Click()
  UnLoad Me

End Sub

Private Sub cmdOK_Click()
 If SaveWorkflow Then
   UnLoad Me
 End If

End Sub

Private Sub cmdValidate_Click()
  ValidateWorkflow False, False, False
End Sub

Private Sub Form_Activate()
  ' Ensure the screen designer form is at the front of the display.
  On Error GoTo ErrorTrap

  Dim iForm As Integer
   
  For iForm = 0 To Forms.Count - 1 Step 1
    If Forms(iForm).Name = "frmWorkflowWFItemProps" _
      Or Forms(iForm).Name = "frmWorkflowWFToolbox" Then
      UnLoad Forms(iForm)
    End If
  Next iForm
  
  Me.ZOrder 0
  
  ' Refresh the menu/toolbar display.
  frmSysMgr.RefreshMenu

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error activating Workflow Designer form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Sub

Private Sub Form_Initialize()
  mfExitToWorkflow = False
  mfExpressionsChanged = False
  
  ReDim miSelectionOrder(0)
      
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Process key strokes.
  On Error GoTo ErrorTrap
  
  Dim sngXMove As Single
  Dim sngYMove As Single
  Dim bHandled As Boolean
  Dim iSelectedElementCount As Integer
  Dim iSelectedLinkCount As Integer
  Dim wfElement As VB.Control
  Dim wfLink As COAWF_Link
  Dim fCtrlPressed As Boolean
  Dim sngRequiredTop As Single
  Dim sngRequiredLeft As Single
  
  fCtrlPressed = ((Shift And vbCtrlMask) > 0)
  
  Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  End Select
  
  If KeyCode = vbKeyF4 Then
    EditMenu "ID_WorkflowProperties"
    bHandled = True
  End If

  If Not bHandled Then
    ' DELETE key deletes any selected controls.
    If (KeyCode = vbKeyDelete) And (abMenu.Tools("ID_WorkflowDelete").Enabled) Then
      EditMenu "ID_WorkflowDelete"
      bHandled = True
    End If
  End If
  
  If Not bHandled Then
    ' CTRL and ARROW keys move the selected controls.
    If fCtrlPressed _
      And (Application.AccessMode = accFull Or Application.AccessMode = accSupportMode) _
      And ((KeyCode = vbKeyLeft) _
        Or (KeyCode = vbKeyRight) _
        Or (KeyCode = vbKeyUp) _
        Or (KeyCode = vbKeyDown) _
        Or (KeyCode = vbKeyZ) _
        Or (KeyCode = vbKeyX) _
        Or (KeyCode = vbKeyC) _
        Or (KeyCode = vbKeyV) _
        Or (KeyCode = vbKeyA)) Then
        
      sngXMove = 0
      sngYMove = 0
  
      ' Determine which movement is being made.
      Select Case KeyCode
        Case vbKeyLeft
          sngXMove = -giSTANDARDMOVEMENT
        Case vbKeyRight
          sngXMove = giSTANDARDMOVEMENT
        Case vbKeyUp
          sngYMove = -giSTANDARDMOVEMENT
        Case vbKeyDown
          sngYMove = giSTANDARDMOVEMENT
        Case vbKeyZ
          If (abMenu.Tools("ID_WorkflowUndo").Enabled) Then
            EditMenu "ID_WorkflowUndo"
          End If
        Case vbKeyX
          If (abMenu.Tools("ID_WorkflowCut").Enabled) Then
            EditMenu "ID_WorkflowCut"
          End If
        Case vbKeyC
          If (abMenu.Tools("ID_WorkflowCopy").Enabled) Then
            EditMenu "ID_WorkflowCopy"
          End If
        Case vbKeyV
          If (abMenu.Tools("ID_WorkflowPaste").Enabled) Then
            EditMenu "ID_WorkflowPaste"
          End If
        Case vbKeyA
          If (abMenu.Tools("ID_SelectAll").Enabled) Then
            EditMenu "ID_SelectAll"
          End If
      End Select
  
      If (sngXMove <> 0) Or (sngYMove <> 0) Then
        Element_KeyMove sngXMove, sngYMove
        
        ResizeCanvas
      End If
  
      bHandled = True
    Else
      bHandled = ArrowSelect(KeyCode)
    End If
  End If

  If Not bHandled Then
    If KeyCode = vbKeyReturn Then
      iSelectedElementCount = SelectedElementCount
      iSelectedLinkCount = SelectedLinkCount
      
      If iSelectedElementCount + iSelectedLinkCount = 1 Then
        If iSelectedElementCount = 1 Then
          For Each wfElement In mcolwfElements
            If (wfElement.Highlighted) And (wfElement.Visible) Then
              ElementEdit wfElement
              Exit For
            End If
          Next wfElement
          Set wfElement = Nothing
        Else
          For Each wfLink In ASRWFLink1
            If (wfLink.Highlighted) And (wfLink.Visible) Then
              ASRWFLink1_DblClick wfLink.Index
              Exit For
            End If
          Next wfLink
          Set wfLink = Nothing
        End If
        
        bHandled = True
      End If
    End If
  End If

  If Not bHandled Then
    If KeyCode = vbKeyPageUp _
      Or KeyCode = vbKeyPageDown _
      Or KeyCode = vbKeyHome _
      Or KeyCode = vbKeyEnd Then

      sngRequiredLeft = picDefinition.Left
      sngRequiredTop = picDefinition.Top
      
      Select Case KeyCode
        Case vbKeyPageUp
          sngRequiredTop = picDefinition.Top + picContainer.Height
        Case vbKeyPageDown
          sngRequiredTop = picDefinition.Top - picContainer.Height
        Case vbKeyHome
          sngRequiredLeft = 0
          If fCtrlPressed Then
            sngRequiredTop = 0
          End If
        Case vbKeyEnd
          sngRequiredLeft = picContainer.Width - picDefinition.Width
          If fCtrlPressed Then
            sngRequiredTop = picContainer.Height - picDefinition.Height
          End If
      End Select
      
      If sngRequiredTop > 0 Then
        sngRequiredTop = 0
      ElseIf sngRequiredTop < picContainer.Height - picDefinition.Height Then
        sngRequiredTop = picContainer.Height - picDefinition.Height
      End If

      If sngRequiredLeft > 0 Then
        sngRequiredLeft = 0
      ElseIf sngRequiredLeft < picContainer.Width - picDefinition.Width Then
        sngRequiredLeft = picContainer.Width - picDefinition.Width
      End If
      
      picDefinition.Top = sngRequiredTop
      picDefinition.Left = sngRequiredLeft
      
      scrollVertical.value = -picDefinition.Top / mdblVerticalScrollRatio
      scrollHorizontal.value = -picDefinition.Left / mdblHorizontalScrollRatio

      bHandled = True
    End If
  End If

  If Not bHandled Then
    bHandled = frmSysMgr.tbMain.OnKeyDown(KeyCode, Shift)
  End If
  
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Function Element_KeyMove(pSngX As Single, pSngY As Single) As Boolean
  ' Move the control.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim wfElement As VB.Control
  Dim wfTempLink As COAWF_Link
  
  fOK = True

  For Each wfElement In mcolwfElements
    If (wfElement.Visible) _
      And (wfElement.Highlighted) Then
      
      wfElement.Move pSngX + wfElement.Left, pSngY + wfElement.Top
    
      For Each wfTempLink In ASRWFLink1
        If wfTempLink.StartElementIndex = wfElement.ControlIndex Or _
          wfTempLink.EndElementIndex = wfElement.ControlIndex Then
      
          FormatLink wfTempLink
        End If
      Next wfTempLink
      Set wfTempLink = Nothing
    End If
  Next wfElement
  Set wfElement = Nothing
  
  ' Flag screen as having changed
  IsChanged = True

TidyUpAndExit:
  Element_KeyMove = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub Form_Load()
  Dim wfTemp As VB.Control
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
  SetBlankIcon Me
  
  mfAppChanged = Application.Changed
  
  ReDim mactlUndoControls(0)
  ReDim mactlClipboardControls(0)
  
  ASRWFLink1(0).AppMajor = App.Major
  ASRWFLink1(0).AppMinor = App.Minor
  ASRWFLink1(0).AppRevision = App.Revision
  
  Set mcolwfElements = New Collection
  Set mcolwfSelectedElements = New Collection
  Set mcolwfSelectedLinks = New Collection
  
  picDefinition.BackColor = vbInactiveTitleBar
  
  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)

  If IsNew Then
    With picDefinition
      .Left = 0
      .Top = 0
    End With
    
    Set wfTemp = AddElement(elem_Begin)
    
    With wfTemp
      .Caption = "BEGIN"
      .Top = 500
      .Left = 500
    End With
        
    If Not mblnLoading Then
      SelectElement wfTemp
      
      ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
      miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
    End If
  End If

  cmdOK.Enabled = IsNew

  scrollVertical.SmallChange = SMALLSCROLL
  scrollHorizontal.SmallChange = SMALLSCROLL
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Reset the mousePointer if its been customised for ElementAdd mode.
  Me.MousePointer = vbNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' If the workflow has been modified then prompt the user
  ' whether or not to save the changes.
  On Error GoTo ErrorTrap
  
  Dim fDeleteWorkflow As Boolean
  Dim iLoop As Integer
  Dim fCancel As Boolean
  
  'JPD 20070327 Fault 12037
  ' Do not unload if there is a WebForm designer still open
  fCancel = False
  
  For iLoop = 0 To Forms.Count - 1
    If Forms(iLoop).Name = "frmWorkflowWFDesigner" Then
      fCancel = True
      Exit For
    End If
  Next iLoop
  
  If fCancel Then
    Cancel = True
    Exit Sub
  End If

  fDeleteWorkflow = mfNewWorkflow And (Not mfChanged)

  If mfChanged Then
    Select Case MsgBox("Apply workflow changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
        mblnDisplayWorkflowOpen = False
      Case vbYes
        Cancel = (Not SaveWorkflow())
        If Cancel = True Then mblnDisplayWorkflowOpen = False Else mblnDisplayWorkflowOpen = True
      Case vbNo
        mblnDisplayWorkflowOpen = True
        fDeleteWorkflow = mfNewWorkflow
        
        ' Restore the original expression definitions
        RestoreOriginalExpressions
        
        ' AE20080611 Fault #13193
        ' We've said no so lets make sure the sure button isnt enabled by us....
        If Not mfAppChanged Then
          Application.Changed = False
        End If
    End Select
  End If

  ' The user was creating a new workflow, but decided not to save changes,
  ' so remove the definition (name, description, etc.)
  If fDeleteWorkflow Then
    DeleteWorkflow
  End If
  
  ' Set the flag that determines whether we need to display the Workflow manager
  ' after the workflow designer is unloaded.
  mfExitToWorkflow = (UnloadMode = vbFormControlMenu) And mblnDisplayWorkflowOpen
  If Not mfChanged Then mfExitToWorkflow = True
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Sub DeleteWorkflow()
  ' Delete the workflow definition.
  On Error GoTo ErrorTrap

  daoDb.Execute "DELETE FROM tmpWorkflowElementItems WHERE elementID IN(SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & mlngWorkflowID & ")"
  daoDb.Execute "DELETE FROM tmpWorkflowElementColumns WHERE elementID IN(SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & mlngWorkflowID & ")"
  daoDb.Execute "DELETE FROM tmpWorkflowElementValidations WHERE elementID IN(SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & mlngWorkflowID & ")"
  daoDb.Execute "DELETE FROM tmpWorkflowElements WHERE workflowID=" & mlngWorkflowID
  daoDb.Execute "DELETE FROM tmpWorkflowLinks WHERE workflowID=" & mlngWorkflowID
  daoDb.Execute "DELETE FROM tmpWorkflows WHERE ID =" & mlngWorkflowID

ExitDeleteWorkflow:
  Exit Sub

ErrorTrap:
  Resume ExitDeleteWorkflow
  
End Sub



Private Function SaveWorkflow() As Boolean
  ' Save the workflow to the local database.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fTransStarted As Boolean
  Dim bHasTargetIdentifier As Boolean
  
  fTransStarted = False
  
  fOK = ValidateWorkflow(True, False, False)

  If fOK Then
    If mfPerge Then
      If WorkflowsWithStatus(WorkflowID, giWFSTATUS_COMPLETE) _
        Or WorkflowsWithStatus(WorkflowID, giWFSTATUS_ERROR) Then
        
        fOK = (MsgBox("Saving the definition will purge all instances of this workflow from the log." & vbCrLf & _
          "Do you wish to continue?", vbQuestion + vbYesNo, App.ProductName) = vbYes)
      End If
    End If
  End If
  
  If fOK Then
    Screen.MousePointer = vbHourglass
  
    With gobjProgress
      .Caption = "Workflow Designer"
      .Bar1Value = 0
      .Bar1MaxValue = 100
      .Bar1Caption = "Saving Workflow Design..."
      .Cancel = False
      .Time = False
      .OpenProgress
    End With
  
    ' Begin the database transaction.
    fTransStarted = True
    daoWS.BeginTrans

    ' Delete the existing element and link definitions for this workflow.
    daoDb.Execute "DELETE FROM tmpWorkflowElementItems WHERE elementID IN(SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & mlngWorkflowID & ")"
    daoDb.Execute "DELETE FROM tmpWorkflowElementItemValues WHERE itemID NOT IN(SELECT ID FROM tmpWorkflowElementItems)"
    daoDb.Execute "DELETE FROM tmpWorkflowElementColumns WHERE elementID IN(SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & mlngWorkflowID & ")"
    daoDb.Execute "DELETE FROM tmpWorkflowElementValidations WHERE elementID IN(SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & mlngWorkflowID & ")"
    daoDb.Execute "DELETE FROM tmpWorkflowElements WHERE workflowID=" & mlngWorkflowID
    daoDb.Execute "DELETE FROM tmpWorkflowLinks WHERE workflowID=" & mlngWorkflowID

    fOK = SaveElementsAndLinks(bHasTargetIdentifier)
    
    ' Find the workflow record.
    With recWorkflowEdit
      .Index = "idxWorkflowID"
      .Seek "=", mlngWorkflowID
      If .NoMatch Then
        Exit Function
      Else
        .Edit
        .Fields("name") = msWorkflowName
        .Fields("description") = msWorkflowDescription
        .Fields("pictureid") = IIf(mlngWorkflowPictureID = 0, Null, mlngWorkflowPictureID)
        .Fields("enabled") = mfWorkflowEnabled
        .Fields("initiationType") = miInitiationType
        .Fields("baseTable") = mlngBaseTableID
        .Fields("changed") = True
        .Fields("perge") = (.Fields("perge") Or mfPerge)
        .Fields("HasTargetIdentifier") = bHasTargetIdentifier
      End If

      .Update
    End With


  End If

ExitSaveWorkflow:
  If fOK Then
    mfNewWorkflow = False
    mfChanged = False
    mfPerge = False
    Application.Changed = True
    frmSysMgr.RefreshMenu

    'Commit transaction
    If fTransStarted Then
      daoWS.CommitTrans dbForceOSFlush
    End If
  Else
    'Rollback transaction
    If fTransStarted Then
      daoWS.Rollback
    End If
  End If

  gobjProgress.CloseProgress
  Screen.MousePointer = vbDefault
  SaveWorkflow = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume ExitSaveWorkflow
  
End Function

Public Function ValidateWorkflow(pfSaving As Boolean, _
  pfSilent As Boolean, _
  pfFix As Boolean) As Boolean
  ' Validate the workflow definition
  On Error GoTo ErrorTrap
  
  Dim wfElement As VB.Control
  Dim wfLink As COAWF_Link
  Dim wfLink2 As COAWF_Link
  Dim sMessage As String
  Dim sSubMessage As String
  Dim avOutboundFlowInfo() As Variant
  Dim iInboundFlowCount As Integer
  Dim avOutboundFlowCounts() As Variant
  Dim iLoop As Integer
  Dim iTerminatorCount As Integer
  Dim fContinue As Boolean
  Dim frmUsage As frmUsage
  Dim iElementIndex As Integer
  Dim sMessagePrefix As String
  Dim fElementOK As Boolean
  Dim fLinkOK As Boolean
  Dim sSQL As String
  Dim rsInfo As New ADODB.Recordset
  Dim bProgressDisplayed As Boolean
       
  ' Punch up a progress bar
  bProgressDisplayed = gobjProgress.Visible
  With gobjProgress
    .Caption = "Workflow Designer"
    .NumberOfBars = 1
    .Bar1Value = 1
    .Bar1MaxValue = 5
    .Bar1Caption = "Validating workflow..."
    .AVI = dbWorkflow
    .MainCaption = "Workflow"
    .Cancel = False
    .Time = False
    .OpenProgress
  End With
    
  sMessage = ""
  iTerminatorCount = 0
  
  CancelElementAddMode
  Screen.MousePointer = vbHourglass
  
  ' Clear the array of validation messages
  'Column 0 = The message
  'Column 1 = Associated element index
  ReDim mavValidationMessages(1, 0)
  mfFixableValidationFailures = False
  
  ' Perform generic validation checks.
  '   There must be at least one terminator.
  '   All elements have the required number of inbound and outbound flows.
  For Each wfElement In mcolwfElements
    
    fElementOK = wfElement.Visible
    If (Not fElementOK) Then
      ' Element might not be .visible but still valid
      ' if this method is called from the the Workflow properties screen.
      fElementOK = (wfElement.ControlIndex > 0)
      
      If fElementOK Then
        If (miLastActionFlag = giACTION_DELETECONTROLS) Then
          For iLoop = 1 To UBound(mactlUndoControls)
            If IsWorkflowElement(mactlUndoControls(iLoop)) Then
              If mactlUndoControls(iLoop).ControlIndex = wfElement.ControlIndex Then
                fElementOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
      End If
      
      'JPD 20060719 Fault 11339
      If fElementOK Then
        For iLoop = 1 To UBound(mactlClipboardControls)
          If IsWorkflowElement(mactlClipboardControls(iLoop)) Then
            If mactlClipboardControls(iLoop).ControlIndex = wfElement.ControlIndex Then
              fElementOK = False
              Exit For
            End If
          End If
        Next iLoop
      End If
    End If
    
    If fElementOK Then
      sMessagePrefix = ValidateElement_MessagePrefix(wfElement)
      
      If wfElement.ElementType = elem_Terminator Then
        iTerminatorCount = iTerminatorCount + 1
      End If
      
      iInboundFlowCount = 0
      ' Get the array of outbound flow information from the start element.
      ' Column 1 = Tag (see enums, or -1 if there's only a single outbound flow)
      ' Column 2 = Direction
      ' Column 3 = XOffset
      ' Column 4 = YOffset
      ' Column 5 = Maximum
      ' Column 6 = Minimum
      ' Column 7 = Description
      avOutboundFlowInfo = wfElement.OutboundFlows_Information
      
      ' Create an array of the count of outbound flows for each outbound flow.
      ReDim avOutboundFlowCounts(2, UBound(avOutboundFlowInfo, 2))
      For iLoop = 1 To UBound(avOutboundFlowInfo, 2)
        avOutboundFlowCounts(1, iLoop) = avOutboundFlowInfo(1, iLoop)
        avOutboundFlowCounts(2, iLoop) = 0
      Next iLoop
      
      For Each wfLink In ASRWFLink1
        fLinkOK = wfLink.Visible
        If (Not fLinkOK) Then
          ' Link might not be .visible but still valid
          ' if this method is called from the Workflow properties screen.
          fLinkOK = True
            
          If (miLastActionFlag = giACTION_DELETECONTROLS) Then
            For iLoop = 1 To UBound(mactlUndoControls)
              If TypeOf mactlUndoControls(iLoop) Is COAWF_Link Then
                If mactlUndoControls(iLoop).Index = wfLink.Index Then
                  fLinkOK = False
                  Exit For
                End If
              End If
            Next iLoop
          End If
          
          If fLinkOK Then
            If (miLastActionFlag = giACTION_SWAPCONTROL) Then
              If UBound(mactlUndoControls) >= 1 Then
                If TypeOf mactlUndoControls(1) Is COAWF_Link Then
                  If mactlUndoControls(1).Index = wfLink.Index Then
                    fLinkOK = False
                  End If
                End If
              End If
            End If
          End If

          'JPD 20060719 Fault 11339
          If fLinkOK Then
            For iLoop = 1 To UBound(mactlClipboardControls)
              If TypeOf mactlClipboardControls(iLoop) Is COAWF_Link Then
                If mactlClipboardControls(iLoop).Index = wfLink.Index Then
                  fLinkOK = False
                  Exit For
                End If
              End If
            Next iLoop
          End If
        End If
        
        If fLinkOK Then
          If wfLink.EndElementIndex = wfElement.ControlIndex Then
            iInboundFlowCount = iInboundFlowCount + 1
          End If
          
          If wfLink.StartElementIndex = wfElement.ControlIndex Then
            For iLoop = 1 To UBound(avOutboundFlowCounts, 2)
              If (wfLink.StartOutboundFlowCode < 0) Or _
                (avOutboundFlowCounts(1, iLoop) = wfLink.StartOutboundFlowCode) Then
                
                avOutboundFlowCounts(2, iLoop) = avOutboundFlowCounts(2, iLoop) + 1
                Exit For
              End If
            Next iLoop
          End If
        End If
      Next wfLink
      Set wfLink = Nothing
      
      ' Check that the maximum number of inbound flows has not been exceeded.
      If (wfElement.InboundFlows_Maximum > -1) And _
        (iInboundFlowCount > wfElement.InboundFlows_Maximum) Then
        
        ValidateWorkflow_AddMessage _
          sMessagePrefix & "Too many inbound flows", _
          wfElement.ControlIndex
      End If
      
      ' Check that the minimum number of inbound flows has been reached.
      If iInboundFlowCount < wfElement.InboundFlows_Minimum Then
        ValidateWorkflow_AddMessage _
          sMessagePrefix & "Not enough inbound flows", _
          wfElement.ControlIndex
      End If
      
      ' Check that the maximum number of outbound flows for each outbound flow has not been exceeded.
      For iLoop = 1 To UBound(avOutboundFlowCounts, 2)
        If (avOutboundFlowInfo(5, iLoop) > -1) And _
          (avOutboundFlowCounts(2, iLoop) > avOutboundFlowInfo(5, iLoop)) Then
          
          If UBound(avOutboundFlowInfo, 2) = 1 Then
            sSubMessage = ""
          Else
            sSubMessage = "'" & avOutboundFlowInfo(7, iLoop) & "' "
          End If
          
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Too many " & sSubMessage & "outbound flows", _
            wfElement.ControlIndex
        End If
      Next iLoop
    
      ' Check that the minimum number of outbound flows for each outbound flow has been reached.
      For iLoop = 1 To UBound(avOutboundFlowCounts, 2)
        If (avOutboundFlowInfo(6, iLoop) > -1) And _
          (avOutboundFlowCounts(2, iLoop) < avOutboundFlowInfo(6, iLoop)) Then
          
          If UBound(avOutboundFlowInfo, 2) = 1 Then
            sSubMessage = ""
          Else
            sSubMessage = "'" & avOutboundFlowInfo(7, iLoop) & "' "
          End If
          
          ValidateWorkflow_AddMessage _
            sMessagePrefix & "Not enough " & sSubMessage & "outbound flows", _
            wfElement.ControlIndex
        End If
      Next iLoop
            
      ' Perform the element specific checks.
      ' NB. All generic checks are made above (eg. number of inbound/outbound links)
      ValidateElement wfElement, pfFix
    End If
  Next wfElement
  Set wfElement = Nothing
  
  If (iTerminatorCount = 0) Then
    ValidateWorkflow_AddMessage _
      "General : No terminator element", _
      -1 ' -1 = non-specific element
  End If
  
  ' Display the validity failures to the user.
  fContinue = (UBound(mavValidationMessages, 2) = 0)
  
  Screen.MousePointer = vbDefault
  
  If Not pfSilent Then
    If fContinue Then
      If pfSaving Then
        If Not mfWorkflowEnabled Then
          
          gobjProgress.CloseProgress
          If MsgBox("The workflow definition is valid." & vbCrLf & _
            "Do you want to enable this definition?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            
            mfWorkflowEnabled = True
          
            sSQL = "SELECT COUNT(*) AS recCount" & _
              " FROM tmpWorkflowTriggeredLinks" & _
              " WHERE tmpWorkflowTriggeredLinks.deleted = FALSE" & _
              " AND tmpWorkflowTriggeredLinks.workflowID = " & CStr(mlngWorkflowID)
            Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            If rsInfo!reccount > 0 Then
              Application.ChangedWorkflowLink = True
            End If
            rsInfo.Close
            Set rsInfo = Nothing

          End If
        End If
      Else
        
        ' Hide the progress bar
        gobjProgress.CloseProgress
      
        MsgBox "The workflow definition is valid.", vbInformation + vbOKOnly, App.ProductName
      End If
    Else
      ' Display the validation messages.
      Set frmUsage = New frmUsage
      frmUsage.ResetList
  
      For iLoop = 1 To UBound(mavValidationMessages, 2)
        frmUsage.AddToList CStr(mavValidationMessages(0, iLoop)), mavValidationMessages(1, iLoop)
      Next iLoop
  
      Screen.MousePointer = vbDefault
      
      frmUsage.Width = (3 * Screen.Width / 4)
      frmUsage.Height = (Me.ScaleHeight / 2)
      
      If mfReadOnly Then
        mfFixableValidationFailures = False
      End If
      
      ' Hide the progress bar
      gobjProgress.CloseProgress
      
      If pfSaving Then
        frmUsage.ShowMessage "Workflow '" & Trim(msWorkflowName) & "'", "The Workflow definition is invalid for the reasons listed below." & _
          vbCrLf & "Saving the definition will force the workflow to be disabled." & _
          vbCrLf & " Do you wish to continue?", UsageCheckObject.Workflow, _
          IIf(mfFixableValidationFailures, USAGEBUTTONS_FIX, 0) + USAGEBUTTONS_PRINT + USAGEBUTTONS_YES + USAGEBUTTONS_NO + USAGEBUTTONS_SELECT, "validation"
      
        fContinue = (frmUsage.Choice = vbYes)
        
        If fContinue Then
          mfWorkflowEnabled = False
        End If
      Else
           
        frmUsage.ShowMessage "Workflow '" & Trim(msWorkflowName) & "'", "The Workflow definition is invalid for the reasons listed below.", _
          UsageCheckObject.Workflow, _
          IIf(mfFixableValidationFailures, USAGEBUTTONS_FIX, 0) + USAGEBUTTONS_PRINT + USAGEBUTTONS_OK + USAGEBUTTONS_SELECT, "validation"
      End If
      
      If frmUsage.Choice = vbRetry Then
        ' Highlight the element 'selected' in the usage check form.
        DeselectAllElements
        
        If frmUsage.Selection >= 0 Then
          iElementIndex = CInt(frmUsage.Selection)
          
          If iElementIndex > 0 Then
'            mcolwfElements(CStr(iElementIndex).HighLighted = True
            SelectElement mcolwfElements(CStr(iElementIndex))
  
            'JPD 20061129 Fault 11533 - Ensure the selected element is visible.
            MoveToItem mcolwfElements(CStr(iElementIndex))
            
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = iElementIndex
          
            RefreshMenu
          End If
        End If
      End If
      
      If frmUsage.Choice = vbIgnore Then
        If MsgBox("This will adjust the size & decimals of all Web Form input items to match the Stored Data columns that refer to them." _
          & vbCrLf & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
          fContinue = ValidateWorkflow(pfSaving, pfSilent, True)
          IsChanged = True
        End If
      End If
      
      UnLoad frmUsage
      Set frmUsage = Nothing
    End If
  End If
  
  
TidyUpAndExit:
  Screen.MousePointer = vbDefault
  gobjProgress.Visible = bProgressDisplayed
  
  ValidateWorkflow = fContinue
  Exit Function
  
ErrorTrap:
  fContinue = True
  Resume TidyUpAndExit
  
End Function


Private Function SaveElementsAndLinks(ByRef bHasTargetIdentifier As Boolean) As Boolean
  ' Save the definition of the given element.
  On Error GoTo ErrorTrap

  Dim fSaveOK As Boolean
  Dim wfElement As VB.Control
  Dim wfLink As COAWF_Link
  Dim alngIndexDirectory() As Long
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim fStartIndexDone As Boolean
  Dim fEndIndexDone As Boolean
  Dim asItems() As String
  Dim asItemValues() As String
  Dim avColumns() As Variant
  Dim asValidations() As String
  Dim lngElementID As Long
  Dim lngItemID As Long
  Dim objMisc As Misc
  
  Set objMisc = New Misc
  
  ReDim alngIndexDirectory(2, 0)
  ' Column 1 = element control index
  ' Column 2 = element record ID
  
  bHasTargetIdentifier = False
  
  ' Save the elements.
  For Each wfElement In mcolwfElements
    ' Do not save the dummy element array controls (index = 0).
    If wfElement.Visible Then
      'Add element definition
      With recWorkflowElementEdit
        .AddNew
  
        ' Use the same element record IDs as before if possible.
        lngElementID = -1
        For iLoop = 1 To UBound(malngIndexDirectory, 2)
          If malngIndexDirectory(1, iLoop) = wfElement.ControlIndex Then
            lngElementID = malngIndexDirectory(2, iLoop)
            Exit For
          End If
        Next iLoop
        If lngElementID >= 0 Then
          .Fields("ID") = lngElementID
        Else
          .Fields("ID") = UniqueColumnValue("tmpWorkflowElements", "ID")
        End If

        .Fields("workflowID") = mlngWorkflowID
        .Fields("type") = wfElement.ElementType
        .Fields("caption") = wfElement.Caption
        .Fields("connectionPairID") = 0
        .Fields("leftCoord") = wfElement.Left
        .Fields("topCoord") = wfElement.Top
        
        Select Case wfElement.ElementType
        Case elem_WebForm
          .Fields("identifier") = wfElement.Identifier
          .Fields("descriptionExprID") = wfElement.DescriptionExprID
          .Fields("descHasWorkflowName") = wfElement.DescriptionHasWorkflowName
          .Fields("descHasElementCaption") = wfElement.DescriptionHasElementCaption
          .Fields("webFormFGColor") = wfElement.WebFormFGColor
          .Fields("webFormBGColor") = wfElement.WebFormBGColor
          .Fields("webFormBGImageID") = wfElement.WebFormBGImageID
          
'           AE20080509 Fault #13162
'          .Fields("webFormDefaultFontName") = wfElement.Font.Name
'          .Fields("webFormDefaultFontSize") = wfElement.Font.Size
'          .Fields("webFormDefaultFontBold") = wfElement.Font.Bold
'          .Fields("webFormDefaultFontItalic") = wfElement.Font.Italic
'          .Fields("webFormDefaultFontStrikeThru") = wfElement.Font.Strikethrough
'          .Fields("webFormDefaultFontUnderline") = wfElement.Font.Underline
          .Fields("webFormBGImageLocation") = wfElement.WebFormBGImageLocation
          .Fields("webFormDefaultFontName") = wfElement.WebFormDefaultFont.Name
          .Fields("webFormDefaultFontSize") = wfElement.WebFormDefaultFont.Size
          .Fields("webFormDefaultFontBold") = wfElement.WebFormDefaultFont.Bold
          .Fields("webFormDefaultFontItalic") = wfElement.WebFormDefaultFont.Italic
          .Fields("webFormDefaultFontStrikeThru") = wfElement.WebFormDefaultFont.Strikethrough
          .Fields("webFormDefaultFontUnderline") = wfElement.WebFormDefaultFont.Underline
          .Fields("webFormWidth") = wfElement.WebFormWidth
          .Fields("webFormHeight") = wfElement.WebFormHeight
          .Fields("TimeoutFrequency") = wfElement.WebFormTimeoutFrequency
          .Fields("TimeoutPeriod") = wfElement.WebFormTimeoutPeriod
          .Fields("TimeoutExcludeWeekend") = wfElement.WebFormTimeoutExcludeWeekend
          
          .Fields("CompletionMessageType") = wfElement.WFCompletionMessageType
          .Fields("CompletionMessage") = wfElement.WFCompletionMessage
          .Fields("SavedForLaterMessageType") = wfElement.WFSavedForLaterMessageType
          .Fields("SavedForLaterMessage") = wfElement.WFSavedForLaterMessage
          .Fields("FollowOnFormsMessageType") = wfElement.WFFollowOnFormsMessageType
          .Fields("FollowOnFormsMessage") = wfElement.WFFollowOnFormsMessage
        
        Case elem_Email
          .Fields("identifier") = wfElement.Identifier
          .Fields("emailID") = wfElement.EmailID
          .Fields("emailCCID") = wfElement.EmailCCID
          .Fields("emailRecord") = wfElement.EmailRecord
          .Fields("emailSubject") = wfElement.EMailSubject
          .Fields("RecSelWebFormIdentifier") = wfElement.RecordSelectorWebFormIdentifier
          .Fields("RecSelIdentifier") = wfElement.RecordSelectorIdentifier
          
          .Fields("Attachment_Type") = wfElement.Attachment_Type
          .Fields("Attachment_File") = wfElement.Attachment_File
          .Fields("Attachment_WFElementIdentifier") = wfElement.Attachment_WFElementIdentifier
          .Fields("Attachment_WFValueIdentifier") = wfElement.Attachment_WFValueIdentifier
          .Fields("Attachment_DBColumnID") = wfElement.Attachment_DBColumnID
          .Fields("Attachment_DBRecord") = wfElement.Attachment_DBRecord
          .Fields("Attachment_DBElement") = wfElement.Attachment_DBElement
          .Fields("Attachment_DBValue") = wfElement.Attachment_DBValue
          
        Case elem_Decision
          .Fields("identifier") = wfElement.Identifier
          .Fields("decisionCaptionType") = wfElement.DecisionCaptionType
          .Fields("trueFlowType") = wfElement.DecisionFlowType
          .Fields("trueFlowIdentifier") = wfElement.TrueFlowIdentifier
          .Fields("trueFlowExprID") = wfElement.DecisionFlowExpressionID
        
        Case elem_StoredData
          .Fields("identifier") = wfElement.Identifier
          .Fields("dataAction") = wfElement.DataAction
          .Fields("dataTableID") = wfElement.DataTableID
          .Fields("dataRecord") = wfElement.DataRecord
          .Fields("RecSelWebFormIdentifier") = wfElement.RecordSelectorWebFormIdentifier
          .Fields("RecSelIdentifier") = wfElement.RecordSelectorIdentifier
          .Fields("DataRecordTable") = wfElement.DataRecordTableID
          .Fields("secondaryDataRecord") = wfElement.SecondaryDataRecord
          .Fields("secondaryRecSelWebFormIdentifier") = wfElement.SecondaryRecordSelectorWebFormIdentifier
          .Fields("secondaryRecSelIdentifier") = wfElement.SecondaryRecordSelectorIdentifier
          .Fields("SecondaryDataRecordTable") = wfElement.SecondaryDataRecordTableID
          .Fields("UseAsTargetIdentifier") = wfElement.UseAsTargetIdentifier
          bHasTargetIdentifier = bHasTargetIdentifier Or wfElement.UseAsTargetIdentifier
                 
        End Select
        .Update
        
        .Bookmark = .LastModified
  
        ReDim Preserve alngIndexDirectory(2, UBound(alngIndexDirectory, 2) + 1)
        alngIndexDirectory(1, UBound(alngIndexDirectory, 2)) = wfElement.ControlIndex
        alngIndexDirectory(2, UBound(alngIndexDirectory, 2)) = .Fields("ID")
        
        lngElementID = .Fields("ID")
      End With
      
      If (wfElement.ElementType = elem_WebForm) Or _
        (wfElement.ElementType = elem_Email) Then
        asItems = wfElement.Items
        
        For iLoop = 1 To UBound(asItems, 2)
          With recWorkflowElementItemEdit
            .AddNew
      
            .Fields("ID") = UniqueColumnValue("tmpWorkflowElementItems", "ID")
            .Fields("elementID") = lngElementID
            
            lngItemID = .Fields("ID")
            
            .Fields("Identifier") = asItems(9, iLoop)
            .Fields("ItemType") = asItems(2, iLoop)
            .Fields("caption") = asItems(3, iLoop)
            .Fields("DBColumnID") = asItems(4, iLoop)
            .Fields("DBRecord") = asItems(5, iLoop)
            
            .Fields("InputType") = val(asItems(6, iLoop))
            .Fields("InputSize") = val(asItems(7, iLoop))
            .Fields("InputDecimals") = val(asItems(8, iLoop))
          
            If val(asItems(6, iLoop)) = giEXPRVALUE_DATE Then
              If Len(asItems(10, iLoop)) > 0 Then
                .Fields("InputDefault") = objMisc.ConvertLocaleDateToSQL(asItems(10, iLoop))
              Else
                .Fields("InputDefault") = ""
              End If
            ElseIf val(asItems(6, iLoop)) = giEXPRVALUE_NUMERIC Then
              If Len(asItems(10, iLoop)) > 0 Then
                .Fields("InputDefault") = UI.ConvertNumberForSQL(asItems(10, iLoop))
              Else
                ' AE20080528 Fault #13185
                '.Fields("InputDefault") = "0"
                .Fields("InputDefault") = asItems(10, iLoop)
              End If
            Else
              .Fields("InputDefault") = asItems(10, iLoop)
            End If
            
            .Fields("WFFormIdentifier") = asItems(11, iLoop)
            .Fields("WFValueIdentifier") = asItems(12, iLoop)
            
            If (wfElement.ElementType = elem_Email) Then
              .Fields("RecSelWebFormIdentifier") = asItems(13, iLoop)
              .Fields("RecSelIdentifier") = asItems(14, iLoop)
            End If
            
            If (wfElement.ElementType = elem_WebForm) Then
              .Fields("LeftCoord") = asItems(13, iLoop)
              .Fields("TopCoord") = asItems(14, iLoop)
              .Fields("Width") = asItems(15, iLoop)
              .Fields("Height") = asItems(16, iLoop)
              .Fields("BackColor") = asItems(17, iLoop)
              .Fields("ForeColor") = asItems(18, iLoop)
              .Fields("FontName") = asItems(19, iLoop)
              .Fields("FontSize") = asItems(20, iLoop)
              .Fields("FontBold") = asItems(21, iLoop)
              .Fields("FontItalic") = asItems(22, iLoop)
              .Fields("FontStrikeThru") = asItems(23, iLoop)
              .Fields("FontUnderline") = asItems(24, iLoop)
              .Fields("PictureID") = asItems(25, iLoop)
              .Fields("PictureBorder") = (asItems(26, iLoop) = vbFixedSingle)
              .Fields("Alignment") = asItems(27, iLoop)
              .Fields("ZOrder") = asItems(28, iLoop)
              .Fields("TabIndex") = asItems(29, iLoop)
              .Fields("BackStyle") = asItems(30, iLoop)
              .Fields("BackColorEven") = asItems(31, iLoop)
              .Fields("BackColorOdd") = asItems(32, iLoop)
              .Fields("ColumnHeaders") = asItems(33, iLoop)
              .Fields("ForeColorEven") = asItems(34, iLoop)
              .Fields("ForeColorOdd") = asItems(35, iLoop)
              .Fields("HeaderBackColor") = asItems(36, iLoop)
              .Fields("HeadFontName") = asItems(37, iLoop)
              .Fields("HeadFontSize") = asItems(38, iLoop)
              .Fields("HeadFontBold") = asItems(39, iLoop)
              .Fields("HeadFontItalic") = asItems(40, iLoop)
              .Fields("HeadFontStrikeThru") = asItems(41, iLoop)
              .Fields("HeadFontUnderline") = asItems(42, iLoop)
              .Fields("Headlines") = asItems(43, iLoop)
              .Fields("TableID") = IIf(asItems(44, iLoop) = vbNullString, 0, asItems(44, iLoop))
              .Fields("ForeColorHighlight") = asItems(45, iLoop)
              .Fields("BackColorHighlight") = asItems(46, iLoop)
              
              If (recWorkflowElementItemEdit.Fields("itemType") = giWFFORMITEM_INPUTVALUE_DROPDOWN) Or _
                (recWorkflowElementItemEdit.Fields("itemType") = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) Then
                
                asItemValues = Split(asItems(47, iLoop), vbTab)
                
                For iLoop3 = 0 To UBound(asItemValues)
                  recWorkflowElementItemValuesEdit.AddNew
                  recWorkflowElementItemValuesEdit.Fields("itemID") = lngItemID
                  recWorkflowElementItemValuesEdit.Fields("value") = asItemValues(iLoop3)
                  recWorkflowElementItemValuesEdit.Fields("sequence") = iLoop3
                  recWorkflowElementItemValuesEdit.Update
                Next iLoop3
                
              End If
              
              .Fields("LookupTableID") = val(asItems(48, iLoop))
              .Fields("LookupColumnID") = val(asItems(49, iLoop))
              .Fields("RecordTableID") = val(asItems(50, iLoop))
              .Fields("Orientation") = val(asItems(51, iLoop))
              .Fields("RecordOrderID") = val(asItems(52, iLoop))
              .Fields("RecordFilterID") = val(asItems(53, iLoop))
              .Fields("Behaviour") = val(asItems(54, iLoop))
              .Fields("Mandatory") = asItems(55, iLoop)
              .Fields("CaptionType") = val(asItems(57, iLoop))
              .Fields("DefaultValueType") = val(asItems(58, iLoop))
              .Fields("VerticalOffsetBehaviour") = val(asItems(59, iLoop))
              .Fields("HorizontalOffsetBehaviour") = val(asItems(60, iLoop))
              .Fields("VerticalOffset") = val(asItems(61, iLoop))
              .Fields("HorizontalOffset") = val(asItems(62, iLoop))
              .Fields("HeightBehaviour") = val(asItems(63, iLoop))
              .Fields("WidthBehaviour") = val(asItems(64, iLoop))
              .Fields("PasswordType") = asItems(65, iLoop)
            
              If (recWorkflowElementItemEdit.Fields("itemType") = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) Then
                
                asItemValues = Split(asItems(66, iLoop), vbTab)
                
                For iLoop3 = 0 To UBound(asItemValues)
                  recWorkflowElementItemValuesEdit.AddNew
                  recWorkflowElementItemValuesEdit.Fields("itemID") = lngItemID
                  recWorkflowElementItemValuesEdit.Fields("value") = asItemValues(iLoop3)
                  recWorkflowElementItemValuesEdit.Fields("sequence") = iLoop3
                  recWorkflowElementItemValuesEdit.Update
                Next iLoop3
                
              End If
            
              .Fields("LookupFilterColumnID") = asItems(67, iLoop)
              .Fields("LookupFilterOperator") = asItems(68, iLoop)
              .Fields("LookupFilterValue") = asItems(69, iLoop)
              .Fields("LookupOrderID") = asItems(80, iLoop)
              .Fields("HotSpotIdentifier") = asItems(81, iLoop)
              .Fields("UseAsTargetIdentifier") = asItems(82, iLoop)
              
              bHasTargetIdentifier = bHasTargetIdentifier Or CBool(asItems(82, iLoop))
            End If
            
            .Fields("CalcID") = val(asItems(56, iLoop))
            .Fields("pageno") = val(asItems(78, iLoop))
            
            If (recWorkflowElementItemEdit.Fields("itemType") = giWFFORMITEM_BUTTON) Then
              .Fields("buttonstyle") = val(asItems(79, iLoop))
            End If

            .Update
          End With
        Next iLoop
      End If
      
      If (wfElement.ElementType = elem_StoredData) Then
        avColumns = wfElement.DataColumns
        
        For iLoop = 1 To UBound(avColumns, 2)
          With recWorkflowElementColumnEdit
            .AddNew
      
            .Fields("ID") = UniqueColumnValue("tmpWorkflowElementColumns", "ID")
            .Fields("elementID") = lngElementID
            .Fields("columnID") = avColumns(3, iLoop)
            .Fields("valueType") = avColumns(4, iLoop)
            .Fields("value") = avColumns(5, iLoop)
            .Fields("WFFormIdentifier") = avColumns(6, iLoop)
            .Fields("WFValueIdentifier") = avColumns(7, iLoop)
            
            If IsNull(avColumns(8, iLoop)) Then
              .Fields("DBColumnID") = 0
            Else
              .Fields("DBColumnID") = val(avColumns(8, iLoop))
            End If
            
            If IsNull(avColumns(9, iLoop)) Then
              .Fields("DBRecord") = 0
            Else
              .Fields("DBRecord") = val(avColumns(9, iLoop))
            End If
            
            .Fields("CalcID") = val(avColumns(10, iLoop))
      
            .Update
          End With
        Next iLoop
      End If
      
      If (wfElement.ElementType = elem_WebForm) Then
        asValidations = wfElement.Validations

        For iLoop = 1 To UBound(asValidations, 2)
          With recWorkflowElementValidationEdit
            .AddNew

            .Fields("ID") = UniqueColumnValue("tmpWorkflowElementValidations", "ID")
            .Fields("elementID") = lngElementID
            .Fields("exprID") = CLng(asValidations(1, iLoop))
            .Fields("type") = CInt(asValidations(2, iLoop))
            .Fields("message") = asValidations(3, iLoop)

            .Update
          End With
        Next iLoop
      End If
    End If
  Next wfElement
  Set wfElement = Nothing
  
  ' Update the connector pair ID values.
  For iLoop = 1 To UBound(alngIndexDirectory, 2)
    With recWorkflowElementEdit
      .Index = "idxElementID"
      .Seek "=", alngIndexDirectory(2, iLoop)

      If Not .NoMatch Then
        Set wfElement = mcolwfElements(CStr(alngIndexDirectory(1, iLoop)))
        
        For iLoop2 = 1 To UBound(alngIndexDirectory, 2)
          If alngIndexDirectory(1, iLoop2) = wfElement.ConnectorPairIndex Then
            .Edit
            .Fields("connectionPairID") = alngIndexDirectory(2, iLoop2)
            .Update
            Exit For
          End If
        Next iLoop2
      End If
    End With
  Next iLoop
  
  ' Save each link.
  For Each wfLink In ASRWFLink1
    ' Do not save the dummy element array control, or those in the clipboard/undo arrays.
    If wfLink.Visible Then
      'Add link definition
      With recWorkflowLinkEdit
        .AddNew

        .Fields("ID") = UniqueColumnValue("tmpWorkflowLinks", "ID")
        
        .Fields("workflowID") = mlngWorkflowID
        .Fields("startOutboundFlowCode") = wfLink.StartOutboundFlowCode

        fStartIndexDone = False
        fEndIndexDone = False
        For iLoop = 1 To UBound(alngIndexDirectory, 2)
          If alngIndexDirectory(1, iLoop) = wfLink.StartElementIndex Then
            .Fields("startElementID") = alngIndexDirectory(2, iLoop)
            fStartIndexDone = True
          End If
          
          If alngIndexDirectory(1, iLoop) = wfLink.EndElementIndex Then
            .Fields("endElementID") = alngIndexDirectory(2, iLoop)
            fEndIndexDone = True
          End If
        
          If fStartIndexDone And fEndIndexDone Then
            Exit For
          End If
        Next iLoop

        .Update
      End With
    End If
  Next wfLink
  Set wfLink = Nothing
  
  fSaveOK = True

TidyUpAndExit:
  Set objMisc = Nothing
  
  SaveElementsAndLinks = fSaveOK
  Exit Function
  
ErrorTrap:
  fSaveOK = False
  Resume TidyUpAndExit
  
End Function


Private Sub Form_Resize()

  'JPD 20030908 Fault 5756
  DisplayApplication

  Dim lngMinWidth As Long
  Dim lngHeight As Long

  If (Me.WindowState = vbMinimized) Or (mblnLoading) Then
    Exit Sub
  End If

  With fraButtons
    .Left = Me.ScaleWidth - .Width - 200
    .Top = Me.ScaleHeight - .Height - 100
  End With

  With picContainer
    .Top = 0
    .Left = 0
    .Width = Me.ScaleWidth - scrollVertical.Width


    lngHeight = fraButtons.Top - .Top - scrollHorizontal.Height - 100
    .Height = Maximum(lngHeight, 0)

    scrollVertical.Top = .Top
    scrollVertical.Left = .Left + .Width
    scrollVertical.Height = .Height

    scrollHorizontal.Left = .Left
    scrollHorizontal.Top = .Top + .Height
    scrollHorizontal.Width = .Width

    If picDefinition.Width < .Width Then
      picDefinition.Width = .Width
    End If
    If picDefinition.Height < .Height Then
      picDefinition.Height = .Height
    End If
  End With

  SetScrollBarValues

  If (scrollVertical.value = scrollVertical.Max) Then
    scrollVertical_Change
  End If

  If (scrollHorizontal.value = scrollHorizontal.Max) Then
    scrollHorizontal_Change
  End If

End Sub

Private Function SetScrollBarValues() As Boolean
  On Error GoTo ErrorTrap
  
  Dim lngMax As Long

  With scrollVertical
    If (picDefinition.Height <= picContainer.Height) Then
      .value = 0
      .Enabled = False
      mdblVerticalScrollRatio = 1
    Else
      .Enabled = True

      If (picDefinition.Height - picContainer.Height) > SCROLLMAX Then
        lngMax = SCROLLMAX
        mdblVerticalScrollRatio = (picDefinition.Height - picContainer.Height) / SCROLLMAX
      Else
        lngMax = picDefinition.Height - picContainer.Height
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
    If (picDefinition.Width <= picContainer.Width) Then
      .value = 0
      .Enabled = False
      mdblHorizontalScrollRatio = 1
    Else
      .Enabled = True

      If (picDefinition.Width - picContainer.Width) > SCROLLMAX Then
        lngMax = SCROLLMAX
        mdblHorizontalScrollRatio = (picDefinition.Width - picContainer.Width) / SCROLLMAX
      Else
        lngMax = picDefinition.Width - picContainer.Width
        mdblHorizontalScrollRatio = 1
      End If

      .Max = lngMax

      If lngMax > (picContainer.Width / mdblHorizontalScrollRatio) Then
        .LargeChange = CInt(picContainer.Width / mdblHorizontalScrollRatio)
      Else
        .LargeChange = CInt(lngMax * 9 / 10)
      End If
    End If
  End With
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Function



Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ErrorTrap
  
  ' Display the Screen manager form if we are not exiting the system.
'  If mfExitToWorkflow Then
  
    With frmSysMgr
      If .frmWorkflowOpen Is Nothing Then
        Set .frmWorkflowOpen = New SystemMgr.frmWorkflowOpen
      End If
            
      .frmWorkflowOpen.WorkflowID = mlngWorkflowID

      'JPD 20060921 Fault 11003
      .frmWorkflowOpen.RefreshWorkflows
      .frmWorkflowOpen.SelectWorkflow

      .frmWorkflowOpen.Show
      .frmWorkflowOpen.SetFocus
      .RefreshMenu
    End With
'  End If

  Unhook Me.hWnd

 
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Sub picDefinition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorTrap

  ' Only handle left button presses here.
  If (Button <> vbLeftButton) Or (Not mblnStarted) Then
    Exit Sub
  End If
  
  ' Deselect all screen controls unless the SHIFT or CTRL keys are pressed.
  If ((Shift And vbShiftMask) = 0) And ((Shift And vbCtrlMask) = 0) Then
    DeselectAllElements
  End If
  
  If Not InElementAddMode Then
    ' Start the multi-selection frame.
    mfMultiSelecting = True
    mlngMultiSelectionXStart = x
    mlngMultiSelectionYStart = y
      
    ' Position and display the multi-selection box.
    With asrboxMultiSelection
      .Left = mlngMultiSelectionXStart
      .Top = mlngMultiSelectionYStart
      .Width = 0
      .Height = 0
      .Visible = True
      .ZOrder 0
    End With
  End If
  
  RefreshMenu

TidyUpAndExit:
  ' Disassociate object variables.
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub


Private Sub picDefinition_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Position and size the multi-selection lines as required.
  On Error GoTo ErrorTrap
  
  Dim lngTop As Long
  Dim lngLeft As Long
  Dim lngRight As Long
  Dim lngBottom As Long
  Dim lngRightLimit As Long
  Dim lngBottomLimit As Long

  mblnDragging = False
  
  If InElementAddMode And Me.MousePointer <> vbCustom Then
    SetElementAddPointer
  End If

  If mfMultiSelecting Then
    ' Calculate the cordinates of the multi-selection area.
    If x < mlngMultiSelectionXStart Then
      lngLeft = x
      lngRight = mlngMultiSelectionXStart
    Else
      lngLeft = mlngMultiSelectionXStart
      lngRight = x
    End If
      
    If y < mlngMultiSelectionYStart Then
      lngTop = y
      lngBottom = mlngMultiSelectionYStart
    Else
      lngTop = mlngMultiSelectionYStart
      lngBottom = y
    End If

    ' Limit the multi-selection area to the form or tab page area.
    lngRightLimit = picDefinition.Width
    lngBottomLimit = picDefinition.Height
      
    If lngLeft < 0 Then lngLeft = 0
    If lngRight > lngRightLimit Then lngRight = lngRightLimit
    If lngTop < 0 Then lngTop = 0
    If lngBottom > lngBottomLimit Then lngBottom = lngBottomLimit
    
    ' Size and position the multi-selection box.
    With asrboxMultiSelection
      .Left = lngLeft
      .Top = lngTop
      .Width = lngRight - lngLeft
      .Height = lngBottom - lngTop
    End With

  End If

  Me.Refresh

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub


Private Sub picDefinition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Select control that lie within the multi-selection area.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fInVerticalBand As Boolean
  Dim fInHorizontalBand As Boolean
  Dim sngSelectionTop As Single
  Dim sngSelectionLeft As Single
  Dim sngSelectionRight As Single
  Dim sngSelectionBottom As Single
  Dim lngXMouse As Long
  Dim lngYMouse As Long
  Dim wfTemp As VB.Control
  Dim wfTemp2 As VB.Control
  Dim wfTempLink As COAWF_Link
  Dim bInSelectionBand As Boolean
  Dim iOriginalPointer As Integer
  
  If Not mblnStarted Then
    Exit Sub
  End If
  
  iOriginalPointer = Screen.MousePointer
  
  Select Case Button
    ' Handle left button presses.
    Case vbLeftButton
    
      If InElementAddMode Then
        
        DeselectAllElements
        
        If abMenu.Tools("ID_WFElement_Terminator").Checked Then
          Set wfTemp = AddElement(elem_Terminator)
        End If
        
        If abMenu.Tools("ID_WFElement_WebForm").Checked Then
          Set wfTemp = AddElement(elem_WebForm)
        End If
        
        If abMenu.Tools("ID_WFElement_Email").Checked Then
          Set wfTemp = AddElement(elem_Email)
        End If
        
        If abMenu.Tools("ID_WFElement_Decision").Checked Then
          Set wfTemp = AddElement(elem_Decision)
        End If
        
        If abMenu.Tools("ID_WFElement_StoredData").Checked Then
          Set wfTemp = AddElement(elem_StoredData)
        End If
        
        If abMenu.Tools("ID_WFElement_SummingJunction").Checked Then
          Set wfTemp = AddElement(elem_SummingJunction)
        End If
        
        If abMenu.Tools("ID_WFElement_Or").Checked Then
          Set wfTemp = AddElement(elem_Or)
        End If
        
        If abMenu.Tools("ID_WFElement_Connector").Checked Then
          Set wfTemp = AddElement(elem_Connector1)
          Set wfTemp2 = AddElement(elem_Connector2)
          
          'JPD 20070713 Fault 12255
          ' Add the first connector to the Undo array as it will have been removed when the
          ' second connector was added
          ReDim Preserve mactlUndoControls(UBound(mactlUndoControls) + 1)
          Set mactlUndoControls(UBound(mactlUndoControls)) = wfTemp
          
          With wfTemp
            .Left = x - .InboundFlow_XOffset
            .Top = y - .InboundFlow_YOffset
            
            wfTemp2.Left = .Left + .Width + 200
            wfTemp2.Top = .Top
            
            .ConnectorPairIndex = wfTemp2.ControlIndex
            wfTemp2.ConnectorPairIndex = .ControlIndex
            
            .Caption = NextConnectorCaption
            wfTemp2.Caption = .Caption
            
            ' AE20080609 Fault #13202
            wfTemp2.Visible = True
            
            '.HighLighted = True
            SelectElement wfTemp
          End With
        End If
          
        If Not wfTemp Is Nothing Then
          wfTemp.Left = x - wfTemp.InboundFlow_XOffset
          wfTemp.Top = y - wfTemp.InboundFlow_YOffset
          
          ' AE20080609 Fault #13202
          wfTemp.Visible = True
          
          ' AE20080722 Fault #13283, #13284, #13285
          If Not mblnLoading Then
            SelectElement wfTemp
      
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = miControlIndex
          End If
    
          'JPD 20070321 Fault 11936
          ResizeCanvas
        End If
      ElseIf mfMultiSelecting Then
        Screen.MousePointer = vbHourglass
    
        ' End the multi-selection.
        mfMultiSelecting = False
        
        With asrboxMultiSelection
          sngSelectionTop = .Top
          sngSelectionBottom = .Top + .Height
          sngSelectionLeft = .Left
          sngSelectionRight = .Left + .Width
          .Visible = False
        End With
        
        If mcolwfElements.Count > 0 Then
          UI.LockWindow Me.hWnd
          
          For Each wfTemp In mcolwfElements
            With wfTemp
              If .Visible Then
                fInVerticalBand = (.Left < sngSelectionRight) And (.Left + .Width > sngSelectionLeft)
                fInHorizontalBand = (.Top < sngSelectionBottom) And (.Top + .Height > sngSelectionTop)
                      
                bInSelectionBand = fInVerticalBand And fInHorizontalBand
                        
                If bInSelectionBand Then
  '                .HighLighted = True
                  SelectElement wfTemp
                  .ZOrder 0
                End If
              End If
            End With
          Next wfTemp
          Set wfTemp = Nothing
        
          UI.UnlockWindow
        End If
        
        ' Mark the screen as having changed.
        frmSysMgr.RefreshMenu
      End If

    ' Handle right button presses.
    Case vbRightButton
      UI.GetMousePos lngXMouse, lngYMouse
      mlngXDrop = x
      mlngYDrop = y
      
      abMenu.Bands("ElementBand").TrackPopup -1, -1
  End Select
  
  RefreshMenu
  
TidyUpAndExit:
  Set wfTemp = Nothing
  Set wfTemp2 = Nothing
  
  ' Close the progress bar
  gobjProgress.CloseProgress
  
  ' Reset the screen mousepointer.
  Screen.MousePointer = iOriginalPointer
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub


Private Sub scrollHorizontal_Change()
  picDefinition.Left = (CLng(scrollHorizontal.value) * -1) * mdblHorizontalScrollRatio

End Sub


Private Sub scrollVertical_Change()
  picDefinition.Top = (CLng(scrollVertical.value) * -1) * mdblVerticalScrollRatio

End Sub


Public Property Get WorkflowID() As Long
  ' Return the current workflow id.
  WorkflowID = mlngWorkflowID
  
End Property

Public Property Let WorkflowID(pLngNewID As Long)
  ' Set the current workflow id.
  mlngWorkflowID = pLngNewID

  ' Load the workflow.
  mblnLoading = True
  mblnStarted = False
  If Not LoadWorkflow Then
    MsgBox "Unable to load workflow." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If

  IsChanged = False
  mfReadOnly = WorkflowsWithStatus(mlngWorkflowID, giWFSTATUS_INPROGRESS) _
    Or (Application.AccessMode <> accFull And Application.AccessMode <> accSupportMode) Or mbLocked

  SetLastActionFlag giACTION_NOACTION
  
  mblnLoading = False
  
  RefreshMenu
  
End Property



Private Function LoadWorkflow() As Boolean
  ' Load elements & links onto the definition screen.
  On Error GoTo ErrorTrap

  Dim fLoadOk As Boolean
  Dim wfElement As VB.Control
  
  Screen.MousePointer = vbHourglass

  ' Find the workflow definition in the database.
  With recWorkflowEdit
    .Index = "idxWorkflowID"
    .Seek "=", mlngWorkflowID
    fLoadOk = (Not .NoMatch)
  End With

  ' Load the workflow properties.
  If fLoadOk Then
    ' Lock the screen refeshing.
    UI.LockWindow Me.hWnd

    'Set form properties from workflow definition
    With recWorkflowEdit

'      mfNewWorkflow = IIf(IsNull(.Fields("new")), True, .Fields("new"))
      msWorkflowName = IIf(IsNull(.Fields("name")), "", .Fields("name"))
      msWorkflowDescription = IIf(IsNull(.Fields("description")), "", .Fields("description"))
      mlngWorkflowPictureID = IIf(IsNull(.Fields("pictureid")), 0, .Fields("pictureid"))
      msExternalInitiationQueryString = IIf(IsNull(.Fields("queryString")), "", .Fields("queryString"))
      mfWorkflowEnabled = IIf(IsNull(.Fields("enabled")), False, .Fields("enabled"))
      miInitiationType = IIf(IsNull(.Fields("initiationType")), WORKFLOWINITIATIONTYPE_MANUAL, .Fields("initiationType"))
      mlngBaseTableID = IIf(IsNull(.Fields("baseTable")), 0, .Fields("baseTable"))
      miRecSelType = IIf(miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL, giWFRECSEL_INITIATOR, _
        IIf(miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED, _
          giWFRECSEL_TRIGGEREDRECORD, giWFRECSEL_UNIDENTIFIED))
      mfChanged = False
      mfPerge = False
      mbLocked = IIf(IsNull(.Fields("locked")), False, .Fields("locked"))
      
      ' Set the screen caption and size.
      Me.Caption = "Workflow Designer - " & IIf(IsNull(.Fields("name")), "unnamed", .Fields("name")) & vbNullString
'''      Me.Height = IIf(IsNull(.Fields("height")), gLngDFLTSCREENHEIGHT, .Fields("height") + 450)
'''      Me.Width = IIf(IsNull(.Fields("width")), gLngDFLTSCREENWIDTH, .Fields("width"))
    
    End With
  End If

  ' Clear the array that records the record ID and control index values.
  ' Column 1 = element control index
  ' Column 2 = element record ID
  ' Column 3 = element type
  ReDim malngIndexDirectory(3, 0)
  
  ReDim maobjOriginalExpressions(0)
  
  ' Load the elements and links
  If Not IsNew Then
    LoadElementsAndLinks
    
    RememberOriginalExpressions
  End If

  DeselectAllElements

  IsChanged = False

TidyUpAndExit:
  ' Unlock the window refreshing.
  UI.UnlockWindow

  ' Position the form.
'''  Me.Top = Int((Forms(0).ScaleHeight - Me.Height) / 2)
'''  Me.Left = Int((Forms(0).ScaleWidth - Me.Width) / 2)

  ' Reset the screen mousepointer.
  Screen.MousePointer = vbDefault

  LoadWorkflow = fLoadOk
  Exit Function

ErrorTrap:
  fLoadOk = False
  MsgBox "Error loading workflow." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function

Public Function LoadElementsAndLinks() As Boolean

  ' Load elements and links.
  On Error GoTo ErrorTrap

  Dim fLoadOk As Boolean
  Dim wfElement As VB.Control
  Dim wfLink As COAWF_Link
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim fStartIndexDone As Boolean
  Dim fEndIndexDone As Boolean
  Dim iStartWFElementType As Integer
  Dim iEndWFElementType As Integer
  Dim avOutboundFlowInfo() As Variant
  Dim iOutboundFlowIndex As Integer
  Dim sngMaxX As Single
  Dim sngMaxY As Single
  Dim asItems() As String
  Dim avColumns() As Variant
  Dim asValidations() As String
  Dim sDescription As String
  Dim objFont As StdFont
  Dim objMisc As Misc
  Dim iControlValueSequence As Integer
  Dim sControlValueList As String
  Dim sTemp As String
  Dim lngArraySize As Long
  
  Set objMisc = New Misc
  
  fLoadOk = True
  sngMaxX = 0
  sngMaxY = 0
  
  If mlngWorkflowID = 0 Then
    LoadElementsAndLinks = True
    Exit Function
  End If

  Screen.MousePointer = vbHourglass

  ' Clear any existing elements and links from the workflow definition.
  ClearFlowchart True

  ' Locate the element definitions
  With recWorkflowElementEdit                  ' Indent 02 - start
    .Index = "idxWorkflowID"
    .Seek ">=", mlngWorkflowID

    If Not .NoMatch Then            ' Indent 03 - start

      ' Add elements to the form
      Do While Not .EOF             ' Indent 04 - start
        If .Fields("workflowID") <> mlngWorkflowID Then
          Exit Do
        End If

        Set wfElement = AddElement(.Fields("type"))
        If Not wfElement Is Nothing Then
          ' Set the control's position & other properties.
          wfElement.Caption = IIf(IsNull(.Fields("caption")), "", .Fields("caption"))
          wfElement.Top = IIf(IsNull(.Fields("topCoord")), 0, .Fields("topCoord"))
          wfElement.Left = IIf(IsNull(.Fields("leftCoord")), 0, .Fields("leftCoord"))

          Select Case wfElement.ElementType
          Case elem_WebForm
            wfElement.Identifier = IIf(IsNull(.Fields("identifier")), "", .Fields("identifier"))
            wfElement.DescriptionExprID = IIf(IsNull(.Fields("descriptionExprID")), 0, .Fields("descriptionExprID"))
            wfElement.DescriptionHasWorkflowName = IIf(IsNull(.Fields("descHasWorkflowName")), False, .Fields("descHasWorkflowName"))
            wfElement.DescriptionHasElementCaption = IIf(IsNull(.Fields("descHasElementCaption")), False, .Fields("descHasElementCaption"))
            wfElement.WebFormFGColor = IIf(IsNull(.Fields("WebFormFGColor")), vbBlack, .Fields("WebFormFGColor"))
            wfElement.WebFormBGColor = IIf(IsNull(.Fields("WebFormBGColor")), vbWhite, .Fields("WebFormBGColor"))
            wfElement.WebFormBGImageID = IIf(IsNull(.Fields("WebFormBGImageID")), 0, .Fields("WebFormBGImageID"))
            wfElement.WebFormBGImageLocation = IIf(IsNull(.Fields("WebFormBGImageLocation")), 0, .Fields("WebFormBGImageLocation"))
            wfElement.WebFormTimeoutFrequency = IIf(IsNull(.Fields("TimeoutFrequency")), 0, .Fields("TimeoutFrequency"))
            wfElement.WebFormTimeoutPeriod = IIf(IsNull(.Fields("TimeoutPeriod")), TIMEOUT_DAY, .Fields("TimeoutPeriod"))
            wfElement.WebFormTimeoutExcludeWeekend = IIf(IsNull(.Fields("TimeoutExcludeWeekend")), False, .Fields("TimeoutExcludeWeekend"))
            wfElement.WebFormWidth = IIf(IsNull(.Fields("WebFormWidth")), 5000, .Fields("WebFormWidth"))
            wfElement.WebFormHeight = IIf(IsNull(.Fields("WebFormHeight")), 5000, .Fields("WebFormHeight"))
            
            Set objFont = New StdFont
            objFont.Name = IIf(IsNull(.Fields("WebFormDefaultFontName")), gobjDefaultScreenFont.Name, .Fields("WebFormDefaultFontName"))
            objFont.Size = IIf(IsNull(.Fields("WebFormDefaultFontSize")), gobjDefaultScreenFont.Size, .Fields("WebFormDefaultFontSize"))
            objFont.Bold = IIf(IsNull(.Fields("WebFormDefaultFontBold")), 0, .Fields("WebFormDefaultFontBold"))
            objFont.Italic = IIf(IsNull(.Fields("WebFormDefaultFontItalic")), 0, .Fields("WebFormDefaultFontItalic"))
            objFont.Strikethrough = IIf(IsNull(.Fields("WebFormDefaultFontStrikeThru")), 0, .Fields("WebFormDefaultFontStrikeThru"))
            objFont.Underline = IIf(IsNull(.Fields("WebFormDefaultFontUnderline")), 0, .Fields("WebFormDefaultFontUnderline"))
           
            Set wfElement.WebFormDefaultFont = objFont
            Set objFont = Nothing
            
            wfElement.WFCompletionMessageType = IIf(IsNull(.Fields("CompletionMessageType")), MESSAGE_SYSTEMDEFAULT, .Fields("CompletionMessageType"))
            wfElement.WFCompletionMessage = IIf(IsNull(.Fields("CompletionMessage")), "", .Fields("CompletionMessage"))
            wfElement.WFSavedForLaterMessageType = IIf(IsNull(.Fields("SavedForLaterMessageType")), MESSAGE_SYSTEMDEFAULT, .Fields("SavedForLaterMessageType"))
            wfElement.WFSavedForLaterMessage = IIf(IsNull(.Fields("SavedForLaterMessage")), "", .Fields("SavedForLaterMessage"))
            wfElement.WFFollowOnFormsMessageType = IIf(IsNull(.Fields("FollowOnFormsMessageType")), MESSAGE_SYSTEMDEFAULT, .Fields("FollowOnFormsMessageType"))
            wfElement.WFFollowOnFormsMessage = IIf(IsNull(.Fields("FollowOnFormsMessage")), "", .Fields("FollowOnFormsMessage"))
            
          Case elem_Email
            wfElement.Identifier = IIf(IsNull(.Fields("identifier")), "", .Fields("identifier"))
            wfElement.EmailID = IIf(IsNull(.Fields("emailID")), 0, .Fields("emailID"))
            wfElement.EmailCCID = IIf(IsNull(.Fields("emailCCID")), 0, .Fields("emailCCID"))
            wfElement.EmailRecord = IIf(IsNull(.Fields("emailRecord")), 0, .Fields("emailRecord"))
            wfElement.EMailSubject = IIf(IsNull(.Fields("emailSubject")), "", .Fields("emailSubject"))
            wfElement.RecordSelectorWebFormIdentifier = IIf(IsNull(.Fields("RecSelWebFormIdentifier")), "", .Fields("RecSelWebFormIdentifier"))
            wfElement.RecordSelectorIdentifier = IIf(IsNull(.Fields("RecSelIdentifier")), "", .Fields("RecSelIdentifier"))
          
            wfElement.Attachment_Type = IIf(IsNull(.Fields("Attachment_Type")), giWFEMAILITEM_UNKNOWN, .Fields("Attachment_Type"))
            wfElement.Attachment_File = IIf(IsNull(.Fields("Attachment_File")), "", .Fields("Attachment_File"))
            wfElement.Attachment_WFElementIdentifier = IIf(IsNull(.Fields("Attachment_WFElementIdentifier")), "", .Fields("Attachment_WFElementIdentifier"))
            wfElement.Attachment_WFValueIdentifier = IIf(IsNull(.Fields("Attachment_WFValueIdentifier")), "", .Fields("Attachment_WFValueIdentifier"))
            wfElement.Attachment_DBColumnID = IIf(IsNull(.Fields("Attachment_DBColumnID")), 0, .Fields("Attachment_DBColumnID"))
            wfElement.Attachment_DBRecord = IIf(IsNull(.Fields("Attachment_DBRecord")), 0, .Fields("Attachment_DBRecord"))
            wfElement.Attachment_DBElement = IIf(IsNull(.Fields("Attachment_DBElement")), "", .Fields("Attachment_DBElement"))
            wfElement.Attachment_DBValue = IIf(IsNull(.Fields("Attachment_DBValue")), "", .Fields("Attachment_DBValue"))
          
          Case elem_Decision
            wfElement.Identifier = IIf(IsNull(.Fields("identifier")), "", .Fields("identifier"))
            wfElement.DecisionCaptionType = IIf(IsNull(.Fields("decisionCaptionType")), decisionCaption_T_F, .Fields("decisionCaptionType"))
            wfElement.DecisionFlowType = IIf(IsNull(.Fields("trueFlowType")), decisionFlowType_Button, .Fields("trueFlowType"))
            wfElement.TrueFlowIdentifier = IIf(IsNull(.Fields("trueFlowIdentifier")), "", .Fields("trueFlowIdentifier"))
            wfElement.DecisionFlowExpressionID = IIf(IsNull(.Fields("trueFlowExprID")), 0, .Fields("trueFlowExprID"))

          Case elem_StoredData
            wfElement.Identifier = IIf(IsNull(.Fields("identifier")), "", .Fields("identifier"))
            wfElement.DataAction = IIf(IsNull(.Fields("dataAction")), DATAACTION_INSERT, .Fields("dataAction"))
            wfElement.DataTableID = IIf(IsNull(.Fields("dataTableID")), 0, .Fields("dataTableID"))
            wfElement.DataRecord = IIf(IsNull(.Fields("dataRecord")), 0, .Fields("dataRecord"))
            wfElement.RecordSelectorWebFormIdentifier = IIf(IsNull(.Fields("RecSelWebFormIdentifier")), "", .Fields("RecSelWebFormIdentifier"))
            wfElement.RecordSelectorIdentifier = IIf(IsNull(.Fields("RecSelIdentifier")), "", .Fields("RecSelIdentifier"))
            wfElement.DataRecordTableID = IIf(IsNull(.Fields("DataRecordTable")), 0, .Fields("DataRecordTable"))
            wfElement.SecondaryDataRecord = IIf(IsNull(.Fields("secondaryDataRecord")), 0, .Fields("secondaryDataRecord"))
            wfElement.SecondaryRecordSelectorWebFormIdentifier = IIf(IsNull(.Fields("secondaryRecSelWebFormIdentifier")), "", .Fields("secondaryRecSelWebFormIdentifier"))
            wfElement.SecondaryRecordSelectorIdentifier = IIf(IsNull(.Fields("secondaryRecSelIdentifier")), "", .Fields("secondaryRecSelIdentifier"))
            wfElement.SecondaryDataRecordTableID = IIf(IsNull(.Fields("secondaryDataRecordTable")), 0, .Fields("secondaryDataRecordTable"))
            wfElement.UseAsTargetIdentifier = IIf(IsNull(.Fields("UseAsTargetIdentifier")), 0, .Fields("UseAsTargetIdentifier"))
            
          End Select
          
          ReDim asItems(0)
          
          If (wfElement.ElementType = elem_WebForm) Or _
            (wfElement.ElementType = elem_Email) Then
            ReDim asItems(WFITEMPROPERTYCOUNT, 0)
            ' Col 0 = Used to store automatically generated sql id.
            ' Col 1 = Description - NOT identifier (use item 9 if you want the identifier)
            ' Col 2 = Item Type
            ' Col 3 = Caption
            ' Col 4 = DB Column ID
            ' Col 5 = DB Record
            ' Col 6 = Input Return Type
            ' Col 7 = Input Size
            ' Col 8 = Input Decimals
            ' Col 9 = Input Identifier
            ' Col 10 = Input Default
            ' Col 11 = WF Form Identifier
            ' Col 12 = WF Value Identifier
            ' Col 13 = Left (webform), RecSelWebForm (email)
            ' Col 14 = Top (webform), RecSelector (email)
            ' Col 15 = Width
            ' Col 16 = Height
            ' Col 17 = Background Colour
            ' Col 18 = Foreground Colour
            ' Col 19 = Font-Name
            ' Col 20 = Font-Size
            ' Col 21 = Font-Bold
            ' Col 22 = Font-Italic
            ' Col 23 = Font-StrikeThru
            ' Col 24 = Font-Underline
            ' Col 25 = PictureID
            ' Col 26 = PictureBorder
            ' Col 27 = Alignment
            ' Col 28 = ZOrder
            ' Col 29 = TabIndex
            ' Col 30 = BackStyle
            ' Col 31 = BackColorEven
            ' Col 32 = BackColorOdd
            ' Col 33 = ColumnHeaders
            ' Col 34 = ForeColorEven
            ' Col 35 = ForeColorOdd
            ' Col 36 = HeaderBackColor
            ' Col 37 = HeadFont-Name
            ' Col 38 = HeadFont-Size
            ' Col 39 = HeadFont-Bold
            ' Col 40 = HeadFont-Italic
            ' Col 41 = HeadFont-StrikeThru
            ' Col 42 = HeadFont-Underline
            ' Col 43 = Headlines
            ' Col 44 = TableID
            ' Col 45 = ForeColorHighlight
            ' Col 46 = BackColorHighlight
            ' Col 47 = Control Values
            ' Col 48 = Lookup Table ID
            ' Col 49 = Lookup Column ID
            ' Col 50 = Record Table ID
            ' Col 51 = Orientation
            ' Col 52 = Record Order ID
            ' Col 53 = Record Filter ID
            ' Col 54 = Behaviour [Button Action = 0 (submit), = 1 (save for later)]
            ' Col 55 = Mandatory
            ' Col 56 = Calculation Expression ID
            ' Col 57 = Caption Type
            ' Col 58 = Default Value Type
            ' Col 59 = Vertical Offset Behaviour [Top = 0, Bottom = 1]
            ' Col 60 = Horizontal Offset Behaviour [Left = 0, Right = 1]
            ' Col 61 = Vertical Offset
            ' Col 62 = Horizontal Offset
            ' Col 63 = Height Behaviour [Fixed = 0, Full = 1]
            ' Col 64 = Width Behaviour [Fixed = 0, Full = 1]
            ' Col 65 = PasswordType [Hide text]
            ' Col 66 = File Extensions
            ' Col 67 = Lookup Filter Column ID
            ' Col 68 = Lookup Filter Operator
            ' Col 69 = Lookup Filter Value
            ' Col 70 = Lookup Order ID
            
            ' NB. IF YOU ADD ANY MORE ROWS TO THIS ARRAY YOU'LL NEED TO CHANGE
            ' THE 'WFITEMPROPERTYCOUNT' CONSTANT
            recWorkflowElementItemEdit.Index = "idxElementID"
            recWorkflowElementItemEdit.Seek ">=", .Fields("ID")
            
            If Not recWorkflowElementItemEdit.NoMatch Then
              Do While Not recWorkflowElementItemEdit.EOF
                If recWorkflowElementItemEdit.Fields("elementID") <> .Fields("ID") Then
                  Exit Do
                End If
              
                ReDim Preserve asItems(WFITEMPROPERTYCOUNT, UBound(asItems, 2) + 1)
                lngArraySize = UBound(asItems, 2)
          
                Select Case recWorkflowElementItemEdit.Fields("itemType")
                  Case giWFFORMITEM_BUTTON
                    sDescription = "Button - '" & recWorkflowElementItemEdit.Fields("caption") & "'"
                  Case giWFFORMITEM_DBVALUE
                    sDescription = "Database value - " & GetColumnName(recWorkflowElementItemEdit.Fields("DBColumnID"))
                  Case giWFFORMITEM_LABEL
                    sDescription = "Label - '" & recWorkflowElementItemEdit.Fields("caption") & "'"
                  Case giWFFORMITEM_INPUTVALUE_CHAR, _
                        giWFFORMITEM_INPUTVALUE_DROPDOWN, _
                        giWFFORMITEM_INPUTVALUE_LOGIC, _
                        giWFFORMITEM_INPUTVALUE_LOOKUP, _
                        giWFFORMITEM_INPUTVALUE_DATE, _
                        giWFFORMITEM_INPUTVALUE_NUMERIC, _
                        giWFFORMITEM_INPUTVALUE_OPTIONGROUP, _
                        giWFFORMITEM_INPUTVALUE_GRID, _
                        giWFFORMITEM_INPUTVALUE_FILEUPLOAD
                    sDescription = "Input value - " & recWorkflowElementItemEdit.Fields("Identifier")
                  Case giWFFORMITEM_WFVALUE, _
                    giWFFORMITEM_WFFILE
                    
                    sDescription = "Workflow value - " & recWorkflowElementItemEdit.Fields("WFFormIdentifier") & "." & recWorkflowElementItemEdit.Fields("WFValueIdentifier")
                  Case giWFFORMITEM_FORMATCODE
                    sDescription = "Formatting - " & FormatDescription(recWorkflowElementItemEdit.Fields("caption"))
                  Case giWFFORMITEM_CALC
                    sDescription = "Calculation - <" & GetExpressionName(recWorkflowElementItemEdit.Fields("calcID").value) & ">"
                  Case Else
                    sDescription = "<unknown>"
                End Select
                
                asItems(0, lngArraySize) = recWorkflowElementItemEdit.Fields("ID").value
                'JPD 20060919 Fault 11355
                'asItems(1, lngArraySize) = recWorkflowElementItemEdit.Fields("Identifier").Value
                asItems(1, lngArraySize) = sDescription
                asItems(2, lngArraySize) = recWorkflowElementItemEdit.Fields("itemType").value
                asItems(3, lngArraySize) = recWorkflowElementItemEdit.Fields("caption").value
                asItems(4, lngArraySize) = recWorkflowElementItemEdit.Fields("DBColumnID").value
                asItems(5, lngArraySize) = recWorkflowElementItemEdit.Fields("DBRecord").value
                asItems(6, lngArraySize) = recWorkflowElementItemEdit.Fields("InputType").value
                asItems(7, lngArraySize) = recWorkflowElementItemEdit.Fields("InputSize").value
                asItems(8, lngArraySize) = recWorkflowElementItemEdit.Fields("InputDecimals").value
                asItems(9, lngArraySize) = recWorkflowElementItemEdit.Fields("Identifier").value
                
                If CInt(asItems(6, lngArraySize)) = giEXPRVALUE_DATE Then
                  If Len(IIf(IsNull(recWorkflowElementItemEdit.Fields("InputDefault").value), "", recWorkflowElementItemEdit.Fields("InputDefault").value)) = 0 Then
                    asItems(10, lngArraySize) = ""
                  Else
                    asItems(10, lngArraySize) = objMisc.ConvertSQLDateToLocale(recWorkflowElementItemEdit.Fields("InputDefault").value)
                  End If
                ElseIf CInt(asItems(6, lngArraySize)) = giEXPRVALUE_NUMERIC Then
                  If Len(IIf(IsNull(recWorkflowElementItemEdit.Fields("InputDefault").value), "", recWorkflowElementItemEdit.Fields("InputDefault").value)) = 0 Then
                    asItems(10, lngArraySize) = "0"
                  Else
                    asItems(10, lngArraySize) = UI.ConvertNumberForDisplay(recWorkflowElementItemEdit.Fields("InputDefault").value)
                  End If
                Else
                  asItems(10, lngArraySize) = recWorkflowElementItemEdit.Fields("InputDefault").value
                End If
                
                asItems(11, lngArraySize) = recWorkflowElementItemEdit.Fields("WFFormIdentifier").value
                asItems(12, lngArraySize) = recWorkflowElementItemEdit.Fields("WFValueIdentifier").value
                
                If (wfElement.ElementType = elem_Email) Then
                  asItems(13, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("RecSelWebFormIdentifier").value), "", recWorkflowElementItemEdit.Fields("RecSelWebFormIdentifier").value)
                  asItems(14, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("RecSelIdentifier").value), "", recWorkflowElementItemEdit.Fields("RecSelIdentifier").value)
                End If
                
                If (wfElement.ElementType = elem_WebForm) Then
                  asItems(13, lngArraySize) = recWorkflowElementItemEdit.Fields("LeftCoord").value
                  asItems(14, lngArraySize) = recWorkflowElementItemEdit.Fields("TopCoord").value
                  asItems(15, lngArraySize) = recWorkflowElementItemEdit.Fields("Width").value
                  asItems(16, lngArraySize) = recWorkflowElementItemEdit.Fields("Height").value
                  asItems(17, lngArraySize) = recWorkflowElementItemEdit.Fields("BackColor").value
                  asItems(18, lngArraySize) = recWorkflowElementItemEdit.Fields("ForeColor").value
                  asItems(19, lngArraySize) = recWorkflowElementItemEdit.Fields("FontName").value
                  asItems(20, lngArraySize) = recWorkflowElementItemEdit.Fields("FontSize").value
                  asItems(21, lngArraySize) = recWorkflowElementItemEdit.Fields("FontBold").value
                  asItems(22, lngArraySize) = recWorkflowElementItemEdit.Fields("FontItalic").value
                  asItems(23, lngArraySize) = recWorkflowElementItemEdit.Fields("FontStrikeThru").value
                  asItems(24, lngArraySize) = recWorkflowElementItemEdit.Fields("FontUnderline").value
                  asItems(25, lngArraySize) = recWorkflowElementItemEdit.Fields("PictureID").value
                  asItems(26, lngArraySize) = CStr(IIf(recWorkflowElementItemEdit.Fields("PictureBorder").value, vbFixedSingle, vbBSNone))
                  asItems(27, lngArraySize) = recWorkflowElementItemEdit.Fields("Alignment").value
                  asItems(28, lngArraySize) = recWorkflowElementItemEdit.Fields("ZOrder").value
                  asItems(29, lngArraySize) = recWorkflowElementItemEdit.Fields("TabIndex").value
                  asItems(30, lngArraySize) = recWorkflowElementItemEdit.Fields("BackStyle").value
                  asItems(31, lngArraySize) = recWorkflowElementItemEdit.Fields("BackColorEven").value
                  asItems(32, lngArraySize) = recWorkflowElementItemEdit.Fields("BackColorOdd").value
                  asItems(33, lngArraySize) = recWorkflowElementItemEdit.Fields("ColumnHeaders").value
                  asItems(34, lngArraySize) = recWorkflowElementItemEdit.Fields("ForeColorEven").value
                  asItems(35, lngArraySize) = recWorkflowElementItemEdit.Fields("ForeColorOdd").value
                  asItems(36, lngArraySize) = recWorkflowElementItemEdit.Fields("HeaderBackColor").value
                  asItems(37, lngArraySize) = recWorkflowElementItemEdit.Fields("HeadFontName").value
                  asItems(38, lngArraySize) = recWorkflowElementItemEdit.Fields("HeadFontSize").value
                  asItems(39, lngArraySize) = recWorkflowElementItemEdit.Fields("HeadFontBold").value
                  asItems(40, lngArraySize) = recWorkflowElementItemEdit.Fields("HeadFontItalic").value
                  asItems(41, lngArraySize) = recWorkflowElementItemEdit.Fields("HeadFontStrikeThru").value
                  asItems(42, lngArraySize) = recWorkflowElementItemEdit.Fields("HeadFontUnderline").value
                  asItems(43, lngArraySize) = recWorkflowElementItemEdit.Fields("HeadLines").value
                  asItems(44, lngArraySize) = recWorkflowElementItemEdit.Fields("TableID").value
                  asItems(45, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("ForeColorHighlight").value), vbHighlightText, recWorkflowElementItemEdit.Fields("ForeColorHighlight").value)
                  asItems(46, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("BackColorHighlight").value), vbHighlight, recWorkflowElementItemEdit.Fields("BackColorHighlight").value)
                  asItems(47, lngArraySize) = vbNullString
                  
                  If (recWorkflowElementItemEdit.Fields("itemType") = giWFFORMITEM_INPUTVALUE_DROPDOWN) Or _
                    (recWorkflowElementItemEdit.Fields("itemType") = giWFFORMITEM_INPUTVALUE_OPTIONGROUP) Then

                    sControlValueList = vbNullString
                    
                    recWorkflowElementItemValuesEdit.Index = "idxItemID"
                    recWorkflowElementItemValuesEdit.Seek ">=", asItems(0, lngArraySize)

                    If Not recWorkflowElementItemValuesEdit.NoMatch Then
                      Do While Not recWorkflowElementItemValuesEdit.EOF
                        If recWorkflowElementItemValuesEdit.Fields("itemID") <> asItems(0, lngArraySize) Then
                          Exit Do
                        End If

                        sControlValueList = sControlValueList & recWorkflowElementItemValuesEdit.Fields("value") & vbTab
                        
                        recWorkflowElementItemValuesEdit.MoveNext
                      Loop
                    End If
                    
                    If Len(sControlValueList) > 0 Then
                      asItems(47, lngArraySize) = Left(sControlValueList, Len(sControlValueList) - 1)
                    End If
                  End If
                  
                  asItems(48, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("LookupTableID").value), 0, recWorkflowElementItemEdit.Fields("LookupTableID").value)
                  asItems(49, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("LookupColumnID").value), vbHighlightText, recWorkflowElementItemEdit.Fields("LookupColumnID").value)
                  asItems(50, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("RecordTableID").value), vbHighlightText, recWorkflowElementItemEdit.Fields("RecordTableID").value)
                  asItems(51, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("Orientation").value), 0, recWorkflowElementItemEdit.Fields("Orientation").value)
                  asItems(52, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("RecordOrderID").value), 0, recWorkflowElementItemEdit.Fields("RecordOrderID").value)
                  asItems(53, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("RecordFilterID").value), 0, recWorkflowElementItemEdit.Fields("RecordFilterID").value)
                  asItems(54, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("behaviour").value), WORKFLOWBUTTONACTION_SUBMIT, recWorkflowElementItemEdit.Fields("behaviour").value)
                  asItems(55, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("mandatory").value), False, recWorkflowElementItemEdit.Fields("mandatory").value)
                  asItems(57, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("captionType").value), giWFDATAVALUE_FIXED, recWorkflowElementItemEdit.Fields("captionType").value)
                  asItems(58, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("defaultValueType").value), giWFDATAVALUE_FIXED, recWorkflowElementItemEdit.Fields("defaultValueType").value)
                  asItems(59, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("VerticalOffsetBehaviour").value), 0, recWorkflowElementItemEdit.Fields("VerticalOffsetBehaviour").value)
                  asItems(60, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("HorizontalOffsetBehaviour").value), 0, recWorkflowElementItemEdit.Fields("HorizontalOffsetBehaviour").value)
                  asItems(61, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("VerticalOffset").value), 0, recWorkflowElementItemEdit.Fields("VerticalOffset").value)
                  asItems(62, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("HorizontalOffset").value), 0, recWorkflowElementItemEdit.Fields("HorizontalOffset").value)
                  asItems(63, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("HeightBehaviour").value), 0, recWorkflowElementItemEdit.Fields("HeightBehaviour").value)
                  asItems(64, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("WidthBehaviour").value), 0, recWorkflowElementItemEdit.Fields("WidthBehaviour").value)
                  asItems(65, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("PasswordType").value), False, recWorkflowElementItemEdit.Fields("PasswordType").value)
                  asItems(66, lngArraySize) = vbNullString
                  asItems(67, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("LookupFilterColumnID").value), 0, recWorkflowElementItemEdit.Fields("LookupFilterColumnID").value)
                  asItems(68, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("LookupFilterOperator").value), 0, recWorkflowElementItemEdit.Fields("LookupFilterOperator").value)
                  asItems(69, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("LookupFilterValue").value), "", recWorkflowElementItemEdit.Fields("LookupFilterValue").value)
                  asItems(80, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("LookupOrderID").value), 0, recWorkflowElementItemEdit.Fields("LookupOrderID").value)
                  asItems(81, lngArraySize) = recWorkflowElementItemEdit.Fields("HotSpotIdentifier").value
                  asItems(82, lngArraySize) = recWorkflowElementItemEdit.Fields("UseAsTargetIdentifier").value

                  If (recWorkflowElementItemEdit.Fields("itemType") = giWFFORMITEM_INPUTVALUE_FILEUPLOAD) Then

                    sControlValueList = vbNullString

                    recWorkflowElementItemValuesEdit.Index = "idxItemID"
                    recWorkflowElementItemValuesEdit.Seek ">=", asItems(0, lngArraySize)

                    If Not recWorkflowElementItemValuesEdit.NoMatch Then
                      Do While Not recWorkflowElementItemValuesEdit.EOF
                        If recWorkflowElementItemValuesEdit.Fields("itemID") <> asItems(0, lngArraySize) Then
                          Exit Do
                        End If

                        sControlValueList = sControlValueList & recWorkflowElementItemValuesEdit.Fields("value") & vbTab

                        recWorkflowElementItemValuesEdit.MoveNext
                      Loop
                    End If

                    If Len(sControlValueList) > 0 Then
                      asItems(66, lngArraySize) = Left(sControlValueList, Len(sControlValueList) - 1)
                    End If
                  End If
                End If
                
                asItems(56, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("calcID").value), 0, recWorkflowElementItemEdit.Fields("calcID").value)
                asItems(78, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("pageno").value), 0, recWorkflowElementItemEdit.Fields("pageno").value)
                asItems(79, lngArraySize) = IIf(IsNull(recWorkflowElementItemEdit.Fields("buttonstyle").value), 0, recWorkflowElementItemEdit.Fields("buttonstyle").value)
                
                recWorkflowElementItemEdit.MoveNext
              Loop
            End If
            
            wfElement.Items = asItems
          End If
          
          ReDim avColumns(0)
          
          If (wfElement.ElementType = elem_StoredData) Then
            ReDim avColumns(10, 0)
            ' Col 1 = Column Description
            ' Col 2 = Value Description
            ' Col 3 = Column ID
            ' Col 4 = Value Type
            ' Col 5 = Value
            ' Col 6 = WF Form Identifier
            ' Col 7 = WF Value Identifier
            ' Col 8 = DB Value Column ID
            ' Col 9 = DB Value Record
            ' Col 10 = CalcID

            recWorkflowElementColumnEdit.Index = "idxElementID"
            recWorkflowElementColumnEdit.Seek ">=", .Fields("ID")

            If Not recWorkflowElementColumnEdit.NoMatch Then
              Do While Not recWorkflowElementColumnEdit.EOF
                If recWorkflowElementColumnEdit.Fields("elementID") <> .Fields("ID") Then
                  Exit Do
                End If

                ReDim Preserve avColumns(10, UBound(avColumns, 2) + 1)

                sDescription = GetColumnName(recWorkflowElementColumnEdit.Fields("columnID"), True)
                avColumns(1, UBound(avColumns, 2)) = sDescription

                sDescription = "<unknown>"
                Select Case recWorkflowElementColumnEdit.Fields("ValueType")
                  Case giWFDATAVALUE_FIXED
                    sDescription = "Fixed value - " & recWorkflowElementColumnEdit.Fields("Value")
                  
                    If (GetColumnDataType(recWorkflowElementColumnEdit.Fields("columnID")) = dtTIMESTAMP) _
                      And UCase(recWorkflowElementColumnEdit.Fields("Value")) <> "NULL" Then
                      sDescription = "Fixed value - " & objMisc.ConvertSQLDateToLocale(recWorkflowElementColumnEdit.Fields("Value"))
                    End If

                  Case giWFDATAVALUE_WFVALUE
                    sDescription = "Workflow value - " & recWorkflowElementColumnEdit.Fields("WFFormIdentifier") & "." & recWorkflowElementColumnEdit.Fields("WFValueIdentifier")
                  
                  Case giWFDATAVALUE_DBVALUE
                    sDescription = "Database value - " & GetColumnName(recWorkflowElementColumnEdit.Fields("DBColumnID"))
                  
                  Case giWFDATAVALUE_CALC
                    sTemp = GetExpressionName(recWorkflowElementColumnEdit.Fields("CalcID"))
                    If Len(Trim(sTemp)) = 0 Then
                      sTemp = "<unknown>"
                    Else
                      sTemp = "<" & sTemp & ">"
                    End If
                    sDescription = "Calculation - " & sTemp
                End Select
                avColumns(2, UBound(avColumns, 2)) = sDescription
                
                avColumns(3, UBound(avColumns, 2)) = recWorkflowElementColumnEdit.Fields("columnID")
                avColumns(4, UBound(avColumns, 2)) = recWorkflowElementColumnEdit.Fields("ValueType")
                avColumns(5, UBound(avColumns, 2)) = recWorkflowElementColumnEdit.Fields("Value")
                avColumns(6, UBound(avColumns, 2)) = recWorkflowElementColumnEdit.Fields("WFFormIdentifier")
                avColumns(7, UBound(avColumns, 2)) = recWorkflowElementColumnEdit.Fields("WFValueIdentifier")
                avColumns(8, UBound(avColumns, 2)) = recWorkflowElementColumnEdit.Fields("DBColumnID")
                avColumns(9, UBound(avColumns, 2)) = recWorkflowElementColumnEdit.Fields("DBRecord")
                avColumns(10, UBound(avColumns, 2)) = recWorkflowElementColumnEdit.Fields("CalcID")

                recWorkflowElementColumnEdit.MoveNext
              Loop
            End If
            
            wfElement.DataColumns = avColumns
          End If

          ReDim asValidations(0)
          
          If (wfElement.ElementType = elem_WebForm) Then
            ReDim asValidations(4, 0)
            ' Col 1 = Expr ID
            ' Col 2 = Type (0=Error, 1=Warning)
            ' Col 3 = Message

            recWorkflowElementValidationEdit.Index = "idxElementID"
            recWorkflowElementValidationEdit.Seek ">=", .Fields("ID")

            If Not recWorkflowElementValidationEdit.NoMatch Then
              Do While Not recWorkflowElementValidationEdit.EOF
                If recWorkflowElementValidationEdit.Fields("elementID") <> .Fields("ID") Then
                  Exit Do
                End If

                ReDim Preserve asValidations(4, UBound(asValidations, 2) + 1)

                asValidations(1, UBound(asValidations, 2)) = recWorkflowElementValidationEdit.Fields("exprID")
                asValidations(2, UBound(asValidations, 2)) = recWorkflowElementValidationEdit.Fields("type")
                asValidations(3, UBound(asValidations, 2)) = recWorkflowElementValidationEdit.Fields("message")

                recWorkflowElementValidationEdit.MoveNext
              Loop
            End If
          
            wfElement.Validations = asValidations
          End If

          ReDim Preserve malngIndexDirectory(3, UBound(malngIndexDirectory, 2) + 1)
          malngIndexDirectory(1, UBound(malngIndexDirectory, 2)) = wfElement.ControlIndex
          malngIndexDirectory(2, UBound(malngIndexDirectory, 2)) = .Fields("ID")
          
          wfElement.Highlighted = False
        
          sngMaxX = IIf(sngMaxX >= wfElement.Left + wfElement.Width, sngMaxX, wfElement.Left + wfElement.Width)
          sngMaxY = IIf(sngMaxY >= wfElement.Top + wfElement.Height, sngMaxY, wfElement.Top + wfElement.Height)
        End If
        
        ' AE20080609 Fault #13202
        wfElement.Visible = True
        Set wfElement = Nothing

        .MoveNext
      Loop       ' Indent 04 - end
    End If       ' Indent 03 - end
  End With       ' Indent 02 - end

  ' Update the connector pair ID values.
  For iLoop = 1 To UBound(malngIndexDirectory, 2)
    With recWorkflowElementEdit
      .Index = "idxElementID"
      .Seek "=", malngIndexDirectory(2, iLoop)

      If Not .NoMatch Then
        For iLoop2 = 1 To UBound(malngIndexDirectory, 2)
          If malngIndexDirectory(2, iLoop2) = .Fields("connectionPairID") Then

            Set wfElement = mcolwfElements(CStr(malngIndexDirectory(1, iLoop)))
            wfElement.ConnectorPairIndex = malngIndexDirectory(1, iLoop2)
            Set wfElement = Nothing
            
            Exit For
          End If
        Next iLoop2
      End If
    End With
  Next iLoop

  ' Load each link.
  With recWorkflowLinkEdit
    .Index = "idxWorkflowID"
    .Seek ">=", mlngWorkflowID

    If Not .NoMatch Then

      ' Add elements to the form
      Do While Not .EOF
        If .Fields("workflowID") <> mlngWorkflowID Then
          Exit Do
        End If

        Load ASRWFLink1(ASRWFLink1.UBound + 1)
        Set wfLink = ASRWFLink1(ASRWFLink1.UBound)

        fStartIndexDone = False
        fEndIndexDone = False
        
        For iLoop = 1 To UBound(malngIndexDirectory, 2)
          If malngIndexDirectory(2, iLoop) = .Fields("startElementID") Then
            wfLink.StartElementIndex = malngIndexDirectory(1, iLoop)
            iStartWFElementType = malngIndexDirectory(3, iLoop)
            fStartIndexDone = True
          End If
          
          If malngIndexDirectory(2, iLoop) = .Fields("endElementID") Then
            wfLink.EndElementIndex = malngIndexDirectory(1, iLoop)
            iEndWFElementType = malngIndexDirectory(3, iLoop)
            fEndIndexDone = True
          End If
          
          If fStartIndexDone And fEndIndexDone Then
            Exit For
          End If
        Next iLoop
        
        wfLink.StartOutboundFlowCode = .Fields("startOutboundFlowCode")

        ' Get the array of outbound flow information from the start element.
        ' Column 1 = Tag (see enums, or -1 if there's only a single outbound flow)
        ' Column 2 = Direction
        ' Column 3 = XOffset
        ' Column 4 = YOffset
        ' Column 5 = Maximum
        ' Column 6 = Minimum
        ' Column 7 = Description
        Set wfElement = mcolwfElements(CStr(wfLink.StartElementIndex))
        avOutboundFlowInfo = wfElement.OutboundFlows_Information
        Set wfElement = Nothing
        
        If wfLink.StartOutboundFlowCode < 0 Then
          iOutboundFlowIndex = 1
        Else
          For iLoop = 1 To UBound(avOutboundFlowInfo, 2)
            If avOutboundFlowInfo(1, iLoop) = wfLink.StartOutboundFlowCode Then
              iOutboundFlowIndex = iLoop
              Exit For
            End If
          Next iLoop
        End If

        wfLink.StartDirection = avOutboundFlowInfo(2, iOutboundFlowIndex)
        
        Set wfElement = mcolwfElements(CStr(wfLink.EndElementIndex))
        wfLink.EndDirection = wfElement.InboundFlow_Direction
        Set wfElement = Nothing
    
        FormatLink wfLink
        
        wfLink.Highlighted = False
        wfLink.Visible = True
        wfLink.ZOrder 1
        
        sngMaxX = IIf(sngMaxX >= wfLink.Left + wfLink.Width, sngMaxX, wfLink.Left + wfLink.Width)
        sngMaxY = IIf(sngMaxY >= wfLink.Top + wfLink.Height, sngMaxY, wfLink.Top + wfLink.Height)
        
        .MoveNext
      Loop
    End If
  End With

  ' Resize the definition picturebox to fit the loaded elements.
  picDefinition.Width = sngMaxX + 200
  picDefinition.Height = sngMaxY + 200
  
  ' Resize the form to fit the definition picturebox.
  If picDefinition.Width >= Screen.Width Or _
    picDefinition.Height > Screen.Height Then
    
    Me.WindowState = vbMaximized
'''  Else
'''    Me.ScaleWidth = picContainer.Left + picDefinition.Width + scrollVertical.Width
'''                   Me.ScaleHeight = picContainer.Top + picDefinition.Height + scrollHorizontal.Height + fraButtons.Height + 5000
  End If
  
  SetScrollBarValues
  
TidyUpAndExit:
  Set objMisc = Nothing

  ' Unlock the window refreshing.
  UI.UnlockWindow

  ' Reset the screen moousepointer.
  Screen.MousePointer = vbDefault

  LoadElementsAndLinks = fLoadOk
  Exit Function

ErrorTrap:
  fLoadOk = False
  MsgBox "Error loading Workflow." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function

Public Property Get IsNew() As Boolean
  ' Return the 'new workflow' flag.
  IsNew = mfNewWorkflow
  
End Property


Public Property Let IsNew(pfNewValue As Boolean)
  ' Set the 'new workflow' flag.
  mfNewWorkflow = pfNewValue
  IsChanged = True
  
End Property


Public Property Get WorkflowName() As String
  ' Return the workflow name
  WorkflowName = msWorkflowName
  
End Property

Public Property Get IsChanged() As Boolean
  ' Return the 'workflow changed' flag.
  IsChanged = mfChanged
  
End Property

Public Property Let IsChanged(pfNewValue As Boolean)
  mfChanged = pfNewValue
  mfPerge = pfNewValue
cmdOK.Enabled = mfChanged
End Property

Public Sub SetChanged(pfPerge As Boolean)
  mfChanged = True
  mfPerge = (mfPerge Or pfPerge)
  cmdOK.Enabled = mfChanged
End Sub

Public Function IsUniqueIdentifier(psIdentifier As String, plngIgnoreElementIndex As Long) As Boolean
  ' Return true if the given Identifier string is unique in this workflow.
  ' Ignore the given element.
  Dim wfElement As VB.Control
  Dim fUnique As Boolean
  
  fUnique = True
  
  For Each wfElement In mcolwfElements
    If (wfElement.Visible) _
      And ElementHasIdentifier(wfElement) _
      And (wfElement.ControlIndex <> plngIgnoreElementIndex) _
      And (UCase(Trim(wfElement.Identifier)) = UCase(Trim(psIdentifier))) Then
      
      fUnique = False
      Exit For
    End If
  Next wfElement
  Set wfElement = Nothing
  
  IsUniqueIdentifier = fUnique
  
End Function
Public Sub AllElements(paWFAllElements As Variant)
  ' Return an array of all elements in the definition.
  Dim wfTempElement As VB.Control
  Dim fElementOK As Boolean
  Dim iLoop As Integer
  
  ReDim paWFAllElements(0)
  
  For Each wfTempElement In mcolwfElements
    With wfTempElement
      fElementOK = .Visible
      
      If (Not fElementOK) Then
        fElementOK = (.ControlIndex > 0)
  
        If fElementOK Then
          If (miLastActionFlag = giACTION_DELETECONTROLS) Then
            For iLoop = 1 To UBound(mactlUndoControls)
              If IsWorkflowElement(mactlUndoControls(iLoop)) Then
                If mactlUndoControls(iLoop).ControlIndex = .ControlIndex Then
                  fElementOK = False
                  Exit For
                End If
              End If
            Next iLoop
          End If
        End If

        If fElementOK Then
          For iLoop = 1 To UBound(mactlClipboardControls)
            If IsWorkflowElement(mactlClipboardControls(iLoop)) Then
              If mactlClipboardControls(iLoop).ControlIndex = .ControlIndex Then
                fElementOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
      End If
      
      If fElementOK Then
        ReDim Preserve paWFAllElements(UBound(paWFAllElements) + 1)
        Set paWFAllElements(UBound(paWFAllElements)) = wfTempElement
      End If
    End With
  Next wfTempElement
  Set wfTempElement = Nothing
  
End Sub

Public Sub SucceedingElements(pwfElement As VB.Control, paWFSucceedingElements As Variant)
  ' Return an array of the elements that succeed (follow) the given element.
  Dim wfLink As COAWF_Link
  Dim fFound As Boolean
  Dim iLoop As Integer
  Dim fLinkOK As Boolean
    
  For Each wfLink In ASRWFLink1
    If (wfLink.StartElementIndex = pwfElement.ControlIndex) Then

      fLinkOK = wfLink.Visible
      If (Not fLinkOK) Then
        ' Link might not be .visible but still valid
        ' if this method is called from the web form designer.
        fLinkOK = True

        If (miLastActionFlag = giACTION_DELETECONTROLS) Then
          For iLoop = 1 To UBound(mactlUndoControls)
            If TypeOf mactlUndoControls(iLoop) Is COAWF_Link Then
              If mactlUndoControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If

        If fLinkOK Then
          If (miLastActionFlag = giACTION_SWAPCONTROL) Then
            If UBound(mactlUndoControls) >= 1 Then
              If TypeOf mactlUndoControls(1) Is COAWF_Link Then
                If mactlUndoControls(1).Index = wfLink.Index Then
                  fLinkOK = False
                End If
              End If
            End If
          End If
        End If

        If fLinkOK Then
          For iLoop = 1 To UBound(mactlClipboardControls)
            If TypeOf mactlClipboardControls(iLoop) Is COAWF_Link Then
              If mactlClipboardControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
      End If

      If fLinkOK Then
        fFound = False

        For iLoop = 1 To UBound(paWFSucceedingElements)
          If paWFSucceedingElements(iLoop).ControlIndex = wfLink.EndElementIndex Then
            fFound = True
            Exit For
          End If
        Next iLoop

        If Not fFound Then
          ReDim Preserve paWFSucceedingElements(UBound(paWFSucceedingElements) + 1)
          Set paWFSucceedingElements(UBound(paWFSucceedingElements)) = mcolwfElements(CStr(wfLink.EndElementIndex))

          If mcolwfElements(CStr(wfLink.EndElementIndex)).ElementType = elem_Connector1 Then
            ReDim Preserve paWFSucceedingElements(UBound(paWFSucceedingElements) + 1)
            Set paWFSucceedingElements(UBound(paWFSucceedingElements)) = mcolwfElements(CStr(mcolwfElements(CStr(wfLink.EndElementIndex)).ConnectorPairIndex))

            SucceedingElements mcolwfElements(CStr(mcolwfElements(CStr(wfLink.EndElementIndex)).ConnectorPairIndex)), paWFSucceedingElements
          Else
            SucceedingElements mcolwfElements(CStr(wfLink.EndElementIndex)), paWFSucceedingElements
          End If
        End If
      End If
    End If
  Next wfLink
  Set wfLink = Nothing
 
End Sub

Public Sub PrecedingElements(pwfElement As VB.Control, paWFPrecedingElements As Variant)
  ' Return an array of the elements that precede the given element.
  Dim wfLink As COAWF_Link
  Dim fFound As Boolean
  Dim iLoop As Integer
  Dim fLinkOK As Boolean
  
  For Each wfLink In ASRWFLink1
    If (wfLink.EndElementIndex = pwfElement.ControlIndex) Then
    
      fLinkOK = wfLink.Visible
      If (Not fLinkOK) Then
        ' Link might not be .visible but still valid
        ' if this method is called from the web form designer.
        fLinkOK = True
          
        If (miLastActionFlag = giACTION_DELETECONTROLS) Then
          For iLoop = 1 To UBound(mactlUndoControls)
            If TypeOf mactlUndoControls(iLoop) Is COAWF_Link Then
              If mactlUndoControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
        
        If fLinkOK Then
          If (miLastActionFlag = giACTION_SWAPCONTROL) Then
            If UBound(mactlUndoControls) >= 1 Then
              If TypeOf mactlUndoControls(1) Is COAWF_Link Then
                If mactlUndoControls(1).Index = wfLink.Index Then
                  fLinkOK = False
                End If
              End If
            End If
          End If
        End If
        
        'JPD 20060719 Fault 11339
        If fLinkOK Then
          For iLoop = 1 To UBound(mactlClipboardControls)
            If TypeOf mactlClipboardControls(iLoop) Is COAWF_Link Then
              If mactlClipboardControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
      End If
      
      If fLinkOK Then
        fFound = False
        
        For iLoop = 1 To UBound(paWFPrecedingElements)
          If paWFPrecedingElements(iLoop).ControlIndex = wfLink.StartElementIndex Then
          
            ' Element is looped back to so can still be referred to.
            If iLoop <> 1 Then
              fFound = True
              Exit For
            End If
          End If
        Next iLoop
        
        If Not fFound Then
          ReDim Preserve paWFPrecedingElements(UBound(paWFPrecedingElements) + 1)
          Set paWFPrecedingElements(UBound(paWFPrecedingElements)) = mcolwfElements(CStr(wfLink.StartElementIndex))
          
          If mcolwfElements(CStr(wfLink.StartElementIndex)).ElementType = elem_Connector2 Then
            ReDim Preserve paWFPrecedingElements(UBound(paWFPrecedingElements) + 1)
            Set paWFPrecedingElements(UBound(paWFPrecedingElements)) = mcolwfElements(CStr(mcolwfElements(CStr(wfLink.StartElementIndex)).ConnectorPairIndex))
          
            PrecedingElements mcolwfElements(CStr(mcolwfElements(CStr(wfLink.StartElementIndex)).ConnectorPairIndex)), paWFPrecedingElements
          Else
            PrecedingElements mcolwfElements(CStr(wfLink.StartElementIndex)), paWFPrecedingElements
          End If
        End If
      End If
    End If
  Next wfLink
  Set wfLink = Nothing
  
End Sub

Private Function CycliclyValid(pwfElement As VB.Control, paWFPrecedingElements As Variant) As Boolean
  ' Check if the first element (index 1) in paWFPrecedingElements is in a flow loop.
  ' If it is, check that the loop contains a user action (ie. a Web Form element)
  Dim wfLink As COAWF_Link
  Dim wfElement As VB.Control
  Dim fFound As Boolean
  Dim iLoop As Integer
  Dim fLinkOK As Boolean
  Dim fCyclic As Boolean
  Dim fValid As Boolean
  Dim iWebFormCount As Integer
  
  fValid = True

  For Each wfLink In ASRWFLink1
    If (wfLink.EndElementIndex = pwfElement.ControlIndex) Then

      fLinkOK = wfLink.Visible
      If (Not fLinkOK) Then
        ' Link might not be .visible but still valid
        ' if this method is called from the web form designer.
        fLinkOK = True

        If (miLastActionFlag = giACTION_DELETECONTROLS) Then
          For iLoop = 1 To UBound(mactlUndoControls)
            If TypeOf mactlUndoControls(iLoop) Is COAWF_Link Then
              If mactlUndoControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If

        If fLinkOK Then
          If (miLastActionFlag = giACTION_SWAPCONTROL) Then
            If UBound(mactlUndoControls) >= 1 Then
              If TypeOf mactlUndoControls(1) Is COAWF_Link Then
                If mactlUndoControls(1).Index = wfLink.Index Then
                  fLinkOK = False
                End If
              End If
            End If
          End If
        End If

        'JPD 20060719 Fault 11339
        If fLinkOK Then
          For iLoop = 1 To UBound(mactlClipboardControls)
            If TypeOf mactlClipboardControls(iLoop) Is COAWF_Link Then
              If mactlClipboardControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
      End If

      If fLinkOK Then
        fFound = False
        fCyclic = False
        iWebFormCount = 0

        For iLoop = 1 To UBound(paWFPrecedingElements)
          If paWFPrecedingElements(iLoop).ControlIndex = wfLink.StartElementIndex Then
            fFound = True
            If (iLoop = 1) Then
              fCyclic = True
            End If
          End If
        
          If paWFPrecedingElements(iLoop).ElementType = elem_WebForm Then
            iWebFormCount = iWebFormCount + 1
          End If
        Next iLoop

        If fFound Then
          If fCyclic And (iWebFormCount = 0) Then
            fValid = False
          End If
        ElseIf iWebFormCount = 0 Then
          ReDim Preserve paWFPrecedingElements(UBound(paWFPrecedingElements) + 1)
          Set paWFPrecedingElements(UBound(paWFPrecedingElements)) = mcolwfElements(CStr(wfLink.StartElementIndex))

          If mcolwfElements(CStr(wfLink.StartElementIndex)).ElementType = elem_Connector2 Then
            ReDim Preserve paWFPrecedingElements(UBound(paWFPrecedingElements) + 1)
            Set paWFPrecedingElements(UBound(paWFPrecedingElements)) = mcolwfElements(CStr(mcolwfElements(CStr(wfLink.StartElementIndex)).ConnectorPairIndex))
            
            fValid = CycliclyValid(mcolwfElements(CStr(mcolwfElements(CStr(wfLink.StartElementIndex)).ConnectorPairIndex)), paWFPrecedingElements)
            
            ReDim Preserve paWFPrecedingElements(UBound(paWFPrecedingElements) - 1)
          Else
            fValid = CycliclyValid(mcolwfElements(CStr(wfLink.StartElementIndex)), paWFPrecedingElements)
          End If

          ReDim Preserve paWFPrecedingElements(UBound(paWFPrecedingElements) - 1)
        Else
          fValid = True
        End If
      End If
    End If
    
    If Not fValid Then
      Exit For
    End If
  Next wfLink
  Set wfLink = Nothing

  CycliclyValid = fValid
  
End Function




Private Sub ImmediatelyPrecedingElements(pwfElement As VB.Control, paWFPrecedingElements As Variant)
  ' Return an array of the elements that IMMEDIATLEY precede the given element.
  ' NB. Connector and Or elements do go into the array, as do their IMMEDIATE predecessors.
  Dim wfLink As COAWF_Link
  Dim fFound As Boolean
  Dim iLoop As Integer
  Dim fLinkOK As Boolean

  For Each wfLink In ASRWFLink1
    If (wfLink.EndElementIndex = pwfElement.ControlIndex) Then

      fLinkOK = wfLink.Visible
      If (Not fLinkOK) Then
        ' Link might not be .visible but still valid
        ' if this method is called from the web form designer.
        fLinkOK = True
          
        If (miLastActionFlag = giACTION_DELETECONTROLS) Then
          For iLoop = 1 To UBound(mactlUndoControls)
            If TypeOf mactlUndoControls(iLoop) Is COAWF_Link Then
              If mactlUndoControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
         
        If fLinkOK Then
          If (miLastActionFlag = giACTION_SWAPCONTROL) Then
            If UBound(mactlUndoControls) >= 1 Then
              If TypeOf mactlUndoControls(1) Is COAWF_Link Then
                If mactlUndoControls(1).Index = wfLink.Index Then
                  fLinkOK = False
                End If
              End If
            End If
          End If
        End If
        
        'JPD 20060719 Fault 11339
        If fLinkOK Then
          For iLoop = 1 To UBound(mactlClipboardControls)
            If TypeOf mactlClipboardControls(iLoop) Is COAWF_Link Then
              If mactlClipboardControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
      End If

      If fLinkOK Then
        fFound = False

        For iLoop = 1 To UBound(paWFPrecedingElements)
          If paWFPrecedingElements(iLoop).ControlIndex = wfLink.StartElementIndex Then
            fFound = True
            Exit For
          End If
        Next iLoop

        If Not fFound Then
          ReDim Preserve paWFPrecedingElements(UBound(paWFPrecedingElements) + 1)
          Set paWFPrecedingElements(UBound(paWFPrecedingElements)) = mcolwfElements(CStr(wfLink.StartElementIndex))

          If mcolwfElements(CStr(wfLink.StartElementIndex)).ElementType = elem_Connector2 Then
            ReDim Preserve paWFPrecedingElements(UBound(paWFPrecedingElements) + 1)
            Set paWFPrecedingElements(UBound(paWFPrecedingElements)) = mcolwfElements(CStr(mcolwfElements(CStr(wfLink.StartElementIndex)).ConnectorPairIndex))

            ImmediatelyPrecedingElements mcolwfElements(CStr(mcolwfElements(CStr(wfLink.StartElementIndex)).ConnectorPairIndex)), paWFPrecedingElements
          
          ElseIf (mcolwfElements(CStr(wfLink.StartElementIndex)).ElementType = elem_Or) _
            Or ((mcolwfElements(CStr(wfLink.StartElementIndex)).ElementType = elem_StoredData) And (glngSQLVersion >= 9)) _
            Or (mcolwfElements(CStr(wfLink.StartElementIndex)).ElementType = elem_Decision) Then
            
            ImmediatelyPrecedingElements mcolwfElements(CStr(wfLink.StartElementIndex)), paWFPrecedingElements
          End If
        End If
      End If
    End If
  Next wfLink
  Set wfLink = Nothing
  
End Sub

Private Sub ImmediatelySucceedingElements(pwfElement As VB.Control, _
  paWFSucceedingElements As Variant, _
  pfIncludeDecisionSuccessors As Boolean, _
  Optional pvStartOutBoundFlowCode As Variant)
  
  ' Return an array of the elements that IMMEDIATLEY succeed the given element.
  ' NB. Connector and Or elements go into the array, as do their IMMEDIATE successors.
  Dim wfLink As COAWF_Link
  Dim fFound As Boolean
  Dim iLoop As Integer
  Dim fLinkOK As Boolean

  For Each wfLink In ASRWFLink1
    If (wfLink.StartElementIndex = pwfElement.ControlIndex) Then

      fLinkOK = wfLink.Visible
      If (Not fLinkOK) Then
        ' Link might not be .visible but still valid
        ' if this method is called from the web form designer.
        fLinkOK = True
          
        If (miLastActionFlag = giACTION_DELETECONTROLS) Then
          For iLoop = 1 To UBound(mactlUndoControls)
            If TypeOf mactlUndoControls(iLoop) Is COAWF_Link Then
              If mactlUndoControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
        
        If fLinkOK Then
          If (miLastActionFlag = giACTION_SWAPCONTROL) Then
            If UBound(mactlUndoControls) >= 1 Then
              If TypeOf mactlUndoControls(1) Is COAWF_Link Then
                If mactlUndoControls(1).Index = wfLink.Index Then
                  fLinkOK = False
                End If
              End If
            End If
          End If
        End If
        
        'JPD 20060719 Fault 11339
        If fLinkOK Then
          For iLoop = 1 To UBound(mactlClipboardControls)
            If TypeOf mactlClipboardControls(iLoop) Is COAWF_Link Then
              If mactlClipboardControls(iLoop).Index = wfLink.Index Then
                fLinkOK = False
                Exit For
              End If
            End If
          Next iLoop
        End If
      End If

      If fLinkOK And Not IsMissing(pvStartOutBoundFlowCode) Then
        ' Check if the link is for the given outboundflow (used for Decision elements)
        fLinkOK = (wfLink.StartOutboundFlowCode = CInt(pvStartOutBoundFlowCode))
        
        If (Not fLinkOK) And (CInt(pvStartOutBoundFlowCode) = 0) Then
          fLinkOK = (wfLink.StartOutboundFlowCode = -1)
        End If
      End If

      If fLinkOK Then
        fFound = False

        For iLoop = 1 To UBound(paWFSucceedingElements)
          If paWFSucceedingElements(iLoop).ControlIndex = wfLink.EndElementIndex Then
            fFound = True
            Exit For
          End If
        Next iLoop

        If Not fFound Then
          ReDim Preserve paWFSucceedingElements(UBound(paWFSucceedingElements) + 1)
          Set paWFSucceedingElements(UBound(paWFSucceedingElements)) = mcolwfElements(CStr(wfLink.EndElementIndex))

          If mcolwfElements(CStr(wfLink.EndElementIndex)).ElementType = elem_Connector1 Then
            ReDim Preserve paWFSucceedingElements(UBound(paWFSucceedingElements) + 1)
            Set paWFSucceedingElements(UBound(paWFSucceedingElements)) = mcolwfElements(CStr(mcolwfElements(CStr(wfLink.EndElementIndex)).ConnectorPairIndex))

            ImmediatelySucceedingElements _
              mcolwfElements(CStr(mcolwfElements(CStr(wfLink.EndElementIndex)).ConnectorPairIndex)), _
              paWFSucceedingElements, _
              pfIncludeDecisionSuccessors
              
          ElseIf (mcolwfElements(CStr(wfLink.EndElementIndex)).ElementType = elem_Or) _
            Or (pfIncludeDecisionSuccessors And (mcolwfElements(CStr(wfLink.EndElementIndex)).ElementType = elem_StoredData) And (glngSQLVersion >= 9)) _
            Or (pfIncludeDecisionSuccessors And (mcolwfElements(CStr(wfLink.EndElementIndex)).ElementType = elem_Decision)) Then
            
            ImmediatelySucceedingElements _
              mcolwfElements(CStr(wfLink.EndElementIndex)), _
              paWFSucceedingElements, _
              pfIncludeDecisionSuccessors
          End If
        End If
      End If
    End If
  Next wfLink
  Set wfLink = Nothing
  
End Sub





Private Function InElementAddMode() As Boolean
  Dim fInElementAddMode As Boolean
  Dim objTool As ActiveBarLibraryCtl.Tool
  
  fInElementAddMode = False
  
  For Each objTool In abMenu.Tools
    If objTool.Checked Then
      fInElementAddMode = True
      Exit For
    End If
  Next objTool
  Set objTool = Nothing
  
  InElementAddMode = fInElementAddMode

End Function

Private Sub ResizeCanvas()
  Dim wfTempElement As VB.Control
  Dim wfTempLink As COAWF_Link
  Dim sngWidthResized As Single
  Dim sngHeightResized As Single
    
  sngWidthResized = 0
  sngHeightResized = 0
     
  For Each wfTempElement In mcolwfElements
    With wfTempElement
      If .Visible Then
        If (.Left < 500) And (sngWidthResized > (.Left - 500)) Then
          sngWidthResized = .Left - 500
        End If
        ' AE20080613 Fault #13216
        'If ((.Left + .Width) > picDefinition.Width) And (sngWidthResized < (.Left + .Width)) Then
        If ((.Left + .Width + 500) > picDefinition.Width) And (sngWidthResized < (.Left + .Width - picDefinition.Width + 500)) Then
          sngWidthResized = .Left + .Width - picDefinition.Width + 500
        End If
        
        If (.Top < 500) And (sngHeightResized > (.Top - 500)) Then
          sngHeightResized = .Top - 500
        End If
        ' AE20080613 Fault #13216
        'If ((.Top + .Height) > picDefinition.Height) And (sngHeightResized < (.Top + .Height)) Then
        If ((.Top + .Height + 500) > picDefinition.Height) And (sngHeightResized < (.Top + .Height - picDefinition.Height + 500)) Then
          sngHeightResized = .Top + .Height - picDefinition.Height + 500
        End If
      End If
    End With
  Next wfTempElement
  Set wfTempElement = Nothing
              
  If sngWidthResized < 0 Then
    sngWidthResized = sngWidthResized
    picDefinition.Width = picDefinition.Width - sngWidthResized
    picDefinition.Left = sngWidthResized
  End If
  If sngWidthResized > 0 Then
    sngWidthResized = sngWidthResized
    picDefinition.Width = picDefinition.Width + sngWidthResized
  End If
  
  If sngHeightResized < 0 Then
    sngHeightResized = sngHeightResized
    picDefinition.Height = picDefinition.Height - sngHeightResized
    picDefinition.Top = sngHeightResized
  End If
  If sngHeightResized > 0 Then
    sngHeightResized = sngHeightResized
    picDefinition.Height = picDefinition.Height + sngHeightResized
  End If
    
  For Each wfTempElement In mcolwfElements
    With wfTempElement
      If .Visible Then
        If (sngWidthResized < 0) Then
          .Left = .Left - sngWidthResized
        End If
        
        If (sngHeightResized < 0) Then
          .Top = .Top - sngHeightResized
        End If
      End If
    End With
  Next wfTempElement
  Set wfTempElement = Nothing
  
  For Each wfTempLink In ASRWFLink1
    With wfTempLink
      If .Visible Then
        If (sngWidthResized < 0) Then
          .Left = .Left - sngWidthResized
        End If
        
        If (sngHeightResized < 0) Then
          .Top = .Top - sngHeightResized
        End If
      End If
    End With
  Next wfTempLink
  Set wfTempLink = Nothing
    
  If (sngWidthResized <> 0) Or (sngHeightResized <> 0) Then
    SetScrollBarValues

    If sngWidthResized < 0 Then
      scrollHorizontal.value = -picDefinition.Left / mdblHorizontalScrollRatio
    End If
    If sngHeightResized < 0 Then
      scrollVertical.value = -picDefinition.Top / mdblVerticalScrollRatio
    End If
  End If

End Sub

Private Sub SetLastActionFlag(piLastActionFlag As UndoActionFlags)
  Dim iIndex As Integer
  Dim ctlControl As VB.Control
  Dim iOldLastActionFlag As UndoActionFlags
  
  iOldLastActionFlag = miLastActionFlag
  miLastActionFlag = piLastActionFlag
  
  RefreshMenu

  If (iOldLastActionFlag = giACTION_DELETECONTROLS) _
    Or ((piLastActionFlag = giACTION_NOACTION) And (Not mblnLoading)) Then
    
    For iIndex = 1 To UBound(mactlUndoControls)
      Set ctlControl = mactlUndoControls(iIndex)
      If Not ctlControl Is Nothing Then
        ' AE20080428 Fault #13136
        'Set mcolwfElements(CStr(ctlControl.ControlIndex)) = Nothing
        
        If IsWorkflowElement(ctlControl) Then
          mcolwfElements.Remove CStr(ctlControl.ControlIndex)
        End If
        
        If ctlControl.Highlighted Then
          If IsWorkflowElement(ctlControl) Then
            mcolwfSelectedElements.Remove CStr(ctlControl.ControlIndex)
          ElseIf TypeOf ctlControl Is COAWF_Link Then
            mcolwfSelectedLinks.Remove CStr(ctlControl.Index)
          End If
        End If
        
        UnLoad ctlControl
      End If
      
      ' Disassociate object variables.
      Set ctlControl = Nothing
    Next iIndex
  
  ElseIf (iOldLastActionFlag = giACTION_SWAPCONTROL) Then
    If UBound(mactlUndoControls) >= 1 Then
      Set ctlControl = mactlUndoControls(1)
      If Not ctlControl Is Nothing Then
        ' AE20080428 Fault #13136
        'Set mcolwfElements(CStr(ctlControl.ControlIndex)) = Nothing
        
        If IsWorkflowElement(ctlControl) Then
          mcolwfElements.Remove CStr(ctlControl.ControlIndex)
        End If
        
        If ctlControl.Highlighted Then
          If IsWorkflowElement(ctlControl) Then
            mcolwfSelectedElements.Remove CStr(ctlControl.ControlIndex)
          ElseIf TypeOf ctlControl Is COAWF_Link Then
            mcolwfSelectedLinks.Remove CStr(ctlControl.Index)
          End If
        End If
        
        UnLoad ctlControl
      End If
      
      ' Disassociate object variables.
      Set ctlControl = Nothing
    End If
  End If
  
  ReDim mactlUndoControls(0)
  
End Sub

Public Property Get InitiationType() As WorkflowInitiationTypes
  InitiationType = miInitiationType
  
End Property

Public Property Let InitiationType(ByVal piNewValue As WorkflowInitiationTypes)
  miInitiationType = piNewValue

End Property

Public Property Get BaseTable() As Long
  BaseTable = mlngBaseTableID
  
End Property

Public Property Let BaseTable(ByVal plngNewValue As Long)
  mlngBaseTableID = plngNewValue

End Property

Private Sub MoveToItem(pctlItem As VB.Control)
  ' Scroll the canvas to make the given element visible.
  Dim fMoved As Boolean
  Dim sngRequiredTop As Single
  Dim sngRequiredLeft As Single
  Dim fAnimatedMove As Boolean
  Dim sngMoveStep_X As Single
  Dim sngMoveStep_Y As Single
  Dim sngDifference_X As Single
  Dim sngDifference_Y As Single
  Dim sngMaxDifference As Single
  
  Const MARGIN = 200
  Const ANIMATEDMOVESTEP = 200
    
  ' JPD - If a move is required, then we do it in little moves to give a nice-ish, animated, scrolly
  ' type move. I've left in the code to move straight to the required element in a single jump, just
  ' in case there are any issues with the scrolly move. To do the single jump instead, simply set
  ' fAnimatedMove to be False instead of True. Could do this as a user option, but it's a bit too poxy
  ' for that really.
  fAnimatedMove = True
  fMoved = False
  sngRequiredTop = picDefinition.Top
  sngRequiredLeft = picDefinition.Left
  
  With pctlItem
    If .Top < -sngRequiredTop Then
      If .Top >= MARGIN Then
        sngRequiredTop = MARGIN - .Top
      Else
        sngRequiredTop = 0
      End If
      
      fMoved = True
    End If
    
    If .Top + .Height + sngRequiredTop > picContainer.Height Then
      If .Top + .Height + MARGIN <= picDefinition.Height Then
        sngRequiredTop = picContainer.Height - (.Top + .Height) - MARGIN
      Else
        sngRequiredTop = picContainer.Height - picDefinition.Height
      End If
      
      fMoved = True
    End If
    
    If .Left < -sngRequiredLeft Then
      If .Left >= MARGIN Then
        sngRequiredLeft = MARGIN - .Left
      Else
        sngRequiredLeft = 0
      End If
      
      fMoved = True
    End If
    
    If .Left + .Width + sngRequiredLeft > picContainer.Width Then
      If .Left + .Width + MARGIN <= picDefinition.Width Then
        sngRequiredLeft = picContainer.Width - (.Left + .Width) - MARGIN
      Else
        sngRequiredLeft = picContainer.Width - picDefinition.Width
      End If
      
      fMoved = True
    End If
  End With
  
  If fMoved Then
    DoEvents
    
    Screen.MousePointer = vbHourglass
    
    If fAnimatedMove Then
      sngDifference_Y = (picDefinition.Top - sngRequiredTop)
      If sngDifference_Y < 0 Then sngDifference_Y = -sngDifference_Y
      
      sngDifference_X = (picDefinition.Left - sngRequiredLeft)
      If sngDifference_X < 0 Then sngDifference_X = -sngDifference_X
      
      sngMaxDifference = IIf(sngDifference_X > sngDifference_Y, sngDifference_X, sngDifference_Y)
      
      sngMoveStep_X = ANIMATEDMOVESTEP * (sngDifference_X / sngMaxDifference)
      sngMoveStep_Y = ANIMATEDMOVESTEP * (sngDifference_Y / sngMaxDifference)
      
      Do While (picDefinition.Top <> sngRequiredTop) Or (picDefinition.Left <> sngRequiredLeft)
        If picDefinition.Top < sngRequiredTop Then
          If picDefinition.Top + sngMoveStep_Y < sngRequiredTop Then
            picDefinition.Top = picDefinition.Top + sngMoveStep_Y
          Else
            picDefinition.Top = sngRequiredTop
          End If
        ElseIf picDefinition.Top > sngRequiredTop Then
          If picDefinition.Top - sngMoveStep_Y > sngRequiredTop Then
            picDefinition.Top = picDefinition.Top - sngMoveStep_Y
          Else
            picDefinition.Top = sngRequiredTop
          End If
        End If
        
        If picDefinition.Left < sngRequiredLeft Then
          If picDefinition.Left + sngMoveStep_X < sngRequiredLeft Then
            picDefinition.Left = picDefinition.Left + sngMoveStep_X
          Else
            picDefinition.Left = sngRequiredLeft
          End If
        ElseIf picDefinition.Left > sngRequiredLeft Then
          If picDefinition.Left - sngMoveStep_X > sngRequiredLeft Then
            picDefinition.Left = picDefinition.Left - sngMoveStep_X
          Else
            picDefinition.Left = sngRequiredLeft
          End If
        End If
        
        scrollVertical.value = -picDefinition.Top / mdblVerticalScrollRatio
        scrollHorizontal.value = -picDefinition.Left / mdblHorizontalScrollRatio
        
        DoEvents
      Loop
    Else
      picDefinition.Top = sngRequiredTop
      picDefinition.Left = sngRequiredLeft
      
      scrollVertical.value = -picDefinition.Top / mdblVerticalScrollRatio
      scrollHorizontal.value = -picDefinition.Left / mdblHorizontalScrollRatio
    End If
  
    Screen.MousePointer = vbDefault
  End If
  
End Sub

Private Sub LocateInitiatorOrTriggeredRecord()
  Dim wfElement As VB.Control
  Dim frmUsage As frmUsage
  Dim iElementIndex As Integer
  Dim avColumns() As Variant
  Dim asElementItems() As String
  Dim asValidations() As String
  Dim alElementExprID() As Long
  Dim sMessage As String
  Dim rsTemp As DAO.Recordset
  Dim objComp As CExprComponent
  
  Dim lngExprID As Long
  Dim i As Integer
  Dim sSQL As String

  ReDim mavValidationMessages(1, 0)
  ReDim alElementExprID(0)

  'sSQL = "SELECT EC.exprID " &
  sSQL = "SELECT EC.componentID " & _
        "FROM tmpComponents EC " & _
        "WHERE EC.exprID in  " & _
        "             (SELECT E.exprID  " & _
        "              FROM tmpExpressions E  " & _
        "              WHERE E.utilityid = " & CStr(mlngWorkflowID) & _
        "              AND E.deleted = FALSE) " & _
        "  AND EC.workflowRecord = " & miRecSelType
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  Do While Not (rsTemp.BOF Or rsTemp.EOF)
    'lngExprID = rsTemp!ExprID
    
    Set objComp = New CExprComponent
    objComp.ComponentID = rsTemp!ComponentID
    lngExprID = objComp.RootExpressionID
    Set objComp = Nothing
    
    ReDim Preserve alElementExprID(UBound(alElementExprID) + 1)
    alElementExprID(UBound(alElementExprID)) = lngExprID

    rsTemp.MoveNext
  Loop
  rsTemp.Close
  Set rsTemp = Nothing
  
  
  For Each wfElement In mcolwfElements
    If wfElement.Visible Then
      
      Select Case wfElement.ElementType
        ' ------------------------------------
        ' Decision element
        ' ------------------------------------
        Case elem_Decision
          
          If (wfElement.DecisionFlowExpressionID > 0) And (wfElement.DecisionFlowType = decisionFlowType_Expression) Then
            
            If WorkflowExpressionIDInArray(CLng(wfElement.DecisionFlowExpressionID), alElementExprID) Then
                ValidateWorkflow_AddMessage _
                    ValidateElement_MessagePrefix(wfElement) & "Calculation - <" & GetExpressionName(wfElement.DecisionFlowExpressionID) & ">", _
                    wfElement.ControlIndex
            End If
    
          End If
        
        ' ------------------------------------
        ' Email element
        ' ------------------------------------
        Case elem_Email
          
          If (wfElement.EmailRecord = miRecSelType) Then
            ValidateWorkflow_AddMessage _
                ValidateElement_MessagePrefix(wfElement) & "Email Record", _
                wfElement.ControlIndex
          End If
          
          ' ------------------------------------
          ' Email element items
          ' ------------------------------------
          asElementItems = wfElement.Items
          
          For i = 1 To UBound(asElementItems, 2)
            
            ' DBValue
            If asElementItems(2, i) = giWFEMAILITEM_DBVALUE And asElementItems(5, i) = miRecSelType Then
              ValidateWorkflow_AddMessage _
                  ValidateElement_MessagePrefix(wfElement) & asElementItems(1, i), _
                  wfElement.ControlIndex
            
            ' Field calculations
            ElseIf asElementItems(2, i) = giWFEMAILITEM_CALCULATION Then
              If WorkflowExpressionIDInArray(CLng(asElementItems(56, i)), alElementExprID) Then
                  ValidateWorkflow_AddMessage _
                      ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)), _
                      wfElement.ControlIndex
              End If
            End If
          Next i
          
        ' ------------------------------------
        ' Stored data element
        ' ------------------------------------
        Case elem_StoredData
          
          avColumns = wfElement.DataColumns
    
          ' Data record
          If (wfElement.DataRecord = 0) Then
            ValidateWorkflow_AddMessage _
                ValidateElement_MessagePrefix(wfElement) & "Primary Record", _
                wfElement.ControlIndex
                
          ' Secondary Data record
          ElseIf (wfElement.SecondaryDataRecord = 0) Then
            ValidateWorkflow_AddMessage _
                ValidateElement_MessagePrefix(wfElement) & "Secondary Record", _
                wfElement.ControlIndex
          End If
          
          ' ------------------------------------
          ' Stored data element columns
          ' ------------------------------------
          For i = 1 To UBound(avColumns, 2)
          
            ' DBValue
            If avColumns(4, i) = giWFDATAVALUE_DBVALUE And avColumns(9, i) = miRecSelType Then
              ValidateWorkflow_AddMessage _
                  ValidateElement_MessagePrefix(wfElement) & "Database value - " & CStr(avColumns(1, i)), _
                  wfElement.ControlIndex
            
            ' Field calculations
            ElseIf avColumns(4, i) = giWFDATAVALUE_CALC Then
              If WorkflowExpressionIDInArray(CLng(avColumns(10, i)), alElementExprID) Then
                  ValidateWorkflow_AddMessage _
                      ValidateElement_MessagePrefix(wfElement) & "Calculated column - " & CStr(avColumns(1, i)), _
                      wfElement.ControlIndex
              End If
              
            End If
          Next i
        
        Case elem_WebForm
                
          ' ------------------------------------
          ' WebForm description
          ' ------------------------------------
          asElementItems = wfElement.Items
          
          If (wfElement.DescriptionExprID > 0) Then
            If WorkflowExpressionIDInArray(CLng(wfElement.DescriptionExprID), alElementExprID) Then
              ValidateWorkflow_AddMessage _
                  ValidateElement_MessagePrefix(wfElement) & _
                  "Description - <" & GetExpressionName(wfElement.DescriptionExprID) & "> - Calculation", _
                  wfElement.ControlIndex
            End If
          End If
          
          ' ------------------------------------
          ' WebForm validations
          ' ------------------------------------
          asValidations = wfElement.Validations
          
          For i = 1 To UBound(asValidations, 2)
          
            If (CLng(asValidations(1, i)) > 0) Then
              If WorkflowExpressionIDInArray(CLng(asValidations(1, i)), alElementExprID) Then
                ValidateWorkflow_AddMessage _
                  ValidateElement_MessagePrefix(wfElement) & _
                  "Validation - <" & GetExpressionName(CLng(asValidations(1, i))) & "> - Calculation", _
                  wfElement.ControlIndex
              End If
            End If
            
          Next i
          
          ' ------------------------------------
          ' WebForm element items
          ' ------------------------------------
          For i = 1 To UBound(asElementItems, 2)
            ' DBValue
            If asElementItems(2, i) = giWFFORMITEM_DBVALUE And asElementItems(5, i) = miRecSelType Then
            
              ValidateWorkflow_AddMessage _
                  ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)), _
                  wfElement.ControlIndex
            
            ' Input grid
            ElseIf asElementItems(2, i) = giWFFORMITEM_INPUTVALUE_GRID Then
              
              ' Record selection
              If asElementItems(5, i) = miRecSelType Then
            
                ValidateWorkflow_AddMessage _
                    ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)) & " - Record", _
                    wfElement.ControlIndex
            
              ' Record selection filter
              End If
              
              If (CLng(asElementItems(53, i)) > 0) Then
              
                If WorkflowExpressionIDInArray(CLng(asElementItems(53, i)), alElementExprID) Then
                  ValidateWorkflow_AddMessage _
                      ValidateElement_MessagePrefix(wfElement) & _
                      CStr(asElementItems(1, i)) & " - <" & GetExpressionName(CLng(asElementItems(53, i))) & "> - Filter", _
                      wfElement.ControlIndex
                      
                End If
              End If
            End If
            
            Select Case asElementItems(2, i)
              ' Label default values
              Case giWFFORMITEM_LABEL
              
                If CInt(asElementItems(57, i)) = giWFDATAVALUE_CALC Then
                  If WorkflowExpressionIDInArray(CLng(asElementItems(56, i)), alElementExprID) Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & Replace(CStr(asElementItems(1, i)), "'", ""), _
                        wfElement.ControlIndex
                  End If
                End If
              
              ' Input items default values
              Case giWFFORMITEM_BUTTON, _
                giWFFORMITEM_INPUTVALUE_CHAR, _
                giWFFORMITEM_INPUTVALUE_NUMERIC, _
                giWFFORMITEM_INPUTVALUE_LOGIC, _
                giWFFORMITEM_INPUTVALUE_DATE, _
                giWFFORMITEM_INPUTVALUE_GRID, _
                giWFFORMITEM_INPUTVALUE_DROPDOWN, _
                giWFFORMITEM_INPUTVALUE_LOOKUP, _
                giWFFORMITEM_INPUTVALUE_OPTIONGROUP
                
                If CInt(asElementItems(58, i)) = giWFDATAVALUE_CALC Then
                  If WorkflowExpressionIDInArray(CLng(asElementItems(56, i)), alElementExprID) Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)) & " - Default value calculation", _
                        wfElement.ControlIndex
                  End If
                End If
            End Select
          Next i
          
      End Select
      
    End If
  Next wfElement
  
  ' Display the validation messages.
  Set frmUsage = New frmUsage
  
  ' AE20080317 Fault #13029
'  frmUsage.Width = (Me.ScaleWidth / 2)
'  frmUsage.Height = (Me.ScaleHeight / 2)
  frmUsage.Width = (3 * Screen.Width / 4)
  frmUsage.Height = (Me.ScaleHeight / 2)
  frmUsage.ResetList
    
  For i = 1 To UBound(mavValidationMessages, 2)
    frmUsage.AddToList CStr(mavValidationMessages(0, i)), mavValidationMessages(1, i)
  Next i

  If UBound(mavValidationMessages, 2) > 0 Then
    sMessage = "The Workflow definition uses "
    sMessage = sMessage & IIf(miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL, "Initiator's Record", "Triggered Record")
    sMessage = sMessage & " in the elements listed below."
    
      frmUsage.ShowMessage "Workflow '" & Trim(msWorkflowName) & "'", sMessage, _
          UsageCheckObject.Workflow, _
          USAGEBUTTONS_PRINT + USAGEBUTTONS_OK + USAGEBUTTONS_SELECT, _
          IIf(miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL, "initiator", "triggered")
      
      miUsageChoice = frmUsage.Choice
      
      If frmUsage.Choice = vbRetry Then
        ' Highlight the element 'selected' in the usage check form.
        DeselectAllElements
                
        If frmUsage.Selection >= 0 Then
          iElementIndex = CInt(frmUsage.Selection)
          
          If iElementIndex > 0 Then
            'mcolwfElements(CStr(iElementIndex).HighLighted = True
            SelectElement mcolwfElements(CStr(iElementIndex))
            
            MoveToItem mcolwfElements(CStr(iElementIndex))
            
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = iElementIndex
          
            RefreshMenu
          End If
        End If
      End If

  Else
      sMessage = "No usage found."
      MsgBox sMessage, vbInformation + vbOKOnly, App.ProductName
  End If
  
  UnLoad frmUsage
  Set frmUsage = Nothing
End Sub

Private Sub LocateElement(psElementIdentifier As String, Optional psItemIdentifier As String)
  Dim wfElement As VB.Control
  Dim asElementItems() As String
  Dim avColumns() As Variant
  Dim asValidations() As String
  Dim alElementExprID() As Long
  Dim rsTemp As DAO.Recordset
  Dim objComp As CExprComponent
  Dim lngExprID As Long
  Dim fLocateItems As Boolean
  Dim i As Integer
  Dim sSQL As String
  
  ReDim mavValidationMessages(1, 0)
  ReDim alElementExprID(0)
  
  ' Get a list of the Workflow's expressions
  sSQL = "SELECT DISTINCT tmpComponents.componentID" & _
    " FROM tmpComponents" & _
    " INNER JOIN tmpExpressions ON tmpComponents.exprID = tmpExpressions.exprID" & _
    " WHERE tmpExpressions.utilityID = " & CStr(mlngWorkflowID) & _
    "   AND (tmpExpressions.type = " & CStr(giEXPR_WORKFLOWCALCULATION) & ")" & _
    "   AND (tmpExpressions.deleted = FALSE)" & _
    "   AND ucase(ltrim(rtrim(tmpComponents.workflowElement))) = '" & Replace(UCase(Trim(psElementIdentifier)), "'", "''") & "'"
      
  If Not (Trim(psItemIdentifier) = vbNullString) Then
    sSQL = sSQL & " AND ucase(ltrim(rtrim(tmpComponents.workflowItem))) = '" & Replace(UCase(Trim(psItemIdentifier)), "'", "''") & "'"
    fLocateItems = True
  End If
  
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  Do While Not (rsTemp.BOF Or rsTemp.EOF)
    Set objComp = New CExprComponent
    objComp.ComponentID = rsTemp!ComponentID
    lngExprID = objComp.RootExpressionID
    Set objComp = Nothing
    
    ReDim Preserve alElementExprID(UBound(alElementExprID) + 1)
    alElementExprID(UBound(alElementExprID)) = lngExprID

    rsTemp.MoveNext
  Loop
  rsTemp.Close
  Set rsTemp = Nothing
  
  For Each wfElement In mcolwfElements
    If wfElement.Visible Then
      With wfElement
        Select Case .ElementType
          Case elem_Decision
            ' ------------------------------------
            ' Decision element
            ' ------------------------------------
            If (wfElement.DecisionFlowExpressionID > 0) And (wfElement.DecisionFlowType = decisionFlowType_Expression) Then
              If CLng(wfElement.DecisionFlowExpressionID) > 0 Then
                If WorkflowExpressionIDInArray(CLng(wfElement.DecisionFlowExpressionID), alElementExprID) Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & "Calculation - <" & GetExpressionName(wfElement.DecisionFlowExpressionID) & ">", _
                        wfElement.ControlIndex
                End If
              End If
            End If
            
          Case elem_Email
            ' ------------------------------------
            ' Email element
            ' ------------------------------------
            If .EmailRecord = giWFRECSEL_IDENTIFIEDRECORD Then
              
              If Not fLocateItems Then
                If UCase(Trim(.RecordSelectorWebFormIdentifier)) = UCase(Trim(psElementIdentifier)) Then
                  ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & "Email Record", _
                        wfElement.ControlIndex
                End If
              ElseIf (UCase(Trim(.RecordSelectorWebFormIdentifier)) = UCase(Trim(psElementIdentifier))) And _
                  (UCase(Trim(.RecordSelectorIdentifier)) = UCase(Trim(psItemIdentifier))) Then
                  
                  ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & "Email Record", _
                        wfElement.ControlIndex
              End If
            End If
            
            ' ------------------------------------
            ' Email element items
            ' ------------------------------------
            asElementItems = wfElement.Items
            
            For i = 1 To UBound(asElementItems, 2)
              
              ' DBValue
              If asElementItems(2, i) = giWFEMAILITEM_DBVALUE And asElementItems(5, i) = giWFRECSEL_IDENTIFIEDRECORD Then
                
                If Not fLocateItems Then
                  If UCase(Trim(asElementItems(13, i))) = UCase(Trim(psElementIdentifier)) Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & asElementItems(1, i), _
                        wfElement.ControlIndex
                  End If
                ElseIf UCase(Trim(asElementItems(13, i))) = UCase(Trim(psElementIdentifier)) And _
                    UCase(Trim(asElementItems(14, i))) = UCase(Trim(psItemIdentifier)) Then
                    
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & asElementItems(1, i), _
                        wfElement.ControlIndex
                End If
              
              ' WFValue
              ElseIf asElementItems(2, i) = giWFEMAILITEM_WFVALUE And UCase(Trim(asElementItems(11, i))) = UCase(Trim(psElementIdentifier)) Then
                  
                  If Not fLocateItems Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & asElementItems(1, i), _
                        wfElement.ControlIndex
                  ElseIf UCase(Trim(asElementItems(12, i))) = UCase(Trim(psItemIdentifier)) Then
                      ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & asElementItems(1, i), _
                        wfElement.ControlIndex
                  End If
                            
              ' Field calculations
              ElseIf asElementItems(2, i) = giWFEMAILITEM_CALCULATION Then
                If CLng(asElementItems(56, i)) > 0 Then
                  If WorkflowExpressionIDInArray(CLng(asElementItems(56, i)), alElementExprID) Then
                      ValidateWorkflow_AddMessage _
                          ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)), _
                          wfElement.ControlIndex
                  End If
                End If
              End If
            Next i
            
          Case elem_StoredData
            ' ------------------------------------
            ' Stored Data element
            ' ------------------------------------
            If .DataRecord = giWFRECSEL_IDENTIFIEDRECORD Then
              If UCase(Trim(.RecordSelectorWebFormIdentifier)) = UCase(Trim(psElementIdentifier)) Then
                If Not fLocateItems Then
                  ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & "Primary Record", _
                        wfElement.ControlIndex
                ElseIf UCase(Trim(.RecordSelectorIdentifier)) = UCase(Trim(psItemIdentifier)) Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & "Primary Record", _
                        wfElement.ControlIndex
                End If
              ElseIf UCase(Trim(.SecondaryRecordSelectorWebFormIdentifier)) = UCase(Trim(psElementIdentifier)) Then
                If Not fLocateItems Then
                  ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & "Secondary Record", _
                        wfElement.ControlIndex
                ElseIf UCase(Trim(.RecordSelectorIdentifier)) = UCase(Trim(psItemIdentifier)) Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & "Primary Record", _
                        wfElement.ControlIndex
                End If
              End If
            End If
                
            ' ------------------------------------
            ' Stored data element columns
            ' ------------------------------------
            avColumns = wfElement.DataColumns
            
            For i = 1 To UBound(avColumns, 2)
            
              ' DBValue
              If avColumns(4, i) = giWFDATAVALUE_DBVALUE And avColumns(9, i) = giWFRECSEL_IDENTIFIEDRECORD Then
                
                If UCase(Trim(avColumns(6, i))) = UCase(Trim(psElementIdentifier)) Then
                  If Not fLocateItems Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & "Database value - " & CStr(avColumns(1, i)), _
                        wfElement.ControlIndex
                  ElseIf UCase(Trim(avColumns(7, i))) = UCase(Trim(psItemIdentifier)) Then
                      ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & "Database value - " & CStr(avColumns(1, i)), _
                        wfElement.ControlIndex
                  End If
                End If
              
              ' WFValue
              ElseIf avColumns(4, i) = giWFDATAVALUE_WFVALUE And UCase(Trim(avColumns(6, i))) = UCase(Trim(psElementIdentifier)) Then
                 
                If Not fLocateItems Then
                  ValidateWorkflow_AddMessage _
                      ValidateElement_MessagePrefix(wfElement) & avColumns(1, i), _
                      wfElement.ControlIndex
                ElseIf UCase(Trim(avColumns(7, i))) = UCase(Trim(psItemIdentifier)) Then
                  ValidateWorkflow_AddMessage _
                      ValidateElement_MessagePrefix(wfElement) & avColumns(1, i), _
                      wfElement.ControlIndex
                End If
              
              ' Field calculations
              ElseIf avColumns(4, i) = giWFDATAVALUE_CALC Then
                If CLng(avColumns(10, i)) > 0 Then
                  If WorkflowExpressionIDInArray(CLng(avColumns(10, i)), alElementExprID) Then
                      ValidateWorkflow_AddMessage _
                          ValidateElement_MessagePrefix(wfElement) & "Calculated column - " & CStr(avColumns(1, i)), _
                          wfElement.ControlIndex
                  End If
                End If
              End If
            Next i
            
          Case elem_WebForm
            ' ------------------------------------
            ' WebForm description
            ' ------------------------------------
            asElementItems = wfElement.Items
            
            If (wfElement.DescriptionExprID > 0) Then
              If WorkflowExpressionIDInArray(CLng(wfElement.DescriptionExprID), alElementExprID) Then
                ValidateWorkflow_AddMessage _
                    ValidateElement_MessagePrefix(wfElement) & _
                    "Description - <" & GetExpressionName(wfElement.DescriptionExprID) & "> - Calculation", _
                    wfElement.ControlIndex
              End If
            End If
            
            ' ------------------------------------
            ' WebForm validations
            ' ------------------------------------
            asValidations = wfElement.Validations
            
            For i = 1 To UBound(asValidations, 2)
            
              If (CLng(asValidations(1, i)) > 0) Then
                If WorkflowExpressionIDInArray(CLng(asValidations(1, i)), alElementExprID) Then
                  ValidateWorkflow_AddMessage _
                    ValidateElement_MessagePrefix(wfElement) & _
                    "Validation - <" & GetExpressionName(CLng(asValidations(1, i))) & "> - Calculation", _
                    wfElement.ControlIndex
                End If
              End If
              
            Next i
            
            ' ------------------------------------
            ' WebForm element items
            ' ------------------------------------
            For i = 1 To UBound(asElementItems, 2)
              ' DBValue
              If asElementItems(2, i) = giWFFORMITEM_DBVALUE And asElementItems(5, i) = giWFRECSEL_IDENTIFIEDRECORD Then
                                
                  If UCase(Trim(asElementItems(13, i))) = UCase(Trim(psElementIdentifier)) Then
                    If Not fLocateItems Then
                      ValidateWorkflow_AddMessage _
                          ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)), _
                          wfElement.ControlIndex
                    ElseIf UCase(Trim(asElementItems(14, i))) = UCase(Trim(psItemIdentifier)) Then
                      ValidateWorkflow_AddMessage _
                          ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)), _
                          wfElement.ControlIndex
                    End If
                  End If
                
              ' WFValue
              ElseIf asElementItems(2, i) = giWFEMAILITEM_WFVALUE And UCase(Trim(asElementItems(11, i))) = UCase(Trim(psElementIdentifier)) Then
                  
                  If Not fLocateItems Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & asElementItems(1, i), _
                        wfElement.ControlIndex
                  ElseIf UCase(Trim(asElementItems(12, i))) = UCase(Trim(psItemIdentifier)) Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & asElementItems(1, i), _
                        wfElement.ControlIndex
                  End If
                      
              ' Input grid
              ElseIf asElementItems(2, i) = giWFFORMITEM_INPUTVALUE_GRID Then
                ' Record selection
                If asElementItems(5, i) = giWFRECSEL_IDENTIFIEDRECORD Then
                  If UCase(Trim(asElementItems(13, i))) = UCase(Trim(psElementIdentifier)) Then
                    If Not fLocateItems Then
                      ValidateWorkflow_AddMessage _
                          ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)) & " - Record", _
                          wfElement.ControlIndex
                    ElseIf UCase(Trim(asElementItems(14, i))) = UCase(Trim(psItemIdentifier)) Then
                      ValidateWorkflow_AddMessage _
                          ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)) & " - Record", _
                          wfElement.ControlIndex
                    End If
                  End If
                  
                End If
                                
                ' Record selection filter
                If (CLng(asElementItems(53, i)) > 0) Then
                  If WorkflowExpressionIDInArray(CLng(asElementItems(53, i)), alElementExprID) Then
                    ValidateWorkflow_AddMessage _
                        ValidateElement_MessagePrefix(wfElement) & _
                        CStr(asElementItems(1, i)) & " - <" & GetExpressionName(CLng(asElementItems(53, i))) & "> - Filter", _
                        wfElement.ControlIndex

                  End If
                End If
              End If
              
              Select Case asElementItems(2, i)
                Case giWFFORMITEM_LABEL
                  ' Label default values
                  If CInt(asElementItems(57, i)) = giWFDATAVALUE_CALC Then
                    If WorkflowExpressionIDInArray(CLng(asElementItems(56, i)), alElementExprID) Then
                      ValidateWorkflow_AddMessage _
                          ValidateElement_MessagePrefix(wfElement) & Replace(CStr(asElementItems(1, i)), "'", ""), _
                          wfElement.ControlIndex
                    End If
                  End If
                
                ' Input items default values
                Case giWFFORMITEM_BUTTON, _
                  giWFFORMITEM_INPUTVALUE_CHAR, _
                  giWFFORMITEM_INPUTVALUE_NUMERIC, _
                  giWFFORMITEM_INPUTVALUE_LOGIC, _
                  giWFFORMITEM_INPUTVALUE_DATE, _
                  giWFFORMITEM_INPUTVALUE_GRID, _
                  giWFFORMITEM_INPUTVALUE_DROPDOWN, _
                  giWFFORMITEM_INPUTVALUE_LOOKUP, _
                  giWFFORMITEM_INPUTVALUE_OPTIONGROUP
                  
                  If CInt(asElementItems(58, i)) = giWFDATAVALUE_CALC Then
                    If WorkflowExpressionIDInArray(CLng(asElementItems(56, i)), alElementExprID) Then
                      ValidateWorkflow_AddMessage _
                          ValidateElement_MessagePrefix(wfElement) & CStr(asElementItems(1, i)) & " - Default value calculation", _
                          wfElement.ControlIndex
                    End If
                  End If
              End Select
            Next i
        End Select
      End With
    End If
  Next wfElement
    
  
  ' Display the validation messages.
  Set frmUsage = New frmUsage
  
  ' AE20080317 Fault #13029
'  frmUsage.Width = (Me.ScaleWidth / 2)
'  frmUsage.Height = (Me.ScaleHeight / 2)
  frmUsage.Width = (3 * Screen.Width / 4)
  frmUsage.Height = (Me.ScaleHeight / 2)
  frmUsage.ResetList
      
  For i = 1 To UBound(mavValidationMessages, 2)
    frmUsage.AddToList CStr(mavValidationMessages(0, i)), mavValidationMessages(1, i)
  Next i

  Dim iElementIndex As Integer
  Dim sMessage As String
  
  If UBound(mavValidationMessages, 2) > 0 Then
    sMessage = "The Workflow definition uses '" & psElementIdentifier & "' in the elements listed below."
    
    frmUsage.ShowMessage "Workflow '" & Trim(msWorkflowName) & "'", sMessage, _
          UsageCheckObject.Workflow, _
          USAGEBUTTONS_PRINT + USAGEBUTTONS_OK + USAGEBUTTONS_SELECT, "details"
      
      ' AE20080317 Fault #13031
      miUsageChoice = frmUsage.Choice
      
      If frmUsage.Choice = vbRetry Then
        ' Highlight the element 'selected' in the usage check form.
        DeselectAllElements
                
        If frmUsage.Selection >= 0 Then
          iElementIndex = CInt(frmUsage.Selection)
          
          If iElementIndex > 0 Then
'            mcolwfElements(CStr(iElementIndex).HighLighted = True
            SelectElement mcolwfElements(CStr(iElementIndex))
            
            MoveToItem mcolwfElements(CStr(iElementIndex))
            
            ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
            miSelectionOrder(UBound(miSelectionOrder)) = iElementIndex
          
            RefreshMenu
          End If
        End If
      End If
  Else
    ' AE20080317 Fault #13030
    'sMessage = "No usage of found."
    sMessage = "No usage found."
    MsgBox sMessage, vbInformation + vbOKOnly, App.ProductName
  End If
  
  UnLoad frmUsage
  Set frmUsage = Nothing
  
End Sub

Private Function WorkflowExpressionIDInArray(plngExprID, palElementExprID() As Long) As Boolean

  Dim lngExprID As Variant
  
  For Each lngExprID In palElementExprID
    If CLng(lngExprID) = plngExprID Then
      WorkflowExpressionIDInArray = True
    End If
  Next lngExprID
  
End Function

Public Sub FindUsage()
  Dim frmWFUsage As frmWorkflowFindUsage
  Set frmWFUsage = New frmWorkflowFindUsage
  
  CancelElementAddMode
  
  miUsageChoice = vbOK
  
  Do While miUsageChoice = vbOK
    frmWFUsage.Initialise mcolwfElements, _
      miWFUsageSelection, _
      msWFUsageElement, _
      msWFUsageItem, _
      miInitiationType
      
    ' AE20080325 Fault #12918
    ' frmWFUsage.Initialise mcolwfElements, frmWFUsage.Selection, frmWFUsage.Element, frmWFUsage.Item
    frmWFUsage.Show vbModal
    
    If frmWFUsage.Choice = vbOK Then
      miWFUsageSelection = frmWFUsage.Selection
      msWFUsageElement = frmWFUsage.Element
      msWFUsageItem = frmWFUsage.Item
      
      Select Case frmWFUsage.Selection
        Case wfRecSelType
          Call LocateInitiatorOrTriggeredRecord
        
        Case wfElement
          Call LocateElement(frmWFUsage.Element)
          
        Case wfWebFormItem
          Call LocateElement(frmWFUsage.Element, frmWFUsage.Item)
      End Select
    Else
      Exit Do
    End If
  Loop
  
  UnLoad frmWFUsage
  Set frmWFUsage = Nothing

End Sub

Private Sub WorkflowElement_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Move the control(s)
  On Error GoTo ErrorTrap
  
  Dim wfTempElement As VB.Control
  Dim wfTempLink As COAWF_Link
  Dim ctl As VB.Control
  Dim contScaleMode As Integer
  Dim contRect As Rect
    
  ' AE20080508 Fault #13153
'  If (Application.AccessMode = accFull Or Application.AccessMode = accSupportMode) Then
  If (Application.AccessMode = accFull Or _
      Application.AccessMode = accSupportMode) And mblnMouseDownFired Then
    
    GetCursorPos currPoint
    contScaleMode = mcolwfElements(CStr(Index)).Container.ScaleMode
        
    ' Move the selected controls if the left button key is down, and the control is selected
    If Button = vbLeftButton Then
      If Not mblnDragging Then
        mblnDragging = True
      End If
      
      ' move the control if they are different
      If (currPoint.x > startPointSingle.x) Or (currPoint.x < startPointSingle.x) _
        Or (currPoint.y > startPointSingle.y) Or (currPoint.y < startPointSingle.y) Then
        
        ' move the control
        With mcolwfElements(CStr(Index))
          .Move .Left + picContainer.ScaleX(currPoint.x - startPointSingle.x, _
            vbPixels, contScaleMode), _
            .Top + picContainer.ScaleY(currPoint.y - startPointSingle.y, _
            vbPixels, contScaleMode)
          
          For Each wfTempLink In ASRWFLink1
            If wfTempLink.StartElementIndex = .ControlIndex Or _
              wfTempLink.EndElementIndex = .ControlIndex Then

              FormatLink wfTempLink
            End If
          Next
          Set wfTempLink = Nothing

          ' refresh container
          .Parent.Refresh
        End With
        LSet startPointSingle = currPoint
        
        SetChanged False
      End If

    End If
    
  End If
  
TidyUpAndExit:
  Screen.MousePointer = vbDefault
  Exit Sub

ErrorTrap:
  Call ClipCursorByNum(0)
  Resume TidyUpAndExit

End Sub

Private Sub WorkflowElement_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  
  mlngXOffset = x
  mlngYOffset = y
  
  GetCursorPos startPointSingle
  GetCursorPos startPointMulti

  If Not mcolwfElements(CStr(Index)).Highlighted Then
    ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
    miSelectionOrder(UBound(miSelectionOrder)) = Index
  End If
  
  If (Shift <> vbShiftMask) And (Shift <> vbCtrlMask) And (Not mcolwfElements(CStr(Index)).Highlighted) Then
    If Not abMenu.Tools("ID_WFElement_Link").Checked Then
      DeselectAllElements
    End If
  End If
    
  SelectElement mcolwfElements(CStr(Index))
  
  If Button = vbLeftButton Then
    
    ' AE20080508 Fault #13153
    mblnMouseDownFired = True
    
    Call GetWindowRect(Me.picContainer.hWnd, WindowRect)
    Call ClipCursor(WindowRect)
  
    ' AE20080613 Fault #11830
'    If Not mcolwfElements(CStr(Index)).HighLighted Then
'      ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
'      miSelectionOrder(UBound(miSelectionOrder)) = Index
'    End If
'
'    SelectElement mcolwfElements(CStr(Index))

    If abMenu.Tools("ID_WFElement_Link").Checked And (SelectedElementCount >= 2) Then
      AddLinks
      DeselectAllElements
    
      ReDim Preserve miSelectionOrder(UBound(miSelectionOrder) + 1)
      miSelectionOrder(UBound(miSelectionOrder)) = Index
      SelectElement mcolwfElements(CStr(Index))
    End If
    
    ' AE20080613 Fault #11830
'    mcolwfElements(CStr(Index)).ZOrder 0
'    RefreshMenu
  End If
  
  mcolwfElements(CStr(Index)).ZOrder 0
  
  RefreshMenu

End Sub

Private Sub WorkflowElement_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Call ClipCursorByNum(0)
  
  Dim wfTempLink As COAWF_Link
  Dim ctl As VB.Control
  Dim contScaleMode As Integer
  Dim contRect As Rect
  Dim lngXMouse As Long
  Dim lngYMouse As Long
  
  If Button = vbLeftButton And mblnDragging Then
    ' restore full mouse movement
    mblnDragging = False
    
    ' AE20080508 Fault #13153
    mblnMouseDownFired = False
    
    Screen.MousePointer = vbHourglass
    UI.LockWindow Me.hWnd
    
    contScaleMode = mcolwfElements(CStr(Index)).Container.ScaleMode
    
    For Each ctl In mcolwfSelectedElements
      If ctl.ControlIndex <> Index Then
        ' move the control if they are different
        If currPoint.x <> startPointMulti.x Or currPoint.y <> startPointMulti.y Then
          ' move the control
          With ctl
            .Move .Left + picContainer.ScaleX(currPoint.x - startPointMulti.x, _
              vbPixels, contScaleMode), _
              .Top + picContainer.ScaleY(currPoint.y - startPointMulti.y, _
              vbPixels, contScaleMode)
                            
            For Each wfTempLink In ASRWFLink1
              If wfTempLink.StartElementIndex = .ControlIndex Or _
                wfTempLink.EndElementIndex = .ControlIndex Then
  
                FormatLink wfTempLink
              End If
            Next
            Set wfTempLink = Nothing

            .Parent.Refresh
          End With
        
          IsChanged = True
        End If
      End If
    Next
    LSet startPointMulti = currPoint

    If Not abMenu.Tools("ID_WFElement_Link").Checked Then
      ResizeCanvas
    End If

    UI.UnlockWindow
    Screen.MousePointer = vbDefault
    
    ' allow background processing
    DoEvents
    
  ElseIf Button = vbRightButton Then ' Handle right button presses.
      UI.GetMousePos lngXMouse, lngYMouse
      mlngXDrop = x
      mlngYDrop = y
      
      abMenu.Bands("ElementBand").TrackPopup -1, -1
  End If
    

  
End Sub

Private Function IsValidCollectionItem(ByRef pcolCollection As Collection, _
  ByRef psIndexKey As String) As Boolean
  ' Return TRUE if the given security table exists in the collection.
  Dim ctlItem As VB.Control
  
  On Error GoTo err_IsValid
  Set ctlItem = pcolCollection.Item(psIndexKey)
  
  IsValidCollectionItem = True
  Set ctlItem = Nothing
  
  Exit Function
  
err_IsValid:
  IsValidCollectionItem = False
  
End Function
Private Function LoadNewElementOfType(piElementType) As VB.Control
  
  Dim wfElement As VB.Control
  Dim objFont As StdFont
  
  miControlIndex = miControlIndex + 1
  
  Select Case piElementType
  Case elem_Begin
    Load ASRWFBeginEnd1(ASRWFBeginEnd1.UBound + 1)
    Set wfElement = ASRWFBeginEnd1(ASRWFBeginEnd1.UBound)
    
    wfElement.ElementType = elem_Begin
    wfElement.Caption = "BEGIN"
  
  Case elem_Terminator
    Load ASRWFBeginEnd1(ASRWFBeginEnd1.UBound + 1)
    Set wfElement = ASRWFBeginEnd1(ASRWFBeginEnd1.UBound)
    
    wfElement.ElementType = elem_Terminator
    wfElement.Caption = "END"
    
  Case elem_WebForm
    Load ASRWFWebform1(ASRWFWebform1.UBound + 1)
    Set wfElement = ASRWFWebform1(ASRWFWebform1.UBound)
    wfElement.Caption = "Web Form"
    wfElement.Identifier = GetUniqueIdentifier(wfElement)
    
    wfElement.WebFormFGColor = 6697779
    wfElement.WebFormBGColor = 16513017
    
    Set objFont = New StdFont
    objFont.Name = gobjDefaultScreenFont.Name
    objFont.Size = gobjDefaultScreenFont.Size
    objFont.Bold = False
    objFont.Italic = False
    objFont.Strikethrough = False
    objFont.Underline = False
    Set wfElement.WebFormDefaultFont = objFont
    Set objFont = Nothing
    
  Case elem_Email
    Load ASRWFEmail1(ASRWFEmail1.UBound + 1)
    Set wfElement = ASRWFEmail1(ASRWFEmail1.UBound)
    
    wfElement.Caption = "Email"
    wfElement.EMailSubject = "OpenHR Workflow"
    If miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
      wfElement.EmailRecord = giWFRECSEL_TRIGGEREDRECORD
    ElseIf miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL Then
      wfElement.EmailRecord = giWFRECSEL_UNIDENTIFIED
    End If

  Case elem_Decision
    Load ASRWFDecision1(ASRWFDecision1.UBound + 1)
    Set wfElement = ASRWFDecision1(ASRWFDecision1.UBound)
    
    wfElement.Caption = "?"
    wfElement.DecisionCaptionType = decisionCaption_Y_N
    wfElement.DecisionFlowType = decisionFlowType_Button
    
  Case elem_StoredData
    Load ASRWFStoredData1(ASRWFStoredData1.UBound + 1)
    Set wfElement = ASRWFStoredData1(ASRWFStoredData1.UBound)
    
    wfElement.Caption = "Stored Data"
    wfElement.Identifier = GetUniqueIdentifier(wfElement)
    
    wfElement.DataAction = DATAACTION_UPDATE
    Select Case miInitiationType
      Case WORKFLOWINITIATIONTYPE_MANUAL
        wfElement.DataRecord = giWFRECSEL_INITIATOR
      Case WORKFLOWINITIATIONTYPE_TRIGGERED
        wfElement.DataRecord = giWFRECSEL_TRIGGEREDRECORD
      Case Else
        wfElement.DataRecord = giWFRECSEL_IDENTIFIEDRECORD
    End Select

  Case elem_SummingJunction, elem_Or, elem_Connector1, elem_Connector2
    Load ASRWFJunctionElement1(ASRWFJunctionElement1.UBound + 1)
    Set wfElement = ASRWFJunctionElement1(ASRWFJunctionElement1.UBound)
    wfElement.ElementType = piElementType
    
  End Select
  
  wfElement.ControlIndex = miControlIndex
  mcolwfElements.Add wfElement, CStr(wfElement.ControlIndex)
  
  Set LoadNewElementOfType = wfElement
        
End Function
