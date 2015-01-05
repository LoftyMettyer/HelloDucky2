VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "actbar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{A48C54F8-25F4-4F50-9112-A9A3B0DBAD63}#1.0#0"; "coa_label.ocx"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "coa_line.ocx"
Object = "{98B2556E-F719-4726-9028-5F2EAB345800}#1.0#0"; "coasd_checkbox.ocx"
Object = "{3EBC9263-7DE3-4E87-8721-81ACE59CD84E}#1.2#0"; "coasd_combo.ocx"
Object = "{3CCEDCBE-4766-494F-84C9-95993D77BD56}#1.0#0"; "coasd_command.ocx"
Object = "{FFAE31F9-C18D-4C20-AAF7-74C1356185D9}#1.1#0"; "coasd_frame.ocx"
Object = "{5F165695-EDF2-40E1-BD8E-8D2E6325BDCF}#1.0#0"; "coasd_image.ocx"
Object = "{32648AC7-4D67-4E6A-A546-1B7783115C22}#1.0#0"; "coasd_ole.ocx"
Object = "{CE18FF03-F3BF-4C4F-81DC-192ED1E1B91F}#1.0#0"; "coasd_optiongroup.ocx"
Object = "{58F88252-94BB-43CE-9EF9-C971F73B93D4}#1.0#0"; "coasd_selection.ocx"
Object = "{714061F3-25A6-4821-B196-7D15DCCDE00E}#1.0#0"; "coasd_selectionbox.ocx"
Object = "{0BE8C79E-5090-4700-B420-B767D1E19561}#1.0#0"; "coasd_spinner.ocx"
Object = "{93EA589D-C793-4EE4-BE53-52A646038BAF}#1.0#0"; "coasd_workingpattern.ocx"
Object = "{AD837810-DD1E-44E0-97C5-854390EA7D3A}#3.2#0"; "coa_navigation.ocx"
Object = "{C1ECF24D-7ECA-4C65-BBFD-DD76B98E3DF2}#1.0#0"; "coasd_colourselector.ocx"
Begin VB.Form frmScrDesigner2 
   AutoRedraw      =   -1  'True
   Caption         =   "Screen Designer"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5027
   Icon            =   "frmScrDesigner2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin SystemMgr.COASD_Label asrDummyTextBox 
      Height          =   315
      Index           =   0
      Left            =   2550
      TabIndex        =   14
      Top             =   570
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
   End
   Begin COASDColSelector.COASD_ColourSelector ASRColourSelector 
      Height          =   315
      Index           =   0
      Left            =   2835
      TabIndex        =   20
      Top             =   3390
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   556
   End
   Begin COANavigation.COA_Navigation ASRDummyNavigation 
      Height          =   510
      Index           =   0
      Left            =   6030
      TabIndex        =   19
      Top             =   2430
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   900
      Caption         =   "Navigate..."
      DisplayType     =   1
      NavigateIn      =   0
      NavigateTo      =   ""
      InScreenDesigner=   -1  'True
      ColumnID        =   0
      ColumnName      =   ""
      Selected        =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontSize        =   8.25
      FontStrikethrough=   0   'False
      FontUnderline   =   -1  'True
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      NavigateOnSave  =   0   'False
   End
   Begin VB.PictureBox picFormIcon 
      Height          =   285
      Left            =   7440
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   3255
      Visible         =   0   'False
      Width           =   285
   End
   Begin COASDSelection.COASD_Selection ASRSelectionMarkers 
      Height          =   840
      Index           =   0
      Left            =   6435
      TabIndex        =   17
      Top             =   300
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1482
   End
   Begin VB.PictureBox picPageContainer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      ScaleHeight     =   315
      ScaleWidth      =   1035
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin COASDOLE.COASD_OLE asrDummyOLEContents 
      Height          =   990
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   2250
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   1746
      Caption         =   "Caption"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
   End
   Begin COASDSelectionBox.COASD_SelectionBox asrBoxMovementMarker 
      Height          =   510
      Index           =   0
      Left            =   5115
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   900
      BorderColor     =   -2147483640
   End
   Begin COASDSelectionBox.COASD_SelectionBox asrboxMultiSelection 
      Height          =   570
      Left            =   5190
      TabIndex        =   2
      Top             =   330
      Visible         =   0   'False
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BorderColor     =   -2147483640
      BorderStyle     =   3
   End
   Begin COALine.COA_Line ASRDummyLine 
      Height          =   30
      Index           =   0
      Left            =   3375
      Top             =   1635
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   53
   End
   Begin COASDOptionGroup.COASD_OptionGroup ASRDummyOptions 
      Height          =   630
      Index           =   0
      Left            =   1290
      TabIndex        =   4
      Top             =   2565
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2381
      _ExtentY        =   1111
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
   Begin COASDWorkingPat.COASD_WorkingPattern ASRCustomDummyWP 
      Height          =   1005
      Index           =   0
      Left            =   1230
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1773
   End
   Begin COASDCommand.COASD_Command asrDummyLink 
      Height          =   330
      Index           =   0
      Left            =   4590
      TabIndex        =   6
      Top             =   2685
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   582
      Caption         =   "Caption"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      ForeColor       =   -2147483630
   End
   Begin COASDFrame.COASD_Frame asrDummyFrame 
      Height          =   315
      Index           =   0
      Left            =   3315
      TabIndex        =   7
      Top             =   2370
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
   End
   Begin COALabel.COA_Label asrDummyLabel 
      Height          =   315
      Index           =   0
      Left            =   2580
      TabIndex        =   8
      Top             =   165
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
   End
   Begin COASDCheckbox.COASD_Checkbox asrDummyCheckBox 
      Height          =   315
      Index           =   0
      Left            =   3690
      TabIndex        =   9
      Top             =   570
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
   End
   Begin COASDSpinner.COASD_Spinner asrDummySpinner 
      Height          =   315
      Index           =   0
      Left            =   3735
      TabIndex        =   10
      Top             =   165
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
   End
   Begin COASDCombo.COASD_Combo asrDummyCombo 
      Height          =   315
      Index           =   0
      Left            =   3315
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
   End
   Begin ComctlLib.TabStrip tabPages 
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      MousePointer    =   1
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
   Begin COALabel.COA_Label asrDummyOLEContents2 
      Height          =   360
      Index           =   0
      Left            =   2550
      TabIndex        =   15
      Top             =   990
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
   End
   Begin COASDImage.COASD_Image asrDummyImage 
      Height          =   315
      Index           =   0
      Left            =   3315
      TabIndex        =   16
      Top             =   1965
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
   End
   Begin COALabel.COA_Label asrDummyPhoto 
      Height          =   360
      Index           =   0
      Left            =   4920
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
   End
   Begin ActiveBarLibraryCtl.ActiveBar abScreen 
      Left            =   5580
      Top             =   3030
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
      Bands           =   "frmScrDesigner2.frx":000C
   End
End
Attribute VB_Name = "frmScrDesigner2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Constants.
Const giMAXTABS = 50
Const giSTANDARDMOVEMENT = 15
Const giMOVEMENTMARKERWIDTH = 20
Const gLngDFLTSCREENHEIGHT = 4900
Const gLngDFLTSCREENWIDTH = 7100
Const gLngAUTOFORMATLABELCOLUMN = 300
Const gLngAUTOFORMATYOFFSET = 100
Const gLngAUTOFORMATYSTART = 300

Private Const MIN_FORM_HEIGHT = 1660
Private Const MIN_FORM_WIDTH = 5520

' Properties.
Private gfAlignToGrid As Boolean
Private gDfltForeColour As ColorConstants
Private giGridX As Long
Private giGridY As Long
Private gLngScreenID As Long
Private gLngTableID As Long
Private gfNewScreen As Boolean
Private gfChangedScreen As Boolean

' Globals.
Private gfLoading As Boolean
Private gfMultiSelecting As Boolean
Private gLngMultiSelectionXStart As Long
Private gLngMultiSelectionYStart As Long
Private gfStretchDown As Boolean
Private gfStretchUp As Boolean
Private gfStretchRight As Boolean
Private gfStretchLeft As Boolean
Private gfMoveSelection As Boolean
Private gLngOldX As Long
Private gLngOldY As Long
Private gfMouseDown As Boolean
Private gfExitToScrMgr As Boolean
Private gfActivating As Boolean
Private giLastActionFlag As UndoActionFlags
Private gfUndo_TabsCreated As Boolean
Private giUndo_ControlIndex As Integer
Private giUndo_ControlAutoLabelIndex As Integer
Private gsUndo_ControlType As String
Private gsUndo_ControlAutoLabelType As String
Private giUndo_TabPageIndex As Integer
Private gsUndo_TabPageCaption As String
Private gavUndo_PastedControls() As Variant
Private gactlUndo_DeletedControls() As VB.Control
Private gactlClipboardControls() As VB.Control
Private gbAutoSendFrameToBack As Boolean

Private mlngLastX As Long
Private mlngLastY As Long
Private mlngXOffset As Long
Private mlngYOffset As Long

Private mbKeyStretching As Boolean
Private mbKeyMoving As Boolean

Private mlngMouseX As Long
Private mlngMouseY As Long

Public Function HasNonSSIntranetControls() As Boolean
  ' Return TRUE if the current screen definition contains
  ' controls that are not permitted in Self-service Intranet screens.
  ' The following controls are NOT permitted :
  '   TabStrips
  '   OLEs
  '   Photos
  '   Images
  Dim ctlControl As VB.Control
  Dim sName As String
  Dim fFound As Boolean
  Dim iIndex As Integer
  
  For Each ctlControl In Me.Controls
    sName = UCase(ctlControl.Name)
    
'    If sName = "TABPAGES" Then
'      If tabPages.Tabs.Count > 0 Then
'        HasNonSSIntranetControls = True
'        Exit Function
'      End If
'    End If
    
'    If sName = "ASRDUMMYPHOTO" Or _
      sName = "ASRDUMMYOLECONTENTS" Or _
      sName = "ASRDUMMYIMAGE" Then
    If sName = "ASRDUMMYIMAGE" Then
    
      ' Do not bother with the dummy screen controls.
      If (ctlControl.Index > 0) Then
        ' Do not bother with controls in the deleted array.
        fFound = False
        For iIndex = 1 To UBound(gactlUndo_DeletedControls)
          If ctlControl Is gactlUndo_DeletedControls(iIndex) Then
            fFound = True
            Exit For
          End If
        Next iIndex
    
        If Not fFound Then
          ' Do not bother with controls in the clipboard array.
          For iIndex = 1 To UBound(gactlClipboardControls)
            If ctlControl Is gactlClipboardControls(iIndex) Then
              fFound = True
              Exit For
            End If
          Next iIndex
        End If
      
        If Not fFound Then
          HasNonSSIntranetControls = True
          Exit Function
        End If
      End If
    End If
  Next ctlControl

  HasNonSSIntranetControls = False

End Function


Public Function IsSSIntranetScreen() As Boolean
  Dim fIsSSIntranetScreen As Boolean

  fIsSSIntranetScreen = False
  
  With recScrEdit
    .Index = "idxScreenID"
    .Seek "=", gLngScreenID
  
    If Not .NoMatch Then
      fIsSSIntranetScreen = IIf(IsNull(.Fields("SSIntranet")), False, .Fields("SSIntranet"))
    End If
  End With

  IsSSIntranetScreen = fIsSSIntranetScreen
  
End Function

Private Function OLEType(iColumnID As Integer) As String

  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", iColumnID
  
    If Not .NoMatch Then
      Select Case !OLEType
        Case OLE_LOCAL
          OLEType = "(Local)"
        Case OLE_SERVER
          OLEType = "(Server)"
        Case OLE_EMBEDDED
          OLEType = "(Linked)"
      End Select
    Else
      OLEType = ""
    End If
  End With

End Function

Private Sub abScreen_Click(ByVal pTool As ActiveBarLibraryCtl.Tool)

  EditMenu pTool.Name

End Sub

Private Sub abScreen_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  
  ' Do not let the user modify the layout.
  Cancel = True

End Sub


Private Sub asrDummyCheckBox_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummyCheckBox(Index), Source, X, Y

End Sub

Private Sub asrDummyCheckBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown asrDummyCheckBox(Index), Button, Shift, X, Y
  
End Sub

Private Sub asrDummyCheckBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove asrDummyCheckBox(Index), Button, X, Y
  
End Sub

Private Sub asrDummyCheckBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummyCheckBox(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyCombo_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummyCombo(Index), Source, X, Y

End Sub

Private Sub asrDummyCombo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown asrDummyCombo(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyCombo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove asrDummyCombo(Index), Button, X, Y

End Sub

Private Sub asrDummyCombo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummyCombo(Index), Button, Shift, X, Y

End Sub


' ASR Navigation Fire events

Private Sub ASRDummyNavigation_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ScreenControl_DragDrop ASRDummyNavigation(Index), Source, X, Y
End Sub

Private Sub ASRDummyNavigation_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRDummyNavigation_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ScreenControl_MouseDown ASRDummyNavigation(Index), Button, Shift, X, Y
End Sub

Private Sub ASRDummyNavigation_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ScreenControl_MouseMove ASRDummyNavigation(Index), Button, X, Y
End Sub

Private Sub ASRDummyNavigation_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ScreenControl_MouseUp ASRDummyNavigation(Index), Button, Shift, X, Y
End Sub


' ASR Colour Selector

Private Sub ASRColourSelector_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ScreenControl_DragDrop ASRColourSelector(Index), Source, X, Y
End Sub

Private Sub ASRColourSelector_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ASRColourSelector_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ScreenControl_MouseDown ASRColourSelector(Index), Button, Shift, X, Y
End Sub

Private Sub ASRColourSelector_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ScreenControl_MouseMove ASRColourSelector(Index), Button, X, Y
End Sub

Private Sub ASRColourSelector_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ScreenControl_MouseUp ASRColourSelector(Index), Button, Shift, X, Y
End Sub




Private Sub asrDummyFrame_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Select the control.
  ScreenControl_MouseDown asrDummyFrame(Index), Button, Shift, X, Y
  'Form_MouseDown Button, Shift, X + asrDummyFrame(Index).Left, Y + asrDummyFrame(Index).Top

End Sub

Private Sub asrDummyFrame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Move the control.
  ScreenControl_MouseMove asrDummyFrame(Index), Button, X, Y
  'Form_MouseMove Button, Shift, X + asrDummyFrame(Index).Left, Y + asrDummyFrame(Index).Top

End Sub

Private Sub asrDummyFrame_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummyFrame(Index), Button, Shift, X, Y
  'Form_MouseUp Button, Shift, X + asrDummyFrame(Index).Left, Y + asrDummyFrame(Index).Top

End Sub

Private Sub asrDummyImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown asrDummyImage(Index), Button, Shift, X, Y
  
End Sub

Private Sub asrDummyImage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove asrDummyImage(Index), Button, X, Y

End Sub

Private Sub asrDummyImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummyImage(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummyLabel(Index), Source, X, Y
  
End Sub

Private Sub asrDummyLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown asrDummyLabel(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove asrDummyLabel(Index), Button, X, Y

End Sub

Private Sub asrDummyLabel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummyLabel(Index), Button, Shift, X, Y
        
End Sub

Private Sub asrDummyLink_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummyLink(Index), Source, X, Y

End Sub

Private Sub asrDummyLink_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown asrDummyLink(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove asrDummyLink(Index), Button, X, Y

End Sub

Private Sub asrDummyLink_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummyLink(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyOLEContents_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummyOLEContents(Index), Source, X, Y

End Sub

Private Sub asrDummyOLEContents_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown asrDummyOLEContents(Index), Button, Shift, X, Y
  
End Sub

Private Sub asrDummyOLEContents_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove asrDummyOLEContents(Index), Button, X, Y

End Sub

Private Sub asrDummyOLEContents_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummyOLEContents(Index), Button, Shift, X, Y

End Sub

Private Sub ASRDummyOptions_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop ASRDummyOptions(Index), Source, X, Y

End Sub

Private Sub ASRDummyOptions_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown ASRDummyOptions(Index), Button, Shift, X, Y

End Sub

Private Sub ASRDummyOptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove ASRDummyOptions(Index), Button, X, Y

End Sub

Private Sub ASRDummyOptions_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp ASRDummyOptions(Index), Button, Shift, X, Y

End Sub



Private Sub asrDummyPhoto_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummyPhoto(Index), Source, X, Y

End Sub

Private Sub asrDummyPhoto_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown asrDummyPhoto(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyPhoto_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove asrDummyPhoto(Index), Button, X, Y

End Sub

Private Sub asrDummyPhoto_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummyPhoto(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummySpinner_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummySpinner(Index), Source, X, Y

End Sub

Private Sub asrDummySpinner_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown asrDummySpinner(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummySpinner_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove asrDummySpinner(Index), Button, X, Y

End Sub

Private Sub asrDummySpinner_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummySpinner(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyTextBox_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummyTextBox(Index), Source, X, Y
  
End Sub

Private Sub asrDummyTextBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' Select the control.
  ScreenControl_MouseDown asrDummyTextBox(Index), Button, Shift, X, Y

End Sub

Private Sub asrDummyTextBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove asrDummyTextBox(Index), Button, X, Y

End Sub

Private Sub asrDummyTextBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp asrDummyTextBox(Index), Button, Shift, X, Y

End Sub

Private Sub ASRCustomDummyWP_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop ASRCustomDummyWP(Index), Source, X, Y

End Sub

Private Sub ASRCustomDummyWP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown ASRCustomDummyWP(Index), Button, Shift, X, Y
  
End Sub

Private Sub ASRCustomDummyWP_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove ASRCustomDummyWP(Index), Button, X, Y
  
End Sub

Private Sub ASRCustomDummyWP_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp ASRCustomDummyWP(Index), Button, Shift, X, Y

End Sub

Private Sub ASRSelectionMarkers_Stretch(Index As Integer, Direction As String, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim iCount As Integer
  Dim lngHeight As Long
  Dim lngWidth As Long
  Dim lngTop As Long
  Dim lngLeft As Long
  Dim bCanStretch As Boolean
  Dim iGridSize As Integer
    
  'UI.LockWindow Me.hWnd
  On Error GoTo CannotStretch
  
  iGridSize = 2
    
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
                  
      If .Visible Then
      
        ' Default sizes for the stretch
        lngTop = .AttachedObject.Top
        lngHeight = .AttachedObject.Height
        lngLeft = .AttachedObject.Left
        lngWidth = .AttachedObject.Width
        bCanStretch = False
              
        Select Case Direction
      
          ' Stretch North West
          Case "TopLeft"
            bCanStretch = Not .HasLockedHeight Or Not .HasLockedWidth
            If bCanStretch Then
              lngHeight = IIf(.Original_Height - Y > .AttachedObject.MinimumHeight, .Original_Height - Y, .AttachedObject.Height)
              lngTop = IIf(.Original_Height - Y > .AttachedObject.MinimumHeight, .Original_Top + Y, .AttachedObject.Top)
              lngLeft = IIf(.Original_Width - X > .AttachedObject.MinimumWidth, .Original_Left + X, .AttachedObject.Left)
              lngWidth = IIf(.Original_Width - X > .AttachedObject.MinimumWidth, .Original_Width - X, .AttachedObject.Width)
            End If
      
          ' Stretch North
          Case "TopCentre"
            bCanStretch = Not .HasLockedHeight And .Original_Height - Y > .AttachedObject.MinimumHeight
            If bCanStretch Then
              lngTop = .Original_Top + Y
              lngHeight = .Original_Height - Y
            End If

          ' Stretch North East
          Case "TopRight"
            If Not .HasLockedHeight And (.Original_Height - Y > .AttachedObject.MinimumHeight) Then
              lngTop = .Original_Top + Y
              lngHeight = .Original_Height - Y
            End If
            
            If Not .HasLockedWidth Then
              lngLeft = .AttachedObject.Left
              lngWidth = .Original_Width + X
              bCanStretch = IIf(lngWidth = .AttachedObject.Width, False, True)
            End If

          Case "CentreLeft"
            bCanStretch = (.Original_Width - X > .AttachedObject.MinimumWidth And Not .HasLockedWidth)
            If bCanStretch Then
              lngLeft = .Original_Left + X
              lngWidth = .Original_Width - X
            End If
          
          Case "CentreRight"
            lngWidth = .Original_Width + X
            bCanStretch = IIf(lngWidth = .AttachedObject.Width And Not .HasLockedWidth, False, True)

          Case "BottomLeft"
            If .Original_Width - X > .AttachedObject.MinimumWidth And Not .HasLockedWidth Then
              '.AttachedObject.Move .Original_Left + X, .AttachedObject.Top, .Original_Width - X, .Original_Height + Y
              lngLeft = .Original_Left + X
              lngWidth = .Original_Width - X
            End If
            lngHeight = .Original_Height + Y
            'bCanStretch = IIf(lngWidth = .AttachedObject.Width And lngHeight = .AttachedObject.Height, False, True)
            bCanStretch = IIf(IsWithin(lngWidth, .AttachedObject.Width, iGridSize) And IsWithin(lngHeight, .AttachedObject.Height, iGridSize), False, Not .HasLockedWidth)
          
          Case "BottomCentre"
            lngHeight = .Original_Height + Y
            bCanStretch = IIf(IsWithin(lngHeight, .AttachedObject.Height, iGridSize), False, Not .HasLockedHeight)

          Case "BottomRight"
            If Not .HasLockedWidth Then
              lngWidth = .Original_Width + X
            End If
                        
            If Not .HasLockedHeight Then
              lngHeight = .Original_Height + Y
            End If
            
            bCanStretch = IIf(IsWithin(lngWidth, .AttachedObject.Width, iGridSize) And IsWithin(lngHeight, .AttachedObject.Height, iGridSize), False, True)
            '.AttachedObject.Move .AttachedObject.Left, .AttachedObject.Top, .Original_Width + X, .Original_Height + Y
      
        End Select
      
        ' Only move the control if it is stretchable
        If bCanStretch Then
          .AttachedObject.Move lngLeft, lngTop, lngWidth, lngHeight
        End If
      
      End If
        
    End With
  Next iCount

  'UI.UnlockWindow
  Exit Sub
  
CannotStretch:
  Exit Sub


End Sub

Private Sub ASRSelectionMarkers_StretchEnd(Index As Integer, Direction As String, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim iCount As Integer
  
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize, .AttachedObject.Width + (.MarkerSize * 2), .AttachedObject.Height + (.MarkerSize * 2)
        .RefreshSelectionMarkers True
      End If
    End With
  Next iCount
 
  ' Flag screen as having changed
  IsChanged = True
 
End Sub

Private Sub ASRSelectionMarkers_StretchStart(Index As Integer, Direction As String, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim iCount As Integer
  
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .SaveOriginalSizes
        .ShowSelectionMarkers False
      End If
    End With
  Next iCount

  ' Store original x,y coordinates
  mlngXOffset = X
  mlngYOffset = Y

End Sub

Private Sub COA_ColourSelector1_GotFocus()

End Sub

Private Sub Form_Activate()
  ' Ensure the screen designer form is at the front of the display.
  On Error GoTo ErrorTrap
  
  Me.ZOrder 0
  
  ' Refresh the properties screen.
  Set frmScrObjProps.CurrentScreen = Me
  frmScrObjProps.RefreshProperties

  ' Refresh the menu/toolbar display.
  frmSysMgr.RefreshMenu

  gfActivating = True

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error activating Screen Designer form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Sub

Private Function DropControl(pVarPageContainer As Variant, pCtlSource As Control, pSngX As Single, pSngY As Single) As Boolean
  ' Drop the given control onto the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iControlType As Long
  Dim lngColumnID As Long
  Dim sCaption As String
  Dim sTableName As String
  Dim sColumnName As String
  Dim objFont As StdFont
  Dim objMisc As New Misc
  Dim ctlControl As VB.Control
  
  ' Deselect all existing controls.
  fOK = DeselectAllControls
  
  If fOK Then
  
    ' Check that a column or standard control is being dropped
    If (pCtlSource Is frmToolbox.trvColumns) Or _
      (pCtlSource Is frmToolbox.trvStandardControls) Then
        
      ' If we are dropping a column control ...
      If pCtlSource Is frmToolbox.trvColumns Then
        
        'Find the definition for the column being dropped
        With frmToolbox.trvColumns.SelectedItem
          lngColumnID = val(Mid(.key, 2))
        End With
          
        With recColEdit
          .Index = "idxColumnID"
          .Seek "=", lngColumnID
          
          If Not .NoMatch Then
          
            ' Add the required control type.
            iControlType = .Fields("controlType")
            Set ctlControl = AddControl(iControlType)
            fOK = Not (ctlControl Is Nothing)
            
            If fOK Then
  
              ' Set the last action flag and enable the Undo menu option.
              If Me.abScreen.Tools("ID_AutoLabel").Checked = True Then
                giLastActionFlag = giACTION_DROPCONTROLAUTOLABEL
              Else
                giLastActionFlag = giACTION_DROPCONTROL
              End If
                
              giUndo_ControlIndex = ctlControl.Index
              gsUndo_ControlType = ctlControl.Name
              
'              giLastActionFlag = giACTION_DROPCONTROL
'              giUndo_ControlIndex = ctlControl.Index
'              gsUndo_ControlType = ctlControl.Name
            
              Set ctlControl.Container = pVarPageContainer
              ctlControl.Left = AlignX(CLng(pSngX))
              ctlControl.Top = AlignY(CLng(pSngY))
              ctlControl.ColumnID = .Fields("columnID").value
                            
              ' Give the control a tooltip.
              sColumnName = .Fields("columnName")
              With recTabEdit
                .Index = "idxTableID"
                .Seek "=", recColEdit.Fields("TableId").value
                    
                If Not .NoMatch Then
                  sTableName = .Fields("tableName").value
                  ctlControl.ToolTipText = sTableName & "." & sColumnName
                End If
              End With
              
             
              ' Initialise the new control's font and forecolour.
              If ScreenControl_HasFont(iControlType) Then
                Set objFont = New StdFont
                objFont.Name = Me.Font.Name
                objFont.Size = Me.Font.Size
                objFont.Bold = Me.Font.Bold
                objFont.Italic = Me.Font.Italic
                Set ctlControl.Font = objFont
                Set objFont = Nothing
              End If
              
              If ScreenControl_HasForeColor(iControlType) Then
                ctlControl.ForeColor = gDfltForeColour
              End If
              
              If ScreenControl_HasCaption(iControlType) Then
                ctlControl.Caption = objMisc.StrReplace(.Fields("columnName").value, "_", " ", False) & vbNullString
              End If
              
              If ScreenControl_HasText(iControlType) Then
                ctlControl.Caption = ctlControl.ToolTipText
                
                If iControlType = giCTRL_OLE Then
'XXXXXXXXXXXXXXXXXXXXXX
                  ctlControl.ButtonCaption = OLEType(ctlControl.ColumnID)
'                  ctlControl.Caption = ctlControl.Caption & gsOLEDISPLAYTYPE_CONTENTS
                End If
              End If
              
              ' Initialise the navigation properties
              If ScreenControl_HasNavigation(iControlType) Then
                ctlControl.ColumnName = GetColumnName(ctlControl.ColumnID, False)
                ctlControl.DisplayType = NavigationDisplayType.Hyperlink
                ctlControl.Caption = "Navigate To..."
                ctlControl.NavigateTo = ctlControl.NavigateTo
              End If
              
              ' JIRA-539 Below piece of code looks strange, but for some reason without it, the background of some controls
              ' goes black. Must be kicking off a refresh somewhere down the line
              If ScreenControl_HasBackColor(iControlType) Then
                ctlControl.BackColor = ctlControl.BackColor
              End If
              
              If ScreenControl_HasOptions(iControlType) Then
'                ctlControl.Options = ReadColumnControlValues(ctlControl.ColumnID)
                ctlControl.SetOptions ReadColumnControlValues(ctlControl.ColumnID)
              End If
              
              ' Default the control's propertes.
              fOK = AutoSizeControl(ctlControl)
                
              If fOK Then
                fOK = SelectControl(ctlControl)
              End If
              
'              If fOK Then
'                ctlControl.Visible = True
'                ctlControl.ZOrder 0
'              End If
            End If
            
            
            ' RH 07/08/00 - Check here if we need to autoadd a label for the ctl
            If Me.abScreen.Tools("ID_AutoLabel").Checked = True Then
            
              Select Case iControlType
              
              ' RH 15/09/00 - BUG 940. Do not drop an autolabel for checkboxes
              '                        Position label below 'top' of control
              
                Case giCTRL_COMBOBOX, giCTRL_SPINNER, giCTRL_TEXTBOX, giCTRL_COLOURPICKER
  
                  AutoLabel pVarPageContainer, pSngX, pSngY + 75, sColumnName
                  
                Case giCTRL_IMAGE, giCTRL_OLE, giCTRL_PHOTO, giCTRL_WORKINGPATTERN
                
                  AutoLabel pVarPageContainer, pSngX, pSngY, sColumnName
                
                Case Else
              
                  ' Dont drop a label automatically
              
              End Select
                         
            End If
            
            If fOK Then
              'TM20010914 Fault 1753
              'The ActiveBar control does mot have the visible property, so to avoid err
              'we only check the visible property of other controls.
              If ctlControl.Name <> "abScreen" Then
                ctlControl.Visible = True
                ctlControl.ZOrder 0
              End If
            End If
            
            Set ctlControl = Nothing
          
          End If
        End With
        
      ' If we are dropping a standard control ...
      ElseIf pCtlSource Is frmToolbox.trvStandardControls Then
        
        ' Add a tab page.
        If frmToolbox.trvStandardControls.SelectedItem.key = "PAGETABCTRL" Then
          fOK = AddTabPage
          
          If fOK Then
            ' Set the last action flag and enable the Undo menu option.
            giLastActionFlag = giACTION_DROPTABPAGE
            giUndo_TabPageIndex = tabPages.Tabs.Count
    
            tabPages.SetFocus
          End If
        Else
        
          ' Add the new control to the screen.
          Select Case frmToolbox.trvStandardControls.SelectedItem.key
            Case "LABELCTRL"
              iControlType = giCTRL_LABEL
              Set ctlControl = AddControl(iControlType)
              sCaption = "Label"
              
            Case "FRAMECTRL"
              iControlType = giCTRL_FRAME
              Set ctlControl = AddControl(iControlType)
              sCaption = "Frame"
              
            Case "IMAGECTRL"
              iControlType = giCTRL_IMAGE
              Set ctlControl = AddControl(iControlType)
          
            Case "LINECTRL"
              iControlType = giCTRL_LINE
              Set ctlControl = AddControl(iControlType)
          
            Case "NAVIGATIONCTRL"
              iControlType = giCTRL_NAVIGATION
              Set ctlControl = AddControl(iControlType)
          
          End Select
  
          fOK = Not (ctlControl Is Nothing)
          
          'Check that a new control was added successfully
          If fOK Then
    
            With ctlControl
  
              ' Set the last action flag and enable the Undo menu option.
              giLastActionFlag = giACTION_DROPCONTROL
              giUndo_ControlIndex = .Index
              gsUndo_ControlType = .Name
            
              Set .Container = pVarPageContainer
              .Left = AlignX(CLng(pSngX))
              .Top = AlignY(CLng(pSngY))
              .ColumnID = 0
              
              ' Initialise the new control's font and forecolour.
              If ScreenControl_HasFont(iControlType) Then
                Set objFont = New StdFont
                objFont.Name = Me.Font.Name
                objFont.Size = Me.Font.Size
                objFont.Bold = Me.Font.Bold
                objFont.Italic = Me.Font.Italic
                Set .Font = objFont
                Set objFont = Nothing
              End If
              
              If ScreenControl_HasForeColor(iControlType) Then
                .ForeColor = gDfltForeColour
              End If
              
              If ScreenControl_HasNavigation(iControlType) Then
                .DisplayType = NavigationDisplayType.Hyperlink
                .NavigateTo = "about:blank"
                sCaption = "Navigate To..."
              End If
              
              If ScreenControl_HasCaption(iControlType) Then
                .Caption = sCaption
              End If
              
              ' Default the control's propertes.
              fOK = AutoSizeControl(ctlControl)
                
              If fOK Then
                fOK = SelectControl(ctlControl)
              End If
              
              If fOK Then
                .Visible = True
                
                ' JDM - 19/08/02 - Fault 4309 - Put frame at the back
                If iControlType = giCTRL_FRAME And gbAutoSendFrameToBack Then
                  .ZOrder 1
                Else
                  .ZOrder 0
                End If
              
              End If
            End With
          End If
          
          ' Disassociate object variables.
          Set ctlControl = Nothing
          
        End If
      End If
    
      ' Set focus on the screen designer form.
      Me.SetFocus
  
    End If
  End If
    
  If fOK Then
    
    ' Mark the screen as having changed.
    gfChangedScreen = True
    frmSysMgr.RefreshMenu
  
    ' Refresh the properties screen.
    Set frmScrObjProps.CurrentScreen = Me
    frmScrObjProps.RefreshProperties
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objMisc = Nothing
  Set objFont = Nothing
  Set ctlControl = Nothing
  ' Return the success/failure value.
  DropControl = fOK
  Exit Function

ErrorTrap:
  ' Flag the error.
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  On Error GoTo ErrorTrap
  
  If CurrentPageContainer Is Me Then
    If Not DropControl(Me, Source, X, Y) Then
      MsgBox "Unable to drop the control." & vbCr & vbCr & _
        Err.Description, vbExclamation + vbOKOnly, App.ProductName
    End If
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_GotFocus()
  ' Refresh the properties screen.
  Set frmScrObjProps.CurrentScreen = Me
  frmScrObjProps.RefreshProperties

End Sub

Private Sub Form_Initialize()
  ' Initialise global variables.
  On Error GoTo ErrorTrap

  gfMultiSelecting = False
  gfExitToScrMgr = False
  
  gbAutoSendFrameToBack = True
  
  ' Initialise properties.
  gfAlignToGrid = True
  giGridX = 40
  giGridY = 40
  
  ASRSelectionMarkers(0).Visible = False
  
  ' Hide the dummy control array controls.
  With asrDummyLabel(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With asrDummyTextBox(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With asrDummyPhoto(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With asrDummyOLEContents(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With asrDummyImage(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With asrDummyFrame(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With asrDummyCombo(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With asrDummySpinner(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With asrDummyCheckBox(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With ASRDummyOptions(0)
    .Left = -.Width
    .Top = -.Height
  End With
  With asrDummyLink(0)
    .Left = -.Width
    .Top = -.Height
  End With
  
  ' Clear the tab strip.
  tabPages.Tabs.Clear
  
  ' Disable the 'undo' menu option until we have somethig to undo.
  giLastActionFlag = giACTION_NOACTION

  ' Initialise gloabl arrays.
  ReDim gactlUndo_DeletedControls(0)
  ReDim gactlClipboardControls(0)
  ReDim gavUndo_PastedControls(2, 0)

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error initialising Screen Designer form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Dim iCount As Integer
  
  ' Complete stretching of selected controls
  If mbKeyStretching Then
    ASRSelectionMarkers_StretchEnd 0, "", 0, 0, 0, 0
  End If

  ' Complete moving of selected controls
  If mbKeyMoving Then
     For iCount = 1 To ASRSelectionMarkers.Count - 1
      With ASRSelectionMarkers(iCount)
        If .Visible Then
          .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize
          .ShowSelectionMarkers True
        End If
      End With
     Next iCount
  End If

  'TM20020102 Fault 2879
  Dim bHandled As Boolean
  
  bHandled = frmSysMgr.tbMain.OnKeyUp(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

End Sub

Private Sub Form_Load()
  ' Update the count of screen designer forms, in the toolbox and properties forms.
  On Error GoTo ErrorTrap
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
  
  frmScrObjProps.ScreenCount = frmScrObjProps.ScreenCount + 1
  frmToolbox.ScreenCount = frmToolbox.ScreenCount + 1
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error loading Screen Designer form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Sub


Public Property Get IsChanged() As Boolean
  ' Return the 'screen changed' flag.
  IsChanged = gfChangedScreen
  
End Property

Public Property Let IsChanged(pfNewValue As Boolean)
  ' Set the 'screen changed' flag.
  gfChangedScreen = pfNewValue
        
  ' Menu may be dependent on the status of the screen.
  frmSysMgr.RefreshMenu
  
End Property

Public Property Get IsNew() As Boolean
  ' Return the 'new screen' flag.
  IsNew = gfNewScreen
  
End Property

Public Property Let IsNew(pfNewValue As Boolean)
  ' Set the 'new screen' flag.
  gfNewScreen = pfNewValue
  
End Property

Public Property Get GridX() As Long
  ' Return the horizontal grid interval.
  GridX = giGridX
  
End Property

Public Property Let GridX(plngGridSize As Long)
  ' Set the horizontal grid interval.
  giGridX = plngGridSize
  
End Property

Public Property Get GridY() As Long
  ' Return the vertical grid interval.
  GridY = giGridY
  
End Property

Public Property Let GridY(plngGridSize As Long)
  ' Set the vertical grid interval.
  giGridY = plngGridSize

End Property

Public Property Get DefaultForeColour() As ColorConstants
  ' Return the default foreground colour.
  DefaultForeColour = gDfltForeColour
  
End Property

Public Property Let DefaultForeColour(ByVal pNewValue As ColorConstants)
  ' Set the default foreground colour.
  gDfltForeColour = pNewValue

End Property


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Process key strokes.
  On Error GoTo ErrorTrap
  
  Dim sngXMove As Single
  Dim sngYMove As Single
  Dim sngXStretch As Single
  Dim sngYStretch As Single
  Dim strDirection As String
  Dim iCount As Integer
  Dim bHandled As Boolean
  
  bHandled = False
  mbKeyMoving = False
  mbKeyStretching = False

  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
  
  ' JDM - 22/08/02 - Fault 4267 - F4 needs to bring up properties dialog
  If KeyCode = vbKeyF4 Then
    EditMenu "ID_ScreenDesignerScreenProperties"
    bHandled = True
  End If

  ' DELETE key deletes any selected controls.
  ' If there are no selected controls then the current tab page is deleted.
  ' If there are no selected controls and no tab pages then nothing happens.
  If Not bHandled Then
    If KeyCode = vbKeyDelete Then
      If SelectedControlsCount > 0 Then
        If Not DeleteSelectedControls Then
          MsgBox "Unable to delete controls." & vbCr & vbCr & _
            Err.Description, vbExclamation + vbOKOnly, App.ProductName
        Else
          bHandled = True
        End If
      Else
        If tabPages.Tabs.Count > 0 Then
          If Not DeleteTabPage(tabPages.SelectedItem.Index, True) Then
            MsgBox "Unable to delete tab page." & vbCr & vbCr & _
              Err.Description, vbExclamation + vbOKOnly, App.ProductName
          End If
          
          bHandled = True
        End If
      End If
    End If
  End If
  
  ' SHIFT and ARROW keys stretch any selected controls.
  If Not bHandled Then
    If ((Shift And vbShiftMask) > 0) Then
      
      ' Determine which stretch is being done.
      Select Case KeyCode
        Case vbKeyLeft
          strDirection = "CentreRight"
          sngXStretch = -giSTANDARDMOVEMENT
        Case vbKeyRight
          strDirection = "CentreRight"
          sngXStretch = giSTANDARDMOVEMENT
        Case vbKeyUp
          strDirection = "BottomCentre"
          sngYStretch = -giSTANDARDMOVEMENT
        Case vbKeyDown
          strDirection = "BottomCentre"
          sngYStretch = giSTANDARDMOVEMENT
      End Select
      
      ' Stretch the selected controls if required.
      If (sngXStretch <> 0) Or (sngYStretch <> 0) Then
        ASRSelectionMarkers_StretchStart 0, strDirection, 0, 0, sngXStretch, sngYStretch
        ASRSelectionMarkers_Stretch 0, strDirection, 0, 0, sngXStretch, sngYStretch
        mbKeyStretching = True
      End If
    End If
  
    ' CTRL and ARROW keys move the selected controls.
    If ((Shift And vbCtrlMask) > 0) Then
  
      mbKeyMoving = True
  
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
  
      End Select
  
      ' Flag the selected selction markers to be moved
      If (sngXMove <> 0) Or (sngYMove <> 0) Then
        For iCount = 1 To ASRSelectionMarkers.Count - 1
          ASRSelectionMarkers(iCount).ShowSelectionMarkers False
        Next iCount
      
        ScreenControl_KeyMove sngXMove, sngYMove
      End If
     
    End If
  
    ' JDM - 20/12/01 - Fault 3315 - CTRL-keys didn't work
    If ((Shift And vbCtrlMask) > 0) Then
    
      Select Case KeyCode
        Case vbKeyZ
          EditMenu "ID_Undo"
        Case vbKeyX
          EditMenu "ID_Cut"
        Case vbKeyC
          EditMenu "ID_Copy"
        Case vbKeyV
          EditMenu "ID_Paste"
        Case vbKeyA
          EditMenu "ID_ScreenSelectAll"
        
      End Select
    
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

Public Property Set DefaultFont(pObjNewValue As Object)
  ' Set the screen's default font.
  Set Me.Font = pObjNewValue
  
End Property

Public Property Get DefaultFont() As Object
  ' Return the screen's default font.
  Set DefaultFont = Me.Font
  
End Property



Public Sub EditMenu(ByVal psMenuOption As String)
  ' Process the menu options.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  Dim lngPictureID As Long
  Dim sFileName As String
  
  Select Case psMenuOption
    
    ' Undo the last deletion, cut or addition of a control.
    Case "ID_Undo"
      If Not UndoLastAction Then
        MsgBox "Unable to undo the last action." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Cut the selected controls.
    Case "ID_Cut"
      If Not CutSelectedControls Then
        MsgBox "Unable to cut controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    ' Copy the selected control.
    Case "ID_Copy"
      If Not CopySelectedControls Then
        MsgBox "Unable to copy controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    ' Paste the control from the clipboard.
    Case "ID_Paste"
      If Not PasteControls Then
        MsgBox "Unable to paste controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    ' Delete the selected control.
    Case "ID_ScreenObjectDelete"
      ' If there are no selected controls then the current tab page is deleted.
      ' If there are no selected controls and no tab pages then nothing happens.
      If SelectedControlsCount > 0 Then
        If Not DeleteSelectedControls Then
          MsgBox "Unable to delete controls." & vbCr & vbCr & _
            Err.Description, vbExclamation + vbOKOnly, App.ProductName
        End If
      Else
        If tabPages.Tabs.Count > 0 Then
          If Not DeleteTabPage(tabPages.SelectedItem.Index, True) Then
            MsgBox "Unable to delete the tab." & vbCr & vbCr & _
              Err.Description, vbExclamation + vbOKOnly, App.ProductName
          End If
        End If
      End If
    
    ' Select all controls on the current screen.
    Case "ID_ScreenSelectAll"
      If Not SelectAllPageControls Then
        MsgBox "Unable to select controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    
    ' Save the current screen.
    Case "ID_Save"
      SaveScreen
    
    ' Display the Screen properties window.
    Case "ID_ScreenDesignerScreenProperties", "ID_ScreenProperties"
      frmScrEdit.ScreenID = gLngScreenID
      frmScrEdit.Show vbModal
      Set frmScrEdit = Nothing
      
      ' The screen name may have been changed so update the form caption, and
      ' also the frmScrOpen screen list if it is loaded.
      recScrEdit.Index = "idxScreenID"
      recScrEdit.Seek "=", gLngScreenID
      
      SetFormCaption Me, "Screen Manager - " & recScrEdit.Fields("name") & vbNullString
  
      For iLoop = 1 To Forms.Count - 1
        If Forms(iLoop).Name = "frmScrOpen" Then
          Forms(iLoop).RefreshScreens
        End If
      Next iLoop
      
    ' Display the object properties window.
    Case "ID_ObjectProperties"
      If Not frmScrObjProps.Visible Then
        frmScrObjProps.Show
      Else
        frmScrObjProps.WindowState = vbNormal
        frmScrObjProps.ZOrder 0
      End If
            
    ' Display the Toolbox window.
    Case "ID_Toolbox"
      If Not frmToolbox.Visible Then
        frmToolbox.Show
      Else
        frmToolbox.WindowState = vbNormal
        frmToolbox.ZOrder 0
      End If
    
    ' Call the pop-up that allows the user to define the object
    ' tab order for the current screen.
    Case "ID_ObjectOrder"
      Set frmTabOrd.CurrentScreen = Me
      frmTabOrd.Show vbModal
      Set frmTabOrd = Nothing
      
    ' AutoFormat all controls on the current screen.
    Case "ID_AutoFormat"
      AutoFormatScreen
      
    ' Display the Editor Options pop-up
    Case "ID_Options"
      Set frmScrEditOpts.CurrentScreen = Me
      frmScrEditOpts.Show vbModal
      Set frmScrEditOpts = Nothing
      
    Case "ID_AutoLabel"
      If mblnAutoLabelling = True Then Exit Sub
      mblnAutoLabelling = True
      'TM20011015 Fault 2959
      'Set the checked property of the AutoLabel button.
      Me.abScreen.Tools("ID_AutoLabel").Checked = Not Me.abScreen.Tools("ID_AutoLabel").Checked
      frmSysMgr.tbMain.Tools("ID_AutoLabel").Checked = Me.abScreen.Tools("ID_AutoLabel").Checked
      mblnAutoLabelling = False
     
    ' Bring selected controls to front
    Case "ID_BringToFront"
      BringSelectedControlsToFront
    
    ' Send selected controls to back
    Case "ID_SendToBack"
      SendSelectedControlsToBack

    ' Left align selected controls
    Case "ID_ScreenControlAlignLeft"
      LeftAlignSelectedControls
    
    ' Centre align selected controls
    Case "ID_ScreenControlAlignCentre"
      CentreAlignSelectedControls

    ' Right align selected controls
    Case "ID_ScreenControlAlignRight"
      RightAlignSelectedControls
     
    ' Top align selected controls
    Case "ID_ScreenControlAlignTop"
      TopAlignSelectedControls
     
    ' Middle align selected controls
    Case "ID_ScreenControlAlignMiddle"
      MiddleAlignSelectedControls
     
    ' Bottom align selected controls
    Case "ID_ScreenControlAlignBottom"
      BottomAlignSelectedControls
     
  End Select
  
  Exit Sub
  
ErrorTrap:

End Sub


Private Sub AutoFormatScreen()
  ' Automatically dumps all relevant database fields onto the
  ' screen with a label attached if required.
  ' Automatically adds extra tab pages when necessary.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iStartPage As Integer
  Dim iColumnCount As Integer
  Dim iControlType As Long
  Dim iCurrentPage As Integer
  Dim lngYPosition As Long
  Dim lngControlColumn As Long
  Dim sngControlHeight As Single
  Dim sSQL As String
  Dim sTableName As String
  Dim sColumnName As String
  Dim varCurrentPageContainer As Variant
  Dim objMisc As New Misc
  Dim objFont As StdFont
  Dim ctlControl As VB.Control
  Dim ctlLabelControl As COA_Label
  Dim rsColumns As DAO.Recordset
  'Dim WaitWindow As WaitMessage.MessageWindow
'  Dim WaitWindow As NewWaitMsg.clsNewWaitMsg
  
  ' Flag that we are automatically formatting the screen.
  gfLoading = True
  
  ' Initialise variables.
  lngYPosition = gLngAUTOFORMATYSTART
  lngControlColumn = Int(Me.Width / 3)
  
  ' Deselect all existing controls.
  fOK = DeselectAllControls
  
  If fOK Then
    ' User the Hourglass mousepointer until everything is finished.
    Screen.MousePointer = vbHourglass
    
    ' Lock the window so that we don't get messy screen updates.
    UI.LockWindow Me.hWnd
    
    ' Get the table name.
    With recTabEdit
      .Index = "idxTableID"
      .Seek "=", gLngTableID
  
      If Not .NoMatch Then
        sTableName = .Fields("tableName")
      Else
        sTableName = ""
      End If
    End With
    
    ' Find the last empty tab page on which to start the automatic format.
    iStartPage = FirstEmptyPage
    fOK = (iStartPage >= 0)
    
    If fOK Then
      ' If there are no controls, then the start page should be 0. So
      ' we need to get rid of any existing tab pages to start with.
      If iStartPage = 0 Then
        For iLoop = tabPages.Tabs.Count To 1 Step -1
    
          ' Unload the tabpage's picture container.
          UnLoad picPageContainer(tabPages.Tabs(iLoop).Tag)
    
          ' Remove the tab from the tabstrip.
          tabPages.Tabs.Remove iLoop
    
          ' Hide the tabstrip if we now have no tabs left.
          If tabPages.Tabs.Count = 0 Then
            tabPages.Visible = False
          End If
        Next iLoop
      End If
      
      ' If there are no existing tabpages, but there are existing controls then
      ' we need to create a new page for the existing controls, and start the
      ' autoFormat on the following page.
      If iStartPage = 1 Then
        fOK = AddTabPage
          
        If fOK Then
          ' Ensure the new pagecontainer is visible.
          picPageContainer(tabPages.Tabs(iStartPage).Tag).Visible = True
          iStartPage = 2
          gfUndo_TabsCreated = True
        End If
      Else
        gfUndo_TabsCreated = False
      End If
    End If
  End If
  
  If fOK Then
    ' Add a new tab page if required.
    If iStartPage > tabPages.Tabs.Count Then
      fOK = AddTabPage
      If fOK Then
        ' Ensure the new pagecontainer is visible.
        picPageContainer(tabPages.Tabs(iStartPage).Tag).Visible = True
      Else
        MsgBox "Unable to add more than " & Trim(Str(giMAXTABS)) & " page tabs." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
    End If
  End If
  
  If fOK Then
    ' Get the first page's pictureBox container.
    iCurrentPage = iStartPage
    If iCurrentPage = 0 Then
      Set varCurrentPageContainer = Me
    Else
      Set varCurrentPageContainer = picPageContainer(tabPages.Tabs(iCurrentPage).Tag)
    End If
  
    ' Read the required column and control information from the database tables.
    ' Get the column count for the primary table.
    sSQL = "SELECT COUNT(columnID) AS columnCount" & _
      " FROM tmpColumns" & _
      " WHERE tableID = " & Trim(Str(gLngTableID)) & _
      " AND deleted = FALSE" & _
      " AND columnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM))
    Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    iColumnCount = rsColumns!ColumnCount
    Set rsColumns = Nothing

    ' Display the progress bar.
    With gobjProgress
      .AVI = dbScreenAutoLayout
      .MainCaption = "Screen Manager"
      .Caption = Application.Name
      .NumberOfBars = 1
      .Bar1MaxValue = iColumnCount
      .Bar1Caption = "Autoformatting screen..."
      .Time = True
      .Cancel = True
      .OpenProgress
    End With
    
    ' Get the column details for the primary table.
    sSQL = "SELECT columnID, columnName, controlType" & _
    " FROM tmpColumns" & _
    " WHERE tableID=" & Trim(Str(gLngTableID)) & _
    " AND deleted = FALSE" & _
    " AND columnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
    " ORDER BY columnName"
    Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
    'Loop though columns and format controls.
    Do While (Not rsColumns.EOF) And fOK
    
      ' Add the required control type.
      iControlType = rsColumns!ControlType
      Set ctlControl = AddControl(iControlType)
      fOK = Not (ctlControl Is Nothing)
      
      If fOK Then
        With ctlControl
          Set .Container = varCurrentPageContainer
          .Left = AlignX(lngControlColumn)
          .Top = AlignY(lngYPosition)
          .ColumnID = rsColumns!ColumnID
          
          ' Give the control a tooltip.
          sColumnName = rsColumns!ColumnName
          If Len(sTableName) = 0 Then
            .ToolTipText = sColumnName
          Else
            .ToolTipText = sTableName & "." & sColumnName
          End If
                     
          ' Initialise the new control's font and forecolour.
          If ScreenControl_HasFont(iControlType) Then
            Set objFont = New StdFont
            objFont.Name = Me.Font.Name
            objFont.Size = Me.Font.Size
            objFont.Bold = Me.Font.Bold
            objFont.Italic = Me.Font.Italic
            Set .Font = objFont
            Set objFont = Nothing
          End If
        
          If ScreenControl_HasForeColor(iControlType) Then
            .ForeColor = gDfltForeColour
          End If
        
          If ScreenControl_HasCaption(iControlType) Then
            .Caption = objMisc.StrReplace(rsColumns!ColumnName, "_", " ", False) & vbNullString
          End If
          
          If ScreenControl_HasText(iControlType) Then
            .Caption = .ToolTipText
            
            If iControlType = giCTRL_OLE Then
              .Caption = .Caption & gsOLEDISPLAYTYPE_CONTENTS
            End If
          End If
                    
          ' Initialise the navigation properties
          If ScreenControl_HasNavigation(iControlType) Then
            ctlControl.ColumnName = GetColumnName(ctlControl.ColumnID, False)
            ctlControl.DisplayType = NavigationDisplayType.Hyperlink
            ctlControl.Caption = "Navigate To..."
            ctlControl.NavigateTo = ctlControl.NavigateTo
          End If
          
          If ScreenControl_HasOptions(iControlType) Then
'            .Options = ReadColumnControlValues(rsColumns!ColumnID)
            .SetOptions ReadColumnControlValues(rsColumns!ColumnID)
          End If
            
          ' Default the control's propertes.
          fOK = AutoSizeControl(ctlControl)
            
          If fOK Then
            .Visible = True
            
            ' If the new control hangs off the bottom of the current page then move it onto
            ' the next page.
            sngControlHeight = .Height
          
            If .Top + sngControlHeight > varCurrentPageContainer.ScaleHeight Then
              ' If we are adding the first page then we need to move existing controls onto this page,
              ' and then create another page for the current control.
              If tabPages.Tabs.Count = 0 Then
                If ScreenControlsCount > 0 Then
                  If AddTabPage Then
                    iStartPage = 1
                    iCurrentPage = iCurrentPage + 1
                  Else
                    ' If we have reached the maximum number of pages allowed then flag this to the user.
                    fOK = False
'                    Set WaitWindow = Nothing
                    gobjProgress.Visible = False
                    MsgBox "Unable to add more than " & Trim(Str(giMAXTABS)) & " page tabs." & vbCr & vbCr & _
                      Err.Description, vbExclamation + vbOKOnly, App.ProductName
                  End If
                End If
              End If
            
              If fOK Then
                ' Create a new page.
                If AddTabPage Then
                  iCurrentPage = iCurrentPage + 1
                  lngYPosition = gLngAUTOFORMATYSTART
                  Set varCurrentPageContainer = picPageContainer(tabPages.Tabs(iCurrentPage).Tag)
                      
                  ' Move the control onto the new page.
                  Set .Container = varCurrentPageContainer
                  .Top = AlignY(lngYPosition)
                Else
                  ' If we have reached the maximum number of pages allowed then flag this to the user.
                  fOK = False
'                  Set WaitWindow = Nothing
                  gobjProgress.Visible = False
                  MsgBox "Unable to add more than " & Trim(Str(giMAXTABS)) & " page tabs." & vbCr & vbCr & _
                    Err.Description, vbExclamation + vbOKOnly, App.ProductName
                End If
              End If
            End If
          End If
                    
          ' Disassociate object variables.
          Set ctlControl = Nothing
          
        End With
        
        ' If everything is still okay check if we need to add a label for the column control.
        If fOK Then
          
          ' Format a label for the new control if it is required.
          If ScreenControl_NeedsLabelling(iControlType) Then
            
            Set ctlLabelControl = AddControl(giCTRL_LABEL)
            fOK = Not (ctlLabelControl Is Nothing)
  
            If fOK Then
        
              With ctlLabelControl
                Set .Container = varCurrentPageContainer
                .Left = AlignX(gLngAUTOFORMATLABELCOLUMN)
                .Top = AlignY(lngYPosition)
                .ColumnID = 0
            
                ' Initialise the new label control's font and forecolour.
                Set objFont = New StdFont
                objFont.Name = Me.Font.Name
                objFont.Size = Me.Font.Size
                objFont.Bold = Me.Font.Bold
                objFont.Italic = Me.Font.Italic
                Set .Font = objFont
                Set objFont = Nothing
                .ForeColor = gDfltForeColour
                .Caption = objMisc.StrReplace(rsColumns!ColumnName, "_", " ", False) & " :"
            
                ' Default the control's propertes.
                fOK = AutoSizeControl(ctlLabelControl)
                .Visible = True
                
              End With
            End If
          
            Set ctlLabelControl = Nothing
          
          End If
              
          lngYPosition = lngYPosition + sngControlHeight + gLngAUTOFORMATYOFFSET
                
        End If
      End If
      
      rsColumns.MoveNext
        
      ' See if the user has cancelled the operation.
'      If Not WaitWindow Is Nothing Then
      If gobjProgress.Visible Then
        ' Update the progress bar.
        gobjProgress.UpdateProgress False
        
'        If WaitWindow.Cancelled Then
        If gobjProgress.Cancelled = True Then
          fOK = False
        End If
      End If
      
    Loop
            
    ' Close the recordset.
    rsColumns.Close
    
  End If
  
  ' Set focus on the first 'autoformatted' page tab.
  If fOK Then
    PageNo = iStartPage
  
    ' Set the last action flag and enable the Undo menu option.
    giLastActionFlag = giACTION_AUTOFORMAT
    giUndo_TabPageIndex = iStartPage
    frmSysMgr.RefreshMenu
  End If
  
TidyUpAndExit:
  ' Mark the screen as having changed.
  gfChangedScreen = True
  
  ' Disassociate all object variables.
  Set varCurrentPageContainer = Nothing
  Set objMisc = Nothing
  Set objFont = Nothing
  Set ctlControl = Nothing
  Set ctlLabelControl = Nothing
  Set rsColumns = Nothing
'  Set WaitWindow = Nothing
  gobjProgress.CloseProgress
  ' Reset the mousepointer.
  Screen.MousePointer = vbDefault
  ' Unlock the window to show the modifications.
  UI.UnlockWindow
  gfLoading = False
  Exit Sub
  
ErrorTrap:
'  Set WaitWindow = Nothing
  gobjProgress.Visible = False
  MsgBox "Failed to AutoFormat the screen." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, Application.Name
  Resume TidyUpAndExit
  
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim VarPageContainer As Variant

  ' Used to work out where to paste controls
  mlngMouseX = X
  mlngMouseY = Y

  ' Only handle left button presses here.
  If Button <> vbLeftButton Then
    Exit Sub
  End If
  
  ' Deselect all screen controls unless the SHIFT or CTRL keys are pressed.
  If ((Shift And vbShiftMask) = 0) And ((Shift And vbCtrlMask) = 0) Then
    fOK = DeselectAllControls
  Else
    'JDM - 16/08/01 - Fault 2455 - Was not setting this flag, and hence could multi-select add...
    fOK = True
  End If
  
  If fOK Then
    ' Start the multi-selection frame.
    gfMultiSelecting = True
    gLngMultiSelectionXStart = X
    gLngMultiSelectionYStart = Y
      
    Set VarPageContainer = CurrentPageContainer
    
    ' Position and display the multi-selection box.
    With asrboxMultiSelection
      .Left = gLngMultiSelectionXStart
      .Top = gLngMultiSelectionYStart
      .Width = 0
      .Height = 0
      Set .Container = VarPageContainer
      .Visible = True
      .ZOrder 0
    End With
  
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set VarPageContainer = Nothing
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit
  
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Position and size the multi-selection lines as required.
  On Error GoTo ErrorTrap
  
  Dim lngTop As Long
  Dim lngLeft As Long
  Dim lngRight As Long
  Dim lngBottom As Long
  Dim lngRightLimit As Long
  Dim lngBottomLimit As Long
  
  If gfMultiSelecting Then

    ' Calculate the cordinates of the multi-selection area.
    If X < gLngMultiSelectionXStart Then
      lngLeft = X
      lngRight = gLngMultiSelectionXStart
    Else
      lngLeft = gLngMultiSelectionXStart
      lngRight = X
    End If
      
    If Y < gLngMultiSelectionYStart Then
      lngTop = Y
      lngBottom = gLngMultiSelectionYStart
    Else
      lngTop = gLngMultiSelectionYStart
      lngBottom = Y
    End If

    ' Limit the multi-selection area to the form or tab page area.
    If tabPages.Tabs.Count > 0 Then
      lngRightLimit = picPageContainer(picPageContainer.UBound).Width - XBorder
      lngBottomLimit = picPageContainer(picPageContainer.UBound).Height - YBorder
    Else
      lngRightLimit = Me.Width - (2 * XFrame) - XBorder
      lngBottomLimit = Me.Height - (2 * YFrame) - CaptionHeight - (4 * YBorder)
    End If
      
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

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
  Dim ctlControl As VB.Control
  Dim VarPageContainer As Variant
  Dim bInSelectionBand As Boolean
  
  Select Case Button
    
    ' Handle left button presses.
    Case vbLeftButton
     
      'TM20010914 Fault
      'Was showing progressbar even on the right click.
      
      'JDM - Fault 2454 - Put up an hourglass
      Screen.MousePointer = vbHourglass
    
      'Open a progress bar
'      With gobjProgress
'        .Caption = "Screen Designer"
''          .NumberOfBars = 1
'        .Bar1Value = 0
'        .Bar1MaxValue = ScreenControlsCount(True)
'        .Bar1Caption = "Selecting Screen Controls..."
'        .Cancel = False
'        .Time = False
'        .OpenProgress
'      End With
      
      ' End the multi-selection.
      gfMultiSelecting = False
    
      Set VarPageContainer = asrboxMultiSelection.Container
      
      ' Hide the multi-selection box and move it onto the form.
      ' NB. This is done so that it is not left on any tabpage containers, thus
      ' making it impossible to unload the tab pages.
      With asrboxMultiSelection
        sngSelectionTop = .Top
        sngSelectionBottom = .Top + .Height
        sngSelectionLeft = .Left
        sngSelectionRight = .Left + .Width
        Set .Container = Me
        .Visible = False
      End With
      
      ' Lock the window refresh.
'      UI.LockWindow Me.hWnd
      
      ' Select thr highlighted controls
      ' JDM - Fault 63 - Restructured if statements to make a little faster on multi-selecting
      ' JDM - 13/08/01 - Commented below because selction markers are now on the
      '                  same level, so no need to refresh??? (Hopefully)
      For Each ctlControl In Me.Controls
        'TM20010914 Fault 1753
        'The ActiveBar control does mot have the visible property, so to avoid err
        'we only check the visible property of other controls.
        If ctlControl.Name <> "abScreen" Then
          If ctlControl.Visible Then
          
            'if Not ctlControl.Name = "asrDummyFrame" Then
          
              If IsScreenControl(ctlControl) Then
                With ctlControl
                  
                  fInVerticalBand = (.Left < sngSelectionRight) And (.Left + .Width > sngSelectionLeft)
                  fInHorizontalBand = (.Top < sngSelectionBottom) And (.Top + .Height > sngSelectionTop)
                  
                  
                  ' Only include the frame if the rubber band crosses a line (i.e. skip if only controls within frame are selected)
                  If ctlControl.Name = "asrDummyFrame" Then
                    
                    'If band is entiterly within selection band dont select the frame
                    bInSelectionBand = IIf((sngSelectionLeft > .Left) And (sngSelectionRight < .Left + .Width) _
                      And (sngSelectionTop > .Top) And (sngSelectionBottom < .Top + .Height), False, fInVerticalBand And fInHorizontalBand)
                    
                  Else
                    bInSelectionBand = fInVerticalBand And fInHorizontalBand
                  End If
                  
                  If bInSelectionBand Then 'fInHorizontalBand And fInVerticalBand Then

                    ' JDM - 20/08/02 - Fault 4309 - Holding down control now deselects controls
                    If ((Shift And vbCtrlMask) = 2) And .Selected Then
                      DeselectControl ctlControl
                      .Selected = False
                    Else
                      .Selected = True
                      SelectControl ctlControl
                    End If
                    
                  End If

                End With
              End If
                          
            'End If
            
          End If
        End If

      Next ctlControl
     
      ' Disassociate object variables.
      Set ctlControl = Nothing
      Set VarPageContainer = Nothing
  
      ' Unlock the window refresh.
      'UI.UnlockWindow
      
      ' Mark the screen as having changed.
'      gfChangedScreen = True
      frmSysMgr.RefreshMenu
      
      ' Refresh the properties screen.
      Set frmScrObjProps.CurrentScreen = Me
      frmScrObjProps.RefreshProperties
 
      ' Handle right button presses.
      Case vbRightButton
        UI.GetMousePos lngXMouse, lngYMouse
'        frmSysMgr.tbMain.PopupMenu "ID_mnuScreenEdit", ssPopupMenuLeftAlign, lngXMouse, lngYMouse
        frmSysMgr.tbMain.Bands("ID_mnuScreenEdit").TrackPopup -1, -1
  End Select
  
TidyUpAndExit:
  
  ' Close the progress bar
  gobjProgress.CloseProgress
  
  ' Disassociate object variables.
  Set ctlControl = Nothing
  Set VarPageContainer = Nothing
  
  ' Reset the screen mousepointer.
  Screen.MousePointer = vbDefault
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' If the screen has been modified then prompt the user
  ' whether or not to save the changes.
  On Error GoTo ErrorTrap
  
  If gfChangedScreen Then
    
   'RH24022000 - Fix if Logitech mouse thumb button is set to 'Close Application'
     
    Select Case MsgBox("Apply screen changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
        mblnDisplayScrOpen = False
      Case vbYes
        Cancel = (Not SaveScreen())
        If Cancel = True Then mblnDisplayScrOpen = False Else mblnDisplayScrOpen = True
      Case vbNo
        mblnDisplayScrOpen = True
    End Select
  End If

  ' Set the flag that determines whether we need to display the screen manager
  ' after the screen designer is unloaded.
  gfExitToScrMgr = (UnloadMode = vbFormControlMenu) And mblnDisplayScrOpen
  If Not gfChangedScreen Then gfExitToScrMgr = True
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit
End Sub

Private Sub Form_Resize()
  ' Resize the form.
  On Error GoTo ErrorTrap
  
  Dim lngMinHeight As Long
  Dim lngMinWidth As Long

  ' Only perform the resize method if the form is not minimized.
  If Me.WindowState <> vbMinimized Then
    
    ' Resize the tabstrip if it is visible.
    If tabPages.Tabs.Count > 0 Then
      TabPages_Resize
    End If
  End If

  '# RH 25/04/00 Bug Fix....resizing form did not flag the screen as changed, but
  '              only set flag if form is visible (ie, size change by user not
  '              when initially loading the screen).
  If Me.Visible = True Then gfChangedScreen = True
  
  ' JPD this is required so that the window state menu is refreshed.
  ' However it makes everything flash so I'd like to change it.
  frmSysMgr.RefreshMenu

  ' Get rid of the icon off the form
  RemoveIcon Me

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error resizing Screen Designer form." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Sub

Private Function TabPages_Resize() As Boolean
  ' Resize the tab pages.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim ctlPictureBox As PictureBox
  
  ' Position and size the tabstrip to fill the form's client area.
  tabPages.Move XFrame, YFrame, Me.ScaleWidth - (XFrame * 2), _
    Me.ScaleHeight - (YFrame * 2)
  
  ' Position and size the picture box containers of the tabstrip.
  For Each ctlPictureBox In picPageContainer
    ctlPictureBox.Move tabPages.ClientLeft, tabPages.ClientTop, _
      tabPages.ClientWidth, tabPages.ClientHeight
  Next ctlPictureBox
          
  fOK = True
  
TidyUpAndExit:
  ' Disassociate object variales.
  Set ctlPictureBox = Nothing
  TabPages_Resize = fOK
  Exit Function

ErrorTrap:
  MsgBox "Error resizing Screen Designer tab pages." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Function

Public Function XFrame() As Double
  ' Return the width of a control frame.
  
  If UI.GetOSVersion = 6 Then
    XFrame = 4 * Screen.TwipsPerPixelX
  Else
    XFrame = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
  End If

End Function

Public Function YFrame() As Double
  ' Return the height of a control frame.
  YFrame = UI.GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY

End Function
Private Function XBorder() As Double
  ' Return the width of a control border.
  XBorder = UI.GetSystemMetrics(SM_CXBORDER) * Screen.TwipsPerPixelX

End Function

Private Function YBorder() As Double
  ' Return the height of a control border.
  YBorder = UI.GetSystemMetrics(SM_CYBORDER) * Screen.TwipsPerPixelY

End Function
Private Function CaptionHeight() As Double
  ' Return the height of a form's caption bar.
  CaptionHeight = UI.GetSystemMetrics(SM_CYSMCAPTION) * Screen.TwipsPerPixelY

End Function

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ErrorTrap
  
  frmScrObjProps.ScreenCount = frmScrObjProps.ScreenCount - 1
  frmToolbox.ScreenCount = frmToolbox.ScreenCount - 1
  
  ' Display the Screen manager form if we are not exiting the system.
  If gfExitToScrMgr Then
  
    With frmSysMgr
      If .frmScrOpen Is Nothing Then
        Set .frmScrOpen = New SystemMgr.frmScrOpen
      End If
      
      .frmScrOpen.Show
      .frmScrOpen.SetFocus
      .RefreshMenu
    End With
    
  End If

  Unhook Me.hWnd

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Sub asrDummyFrame_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummyFrame(Index), Source, X, Y
  
End Sub

Private Sub asrDummyImage_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop asrDummyImage(Index), Source, X, Y
  
End Sub


Private Sub picPageContainer_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  On Error GoTo ErrorTrap
  
  If Not DropControl(picPageContainer(Index), Source, X, Y) Then
    MsgBox "Unable to drop the control." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit

End Sub

Private Sub picPageContainer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseDown event to the parent form.
  Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub picPageContainer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseMove event to the parent form.
  Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub picPageContainer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pass the MouseUp event to the parent form.
  Form_MouseUp Button, Shift, X, Y
End Sub

Private Function DeleteSelectedControls(Optional pbIsCutting As Boolean) As Boolean
  ' Delete the selected controls.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iIndex2 As Integer
  Dim avScreenControls() As Variant
  Dim ctlControl As VB.Control
  Dim iSelectedControls As Integer

  ' How many controls do we have
  iSelectedControls = SelectedControlsCount

 ' Open a progress bar
  With gobjProgress
    .Caption = "Screen Designer"
    .Bar1Value = 0
    .Bar1MaxValue = iSelectedControls
    .Bar1Caption = IIf(pbIsCutting, "Cutting", "Deleting") & " Screen Controls..."
    .Cancel = False
    .Time = False
    .OpenProgress
  End With
  
  fOK = True
  
  CurrentPageContainer.SetFocus
  
  ' Do nothing if there are no selected controls.
  If iSelectedControls > 0 Then
  
    ' Clear the array of deleted controls.
    For iIndex = 1 To UBound(gactlUndo_DeletedControls)
      Set ctlControl = gactlUndo_DeletedControls(iIndex)
      UnLoad ctlControl
      ' Disassociate object variables.
      Set ctlControl = Nothing
    Next iIndex
    ReDim gactlUndo_DeletedControls(0)
    
    ' Construct an array of the selected controls.
    ReDim avScreenControls(0)
    For Each ctlControl In Me.Controls
      If IsScreenControl(ctlControl) Then
        If ctlControl.Selected Then
          iIndex = UBound(avScreenControls) + 1
          ReDim Preserve avScreenControls(iIndex)
          Set avScreenControls(iIndex) = ctlControl
        End If
      End If
    Next ctlControl
    
    ' Disassociate object variables.
    Set ctlControl = Nothing

    ' Move all selected screen controls from the screen into the array of deleted controls.
    For iIndex = 1 To UBound(avScreenControls)
           
      Set ctlControl = avScreenControls(iIndex)

      iIndex2 = UBound(gactlUndo_DeletedControls) + 1
      ReDim Preserve gactlUndo_DeletedControls(iIndex2)
      Set gactlUndo_DeletedControls(iIndex2) = ctlControl

      With ctlControl
        If ctlControl.Tag > 0 Then
          
          ' Hide the selection markers
          ASRSelectionMarkers(ctlControl.Tag).Visible = False
          ASRSelectionMarkers(ctlControl.Tag).AttachedObject = gactlUndo_DeletedControls(iIndex2)
          
          ' Unload the control's selection markers.
          gobjProgress.UpdateProgress
          
          If Not fOK Then
            Exit For
          End If
        End If
  
        '.Tag = 0
        .Visible = False
        .Selected = False
      End With
      
      ' Disassociate object variables.
      Set ctlControl = Nothing
    Next iIndex

    ' Mark the screen as having changed.
    gfChangedScreen = True
    
    If fOK Then
      ' Set the last action flag and enable the Undo menu option.
      giLastActionFlag = giACTION_DELETECONTROLS
      giUndo_TabPageIndex = PageNo
      frmSysMgr.RefreshMenu
    End If
    
  End If
  
TidyUpAndExit:
  
  ' Close progress bar
  gobjProgress.CloseProgress
  
  ' Disassociate object variables.
  Set ctlControl = Nothing
  ' Return the success/failure value.
  DeleteSelectedControls = fOK
  Exit Function
  
ErrorTrap:
  ' Flag the error.
  fOK = False
  Resume TidyUpAndExit
  
End Function

' Return a count of the number of selected controls.
Public Function SelectedControlsCount() As Integer
  On Error GoTo ErrorTrap
  
  Dim iCount As Integer
  Dim iSelectedControls As Integer
'  Dim ctlControl As VB.Control
  
  ' Initialize the count.
  iSelectedControls = 0
  
  ' Count the number of custom screen controls that are selected.
'  For Each ctlControl In Me.Controls
'    If IsScreenControl(ctlControl) Then
'      If ctlControl.Selected Then
'        iCount = iCount + 1
'      End If
'    End If
'  Next ctlControl
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    iSelectedControls = iSelectedControls + IIf(ASRSelectionMarkers(iCount).Visible, 1, 0)
  Next iCount
        
TidyUpAndExit:
  SelectedControlsCount = iSelectedControls
  Exit Function
  
ErrorTrap:
  iSelectedControls = 0
  Resume TidyUpAndExit
  
End Function
Public Function ScreenControlsCount(Optional pbShowThisPageOnly As Boolean) As Integer
  ' Return a count of the number of selected controls.
  On Error GoTo ErrorTrap
  
  Dim iCount As Integer
  Dim ctlControl As VB.Control
  
  ' Initialize the count.
  iCount = 0
  
  ' Count the number of custom screen controls that are selected.
  For Each ctlControl In Me.Controls
    'TM20010914 Fault 1753
    'The ActiveBar control does mot have the visible property, so to avoid err
    'we only check the visible property of other controls.
    If ctlControl.Name <> "abScreen" Then
      If ctlControl.Visible Or Not pbShowThisPageOnly Then
        If IsScreenControl(ctlControl) Then
          iCount = iCount + 1
        End If
      End If
    End If
  Next ctlControl
        
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  ScreenControlsCount = iCount
  Exit Function
  
ErrorTrap:
  iCount = 0
  Resume TidyUpAndExit

End Function

Public Function ClipboardControlsCount() As Integer
  ' Return a count of the number of controls in the clipboard control.
  ClipboardControlsCount = UBound(gactlClipboardControls)
  
End Function


Private Function LoadScreen() As Boolean
  ' Load controls onto the screen.
  'On Error GoTo ErrorTrap
    
  Dim fLoadOk As Boolean
  Dim iPageNo As Integer
  Dim iCtrlType As Long
  Dim iDisplayType As Integer
  Dim lngTableID As Long
  Dim lngPictureID As Long
  Dim sFileName As String
  Dim sTableName As String
  Dim sColumnName As String
  Dim objFont As StdFont
  Dim ctlControl As VB.Control
  Dim iNextIndex As Integer
  Dim iRecordCount As Integer
  Dim iTabPages As Integer
  Dim iCount As Integer
  
  iNextIndex = 1

  If gLngScreenID = 0 Then
    LoadScreen = True
    Exit Function
  End If

  Screen.MousePointer = vbHourglass
  
  ' Find the screen definition in the database.
  With recScrEdit
    .Index = "idxScreenID"
    .Seek "=", gLngScreenID
    fLoadOk = (Not .NoMatch)
  End With
  
  ' Load the screen properties.
  If fLoadOk Then
    
    ' Lock the screen refeshing.
    UI.LockWindow Me.hWnd
    
    'Set form properties from screen definition
    With recScrEdit
      
      gfNewScreen = IIf(IsNull(.Fields("new")), True, .Fields("new"))
      gfChangedScreen = False
      
      ' Set the screen caption and size.
      SetFormCaption Me, "Screen Manager - " & IIf(IsNull(.Fields("name")), "unnamed", .Fields("name")) & vbNullString
      
      Me.Height = IIf(IsNull(.Fields("height")), gLngDFLTSCREENHEIGHT, .Fields("height") + 450) + IIf(UI.GetOSVersion = 6, 240, 0)
      Me.Width = IIf(IsNull(.Fields("width")), gLngDFLTSCREENWIDTH, .Fields("width") + IIf(UI.GetOSVersion = 6, 240, 0))
           
      gLngTableID = IIf(IsNull(.Fields("tableID")), 0, .Fields("tableID"))
      giGridX = IIf(IsNull(.Fields("gridX")), 40, .Fields("gridX"))
      giGridY = IIf(IsNull(.Fields("gridY")), 40, .Fields("gridY"))
      gfAlignToGrid = IIf(IsNull(.Fields("alignToGrid")), True, .Fields("alignToGrid"))
      
      ' Read the default foreground colour and font options.
      gDfltForeColour = IIf(IsNull(.Fields("dfltForeColour")), vbBlack, .Fields("dfltForeColour"))
            
      Me.Font.Name = IIf(IsNull(.Fields("dfltFontName")), gobjDefaultScreenFont.Name, .Fields("dfltFontName"))
      Me.Font.Size = IIf(IsNull(.Fields("dfltFontSize")), gobjDefaultScreenFont.Size, .Fields("dfltFontSize"))
      Me.Font.Bold = IIf(IsNull(.Fields("dfltFontBold")), False, .Fields("dfltFontBold"))
      Me.Font.Italic = IIf(IsNull(.Fields("dfltFontItalic")), False, .Fields("dfltFontItalic"))
       
      ' Set tabstrip properties.
      tabPages.Font.Name = IIf(IsNull(.Fields("fontName")), DefaultFont.Name, .Fields("fontName"))
      tabPages.Font.Size = IIf(IsNull(.Fields("fontSize")), DefaultFont.Size, .Fields("fontSize"))
      tabPages.Font.Bold = IIf(IsNull(.Fields("fontBold")), DefaultFont.Bold, .Fields("fontBold"))
      tabPages.Font.Italic = IIf(IsNull(.Fields("fontItalic")), DefaultFont.Italic, .Fields("fontItalic"))
      tabPages.Font.Strikethrough = IIf(IsNull(.Fields("fontStrikeThru")), DefaultFont.Strikethrough, .Fields("fontStrikeThru"))
      tabPages.Font.Underline = IIf(IsNull(.Fields("fontUnderline")), DefaultFont.Underline, .Fields("fontUnderline"))
      tabPages.TabIndex = 0
      iTabPages = 0
      
       ' Add tab pages if required.
      With recPageCaptEdit
        .Index = "idxScreenPage"
        .Seek "=", gLngScreenID, 1
        
        If Not .NoMatch Then
        
          'Loop through Page Definitions and add pages to the screen.
          Do While Not .EOF
              
            If .Fields("screenID") <> gLngScreenID Then
              Exit Do
            End If
            
            fLoadOk = AddTabPage
            If fLoadOk Then
              'MH20020527 Fault 1862
              'tabPages.Tabs(tabPages.Tabs.Count).Caption = IIf(IsNull(.Fields("caption")), "", .Fields("caption"))
              tabPages.Tabs(tabPages.Tabs.Count).Caption = IIf(IsNull(.Fields("caption")), "", Replace(.Fields("caption"), "&", "&&"))
            End If
            
            .MoveNext
            
            iTabPages = iTabPages + 1
          
          Loop
          
          ' Select the first page.
          PageNo = 1

        End If
      End With
    End With
  End If
  
  ' Load the controls for each tab page
  If iTabPages = 0 Then
    LoadTabPage 0
  End If

  IsChanged = False
  
TidyUpAndExit:

  ' Unlock the window refreshing.
  UI.UnlockWindow
  
  ' Position the form.
  Me.Top = Int((Forms(0).ScaleHeight - Me.Height) / 2)
  Me.Left = Int((Forms(0).ScaleWidth - Me.Width) / 2)
    
  ' Reset the screen moousepointer.
  Screen.MousePointer = vbDefault
  
  LoadScreen = fLoadOk
  Exit Function
  
ErrorTrap:
  fLoadOk = False
  MsgBox "Error loading Screen." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function

Private Function ReadColumnControlValues(plngColumnID As Long) As Variant
  ' Return an array of control values for the given column.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  Dim avValues As Variant
  Dim asResults() As String
  Dim sSQL As String
  Dim rsControlValues As DAO.Recordset
  
  ' Pull the column control values from the database.
  sSQL = "SELECT value" & _
    " FROM tmpControlValues" & _
    " WHERE columnID = " & plngColumnID & _
    " ORDER BY sequence"
  Set rsControlValues = daoDb.OpenRecordset(sSQL, dbOpenSnapshot, dbReadOnly)
  
  ' Load the control values into an array
  'avValues = rsControlValues.GetRows(100)
  avValues = rsControlValues.GetRows(rsControlValues.RecordCount)

  ' Copy the required values from the 2-dimensional variant array, into
  ' a 1-dimensional string array.
  ReDim asResults(UBound(avValues, 2))
  For iLoop = LBound(avValues, 2) To UBound(avValues, 2)
    asResults(iLoop) = CStr(avValues(0, iLoop))
  Next iLoop

TidyUpAndExit:
  rsControlValues.Close
  Set rsControlValues = Nothing
  ReadColumnControlValues = asResults
  Exit Function
  
ErrorTrap:
  MsgBox "Error reading column control values." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  ReDim asResults(0)
  Resume TidyUpAndExit
  
End Function

Public Property Get TableID() As Long
  ' Return the screen's table id.
  TableID = gLngTableID
  
End Property

Public Property Let TableID(plngNewValue As Long)
  ' Return the screen's table id.
  gLngTableID = plngNewValue

End Property

Public Property Get AlignToGrid() As Boolean
  ' Return the value of the 'align to grid' property.
  AlignToGrid = gfAlignToGrid
  
End Property

Public Property Let AlignToGrid(ByVal pfAlignToGrid As Boolean)
  ' Set the value of the 'align to grid' property.
  gfAlignToGrid = pfAlignToGrid
  
End Property

Private Function CopySelectedControls() As Boolean
  ' Copy the selected controls to the clipboard array.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iControlType As Long
  Dim ctlSourceControl As VB.Control
  Dim ctlCopiedControl As VB.Control
  
  ' Do nothing if no controls are selected.
  If SelectedControlsCount = 0 Then
    CopySelectedControls = True
    Exit Function
  End If
  
  ' Clear the array of copied controls.
  For iIndex = 1 To UBound(gactlClipboardControls)
    Set ctlCopiedControl = gactlClipboardControls(iIndex)
    UnLoad ctlCopiedControl
  Next iIndex
  ReDim gactlClipboardControls(0)
  ' Disassociate object variables.
  Set ctlCopiedControl = Nothing
  
  ' Create a copy of each selected control in the array.
  For Each ctlSourceControl In Me.Controls
    If IsScreenControl(ctlSourceControl) Then
      If ctlSourceControl.Selected Then
      
        iControlType = ScreenControl_Type(ctlSourceControl)
        
        ' Create a new instance of the required control type.
        Set ctlCopiedControl = AddControl(iControlType)
        
        fOK = Not (ctlCopiedControl Is Nothing)
        
        If fOK Then
          ' Copy the properties from the selected control to the new control.
          fOK = CopyControlProperties(ctlSourceControl, ctlCopiedControl)
          
          iIndex = UBound(gactlClipboardControls) + 1
          ReDim Preserve gactlClipboardControls(iIndex)
          Set gactlClipboardControls(iIndex) = ctlCopiedControl
        Else
          Exit For
        End If
        
        Set ctlCopiedControl = Nothing
      
      End If
    End If
  Next ctlSourceControl

TidyUpAndExit:
  If Not fOK Then
    ' Clear the array of copied controls.
    For iIndex = 1 To UBound(gactlClipboardControls)
      Set ctlCopiedControl = gactlClipboardControls(iIndex)
      UnLoad ctlCopiedControl
    Next iIndex
    ReDim gactlClipboardControls(0)
  End If
  ' Disassociate object variables.
  Set ctlSourceControl = Nothing
  Set ctlCopiedControl = Nothing
  CopySelectedControls = fOK
  frmSysMgr.RefreshMenu
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Private Function CopyControlProperties(pCtlSource As VB.Control, pCtlDestination As VB.Control) As Boolean
  ' Copy the properties from the pCtlSource control to the pCtlDestination control.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iControlType As Long
  Dim sFileName As String
  
  ' Get the given control's type.
  iControlType = ScreenControl_Type(pCtlSource)

  With pCtlDestination
  
    ' Copy the source control's properties to the destination control.
    If ScreenControl_HasAlignment(iControlType) Or ScreenControl_HasOrientation(iControlType) Then
      .Alignment = pCtlSource.Alignment
    End If
    
    If ScreenControl_HasBackColor(iControlType) Then
      .BackColor = pCtlSource.BackColor
    End If
     
    If ScreenControl_HasBorderStyle(iControlType) Then
      .BorderStyle = pCtlSource.BorderStyle
    End If
    
    
    'NPG20071030
    If ScreenControl_HasReadOnly(iControlType) Then
      .Read_Only = pCtlSource.Read_Only
    End If

    If ScreenControl_HasCaption(iControlType) Or _
      ScreenControl_HasText(iControlType) Then
      .Caption = pCtlSource.Caption
    End If
    
    If ScreenControl_HasFont(iControlType) Then
      Set .Font = pCtlSource.Font
    End If
    
    If ScreenControl_HasForeColor(iControlType) Then
      .ForeColor = pCtlSource.ForeColor
    End If
    
    If ScreenControl_HasOptions(iControlType) Then
'      .Options = pCtlSource.Options

      ' JDM - 14/08/01 - Fault 1949 - Changed to read control values
      .SetOptions ReadColumnControlValues(pCtlSource.ColumnID)
    End If
        
    If ScreenControl_HasNavigation(iControlType) Then
      .DisplayType = pCtlSource.DisplayType
      .ColumnName = pCtlSource.ColumnName
      .NavigateTo = pCtlSource.NavigateTo
      .NavigateIn = pCtlSource.NavigateIn
      .NavigateOnSave = pCtlSource.NavigateOnSave
    End If
    
    If ScreenControl_HasPicture(iControlType) Then
      .PictureID = pCtlSource.PictureID
      If .PictureID > 0 Then
        recPictEdit.Index = "idxID"
        recPictEdit.Seek "=", .PictureID
                    
        If Not recPictEdit.NoMatch Then
          sFileName = ReadPicture
          .Picture = sFileName
          Kill sFileName
        End If
      Else
        .Picture = "No picture"
      End If
    End If
  
    ' Copy the source control's position and dimension's to the destination control.
    .ColumnID = pCtlSource.ColumnID
    .Top = pCtlSource.Top
    .Left = pCtlSource.Left
    .Height = pCtlSource.Height
    .Width = pCtlSource.Width
  
    .ToolTipText = pCtlSource.ToolTipText

    ' Force the value of some of the destination control's properties.
    .Selected = False
    .Visible = False
    
  End With
  
  fOK = True
  
TidyUpAndExit:
  CopyControlProperties = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function PasteControls() As Boolean
  ' Paste the controls from the clipboard onto the current page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iIndex2 As Integer
  Dim iControlType As Long
  Dim lngXOffset As Long
  Dim lngYOffset As Long
  Dim ctlControl As VB.Control
  Dim ctlNewControl As VB.Control
  Dim VarPageContainer As Variant
  
  ' Do nothing if there's nothing in the clipboard.
  If ClipboardControlsCount = 0 Then
    PasteControls = True
    Exit Function
  End If
  
 ' Open a progress bar
  With gobjProgress
    .Caption = "Screen Designer"
    .Bar1Value = 0
    .Bar1MaxValue = ClipboardControlsCount
    .Bar1Caption = "Pasting Controls..."
    .Cancel = False
    .Time = False
    .OpenProgress
  End With
  
  ' Lock the forms refreshing.
  UI.LockWindow Me.hWnd
  
  ' Get the current page container.
  Set VarPageContainer = CurrentPageContainer
  
  ' Get the offset for the new positions of the controls.
  lngXOffset = VarPageContainer.Width
  lngYOffset = VarPageContainer.Height
  
  For iIndex = 1 To UBound(gactlClipboardControls)
    Set ctlControl = gactlClipboardControls(iIndex)
    With ctlControl
      If .Left < lngXOffset Then
        lngXOffset = .Left
      End If
      If .Top < lngYOffset Then
        lngYOffset = .Top
      End If
    End With
  Next iIndex
  
  Set ctlControl = Nothing
  
  ' Deselect all existing controls.
  fOK = DeselectAllControls
  
  If fOK Then
  
    ReDim gavUndo_PastedControls(2, 0)
  
    ' Drop each control from the clipboard onto the current page.
    For iIndex = 1 To UBound(gactlClipboardControls)
    
      Set ctlControl = gactlClipboardControls(iIndex)
      
      ' Add the required control type.
      iControlType = ScreenControl_Type(ctlControl)
      Set ctlNewControl = AddControl(iControlType)
      
      fOK = Not (ctlNewControl Is Nothing)
      If fOK Then
      
        fOK = CopyControlProperties(ctlControl, ctlNewControl)
    
        If fOK Then
          With ctlNewControl
            Set .Container = VarPageContainer
            .Left = AlignX(.Left - lngXOffset)
            .Top = AlignY(.Top - lngYOffset)
            
            fOK = SelectControl(ctlNewControl)
            
            If fOK Then
              .Visible = True
            
              iIndex2 = UBound(gavUndo_PastedControls, 2) + 1
              ReDim Preserve gavUndo_PastedControls(2, iIndex2)
              gavUndo_PastedControls(1, iIndex2) = .Name
              gavUndo_PastedControls(2, iIndex2) = .Index
            End If
          End With
        End If
      End If
      
      If Not fOK Then
        Exit For
      End If
      
      'Update the progress bar
      gobjProgress.UpdateProgress
      
    Next iIndex
  End If

  If fOK Then
    ' Mark the screen as having changed.
    gfChangedScreen = True

    ' Set the last action flag and enable the Undo menu option.
    giLastActionFlag = giACTION_PASTECONTROLS
    frmSysMgr.RefreshMenu
  End If

TidyUpAndExit:
  
  ' Unlock the forms refreshing.
  UI.UnlockWindow
  
  'Close the progress bar
  gobjProgress.CloseProgress
  
  ' Disassociate object variables.
  Set ctlControl = Nothing
  Set VarPageContainer = Nothing
  PasteControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function AddTabPage(Optional piTabPageIndex As Integer) As Boolean
  ' Add a tab to the page. If none exist then move all existing controls onto
  ' the new tab.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fControlsMoved As Boolean
  Dim iContainerIndex As Integer
  Dim ctlControl As VB.Control
  Dim iCount As Integer
  
  ' Do not exceed the maximum number of pages.
  If tabPages.Tabs.Count = giMAXTABS Then
    ' Flag the error to the user if we are not just loading the screen.
    If Not gfLoading Then
      MsgBox "Unable to add more than " & Trim(Str(giMAXTABS)) & " page tabs."
    End If
    
    AddTabPage = False
    Exit Function
  End If
  
  ' Get the index of the new tab page.
  If (IsMissing(piTabPageIndex)) Or (piTabPageIndex = 0) Then
    piTabPageIndex = tabPages.Tabs.Count + 1
  ElseIf (piTabPageIndex > tabPages.Tabs.Count + 1) Then
    piTabPageIndex = tabPages.Tabs.Count + 1
  End If
  
  ' Get the index of the new page's pictureBox container.
  iContainerIndex = picPageContainer.UBound + 1
  
  ' If we are adding the first tab page then move all existing controls onto this page
  If tabPages.Tabs.Count = 0 Then
  
    ' Add the new tab, and initialise its caption.
    tabPages.Tabs.Add
    tabPages.Tabs(1).Caption = "Page 1"
    
    ' Add a picture control to the tab page to contain its controls, and initialise its properties.
    Load picPageContainer(iContainerIndex)
    With picPageContainer(iContainerIndex)
      .BackColor = Me.BackColor
      .BorderStyle = 0
      .Left = tabPages.ClientLeft
      .Top = tabPages.ClientTop
      .Width = tabPages.ClientWidth
      .Height = tabPages.ClientHeight
      .Visible = False
      .ZOrder 1
    End With
    
    ' Move all screen controls onto the new tab page's picture container.
    GetControlLevel (Me.hWnd)

    fControlsMoved = False
    For Each ctlControl In Me.Controls
      If IsScreenControl(ctlControl) Then
        Set ctlControl.Container = picPageContainer(iContainerIndex)
        fControlsMoved = True
      End If
    Next ctlControl
    ' Disassociate object variables.
    Set ctlControl = Nothing

    ' Ensure that the z-order of the controls is the same as before.
    SetControlLevel

    ' If we moving controls from the form onto the new tabpage then increase the
    ' form dimensions to allow for the tabs.
    If fControlsMoved Then
      With Me
        .Height = .Height + (tabPages.Height - tabPages.ClientHeight) + (2 * YFrame)
        .Width = .Width + (4 * XFrame)
      End With
    
      ' JDM - 22/08/02 - Fault 4264 - Refresh the selection markers
      For iCount = 1 To ASRSelectionMarkers.Count - 1
        With ASRSelectionMarkers(iCount)
          Set .Container = picPageContainer(iContainerIndex)
        End With
      Next iCount
    
    End If
    
  Else
    ' Add the new tab.
    tabPages.Tabs.Add piTabPageIndex, , "Page " & tabPages.Tabs.Count + 1
    
    ' Add a picture control to the tab page to contain its controls, and initialise its properties.
    Load picPageContainer(iContainerIndex)
    With picPageContainer(iContainerIndex)
      .BackColor = Me.BackColor
      .BorderStyle = 0
      .Left = tabPages.ClientLeft
      .Top = tabPages.ClientTop
      .Width = tabPages.ClientWidth
      .Height = tabPages.ClientHeight
      .Visible = False
      .ZOrder 1
    End With

  End If
  
  ' Set the 'tag' property of the tab page. We use to relate a tab page
  ' with its associated picture container control.
  tabPages.Tabs(piTabPageIndex).Tag = iContainerIndex

  ' Display the tab strip.
  fOK = TabPages_Resize
  tabPages.Visible = True
  
  ' Select the new page if we are not just loading the screen.
  If Not gfLoading Then
    PageNo = piTabPageIndex
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  AddTabPage = fOK
  Exit Function

ErrorTrap:
  MsgBox "Error adding tab page." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function DeleteTabPage(piTabIndex As Integer, pfPromptUser As Boolean) As Boolean
  ' Delete the current tab page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fConfirmed As Boolean
  Dim iTag As Integer
  Dim iIndex As Integer
  Dim iIndex2 As Integer
  Dim ctlControl As VB.Control
  Dim actlScreenControls() As VB.Control
  Dim ctlPageContainer As VB.PictureBox
  Dim strCaption As String
  
  fOK = True
  
  CurrentPageContainer.SetFocus
  
  ' Get the given tab page's container control.
  Set ctlPageContainer = picPageContainer(tabPages.Tabs(piTabIndex).Tag)
  strCaption = tabPages.Tabs(tabPages.SelectedItem.Index).Caption
    
  ' Construct an array of the given tab page's screen controls.
  ReDim actlScreenControls(0)
  For Each ctlControl In Me.Controls
    If IsScreenControl(ctlControl) Then
      If ctlControl.Container Is ctlPageContainer Then
        iIndex = UBound(actlScreenControls) + 1
        ReDim Preserve actlScreenControls(iIndex)
        Set actlScreenControls(iIndex) = ctlControl
      End If
    End If
  Next ctlControl
  ' Disassociate object variables.
  Set ctlControl = Nothing

  ' Prompt the user for confirmation if the page contains controls.
  If (UBound(actlScreenControls) > 0) And (pfPromptUser) Then
    fConfirmed = (MsgBox("The page '" & strCaption & "' contains controls." & _
      vbCr & vbCr & "Are you sure you want to delete it ?", _
      vbQuestion + vbYesNo, Me.Caption) = vbYes)
  Else
    fConfirmed = True
  End If
    
  If fConfirmed Then
    ' Clear the array of deleted controls.
    For iIndex = 1 To UBound(gactlUndo_DeletedControls)
      Set ctlControl = gactlUndo_DeletedControls(iIndex)
      UnLoad ctlControl
      Set ctlControl = Nothing
    Next iIndex
    ReDim gactlUndo_DeletedControls(0)

    ' Delete all controls on this page.
    For iIndex = 1 To UBound(actlScreenControls)
      Set ctlControl = actlScreenControls(iIndex)
      
      With ctlControl
        iTag = val(.Tag)
      
        If iTag > 0 Then
          ' Unload the control's selection markers.
          fOK = True
          
          If Not fOK Then
            Exit For
          End If
        End If

        '.Tag = 0
        .Visible = False
        .Selected = False
        Set .Container = Me
      End With
      
      iIndex2 = UBound(gactlUndo_DeletedControls) + 1
      ReDim Preserve gactlUndo_DeletedControls(iIndex2)
      Set gactlUndo_DeletedControls(iIndex2) = ctlControl
    
      ' Disassociate object variables.
      Set ctlControl = Nothing
    Next iIndex

    If fOK Then
      ' Delete the page container.
      'UnLoad picPageContainer(tabPages.Tabs(piTabIndex).Tag)
      picPageContainer(piTabIndex).Visible = False
  
      ' Remember the tabpage caption.
      gsUndo_TabPageCaption = tabPages.Tabs(piTabIndex).Caption
      
      ' Remove the tab from the tabstrip.
      tabPages.Tabs.Remove piTabIndex
  
      ' Hide the tabstrip if we now have no tabs left.
      ' Otherwise select the first tab page.
      If tabPages.Tabs.Count = 0 Then
        tabPages.Visible = False
        PageNo = 0
      Else
        PageNo = 1
      End If
          
      ' Mark the screen as having changed.
      gfChangedScreen = True
      
      ' Set the last action flag and enable the Undo menu option.
      giLastActionFlag = giACTION_DELETETABPAGE
      giUndo_TabPageIndex = piTabIndex
      frmSysMgr.RefreshMenu
    End If
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  Set ctlPageContainer = Nothing
  ' Return the success/failure value.
  DeleteTabPage = fOK
  Exit Function
  
ErrorTrap:
  ' Flag the error.
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Property Let PageNo(piPageNumber As Integer)
  ' Set the tabstrip page number.
  On Error GoTo ErrorTrap
  
  Dim iPageTag As Integer
  Dim ctlPictureBox As PictureBox
  
  ' Do nothing if there are no tabpages.
  If tabPages.Tabs.Count > 0 Then
    
    ' If the given page number is not valid, just select the first page.
    If piPageNumber > tabPages.Tabs.Count Then
      piPageNumber = 1
    End If
    
    iPageTag = tabPages.Tabs(piPageNumber).Tag
    
    ' Position and size the picture box containers of the tabstrip.
    For Each ctlPictureBox In picPageContainer
      With ctlPictureBox
        If .Index = iPageTag Then
          .Enabled = True
          .Visible = True
          .ZOrder 0
        Else
          .Enabled = False
          .Visible = False
        End If
      End With
    Next ctlPictureBox
    
    tabPages.Tabs(piPageNumber).Selected = True
      
    ' If the page has changed then ensure that the old page
    ' controls are deselected.
    DeselectAllControls
  Else
    For Each ctlPictureBox In picPageContainer
      With ctlPictureBox
        .Enabled = False
        .Visible = False
      End With
    Next ctlPictureBox
  End If

  frmSysMgr.RefreshMenu
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlPictureBox = Nothing
  Exit Property
  
ErrorTrap:
  MsgBox "Error setting page number." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Property

Public Property Get PageNo() As Integer
  ' Return the current tabstrip page number.
  On Error GoTo ErrorTrap
  
  Dim iPageNo As Integer
  
  If tabPages.Tabs.Count = 0 Then
    iPageNo = 0
  Else
    iPageNo = tabPages.SelectedItem.Index
  End If
  
TidyUpAndExit:
  PageNo = iPageNo
  Exit Function

ErrorTrap:
  iPageNo = 0
  Resume TidyUpAndExit
  
End Property

Private Sub tabPages_Click()
  
  Dim iOldPage As Integer
  
  ' Select a tab page.
  Static fInClick As Boolean
  
  If Not fInClick Then
    fInClick = True
  
    tabPages.Enabled = False
    Screen.MousePointer = vbHourglass
  
    ' Set the active page.
    PageNo = tabPages.SelectedItem.Index
      
    ' Load the controls for this page
    LoadTabPage (tabPages.SelectedItem.Index)
      
    ' Refresh the properties screen.
    Set frmScrObjProps.CurrentScreen = Me
    frmScrObjProps.RefreshProperties
  
    tabPages.Enabled = True
    Screen.MousePointer = vbDefault
  
    fInClick = False
    
    
    
  End If
  

  
End Sub

Private Function DeleteControl(pctlControl As VB.Control) As Boolean
  ' Delete the given screen control.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  
  fOK = True
  
  ' Get the index of the given control.
  iIndex = val(pctlControl.Tag)
  
  ' Do not delete the control array dummy (index = 0).
  If pctlControl.Index = 0 Then
    DeleteControl = True
    Exit Function
  End If
  
  ' Hide the selection markers
  If Not pctlControl.Tag = "" Then
    ASRSelectionMarkers(pctlControl.Tag).Visible = False
  End If
  
  ' Unload the screen control.
  UnLoad pctlControl
  
  If iIndex > 0 Then
    ' Unload the control's selection markers.
    fOK = True
  End If
        
TidyUpAndExit:
  DeleteControl = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function AddControl(piControlType As Long) As VB.Control
  On Error GoTo ErrorTrap
  
  Select Case piControlType
  
    Case giCTRL_CHECKBOX
      Load asrDummyCheckBox(asrDummyCheckBox.UBound + 1)
      Set AddControl = asrDummyCheckBox(asrDummyCheckBox.UBound)
    
    Case giCTRL_COMBOBOX
      Load asrDummyCombo(asrDummyCombo.UBound + 1)
      Set AddControl = asrDummyCombo(asrDummyCombo.UBound)
    
    Case giCTRL_IMAGE
      Load asrDummyImage(asrDummyImage.UBound + 1)
      Set AddControl = asrDummyImage(asrDummyImage.UBound)
    
    Case giCTRL_OLE
      Load asrDummyOLEContents(asrDummyOLEContents.UBound + 1)
      Set AddControl = asrDummyOLEContents(asrDummyOLEContents.UBound)
      
    Case giCTRL_PHOTO
      Load asrDummyPhoto(asrDummyPhoto.UBound + 1)
      Set AddControl = asrDummyPhoto(asrDummyPhoto.UBound)
      
      ' AE20080519 Fault #13166
      AddControl.BackColor = vbWindowBackground
      
    Case giCTRL_OPTIONGROUP
      Load ASRDummyOptions(ASRDummyOptions.UBound + 1)
      Set AddControl = ASRDummyOptions(ASRDummyOptions.UBound)
      
    Case giCTRL_SPINNER
      Load asrDummySpinner(asrDummySpinner.UBound + 1)
      Set AddControl = asrDummySpinner(asrDummySpinner.UBound)
    
    Case giCTRL_TEXTBOX
      Load asrDummyTextBox(asrDummyTextBox.UBound + 1)
      Set AddControl = asrDummyTextBox(asrDummyTextBox.UBound)
    
    Case giCTRL_WORKINGPATTERN
      Load ASRCustomDummyWP(ASRCustomDummyWP.UBound + 1)
      Set AddControl = ASRCustomDummyWP(ASRCustomDummyWP.UBound)
    
    Case giCTRL_LABEL
      Load asrDummyLabel(asrDummyLabel.UBound + 1)
      Set AddControl = asrDummyLabel(asrDummyLabel.UBound)
      With AddControl
        .BorderStyle = vbBSNone
        .BackColor = Me.BackColor
      End With
    
    Case giCTRL_FRAME
      Load asrDummyFrame(asrDummyFrame.UBound + 1)
      Set AddControl = asrDummyFrame(asrDummyFrame.UBound)
      
    Case giCTRL_LINK
      Load asrDummyLink(asrDummyLink.UBound + 1)
      Set AddControl = asrDummyLink(asrDummyLink.UBound)
      
    Case giCTRL_LINE
      Load ASRDummyLine(ASRDummyLine.UBound + 1)
      Set AddControl = ASRDummyLine(ASRDummyLine.UBound)
      
    Case giCTRL_NAVIGATION
      Load ASRDummyNavigation(ASRDummyNavigation.UBound + 1)
      Set AddControl = ASRDummyNavigation(ASRDummyNavigation.UBound)
      
    Case giCTRL_COLOURPICKER
      Load ASRColourSelector(ASRColourSelector.UBound + 1)
      Set AddControl = ASRColourSelector(ASRColourSelector.UBound)
      
  End Select
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  MsgBox "Unable to load control type " & Trim(Str(piControlType)) & "." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, Application.Name
  Set AddControl = Nothing
  Resume TidyUpAndExit

  
End Function

Public Property Get ScreenID() As Long
  ' Return the current screen id.
  ScreenID = gLngScreenID
  
End Property

Public Property Let ScreenID(pLngNewID As Long)
  ' Set the current screen id.
  gLngScreenID = pLngNewID

  ' Load the screen.
  gfLoading = True
  If Not LoadScreen Then
    MsgBox "Unable to load screen." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  gfLoading = False

  ' RH 26/01/01 - BUG 987
  IsChanged = False
  
End Property

Private Sub tabPages_GotFocus()
  
  ' Do nothing if we are just activating the form.
  If gfActivating Then
    gfActivating = False
    Exit Sub
  End If
  
  ' Deselect all controls.
  If tabPages.Tabs.Count > 0 Then
    If PageNo <> tabPages.SelectedItem.Index Then
      DeselectAllControls
  
      ' Refresh the menu.
      frmSysMgr.RefreshMenu
      
      ' Refresh the properties screen.
      Set frmScrObjProps.CurrentScreen = Me
      frmScrObjProps.RefreshProperties
    End If
  End If

End Sub



Private Function CurrentPageContainer() As Variant
  ' Return the current page container.
  If tabPages.Tabs.Count > 0 Then
    Set CurrentPageContainer = picPageContainer(tabPages.SelectedItem.Tag)
  Else
    Set CurrentPageContainer = Me
  End If
  
End Function

Private Function CutSelectedControls() As Boolean
  ' Cut the selected controls.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Copy the selected controls to the clipboard.
  fOK = CopySelectedControls
  
  If fOK Then
    ' Delete the selected controls.
    fOK = DeleteSelectedControls(True)
  End If
  
  If fOK Then
    ' Set the last action flag and enable the Undo menu option.
    giLastActionFlag = giACTION_CUTCONTROLS
    giUndo_TabPageIndex = PageNo
    frmSysMgr.RefreshMenu
  End If

TidyUpAndExit:
  CutSelectedControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function AlignX(pLngX As Long) As Long
  ' Return the given X coordinate aligned to the X grid if required.
  If gfAlignToGrid Then
    AlignX = pLngX - (pLngX Mod giGridX)
  Else
    AlignX = pLngX
  End If
  
End Function
Private Function AlignY(pLngY As Long) As Long
  ' Return the given Y coordinate aligned to the Y grid if required.
  If gfAlignToGrid Then
    AlignY = pLngY - (pLngY Mod giGridY)
  Else
    AlignY = pLngY
  End If
  
End Function

' Validates the screen
Private Function ValidateScreen() As Boolean

  Dim bOK As Boolean
  Dim ctlControl As Control
  Dim bControlsFound As Boolean
  
  bOK = True
  bControlsFound = False
    
  ' Save each screen control.
  For Each ctlControl In Me.Controls
    If IsScreenControl(ctlControl) Then
      bControlsFound = True
      Exit For
    End If
  Next ctlControl
  
  If Not bControlsFound Then
    MsgBox "You cannot save a screen with no controls.", vbExclamation, Me.Caption
    bOK = False
  End If
  
  ValidateScreen = bOK

End Function

Private Function SaveScreen() As Boolean
  ' Save the screen to the local database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iPageNo As Integer
  Dim iCount As Integer
  
  If Not ValidateScreen Then
    Exit Function
  End If
  
  Screen.MousePointer = vbHourglass
  
  With gobjProgress
    .Caption = "Screen Designer"
    .Bar1Value = 0
    .Bar1MaxValue = 100
    .Bar1Caption = "Saving Screen Design..."
    .Cancel = False
    .Time = False
    .OpenProgress
  End With
  
  ' Load existing tab pages, so the save is successful
  If tabPages.Tabs.Count > 0 Then
    For iCount = 1 To tabPages.Tabs.Count
    'For iCount = 1 To picPageContainer.Count - 1
    'For iCount = 0 To picPageContainer.Count - 1
      LoadTabPage (iCount)
    Next iCount
  End If
  
  ' Begin the database transaction.
  daoWS.BeginTrans
  
  ' Set then level of each control.
  If tabPages.Tabs.Count = 0 Then
    fOK = GetControlLevel(Me.hWnd)
  Else
    For iPageNo = 1 To tabPages.Tabs.Count
      fOK = GetControlLevel(picPageContainer(tabPages.Tabs(iPageNo).Tag).hWnd)
    Next iPageNo
  End If
  
  If fOK Then
  
    ' Find the screen record.
    With recScrEdit
      .Index = "idxScreenID"
      .Seek "=", gLngScreenID
      If .NoMatch Then
        Exit Function
      Else
        .Edit
        .Fields("changed") = True
      End If
    End With
      
    ' Update screen details.
    With recScrEdit
      
      '####
      .Fields("height") = Me.Height - 450 - IIf(UI.GetOSVersion = 6, 240, 0)
      .Fields("width") = Me.Width - IIf(UI.GetOSVersion = 6, 240, 0)
      .Fields("fontName") = tabPages.Font.Name
      .Fields("fontSize") = tabPages.Font.Size
      .Fields("fontBold") = tabPages.Font.Bold
      .Fields("fontitalic") = tabPages.Font.Italic
      .Fields("fontStrikeThru") = tabPages.Font.Strikethrough
      .Fields("fontUnderline") = tabPages.Font.Underline
      .Fields("gridX") = GridX
      .Fields("gridY") = GridY
      .Fields("alignToGrid") = IIf(AlignToGrid, 1, 0)
      .Fields("dfltForeColour") = gDfltForeColour
      .Fields("dfltFontName") = Me.Font.Name
      .Fields("dfltFontSize") = Me.Font.Size
      .Fields("dfltFontBold") = IIf(Me.Font.Bold, 1, 0)
      .Fields("dfltFontItalic") = IIf(Me.Font.Italic, 1, 0)
      
      .Update
    
    End With
        
    ' Delete the existing page caption definitions for this screen.
    daoDb.Execute "DELETE FROM tmpPageCaptions WHERE screenID = " & gLngScreenID
  
    ' Save the page captions.
    If tabPages.Tabs.Count > 0 Then
      With recPageCaptEdit
        For iPageNo = 1 To tabPages.Tabs.Count
          .AddNew
          .Fields("screenID") = gLngScreenID
          .Fields("pageIndexID") = iPageNo
          'MH20020527 Fault 1862
          '.Fields("caption") = tabPages.Tabs(iPageNo).Caption
          .Fields("caption") = Replace(tabPages.Tabs(iPageNo).Caption, "&&", "&")
          .Update
        Next iPageNo
      End With
    End If
  
    ' Delete the existing control definitions for this screen.
    daoDb.Execute "DELETE FROM tmpControls WHERE screenID=" & gLngScreenID
      
    fOK = SaveControls
  End If
  
ExitSaveScreen:
  If fOK Then
    gfNewScreen = False
    gfChangedScreen = False
    Application.Changed = True
    frmSysMgr.RefreshMenu
  
    'Commit transaction
    daoWS.CommitTrans dbForceOSFlush
  Else
    'Rollback transaction
    daoWS.Rollback
  End If
  
  gobjProgress.CloseProgress
  Screen.MousePointer = vbDefault
  SaveScreen = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume ExitSaveScreen
  
End Function


Private Function SelectAllPageControls() As Boolean
  ' Select all controls on the current page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim ctlControl As VB.Control
  Dim VarPageContainer As Variant
  
  fOK = True
  
  'JDM - Fault 2454 - Put up an hourglass
  Screen.MousePointer = vbHourglass
   
  ' Get the current page container.
  Set VarPageContainer = CurrentPageContainer
  
  ' Select all the controls on this container
  For Each ctlControl In Me.Controls
    'TM20010914 Fault 1753
    'The ActiveBar control does mot have the visible property, so to avoid err
    'we only check the visible property of other controls.
    If ctlControl.Name <> "abScreen" Then
      If ctlControl.Visible Then
        If IsScreenControl(ctlControl) Then
          If ctlControl.Container Is VarPageContainer Then
            ctlControl.Selected = True
            SelectControl ctlControl
          End If
        End If
      End If
    End If
  Next ctlControl
  
  'Refresh the menu
  frmSysMgr.RefreshMenu
  
TidyUpAndExit:

  'Reset the mousepointer
  Screen.MousePointer = vbDefault
  
  ' Disassociate object variables.
  Set ctlControl = Nothing
  SelectAllPageControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function UndoLastAction() As Boolean
  ' Undo the last action.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  
  Select Case giLastActionFlag
  
    ' Undo the previous TabPage Drop.
    Case giACTION_DROPTABPAGE
      If Not UndoDropTabPage Then
        MsgBox "Unable to undo Drop Tab Page." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Undo the previous control Drop.
    Case giACTION_DROPCONTROL, giACTION_DROPCONTROLAUTOLABEL
      If Not UndoDropControl Then
        MsgBox "Unable to undo Drop Control." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Undo the previous control Cut.
    Case giACTION_CUTCONTROLS
      If Not UndoCutControls Then
        MsgBox "Unable to undo Cut Controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Undo the previous control Paste.
    Case giACTION_PASTECONTROLS
      If Not UndoPasteControls Then
        MsgBox "Unable to undo Paste Controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Undo the previous Tab Page Delete.
    Case giACTION_DELETETABPAGE
      If Not UndoDeleteTabPage Then
        MsgBox "Unable to undo Delete Tab Pages." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Undo the previous control Delete.
    Case giACTION_DELETECONTROLS
      If Not UndoDeleteControls Then
        MsgBox "Unable to undo Delete Controls." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
    ' Undo the previous AutoFormat.
    Case giACTION_AUTOFORMAT
      If Not UndoAutoFormat Then
        MsgBox "Unable to undo AutoFormat." & vbCr & vbCr & _
          Err.Description, vbExclamation + vbOKOnly, App.ProductName
      End If
      
  End Select
  
  ' Clear the last action flag.
  giLastActionFlag = giACTION_NOACTION
  
  ' Disable the Undo button on the menubar.
  frmSysMgr.RefreshMenu
  
  ' Refresh the properties screen.
  Set frmScrObjProps.CurrentScreen = Me
  frmScrObjProps.RefreshProperties

  fOK = True
  
TidyUpAndExit:
  UndoLastAction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function GetControlPageNo(pctlControl As VB.Control) As Integer
  ' Return the page number on which the given control is located.
  ' =0 - no tab pages. ie. the form itself.
  ' >0 - the tab page index.
  On Error GoTo ErrorTrap
  
  Dim iPageNo As Integer
  Dim objTabPage As Object
  
  iPageNo = 0
        
  If (tabPages.Tabs.Count > 0) And (Not pctlControl.Container Is Me) Then
    For Each objTabPage In tabPages.Tabs
      If objTabPage.Tag = pctlControl.Container.Index Then
        iPageNo = objTabPage.Index
      End If
    Next objTabPage
  End If

TidyUpAndExit:
  ' Disassociate object variables.
  Set objTabPage = Nothing
  ' Return the page number.
  GetControlPageNo = iPageNo
  Exit Function

ErrorTrap:
  iPageNo = 0
  Resume TidyUpAndExit
  
End Function

Public Function AutoSizeControl(pctlControl As VB.Control) As Boolean
  ' Initialise the given control's properties.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iType As Long
  Dim iDigits As Integer
  Dim iMinLength As Integer
  Dim iMaxLength As Integer
  Dim lngColumnID As Long
  Dim lngMinWidth As Long
  Dim lngMinHeight As Long
  Dim iExtraWidth As Integer
  Dim sngWidth As Single
  Dim iLoop As Integer
  Dim fLiteral As Boolean
  Dim sMask As String
  
  lngColumnID = pctlControl.ColumnID
  iType = ScreenControl_Type(pctlControl)
    
  ' If we are initialising a column control.
  If lngColumnID >= 0 Then
      
    With recColEdit
      .Index = "idxColumnID"
      .Seek "=", lngColumnID
          
      If Not .NoMatch Then
            
        ' Set the width of the new control.
        Select Case iType
          
          Case giCTRL_TEXTBOX, giCTRL_COMBOBOX
            ' If the column has size, then set the control
            ' width to column size * average character width.
            If .Fields("size") > 0 Then
              
              If .Fields("datatype") = dtVARCHAR Then
                If .Fields("Multiline") Then
                  pctlControl.Width = TextWidth(String(500, "W")) + (2 * XFrame)
                  pctlControl.Height = Me.TextHeight("W") * 3
                Else
                  If Len(.Fields("Mask")) > 0 Then
                    fLiteral = False
                    sMask = .Fields("Mask")
                    
                    For iLoop = 1 To Len(sMask)
                      If fLiteral Then
                        sngWidth = sngWidth + TextWidth(String(1, Mid(sMask, iLoop, 1)))
                        fLiteral = False
                      Else
                        Select Case Mid(sMask, iLoop, 1)
                          Case "A"
                            sngWidth = sngWidth + TextWidth(String(1, "W"))
                          Case "a"
                            sngWidth = sngWidth + TextWidth(String(1, "w"))
                          Case "9"
                            sngWidth = sngWidth + TextWidth(String(1, "8"))
                          Case "#"
                            sngWidth = sngWidth + TextWidth(String(1, "8"))
                          Case "B"
                            sngWidth = sngWidth + TextWidth(String(1, "0"))
                          Case "\"
                            fLiteral = True
                          Case Else
                            sngWidth = sngWidth + TextWidth(String(1, Mid(sMask, iLoop, 1)))
                        End Select
                      End If
                    Next iLoop
                  
                    pctlControl.Width = sngWidth + (2 * XFrame)
                  Else
                    pctlControl.Width = Default_ColumnWidth_Textbox(.Fields("size").value)
                  End If
                End If
              Else
              
                pctlControl.Width = Default_ColumnWidth_Numeric(.Fields("size").value, .Fields("decimals").value, .Fields("Use1000Separator").value)

              End If
            End If
                
            If .Fields("dataType") = dtTIMESTAMP Then
              pctlControl.Width = TextWidth("28/12/2000") + (4 * XFrame) + 255
            End If
                
          Case giCTRL_SPINNER
            iMinLength = Len(Trim(Str(.Fields("spinnerMinimum"))))
            iMaxLength = Len(Trim(Str(.Fields("spinnerMaximum"))))
            iDigits = IIf(iMinLength > iMaxLength, iMinLength, iMaxLength)
           ' pctlControl.Width = (iDigits * UI.GetMaxCharWidth(Me.hDC)) + (2 * XFrame)
            pctlControl.Width = TextWidth(String(iDigits, "8")) + (2 * XFrame)
                  
        End Select
        
      End If
    End With
  End If
            
  Select Case iType
    ' Set the control to have the minimum width and height for labels.
    Case giCTRL_LABEL
      lngMinWidth = TextWidth(pctlControl.Caption)
      lngMinWidth = IIf(lngMinWidth < 255, 255, lngMinWidth)
      pctlControl.Width = lngMinWidth
      lngMinHeight = Me.TextHeight(pctlControl.Caption)
      lngMinHeight = IIf(lngMinHeight < 195, 195, lngMinHeight)
      pctlControl.Height = lngMinHeight
                
    ' Set the control to have the minimum height for textboxes.
    ' Do not set width.
    Case giCTRL_TEXTBOX
      'pctlControl.Height = UI.GetCharHeight(Me.hDC) + (2 * YFrame)
'      pctlControl.Height = pctlControl.MinimumHeight
                
    Case giCTRL_COMBOBOX
                
    ' Set the control to have the minimum height for spinners.
    ' Do not set width.
    Case giCTRL_SPINNER
      pctlControl.Height = pctlControl.MinimumHeight
                
    ' Set the control to have the minimum width and height for check boxes.
    Case giCTRL_CHECKBOX
      lngMinWidth = 360 + TextWidth("W" & pctlControl.Caption)
      pctlControl.Width = lngMinWidth
      lngMinHeight = UI.GetCharHeight(Me.hDC)
      If lngMinHeight < 285 Then lngMinHeight = 285
      pctlControl.Height = lngMinHeight
               
    Case giCTRL_OPTIONGROUP
    Case giCTRL_PHOTO
    Case giCTRL_OLE
    Case giCTRL_FRAME
    Case giCTRL_IMAGE
    
    Case giCTRL_WORKINGPATTERN
    
    Case giCTRL_LINK
      pctlControl.Width = (Len(pctlControl.Caption) * UI.GetAvgCharWidth(Me.hDC)) + (2 * XFrame)
      pctlControl.Height = UI.GetCharHeight(Me.hDC) + (2 * YFrame)
                
    Case giCTRL_LINE
      pctlControl.Alignment = 1
      pctlControl.Length = 1000
      pctlControl.Width = 1000
      
  End Select
          
  ' Ensure the control does not extend past the right-hand edge
  ' of the parent container.
  With pctlControl
    If .Left + .Width > .Container.Width Then
      .Width = .Container.Width - .Left
    End If
  End With
  
  fOK = True
  
TidyUpAndExit:
  AutoSizeControl = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Property Get UndoAction() As UndoActionFlags
  ' Return the key that identifies the alast action that can be 'undone'.
  UndoAction = giLastActionFlag
  
End Property


Private Function UndoDropTabPage() As Boolean
  ' Delete the last Tab Page that was created.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim ctlControl As VB.Control
  
  ' If there is only one tab page, then moved all contained controls onto the
  ' form itself.
  If giUndo_TabPageIndex = 1 Then
  
    For Each ctlControl In Me.Controls
      If IsScreenControl(ctlControl) Then
        Set ctlControl.Container = Me
        
        ' JDM - 22/08/02 - Fault 4265 - Put frame to back when undoing add tab page
        If ctlControl.Name = "asrDummyFrame" Then
          ctlControl.ZOrder 1
        End If
        
      End If
    Next ctlControl
     
    ' Adjust the form's dimensions.
    Me.Height = Me.Height - (tabPages.Height - tabPages.ClientHeight) - (2 * YFrame)
    Me.Width = Me.Width - (4 * XFrame)
     
    ' Disassociate object variables.
    Set ctlControl = Nothing
  End If
      
  ' Delete the tab page.
  fOK = DeleteTabPage(giUndo_TabPageIndex, False)




TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  UndoDropTabPage = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function UndoDropControl() As Boolean
  ' Delete the last control that was dropped on the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  If giLastActionFlag = giACTION_DROPCONTROLAUTOLABEL Then
    fOK = DeleteControl(asrDummyLabel(giUndo_ControlAutoLabelIndex))
  End If
  
  Select Case gsUndo_ControlType
    Case "asrDummyLabel"
      fOK = DeleteControl(asrDummyLabel(giUndo_ControlIndex))
    Case "asrDummyTextBox"
      fOK = DeleteControl(asrDummyTextBox(giUndo_ControlIndex))
    Case "asrDummyPhoto"
      fOK = DeleteControl(asrDummyPhoto(giUndo_ControlIndex))
    Case "asrDummyOLEContents"
      fOK = DeleteControl(asrDummyOLEContents(giUndo_ControlIndex))
    Case "asrDummyImage"
      fOK = DeleteControl(asrDummyImage(giUndo_ControlIndex))
    Case "asrDummyFrame"
      fOK = DeleteControl(asrDummyFrame(giUndo_ControlIndex))
    Case "asrDummyCombo"
      fOK = DeleteControl(asrDummyCombo(giUndo_ControlIndex))
    Case "asrDummySpinner"
      fOK = DeleteControl(asrDummySpinner(giUndo_ControlIndex))
    Case "asrDummyCheckBox"
      fOK = DeleteControl(asrDummyCheckBox(giUndo_ControlIndex))
    Case "ASRDummyOptions"
      fOK = DeleteControl(ASRDummyOptions(giUndo_ControlIndex))
    Case "asrDummyLink"
      fOK = DeleteControl(asrDummyLink(giUndo_ControlIndex))
      
    'JDM - 22/08/02 - Fault 3960 - Not removing working pattern control
    Case "ASRCustomDummyWP"
      fOK = DeleteControl(ASRCustomDummyWP(giUndo_ControlIndex))
  
    'JPD 20070102 Fault 5857
    Case "ASRDummyLine"
      fOK = DeleteControl(ASRDummyLine(giUndo_ControlIndex))
  End Select

TidyUpAndExit:
  UndoDropControl = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function UndoPasteControls() As Boolean
  ' Delete the last controls that were pasted on the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iIndex2 As Integer
  
  fOK = True
  
  ' Delete the pasted controls.
  For iIndex = 1 To UBound(gavUndo_PastedControls, 2)
  
    iIndex2 = gavUndo_PastedControls(2, iIndex)
    
    Select Case gavUndo_PastedControls(1, iIndex)
      Case "asrDummyLabel"
        fOK = DeleteControl(asrDummyLabel(iIndex2))
      Case "asrDummyTextBox"
        fOK = DeleteControl(asrDummyTextBox(iIndex2))
      Case "asrDummyPhoto"
        fOK = DeleteControl(asrDummyPhoto(iIndex2))
      Case "asrDummyOLEContents"
        fOK = DeleteControl(asrDummyOLEContents(iIndex2))
      Case "asrDummyImage"
        fOK = DeleteControl(asrDummyImage(iIndex2))
      Case "asrDummyFrame"
        fOK = DeleteControl(asrDummyFrame(iIndex2))
      Case "asrDummyCombo"
        fOK = DeleteControl(asrDummyCombo(iIndex2))
      Case "asrDummySpinner"
        fOK = DeleteControl(asrDummySpinner(iIndex2))
      Case "asrDummyCheckBox"
        fOK = DeleteControl(asrDummyCheckBox(iIndex2))
      Case "ASRDummyOptions"
        fOK = DeleteControl(ASRDummyOptions(iIndex2))
      Case "asrDummyLink"
        fOK = DeleteControl(asrDummyLink(iIndex2))
    
      'JDM - 22/08/02 - Fault 3959 - Not removing working pattern control
      Case "ASRCustomDummyWP"
        fOK = DeleteControl(ASRCustomDummyWP(iIndex2))
    
    End Select
  Next iIndex
  
TidyUpAndExit:
  UndoPasteControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Private Function UndoCutControls() As Boolean
  ' Paste the cut controls back onto their original page.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = UndoDeleteControls

TidyUpAndExit:
  UndoCutControls = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function UndoAutoFormat() As Boolean
  ' Delete any AutoFormatted tab pages.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim ctlControl As VB.Control
  
  fOK = True
  
  ' Delete any AutoFormatted tab pages.
  If giUndo_TabPageIndex > 0 Then
      
    For iLoop = tabPages.Tabs.Count To giUndo_TabPageIndex Step -1
      fOK = DeleteTabPage(iLoop, False)
      
      If Not fOK Then
        Exit For
      End If
    Next iLoop
        
    ' Move the original controls onto the form, and delete the last tab page
    ' if it was added as part of the AutoFormat.
    If fOK And gfUndo_TabsCreated Then
    
      For Each ctlControl In Me.Controls
        If IsScreenControl(ctlControl) Then
          Set ctlControl.Container = Me
        
          ' JDM - 16/09/03 - Fault 6258 - Push frame to back
          If ctlControl.Name = "asrDummyFrame" Then
            ctlControl.ZOrder 1
          End If
        
        End If
      Next ctlControl
      ' Disassociate object variables.
      Set ctlControl = Nothing
        
      ' Adjust the form's dimensions.
      Me.Height = Me.Height - (tabPages.Height - tabPages.ClientHeight) - (2 * YFrame)
      Me.Width = Me.Width - (4 * XFrame)
      
      ' Delete the first AutoFormatted tab page.
      fOK = DeleteTabPage(1, False)
    
    End If
  Else
    ' Delete all AutoFormatted controls.
    For Each ctlControl In Me.Controls
      If IsScreenControl(ctlControl) Then
        fOK = DeleteControl(ctlControl)
      End If
      
      If Not fOK Then
        Exit For
      End If
    Next ctlControl
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  UndoAutoFormat = fOK
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

 ' Open a progress bar
  With gobjProgress
    .Caption = "Screen Designer"
    .Bar1Value = 0
    .Bar1MaxValue = UBound(gactlUndo_DeletedControls)
    .Bar1Caption = "Undoing Screen Control Deletion..."
    .Cancel = False
    .Time = False
    .OpenProgress
  End With
  
  ' Go to the page from which the controls were deleted.
  PageNo = giUndo_TabPageIndex

  ' Restore the deleted controls to their original positions.
  For iIndex = 1 To UBound(gactlUndo_DeletedControls)

    Set ctlNewControl = gactlUndo_DeletedControls(iIndex)
    ctlNewControl.Visible = True
    fOK = SelectControl(ctlNewControl)
    
    ' Disassociate object variables.
'    Set ctlNewControl = Nothing
    
    Set gactlUndo_DeletedControls(iIndex) = Nothing
  
    If Not fOK Then
      Exit For
    End If

    'Update the progress bar
    gobjProgress.UpdateProgress

  Next iIndex

  ' Clear the array of deleted controls.
  ReDim gactlUndo_DeletedControls(0)

TidyUpAndExit:

  'Close the progress bar
  gobjProgress.CloseProgress

  UndoDeleteControls = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function UndoDeleteTabPage() As Boolean
  ' Recreate the last tab page that was deleted.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim picContainer As PictureBox
  Dim ctlNewControl As VB.Control

  fOK = AddTabPage(giUndo_TabPageIndex)
  
  If fOK Then
  
    ' Restore the original page caption.
    tabPages.Tabs(giUndo_TabPageIndex).Caption = gsUndo_TabPageCaption
     
    ' Recreate the controls that were on this page when it was deleted.
    Set picContainer = picPageContainer(tabPages.Tabs(giUndo_TabPageIndex).Tag)
  
    ' Restore the deleted controls to their original positions.
    For iIndex = 1 To UBound(gactlUndo_DeletedControls)
    
      Set ctlNewControl = gactlUndo_DeletedControls(iIndex)
      ctlNewControl.Visible = True
      Set ctlNewControl.Container = picContainer
      fOK = SelectControl(ctlNewControl)
        
      ' Disassociate object variables.
      Set ctlNewControl = Nothing
        
      Set gactlUndo_DeletedControls(iIndex) = Nothing
      
      If Not fOK Then
        Exit For
      End If
    Next iIndex
    
    ' Clear the array of deleted controls.
    ReDim gactlUndo_DeletedControls(0)
              
    ' Go to the page.
    PageNo = giUndo_TabPageIndex
  
  End If

TidyUpAndExit:
  ' Disassociate object varables.
  Set ctlNewControl = Nothing
  Set picContainer = Nothing
  UndoDeleteTabPage = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function GetControlLevel(pLngHWnd As Long) As Boolean
  ' Determine the control level of each screen control. Set the 'controlLevel' property
  ' of the screen controls with the determined value.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim iCounter As Integer
  Dim lngChildHWnd As Long
  Dim actlScreenControls() As VB.Control
  Dim ctlControl As VB.Control
  
  ' Create an array of the screen control's.
  ReDim actlScreenControls(0)
    
  ' Construct an array of the screen controls.
  For Each ctlControl In Me.Controls
    If IsScreenControl(ctlControl) Then
      iIndex = UBound(actlScreenControls) + 1
      ReDim Preserve actlScreenControls(iIndex)
      Set actlScreenControls(iIndex) = ctlControl
    End If
  Next ctlControl
    
  ' Disassociate object variables.
  Set ctlControl = Nothing
  
  iCounter = 1
  
  ' Get the hWnd of the first child window of the given page.
  lngChildHWnd = UI.GetChildWindowHWnd(pLngHWnd, GW_CHILD)
    
  ' Find all the child windows of the screen designer.
  Do While lngChildHWnd <> 0
    ' Check if the child window is a screen control.
    For iLoop = 1 To UBound(actlScreenControls)
      Set ctlControl = actlScreenControls(iLoop)
      If lngChildHWnd = ctlControl.hWnd Then
        ctlControl.ControlLevel = iCounter
        iCounter = iCounter + 1
        Exit For
      End If
      Set ctlControl = Nothing
    Next iLoop
    
    ' Get the hWnd of the next child window of the screen designer.
    lngChildHWnd = UI.GetChildWindowHWnd(lngChildHWnd, GW_HWNDNEXT)
  Loop

  fOK = True
  
TidyUpAndExit:
  GetControlLevel = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function SetControlLevel() As Boolean
  ' Set the correct z-order for each control.
  ' The controlLevel property of each control will determine the z-order of each control, but
  ' we now need to actually set that z-order value.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLevel As Integer
  Dim iMaxLevel As Integer
  Dim ctlControl As VB.Control
  
  ' Initialise the array of control information.
  iMaxLevel = 0
  
  ' Find the highest control level.
  For Each ctlControl In Me.Controls
    With ctlControl
      If IsScreenControl(ctlControl) Then
        If ctlControl.ControlLevel > iMaxLevel Then iMaxLevel = ctlControl.ControlLevel
      End If
    End With
  Next ctlControl
  ' Disassociate object variables.
  Set ctlControl = Nothing

  ' Set the z-order for each control.
  For iLevel = iMaxLevel To 0 Step -1
    For Each ctlControl In Me.Controls
      If IsScreenControl(ctlControl) Then
        If ctlControl.ControlLevel = iLevel Then
          ctlControl.ZOrder 0
        End If
        
        ' JDM - 02/09/02 - Fault 4347 - Push frame to back
        If ctlControl.Name = "asrDummyFrame" Then
          ctlControl.ZOrder 1
        End If
        
        ' JDM - 16/05/02 - Fault 10050 - Pull labels to front
        If ctlControl.Name = "asrDummyLabel" Then
          ctlControl.ZOrder 0
        End If
        
      End If
    Next ctlControl
  Next iLevel
  
  fOK = True
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  SetControlLevel = fOK
  Exit Function
  
ErrorTrap:
  MsgBox "Error setting control level." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ScreenControl_NeedsLabelling(piControlType As Long) As Boolean
  ' Returns true if the current control type needs labelling on the screen.
  On Error GoTo ErrorTrap
  
  Dim fNeedsLabelling  As Boolean
  
  fNeedsLabelling = (piControlType = giCTRL_COMBOBOX) Or _
    (piControlType = giCTRL_IMAGE) Or _
    (piControlType = giCTRL_OLE) Or _
    (piControlType = giCTRL_SPINNER) Or _
    (piControlType = giCTRL_PHOTO) Or _
    (piControlType = giCTRL_TEXTBOX)

TidyUpAndExit:
  ScreenControl_NeedsLabelling = fNeedsLabelling
  Exit Function
  
ErrorTrap:
  fNeedsLabelling = False
  Resume TidyUpAndExit
  
End Function

Private Function FirstEmptyPage() As Integer
  ' Return the number of the last empty page.
  On Error GoTo ErrorTrap
  
  Dim iPageNo As Integer
  Dim iFirstPage As Integer
  Dim ctlControl As VB.Control
  
  iFirstPage = 0
  
  For Each ctlControl In Me.Controls
    If IsScreenControl(ctlControl) Then
      ' Get the control's page no.
      iPageNo = GetControlPageNo(ctlControl)
      If iFirstPage < (iPageNo + 1) Then
        iFirstPage = (iPageNo + 1)
      End If
    End If
  Next ctlControl

TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  FirstEmptyPage = iFirstPage
  Exit Function

ErrorTrap:
  iFirstPage = -1
  Resume TidyUpAndExit

End Function

Public Function IsScreenControl(pctlControl As VB.Control) As Boolean
  ' Return true if the given control is a screen control.
  On Error GoTo ErrorTrap
  
  Dim fIsScreenControl As Boolean
  Dim iIndex As Integer
  Dim sName As String
  
  sName = pctlControl.Name
  fIsScreenControl = False
   
  If sName = "asrDummyLabel" Or _
    sName = "asrDummyTextBox" Or _
    sName = "asrDummyPhoto" Or _
    sName = "asrDummyOLEContents" Or _
    sName = "asrDummyImage" Or _
    sName = "asrDummyFrame" Or _
    sName = "asrDummyCombo" Or _
    sName = "asrDummySpinner" Or _
    sName = "asrDummyCheckBox" Or _
    sName = "asrDummyLink" Or _
    sName = "ASRCustomDummyWP" Or _
    sName = "ASRDummyLine" Or _
    sName = "ASRDummyNavigation" Or _
    sName = "ASRColourSelector" Or _
    sName = "ASRDummyOptions" Then
    
    ' Do not bother with the dummy screen controls.
    If (pctlControl.Index > 0) Then

      fIsScreenControl = True
      
      ' Do not bother with controls in the deleted array.
      For iIndex = 1 To UBound(gactlUndo_DeletedControls)
        If pctlControl Is gactlUndo_DeletedControls(iIndex) Then
          fIsScreenControl = False
          Exit For
        End If
      Next iIndex
    
      If fIsScreenControl Then
        ' Do not bother with controls in the clipboard array.
        For iIndex = 1 To UBound(gactlClipboardControls)
          If pctlControl Is gactlClipboardControls(iIndex) Then
            fIsScreenControl = False
            Exit For
          End If
        Next iIndex
      End If
      
    End If
  End If
  
TidyUpAndExit:
  IsScreenControl = fIsScreenControl
  Exit Function

ErrorTrap:
  fIsScreenControl = False
  Resume TidyUpAndExit
  
End Function

Public Function ScreenControl_HasForeColor(piControlType As Long) As Boolean
  ' Return true if the given control has a ForeColor property.
  ScreenControl_HasForeColor = _
    (piControlType = giCTRL_CHECKBOX) Or _
    (piControlType = giCTRL_COMBOBOX) Or _
    (piControlType = giCTRL_FRAME) Or _
    (piControlType = giCTRL_LABEL) Or _
    (piControlType = giCTRL_OPTIONGROUP) Or _
    (piControlType = giCTRL_SPINNER) Or _
    (piControlType = giCTRL_WORKINGPATTERN) Or _
    (piControlType = giCTRL_NAVIGATION) Or _
    (piControlType = giCTRL_TEXTBOX)

End Function

Public Function ScreenControl_HasDisplayType(piControlType As Long) As Boolean
  ' Return true if the given control has a DisplayType property.
  ScreenControl_HasDisplayType = _
    (piControlType = giCTRL_NAVIGATION)

End Function

Public Function ScreenControl_HasNavigation(piControlType As Long) As Boolean
  ' Return true if the given control has a NavigateTo property.
  ScreenControl_HasNavigation = (piControlType = giCTRL_NAVIGATION)

End Function

Public Function ScreenControl_HasBackColor(piControlType As Long) As Boolean
  ' Return true if the given control has a BackColor property.
  ScreenControl_HasBackColor = _
    (piControlType = giCTRL_CHECKBOX) Or _
    (piControlType = giCTRL_COMBOBOX) Or _
    (piControlType = giCTRL_FRAME) Or _
    (piControlType = giCTRL_LABEL) Or _
    (piControlType = giCTRL_OPTIONGROUP) Or _
    (piControlType = giCTRL_SPINNER) Or _
    (piControlType = giCTRL_WORKINGPATTERN) Or _
    (piControlType = giCTRL_TEXTBOX)

End Function

Public Function ScreenControl_HasFont(piControlType As Long) As Boolean
  ' Return true if the given control has a Font property.
  ScreenControl_HasFont = _
    (piControlType = giCTRL_CHECKBOX) Or _
    (piControlType = giCTRL_COMBOBOX) Or _
    (piControlType = giCTRL_FRAME) Or _
    (piControlType = giCTRL_LABEL) Or _
    (piControlType = giCTRL_OPTIONGROUP) Or _
    (piControlType = giCTRL_SPINNER) Or _
    (piControlType = giCTRL_LINK) Or _
    (piControlType = giCTRL_WORKINGPATTERN) Or _
    (piControlType = giCTRL_NAVIGATION) Or _
    (piControlType = giCTRL_TEXTBOX)

End Function

Public Function ScreenControl_IsTabStop(piControlType As Long) As Boolean
  ' Return true if the given control has a TabStop property.
  ScreenControl_IsTabStop = (piControlType = giCTRL_CHECKBOX) Or _
    (piControlType = giCTRL_COMBOBOX) Or _
    (piControlType = giCTRL_OLE) Or _
    (piControlType = giCTRL_PHOTO) Or _
    (piControlType = giCTRL_OPTIONGROUP) Or _
    (piControlType = giCTRL_SPINNER) Or _
    (piControlType = giCTRL_LINK) Or _
    (piControlType = giCTRL_WORKINGPATTERN) Or _
    (piControlType = giCTRL_NAVIGATION) Or _
    (piControlType = giCTRL_COLOURPICKER) Or _
    (piControlType = giCTRL_TEXTBOX)

End Function

Public Function ScreenControl_HasAlignment(piControlType As Long) As Boolean
  ' Return true if the given control has an Alignment property.
  ScreenControl_HasAlignment = _
    (piControlType = giCTRL_CHECKBOX) Or _
    (piControlType = giCTRL_SPINNER)

End Function

Public Function ScreenControl_HasOrientation(piControlType As Long) As Boolean
  ' Return true if the given control has an Orientation property.
  ScreenControl_HasOrientation = _
    (piControlType = giCTRL_OPTIONGROUP) Or _
    (piControlType = giCTRL_LINE)


End Function


Public Function ScreenControl_HasOptions(piControlType As Long) As Boolean
  ' Return true if the given control has an Options property.
  ScreenControl_HasOptions = (piControlType = giCTRL_OPTIONGROUP)

End Function

Public Function ScreenControl_HasPicture(piControlType As Long) As Boolean
  ' Return true if the given control has a Picture property.
  ScreenControl_HasPicture = (piControlType = giCTRL_IMAGE)

End Function

Public Function ScreenControl_HasWidth(piControlType As Long) As Boolean
  ' Return true if the given control has a Width property.
  ScreenControl_HasWidth = (piControlType = giCTRL_CHECKBOX) Or _
    (piControlType = giCTRL_COMBOBOX) Or _
    (piControlType = giCTRL_FRAME) Or _
    (piControlType = giCTRL_IMAGE) Or _
    (piControlType = giCTRL_LABEL) Or _
    (piControlType = giCTRL_OLE) Or _
    (piControlType = giCTRL_PHOTO) Or _
    (piControlType = giCTRL_LINK) Or _
    (piControlType = giCTRL_LINE) Or _
    (piControlType = giCTRL_SPINNER) Or _
    (piControlType = giCTRL_NAVIGATION) Or _
    (piControlType = giCTRL_COLOURPICKER) Or _
    (piControlType = giCTRL_TEXTBOX)

End Function



Public Function ScreenControl_HasBorderStyle(piControlType As Long) As Boolean
  ' Return true if the given control has a BorderStyle property.
'  ScreenControl_HasBorderStyle = _
'    (piControlType = giCTRL_IMAGE) Or _
'    (piControlType = giCTRL_OLE) Or _
'    (piControlType = giCTRL_PHOTO) Or _
'    (piControlType = giCTRL_WORKINGPATTERN) Or _
'    (piControlType = giCTRL_OPTIONGROUP)

  ScreenControl_HasBorderStyle = _
    (piControlType = giCTRL_IMAGE) Or _
    (piControlType = giCTRL_PHOTO) Or _
    (piControlType = giCTRL_WORKINGPATTERN) Or _
    (piControlType = giCTRL_OPTIONGROUP)

End Function

Public Function ScreenControl_HasReadOnly(piControlType As Long) As Boolean
  'NPG20071022
  ' Return true if the given control has a Width property.
  ScreenControl_HasReadOnly = _
    (piControlType = giCTRL_CHECKBOX) Or _
    (piControlType = giCTRL_COMBOBOX) Or _
    (piControlType = giCTRL_OPTIONGROUP) Or _
    (piControlType = giCTRL_OLE) Or _
    (piControlType = giCTRL_TEXTBOX) Or _
    (piControlType = giCTRL_SPINNER) Or _
    (piControlType = giCTRL_PHOTO) Or _
    (piControlType = giCTRL_COLOURPICKER) Or _
    (piControlType = giCTRL_WORKINGPATTERN)

'NPG20071210 Fault 12694
'    (piControlType = giCTRL_LINK) Or _


End Function

Public Function ScreenControl_HasCaption(piControlType As Long) As Boolean
  ' Return true if the given control has a Caption property.
  ScreenControl_HasCaption = _
    (piControlType = giCTRL_CHECKBOX) Or _
    (piControlType = giCTRL_FRAME) Or _
    (piControlType = giCTRL_LABEL) Or _
    (piControlType = giCTRL_LINK) Or _
    (piControlType = giCTRL_NAVIGATION) Or _
    (piControlType = giCTRL_OPTIONGROUP)

End Function

Public Function ScreenControl_HasText(piControlType As Long) As Boolean
  ' Return true if the given control has a Caption property.
  ScreenControl_HasText = _
    (piControlType = giCTRL_TEXTBOX) Or _
    (piControlType = giCTRL_PHOTO) Or _
    (piControlType = giCTRL_OLE) Or _
    (piControlType = giCTRL_COMBOBOX) Or _
    (piControlType = giCTRL_SPINNER)

End Function

Public Function ScreenControl_Type(pctlControl As VB.Control) As Long
  ' Return the control type of the given control.
  Select Case pctlControl.Name
    Case "asrDummyLabel"
      ScreenControl_Type = giCTRL_LABEL
    Case "asrDummyTextBox"
      ScreenControl_Type = giCTRL_TEXTBOX
    Case "asrDummyOLEContents"
      ScreenControl_Type = giCTRL_OLE
    Case "asrDummyPhoto"
      ScreenControl_Type = giCTRL_PHOTO
    Case "asrDummyImage"
      ScreenControl_Type = giCTRL_IMAGE
    Case "asrDummyFrame"
      ScreenControl_Type = giCTRL_FRAME
    Case "asrDummyCombo"
      ScreenControl_Type = giCTRL_COMBOBOX
    Case "asrDummySpinner"
      ScreenControl_Type = giCTRL_SPINNER
    Case "asrDummyCheckBox"
      ScreenControl_Type = giCTRL_CHECKBOX
    Case "ASRDummyOptions"
      ScreenControl_Type = giCTRL_OPTIONGROUP
    Case "asrDummyLink"
      ScreenControl_Type = giCTRL_LINK
    Case "ASRCustomDummyWP"
      ScreenControl_Type = giCTRL_WORKINGPATTERN
    Case "ASRDummyLine"
      ScreenControl_Type = giCTRL_LINE
    Case "ASRDummyNavigation"
      ScreenControl_Type = giCTRL_NAVIGATION
    Case "ASRColourSelector"
      ScreenControl_Type = giCTRL_COLOURPICKER
  End Select
  
End Function

Private Function SaveControl(pctlControl As VB.Control) As Boolean
  ' Save the definition of the given screen control to the database.
  On Error GoTo ErrorTrap
  
  Dim fSaveOK As Boolean
  Dim iControlType As Long
  Dim objFont As StdFont
  Dim vNull As Variant
  
  ' Do not save the dummy control array controls (index = 0).
  ' Do not save the clipboard controls.
  If (pctlControl.Index > 0) Then
    
    iControlType = ScreenControl_Type(pctlControl)
  
    'Add control definition
    With recCtrlEdit
      
      .AddNew
        
      .Fields("screenID") = gLngScreenID
      .Fields("pageNo") = GetControlPageNo(pctlControl)
      .Fields("controlLevel") = pctlControl.ControlLevel
        
      .Fields("columnID") = pctlControl.ColumnID
          
      ' Get the table ID.
      recColEdit.Index = "idxColumnID"
      recColEdit.Seek "=", pctlControl.ColumnID
      If Not recColEdit.NoMatch Then
        .Fields("tableID") = recColEdit.Fields("tableID")
      End If
        
      .Fields("controlType") = iControlType
      .Fields("controlIndex") = 0           ' No longer used.
      .Fields("topCoord") = pctlControl.Top
      .Fields("leftCoord") = pctlControl.Left
      .Fields("height") = pctlControl.Height
      .Fields("width") = pctlControl.Width
      .Fields("tabIndex") = pctlControl.TabIndex
      
      If (ScreenControl_HasCaption(iControlType)) Or _
        (ScreenControl_HasText(iControlType)) Then
        .Fields("caption") = IIf(Len(pctlControl.Caption) > 0, pctlControl.Caption, vNull)
      End If
      
      If ScreenControl_HasBackColor(iControlType) Then
        .Fields("backColor") = pctlControl.BackColor
      End If
      
      If ScreenControl_HasForeColor(iControlType) Then
        .Fields("foreColor") = pctlControl.ForeColor
      End If

      If ScreenControl_HasDisplayType(iControlType) Then
        .Fields("displayType") = pctlControl.DisplayType
      End If

      If ScreenControl_HasNavigation(iControlType) Then
        .Fields("NavigateTo") = pctlControl.NavigateTo
        .Fields("NavigateIn") = pctlControl.NavigateIn
        .Fields("NavigateOnSave") = pctlControl.NavigateOnSave
      End If

      If ScreenControl_HasBorderStyle(iControlType) Then
        .Fields("borderStyle") = pctlControl.BorderStyle
      End If
          
      'NPG20071023
      If ScreenControl_HasReadOnly(iControlType) Then
        .Fields("readOnly") = pctlControl.Read_Only
      End If
      
      If ScreenControl_HasFont(iControlType) Then
        Set objFont = pctlControl.Font
        .Fields("fontName") = objFont.Name
        .Fields("fontSize") = objFont.Size
        .Fields("fontBold") = objFont.Bold
        .Fields("fontItalic") = objFont.Italic
        .Fields("fontStrikeThru") = objFont.Strikethrough
        .Fields("fontUnderline") = objFont.Underline
        Set objFont = Nothing
      End If
      
      If ScreenControl_HasAlignment(iControlType) Then
        .Fields("alignment") = pctlControl.Alignment
      End If

      If ScreenControl_HasOrientation(iControlType) Then
        .Fields("alignment") = pctlControl.Alignment
      End If

      If ScreenControl_HasPicture(iControlType) Then
        .Fields("pictureID") = pctlControl.PictureID
      End If

      .Update
    End With
  End If

  fSaveOK = True
  
TidyUpAndExit:
  Set objFont = Nothing
  SaveControl = fSaveOK
  Exit Function
  
ErrorTrap:
  fSaveOK = False
  Resume TidyUpAndExit
  
End Function

Private Function SaveControls() As Boolean
  ' Save the definition of each instance of each type of screen control to the database.
  On Error GoTo ErrorTrap
  
  Dim fSaveOK As Boolean
  Dim ctlControl As VB.Control
  
  fSaveOK = True
  
  ' Save each screen control.
  For Each ctlControl In Me.Controls
    If fSaveOK And IsScreenControl(ctlControl) Then
      fSaveOK = SaveControl(ctlControl)
    End If
  Next ctlControl
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set ctlControl = Nothing
  SaveControls = fSaveOK
  Exit Function
  
ErrorTrap:
  fSaveOK = False
  Resume TidyUpAndExit
  
End Function

Private Function DeselectAllControls(Optional pctlException As VB.Control) As Boolean
  
  Dim iCount As Integer
  Dim ctlControl As Control

  ' Hide all the selection markers
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    ASRSelectionMarkers(iCount).Visible = False
  Next iCount
  
  ' Deselect the controls
  For Each ctlControl In Me.Controls
    If IsScreenControl(ctlControl) Then
      With ctlControl
        ctlControl.Selected = False
      End With
    End If
  Next ctlControl
  
  DeselectAllControls = True
  
End Function

Public Function ScreenControl_HasHeight(piControlType As Long) As Boolean
  ' Return true if the given control has a Height property.
  ScreenControl_HasHeight = (piControlType = giCTRL_CHECKBOX) Or _
    (piControlType = giCTRL_FRAME) Or _
    (piControlType = giCTRL_IMAGE) Or _
    (piControlType = giCTRL_LABEL) Or _
    (piControlType = giCTRL_OLE) Or _
    (piControlType = giCTRL_PHOTO) Or _
    (piControlType = giCTRL_LINK) Or _
    (piControlType = giCTRL_LINE) Or _
    (piControlType = giCTRL_SPINNER) Or _
    (piControlType = giCTRL_NAVIGATION) Or _
    (piControlType = giCTRL_TEXTBOX)

End Function

Private Function ScreenControl_DragDrop(pctlControl As VB.Control, pCtlSource As Control, pSngX As Single, pSngY As Single) As Boolean
  ' Drop a control onto the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  With pctlControl
    fOK = DropControl(.Container, pCtlSource, pSngX + .Left, pSngY + .Top)
  End With
  
TidyUpAndExit:
  If Not fOK Then
    MsgBox "Unable to drop the control." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  ScreenControl_DragDrop = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function ScreenControl_MouseMove(pctlControl As VB.Control, pButton As Integer, pSngX As Single, pSngY As Single) As Boolean
  ' Move the control.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iCount As Integer
  Dim lngNewX As Long
  Dim lngNewY As Long
  
  ' Remove the original offset of the mouse cursor
  pSngX = pSngX - mlngXOffset
  pSngY = pSngY - mlngYOffset
  
  fOK = True
  
  ' Only run if the mouse pointer has moved significantly
  If (mlngLastX > pSngX + giGridX) Or (mlngLastX < pSngX - giGridX) _
      Or (mlngLastY > pSngY + giGridY) Or (mlngLastY < pSngY - giGridY) Then
 
    ' Move the selected controls if the left button key is down, and the control is selected
    If pButton = vbLeftButton And pctlControl.Selected Then
    
      For iCount = 1 To ASRSelectionMarkers.Count - 1
        With ASRSelectionMarkers(iCount)
          If .Visible Then
          
            lngNewX = AlignX(pSngX + .AttachedObject.Left)
            lngNewY = AlignX(pSngY + .AttachedObject.Top)
            .AttachedObject.Move lngNewX, lngNewY
          End If
        End With
      Next iCount
    
      gfMoveSelection = True

    End If
      
  End If

TidyUpAndExit:
  ScreenControl_MouseMove = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ScreenControl_MouseUp(pctlControl As VB.Control, piButton As Integer, piShift As Integer, X As Single, Y As Single) As Boolean
  ' Actually move the selected controls to the positions of their movement frames.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iIndex As Integer
  Dim lngXMouse As Long
  Dim lngYMouse As Long
  Dim ctlControl As VB.Control
  Dim avScreenControls() As Variant
  Dim iCount As Integer
  
  fOK = True

  Select Case piButton
    
    ' Handle left button presses.
    Case vbLeftButton

      ' Deselect all OTHER screen controls if the CTRL or SHIFT keys are not pressed,
      ' and if we do not already have the control selected as part of a multiple selection.
      If Not gfMoveSelection Then
      
        If ((piShift And vbShiftMask) = 0) And ((piShift And vbCtrlMask) = 0) Then
          DeselectAllControls
        End If
          
        ' Toggle this control if the shift/ctrl key is pressed
        If ((piShift And vbShiftMask) <> 0) Or ((piShift And vbCtrlMask) <> 0) Then
          pctlControl.Selected = Not pctlControl.Selected
          'Debug.Print pctlControl.Selected
        Else
          DeselectAllControls
          pctlControl.Selected = True
        End If
        
        ' JDM - 20/08/02 - Fault 4309 - Holding down control now selects/deselects controls
        If pctlControl.Selected Then
          SelectControl pctlControl
        Else
          DeselectControl pctlControl
        End If
        
      Else

        ' End placementing of all selected objects
        For iCount = 1 To ASRSelectionMarkers.Count - 1
          With ASRSelectionMarkers(iCount)
            If .Visible Then
              .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize
            End If
          End With
        Next iCount
        
        ' Flag screen as having changed
        IsChanged = True
        
      End If

    ' Show all selected selection markers
    For iCount = 1 To ASRSelectionMarkers.Count - 1
      ASRSelectionMarkers(iCount).ShowSelectionMarkers True
    Next iCount
    
    ' Refresh the properties screen.
    frmSysMgr.RefreshMenu
    Set frmScrObjProps.CurrentScreen = Me
    frmScrObjProps.RefreshProperties
      
    ' Handle right button presses.
    Case vbRightButton
      UI.GetMousePos lngXMouse, lngYMouse
'      frmSysMgr.tbMain.PopupMenu "ID_mnuScreenEdit", ssPopupMenuLeftAlign, lngXMouse, lngYMouse
      frmSysMgr.tbMain.Bands("ID_mnuScreenEdit").TrackPopup -1, -1
  End Select

  gfMoveSelection = False

TidyUpAndExit:

  ' Stop moving the control.
  gfMoveSelection = False

  ' Disassociate object variables.
  Set ctlControl = Nothing
  UI.UnlockWindow
  ScreenControl_MouseUp = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function AutoLabel(pVarPageContainer As Variant, pSngX As Single, pSngY As Single, sCaption As String) As Boolean
  
  ' Drop the given control onto the screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iControlType As Long
  Dim lngColumnID As Long
  Dim sTableName As String
  Dim sColumnName As String
  Dim objFont As StdFont
  Dim objMisc As New Misc
  Dim ctlControl As VB.Control
  
  fOK = True
  
  ' Deselect all existing controls.
  'fOK = DeselectAllControls
  
  If fOK Then
  
    iControlType = giCTRL_LABEL
    Set ctlControl = AddControl(iControlType)
              
    fOK = Not (ctlControl Is Nothing)
          
    'Check that a new control was added successfully
    If fOK Then
  
      With ctlControl

        Set .Container = pVarPageContainer
        .Left = AlignX((CLng(pSngX) - TextWidth(sCaption + Space(5))))
        If .Left < 0 Then
          .Left = CLng(pSngX)
          .Top = AlignY((CLng(pSngY) - (Me.TextHeight(sCaption) + 20)))
        Else
          .Top = AlignY(CLng(pSngY))
        End If
        
        .ColumnID = 0
            
        ' Initialise the new control's font and forecolour.
        If ScreenControl_HasFont(iControlType) Then
          Set objFont = New StdFont
          objFont.Name = Me.Font.Name
          objFont.Size = Me.Font.Size
          objFont.Bold = Me.Font.Bold
          objFont.Italic = Me.Font.Italic
          Set .Font = objFont
          Set objFont = Nothing
        End If
            
        If ScreenControl_HasForeColor(iControlType) Then
          .ForeColor = gDfltForeColour
        End If
        
        If ScreenControl_HasCaption(iControlType) Then
          'MH20030416 Added colon to the end of the autolabel
          '''.Caption = Replace(sCaption, "_", " ")
          ''' RH 08/11/00 - BUG 1313 - Ensure the & char doesnt appear as sht/cut key
          '.Caption = Replace(Replace(sCaption, "_", " "), "&", "&&")
          .Caption = Replace(Replace(sCaption, "_", " "), "&", "&&") & " :"
        End If
            
        ' Default the control's propertes.
        fOK = AutoSizeControl(ctlControl)
              
        If fOK Then
          fOK = SelectControl(ctlControl)
        End If
            
        If fOK Then
          .Visible = True
          .ZOrder 0
        End If
      
        If giLastActionFlag = giACTION_DROPCONTROLAUTOLABEL Then
          giUndo_ControlAutoLabelIndex = .Index
          gsUndo_ControlAutoLabelType = .Name
        Else
          giUndo_ControlAutoLabelIndex = .Index
          gsUndo_ControlAutoLabelType = ""
        End If
      
      End With
      
    End If
          
    ' Disassociate object variables.
    Set ctlControl = Nothing
          
  End If
    
  ' Set focus on the screen designer form.
  Me.SetFocus
  
    
  If fOK Then
    ' Mark the screen as having changed.
    gfChangedScreen = True
    frmSysMgr.RefreshMenu
  
    ' Refresh the properties screen.
    Set frmScrObjProps.CurrentScreen = Me
    frmScrObjProps.RefreshProperties
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objMisc = Nothing
  Set objFont = Nothing
  Set ctlControl = Nothing
  ' Return the success/failure value.
  AutoLabel = fOK
  Exit Function

ErrorTrap:
  ' Flag the error.
  fOK = False
    
  MsgBox "Could not automatically add a label for this control." & vbCrLf & _
         Err.Description, vbExclamation + vbOKOnly, Application.Name
         
  Resume TidyUpAndExit
  
End Function


Private Sub ASRDummyLine_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  ' Drop a control onto the screen.
  ScreenControl_DragDrop ASRDummyLine(Index), Source, X, Y
  
End Sub

Private Sub ASRDummyLine_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Select the control.
  ScreenControl_MouseDown ASRDummyLine(Index), Button, Shift, X, Y

End Sub

Private Sub ASRDummyLine_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the control.
  ScreenControl_MouseMove ASRDummyLine(Index), Button, X, Y

End Sub

Private Sub ASRDummyLine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move the selected controls.
  ScreenControl_MouseUp ASRDummyLine(Index), Button, Shift, X, Y

End Sub

Public Function ScreenControl_HasTabNumber(piControlType As Long) As Boolean
  ' Return true if the given control has a Tab Index property.
  ScreenControl_HasTabNumber = (piControlType = giCTRL_TAB)

End Function
Private Sub asrDummyCheckBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub
Private Sub asrDummyTextBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub
Private Sub asrDummyCombo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub
Private Sub asrDummySpinner_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub
Private Sub asrCustomDummyWP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub
Private Sub asrDummyPhoto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub
Private Sub asrDummyOLEContents_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub
Private Sub asrDummyLine_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub
Private Sub asrDummyLabel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub asrDummyLink_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub asrDummyImage_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub ScreenControl_MouseDown(pctlControl As VB.Control, piButton As Integer, piShift As Integer, pSngX As Single, pSngY As Single)

  Dim iCount As Integer

  ' Only handle left button presses here.
  If piButton <> vbLeftButton Then
    Exit Sub
  End If

  mlngXOffset = pSngX
  mlngYOffset = pSngY

  ' Flag the selected selction markers to be moved
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    ASRSelectionMarkers(iCount).ShowSelectionMarkers False
  Next iCount

End Sub

Public Function SelectControl(pctlControl As VB.Control) As Boolean
  
  Dim iIndex As Integer
  Dim iCount As Integer
  Dim objMarkers As Object
  
  ' Have selection markers for this control already been created
  If pctlControl.Selected Then
  
    If pctlControl.Tag = "" Then
    
      iIndex = ASRSelectionMarkers.Count
      Load ASRSelectionMarkers(iIndex)
      
      With ASRSelectionMarkers(iIndex)
        Set .Container = pctlControl.Container
        .AttachedObject = pctlControl
        
        If Not pctlControl.Name = "asrDummyFrame" Then
          .AttachedObject.ZOrder 0
        End If
          
        .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize, .AttachedObject.Width + (.MarkerSize * 2), .AttachedObject.Height + (.MarkerSize * 2)
        .RefreshSelectionMarkers True
        .ZOrder 0
        .Visible = True
      End With
    
      pctlControl.Tag = iIndex
    
    Else
      With ASRSelectionMarkers(pctlControl.Tag)
        'JPD 20050815 Fault 10245
        ' Ensure the selection markers are in the same container
        ' as the control - this can get out of synch sometimes.
        Set .Container = pctlControl.Container
        
        If Not .AttachedObject.Name = "asrDummyFrame" Then
          .AttachedObject.ZOrder 0
        End If
        
        .ZOrder 0
        .Visible = True
      End With
    End If

  End If
  
  SelectControl = True
  
End Function

Public Function LoadTabPage(piPageNumber As Integer) As Boolean
  ' Load controls onto the selected tab page.
  'On Error GoTo ErrorTrap
    
  Dim fLoadOk As Boolean
  Dim iPageNo As Integer
  Dim iCtrlType As Long
  Dim iDisplayType As Integer
  Dim lngTableID As Long
  Dim lngPictureID As Long
  Dim sFileName As String
  Dim sTableName As String
  Dim sColumnName As String
  Dim objFont As StdFont
  Dim ctlControl As VB.Control
  Dim iNextIndex As Integer
  Dim iRecordCount As Integer
  Dim iCount As Integer
  Dim iOriginalPageNumber As Integer
    
  iNextIndex = 1
  fLoadOk = True

  If gLngScreenID = 0 Then
    LoadTabPage = True
    Exit Function
  End If
 
  If tabPages.Tabs.Count > 0 Then
    iOriginalPageNumber = tabPages.Tabs(piPageNumber).Tag
  Else
    iOriginalPageNumber = 0
  End If

 'tabPages.SelectedItem.Index
 
  ' Have the controls on this page already been loaded
  If picPageContainer(iOriginalPageNumber).Tag = "loaded" And picPageContainer.Count > 0 Then
    LoadTabPage = True
    Exit Function
  End If

  Screen.MousePointer = vbHourglass
  
  ' Load the screen controls if everything is okay so far.
  If fLoadOk Then                     ' Indent 01 - start
  
    ' Locate the control definitions for the current screen.
    With recCtrlEdit                  ' Indent 02 - start
    
      iRecordCount = .RecordCount
    
      .Index = "idxTabIndex"
      .Seek ">=", gLngScreenID
      
      ' Flag this container has been loaded
       'If Not iOriginalPageNumber = 0 Then
         picPageContainer(iOriginalPageNumber).Tag = "loaded"
       'End If
      
      If Not .NoMatch Then            ' Indent 03 - start
      
        ' Add controls to the form for each control defined in the database.
        Do While Not .EOF             ' Indent 04 - start
        
          If .Fields("screenID").value <> gLngScreenID Then
            Exit Do
          End If
            
          ' Only Load controls for selected page
          If .Fields("pageNo").value = iOriginalPageNumber Then
                           
            ' Get the control's type.
            iCtrlType = IIf(IsNull(.Fields("controlType").value), giCTRL_TEXTBOX, .Fields("controlType").value)
              
            ' Create the new control.
            Set ctlControl = AddControl(iCtrlType)
      
            If Not ctlControl Is Nothing Then             ' Indent 05 - start

              ' Set the page container of the page that contains the control.
              iPageNo = IIf(IsNull(.Fields("pageNo").value), 0, .Fields("pageNo").value)
              If iPageNo = 0 Then
                Set ctlControl.Container = Me
              Else
                Set ctlControl.Container = picPageContainer(iOriginalPageNumber)
              End If

              ' Set the control's level in the z-order.
              'ctlControl.ControlLevel = .Fields("controlLevel").value

              ' Set the control's size.
'              ctlControl.Top = IIf(IsNull(.Fields("topCoord").value), 0, .Fields("topCoord").value)
'              ctlControl.Left = IIf(IsNull(.Fields("leftCoord").value), 0, .Fields("leftCoord").value)
'
'              ' Set the control's dimensions.
'              ctlControl.Height = IIf(IsNull(.Fields("height").value), 0, .Fields("height").value)
'              ctlControl.Width = IIf(IsNull(.Fields("width").value), 0, .Fields("width").value)
              ctlControl.Move IIf(IsNull(.Fields("leftCoord").value), 0, .Fields("leftCoord").value), IIf(IsNull(.Fields("topCoord").value), 0, .Fields("topCoord").value), _
                  IIf(IsNull(.Fields("width").value), 0, .Fields("width").value), IIf(IsNull(.Fields("height").value), 0, .Fields("height").value)


              ' Set the controls tab index.
              ctlControl.TabIndex = iNextIndex
              If (Not IsNull(.Fields("tabIndex").value)) And _
                (ScreenControl_IsTabStop(iCtrlType)) Then
                iNextIndex = iNextIndex + 1
              End If

              ' Set the control's column and table IDs.
              lngTableID = IIf(IsNull(.Fields("tableID").value), 0, .Fields("tableID").value)
              ctlControl.ColumnID = IIf(IsNull(.Fields("columnID").value), 0, .Fields("columnID").value)

              ' Give the control a tooltip if it is associated with a column.
              With recColEdit
                .Index = "idxColumnID"
                .Seek "=", ctlControl.ColumnID

                If Not .NoMatch Then
                  sColumnName = .Fields("columnName").value

                  With recTabEdit
                    .Index = "idxTableID"
                    .Seek "=", lngTableID

                    If Not .NoMatch Then
                      sTableName = .Fields("tableName").value
                      ctlControl.ToolTipText = sTableName & "." & sColumnName
                    End If

                  End With
                End If
              End With

              ' Set the controls caption.
              If (ScreenControl_HasCaption(iCtrlType)) Then
                ctlControl.Caption = IIf(IsNull(.Fields("caption").value), "", .Fields("caption").value & vbNullString)
              End If

              If (ScreenControl_HasText(iCtrlType)) Then
                ctlControl.Caption = ctlControl.ToolTipText
                If iCtrlType = giCTRL_OLE Then
                  ctlControl.ButtonCaption = OLEType(ctlControl.ColumnID)
                End If
              End If

              ' Set the BackColor and ForeColor properties.
              If ScreenControl_HasBackColor(iCtrlType) Then
                ctlControl.BackColor = IIf(IsNull(.Fields("backColor").value), Me.BackColor, .Fields("backColor").value)
              End If

              If ScreenControl_HasForeColor(iCtrlType) Then
                ctlControl.ForeColor = IIf(IsNull(.Fields("foreColor").value), Me.ForeColor, .Fields("foreColor").value)
              End If

              ' Font properties.
              If ScreenControl_HasFont(iCtrlType) Then
                Set objFont = New StdFont
                objFont.Name = IIf(IsNull(.Fields("fontName").value), gobjDefaultScreenFont.Name, .Fields("fontName").value)
                objFont.Size = IIf(IsNull(.Fields("fontSize").value), gobjDefaultScreenFont.Size, .Fields("fontSize").value)
                objFont.Bold = IIf(IsNull(.Fields("fontBold").value), False, .Fields("fontBold").value)
                objFont.Italic = IIf(IsNull(.Fields("fontItalic").value), False, .Fields("fontItalic").value)
                objFont.Strikethrough = IIf(IsNull(.Fields("fontStrikeThru").value), False, .Fields("fontStrikeThru").value)
                objFont.Underline = IIf(IsNull(.Fields("fontUnderline").value), False, .Fields("fontUnderline").value)
                Set ctlControl.Font = objFont
                Set objFont = Nothing
              End If

              ' Set the BorderStyle property.
              If ScreenControl_HasBorderStyle(iCtrlType) Then
                ctlControl.BorderStyle = IIf(IsNull(.Fields("borderStyle").value), vbFixedSingle, .Fields("borderStyle").value)
              End If


              'NPG20071023
              ' Set the ReadOnly property.
              If ScreenControl_HasReadOnly(iCtrlType) Then
                ctlControl.Read_Only = IIf(IsNull(.Fields("readOnly").value), False, .Fields("readOnly").value)
              End If


              ' Set the Alignment property.
              If ScreenControl_HasAlignment(iCtrlType) Then
                If Not IsNull(.Fields("alignment").value) Then
                  ctlControl.Alignment = .Fields("alignment").value
                End If
              End If

              ' Set the Orientation property.
              If ScreenControl_HasOrientation(iCtrlType) Then
                If Not IsNull(.Fields("alignment").value) Then
                  ctlControl.Alignment = .Fields("alignment").value

                  ' Height/Width required to be set again after alignment property...
                  ' Bug in the line control..only happens in ScrDsgnr, not in the
                  ' test project !
                  If iCtrlType = giCTRL_LINE Then
                    ctlControl.Height = IIf(IsNull(.Fields("height").value), 0, .Fields("height").value)
                    ctlControl.Width = IIf(IsNull(.Fields("width").value), 0, .Fields("width").value)
                  End If

                End If
              End If

              ' Set the Picture property.
              If ScreenControl_HasPicture(iCtrlType) Then
                ctlControl.PictureID = IIf(IsNull(.Fields("pictureID").value), 0, .Fields("pictureID").value)
                If ctlControl.PictureID > 0 Then

                  recPictEdit.Index = "idxID"
                  recPictEdit.Seek "=", ctlControl.PictureID

                  If Not recPictEdit.NoMatch Then
                    sFileName = ReadPicture
                    ctlControl.Picture = sFileName
                    Kill sFileName
                  End If

                End If
              End If

              ' Set the control's Options property.
              If ScreenControl_HasOptions(iCtrlType) Then
                recColEdit.Index = "idxColumnID"
                recColEdit.Seek "=", .Fields("columnID").value

                If Not recColEdit.NoMatch Then
                  'ctlControl.Options = ReadColumnControlValues(recColEdit.Fields("columnID"))
                  ctlControl.SetOptions ReadColumnControlValues(recColEdit.Fields("columnID").value)
                End If
              End If

              ' Set the controls Display type properties
              If ScreenControl_HasDisplayType(iCtrlType) Then
                ctlControl.DisplayType = IIf(IsNull(.Fields("DisplayType").value), NavigationDisplayType.Button, .Fields("DisplayType").value)
              End If

              ' Set the controls navigate properties
              If ScreenControl_HasNavigation(iCtrlType) Then
                ctlControl.ColumnName = GetColumnName(ctlControl.ColumnID, False)
                ctlControl.NavigateTo = IIf(IsNull(.Fields("NavigateTo").value), vbNullString, .Fields("NavigateTo").value)
                ctlControl.NavigateIn = IIf(IsNull(.Fields("NavigateIn").value), NavigateIn.URL, .Fields("NavigateIn").value)
                ctlControl.NavigateOnSave = IIf(IsNull(.Fields("NavigateOnSave").value), vbNo, .Fields("NavigateOnSave").value)
              End If


            'TM20010914 Fault 1753
            'The ActiveBar control does mot have the visible property, so to avoid err
            'we only check the visible property of other controls.
              If ctlControl.Name <> "abScreen" Then
                ctlControl.Visible = True
              End If
            End If       ' Indent 05 - end

            ' Disassociate object variables.
            Set ctlControl = Nothing

          End If
          
          .MoveNext
        Loop       ' Indent 04 - end
      End If       ' Indent 03 - end
    End With       ' Indent 02 - end
    
    ' Set the correct z-order for each control.
    fLoadOk = SetControlLevel
    
  End If       ' Indent 01 - end

TidyUpAndExit:

  ' Unlock the window refreshing.
  UI.UnlockWindow
    
  ' Reset the screen moousepointer.
  Screen.MousePointer = vbDefault
  
  LoadTabPage = fLoadOk
  Exit Function
  
ErrorTrap:
  fLoadOk = False
  MsgBox "Error loading Screen." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit

End Function

Private Function ScreenControl_KeyMove(pSngX As Single, pSngY As Single) As Boolean
  ' Move the control.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iCount As Integer
  
  fOK = True
  
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
    
        If .AttachedObject.Selected Then
          .AttachedObject.Move pSngX + .AttachedObject.Left, pSngY + .AttachedObject.Top
        End If
      
      End If
    End With
  Next iCount
  
  ' Flag screen as having changed
  IsChanged = True

TidyUpAndExit:
  ScreenControl_KeyMove = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

' Scroll through each selected control and send to back
Private Function SendSelectedControlsToBack()

  Dim iCount As Integer
  
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .AttachedObject.ZOrder 1
      End If
    End With
  Next iCount

End Function

' Scroll through each selected control and bring to front
Private Function BringSelectedControlsToFront()

  Dim iCount As Integer
  
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .AttachedObject.ZOrder 0
      End If
    End With
  Next iCount

End Function

Private Function DeselectControl(pctlControl As VB.Control) As Boolean
 
  ' Deselect current control
  ASRSelectionMarkers(pctlControl.Tag).Visible = False
  pctlControl.Selected = False

  DeselectControl = True
  
End Function

' Does this screen have any user controls on it
Public Function ScreenHasControls() As Boolean

  Dim ctlControl As VB.Control
  
  For Each ctlControl In Me.Controls
    If IsScreenControl(ctlControl) Then
      ScreenHasControls = True
      Exit Function
    End If
  Next ctlControl

End Function


' Left align the selected controls
Private Function LeftAlignSelectedControls()

  Dim iCount As Integer
  Dim lngLeft As Long
  Dim lngTop As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
        lngLeft = .Left
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Left = lngLeft
        .AttachedObject.Left = lngLeft + .MarkerSize
        Application.Changed = True
      End If
    End With
  Next iCount
    
  'NHRD02122002 Fault 4675
  ' Mark the screen as having changed.
  gfChangedScreen = True
  frmSysMgr.RefreshMenu

End Function


' Right align the selected controls
Private Function RightAlignSelectedControls()

  Dim iCount As Integer
  Dim lngRight As Long
  Dim lngTop As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
        lngRight = .Left + .Width
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Left = lngRight - .Width
        .AttachedObject.Left = .Left + .MarkerSize
      End If
    End With
  Next iCount

  'NHRD02122002 Fault 4675
  ' Mark the screen as having changed.
  gfChangedScreen = True
  frmSysMgr.RefreshMenu

End Function


' Centre align the selected controls
Private Function CentreAlignSelectedControls()

  Dim iCount As Integer
  Dim lngCentre As Long
  Dim lngTop As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
        lngCentre = .Left + (.Width / 2)
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Left = lngCentre - (.Width / 2)
        .AttachedObject.Left = .Left + .MarkerSize
      End If
    End With
  Next iCount

  'NHRD02122002 Fault 4675
  ' Mark the screen as having changed.
  gfChangedScreen = True
  frmSysMgr.RefreshMenu

End Function

' Top align the selected controls
Private Function TopAlignSelectedControls()

  Dim iCount As Integer
  Dim lngTop As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Top = lngTop
        .AttachedObject.Top = .Top + .MarkerSize
      End If
    End With
  Next iCount

  'NHRD02122002 Fault 4675
  ' Mark the screen as having changed.
  gfChangedScreen = True
  frmSysMgr.RefreshMenu

End Function

' Middle align the selected controls
Private Function MiddleAlignSelectedControls()

  Dim iCount As Integer
  Dim lngTop As Long
  Dim lngMiddle As Long

  'Find out the topmost control - this is used as the align point
  lngTop = 9999999
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top < lngTop Then
        lngTop = .Top
        lngMiddle = .Top + (.Height / 2)
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Top = lngMiddle - (.Height / 2)
        .AttachedObject.Top = (lngMiddle - (.Height / 2)) + .MarkerSize
      End If
    End With
  Next iCount

  'NHRD02122002 Fault 4675
  ' Mark the screen as having changed.
  gfChangedScreen = True
  frmSysMgr.RefreshMenu

End Function

' Bottom align the selected controls
Private Function BottomAlignSelectedControls()

  Dim iCount As Integer
  Dim lngBottom As Long

  'Find out the bottom most control - this is used as the align point
  lngBottom = 0
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible And .Top + .Height > lngBottom Then
        lngBottom = .Top + .Height
      End If
    End With
  Next iCount

  'Move the left property to everything matches
  For iCount = 1 To ASRSelectionMarkers.Count - 1
    With ASRSelectionMarkers(iCount)
      If .Visible Then
        .Top = lngBottom - .Height
        .AttachedObject.Top = (lngBottom - .Height) + .MarkerSize
      End If
    End With
  Next iCount

  'NHRD02122002 Fault 4675
  ' Mark the screen as having changed.
  gfChangedScreen = True
  frmSysMgr.RefreshMenu

End Function

' Default column width following font change to Verdana (Textbox)
Public Function Default_ColumnWidth_Textbox(ByRef plngColumnWidth As Long) As Long
  Default_ColumnWidth_Textbox = CLng(((plngColumnWidth + 1) * 95 + 105) / 10) * 10
End Function

' Default column width following font change to Verdana (Textbox)
Public Function Default_ColumnWidth_Numeric(ByRef plngNumeric As Long, ByRef plngDecimals As Long, ByRef pbSeperators As Boolean) As Long

  Dim lngSeperators As Long
  Dim lngWidth As Long

  lngSeperators = 60 * IIf(pbSeperators, plngNumeric / 3, 0)
  lngWidth = plngNumeric + IIf(plngDecimals > 0, plngDecimals + 1, 0) + 1

  Default_ColumnWidth_Numeric = (plngNumeric * 105) + 120 + 60 + lngSeperators
End Function




