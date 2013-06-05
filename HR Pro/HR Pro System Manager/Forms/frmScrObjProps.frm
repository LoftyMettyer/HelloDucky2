VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{051CE3FC-5250-4486-9533-4E0723733DFA}#1.0#0"; "COA_ColourPicker.ocx"
Begin VB.Form frmScrObjProps 
   Caption         =   "Control Properties"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1030
   Icon            =   "frmScrObjProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   4620
   Visible         =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   4230
      Width           =   4620
      _ExtentX        =   8149
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
      Left            =   45
      TabIndex        =   0
      Top             =   510
      Width           =   4545
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   3
      stylesets.count =   4
      stylesets(0).Name=   "ssetBackColorValue"
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
      stylesets(0).Picture=   "frmScrObjProps.frx":000C
      stylesets(1).Name=   "ssetForeColorValue"
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
      stylesets(1).Picture=   "frmScrObjProps.frx":0028
      stylesets(2).Name=   "ssetActiveRow"
      stylesets(2).ForeColor=   -2147483639
      stylesets(2).BackColor=   -2147483646
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
      stylesets(2).Picture=   "frmScrObjProps.frx":0044
      stylesets(3).Name=   "ssetDormantRow"
      stylesets(3).ForeColor=   0
      stylesets(3).BackColor=   16777215
      stylesets(3).HasFont=   -1  'True
      BeginProperty stylesets(3).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(3).Picture=   "frmScrObjProps.frx":0060
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
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "Property"
      Columns(0).Name =   "colProperties"
      Columns(0).AllowSizing=   0   'False
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(0).Style=   4
      Columns(0).BackColor=   16777215
      Columns(1).Width=   3200
      Columns(1).Caption=   "Value"
      Columns(1).Name =   "colValues"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   32767
      Columns(1).Locked=   -1  'True
      Columns(1).Style=   3
      Columns(1).Row.Count=   2
      Columns(1).Col.Count=   2
      Columns(1).Row(0).Col(0)=   "True"
      Columns(1).Row(1).Col(0)=   "False"
      Columns(1).BackColor=   16777215
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Tag"
      Columns(2).Name =   "colTag"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      UseDefaults     =   0   'False
      _ExtentX        =   8017
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
   Begin COAColourPicker.COA_ColourPicker ColorPicker 
      Left            =   720
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblSizeTester 
      Caption         =   "<Size Tester>"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1665
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmScrObjProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Properties.
Private gFrmScreen As frmScrDesigner2
Private giScreenCount As Integer

' Globals.
Dim giCXFrame As Integer
Dim giCYFrame As Integer
Dim giCXBorder As Integer
Dim giCYBorder As Integer
Dim giAlignment As Integer
Dim giOrientation As Integer
Dim gColBackColor As OLE_COLOR
Dim giBorderStyle As Integer
Dim mblnReadOnly As Boolean   'NPG20071022

Dim gsCaption As String
Dim gsNavigateTo As String
Dim giDisplayType As NavigationDisplayType
Dim giNavigateIn As NavigateIn
Dim gbNavigateOnSave As Boolean
Dim gObjFont As StdFont
Dim gColForeColor As OLE_COLOR
Dim gLngHeight As Long
Dim gLngLeft As Long
Dim glngPictureID As Long
Dim gLngTop As Long
Dim gLngWidth As Long
Dim glngTabNumber As Long

' Alignment text constants.
Const gsALIGNMENTLEFTTEXT = "Left Alignment"
Const gsALIGNMENTRIGHTTEXT = "Right Alignment"
' Orientation text constants.
Const gsORIENTATIONVERTICAL = "Vertical"
Const gsORIENTATIONHORIZONTAL = "Horizontal"
' BorderStyle text constants.
Const gsBORDERSTYLENONETEXT = "No"
Const gsBORDERSTYLEFIXEDSINGLETEXT = "Yes"

'NPG20071022 - ReadOnly text constants
Const gsREADONLYNOTEXT = "No"
Const gsREADONLYYESTEXT = "Yes"

' Display Type text constants.
Const gsDISPLAYTYPE_HYPERLINK = "Hyperlink"
Const gsDISPLAYTYPE_BUTTON = "Button"
Const gsDISPLAYTYPE_BROWSER = "Browser"
Const gsDISPLAYTYPE_HIDDEN = "Hidden"

' Navigate In constants
Const gsNAVIGATE_URL = "External"
Const gsNAVIGATE_MENUBAR = "Internal"

' Picture text constants.
Const gsPICTURENONETEXT = "(None)"
Const gsPICTURESELECTEDTEXT = "(Picture)"

' Property type constants.
Const giPROPID_ALIGNMENT = 1
Const giPROPID_BACKCOLOR = 2
Const giPROPID_BORDERSTYLE = 3
Const giPROPID_CAPTION = 4
Const giPROPID_DISPLAYTYPE = 5
Const giPROPID_FONT = 6
Const giPROPID_FORECOLOR = 7
Const giPROPID_HEIGHT = 8
Const giPROPID_LEFT = 9
Const giPROPID_PICTURE = 10
Const giPROPID_TOP = 11
Const giPROPID_WIDTH = 12
Const giPROPID_ORIENTATION = 13
Const giPROPID_TABNUMBER = 14
Const giPROPID_READONLY = 15      'NPG20071022
Const giPROPID_NAVIGATETO = 16
Const giPROPID_NAVIGATEIN = 17
Const giPROPID_NAVIGATEONSAVE = 18

Private Const MIN_FORM_HEIGHT = 2850
Private Const MIN_FORM_WIDTH = 2850

Private miCurrentRowFormat As Integer
Public Property Get CurrentScreen() As frmScrDesigner2
  ' Return the current screen.
  Set CurrentScreen = gFrmScreen
  
End Property

Public Property Set CurrentScreen(pFrmNewValue As frmScrDesigner2)
  ' Set the current screen property of the form.
  Set gFrmScreen = pFrmNewValue
  
End Property

Public Property Get ScreenCount() As Integer
  ' return the number of screen designers that are open.
  ScreenCount = giScreenCount
  
End Property

Public Property Let ScreenCount(piNewValue As Integer)
  ' Set the number of screen designers that are open.
  giScreenCount = piNewValue
  If giScreenCount < 1 Then
    UnLoad Me
  End If
  
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  Dim bHandled As Boolean
'
'  bHandled = frmSysMgr.tbMain.OnKeyDown(KeyCode, Shift)
'  If bHandled Then
'    KeyCode = 0
'    Shift = 0
'  End If

End Sub

Private Sub Form_Load()
  On Error GoTo ErrorTrap
  
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
  
  'Get dimensions of windows borders
  giCXFrame = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
  giCYFrame = UI.GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY
  giCXBorder = UI.GetSystemMetrics(SM_CXBORDER) * Screen.TwipsPerPixelX
  giCYBorder = UI.GetSystemMetrics(SM_CYBORDER) * Screen.TwipsPerPixelY
  
  'Position and size form
  Me.Move 0, 0, 3750, 4250

  ' Format properties grid.
  With ssGridProperties
    .Top = 0
    .Left = 0

    .Columns(0).Width = 1600
'    .Columns(1).Width = (.Width - .Columns(0).Width) - giCXFrame
  End With
  
  ' Resize the form
  Form_Resize
  
  ' Position the form at the right hand side of the screen.
  Me.Left = Screen.Width - Me.Width - giCXFrame
  
TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_Resize()
  On Error GoTo ErrorTrap

  Dim lngColumnWidth  As Long
  
  'If Not Initialising And Me.WindowState <> vbMinimized Then
  If Me.WindowState <> vbMinimized Then
    
    With ssGridProperties
      .Width = Me.ScaleWidth
      .Height = Me.ScaleHeight - .Top - StatusBar1.Height
      
      lngColumnWidth = .Width - .Columns(0).Width - (giCXBorder * 2)
    
      ' Cater for the display of the vertical scroll bar.
'      If .Rows > .VisibleRows Then
'        lngColumnWidth = lngColumnWidth - _
'          (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX)
'      End If
    
      .Columns(1).Width = lngColumnWidth
    End With
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
  Set gObjFont = Nothing
  Set frmScrObjProps = Nothing

  Unhook Me.hWnd

End Sub


Private Sub ssGridProperties_BeforeRowColChange(Cancel As Integer)
  'JPD 20050812 Fault 10175
  ' Ensure the grid has the correct column properties applied for the selected property.
  ' Note that fault 10175 occurred because the ComboDropDown event fires BEFORE the rowColChange event.
  ' The BeforeRowColChange event fires BEFORE the ComboDropDown event, so we'll refresh the column properties here if we need to.
  ' Note that we only call RowColChange if we need to, otherwise infinite looping occurs!
  If miCurrentRowFormat <> Val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark)) Then
    ssGridProperties_RowColChange 0, 0
  End If

End Sub

Private Sub ssGridProperties_Change()

  ' RH 08/08/00 - FAULT 55 - Captions now update in real time (ie, as you
  '                          type them into the properties window.
  
  If Val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark)) = giPROPID_CAPTION Then
    gsCaption = ssGridProperties.ActiveCell.Text
    UpdateControls giPROPID_CAPTION
  End If

'MH20060915 Fault 11422
'  If Val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark)) = giPROPID_TABNUMBER Then
'    If ValidIntegerString(ssGridProperties.ActiveCell.Text) Then
'      glngTabNumber = Val(ssGridProperties.ActiveCell.Text)
'      UpdateControls giPROPID_TABNUMBER
'    End If
'
'  End If
    
End Sub

Private Sub ssGridProperties_BeforeUpdate(Cancel As Integer)
  ' Update the selected controls with the new value.
  On Error GoTo ErrorTrap
  
  ' Debug.Print "ssGridProperties_BeforeUpdate"
 
 
  Dim fUpdateControls As Boolean
  Dim iPropertyTag As Integer
  Dim sNewValue As String

  ' Read the new property value from the grid.
  iPropertyTag = Val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  sNewValue = ssGridProperties.ActiveCell.Text
  fUpdateControls = True
  
  ' Update the required global variable with the new value if required.
  Select Case iPropertyTag
    Case giPROPID_ALIGNMENT
        
    Case giPROPID_BACKCOLOR
        
    Case giPROPID_BORDERSTYLE
    
    Case giPROPID_READONLY      'NPG20071022
        
    Case giPROPID_CAPTION
      gsCaption = sNewValue
        
    Case giPROPID_DISPLAYTYPE
        
    Case giPROPID_NAVIGATETO
      gsNavigateTo = sNewValue
        
    Case giPROPID_NAVIGATEONSAVE
        
    Case giPROPID_FONT
        
    Case giPROPID_FORECOLOR
        
    Case giPROPID_HEIGHT
      If ValidIntegerString(sNewValue) Then
        gLngHeight = Val(sNewValue)
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(gLngHeight))
        fUpdateControls = False
      End If
      
    Case giPROPID_LEFT
      If ValidIntegerString(sNewValue) Then
        gLngLeft = Val(sNewValue)
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(gLngLeft))
        fUpdateControls = False
      End If
        
    Case giPROPID_PICTURE
        
    Case giPROPID_TOP
      If ValidIntegerString(sNewValue) Then
        gLngTop = Val(sNewValue)
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(gLngTop))
        fUpdateControls = False
      End If
        
    Case giPROPID_WIDTH
      If ValidIntegerString(sNewValue) Then
        gLngWidth = Val(sNewValue)
      Else
        ssGridProperties.ActiveCell.Text = Trim(Str(gLngWidth))
        fUpdateControls = False
      End If
      
    Case giPROPID_TABNUMBER
      If ValidIntegerString(sNewValue) Then
        glngTabNumber = Val(sNewValue)
        
        'MH20060915 Fault 11422
        UpdateControls giPROPID_TABNUMBER
      Else
        fUpdateControls = False
      End If
    
    Case Else
      
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

  Dim iPropertyTag As Integer
  
  iPropertyTag = Val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  
  ' Display the form required for changing the current property.
  ' Read the new property value from the form into a global variable.
  Select Case iPropertyTag
    
    Case giPROPID_BACKCOLOR
      ' Display the Background Colour dialogue box.
'      With comDlgBox
'        .Flags = cdlCCRGBInit
'        .Color = gColBackColor
'        .ShowColor
'        gColBackColor = .Color
'      End With

    With ColorPicker
      .Color = gColBackColor
      .ShowPalette
      gColBackColor = .Color
    End With

      ' Update the grid display.
      With ssGridProperties
        .StyleSets("ssetBackColorValue").BackColor = gColBackColor
        .Columns(1).CellStyleSet "ssetBackColorValue", .Row
      End With

    Case giPROPID_FONT
      ' Display the Font dialogue box.
      With comDlgBox
        .FontName = gObjFont.Name
        .FontSize = gObjFont.Size
        .FontBold = gObjFont.Bold
        .FontItalic = gObjFont.Italic
        .FontUnderline = gObjFont.Underline
        .FontStrikethru = gObjFont.Strikethrough
        .Flags = cdlCFScreenFonts Or cdlCFEffects
        .ShowFont
        gObjFont.Name = .FontName
        gObjFont.Size = .FontSize
        gObjFont.Bold = .FontBold
        gObjFont.Italic = .FontItalic
        gObjFont.Underline = .FontUnderline
        gObjFont.Strikethrough = .FontStrikethru
      End With
      ' Update the grid display.
      ssGridProperties.Columns(1).Text = GetFontDescription
      
    Case giPROPID_FORECOLOR
      ' Display the Foreground Colour dialogue box.
'      With comDlgBox
'        .Flags = cdlCCRGBInit
'        .Color = gColForeColor
'        .ShowColor
'        gColForeColor = .Color
'      End With

      ' AE20080331 Fault #4604, #10170
      With ColorPicker
        .Color = gColForeColor
        .ShowPalette
        gColForeColor = .Color
      End With

      ' Update the grid display.
      With ssGridProperties
        .StyleSets("ssetForeColorValue").BackColor = gColForeColor
        .Columns(1).CellStyleSet "ssetForeColorValue", .Row
      End With
      
    Case giPROPID_PICTURE
      ' Display the Picture selection form.
      frmPictSel.SelectedPicture = glngPictureID
      frmPictSel.ExcludedExtensions = ".gif"
      frmPictSel.Show vbModal
      If frmPictSel.SelectedPicture > 0 Then
        With recPictEdit
          .Index = "idxID"
          .Seek "=", frmPictSel.SelectedPicture
          If Not .NoMatch Then
            glngPictureID = !PictureID
            ssGridProperties.Columns(1).Text = !Name
          End If
        End With
      End If
      
      Set frmPictSel = Nothing
        
  End Select

  ' Update the selected controls with the new property value.
  UpdateControls iPropertyTag

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  ' User pressed cancel.
  Err = False
  Resume TidyUpAndExit
  
End Sub


Private Sub ssGridProperties_Click()

  frmScrObjProps.SetFocus
  frmSysMgr.RefreshMenu
  
End Sub

Private Sub ssGridProperties_ComboCloseUp()
  ' Process the combo selection in the properties grid.
  On Error GoTo ErrorTrap
  
  ' Debug.Print "ssGridProperties_ComboCloseUp"
  
  Dim iPropertyTag As Integer
  Dim sNewValue As String
  
  iPropertyTag = Val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  sNewValue = ssGridProperties.ActiveCell.Text
  
  ' Read the new property value from the grid.
  Select Case iPropertyTag
    
    Case giPROPID_ALIGNMENT
      If sNewValue = gsALIGNMENTLEFTTEXT Then
        giAlignment = vbLeftJustify
      Else
        giAlignment = vbRightJustify
      End If
          
    Case giPROPID_BORDERSTYLE
      If sNewValue = gsBORDERSTYLENONETEXT Then
        giBorderStyle = vbBSNone
      Else
        giBorderStyle = vbFixedSingle
      End If
          
    Case giPROPID_READONLY
      If sNewValue = gsREADONLYNOTEXT Then
        mblnReadOnly = False
      Else
        mblnReadOnly = True
      End If

    Case giPROPID_DISPLAYTYPE
      Select Case sNewValue
        Case gsDISPLAYTYPE_HYPERLINK
          giDisplayType = NavigationDisplayType.Hyperlink
        Case gsDISPLAYTYPE_BUTTON
          giDisplayType = NavigationDisplayType.Button
        Case gsDISPLAYTYPE_BROWSER
          giDisplayType = NavigationDisplayType.Browser
        Case gsDISPLAYTYPE_HIDDEN
          giDisplayType = NavigationDisplayType.Hidden
      End Select

    Case giPROPID_NAVIGATEIN
      Select Case sNewValue
        Case gsNAVIGATE_URL
          giNavigateIn = NavigateIn.URL
        Case gsNAVIGATE_MENUBAR
          giNavigateIn = NavigateIn.MenuBar
      End Select

    Case giPROPID_NAVIGATEONSAVE
      If sNewValue = gsREADONLYNOTEXT Then
        gbNavigateOnSave = False
      Else
        gbNavigateOnSave = True
      End If

    Case giPROPID_ORIENTATION
      If sNewValue = gsORIENTATIONVERTICAL Then
        giOrientation = 0
      Else
        giOrientation = 1
      End If
                  
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
  
  Dim iPropertyTag As Integer
  
  ' Debug.Print "ssGridProperties_DblClick"
  
  
  iPropertyTag = Val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  
  ' Read the new property value from the grid.
  Select Case iPropertyTag
    
    Case giPROPID_ALIGNMENT
      If giAlignment = vbRightJustify Then
        giAlignment = vbLeftJustify
        ssGridProperties.Columns(1).Text = gsALIGNMENTLEFTTEXT
      Else
        giAlignment = vbRightJustify
        ssGridProperties.Columns(1).Text = gsALIGNMENTRIGHTTEXT
      End If
      
    Case giPROPID_BACKCOLOR
      ssGridProperties_BtnClick
      
    Case giPROPID_BORDERSTYLE
      If giBorderStyle = vbFixedSingle Then
        giBorderStyle = vbBSNone
        ssGridProperties.Columns(1).Text = gsBORDERSTYLENONETEXT
      Else
        giBorderStyle = vbFixedSingle
        ssGridProperties.Columns(1).Text = gsBORDERSTYLEFIXEDSINGLETEXT
      End If


    'NPG20071022
    Case giPROPID_READONLY
      If mblnReadOnly = True Then
        mblnReadOnly = False
        ssGridProperties.Columns(1).Text = gsREADONLYNOTEXT
      Else
        mblnReadOnly = True
        ssGridProperties.Columns(1).Text = gsREADONLYYESTEXT
      End If

    Case giPROPID_DISPLAYTYPE
      Select Case giDisplayType
        Case NavigationDisplayType.Hyperlink
          ssGridProperties.Columns(1).Text = gsDISPLAYTYPE_HYPERLINK
        Case NavigationDisplayType.Button
          ssGridProperties.Columns(1).Text = gsDISPLAYTYPE_BUTTON
        Case NavigationDisplayType.Browser
          ssGridProperties.Columns(1).Text = gsDISPLAYTYPE_BROWSER
        Case NavigationDisplayType.Hidden
          ssGridProperties.Columns(1).Text = gsDISPLAYTYPE_HIDDEN
      End Select

    Case giPROPID_NAVIGATEONSAVE
      If gbNavigateOnSave = True Then
        ssGridProperties.Columns(1).Text = gsREADONLYNOTEXT
      Else
        ssGridProperties.Columns(1).Text = gsREADONLYYESTEXT
      End If

    Case giPROPID_FONT
      ssGridProperties_BtnClick
    
    Case giPROPID_FORECOLOR
      ssGridProperties_BtnClick
    
    Case giPROPID_ORIENTATION
      If giOrientation = 0 Then
        giOrientation = 1
        ssGridProperties.Columns(1).Text = gsORIENTATIONHORIZONTAL
      Else
        giOrientation = 0
        ssGridProperties.Columns(1).Text = gsORIENTATIONVERTICAL
      End If
    
    Case giPROPID_PICTURE
      ssGridProperties_BtnClick
    
  End Select

  ' Update the selected controls with the new property value.
  UpdateControls iPropertyTag

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ssGridProperties_KeyPress(KeyAscii As Integer)
  ' Process the ENTER key press.
  On Error GoTo ErrorTrap
  
  If KeyAscii = vbKeyReturn Then
    
    ssGridProperties_BeforeUpdate 0
    
'    With ssGridProperties
'      .ActiveCell.SelStart = 0
'      .ActiveCell.SelLength = Len(.ActiveCell.Text)
'    End With
   
   If Val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark)) = giPROPID_CAPTION Then
      With ssGridProperties
        .ActiveCell.SelStart = Len(.ActiveCell.Text)
      End With
    Else
      With ssGridProperties
        .ActiveCell.SelStart = 0
        .ActiveCell.SelLength = Len(.ActiveCell.Text)
      End With
    End If
   
  End If
  
TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ssGridProperties_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorTrap
  
  ' RH23022000 - Stop the propery window from selected all the text if the user
  '              is editing the caption, height, left, top, or width property of
  '              a control
  
  Dim iPropertyTag As Integer
  iPropertyTag = Val(ssGridProperties.Columns(2).CellText(ssGridProperties.Bookmark))
  
  Select Case iPropertyTag
    Case giPROPID_CAPTION, giPROPID_HEIGHT, giPROPID_LEFT, giPROPID_TOP, giPROPID_WIDTH, giPROPID_TABNUMBER, giPROPID_NAVIGATETO
    Case Else
      With ssGridProperties
        .ActiveCell.SelStart = 0
        .ActiveCell.SelLength = Len(.ActiveCell.Text)
      End With
  End Select

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub ssGridProperties_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  ' Configure the grid for the currently selected row.
  On Error GoTo ErrorTrap

  Dim iLoop As Integer

  ' Debug.Print "ssGridProperties_RowColChange"
  
  
  With ssGridProperties
    ' Set the styleSet of the rows to show which is selected.
    For iLoop = 0 To ssGridProperties.Rows
      If iLoop = .Row Then
        .Columns(0).CellStyleSet "ssetActiveRow", iLoop
      Else
        .Columns(0).CellStyleSet "ssetDormatRow", iLoop
      End If
    Next iLoop

    miCurrentRowFormat = Val(.Columns(2).CellText(.Bookmark))
    
    Select Case Val(.Columns(2).CellText(.Bookmark))
      Case giPROPID_ALIGNMENT
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).List(0) = gsALIGNMENTLEFTTEXT
        .Columns(1).List(1) = gsALIGNMENTRIGHTTEXT

      Case giPROPID_BACKCOLOR
        .Columns(1).Style = ssStyleEditButton
        .Columns(1).Locked = True

      Case giPROPID_BORDERSTYLE
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).List(0) = gsBORDERSTYLENONETEXT
        .Columns(1).List(1) = gsBORDERSTYLEFIXEDSINGLETEXT
      
      Case giPROPID_READONLY
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem gsREADONLYNOTEXT, False
        .Columns(1).AddItem gsREADONLYYESTEXT, True
      
      Case giPROPID_CAPTION
        .Columns(1).Locked = False
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEdit

      Case giPROPID_DISPLAYTYPE
        .Columns(1).Locked = False
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem gsDISPLAYTYPE_HYPERLINK, NavigationDisplayType.Hyperlink
        .Columns(1).AddItem gsDISPLAYTYPE_BUTTON, NavigationDisplayType.Button
        .Columns(1).AddItem gsDISPLAYTYPE_BROWSER, NavigationDisplayType.Browser
        .Columns(1).AddItem gsDISPLAYTYPE_HIDDEN, NavigationDisplayType.Hidden

      Case giPROPID_NAVIGATEIN
        .Columns(1).Locked = False
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem gsNAVIGATE_URL, NavigateIn.URL
        .Columns(1).AddItem gsNAVIGATE_MENUBAR, NavigateIn.MenuBar

      Case giPROPID_NAVIGATETO
        .Columns(1).Locked = False
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEdit

      Case giPROPID_NAVIGATEONSAVE
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).RemoveAll
        .Columns(1).AddItem gsREADONLYNOTEXT, False
        .Columns(1).AddItem gsREADONLYYESTEXT, True

      Case giPROPID_FONT
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEditButton

      Case giPROPID_FORECOLOR
        .Columns(1).Locked = True
        .Columns(1).Style = ssStyleEditButton

      Case giPROPID_HEIGHT
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      Case giPROPID_LEFT
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      Case giPROPID_ORIENTATION
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleComboBox
        .Columns(1).List(0) = gsORIENTATIONVERTICAL
        .Columns(1).List(1) = gsORIENTATIONHORIZONTAL

      Case giPROPID_PICTURE
        .Columns(1).Locked = True
        .Columns(1).DataType = vbString
        .Columns(1).Style = ssStyleEditButton

      Case giPROPID_TOP
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      Case giPROPID_WIDTH
        .Columns(1).Locked = False
        .Columns(1).DataType = vbLong
        .Columns(1).Style = ssStyleEdit

      Case giPROPID_TABNUMBER
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

'   If Val(.Columns(2).CellText(.Bookmark)) <> giPROPID_CAPTION Then
   If Val(.Columns(2).CellText(.Bookmark)) = giPROPID_CAPTION Then
      With ssGridProperties
        '.ActiveCell.SelStart = 0
        .ActiveCell.SelStart = Len(.ActiveCell.Text)
        '.ActiveCell.SelLength = Len(.ActiveCell.Text)
      End With
    Else
      With ssGridProperties
        .ActiveCell.SelStart = 0
        .ActiveCell.SelLength = Len(.ActiveCell.Text)
      End With

    End If

  End With

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit

End Sub

Public Function EditMenu(psMenuOption As String) As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Pass all menu evants onto the active screen designer.
  gFrmScreen.EditMenu psMenuOption
  
TidyUpAndExit:
  EditMenu = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function



Public Function RefreshProperties(Optional StayOnSameLine As Boolean) As Boolean
  ' Refresh the properties grid for the selected controls in the screen designer.
  On Error GoTo ErrorTrap
  
  ' Debug.Print "RefreshProperties"
  
  
  Dim fOK As Boolean
  Dim fPassedOnce As Boolean
  Dim iLoop As Integer
  Dim iDisplayType As NavigationDisplayType
  Dim iControlType As Integer
  Dim sDescription As String
  Dim objCtlFont As StdFont
  Dim ctlControl As VB.Control
  Dim avProperties(18, 3) As Variant
  
  'JDM - Fault 63 - Cache the SelectedControlCount to speed things up a bit.
  Dim iSelectedControlCount As Integer
  
  '22/07 RH to store which row the user is on after pressing enter
  Dim CurrentRow As Integer
  
  ' Initialise the property array.
  ' NB. This array has a row for each property, and 3 columns.
  ' Column 1 - True if the selected controls all have the property.
  ' Column 2 - True if selected controls have different values for the property.
  ' Column 3 - row number in the grid for the property.
  For iLoop = 1 To UBound(avProperties)
    avProperties(iLoop, 1) = False
    avProperties(iLoop, 2) = False
    avProperties(iLoop, 3) = -1
  Next iLoop
  
  ' Ensure the global font object is instantiated.
  If gObjFont Is Nothing Then
    Set gObjFont = New StdFont
  End If

  ' Determine which properties should be displayed in the grid, and their values.
  If Not (gFrmScreen Is Nothing) Then
  
    iSelectedControlCount = gFrmScreen.SelectedControlsCount
  
    ' Process each selected control.
    If iSelectedControlCount > 0 Then
      
      ' Initialise the property array.
      For iLoop = 1 To UBound(avProperties)
        avProperties(iLoop, 1) = True
      Next iLoop
      
      fPassedOnce = False
      
      ' Get the properties of each screen control type.
      For Each ctlControl In gFrmScreen.Controls
        
        If gFrmScreen.IsScreenControl(ctlControl) Then
        
          ' Get the screen control type.
          iControlType = gFrmScreen.ScreenControl_Type(ctlControl)
          
          With ctlControl
          
            If (.Selected) Then
            
            ' RH 18/09/00 - SUG 757 - Show name of currently highlighted control in the
            '                         properties window status bar. This is because for some
            '                         reason, tooltips dont display in the exe
            ' JDM - Fault 2452 - Put up correct message in properties window
            If iSelectedControlCount = 1 Then
              Me.StatusBar1.SimpleText = "" & GetTableColumnName(.ColumnID)
              If Me.StatusBar1.SimpleText = "" Then
                Me.Caption = "<Non Field Control>"
              Else
                Me.Caption = "Control Properties"
              End If
            ElseIf iSelectedControlCount > 1 Then
              Me.StatusBar1.SimpleText = "<Multiple Controls>"
              Me.Caption = "<Multiple Controls>"
            End If
            
              ' Read the Alignment property from the control if required.
              If avProperties(giPROPID_ALIGNMENT, 1) Then
                If gFrmScreen.ScreenControl_HasAlignment(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_ALIGNMENT, 2)) Then
                    If giAlignment <> .Alignment Then
                      avProperties(giPROPID_ALIGNMENT, 2) = True
                    End If
                  Else
                    giAlignment = .Alignment
                  End If
                Else
                  avProperties(giPROPID_ALIGNMENT, 1) = False
                End If
              End If
            
              ' Read the BackColor property from the control if required.
              If avProperties(giPROPID_BACKCOLOR, 1) Then
                If gFrmScreen.ScreenControl_HasBackColor(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_BACKCOLOR, 2)) Then
                    If gColBackColor <> .BackColor Then
                      avProperties(giPROPID_BACKCOLOR, 2) = True
                    End If
                  Else
                    gColBackColor = .BackColor
                  End If
                Else
                  avProperties(giPROPID_BACKCOLOR, 1) = False
                End If
              End If
              
              ' Read the BorderStyle property from the control if required.
              If avProperties(giPROPID_BORDERSTYLE, 1) Then
                If gFrmScreen.ScreenControl_HasBorderStyle(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_BORDERSTYLE, 2)) Then
                    If giBorderStyle <> .BorderStyle Then
                      avProperties(giPROPID_BORDERSTYLE, 2) = True
                    End If
                  Else
                    giBorderStyle = .BorderStyle
                  End If
                Else
                  avProperties(giPROPID_BORDERSTYLE, 1) = False
                End If
              End If
                          
                          
              'NPG20071022 - Read the READONLY property from the control if required.
              If avProperties(giPROPID_READONLY, 1) Then
                If gFrmScreen.ScreenControl_HasReadOnly(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_READONLY, 2)) Then
                    If mblnReadOnly <> .Read_Only Then
                      avProperties(giPROPID_READONLY, 2) = True
                    End If
                  Else
                    mblnReadOnly = .Read_Only
                  End If
                Else
                  avProperties(giPROPID_READONLY, 1) = False
                End If
              End If
                          
                          
              ' Read the Caption property from the control if required.
              If avProperties(giPROPID_CAPTION, 1) Then
                If gFrmScreen.ScreenControl_HasCaption(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_CAPTION, 2)) Then
                    If gsCaption <> .Caption Then
                      avProperties(giPROPID_CAPTION, 2) = True
                    End If
                  Else
                    'gsCaption = .Caption
                    gsCaption = Replace(.Caption, "&&", "&")
                  End If
                Else
                  avProperties(giPROPID_CAPTION, 1) = False
                End If
              End If
              
              
              ' Read the NavigateTo property from the control if required.
              If avProperties(giPROPID_NAVIGATETO, 1) Then
                If gFrmScreen.ScreenControl_HasNavigation(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_NAVIGATETO, 2)) Then
                    If gsNavigateTo <> .NavigateTo Then
                      avProperties(giPROPID_NAVIGATETO, 2) = True
                    End If
                  Else
                    gsNavigateTo = .NavigateTo
                  End If
                Else
                  avProperties(giPROPID_NAVIGATETO, 1) = False
                End If
              End If
              
              
              ' Read the NavigateTo property from the control if required.
              If avProperties(giPROPID_NAVIGATEIN, 1) Then
                If gFrmScreen.ScreenControl_HasNavigation(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_NAVIGATEIN, 2)) Then
                    If giNavigateIn <> .NavigateIn Then
                      avProperties(giPROPID_NAVIGATEIN, 2) = True
                    End If
                  Else
                    giNavigateIn = .NavigateIn
                  End If
                Else
                  avProperties(giPROPID_NAVIGATEIN, 1) = False
                End If
              End If
              
              
              ' Read the DisplayType property from the control if required.
              If avProperties(giPROPID_DISPLAYTYPE, 1) Then
                If gFrmScreen.ScreenControl_HasDisplayType(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_DISPLAYTYPE, 2)) Then
                    If giDisplayType <> .DisplayType Then
                      avProperties(giPROPID_DISPLAYTYPE, 2) = True
                    End If
                  Else
                    giDisplayType = .DisplayType
                  End If
                Else
                  avProperties(giPROPID_DISPLAYTYPE, 1) = False
                End If
              End If
              
              ' Read the NavigateOnSave property
              If avProperties(giPROPID_NAVIGATEONSAVE, 1) Then
                If gFrmScreen.ScreenControl_HasNavigation(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_NAVIGATEONSAVE, 2)) Then
                    If gbNavigateOnSave <> .NavigateOnSave Then
                      avProperties(giPROPID_NAVIGATEONSAVE, 2) = True
                    End If
                  Else
                    gbNavigateOnSave = .NavigateOnSave
                  End If
                Else
                  avProperties(giPROPID_NAVIGATEONSAVE, 1) = False
                End If
              End If
              
              ' Read the Font property from the control if required.
              If avProperties(giPROPID_FONT, 1) Then
                If gFrmScreen.ScreenControl_HasFont(iControlType) Then
                  Set objCtlFont = .Font
                  
                  If (fPassedOnce) And (Not avProperties(giPROPID_FONT, 2)) Then
                    If (objCtlFont.Name <> gObjFont.Name) Or _
                      (objCtlFont.Size <> gObjFont.Size) Or _
                      (objCtlFont.Bold <> gObjFont.Bold) Or _
                      (objCtlFont.Italic <> gObjFont.Italic) Or _
                      (objCtlFont.Strikethrough <> gObjFont.Strikethrough) Or _
                      (objCtlFont.Underline <> gObjFont.Underline) Then
                      avProperties(giPROPID_FONT, 2) = True
                    End If
                  Else
                  
                    gObjFont.Name = objCtlFont.Name
                    gObjFont.Size = objCtlFont.Size
                    gObjFont.Bold = objCtlFont.Bold
                    gObjFont.Italic = objCtlFont.Italic
                    gObjFont.Strikethrough = objCtlFont.Strikethrough
                    gObjFont.Underline = objCtlFont.Underline
                  End If
                  
                  Set objCtlFont = Nothing
                Else
                  avProperties(giPROPID_FONT, 1) = False
                End If
              End If
              
              ' Read the ForeColor property from the control if required.
              If avProperties(giPROPID_FORECOLOR, 1) Then
                If gFrmScreen.ScreenControl_HasForeColor(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_FORECOLOR, 2)) Then
                    If gColForeColor <> .ForeColor Then
                      avProperties(giPROPID_FORECOLOR, 2) = True
                    End If
                  Else
                    gColForeColor = .ForeColor
                  End If
                Else
                  avProperties(giPROPID_FORECOLOR, 1) = False
                End If
              End If
                          
              ' Read the Height property from the control if required.
              If avProperties(giPROPID_HEIGHT, 1) Then
                If gFrmScreen.ScreenControl_HasHeight(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_HEIGHT, 2)) Then
                    If gLngHeight <> .Height Then
                      avProperties(giPROPID_HEIGHT, 2) = True
                    End If
                  Else
                    gLngHeight = .Height
                  End If
                Else
                  avProperties(giPROPID_HEIGHT, 1) = False
                End If
              End If
              
              ' Read the Left property from the control.
              avProperties(giPROPID_LEFT, 1) = True
              If (fPassedOnce) And (Not avProperties(giPROPID_LEFT, 2)) Then
                If gLngLeft <> .Left Then
                  avProperties(giPROPID_LEFT, 2) = True
                End If
              Else
                gLngLeft = .Left
              End If
              
              ' Read the Orientation property from the control if required.
              If avProperties(giPROPID_ORIENTATION, 1) Then
                If gFrmScreen.ScreenControl_HasOrientation(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_ORIENTATION, 2)) Then
                    If giOrientation <> .Alignment Then
                      avProperties(giPROPID_ORIENTATION, 2) = True
                    End If
                  Else
                    giOrientation = .Alignment
                  End If
                Else
                  avProperties(giPROPID_ORIENTATION, 1) = False
                End If
              End If
            
              ' Read the Picture property from the control if required.
              If avProperties(giPROPID_PICTURE, 1) Then
                If gFrmScreen.ScreenControl_HasPicture(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_PICTURE, 2)) Then
                    If glngPictureID <> .PictureID Then
                      avProperties(giPROPID_PICTURE, 2) = True
                    End If
                  Else
                    glngPictureID = .PictureID
                  End If
                Else
                  avProperties(giPROPID_PICTURE, 1) = False
                End If
              End If
              
              ' Read the Top property from the control.
              avProperties(giPROPID_TOP, 1) = True
              If (fPassedOnce) And (Not avProperties(giPROPID_TOP, 2)) Then
                If gLngTop <> .Top Then
                  avProperties(giPROPID_TOP, 2) = True
                End If
              Else
                gLngTop = .Top
              End If
              
              ' Read the Width property from the control if required.
              If avProperties(giPROPID_WIDTH, 1) Then
                If gFrmScreen.ScreenControl_HasWidth(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_WIDTH, 2)) Then
                    If gLngWidth <> .Width Then
                      avProperties(giPROPID_WIDTH, 2) = True
                    End If
                  Else
                    gLngWidth = .Width
                  End If
                Else
                  avProperties(giPROPID_WIDTH, 1) = False
                End If
              End If
              
              ' Read the Page Tab Number property from the control if required.
              If avProperties(giPROPID_TABNUMBER, 1) Then
                If gFrmScreen.ScreenControl_HasTabNumber(iControlType) Then
                  If (fPassedOnce) And (Not avProperties(giPROPID_TABNUMBER, 2)) Then
                    If glngTabNumber <> .SelectedItem.Index Then
                      avProperties(giPROPID_TABNUMBER, 2) = True
                    End If
                  Else
                    glngTabNumber = .SelectedItem.Index
                  End If
                Else
                  avProperties(giPROPID_TABNUMBER, 1) = False
                End If
              End If
              
              ' Flag that we have read a control's properties.
              fPassedOnce = True
  
            End If
          End With
        End If
      Next ctlControl
      
      ' Disassociate object variables.
      Set ctlControl = Nothing
    Else
  
      ' The screen's tab strip is the active control.
      Me.StatusBar1.SimpleText = "Tab"
      
      If gFrmScreen.tabPages.Tabs.Count > 0 Then
          
        With gFrmScreen.tabPages
          
          ' Read the Caption property from the tab strip.
          avProperties(giPROPID_CAPTION, 1) = True
          
          'MH20020527 Fault 1862
          'gsCaption = .SelectedItem.Caption
          gsCaption = Replace(.SelectedItem.Caption, "&&", "&")
          
          ' Read the Font property from the tab strip.
          avProperties(giPROPID_FONT, 1) = True
          
          gObjFont.Name = .Font.Name
          gObjFont.Size = .Font.Size
          gObjFont.Bold = .Font.Bold
          gObjFont.Italic = .Font.Italic
          gObjFont.Strikethrough = .Font.Strikethrough
          gObjFont.Underline = .Font.Underline
        
          ' Read the tab number from the tab strip
          avProperties(giPROPID_TABNUMBER, 1) = True
          avProperties(giPROPID_TABNUMBER, 2) = False
          glngTabNumber = .SelectedItem.Index
        
        End With
      End If
    End If
  End If
  
  '
  ' Add the required rows to the properties grid.
  
  CurrentRow = ssGridProperties.Row
  '
  ' Clear the properties grid.
  ssGridProperties.RemoveAll
  
  ' Add the Alignment property row to the properties grid if required.
  If avProperties(giPROPID_ALIGNMENT, 1) Then
    If avProperties(giPROPID_ALIGNMENT, 2) Then
      sDescription = ""
    Else
      If giAlignment = vbLeftJustify Then
        sDescription = gsALIGNMENTLEFTTEXT
      Else
        sDescription = gsALIGNMENTRIGHTTEXT
      End If
    End If
    
    ssGridProperties.AddItem "Alignment" & vbTab & sDescription & vbTab & Str(giPROPID_ALIGNMENT)
    avProperties(giPROPID_ALIGNMENT, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the BackColor property row to the properties grid if required.
  If avProperties(giPROPID_BACKCOLOR, 1) Then
    If avProperties(giPROPID_BACKCOLOR, 2) Then
      ssGridProperties.StyleSets("ssetBackColorValue").BackColor = vbWhite
    Else
      ssGridProperties.StyleSets("ssetBackColorValue").BackColor = gColBackColor
    End If

    ssGridProperties.AddItem "Background Colour" & vbTab & "" & vbTab & Str(giPROPID_BACKCOLOR)
    avProperties(giPROPID_BACKCOLOR, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the BorderStyle property row to the properties grid if required.
  If avProperties(giPROPID_BORDERSTYLE, 1) Then
    If avProperties(giPROPID_BORDERSTYLE, 2) Then
      sDescription = ""
    Else
      If giBorderStyle = vbBSNone Then
        sDescription = gsBORDERSTYLENONETEXT
      Else
        sDescription = gsBORDERSTYLEFIXEDSINGLETEXT
      End If
    End If
    
    ssGridProperties.AddItem "Border" & vbTab & sDescription & vbTab & Str(giPROPID_BORDERSTYLE)
    avProperties(giPROPID_BORDERSTYLE, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the Caption property row to the properties grid if required.
  If avProperties(giPROPID_CAPTION, 1) Then
    If avProperties(giPROPID_CAPTION, 2) Then
      sDescription = ""
    Else
      sDescription = gsCaption
    End If
    
    ssGridProperties.AddItem "Caption" & vbTab & sDescription & vbTab & Str(giPROPID_CAPTION)
    avProperties(giPROPID_CAPTION, 3) = ssGridProperties.Rows - 1
  End If
  
  
  
  ' Add the NavigateOnSave property row to the properties grid if required.
  If avProperties(giPROPID_NAVIGATEONSAVE, 1) Then
    If avProperties(giPROPID_NAVIGATEONSAVE, 2) Then
      sDescription = ""
    Else
      If gbNavigateOnSave = False Then
        sDescription = gsREADONLYNOTEXT
      Else
        sDescription = gsREADONLYYESTEXT
      End If
    End If
    
    ssGridProperties.AddItem "Navigate On Save" & vbTab & sDescription & vbTab & Str(giPROPID_NAVIGATEONSAVE)
    avProperties(giPROPID_NAVIGATEONSAVE, 3) = ssGridProperties.Rows - 1
  End If
  
    
  ' Add the DisplayType property row to the properties grid if required.
  If avProperties(giPROPID_DISPLAYTYPE, 1) Then
    If avProperties(giPROPID_DISPLAYTYPE, 2) Then
      sDescription = ""
    Else
      Select Case giDisplayType
        Case NavigationDisplayType.Hyperlink
          sDescription = gsDISPLAYTYPE_HYPERLINK
        Case NavigationDisplayType.Button
          sDescription = gsDISPLAYTYPE_BUTTON
        Case NavigationDisplayType.Browser
          sDescription = gsDISPLAYTYPE_BROWSER
        Case NavigationDisplayType.Hidden
          sDescription = gsDISPLAYTYPE_HIDDEN
      End Select

    End If

    ssGridProperties.AddItem "Display Type" & vbTab & sDescription & vbTab & Str(giPROPID_DISPLAYTYPE)
    avProperties(giPROPID_DISPLAYTYPE, 3) = ssGridProperties.Rows - 1
  End If
  
  
  ' Add the NavigateTo property row to the properties grid if required.
  If avProperties(giPROPID_NAVIGATETO, 1) Then
    If avProperties(giPROPID_NAVIGATETO, 2) Then
      sDescription = ""
    Else
      sDescription = gsNavigateTo
    End If
    
    ssGridProperties.AddItem "Navigate To" & vbTab & sDescription & vbTab & Str(giPROPID_NAVIGATETO)
    avProperties(giPROPID_NAVIGATETO, 3) = ssGridProperties.Rows - 1
  End If
  
  
  ' Add the NavigateTo property row to the properties grid if required.
  If avProperties(giPROPID_NAVIGATEIN, 1) Then
    If avProperties(giPROPID_NAVIGATEIN, 2) Then
      sDescription = ""
    Else
      Select Case giNavigateIn
        Case NavigateIn.URL
          sDescription = gsNAVIGATE_URL
        Case NavigateIn.MenuBar
          sDescription = gsNAVIGATE_MENUBAR
      End Select
    End If
    
    ssGridProperties.AddItem "Destination" & vbTab & sDescription & vbTab & Str(giPROPID_NAVIGATEIN)
    avProperties(giPROPID_NAVIGATEIN, 3) = ssGridProperties.Rows - 1
  End If
  
  
  ' Add the Font property row to the properties grid if required.
  If avProperties(giPROPID_FONT, 1) Then
    If avProperties(giPROPID_FONT, 2) Then
      sDescription = ""
    Else
      sDescription = GetFontDescription
    End If
    
    ssGridProperties.AddItem "Font" & vbTab & sDescription & vbTab & Str(giPROPID_FONT)
    avProperties(giPROPID_FONT, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the ForeColor property row to the properties grid if required.
  If avProperties(giPROPID_FORECOLOR, 1) Then
    If avProperties(giPROPID_FORECOLOR, 2) Then
      ssGridProperties.StyleSets("ssetForeColorValue").BackColor = vbWhite
    Else
      ssGridProperties.StyleSets("ssetForeColorValue").BackColor = gColForeColor
    End If

    ssGridProperties.AddItem "Foreground Colour" & vbTab & "" & vbTab & Str(giPROPID_FORECOLOR)
    avProperties(giPROPID_FORECOLOR, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the Height property row to the properties grid if required.
  If avProperties(giPROPID_HEIGHT, 1) Then
    If avProperties(giPROPID_HEIGHT, 2) Then
      sDescription = ""
    Else
      sDescription = Trim(Str(gLngHeight))
    End If
    
    ssGridProperties.AddItem "Height" & vbTab & sDescription & vbTab & Str(giPROPID_HEIGHT)
    avProperties(giPROPID_HEIGHT, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the Left property row to the properties grid if required.
  If avProperties(giPROPID_LEFT, 1) Then
    If avProperties(giPROPID_LEFT, 2) Then
      sDescription = ""
    Else
      sDescription = Trim(Str(gLngLeft))
    End If
    
    ssGridProperties.AddItem "Left" & vbTab & sDescription & vbTab & Str(giPROPID_LEFT)
    avProperties(giPROPID_LEFT, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the Orientation property row to the properties grid if required.
  If avProperties(giPROPID_ORIENTATION, 1) Then
    If avProperties(giPROPID_ORIENTATION, 2) Then
      sDescription = ""
    Else
      If giOrientation = 0 Then
        sDescription = gsORIENTATIONVERTICAL
      Else
        sDescription = gsORIENTATIONHORIZONTAL
      End If
    End If
    
    ssGridProperties.AddItem "Orientation" & vbTab & sDescription & vbTab & Str(giPROPID_ORIENTATION)
    avProperties(giPROPID_ORIENTATION, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the Picture property row to the properties grid if required.
  If avProperties(giPROPID_PICTURE, 1) Then
    If avProperties(giPROPID_PICTURE, 2) Then
      sDescription = ""
    Else
      If glngPictureID = 0 Then
        sDescription = gsPICTURENONETEXT
      Else
        With recPictEdit
          .Index = "idxID"
          .Seek "=", glngPictureID
          If Not .NoMatch Then
            sDescription = !Name
          Else
            sDescription = gsPICTURESELECTEDTEXT
          End If
        End With
      End If
    End If
    
    ssGridProperties.AddItem "Picture" & vbTab & sDescription & vbTab & Str(giPROPID_PICTURE)
    avProperties(giPROPID_PICTURE, 3) = ssGridProperties.Rows - 1
  End If


  'NPG20071022 -  Add the ReadOnly property row to the properties grid if required.
  If avProperties(giPROPID_READONLY, 1) Then
    If avProperties(giPROPID_READONLY, 2) Then
      sDescription = ""
    Else
      If mblnReadOnly = False Then
        sDescription = gsREADONLYNOTEXT
      Else
        sDescription = gsREADONLYYESTEXT
      End If
    End If
    
    ssGridProperties.AddItem "Read Only" & vbTab & sDescription & vbTab & Str(giPROPID_READONLY)
    avProperties(giPROPID_READONLY, 3) = ssGridProperties.Rows - 1
  End If
  
  
  
  ' Add the Top property row to the properties grid if required.
  If avProperties(giPROPID_TOP, 1) Then
    If avProperties(giPROPID_TOP, 2) Then
      sDescription = ""
    Else
      sDescription = Trim(Str(gLngTop))
    End If
    
    ssGridProperties.AddItem "Top" & vbTab & sDescription & vbTab & Str(giPROPID_TOP)
    avProperties(giPROPID_TOP, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the Width property row to the properties grid if required.
  If avProperties(giPROPID_WIDTH, 1) Then
    If avProperties(giPROPID_WIDTH, 2) Then
      sDescription = ""
    Else
      sDescription = Trim(Str(gLngWidth))
    End If
    
    ssGridProperties.AddItem "Width" & vbTab & sDescription & vbTab & Str(giPROPID_WIDTH)
    avProperties(giPROPID_WIDTH, 3) = ssGridProperties.Rows - 1
  End If
  
  ' Add the Page Number property row to the properties grid if required.
  If avProperties(giPROPID_TABNUMBER, 1) Then
    If avProperties(giPROPID_TABNUMBER, 2) Then
      sDescription = ""
    Else
      sDescription = Trim(Str(glngTabNumber))
    End If
    
    ssGridProperties.AddItem "Page Tab Order" & vbTab & sDescription & vbTab & Str(giPROPID_TABNUMBER)
    avProperties(giPROPID_TABNUMBER, 3) = ssGridProperties.Rows - 1
  End If
  
  '
  ' Set the stylesets for the grid cells.
  ' NB. This is done after all of the rows are added to the grid, as
  ' the AddItem method appears to reset any cellStyleSet settings that already existed.
  ' Set the stylesets for the grid cells.
  '
  If avProperties(giPROPID_BACKCOLOR, 1) Then
    ssGridProperties.Columns(1).CellStyleSet "ssetBackColorValue", avProperties(giPROPID_BACKCOLOR, 3)
  End If

  If avProperties(giPROPID_FORECOLOR, 1) Then
    ssGridProperties.Columns(1).CellStyleSet "ssetForeColorValue", avProperties(giPROPID_FORECOLOR, 3)
  End If
  
  fOK = True

  'Set the row number back to what it was before. 22/07
  If StayOnSameLine = True Then
    ssGridProperties.Row = CurrentRow
  End If
  CurrentRow = 0

  'JPD 20050812 Fault 10175
  ' Ensure the grid has the correct column properties applied for the selected property.
  ' Note that we to do this here even if the selected row/col has not changed, as the property
  ' displayed in the selected row may have.
  If frmSysMgr.ActiveForm Is Me Then
    ssGridProperties_RowColChange 0, 0
  End If
  
  ' Exterminate the form icon
  RemoveIcon Me
  
TidyUpAndExit:
  Set objCtlFont = Nothing
  Set ctlControl = Nothing
  RefreshProperties = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function UpdateControls(piProperty As Integer) As Boolean
  ' Update the screen controls with the given property value from the grid.
  On Error GoTo ErrorTrap
  
  ' Debug.Print "UpdateControls"
  
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim sFileName As String
  Dim objFont As StdFont
  Dim ctlControl As VB.Control
  Dim actlSelectedControls() As VB.Control
  Dim blnDontRefresh As Boolean
  Dim lngTargetPageNumber As Long
  Dim sCaption As String
  Dim iControlType As Integer
  
  Dim iCount As Integer

  fOK = Not (gFrmScreen Is Nothing)

  If fOK Then
    If gFrmScreen.SelectedControlsCount > 0 Then
      ' Construct an array of selected controls.
      ReDim actlSelectedControls(0)

      For Each ctlControl In gFrmScreen.Controls
        If gFrmScreen.IsScreenControl(ctlControl) Then
          If ctlControl.Selected Then
            iIndex = UBound(actlSelectedControls) + 1
            ReDim Preserve actlSelectedControls(iIndex)
            Set actlSelectedControls(iIndex) = ctlControl
          End If
        End If
      Next ctlControl
      ' Disassociate object variables.
      Set ctlControl = Nothing

      For iIndex = 1 To UBound(actlSelectedControls)

        Set ctlControl = actlSelectedControls(iIndex)
        With ctlControl
          iControlType = gFrmScreen.ScreenControl_Type(ctlControl)

          ' Update the control with the new property value.
          Select Case piProperty
            Case giPROPID_ALIGNMENT
              .Alignment = giAlignment

            Case giPROPID_BACKCOLOR
              .BackColor = gColBackColor

            Case giPROPID_BORDERSTYLE
              .BorderStyle = giBorderStyle


            'NPG20071022
            Case giPROPID_READONLY
              .Read_Only = mblnReadOnly


            Case giPROPID_CAPTION
              .Caption = Replace(gsCaption, "&", "&&")
              blnDontRefresh = True

              If iControlType = giCTRL_LABEL Then
                lblSizeTester.Width = .Width
                lblSizeTester.Height = .Height
                lblSizeTester.Caption = ""
                Set lblSizeTester.Font = .Font
                lblSizeTester.WordWrap = True

                lblSizeTester.AutoSize = True
                lblSizeTester.Caption = .Caption
                lblSizeTester.AutoSize = False

                .Width = lblSizeTester.Width
                .Height = lblSizeTester.Height
              End If

            Case giPROPID_DISPLAYTYPE
              .DisplayType = giDisplayType

            Case giPROPID_NAVIGATETO
              .NavigateTo = gsNavigateTo

            Case giPROPID_NAVIGATEIN
              .NavigateIn = giNavigateIn

            Case giPROPID_NAVIGATEONSAVE
              .NavigateOnSave = gbNavigateOnSave

            Case giPROPID_FONT
              Set objFont = New StdFont
              objFont.Name = gObjFont.Name
              objFont.Size = gObjFont.Size
              objFont.Bold = gObjFont.Bold
              objFont.Italic = gObjFont.Italic
              objFont.Strikethrough = gObjFont.Strikethrough
              objFont.Underline = gObjFont.Underline
              Set .Font = objFont
              Set objFont = Nothing

              If iControlType = giCTRL_LABEL Then
                lblSizeTester.Width = .Width
                lblSizeTester.Height = .Height
                lblSizeTester.Caption = ""
                Set lblSizeTester.Font = .Font
                lblSizeTester.WordWrap = True

                lblSizeTester.AutoSize = True
                lblSizeTester.Caption = .Caption
                lblSizeTester.AutoSize = False

                .Width = lblSizeTester.Width
                .Height = lblSizeTester.Height
              End If

            Case giPROPID_FORECOLOR
              .ForeColor = gColForeColor

            Case giPROPID_HEIGHT
              .Height = gLngHeight

            Case giPROPID_LEFT
              .Left = gLngLeft

            Case giPROPID_ORIENTATION
              ' AE20080221 Fault #12317
              '.Alignment = giOrientation
              If giOrientation <> .Alignment Then
                .Alignment = giOrientation
              End If

            Case giPROPID_PICTURE
              If glngPictureID > 0 Then
                recPictEdit.Index = "idxID"
                recPictEdit.Seek "=", glngPictureID

                If Not recPictEdit.NoMatch Then
                  sFileName = ReadPicture
                  .PictureID = glngPictureID
                  .Picture = sFileName
                  Kill sFileName
                End If
              End If

            Case giPROPID_TOP
              .Top = gLngTop

            Case giPROPID_WIDTH
              .Width = gLngWidth

          End Select
        End With
       
        ' Refresh the selection markers
        For iCount = 1 To gFrmScreen.ASRSelectionMarkers.Count - 1
         With gFrmScreen.ASRSelectionMarkers(iCount)
           If .Visible Then
             .Move .AttachedObject.Left - .MarkerSize, .AttachedObject.Top - .MarkerSize, .AttachedObject.Width + (.MarkerSize * 2), .AttachedObject.Height + (.MarkerSize * 2)
             .RefreshSelectionMarkers True
           End If
         End With
        Next iCount

        If Not fOK Then
          Exit For
        End If

      Next iIndex
    Else
      With gFrmScreen.tabPages
        ' The screen's tab strip is the active control.
        If .Tabs.Count > 0 Then

          ' Update the tab strip with the new property value.
          Select Case piProperty
            Case giPROPID_CAPTION

              'MH20020520 Fault 1862
              '.SelectedItem.Caption = gsCaption
              .SelectedItem.Caption = Replace(gsCaption, "&", "&&")

              ' RH 15/09/00 - Dont refresh object properties if caption of tab
              '               is changed. BUG 936
              blnDontRefresh = True

            Case giPROPID_FONT
              .Font.Name = gObjFont.Name
              .Font.Size = gObjFont.Size
              .Font.Bold = gObjFont.Bold
              .Font.Italic = gObjFont.Italic
              .Font.Strikethrough = gObjFont.Strikethrough
              .Font.Underline = gObjFont.Underline

            Case giPROPID_TABNUMBER

              lngTargetPageNumber = gFrmScreen.tabPages.Tabs(.SelectedItem.Index).Tag
              sCaption = gFrmScreen.tabPages.Tabs(.SelectedItem.Index).Caption

              ' Set the current tab details
              gFrmScreen.tabPages.Tabs(.SelectedItem.Index).Tag = gFrmScreen.tabPages.Tabs(glngTabNumber).Tag
              gFrmScreen.tabPages.Tabs(.SelectedItem.Index).Caption = gFrmScreen.tabPages.Tabs(glngTabNumber).Caption

              ' Set the original page details
              gFrmScreen.tabPages.Tabs(glngTabNumber).Caption = sCaption
              gFrmScreen.tabPages.Tabs(glngTabNumber).Tag = lngTargetPageNumber

              ' Select the new page
              gFrmScreen.PageNo = glngTabNumber
              gFrmScreen.IsChanged = True
              Me.SetFocus
              ssGridProperties.MoveLast
              blnDontRefresh = True

          End Select
        End If
      End With
    End If
  
    ' JDM - 21/08/02 - Fault 4275 - Update save button
    gFrmScreen.IsChanged = True

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

Private Function GetFontDescription() As String
  ' Return the test description of the current font for display in the grid.
  On Error GoTo ErrorTrap
  
  Dim sFontDescription As String
  
  If Not gObjFont Is Nothing Then
    With gObjFont
    
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






Private Function ValidIntegerString(psString As String) As Boolean
  ' Return true if the given string is a string.
  On Error GoTo ErrorTrap
  
  Dim fValid As Boolean
  Dim lngValueOfString As Long
  Dim sStringOfValue As String
  
  psString = Trim(psString)
  lngValueOfString = Val(psString)
  sStringOfValue = Trim(Str(lngValueOfString))
  
  fValid = (psString = sStringOfValue)
  
TidyUpAndExit:
  ValidIntegerString = fValid
  Exit Function

ErrorTrap:
  fValid = False
  Resume TidyUpAndExit

End Function
