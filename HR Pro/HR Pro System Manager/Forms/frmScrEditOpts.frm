VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{051CE3FC-5250-4486-9533-4E0723733DFA}#1.0#0"; "COA_ColourPicker.ocx"
Begin VB.Form frmScrEditOpts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Edit Options"
   ClientHeight    =   3705
   ClientLeft      =   1605
   ClientTop       =   2505
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
   HelpContextID   =   5029
   Icon            =   "frmScrEditOpts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   Begin COAColourPicker.COA_ColourPicker ColorPicker 
      Left            =   120
      Top             =   3120
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin TabDlg.SSTab ssTabScreenEditOptions 
      Height          =   3000
      Left            =   50
      TabIndex        =   6
      Top             =   50
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   5292
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Default &Font"
      TabPicture(0)   =   "frmScrEditOpts.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDefaultFontPage"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Screen &Grid"
      TabPicture(1)   =   "frmScrEditOpts.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraScreenGridPage"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraDefaultFontPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   2600
         Left            =   -74950
         TabIndex        =   11
         Top             =   325
         Width           =   3300
         Begin VB.TextBox txtForeColor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   200
            TabIndex        =   17
            Text            =   "Foreground colour"
            Top             =   1000
            Width           =   2500
         End
         Begin VB.CommandButton cmdForeColor 
            Caption         =   "..."
            Height          =   315
            Left            =   2715
            TabIndex        =   3
            Top             =   1000
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.CheckBox chkItalic 
            Caption         =   "&Italic"
            Height          =   315
            Left            =   2000
            TabIndex        =   2
            Top             =   600
            Width           =   750
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "&Bold"
            Height          =   315
            Left            =   200
            TabIndex        =   1
            Top             =   600
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "..."
            Height          =   315
            Left            =   2715
            TabIndex        =   0
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.Frame fraFontSample 
            Caption         =   "Sample :"
            Height          =   1000
            Left            =   200
            TabIndex        =   15
            Top             =   1400
            Width           =   2825
            Begin VB.Label lblFontSample 
               BackStyle       =   0  'Transparent
               Caption         =   "AaBbCcXxYyZz"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   500
               Left            =   200
               TabIndex        =   16
               Top             =   300
               Width           =   2400
            End
         End
         Begin VB.TextBox txtFont 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   200
            TabIndex        =   14
            Top             =   200
            Width           =   2500
         End
         Begin MSComDlg.CommonDialog comDlgBox 
            Left            =   1410
            Top             =   555
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
      End
      Begin VB.Frame fraScreenGridPage 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2600
         Left            =   50
         TabIndex        =   10
         Top             =   325
         Width           =   3300
         Begin VB.CheckBox chkAlignToGrid 
            Caption         =   "&Align Controls to Grid"
            Height          =   315
            Left            =   200
            TabIndex        =   9
            Top             =   1200
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.TextBox txtHeight 
            Height          =   315
            Left            =   1125
            MaxLength       =   4
            TabIndex        =   8
            Text            =   "Text"
            Top             =   700
            Width           =   495
         End
         Begin VB.TextBox txtWidth 
            Height          =   315
            Left            =   1125
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "Text"
            Top             =   200
            Width           =   495
         End
         Begin VB.Label lblWidth 
            BackStyle       =   0  'Transparent
            Caption         =   "Width :"
            Height          =   195
            Left            =   195
            TabIndex        =   13
            Top             =   255
            Width           =   750
         End
         Begin VB.Label lblHeight 
            BackStyle       =   0  'Transparent
            Caption         =   "Height :"
            Height          =   195
            Left            =   195
            TabIndex        =   12
            Top             =   765
            Width           =   795
         End
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   975
      TabIndex        =   4
      Top             =   3200
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2250
      TabIndex        =   5
      Top             =   3200
      Width           =   1200
   End
End
Attribute VB_Name = "frmScrEditOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private globals.
Private gFrmScreen As frmScrDesigner2
Private gObjFont As Object
Private gForeColour As ColorConstants
Private mblnOK

Private Sub chkBold_Click()
  ' Update the default font object.
  gObjFont.Bold = (chkBold.value = 1)
  RefreshFontPage
  
End Sub

Private Sub chkItalic_Click()
  ' Update the default font object.
  gObjFont.Italic = (chkItalic.value = 1)
  RefreshFontPage
  
End Sub

Private Sub cmdCancel_Click()
  ' Unload the form without saving the changes.
  UnLoad Me
  
End Sub


Private Sub cmdFont_Click()
  ' Call the font dialogue box.
  ' Trap the error caused when the dialogue box is cancelled.
  On Error GoTo ErrorTrap
  
  With comDlgBox
  
    ' Set the font properties of the dialogue box.
    .Flags = cdlCFScreenFonts Or cdlCFLimitSize
    .Max = 24
    .FontName = gObjFont.Name
    .FontSize = gObjFont.Size
    .FontBold = gObjFont.Bold
    .FontItalic = gObjFont.Italic
    .Color = gForeColour
  
    ' Display the dialogue box.
    .ShowFont
  
    ' Read the font properties of the dialogue box.
    gObjFont.Name = .FontName
    gObjFont.Size = .FontSize
    gObjFont.Bold = .FontBold
    gObjFont.Italic = .FontItalic
    gForeColour = .Color
    
  End With
      
  ' Refresh the font page.
  RefreshFontPage

ErrorTrap:
  ' User pressed cancel.
  
End Sub


Private Sub cmdForeColor_Click()
  ' Call the colour dialogue box.
  ' Trap the error caused when the dialogue box is cancelled.
  On Error GoTo ErrorTrap
  
'  With comDlgBox
'
'    ' Set the colour properties of the dialogue box.
'    .Flags = cdlCCRGBInit
'    .Color = gForeColour
'
'    ' Display the dialogue box.
'    .ShowColor
'
'    ' Read the colour properties of the dialogue box.
'    gForeColour = .Color
'
'  End With
  
  ' AE20080331 Fault #4604, #10170
  With ColorPicker
    ' Set the colour properties of the dialogue box.
    .Color = gForeColour
    ' Display the dialogue box.
    .ShowPalette
    ' Read the colour properties of the dialogue box.
    gForeColour = .Color
  End With
  
  ' Refresh the font page.
  RefreshFontPage

ErrorTrap:
  ' User pressed cancel.

End Sub


Private Sub cmdOk_Click()
  ' Save changes and unload the form.
  mblnOK = True
  
  SaveChanges
  
  If mblnOK Then UnLoad Me
  
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

  Set gObjFont = New StdFont

  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
  
  ' Call the sub-routine to initialise the screen controls.
  Call initialiseControls

  ' Ensure the frames on each of the tab pages have the same
  ' background colour as the tab pages themselves.
  fraDefaultFontPage.BackColor = ssTabScreenEditOptions.BackColor
  fraScreenGridPage.BackColor = ssTabScreenEditOptions.BackColor
  
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr
  
End Sub

Private Sub initialiseControls()
  Dim objScreenFont As StdFont
  
  Set objScreenFont = gFrmScreen.DefaultFont
  
  ' Initialise the default font controls.
  gObjFont.Name = objScreenFont.Name
  gObjFont.Bold = objScreenFont.Bold
  gObjFont.Italic = objScreenFont.Italic
  gObjFont.Size = objScreenFont.Size
  
  Set objScreenFont = Nothing
  
  gForeColour = gFrmScreen.DefaultForeColour
  
  ssTabScreenEditOptions.Tab = 0
  RefreshFontPage
   
  ' Initialise the screen grid controls.
  txtWidth.Text = gFrmScreen.GridX
  txtHeight.Text = gFrmScreen.GridY
  chkAlignToGrid.value = IIf(gFrmScreen.AlignToGrid, vbChecked, vbUnchecked)
  

End Sub


Public Property Get CurrentScreen() As frmScrDesigner2
  ' Return the current screen associated with this pop-up.
  Set CurrentScreen = gFrmScreen
End Property

Public Property Set CurrentScreen(pScrForm As frmScrDesigner2)
  Set gFrmScreen = pScrForm
End Property




Private Function SaveChanges()
Dim ErrorString As String
  ' Update the screen designer with the edit option changes.
  'On Error GoTo ErrorTrap
  ' Update the screen manager's grid size properties.
  If (val(txtWidth.Text) > 9999) Or (val(txtHeight.Text) > 9999) Then
    'ErrorString = "The Width or Height you have specified is too large." & vbCrLf & "A figure lower than 5000 is more practical for the screen designer."
    ErrorString = "Value too high."
    MsgBox ErrorString, vbExclamation + vbOKOnly, "HR Pro"
    mblnOK = False
    Exit Function
  Else
    gFrmScreen.GridX = val(txtWidth.Text)
    gFrmScreen.GridY = val(txtHeight.Text)
  End If
  
  ' Update the screen manager's align to grid property.
  gFrmScreen.AlignToGrid = IIf(chkAlignToGrid.value = vbChecked, True, False)

  ' Update the screen designer's default font properties.
  With gFrmScreen.DefaultFont
    .Name = gObjFont.Name
    .Bold = gObjFont.Bold
    .Italic = gObjFont.Italic
    .Size = gObjFont.Size
  End With
  gFrmScreen.DefaultForeColour = gForeColour

  ' Mark the screen as being changed so that the
  ' new edit options are saved.
  gFrmScreen.IsChanged = True
  
End Function

Private Sub fraDefaulFontPage_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Disassociate object variables.
  Set gFrmScreen = Nothing
  Set gObjFont = Nothing

End Sub

Private Sub ssTabScreenEditOptions_Click(PreviousTab As Integer)
  Dim fDefaultFontPage As Boolean
  Dim fScreenGridPage As Boolean
  
  ' Determine which tab is selected.
  fDefaultFontPage = (ssTabScreenEditOptions.Tab = 0)
  fScreenGridPage = (ssTabScreenEditOptions.Tab = 1)
  
  ' Enable, and make visible the selected tab.
  fraDefaultFontPage.Enabled = fDefaultFontPage
  fraDefaultFontPage.Visible = fDefaultFontPage
  
  fraScreenGridPage.Enabled = fScreenGridPage
  fraScreenGridPage.Visible = fScreenGridPage

End Sub


Private Sub RefreshFontPage()
  ' Refresh the font controls with the current font options.
  If Not gObjFont Is Nothing Then
  
    txtFont.Text = gObjFont.Name & ", " & CInt(gObjFont.Size)
    chkBold.value = IIf(gObjFont.Bold, 1, 0)
    chkItalic.value = IIf(gObjFont.Italic, 1, 0)
    txtForeColor.BackColor = gForeColour
    txtForeColor.ForeColor = UI.GetInverseColor(gForeColour)
    lblFontSample.ForeColor = gForeColour
    lblFontSample.Font.Name = gObjFont.Name
    lblFontSample.Font.Bold = gObjFont.Bold
    lblFontSample.Font.Italic = gObjFont.Italic
    lblFontSample.Font.Size = gObjFont.Size
    lblFontSample.Left = 200
    lblFontSample.Top = 300
    
  End If
  
End Sub

Private Sub txtHeight_GotFocus()
  
  ' Select the whole string.
  UI.txtSelText

End Sub


Private Sub txtWidth_GotFocus()
  
  ' Select the whole string.
  UI.txtSelText

End Sub



