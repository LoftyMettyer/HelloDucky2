VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'PUBLIC SUBROUTINE...

'PrintStart   : Prompts the user for print options

'PrintHeader  : Used once at the top of the page

'PrintTitle   : Can be used for sub titles anywhere on the page
'               (NOTE: Standard is to use one of these for each tab on definition)

'PrintNormal  : Ordinary text
'               (NOTE: Everything before colun will be bold,
'                      everything after non bold)

'PrintBold    : All bold text

'PrintNonBold : All non bold text

'PrintEnd     : Will print a footer on the last page and finish the print job


Option Explicit

Private mlngPageNum As Long
Private mlngBottom As Long
Private mblnIsOK As Boolean

Public Enum PrintFontFormat
  pffNormal = 0
  pffBold = 1
  pffNonBold = 2
End Enum

Private miTabsOnPage As Integer

Private mstrNormalFontName As String
Private miNormalFontSize As Integer
Private mlngNormalFontColour As Long
Private mstrSelectedPrinter As String

Public Property Get IsOK() As Boolean
  
  IsOK = mblnIsOK
  
End Property


Public Sub PrintHeader(strTitle As String)

  'Dim fOK As Boolean

  On Error GoTo LocalErr

  With Printer
    .Font = "Verdana"
    
    ' Print the header.
    .FontSize = 12
    .FontBold = True
    .FontUnderline = True
    
    .CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(strTitle) / 2)
    .CurrentY = giPRINT_YINDENT
    Printer.Print strTitle & vbCrLf

    'This line will work out the Y difference before the title was printed and
    'after the title was printed.  It will be times by 3 so that a blank space
    'is left just above the footer text
    mlngBottom = .ScaleHeight - ((.CurrentY - giPRINT_YINDENT) * 3)
    
    
    .FontSize = 10
    .FontBold = False
    .FontUnderline = False

  End With

Exit Sub

LocalErr:
  Stop

End Sub


Public Sub PrintTitle(strText As String)

  Printer.Print
  Printer.FontSize = 12
  Printer.FontUnderline = True
  PrintText strText, pffBold
  Printer.FontUnderline = False
  Printer.FontSize = 10
  Printer.Print

End Sub


Public Sub PrintNormal(Optional strText As String)
  PrintText strText
End Sub

Public Sub PrintBold(Optional strText As String)
  PrintText strText, pffBold
End Sub

Public Sub PrintNonBold(Optional strText As String)
  PrintText strText, pffNonBold
End Sub


Public Sub PrintEnd()

  With Printer
    Do While .CurrentY < mlngBottom
      Printer.Print vbCr
    Loop
    Call FooterText
    .EndDoc
  End With

  Dim objDefPrinter As cSetDfltPrinter

  If gstrDefaultPrinterName <> vbNullString Then
    Set objDefPrinter = New cSetDfltPrinter
    Do
      objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
    Loop While Printer.DeviceName <> gstrDefaultPrinterName
    Set objDefPrinter = Nothing
  End If

End Sub

Public Sub PrintNewPage()
  
  With Printer
    Do While .CurrentY < mlngBottom
      Printer.Print vbCr
    Loop
    Call FooterText
    Printer.NewPage
  End With
  
End Sub


Private Sub PrintText(strText As String, Optional pffStyle As PrintFontFormat = pffNormal)

  Dim strChar As String
  Dim intCount As Integer
  Dim lngLeft As Long
  Dim lngCentre As Long
  Dim lngRight As Long
  
  Dim lngTabSize As Long
  Dim lngNextTabLocation As Long
  Dim lngTabsTaken As Long
  
  With Printer

    Call CheckEndOfPage
    
    ' Define print format
    .Font = mstrNormalFontName
    .FontSize = miNormalFontSize
    .ForeColor = mlngNormalFontColour
        
    'Set up margins
    lngLeft = giPRINT_XINDENT
    lngCentre = .ScaleWidth / 2
    lngRight = .ScaleWidth - (giPRINT_XINDENT * 1.5)

    lngTabSize = .ScaleWidth / miTabsOnPage
    lngNextTabLocation = lngTabSize
    lngTabsTaken = 0

    .FontBold = (pffStyle <> pffNonBold)
    .CurrentX = lngLeft

    intCount = 1
    Do While intCount <= Len(strText)
      strChar = Mid$(strText, intCount, 1)

      Select Case strChar
      Case " "
        'Check if gone past right margin
        Call CheckEndOfLine(lngLeft, lngRight)

      Case vbTab
      
        ' Move to next tab location
        If lngTabsTaken <= miTabsOnPage Then
          .FontBold = (pffStyle = pffNormal) Or (pffStyle = pffBold)
          .CurrentX = lngNextTabLocation
          lngNextTabLocation = lngNextTabLocation + lngTabSize
          lngTabsTaken = lngTabsTaken + 1
        Else
          Call CheckEndOfLine(lngLeft, lngRight)
        End If
             
      Case ":"
        'Only bold after ':' if force bold
        Printer.Print strChar;
        .FontBold = (pffStyle = pffBold)

      Case Else
        Printer.Print strChar;
      
      End Select

      'Make sure that we are not less than left margin !
      If .CurrentX < lngLeft Then
        .CurrentX = lngLeft
      End If
      
      Call CheckEndOfPage
      intCount = intCount + 1
    Loop

    Printer.Print
    Call CheckEndOfPage
    
  End With

End Sub


Private Sub FooterText()
  
  Dim strPageNum As String
  
  mlngPageNum = mlngPageNum + 1
  strPageNum = "Page " & CStr(mlngPageNum)

  Printer.Font = "Verdana"
  Printer.FontSize = 8
  Printer.Print " "
  Printer.FontBold = False
  
  Printer.CurrentX = giPRINT_XINDENT
  Printer.Print "Printed on " & Format(Now, DateFormat) & _
                " at " & Format(Now, "hh:nn") & " by " & gsUserName;
  
  Printer.CurrentX = (Printer.ScaleWidth - giPRINT_XINDENT) - Printer.TextWidth(strPageNum)
  Printer.Print strPageNum

  Printer.FontSize = 10

End Sub


Private Sub CheckEndOfLine(lngLeft As Long, lngRight As Long)
  'Carriage Return if line too long
  If Printer.CurrentX > lngRight Then
    Printer.Print " "
    Printer.CurrentX = lngLeft
  Else
    Printer.Print " ";
  End If
End Sub


Private Sub CheckEndOfPage()
  If Printer.CurrentY > mlngBottom Then
    Call FooterText
    Printer.NewPage
    Printer.CurrentY = giPRINT_YINDENT
    Printer.CurrentX = giPRINT_XINDENT
  End If
End Sub

Private Sub Class_Initialize()

  Dim pstrError As String
  
  On Error GoTo Init_ERROR
  
  If Printer.CurrentX = 0 Then
  End If
  
  mblnIsOK = True
  'NHRD17072003 Fault 6155
  miTabsOnPage = 2
  ResetFontToDefault
  
  Exit Sub
  
Init_ERROR:
  
  mblnIsOK = False
  
  Select Case Err.Number
    Case 482: pstrError = "Printer Error : Please check your printer connection."
    Case Else: pstrError = Err.Description
  End Select
  
  If Not gblnBatchMode Then MsgBox pstrError, vbExclamation + vbOKOnly, "HR Pro"
  
End Sub

Public Function PrintStart(ByVal pbForceOptions As Boolean) As Boolean

  Dim frmOptions As frmPrintOptions   'MH20040309 Fault 8143

  On Local Error GoTo LocalErr

  Set frmOptions = New frmPrintOptions

  If gbPrinterPrompt Or pbForceOptions Then
    
    Screen.MousePointer = vbDefault
    
    With frmOptions
      .PrintDefinition
      .Show vbModal
      If .Cancelled Then
        PrintStart = False
      Else
        Printer.Copies = .PrintCopies
        Printer.Orientation = IIf(.PrintPortrait, vbPRORPortrait, vbPRORLandscape)
        'MH20050114 Fault 9545
        'Last check to make sure printer is okay
        Printer.Print vbNullString
        PrintStart = True
      End If
    End With
  
    'Tidy up
    Unload frmOptions
  Else
    PrintStart = True
    Printer.Copies = 1
    Printer.Orientation = vbPRORPortrait
  End If

  mstrSelectedPrinter = Printer.DeviceName
  Set frmOptions = Nothing

Exit Function

LocalErr:
  mblnIsOK = False
  Set frmOptions = Nothing
  MsgBox Err.Description, vbCritical

End Function

Public Sub PrintConfirm(pstrMessage As String, pstrHeading As String)
  
  Dim strMBText As String
  Dim msb As frmMessageBox
  Dim iShowMeAgain As Integer
  
  If gbPrinterConfirm And Not gblnBatchMode Then
    
    'MH20040127 Fault 7941
    'strMBText = pstrMessage & " printing complete." _
      & vbCrLf & vbCrLf & "(" & Printer.DeviceName & ")"
    strMBText = pstrMessage & " printing complete." _
      & vbCrLf & vbCrLf & "(" & mstrSelectedPrinter & ")"
  
    iShowMeAgain = IIf(gbPrinterConfirm, 1, 0)
    If iShowMeAgain = 1 Then
      Set msb = New frmMessageBox
      'TM20020829 Fault 4293 - the checkbox should not be checked.
      iShowMeAgain = 0
      msb.MessageBox strMBText, vbInformation, App.ProductName, iShowMeAgain, "Don't show me this confirmation again."
      gbPrinterConfirm = IIf(iShowMeAgain = 0, True, False)
      SavePCSetting "Printer", "Confirm", gbPrinterConfirm
      Set msb = Nothing
    End If
  
  End If

End Sub
Public Function PrintStart_AbsenceCalendar() As Boolean

  Dim frmOptions As frmPrintOptions   'MH20040309 Fault 8143

  Set frmOptions = New frmPrintOptions

  If gbPrinterPrompt Then
    
    Screen.MousePointer = vbDefault
    
    With frmOptions
      
      .chkCollate.Value = vbUnchecked
      .chkHeadingsEveryPage.Value = vbUnchecked
      .chkGridlines.Value = vbUnchecked
      
      .chkCollate.Enabled = False
      .chkHeadingsEveryPage.Enabled = False
      .chkGridlines.Enabled = False
      
      .txtMargin(0).Enabled = False
      .txtMargin(1).Enabled = False
      .txtMargin(2).Enabled = False
      .txtMargin(3).Enabled = False

      .Show vbModal
      If .Cancelled Then
        PrintStart_AbsenceCalendar = False
      Else
        Printer.Copies = .PrintCopies
        Printer.Orientation = IIf(.PrintPortrait, vbPRORPortrait, vbPRORLandscape)
        Printer.ColorMode = IIf(.chkShading.Value = 1, vbPRCMColor, vbPRCMMonochrome)
        PrintStart_AbsenceCalendar = True
      End If
    End With
  
    'Tidy up
    Unload frmOptions
  Else
    PrintStart_AbsenceCalendar = True
    Printer.Copies = 1
    Printer.Orientation = vbPRORLandscape
    Printer.ColorMode = vbPRCMColor
  End If

  mstrSelectedPrinter = Printer.DeviceName
  Set frmOptions = Nothing

End Function

Public Property Let TabsOnPage(ByVal piNewValue As Integer)
  miTabsOnPage = piNewValue
End Property

Public Property Let Font(ByVal pstrFontName As String)
  mstrNormalFontName = pstrFontName
End Property

Public Property Let FontSize(ByVal piFontSize As Integer)
  miNormalFontSize = piFontSize
End Property

Public Property Let FontColour(ByVal plngFontColour As Long)
  mlngNormalFontColour = plngFontColour
End Property

Public Sub ResetFontToDefault()

  ' Default print style
  mstrNormalFontName = "Verdana"
  miNormalFontSize = 10
  mlngNormalFontColour = 0

End Sub
