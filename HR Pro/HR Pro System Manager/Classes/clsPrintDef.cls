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

Private miTabsOnPage As Integer

Private mintMarginTop As Integer
Private mintMarginBottom As Integer
Private mintMarginLeft As Integer
Private mintMarginRight As Integer
Public Property Let PageNumber(plngNewValue As Long)
  mlngPageNum = plngNewValue
  
End Property

Public Function PrintStart(ByVal pbForceOptions As Boolean) As Boolean

  If pbForceOptions Then
    
    Screen.MousePointer = vbDefault
    
    With frmPrintOptions
      .PrintDefinition
      .Show vbModal
      If .Cancelled Then
        PrintStart = False
      Else
        Printer.Copies = .PrintCopies
        Printer.Orientation = IIf(.PrintPortrait, vbPRORPortrait, vbPRORLandscape)
        PrintStart = True
        
        mintMarginTop = .MarginTop
        mintMarginBottom = .MarginBottom
        mintMarginLeft = .MarginLeft
        mintMarginRight = .MarginRight
      End If
    End With
  
    'Tidy up
    UnLoad frmPrintOptions
    Set frmPrintOptions = Nothing
  Else
    PrintStart = True
    Printer.Copies = 1
    Printer.Orientation = vbPRORPortrait
  End If

End Function
Public Property Get IsOK() As Boolean
  
  IsOK = mblnIsOK
  
End Property


Public Sub PrintHeader(strTitle As String)

  Dim fOK As Boolean

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

Public Property Get MarginBottom() As Integer
  MarginBottom = mintMarginBottom
End Property

Public Property Get MarginLeft() As Integer
  MarginLeft = mintMarginLeft
End Property

Public Property Get MarginTop() As Integer
  MarginTop = mintMarginTop
End Property

Public Property Get MarginRight() As Integer
  MarginRight = mintMarginRight
End Property


Public Sub PrintConfirm(pstrMessage As String, pstrHeading As String)

'Dim strMBText As String
'Dim msb As frmMessageBox
'Dim iShowMeAgain As Integer
'
'If gbPrinterConfirm And Not mblnInBatchMode Then
'
'  strMBText = pstrMessage & " printing complete." _
'    & vbCrLf & vbCrLf & "(" & Printer.DeviceName & ")"
'
'  iShowMeAgain = IIf(gbPrinterConfirm, 1, 0)
'  If iShowMeAgain = 1 Then
'    Set msb = New frmMessageBox
'    msb.MessageBox strMBText, vbInformation, pstrHeading, iShowMeAgain, "Don't show me this confirmation again."
'    gbPrinterConfirm = IIf(iShowMeAgain = 1, True, False)
'    SavePCSetting "Printer", "Confirm", gbPrinterConfirm
'    Set msb = Nothing
'  End If
'
'End If

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

  Set objDefPrinter = New cSetDfltPrinter
  Do
    objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
  Loop While Printer.DeviceName <> gstrDefaultPrinterName
  Set objDefPrinter = Nothing

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
      
'        'Jump to the centre of the page
'        'If gone past centre then print a space
'        If .CurrentX <= lngCentre Then
'          .CurrentX = lngCentre
'        Else
'          Call CheckEndOfLine(lngLeft, lngRight)
'        End If
        
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

  Printer.FontSize = 8
  Printer.Print " "
  Printer.FontBold = False
  Printer.FontUnderline = False
  Printer.FontStrikethru = False
  
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
  miTabsOnPage = 2
  
  Exit Sub
  
Init_ERROR:
  
  mblnIsOK = False
  
  Select Case Err.Number
    Case 482: pstrError = "Printer Error : Please check your printer connection."
    Case Else: pstrError = Err.Description
  End Select
  
  MsgBox pstrError, vbExclamation + vbOKOnly, Application.Name
  
End Sub

Public Property Let TabsOnPage(ByVal piNewValue As Integer)
  miTabsOnPage = piNewValue
End Property
