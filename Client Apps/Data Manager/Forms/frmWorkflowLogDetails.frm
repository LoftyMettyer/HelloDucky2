VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmWorkflowLogDetails 
   Caption         =   "Workflow Log Details"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1140
   Icon            =   "frmWorkflowLogDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2650
      Left            =   10200
      TabIndex        =   8
      Top             =   200
      Width           =   1200
      Begin VB.CommandButton cmdSucceeding 
         Caption         =   "&Succeeding"
         Height          =   400
         Left            =   0
         TabIndex        =   12
         Top             =   2000
         Width           =   1200
      End
      Begin VB.CommandButton cmdPreceding 
         Caption         =   "&Preceding"
         Height          =   400
         Left            =   0
         TabIndex        =   11
         Top             =   1500
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   400
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View..."
         Height          =   400
         Left            =   0
         TabIndex        =   10
         Top             =   500
         Width           =   1200
      End
   End
   Begin VB.Frame fraFilters 
      Caption         =   "Filters :"
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10020
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   500
         Width           =   2600
      End
      Begin VB.ComboBox cboElementType 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   500
         Width           =   2200
      End
      Begin VB.ComboBox cboCaption 
         Height          =   315
         ItemData        =   "frmWorkflowLogDetails.frx":000C
         Left            =   3900
         List            =   "frmWorkflowLogDetails.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   500
         Width           =   2200
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   195
         Left            =   7200
         TabIndex        =   5
         Top             =   250
         Width           =   705
      End
      Begin VB.Label lblElementType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Element Type :"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   250
         Width           =   1305
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Caption :"
         Height          =   195
         Left            =   3900
         TabIndex        =   3
         Top             =   250
         Width           =   1230
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   13
      Top             =   4845
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   19844
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid grdWorkflowLog 
      Height          =   3525
      Left            =   120
      TabIndex        =   7
      Top             =   1185
      Width           =   10020
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   7
      stylesets.count =   2
      stylesets(0).Name=   "ssetDormant"
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
      stylesets(0).Picture=   "frmWorkflowLogDetails.frx":0010
      stylesets(1).Name=   "ssetActive"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   -2147483646
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
      stylesets(1).Picture=   "frmWorkflowLogDetails.frx":002C
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   3
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      MaxSelectedRows =   0
      StyleSet        =   "ssetDormant"
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      ActiveRowStyleSet=   "ssetActive"
      CaptionAlignment=   0
      Columns.Count   =   7
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ID"
      Columns(0).Name =   "ID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   979
      Columns(1).Caption=   "Step"
      Columns(1).Name =   "Index"
      Columns(1).Alignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2646
      Columns(2).Caption=   "Element Type"
      Columns(2).Name =   "Element Type"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4260
      Columns(3).Caption=   "Caption"
      Columns(3).Name =   "Caption"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3519
      Columns(4).Caption=   "Status"
      Columns(4).Name =   "Status"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2831
      Columns(5).Caption=   "Preceding Elements"
      Columns(5).Name =   "Preceding Elements"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2990
      Columns(6).Caption=   "Succeeding Elements"
      Columns(6).Name =   "Succeeding Elements"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   17674
      _ExtentY        =   6218
      _StockProps     =   79
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
   Begin ActiveBarLibraryCtl.ActiveBar abWorkflowLog 
      Left            =   10230
      Top             =   2805
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
      Bands           =   "frmWorkflowLogDetails.frx":0048
   End
End
Attribute VB_Name = "frmWorkflowLogDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Flag to prevent grid refreshing when combos are being populated and set initially
Private mblnLoading As Boolean

' Must be public so the details form can change the bookmark of the recordset
Public mrstInstanceSteps As Recordset
Public mrstLinks As Recordset


' Data access class
Private mclsData As New clsDataAccess

' Variables to hold the column clicked on, its field and the order to sort the grid
Private pstrOrderField As String
Private pstrOrderOrder As String
Private mintSortColumnIndex As Integer

Private mlngWorkflowInstanceID As Long

Private mfrmWorkflowLog As frmWorkflowLog

Private mavarElementIndex() As Variant
Private mlngNextIndex As Long

Private malngTempElements() As Long

Private Sub AddRowToGrid()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowDetailsLog.AddRowToGrid()"

  Dim fGoodRecord As Boolean
  Dim lngElementIndex As Long
  Dim sPrecedingElements As String
  Dim sSucceedingElements As String
  Dim sAddItemString As String
        
  fGoodRecord = mrstInstanceSteps.Fields("type") <> elem_Connector1 _
    And mrstInstanceSteps.Fields("type") <> elem_Connector2
        
  If fGoodRecord And cboElementType.ListIndex > 0 Then
    If mrstInstanceSteps.Fields("typeDesc") <> cboElementType.Text Then
      fGoodRecord = False
    End If
  End If
        
  If fGoodRecord And cboCaption.ListIndex > 0 Then
    If cboCaption.ListIndex = 1 Then
      ' Blank captions
      If Len(Trim(mrstInstanceSteps.Fields("caption"))) > 0 Then
        fGoodRecord = False
      End If
    Else
      If mrstInstanceSteps.Fields("caption") <> cboCaption.Text Then
        fGoodRecord = False
      End If
    End If
  End If
        
  If fGoodRecord And cboStatus.ListIndex > 0 Then
    If mrstInstanceSteps.Fields("status") <> cboStatus.ItemData(cboStatus.ListIndex) Then
      fGoodRecord = False
    End If
  End If

  If fGoodRecord Then
    lngElementIndex = GetElementIndex(mrstInstanceSteps.Fields("elementID"))
    sPrecedingElements = GetPrecedingElementIndexes(mrstInstanceSteps.Fields("elementID"))
    sSucceedingElements = GetSucceedingElementIndexes(mrstInstanceSteps.Fields("elementID"))

    sAddItemString = CStr(mrstInstanceSteps.Fields("stepID")) & _
      vbTab & CStr(lngElementIndex) & _
      vbTab & mrstInstanceSteps.Fields("typeDesc") & _
      vbTab & mrstInstanceSteps.Fields("caption") & _
      vbTab & mrstInstanceSteps.Fields("statusDesc") & _
      vbTab & sPrecedingElements & _
      vbTab & sSucceedingElements

    grdWorkflowLog.AddItem sAddItemString
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
End Sub

Private Function GetPrecedingElementIndexes(plngElementID As Long) As String
  Dim iLoop As Integer
  Dim sIndexes As String
  
  sIndexes = ""
  
  For iLoop = 0 To UBound(mavarElementIndex, 2)
    If Not IsEmpty(mavarElementIndex(0, iLoop)) Then
      If CLng(mavarElementIndex(0, iLoop)) = plngElementID Then
        sIndexes = CStr(mavarElementIndex(2, iLoop))
        Exit For
      End If
    End If
  Next iLoop
  
  GetPrecedingElementIndexes = sIndexes
  
End Function

Private Sub GetPrecedingElements(plngElementID As Long)
  Dim iLoop As Integer
  Dim fFound As Boolean
  Dim varBookmark As Variant
  
  With mrstLinks
    .MoveFirst
    
    Do Until .EOF
      If .Fields("endElementID") = plngElementID Then
        
        If .Fields("startElementType") = elem_Connector1 _
          Or .Fields("startElementType") = elem_Connector2 Then
        
          varBookmark = mrstLinks.Bookmark
          GetPrecedingElements .Fields("startElementID")
          .Bookmark = varBookmark
        Else
          fFound = False
          
          For iLoop = 1 To UBound(malngTempElements)
            If malngTempElements(iLoop) = .Fields("startElementID") Then
              fFound = True
              Exit For
            End If
          Next iLoop
          
          If Not fFound Then
            ReDim Preserve malngTempElements(UBound(malngTempElements) + 1)
            malngTempElements(UBound(malngTempElements)) = .Fields("startElementID")
          End If
        End If
      End If
      
      .MoveNext
    Loop
    
  End With
  
End Sub

Private Sub GetSucceedingElements(plngElementID As Long)
  Dim iLoop As Integer
  Dim fFound As Boolean
  Dim varBookmark As Variant
  
  With mrstLinks
    .MoveFirst
    
    Do Until .EOF
      If .Fields("startElementID") = plngElementID Then
        
        If .Fields("endElementType") = elem_Connector1 _
          Or .Fields("endElementType") = elem_Connector2 Then
        
          varBookmark = mrstLinks.Bookmark
          GetSucceedingElements .Fields("endElementID")
          .Bookmark = varBookmark
        Else
          fFound = False
          
          For iLoop = 1 To UBound(malngTempElements)
            If malngTempElements(iLoop) = .Fields("endElementID") Then
              fFound = True
              Exit For
            End If
          Next iLoop
          
          If Not fFound Then
            ReDim Preserve malngTempElements(UBound(malngTempElements) + 1)
            malngTempElements(UBound(malngTempElements)) = .Fields("endElementID")
          End If
        End If
      End If
      
      .MoveNext
    Loop
    
  End With
  
End Sub


Private Function GetSucceedingElementIndexes(plngElementID As Long) As String
  Dim iLoop As Integer
  Dim sIndexes As String
  
  sIndexes = ""
  
  For iLoop = 0 To UBound(mavarElementIndex, 2)
    If Not IsEmpty(mavarElementIndex(0, iLoop)) Then
      If CLng(mavarElementIndex(0, iLoop)) = plngElementID Then
        sIndexes = CStr(mavarElementIndex(3, iLoop))
        Exit For
      End If
    End If
  Next iLoop
  
  GetSucceedingElementIndexes = sIndexes
  
End Function


Private Function GetElementIndex(plngElementID As Long) As Long
  Dim iLoop As Integer
  Dim lngIndex As Long
  
  lngIndex = 0
  
  For iLoop = 0 To UBound(mavarElementIndex, 2)
    If Not IsEmpty(mavarElementIndex(0, iLoop)) Then
      If CLng(mavarElementIndex(0, iLoop)) = plngElementID Then
        lngIndex = CLng(mavarElementIndex(1, iLoop))
        Exit For
      End If
    End If
  Next iLoop
  
  GetElementIndex = lngIndex
  
End Function


Private Sub IndexElement(plngElementID As Long, piElementType As ElementType)
  ' Add the given element to the index array if it is not already there.
  Dim iLoop As Integer
  Dim fFound As Boolean
  Dim varBookmark As Variant
  
  fFound = False
  For iLoop = 0 To UBound(mavarElementIndex, 2)
    If Not IsEmpty(mavarElementIndex(0, iLoop)) Then
      If CLng(mavarElementIndex(0, iLoop)) = plngElementID Then
        fFound = True
        Exit For
      End If
    End If
  Next iLoop
  
  If Not fFound Then
    ' Add the given element to the index array as it is not already there.
    ' Column 0 = Element ID
    ' Column 1 = Index
    ' Column 2 = Preceding Element Indexes (comma delimited)
    ' Column 3 = Succeeding Element Indexes (comma delimited)
    ReDim Preserve mavarElementIndex(3, UBound(mavarElementIndex, 2) + 1)
    mavarElementIndex(0, UBound(mavarElementIndex, 2)) = plngElementID
    mavarElementIndex(2, UBound(mavarElementIndex, 2)) = ""
    mavarElementIndex(3, UBound(mavarElementIndex, 2)) = ""
    
    If piElementType = elem_Connector1 _
      Or piElementType = elem_Connector2 Then
    
      mavarElementIndex(1, UBound(mavarElementIndex, 2)) = -1
    Else
      mlngNextIndex = mlngNextIndex + 1
      mavarElementIndex(1, UBound(mavarElementIndex, 2)) = mlngNextIndex
    End If
    
    ' Add to the index array any
    mrstLinks.MoveFirst
    Do Until mrstLinks.EOF
      If mrstLinks.Fields("startElementID") = plngElementID Then
        varBookmark = mrstLinks.Bookmark
      
        ' Add to the index array the elements that follow this one.
        IndexElement mrstLinks.Fields("endElementID"), mrstLinks.Fields("endElementType")
        
        mrstLinks.Bookmark = varBookmark
      End If
      
      mrstLinks.MoveNext
    Loop
  End If
  
End Sub

Private Sub IndexPrecedingSucceedingElements(plngElementID As Long)
  Dim sPrecedingElementIndexes As String
  Dim sSucceedingElementIndexes As String
  Dim iLoop As Integer
  
  sPrecedingElementIndexes = ""
  ReDim malngTempElements(0)
  GetPrecedingElements plngElementID
  
  ' Replace the element IDs with the index values.
  For iLoop = 1 To UBound(malngTempElements)
    malngTempElements(iLoop) = GetElementIndex(malngTempElements(iLoop))
  Next iLoop
  
  ' Sort the array into index order.
  ShellSortArray malngTempElements
  
  ' Create the string
  For iLoop = 1 To UBound(malngTempElements)
    sPrecedingElementIndexes = sPrecedingElementIndexes & _
      IIf(Len(sPrecedingElementIndexes) > 0, ", ", "") & _
      CStr(malngTempElements(iLoop))
  Next iLoop
  
  sSucceedingElementIndexes = ""
  ReDim malngTempElements(0)
  GetSucceedingElements plngElementID
  
  ' Replace the element IDs with the index values.
  For iLoop = 1 To UBound(malngTempElements)
    malngTempElements(iLoop) = GetElementIndex(malngTempElements(iLoop))
  Next iLoop
  
  ' Sort the array into index order.
  ShellSortArray malngTempElements

  ' Create the string
  For iLoop = 1 To UBound(malngTempElements)
    sSucceedingElementIndexes = sSucceedingElementIndexes & _
      IIf(Len(sSucceedingElementIndexes) > 0, ", ", "") & _
      CStr(malngTempElements(iLoop))
  Next iLoop

  ' Store the string of preceding/succeeding elements in the array.
  For iLoop = 0 To UBound(mavarElementIndex, 2)
    If Not IsEmpty(mavarElementIndex(0, iLoop)) Then
      If CLng(mavarElementIndex(0, iLoop)) = plngElementID Then
        mavarElementIndex(2, iLoop) = sPrecedingElementIndexes
        mavarElementIndex(3, iLoop) = sSucceedingElementIndexes
        
        Exit For
      End If
    End If
  Next iLoop
  
End Sub

Private Sub ShellSortArray(vArray As Variant)
  Dim lLoop1 As Long
  Dim lHold As Long
  Dim lHValue As Long
  Dim varTemp As Variant

  lHValue = LBound(vArray)
  
  Do
    lHValue = 3 * lHValue + 1
  Loop Until lHValue > UBound(vArray)
  
  Do
    lHValue = lHValue / 3
    
    For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
      varTemp = vArray(lLoop1)
      lHold = lLoop1
      
      Do While (vArray(lHold - lHValue) > varTemp)
        vArray(lHold) = vArray(lHold - lHValue)
        lHold = lHold - lHValue
        If lHold < lHValue Then Exit Do
      Loop
      
      vArray(lHold) = varTemp
    Next lLoop1
  Loop Until lHValue = LBound(vArray)
  
End Sub


Private Sub ShellSortArrayByColumn(vArray As Variant, piColumn As Integer)
  Dim lLoop1 As Long
  Dim lLoop2 As Long
  Dim lHold As Long
  Dim lHValue As Long
  Dim avarTemp() As Variant

  ReDim avarTemp(UBound(vArray, 1))

  lHValue = LBound(vArray, 2)
  
  Do
    lHValue = 3 * lHValue + 1
  Loop Until lHValue > UBound(vArray, 2)
  
  Do
    lHValue = lHValue / 3
    
    For lLoop1 = lHValue + LBound(vArray, 2) To UBound(vArray, 2)
      For lLoop2 = 0 To UBound(vArray, 1)
        avarTemp(lLoop2) = vArray(lLoop2, lLoop1)
      Next lLoop2
      
      lHold = lLoop1
      
      Do While (vArray(piColumn, lHold - lHValue) > avarTemp(piColumn))
        For lLoop2 = 0 To UBound(vArray, 1)
          vArray(lLoop2, lHold) = vArray(lLoop2, lHold - lHValue)
        Next lLoop2
        
        lHold = lHold - lHValue
        If lHold < lHValue Then Exit Do
      Loop

      For lLoop2 = 0 To UBound(vArray, 1)
        vArray(lLoop2, lHold) = avarTemp(lLoop2)
      Next lLoop2
    Next lLoop1
  Loop Until lHValue = LBound(vArray, 2)
  
End Sub



Public Function Initialise(plngKey As Long, pfrmWorkflowLog As frmWorkflowLog) As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim rstElementTypes As Recordset
  Dim rstCaptions As Recordset
  Dim sSQL As String
  Dim iLoop As Integer
  
  fOK = True

  mblnLoading = True

  ' Let user know we are doing something, and dont redraw the form until the controls
  ' have all been repositioned
  Screen.MousePointer = vbHourglass
  Me.AutoRedraw = False

  mlngWorkflowInstanceID = plngKey

  Set mfrmWorkflowLog = pfrmWorkflowLog

  'Add all available element types to the Element combo
  sSQL = "SELECT DISTINCT" & _
    "  CASE" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 0 THEN 'Begin'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 1 THEN 'Terminator'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 2 THEN 'Web Form'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 3 THEN 'Email'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 4 THEN 'Decision'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 5 THEN 'Stored Data'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 6 THEN 'And'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 7 THEN 'Or'" & _
    "  END AS [typeDesc]" & _
    " FROM ASRSysWorkflowElements" & _
    " INNER JOIN ASRSysWorkflowInstances ON ASRSysWorkflowElements.workflowID = ASRSysWorkflowInstances.workflowID" & _
    " WHERE ASRSysWorkflowInstances.ID = " & CStr(mlngWorkflowInstanceID) & _
    "  AND [ASRSysWorkflowElements].[Type] <= 7 " & _
    " ORDER BY [typeDesc]"

  cboElementType.Clear
  cboElementType.AddItem "<All>"
  Set rstElementTypes = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Do Until rstElementTypes.EOF
    cboElementType.AddItem rstElementTypes.Fields("typeDesc")
    rstElementTypes.MoveNext
  Loop
  rstElementTypes.Close
  Set rstElementTypes = Nothing

  cboElementType.ListIndex = 0

  'Add all available captions to the Caption combo
  sSQL = "SELECT DISTINCT [ASRSysWorkflowElements].[Caption]" & _
    " FROM ASRSysWorkflowElements" & _
    " INNER JOIN ASRSysWorkflowInstances ON ASRSysWorkflowElements.workflowID = ASRSysWorkflowInstances.workflowID" & _
    " WHERE ASRSysWorkflowInstances.ID = " & CStr(mlngWorkflowInstanceID) & _
    "  AND [ASRSysWorkflowElements].[Type] <= 7" & _
    "  AND len(ltrim(rtrim(isnull([ASRSysWorkflowElements].[Caption], '')))) > 0" & _
    " ORDER BY [Caption]"
  
  With cboCaption
    .Clear
    .AddItem "<All>"
    .AddItem "<Blank>"

    Set rstCaptions = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Do Until rstCaptions.EOF
      .AddItem rstCaptions.Fields("caption")
      rstCaptions.MoveNext
    Loop
    Set rstCaptions = Nothing
    
    .ListIndex = 0
  End With

  'Add all available statuses to the Status combo
  With cboStatus
    .Clear
    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_ALL)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_ALL

    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_COMPLETED)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_COMPLETED

    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_FAILED)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_FAILED

    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_FAILEDACTION)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_FAILEDACTION

    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_INPROGRESS)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_INPROGRESS

    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_ONHOLD)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_ONHOLD

    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_PENDINGENGINEACTION)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_PENDINGENGINEACTION

    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_PENDINGUSERACTION)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_PENDINGUSERACTION

    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_PENDINGUSERCOMPLETION)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_PENDINGUSERCOMPLETION

    .AddItem WorkflowStepStatusDescription(giWFSTEPSTATUS_TIMEOUT)
    .ItemData(.NewIndex) = giWFSTEPSTATUS_TIMEOUT

    .ListIndex = 0
  End With

  ' Let user know we have finished, and can now redraw the form
  Screen.MousePointer = vbDefault
  Me.AutoRedraw = True

  mblnLoading = False
  Initialise = fOK

  'Set default sort order to be date desc
  pstrOrderField = grdWorkflowLog.Columns(1).Caption
  mintSortColumnIndex = 1
  pstrOrderOrder = "ASC"
  
  'Read the workflow link details. Use this to create an index of the elements.
  sSQL = "SELECT ASRSysWorkflowLinks.startElementID," & _
    "  ASRSysWorkflowLinks.endElementID," & _
    "  startEl.type AS [startElementType]," & _
    "  endEl.type AS [endElementType]" & _
    " FROM ASRSysWorkflowLinks" & _
    " INNER JOIN ASRSysWorkflowInstances ON ASRSysWorkflowLinks.workflowID = ASRSysWorkflowInstances.workflowID" & _
    "   AND ASRSysWorkflowInstances.ID = " & CStr(mlngWorkflowInstanceID) & _
    " INNER JOIN ASRSysWorkflowElements startEl ON ASRSysWorkflowLinks.startElementID = startEl.ID" & _
    " INNER JOIN ASRSysWorkflowElements endEl ON ASRSysWorkflowLinks.endElementID = endEl.ID" & _
    " UNION" & _
    " SELECT startEl.ID AS [startElementID]," & _
    "   startEl.connectionPairID AS [endElementID]," & _
    "   startEl.type AS [startElementType]," & _
    "   endEl.type AS [endElementType]" & _
    " FROM ASRSysWorkflowElements startEl" & _
    " INNER JOIN ASRSysWorkflowInstances ON startEl.workflowID = ASRSysWorkflowInstances.workflowID" & _
    "   AND ASRSysWorkflowInstances.ID = " & CStr(mlngWorkflowInstanceID) & _
    " INNER JOIN ASRSysWorkflowElements endEl ON startEl.connectionPairID = endEl.ID" & _
    " WHERE startEl.Type = 8" & _
    " ORDER BY [startElementType]"
  
  Set mrstLinks = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  ReDim mavarElementIndex(3, 0)
  mlngNextIndex = 0
  If Not (mrstLinks.EOF And mrstLinks.BOF) Then
    ' Add to the index array the Begin element (all others will follow)
    IndexElement mrstLinks.Fields("startElementID"), mrstLinks.Fields("startElementType")
  End If
  
  For iLoop = 0 To UBound(mavarElementIndex, 2)
    If Not IsEmpty(mavarElementIndex(0, iLoop)) Then
      IndexPrecedingSucceedingElements CLng(mavarElementIndex(0, iLoop))
    End If
  Next iLoop
  
  RefreshLog

TidyUpAndExit:
  Exit Function

ErrorTrap:
  Initialise = False
  COAMsgBox "Error retrieving detail entries for this workflow." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Workflow Log"
  GoTo TidyUpAndExit

End Function


Private Function RefreshGrid() As Boolean
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowDetailsLog.RefreshGrid()"

  ' Populate the grid using filter/sort criteria as set by the user
  Dim strStatusBarText As String
  Dim sOrderString As String
  Dim fOrderByArray As Boolean
  Dim iLoop As Integer
  Dim lngElementID As Long
  Dim lngStartPosition As Long
  Dim lngEndPosition As Long
  Dim lngStep As Long
  
  If mblnLoading = True Then GoTo TidyUpAndExit

  Screen.MousePointer = vbHourglass

  grdWorkflowLog.Redraw = False
  grdWorkflowLog.RemoveAll

  If Not mrstInstanceSteps Is Nothing Then

    fOrderByArray = False
    sOrderString = ""
    Select Case pstrOrderField
      Case "Caption"
        sOrderString = "caption"
      Case "Element Type"
        sOrderString = "typeDesc"
      Case "Status"
        sOrderString = "statusDesc"
      Case Else
        fOrderByArray = True
    End Select
    
    If fOrderByArray Then
      Select Case pstrOrderField
        Case "Step"
          ShellSortArrayByColumn mavarElementIndex, 1
        Case "Preceding Elements"
          ShellSortArrayByColumn mavarElementIndex, 2
        Case "Succeeding Elements"
          ShellSortArrayByColumn mavarElementIndex, 3
      End Select
      
      If pstrOrderOrder = "ASC" Then
        lngStartPosition = 0
        lngEndPosition = UBound(mavarElementIndex, 2)
        lngStep = 1
      Else
        lngStartPosition = UBound(mavarElementIndex, 2)
        lngEndPosition = 0
        lngStep = -1
      End If
      
      For iLoop = lngStartPosition To lngEndPosition Step lngStep
        If (Not IsEmpty(mavarElementIndex(0, iLoop))) _
          And (Not (mrstInstanceSteps.BOF And mrstInstanceSteps.EOF)) Then
          lngElementID = mavarElementIndex(0, iLoop)
          mrstInstanceSteps.MoveFirst
          mrstInstanceSteps.Find "[elementID] = " & CStr(lngElementID)
    
          If Not mrstInstanceSteps.EOF Then
            AddRowToGrid
          End If
        End If
      Next iLoop
    Else
      If Len(sOrderString) > 0 Then
        sOrderString = sOrderString & " " & pstrOrderOrder
      End If
      mrstInstanceSteps.Sort = sOrderString

      mrstInstanceSteps.MoveFirst
      
      Do Until mrstInstanceSteps.EOF
        AddRowToGrid
        
        mrstInstanceSteps.MoveNext
      Loop
    End If
  End If
    
  With grdWorkflowLog
    .Redraw = True

    If .Rows > 0 Then
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With

  strStatusBarText = vbNullString
  strStatusBarText = strStatusBarText & " " & grdWorkflowLog.Rows & " step" & IIf(grdWorkflowLog.Rows > 1 Or grdWorkflowLog.Rows = 0, "s", "")
  If grdWorkflowLog.Rows > 1 Then
    strStatusBarText = strStatusBarText & " sorted by "
    strStatusBarText = strStatusBarText & pstrOrderField & " "
    strStatusBarText = strStatusBarText & "in "
    strStatusBarText = strStatusBarText & IIf(pstrOrderOrder = "ASC", "ascending", "descending")
    strStatusBarText = strStatusBarText & " order"
  End If

  StatusBar1.SimpleText = strStatusBarText

  RefreshButtons

  DoColumnSizes

  Screen.MousePointer = vbDefault

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

Private Sub RefreshLog()
  Dim sSQL As String
  Dim varBookmark As Variant
  
  'Read the workflow element, instance and link details into arrays.
  sSQL = "SELECT ASRSysWorkflowElements.ID AS [elementID]," & _
    "  ASRSysWorkflowElements.type," & _
    "  isnull(ASRSysWorkflowElements.caption, '') AS [caption]," & _
    "  ASRSysWorkflowInstanceSteps.ID AS [stepID]," & _
    "  ASRSysWorkflowInstanceSteps.status," & _
    "  CASE" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 0 THEN 'Begin'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 1 THEN 'Terminator'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 2 THEN 'Web Form'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 3 THEN 'Email'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 4 THEN 'Decision'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 5 THEN 'Stored Data'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 6 THEN 'And'" & _
    "    WHEN [ASRSysWorkflowElements].[Type] = 7 THEN 'Or'" & _
    "  END AS [typeDesc],"
    
  sSQL = sSQL & _
    "  CASE" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_ONHOLD) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_ONHOLD) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_PENDINGENGINEACTION) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_PENDINGENGINEACTION) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_PENDINGUSERACTION) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_PENDINGUSERACTION) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_COMPLETED) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_COMPLETED) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_FAILED) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_FAILED) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_FAILEDACTION) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_FAILEDACTION) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_INPROGRESS) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_INPROGRESS) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_TIMEOUT) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_TIMEOUT) & "'" & _
    "    WHEN [ASRSysWorkflowInstanceSteps].[status] = " & CStr(giWFSTEPSTATUS_PENDINGUSERCOMPLETION) & " THEN '" & WorkflowStepStatusDescription(giWFSTEPSTATUS_PENDINGUSERCOMPLETION) & "'" & _
    "  END AS [statusDesc]" & _
    " FROM ASRSysWorkflowElements" & _
    " INNER JOIN ASRSysWorkflowInstanceSteps ON ASRSysWorkflowElements.ID = ASRSysWorkflowInstanceSteps.elementID" & _
    " WHERE ASRSysWorkflowInstanceSteps.instanceID = " & CStr(mlngWorkflowInstanceID) & _
    " ORDER BY ASRSysWorkflowElements.type"
    
  If Not mrstInstanceSteps Is Nothing Then
    If mrstInstanceSteps.State <> adStateClosed Then
      mrstInstanceSteps.Close
    End If
    Set mrstInstanceSteps = Nothing
  End If
  Set mrstInstanceSteps = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
  If grdWorkflowLog.SelBookmarks.Count = 1 Then
    varBookmark = grdWorkflowLog.SelBookmarks.Item(0)
  End If
  
  'Populate the grid
  RefreshGrid
  
  If Not IsEmpty(varBookmark) Then
    grdWorkflowLog.SelBookmarks.RemoveAll
    grdWorkflowLog.Bookmark = varBookmark
    grdWorkflowLog.SelBookmarks.Add varBookmark
  End If
  
End Sub

Private Function ViewWorkflowStep()
  
  Dim frmDetails As frmWorkflowStepDetails

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.ViewWorkflowStep()"

  If IsNumeric(grdWorkflowLog.Columns("ID").Value) Then
    Set frmDetails = New frmWorkflowStepDetails
    If frmDetails.Initialise(grdWorkflowLog.Columns("ID").Value, Me, grdWorkflowLog.Columns("Element Type").Value) Then
      frmDetails.Show vbModal
      
      RefreshLog
    Else
      RefreshLog
      Unload frmDetails
    End If
  End If

TidyUpAndExit:
  Set frmDetails = Nothing
  gobjErrorStack.PopStack
  Exit Function

ErrorTrap:
  gobjErrorStack.HandleError
  
End Function



Private Function HighlightPrecedingWorkflowSteps()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.HighlightPrecedingWorkflowSteps()"

  Dim iLoop As Integer
  Dim sAllPrecedingIndexes As String
  Dim sSubPrecedingIndexes As String
  Dim varBookmark As Variant
  Dim varFirstBookmark As Variant
  Dim sTemp As String
  
  sAllPrecedingIndexes = ", "
  
  With grdWorkflowLog
    ' Construct a comma-delimted string of the element indexes that precede the selected elements
    For iLoop = 0 To .SelBookmarks.Count - 1
      sSubPrecedingIndexes = .Columns("Preceding Elements").CellText(.SelBookmarks(iLoop))
    
      If Len(sSubPrecedingIndexes) > 0 Then
        sAllPrecedingIndexes = sAllPrecedingIndexes & _
          sSubPrecedingIndexes & _
          ", "
      End If
    Next iLoop
    sAllPrecedingIndexes = Replace(sAllPrecedingIndexes, " ", "")
  
    If Len(sAllPrecedingIndexes) > 1 Then
      ' Clear the selected rows, and select any rows for elements that precede the previously selected elements
      .Redraw = False
      .SelBookmarks.RemoveAll
      .MoveFirst
    
      For iLoop = 0 To .Rows - 1
        varBookmark = .Bookmark
          
        sTemp = .Columns("Index").CellText(varBookmark)
        If InStr(sAllPrecedingIndexes, "," & sTemp & ",") > 0 Then
          .SelBookmarks.Add varBookmark
        End If
        
        .MoveNext
      Next iLoop
      
      ' If there are no preceding elements, just select the top row.
      If .SelBookmarks.Count = 0 Then
        .MoveFirst
        .SelBookmarks.Add .AddItemBookmark(0)
      Else
        .Bookmark = .SelBookmarks(0)
      End If
      
      .Redraw = True
    End If
  End With
  
  RefreshButtons
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function

ErrorTrap:
  gobjErrorStack.HandleError
  
End Function


Private Function HighlightSucceedingWorkflowSteps()
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.HighlightSucceedingWorkflowSteps()"

  Dim iLoop As Integer
  Dim sAllSucceedingIndexes As String
  Dim sSubSucceedingIndexes As String
  Dim varBookmark As Variant
  Dim varFirstBookmark As Variant
  Dim sTemp As String
  
  sAllSucceedingIndexes = ", "
  
  With grdWorkflowLog
    ' Construct a comma-delimted string of the element indexes that succeed the selected elements
    For iLoop = 0 To .SelBookmarks.Count - 1
      sSubSucceedingIndexes = .Columns("Succeeding Elements").CellText(.SelBookmarks(iLoop))
    
      If Len(sSubSucceedingIndexes) > 0 Then
        sAllSucceedingIndexes = sAllSucceedingIndexes & _
          sSubSucceedingIndexes & _
          ", "
      Else
        sAllSucceedingIndexes = sAllSucceedingIndexes & _
          .Columns("Index").CellText(.SelBookmarks(iLoop)) & _
          ", "
      End If
    Next iLoop
    sAllSucceedingIndexes = Replace(sAllSucceedingIndexes, " ", "")
  
    If Len(sAllSucceedingIndexes) > 1 Then
      ' Clear the selected rows, and select any rows for elements that succeed the previously selected elements
      .Redraw = False
      .SelBookmarks.RemoveAll
      .MoveFirst
    
      For iLoop = 0 To .Rows - 1
        varBookmark = .Bookmark
          
        sTemp = .Columns("Index").CellText(varBookmark)
        If InStr(sAllSucceedingIndexes, "," & sTemp & ",") > 0 Then
          .SelBookmarks.Add varBookmark
        End If
        
        .MoveNext
      Next iLoop
      
      .Bookmark = .SelBookmarks(0)
      
      .Redraw = True
    End If
  End With
  
  RefreshButtons
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function

ErrorTrap:
  gobjErrorStack.HandleError
  
End Function


Private Sub DoColumnSizes()
  On Error GoTo ErrorTrap

  Dim lngAvailableWidth As Long

  gobjErrorStack.PushStack "frmWorkflowLogDetails.DoColumnSizes()"

  With grdWorkflowLog
    lngAvailableWidth = grdWorkflowLog.Width - (270 + .Columns(1).Width + .Columns(2).Width + .Columns(4).Width)

    .Columns(3).Width = (lngAvailableWidth * 0.4)
    .Columns(5).Width = (lngAvailableWidth * 0.3)
    .Columns(6).Width = (lngAvailableWidth * 0.3)
  End With

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub



Private Sub RefreshButtons()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.RefreshButtons()"

  cmdView.Enabled = (grdWorkflowLog.SelBookmarks.Count = 1)
  cmdPreceding.Enabled = (grdWorkflowLog.SelBookmarks.Count >= 1)
  cmdSucceeding.Enabled = (grdWorkflowLog.SelBookmarks.Count >= 1)

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub




Private Sub abWorkflowLog_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.abWorkflowLog_Click(Tool)", Array(Tool)

  Select Case Tool.Name
    Case "View"
      ViewWorkflowStep

    Case "Preceding Elements"
      HighlightPrecedingWorkflowSteps

    Case "Succeeding Elements"
      HighlightSucceedingWorkflowSteps

  End Select

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub cboCaption_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.cboCaption_Click()"

  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub cboElementType_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.cboElementType_Click()"

  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub cboStatus_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.cboStatus_Click()"

  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub cmdOK_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.cmdOK_Click()"

  Unload Me

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub cmdPreceding_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.cmdPreceding_Click()"

  HighlightPrecedingWorkflowSteps

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
End Sub

Private Sub cmdSucceeding_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.cmdSucceeding_Click()"

  HighlightSucceedingWorkflowSteps

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
End Sub


Private Sub cmdView_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.cmdView_Click()"

  ViewWorkflowStep

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub Form_Activate()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.Form_Activate()"

  DoColumnSizes

  UI.RemoveClipping

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = vbKeyEscape
    Unload Me
  Case KeyCode = vbKeyF5
    RefreshLog
End Select

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub Form_Load()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowDetailsLog.Form_Load()"

  Hook Me.hWnd, 12500, 5550
  
  Set mclsData = New clsDataAccess

  fraButtons.BackColor = Me.BackColor
  
  'Set height and width to last saved. Form is centred on screen
  Me.Height = GetPCSetting("WorkflowLogDetails", "Height", Me.Height)
  Me.Width = GetPCSetting("WorkflowLogDetails", "Width", Me.Width)

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode)

  ' Save the window size ready to recall next time user views the workflow log
  SavePCSetting "WorkflowLogDetails", "Height", Me.Height
  SavePCSetting "WorkflowLogDetails", "Width", Me.Width
  UI.RemoveClipping

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub Form_Resize()
  
  Const lngGap As Long = 120
  Const COMBO_GAP As Integer = 170
  Dim lngComboWidth As Long
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.Form_Resize()"

  'JPD 20030908 Fault 5756
  DisplayApplication

  ' Ensure form does not get too small/big. Also reposition controls as necessary
'  UI.ClipForForm Me, 5550, 11700
'  If Me.Width < 11900 Then Me.Width = 11900
'  If Me.Width > Screen.Width Then Me.Width = (Screen.Width - 200)
'  If Me.Height < 5550 Then Me.Height = 5550
'  If Me.Height > Screen.Height Then Me.Height = (Screen.Height - 500)

  fraButtons.Left = Me.ScaleWidth - (fraButtons.Width + lngGap)

  fraFilters.Width = fraButtons.Left - (lngGap * 2)
  
  lngComboWidth = (fraFilters.Width - (COMBO_GAP * 4)) / 3
  
  cboElementType.Move COMBO_GAP, 500, lngComboWidth
  lblElementType.Left = cboElementType.Left
  
  cboCaption.Move cboElementType.Left + cboElementType.Width + COMBO_GAP, 500, lngComboWidth
  lblCaption.Left = cboCaption.Left
  
  cboStatus.Move cboCaption.Left + cboCaption.Width + COMBO_GAP, 500, lngComboWidth
  lblStatus.Left = cboStatus.Left
  
  
  grdWorkflowLog.Width = fraFilters.Width
  grdWorkflowLog.Height = Me.ScaleHeight - (fraFilters.Height + StatusBar1.Height + (lngGap * 3))

  DoColumnSizes

  Me.Refresh
  grdWorkflowLog.Refresh

  ' Get rid of the icon off the form
  RemoveIcon Me

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

Private Sub grdWorkflowLog_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowDetailsLog.grdWorkflowLog_Click()"

  RefreshButtons
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Sub grdWorkflowLog_DblClick()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.grdWorkflowLog_DblClick()"

  If cmdView.Enabled Then
    ViewWorkflowStep
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub grdWorkflowLog_HeadClick(ByVal ColIndex As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.grdWorkflowLog_HeadClick(ColIndex)", Array(ColIndex)

  ' Set the sort criteria depending on the column header clicked and refresh the grid
  pstrOrderField = grdWorkflowLog.Columns(ColIndex).Caption

  If ColIndex = mintSortColumnIndex Then
    If pstrOrderOrder = "ASC" Then pstrOrderOrder = "DESC" Else pstrOrderOrder = "ASC"
  End If

  mintSortColumnIndex = ColIndex

  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub grdWorkflowLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.grdWorkflowLog_MouseUp(Button,Shift,X,Y)", Array(Button, Shift, X, Y)

 If (Button = vbRightButton) And (Y > Me.grdWorkflowLog.RowHeight) Then
    ' Enable/disable the required tools.
    With Me.abWorkflowLog.Bands("bndWorkflowLog")
      .Tools("View").Enabled = Me.cmdView.Enabled
      .Tools("Preceding Elements").Enabled = cmdPreceding.Enabled
      .Tools("Succeeding Elements").Enabled = cmdSucceeding.Enabled
      .TrackPopup -1, -1
    End With
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub grdWorkflowLog_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLogDetails.grdWorkflowLog_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol)

  RefreshButtons

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub



