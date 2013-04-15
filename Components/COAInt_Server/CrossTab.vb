Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("CrossTab_NET.CrossTab")> Public Class CrossTab


  Private mstrSQLSelect As String
  Private mstrSQLFrom As String
  Private mstrSQLJoin As String
  Private mstrSQLWhere As String

  Private mrsCrossTabColumns As ADODB.Recordset
  Private mrsMergeData As ADODB.Recordset
  Private mlngCrossTabID As Integer
  Private mblnBatchMode As Boolean
  Private mblnNoRecords As Boolean

  Private fOK As Boolean
  Private mstrStatusMessage As String
  Private mblnUserCancelled As Boolean

  Private mlngCrossTabType As Declarations.CrossTabType
  Private mstrTempTableName As String
  Private mcmdUpdateRecDescs As ADODB.Command

  Private mstrBaseTable As String
  Private mlngBaseTableID As Integer
  Private rsCrossTabData As ADODB.Recordset
  Private mblnIntersection As Boolean
  Private mblnPageBreak As Boolean
  Private mblnShowAllPagesTogether As Boolean
  Private mstrReportStartDate As String
  Private mstrReportEndDate As String

  Private mstrCrossTabName As String
  'Private mlngIntersectionType As Long
  Private mblnShowPercentage As Boolean
  Private mblnPercentageofPage As Boolean
  Private mbUse1000Separator As Boolean
  Private mblnSuppressZeros As Boolean
  Private mlngRecordDescExprID As Integer
  Private mstrPicklistFilter As String
  Private mstrPicklistFilterName As String
  Private mblnChkPicklistFilter As Boolean

  'Private mintDefaultOutput As Integer
  'Private mintDefaultExportTo As Integer
  'Private mblnDefaultSave As Boolean
  'Private mstrDefaultSaveAs As String
  'Private mblnDefaultCloseApp As Boolean
  Private mblnOutputPreview As Boolean
  Private mlngOutputFormat As Integer
  Private mblnOutputScreen As Boolean
  Private mblnOutputPrinter As Boolean
  Private mstrOutputPrinterName As String
  Private mblnOutputSave As Boolean
  Private mlngOutputSaveExisting As Integer
  Private mblnOutputEmail As Boolean
  Private mlngOutputEmailID As Integer
  'Private mstrOutputEmailAddr As String
  Private mstrOutputEmailName As String
  Private mstrOutputEmailSubject As String
  Private mstrOutputEmailAttachAs As String
  Private mstrOutputFilename As String
  Private mstrOutputPivotArray() As String

  Private mlngTableViews() As Integer
  Private mstrWhereIDs As String

  Private Const HOR As Short = 0 'Horizontal
  Private Const VER As Short = 1 'Vertical
  Private Const PGB As Short = 2 'Page Break
  Private Const INS As Short = 3 'Intersection

  Private Const TYPECOUNT As Short = 0
  Private Const TYPEAVERAGE As Short = 1
  Private Const TYPEMAXIMUM As Short = 2
  Private Const TYPEMINIMUM As Short = 3
  Private Const TYPETOTAL As Short = 4

  Private mvarHeadings(2) As Object
  Private mvarSearches(2) As Object

  Private mdblHorTotal(,,) As Double
  Private mdblVerTotal(,,) As Double
  Private mdblPgbTotal(,,) As Double
  Private mdblPageTotal(,) As Double
  Private mdblGrandTotal() As Double

  Private mdblDataArray(,,,) As Double
  Private mstrOutput() As String

  Private mlngIntersectionDecimals As Integer
  Private mstrIntersectionMask As String
  Private mdblPercentageFactor As Double
  'Private mblnInvalidPicklistFilter As Boolean

  Private mstrType() As String 'e.g. mstrtype(TYPETOTAL) = for example: "Total"
  Private mlngColID(3) As Integer
  Private mstrColName(3) As String 'e.g. mstrColName(INS) = "Salary" (the name of the intersection field)
  Private mlngColDataType(3) As String
  Private mstrColCase(3) As String
  Private mstrFormat(3) As String
  Private mdblMin(2) As Double
  Private mdblMax(2) As Double
  Private mdblStep(2) As Double

  ' Classes
  Private mclsData As clsDataAccess
  Private mobjEventLog As clsEventLog

  Private mlngType As Integer

  Private msAbsenceBreakdownTypes As String

  Private mstrOutputArray_Data() As Object
  Private mvarPrompts(,) As Object
  Private mstrClientDateFormat As String
  Private mstrLocalDecimalSeparator As String

  ' Array holding the User Defined functions that are needed for this report
  Private mastrUDFsRequired() As String

  Public WriteOnly Property Connection() As Object
    Set(ByVal Value As Object)

      If ASRDEVELOPMENT Then
        gADOCon = New ADODB.Connection
        'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gADOCon.Open(Value)
      Else
        gADOCon = Value
      End If

    End Set
  End Property

  Public WriteOnly Property CustomReportID() As Integer
    Set(ByVal Value As Integer)
      mlngCrossTabID = Value
    End Set
  End Property

  Public WriteOnly Property FailedMessage() As String
    Set(ByVal value As String)
      mobjEventLog.AddDetailEntry(value)
    End Set
  End Property

  Public WriteOnly Property ClientDateFormat() As String
    Set(ByVal value As String)
      mstrClientDateFormat = value
    End Set
  End Property

  Public WriteOnly Property LocalDecimalSeparator() As String
    Set(ByVal value As String)
      mstrLocalDecimalSeparator = value
    End Set
  End Property

  Public ReadOnly Property NoRecords() As Boolean
    Get
      NoRecords = mblnNoRecords
    End Get
  End Property

  Public WriteOnly Property CrossTabID() As Integer
    Set(ByVal Value As Integer)
      mlngCrossTabID = Value
    End Set
  End Property


  Public WriteOnly Property Username() As String
    Set(ByVal Value As String)
      gsUsername = Value
    End Set
  End Property

  Public ReadOnly Property ErrorString() As String
    Get
      ErrorString = mstrStatusMessage
    End Get
  End Property

  Public ReadOnly Property UserCancelled() As Boolean
    Get
      UserCancelled = mblnUserCancelled
    End Get
  End Property

  Public Property EventLogID() As Integer
    Get
      EventLogID = mobjEventLog.EventLogID
    End Get
    Set(ByVal value As Integer)
      mobjEventLog.EventLogID = value
    End Set
  End Property

  Public Property IntersectionType() As Integer
    Get
      IntersectionType = mlngType
    End Get
    Set(ByVal value As Integer)
      mlngType = value
    End Set
  End Property

  Public Property ShowPercentage() As Boolean
    Get
      ShowPercentage = mblnShowPercentage
    End Get
    Set(ByVal value As Boolean)
      mblnShowPercentage = value
    End Set
  End Property

  Public Property PercentageOfPage() As Boolean
    Get
      PercentageOfPage = mblnPercentageofPage
    End Get
    Set(ByVal value As Boolean)
      mblnPercentageofPage = value
    End Set
  End Property

  Public Property SuppressZeros() As Boolean
    Get
      SuppressZeros = mblnSuppressZeros
    End Get
    Set(ByVal value As Boolean)
      mblnSuppressZeros = value
    End Set
  End Property

  Public ReadOnly Property OutputArrayData(ByVal lngIndex As Integer) As String
    Get
      'UPGRADE_WARNING: Couldn't resolve default property of object lngIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      OutputArrayData = mstrOutput(lngIndex)
    End Get
  End Property

  Public ReadOnly Property OutputArrayDataUBound() As String
    Get
      OutputArrayDataUBound = CStr(UBound(mstrOutput))
    End Get
  End Property

  Public ReadOnly Property CrossTabName() As String
    Get

      Dim strOutput As String

      strOutput = mstrCrossTabName

      If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown Then
        strOutput = strOutput & " (" & ConvertSQLDateToLocale(mstrReportStartDate) & " -> " & ConvertSQLDateToLocale(mstrReportEndDate) & ")"
      End If

      If mblnChkPicklistFilter Then
        strOutput = strOutput & mstrPicklistFilterName
      End If

      CrossTabName = strOutput

    End Get
  End Property

  Public ReadOnly Property ColumnHeading(ByVal lngArray As Integer, ByVal lngIndex As Integer) As String
    Get
      'ColumnHeading = Replace(mvarHeadings(lngArray)(CLng(lngIndex)), Chr(34), Chr(34) & " + String.fromCharCode(34) + " & Chr(34))
      'UPGRADE_WARNING: Couldn't resolve default property of object lngIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      ColumnHeading = mvarHeadings(lngArray)(lngIndex)
    End Get
  End Property

  Public ReadOnly Property ColumnDataType(ByVal lngIndex As Integer) As Integer
    Get
      'UPGRADE_WARNING: Couldn't resolve default property of object lngIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      ColumnDataType = CInt(mlngColDataType(lngIndex))
    End Get
  End Property

  Public ReadOnly Property ColumnHeadingUbound(ByVal lngIndex As Integer) As Integer
    Get
      If Not mvarHeadings(lngIndex) Is Nothing Then
        ColumnHeadingUbound = UBound(mvarHeadings(lngIndex))
      Else
        ColumnHeadingUbound = 0
      End If
    End Get
  End Property

  Public ReadOnly Property PageBreakColumn() As Boolean
    Get
      PageBreakColumn = mblnPageBreak
    End Get
  End Property

  Public ReadOnly Property PageBreakColumnName() As String
    Get
      PageBreakColumnName = IIf(mblnPageBreak, Replace(mstrColName(PGB), "_", " "), "<None>")
    End Get
  End Property

  Public ReadOnly Property IntersectionColumn() As Boolean
    Get
      IntersectionColumn = mblnIntersection
    End Get
  End Property

  Public ReadOnly Property IntersectionColumnName() As String
    Get
      IntersectionColumnName = IIf(mblnIntersection, Replace(mstrColName(INS), "_", " "), "<None>")
    End Get
  End Property

  Public ReadOnly Property HorizontalColumnName() As String
    Get

      If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown Then
        HorizontalColumnName = "Day"
      Else
        HorizontalColumnName = Replace(mstrColName(HOR), "_", " ")
      End If

    End Get
  End Property

  Public ReadOnly Property VerticalColumnName() As String
    Get
      VerticalColumnName = Replace(mstrColName(VER), "_", " ")
    End Get
  End Property

  Public ReadOnly Property BaseTableName() As String
    Get
      BaseTableName = mstrBaseTable
    End Get
  End Property

  Public ReadOnly Property RecordDescExprID() As Integer
    Get
      RecordDescExprID = mlngRecordDescExprID
    End Get
  End Property

  ' What type of cross tab are we running as
  Public ReadOnly Property CrossTabType() As Integer
    Get
      CrossTabType = mlngCrossTabType
    End Get
  End Property


  Public ReadOnly Property OutputPreview() As Boolean
    Get
      OutputPreview = mblnOutputPreview
    End Get
  End Property

  Public ReadOnly Property OutputFormat() As Integer
    Get
      OutputFormat = mlngOutputFormat
    End Get
  End Property

  Public ReadOnly Property OutputScreen() As Boolean
    Get
      OutputScreen = mblnOutputScreen
    End Get
  End Property

  Public ReadOnly Property OutputPrinter() As Boolean
    Get
      OutputPrinter = mblnOutputPrinter
    End Get
  End Property

  Public ReadOnly Property OutputPrinterName() As String
    Get
      OutputPrinterName = mstrOutputPrinterName
    End Get
  End Property

  Public ReadOnly Property OutputSave() As Boolean
    Get
      OutputSave = mblnOutputSave
    End Get
  End Property

  Public ReadOnly Property OutputSaveExisting() As Integer
    Get
      OutputSaveExisting = mlngOutputSaveExisting
    End Get
  End Property

  Public ReadOnly Property OutputEmail() As Boolean
    Get
      OutputEmail = mblnOutputEmail
    End Get
  End Property

  Public ReadOnly Property OutputEmailID() As Integer
    Get
      OutputEmailID = mlngOutputEmailID
    End Get
  End Property

  Public ReadOnly Property OutputEmailGroupName() As String
    Get
      OutputEmailGroupName = mstrOutputEmailName
    End Get
  End Property

  Public ReadOnly Property OutputEmailSubject() As String
    Get
      OutputEmailSubject = mstrOutputEmailSubject
    End Get
  End Property

  Public ReadOnly Property OutputEmailAttachAs() As String
    Get
      OutputEmailAttachAs = mstrOutputEmailAttachAs
    End Get
  End Property

  Public ReadOnly Property OutputFilename() As String
    Get
      OutputFilename = mstrOutputFilename
    End Get
  End Property

  Public ReadOnly Property IntersectionDecimals() As Integer
    Get
      IntersectionDecimals = mlngIntersectionDecimals
    End Get
  End Property


  Public ReadOnly Property Heading(ByVal lngCol As Integer, ByVal lngRow As Integer) As String
    Get
      'Heading = Replace(mvarHeadings(lngCol)(lngRow), Chr(34), Chr(34) & " + String.fromCharCode(34) + " & Chr(34))
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      Heading = mvarHeadings(lngCol)(lngRow)
    End Get
  End Property

  Public ReadOnly Property DataArrayCols() As Integer
    Get
      DataArrayCols = UBound(Split(mstrOutput(0), vbTab))
    End Get
  End Property

  Public ReadOnly Property DataArrayRows() As Integer
    Get
      DataArrayRows = UBound(mstrOutput)
    End Get
  End Property

  Public ReadOnly Property DataArray(ByVal lngCol As Integer, ByVal lngRow As Integer) As String
    Get
      DataArray = Split(mstrOutput(lngRow), vbTab)(lngCol)
    End Get
  End Property


  Public Property Use1000Separator() As Boolean
    Get
      Use1000Separator = mbUse1000Separator
    End Get
    Set(ByVal value As Boolean)
      mbUse1000Separator = Value
    End Set
  End Property

  Public ReadOnly Property OutputPivotArrayData(ByVal lngIndex As Integer) As String
    Get
      'UPGRADE_WARNING: Couldn't resolve default property of object lngIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      OutputPivotArrayData = mstrOutputPivotArray(lngIndex)
    End Get
  End Property

  Public ReadOnly Property OutputPivotArrayDataUBound() As String
    Get
      OutputPivotArrayDataUBound = CStr(UBound(mstrOutputPivotArray))
    End Get
  End Property

  Public Function EventLogAddHeader() As Integer

    ' JDM - 05/12/02 - Fault 4840 - Wrong report type in event log
    If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown Then
      mobjEventLog.AddHeader(clsEventLog.EventLog_Type.eltStandardReport, "Absence Breakdown")
    Else
      mobjEventLog.AddHeader(clsEventLog.EventLog_Type.eltCrossTab, mstrCrossTabName)
    End If

    EventLogAddHeader = mobjEventLog.EventLogID
  End Function

  Public Sub EventLogChangeHeaderStatus(ByRef lngStatus As Integer)
    mobjEventLog.ChangeHeaderStatus(lngStatus)
  End Sub

  'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
  Private Sub Class_Initialize_Renamed()

    ' Initialise the the classes/arrays to be used
    mclsData = New clsDataAccess
    mobjEventLog = New clsEventLog

    ReDim mstrOutputArray_Data(0)

  End Sub
  Public Sub New()
    MyBase.New()
    Class_Initialize_Renamed()
  End Sub

  'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
  Private Sub Class_Terminate_Renamed()

    ' Clear references to classes and clear collection objects
    'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mclsData = Nothing
    'UPGRADE_NOTE: Object mobjEventLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mobjEventLog = Nothing

    ' JPD20030313 Do not drop the tables & columns collections as they can be reused.
    'Set gcoTablePrivileges = Nothing
    'Set gcolColumnPrivilegesCollection = Nothing

  End Sub
  Protected Overrides Sub Finalize()
    Class_Terminate_Renamed()
    MyBase.Finalize()
  End Sub

  Public Function SetPromptedValues(ByRef pavPromptedValues As Object) As Boolean

    ' Purpose : This function calls the individual functions that
    '           generate the components of the main SQL string.
    On Error GoTo ErrorTrap

    Dim fOK As Boolean
    Dim iLoop As Short
    Dim iDataType As Short
    Dim lngComponentID As Integer

    fOK = True

    ReDim mvarPrompts(1, 0)

    If IsArray(pavPromptedValues) Then
      ReDim mvarPrompts(1, UBound(pavPromptedValues, 2))

      For iLoop = 0 To UBound(pavPromptedValues, 2)
        ' Get the prompt data type.
        'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Len(Trim(Mid(pavPromptedValues(0, iLoop), 10))) > 0 Then
          'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          lngComponentID = CInt(Mid(pavPromptedValues(0, iLoop), 10))
          'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          iDataType = CShort(Mid(pavPromptedValues(0, iLoop), 8, 1))

          'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mvarPrompts(0, iLoop) = lngComponentID

          ' NB. Locale to server conversions are done on the client.
          Select Case iDataType
            Case 2
              ' Numeric.
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = CDbl(pavPromptedValues(1, iLoop))
            Case 3
              ' Logic.
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = (UCase(CStr(pavPromptedValues(1, iLoop))) = "TRUE")
            Case 4
              ' Date.
              ' JPD 20040212 Fault 8082 - DO NOT CONVERT DATE PROMPTED VALUES
              ' THEY ARE PASSED IN FROM THE ASPs AS STRING VALUES IN THE CORRECT
              ' FORMAT (mm/dd/yyyy) AND DOING ANY KIND OF CONVERSION JUST SCREWS
              ' THINGS UP.
              'mvarPrompts(1, iLoop) = CDate(pavPromptedValues(1, iLoop))
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = pavPromptedValues(1, iLoop)
            Case Else
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = CStr(pavPromptedValues(1, iLoop))
          End Select
        Else
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mvarPrompts(0, iLoop) = 0
        End If
      Next iLoop
    End If

    SetPromptedValues = fOK

    Exit Function

ErrorTrap:
    mstrStatusMessage = "Error whilst setting prompted values. " & Err.Description
    fOK = False
    SetPromptedValues = False

  End Function


  Public Function RetreiveDefinition() As Boolean

    On Error GoTo LocalErr

    Dim rsCrossTabDefinition As ADODB.Recordset
    Dim strSQL As String

    ReDim mastrUDFsRequired(0)

    ' Define this cross tab as a normal one
    mlngCrossTabType = Declarations.CrossTabType.cttNormal

    strSQL = "SELECT ASRSysCrossTab.*, " & "'TableName' = ASRSysTables.TableName, " & "'RecordDescExprID' = ASRSysTables.RecordDescExprID, " & "'IntersectionColName' = ASRSysColumns.ColumnName, " & "'IntersectionDecimals' = ASRSysColumns.Decimals " & "FROM ASRSysCrossTab " & "JOIN ASRSysTables ON ASRSysCrossTab.TableID = ASRSysTables.TableID " & "LEFT OUTER JOIN ASRSysColumns ON ASRSysCrossTab.IntersectionColID = ASRSysColumns.ColumnID " & "WHERE CrossTabID = " & CStr(mlngCrossTabID)

    rsCrossTabDefinition = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
    If rsCrossTabDefinition.BOF And rsCrossTabDefinition.EOF Then
      'UPGRADE_NOTE: Object rsCrossTabDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      rsCrossTabDefinition = Nothing
      mstrStatusMessage = "This definition has been deleted by another user"
      RetreiveDefinition = False
      Exit Function
    End If

    With rsCrossTabDefinition

      If LCase(CType(.Fields("Username").Value, String)) <> LCase(gsUsername) And CurrentUserAccess(modUtilAccessLog.UtilityType.utlCrossTab, mlngCrossTabID) = ACCESS_HIDDEN Then
        'UPGRADE_NOTE: Object rsCrossTabDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsCrossTabDefinition = Nothing
        mstrStatusMessage = "This definition has been made hidden by another user."
        RetreiveDefinition = False
        Exit Function
      End If

      mlngBaseTableID = .Fields("TableID").Value
      mstrBaseTable = .Fields("TableName").Value
      mlngRecordDescExprID = .Fields("RecordDescExprID").Value
      mstrCrossTabName = .Fields("Name").Value
      mblnChkPicklistFilter = .Fields("PrintFilterHeader").Value

      mblnShowPercentage = .Fields("Percentage").Value
      mblnPercentageofPage = .Fields("PercentageOfPage").Value
      mblnSuppressZeros = .Fields("SuppressZeros").Value
      mbUse1000Separator = .Fields("ThousandSeparators").Value

      'mintDefaultOutput = !DefaultOutput
      'mintDefaultExportTo = !DefaultExportTo
      'mblnDefaultSave = !DefaultSave
      'mstrDefaultSaveAs = !DefaultSaveAs
      'mblnDefaultCloseApp = !DefaultCloseApp
      mblnOutputPreview = .Fields("OutputPreview").Value
      mlngOutputFormat = .Fields("OutputFormat").Value
      mblnOutputScreen = .Fields("OutputScreen").Value
      mblnOutputPrinter = .Fields("OutputPrinter").Value
      mstrOutputPrinterName = .Fields("OutputPrinterName").Value
      mblnOutputSave = .Fields("OutputSave").Value
      mlngOutputSaveExisting = .Fields("OutputSaveExisting").Value
      mblnOutputEmail = .Fields("OutputEmail").Value
      mlngOutputEmailID = .Fields("OutputEmailAddr").Value
      'mstrOutputEmailAddr = GetEmailGroupAddresses(!OutputEmailAddr)
      mstrOutputEmailName = GetEmailGroupName(.Fields("OutputEmailAddr").Value)
      mstrOutputEmailSubject = .Fields("OutputEmailSubject").Value
      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      mstrOutputEmailAttachAs = IIf(IsDbNull(.Fields("OutputEmailAttachAs").Value), vbNullString, .Fields("OutputEmailAttachAs").Value)
      mstrOutputFilename = .Fields("OutputFilename").Value

      mblnOutputPreview = (.Fields("OutputPreview").Value Or (mlngOutputFormat = Declarations.OutputFormats.fmtDataOnly And mblnOutputScreen))

      mlngColID(HOR) = .Fields("HorizontalColID").Value
      mdblMin(HOR) = Val(.Fields("HorizontalStart").Value)
      mdblMax(HOR) = Val(.Fields("HorizontalStop").Value)
      mdblStep(HOR) = Val(.Fields("HorizontalStep").Value)
      mstrColName(HOR) = datGeneral.GetColumnName(mlngColID(HOR))
      mlngColDataType(HOR) = CStr(datGeneral.GetDataType(mlngBaseTableID, mlngColID(HOR)))
      mstrFormat(HOR) = GetFormat(mlngColID(HOR))

      mlngColID(VER) = .Fields("VerticalColID").Value
      mdblMin(VER) = Val(.Fields("VerticalStart").Value)
      mdblMax(VER) = Val(.Fields("VerticalStop").Value)
      mdblStep(VER) = Val(.Fields("VerticalStep").Value)
      mstrColName(VER) = datGeneral.GetColumnName(mlngColID(VER))
      mlngColDataType(VER) = CStr(datGeneral.GetDataType(mlngBaseTableID, mlngColID(VER)))
      mstrFormat(VER) = GetFormat(mlngColID(VER))

      mlngColID(PGB) = .Fields("PageBreakColID").Value
      mblnPageBreak = (mlngColID(PGB) > 0)
      If mblnPageBreak Then
        mstrColName(PGB) = datGeneral.GetColumnName(mlngColID(PGB))
        mlngColDataType(PGB) = CStr(datGeneral.GetDataType(mlngBaseTableID, mlngColID(PGB)))
        mstrFormat(PGB) = GetFormat(mlngColID(PGB))
        mdblMin(PGB) = Val(.Fields("PageBreakStart").Value)
        mdblMax(PGB) = Val(.Fields("PageBreakStop").Value)
        mdblStep(PGB) = Val(.Fields("PageBreakStep").Value)
      End If

      mblnIntersection = (.Fields("IntersectionColID").Value > 0)
      If mblnIntersection Then
        mlngType = .Fields("IntersectionType").Value
        mlngColID(INS) = .Fields("IntersectionColID").Value
        mstrColName(INS) = .Fields("IntersectionColName").Value
        'mstrIntersectionMask = String$(20, "#") & "0." & _
        'String$(CLng(!IntersectionDecimals), "0")
        mlngIntersectionDecimals = CInt(.Fields("IntersectionDecimals").Value)
        mstrIntersectionMask = New String("#", 20) & "0"
        If CInt(.Fields("IntersectionDecimals").Value) > 0 Then
          'mstrIntersectionMask = mstrIntersectionMask & _
          'UI.GetSystemDecimalSeparator & String$(CLng(!IntersectionDecimals), "0")
          mstrIntersectionMask = mstrIntersectionMask & "." & New String("0", CInt(.Fields("IntersectionDecimals").Value))
        End If
      Else
        mlngType = 0
      End If

      fOK = IsRecordSelectionValid(.Fields("PickListID").Value, .Fields("FilterID").Value)
      If fOK = False Then
        Exit Function
      End If

      mstrPicklistFilter = GetPicklistFilterSelect(.Fields("PickListID").Value, .Fields("FilterID").Value)
      If fOK = False Then
        Exit Function
      End If

    End With

    Call UtilUpdateLastRun(modUtilAccessLog.UtilityType.utlCrossTab, mlngCrossTabID)

TidyAndExit:
    RetreiveDefinition = fOK
    'UPGRADE_NOTE: Object rsCrossTabDefinition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsCrossTabDefinition = Nothing

    Exit Function

LocalErr:
    mstrStatusMessage = "Error reading Cross Tab definition"
    fOK = False
    Resume TidyAndExit

  End Function

  Private Function IsRecordSelectionValid(ByVal lngPicklistID As Integer, ByVal lngFilterID As Integer) As Boolean

    Dim sSQL As String
    Dim lCount As Integer
    Dim rsTemp As ADODB.Recordset
    Dim iResult As modUtilityAccess.RecordSelectionValidityCodes
    Dim fCurrentUserIsSysSecMgr As Boolean

    fCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr()

    ' Filter
    If lngFilterID > 0 Then
      iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_FILTER, lngFilterID)
      Select Case iResult
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
          mstrStatusMessage = "The base table filter used in this definition has been deleted by another user."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
          mstrStatusMessage = "The base table filter used in this definition is invalid."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
          If Not fCurrentUserIsSysSecMgr Then
            mstrStatusMessage = "The base table filter used in this definition has been made hidden by another user."
          End If
      End Select
    ElseIf lngPicklistID > 0 Then
      iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_PICKLIST, lngPicklistID)
      Select Case iResult
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
          mstrStatusMessage = "The base table picklist used in this definition has been deleted by another user."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
          mstrStatusMessage = "The base table picklist used in this definition is invalid."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
          If Not fCurrentUserIsSysSecMgr Then
            mstrStatusMessage = "The base table picklist used in this definition has been made hidden by another user."
          End If
      End Select
    End If

    IsRecordSelectionValid = (Len(mstrStatusMessage) = 0)

  End Function

  Private Function GetPicklistFilterSelect(ByVal lngPicklistID As Integer, ByVal lngFilterID As Integer) As String

    Dim rsTemp As ADODB.Recordset

    If lngPicklistID > 0 Then

      mstrStatusMessage = IsPicklistValid(lngPicklistID)
      If mstrStatusMessage <> vbNullString Then
        'mblnInvalidPicklistFilter = True
        fOK = False
        Exit Function
      End If

      'Get List of IDs from Picklist
      rsTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & lngPicklistID, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
      fOK = Not (rsTemp.BOF And rsTemp.EOF)

      If Not fOK Then
        mstrStatusMessage = "The base table picklist contains no records."
      Else
        GetPicklistFilterSelect = vbNullString
        Do While Not rsTemp.EOF
          GetPicklistFilterSelect = GetPicklistFilterSelect & IIf(Len(GetPicklistFilterSelect) > 0, ", ", "") & rsTemp.Fields(0).Value
          rsTemp.MoveNext()
        Loop
      End If

      rsTemp.Close()
      'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      rsTemp = Nothing

      'MH20020704 Fault 4022
      rsTemp = mclsData.OpenRecordset("SELECT name from ASRSysPicklistName WHERE PicklistID = " & CStr(lngPicklistID), ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
      mstrPicklistFilterName = " (Base Table Picklist : " & rsTemp.Fields("Name").Value & ")"
      rsTemp.Close()
      'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      rsTemp = Nothing

    ElseIf lngFilterID > 0 Then

      mstrStatusMessage = IsFilterValid(lngFilterID)
      If mstrStatusMessage <> vbNullString Then
        'mblnInvalidPicklistFilter = True
        fOK = False
        Exit Function
      End If

      'Get list of IDs from Filter
      fOK = datGeneral.FilteredIDs(lngFilterID, GetPicklistFilterSelect, mvarPrompts)

      ' Generate any UDFs that are used in this filter
      If fOK Then
        datGeneral.FilterUDFs(lngFilterID, mastrUDFsRequired)
      End If

      If Not fOK Then
        ' Permission denied on something in the filter.
        mstrStatusMessage = "You do not have permission to use the '" & datGeneral.GetFilterName(lngFilterID) & "' filter."
      End If

      'MH20020704 Fault 4022
      rsTemp = mclsData.OpenRecordset("SELECT Name from ASRSysExpressions WHERE ExprID = " & CStr(lngFilterID), ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
      mstrPicklistFilterName = " (Base Table Filter : " & rsTemp.Fields("Name").Value & ")"
      rsTemp.Close()
      'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      rsTemp = Nothing

    Else
      mstrPicklistFilterName = " (No Picklist Or Filter Selected)"

    End If

  End Function

  Private Function GetFormat(ByVal lngColumnID As Integer) As String

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    strSQL = "SELECT DataType, Size, Decimals FROM ASRSysColumns Where ColumnID = " & CStr(lngColumnID)
    rsTemp = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    Select Case rsTemp.Fields("DataType").Value
      Case Declarations.SQLDataType.sqlNumeric
        GetFormat = New String("#", rsTemp.Fields("Size").Value - 1) & "0"
        If rsTemp.Fields("Decimals").Value > 0 Then
          'GetFormat = GetFormat & UI.GetSystemDecimalSeparator & String$(rsTemp!Decimals, "0")
          GetFormat = GetFormat & "." & New String("0", rsTemp.Fields("Decimals").Value)
        End If

      Case Declarations.SQLDataType.sqlInteger
        GetFormat = New String("#", 9) & "0"

    End Select

    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing

  End Function

  Public Function CreateTempTable() As Boolean

    Dim strColumn(,) As String
    Dim strSQL As String
    Dim lngMax As Integer

    On Error GoTo LocalErr

    lngMax = 2
    ReDim strColumn(2, lngMax)

    strColumn(1, 0) = "ID"
    strColumn(2, 0) = "ID"

    strColumn(1, 1) = mstrColName(HOR)
    strColumn(2, 1) = "Hor"

    strColumn(1, 2) = mstrColName(VER)
    strColumn(2, 2) = "Ver"

    If mblnPageBreak Then
      lngMax = lngMax + 1
      ReDim Preserve strColumn(2, lngMax)

      strColumn(1, lngMax) = mstrColName(PGB)
      strColumn(2, lngMax) = "Pgb"
    End If

    If mblnIntersection Then
      lngMax = lngMax + 1
      ReDim Preserve strColumn(2, lngMax)

      strColumn(1, lngMax) = mstrColName(INS)
      strColumn(2, lngMax) = "Ins"
    End If

    'MH20020321 Remmed out for INT
    If mlngCrossTabType <> Declarations.CrossTabType.cttNormal Then
      If mlngCrossTabType <> Declarations.CrossTabType.cttAbsenceBreakdown Then
        lngMax = lngMax + 2
        ReDim Preserve strColumn(2, lngMax)

        strColumn(1, lngMax - 1) = gsPersonnelStartDateColumnName
        strColumn(2, lngMax - 1) = "StartDate"

        strColumn(1, lngMax) = gsPersonnelLeavingDateColumnName
        strColumn(2, lngMax) = "LeavingDate"
      End If

      If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown Then
        lngMax = lngMax + 7
        ReDim Preserve strColumn(2, lngMax)

        strColumn(1, lngMax) = gsAbsenceDurationColumnName
        strColumn(2, lngMax) = "Value"

        strColumn(1, lngMax - 4) = gsAbsenceStartDateColumnName
        strColumn(2, lngMax - 4) = "Start_Date"

        strColumn(1, lngMax - 3) = gsAbsenceStartSessionColumnName
        strColumn(2, lngMax - 3) = "Start_Session"

        strColumn(1, lngMax - 2) = gsAbsenceEndDateColumnName
        strColumn(2, lngMax - 2) = "End_Date"

        strColumn(1, lngMax - 1) = gsAbsenceEndSessionColumnName
        strColumn(2, lngMax - 1) = "End_Session"

        strColumn(1, lngMax - 5) = "ID_" & Trim(Str(glngPersonnelTableID))
        strColumn(2, lngMax - 5) = "Personnel_ID"

        strColumn(1, lngMax - 6) = gsAbsenceDurationColumnName ' Used to hold the day number (1=Mon, 2=Tues etc.)
        strColumn(2, lngMax - 6) = "Day_Number"


      End If

    End If

    fOK = True
    Call GetSQL2(strColumn)
    If fOK = False Then
      CreateTempTable = False
      Exit Function
    End If

    mstrTempTableName = datGeneral.UniqueSQLObjectName("ASRSysTempCrossTab", 3)
    mstrSQLSelect = mstrSQLSelect & ", " & "space(255) as 'RecDesc' INTO " & mstrTempTableName


    strSQL = "SELECT " & mstrSQLSelect & vbNewLine & " FROM " & mstrSQLFrom & vbNewLine & mstrSQLJoin & vbNewLine & mstrSQLWhere

    'MH20010327 Seems that it might be moving on pass this line of code too
    'quickly so I've tried returning the number of rows effected to make
    'sure that it completes fully
    'mclsData.Execute strSQL
    mclsData.ExecuteSqlReturnAffected(strSQL)

    'Dim tt As Double
    'tt = Timer + 2
    'Do While Timer < tt
    '  DoEvents
    'Loop

    '  Set mcmdUpdateRecDescs = New ADODB.Command
    '  mcmdUpdateRecDescs.ActiveConnection = gADOCon
    '  mcmdUpdateRecDescs.CommandText = strSQL
    '  mcmdUpdateRecDescs.CommandTimeout = 0
    '  mcmdUpdateRecDescs.Execute , , adAsyncExecute
    '
    '  Do While mcmdUpdateRecDescs.State = adStateExecuting
    '    DoEvents
    '  Loop
    '
    '  Set mcmdUpdateRecDescs = Nothing


    strSQL = "SELECT * FROM " & mstrTempTableName

    'MH20020321 Remmed out for INT
    '  If mlngCrossTabType <> cttNormal Then
    '    If mlngCrossTabType <> cttAbsenceBreakdown Then
    '      strSQL = strSQL & " WHERE " & _
    ''        SQLEmployedAtStartOfReport("startdate", "leavingdate")
    '    End If
    '  End If

    rsCrossTabData = New ADODB.Recordset
    rsCrossTabData.let_ActiveConnection(gADOCon)
    'rsCrossTabData.Properties("Preserve On Commit") = True
    'rsCrossTabData.Properties("Preserve On Abort") = True
    ' rsCrossTabData.Open strSQL, , adOpenKeyset, adLockReadOnly, adCmdText
    rsCrossTabData = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

    If rsCrossTabData.EOF Then
      mstrStatusMessage = "No records meet selection criteria"
      mblnNoRecords = True
      mobjEventLog.AddDetailEntry("Completed successfully. " & mstrStatusMessage)
      mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsSuccessful)
      fOK = False
    End If

    'Check if we might need record description...
    'If mblnBatchMode = False And mlngRecordDescExprID > 0 Then
    If fOK Then
      mcmdUpdateRecDescs = New ADODB.Command
      mcmdUpdateRecDescs.let_ActiveConnection(gADOCon)
      mcmdUpdateRecDescs.CommandText = "EXEC dbo.sp_ASRCrossTabsRecDescs '" & mstrTempTableName & "', " & CStr(mlngRecordDescExprID)
      mcmdUpdateRecDescs.CommandTimeout = 0

      'MH20020321 Remmed out for INT (but would prefer record descriptions now anyway!)
      ' JDM - Fault 2421 - Must have record description calculated NOW.
      '    If mlngCrossTabType = cttAbsenceBreakdown Then
      mcmdUpdateRecDescs.Execute()
      '    Else
      '      mcmdUpdateRecDescs.Execute , , adAsyncExecute
      '    End If
      'End If

    End If

    CreateTempTable = fOK

    Exit Function

LocalErr:
    'mstrStatusMessage = "Error retrieving data" & vbNewLine & _
    '"(" & Err.Description & ")"
    mstrStatusMessage = Err.Description
    CreateTempTable = False

  End Function

  Private Sub GetSQL2(ByRef strCol(,) As String)

    Dim objTableView As CTablePrivilege
    Dim objColumnPrivileges As CColumnPrivileges
    Dim sRealSource As String
    Dim sSource As String
    Dim lngCount As Integer
    Dim fColumnOK As Boolean
    Dim alngTableViews(,) As Integer
    Dim iNextIndex As Short
    Dim fFound As Boolean

    Dim sCaseStatement As String
    Dim sWhereColumn As String
    Dim strSelectedRecords As String
    Dim sWhereIDs As String
    Dim blnOK As Boolean
    Dim strColumn As String
    Dim blnCharColumn As Boolean


    On Error GoTo LocalErr

    fOK = True
    ReDim alngTableViews(2, 0)

    mstrSQLFrom = gcoTablePrivileges.Item(mstrBaseTable).RealSource
    mstrSQLSelect = vbNullString
    mstrSQLJoin = vbNullString
    Dim asViews(0) As Object

    blnCharColumn = (Val(mlngColDataType(lngCount)) = Declarations.SQLDataType.sqlVarChar)


    For lngCount = 0 To UBound(strCol, 2)

      objColumnPrivileges = GetColumnPrivileges(mstrBaseTable)
      fColumnOK = objColumnPrivileges.IsValid(strCol(1, lngCount))
      If fColumnOK Then
        fColumnOK = objColumnPrivileges.Item(strCol(1, lngCount)).AllowSelect

        If fColumnOK Then
          fColumnOK = gcoTablePrivileges.Item(mstrBaseTable).AllowSelect
        End If

      End If

      'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      objColumnPrivileges = Nothing

      If lngCount <= UBound(mlngColDataType) Then
        blnCharColumn = (Val(mlngColDataType(lngCount)) = Declarations.SQLDataType.sqlVarChar)
      End If

      If fColumnOK Then
        ' The column can be read from the base table/view, or directly from a parent table.
        ' Add the column to the column list.

        If strSelectedRecords = vbNullString And mstrPicklistFilter <> vbNullString Then

          'MH20020321 Remmed out for INT
          If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown Then
            strSelectedRecords = mstrSQLFrom & ".ID_" & Trim(Str(glngPersonnelTableID)) & " IN (" & mstrPicklistFilter & ")"
          Else
            strSelectedRecords = mstrSQLFrom & ".ID IN (" & mstrPicklistFilter & ")"
          End If

        End If

        'mstrSQLSelect = mstrSQLSelect & _
        'IIf(Len(mstrSQLSelect) > 0, ", ", "") & _
        'mstrSQLFrom & "." & strCol(1, lngCount) & _
        '" AS '" & strCol(2, lngCount) & "'"

        strColumn = mstrSQLFrom & "." & strCol(1, lngCount)
        If blnCharColumn Then
          strColumn = FormatSQLColumn(strColumn)
        End If

        mstrSQLSelect = mstrSQLSelect & IIf(Len(mstrSQLSelect) > 0, ", ", "") & strColumn & " AS '" & strCol(2, lngCount) & "'"

      Else

        ReDim asViews(0)
        For Each objTableView In gcoTablePrivileges.Collection

          'Loop thru all of the views for this table where the user has select access
          If (Not objTableView.IsTable) And (objTableView.TableID = mlngBaseTableID) And (objTableView.AllowSelect) Then

            sSource = objTableView.ViewName

            ' Get the column permission for the view.
            objColumnPrivileges = GetColumnPrivileges(sSource)

            If objColumnPrivileges.IsValid(strCol(1, lngCount)) Then
              If objColumnPrivileges.Item(strCol(1, lngCount)).AllowSelect Then
                ' Add the view info to an array to be put into the column list or order code below.
                iNextIndex = UBound(asViews) + 1
                ReDim Preserve asViews(iNextIndex)
                'UPGRADE_WARNING: Couldn't resolve default property of object asViews(iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                asViews(iNextIndex) = sSource


                '=== This is the join code section ===
                ' Add the view to the Join code.
                ' Check if the view has already been added to the join code.
                fFound = False
                For iNextIndex = 1 To UBound(alngTableViews, 2)
                  If alngTableViews(2, iNextIndex) = objTableView.ViewID Then
                    fFound = True
                    Exit For
                  End If
                Next iNextIndex

                If Not fFound Then
                  ' The view has not yet been added to the join code, so add it to the array and the join code.
                  ' (also include the picklist info)

                  iNextIndex = UBound(alngTableViews, 2) + 1
                  ReDim Preserve alngTableViews(2, iNextIndex)
                  alngTableViews(1, iNextIndex) = 1
                  alngTableViews(2, iNextIndex) = objTableView.ViewID

                  mstrSQLJoin = mstrSQLJoin & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & mstrSQLFrom & ".ID = " & sSource & ".ID"

                  sWhereIDs = sWhereIDs & IIf(sWhereIDs <> vbNullString, " OR ", vbNullString) & mstrSQLFrom & ".ID IN (SELECT ID FROM " & sSource & ")"

                  'If mstrPicklistFilter <> vbNullString Then
                  strSelectedRecords = strSelectedRecords & IIf(strSelectedRecords <> vbNullString, " OR ", vbNullString) & "(" & IIf(mstrPicklistFilter <> vbNullString, sSource & ".ID IN (" & mstrPicklistFilter & ") AND ", vbNullString) & sSource & ".ID > 0)"
                  'End If

                End If
              End If
              '=== End of Join Code ===


              'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
              objColumnPrivileges = Nothing
            End If

          End If
        Next objTableView
        'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objTableView = Nothing

        ' The current user does have permission to 'read' the column through a/some view(s) on the
        ' table.
        If UBound(asViews) = 0 Then
          fOK = False
          'MH20010716 Fault 2497
          'If its the ID column they they don't have any access to the table.
          'mstrStatusMessage = "You do not have permission to see the column '" & strCol(1, lngCount) & "' " & _
          '"either directly or through any views." & vbNewLine
          mstrStatusMessage = "You do not have permission to see the " & IIf(strCol(1, lngCount) = "ID", "table '" & mstrBaseTable, "column '" & strCol(1, lngCount)) & "' either directly or through any views." & vbNewLine
          Exit Sub
        Else
          ' Add the column to the column list.
          '        sCaseStatement = "CASE"
          '        sWhereColumn = vbNullString
          '        For iNextIndex = 1 To UBound(asViews)
          '          sCaseStatement = sCaseStatement & _
          ''            " WHEN NOT " & asViews(iNextIndex) & "." & strCol(1, lngCount) & " IS NULL THEN " & asViews(iNextIndex) & "." & strCol(1, lngCount) & vbNewLine
          '        Next iNextIndex
          '
          '        If Len(sCaseStatement) > 0 Then
          '          sCaseStatement = sCaseStatement & _
          ''            " ELSE NULL END AS " & _
          ''            "'" & strCol(2, lngCount) & "'"
          '
          '          mstrSQLSelect = mstrSQLSelect & _
          ''            IIf(Len(mstrSQLSelect) > 0, ", ", "") & vbNewLine & _
          ''            sCaseStatement
          '        End If
          sCaseStatement = ""
          For iNextIndex = 1 To UBound(asViews)
            'UPGRADE_WARNING: Couldn't resolve default property of object asViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sCaseStatement = sCaseStatement & IIf(sCaseStatement <> "", vbCrLf & " , ", "") & asViews(iNextIndex) & "." & strCol(1, lngCount)
          Next iNextIndex

          If Len(sCaseStatement) > 0 Then
            ' NPG20100820 Fault HRPRO 1080
            '          If InStr(strColumn, ",") > 0 Then
            '            strColumn = "COALESCE(" & sCaseStatement & ")"
            '          Else
            strColumn = "COALESCE(" & sCaseStatement & ", NULL)"
            '          End If

            If blnCharColumn Then
              strColumn = FormatSQLColumn(strColumn)
            End If

            mstrSQLSelect = mstrSQLSelect & IIf(Len(mstrSQLSelect) > 0, ", ", "") & vbCrLf & strColumn & "AS '" & strCol(2, lngCount) & "'"
          End If

        End If
      End If
    Next

    If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown And Not msAbsenceBreakdownTypes = vbNullString Then
      mstrSQLWhere = mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & "(UPPER(" & gsAbsenceTypeColumnName & ") IN " & msAbsenceBreakdownTypes & ")"
    End If

    If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown Then
      mstrSQLWhere = mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & "( " & gsAbsenceStartDateColumnName & " <= CONVERT(datetime, '" & mstrReportEndDate & "'))" & "And (" & gsAbsenceEndDateColumnName & " >= CONVERT(datetime, '" & mstrReportStartDate & "') OR " & gsAbsenceEndDateColumnName & " IS NULL)"
      '" >= CONVERT(datetime, '" & mstrReportStartDate & "'))"
      'MH20060626 Fault 11260
    End If

    If strSelectedRecords <> vbNullString Then
      mstrSQLWhere = mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & "(" & strSelectedRecords & ")"
    End If

    Exit Sub

LocalErr:
    mstrStatusMessage = "Error retrieving data"
    fOK = False

  End Sub

  Public Function GetHeadingsAndSearches() As Boolean

    Dim strHeading() As String
    Dim strSearch() As String
    Dim lngLoop As Integer


    On Error GoTo LocalErr

    For lngLoop = 0 To 2

      ReDim strHeading(0)
      ReDim strSearch(0)

      If lngLoop = 2 And mblnPageBreak = False Then
        'When no page break field is specified
        strHeading(0) = "<None>"

      Else
        GetHeadingsAndSearchesForColumns(lngLoop, strHeading, strSearch)

      End If

      'Store each array in an array of variants (an array in an array!)
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(lngLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      mvarHeadings(lngLoop) = VB6.CopyArray(strHeading)
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches(lngLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      mvarSearches(lngLoop) = VB6.CopyArray(strSearch)

    Next

    GetHeadingsAndSearches = fOK
    Exit Function

LocalErr:
    mstrStatusMessage = "Error building headings and search arrays"
    GetHeadingsAndSearches = False

  End Function

  Private Sub GetHeadingsAndSearchesForColumns(ByRef lngLoop As Integer, ByRef strHeading() As String, ByRef strSearch() As String)

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strFieldValue As String
    Dim lngCount As Integer
    Dim dblGroup As Double
    Dim dblGroupMax As Double
    Dim dblUnit As Double
    Dim strColumnName As String
    Dim strWhereEmpty As String


    'UPGRADE_WARNING: Couldn't resolve default property of object Choose(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    strColumnName = Choose(lngLoop + 1, "Hor", "Ver", "Pgb")

    If mdblMin(lngLoop) = 0 And mdblMax(lngLoop) = 0 Then

      lngCount = 0

      strWhereEmpty = strColumnName & " IS NULL"
      If mlngColDataType(lngLoop) <> CStr(Declarations.SQLDataType.sqlNumeric) And mlngColDataType(lngLoop) <> CStr(Declarations.SQLDataType.sqlInteger) And mlngColDataType(lngLoop) <> CStr(Declarations.SQLDataType.sqlBoolean) Then
        strWhereEmpty = strWhereEmpty & " OR RTrim(" & strColumnName & ") = ''"
      End If

      'MH20010327 Always add <empty> to see if that helps problems
      '''Check for Empty
      ''strSQL = "SELECT DISTINCT " & strColumnName & _
      '''         " FROM " & mstrTempTableName & _
      '''         " WHERE " & strWhereEmpty
      ''Set rsTemp = mclsData.OpenRecordset(strSQL, adOpenDynamic, adLockReadOnly)

      ''If Not (rsTemp.BOF And rsTemp.EOF) Then

      ' Don't put in empty clauses if we're running an absence breakdown
      If mlngCrossTabType <> Declarations.CrossTabType.cttAbsenceBreakdown Then
        ReDim Preserve strHeading(lngCount)
        ReDim Preserve strSearch(lngCount)
        strHeading(lngCount) = "<Empty>"
        strSearch(lngCount) = strWhereEmpty
        lngCount = lngCount + 1
      End If
      ''End If

      If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown And strColumnName = "Hor" Then
        strSQL = "SELECT DISTINCT " & FormatSQLColumn(strColumnName) & ",Day_Number, DisplayOrder" & " FROM " & mstrTempTableName & " ORDER BY DisplayOrder"
      Else
        strSQL = "SELECT DISTINCT " & FormatSQLColumn(strColumnName) & " FROM " & mstrTempTableName & " ORDER BY 1"
      End If

      rsTemp = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)

      With rsTemp

        If .BOF And .EOF Then
          Exit Sub
        End If

        .MoveFirst()
        Do While Not .EOF 'And Not gobjProgress.Cancelled

          'MH20010213 Had to make this change so that working pattern would work
          'The field has spaces at the begining
          'strFieldValue = IIf(IsNull(.Fields(0).Value), vbNullString, Trim(.Fields(0).Value))
          'If strFieldValue <> vbNullString Then
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          strFieldValue = IIf(IsDbNull(.Fields(0).Value), vbNullString, .Fields(0).Value)

          If Trim(strFieldValue) <> vbNullString Then
            ReDim Preserve strHeading(lngCount)
            ReDim Preserve strSearch(lngCount)

            Select Case mlngColDataType(lngLoop)
              Case CStr(Declarations.SQLDataType.sqlDate)
                'strHeading(lngCount) = Format(.Fields(0).Value, mclsUI.DateFormat)
                strHeading(lngCount) = VB6.Format(.Fields(0).Value, mstrClientDateFormat)
                strSearch(lngCount) = strColumnName & " = '" & VB6.Format(.Fields(0).Value, "mm/dd/yyyy") & "'"

              Case CStr(Declarations.SQLDataType.sqlBoolean)
                strHeading(lngCount) = IIf(.Fields(0).Value, "True", "False")
                strSearch(lngCount) = strColumnName & " = " & IIf(.Fields(0).Value, "1", "0")

              Case CStr(Declarations.SQLDataType.sqlNumeric), CStr(Declarations.SQLDataType.sqlInteger)
                'strHeading(lngCount) = .Fields(0).Value
                'strSearch(lngCount) = strColumnName & " = " & .Fields(0).Value

                strHeading(lngCount) = datGeneral.ConvertNumberForDisplay(VB6.Format(.Fields(0).Value, mstrFormat(lngLoop)))
                strSearch(lngCount) = strColumnName & " = " & datGeneral.ConvertNumberForSQL(.Fields(0).Value)

              Case Else
                '''MH20011008 Fault 2922
                '''strHeading(lngCount) = .Fields(0).Value
                '''strSearch(lngCount) = strColumnName & " = '" & Replace(strFieldValue, "'", "''") & "'"

                'MH20050105 Fault 9608
                'strHeading(lngCount) = Trim(.Fields(0).Value)
                'strSearch(lngCount) = strColumnName & " = '" & Trim(Replace(strFieldValue, "'", "''")) & "'"
                'strHeading(lngCount) = Trim(.Fields(0).Value)
                'strSearch(lngCount) = "ltrim(rtrim(" & strColumnName & ")) = '" & Trim(Replace(strFieldValue, "'", "''")) & "'"
                strHeading(lngCount) = FormatString(.Fields(0).Value)
                strSearch(lngCount) = FormatSQLColumn(strColumnName) & " = '" & Replace(strFieldValue, "'", "''") & "'"

            End Select

            lngCount = lngCount + 1

            '        Else
            '
            '          ReDim Preserve strHeading(lngCount) As String
            '          ReDim Preserve strSearch(lngCount) As String
            '          strHeading(lngCount) = "<Empty>"
            '          strSearch(lngCount) = strColumnName & " IS NULL"
            '
            '          If mlngColDataType(lngLoop) <> sqlNumeric And _
            ''             mlngColDataType(lngLoop) <> sqlInteger And _
            ''             mlngColDataType(lngLoop) <> sqlBoolean Then
            '            strSearch(lngCount) = strSearch(lngCount) & _
            ''                " OR RTrim(" & strColumnName & ") = ''"
            '          End If
            '
            '          lngCount = lngCount + 1
            '
          End If

          .MoveNext()
        Loop
      End With

    Else

      ReDim Preserve strHeading(1)
      ReDim Preserve strSearch(1)

      'First element of range for null values...
      strHeading(0) = "<Empty>"
      strSearch(0) = strColumnName & " IS NULL"

      'Second element of range for those less than minimum value of range...
      strHeading(1) = "< " & datGeneral.ConvertNumberForDisplay(VB6.Format(mdblMin(lngLoop), mstrFormat(lngLoop)))
      'MH20010411 Fault 1978 Convert to int stops overflow error !
      'strSearch(1) = "Convert(int," & strColumnName & ") < " & datGeneral.ConvertNumberForSQL(mdblMin(lngLoop))
      strSearch(1) = "Convert(float," & strColumnName & ") < " & datGeneral.ConvertNumberForSQL(CStr(mdblMin(lngLoop)))

      dblUnit = GetSmallestUnit(lngLoop)

      If mdblStep(lngLoop) = 0 Then
        mstrStatusMessage = "Step value for " & strColumnName & " column cannot be zero"
        fOK = False
        Exit Sub
      End If

      lngCount = 2
      For dblGroup = mdblMin(lngLoop) To mdblMax(lngLoop) Step mdblStep(lngLoop)
        ReDim Preserve strHeading(lngCount)
        ReDim Preserve strSearch(lngCount)
        dblGroupMax = dblGroup + mdblStep(lngLoop) - dblUnit
        strHeading(lngCount) = datGeneral.ConvertNumberForDisplay(VB6.Format(dblGroup, mstrFormat(lngLoop))) & IIf(dblGroupMax <> dblGroup, " - " & datGeneral.ConvertNumberForDisplay(VB6.Format(dblGroupMax, mstrFormat(lngLoop))), "")
        'MH20010411 Fault 1978 Convert to int stops overflow error !
        'strSearch(lngCount) = "Convert(int," & strColumnName & ") BETWEEN " & _
        'datGeneral.ConvertNumberForSQL(dblGroup) & " AND " & datGeneral.ConvertNumberForSQL(dblGroupMax)
        strSearch(lngCount) = "Convert(float," & strColumnName & ") BETWEEN " & datGeneral.ConvertNumberForSQL(CStr(dblGroup)) & " AND " & datGeneral.ConvertNumberForSQL(CStr(dblGroupMax))

        lngCount = lngCount + 1
      Next

      ReDim Preserve strHeading(lngCount)
      ReDim Preserve strSearch(lngCount)
      'Last element of range for those more than maximum value of range...
      strHeading(lngCount) = "> " & datGeneral.ConvertNumberForDisplay(VB6.Format(dblGroup - dblUnit, mstrFormat(lngLoop)))
      'MH20010411 Fault 1978 Convert to int stops overflow error !
      'strSearch(lngCount) = "Convert(int," & strColumnName & ") > " & datGeneral.ConvertNumberForSQL(dblGroup - dblUnit)
      strSearch(lngCount) = "Convert(float," & strColumnName & ") > " & datGeneral.ConvertNumberForSQL(CStr(dblGroup - dblUnit))

      lngCount = lngCount + 1
    End If

  End Sub

  Private Function GetSmallestUnit(ByRef lngLoop As Integer) As Double

    'e.g. mstrFormat(lngLoop) = 0.0,   GetSmallestUnit = 0.1
    '     mstrFormat(lngLoop) = 0.000, GetSmallestUnit = 0.001
    '     mstrFormat(lngLoop) = #0,    GetSmallestUnit = 1
    '     mstrFormat(lngLoop) = #0.00, GetSmallestUnit = 0.01

    '  Dim intFound As Integer
    '
    '  intFound = InStr(mstrFormat(lngLoop), UI.GetSystemDecimalSeparator)
    '  If intFound > 0 Then
    '    GetSmallestUnit = Mid$(mstrFormat(lngLoop), intFound, Len(mstrFormat(lngLoop)) - intFound) & "1"
    '  Else
    '    GetSmallestUnit = 1
    '  End If

    Dim strTemp As String
    Dim intFound As Short

    intFound = InStr(mstrFormat(lngLoop), ".")
    If intFound > 0 Then
      strTemp = Mid(mstrFormat(lngLoop), intFound, Len(mstrFormat(lngLoop)) - intFound) & "1"
      GetSmallestUnit = CDbl(datGeneral.ConvertNumberForDisplay(strTemp))
    Else
      GetSmallestUnit = 1
    End If

  End Function

  Public Function BuildTypeArray() As Boolean

    On Error GoTo LocalErr

    If mblnIntersection Then

      ReDim mstrType(4)
      mstrType(TYPECOUNT) = "Count"
      mstrType(TYPEAVERAGE) = "Average"
      mstrType(TYPEMAXIMUM) = "Maximum"
      mstrType(TYPEMINIMUM) = "Minimum"
      mstrType(TYPETOTAL) = "Total"

    Else

      ReDim mstrType(0)
      mstrType(TYPECOUNT) = "Count"

    End If

    BuildTypeArray = True

    Exit Function

LocalErr:
    mstrStatusMessage = "Error building type array"
    BuildTypeArray = False

  End Function

  Public Function BuildDataArrays() As Boolean

    Dim strTempValue As String
    Dim dblThisIntersectionVal As Double

    Dim lngCol As Integer
    Dim lngRow As Integer
    Dim lngPage As Integer
    Dim lngNumCols As Integer
    Dim lngNumRows As Integer
    Dim lngNumPages As Integer


    On Error GoTo LocalErr

    lngNumCols = UBound(mvarHeadings(0))
    lngNumRows = UBound(mvarHeadings(1))
    lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(2)), 0)

    ReDim mdblDataArray(lngNumCols, lngNumRows, lngNumPages, 4)
    ReDim mdblHorTotal(lngNumCols, lngNumPages, 4)
    ReDim mdblVerTotal(lngNumRows, lngNumPages, 4)
    ReDim mdblPgbTotal(lngNumCols, lngNumRows + 1, 4) '+1 for totals !
    ReDim mdblPageTotal(lngNumPages, 4)
    ReDim mdblGrandTotal(4)

    With rsCrossTabData

      .MoveFirst()
      Do While Not .EOF

        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        strTempValue = IIf(Not IsDbNull(.Fields("HOR").Value), .Fields("HOR").Value, vbNullString)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lngCol = GetGroupNumber(strTempValue, HOR)

        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        strTempValue = IIf(Not IsDbNull(.Fields("VER").Value), .Fields("VER").Value, vbNullString)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lngRow = GetGroupNumber(strTempValue, VER)

        If mblnPageBreak Then
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          strTempValue = IIf(Not IsDbNull(.Fields("PGB").Value), .Fields("PGB").Value, vbNullString)
          'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          lngPage = GetGroupNumber(strTempValue, PGB)
        Else
          lngPage = 0
        End If

        'Count
        mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) = mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) + 1
        mdblHorTotal(lngCol, lngPage, TYPECOUNT) = mdblHorTotal(lngCol, lngPage, TYPECOUNT) + 1
        mdblVerTotal(lngRow, lngPage, TYPECOUNT) = mdblVerTotal(lngRow, lngPage, TYPECOUNT) + 1
        mdblPgbTotal(lngCol, lngRow, TYPECOUNT) = mdblPgbTotal(lngCol, lngRow, TYPECOUNT) + 1
        mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) = mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) + 1
        mdblPageTotal(lngPage, TYPECOUNT) = mdblPageTotal(lngPage, TYPECOUNT) + 1
        mdblGrandTotal(TYPECOUNT) = mdblGrandTotal(TYPECOUNT) + 1

        'If mblnIntersection And IsNull(.Fields(.Fields.Count - 1)) = False Then
        If mblnIntersection Then

          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          If IsDbNull(.Fields("INS").Value) Then
            dblThisIntersectionVal = 0
          Else
            dblThisIntersectionVal = Val(datGeneral.ConvertNumberForSQL(.Fields("INS").Value))
          End If


          'Total
          mdblDataArray(lngCol, lngRow, lngPage, TYPETOTAL) = mdblDataArray(lngCol, lngRow, lngPage, TYPETOTAL) + dblThisIntersectionVal
          mdblHorTotal(lngCol, lngPage, TYPETOTAL) = mdblHorTotal(lngCol, lngPage, TYPETOTAL) + dblThisIntersectionVal
          mdblVerTotal(lngRow, lngPage, TYPETOTAL) = mdblVerTotal(lngRow, lngPage, TYPETOTAL) + dblThisIntersectionVal
          mdblPgbTotal(lngCol, lngRow, TYPETOTAL) = mdblPgbTotal(lngCol, lngRow, TYPETOTAL) + dblThisIntersectionVal
          mdblPgbTotal(lngCol, lngNumRows + 1, TYPETOTAL) = mdblPgbTotal(lngCol, lngNumRows + 1, TYPETOTAL) + dblThisIntersectionVal
          mdblPageTotal(lngPage, TYPETOTAL) = mdblPageTotal(lngPage, TYPETOTAL) + dblThisIntersectionVal
          mdblGrandTotal(TYPETOTAL) = mdblGrandTotal(TYPETOTAL) + dblThisIntersectionVal

          'Average
          mdblDataArray(lngCol, lngRow, lngPage, TYPEAVERAGE) = mdblDataArray(lngCol, lngRow, lngPage, TYPETOTAL) / mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT)
          mdblHorTotal(lngCol, lngPage, TYPEAVERAGE) = mdblHorTotal(lngCol, lngPage, TYPETOTAL) / mdblHorTotal(lngCol, lngPage, TYPECOUNT)
          mdblVerTotal(lngRow, lngPage, TYPEAVERAGE) = mdblVerTotal(lngRow, lngPage, TYPETOTAL) / mdblVerTotal(lngRow, lngPage, TYPECOUNT)
          mdblPgbTotal(lngCol, lngRow, TYPEAVERAGE) = mdblPgbTotal(lngCol, lngRow, TYPETOTAL) / mdblPgbTotal(lngCol, lngRow, TYPECOUNT)
          mdblPgbTotal(lngCol, lngNumRows + 1, TYPEAVERAGE) = mdblPgbTotal(lngCol, lngNumRows + 1, TYPETOTAL) / mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT)
          mdblPageTotal(lngPage, TYPEAVERAGE) = mdblPageTotal(lngPage, TYPETOTAL) / mdblPageTotal(lngPage, TYPECOUNT)
          mdblGrandTotal(TYPEAVERAGE) = mdblGrandTotal(TYPETOTAL) / mdblGrandTotal(TYPECOUNT)

          'Minimum
          If dblThisIntersectionVal < mdblDataArray(lngCol, lngRow, lngPage, TYPEMINIMUM) Or mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) = 1 Then
            mdblDataArray(lngCol, lngRow, lngPage, TYPEMINIMUM) = dblThisIntersectionVal
            If dblThisIntersectionVal < mdblHorTotal(lngCol, lngPage, TYPEMINIMUM) Or mdblHorTotal(lngCol, lngPage, TYPECOUNT) = 1 Then
              mdblHorTotal(lngCol, lngPage, TYPEMINIMUM) = dblThisIntersectionVal
            End If
            If dblThisIntersectionVal < mdblVerTotal(lngRow, lngPage, TYPEMINIMUM) Or mdblVerTotal(lngRow, lngPage, TYPECOUNT) = 1 Then
              mdblVerTotal(lngRow, lngPage, TYPEMINIMUM) = dblThisIntersectionVal
            End If
            If dblThisIntersectionVal < mdblPgbTotal(lngCol, lngRow, TYPEMINIMUM) Or mdblPgbTotal(lngCol, lngRow, TYPECOUNT) = 1 Then
              mdblPgbTotal(lngCol, lngRow, TYPEMINIMUM) = dblThisIntersectionVal
            End If
            If dblThisIntersectionVal < mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMINIMUM) Or mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) = 1 Then
              mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMINIMUM) = dblThisIntersectionVal
            End If
            If dblThisIntersectionVal < mdblPageTotal(lngPage, TYPEMINIMUM) Or mdblPageTotal(lngPage, TYPECOUNT) = 1 Then
              mdblPageTotal(lngPage, TYPEMINIMUM) = dblThisIntersectionVal
              If dblThisIntersectionVal < mdblGrandTotal(TYPEMINIMUM) Or mdblGrandTotal(TYPECOUNT) = 1 Then
                mdblGrandTotal(TYPEMINIMUM) = dblThisIntersectionVal
              End If
            End If
          End If

          'Maximum
          If dblThisIntersectionVal > mdblDataArray(lngCol, lngRow, lngPage, TYPEMAXIMUM) Or mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) = 1 Then
            mdblDataArray(lngCol, lngRow, lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
            If dblThisIntersectionVal > mdblHorTotal(lngCol, lngPage, TYPEMAXIMUM) Or mdblHorTotal(lngCol, lngPage, TYPECOUNT) = 1 Then
              mdblHorTotal(lngCol, lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
            End If
            If dblThisIntersectionVal > mdblVerTotal(lngRow, lngPage, TYPEMAXIMUM) Or mdblVerTotal(lngRow, lngPage, TYPECOUNT) = 1 Then
              mdblVerTotal(lngRow, lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
            End If
            If dblThisIntersectionVal > mdblPgbTotal(lngCol, lngRow, TYPEMAXIMUM) Or mdblPgbTotal(lngCol, lngRow, TYPECOUNT) = 1 Then
              mdblPgbTotal(lngCol, lngRow, TYPEMAXIMUM) = dblThisIntersectionVal
            End If
            If dblThisIntersectionVal > mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMAXIMUM) Or mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) = 1 Then
              mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMAXIMUM) = dblThisIntersectionVal
            End If
            If dblThisIntersectionVal > mdblPageTotal(lngPage, TYPEMAXIMUM) Or mdblPageTotal(lngPage, TYPECOUNT) = 1 Then
              mdblPageTotal(lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
              If dblThisIntersectionVal > mdblGrandTotal(TYPEMAXIMUM) Or mdblGrandTotal(TYPECOUNT) = 1 Then
                mdblGrandTotal(TYPEMAXIMUM) = dblThisIntersectionVal
              End If
            End If
          End If
        End If

        .MoveNext()
      Loop

    End With

    'UPGRADE_NOTE: Object rsCrossTabData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsCrossTabData = Nothing
    BuildDataArrays = True

    Exit Function

LocalErr:
    mstrStatusMessage = "Error processing data"
    BuildDataArrays = False

  End Function

  Private Function GetGroupNumber(ByRef strValue As String, ByRef Index As Short) As Object

    'This returns which array element a particular value should be added to
    'Examples:
    '
    ' value = null, Minimum = 16, Maximum = 70, Step = 5
    '    GetGroupNumber = 0
    '
    ' value = 11, Minimum = 16, Maximum = 70, Step = 5
    '    GetGroupNumber = 1
    '
    ' value = 18, Minimum = 16, Maximum = 70, Step = 5
    '    GetGroupNumber = 2
    '
    ' value = 26, Minimum = 16, Maximum = 70, Step = 5
    '    GetGroupNumber = 4
    '
    ' value = 92, Minimum = 16, Maximum = 70, Step = 5
    '    GetGroupNumber = 13

    On Error GoTo LocalErr

    Dim dblValue As Double
    Dim lngCount As Integer
    Dim dblLoop As Double
    Dim rsTemp As ADODB.Recordset

    'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    GetGroupNumber = 0
    'GetGroupNumber = IIf(strValue = vbNullString, 0, -1)

    'Non range column
    If mdblMin(Index) = 0 And mdblMax(Index) = 0 Then

      For lngCount = 0 To UBound(mvarHeadings(Index))

        Select Case mlngColDataType(Index)
          Case CStr(Declarations.SQLDataType.sqlDate)

            'JDM - 22/10/2003 - Fault 7246 - Below fix seems to gone wrong...
            'MH20020212 Fault 4893
            'If mvarHeadings(Index)(lngCount) = Format(strValue, DateFormat) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(Index)(lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If mvarHeadings(Index)(lngCount) = VB6.Format(strValue, mstrClientDateFormat) Then
              'If mvarHeadings(Index)(lngCount) = strValue Then
              'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              GetGroupNumber = lngCount
              Exit For
            End If

          Case CStr(Declarations.SQLDataType.sqlNumeric), CStr(Declarations.SQLDataType.sqlInteger)
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If UCase(mvarHeadings(Index)(lngCount)) = datGeneral.ConvertNumberForDisplay(VB6.Format(strValue, mstrFormat(Index))) Then
              'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              GetGroupNumber = lngCount
              Exit For
            End If

          Case CStr(Declarations.SQLDataType.sqlBoolean)
            Select Case LCase(strValue)
              Case ""
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(Index)(lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If mvarHeadings(Index)(lngCount) = "<Empty>" Then
                  'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  GetGroupNumber = lngCount
                  Exit For
                End If
              Case "false", "0"
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(Index)(lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If mvarHeadings(Index)(lngCount) = "False" Then
                  'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  GetGroupNumber = lngCount
                  Exit For
                End If
              Case Else
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(Index)(lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If mvarHeadings(Index)(lngCount) = "True" Then
                  'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  GetGroupNumber = lngCount
                  Exit For
                End If
            End Select

          Case Else

            'MH20021018 Faults 4532 & 4533
            'If LCase(mvarHeadings(Index)(lngCount)) = LCase(Trim(strValue)) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If LCase(mvarHeadings(Index)(lngCount)) = LCase(FormatString(strValue)) Then
              'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              GetGroupNumber = lngCount
              Exit For
            End If

        End Select

      Next

    Else 'Numeric ranges...

      dblValue = Val(strValue)
      If strValue = vbNullString Then
        'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetGroupNumber = 0
        Exit Function
      ElseIf dblValue < mdblMin(Index) Then
        'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetGroupNumber = 1
        Exit Function
      End If

      lngCount = 1
      For dblLoop = mdblMin(Index) To mdblMax(Index) Step mdblStep(Index)
        lngCount = lngCount + 1
        'If dblValue >= dblLoop And dblValue <= dblLoop + mdblStep(Index) Then
        If dblValue >= dblLoop And dblValue < dblLoop + mdblStep(Index) Then
          'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          GetGroupNumber = lngCount
          Exit Function
        End If
      Next
      'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      GetGroupNumber = lngCount + 1

    End If


    Exit Function

LocalErr:
    mstrStatusMessage = "Error grouping data <" & strValue & ">"
    fOK = False

  End Function

  Public Function HTMLText(ByRef strInput As String) As String

    HTMLText = Replace(strInput, "<", "&LT;")
    HTMLText = Replace(HTMLText, ">", "&GT;")
    HTMLText = Replace(HTMLText, vbTab, "</TD><TD>")
    HTMLText = Replace(HTMLText, "<TD></TD>", "<TD>&nbsp;</TD>")
    If Left(HTMLText, 5) = "</TD>" Then
      HTMLText = "&nbsp;" & HTMLText
    End If
    If Right(HTMLText, 4) = "<TD>" Then
      HTMLText = HTMLText & "&nbsp;"
    End If

  End Function

  Public Sub BuildOutputStrings(ByRef lngSinglePage As Integer)

    Const strDelim As String = vbTab
    Dim strTempDelim As String

    Dim lngNumCols As Integer
    Dim lngNumRows As Integer
    Dim lngNumPages As Integer

    Dim lngCol As Integer
    Dim lngRow As Integer
    Dim lngPage As Integer
    Dim lngPointer As Integer
    Dim lngTYPE As Integer

    Dim sngAverage As Single
    Dim iAverageColumn As Short

    On Error GoTo LocalErr

    lngNumCols = UBound(mvarHeadings(HOR))
    lngNumRows = UBound(mvarHeadings(VER))
    lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(PGB)), 0)
    iAverageColumn = lngNumCols - 1

    ' JDM - 22/06/01 - Fault 2476 - Display totals instead
    If mlngCrossTabType <> Declarations.CrossTabType.cttAbsenceBreakdown Then
      lngTYPE = mlngType 'cboType.ItemData(cboType.ListIndex)
    Else
      lngTYPE = TYPETOTAL
    End If

    'mdblPercentageFactor will be used in FORMATCELL, if required
    GetPercentageFactor(lngSinglePage, lngTYPE)

    ReDim mstrOutput(lngNumRows + 2)

    'Add First Column details (Vertical headings)
    mstrOutput(0) = strDelim & mstrOutput(0)
    For lngRow = 0 To lngNumRows
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      mstrOutput(lngRow + 1) = Trim(mvarHeadings(VER)(lngRow)) & strDelim & mstrOutput(lngRow + 1)
    Next
    mstrOutput(lngNumRows + 2) = IIf(mlngCrossTabType = Declarations.CrossTabType.cttNormal, mstrType(mlngType), "Total") & strDelim & mstrOutput(lngNumRows + 2)

    If mblnShowAllPagesTogether Then

      'Now add the main row data
      For lngPage = 0 To lngNumPages
        For lngCol = 0 To lngNumCols

          strTempDelim = IIf(lngCol < lngNumCols Or lngPage < lngNumPages, strDelim, "")

          'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mstrOutput(0) = mstrOutput(0) & Trim(mvarHeadings(0)(lngCol)) & strTempDelim


          For lngRow = 0 To lngNumRows
            mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & FormatCell(mdblDataArray(lngCol, lngRow, lngPage, lngTYPE), lngCol) & strTempDelim
          Next

          mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & FormatCell(mdblHorTotal(lngCol, lngPage, lngTYPE), lngCol) & strTempDelim

        Next
      Next


      If mblnPageBreak Then
        For lngCol = 0 To lngNumCols
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mstrOutput(0) = mstrOutput(0) & strDelim & Trim(mvarHeadings(0)(lngCol))

          For lngRow = 0 To lngNumRows + 1
            mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & strDelim & FormatCell(mdblPgbTotal(lngCol, lngRow, lngTYPE), lngCol)
          Next
        Next
      End If

    Else
      'Now add the main row data
      For lngCol = 0 To lngNumCols
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mstrOutput(0) = mstrOutput(0) & Trim(mvarHeadings(0)(lngCol)) & IIf(lngCol <> lngNumCols, strDelim, "")
        For lngRow = 0 To lngNumRows
          mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & FormatCell(mdblDataArray(lngCol, lngRow, lngSinglePage, lngTYPE)) & IIf(lngCol <> lngNumCols, strDelim, "")
        Next

        ' JDM - 10/09/2003 - Fault 7048 - Make the average column not total up.
        If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown And lngCol = iAverageColumn Then
          sngAverage = mdblHorTotal(lngCol - 1, lngSinglePage, TYPETOTAL) / mdblHorTotal(lngCol, lngSinglePage, TYPECOUNT)
          mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & FormatCell(sngAverage) & IIf(lngCol <> lngNumCols, strDelim, "")
        Else
          mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & FormatCell(mdblHorTotal(lngCol, lngSinglePage, lngTYPE)) & IIf(lngCol <> lngNumCols, strDelim, "")
        End If

      Next

      'Add the last column details (Vertical totals)
      'If mlngCrossTabType <> cttAbsenceBreakdown Then
      If mlngCrossTabType = Declarations.CrossTabType.cttNormal Then
        mstrOutput(0) = mstrOutput(0) & strDelim & mstrType(mlngType)
        For lngRow = 0 To lngNumRows
          mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & strDelim & FormatCell(mdblVerTotal(lngRow, lngSinglePage, lngTYPE))
        Next
        mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & strDelim & FormatCell(mdblPageTotal(lngSinglePage, lngTYPE))
      End If
    End If

    Exit Sub

LocalErr:
    mstrStatusMessage = "Error building output strings (" & Err.Description & ")"
    'Resume 0
    fOK = False

  End Sub

  Private Function FormatCell(ByVal dblCellValue As Double, Optional ByRef lngHOR As Integer = 0) As String

    Dim strMask As String

    On Error GoTo LocalErr

    strMask = vbNullString
    FormatCell = vbNullString


    If dblCellValue <> 0 Or mblnSuppressZeros = False Then


      If mlngCrossTabType <> Declarations.CrossTabType.cttNormal Then

        ' 1000 seperators
        'strMask = IIf(mbUse1000Separator, "#,0", "#0")

        If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown Then
          'strMask = String$(20, "#") & "0.0"
          strMask = IIf(mbUse1000Separator, "#,", "#") & "0.00"
        Else
          'strMask = String$(20, "#") & "0"
          strMask = IIf(mbUse1000Separator, "#,", "#") & "0"

          If lngHOR = 2 Then
            'strMask = String$(20, "#") & "0" & UI.GetSystemDecimalSeparator & "00%"
            strMask = New String("#", 20) & "0.00%"
          ElseIf lngHOR = 0 And mlngCrossTabType = Declarations.CrossTabType.cttTurnover Then
            'strMask = String$(20, "#") & "0" & UI.GetSystemDecimalSeparator & "0"
            strMask = New String("#", 20) & "0.0"
          End If
        End If

      Else

        ' 1000 seperators
        strMask = IIf(mbUse1000Separator, "#,0", "#0")

        If mblnShowPercentage Then
          'If percentage
          dblCellValue = dblCellValue * mdblPercentageFactor
          'strMask = strMask & UI.GetSystemDecimalSeparator & "00%"
          strMask = strMask & ".00%"

        ElseIf mlngType > 0 Then
          'if not count then
          'value should be displayed as per field definition
          'strMask = mstrIntersectionMask

          If mlngIntersectionDecimals > 0 Then
            strMask = strMask & "." & New String("0", mlngIntersectionDecimals)
          End If

        End If

      End If

      If strMask <> vbNullString Then
        FormatCell = VB6.Format(dblCellValue, strMask)
      End If

    End If


    Exit Function

LocalErr:
    mstrStatusMessage = "Error formatting data"
    fOK = False

  End Function

  Private Sub GetPercentageFactor(ByRef lngPage As Integer, ByRef mlngType As Integer)

    'mdblPercentageFactor will be used in FORMATCELL, if required
    mdblPercentageFactor = 0
    If mblnShowPercentage Then
      If mblnPercentageofPage Then
        If mdblPageTotal(lngPage, mlngType) > 0 Then
          mdblPercentageFactor = 1 / mdblPageTotal(lngPage, mlngType)
        End If
      Else
        If mdblGrandTotal(mlngType) > 0 Then
          mdblPercentageFactor = 1 / mdblGrandTotal(mlngType)
        End If
      End If
    End If

  End Sub

  Public Sub BuildBreakdownStrings(ByRef lngHOR As Integer, ByRef lngVER As Integer, ByRef lngPGB As Integer)

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim objColumnPrivileges As CColumnPrivileges
    Dim strOutput As String

    Dim strColumnName() As String
    Dim strWhere As String
    Dim lngCount As Integer


    On Error GoTo LocalErr

    'BuildBreakdownStrings = False

    strWhere = vbNullString
    If lngHOR <= UBound(mvarSearches(HOR)) Then
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      strWhere = IIf(strWhere = vbNullString, " WHERE ", strWhere & " AND ") & "(" & mvarSearches(HOR)(lngHOR) & ")"
    End If

    If lngVER <= UBound(mvarSearches(VER)) Then
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      strWhere = IIf(strWhere = vbNullString, " WHERE ", strWhere & " AND ") & "(" & mvarSearches(VER)(lngVER) & ")"
    End If

    If mblnPageBreak Then
      If lngPGB <= UBound(mvarSearches(PGB)) Then
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strWhere = IIf(strWhere = vbNullString, " WHERE ", strWhere & " AND ") & "(" & mvarSearches(PGB)(lngPGB) & ")"
      End If
    End If


    strSQL = "SELECT * FROM " & mstrTempTableName & strWhere & " ORDER BY "

    Select Case mlngType
      Case TYPEMINIMUM : strSQL = strSQL & "Ins, "
      Case TYPEMAXIMUM : strSQL = strSQL & "Ins DESC, "
    End Select
    strSQL = strSQL & "RecDesc"

    rsTemp = mclsData.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)


    ReDim mstrOutput(0)
    lngCount = 0

    With rsTemp
      If Not .EOF Then .MoveFirst()
      Do While Not .EOF

        strOutput = .Fields("RecDesc").Value

        ' Build output string for absence breakdown
        If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown Then

          strOutput = strOutput & vbTab

          ' Add absence start date
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          If IsDbNull(.Fields("Start_Date").Value) Then
            strOutput = strOutput & vbTab
          Else
            strOutput = strOutput & VB6.Format(.Fields("Start_Date").Value, mstrClientDateFormat) & vbTab
          End If

          ' Add absence end date
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          If IsDbNull(.Fields("End_Date").Value) Then
            strOutput = strOutput & vbTab
          Else
            strOutput = strOutput & VB6.Format(.Fields("End_Date").Value, mstrClientDateFormat) & vbTab
          End If

          ' Add occurences
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          If IsDbNull(.Fields("Value").Value) Then
            strOutput = strOutput & vbTab
          Else
            'MH20040128 Fault 7995 - Round average to 2 decimal places
            'strOutput = strOutput & .Fields("Value").Value & vbTab
            strOutput = strOutput & VB6.Format(.Fields("Value").Value, "0.00") & vbTab
          End If

        End If

        If mblnIntersection Then
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          If Not IsDbNull(.Fields("Ins").Value) Then
            strOutput = strOutput & vbTab & VB6.Format(.Fields("Ins").Value, mstrIntersectionMask)
          End If
        End If

        lngCount = lngCount + 1
        ReDim Preserve mstrOutput(lngCount)
        mstrOutput(lngCount) = strOutput

        .MoveNext()
      Loop
    End With

    'BuildBreakdownStrings = True

    Exit Sub

LocalErr:
    mstrStatusMessage = "Error reading breakdown"
    'PopulateCellBreakdown2 = False

  End Sub

  Public Sub CreateTablesCollection()

    ' JPD20030313 Do not drop the tables & columns collections as they can be reused.
    'Set gcoTablePrivileges = Nothing
    'Set gcolColumnPrivilegesCollection = Nothing

    SetupTablesCollection()

  End Sub

  Public Function AbsenceBreakdownRunStoredProcedure() As Boolean

    ' Purpose : To re-jig the selected records from the normal cross tab so it can be used in the standard
    '           crosstab output.

    On Error GoTo LocalErr

    Dim datData As clsDataAccess
    Dim sSQL As String
    Dim iAffectedRecords As Short
    Dim lngAffectedRecords As Integer

    datData = New clsDataAccess

    'Fire off the stored procedure to turn the current data into something the crosstab code will like.
    sSQL = "EXECUTE sp_ASR_AbsenceBreakdown_Run '" & mstrReportStartDate & "','" & mstrReportEndDate & "','" & mstrTempTableName & "'"
    datData.ExecuteSql((sSQL))

    ' Check that records exist (in case there's no working pattern and such like)
    Dim rsCrossTabData As ADODB.Recordset
    rsCrossTabData = New ADODB.Recordset
    rsCrossTabData.let_ActiveConnection(gADOCon)
    rsCrossTabData.Open("Select * From " & mstrTempTableName, , ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)

    If rsCrossTabData.EOF Then
      mstrStatusMessage = "No records meet selection criteria"
      mblnNoRecords = True
      fOK = False
    End If

    ' Fault 2422 - Switch days into language of client machine (server independant)
    ' JDM - 19/06/01 - Fault 2472 - Whoops, missed out some error checking...
    If fOK Then
      With rsCrossTabData
        .MoveFirst()
        Do While Not .EOF

          If .Fields("Day_Number").Value < 8 Then
            .Fields("HOR").Value = WeekDayName(.Fields("Day_Number").Value, False, FirstDayOfWeek.Monday)
          End If

          .MoveNext()
        Loop

      End With
    End If

    AbsenceBreakdownRunStoredProcedure = True
    Exit Function

LocalErr:
    mstrStatusMessage = "Error running stored procedure in database"
    AbsenceBreakdownRunStoredProcedure = False

  End Function

  Public Sub AbsenceBreakdownSizeColumns()

    Dim lngWidth As Integer
    Dim intCount As Short


    '  With frmCrossTabRun.SSDBGrid1
    '
    '    lngWidth = (.Width - .Columns(1).Width - 5) / (.Columns.Count - 1)
    '
    '    For intCount = 1 To .Cols - 1
    '      .Columns(intCount).Width = lngWidth - 2
    '    Next intCount
    '
    '  End With

  End Sub

  Public Function AbsenceBreakdownBuildDataArrays() As Boolean

    Dim strTempValue As String
    Dim dblThisIntersectionVal As Double

    Dim lngCol As Integer
    Dim lngRow As Integer
    Dim lngPage As Integer
    Dim lngNumCols As Integer
    Dim lngNumRows As Integer
    Dim lngNumPages As Integer


    On Error GoTo LocalErr

    lngNumCols = UBound(mvarHeadings(0))
    lngNumRows = UBound(mvarHeadings(1))
    lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(2)), 0)

    ReDim mdblDataArray(lngNumCols, lngNumRows, lngNumPages, 4)
    ReDim mdblHorTotal(lngNumCols, lngNumPages, 4)
    ReDim mdblVerTotal(lngNumRows, lngNumPages, 4)
    ReDim mdblPgbTotal(lngNumCols, lngNumRows + 1, 4) '+1 for totals !
    ReDim mdblPageTotal(lngNumPages, 4)
    ReDim mdblGrandTotal(4)

    ' Because the stored procedure has run we need to requery the recordset
    rsCrossTabData.Requery()

    With rsCrossTabData

      If .BOF And .EOF Then
        AbsenceBreakdownBuildDataArrays = False
        Exit Function
      End If

      .MoveFirst()
      Do While Not .EOF

        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        strTempValue = IIf(Not IsDbNull(.Fields("HOR").Value), .Fields("HOR").Value, vbNullString)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lngCol = GetGroupNumber(strTempValue, HOR)

        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        strTempValue = IIf(Not IsDbNull(.Fields("VER").Value), .Fields("VER").Value, vbNullString)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lngRow = GetGroupNumber(strTempValue, VER)

        'Count
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        mdblDataArray(lngCol, lngRow, 0, TYPECOUNT) = mdblDataArray(lngCol, lngRow, 0, TYPECOUNT) + IIf(Not IsDbNull(.Fields("VALUE").Value), .Fields("VALUE").Value, 143)
        mdblHorTotal(lngCol, 0, TYPECOUNT) = mdblHorTotal(lngCol, 0, TYPECOUNT) + 1
        mdblVerTotal(lngRow, 0, TYPECOUNT) = mdblVerTotal(lngRow, 0, TYPECOUNT) + 1

        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        mdblDataArray(lngCol, lngRow, 0, TYPETOTAL) = mdblDataArray(lngCol, lngRow, 0, TYPETOTAL) + IIf(Not IsDbNull(.Fields("VALUE").Value), .Fields("VALUE").Value, 143)
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        mdblHorTotal(lngCol, 0, TYPETOTAL) = mdblHorTotal(lngCol, 0, TYPETOTAL) + IIf(Not IsDbNull(.Fields("VALUE").Value), .Fields("VALUE").Value, 0)
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        mdblVerTotal(lngRow, 0, TYPETOTAL) = mdblVerTotal(lngRow, 0, TYPETOTAL) + IIf(Not IsDbNull(.Fields("VALUE").Value), .Fields("VALUE").Value, 0)

        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        mdblDataArray(lngCol, lngRow, lngPage, TYPEAVERAGE) = mdblDataArray(lngCol, lngRow, lngPage, TYPEAVERAGE) + IIf(Not IsDbNull(.Fields("VALUE").Value), .Fields("VALUE").Value, 143)
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        mdblHorTotal(lngCol, lngPage, TYPEAVERAGE) = mdblHorTotal(lngCol, lngPage, TYPEAVERAGE) + IIf(Not IsDbNull(.Fields("VALUE").Value), .Fields("VALUE").Value, 0)
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        mdblVerTotal(lngRow, lngPage, TYPEAVERAGE) = mdblVerTotal(lngRow, lngPage, TYPEAVERAGE) + IIf(Not IsDbNull(.Fields("VALUE").Value), .Fields("VALUE").Value, 0)

        .MoveNext()
      Loop

    End With

    'UPGRADE_NOTE: Object rsCrossTabData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsCrossTabData = Nothing

    AbsenceBreakdownBuildDataArrays = True
    Exit Function

LocalErr:
    mstrStatusMessage = "Error processing data"
    AbsenceBreakdownBuildDataArrays = False

  End Function

  Public Function AbsenceBreakdownGetHeadingsAndSearches() As Boolean

    Dim strHeading() As String
    Dim strSearch() As String
    Dim lngLoop As Integer


    On Error GoTo LocalErr

    For lngLoop = 0 To 2

      ReDim strHeading(0)
      ReDim strSearch(0)

      If lngLoop = 2 And mblnPageBreak = False Then
        'When no page break field is specified
        strHeading(0) = "<None>"
      Else
        GetHeadingsAndSearchesForColumns(lngLoop, strHeading, strSearch)
      End If


      'Store each array in an array of variants (an array in an array!)
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings(lngLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      mvarHeadings(lngLoop) = VB6.CopyArray(strHeading)
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarSearches(lngLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      mvarSearches(lngLoop) = VB6.CopyArray(strSearch)

    Next

    AbsenceBreakdownGetHeadingsAndSearches = True
    Exit Function

LocalErr:
    mstrStatusMessage = "Error building headings and search arrays"
    AbsenceBreakdownGetHeadingsAndSearches = False

  End Function

  Public Function AbsenceBreakdownRetreiveDefinition(ByRef pdtStartDate As Object, ByRef pdtEndDate As Object, ByRef plngHorColID As Object, ByRef plngVerColID As Object, ByRef plngPicklistID As Object, ByRef plngFilterID As Object, ByRef plngPersonnelID As Object, ByRef pstrIncludedTypes As Object) As Boolean

    Dim iCount As Short
    Dim lngHorColID As Integer
    Dim lngVerColID As Integer
    Dim lngPicklistID As Integer
    Dim lngFilterID As Integer
    Dim lngPersonnelID As Integer
    Dim dtStartDate As Date
    Dim dtEndDate As Date
    Dim strIncludedTypes As String

    ReDim mastrUDFsRequired(0)

    ' Read the module parameters
    ReadPersonnelParameters()
    ReadAbsenceParameters()

    ' Define this cross tab as an absence breakdown
    mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown

    ' Initialse the ok variable
    fOK = True

    ' Convert variants into correct types
    lngHorColID = IIf(IsNumeric(plngHorColID), plngHorColID, 0)
    lngVerColID = IIf(IsNumeric(plngVerColID), plngVerColID, 0)
    lngPicklistID = IIf(IsNumeric(plngPicklistID), plngPicklistID, 0)
    lngFilterID = IIf(IsNumeric(plngFilterID), plngFilterID, 0)
    lngPersonnelID = IIf(IsNumeric(plngPersonnelID), plngPersonnelID, 0)

    ' Force the inputted string to be formatted correctly
    'UPGRADE_WARNING: Couldn't resolve default property of object pstrIncludedTypes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    pstrIncludedTypes = Trim(pstrIncludedTypes)
    'UPGRADE_WARNING: Couldn't resolve default property of object pstrIncludedTypes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    strIncludedTypes = Replace(pstrIncludedTypes, "'", "''")
    strIncludedTypes = "'" & Replace(strIncludedTypes, ",", "','")
    strIncludedTypes = Mid(strIncludedTypes, 1, Len(strIncludedTypes) - 2)

    mstrCrossTabName = "Absence Breakdown"

    ' Dates coming in are always in SQL (American) format
    'UPGRADE_WARNING: Couldn't resolve default property of object pdtStartDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mstrReportStartDate = pdtStartDate
    'UPGRADE_WARNING: Couldn't resolve default property of object pdtEndDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mstrReportEndDate = pdtEndDate

    'JPD 20041214 - ensure no injection can take place.
    mstrReportStartDate = Replace(mstrReportStartDate, "'", "''")
    mstrReportEndDate = Replace(mstrReportEndDate, "'", "''")

    'MH20040129 Fault 7857
    'UPGRADE_WARNING: Couldn't resolve default property of object pdtEndDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    'UPGRADE_WARNING: Couldn't resolve default property of object pdtStartDate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
    If DateDiff(Microsoft.VisualBasic.DateInterval.Day, datGeneral.ConvertSQLDateToSystemFormat(CStr(pdtStartDate)), datGeneral.ConvertSQLDateToSystemFormat(CStr(pdtEndDate))) < 0 Then
      mstrStatusMessage = "The report end date is before the report start date."
      fOK = False
      Exit Function
    End If


    mlngBaseTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE))
    mstrBaseTable = datGeneral.GetTableName(mlngBaseTableID)

    mlngRecordDescExprID = CInt(datGeneral.GetRecDescExprID(mlngBaseTableID))

    ' Add the different reason types
    msAbsenceBreakdownTypes = "(" & IIf(Len(strIncludedTypes) = 0, "''", strIncludedTypes) & ")"

    ' Load the appropraite records
    If lngPersonnelID > 0 Then
      mstrPicklistFilter = CStr(lngPersonnelID)
    Else
      mstrPicklistFilter = GetPicklistFilterSelect(lngPicklistID, lngFilterID)
    End If

    If fOK = False Then
      Exit Function
    End If

    mlngColID(HOR) = lngHorColID
    mstrColName(HOR) = datGeneral.GetColumnName(lngHorColID)
    mlngColDataType(HOR) = CStr(datGeneral.GetDataType(mlngBaseTableID, lngHorColID))
    mstrFormat(HOR) = GetFormat(mlngColID(HOR))

    mlngColID(VER) = lngVerColID
    mstrColName(VER) = datGeneral.GetColumnName(lngVerColID)
    mlngColDataType(VER) = CStr(datGeneral.GetDataType(mlngBaseTableID, lngVerColID))
    mstrFormat(VER) = GetFormat(mlngColID(VER))

    mlngIntersectionDecimals = 2
    mblnIntersection = False
    mblnShowAllPagesTogether = False

    AbsenceBreakdownRetreiveDefinition = True

  End Function

  Public Function SetAbsenceBreakDownDisplayOptions(ByRef pbShowBasePicklistFilter As Object) As Boolean

    ' Set Report Display Options
    'UPGRADE_WARNING: Couldn't resolve default property of object pbShowBasePicklistFilter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mblnChkPicklistFilter = pbShowBasePicklistFilter
    SetAbsenceBreakDownDisplayOptions = True

  End Function


  Private Function ConvertSQLDateToLocale(ByRef psSQLDate As String) As String
    ' Convert the given date string (mm/dd/yyyy) into the locale format.
    ' NB. This function assumes a sensible locale format is used.
    Dim fDaysDone As Boolean
    Dim fMonthsDone As Boolean
    Dim fYearsDone As Boolean
    Dim iLoop As Short
    Dim sFormattedDate As String

    sFormattedDate = ""

    ' Get the locale's date format.
    fDaysDone = False
    fMonthsDone = False
    fYearsDone = False

    For iLoop = 1 To Len(mstrClientDateFormat)
      Select Case UCase(Mid(mstrClientDateFormat, iLoop, 1))
        Case "D"
          If Not fDaysDone Then
            sFormattedDate = sFormattedDate & Mid(psSQLDate, 4, 2)
            fDaysDone = True
          End If

        Case "M"
          If Not fMonthsDone Then
            sFormattedDate = sFormattedDate & Mid(psSQLDate, 1, 2)
            fMonthsDone = True
          End If

        Case "Y"
          If Not fYearsDone Then
            sFormattedDate = sFormattedDate & Mid(psSQLDate, 7, 4)
            fYearsDone = True
          End If

        Case Else
          sFormattedDate = sFormattedDate & Mid(mstrClientDateFormat, iLoop, 1)
      End Select
    Next iLoop

    ConvertSQLDateToLocale = sFormattedDate

  End Function

  ' Function which we use to pass in the default output parameters (Standard reports read from the defintion table,
  '    which don't exist for standard reports)
  Public Function SetAbsenceBreakDownDefaultOutputOptions(ByRef pbOutputPreview As Object, ByRef plngOutputFormat As Object, ByRef pblnOutputScreen As Object, ByRef pblnOutputPrinter As Object, ByRef pstrOutputPrinterName As Object, ByRef pblnOutputSave As Object, ByRef plngOutputSaveExisting As Object, ByRef pblnOutputEmail As Object, ByRef plngOutputEmailID As Object, ByRef pstrOutputEmailName As Object, ByRef pstrOutputEmailSubject As Object, ByRef pstrOutputEmailAttachAs As Object, ByRef pstrOutputFilename As Object) As Boolean

    'UPGRADE_WARNING: Couldn't resolve default property of object pbOutputPreview. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mblnOutputPreview = pbOutputPreview
    'UPGRADE_WARNING: Couldn't resolve default property of object plngOutputFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mlngOutputFormat = plngOutputFormat
    'UPGRADE_WARNING: Couldn't resolve default property of object pblnOutputScreen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mblnOutputScreen = pblnOutputScreen
    'UPGRADE_WARNING: Couldn't resolve default property of object pblnOutputPrinter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mblnOutputPrinter = pblnOutputPrinter
    'UPGRADE_WARNING: Couldn't resolve default property of object pstrOutputPrinterName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mstrOutputPrinterName = pstrOutputPrinterName
    'UPGRADE_WARNING: Couldn't resolve default property of object pblnOutputSave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mblnOutputSave = pblnOutputSave
    'UPGRADE_WARNING: Couldn't resolve default property of object plngOutputSaveExisting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mlngOutputSaveExisting = plngOutputSaveExisting
    'UPGRADE_WARNING: Couldn't resolve default property of object pblnOutputEmail. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mblnOutputEmail = pblnOutputEmail
    'UPGRADE_WARNING: Couldn't resolve default property of object plngOutputEmailID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mlngOutputEmailID = IIf(plngOutputEmailID = "", 0, plngOutputEmailID)
    mstrOutputEmailName = GetEmailGroupName(mlngOutputEmailID)
    'UPGRADE_WARNING: Couldn't resolve default property of object pstrOutputEmailSubject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mstrOutputEmailSubject = pstrOutputEmailSubject
    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
    mstrOutputEmailAttachAs = IIf(IsDbNull(pstrOutputEmailAttachAs), vbNullString, pstrOutputEmailAttachAs)
    'UPGRADE_WARNING: Couldn't resolve default property of object pstrOutputFilename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mstrOutputFilename = pstrOutputFilename

    ' JDM - 17/11/03 - Fault 7564 - Going to printer automatically
    mblnOutputPreview = (pbOutputPreview Or (mlngOutputFormat = Declarations.OutputFormats.fmtDataOnly And mblnOutputScreen))

  End Function

  Public Function UDFFunctions(ByRef pbCreate As Boolean) As Boolean

    On Error GoTo UDFFunctions_ERROR

    Dim iCount As Short
    Dim strDropCode As String
    Dim strFunctionName As String
    Dim sUDFCode As String
    Dim datData As clsDataAccess
    Dim iStart As Short
    Dim iEnd As Short
    Dim strFunctionNumber As String

    Const FUNCTIONPREFIX As String = "udf_ASRSys_"

    If gbEnableUDFFunctions Then

      For iCount = 1 To UBound(mastrUDFsRequired)

        'JPD 20060110 Fault 10509
        'iStart = Len("CREATE FUNCTION udf_ASRSys_") + 1
        iStart = InStr(mastrUDFsRequired(iCount), FUNCTIONPREFIX) + Len(FUNCTIONPREFIX)
        iEnd = InStr(1, Mid(mastrUDFsRequired(iCount), 1, 1000), "(@Per")
        strFunctionNumber = Mid(mastrUDFsRequired(iCount), iStart, iEnd - iStart)
        strFunctionName = FUNCTIONPREFIX & strFunctionNumber

        'Drop existing function (could exist if the expression is used more than once in a report)
        strDropCode = "IF EXISTS" & " (SELECT *" & "   FROM sysobjects" & "   WHERE id = object_id('[" & Replace(gsUsername, "'", "''") & "]." & strFunctionName & "')" & "     AND sysstat & 0xf = 0)" & " DROP FUNCTION [" & gsUsername & "]." & strFunctionName
        mclsData.ExecuteSql(strDropCode)

        ' Create the new function
        'UPGRADE_WARNING: Couldn't resolve default property of object pbCreate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If pbCreate Then
          sUDFCode = mastrUDFsRequired(iCount)
          mclsData.ExecuteSql(sUDFCode)
        End If

      Next iCount
    End If

    'UPGRADE_WARNING: Couldn't resolve default property of object UDFFunctions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    UDFFunctions = True
    Exit Function

UDFFunctions_ERROR:
    'UPGRADE_WARNING: Couldn't resolve default property of object UDFFunctions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    UDFFunctions = False

  End Function


  Public Sub GetPivotRecordset()

    Dim rsPivot As ADODB.Recordset
    Dim strSQL As String

    Dim strOutput(,) As String
    Dim strPageValue As String
    Dim lngGroupNum As Integer
    Dim lngCol As Integer
    Dim lngRow As Integer


    strSQL = "SELECT HOR as 'Horizontal', VER as 'Vertical'" & IIf(mblnPageBreak, ", PGB as 'Page Break'", vbNullString) & ", RecDesc as 'Record Description'" & IIf(mblnIntersection, ", Ins as 'Intersection'", vbNullString) & IIf(mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown, ", Value as 'Duration'", vbNullString) & " FROM " & mstrTempTableName

    If mlngCrossTabType = Declarations.CrossTabType.cttAbsenceBreakdown Then
      strSQL = strSQL & " WHERE NOT HOR IN ('Total','Count','Average')"

    ElseIf mblnPageBreak Then
      'If objOutput.cboPageBreak.ListIndex > 0 Then
      '  strSQL = strSQL & " WHERE " & mvarSearches(PGB)(objOutput.cboPageBreak.ListIndex - 1)
      'End If
      strSQL = strSQL & " ORDER BY PGB"
    End If

    rsPivot = datGeneral.GetReadOnlyRecords(strSQL)

    '------------

    With rsPivot

      ReDim mstrOutputPivotArray(0)

      If Not mblnPageBreak Then
        lngRow = 1
        ReDim strOutput(.Fields.Count - 1, 0)
        For lngCol = 0 To .Fields.Count - 1
          strOutput(lngCol, 0) = .Fields(lngCol).Name
        Next
      End If

      .MoveFirst()
      Do While Not .EOF

        If mblnPageBreak Then
          If strPageValue <> .Fields("Page Break").Value Then

            If strPageValue <> vbNullString Then

              '              objOutput.AddPage Replace(SSDBGrid1.Caption, "&&", "&"), strPageValue
              '              objOutput.DataArray strOutput

              PivotAddToArray("  window.parent.parent.ASRIntranetOutput.AddPage(""" & Replace(Me.CrossTabName, """", "\""") & """, """ & Left(Replace(strPageValue, """", "\"""), 255) & """);")
              PivotAddToArray("  window.parent.parent.ASRIntranetOutput.ArrayDim(" & CStr(UBound(strOutput, 1)) & ", " & CStr(UBound(strOutput, 2)) & ");")
              For lngCol = 0 To UBound(strOutput, 1)
                For lngRow = 0 To UBound(strOutput, 2)
                  PivotAddToArray("  window.parent.parent.ASRIntranetOutput.ArrayAddTo(" & CStr(lngCol) & ", " & CStr(lngRow) & ", """ & Left(Replace(strOutput(lngCol, lngRow), """", "\"""), 255) & """);")
                Next
              Next

              PivotAddToArray("  window.parent.parent.ASRIntranetOutput.DataArray();")

            End If
            strPageValue = .Fields("Page Break").Value

            lngRow = 1
            ReDim strOutput(.Fields.Count - 1, 0)
            For lngCol = 0 To .Fields.Count - 1
              strOutput(lngCol, 0) = .Fields(lngCol).Name
            Next

          End If
        Else
          strPageValue = mstrBaseTable

        End If

        ReDim Preserve strOutput(.Fields.Count - 1, lngRow)
        For lngCol = 0 To .Fields.Count - 1

          'MH20070226 Fault 11962
          'If lngCol <= UBound(mvarHeadings) Then
          If lngCol < 2 Or (lngCol = 2 And mblnPageBreak) Then

            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            lngGroupNum = GetGroupNumber(CStr(IIf(IsDbNull(.Fields(lngCol).Value), vbNullString, .Fields(lngCol).Value)), CShort(lngCol))
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strOutput(lngCol, lngRow) = mvarHeadings(lngCol)(lngGroupNum)
          Else
            strOutput(lngCol, lngRow) = .Fields(lngCol).Value
          End If
        Next
        lngRow = lngRow + 1
        .MoveNext()
      Loop
    End With

    '    objOutput.AddPage Replace(SSDBGrid1.Caption, "&&", "&"), IIf(strPageValue <> vbNullString, strPageValue, mstrCrossTabName)
    '    objOutput.DataArray strOutput
    PivotAddToArray("  window.parent.parent.ASRIntranetOutput.AddPage(""" & Replace(Me.CrossTabName, """", "\""") & """, """ & Replace(strPageValue, """", "\""") & """);")

    PivotAddToArray("  window.parent.parent.ASRIntranetOutput.ArrayDim(" & CStr(UBound(strOutput, 1)) & ", " & CStr(UBound(strOutput, 2)) & ");")
    For lngCol = 0 To UBound(strOutput, 1)
      For lngRow = 0 To UBound(strOutput, 2)
        PivotAddToArray("  window.parent.parent.ASRIntranetOutput.ArrayAddTo(" & CStr(lngCol) & ", " & CStr(lngRow) & ", """ & Left(Replace(strOutput(lngCol, lngRow), """", "\"""), 255) & """);")
      Next
    Next

    PivotAddToArray("  window.parent.parent.ASRIntranetOutput.DataArray();")

    rsPivot.Close()
    'UPGRADE_NOTE: Object rsPivot may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsPivot = Nothing

  End Sub


  Private Function PivotAddToArray(ByRef strInput As String) As Object

    Dim lngIndex As Integer

    lngIndex = UBound(mstrOutputPivotArray) + 1
    ReDim Preserve mstrOutputPivotArray(lngIndex)
    'mstrOutputPivotArray(lngIndex) = Replace(strInput, String(1, 34), String(2, 34)) & vbNewLine
    mstrOutputPivotArray(lngIndex) = strInput & vbNewLine

  End Function


  Private Function FormatSQLColumn(ByVal sColumn As String) As String

    Dim sReturnValue As String

    'sReturnValue = Left(Trim(Replace(sHeading, Chr(32), "")), 255)
    sReturnValue = sColumn
    sReturnValue = "left(rtrim(" & sReturnValue & "), 100)"
    sReturnValue = "replace(" & sReturnValue & ",char(9),'')"
    sReturnValue = "replace(" & sReturnValue & ",char(10),'')"
    sReturnValue = "replace(" & sReturnValue & ",char(13),'')"

    FormatSQLColumn = sReturnValue

  End Function


  Private Function FormatString(ByVal sHeading As String) As String

    Dim sReturnValue As String

    'sReturnValue = Left(Trim(Replace(sHeading, Chr(32), "")), 255)
    sReturnValue = Left(Trim(sHeading), 100)
    sReturnValue = Replace(sReturnValue, Chr(9), "")
    sReturnValue = Replace(sReturnValue, Chr(10), "")
    sReturnValue = Replace(sReturnValue, Chr(13), "")

    'Intranet Only
    'sReturnValue = HTMLEncode(sReturnValue)

    FormatString = sReturnValue

  End Function
End Class