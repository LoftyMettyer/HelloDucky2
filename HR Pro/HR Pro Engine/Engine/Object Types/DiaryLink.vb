Namespace Things
  <Serializable()> _
  Public Class DiaryLink
    Inherits Things.Base

    Public Property Column As Things.Column
    Public Property Comment As String
    Public Property Offset As Integer
    Public Property OffsetType As DateOffsetType
    Public Property Reminder As Boolean
    Public Property Filter As Things.Expression
    Public Property EffectiveDate As DateTime
    Public Property CheckLeavingDate As Boolean

    Public UDF As ScriptDB.GeneratedUDF

    Public Sub Generate()

      'Dim sCode As String
      'Dim sFilterCode As String
      'Dim aryFilters As New ArrayList
      'Dim objLeavingDate As Things.Column = Nothing


      '' Add offset
      'Select Case OffsetType
      '  Case Things.DateOffsetType.Week
      '    sCode = String.Format("DATEADD(WW, {1}, DATEADD(D, 0, DATEDIFF(D, 0, [{0}])))", Column.Name, CInt(Offset))
      '  Case Things.DateOffsetType.Month
      '    sCode = String.Format("DATEADD(MM, {1}, DATEADD(D, 0, DATEDIFF(D, 0, [{0}])))", Column.Name, CInt(Offset))
      '  Case Things.DateOffsetType.Year
      '    sCode = String.Format("DATEADD(DD, {1}, DATEADD(D, 0, DATEDIFF(D, 0, [{0}])))", Column.Name, CInt(Offset))
      '  Case Else
      '    sCode = String.Format("DATEADD(DD, {1}, DATEADD(D, 0, DATEDIFF(D, 0, [{0}])))", Column.Name, CInt(Offset))
      'End Select

      '' Add date effective
      'If Not EffectiveDate.ToString Is Nothing Then
      '  aryFilters.Add(String.Format("([{0}] > convert(datetime,'{1}'))", Column.Name, Format(EffectiveDate, "yyyy-MM-dd")))
      'End If

      '' Add omit after employee leaving date
      'objLeavingDate = Globals.ModuleSetup.GetSetting("MODULE_PERSONNEL", "Param_FieldsLeavingDate").Column
      'If Not objLeavingDate Is Nothing Then
      '  aryFilters.Add(String.Format("([{0}] > GETDATE() OR [{0}] IS NULL)", objLeavingDate.Name))
      'End If


      ''if theres a filter we need a more complex UDF - switch to complex mode and pass in single paramter of the record id and generate a udf for the diary link
      '' Add filter
      'If Not Filter Is Nothing Then
      '  Filter.ExpressionType = ScriptDB.ExpressionType.DiaryFilter
      '  Filter.EmbedDependencies = False
      '  Filter.InlineScript = True
      '  Filter.GenerateCode()

      '  If Filter.IsComplex Then
      '    sFilterCode = String.Format("udfdiarylinkfilter_{0}{1}({2})", Me.Parent.Name, Filter.Name, Filter.Parameters)
      '  Else
      '    sFilterCode = "(" & Filter.UDF.Code & ") = 1"
      '  End If
      '  aryFilters.Add(sFilterCode)
      'End If

      '' Merge the filters
      'If aryFilters.Count > 0 Then
      '  sCode = String.Format("CASE WHEN {0} THEN {1} ELSE NULL END ", String.Join(" AND ", aryFilters.ToArray), sCode)
      'End If


      '' Attach to UDF object
      'With UDF
      '  .SelectCode = sCode
      '  .CallingCode = sCode
      'End With

      ' eefective date could be something like this?
      '	, CASE WHEN [Start_Date] < GETDATE() THEN DATEADD(DD, 20, DATEADD(D, 0, DATEDIFF(D, 0, [Start_Date]))) ELSE NULL END AS [17]

    End Sub


    Public Overrides ReadOnly Property Type As Things.Enums.Type
      Get
        Return Enums.Type.DiaryLink
      End Get
    End Property

  End Class
End Namespace
