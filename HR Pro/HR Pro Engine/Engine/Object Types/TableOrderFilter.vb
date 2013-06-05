Namespace Things

  <Serializable()>
  Public Class TableOrderFilter
    Inherits Base

    Public Property Table As Table
    Public Property ComponentNumber As Long
    Public UDF As ScriptDB.GeneratedUDF
    Public RowDetails As ChildRowDetails
    Public Property IncludedColumns As ICollection(Of Column)

    Public Sub New()
      IncludedColumns = New Collection(Of Column)
    End Sub

    Public Overrides Property Name As String
      Get
        Dim sName As String

        sName = String.Format("udftab_{0}", Table.Name)

        If Not RowDetails.Order Is Nothing Then
          sName = sName + "_" + String.Format("{0}({1})", RowDetails.Order.Name, RowDetails.Order.ID)
        End If

        If Not RowDetails.Filter Is Nothing Then
          sName = sName + "_" + String.Format("{0}({1})", RowDetails.Filter.Name, RowDetails.Filter.ID)
        End If

        If Not RowDetails.Relation Is Nothing Then
          sName = String.Format("{0}_{1}", sName, RowDetails.Relation.ParentID)
        End If

        Select Case RowDetails.RowSelection
          Case ScriptDB.ColumnRowSelection.First, ScriptDB.ColumnRowSelection.Last
            sName = String.Format("{0}_{1}", sName, RowDetails.RowSelection)

          Case ScriptDB.ColumnRowSelection.Specific
            sName = String.Format("{0}_line_{1}", sName, RowDetails.RowNumber)

          Case Else
            sName = String.Format("{0}_{1}", sName, RowDetails.RowSelection)

        End Select

        Return sName
      End Get
      Set(ByVal value As String)
        MyBase.Name = value
      End Set
    End Property

    Public Sub GenerateCode()

      Dim objOrderItem As TableOrderItem
      Dim objColumn As Column
      Dim aryOrderBy As New ArrayList
      Dim aryColumnList As New ArrayList
      Dim aryReturnDefintion As New ArrayList
      Dim aryParameters As New ArrayList
      Dim aryWheres As New ArrayList
      Dim aryJoins As New ArrayList
      Dim sRowSelection As String = vbNullString
      Dim bReverseOrder As Boolean
      Dim sOptions As String = ""
      Dim objIndex As New Index

      ' What type of rows to retrieve
      Select Case RowDetails.RowSelection
        Case ScriptDB.ColumnRowSelection.First
          sRowSelection = " TOP 1"
        Case ScriptDB.ColumnRowSelection.Last
          sRowSelection = " TOP 1"
          bReverseOrder = True
      End Select

      sOptions = "--WITH SCHEMABINDING"

      ' Build the where clause
      If Not RowDetails.Filter Is Nothing Then
        RowDetails.Filter.AssociatedColumn = Me.Table.Columns(0)
        RowDetails.Filter.ExpressionType = ScriptDB.ExpressionType.ColumnFilter
        RowDetails.Filter.GenerateCode()

        'If RowDetails.Filter.IsComplex Then
        '  aryWheres.Add(String.Format("({0} = 1)", RowDetails.Filter.UDF.CallingCode))
        'Else
        aryWheres.Add(String.Format("({0} = 1)", RowDetails.Filter.UDF.SelectCode))
        '    End If

        aryJoins.Add(RowDetails.Filter.UDF.JoinCode)
      End If

      ' Build the order by clause
      If Not RowDetails.Order Is Nothing And _
          Not (RowDetails.RowSelection = ScriptDB.ColumnRowSelection.Total Or RowDetails.RowSelection = ScriptDB.ColumnRowSelection.Count) Then
        For Each objOrderItem In RowDetails.Order.Items
          If objOrderItem.ColumnType = "O" And Not objOrderItem.Column Is Nothing Then

            If Not objOrderItem.Column Is Nothing And objOrderItem.Column.Table Is Me.Table Then
              Select Case objOrderItem.Ascending
                Case Order.Ascending
                  aryOrderBy.Add(String.Format("base.[{1}] {2}", objOrderItem.Column.Table.Name, objOrderItem.Column.Name, If(bReverseOrder, "DESC", "ASC")))
                Case Else
                  aryOrderBy.Add(String.Format("base.[{1}] {2}", objOrderItem.Column.Table.Name, objOrderItem.Column.Name, If(bReverseOrder, "ASC", "DESC")))
              End Select
              objIndex.Columns.AddIfNew(objOrderItem.Column)
            End If
          End If
        Next
        aryOrderBy.Add("base.[ID] ASC")

      End If

      ' Add foreign key
      If Not RowDetails.Relation Is Nothing Then
        aryParameters.Add(String.Format("@prm_ID_{0} integer", RowDetails.Relation.ParentID))
        aryWheres.Add(String.Format("[ID_{0}] = @prm_ID_{0}", RowDetails.Relation.ParentID))
        objIndex.Relations.AddIfNew(RowDetails.Relation)
      End If

      ' Add the included columns
      For Each objColumn In IncludedColumns
        Select Case RowDetails.RowSelection
          Case ScriptDB.ColumnRowSelection.Count
            aryColumnList.Add(String.Format("COUNT(ISNULL(base.[{0}],0))", objColumn.Name))
            aryReturnDefintion.Add(String.Format("[{0}] numeric(38,8)", objColumn.Name))
          Case ScriptDB.ColumnRowSelection.Total
            aryColumnList.Add(String.Format("SUM(ISNULL(base.[{0}],0))", objColumn.Name))
            aryReturnDefintion.Add(String.Format("[{0}] numeric(38,8)", objColumn.Name))
          Case Else
            aryColumnList.Add(String.Format("base.[{0}]", objColumn.Name))
            aryReturnDefintion.Add(String.Format("[{0}] {1}", objColumn.Name, objColumn.DataTypeSyntax))
            objIndex.IncludedColumns.AddIfNew(objColumn)
        End Select

      Next

      ' Create index for this object
      objIndex.Name = String.Format("IDX_{0}", Me.Name)
      objIndex.IncludePrimaryKey = False
      objIndex.IsTableIndex = True
      objIndex.IsClustered = False

         With UDF
            .Name = "[dbo].[" & Me.Name & "]"
            .WhereCode = If(aryWheres.Count > 0, "WHERE ", "") & String.Join(" AND ", aryWheres.ToArray())
            .OrderCode = If(aryOrderBy.Count > 0, "ORDER BY ", "") & String.Join(", ", aryOrderBy.ToArray())

            If RowDetails.RowSelection = ScriptDB.ColumnRowSelection.Specific Then

               .Code = String.Format("CREATE FUNCTION dbo.[{0}]({1})" & vbNewLine & _
                  "RETURNS @results TABLE({2})" & vbNewLine & _
                  "{9}" & vbNewLine & _
                  "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                  "DECLARE @forcerefresh bit;" & vbNewLine & _
                  "SET @forcerefresh = 1;" & vbNewLine & vbNewLine & _
                  "WITH base AS (" & vbNewLine & _
                  "    SELECT {3}, [rownumber] = ROW_NUMBER() OVER ({7})" & vbNewLine & _
                  "    FROM {4} base" & vbNewLine & _
                  "    {5}" & vbNewLine & _
                  "    {6})" & vbNewLine & _
                  "INSERT @Results SELECT {3}" & vbNewLine & _
                  "        FROM base" & vbNewLine & _
                  "        WHERE [rownumber] = {10}" & vbNewLine & _
                  "    RETURN;" & vbNewLine & _
                  "END" _
                 , Me.Name, String.Join(", ", aryParameters.ToArray()) _
                 , String.Join(", ", aryReturnDefintion.ToArray()), String.Join(", ", aryColumnList.ToArray()) _
                 , Me.Table.Name, String.Join(vbNewLine, aryJoins.ToArray()), .WhereCode, .OrderCode, sRowSelection _
                 , sOptions, RowDetails.RowNumber)

            Else
               .Code = String.Format("CREATE FUNCTION dbo.[{0}]({1})" & vbNewLine & _
                              "RETURNS @results TABLE({2})" & vbNewLine & _
                              "{9}" & vbNewLine & _
                              "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                              "DECLARE @forcerefresh bit;" & vbNewLine & _
                              "SET @forcerefresh = 1;" & vbNewLine & vbNewLine & _
                              "INSERT @Results SELECT{8} {3}" & vbNewLine & _
                              "        FROM dbo.[{4}] base" & vbNewLine & _
                              "        {5}" & vbNewLine & _
                              "        {6}" & vbNewLine & _
                              "        {7}" & vbNewLine & _
                              "    RETURN;" & vbNewLine & _
                              "END" _
                             , Me.Name, String.Join(", ", aryParameters.ToArray()) _
                             , String.Join(", ", aryReturnDefintion.ToArray()), String.Join(", ", aryColumnList.ToArray()) _
                             , Me.Table.Name, String.Join(vbNewLine, aryJoins.ToArray()), .WhereCode, .OrderCode, sRowSelection _
                             , sOptions)
            End If

            ' Add the index
            Me.Table.Indexes.Add(objIndex)

         End With

    End Sub


    '         bReverseOrder = False

    '' What type/line number are we dealing with?
    '        Select Case [Component].ChildRowDetails.RowSelection

    '          Case ScriptDB.ColumnRowSelection.First
    '            sPartCode = String.Format("{0}SELECT TOP 1 @part_{1} = base.[{2}]" & vbNewLine _
    '                , [CodeCluster].Indentation, iPartNumber, objThisColumn.Name)

    '          Case ScriptDB.ColumnRowSelection.Last
    '            sPartCode = String.Format("{0}SELECT TOP 1 @part_{1} = base.[{2}]" & vbNewLine _
    '                , [CodeCluster].Indentation, iPartNumber, objThisColumn.Name)
    '            bReverseOrder = True

    '          Case ScriptDB.ColumnRowSelection.Specific
    '            sPartCode = String.Format("{0}SELECT TOP {3} @part_{1} = base.[{2}]" & vbNewLine _
    '                , [CodeCluster].Indentation, iPartNumber, objThisColumn.Name, Component.ChildRowDetails.RowNumber)

    '          Case ScriptDB.ColumnRowSelection.Total
    '            sPartCode = String.Format("{0}SELECT @part_{1} = SUM(base.[{2}])" & vbNewLine _
    '                , [CodeCluster].Indentation, iPartNumber, objThisColumn.Name)
    '            bIsSummaryColumn = True

    '          Case ScriptDB.ColumnRowSelection.Count
    '            sPartCode = String.Format("{0}SELECT @part_{1} = COUNT(base.[{2}])" & vbNewLine _
    '                , [CodeCluster].Indentation, iPartNumber, objThisColumn.Name)
    '            bIsSummaryColumn = True

    '          Case Else
    '            sPartCode = String.Format("{0}SELECT TOP 1 @part_{1} = base.[{2}]" & vbNewLine _
    '                        , [CodeCluster].Indentation, iPartNumber, objThisColumn.Name)

    '        End Select



  End Class
End Namespace