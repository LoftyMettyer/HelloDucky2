Namespace Things
  Public Class TableOrderFilter
    Inherits Things.Base

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.TableOrderFilter
      End Get
    End Property

    '    Public Relation As Things.Relation
    '   Public Order As Things.TableOrder
    '  Public Filter As Things.Expression
    Public ComponentNumber As Long
    Public UDF As ScriptDB.GeneratedUDF
    Public RowDetails As Things.ChildRowDetails

    Public ReadOnly Property Table As Things.Table
      Get
        Return Parent
      End Get
    End Property

    Public ReadOnly Property IncludedColumns As Things.Collection
      Get
        Return Me.Objects
      End Get
    End Property

    Public Overrides Property Name As String
      Get
        Dim sName As String

        sName = String.Format("udftab_{0}", Parent.Name)

        If Not RowDetails.Order Is Nothing Then
          sName = sName + "_" + RowDetails.Order.Name
        End If

        If Not RowDetails.Filter Is Nothing Then
          sName = sName + "_" + RowDetails.Filter.Name
        End If

        If Not RowDetails.Relation Is Nothing Then
          sName = String.Format("{0}_{1}", sName, CInt(RowDetails.Relation.ParentID))
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

      Dim objOrderItem As Things.TableOrderItem
      Dim objColumn As Things.Column
      Dim aryOrderBy As New ArrayList
      Dim aryColumnList As New ArrayList
      Dim aryReturnDefintion As New ArrayList
      Dim aryParameters As New ArrayList
      Dim aryWheres As New ArrayList
      Dim aryJoins As New ArrayList
      Dim sRowSelection As String = vbNullString
      Dim bReverseOrder As Boolean = False
      Dim sOptions As String = ""
      Dim objIndex As New Things.Index

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
        RowDetails.Filter.AssociatedColumn = Me.Parent.Objects(Enums.Type.Column)(0)
        RowDetails.Filter.ExpressionType = ScriptDB.ExpressionType.ColumnFilter
        RowDetails.Filter.GenerateCode()
        aryWheres.Add(String.Format("({0} = 1)", RowDetails.Filter.UDF.SelectCode))
        aryJoins.Add(RowDetails.Filter.UDF.JoinCode)
      End If

      ' Build the order by clause
      If Not RowDetails.Order Is Nothing And _
          Not (RowDetails.RowSelection = ScriptDB.ColumnRowSelection.Total Or RowDetails.RowSelection = ScriptDB.ColumnRowSelection.Count) Then
        For Each objOrderItem In RowDetails.Order.Objects
          If objOrderItem.ColumnType = "O" Then
            Select Case objOrderItem.Ascending
              Case Enums.Order.Ascending
                aryOrderBy.Add(String.Format("[{0}].[{1}] {2}", objOrderItem.Column.Table, objOrderItem.Column.Name, IIf(bReverseOrder, "DESC", "ASC")))
              Case Else
                aryOrderBy.Add(String.Format("[{0}].[{1}] {2}", objOrderItem.Column.Table, objOrderItem.Column.Name, IIf(bReverseOrder, "ASC", "DESC")))
            End Select
            objIndex.Columns.AddIfNew(objOrderItem.Column)
          End If
        Next
      End If

      ' Add foreign key
      If Not RowDetails.Relation Is Nothing Then
        aryParameters.Add(String.Format("@prmID_{0} integer", CInt(RowDetails.Relation.ParentID)))
        aryWheres.Add(String.Format("[ID_{0}] = @prmID_{0}", CInt(RowDetails.Relation.ParentID)))
        objIndex.Relations.AddIfNew(RowDetails.Relation)
      End If

      ' Add the included columns
      For Each objColumn In IncludedColumns
        Select Case RowDetails.RowSelection
          Case ScriptDB.ColumnRowSelection.Count
            aryColumnList.Add(String.Format("COUNT(base.[{0}])", objColumn.Name))
            aryReturnDefintion.Add(String.Format("[{0}] numeric(38,8)", objColumn.Name))
          Case ScriptDB.ColumnRowSelection.Total
            aryColumnList.Add(String.Format("SUM(base.[{0}])", objColumn.Name))
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
        .WhereCode = IIf(aryWheres.Count > 0, "WHERE ", "") & String.Join(" AND ", aryWheres.ToArray())
        .OrderCode = IIf(aryOrderBy.Count > 0, "ORDER BY ", "") & String.Join(", ", aryOrderBy.ToArray())

        If RowDetails.RowSelection = ScriptDB.ColumnRowSelection.Specific Then

          .Code = String.Format("CREATE FUNCTION dbo.[{0}]({1})" & vbNewLine & _
             "RETURNS @results TABLE({2})" & vbNewLine & _
             "{9}" & vbNewLine & _
             "AS" & vbNewLine & "BEGIN" & vbNewLine & _
             "WITH base AS (" & vbNewLine & _
             "    SELECT *, [rownumber] = ROW_NUMBER() OVER ({7})" & vbNewLine & _
             "    FROM {4} base" & vbNewLine & _
             "    {6})" & vbNewLine & _
             "INSERT @Results SELECT {3}" & vbNewLine & _
             "        FROM base" & vbNewLine & _
             "        WHERE [rownumber] = {10}" & vbNewLine & _
             "    RETURN;" & vbNewLine & _
             "END" _
            , Me.Name, String.Join(", ", aryParameters.ToArray()) _
            , String.Join(", ", aryReturnDefintion.ToArray()), String.Join(", ", aryColumnList.ToArray()) _
            , Me.Parent.Name, String.Join(vbNewLine, aryJoins.ToArray()), .WhereCode, .OrderCode, sRowSelection _
            , sOptions, RowDetails.RowNumber)

        Else
          .Code = String.Format("CREATE FUNCTION dbo.[{0}]({1})" & vbNewLine & _
                         "RETURNS @results TABLE({2})" & vbNewLine & _
                         "{9}" & vbNewLine & _
                         "AS" & vbNewLine & "BEGIN" & vbNewLine & _
                         "INSERT @Results SELECT{8} {3}" & vbNewLine & _
                         "        FROM dbo.[{4}] base" & vbNewLine & _
                         "        {5}" & vbNewLine & _
                         "        {6}" & vbNewLine & _
                         "        {7}" & vbNewLine & _
                         "    RETURN;" & vbNewLine & _
                         "END" _
                        , Me.Name, String.Join(", ", aryParameters.ToArray()) _
                        , String.Join(", ", aryReturnDefintion.ToArray()), String.Join(", ", aryColumnList.ToArray()) _
                        , Me.Parent.Name, String.Join(vbNewLine, aryJoins.ToArray()), .WhereCode, .OrderCode, sRowSelection _
                        , sOptions)
        End If

        ' Add the index
        Me.Parent.Objects.Add(objIndex)

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