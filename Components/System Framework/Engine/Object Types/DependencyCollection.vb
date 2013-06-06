Imports System.Text.RegularExpressions

Public Structure Dependency
  Property PartNumber As Integer
  Property Declaration As String
  Property Code As String
  Property Thing As Object
  Property Filter As Expression
  Property Order As TableOrder
  Property RowNumber As Integer
End Structure

Public Class ExpressionDependencies

  Private ReadOnly _colThings As New Collection(Of Dependency)
  Private ReadOnly _colCode As New Collection(Of ScriptDB.GeneratedUdf)

  Private ReadOnly _dicColumns As New Dictionary(Of Column, Long)

  Private ReadOnly _dicTables As New Dictionary(Of Table, Long)
  Private ReadOnly _dicExpressions As New Dictionary(Of Expression, Long)
  Private ReadOnly _dicRelations As New Dictionary(Of Relation, Long)

  ' Code expressions
  Public Overloads Function Add(ByVal pobjExpression As Expression) As String

    Dim iPartNumber As Integer = 0
    Dim code As New ScriptDB.GeneratedUdf
    Dim bFound As Boolean = False

    code.SelectCode = pobjExpression.Udf.SelectCode
    code.FromCode = pobjExpression.Udf.FromCode
    code.JoinCode = pobjExpression.Udf.JoinCode
    code.WhereCode = pobjExpression.Udf.WhereCode

    ' Is this column/filter/order/line no in the dependcy stack
    For Each objThing As ScriptDB.GeneratedUdf In _colCode
      If objThing.SelectCode = code.SelectCode And objThing.FromCode = code.FromCode And objThing.JoinCode = code.JoinCode And objThing.WhereCode = code.WhereCode Then
        iPartNumber = objThing.PartNumber
        bFound = True
        Exit For
      End If
    Next

    ' Not in the list - add it and return this value
    If Not bFound Then

      iPartNumber = _colCode.Count
      code.PartNumber = iPartNumber

      code.Declaration = String.Format("@part_{0} {1}", iPartNumber, pobjExpression.DataTypeSyntax)
      code.Code = String.Format("SELECT @part_{0} = {1} " & _
        "{2}" & vbNewLine & _
        "{3}" & vbNewLine & _
        "{4}" & vbNewLine _
        , iPartNumber _
        , code.SelectCode, code.FromCode, code.JoinCode, code.WhereCode)

      code.Code = Regex.Replace(code.Code, "\s*(\n)", "$1")

      _colCode.Add(code)
    End If

    Return String.Format("@part_{0}", iPartNumber)

  End Function

  ' Child Columns
  Public Overloads Function Add(ByVal child As ChildRowDetails) As Integer

    Dim objDepends As New Dependency
    Dim objOrderFilter As TableOrderFilter
    Dim iPartNumber As Integer = 0
    Dim bFound As Boolean = False
    Dim sTypesafeCode As String
    Dim aryParameters As New ArrayList

    objDepends.Thing = child.Column
    objDepends.Filter = child.Filter
    objDepends.Order = child.Order
    objDepends.RowNumber = child.RowNumber

    ' Is this column/filter/order/line no in the dependcy stack
    For Each objThing As Dependency In _colThings
      If objThing.Thing Is child.Column And objThing.Filter Is child.Filter And objThing.Order Is child.Order And objThing.RowNumber = child.RowNumber Then
        iPartNumber = objThing.PartNumber
        bFound = True
        Exit For
      End If
    Next

    ' Not in the list - add it and return this value
    If Not bFound Then

      aryParameters.Add(String.Format("@prm_ID"))

      iPartNumber = _colThings.Count
      objDepends.PartNumber = iPartNumber

      If child.RowSelection = ScriptDB.ColumnRowSelection.Total Or child.RowSelection = ScriptDB.ColumnRowSelection.Count Then
        objDepends.Declaration = String.Format("@child_{0} numeric(38,8)", iPartNumber)
        sTypesafeCode = "0"
      Else
        objDepends.Declaration = String.Format("@child_{0} {1}", iPartNumber, child.Column.DataTypeSyntax)
        sTypesafeCode = child.Column.SafeReturnType
      End If


      ' Calculate the extra parameters we require if there's a filter attached
      If Not child.Filter Is Nothing Then

        child.Filter.AssociatedColumn = child.Filter.BaseTable.Columns(0)
        child.Filter.ExpressionType = ScriptDB.ExpressionType.ColumnFilter
        child.Filter.GenerateCodeForColumn()

        ' Add the dependent columns
        For Each objColumn In child.Filter.Dependencies.Columns
          If objColumn.Table Is child.BaseTable Then
            'aryParameters.Add(String.Format("@prm_{0}", objColumn.Name))
          End If
        Next

      End If

      objOrderFilter = child.Column.Table.TableOrderFilter(child)
      objOrderFilter.IncludedColumns.AddIfNew(child.Column)

      objDepends.Code = String.Format("SET @child_{0} = {2};" & vbNewLine & _
          "SELECT @child_{0} = ISNULL(base.[{1}],{2})" & vbNewLine & _
          "    FROM [dbo].[{3}]({4}) base" & vbNewLine _
          , iPartNumber.ToString, child.Column.Name, sTypesafeCode, objOrderFilter.Name _
          , String.Join(", ", aryParameters.ToArray()))

      _colThings.Add(objDepends)
    End If

    Return iPartNumber

  End Function

  ' Adds tables
  Public Overloads Sub Add(ByVal dependency As Table)
    If _dicTables.ContainsKey(dependency) Then
      _dicTables(dependency) = _dicTables(dependency) + 1
    Else
      _dicTables.Add(dependency, 0)
    End If
  End Sub

  ' Adds columns
  Public Overloads Sub Add(ByVal dependency As Column)
    If _dicColumns.ContainsKey(dependency) Then
      _dicColumns(dependency) = _dicColumns(dependency) + 1
    Else
      _dicColumns.Add(dependency, 0)
    End If
  End Sub

  ' Adds relation
  Public Overloads Sub Add(ByVal dependency As Relation)
    If _dicRelations.ContainsKey(dependency) Then
      _dicRelations(dependency) = _dicRelations(dependency) + 1
    Else
      _dicRelations.Add(dependency, 0)
    End If
  End Sub


  Public Sub Clear()
    _colThings.Clear()
    _dicColumns.Clear()
    _dicTables.Clear()
    _dicExpressions.Clear()
    _dicRelations.Clear()
  End Sub

  Public ReadOnly Property Statements As ICollection(Of ScriptDB.GeneratedUdf)
    Get
      Return _colCode
    End Get
  End Property

  Public ReadOnly Property ChildRowDetails As ICollection(Of Dependency)
    Get
      Return _colThings
    End Get
  End Property

  Public ReadOnly Property Relations As ICollection(Of Relation)
    Get
      Return _dicRelations.Keys
    End Get
  End Property

  Public ReadOnly Property Expressions As ICollection(Of Expression)
    Get
      Return _dicExpressions.Keys
    End Get
  End Property

  Public ReadOnly Property Tables As ICollection(Of Table)
    Get
      Return _dicTables.Keys
    End Get
  End Property

  Public ReadOnly Property Columns As ICollection(Of Column)
    Get
      Return _dicColumns.Keys
    End Get
  End Property

End Class
