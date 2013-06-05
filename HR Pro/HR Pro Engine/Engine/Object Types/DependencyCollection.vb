Imports System.Text.RegularExpressions

Public Structure Dependency
  Property PartNumber As Integer
  Property Declaration As String
  Property Code As String

  Property Thing As Object
  Property Filter As Things.Expression
  Property Order As Things.TableOrder
  Property RowNumber As Integer
End Structure


Public Class ExpressionDependencies

  Private colThings As New Collection(Of Dependency)
  Private colCode As New Collection(Of ScriptDB.GeneratedUDF)

  Private dicColumns As New Dictionary(Of Column, Long)

  Private dicTables As New Dictionary(Of Table, Long)
  Private dicExpressions As New Dictionary(Of Expression, Long)
  Private dicRelations As New Dictionary(Of Relation, Long)

  ' Code expressions
  Public Overloads Function Add(ByVal pobjExpression As Things.Expression) As String

    Dim iPartNumber As Integer = 0
    Dim code As New ScriptDB.GeneratedUDF
    Dim bFound As Boolean = False

    code.SelectCode = pobjExpression.UDF.SelectCode
    code.FromCode = pobjExpression.UDF.FromCode
    code.JoinCode = pobjExpression.UDF.JoinCode
    code.WhereCode = pobjExpression.UDF.WhereCode

    ' Is this column/filter/order/line no in the dependcy stack
    For Each objThing As ScriptDB.GeneratedUDF In colCode
      If objThing.SelectCode = code.SelectCode And objThing.FromCode = code.FromCode And objThing.JoinCode = code.JoinCode And objThing.WhereCode = code.WhereCode Then
        iPartNumber = objThing.PartNumber
        bFound = True
        Exit For
      End If
    Next

    ' Not in the list - add it and return this value
    If Not bFound Then

      iPartNumber = colCode.Count
      code.PartNumber = iPartNumber

      code.Declaration = String.Format("@part_{0} {1}", iPartNumber, pobjExpression.DataTypeSyntax)
      code.Code = String.Format("SELECT @part_{0} = {1} " & _
        "{2}" & vbNewLine & _
        "{3}" & vbNewLine & _
        "{4}" & vbNewLine _
        , iPartNumber _
        , code.SelectCode, code.FromCode, code.JoinCode, code.WhereCode)

      code.Code = Regex.Replace(code.Code, "\s*(\n)", "$1")

      colCode.Add(code)
    End If

    Return String.Format("@part_{0}", iPartNumber)

  End Function

  ' Child Columns
  Public Overloads Function Add(ByVal Child As ChildRowDetails) As Integer

    Dim objDepends As New Dependency
    Dim objOrderFilter As TableOrderFilter
    Dim iPartNumber As Integer = 0
    Dim bFound As Boolean = False
    Dim sTypesafeCode As String
    objDepends.Thing = Child.Column
    objDepends.Filter = Child.Filter
    objDepends.Order = Child.Order
    objDepends.RowNumber = Child.RowNumber

    ' Is this column/filter/order/line no in the dependcy stack
    For Each objThing As Dependency In colThings
      If objThing.Thing Is Child.Column And objThing.Filter Is Child.Filter And objThing.Order Is Child.Order And objThing.RowNumber = Child.RowNumber Then
        iPartNumber = objThing.PartNumber
        bFound = True
        Exit For
      End If
    Next

    ' Not in the list - add it and return this value
    If Not bFound Then

      iPartNumber = colThings.Count
      objDepends.PartNumber = iPartNumber

      If Child.RowSelection = ScriptDB.ColumnRowSelection.Total Or Child.RowSelection = ScriptDB.ColumnRowSelection.Count Then
        objDepends.Declaration = String.Format("@child_{0} numeric(38,8)", iPartNumber)
        sTypesafeCode = "0"
      Else
        objDepends.Declaration = String.Format("@child_{0} {1}", iPartNumber, Child.Column.DataTypeSyntax)
        sTypesafeCode = Child.Column.SafeReturnType
      End If


      objOrderFilter = Child.Column.Table.TableOrderFilter(Child)
      objOrderFilter.IncludedColumns.AddIfNew(Child.Column)

      objDepends.Code = String.Format("SELECT @child_{0} = ISNULL(base.[{1}],{2})" & vbNewLine & _
          "FROM [dbo].[{3}](@prm_ID) base" _
          , iPartNumber.ToString, Child.Column.Name, sTypesafeCode, objOrderFilter.Name)

      colThings.Add(objDepends)
    End If

    Return iPartNumber

  End Function

  ' Adds tables
  Public Overloads Sub Add(ByVal Dependency As Table)
    If dicTables.ContainsKey(Dependency) Then
      dicTables(Dependency) = dicTables(Dependency) + 1
    Else
      dicTables.Add(Dependency, 0)
    End If
  End Sub

  ' Adds columns
  Public Overloads Sub Add(ByVal Dependency As Column)
    If dicColumns.ContainsKey(Dependency) Then
      dicColumns(Dependency) = dicColumns(Dependency) + 1
    Else
      dicColumns.Add(Dependency, 0)
    End If
  End Sub

  ' Adds relation
  Public Overloads Sub Add(ByVal Dependency As Relation)
    If dicRelations.ContainsKey(Dependency) Then
      dicRelations(Dependency) = dicRelations(Dependency) + 1
    Else
      dicRelations.Add(Dependency, 0)
    End If
  End Sub


  Public Sub Clear()
    colThings.Clear()
    dicColumns.Clear()
    dicTables.Clear()
    dicExpressions.Clear()
    dicRelations.Clear()
  End Sub

  Public ReadOnly Property Statements As ICollection(Of ScriptDB.GeneratedUDF)
    Get
      Return colCode
    End Get
  End Property

  Public ReadOnly Property ChildRowDetails As ICollection(Of Dependency)
    Get
      Return colThings
    End Get
  End Property

  Public ReadOnly Property Relations As ICollection(Of Relation)
    Get
      Return dicRelations.Keys
    End Get
  End Property

  Public ReadOnly Property Expressions As ICollection(Of Expression)
    Get
      Return dicExpressions.Keys
    End Get
  End Property

  Public ReadOnly Property Tables As ICollection(Of Table)
    Get
      Return dicTables.Keys
    End Get
  End Property

  Public ReadOnly Property Columns As ICollection(Of Column)
    Get
      Return dicColumns.Keys
    End Get
  End Property

End Class
