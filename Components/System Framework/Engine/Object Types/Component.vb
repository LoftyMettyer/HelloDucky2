Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
  Public Class Component
    Inherits Base

    Public Property SubType As ScriptDB.ComponentTypes
    Public Property ReturnType As ScriptDB.ComponentValueTypes
    Public Property FunctionId As Integer
    Public Property OperatorId As Integer
    Public Property CalculationId As Integer
    Public Property FilterId As Integer
    Public Property ValueType As ScriptDB.ComponentValueTypes
    Public Property ValueString As String
    Public Property ValueDate As Date
    Public Property ValueLogic As Boolean

    Public Property TableId As Integer
    Public Property ColumnId As Integer
    Public ChildRowDetails As ChildRowDetails

    Public Property IsColumnByReference As Boolean
    Public Property LookupTableId As Integer
    Public Property LookupColumnId As Integer

    Public Property BaseExpression As Expression
    Public Property AssociatedColumn As Column

    Public Property IsSchemaBound As Boolean = True
    Public Property IsTimeDependant As Boolean

    Public Property Parent As Component
    Public Property Components As ICollection(Of Component)
    Public Property Level As Long = 0

    Private _mdecValueNumeric As Decimal = 0

    Private _convertSubComponents As Boolean = True
    Public Shared DepthCharge As Long

    Public Sub New()
      Components = New Collection(Of Component)
    End Sub

    Public Property ValueNumeric As String
      Get

        Dim sValue As String

        sValue = _mdecValueNumeric.ToString.Replace(",", ".")

        ' JIRA-1976 - SQL interprets values as integer if no decimal place - causes problems with divisions.
        If sValue.IndexOf(".") = -1 Then
          sValue = String.Format("{0}.0", sValue)
        End If

        Return sValue
      End Get
      Set(ByVal value As String)
        _mdecValueNumeric = CDec(value)
      End Set
    End Property

#Region "Convert components to expressions"

    Public Sub ConvertToExpression()

      DepthCharge = 0

      Dim objRecursiveComponents As New Collection(Of Base)
      objRecursiveComponents.Add(AssociatedColumn)
      ConvertToExpression(0, objRecursiveComponents)

    End Sub

    Public Sub ConvertToExpression(ByRef Level As Long, ByRef recursion As Collection(Of Base))

      Dim objExpression As Expression
      Dim objColumn As Column
      Dim objTable As Table
      Dim bConvertsSubComponents As Boolean
      Dim lngThisLevel As Long

      Try

        Level = Level + 1
        lngThisLevel = Level

        If SubType = ScriptDB.ComponentTypes.Calculation Then

          objExpression = Expressions.GetById(CalculationId).Clone
          Components = objExpression.Components
          TableId = objExpression.TableId
          ReturnType = objExpression.ReturnType
          SubType = ScriptDB.ComponentTypes.Expression

          If TableId <> BaseExpression.TableId Then
            ErrorLog.Add(ErrorHandler.Section.UdFs, "", ErrorHandler.Severity.Warning _
            , String.Format("Error creating calculation for {0}.{1} ", BaseExpression.BaseTable.Name, BaseExpression.AssociatedColumn.Name) _
              , "This is likely to be caused by copying a table and a calculation reference is still attached to the original column. In the associated calculation try re-selecting any calculations.")
            BaseExpression.IsValid = False
          End If

        ElseIf SubType = ScriptDB.ComponentTypes.Expression Then

          For Each objComponent As Component In Components
            DepthCharge = DepthCharge + 1
            objComponent.ConvertToExpression(Level, recursion)
            DepthCharge = DepthCharge - 1
          Next

        ElseIf SubType = ScriptDB.ComponentTypes.Function Then

          For Each objComponent As Component In Components
            DepthCharge = DepthCharge + 1
            objComponent.ConvertToExpression(Level, recursion)
            DepthCharge = DepthCharge - 1
          Next

          ' Pull a calculated column directly in as an expression
        ElseIf SubType = ScriptDB.ComponentTypes.Column And _convertSubComponents Then

          If Not IsColumnByReference Then

            objTable = Tables.GetById(TableId)
            objColumn = objTable.Columns.GetById(ColumnId)

            ' We've got ourselves into a recursive loop somehow
            If recursion.Contains(objColumn) Then

            ElseIf objColumn.IsCalculated Then
              If objColumn.Table Is BaseExpression.BaseTable Then

                objExpression = objColumn.Table.Expressions.GetById(objColumn.CalcId).Clone
                Components = objExpression.Components
                ReturnType = objColumn.ComponentReturnType
                bConvertsSubComponents = Not recursion.Contains(objColumn)

                recursion.AddIfNew(objColumn)

                For Each objComponent As Component In Components
                  objComponent._convertSubComponents = bConvertsSubComponents
                  DepthCharge = DepthCharge + 1
                  objComponent.ConvertToExpression(Level, recursion)
                  DepthCharge = DepthCharge - 1
                Next

                If lngThisLevel < Level Then
                  recursion.Remove(objColumn)
                End If

                SubType = ScriptDB.ComponentTypes.ConvertedCalculatedColumn

              End If

            End If
          End If

        End If

      Catch ex As Exception

        ErrorLog.Add(ErrorHandler.Section.UdFs, "", ErrorHandler.Severity.Error, "Calculation not found", CStr(CalculationId))

      End Try


    End Sub

#End Region

#Region "Cloning"

    Public Function Clone() As Component

      Dim objClone As New Component

      ' Clone component properties (shallow clone)
      objClone = CType(MemberwiseClone(), Component)

      ' Clone the child nodes (deep clone)
      objClone.Components = New Collection(Of Component)
      For Each objComponent As Component In Me.Components
        objClone.Components.Add(CType(objComponent.Clone, Component))
      Next

      Return objClone
    End Function

    Public Function CloneComponents() As ICollection(Of Component)

      Dim objClone As New Collection(Of Component)

      ' Clone the child nodes
      For Each objComponent As Component In Components
        objClone.Add(CType(objComponent.Clone, Component))
      Next

      Return objClone

    End Function

    ' Set the root expression on all nodes for this expression
    Friend Sub SetRootNode(ByRef rootNode As Expression)

      BaseExpression = rootNode

      ' Clone the child nodes
      For Each objComponent As Component In Components
        objComponent.SetRootNode(rootNode)
      Next

    End Sub


#End Region

  End Class
End Namespace

