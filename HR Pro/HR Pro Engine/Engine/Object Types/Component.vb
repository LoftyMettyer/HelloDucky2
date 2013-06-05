Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
  Public Class Component
    Inherits Base

    Public Property SubType As ScriptDB.ComponentTypes
    Public Property ReturnType As ScriptDB.ComponentValueTypes
    Public Property FunctionID As Integer
    Public Property OperatorID As Integer
    Public Property CalculationID As Integer
    Public Property ValueType As ScriptDB.ComponentValueTypes
    Public Property ValueString As String
    Public Property ValueDate As Date
    Public Property ValueLogic As Boolean

    Public Property TableID As Integer
    Public Property ColumnID As Integer
    Public ChildRowDetails As ChildRowDetails

    Public Property IsColumnByReference As Boolean
    Public Property LookupTableID As Integer
    Public Property LookupColumnID As Integer

    Public Property BaseExpression As Expression
    Public Property AssociatedColumn As Column

    Public Property IsSchemaBound As Boolean = True
    Public Property IsTimeDependant As Boolean

    Public Property Parent As Component
    Public Property Components As ICollection(Of Component)
    Public Property Level As Long = 0

    Private mdecValueNumeric As Decimal = 0

    Private ConvertSubComponents As Boolean = True

    Public Sub New()
      Components = New Collection(Of Component)
    End Sub

    Public Property ValueNumeric As String
      Get

        Dim sValue As String

        sValue = mdecValueNumeric.ToString.Replace(",", ".")

        ' JIRA-1976 - SQL interprets values as integer if no decimal place - causes problems with divisions.
        If sValue.IndexOf(".") = -1 Then
          sValue = String.Format("{0}.0", sValue)
        End If

        Return sValue
      End Get
      Set(ByVal value As String)
        mdecValueNumeric = CDec(value)
      End Set
    End Property

#Region "Convert components to expressions"

    Public Sub ConvertToExpression()

      Dim objRecursiveComponents As New Collection(Of Base)
      objRecursiveComponents.Add(Me.AssociatedColumn)
      ConvertToExpression(0, objRecursiveComponents)

    End Sub

    Public Sub ConvertToExpression(ByRef Level As Long, ByRef Recursion As Collection(Of Base))

      Dim objExpression As Expression
      Dim objColumn As Column
      Dim objTable As Table
      Dim objClone As New Collection(Of Component)
      Dim bConvertsSubComponents As Boolean
      Dim lngThisLevel As Long

      Try

        Level = Level + 1
        lngThisLevel = Level

        If Me.SubType = ScriptDB.ComponentTypes.Calculation Then

          objExpression = Globals.Expressions.GetById(Me.CalculationID).Clone
          Me.Components = objExpression.Components
          Me.TableID = objExpression.TableID
          Me.ReturnType = objExpression.ReturnType
          Me.SubType = ScriptDB.ComponentTypes.Expression

          If Me.TableID <> Me.BaseExpression.TableID Then
            Globals.ErrorLog.Add(ErrorHandler.Section.UDFs, "", ErrorHandler.Severity.Warning _
            , String.Format("Error creating calculation for {0}.{1} ", Me.BaseExpression.BaseTable.Name, Me.BaseExpression.AssociatedColumn.Name) _
              , "This is likely caused by copying a table and a calculation reference is still attached to the original column. In the associated calculation try re-selecting any calculations")
            Me.BaseExpression.IsValid = False
          End If

        ElseIf Me.SubType = ScriptDB.ComponentTypes.Expression Then

          For Each objComponent As Component In Me.Components
            objComponent.ConvertToExpression(Level, Recursion)
          Next

        ElseIf Me.SubType = ScriptDB.ComponentTypes.Function Then

          For Each objComponent As Component In Me.Components
            objComponent.ConvertToExpression(Level, Recursion)
          Next

          ' Pull a calculated column directly in as an expression
        ElseIf Me.SubType = ScriptDB.ComponentTypes.Column And Me.ConvertSubComponents Then

          If Not Me.IsColumnByReference Then

            objTable = Globals.Tables.GetById(Me.TableID)
            objColumn = objTable.Columns.GetById(Me.ColumnID)

            ' We've got ourselves into a recursive loop somehow
            If Recursion.Contains(objColumn) Then

            ElseIf objColumn.IsCalculated Then
              If objColumn.Table Is Me.BaseExpression.BaseTable Then

                objExpression = objColumn.Table.Expressions.GetById(objColumn.CalcID).Clone
                Me.Components = objExpression.Components
                Me.ReturnType = objColumn.ComponentReturnType
                bConvertsSubComponents = Not Recursion.Contains(objColumn)

                Recursion.AddIfNew(objColumn)

                For Each objComponent As Component In Me.Components
                  objComponent.ConvertSubComponents = bConvertsSubComponents
                  objComponent.ConvertToExpression(Level, Recursion)
                Next

                If lngThisLevel < Level Then
                  Recursion.Remove(objColumn)
                End If

                Me.SubType = ScriptDB.ComponentTypes.ConvertedCalculatedColumn

              End If
            End If
          End If

        End If

      Catch ex As Exception

        Globals.ErrorLog.Add(ErrorHandler.Section.UDFs, "", ErrorHandler.Severity.Error, "Calculation not found", CStr(Me.CalculationID))

      End Try


    End Sub

#End Region

#Region "Cloning"

    Public Function Clone() As Component

      Dim objClone As New Component

      ' Clone component properties (shallow clone)
      objClone = CType(Me.MemberwiseClone(), Component)

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
      For Each objComponent As Component In Me.Components
        objClone.Add(CType(objComponent.Clone, Component))
      Next

      Return objClone

    End Function

    ' Set the root expression on all nodes for this expression
    Friend Sub SetRootNode(ByRef RootNode As Expression)

      Me.BaseExpression = RootNode

      ' Clone the child nodes
      For Each objComponent As Component In Me.Components
        objComponent.SetRootNode(RootNode)
      Next

    End Sub


#End Region

  End Class
End Namespace

