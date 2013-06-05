Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()> _
  Public Class Component
    Inherits Things.Base
    Implements ICloneable

    Public Property ReturnType As ScriptDB.ComponentValueTypes
    Public Property FunctionID As Integer
    Public Property OperatorID As Integer
    Public Property CalculationID As Integer
    Public Property ValueType As ScriptDB.ComponentValueTypes
    Public Property ValueNumeric As Double
    Public Property ValueString As String
    Public Property ValueDate As Date
    Public Property ValueLogic As Boolean

    Public Property TableID As Integer
    Public Property ColumnID As Integer
    Public ChildRowDetails As ChildRowDetails

    Public Property IsColumnByReference As Boolean
    Public Property LookupTableID As Integer
    Public Property LookupColumnID As Integer

    Public Property BaseExpression As Things.Expression

    Public Property IsSchemaBound As Boolean = True
    Public Property IsTimeDependant As Boolean = False

    Public Property Components As New List(Of Component)

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Component
      End Get
    End Property

    Public Function Clone() As Object Implements System.ICloneable.Clone
      Return Me.MemberwiseClone
    End Function

    Public ReadOnly Property SafeReturnType As String
      Get

        Dim sqlType As String = String.Empty

        Select Case CInt(Me.ReturnType)
          Case ScriptDB.ComponentValueTypes.String
            sqlType = "''"

          Case ScriptDB.ComponentValueTypes.Numeric
            sqlType = "0"

          Case ScriptDB.ComponentValueTypes.Date
            sqlType = "NULL"

          Case Else
            sqlType = "0"

        End Select

        Return sqlType

      End Get
    End Property

  End Class
End Namespace

