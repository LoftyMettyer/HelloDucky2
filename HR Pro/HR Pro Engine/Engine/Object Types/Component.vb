Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
  Public Class Component
    Inherits Base
    Implements ICloneable

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

    Public Property IsSchemaBound As Boolean = True
    Public Property IsTimeDependant As Boolean

    Public Property Components As ICollection(Of Component)
    Public Property Level As Long = 0

    Private mdblValueNumeric As Double = 0

    Public Sub New()
      Components = New Collection(Of Component)
    End Sub

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

    Public Property ValueNumeric As String
      Get

        Dim sValue As String

        sValue = mdblValueNumeric.ToString.Replace(",", ".")

        ' JIRA-1976 - SQL interprets values as integer if no decimal place - causes problems with divisions.
        If sValue.IndexOf(".") = -1 Then
          sValue = String.Format("{0}.0", sValue)
        End If

        Return sValue
      End Get
      Set(ByVal value As String)
        mdblValueNumeric = CDbl(value)
      End Set
    End Property

  End Class
End Namespace

