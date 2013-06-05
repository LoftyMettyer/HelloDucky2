'Imports System.IO
'Imports System.Runtime.Serialization.Formatters.Binary
'Imports System.Runtime.Serialization

Namespace Things
  Public Class Component
    Inherits Things.Base
    Implements ICloneable

    '    Public SubType As ScriptDB.ComponentTypes

    'Public ReturnType As ScriptDB.ColumnTypes

    Public FunctionID As HCMGuid
    Public OperatorID As HCMGuid
    Public CalculationID As HCMGuid

    Public ReturnType As ScriptDB.ComponentValueTypes
    'Public IsScriptSafe As Boolean = True
    Public BypassValidation As Boolean = False
    Public IsComplex As Boolean = False
    Public IsScriptSafe As Boolean = True

    Public ValueType As ScriptDB.ComponentValueTypes
    Public ValueNumeric As Integer
    Public ValueString As String
    Public ValueDate As Date
    Public ValueLogic As Boolean

    Public TableID As HCMGuid
    Public ColumnID As HCMGuid
    Public ColumnFilterID As HCMGuid
    Public ColumnOrderID As HCMGuid
    Public ColumnAggregiateType As ScriptDB.AggregiateNumeric
    Public IsColumnByReference As Boolean

    Public LookupTableID As HCMGuid
    Public LookupColumnID As HCMGuid

    <System.Xml.Serialization.XmlIgnore()> _
    Public EmbedDependencies As Boolean = True
    <System.Xml.Serialization.XmlIgnore()> _
    Public InlineScript As Boolean = False


    <System.Xml.Serialization.XmlIgnore()> _
    Public BaseExpression As Things.Expression

    'Public Overrides Function Commit() As Boolean
    'End Function

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Component
      End Get
    End Property

    'Public Property BaseExpression
    '  Set(ByVal value)

    '    For Each objChild As Things.Component In Me.Objects
    '      objChild.BaseExpression = value
    '    Next

    '  End Set
    '  Get
    '    BaseExpression = objBaseExpression
    '  End Get
    'End Property

    Public Sub SetBaseExpression(ByRef objBaseExpression As Things.Component)

      ' Attach the base component info
      Me.BaseExpression = objBaseExpression
      For Each objComponent As Things.Component In Me.Objects
        objComponent.SetBaseExpression(objBaseExpression)
        'objComponent.BaseExpression = Me.BaseExpression
      Next

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone
      Return Me.MemberwiseClone
    End Function

    Public ReadOnly Property ToExpression() As Things.Expression
      Get

        Dim objExpression As New Things.Expression

        objExpression.ID = Me.ID
        objExpression.FunctionID = Me.FunctionID
        objExpression.Objects = Me.Objects
        objExpression.ReturnType = Me.ReturnType
        objExpression.ExpressionType = ScriptDB.ExpressionType.Mask

        Return objExpression

      End Get
    End Property


    'Public ReadOnly Property DataTypeSyntax As String
    '  Get

    '    Dim sSQLType As String = String.Empty

    '    Select Case Me.ReturnType
    '      Case ScriptDB.ComponentValueTypes.String, ScriptDB.ComponentValueTypes.Component_String
    '        sSQLType = "[varchar](MAX)"

    '      Case ScriptDB.ComponentValueTypes.Numeric, ScriptDB.ComponentValueTypes.Component_Numeric
    '        sSQLType = String.Format("[numeric]({38},{10})")

    '      Case ScriptDB.ComponentValueTypes.Date, ScriptDB.ComponentValueTypes.Component_Date
    '        sSQLType = "[datetime]"

    '      Case ScriptDB.ComponentValueTypes.Logic, ScriptDB.ComponentValueTypes.Component_Logic
    '        sSQLType = "[bit]"
    '    End Select

    '    Return sSQLType

    '  End Get

    'End Property

    'Public Shared Function DeepClone(ByVal obj As Object) As Object
    '  Dim memStream As MemoryStream = New MemoryStream
    '  Dim binaryFormatter As BinaryFormatter = New BinaryFormatter(Nothing, New StreamingContext(StreamingContextStates.Clone))
    '  binaryFormatter.Serialize(memStream, obj)
    '  memStream.Seek(0, SeekOrigin.Begin)
    '  Return binaryFormatter.Deserialize(memStream)
    'End Function

  End Class
End Namespace

