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
    'Public IsComplex As Boolean = False
    'Public IsScriptSafe As Boolean = True

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
    Public SpecificLine As Integer
    Public IsColumnByReference As Boolean
    'Public IsEvaluated As Boolean = False

    Public LookupTableID As HCMGuid
    Public LookupColumnID As HCMGuid

    '<System.Xml.Serialization.XmlIgnore()> _
    'Public EmbedDependencies As Boolean = True

    <System.Xml.Serialization.XmlIgnore()> _
Public InlineScript As Boolean = False

    <System.Xml.Serialization.XmlIgnore()> _
    Public BaseExpression As Things.Expression

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Component
      End Get
    End Property

    Public Sub SetBaseExpression(ByRef objBaseExpression As Things.Component)

      ' Attach the base component info
      Me.BaseExpression = objBaseExpression
      For Each objComponent As Things.Component In Me.Objects
        objComponent.SetBaseExpression(objBaseExpression)
      Next

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone
      Return Me.MemberwiseClone
    End Function

    'Public ReadOnly Property ToExpression() As Things.Expression
    '  Get

    '    Dim objExpression As New Things.Expression

    '    objExpression.ID = Me.ID
    '    objExpression.FunctionID = Me.FunctionID
    '    objExpression.Objects = Me.Objects
    '    objExpression.ReturnType = Me.ReturnType
    '    objExpression.ExpressionType = ScriptDB.ExpressionType.Mask

    '    Return objExpression

    '  End Get
    'End Property

  End Class
End Namespace

