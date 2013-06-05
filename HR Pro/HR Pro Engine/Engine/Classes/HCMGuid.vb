Imports System.Runtime.InteropServices

<Serializable(), ComVisible(True)> _
Public Structure HCMGuid
  Implements IConvertible

  Private mintValue As Integer

  Default Public Shadows Property Value(ByVal sString As String) As Integer
    <System.Diagnostics.DebuggerStepThrough()> _
    Get
      Return mintValue
    End Get
    <System.Diagnostics.DebuggerStepThrough()> _
    Set(ByVal value As Integer)
      mintValue = value
    End Set
  End Property

  Public ReadOnly Property PadLeft() As String
    Get
      Return mintValue.ToString.PadLeft(8, "0")
    End Get
  End Property

#Region "IConvertible Interface"

  Public Function GetTypeCode() As TypeCode Implements IConvertible.GetTypeCode
    Return TypeCode.String
  End Function

  Function ToBoolean(ByVal provider As IFormatProvider) As Boolean Implements IConvertible.ToBoolean
    Return False
  End Function

  Function ToByte(ByVal provider As IFormatProvider) As Byte Implements IConvertible.ToByte
    Return Convert.ToByte(0)
  End Function

  Function ToChar(ByVal provider As IFormatProvider) As Char Implements IConvertible.ToChar
    Return Convert.ToChar(0)
  End Function

  Function ToDateTime(ByVal provider As IFormatProvider) As DateTime Implements IConvertible.ToDateTime
    Return Convert.ToDateTime(0)
  End Function

  Function ToDecimal(ByVal provider As IFormatProvider) As Decimal Implements IConvertible.ToDecimal
    Return Convert.ToDecimal(0)
  End Function

  Function ToDouble(ByVal provider As IFormatProvider) As Double Implements IConvertible.ToDouble
    Return 0
  End Function

  Function ToInt16(ByVal provider As IFormatProvider) As Short Implements IConvertible.ToInt16
    Return Convert.ToInt16(0)
  End Function

  Function ToInt32(ByVal provider As IFormatProvider) As Integer Implements IConvertible.ToInt32
    Return Convert.ToInt32(0)
  End Function

  Function ToInt64(ByVal provider As IFormatProvider) As Long Implements IConvertible.ToInt64
    Return Convert.ToInt64(0)
  End Function

  <CLSCompliant(False)> _
  Function ToSByte(ByVal provider As IFormatProvider) As SByte Implements IConvertible.ToSByte
    Return Convert.ToSByte(0)
  End Function

  Function ToSingle(ByVal provider As IFormatProvider) As Single Implements IConvertible.ToSingle
    Return Convert.ToSingle(0)
  End Function

  Overloads Function iConvertible_ToString(ByVal provider As IFormatProvider) As String Implements IConvertible.ToString
    Return mintValue.ToString
  End Function

  Function ToType(ByVal conversionType As Type, ByVal provider As IFormatProvider) As Object Implements IConvertible.ToType
    Return Convert.ChangeType(0, conversionType)
  End Function

  <CLSCompliant(False)> _
  Function ToUInt16(ByVal provider As IFormatProvider) As UInt16 Implements IConvertible.ToUInt16
    Return Convert.ToUInt16(0)
  End Function

  <CLSCompliant(False)> _
  Function ToUInt32(ByVal provider As IFormatProvider) As UInt32 Implements IConvertible.ToUInt32
    Return Convert.ToUInt32(0)
  End Function

  <CLSCompliant(False)> _
  Function ToUInt64(ByVal provider As IFormatProvider) As UInt64 Implements IConvertible.ToUInt64
    Return Convert.ToUInt64(0)
  End Function

#End Region

#Region "Casting Operators"

  Public Shared Narrowing Operator CType(ByVal x As String) As HCMGuid
    Dim r As New HCMGuid
    r.Value(0) = x
    Return r
  End Operator

  '<System.Diagnostics.DebuggerStepThrough()> _
  Public Shared Widening Operator CType(ByVal x As HCMGuid) As String
    Dim r As String
    r = x(0).ToString
    Return r
  End Operator

#End Region

#Region "Comparison Operators"
  <System.Diagnostics.DebuggerStepThrough()> _
  Public Shared Operator =(ByVal x As HCMGuid, ByVal y As HCMGuid) As Boolean
    Dim r As Boolean
    r = (x(0) = y(0))
    Return r
  End Operator

  <System.Diagnostics.DebuggerStepThrough()> _
  Public Shared Operator <>(ByVal x As HCMGuid, ByVal y As HCMGuid) As Boolean
    Dim r As Boolean
    r = (x(0) <> y(0))
    Return r
  End Operator

  <System.Diagnostics.DebuggerStepThrough()> _
  Public Shared Operator =(ByVal x As Integer, ByVal y As HCMGuid) As Boolean
    Dim r As Boolean
    r = (x = y(0))
    Return r
  End Operator

  <System.Diagnostics.DebuggerStepThrough()> _
  Public Shared Operator <>(ByVal x As Integer, ByVal y As HCMGuid) As Boolean
    Dim r As Boolean
    r = (x <> y(0))
    Return r
  End Operator

  <System.Diagnostics.DebuggerStepThrough()> _
  Public Shared Operator =(ByVal x As HCMGuid, ByVal y As Integer) As Boolean
    Dim r As Boolean
    r = (x(0) = y)
    Return r
  End Operator

  <System.Diagnostics.DebuggerStepThrough()> _
  Public Shared Operator <>(ByVal x As HCMGuid, ByVal y As Integer) As Boolean
    Dim r As Boolean
    r = (x(0) <> y)
    Return r
  End Operator

  '  <System.Diagnostics.DebuggerStepThrough()> _
  '  Public Shared Operator =(ByVal x As HCMGuid, ByVal y As String) As Boolean
  '    Dim r As Boolean
  '    r = (x(0) = y)
  '    Return r
  '  End Operator

  '  <System.Diagnostics.DebuggerStepThrough()> _
  '  Public Shared Operator <>(ByVal x As HCMGuid, ByVal y As String) As Boolean
  '    Dim r As Boolean
  '    r = (x(0) <> y)
  '    Return r
  '  End Operator

  '  <System.Diagnostics.DebuggerStepThrough()> _
  '  Public Shared Operator =(ByVal x As Object, ByVal y As HCMGuid) As Boolean
  '    Dim r As Boolean
  '    r = (x.ToString = y(0))
  '    Return r
  '  End Operator

  '  <System.Diagnostics.DebuggerStepThrough()> _
  '  Public Shared Operator <>(ByVal x As Object, ByVal y As HCMGuid) As Boolean
  '    Dim r As Boolean
  '    r = (x.ToString <> y(0))
  '    Return r
  '  End Operator

  '  <System.Diagnostics.DebuggerStepThrough()> _
  '  Public Shared Operator =(ByVal x As System.DBNull, ByVal y As HCMGuid) As Boolean
  '    Dim r As Boolean
  '    r = (x.ToString = y(0))
  '    Return r
  '  End Operator

  '  <System.Diagnostics.DebuggerStepThrough()> _
  '  Public Shared Operator <>(ByVal x As System.DBNull, ByVal y As HCMGuid) As Boolean
  '    Dim r As Boolean
  '    r = (x.ToString <> y(0))
  '    Return r
  '  End Operator

#End Region

#Region "Manipulation Operators"

  Public Shared Operator &(ByVal Left As String, ByVal Right As HCMGuid) As String
    Return New String(Left & Right(0))
  End Operator



#End Region

End Structure
