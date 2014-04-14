Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports Microsoft.SqlServer.Server
Imports Microsoft.SqlServer.Server.SqlFunctionAttribute

Partial Public Class UserDefinedFunctions

  <SqlFunction(Name:="udfstat_isfieldpopulated", IsDeterministic:=True, DataAccess:=DataAccessKind.None)> _
  Public Shared Function IsFieldPopulated(ByVal Value As String, ByVal DataType As Integer) As SqlTypes.SqlBoolean

    Dim bOK As Boolean = False

    If Value Is Nothing Then
      bOK = False
    Else

      Select Case DataType

        Case 1 ' String
          bOK = (Len(Value.Trim()) > 0)

        Case 2 ' Numeric
          If IsNumeric(Value) Then
            bOK = Not (CDbl(Value) = 0)
          Else
            bOK = False
          End If

        Case 3 ' Logic
          bOK = (Value = "1")

        Case 4 ' Date
          bOK = IsDate(Value)

      End Select
    End If

    Return bOK

  End Function

  <SqlFunction(Name:="udfstat_isfieldempty", IsDeterministic:=True, DataAccess:=DataAccessKind.None)> _
  Public Shared Function IsFieldEmpty(ByVal Value As String, ByVal DataType As Integer) As SqlTypes.SqlBoolean

    Return Not IsFieldPopulated(Value, DataType)

  End Function

  <SqlFunction(Name:="udfstat_convertcharactertonumeric", IsDeterministic:=True, DataAccess:=DataAccessKind.None)> _
  Public Shared Function ConvertCharacterToNumeric(ByVal Value As String) As SqlTypes.SqlDouble

    If IsNumeric(Value) Then
      Return CDbl(Value)
    Else
      Return 0
    End If

  End Function

  <SqlFunction(Name:="udfstat_divideby", IsDeterministic:=True, DataAccess:=DataAccessKind.None)> _
  Public Shared Function DivideBy(ByVal Value As SqlTypes.SqlDouble, ByVal DivideByValue As SqlTypes.SqlDouble) As SqlTypes.SqlDouble

		If DivideByValue = 0 Then
			Return 0
		Else
			Return Value / DivideByValue
		End If

  End Function





End Class
