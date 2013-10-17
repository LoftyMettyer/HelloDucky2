Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums

Namespace Metadata
	Public Class [Operator]
		Inherits Base
			Public ReturnType As ExpressionValueTypes
			Public Precedence As Integer
			Public OperandCount As Integer
			Public SPName As String
			Public SQLCode As String
			Public SQLType As String
			Public CheckDivideByZero As Boolean
			Public SQLFixedParam1 As String
			Public CastAsFloat As Boolean
			Public Parameters As ICollection(Of OperatorParameter)
	End Class
End Namespace
