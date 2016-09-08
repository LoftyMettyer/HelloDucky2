Option Strict On
Option Explicit On

Imports DMI.NET.Code.Attributes
Imports HR.Intranet.Server.Expressions
Imports HR.Intranet.Server
Namespace Classes
	Public Class ReportColumnItem
		Implements IJsonSerialize
		Implements IReportDetail

		Public Property ReportID As Integer Implements IReportDetail.ReportID
		Public Property ReportType As UtilityType Implements IReportDetail.ReportType

		Public Property ID As Integer Implements IJsonSerialize.ID
      Public Property TableID As Integer

      Public Property ViewID As Integer

      Public Property ViewName As String

      Public Property IsViewColumn As Boolean

      Public Property IsExpression As Boolean
      Public Property Name As String
		Public Property Sequence As Integer

      <ExcludeChar("/,.!@#$%")>
      Public Property Heading As String

      Public Property ColumnValue As String

      Public Property DataType As ColumnDataType
      Public Property Size As Long
		Public Property ColumnSize As Long
		Public Property Decimals As Integer
		Public Property IsAverage As Boolean
		Public Property IsCount As Boolean
		Public Property IsTotal As Boolean
		Public Property IsHidden As Boolean
		Public Property IsGroupWithNext As Boolean
      Public Property IsRepeated As Boolean

      ''Organisation Columns Properties
      Public Property ColumnID As Integer = 0
      Public Property Prefix As String
      Public Property Suffix As String
      Public Property FontSize As Integer = 11
      Public Property Height As Integer = 1
      Public Property DefaultHeight As Integer = 1
      'Public Property ConcatenateWithNext As Boolean

      ''' <summary>
      ''' Gets/Sets the access rights for the column (E.g. HD/RW/RO)
      ''' </summary>
      ''' <value>The column access value</value>
      Public Property Access As String


		''' <summary>
		''' True if expression column needs to be validated to get its datatype, false otherwise.
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ValidateExpressionDataType As Boolean

		''' <summary>
		''' Return true if the datatype of column/expression is numeric, else false.
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property IsNumeric As Boolean
			Get
				' Check if the data type of the caculation/Expression is numeric or not whilst saving the report.
				If ValidateExpressionDataType Then

					Dim objCalc As clsExprCalculation = New clsExprCalculation(_objSessionInfo)
					objCalc.CalculationID = ID

					'Return true if any of the column datatype used in the calculation/expression as numeric, else false.
					Return (objCalc.ReturnType() = ColumnDataType.sqlNumeric)

				Else
					Return DataType = ColumnDataType.sqlInteger OrElse DataType = ColumnDataType.sqlNumeric
				End If
			End Get
		End Property

		Private ReadOnly Property _objSessionInfo As SessionInfo
			Get
				Return CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			End Get
		End Property
        Public Function Clone() As ReportColumnItem
            Return DirectCast(Me.MemberwiseClone, ReportColumnItem)
        End Function
    End Class
End Namespace