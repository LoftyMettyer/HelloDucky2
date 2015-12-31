Option Strict On
Option Explicit On

Imports System.Collections.ObjectModel
Imports HR.Intranet.Server.Metadata

Friend Class clsMatchRelation
	
	Private lngTable1ID As Integer
	Private strTable1Name As String
	Private strTable1RealSource As String
	Private lngTable2ID As Integer
	Private strTable2Name As String
	Private strTable2RealSource As String
	Private lngRequiredExprID As Integer
	Private lngPreferredExprID As Integer
	Private lngMatchScoreExprID As Integer
	Private colBreakdownColumns As Collection(Of DisplayColumn)
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		colBreakdownColumns = New Collection(Of DisplayColumn)
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object colBreakdownColumns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		colBreakdownColumns = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public Property Table1ID() As Integer
		Get
			Table1ID = lngTable1ID
		End Get
		Set(ByVal Value As Integer)
			lngTable1ID = Value
		End Set
	End Property
	
	
	Public Property Table1Name() As String
		Get
			Table1Name = strTable1Name
		End Get
		Set(ByVal Value As String)
			strTable1Name = Value
		End Set
	End Property
	
	
	Public Property Table1RealSource() As String
		Get
			Table1RealSource = strTable1RealSource
		End Get
		Set(ByVal Value As String)
			strTable1RealSource = Value
		End Set
	End Property
	
	
	Public Property Table2ID() As Integer
		Get
			Table2ID = lngTable2ID
		End Get
		Set(ByVal Value As Integer)
			lngTable2ID = Value
		End Set
	End Property
	
	
	Public Property Table2Name() As String
		Get
			Table2Name = strTable2Name
		End Get
		Set(ByVal Value As String)
			strTable2Name = Value
		End Set
	End Property
	
	
	Public Property Table2RealSource() As String
		Get
			Table2RealSource = strTable2RealSource
		End Get
		Set(ByVal Value As String)
			strTable2RealSource = Value
		End Set
	End Property
	
	
	Public Property RequiredExprID() As Integer
		Get
			RequiredExprID = lngRequiredExprID
		End Get
		Set(ByVal Value As Integer)
			lngRequiredExprID = Value
		End Set
	End Property
	
	
	Public Property PreferredExprID() As Integer
		Get
			PreferredExprID = lngPreferredExprID
		End Get
		Set(ByVal Value As Integer)
			lngPreferredExprID = Value
		End Set
	End Property
	
	
	Public Property MatchScoreID() As Integer
		Get
			MatchScoreID = lngMatchScoreExprID
		End Get
		Set(ByVal Value As Integer)
			lngMatchScoreExprID = Value
		End Set
	End Property
	
	
	Public Property BreakdownColumns() As Collection(Of DisplayColumn)
		Get
			BreakdownColumns = colBreakdownColumns
		End Get
		Set(ByVal Value As Collection(Of DisplayColumn))
			colBreakdownColumns = Value
		End Set
	End Property
End Class