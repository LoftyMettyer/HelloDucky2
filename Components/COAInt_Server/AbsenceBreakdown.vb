Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("AbsenceBreakdown_NET.AbsenceBreakdown")> Public Class AbsenceBreakdown
	
	Private mlngPersonnelRecordID As Integer
	
	Private mstrHexColour_BoxBackground As String
	
	Public Function HTML_RecordSelection() As Object
		
		Dim strHTML As String
		Dim iBoxWidth As Short
		
		strHTML = ""
		
		'UPGRADE_WARNING: Couldn't resolve default property of object HTML_RecordSelection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HTML_RecordSelection = strHTML
		
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		' Define the base colours
		mstrHexColour_BoxBackground = "ThreeDFace"
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public WriteOnly Property RecordID() As Object
		Set(ByVal Value As Object)
			
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsNumeric(Value) And Not IsNothing(Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object piRecordID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mlngPersonnelRecordID = Value
			End If
			
		End Set
	End Property
	Public WriteOnly Property Connection() As Object
		Set(ByVal Value As Object)
			
			' Connection object passed in from the asp page
			
			' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
			If ASRDEVELOPMENT Then
				gADOCon = New ADODB.Connection
				'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gADOCon.Open(Value)
			Else
				gADOCon = Value
			End If
			
		End Set
	End Property
End Class