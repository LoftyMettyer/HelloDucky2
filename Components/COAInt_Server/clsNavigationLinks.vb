Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsNavigationLinks_NET.clsNavigationLinks")> Public Class clsNavigationLinks
	
	Public Enum NavigationLinkType
		intHYPERLINK = 0
		intBUTTON = 1
		intDROPDOWNLINK = 2
		intDOCUMENTDISPLAY = 3
	End Enum
	
	Private mlngSSITableID As Integer
	Private mlngSSIViewID As Integer
	
	Public WriteOnly Property SSITableID() As Integer
		Set(ByVal Value As Integer)
			
			If Value <> mlngSSITableID Then
				ClearLinks()
			End If
			mlngSSITableID = Value
			
		End Set
	End Property
	
	Public WriteOnly Property SSIViewID() As Integer
		Set(ByVal Value As Integer)
			
			If Value <> mlngSSIViewID Then
				ClearLinks()
			End If
			mlngSSIViewID = Value
			
		End Set
	End Property
	
	Public WriteOnly Property Connection() As Object
		Set(ByVal Value As Object)
			
			' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
			If ASRDEVELOPMENT Then
				
				gADOCon = New ADODB.Connection
				
				'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gADOCon.Open(Value)
				
				CreateASRDev_SysProtects(gADOCon)
				
			Else
				gADOCon = Value
				
			End If
			
		End Set
	End Property
	
	Public Sub ClearLinks()
		
		'UPGRADE_NOTE: Object gcolLinks may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		gcolLinks = Nothing
		'UPGRADE_NOTE: Object gcolNavigationLinks may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		gcolNavigationLinks = Nothing
		
	End Sub
	
	' Loads all of the links and documents for this user session
	Public Sub LoadLinks()
		
		Dim rsLinks As ADODB.Recordset
		Dim sSQL As String
		Dim objLink As clsNavigationLink
		
		If Not gcolLinks Is Nothing Then
			Exit Sub
		End If
		
		gcolLinks = New Collection
		
		sSQL = "EXEC spASRIntGetLinks " & mlngSSITableID & ", " & mlngSSIViewID
		rsLinks = New ADODB.Recordset
		rsLinks.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		
		With rsLinks
			Do While Not (.EOF Or .BOF)
				objLink = New clsNavigationLink
				
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.BaseTable = IIf(Not IsDbNull(.Fields("BaseTable").Value), .Fields("BaseTable").Value, "")
				objLink.ID = .Fields("ID").Value
				objLink.DrillDownHidden = .Fields("DrillDownHidden").Value
				objLink.LinkOrder = .Fields("LinkOrder").Value
				objLink.LinkType = .Fields("LinkType").Value
				objLink.NewWindow = .Fields("NewWindow").Value
				objLink.PageTitle = .Fields("PageTitle").Value
				objLink.Prompt = .Fields("Prompt").Value
				objLink.ScreenID = .Fields("ScreenID").Value
				objLink.Text = .Fields("Text").Value
				objLink.URL = .Fields("URL").Value
				objLink.UtilityID = .Fields("UtilityID").Value
				objLink.UtilityType = .Fields("UtilityType").Value
				objLink.EmailAddress = .Fields("EmailAddress").Value
				objLink.EmailSubject = .Fields("EmailSubject").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.AppFilePath = IIf(IsDbNull(.Fields("AppFilePath").Value), "", .Fields("AppFilePath").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.AppParameters = IIf(IsDbNull(.Fields("AppParameters").Value), "", .Fields("AppParameters").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.DocumentFilePath = IIf(IsDbNull(.Fields("DocumentFilePath").Value), "", .Fields("DocumentFilePath").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.DisplayDocumentHyperlink = IIf(IsDbNull(.Fields("DisplayDocumentHyperlink").Value), False, .Fields("DisplayDocumentHyperlink").Value)
				' objLink.IsSeparator = IIf(IsNull(.Fields("IsSeparator").Value), False, .Fields("IsSeparator").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.SeparatorOrientation = IIf(IsDbNull(.Fields("SeparatorOrientation").Value), 0, .Fields("SeparatorOrientation").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.PictureID = IIf(IsDbNull(.Fields("PictureID").Value), 0, .Fields("PictureID").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ShowLegend = IIf(IsDbNull(.Fields("Chart_ShowLegend").Value), False, .Fields("Chart_ShowLegend").Value)
				objLink.Chart_Type = .Fields("Chart_Type").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ShowGrid = IIf(IsDbNull(.Fields("Chart_ShowGrid").Value), False, .Fields("Chart_ShowGrid").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_StackSeries = IIf(IsDbNull(.Fields("Chart_StackSeries").Value), False, .Fields("Chart_StackSeries").Value)
				objLink.Chart_ViewID = .Fields("Chart_ViewID").Value
				objLink.Chart_TableID = .Fields("Chart_TableID").Value
				objLink.Chart_ColumnID = .Fields("Chart_ColumnID").Value
				objLink.Chart_FilterID = .Fields("Chart_FilterID").Value
				objLink.Chart_AggregateType = .Fields("Chart_AggregateType").Value
				objLink.Element_Type = .Fields("Element_Type").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ShowValues = IIf(IsDbNull(.Fields("Chart_ShowValues").Value), False, .Fields("Chart_ShowValues").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.UseFormatting = IIf(IsDbNull(.Fields("UseFormatting").Value), False, .Fields("UseFormatting").Value)
				objLink.Formatting_DecimalPlaces = .Fields("Formatting_DecimalPlaces").Value
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Formatting_Use1000Separator = IIf(IsDbNull(.Fields("Formatting_Use1000Separator").Value), False, .Fields("Formatting_Use1000Separator").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Formatting_Prefix = IIf(IsDbNull(.Fields("Formatting_Prefix").Value), "", .Fields("Formatting_Prefix").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Formatting_Suffix = IIf(IsDbNull(.Fields("Formatting_Suffix").Value), "", .Fields("Formatting_Suffix").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.UseConditionalFormatting = IIf(IsDbNull(.Fields("UseConditionalFormatting").Value), False, .Fields("UseConditionalFormatting").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Operator_1 = IIf(IsDbNull(.Fields("ConditionalFormatting_Operator_1").Value), "", .Fields("ConditionalFormatting_Operator_1").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Value_1 = IIf(IsDbNull(.Fields("ConditionalFormatting_Value_1").Value), "", .Fields("ConditionalFormatting_Value_1").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Style_1 = IIf(IsDbNull(.Fields("ConditionalFormatting_Style_1").Value), "", .Fields("ConditionalFormatting_Style_1").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Colour_1 = IIf(IsDbNull(.Fields("ConditionalFormatting_Colour_1").Value), "", .Fields("ConditionalFormatting_Colour_1").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Operator_2 = IIf(IsDbNull(.Fields("ConditionalFormatting_Operator_2").Value), "", .Fields("ConditionalFormatting_Operator_2").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Value_2 = IIf(IsDbNull(.Fields("ConditionalFormatting_Value_2").Value), "", .Fields("ConditionalFormatting_Value_2").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Style_2 = IIf(IsDbNull(.Fields("ConditionalFormatting_Style_2").Value), "", .Fields("ConditionalFormatting_Style_2").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Colour_2 = IIf(IsDbNull(.Fields("ConditionalFormatting_Colour_2").Value), "", .Fields("ConditionalFormatting_Colour_2").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Operator_3 = IIf(IsDbNull(.Fields("ConditionalFormatting_Operator_3").Value), "", .Fields("ConditionalFormatting_Operator_3").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Value_3 = IIf(IsDbNull(.Fields("ConditionalFormatting_Value_3").Value), "", .Fields("ConditionalFormatting_Value_3").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Style_3 = IIf(IsDbNull(.Fields("ConditionalFormatting_Style_3").Value), "", .Fields("ConditionalFormatting_Style_3").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Colour_3 = IIf(IsDbNull(.Fields("ConditionalFormatting_Colour_3").Value), "", .Fields("ConditionalFormatting_Colour_3").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.SeparatorColour = IIf(IsDbNull(.Fields("SeparatorColour").Value), "", .Fields("SeparatorColour").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColumnName = IIf(IsDbNull(.Fields("Chart_ColumnName").Value), "", .Fields("Chart_ColumnName").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColumnName_2 = datGeneral.GetColumnName(CInt(IIf(IsDbNull(.Fields("Chart_ColumnID_2").Value), 0, .Fields("Chart_ColumnID_2").Value)))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.InitialDisplayMode = IIf(IsDbNull(.Fields("InitialDisplayMode").Value), 0, .Fields("InitialDisplayMode").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_TableID_2 = IIf(IsDbNull(.Fields("Chart_TableID_2").Value), 0, .Fields("Chart_TableID_2").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColumnID_2 = IIf(IsDbNull(.Fields("Chart_ColumnID_2").Value), 0, .Fields("Chart_ColumnID_2").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_TableID_3 = IIf(IsDbNull(.Fields("Chart_TableID_3").Value), 0, .Fields("Chart_TableID_3").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColumnID_3 = IIf(IsDbNull(.Fields("Chart_ColumnID_3").Value), 0, .Fields("Chart_ColumnID_3").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_SortOrderID = IIf(IsDbNull(.Fields("Chart_SortOrderID").Value), 0, .Fields("Chart_SortOrderID").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_SortDirection = IIf(IsDbNull(.Fields("Chart_SortDirection").Value), 0, .Fields("Chart_SortDirection").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColourID = IIf(IsDbNull(.Fields("Chart_ColourID").Value), 0, .Fields("Chart_ColourID").Value)
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ShowPercentages = IIf(IsDbNull(.Fields("Chart_ShowPercentages").Value), False, .Fields("Chart_ShowPercentages").Value)
				gcolLinks.Add(objLink)
				.MoveNext()
			Loop 
		End With
		
		rsLinks.Close()
		'UPGRADE_NOTE: Object rsLinks may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLinks = Nothing
		
	End Sub
	
	' Loads all of the navigation links for this user session
	Public Sub LoadNavigationLinks()
		
		Dim rsLinks As ADODB.Recordset
		Dim sSQL As String
		Dim objLink As clsNavigationLink
		
		If Not gcolNavigationLinks Is Nothing Then
			Exit Sub
		End If
		
		gcolNavigationLinks = New Collection
		
		sSQL = "EXEC spASRIntGetNavigationLinks " & mlngSSITableID & ", " & mlngSSIViewID
		rsLinks = New ADODB.Recordset
		rsLinks.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		
		With rsLinks
			Do While Not (.EOF Or .BOF)
				objLink = New clsNavigationLink
				
				objLink.LinkType = .Fields("LinkType").Value
				objLink.Text1 = .Fields("Text1").Value
				objLink.Text2 = .Fields("Text2").Value
				objLink.SingleRecord = .Fields("SingleRecord").Value
				objLink.LinkToFind = .Fields("LinkToFind").Value
				objLink.TableID = .Fields("TableID").Value
				objLink.ViewID = .Fields("ViewID").Value
				objLink.PrimarySequence = .Fields("PrimarySequence").Value
				objLink.SecondarySequence = .Fields("SecondarySequence").Value
				objLink.FindPage = .Fields("FindPage").Value
				gcolNavigationLinks.Add(objLink)
				.MoveNext()
			Loop 
		End With
		
		rsLinks.Close()
		'UPGRADE_NOTE: Object rsLinks may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLinks = Nothing
		
	End Sub
	
	Public Function GetNavigationLinks(ByRef piLinkType As NavigationLinkType, ByRef pbShowFindPages As Boolean) As Collection
		
		Dim objLink As clsNavigationLink
		Dim objLinks As Collection
		
		objLinks = New Collection
		
		For	Each objLink In gcolNavigationLinks
			If objLink.LinkType = piLinkType And (objLink.FindPage = pbShowFindPages Or pbShowFindPages) Then
				objLinks.Add(objLink)
			End If
		Next objLink
		
		GetNavigationLinks = objLinks
		
	End Function
	
  Public Function GetLinks(ByRef piLinkType As NavigationLinkType) As Collection

    Dim objLink As clsNavigationLink
    Dim objLinks As Collection

    objLinks = New Collection

    For Each objLink In gcolLinks
      If objLink.LinkType = piLinkType Then
        objLinks.Add(objLink)
      End If
    Next objLink

    GetLinks = objLinks

  End Function
	
	Public Function GetDocuments(ByRef piLinkType As NavigationLinkType) As Collection
		
		Dim objLink As clsNavigationLink
		Dim objLinks As Collection
		
		objLinks = New Collection
		
		For	Each objLink In gcolLinks
			If objLink.LinkType = piLinkType Then
				objLinks.Add(objLink)
			End If
		Next objLink
		
		GetDocuments = objLinks
		
	End Function
End Class