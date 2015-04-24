Option Strict Off
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports System.Data.SqlClient

Public Class clsNavigationLinks
	Inherits BaseForDMI

	Private _colLinks As List(Of Link)
	Private _colNavigationLinks As List(Of Link)

	Private mlngSSITableID As Integer
	Private mlngSSIViewID As Integer

	Public WriteOnly Property SSITableID() As Integer
		Set(Value As Integer)

			If Value <> mlngSSITableID Then
				ClearLinks()
			End If
			mlngSSITableID = Value

		End Set
	End Property

	Public WriteOnly Property SSIViewID() As Integer
		Set(Value As Integer)

			If Value <> mlngSSIViewID Then
				ClearLinks()
			End If
			mlngSSIViewID = Value

		End Set
	End Property

	Public Sub ClearLinks()

		_colLinks = Nothing
		_colNavigationLinks = Nothing

	End Sub

	Public ReadOnly Property ColLinks() As List(Of Link)
		Get
			Return _colLinks
		End Get
	End Property

	' Loads all of the links and documents for this user session
	Public Sub LoadLinks()

		Dim objLink As Link
		_colLinks = New List(Of Link)

		Dim prmTableID = New SqlParameter("plngTableID", SqlDbType.Int)
		Dim prmViewID = New SqlParameter("plngViewID", SqlDbType.Int)

		prmTableID.Value = mlngSSITableID
		prmViewID.Value = mlngSSIViewID

		Using rsLinks = DB.GetDataTable("spASRIntGetLinks", CommandType.StoredProcedure, prmTableID, prmViewID)

			For Each objRow As DataRow In rsLinks.Rows

				objLink = New Link

				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.BaseTableID = IIf(Not IsDBNull(objRow("BaseTable")), objRow("BaseTable"), 0)
				objLink.ID = CInt(objRow("ID"))
				objLink.DrillDownHidden = CBool(objRow("DrillDownHidden"))
				objLink.LinkOrder = CShort(objRow("LinkOrder"))
				objLink.LinkType = CShort(objRow("LinkType"))
				objLink.NewWindow = CBool(objRow("NewWindow"))
				objLink.PageTitle = objRow("PageTitle").ToString()
				objLink.Prompt = objRow("Prompt").ToString()
				objLink.ScreenID = CInt(objRow("ScreenID"))
				objLink.Text = objRow("Text").ToString()
				objLink.URL = objRow("URL").ToString()
				objLink.UtilityID = CInt(objRow("UtilityID"))
				objLink.UtilityType = CShort(objRow("UtilityType"))
				objLink.EmailAddress = objRow("EmailAddress").ToString()
				objLink.EmailSubject = objRow("EmailSubject").ToString()
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.AppFilePath = IIf(IsDBNull(objRow("AppFilePath")), "", objRow("AppFilePath"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.AppParameters = IIf(IsDBNull(objRow("AppParameters")), "", objRow("AppParameters"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.DocumentFilePath = IIf(IsDBNull(objRow("DocumentFilePath")), "", objRow("DocumentFilePath"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.DisplayDocumentHyperlink = IIf(IsDBNull(objRow("DisplayDocumentHyperlink")), False, objRow("DisplayDocumentHyperlink"))
				' objLink.IsSeparator = IIf(IsNull(objRow("IsSeparator")), False, objRow("IsSeparator"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.SeparatorOrientation = IIf(IsDBNull(objRow("SeparatorOrientation")), 0, objRow("SeparatorOrientation"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.PictureID = IIf(IsDBNull(objRow("PictureID")), 0, objRow("PictureID"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ShowLegend = IIf(IsDBNull(objRow("Chart_ShowLegend")), False, objRow("Chart_ShowLegend"))
				objLink.Chart_Type = objRow("Chart_Type")
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ShowGrid = IIf(IsDBNull(objRow("Chart_ShowGrid")), False, objRow("Chart_ShowGrid"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_StackSeries = IIf(IsDBNull(objRow("Chart_StackSeries")), False, objRow("Chart_StackSeries"))
				objLink.Chart_ViewID = CInt(objRow("Chart_ViewID"))
				objLink.Chart_TableID = CInt(objRow("Chart_TableID"))
				objLink.Chart_ColumnID = CInt(objRow("Chart_ColumnID"))
				objLink.Chart_FilterID = CInt(objRow("Chart_FilterID"))
				objLink.Chart_AggregateType = objRow("Chart_AggregateType")
				objLink.Element_Type = objRow("Element_Type")
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ShowValues = IIf(IsDBNull(objRow("Chart_ShowValues")), False, objRow("Chart_ShowValues"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.UseFormatting = IIf(IsDBNull(objRow("UseFormatting")), False, objRow("UseFormatting"))
				objLink.Formatting_DecimalPlaces = objRow("Formatting_DecimalPlaces")
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Formatting_Use1000Separator = IIf(IsDBNull(objRow("Formatting_Use1000Separator")), False, objRow("Formatting_Use1000Separator"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Formatting_Prefix = IIf(IsDBNull(objRow("Formatting_Prefix")), "", objRow("Formatting_Prefix"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Formatting_Suffix = IIf(IsDBNull(objRow("Formatting_Suffix")), "", objRow("Formatting_Suffix"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.UseConditionalFormatting = IIf(IsDBNull(objRow("UseConditionalFormatting")), False, objRow("UseConditionalFormatting"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Operator_1 = IIf(IsDBNull(objRow("ConditionalFormatting_Operator_1")), "", objRow("ConditionalFormatting_Operator_1"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Value_1 = IIf(IsDBNull(objRow("ConditionalFormatting_Value_1")), "", objRow("ConditionalFormatting_Value_1"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Style_1 = IIf(IsDBNull(objRow("ConditionalFormatting_Style_1")), "", objRow("ConditionalFormatting_Style_1"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Colour_1 = IIf(IsDBNull(objRow("ConditionalFormatting_Colour_1")), "", objRow("ConditionalFormatting_Colour_1"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Operator_2 = IIf(IsDBNull(objRow("ConditionalFormatting_Operator_2")), "", objRow("ConditionalFormatting_Operator_2"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Value_2 = IIf(IsDBNull(objRow("ConditionalFormatting_Value_2")), "", objRow("ConditionalFormatting_Value_2"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Style_2 = IIf(IsDBNull(objRow("ConditionalFormatting_Style_2")), "", objRow("ConditionalFormatting_Style_2"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Colour_2 = IIf(IsDBNull(objRow("ConditionalFormatting_Colour_2")), "", objRow("ConditionalFormatting_Colour_2"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Operator_3 = IIf(IsDBNull(objRow("ConditionalFormatting_Operator_3")), "", objRow("ConditionalFormatting_Operator_3"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Value_3 = IIf(IsDBNull(objRow("ConditionalFormatting_Value_3")), "", objRow("ConditionalFormatting_Value_3"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Style_3 = IIf(IsDBNull(objRow("ConditionalFormatting_Style_3")), "", objRow("ConditionalFormatting_Style_3"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.ConditionalFormatting_Colour_3 = IIf(IsDBNull(objRow("ConditionalFormatting_Colour_3")), "", objRow("ConditionalFormatting_Colour_3"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.SeparatorColour = IIf(IsDBNull(objRow("SeparatorColour")), "", objRow("SeparatorColour"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColumnName = IIf(IsDBNull(objRow("Chart_ColumnName")), "", objRow("Chart_ColumnName"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColumnName_2 = GetColumnName(CInt(IIf(IsDBNull(objRow("Chart_ColumnID_2")), 0, objRow("Chart_ColumnID_2"))))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.InitialDisplayMode = IIf(IsDBNull(objRow("InitialDisplayMode")), 0, objRow("InitialDisplayMode"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_TableID_2 = IIf(IsDBNull(objRow("Chart_TableID_2")), 0, objRow("Chart_TableID_2"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColumnID_2 = IIf(IsDBNull(objRow("Chart_ColumnID_2")), 0, objRow("Chart_ColumnID_2"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_TableID_3 = IIf(IsDBNull(objRow("Chart_TableID_3")), 0, objRow("Chart_TableID_3"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColumnID_3 = IIf(IsDBNull(objRow("Chart_ColumnID_3")), 0, objRow("Chart_ColumnID_3"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_SortOrderID = IIf(IsDBNull(objRow("Chart_SortOrderID")), 0, objRow("Chart_SortOrderID"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_SortDirection = IIf(IsDBNull(objRow("Chart_SortDirection")), 0, objRow("Chart_SortDirection"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ColourID = IIf(IsDBNull(objRow("Chart_ColourID")), 0, objRow("Chart_ColourID"))
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				objLink.Chart_ShowPercentages = IIf(IsDBNull(objRow("Chart_ShowPercentages")), False, objRow("Chart_ShowPercentages"))
				_colLinks.Add(objLink)
			Next

		End Using
	End Sub

	' Loads all of the navigation links for this user session
	Public Sub LoadNavigationLinks()

		Dim objLink As Link

		_colNavigationLinks = New List(Of Link)

		Dim prmTableID = New SqlParameter("plngTableID", SqlDbType.Int)
		Dim prmViewID = New SqlParameter("plngViewID", SqlDbType.Int)

		prmTableID.Value = mlngSSITableID
		prmViewID.Value = mlngSSIViewID

		Using rsLinks = DB.GetDataTable("spASRIntGetNavigationLinks", CommandType.StoredProcedure, prmTableID, prmViewID)

			For Each objRow As DataRow In rsLinks.Rows
				objLink = New Link
				objLink.LinkType = CShort(objRow("LinkType"))
				objLink.Text1 = objRow("Text1").ToString()
				objLink.Text2 = objRow("Text2").ToString()
				objLink.SingleRecord = CShort(objRow("SingleRecord"))
				objLink.LinkToFind = CShort(objRow("LinkToFind"))
				objLink.TableID = CInt(objRow("TableID"))
				objLink.ViewID = CInt(objRow("ViewID"))
				objLink.PrimarySequence = CShort(objRow("PrimarySequence"))
				objLink.SecondarySequence = CShort(objRow("SecondarySequence"))
				objLink.FindPage = CShort(objRow("FindPage"))
				_colNavigationLinks.Add(objLink)

			Next
		End Using

	End Sub

	Public Function GetNavigationLinks(pbShowFindPages As Boolean, piLinkType As LinkType) As List(Of Link)
		Return _colNavigationLinks.FindAll(Function(n) (n.FindPage = pbShowFindPages Or pbShowFindPages) And n.LinkType = piLinkType)
	End Function

	Public Function NavigationLinks(pbShowFindPages As Boolean) As List(Of Link)
		Return _colNavigationLinks
	End Function

	Public Function GetLinks(piLinkType As LinkType) As List(Of Link)
		Return _colLinks.FindAll(Function(n) n.LinkType = piLinkType)
	End Function

	Public Function GetAllLinks() As List(Of Link)
		Return _colLinks
	End Function

End Class