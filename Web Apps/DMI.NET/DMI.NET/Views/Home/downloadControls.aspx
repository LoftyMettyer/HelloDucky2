<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
	<%= GetPageTitle("downloadControls")%>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

	<%@  Language="VBScript" %>
	<html>
	<head>
		<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" type="text/css" href="OpenHR.css">
		<title>OpenHR Intranet</title>
		<meta http-equiv="X-UA-Compatible" content="IE=5">
		<!-- LPK file -->
		<object
			classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
			id="Microsoft_Licensed_Class_Manager_1_0"
			viewastext>
			<param name="LPKPath" value="lpks/main.lpk">
		</object>

		<!-- Client-side Intranet general functions DLL -->
		<!--#include file="include\ctl_ASRIntranetFunctions.txt"-->

		<!-- Client-side Intranet print functions DLL -->
		<!--#include file="include\ctl_ASRIntranetPrintFunctions.txt"-->

		<!-- Client-side Intranet output functions DLL -->
		<!--#include file="include\ctl_ASRIntranetOutput.txt"-->

		<!-- Codejock controls -->
		<object
			id="ctlCodeJock_PushButton"
			classid="CLSID:3E8187B5-3C15-4233-B811-9CB3F929A28B"
			codebase="cabs/COAInt_Client.cab#version=13,1,0,0"
			style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			viewastext>
			<param name="_ExtentX" value="3307">
			<param name="_ExtentY" value="1323">
		</object>

		<!-- Grid control -->
		<object
			classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
			codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
			id="ssOleDBGrid"
			name="ssOleDBGrid"
			style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"
			viewastext>
			<param name="ScrollBars" value="4">
			<param name="_Version" value="196617">
			<param name="DataMode" value="2">
			<param name="Cols" value="0">
			<param name="Rows" value="0">
			<param name="BorderStyle" value="1">
			<param name="RecordSelectors" value="0">
			<param name="GroupHeaders" value="0">
			<param name="ColumnHeaders" value="1">
			<param name="GroupHeadLines" value="1">
			<param name="HeadLines" value="1">
			<param name="FieldDelimiter" value="(None)">
			<param name="FieldSeparator" value="(Tab)">
			<param name="Row.Count" value="0">
			<param name="Col.Count" value="1">
			<param name="stylesets.count" value="0">
			<param name="TagVariant" value="EMPTY">
			<param name="UseGroups" value="0">
			<param name="HeadFont3D" value="0">
			<param name="Font3D" value="0">
			<param name="DividerType" value="3">
			<param name="DividerStyle" value="1">
			<param name="DefColWidth" value="0">
			<param name="BeveColorScheme" value="2">
			<param name="BevelColorFrame" value="-2147483642">
			<param name="BevelColorHighlight" value="-2147483628">
			<param name="BevelColorShadow" value="-2147483632">
			<param name="BevelColorFace" value="-2147483633">
			<param name="CheckBox3D" value="-1">
			<param name="AllowAddNew" value="0">
			<param name="AllowDelete" value="0">
			<param name="AllowUpdate" value="0">
			<param name="MultiLine" value="0">
			<param name="ActiveCellStyleSet" value="">
			<param name="RowSelectionStyle" value="0">
			<param name="AllowRowSizing" value="0">
			<param name="AllowGroupSizing" value="0">
			<param name="AllowColumnSizing" value="-1">
			<param name="AllowGroupMoving" value="0">
			<param name="AllowColumnMoving" value="0">
			<param name="AllowGroupSwapping" value="0">
			<param name="AllowColumnSwapping" value="0">
			<param name="AllowGroupShrinking" value="0">
			<param name="AllowColumnShrinking" value="0">
			<param name="AllowDragDrop" value="0">
			<param name="UseExactRowCount" value="-1">
			<param name="SelectTypeCol" value="0">
			<param name="SelectTypeRow" value="1">
			<param name="SelectByCell" value="-1">
			<param name="BalloonHelp" value="0">
			<param name="RowNavigation" value="1">
			<param name="CellNavigation" value="0">
			<param name="MaxSelectedRows" value="1">
			<param name="HeadStyleSet" value="">
			<param name="StyleSet" value="">
			<param name="ForeColorEven" value="0">
			<param name="ForeColorOdd" value="0">
			<param name="BackColorEven" value="16777215">
			<param name="BackColorOdd" value="16777215">
			<param name="Levels" value="1">
			<param name="RowHeight" value="503">
			<param name="ExtraHeight" value="0">
			<param name="ActiveRowStyleSet" value="">
			<param name="CaptionAlignment" value="2">
			<param name="SplitterPos" value="0">
			<param name="SplitterVisible" value="0">
			<param name="Columns.Count" value="1">
			<param name="Columns(0).Width" value="6500">
			<param name="Columns(0).Visible" value="-1">
			<param name="Columns(0).Columns.Count" value="1">
			<param name="Columns(0).Caption" value="">
			<param name="Columns(0).Name" value="">
			<param name="Columns(0).Alignment" value="0">
			<param name="Columns(0).CaptionAlignment" value="3">
			<param name="Columns(0).Bound" value="0">
			<param name="Columns(0).AllowSizing" value="1">
			<param name="Columns(0).DataField" value="Column 0">
			<param name="Columns(0).DataType" value="8">
			<param name="Columns(0).Level" value="0">
			<param name="Columns(0).NumberFormat" value="">
			<param name="Columns(0).Case" value="0">
			<param name="Columns(0).FieldLen" value="256">
			<param name="Columns(0).VertScrollBar" value="0">
			<param name="Columns(0).Locked" value="0">
			<param name="Columns(0).Style" value="0">
			<param name="Columns(0).ButtonsAlways" value="0">
			<param name="Columns(0).RowCount" value="0">
			<param name="Columns(0).ColCount" value="1">
			<param name="Columns(0).HasHeadForeColor" value="0">
			<param name="Columns(0).HasHeadBackColor" value="0">
			<param name="Columns(0).HasForeColor" value="0">
			<param name="Columns(0).HasBackColor" value="0">
			<param name="Columns(0).HeadForeColor" value="0">
			<param name="Columns(0).HeadBackColor" value="0">
			<param name="Columns(0).ForeColor" value="0">
			<param name="Columns(0).BackColor" value="0">
			<param name="Columns(0).HeadStyleSet" value="">
			<param name="Columns(0).StyleSet" value="">
			<param name="Columns(0).Nullable" value="1">
			<param name="Columns(0).Mask" value="">
			<param name="Columns(0).PromptInclude" value="0">
			<param name="Columns(0).ClipMode" value="0">
			<param name="Columns(0).PromptChar" value="95">
			<param name="UseDefaults" value="-1">
			<param name="TabNavigation" value="1">
			<param name="BatchUpdate" value="0">
			<param name="_ExtentX" value="16087">
			<param name="_ExtentY" value="4630">
			<param name="_StockProps" value="79">
			<param name="Caption" value="">
			<param name="ForeColor" value="0">
			<param name="BackColor" value="16777215">
			<param name="Enabled" value="-1">
			<param name="DataMember" value="">
		</object>


		<!--Intranet Lookup Control -->
		<object
			id="ctlIntLookup1"
			classid="CLSID:F513059D-1E77-4F0A-8681-21B2BA7E5986"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,3"
			style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			viewastext>
			<param name="_ExtentX" value="4233">
			<param name="_ExtentY" value="953">
			<param name="BackColor" value="16777215">
			<param name="ForeColor" value="0">
			<param name="Enabled" value="-1">
			<param name="Text" value="">
		</object>

		<!-- New Option Group 3 -->
		<object
			id="ctlNewOptionGroup3"
			classid="CLSID:0B21DBA2-E1A7-47C7-B62F-BC522820590A"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,2"
			style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			viewastext>
			<param name="_ExtentX" value="3307">
			<param name="_ExtentY" value="1323">
		</object>

		<!-- Image control -->
		<object
			id="ASRUserImage1"
			classid="CLSID:8FF15C8D-49D5-4B79-8419-C36C26654283"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,7"
			viewastext>
			<param name="_ExtentX" value="2619">
			<param name="_ExtentY" value="2619">
			<param name="ForeColor" value="0">
			<param name="Enabled" value="-1">
			<param name="BorderStyle" value="0">
			<param name="ASRDataField" value="0">
		</object>

		<!-- Menu control (base) -->
		<object
			classid="clsid:E4F874A0-56ED-11D0-9C43-00A0C90F29FC"
			codebase="cabs/COAInt_Client.cab#Version=1,0,6,5"
			height="32"
			id="abMainMenuBase"
			name="abMainMenuBase"
			style="LEFT: 0px; TOP: 0px"
			width="32"
			viewastext>
		</object>

		<%If Request.QueryString("fromMenu") = "false" Then%>
		<!--NPG20111014 Fault HRPRO-1799, don't reload the menu control if already loaded.-->
		<!-- Menu control -->
		<object
			classid="clsid:6976CB54-C39B-4181-B1DC-1A829068E2E7"
			codebase="cabs/COAInt_Client.cab#Version=1,0,0,5"
			id="abMainMenu"
			name="abMainMenu"
			style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			viewastext>
		</object>
		<%End If%>

		<!-- Mask control -->
		<object
			classid="clsid:66A90C04-346D-11D2-9BC0-00A024695830"
			codebase="cabs/COAInt_Client.cab#version=6,0,1,1"
			id="TDBMask1"
			viewastext>
		</object>

		<!-- Number control -->
		<object
			classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
			codebase="cabs/COAInt_Client.cab#version=6,0,1,1"
			id="TDBNumber1"
			viewastext>
		</object>

		<!-- ASR Spinner control -->
		<object
			classid="CLSID:C25C3704-2AA7-44E5-943A-B40B14E2348F"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,3"
			id="ASRSpinner1"
			style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			viewastext>
		</object>

		<!-- Date control -->
		<object
			classid="clsid:A49CE0E4-C0F9-11D2-B0EA-00A024695830"
			codebase="cabs/COAInt_Client.cab#version=6,0,1,1"
			id="TDBDate1"
			style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			viewastext>
		</object>

		<!-- Dialogue control -->
		<object
			classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB"
			codebase="cabs/COAInt_Client.cab#Version=1,0,0,0"
			id="dialog"
			style="LEFT: 0px; TOP: 0px"
			viewastext>
			<param name="_ExtentX" value="847">
			<param name="_ExtentY" value="847">
			<param name="_Version" value="393216">
			<param name="CancelError" value="0">
			<param name="Color" value="0">
			<param name="Copies" value="1">
			<param name="DefaultExt" value="">
			<param name="DialogTitle" value="">
			<param name="FileName" value="">
			<param name="Filter" value="">
			<param name="FilterIndex" value="0">
			<param name="Flags" value="0">
			<param name="FontBold" value="0">
			<param name="FontItalic" value="0">
			<param name="FontName" value="">
			<param name="FontSize" value="8">
			<param name="FontStrikeThru" value="0">
			<param name="FontUnderLine" value="0">
			<param name="FromPage" value="0">
			<param name="HelpCommand" value="0">
			<param name="HelpContext" value="0">
			<param name="HelpFile" value="">
			<param name="HelpKey" value="">
			<param name="InitDir" value="">
			<param name="Max" value="0">
			<param name="Min" value="0">
			<param name="MaxFileSize" value="260">
			<param name="PrinterDefault" value="1">
			<param name="ToPage" value="0">
			<param name="Orientation" value="1">
		</object>

		<!-- Tree control -->
		<object
			id="SSTree"
			classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B"
			codebase="cabs/COAInt_Client.cab#version=1,0,2,24"
			style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 100%"
			viewastext>
			<param name="_ExtentX" value="2646">
			<param name="_ExtentY" value="1323">
			<param name="_Version" value="65538">
			<param name="BackColor" value="-2147483643">
			<param name="ForeColor" value="-2147483640">
			<param name="ImagesMaskColor" value="12632256">
			<param name="PictureBackgroundMaskColor" value="12632256">
			<param name="Appearance" value="1">
			<param name="BorderStyle" value="0">
			<param name="LabelEdit" value="1">
			<param name="LineStyle" value="0">
			<param name="LineType" value="1">
			<param name="MousePointer" value="0">
			<param name="NodeSelectionStyle" value="2">
			<param name="PictureAlignment" value="0">
			<param name="ScrollStyle" value="0">
			<param name="Style" value="6">
			<param name="IndentationStyle" value="0">
			<param name="TreeTips" value="3">
			<param name="PictureBackgroundStyle" value="0">
			<param name="Indentation" value="38">
			<param name="MaxLines" value="1">
			<param name="TreeTipDelay" value="500">
			<param name="ImageCount" value="0">
			<param name="ImageListIndex" value="-1">
			<param name="OLEDragMode" value="0">
			<param name="OLEDropMode" value="0">
			<param name="AllowDelete" value="0">
			<param name="AutoSearch" value="0">
			<param name="Enabled" value="-1">
			<param name="HideSelection" value="0">
			<param name="ImagesUseMask" value="0">
			<param name="Redraw" value="-1">
			<param name="UseImageList" value="-1">
			<param name="PictureBackgroundUseMask" value="0">
			<param name="HasFont" value="0">
			<param name="HasMouseIcon" value="0">
			<param name="HasPictureBackground" value="0">
			<param name="PathSeparator" value="\">
			<param name="TabStops" value="32">
			<param name="ImageList" value="<None>">
			<param name="LoadStyleRoot" value="1">
			<param name="Sorted" value="0">
			<param name="OnDemandDiscardBuffer" value="10">
		</object>

		<object
			id="SSTree2"
			classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B"
			codebase="cabs/COAInt_Client.cab#version=1,0,2,24"
			style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 100%"
			viewastext>
			<param name="_ExtentX" value="2646">
			<param name="_ExtentY" value="1323">
			<param name="_Version" value="65538">
			<param name="BackColor" value="-2147483643">
			<param name="ForeColor" value="-2147483640">
			<param name="ImagesMaskColor" value="12632256">
			<param name="PictureBackgroundMaskColor" value="12632256">
			<param name="Appearance" value="1">
			<param name="BorderStyle" value="0">
			<param name="LabelEdit" value="1">
			<param name="LineStyle" value="0">
			<param name="LineType" value="1">
			<param name="MousePointer" value="0">
			<param name="NodeSelectionStyle" value="2">
			<param name="PictureAlignment" value="0">
			<param name="ScrollStyle" value="0">
			<param name="Style" value="6">
			<param name="IndentationStyle" value="0">
			<param name="TreeTips" value="3">
			<param name="PictureBackgroundStyle" value="0">
			<param name="Indentation" value="38">
			<param name="MaxLines" value="1">
			<param name="TreeTipDelay" value="500">
			<param name="ImageCount" value="0">
			<param name="ImageListIndex" value="-1">
			<param name="OLEDragMode" value="0">
			<param name="OLEDropMode" value="0">
			<param name="AllowDelete" value="0">
			<param name="AutoSearch" value="0">
			<param name="Enabled" value="-1">
			<param name="HideSelection" value="0">
			<param name="ImagesUseMask" value="0">
			<param name="Redraw" value="-1">
			<param name="UseImageList" value="-1">
			<param name="PictureBackgroundUseMask" value="0">
			<param name="HasFont" value="0">
			<param name="HasMouseIcon" value="0">
			<param name="HasPictureBackground" value="0">
			<param name="PathSeparator" value="\">
			<param name="TabStops" value="32">
			<param name="ImageList" value="<None>">
			<param name="LoadStyleRoot" value="1">
			<param name="Sorted" value="0">
			<param name="OnDemandDiscardBuffer" value="10">
		</object>


		<!-- Navigation control -->
		<object
			classid="CLSID:0F3914E5-4334-4E10-8DE8-538F9B828CD9"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,10"
			id="ctlNavigation"
			name="ctlNavigation"
			style="WIDTH: 100%; visibility: hidden; display: none"
			viewastext>
		</object>


		<!-- Calendar Reports Record control -->
		<object
			classid="CLSID:252D73AF-D7C6-4833-8539-A2C0293950B1"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,2"
			id="ctlCalRec"
			name="ctlCalRec"
			style="WIDTH: 100%"
			width="100%">
		</object>


		<!-- Calendar Reports Key control -->
		<object
			classid="CLSID:8E2F1EF1-3812-4678-A084-16384DE3EA6D"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,2"
			id="ctlKey"
			name="ctlKey"
			style="WIDTH: 100%; visibility: hidden; display: none"
			width="100%"
			height="85%">
		</object>

		<!-- Intranet OLE Embedded control -->
		<object
			id="ASRIntOLE1"
			name="ASRIntOLE1"
			classid="CLSID:5A5DF13B-2C9C-49E8-9C51-809E909F1123"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,15"
			style="WIDTH: 100%; visibility: hidden; display: none"
			viewastext>
		</object>

		<!-- Calendar Reports Dates control -->
		<object
			classid="CLSID:41021C13-8D42-4364-8388-9506F0755AE3"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,2"
			id="ctlDates"
			name="ctlDates"
			style="WIDTH: 100%"
			width="100%"
			viewastext>
		</object>


		<!-- Record Edit control -->
		<object
			classid="CLSID:2D0A5ED7-6669-481F-9A5D-19BA14E92364"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,162"
			id="ctlRecordEdit"
			viewastext>
			<param name="_ExtentX" value="16007">
			<param name="_ExtentY" value="6403">
			<param name="TabCount" value="0">
			<param name="TabCaptions" value="">
			<param name="BorderStyle" value="0">
		</object>


		<!-- Self-service Record Edit control -->
		<object
			id="ctlSSIRecordEdit"
			style="LEFT: 0px; TOP: 0px"
			classid="CLSID:F1D5B565-EBE9-4295-97AF-60F6F8A126E1"
			codebase="cabs/COAInt_Client.cab#version=1,0,0,162"
			viewastext>
			<param name="_ExtentX" value="160070">
			<param name="_ExtentY" value="7594">
			<param name="TabCount" value="0">
			<param name="TabCaptions" value="">
			<param name="BorderStyle" value="0">
		</object>


		<script for="window" event="onload" language="JavaScript">
<!--

	var temp;
	var sMessage = "";

	/* Test the CodeJock controls */
	try 
	{
		temp = ctlCodeJock_PushButton.Caption;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'CodeJock controls'.";
		setStatus(2,sMessage);
		return;
	}
	
	/* Test the Client-side Intranet general functions DLL */
	try 
	{
		temp = ASRIntranetFunctions.LocaleDateFormat;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Client-side General Functions DLL'.";
		setStatus(2,sMessage);
		return;
	}
	
	/* Test the Client-side Intranet printer functions DLL */
	try 
	{
		temp = ASRIntranetPrintFunctions.ClipboardGetText();
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Client-side Printer Functions DLL'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Client-side Intranet output functions DLL */
	try 
	{
		//temp = ASRIntranetOutput.UserName();
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Client-side Output Functions DLL'.";
		setStatus(2,sMessage);
		return;
	}
		
		
	/* Test the Grid control */
	try 
	{
		temp = ssOleDBGrid.Cols;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Grid control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Image control */
	try 
	{
		temp = ASRUserImage1.Picture;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Image control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Intranet Lookup control */
	try 
	{
		temp = ctlIntLookup1.Text;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Intranet Lookup control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Option Group control */
	try 
	{
		temp = ctlNewOptionGroup3.Caption;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Option Group control'.";
		setStatus(2,sMessage);
		return;
	}
	
	//NPG20111014 Fault HRPRO-1799, don't reload the menu control if already loaded.
	if(window.location.search.substring(1) != "fromMenu=true") {
		/* Test the Menu control */
		try 
		{
			temp = abMainMenu.Bands;
			if (typeof(temp) == "undefined")
			{
				throw e;
			}
		}
		catch(e) 
		{
			sMessage = "Error downloading 'Menu control'.";
			setStatus(2,sMessage);
			return;
		}
	}

	/* Test the Navigation control */
	try 
	{
		temp = ctlNavigation.NavigateTo;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Navigation control'.";
		setStatus(2,sMessage);
		return;
	}


	/* Test the Mask control */
	try	
	{
		temp = TDBMask1.Value;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Mask control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Number control */
	try 
	{
		temp = TDBNumber1.Value;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Number control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the ASR Spinner control */
	try 
	{
		temp = ASRSpinner1.Value;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Spinner control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Date control */
	try 
	{
		temp = TDBDate1.Text;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Date control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Record Edit control */
	try 
	{
		temp = ctlRecordEdit.changed;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Record Edit control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the SSI Record Edit control */
	try 
	{
		temp = ctlSSIRecordEdit.Changed;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Self-service Record Edit control'.";
		setStatus(2,sMessage);
		return;
	}
		
	/* Test the Dialogue control */
	try 
	{
		temp = dialog.FileName;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Dialogue control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Tree control */
	try 
	{
		temp = SSTree.Nodes;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Tree control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Calendar Report Record control */
	try 
	{
		temp = ctlCalRec.ClientDateFormat;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Calendar Report Record control'.";
		setStatus(2,sMessage);
		return;
	}
	
	/* Test the Calendar Report Key control */
	try 
	{
		temp = ctlKey.CaptionsVisible;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Calendar Report Key control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Calendar Report Dates control */
	try 
	{
		temp = ctlDates.ClientDateFormat;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Calendar Report Dates control'.";
		setStatus(2,sMessage);
		return;
	}

	/* Test the Embedded OLE control */
	try 
	{
		temp = ASRIntOLE1.FileName;
		if (typeof(temp) == "undefined")
		{
			throw e;
		}
	}
	catch(e) 
	{
		sMessage = "Error downloading 'Embedded OLE control'.";
		setStatus(2,sMessage);
		return;
	}

		
	/* Controls all down okay. */
	sMessage = "All controls downloaded successfully.";
	setStatus(1,sMessage);
	return;

	-->
		</script>

		<script language="JavaScript">
<!--

	function setStatus(piStatus, psMessage)
	{

		if (piStatus == 0)
		{
			sButtonText = 'Cancel';
		}
		else if (piStatus == 1)
		{
			sButtonText = 'OK';
		}
		else if (piStatus == 2)
		{ 
			sButtonText = 'Retry';
		}
	
		window.parent.frames("downloadControlsStatusFrame").document.getElementById('txtDownloadStatus').value = piStatus;
		window.parent.frames("downloadControlsStatusFrame").document.getElementById('txtMessage').innerText = psMessage;
		var objButton = window.parent.frames("downloadControlsStatusFrame").document.getElementById('tdButton');

		objButton.value = sButtonText;
	
		return;
	}

	-->
		</script>

	</head>

	<body topmargin="10" bottommargin="10" leftmargin="10" rightmargin="10">
		<form action="downloadControls.asp" method="post" id="frmRefresh" name="frmRefresh">
		</form>
	</body>
	</html>

	<!-- Embeds createActiveX.js script reference -->
	<!--#include file="include\ctl_CreateControl.txt"-->


</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="FixedLinksContent" runat="server">
</asp:Content>
