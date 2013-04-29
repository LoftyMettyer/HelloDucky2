<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
downloadControls
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="OpenHR.css">
<TITLE>OpenHR Intranet</TITLE>
<meta http-equiv="X-UA-Compatible" content="IE=5">
<!-- LPK file -->
<OBJECT 
	classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
	id="Microsoft_Licensed_Class_Manager_1_0" 
	VIEWASTEXT>
		<PARAM NAME="LPKPath" VALUE="lpks/main.lpk">
</OBJECT>

<!-- Client-side Intranet general functions DLL -->
<!--#include file="include\ctl_ASRIntranetFunctions.txt"-->

<!-- Client-side Intranet print functions DLL -->
<!--#include file="include\ctl_ASRIntranetPrintFunctions.txt"-->

<!-- Client-side Intranet output functions DLL -->
<!--#include file="include\ctl_ASRIntranetOutput.txt"-->

<!-- Codejock controls -->
<OBJECT 
	id=ctlCodeJock_PushButton
	CLASSID="CLSID:3E8187B5-3C15-4233-B811-9CB3F929A28B"
	codebase="cabs/COAInt_Client.cab#version=13,1,0,0"
	style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden" 
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="3307">
		<PARAM NAME="_ExtentY" VALUE="1323">
</OBJECT>

<!-- Grid control -->
<OBJECT 
	classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
	codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
	id=ssOleDBGrid
	name=ssOleDBGrid
	style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%" 
	VIEWASTEXT>
		<PARAM NAME="ScrollBars" VALUE="4">
		<PARAM NAME="_Version" VALUE="196617">
		<PARAM NAME="DataMode" VALUE="2">
		<PARAM NAME="Cols" VALUE="0">
		<PARAM NAME="Rows" VALUE="0">
		<PARAM NAME="BorderStyle" VALUE="1">
		<PARAM NAME="RecordSelectors" VALUE="0">
		<PARAM NAME="GroupHeaders" VALUE="0">
		<PARAM NAME="ColumnHeaders" VALUE="1">
		<PARAM NAME="GroupHeadLines" VALUE="1">
		<PARAM NAME="HeadLines" VALUE="1">
		<PARAM NAME="FieldDelimiter" VALUE="(None)">
		<PARAM NAME="FieldSeparator" VALUE="(Tab)">
		<PARAM NAME="Row.Count" VALUE="0">
		<PARAM NAME="Col.Count" VALUE="1">
		<PARAM NAME="stylesets.count" VALUE="0">
		<PARAM NAME="TagVariant" VALUE="EMPTY">
		<PARAM NAME="UseGroups" VALUE="0">
		<PARAM NAME="HeadFont3D" VALUE="0">
		<PARAM NAME="Font3D" VALUE="0">
		<PARAM NAME="DividerType" VALUE="3">
		<PARAM NAME="DividerStyle" VALUE="1">
		<PARAM NAME="DefColWidth" VALUE="0">
		<PARAM NAME="BeveColorScheme" VALUE="2">
		<PARAM NAME="BevelColorFrame" VALUE="-2147483642">
		<PARAM NAME="BevelColorHighlight" VALUE="-2147483628">
		<PARAM NAME="BevelColorShadow" VALUE="-2147483632">
		<PARAM NAME="BevelColorFace" VALUE="-2147483633">
		<PARAM NAME="CheckBox3D" VALUE="-1">
		<PARAM NAME="AllowAddNew" VALUE="0">
		<PARAM NAME="AllowDelete" VALUE="0">
		<PARAM NAME="AllowUpdate" VALUE="0">
		<PARAM NAME="MultiLine" VALUE="0">
		<PARAM NAME="ActiveCellStyleSet" VALUE="">
		<PARAM NAME="RowSelectionStyle" VALUE="0">
		<PARAM NAME="AllowRowSizing" VALUE="0">
		<PARAM NAME="AllowGroupSizing" VALUE="0">
		<PARAM NAME="AllowColumnSizing" VALUE="-1">
		<PARAM NAME="AllowGroupMoving" VALUE="0">
		<PARAM NAME="AllowColumnMoving" VALUE="0">
		<PARAM NAME="AllowGroupSwapping" VALUE="0">
		<PARAM NAME="AllowColumnSwapping" VALUE="0">
		<PARAM NAME="AllowGroupShrinking" VALUE="0">
		<PARAM NAME="AllowColumnShrinking" VALUE="0">
		<PARAM NAME="AllowDragDrop" VALUE="0">
		<PARAM NAME="UseExactRowCount" VALUE="-1">
		<PARAM NAME="SelectTypeCol" VALUE="0">
		<PARAM NAME="SelectTypeRow" VALUE="1">
		<PARAM NAME="SelectByCell" VALUE="-1">
		<PARAM NAME="BalloonHelp" VALUE="0">
		<PARAM NAME="RowNavigation" VALUE="1">
		<PARAM NAME="CellNavigation" VALUE="0">
		<PARAM NAME="MaxSelectedRows" VALUE="1">
		<PARAM NAME="HeadStyleSet" VALUE="">
		<PARAM NAME="StyleSet" VALUE="">
		<PARAM NAME="ForeColorEven" VALUE="0">
		<PARAM NAME="ForeColorOdd" VALUE="0">
		<PARAM NAME="BackColorEven" VALUE="16777215">
		<PARAM NAME="BackColorOdd" VALUE="16777215">
		<PARAM NAME="Levels" VALUE="1">
		<PARAM NAME="RowHeight" VALUE="503">
		<PARAM NAME="ExtraHeight" VALUE="0">
		<PARAM NAME="ActiveRowStyleSet" VALUE="">
		<PARAM NAME="CaptionAlignment" VALUE="2">
		<PARAM NAME="SplitterPos" VALUE="0">
		<PARAM NAME="SplitterVisible" VALUE="0">
		<PARAM NAME="Columns.Count" VALUE="1">			
		<PARAM NAME="Columns(0).Width" VALUE="6500">
		<PARAM NAME="Columns(0).Visible" VALUE="-1">
		<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
		<PARAM NAME="Columns(0).Caption" VALUE="">
		<PARAM NAME="Columns(0).Name" VALUE="">
		<PARAM NAME="Columns(0).Alignment" VALUE="0">
		<PARAM NAME="Columns(0).CaptionAlignment" VALUE="3">
		<PARAM NAME="Columns(0).Bound" VALUE="0">
		<PARAM NAME="Columns(0).AllowSizing" VALUE="1">
		<PARAM NAME="Columns(0).DataField" VALUE="Column 0">
		<PARAM NAME="Columns(0).DataType" VALUE="8">
		<PARAM NAME="Columns(0).Level" VALUE="0">
		<PARAM NAME="Columns(0).NumberFormat" VALUE="">
		<PARAM NAME="Columns(0).Case" VALUE="0">
		<PARAM NAME="Columns(0).FieldLen" VALUE="256">
		<PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
		<PARAM NAME="Columns(0).Locked" VALUE="0">
		<PARAM NAME="Columns(0).Style" VALUE="0">
		<PARAM NAME="Columns(0).ButtonsAlways" VALUE="0">
		<PARAM NAME="Columns(0).RowCount" VALUE="0">
		<PARAM NAME="Columns(0).ColCount" VALUE="1">
		<PARAM NAME="Columns(0).HasHeadForeColor" VALUE="0">
		<PARAM NAME="Columns(0).HasHeadBackColor" VALUE="0">
		<PARAM NAME="Columns(0).HasForeColor" VALUE="0">
		<PARAM NAME="Columns(0).HasBackColor" VALUE="0">
		<PARAM NAME="Columns(0).HeadForeColor" VALUE="0">
		<PARAM NAME="Columns(0).HeadBackColor" VALUE="0">
		<PARAM NAME="Columns(0).ForeColor" VALUE="0">
		<PARAM NAME="Columns(0).BackColor" VALUE="0">
		<PARAM NAME="Columns(0).HeadStyleSet" VALUE="">
		<PARAM NAME="Columns(0).StyleSet" VALUE="">
		<PARAM NAME="Columns(0).Nullable" VALUE="1">
		<PARAM NAME="Columns(0).Mask" VALUE="">
		<PARAM NAME="Columns(0).PromptInclude" VALUE="0">
		<PARAM NAME="Columns(0).ClipMode" VALUE="0">
		<PARAM NAME="Columns(0).PromptChar" VALUE="95">
		<PARAM NAME="UseDefaults" VALUE="-1">
		<PARAM NAME="TabNavigation" VALUE="1">
		<PARAM NAME="BatchUpdate" VALUE="0">
		<PARAM NAME="_ExtentX" VALUE="16087">
		<PARAM NAME="_ExtentY" VALUE="4630">
		<PARAM NAME="_StockProps" VALUE="79">
		<PARAM NAME="Caption" VALUE="">
		<PARAM NAME="ForeColor" VALUE="0">
		<PARAM NAME="BackColor" VALUE="16777215">
		<PARAM NAME="Enabled" VALUE="-1">
		<PARAM NAME="DataMember" VALUE="">
</OBJECT>


<!--Intranet Lookup Control -->
<OBJECT 
	id=ctlIntLookup1 
	CLASSID="CLSID:F513059D-1E77-4F0A-8681-21B2BA7E5986"
	 codebase="cabs/COAInt_Client.cab#version=1,0,0,3" 
	style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden" 
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="4233">
		<PARAM NAME="_ExtentY" VALUE="953">
		<PARAM NAME="BackColor" VALUE="16777215">
		<PARAM NAME="ForeColor" VALUE="0">
		<PARAM NAME="Enabled" VALUE="-1">
		<PARAM NAME="Text" VALUE="">
</OBJECT>

<!-- New Option Group 3 -->
<OBJECT 
	id=ctlNewOptionGroup3
	CLASSID="CLSID:0B21DBA2-E1A7-47C7-B62F-BC522820590A"
	codebase="cabs/COAInt_Client.cab#version=1,0,0,2"
	style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden" 
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="3307">
		<PARAM NAME="_ExtentY" VALUE="1323">
</OBJECT>

<!-- Image control -->
<OBJECT 
	id=ASRUserImage1 
	CLASSID="CLSID:8FF15C8D-49D5-4B79-8419-C36C26654283"
	CODEBASE="cabs/COAInt_Client.cab#version=1,0,0,7"
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="2619">
		<PARAM NAME="_ExtentY" VALUE="2619">
		<PARAM NAME="ForeColor" VALUE="0">
		<PARAM NAME="Enabled" VALUE="-1">
		<PARAM NAME="BorderStyle" VALUE="0">
		<PARAM NAME="ASRDataField" VALUE="0">
</OBJECT>

<!-- Menu control (base) -->
<OBJECT 
	classid="clsid:E4F874A0-56ED-11D0-9C43-00A0C90F29FC" 
	codebase="cabs/COAInt_Client.cab#Version=1,0,6,5" 
	height=32 
	id=abMainMenuBase 
	name=abMainMenuBase 
	style="LEFT: 0px; TOP: 0px" 
	width=32 
	VIEWASTEXT>
</OBJECT>

<%if Request.QueryString("fromMenu") = "false" then  %>
<!--NPG20111014 Fault HRPRO-1799, don't reload the menu control if already loaded.-->
<!-- Menu control -->
<OBJECT 
	classid="clsid:6976CB54-C39B-4181-B1DC-1A829068E2E7"
	codebase="cabs/COAInt_Client.cab#Version=1,0,0,5" 
	id=abMainMenu 
	name=abMainMenu
	style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden" 
	VIEWASTEXT>
</OBJECT>
<%end if %>

<!-- Mask control -->
<OBJECT 
	classid=clsid:66A90C04-346D-11D2-9BC0-00A024695830 
	codebase="cabs/COAInt_Client.cab#version=6,0,1,1" 
	id=TDBMask1 
	VIEWASTEXT>
</OBJECT>

<!-- Number control -->
<OBJECT 
	classid=clsid:49CBFCC2-1337-11D2-9BBF-00A024695830 
	codebase="cabs/COAInt_Client.cab#version=6,0,1,1" 
	id=TDBNumber1 
	VIEWASTEXT>
</OBJECT>

<!-- ASR Spinner control -->
<OBJECT 
	CLASSID="CLSID:C25C3704-2AA7-44E5-943A-B40B14E2348F"
	CODEBASE="cabs/COAInt_Client.cab#version=1,0,0,3"
	id=ASRSpinner1 
	style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden" 
	VIEWASTEXT>
</OBJECT>

<!-- Date control -->
<OBJECT 
	classid="clsid:A49CE0E4-C0F9-11D2-B0EA-00A024695830" 
	codebase="cabs/COAInt_Client.cab#version=6,0,1,1" 
	id=TDBDate1 
	style="LEFT: 0px; TOP: 0px; VISIBILITY: hidden" 
	VIEWASTEXT>
</OBJECT>

<!-- Dialogue control -->
<OBJECT 
	classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB" 
  codebase="cabs/COAInt_Client.cab#Version=1,0,0,0"
	id=dialog 
	style="LEFT: 0px; TOP: 0px" 
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="847">
		<PARAM NAME="_ExtentY" VALUE="847">
		<PARAM NAME="_Version" VALUE="393216">
		<PARAM NAME="CancelError" VALUE="0">
		<PARAM NAME="Color" VALUE="0">
		<PARAM NAME="Copies" VALUE="1">
		<PARAM NAME="DefaultExt" VALUE="">
		<PARAM NAME="DialogTitle" VALUE="">
		<PARAM NAME="FileName" VALUE="">
		<PARAM NAME="Filter" VALUE="">
		<PARAM NAME="FilterIndex" VALUE="0">
		<PARAM NAME="Flags" VALUE="0">
		<PARAM NAME="FontBold" VALUE="0">
		<PARAM NAME="FontItalic" VALUE="0">
		<PARAM NAME="FontName" VALUE="">
		<PARAM NAME="FontSize" VALUE="8">
		<PARAM NAME="FontStrikeThru" VALUE="0">
		<PARAM NAME="FontUnderLine" VALUE="0">
		<PARAM NAME="FromPage" VALUE="0">
		<PARAM NAME="HelpCommand" VALUE="0">
		<PARAM NAME="HelpContext" VALUE="0">
		<PARAM NAME="HelpFile" VALUE="">
		<PARAM NAME="HelpKey" VALUE="">
		<PARAM NAME="InitDir" VALUE="">
		<PARAM NAME="Max" VALUE="0">
		<PARAM NAME="Min" VALUE="0">
		<PARAM NAME="MaxFileSize" VALUE="260">
		<PARAM NAME="PrinterDefault" VALUE="1">
		<PARAM NAME="ToPage" VALUE="0">
		<PARAM NAME="Orientation" VALUE="1">
</OBJECT>

<!-- Tree control -->
<OBJECT 
	id=SSTree    
	classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" 
	codebase="cabs/COAInt_Client.cab#version=1,0,2,24" 
	style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" 
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="2646">
		<PARAM NAME="_ExtentY" VALUE="1323">
		<PARAM NAME="_Version" VALUE="65538">
		<PARAM NAME="BackColor" VALUE="-2147483643">
		<PARAM NAME="ForeColor" VALUE="-2147483640">
		<PARAM NAME="ImagesMaskColor" VALUE="12632256">
		<PARAM NAME="PictureBackgroundMaskColor" VALUE="12632256">
		<PARAM NAME="Appearance" VALUE="1">
		<PARAM NAME="BorderStyle" VALUE="0">
		<PARAM NAME="LabelEdit" VALUE="1">
		<PARAM NAME="LineStyle" VALUE="0">
		<PARAM NAME="LineType" VALUE="1">
		<PARAM NAME="MousePointer" VALUE="0">
		<PARAM NAME="NodeSelectionStyle" VALUE="2">
		<PARAM NAME="PictureAlignment" VALUE="0">
		<PARAM NAME="ScrollStyle" VALUE="0">
		<PARAM NAME="Style" VALUE="6">
		<PARAM NAME="IndentationStyle" VALUE="0">
		<PARAM NAME="TreeTips" VALUE="3">
		<PARAM NAME="PictureBackgroundStyle" VALUE="0">
		<PARAM NAME="Indentation" VALUE="38">
		<PARAM NAME="MaxLines" VALUE="1">
		<PARAM NAME="TreeTipDelay" VALUE="500">
		<PARAM NAME="ImageCount" VALUE="0">
		<PARAM NAME="ImageListIndex" VALUE="-1">
		<PARAM NAME="OLEDragMode" VALUE="0">
		<PARAM NAME="OLEDropMode" VALUE="0">
		<PARAM NAME="AllowDelete" VALUE="0">
		<PARAM NAME="AutoSearch" VALUE="0">
		<PARAM NAME="Enabled" VALUE="-1">
		<PARAM NAME="HideSelection" VALUE="0">
		<PARAM NAME="ImagesUseMask" VALUE="0">
		<PARAM NAME="Redraw" VALUE="-1">
		<PARAM NAME="UseImageList" VALUE="-1">
		<PARAM NAME="PictureBackgroundUseMask" VALUE="0">
		<PARAM NAME="HasFont" VALUE="0">
		<PARAM NAME="HasMouseIcon" VALUE="0">
		<PARAM NAME="HasPictureBackground" VALUE="0">
		<PARAM NAME="PathSeparator" VALUE="\">
		<PARAM NAME="TabStops" VALUE="32">
		<PARAM NAME="ImageList" VALUE="<None>">
		<PARAM NAME="LoadStyleRoot" VALUE="1">
		<PARAM NAME="Sorted" VALUE="0">
		<PARAM NAME="OnDemandDiscardBuffer" VALUE="10">
</OBJECT>

<OBJECT 
	id=SSTree2
	classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" 
	codebase="cabs/COAInt_Client.cab#version=1,0,2,24" 
	style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" 
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="2646">
		<PARAM NAME="_ExtentY" VALUE="1323">
		<PARAM NAME="_Version" VALUE="65538">
		<PARAM NAME="BackColor" VALUE="-2147483643">
		<PARAM NAME="ForeColor" VALUE="-2147483640">
		<PARAM NAME="ImagesMaskColor" VALUE="12632256">
		<PARAM NAME="PictureBackgroundMaskColor" VALUE="12632256">
		<PARAM NAME="Appearance" VALUE="1">
		<PARAM NAME="BorderStyle" VALUE="0">
		<PARAM NAME="LabelEdit" VALUE="1">
		<PARAM NAME="LineStyle" VALUE="0">
		<PARAM NAME="LineType" VALUE="1">
		<PARAM NAME="MousePointer" VALUE="0">
		<PARAM NAME="NodeSelectionStyle" VALUE="2">
		<PARAM NAME="PictureAlignment" VALUE="0">
		<PARAM NAME="ScrollStyle" VALUE="0">
		<PARAM NAME="Style" VALUE="6">
		<PARAM NAME="IndentationStyle" VALUE="0">
		<PARAM NAME="TreeTips" VALUE="3">
		<PARAM NAME="PictureBackgroundStyle" VALUE="0">
		<PARAM NAME="Indentation" VALUE="38">
		<PARAM NAME="MaxLines" VALUE="1">
		<PARAM NAME="TreeTipDelay" VALUE="500">
		<PARAM NAME="ImageCount" VALUE="0">
		<PARAM NAME="ImageListIndex" VALUE="-1">
		<PARAM NAME="OLEDragMode" VALUE="0">
		<PARAM NAME="OLEDropMode" VALUE="0">
		<PARAM NAME="AllowDelete" VALUE="0">
		<PARAM NAME="AutoSearch" VALUE="0">
		<PARAM NAME="Enabled" VALUE="-1">
		<PARAM NAME="HideSelection" VALUE="0">
		<PARAM NAME="ImagesUseMask" VALUE="0">
		<PARAM NAME="Redraw" VALUE="-1">
		<PARAM NAME="UseImageList" VALUE="-1">
		<PARAM NAME="PictureBackgroundUseMask" VALUE="0">
		<PARAM NAME="HasFont" VALUE="0">
		<PARAM NAME="HasMouseIcon" VALUE="0">
		<PARAM NAME="HasPictureBackground" VALUE="0">
		<PARAM NAME="PathSeparator" VALUE="\">
		<PARAM NAME="TabStops" VALUE="32">
		<PARAM NAME="ImageList" VALUE="<None>">
		<PARAM NAME="LoadStyleRoot" VALUE="1">
		<PARAM NAME="Sorted" VALUE="0">
		<PARAM NAME="OnDemandDiscardBuffer" VALUE="10">
</OBJECT>


<!-- Navigation control -->
<OBJECT 
	CLASSID="CLSID:0F3914E5-4334-4E10-8DE8-538F9B828CD9"
	CODEBASE="cabs/COAInt_Client.cab#version=1,0,0,10"
	ID=ctlNavigation
	name=ctlNavigation
	style="WIDTH: 100%;visibility:hidden;display:none"
	VIEWASTEXT>
</OBJECT>


<!-- Calendar Reports Record control -->
<OBJECT 
	CLASSID="CLSID:252D73AF-D7C6-4833-8539-A2C0293950B1"
	CODEBASE="cabs/COAInt_Client.cab#version=1,0,0,2"
	id=ctlCalRec
	name=ctlCalRec
	style="WIDTH: 100%"
	width="100%">
</OBJECT>
 
 
<!-- Calendar Reports Key control -->
<OBJECT 
	CLASSID="CLSID:8E2F1EF1-3812-4678-A084-16384DE3EA6D"
	CODEBASE="cabs/COAInt_Client.cab#version=1,0,0,2"
	ID=ctlKey 
	Name=ctlKey
	style="WIDTH: 100%;visibility:hidden;display:none"
	width=100% 
	height=85%>
</OBJECT>

<!-- Intranet OLE Embedded control -->
<OBJECT 
	ID=ASRIntOLE1
	name=ASRIntOLE1
	CLASSID="CLSID:5A5DF13B-2C9C-49E8-9C51-809E909F1123"
	CODEBASE="cabs/COAInt_Client.cab#version=1,0,0,15"
	style="WIDTH: 100%;visibility:hidden;display:none"
	VIEWASTEXT>
</OBJECT>

<!-- Calendar Reports Dates control -->
<OBJECT 
	CLASSID="CLSID:41021C13-8D42-4364-8388-9506F0755AE3"
	CODEBASE="cabs/COAInt_Client.cab#version=1,0,0,2" 
	id=ctlDates 
	name=ctlDates 
	style="WIDTH: 100%" 
	width="100%"
	VIEWASTEXT>
</OBJECT>


<!-- Record Edit control -->
<OBJECT 
	CLASSID="CLSID:2D0A5ED7-6669-481F-9A5D-19BA14E92364"
	CODEBASE="cabs/COAInt_Client.cab#version=1,0,0,162"
	id=ctlRecordEdit 
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="16007">
		<PARAM NAME="_ExtentY" VALUE="6403">
		<PARAM NAME="TabCount" VALUE="0">
		<PARAM NAME="TabCaptions" VALUE="">
		<PARAM NAME="BorderStyle" VALUE="0">
</OBJECT>


<!-- Self-service Record Edit control -->
<OBJECT 
	id=ctlSSIRecordEdit 
	style="LEFT: 0px; TOP: 0px"     
	CLASSID="CLSID:F1D5B565-EBE9-4295-97AF-60F6F8A126E1"
	 CODEBASE="cabs/COAInt_Client.cab#version=1,0,0,162"
	VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="160070">
		<PARAM NAME="_ExtentY" VALUE="7594">
		<PARAM NAME="TabCount" VALUE="0">
		<PARAM NAME="TabCaptions" VALUE="">
		<PARAM NAME="BorderStyle" VALUE="0">
</OBJECT>


<SCRIPT FOR=window EVENT=onload LANGUAGE=JavaScript>
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
</SCRIPT>

<SCRIPT LANGUAGE=JavaScript>
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
</SCRIPT>

</HEAD>

<BODY TOPMARGIN=10 BOTTOMMARGIN=10 LEFTMARGIN=10 RIGHTMARGIN=10>
<FORM action="downloadControls.asp" method=post id=frmRefresh name=frmRefresh>

</FORM>
</BODY>
</HTML>

<!-- Embeds createActiveX.js script reference -->
<!--#include file="include\ctl_CreateControl.txt"-->


</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="FixedLinksContent" runat="server">
</asp:Content>
