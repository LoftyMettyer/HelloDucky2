<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%--licence manager reference for activeX--%>
<object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
	id="Microsoft_Licensed_Class_Manager_1_0"
	viewastext>
	<param name="LPKPath" value="<%: Url.Content("~/lpks/ssmain.lpk")%>">
</object>

<OBJECT 
	ID="ASRIntranetPrintFunctions"
	CLASSID="CLSID:6B26B8A7-7A9B-42D2-9528-18BC6037DF49"
	CODEBASE="cabs/COAInt_Client.CAB#version=1,0,0,147" 
	VIEWASTEXT>
</OBJECT>
