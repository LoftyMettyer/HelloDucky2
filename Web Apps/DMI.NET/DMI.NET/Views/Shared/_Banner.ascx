<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%If Session("Config-banner-justification") = "left" Then%>
<div style="float: left;">
	<img src="<%=session("TopBarFile")%>" width="<%=session("Config-banner-graphic-left-width")%>" height="44px" alt=""></div>
<div style="float: left;">
	<img src="<%=Session("LogoFile")%>" width="<%=Session("Config-banner-graphic-right-width")%>" height="44px" alt=""></div>
<%ElseIf Session("Config-banner-justification") = "right" Then%>
<div style="float: right;">
	<img src="<%=session("TopBarFile")%>" width="<%=session("Config-banner-graphic-left-width")%>" height="44px" alt=""></div>
<div style="float: right;">
	<img src="<%=Session("LogoFile")%>" width="<%=Session("Config-banner-graphic-right-width")%>" height="44px" alt=""></div>
<%ElseIf Session("Config-banner-justification") = "justify" Then%>
<div style="float: left;">
	<img src="<%=session("TopBarFile")%>" width="<%=session("Config-banner-graphic-left-width")%>" height="44px" alt=""></div>
<div style="float: right;">
	<img src="<%=Session("LogoFile")%>" width="<%=Session("Config-banner-graphic-right-width")%>" height="44px" alt=""></div>
<%Else
		Dim styleWidth = CInt(Session("Config-banner-graphic-left-width")) + CInt(Session("Config-banner-graphic-right-width")) & "px"%>
<div style="width: <%=styleWidth%>; margin: 0 auto;">
	<div style="float: left;">
		<img src="<%=session("TopBarFile")%>" width="<%=session("Config-banner-graphic-left-width")%>" height="44px" alt=""></div>
	<div style="float: left;">
		<img src="<%=Session("LogoFile")%>" width="<%=Session("Config-banner-graphic-right-width")%>" height="44px" alt=""></div>
</div>
<%End If%>
