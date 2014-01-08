<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="ADODB" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script type="text/javascript">

	function formatDocumentDisplay(fShowDocDisplay) {

		if (fShowDocDisplay == 'false') {
			//Hide parent div
			$('#documentDisplay').hide();
		}
		else {			
			//available height
			var availHeight = document.getElementById('divResize').clientHeight;

			//number of frames		
			var colIFrames = $('iframe[id^="ifDocumentDisplay"]');
			var frameCount = colIFrames.length;

			var newFrameHeight = availHeight / frameCount;

			for (var i = 0; i < colIFrames.length; i++) {

				var objIframe = colIFrames[i];
				objIframe.style.width = '100%';

				var docLinkLabelID = new String('divDocumentHyperlink' + (i + 1));
				var docLinkLabel = document.getElementById(docLinkLabelID);
				if (docLinkLabel) {
					objIframe.style.height = (newFrameHeight - docLinkLabel.offsetHeight - 10) + 'px';
				} else {
					objIframe.style.height = (newFrameHeight) + 'px';
				}

			}

			return;
		}
	}

	function openDocument(docUrl) {
		window.open(docUrl);
		return;
	}
</script>

<div>

	<div>
		<table style="width: 100%; outline-style: none; border-style: none; border-width: 0; position: absolute; padding: 0; margin: 0 0 0 6px;">
			<% 
				' Get the Documents collection
				Dim objNavigation = New HR.Intranet.Server.clsNavigationLinks
				objNavigation.SessionInfo = CType(Session("SessionContext"), SessionInfo)

				Dim iCount As Integer = 0
												
				For Each objLink In objNavigation.GetLinks(LinkType.DocumentDisplay)
		
			%>
			<tr>
				<td style="vertical-align: top">
					<iframe id="ifDocumentDisplay<%=iCount %>" style="border: 0; margin: 0;"
						src="<%=objLink.DocumentFilePath%>" style="visibility: visible; z-index: 1; display: block; padding: 0; margin: 0; position: relative; width: 100%; height: 100%; border-style: none; border-width: 0;"></iframe>
				</td>
			</tr>
			<%
				If objLink.DisplayDocumentHyperlink Then
			%>
			<tr>
				<td>
					<label id="divDocumentHyperlink<%=iCount %>"
						onclick="openDocument('<%=objLink.DocumentFilePath%>');"
						style="cursor: pointer; position: static;">
						<%=objLink.Text%>
					</label>
				</td>
			</tr>
			<%    
			End If
			
			iCount += 1
		Next

			%>
		</table>
	</div>
</div>

<script type="text/javascript">
	<%
	If (iCount > 0) Then
		Response.Write("formatDocumentDisplay('true');")
	Else
		Response.Write("formatDocumentDisplay('false');")
	End If
%>
</script>
