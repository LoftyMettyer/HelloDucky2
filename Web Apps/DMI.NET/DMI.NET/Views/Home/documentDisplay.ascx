<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="ADODB" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>

<script type="text/javascript">
	function formatDocumentDisplay(fShowDocDisplay) {
		
		if (fShowDocDisplay == 'False') {
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
				objNavigation.Connection = CType(Session("databaseConnection"), Connection)
												
				Dim objDocumentInfo = objNavigation.GetDocuments(NavigationLinkType.DocumentDisplay)

				For iCount = 1 To objDocumentInfo.Count
		
			%>
			<tr>
				<td style="vertical-align: top">
					<iframe id="ifDocumentDisplay<%=iCount %>" style="border: 0; margin: 0;"
						src="<%=objDocumentInfo(iCount).DocumentFilePath%>" style="visibility: visible; z-index: 1; display: block; padding: 0; margin: 0; position: relative; width: 100%; height: 100%; border-style: none; border-width: 0;"></iframe>
				</td>
			</tr>
			<%	 If objDocumentInfo(iCount).DisplayDocumentHyperlink = True Then%>
			<tr>
				<td>
					<label id="divDocumentHyperlink<%=iCount %>"
						onclick="openDocument('<%=objDocumentInfo(iCount).DocumentFilePath %>');"
						style="cursor: pointer; position: static;">
						<%=objDocumentInfo(iCount).Text %>
					</label>
				</td>
			</tr>
			<%    
			End If
		Next

			%>
		</table>
	</div>
</div>

<%Dim fShowDocDisplay As Boolean = (objDocumentInfo.Count > 0)%>
<script type="text/javascript"> formatDocumentDisplay('<%=fShowDocDisplay%>');</script>
