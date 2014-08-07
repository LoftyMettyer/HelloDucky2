<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<script runat="server">
	Private Function GetEmailSelection() As String
		Dim emailSelectionHtmlTable As New StringBuilder 'Used to construct the (temporary) HTML table that will be transformed into a jQuery grid table
		Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)		'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
		
		'Get the records.
		Dim rstDefSelRecords As DataTable = objDataAccess.GetDataTable("spASRIntGetRecordSelection", CommandType.StoredProcedure _
			, New SqlParameter("@psType", SqlDbType.VarChar, 255) With {.Value = "EMAIL"} _
			, New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = 0})

		'Create an HTML table
		With emailSelectionHtmlTable
			.Append("<table id=""EmailSelectionTable"">")
			.Append("<tr>")
			.Append("<th id=""EmailGroupIDHeader"">EmailGroupID</th>")
			.Append("<th id=""NameHeader"">Name</th>")
			.Append("<th id=""FullNameHeader"">FullName</th>")
			.Append("</tr>")
		End With

        'Populate the table
		Dim i As Integer = 1
		For Each r As DataRow In rstDefSelRecords.Rows
			With emailSelectionHtmlTable
				.Append("<tr>")
				.Append("<td id='Row" & i & "'>" & r("ID").ToString & "</td>")
				.Append("<td>" & r("name").ToString.Replace("_", " ").Replace("""", "&quot;") & "</td>")
				'Loop around and add on fullemailaddresses to the grid
				If r("ID") < 1 Then
				Else
					Try
						Dim rstEmailAddr = objDataAccess.GetDataTable("spASRIntGetEmailGroupAddresses", CommandType.StoredProcedure _
						, New SqlParameter("EmailGroupID", SqlDbType.Int) With {.Value = r("ID")})
						Dim x As Integer = 1
						If Not rstEmailAddr Is Nothing Then
							If rstEmailAddr.Rows.Count < 2 Then
								For Each objRow In rstEmailAddr.Rows
									If x > 0 Then
										.Append("<td>" & objRow(0).ToString.Replace("_", " ").Replace("""", "&quot;") & "</td>")
									End If
									x += 1
								Next
							Else
								'Append two email addresses
								Dim buildMultipleEmailString As String = ""
								For Each objRow In rstEmailAddr.Rows
									If buildMultipleEmailString.Length > 0 Then buildMultipleEmailString = buildMultipleEmailString + ";"
									buildMultipleEmailString = buildMultipleEmailString + objRow(0).ToString.Replace("_", " ").Replace("""", "&quot;")
									x += 1
								Next
								If x > 0 Then
									.Append("<td>" & buildMultipleEmailString.ToString.Replace("_", " ").Replace("""", "&quot;") & "</td>")
								End If
							End If
						End If
                            
					Catch ex As Exception
						Dim sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(ex.Message)
					End Try
				End If
				.Append("</tr>")
				i += 1
			End With
		Next

		emailSelectionHtmlTable.Append("</table>")

		Return emailSelectionHtmlTable.ToString
	End Function
</script>

<!DOCTYPE html>
<html>
<head>
    <title>OpenHR</title>
    <script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
    <script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
    <script src="<%: Url.LatestContent("~/bundles/Microsoft")%>" type="text/javascript"></script>
    <script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>

    <link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
    <script id="officebarscript" src="<%: Url.LatestContent("~/Scripts/officebar/jquery.officebar.js")%>" type="text/javascript"></script>

    <script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>
    <link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
    <link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
    <link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
    <link href="<%= Url.LatestContent("~/Content/font-awesome.min.css")%>" rel="stylesheet" type="text/css" />

    <%--jQuery Grid Stylesheet--%>
    <link href="<%: Url.LatestContent("~/Content/ui.jqgrid.css")%>" rel="stylesheet" type="text/css" />

    <%--Placeholders for theme and layout--%>
    <link id="layoutLink" rel="stylesheet" type="text/css" />
    <link id="themeLink" rel="stylesheet" type="text/css" />
    <link id="WireframethemeLink" rel="stylesheet" type="text/css" />

    <script type="text/javascript">
        window.onload = function () {
            //Get some cookies that we need to determine the CSS to apply
            var SSIMode = OpenHR.getCookie("SSIMode");
            var currentLayout = OpenHR.getCookie("currentLayout");
            var currentTheme = OpenHR.getCookie("currentTheme");
            var cookiewireframeTheme = OpenHR.getCookie("cookiewireframeTheme");
            var cookieapplyWireframeTheme = OpenHR.getCookie("cookieapplyWireframeTheme");

            if (($("#fixedlinksframe").length > 0) && (currentLayout != "winkit"))
                $("link[id=DMIthemeLink]").attr({ href: "" });

            //The logic below is taken from Site.Master, it should be abstracted somewhere else, but no time to do that now
            if (SSIMode != "True") {
                $("link[id=layoutLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/layouts/winkit.css")%>" });
                $("link[id=themeLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/themes/white.css")%>" });
                $('body').addClass('DMI');
            } else {
                switch (OpenHR.getCookie("Intranet_Layout")) {
                    case "winkit":
                        $("link[id=layoutLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/layouts/winkit.css")%>" });
                        $("link[id=SSIthemeLink]").attr({ href: "<%:Url.LatestContent("~/Content/themes/redmond-segoe/jquery-ui.min.css")%>" });
                        $("link[id=DMIthemeLink]").attr({ href: "<%:Url.LatestContent("~/Content/themes/redmond-segoe/jquery-ui.min.css")%>" });
                        break;
                    case "wireframe":
                        if (cookieapplyWireframeTheme == "true") $("link[id=WireframethemeLink]").attr({ href: "../Content/DashboardStyles/themes/upgraded.css" });

                        $("link[id=layoutLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/layouts/wireframe.css")%>" });
                        $("link[id=SSIthemeLink]").attr({ href: "../Content/themes/" + cookiewireframeTheme + "/jquery-ui.min.css" });
                        $("link[id=DMIthemeLink]").attr({ href: "../Content/themes/" + cookiewireframeTheme + "/jquery-ui.min.css" });
                        break;
                    case "tiles":
                        $("link[id=layoutLink]").attr({ href: "<%:Url.LatestContent("~/Content/DashboardStyles/layouts/tiles.css")%>" });
                        $("link[id=SSIthemeLink]").attr({ href: "<%:Url.LatestContent("~/Content/themes/start/jquery-ui.min.css")%>" });
                        $("link[id=DMIthemeLink]").attr({ href: "<%:Url.LatestContent("~/Content/themes/start/jquery-ui.min.css")%>" });
                        break;
                }

                switch (currentTheme) {
                    case "red":
                        $("link[id=themeLink]").attr({ href: "<%: Url.LatestContent("~/Content/DashboardStyles/themes/Red.css")%>" });
                    break;
                case "blue":
                    $("link[id=themeLink]").attr({ href: "<%: Url.LatestContent("~/Content/DashboardStyles/themes/Blue.css")%>" });
                    break;
                case "white":
                    $("link[id=themeLink]").attr({ href: "<%: Url.LatestContent("~/Content/DashboardStyles/themes/White.css")%>" });
                default:
                    break;
            }
        }
        };

    function setForm() {
        var frmPopup = document.getElementById("frmPopup");

        if (window.dialogArguments.document.getElementById('txtAbsenceEmailGroup') != null) {
            window.dialogArguments.document.getElementById('txtAbsenceEmailGroup').value = frmPopup.txtSelectedName.value;
            window.dialogArguments.document.getElementById('txtAbsenceEmailGroupID').value = frmPopup.txtSelectedID.value;
        }

        if (window.dialogArguments.document.getElementById('txtEmailGroup') != null) {
            window.dialogArguments.document.getElementById('txtEmailGroup').value = frmPopup.txtSelectedName.value;
            window.dialogArguments.document.getElementById('txtEmailGroupID').value = frmPopup.txtSelectedID.value;
        }

        self.close();
        return false;
    }
    </script>
</head>

<body id="bdyMain" name="bdyMain" <%=session("BodyColour")%> leftmargin="20" topmargin="20" bottommargin="20" rightmargin="20" style="overflow: hidden">
    <form id="frmPopup" name="frmPopup" onsubmit="return setForm();" style="visibility: hidden; display: none">
        <input type="hidden" id="txtSelectedID" name="txtSelectedID">
        <input type="hidden" id="txtSelectedName" name="txtSelectedName">
        <input type="hidden" id="txtSelectedAccess" name="txtSelectedAccess">
        <input type="hidden" id="txtSelectedUserName" name="txtSelectedUserName">
    </form>

    <div style="text-align: center">
        <h3></h3>
    </div>

    <div style="margin-left: 15px;">
        <%=GetEmailSelection()%>
    </div>

    <div style="margin-top: 10px; margin-right: 20px; float: right;">
        <input id="cmdok" type="button" value="OK" name="cmdok"
            style="width: 80px"
            class="button"
            onclick="emailEvent();" />
        <input id="cmdcancel" type="button" value="Cancel" name="cmdcancel"
            style="width: 80px"
            class="button"
            onclick="self.close();" />
    </div>

    <form name="frmEmailDetails" id="frmEmailDetails" style="visibility: hidden; display: none; width: 100%">
        <%
            'Get the required Email information
            Dim sErrorDescription As String = ""
            Dim sEmailInfo As String = vbNullString
            Dim iLastEventID As Integer = -1
            Dim iDetailCount As Integer = 0
            Dim EventCounter As Integer = 0
            Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

            Try
                Dim prmSubject = New SqlParameter("psSubject", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
                Dim rsEmailDetails = objDataAccess.GetFromSP("spASRIntGetEventLogEmailInfo" _
                , New SqlParameter("psSelectedIDs", SqlDbType.VarChar, -1) With {.Value = Request("txtSelectedEventIDs")} _
                , prmSubject _
                , New SqlParameter("psOrderColumn", SqlDbType.VarChar, -1) With {.Value = CStr(Request("txtEmailOrderColumn"))} _
                , New SqlParameter("psOrderOrder", SqlDbType.VarChar, -1) With {.Value = CStr(Request("txtEmailOrderOrder"))})
			
                If rsEmailDetails.Rows.Count > 0 Then
                    For Each objRow As DataRow In rsEmailDetails.Rows
                        If iLastEventID <> CInt(objRow("ID")) Then
                            EventCounter = EventCounter + 1
                            Response.Write(CStr(EventCounter))
                            sEmailInfo = sEmailInfo & StrDup(Len(objRow("Name").ToString()) + 30, "-") & vbCrLf
                            sEmailInfo = sEmailInfo & "Event Name : " & objRow("Name").ToString() & vbCrLf
                            sEmailInfo = sEmailInfo & StrDup(Len(objRow("Name").ToString()) + 30, "-") & vbCrLf
                            sEmailInfo = sEmailInfo & "Mode :		" & objRow("Mode").ToString() & vbCrLf & vbCrLf
                            sEmailInfo = sEmailInfo & "Start Time :	" & ConvertSQLDateToLocale(objRow("DateTime")) & " " & ConvertSqlDateToTime(objRow("DateTime")) & vbCrLf
                            If IsDBNull(objRow("EndTime")) Then
                                sEmailInfo = sEmailInfo & "End Time :	" & vbCrLf
                            Else
                                sEmailInfo = sEmailInfo & "End Time :	" & ConvertSQLDateToLocale(objRow("DateTime")) & " " & ConvertSqlDateToTime(objRow("EndTime")) & vbCrLf
                            End If
                            sEmailInfo = sEmailInfo & "Duration :	" & FormatEventDuration(CInt(objRow("Duration"))) & vbCrLf
                            sEmailInfo = sEmailInfo & "Type :		" & objRow("Type").ToString() & vbCrLf
                            sEmailInfo = sEmailInfo & "Status :		" & objRow("Status").ToString() & vbCrLf
                            sEmailInfo = sEmailInfo & "User name :	" & objRow("Username").ToString() & vbCrLf & vbCrLf
                            If Request("txtFromMain") = 0 Then
                                If Request("txtBatchy") Then
                                    sEmailInfo = sEmailInfo & Request("txtBatchInfo") & vbCrLf
                                End If
                            Else
                                If (Not IsDBNull(objRow("BatchName"))) And (Len(objRow("BatchName").ToString()) > 0) Then
                                    sEmailInfo = sEmailInfo & "Batch Job Name	: " & objRow("BatchName").ToString() & vbCrLf & vbCrLf
                                End If
                            End If
										
                            sEmailInfo = sEmailInfo & "Records Successful :	" & objRow("SuccessCount").ToString() & vbCrLf
                            sEmailInfo = sEmailInfo & "Records Failed :		" & objRow("FailCount").ToString() & vbCrLf & vbCrLf
                            sEmailInfo = sEmailInfo & "Details : " & vbCrLf & vbCrLf
                            iLastEventID = CInt(objRow("ID"))
                            iDetailCount = 0
                        End If
				
                        iDetailCount += 1
				
                        If objRow("count") > 0 Then
                            If (Not IsDBNull(objRow("Notes"))) And (Len(objRow("Notes")) > 0) Then
                                sEmailInfo = sEmailInfo & "*** Log Entry " & CStr(iDetailCount) & " of " & CStr(objRow("count")) & " ***" & vbCrLf
                                sEmailInfo = sEmailInfo & objRow("Notes").ToString()
                            End If
                        Else
                            sEmailInfo = sEmailInfo & "There are no details for this event log entry" & vbCrLf
                        End If
                        sEmailInfo = sEmailInfo & vbCrLf & vbCrLf & vbCrLf
                    Next
                    Response.Write("<input  name=txtEventDeleted id=txtEventDeleted value=0>" & vbCrLf)
                Else
                    Response.Write("<input  name=txtEventDeleted id=txtEventDeleted value=1>" & vbCrLf)
                End If

                Response.Write("<input  name=txtBody id=txtBody value=""" & Replace(sEmailInfo, """", "&quot;") & """>" & vbCrLf)
                Response.Write("<input  name=txtSubject id=txtSubject value=""" & Replace(prmSubject.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
            Catch ex As Exception
                sErrorDescription = "Error getting the event log records." & vbCrLf & FormatError(ex.Message)
            End Try
        %>
    </form>

    <form id="frmFromOpener" name="frmFromOpener" style="visibility: hidden; display: none">
        <input type="hidden" id="calcEmailCurrentID" name="calcEmailCurrentID" value='<%= Request("emailSelCurrentID") %>'>
    </form>

    <input type="hidden" id="txtTicker" name="txtTicker" value="0">
    <input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">
</body>

<script type="text/javascript">
    // Table to jQuery grid
    tableToGrid("#EmailSelectionTable", {
        multiselect: true,
        onSelectRow: function (rowID) { refreshControls(); },
        onSelectAll: function (rowID) { refreshControls(); },
        ondblClickRow: function (rowID) { },
        colNames: ['EmailGroupIDHeader', 'Name', 'FullNameHeader', 'To', 'Cc', 'Bcc'],
        colModel: [
			{ name: 'EmailGroupIDHeader', hidden: true },
			{ name: 'NameHeader', sortable: false },
            { name: 'FullNameHeader', sortable: false, hidden: true },
            { name: 'to', edittype: 'checkbox', index: 'to', editoptions: { value: "True:False" }, formatter: 'checkbox', formatoptions: { disabled: false }, align: 'center', width: 20 },
            { name: 'cc', edittype: 'checkbox', index: 'cc', editoptions: { value: "True:False" }, formatter: 'checkbox', formatoptions: { disabled: false }, align: 'center', width: 20 },
            { name: 'bcc', edittype: 'checkbox', index: 'bcc', editoptions: { value: "True:False" }, formatter: 'checkbox', formatoptions: { disabled: false }, align: 'center', width: 20 }
        ],
        rowNum: 1000,   //TODO set this to blocksize...
        height: 320,
        width: (screen.width) / 3 + 5,
        beforeSelectRow: function (rowid, e) { return false; }
    });

    function refreshControls() {
        var sSelectionList = jQuery("#EmailSelectionTable").jqGrid('getGridParam', 'selarrrow');
        sSelectionList = (sSelectionList == null ? '' : sSelectionList);
    }

    function emailEvent() {
        var sTo = getEmails(5);
        var sCC = getEmails(6);
        var sBCC = getEmails(7);
        var sSubject = getSubject();
        var sBody = getBody();
        window.dialogArguments.OpenHR.sendMail(sTo, sSubject, sBody, sCC, sBCC);
        self.close();
        return true;
    }

    function getEmails(typeIndex) {
        var localList = ""
        $('#EmailSelectionTable').find('td:nth-child(' + typeIndex + ')').each(function () {
            $(this).find('input:checked').each(function () {
                localList += $(this).parent().parent().find('td:nth-child(' + 4 + ')').text() + ";";
            });
        });
        return localList;
    }

    function getSubject() {
        return frmEmailDetails.txtSubject.value;
    }

    function getBody() {
        return frmEmailDetails.txtBody.value;
    }

    $(".button").button();
    //Hide the EmailGroup table header and its column
    $('.ui-jqgrid-htable tr th:nth-child(1)').hide();
    $('#EmailSelectionTable tr td:nth-child(1)').hide();
</script>

</html>
