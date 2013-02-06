<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>


<%
    ' following sessions vars:
'
' UtilType    - 0-13 (see UtilityType code in DATMGR .exe
' UtilName    - <the name of the utility>
' UtilID      - <the id of the utility>
' Action      - run/delete
session("utiltype") = Request.Form("utiltype")
session("utilid") = Request.Form("utilid")
session("utilname") = Request.Form("utilname")
session("action") = Request.Form("action")

' Write the prompted values from the calling form into a session variable.
' NB. The prompts are written into an array and this array is written to a 
    ' session variables with the name 'Prompts_<util type>_<util id>.
    Dim sKey

    Dim aPrompts(1, 0)
    Dim j = 0
redim preserve aPrompts(1, 0)
    For i = 0 To (Request.Form.Count) - 1
        sKey = Request.Form.Keys(i)
        If ((UCase(Left(sKey, 7)) = "PROMPT_") And (Mid(sKey, 8, 1) <> "3")) Or _
            (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
            ReDim Preserve aPrompts(1, j)
		
            If (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
                aPrompts(0, j) = "prompt_3_" & Mid(sKey, 11)
                aPrompts(1, j) = UCase(Request.Form.Item(i))
            Else
                aPrompts(0, j) = sKey
                Select Case Mid(sKey, 8, 1)
                    Case "2"
                        ' Numeric. Replace locale decimal point with '.'
                        aPrompts(1, j) = Replace(Request.Form.Item(i), Session("LocaleDecimalSeparator"), ".")
                    Case "4"
                        ' Date. Reformat to match SQL's mm/dd/yyyy format.
                        'aPrompts(1, j) = convertLocaleDateToSQL(Request.Form.Item(i))
                        'aPrompts(1, j) = convertLocaleDateToSQL(Request.Form.Item(i))
                        'TODO convertdatetosqlformat - function needs to be central - can't append to end of files anymore!
                    Case Else
                        aPrompts(1, j) = Request.Form.Item(i)
                End Select
            End If
            j = j + 1
        End If
    Next
sKey = "Prompts_" & Request.Form("utiltype") & "_" & Request.Form("utilid") 
session(sKey) = aPrompts
%>


    
    
<script type="text/javascript">

    function util_run_window_onload() {
        $("#workframe").attr("data-framesource", "UTIL_RUN");
    }    

    function raiseError(sErrorDesc, fok, fcancelled) 
    {
        frmError.txtUtilTypeDesc.value = window.frames("top").frmPopup.txtUtilTypeDesc.value;
        frmError.txtErrorDesc.value = sErrorDesc;
        frmError.txtOK.value = fok;
        frmError.txtUserCancelled.value = fcancelled;
        var sTarget = new String("errorMessage");
        frmError.target = sTarget;
        NewWindow('', sTarget,'500','200','no');
        frmError.submit();
        self.close();
        return;
    }

    function pausecomp(millis) 
    {
        var date = new Date();
        var curDate = null;

        do 
        { 
            curDate = new Date(); 
        } while(curDate-date < millis);
    } 

    function NewWindow(mypage, myname, w, h, scroll) 
    {
        var winl = (screen.width - w) / 2;
        var wint = (screen.height - h) / 2;
        winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resizable';
        win = window.open(mypage, myname, winprops);

        if (parseInt(navigator.appVersion) >= 4) 
        {
            pausecomp(300);
            win.window.focus(); 
        }
    }

    function ShowWaitFrame(sMessage)
    {
        var fs = window.parent.document.all.item("myframeset");

        if (fs) 
        {
            fs.rows = "*,0,0";
        }
	
        try 
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);		
        } 
        catch(e) {}
    }

    function ShowOutputOptionsFrame(sURL)
    {
        //frames("outputoptions").location.replace(sURL);
        var fsOptions = window.parent.document.all.item("outputoptions");
        if (fsOptions)	
        {
            fsOptions.src = sURL
        }

        var fs = window.parent.document.all.item("myframeset");
        if (fs) {
            fs.rows = "0,0,*";
        }

        try 
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);		
        } 
        catch(e) {}
    }

    function ShowDataFrame()
    {
        var fs = window.parent.document.all.item("myframeset");
        if (fs) 
        {
            fs.rows = "0,*,0";
        }
	
        try 
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);		
        } 
        catch(e) {}
    }

</script>

<FORM id=frmError name=frmError action="util_run_error" method=post>
	<INPUT type="hidden" id=txtUtilTypeDesc name=txtUtilTypeDesc>
	<INPUT type="hidden" id=txtEventLogID name=txtEventLogID>
	<INPUT type="hidden" id=txtOK name=txtOK>
	<INPUT type="hidden" id=txtUserCancelled name=txtUserCancelled>
	<INPUT type="hidden" id=txtErrorDesc name=txtErrorDesc>
</FORM>

<div id="reportframeset">

    <div id="top">       
        <%html.RenderPartial("~/Views/Home/progress.ascx")%>
    </div>
    
    <div id="main" data-framesource="util_run">
    <%   
        If Session("utiltype") = "1" Then
            Html.RenderPartial("~/Views/Home/util_run_crosstabsMain.ascx")
        ElseIf Session("utiltype") = "2" Then
            Html.RenderPartial("~/Views/Home/util_run_customreportsMain.ascx")
        ElseIf Session("utiltype") = "3" Then
            Html.RenderPartial("~/Views/Home/util_run_datatransfer.ascx")
        ElseIf Session("utiltype") = "4" Then
            Html.RenderPartial("~/Views/Home/util_run_export.ascx")
        ElseIf Session("utiltype") = "5" Then
            Html.RenderPartial("~/Views/Home/util_run_globaladd.ascx")
        ElseIf Session("utiltype") = "6" Then
            Html.RenderPartial("~/Views/Home/util_run_globalupdate.ascx")
        ElseIf Session("utiltype") = "7" Then
            Html.RenderPartial("~/Views/Home/util_run_globaldelete.ascx")
        ElseIf Session("utiltype") = "8" Then
            Html.RenderPartial("~/Views/Home/util_run_import.ascx")
        ElseIf Session("utiltype") = "9" Then
            Html.RenderPartial("~/Views/Home/util_run_mailmerge.ascx")
        ElseIf Session("utiltype") = "15" Then
            Html.RenderPartial("~/Views/Home/stdrpt_run_AbsenceBreakdown.ascx")
        ElseIf Session("utiltype") = "16" Then
            Html.RenderPartial("~/Views/Home/util_run_customreportsMain.ascx")
        ElseIf Session("utiltype") = "17" Then
            Html.RenderPartial("~/Views/Home/util_run_calendarreport_main.ascx")
        End If
    %>
    </div>

	<div id="outputoptions"></div>

    <form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
        <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
    </form>

</div>
