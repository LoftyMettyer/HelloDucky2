<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<script type="text/javascript">

	$(document).ready(function () {

		//Read items from existing stylesheet if required:
		var currentFontWeight = "normal";
		if ($(".ui-state-default").css("font-weight") > 400) currentFontWeight = "bold";

		//Set theme editor options to existing stylesheet		
		$("#fwDefault").val(currentFontWeight);
		$("#fsDefault").val($(".ui-widget").css("font-size"));  //note: always shown in px :S
		$("#ffDefault").val($(".ui-widget").css("font-family"));
		$("#cornerRadius").val($(".ui-corner-all").css("border-bottom-left-radius"));

		//Change events
		$("#fwDefault").change(function () {
			if ($(this).val() == "bold") {
				$(".ui-state-default").css("font-weight", "bold");
				$(".ui-widget-content").css("font-weight", "bold");
				$(".ui-widget-header").css("font-weight", "bold");
			} else {
				$(".ui-state-default").css("font-weight", "normal");
				$(".ui-widget-content").css("font-weight", "normal");
				$(".ui-widget-header").css("font-weight", "normal");
			}
		});

		$("#fsDefault").change(function () {
			$(".ui-widget").css("font-size", $("#fsDefault").val());
		});

		$("#ffDefault").change(function () {
			$(".ui-widget").css("font-family", $("#ffDefault").val());
		});

		$("#cornerRadius").change(function () {
			$(".ui-corner-all, .ui-corner-top, .ui-corner-left, .ui-corner-tl").css("border-radius", $("#cornerRadius").val());
			$(".ui-corner-right, .ui-corner-tr").css("border-radius", $("#cornerRadius").val());
			$(".ui-corner-bottom, .ui-corner-left, .ui-corner-bl").css("border-radius", $("#cornerRadius").val());
			$(".ui-corner-right, .ui-corner-br").css("border-radius", $("#cornerRadius").val());
		});

		$("input[type=submit], input[type=button], button, input[type=file]").button();

		if (window.cookieapplyWireframeTheme == "true") $('#chkAppywireframetheme').prop('checked', true);

	});

	function themeEditor_window_onload() {
		 
	  document.getElementById('cmbLayout').value = window.IntranetLayout;
		document.getElementById('cmbTheme').value = window.IntranetWireframeTheme;
		toggleCombos();

	}

	function saveLayoutandTheme() {

		try { changeLayout($("#cmbLayout :selected").text()); } catch (e) { }		
			if ($("#cmbLayout :selected").val() == "wireframe") {	      
					try { changeTheme($("#cmbTheme :selected").val()); } catch (e) { }      
				 
			}

			$(".DashContent").fadeOut("slow");
			$(".DashContent").promise().done(function () {
			  window.location = "MainSSI";
			});

	}

	function toggleCombos() {
			
			//theme selection is allowed in wirefamemode only
			if ($("#cmbLayout :selected").text() != "") {
			      button_disable($("#btnDiv2OK")[0], false);
					if ($("#cmbLayout :selected").text() != "wireframe") {
							document.getElementById("cmbTheme").disabled = true;
							document.getElementById("cmbTheme").value = "";
					}
					else {
							document.getElementById("cmbTheme").disabled = false;
							if ($("#cmbTheme :selected").text() == "") {
							   button_disable($("#btnDiv2OK")[0], true);
			      	  }
							else {
							   button_disable($("#btnDiv2OK")[0], false);
							}
					}
			}
			else {
					document.getElementById("cmbTheme").value = "";
					button_disable($("#btnDiv2OK")[0], true);
			}
	}

	function themeEditor_cancelClick() {
		 $("#divthemeRoller").dialog("close");
			return false;
	}

</script>

<style>
	.ffDefault {
		width: 100%;
	}

	.field-group {
		border-bottom: 1px dotted lightgray;
		padding: 2px;
	}
</style>
<form method="POST" action="importTheme_Submit" name="themeroller" id="themeroller" enctype="multipart/form-data">
	<div class="application">
		<div id="themeeditoraccordion">
			<p>Choose a predefined theme:</p>
			<span>Layout:<select id="cmbLayout" style="width: 220px; margin-left: 40px;" onchange="toggleCombos()"><option></option>
				<option value="wireframe">wireframe</option>
				<option value="winkit">winkit</option>
				<option value="tiles">tiles</option>
			</select></span>
			
			<br /><br />

			<span>Theme:<select id="cmbTheme" style="width: 220px; margin-left: 40px;" onchange="toggleCombos()"><option></option>
				<%
					for each strTheme in Session("ui-dynamic-themes")
						Response.Write("<OPTION VALUE = " & """" & strTheme & """" & ">" & strTheme & "</OPTION>")
					Next
				%>
			</select>
			</span>
			
			<br /><br /><br />
			<hr />
			
			<div id="divSaveButtons" style="text-align: right">
				<input id="btnDiv2OK" name="btnDiv2OK" type="button" class="btn" value="OK" onclick="saveLayoutandTheme()" />
				<input id="btnDiv2Cancel" name="btnDiv2Cancel" type="button" class="btn" value="Cancel" onclick="themeEditor_cancelClick()" />
			</div>
		</div>
	</div>
	<%=Html.AntiForgeryToken()%>
</form>

<script type="text/javascript">
		themeEditor_window_onload();

</script>
