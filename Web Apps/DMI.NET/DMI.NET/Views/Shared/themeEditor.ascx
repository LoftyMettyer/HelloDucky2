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

		//Apply accordion functionality
		//$("#themeeditoraccordion").accordion({ heightStyle: "content" });
		
		if (window.cookieapplyWireframeTheme == "true") $('#chkAppywireframetheme').prop('checked', true);

	});

	function themeEditor_window_onload() {
		 
			var currentLayout = OpenHR.getCookie("Intranet_Layout");
			var currentTheme = OpenHR.getCookie("Intranet_Wireframe_Theme");
			
			if (currentLayout != null) {
					document.getElementById('cmbLayout').value = currentLayout;
			}

			if (currentTheme != null) {
					document.getElementById('cmbTheme').value = currentTheme;
			}	    
			toggleCombos();
	}

	function saveLayoutandTheme() {
		 
		try { changeLayout($("#cmbLayout :selected").text()); } catch (e) { }		
			if ($("#cmbLayout :selected").val() == "wireframe") {	      
					try { changeTheme($("#cmbTheme :selected").val()); } catch (e) { }      
				 
			}
			themeEditor_window_onload();
	}

	function toggleCombos() {
			
			//theme selection is allowed in wirefamemode only
			if ($("#cmbLayout :selected").text() != "") {
					document.getElementById("btnDiv2OK").disabled = false;
					if ($("#cmbLayout :selected").text() != "wireframe") {
							document.getElementById("cmbTheme").disabled = true;
							document.getElementById("cmbTheme").value = "";
					}
					else {
							document.getElementById("cmbTheme").disabled = false;
							if ($("#cmbTheme :selected").text() == "") {
									document.getElementById("btnDiv2OK").disabled = true;
							}
							else {
									document.getElementById("btnDiv2OK").disabled = false;
							}
					}
			}
			else {
					document.getElementById("cmbTheme").value = "";
					document.getElementById("btnDiv2OK").disabled = true;
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
		<div id="themeeditoraccordion" style="margin: 30px;">
			<p>Choose a predefined theme:</p>
			<span>Layout:<select id="cmbLayout" style="width: 150px; margin-left: 40px;" onChange="toggleCombos()"><option></option>
				<option value="wireframe">wireframe</option>
				<option value="winkit">winkit</option>
				<option value="tiles">tiles</option>
			</select></span>
			<br />
			<br />
			<span>Theme:<select  id="cmbTheme" style="width: 150px; margin-left: 40px;"  onChange="toggleCombos()"><option></option>
				<option value="ABS">ABS</option>
				<option value="activeX">ActiveX</option>
				<option value="black-tie">Black-Tie</option>
				<option value="blitzer">Blitzer</option>
				<option value="cupertino">Cupertino</option>
				<option value="dark-hive">Dark-Hive</option>
				<option value="dot-luv">Dot-Luv</option>
				<option value="eggplant">Eggplant</option>
				<option value="excite-bike">Excite-Bike</option>
				<option value="flick">Flick</option>
				<option value="hot-sneaks">Hot-Sneaks</option>
				<option value="humanity">Humanity</option>
				<option value="le-frog">le-Frog</option>
				<option value="mint-choc">Mint-Choc</option>
				<option value="overcast">Overcast</option>
				<option value="pepper-grinder">Pepper-Grinder</option>
				<option value="pink-pip">Pink-Pip</option>
				<option value="redmond">Redmond</option>
				<option value="redmond-segoe">Redmond-Segoe</option>
				<option value="smoothness">Smoothness</option>
				<option value="south-street">South-Street</option>
				<option value="start">Start</option>
				<option value="sunny">Sunny</option>
				<option value="swanky-purse">Swanky-Purse</option>
				<option value="trontastic">Trontastic</option>
				<option value="ui-darkness">ui-Darkness</option>
				<option value="ui-lightness">ui-Lightness</option>
				<option value="vader">Vader</option>
			</select>
			<br />
				
			</span>
			<hr />
			
						<div id ="divSaveButtons" style="text-align:right">
							
								<input id="btnDiv2OK" name="btnDiv2OK" type="button" class="btn" value="OK" onclick="saveLayoutandTheme()" />	
								<input id="btnDiv2Cancel" name="btnDiv2Cancel" type="button" class="btn" value="Cancel" onclick="themeEditor_cancelClick()"/>
						
								</div>
		</div>
	</div>
	<%=Html.AntiForgeryToken()%>
</form>
<!-- /themeroller -->

<script type="text/javascript">
		themeEditor_window_onload();

</script>
