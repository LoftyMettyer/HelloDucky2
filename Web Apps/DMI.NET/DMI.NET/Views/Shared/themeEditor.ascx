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

		$("#fsDefault").change(function() {
			$(".ui-widget").css("font-size", $("#fsDefault").val());
		});

		$("#ffDefault").change(function() {
			$(".ui-widget").css("font-family", $("#ffDefault").val());
		});
		
		$("#cornerRadius").change(function () {
			$(".ui-corner-all, .ui-corner-top, .ui-corner-left, .ui-corner-tl").css("border-radius", $("#cornerRadius").val());
			$(".ui-corner-right, .ui-corner-tr").css("border-radius", $("#cornerRadius").val());
			$(".ui-corner-bottom, .ui-corner-left, .ui-corner-bl").css("border-radius", $("#cornerRadius").val());
			$(".ui-corner-right, .ui-corner-br").css("border-radius", $("#cornerRadius").val());					
		});



		//Apply accordion functionality
		$("#themeeditoraccordion").accordion({ heightStyle: "content" });

	});

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

<div class="application">
	<div id="themeeditoraccordion">

		<h3>Font Settings</h3>
		<!-- /theme group header -->
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group clearfix">
				<label for="ffDefault">Family</label>
				<input name="ffDefault" class="ffDefault" id="ffDefault" type="text" size="8" value="Verdana,Arial,sans-serif">
			</div>			
			<div class="field-group clearfix">
				<label for="fwDefault">Weight</label>
				<select name="fwDefault" class="fwDefault" id="fwDefault">

					<option value="normal" selected="">normal</option>

					<option value="bold">bold</option>

				</select>
			</div>
			<div class="field-group clearfix">
				<label for="fsDefault">Size</label>
				<input name="fsDefault" class="fsDefault" id="fsDefault" type="text" size="3" value="1.1em">
			</div>
		</div>
		<!-- /theme group content -->

		<!-- /theme group -->

		<h3>Corner Radius</h3>
		<!-- /theme group header -->
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-corners clearfix">
				<label for="cornerRadius">Corners:</label>
				<input name="cornerRadius" class="cornerRadius" id="cornerRadius" type="text" value="4px">
			</div>
			<p id="cornerWarning"><em><strong>Note:</strong> ThemeRoller uses CSS3 border-radius for corner rounding, which is not supported by Internet Explorer 7 or 8.</em></p>
		</div>
		<!-- /theme group content -->



		<h3>Header/Toolbar</h3>
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-background clearfix">
				<label class="background" for="bgColorHeader">Background color &amp; texture</label>
				<div class="hasPicker">
					<input name="bgColorHeader" class="hex" id="bgColorHeader" style="color: rgb(0, 0, 0); background-color: rgb(204, 204, 204);" type="text" value="#cccccc">
				</div>
				<select name="bgTextureHeader" class="texture">

					<option value="flat" data-textureheight="100" data-texturewidth="40">flat</option>

					<option value="glass" data-textureheight="400" data-texturewidth="1">glass</option>

					<option selected='"selected"' value="highlight_soft" data-textureheight="100" data-texturewidth="1">highlight_soft</option>

					<option value="highlight_hard" data-textureheight="100" data-texturewidth="1">highlight_hard</option>

					<option value="inset_soft" data-textureheight="100" data-texturewidth="1">inset_soft</option>

					<option value="inset_hard" data-textureheight="100" data-texturewidth="1">inset_hard</option>

					<option value="diagonals_small" data-textureheight="40" data-texturewidth="40">diagonals_small</option>

					<option value="diagonals_medium" data-textureheight="40" data-texturewidth="40">diagonals_medium</option>

					<option value="diagonals_thick" data-textureheight="40" data-texturewidth="40">diagonals_thick</option>

					<option value="dots_small" data-textureheight="2" data-texturewidth="2">dots_small</option>

					<option value="dots_medium" data-textureheight="4" data-texturewidth="4">dots_medium</option>

					<option value="white_lines" data-textureheight="100" data-texturewidth="40">white_lines</option>

					<option value="gloss_wave" data-textureheight="100" data-texturewidth="500">gloss_wave</option>

					<option value="diamond" data-textureheight="8" data-texturewidth="10">diamond</option>

					<option value="loop" data-textureheight="21" data-texturewidth="21">loop</option>

					<option value="carbon_fiber" data-textureheight="9" data-texturewidth="8">carbon_fiber</option>

					<option value="diagonal_maze" data-textureheight="10" data-texturewidth="10">diagonal_maze</option>

					<option value="diamond_ripple" data-textureheight="22" data-texturewidth="22">diamond_ripple</option>

					<option value="hexagon" data-textureheight="10" data-texturewidth="12">hexagon</option>

					<option value="layered_circles" data-textureheight="13" data-texturewidth="13">layered_circles</option>

					<option value="3D_boxes" data-textureheight="10" data-texturewidth="12">3D_boxes</option>

					<option value="glow_ball" data-textureheight="16" data-texturewidth="16">glow_ball</option>

					<option value="spotlight" data-textureheight="16" data-texturewidth="16">spotlight</option>

					<option value="fine_grain" data-textureheight="60" data-texturewidth="60">fine_grain</option>

				</select><div title="highlight_soft" class="texturePicker" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);'>
					<a href="#"></a>
					<ul style="display: none;">
						<li class="flat" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="flat" href="#">flat</a></li>
						<li class="glass" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);' data-textureheight="400" data-texturewidth="1"><a title="glass" href="#">glass</a></li>
						<li class="highlight_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_soft" href="#">highlight_soft</a></li>
						<li class="highlight_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_hard" href="#">highlight_hard</a></li>
						<li class="inset_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_soft" href="#">inset_soft</a></li>
						<li class="inset_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_hard" href="#">inset_hard</a></li>
						<li class="diagonals_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-small_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_small" href="#">diagonals_small</a></li>
						<li class="diagonals_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-medium_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_medium" href="#">diagonals_medium</a></li>
						<li class="diagonals_thick" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-thick_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_thick" href="#">diagonals_thick</a></li>
						<li class="dots_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-small_100_555_2x2.png") 50% 50% rgb(85, 85, 85);' data-textureheight="2" data-texturewidth="2"><a title="dots_small" href="#">dots_small</a></li>
						<li class="dots_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-medium_100_555_4x4.png") 50% 50% rgb(85, 85, 85);' data-textureheight="4" data-texturewidth="4"><a title="dots_medium" href="#">dots_medium</a></li>
						<li class="white_lines" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_white-lines_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="white_lines" href="#">white_lines</a></li>
						<li class="gloss_wave" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_gloss-wave_100_555_500x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="500"><a title="gloss_wave" href="#">gloss_wave</a></li>
						<li class="diamond" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond_100_555_10x8.png") 50% 50% rgb(85, 85, 85);' data-textureheight="8" data-texturewidth="10"><a title="diamond" href="#">diamond</a></li>
						<li class="loop" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_loop_100_555_21x21.png") 50% 50% rgb(85, 85, 85);' data-textureheight="21" data-texturewidth="21"><a title="loop" href="#">loop</a></li>
						<li class="carbon_fiber" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_carbon-fiber_100_555_8x9.png") 50% 50% rgb(85, 85, 85);' data-textureheight="9" data-texturewidth="8"><a title="carbon_fiber" href="#">carbon_fiber</a></li>
						<li class="diagonal_maze" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonal-maze_100_555_10x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="10"><a title="diagonal_maze" href="#">diagonal_maze</a></li>
						<li class="diamond_ripple" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond-ripple_100_555_22x22.png") 50% 50% rgb(85, 85, 85);' data-textureheight="22" data-texturewidth="22"><a title="diamond_ripple" href="#">diamond_ripple</a></li>
						<li class="hexagon" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_hexagon_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="hexagon" href="#">hexagon</a></li>
						<li class="layered_circles" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_layered-circles_100_555_13x13.png") 50% 50% rgb(85, 85, 85);' data-textureheight="13" data-texturewidth="13"><a title="layered_circles" href="#">layered_circles</a></li>
						<li class="3D_boxes" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_3D-boxes_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="3D_boxes" href="#">3D_boxes</a></li>
						<li class="glow_ball" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glow-ball_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="glow_ball" href="#">glow_ball</a></li>
						<li class="spotlight" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_spotlight_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="spotlight" href="#">spotlight</a></li>
						<li class="fine_grain" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_fine-grain_100_555_60x60.png") 50% 50% rgb(85, 85, 85);' data-textureheight="60" data-texturewidth="60"><a title="fine_grain" href="#">fine_grain</a></li>
					</ul>
				</div>
				<input name="bgImgOpacityHeader" class="opacity" type="text" value="75">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group field-group-border clearfix">
				<label for="borderColorHeader">Border</label>
				<div class="hasPicker">
					<input name="borderColorHeader" class="hex" id="borderColorHeader" style="color: rgb(0, 0, 0); background-color: rgb(170, 170, 170);" type="text" size="6" value="#aaaaaa">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="fcHeader">Text</label>
				<div class="hasPicker">
					<input name="fcHeader" class="hex" id="fcHeader" style="color: rgb(255, 255, 255); background-color: rgb(34, 34, 34);" type="text" size="6" value="#222222">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="iconColorHeader">Icon</label>
				<div class="hasPicker">
					<input name="iconColorHeader" class="hex" id="iconColorHeader" style="color: rgb(255, 255, 255); background-color: rgb(34, 34, 34);" type="text" size="6" value="#222222">
				</div>
			</div>
		</div>
		<!-- /theme group content -->


		<h3>Content</h3>
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-background clearfix">
				<label class="background" for="bgColorContent">Background color &amp; texture</label>
				<div class="hasPicker">
					<input name="bgColorContent" class="hex" id="bgColorContent" style="color: rgb(0, 0, 0); background-color: rgb(255, 255, 255);" type="text" value="#ffffff">
				</div>
				<select name="bgTextureContent" class="texture">

					<option selected='"selected"' value="flat" data-textureheight="100" data-texturewidth="40">flat</option>

					<option value="glass" data-textureheight="400" data-texturewidth="1">glass</option>

					<option value="highlight_soft" data-textureheight="100" data-texturewidth="1">highlight_soft</option>

					<option value="highlight_hard" data-textureheight="100" data-texturewidth="1">highlight_hard</option>

					<option value="inset_soft" data-textureheight="100" data-texturewidth="1">inset_soft</option>

					<option value="inset_hard" data-textureheight="100" data-texturewidth="1">inset_hard</option>

					<option value="diagonals_small" data-textureheight="40" data-texturewidth="40">diagonals_small</option>

					<option value="diagonals_medium" data-textureheight="40" data-texturewidth="40">diagonals_medium</option>

					<option value="diagonals_thick" data-textureheight="40" data-texturewidth="40">diagonals_thick</option>

					<option value="dots_small" data-textureheight="2" data-texturewidth="2">dots_small</option>

					<option value="dots_medium" data-textureheight="4" data-texturewidth="4">dots_medium</option>

					<option value="white_lines" data-textureheight="100" data-texturewidth="40">white_lines</option>

					<option value="gloss_wave" data-textureheight="100" data-texturewidth="500">gloss_wave</option>

					<option value="diamond" data-textureheight="8" data-texturewidth="10">diamond</option>

					<option value="loop" data-textureheight="21" data-texturewidth="21">loop</option>

					<option value="carbon_fiber" data-textureheight="9" data-texturewidth="8">carbon_fiber</option>

					<option value="diagonal_maze" data-textureheight="10" data-texturewidth="10">diagonal_maze</option>

					<option value="diamond_ripple" data-textureheight="22" data-texturewidth="22">diamond_ripple</option>

					<option value="hexagon" data-textureheight="10" data-texturewidth="12">hexagon</option>

					<option value="layered_circles" data-textureheight="13" data-texturewidth="13">layered_circles</option>

					<option value="3D_boxes" data-textureheight="10" data-texturewidth="12">3D_boxes</option>

					<option value="glow_ball" data-textureheight="16" data-texturewidth="16">glow_ball</option>

					<option value="spotlight" data-textureheight="16" data-texturewidth="16">spotlight</option>

					<option value="fine_grain" data-textureheight="60" data-texturewidth="60">fine_grain</option>

				</select><div title="flat" class="texturePicker" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);'>
					<a href="#"></a>
					<ul style="display: none;">
						<li class="flat" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="flat" href="#">flat</a></li>
						<li class="glass" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);' data-textureheight="400" data-texturewidth="1"><a title="glass" href="#">glass</a></li>
						<li class="highlight_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_soft" href="#">highlight_soft</a></li>
						<li class="highlight_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_hard" href="#">highlight_hard</a></li>
						<li class="inset_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_soft" href="#">inset_soft</a></li>
						<li class="inset_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_hard" href="#">inset_hard</a></li>
						<li class="diagonals_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-small_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_small" href="#">diagonals_small</a></li>
						<li class="diagonals_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-medium_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_medium" href="#">diagonals_medium</a></li>
						<li class="diagonals_thick" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-thick_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_thick" href="#">diagonals_thick</a></li>
						<li class="dots_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-small_100_555_2x2.png") 50% 50% rgb(85, 85, 85);' data-textureheight="2" data-texturewidth="2"><a title="dots_small" href="#">dots_small</a></li>
						<li class="dots_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-medium_100_555_4x4.png") 50% 50% rgb(85, 85, 85);' data-textureheight="4" data-texturewidth="4"><a title="dots_medium" href="#">dots_medium</a></li>
						<li class="white_lines" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_white-lines_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="white_lines" href="#">white_lines</a></li>
						<li class="gloss_wave" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_gloss-wave_100_555_500x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="500"><a title="gloss_wave" href="#">gloss_wave</a></li>
						<li class="diamond" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond_100_555_10x8.png") 50% 50% rgb(85, 85, 85);' data-textureheight="8" data-texturewidth="10"><a title="diamond" href="#">diamond</a></li>
						<li class="loop" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_loop_100_555_21x21.png") 50% 50% rgb(85, 85, 85);' data-textureheight="21" data-texturewidth="21"><a title="loop" href="#">loop</a></li>
						<li class="carbon_fiber" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_carbon-fiber_100_555_8x9.png") 50% 50% rgb(85, 85, 85);' data-textureheight="9" data-texturewidth="8"><a title="carbon_fiber" href="#">carbon_fiber</a></li>
						<li class="diagonal_maze" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonal-maze_100_555_10x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="10"><a title="diagonal_maze" href="#">diagonal_maze</a></li>
						<li class="diamond_ripple" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond-ripple_100_555_22x22.png") 50% 50% rgb(85, 85, 85);' data-textureheight="22" data-texturewidth="22"><a title="diamond_ripple" href="#">diamond_ripple</a></li>
						<li class="hexagon" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_hexagon_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="hexagon" href="#">hexagon</a></li>
						<li class="layered_circles" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_layered-circles_100_555_13x13.png") 50% 50% rgb(85, 85, 85);' data-textureheight="13" data-texturewidth="13"><a title="layered_circles" href="#">layered_circles</a></li>
						<li class="3D_boxes" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_3D-boxes_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="3D_boxes" href="#">3D_boxes</a></li>
						<li class="glow_ball" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glow-ball_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="glow_ball" href="#">glow_ball</a></li>
						<li class="spotlight" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_spotlight_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="spotlight" href="#">spotlight</a></li>
						<li class="fine_grain" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_fine-grain_100_555_60x60.png") 50% 50% rgb(85, 85, 85);' data-textureheight="60" data-texturewidth="60"><a title="fine_grain" href="#">fine_grain</a></li>
					</ul>
				</div>
				<input name="bgImgOpacityContent" class="opacity" type="text" value="75">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group field-group-border clearfix">
				<label for="borderColorContent">Border</label>
				<div class="hasPicker">
					<input name="borderColorContent" class="hex" id="borderColorContent" style="color: rgb(0, 0, 0); background-color: rgb(170, 170, 170);" type="text" size="6" value="#aaaaaa">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="fcContent">Text</label>
				<div class="hasPicker">
					<input name="fcContent" class="hex" id="fcContent" style="color: rgb(255, 255, 255); background-color: rgb(34, 34, 34);" type="text" size="6" value="#222222">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="iconColorContent">Icon</label>
				<div class="hasPicker">
					<input name="iconColorContent" class="hex" id="iconColorContent" style="color: rgb(255, 255, 255); background-color: rgb(34, 34, 34);" type="text" size="6" value="#222222">
				</div>
			</div>
		</div>
		<!-- /theme group content -->

		<h3>Clickable: default state</h3>
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-background clearfix">
				<label class="background" for="bgColorDefault">Background color &amp; texture</label>
				<div class="hasPicker">
					<input name="bgColorDefault" class="hex" id="bgColorDefault" style="color: rgb(0, 0, 0); background-color: rgb(230, 230, 230);" type="text" value="#e6e6e6">
				</div>
				<select name="bgTextureDefault" class="texture">

					<option value="flat" data-textureheight="100" data-texturewidth="40">flat</option>

					<option selected='"selected"' value="glass" data-textureheight="400" data-texturewidth="1">glass</option>

					<option value="highlight_soft" data-textureheight="100" data-texturewidth="1">highlight_soft</option>

					<option value="highlight_hard" data-textureheight="100" data-texturewidth="1">highlight_hard</option>

					<option value="inset_soft" data-textureheight="100" data-texturewidth="1">inset_soft</option>

					<option value="inset_hard" data-textureheight="100" data-texturewidth="1">inset_hard</option>

					<option value="diagonals_small" data-textureheight="40" data-texturewidth="40">diagonals_small</option>

					<option value="diagonals_medium" data-textureheight="40" data-texturewidth="40">diagonals_medium</option>

					<option value="diagonals_thick" data-textureheight="40" data-texturewidth="40">diagonals_thick</option>

					<option value="dots_small" data-textureheight="2" data-texturewidth="2">dots_small</option>

					<option value="dots_medium" data-textureheight="4" data-texturewidth="4">dots_medium</option>

					<option value="white_lines" data-textureheight="100" data-texturewidth="40">white_lines</option>

					<option value="gloss_wave" data-textureheight="100" data-texturewidth="500">gloss_wave</option>

					<option value="diamond" data-textureheight="8" data-texturewidth="10">diamond</option>

					<option value="loop" data-textureheight="21" data-texturewidth="21">loop</option>

					<option value="carbon_fiber" data-textureheight="9" data-texturewidth="8">carbon_fiber</option>

					<option value="diagonal_maze" data-textureheight="10" data-texturewidth="10">diagonal_maze</option>

					<option value="diamond_ripple" data-textureheight="22" data-texturewidth="22">diamond_ripple</option>

					<option value="hexagon" data-textureheight="10" data-texturewidth="12">hexagon</option>

					<option value="layered_circles" data-textureheight="13" data-texturewidth="13">layered_circles</option>

					<option value="3D_boxes" data-textureheight="10" data-texturewidth="12">3D_boxes</option>

					<option value="glow_ball" data-textureheight="16" data-texturewidth="16">glow_ball</option>

					<option value="spotlight" data-textureheight="16" data-texturewidth="16">spotlight</option>

					<option value="fine_grain" data-textureheight="60" data-texturewidth="60">fine_grain</option>

				</select><div title="glass" class="texturePicker" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);'>
					<a href="#"></a>
					<ul style="display: none;">
						<li class="flat" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="flat" href="#">flat</a></li>
						<li class="glass" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);' data-textureheight="400" data-texturewidth="1"><a title="glass" href="#">glass</a></li>
						<li class="highlight_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_soft" href="#">highlight_soft</a></li>
						<li class="highlight_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_hard" href="#">highlight_hard</a></li>
						<li class="inset_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_soft" href="#">inset_soft</a></li>
						<li class="inset_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_hard" href="#">inset_hard</a></li>
						<li class="diagonals_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-small_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_small" href="#">diagonals_small</a></li>
						<li class="diagonals_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-medium_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_medium" href="#">diagonals_medium</a></li>
						<li class="diagonals_thick" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-thick_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_thick" href="#">diagonals_thick</a></li>
						<li class="dots_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-small_100_555_2x2.png") 50% 50% rgb(85, 85, 85);' data-textureheight="2" data-texturewidth="2"><a title="dots_small" href="#">dots_small</a></li>
						<li class="dots_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-medium_100_555_4x4.png") 50% 50% rgb(85, 85, 85);' data-textureheight="4" data-texturewidth="4"><a title="dots_medium" href="#">dots_medium</a></li>
						<li class="white_lines" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_white-lines_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="white_lines" href="#">white_lines</a></li>
						<li class="gloss_wave" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_gloss-wave_100_555_500x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="500"><a title="gloss_wave" href="#">gloss_wave</a></li>
						<li class="diamond" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond_100_555_10x8.png") 50% 50% rgb(85, 85, 85);' data-textureheight="8" data-texturewidth="10"><a title="diamond" href="#">diamond</a></li>
						<li class="loop" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_loop_100_555_21x21.png") 50% 50% rgb(85, 85, 85);' data-textureheight="21" data-texturewidth="21"><a title="loop" href="#">loop</a></li>
						<li class="carbon_fiber" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_carbon-fiber_100_555_8x9.png") 50% 50% rgb(85, 85, 85);' data-textureheight="9" data-texturewidth="8"><a title="carbon_fiber" href="#">carbon_fiber</a></li>
						<li class="diagonal_maze" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonal-maze_100_555_10x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="10"><a title="diagonal_maze" href="#">diagonal_maze</a></li>
						<li class="diamond_ripple" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond-ripple_100_555_22x22.png") 50% 50% rgb(85, 85, 85);' data-textureheight="22" data-texturewidth="22"><a title="diamond_ripple" href="#">diamond_ripple</a></li>
						<li class="hexagon" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_hexagon_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="hexagon" href="#">hexagon</a></li>
						<li class="layered_circles" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_layered-circles_100_555_13x13.png") 50% 50% rgb(85, 85, 85);' data-textureheight="13" data-texturewidth="13"><a title="layered_circles" href="#">layered_circles</a></li>
						<li class="3D_boxes" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_3D-boxes_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="3D_boxes" href="#">3D_boxes</a></li>
						<li class="glow_ball" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glow-ball_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="glow_ball" href="#">glow_ball</a></li>
						<li class="spotlight" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_spotlight_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="spotlight" href="#">spotlight</a></li>
						<li class="fine_grain" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_fine-grain_100_555_60x60.png") 50% 50% rgb(85, 85, 85);' data-textureheight="60" data-texturewidth="60"><a title="fine_grain" href="#">fine_grain</a></li>
					</ul>
				</div>
				<input name="bgImgOpacityDefault" class="opacity" type="text" value="75">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group field-group-border clearfix">
				<label for="borderColorDefault">Border</label>
				<div class="hasPicker">
					<input name="borderColorDefault" class="hex" id="borderColorDefault" style="color: rgb(0, 0, 0); background-color: rgb(211, 211, 211);" type="text" size="6" value="#d3d3d3">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="fcDefault">Text</label>
				<div class="hasPicker">
					<input name="fcDefault" class="hex" id="fcDefault" style="color: rgb(255, 255, 255); background-color: rgb(85, 85, 85);" type="text" size="6" value="#555555">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="iconColorDefault">Icon</label>
				<div class="hasPicker">
					<input name="iconColorDefault" class="hex" id="iconColorDefault" style="color: rgb(0, 0, 0); background-color: rgb(136, 136, 136);" type="text" size="6" value="#888888">
				</div>
			</div>
		</div>
		<!-- /theme group content -->


		<h3>Clickable: hover state</h3>
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-background clearfix">
				<label class="background" for="bgColorHover">Background color &amp; texture</label>
				<div class="hasPicker">
					<input name="bgColorHover" class="hex" id="bgColorHover" style="color: rgb(0, 0, 0); background-color: rgb(218, 218, 218);" type="text" value="#dadada">
				</div>
				<select name="bgTextureHover" class="texture">

					<option value="flat" data-textureheight="100" data-texturewidth="40">flat</option>

					<option selected='"selected"' value="glass" data-textureheight="400" data-texturewidth="1">glass</option>

					<option value="highlight_soft" data-textureheight="100" data-texturewidth="1">highlight_soft</option>

					<option value="highlight_hard" data-textureheight="100" data-texturewidth="1">highlight_hard</option>

					<option value="inset_soft" data-textureheight="100" data-texturewidth="1">inset_soft</option>

					<option value="inset_hard" data-textureheight="100" data-texturewidth="1">inset_hard</option>

					<option value="diagonals_small" data-textureheight="40" data-texturewidth="40">diagonals_small</option>

					<option value="diagonals_medium" data-textureheight="40" data-texturewidth="40">diagonals_medium</option>

					<option value="diagonals_thick" data-textureheight="40" data-texturewidth="40">diagonals_thick</option>

					<option value="dots_small" data-textureheight="2" data-texturewidth="2">dots_small</option>

					<option value="dots_medium" data-textureheight="4" data-texturewidth="4">dots_medium</option>

					<option value="white_lines" data-textureheight="100" data-texturewidth="40">white_lines</option>

					<option value="gloss_wave" data-textureheight="100" data-texturewidth="500">gloss_wave</option>

					<option value="diamond" data-textureheight="8" data-texturewidth="10">diamond</option>

					<option value="loop" data-textureheight="21" data-texturewidth="21">loop</option>

					<option value="carbon_fiber" data-textureheight="9" data-texturewidth="8">carbon_fiber</option>

					<option value="diagonal_maze" data-textureheight="10" data-texturewidth="10">diagonal_maze</option>

					<option value="diamond_ripple" data-textureheight="22" data-texturewidth="22">diamond_ripple</option>

					<option value="hexagon" data-textureheight="10" data-texturewidth="12">hexagon</option>

					<option value="layered_circles" data-textureheight="13" data-texturewidth="13">layered_circles</option>

					<option value="3D_boxes" data-textureheight="10" data-texturewidth="12">3D_boxes</option>

					<option value="glow_ball" data-textureheight="16" data-texturewidth="16">glow_ball</option>

					<option value="spotlight" data-textureheight="16" data-texturewidth="16">spotlight</option>

					<option value="fine_grain" data-textureheight="60" data-texturewidth="60">fine_grain</option>

				</select><div title="glass" class="texturePicker" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);'>
					<a href="#"></a>
					<ul style="display: none;">
						<li class="flat" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="flat" href="#">flat</a></li>
						<li class="glass" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);' data-textureheight="400" data-texturewidth="1"><a title="glass" href="#">glass</a></li>
						<li class="highlight_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_soft" href="#">highlight_soft</a></li>
						<li class="highlight_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_hard" href="#">highlight_hard</a></li>
						<li class="inset_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_soft" href="#">inset_soft</a></li>
						<li class="inset_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_hard" href="#">inset_hard</a></li>
						<li class="diagonals_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-small_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_small" href="#">diagonals_small</a></li>
						<li class="diagonals_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-medium_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_medium" href="#">diagonals_medium</a></li>
						<li class="diagonals_thick" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-thick_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_thick" href="#">diagonals_thick</a></li>
						<li class="dots_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-small_100_555_2x2.png") 50% 50% rgb(85, 85, 85);' data-textureheight="2" data-texturewidth="2"><a title="dots_small" href="#">dots_small</a></li>
						<li class="dots_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-medium_100_555_4x4.png") 50% 50% rgb(85, 85, 85);' data-textureheight="4" data-texturewidth="4"><a title="dots_medium" href="#">dots_medium</a></li>
						<li class="white_lines" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_white-lines_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="white_lines" href="#">white_lines</a></li>
						<li class="gloss_wave" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_gloss-wave_100_555_500x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="500"><a title="gloss_wave" href="#">gloss_wave</a></li>
						<li class="diamond" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond_100_555_10x8.png") 50% 50% rgb(85, 85, 85);' data-textureheight="8" data-texturewidth="10"><a title="diamond" href="#">diamond</a></li>
						<li class="loop" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_loop_100_555_21x21.png") 50% 50% rgb(85, 85, 85);' data-textureheight="21" data-texturewidth="21"><a title="loop" href="#">loop</a></li>
						<li class="carbon_fiber" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_carbon-fiber_100_555_8x9.png") 50% 50% rgb(85, 85, 85);' data-textureheight="9" data-texturewidth="8"><a title="carbon_fiber" href="#">carbon_fiber</a></li>
						<li class="diagonal_maze" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonal-maze_100_555_10x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="10"><a title="diagonal_maze" href="#">diagonal_maze</a></li>
						<li class="diamond_ripple" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond-ripple_100_555_22x22.png") 50% 50% rgb(85, 85, 85);' data-textureheight="22" data-texturewidth="22"><a title="diamond_ripple" href="#">diamond_ripple</a></li>
						<li class="hexagon" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_hexagon_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="hexagon" href="#">hexagon</a></li>
						<li class="layered_circles" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_layered-circles_100_555_13x13.png") 50% 50% rgb(85, 85, 85);' data-textureheight="13" data-texturewidth="13"><a title="layered_circles" href="#">layered_circles</a></li>
						<li class="3D_boxes" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_3D-boxes_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="3D_boxes" href="#">3D_boxes</a></li>
						<li class="glow_ball" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glow-ball_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="glow_ball" href="#">glow_ball</a></li>
						<li class="spotlight" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_spotlight_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="spotlight" href="#">spotlight</a></li>
						<li class="fine_grain" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_fine-grain_100_555_60x60.png") 50% 50% rgb(85, 85, 85);' data-textureheight="60" data-texturewidth="60"><a title="fine_grain" href="#">fine_grain</a></li>
					</ul>
				</div>
				<input name="bgImgOpacityHover" class="opacity" type="text" value="75">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group field-group-border clearfix">
				<label for="borderColorHover">Border</label>
				<div class="hasPicker">
					<input name="borderColorHover" class="hex" id="borderColorHover" style="color: rgb(0, 0, 0); background-color: rgb(153, 153, 153);" type="text" size="6" value="#999999">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="fcHover">Text</label>
				<div class="hasPicker">
					<input name="fcHover" class="hex" id="fcHover" style="color: rgb(255, 255, 255); background-color: rgb(33, 33, 33);" type="text" size="6" value="#212121">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="iconColorHover">Icon</label>
				<div class="hasPicker">
					<input name="iconColorHover" class="hex" id="iconColorHover" style="color: rgb(255, 255, 255); background-color: rgb(69, 69, 69);" type="text" size="6" value="#454545">
				</div>
			</div>
		</div>
		<!-- /theme group content -->


		<h3>Clickable: active state</h3>
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-background clearfix">
				<label class="background" for="bgColorActive">Background color &amp; texture</label>
				<div class="hasPicker">
					<input name="bgColorActive" class="hex" id="bgColorActive" style="color: rgb(0, 0, 0); background-color: rgb(255, 255, 255);" type="text" value="#ffffff">
				</div>
				<select name="bgTextureActive" class="texture">

					<option value="flat" data-textureheight="100" data-texturewidth="40">flat</option>

					<option selected='"selected"' value="glass" data-textureheight="400" data-texturewidth="1">glass</option>

					<option value="highlight_soft" data-textureheight="100" data-texturewidth="1">highlight_soft</option>

					<option value="highlight_hard" data-textureheight="100" data-texturewidth="1">highlight_hard</option>

					<option value="inset_soft" data-textureheight="100" data-texturewidth="1">inset_soft</option>

					<option value="inset_hard" data-textureheight="100" data-texturewidth="1">inset_hard</option>

					<option value="diagonals_small" data-textureheight="40" data-texturewidth="40">diagonals_small</option>

					<option value="diagonals_medium" data-textureheight="40" data-texturewidth="40">diagonals_medium</option>

					<option value="diagonals_thick" data-textureheight="40" data-texturewidth="40">diagonals_thick</option>

					<option value="dots_small" data-textureheight="2" data-texturewidth="2">dots_small</option>

					<option value="dots_medium" data-textureheight="4" data-texturewidth="4">dots_medium</option>

					<option value="white_lines" data-textureheight="100" data-texturewidth="40">white_lines</option>

					<option value="gloss_wave" data-textureheight="100" data-texturewidth="500">gloss_wave</option>

					<option value="diamond" data-textureheight="8" data-texturewidth="10">diamond</option>

					<option value="loop" data-textureheight="21" data-texturewidth="21">loop</option>

					<option value="carbon_fiber" data-textureheight="9" data-texturewidth="8">carbon_fiber</option>

					<option value="diagonal_maze" data-textureheight="10" data-texturewidth="10">diagonal_maze</option>

					<option value="diamond_ripple" data-textureheight="22" data-texturewidth="22">diamond_ripple</option>

					<option value="hexagon" data-textureheight="10" data-texturewidth="12">hexagon</option>

					<option value="layered_circles" data-textureheight="13" data-texturewidth="13">layered_circles</option>

					<option value="3D_boxes" data-textureheight="10" data-texturewidth="12">3D_boxes</option>

					<option value="glow_ball" data-textureheight="16" data-texturewidth="16">glow_ball</option>

					<option value="spotlight" data-textureheight="16" data-texturewidth="16">spotlight</option>

					<option value="fine_grain" data-textureheight="60" data-texturewidth="60">fine_grain</option>

				</select><div title="glass" class="texturePicker" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);'>
					<a href="#"></a>
					<ul style="display: none;">
						<li class="flat" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="flat" href="#">flat</a></li>
						<li class="glass" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);' data-textureheight="400" data-texturewidth="1"><a title="glass" href="#">glass</a></li>
						<li class="highlight_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_soft" href="#">highlight_soft</a></li>
						<li class="highlight_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_hard" href="#">highlight_hard</a></li>
						<li class="inset_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_soft" href="#">inset_soft</a></li>
						<li class="inset_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_hard" href="#">inset_hard</a></li>
						<li class="diagonals_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-small_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_small" href="#">diagonals_small</a></li>
						<li class="diagonals_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-medium_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_medium" href="#">diagonals_medium</a></li>
						<li class="diagonals_thick" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-thick_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_thick" href="#">diagonals_thick</a></li>
						<li class="dots_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-small_100_555_2x2.png") 50% 50% rgb(85, 85, 85);' data-textureheight="2" data-texturewidth="2"><a title="dots_small" href="#">dots_small</a></li>
						<li class="dots_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-medium_100_555_4x4.png") 50% 50% rgb(85, 85, 85);' data-textureheight="4" data-texturewidth="4"><a title="dots_medium" href="#">dots_medium</a></li>
						<li class="white_lines" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_white-lines_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="white_lines" href="#">white_lines</a></li>
						<li class="gloss_wave" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_gloss-wave_100_555_500x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="500"><a title="gloss_wave" href="#">gloss_wave</a></li>
						<li class="diamond" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond_100_555_10x8.png") 50% 50% rgb(85, 85, 85);' data-textureheight="8" data-texturewidth="10"><a title="diamond" href="#">diamond</a></li>
						<li class="loop" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_loop_100_555_21x21.png") 50% 50% rgb(85, 85, 85);' data-textureheight="21" data-texturewidth="21"><a title="loop" href="#">loop</a></li>
						<li class="carbon_fiber" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_carbon-fiber_100_555_8x9.png") 50% 50% rgb(85, 85, 85);' data-textureheight="9" data-texturewidth="8"><a title="carbon_fiber" href="#">carbon_fiber</a></li>
						<li class="diagonal_maze" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonal-maze_100_555_10x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="10"><a title="diagonal_maze" href="#">diagonal_maze</a></li>
						<li class="diamond_ripple" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond-ripple_100_555_22x22.png") 50% 50% rgb(85, 85, 85);' data-textureheight="22" data-texturewidth="22"><a title="diamond_ripple" href="#">diamond_ripple</a></li>
						<li class="hexagon" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_hexagon_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="hexagon" href="#">hexagon</a></li>
						<li class="layered_circles" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_layered-circles_100_555_13x13.png") 50% 50% rgb(85, 85, 85);' data-textureheight="13" data-texturewidth="13"><a title="layered_circles" href="#">layered_circles</a></li>
						<li class="3D_boxes" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_3D-boxes_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="3D_boxes" href="#">3D_boxes</a></li>
						<li class="glow_ball" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glow-ball_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="glow_ball" href="#">glow_ball</a></li>
						<li class="spotlight" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_spotlight_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="spotlight" href="#">spotlight</a></li>
						<li class="fine_grain" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_fine-grain_100_555_60x60.png") 50% 50% rgb(85, 85, 85);' data-textureheight="60" data-texturewidth="60"><a title="fine_grain" href="#">fine_grain</a></li>
					</ul>
				</div>
				<input name="bgImgOpacityActive" class="opacity" type="text" value="65">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group field-group-border clearfix">
				<label for="borderColorActive">Border</label>
				<div class="hasPicker">
					<input name="borderColorActive" class="hex" id="borderColorActive" style="color: rgb(0, 0, 0); background-color: rgb(170, 170, 170);" type="text" size="6" value="#aaaaaa">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="fcActive">Text</label>
				<div class="hasPicker">
					<input name="fcActive" class="hex" id="fcActive" style="color: rgb(255, 255, 255); background-color: rgb(33, 33, 33);" type="text" size="6" value="#212121">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="iconColorActive">Icon</label>
				<div class="hasPicker">
					<input name="iconColorActive" class="hex" id="iconColorActive" style="color: rgb(255, 255, 255); background-color: rgb(69, 69, 69);" type="text" size="6" value="#454545">
				</div>
			</div>
		</div>
		<!-- /theme group content -->


		<h3>Highlight</h3>
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-background clearfix">
				<label class="background" for="bgColorHighlight">Background color &amp; texture</label>
				<div class="hasPicker">
					<input name="bgColorHighlight" class="hex" id="bgColorHighlight" style="color: rgb(0, 0, 0); background-color: rgb(251, 249, 238);" type="text" value="#fbf9ee">
				</div>
				<select name="bgTextureHighlight" class="texture">

					<option value="flat" data-textureheight="100" data-texturewidth="40">flat</option>

					<option selected='"selected"' value="glass" data-textureheight="400" data-texturewidth="1">glass</option>

					<option value="highlight_soft" data-textureheight="100" data-texturewidth="1">highlight_soft</option>

					<option value="highlight_hard" data-textureheight="100" data-texturewidth="1">highlight_hard</option>

					<option value="inset_soft" data-textureheight="100" data-texturewidth="1">inset_soft</option>

					<option value="inset_hard" data-textureheight="100" data-texturewidth="1">inset_hard</option>

					<option value="diagonals_small" data-textureheight="40" data-texturewidth="40">diagonals_small</option>

					<option value="diagonals_medium" data-textureheight="40" data-texturewidth="40">diagonals_medium</option>

					<option value="diagonals_thick" data-textureheight="40" data-texturewidth="40">diagonals_thick</option>

					<option value="dots_small" data-textureheight="2" data-texturewidth="2">dots_small</option>

					<option value="dots_medium" data-textureheight="4" data-texturewidth="4">dots_medium</option>

					<option value="white_lines" data-textureheight="100" data-texturewidth="40">white_lines</option>

					<option value="gloss_wave" data-textureheight="100" data-texturewidth="500">gloss_wave</option>

					<option value="diamond" data-textureheight="8" data-texturewidth="10">diamond</option>

					<option value="loop" data-textureheight="21" data-texturewidth="21">loop</option>

					<option value="carbon_fiber" data-textureheight="9" data-texturewidth="8">carbon_fiber</option>

					<option value="diagonal_maze" data-textureheight="10" data-texturewidth="10">diagonal_maze</option>

					<option value="diamond_ripple" data-textureheight="22" data-texturewidth="22">diamond_ripple</option>

					<option value="hexagon" data-textureheight="10" data-texturewidth="12">hexagon</option>

					<option value="layered_circles" data-textureheight="13" data-texturewidth="13">layered_circles</option>

					<option value="3D_boxes" data-textureheight="10" data-texturewidth="12">3D_boxes</option>

					<option value="glow_ball" data-textureheight="16" data-texturewidth="16">glow_ball</option>

					<option value="spotlight" data-textureheight="16" data-texturewidth="16">spotlight</option>

					<option value="fine_grain" data-textureheight="60" data-texturewidth="60">fine_grain</option>

				</select><div title="glass" class="texturePicker" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);'>
					<a href="#"></a>
					<ul style="display: none;">
						<li class="flat" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="flat" href="#">flat</a></li>
						<li class="glass" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);' data-textureheight="400" data-texturewidth="1"><a title="glass" href="#">glass</a></li>
						<li class="highlight_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_soft" href="#">highlight_soft</a></li>
						<li class="highlight_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_hard" href="#">highlight_hard</a></li>
						<li class="inset_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_soft" href="#">inset_soft</a></li>
						<li class="inset_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_hard" href="#">inset_hard</a></li>
						<li class="diagonals_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-small_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_small" href="#">diagonals_small</a></li>
						<li class="diagonals_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-medium_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_medium" href="#">diagonals_medium</a></li>
						<li class="diagonals_thick" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-thick_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_thick" href="#">diagonals_thick</a></li>
						<li class="dots_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-small_100_555_2x2.png") 50% 50% rgb(85, 85, 85);' data-textureheight="2" data-texturewidth="2"><a title="dots_small" href="#">dots_small</a></li>
						<li class="dots_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-medium_100_555_4x4.png") 50% 50% rgb(85, 85, 85);' data-textureheight="4" data-texturewidth="4"><a title="dots_medium" href="#">dots_medium</a></li>
						<li class="white_lines" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_white-lines_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="white_lines" href="#">white_lines</a></li>
						<li class="gloss_wave" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_gloss-wave_100_555_500x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="500"><a title="gloss_wave" href="#">gloss_wave</a></li>
						<li class="diamond" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond_100_555_10x8.png") 50% 50% rgb(85, 85, 85);' data-textureheight="8" data-texturewidth="10"><a title="diamond" href="#">diamond</a></li>
						<li class="loop" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_loop_100_555_21x21.png") 50% 50% rgb(85, 85, 85);' data-textureheight="21" data-texturewidth="21"><a title="loop" href="#">loop</a></li>
						<li class="carbon_fiber" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_carbon-fiber_100_555_8x9.png") 50% 50% rgb(85, 85, 85);' data-textureheight="9" data-texturewidth="8"><a title="carbon_fiber" href="#">carbon_fiber</a></li>
						<li class="diagonal_maze" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonal-maze_100_555_10x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="10"><a title="diagonal_maze" href="#">diagonal_maze</a></li>
						<li class="diamond_ripple" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond-ripple_100_555_22x22.png") 50% 50% rgb(85, 85, 85);' data-textureheight="22" data-texturewidth="22"><a title="diamond_ripple" href="#">diamond_ripple</a></li>
						<li class="hexagon" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_hexagon_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="hexagon" href="#">hexagon</a></li>
						<li class="layered_circles" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_layered-circles_100_555_13x13.png") 50% 50% rgb(85, 85, 85);' data-textureheight="13" data-texturewidth="13"><a title="layered_circles" href="#">layered_circles</a></li>
						<li class="3D_boxes" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_3D-boxes_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="3D_boxes" href="#">3D_boxes</a></li>
						<li class="glow_ball" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glow-ball_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="glow_ball" href="#">glow_ball</a></li>
						<li class="spotlight" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_spotlight_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="spotlight" href="#">spotlight</a></li>
						<li class="fine_grain" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_fine-grain_100_555_60x60.png") 50% 50% rgb(85, 85, 85);' data-textureheight="60" data-texturewidth="60"><a title="fine_grain" href="#">fine_grain</a></li>
					</ul>
				</div>
				<input name="bgImgOpacityHighlight" class="opacity" type="text" value="55">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group field-group-border clearfix">
				<label for="borderColorHighlight">Border</label>
				<div class="hasPicker">
					<input name="borderColorHighlight" class="hex" id="borderColorHighlight" style="color: rgb(0, 0, 0); background-color: rgb(252, 239, 161);" type="text" size="6" value="#fcefa1">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="fcHighlight">Text</label>
				<div class="hasPicker">
					<input name="fcHighlight" class="hex" id="fcHighlight" style="color: rgb(255, 255, 255); background-color: rgb(54, 54, 54);" type="text" size="6" value="#363636">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="iconColorHighlight">Icon</label>
				<div class="hasPicker">
					<input name="iconColorHighlight" class="hex" id="iconColorHighlight" style="color: rgb(0, 0, 0); background-color: rgb(46, 131, 255);" type="text" size="6" value="#2e83ff">
				</div>
			</div>
		</div>
		<!-- /theme group content -->

		<h3>Error</h3>
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-background clearfix">
				<label class="background" for="bgColorError">Background color &amp; texture</label>
				<div class="hasPicker">
					<input name="bgColorError" class="hex" id="bgColorError" style="color: rgb(0, 0, 0); background-color: rgb(254, 241, 236);" type="text" value="#fef1ec">
				</div>
				<select name="bgTextureError" class="texture">

					<option value="flat" data-textureheight="100" data-texturewidth="40">flat</option>

					<option selected='"selected"' value="glass" data-textureheight="400" data-texturewidth="1">glass</option>

					<option value="highlight_soft" data-textureheight="100" data-texturewidth="1">highlight_soft</option>

					<option value="highlight_hard" data-textureheight="100" data-texturewidth="1">highlight_hard</option>

					<option value="inset_soft" data-textureheight="100" data-texturewidth="1">inset_soft</option>

					<option value="inset_hard" data-textureheight="100" data-texturewidth="1">inset_hard</option>

					<option value="diagonals_small" data-textureheight="40" data-texturewidth="40">diagonals_small</option>

					<option value="diagonals_medium" data-textureheight="40" data-texturewidth="40">diagonals_medium</option>

					<option value="diagonals_thick" data-textureheight="40" data-texturewidth="40">diagonals_thick</option>

					<option value="dots_small" data-textureheight="2" data-texturewidth="2">dots_small</option>

					<option value="dots_medium" data-textureheight="4" data-texturewidth="4">dots_medium</option>

					<option value="white_lines" data-textureheight="100" data-texturewidth="40">white_lines</option>

					<option value="gloss_wave" data-textureheight="100" data-texturewidth="500">gloss_wave</option>

					<option value="diamond" data-textureheight="8" data-texturewidth="10">diamond</option>

					<option value="loop" data-textureheight="21" data-texturewidth="21">loop</option>

					<option value="carbon_fiber" data-textureheight="9" data-texturewidth="8">carbon_fiber</option>

					<option value="diagonal_maze" data-textureheight="10" data-texturewidth="10">diagonal_maze</option>

					<option value="diamond_ripple" data-textureheight="22" data-texturewidth="22">diamond_ripple</option>

					<option value="hexagon" data-textureheight="10" data-texturewidth="12">hexagon</option>

					<option value="layered_circles" data-textureheight="13" data-texturewidth="13">layered_circles</option>

					<option value="3D_boxes" data-textureheight="10" data-texturewidth="12">3D_boxes</option>

					<option value="glow_ball" data-textureheight="16" data-texturewidth="16">glow_ball</option>

					<option value="spotlight" data-textureheight="16" data-texturewidth="16">spotlight</option>

					<option value="fine_grain" data-textureheight="60" data-texturewidth="60">fine_grain</option>

				</select><div title="glass" class="texturePicker" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);'>
					<a href="#"></a>
					<ul style="display: none;">
						<li class="flat" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="flat" href="#">flat</a></li>
						<li class="glass" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glass_100_555_1x400.png") 50% 50% rgb(85, 85, 85);' data-textureheight="400" data-texturewidth="1"><a title="glass" href="#">glass</a></li>
						<li class="highlight_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_soft" href="#">highlight_soft</a></li>
						<li class="highlight_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_hard" href="#">highlight_hard</a></li>
						<li class="inset_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_soft" href="#">inset_soft</a></li>
						<li class="inset_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_hard" href="#">inset_hard</a></li>
						<li class="diagonals_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-small_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_small" href="#">diagonals_small</a></li>
						<li class="diagonals_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-medium_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_medium" href="#">diagonals_medium</a></li>
						<li class="diagonals_thick" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-thick_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_thick" href="#">diagonals_thick</a></li>
						<li class="dots_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-small_100_555_2x2.png") 50% 50% rgb(85, 85, 85);' data-textureheight="2" data-texturewidth="2"><a title="dots_small" href="#">dots_small</a></li>
						<li class="dots_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-medium_100_555_4x4.png") 50% 50% rgb(85, 85, 85);' data-textureheight="4" data-texturewidth="4"><a title="dots_medium" href="#">dots_medium</a></li>
						<li class="white_lines" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_white-lines_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="white_lines" href="#">white_lines</a></li>
						<li class="gloss_wave" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_gloss-wave_100_555_500x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="500"><a title="gloss_wave" href="#">gloss_wave</a></li>
						<li class="diamond" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond_100_555_10x8.png") 50% 50% rgb(85, 85, 85);' data-textureheight="8" data-texturewidth="10"><a title="diamond" href="#">diamond</a></li>
						<li class="loop" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_loop_100_555_21x21.png") 50% 50% rgb(85, 85, 85);' data-textureheight="21" data-texturewidth="21"><a title="loop" href="#">loop</a></li>
						<li class="carbon_fiber" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_carbon-fiber_100_555_8x9.png") 50% 50% rgb(85, 85, 85);' data-textureheight="9" data-texturewidth="8"><a title="carbon_fiber" href="#">carbon_fiber</a></li>
						<li class="diagonal_maze" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonal-maze_100_555_10x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="10"><a title="diagonal_maze" href="#">diagonal_maze</a></li>
						<li class="diamond_ripple" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond-ripple_100_555_22x22.png") 50% 50% rgb(85, 85, 85);' data-textureheight="22" data-texturewidth="22"><a title="diamond_ripple" href="#">diamond_ripple</a></li>
						<li class="hexagon" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_hexagon_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="hexagon" href="#">hexagon</a></li>
						<li class="layered_circles" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_layered-circles_100_555_13x13.png") 50% 50% rgb(85, 85, 85);' data-textureheight="13" data-texturewidth="13"><a title="layered_circles" href="#">layered_circles</a></li>
						<li class="3D_boxes" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_3D-boxes_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="3D_boxes" href="#">3D_boxes</a></li>
						<li class="glow_ball" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glow-ball_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="glow_ball" href="#">glow_ball</a></li>
						<li class="spotlight" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_spotlight_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="spotlight" href="#">spotlight</a></li>
						<li class="fine_grain" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_fine-grain_100_555_60x60.png") 50% 50% rgb(85, 85, 85);' data-textureheight="60" data-texturewidth="60"><a title="fine_grain" href="#">fine_grain</a></li>
					</ul>
				</div>
				<input name="bgImgOpacityError" class="opacity" type="text" value="95">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group field-group-border clearfix">
				<label for="borderColorError">Border</label>
				<div class="hasPicker">
					<input name="borderColorError" class="hex" id="borderColorError" style="color: rgb(255, 255, 255); background-color: rgb(205, 10, 10);" type="text" size="6" value="#cd0a0a">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="fcError">Text</label>
				<div class="hasPicker">
					<input name="fcError" class="hex" id="fcError" style="color: rgb(255, 255, 255); background-color: rgb(205, 10, 10);" type="text" size="6" value="#cd0a0a">
				</div>
			</div>
			<div class="field-group clearfix">
				<label for="iconColorError">Icon</label>
				<div class="hasPicker">
					<input name="iconColorError" class="hex" id="iconColorError" style="color: rgb(255, 255, 255); background-color: rgb(205, 10, 10);" type="text" size="6" value="#cd0a0a">
				</div>
			</div>
		</div>
		<!-- /theme group content -->


		<h3>Modal Screen for Overlays </h3>
		<!-- /theme group Overlay -->
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-background clearfix">
				<label class="background" for="bgColorOverlay">Background color &amp; texture</label>
				<div class="hasPicker">
					<input name="bgColorOverlay" class="hex" id="bgColorOverlay" style="color: rgb(0, 0, 0); background-color: rgb(170, 170, 170);" type="text" value="#aaaaaa">
				</div>
				<select name="bgTextureOverlay" class="texture">

					<option selected='"selected"' value="flat" data-textureheight="100" data-texturewidth="40">flat</option>

					<option value="highlight_soft" data-textureheight="100" data-texturewidth="1">highlight_soft</option>

					<option value="highlight_hard" data-textureheight="100" data-texturewidth="1">highlight_hard</option>

					<option value="inset_soft" data-textureheight="100" data-texturewidth="1">inset_soft</option>

					<option value="inset_hard" data-textureheight="100" data-texturewidth="1">inset_hard</option>

					<option value="diagonals_small" data-textureheight="40" data-texturewidth="40">diagonals_small</option>

					<option value="diagonals_medium" data-textureheight="40" data-texturewidth="40">diagonals_medium</option>

					<option value="diagonals_thick" data-textureheight="40" data-texturewidth="40">diagonals_thick</option>

					<option value="dots_small" data-textureheight="2" data-texturewidth="2">dots_small</option>

					<option value="dots_medium" data-textureheight="4" data-texturewidth="4">dots_medium</option>

					<option value="white_lines" data-textureheight="100" data-texturewidth="40">white_lines</option>

					<option value="gloss_wave" data-textureheight="100" data-texturewidth="500">gloss_wave</option>

					<option value="diamond" data-textureheight="8" data-texturewidth="10">diamond</option>

					<option value="loop" data-textureheight="21" data-texturewidth="21">loop</option>

					<option value="carbon_fiber" data-textureheight="9" data-texturewidth="8">carbon_fiber</option>

					<option value="diagonal_maze" data-textureheight="10" data-texturewidth="10">diagonal_maze</option>

					<option value="diamond_ripple" data-textureheight="22" data-texturewidth="22">diamond_ripple</option>

					<option value="hexagon" data-textureheight="10" data-texturewidth="12">hexagon</option>

					<option value="layered_circles" data-textureheight="13" data-texturewidth="13">layered_circles</option>

					<option value="3D_boxes" data-textureheight="10" data-texturewidth="12">3D_boxes</option>

					<option value="glow_ball" data-textureheight="16" data-texturewidth="16">glow_ball</option>

					<option value="spotlight" data-textureheight="16" data-texturewidth="16">spotlight</option>

					<option value="fine_grain" data-textureheight="60" data-texturewidth="60">fine_grain</option>

				</select><div title="flat" class="texturePicker" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);'>
					<a href="#"></a>
					<ul style="display: none;">
						<li class="flat" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="flat" href="#">flat</a></li>
						<li class="highlight_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_soft" href="#">highlight_soft</a></li>
						<li class="highlight_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_hard" href="#">highlight_hard</a></li>
						<li class="inset_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_soft" href="#">inset_soft</a></li>
						<li class="inset_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_hard" href="#">inset_hard</a></li>
						<li class="diagonals_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-small_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_small" href="#">diagonals_small</a></li>
						<li class="diagonals_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-medium_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_medium" href="#">diagonals_medium</a></li>
						<li class="diagonals_thick" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-thick_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_thick" href="#">diagonals_thick</a></li>
						<li class="dots_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-small_100_555_2x2.png") 50% 50% rgb(85, 85, 85);' data-textureheight="2" data-texturewidth="2"><a title="dots_small" href="#">dots_small</a></li>
						<li class="dots_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-medium_100_555_4x4.png") 50% 50% rgb(85, 85, 85);' data-textureheight="4" data-texturewidth="4"><a title="dots_medium" href="#">dots_medium</a></li>
						<li class="white_lines" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_white-lines_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="white_lines" href="#">white_lines</a></li>
						<li class="gloss_wave" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_gloss-wave_100_555_500x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="500"><a title="gloss_wave" href="#">gloss_wave</a></li>
						<li class="diamond" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond_100_555_10x8.png") 50% 50% rgb(85, 85, 85);' data-textureheight="8" data-texturewidth="10"><a title="diamond" href="#">diamond</a></li>
						<li class="loop" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_loop_100_555_21x21.png") 50% 50% rgb(85, 85, 85);' data-textureheight="21" data-texturewidth="21"><a title="loop" href="#">loop</a></li>
						<li class="carbon_fiber" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_carbon-fiber_100_555_8x9.png") 50% 50% rgb(85, 85, 85);' data-textureheight="9" data-texturewidth="8"><a title="carbon_fiber" href="#">carbon_fiber</a></li>
						<li class="diagonal_maze" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonal-maze_100_555_10x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="10"><a title="diagonal_maze" href="#">diagonal_maze</a></li>
						<li class="diamond_ripple" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond-ripple_100_555_22x22.png") 50% 50% rgb(85, 85, 85);' data-textureheight="22" data-texturewidth="22"><a title="diamond_ripple" href="#">diamond_ripple</a></li>
						<li class="hexagon" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_hexagon_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="hexagon" href="#">hexagon</a></li>
						<li class="layered_circles" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_layered-circles_100_555_13x13.png") 50% 50% rgb(85, 85, 85);' data-textureheight="13" data-texturewidth="13"><a title="layered_circles" href="#">layered_circles</a></li>
						<li class="3D_boxes" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_3D-boxes_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="3D_boxes" href="#">3D_boxes</a></li>
						<li class="glow_ball" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glow-ball_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="glow_ball" href="#">glow_ball</a></li>
						<li class="spotlight" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_spotlight_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="spotlight" href="#">spotlight</a></li>
						<li class="fine_grain" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_fine-grain_100_555_60x60.png") 50% 50% rgb(85, 85, 85);' data-textureheight="60" data-texturewidth="60"><a title="fine_grain" href="#">fine_grain</a></li>
					</ul>
				</div>
				<input name="bgImgOpacityOverlay" class="opacity" type="text" value="0">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group field-group-opacity clearfix">
				<label for="opacityOverlay">Overlay Opacity:</label>
				<input name="opacityOverlay" class="opacity" id="opacityOverlay" type="text" value="30">
				<span class="opacity-per">%</span>
			</div>
		</div>
		<!-- /theme group Overlay -->

		<h3>Drop Shadows</h3>
		<!-- /theme group Shadow -->
		<div class="theme-group-content corner-bottom clearfix">
			<div class="field-group field-group-background clearfix">
				<label class="background" for="bgColorShadow">Background color &amp; texture</label>
				<div class="hasPicker">
					<input name="bgColorShadow" class="hex" id="bgColorShadow" style="color: rgb(0, 0, 0); background-color: rgb(170, 170, 170);" type="text" value="#aaaaaa">
				</div>
				<select name="bgTextureShadow" class="texture">

					<option selected='"selected"' value="flat" data-textureheight="100" data-texturewidth="40">flat</option>

					<option value="highlight_soft" data-textureheight="100" data-texturewidth="1">highlight_soft</option>

					<option value="highlight_hard" data-textureheight="100" data-texturewidth="1">highlight_hard</option>

					<option value="inset_soft" data-textureheight="100" data-texturewidth="1">inset_soft</option>

					<option value="inset_hard" data-textureheight="100" data-texturewidth="1">inset_hard</option>

					<option value="diagonals_small" data-textureheight="40" data-texturewidth="40">diagonals_small</option>

					<option value="diagonals_medium" data-textureheight="40" data-texturewidth="40">diagonals_medium</option>

					<option value="diagonals_thick" data-textureheight="40" data-texturewidth="40">diagonals_thick</option>

					<option value="dots_small" data-textureheight="2" data-texturewidth="2">dots_small</option>

					<option value="dots_medium" data-textureheight="4" data-texturewidth="4">dots_medium</option>

					<option value="white_lines" data-textureheight="100" data-texturewidth="40">white_lines</option>

					<option value="gloss_wave" data-textureheight="100" data-texturewidth="500">gloss_wave</option>

					<option value="diamond" data-textureheight="8" data-texturewidth="10">diamond</option>

					<option value="loop" data-textureheight="21" data-texturewidth="21">loop</option>

					<option value="carbon_fiber" data-textureheight="9" data-texturewidth="8">carbon_fiber</option>

					<option value="diagonal_maze" data-textureheight="10" data-texturewidth="10">diagonal_maze</option>

					<option value="diamond_ripple" data-textureheight="22" data-texturewidth="22">diamond_ripple</option>

					<option value="hexagon" data-textureheight="10" data-texturewidth="12">hexagon</option>

					<option value="layered_circles" data-textureheight="13" data-texturewidth="13">layered_circles</option>

					<option value="3D_boxes" data-textureheight="10" data-texturewidth="12">3D_boxes</option>

					<option value="glow_ball" data-textureheight="16" data-texturewidth="16">glow_ball</option>

					<option value="spotlight" data-textureheight="16" data-texturewidth="16">spotlight</option>

					<option value="fine_grain" data-textureheight="60" data-texturewidth="60">fine_grain</option>

				</select><div title="flat" class="texturePicker" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);'>
					<a href="#"></a>
					<ul style="display: none;">
						<li class="flat" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_flat_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="flat" href="#">flat</a></li>
						<li class="highlight_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_soft" href="#">highlight_soft</a></li>
						<li class="highlight_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_highlight-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="highlight_hard" href="#">highlight_hard</a></li>
						<li class="inset_soft" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-soft_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_soft" href="#">inset_soft</a></li>
						<li class="inset_hard" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_inset-hard_100_555_1x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="1"><a title="inset_hard" href="#">inset_hard</a></li>
						<li class="diagonals_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-small_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_small" href="#">diagonals_small</a></li>
						<li class="diagonals_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-medium_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_medium" href="#">diagonals_medium</a></li>
						<li class="diagonals_thick" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonals-thick_100_555_40x40.png") 50% 50% rgb(85, 85, 85);' data-textureheight="40" data-texturewidth="40"><a title="diagonals_thick" href="#">diagonals_thick</a></li>
						<li class="dots_small" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-small_100_555_2x2.png") 50% 50% rgb(85, 85, 85);' data-textureheight="2" data-texturewidth="2"><a title="dots_small" href="#">dots_small</a></li>
						<li class="dots_medium" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_dots-medium_100_555_4x4.png") 50% 50% rgb(85, 85, 85);' data-textureheight="4" data-texturewidth="4"><a title="dots_medium" href="#">dots_medium</a></li>
						<li class="white_lines" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_white-lines_100_555_40x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="40"><a title="white_lines" href="#">white_lines</a></li>
						<li class="gloss_wave" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_gloss-wave_100_555_500x100.png") 50% 50% rgb(85, 85, 85);' data-textureheight="100" data-texturewidth="500"><a title="gloss_wave" href="#">gloss_wave</a></li>
						<li class="diamond" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond_100_555_10x8.png") 50% 50% rgb(85, 85, 85);' data-textureheight="8" data-texturewidth="10"><a title="diamond" href="#">diamond</a></li>
						<li class="loop" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_loop_100_555_21x21.png") 50% 50% rgb(85, 85, 85);' data-textureheight="21" data-texturewidth="21"><a title="loop" href="#">loop</a></li>
						<li class="carbon_fiber" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_carbon-fiber_100_555_8x9.png") 50% 50% rgb(85, 85, 85);' data-textureheight="9" data-texturewidth="8"><a title="carbon_fiber" href="#">carbon_fiber</a></li>
						<li class="diagonal_maze" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diagonal-maze_100_555_10x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="10"><a title="diagonal_maze" href="#">diagonal_maze</a></li>
						<li class="diamond_ripple" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_diamond-ripple_100_555_22x22.png") 50% 50% rgb(85, 85, 85);' data-textureheight="22" data-texturewidth="22"><a title="diamond_ripple" href="#">diamond_ripple</a></li>
						<li class="hexagon" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_hexagon_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="hexagon" href="#">hexagon</a></li>
						<li class="layered_circles" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_layered-circles_100_555_13x13.png") 50% 50% rgb(85, 85, 85);' data-textureheight="13" data-texturewidth="13"><a title="layered_circles" href="#">layered_circles</a></li>
						<li class="3D_boxes" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_3D-boxes_100_555_12x10.png") 50% 50% rgb(85, 85, 85);' data-textureheight="10" data-texturewidth="12"><a title="3D_boxes" href="#">3D_boxes</a></li>
						<li class="glow_ball" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_glow-ball_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="glow_ball" href="#">glow_ball</a></li>
						<li class="spotlight" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_spotlight_100_555_16x16.png") 50% 50% rgb(85, 85, 85);' data-textureheight="16" data-texturewidth="16"><a title="spotlight" href="#">spotlight</a></li>
						<li class="fine_grain" style='background: url("http://download.jqueryui.com/themeroller/images/ui-bg_fine-grain_100_555_60x60.png") 50% 50% rgb(85, 85, 85);' data-textureheight="60" data-texturewidth="60"><a title="fine_grain" href="#">fine_grain</a></li>
					</ul>
				</div>
				<input name="bgImgOpacityShadow" class="opacity" id="bgImgOpacityShadow" type="text" value="0">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group field-group-opacity clearfix">
				<label for="opacityShadow">Shadow Opacity:</label>
				<input name="opacityShadow" class="opacity" id="opacityShadow" type="text" value="30">
				<span class="opacity-per">%</span>
			</div>
			<div class="field-group clearfix">
				<label for="thicknessShadow">Shadow Thickness:</label>
				<input name="thicknessShadow" class="offset" id="thicknessShadow" type="text" value="8px">
			</div>
			<div class="field-group clearfix">
				<label for="offsetTopShadow">Top Offset:</label>
				<input name="offsetTopShadow" class="offset" id="offsetTopShadow" type="text" value="-8px">
			</div>
			<div class="field-group clearfix">
				<label for="offsetLeftShadow">Left Offset:</label>
				<input name="offsetLeftShadow" class="offset" id="offsetLeftShadow" type="text" value="-8px">
			</div>
			<div class="field-group field-group-corners clearfix">
				<label for="cornerRadiusShadow">Corners:</label>
				<input name="cornerRadiusShadow" class="cornerRadius" id="cornerRadiusShadow" type="text" value="8px">
			</div>
		</div>
		<!-- /theme group Shadow -->

		
	</div>

<button name="submit" id="submitBtn" type="submit">Preview Changes</button>

	<!-- /themeroller -->
</div>