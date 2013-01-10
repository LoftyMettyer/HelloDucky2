<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage(Of DMI.NET.NavLinksViewModel)" %>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

	
	<link href="<%= url.content("~/Content/jquery.gridster.css") %>" rel="stylesheet" type="text/css" />
	<script src="<%: Url.Content("~/Scripts/jquery.gridster.js") %>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/Scripts/jquery.mousewheel.js") %>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/Scripts/jquery.flip.js") %>" type="text/javascript"></script>



	<script type="text/javascript">
		function loadjscssfile(filename, filetype) {
			var fileref;
			
			if (filetype == "ajax") {
				fileref = document.createElement("script");
				fileref.setAttribute("type", "text/javascript");
				fileref.innerHTML = filename;
			}
			else if (filetype == "js") { //if filename is a external JavaScript file
				fileref = document.createElement('script');
				fileref.id = "myScript1";
				fileref.setAttribute("type", "text/javascript");
				fileref.setAttribute("src", filename);
			}
			else if (filetype == "css") { //if filename is an external CSS file
				fileref = document.createElement("link");
				fileref.setAttribute("rel", "stylesheet");
				fileref.setAttribute("type", "text/css");
				fileref.setAttribute("href", filename);
			}
			if (typeof fileref != "undefined") {
				document.getElementsByTagName("head")[0].appendChild(fileref);
			}
		}

		$(document).ready(function () {
			if (window.currentLayout == "tiles") {
				setupTiles();
			}
			else {
				// for wireframe layout, convert the dropdownlinks to a <select> element
				$(function () {
					$('ul.DropDownListMenu').each(function () {
						var $select = $('<select />');

						$(this).find('a').each(function () {
							var $option = $('<option />');
							$option.attr('value', $(this).attr('href')).html($(this).html());
							$select.append($option);
						});

						$(this).replaceWith($select);
					});
				});
			}

			$(".DashContent").fadeIn("slow");


		});

		function setupTiles() { 
				//apply the gridster functionality.
				griditup(true);

				//add mousewheel scrollability to the main content window
				$(".DashContent").mousewheel(function (event, delta) {
					this.scrollLeft -= (delta * 30);
					event.preventDefault();
				});

				//Add flippy stuff
				$(".flipTile").hover(function () {
					$(this).flip({
						direction: 'tb'
					});
				});		
			
				//hide the ribbon - they're charms in this layout
			//$("#fixedlinks").hide();

		}

		function griditup(mode) {
			if (mode == true) {
				$(".gridster ul").gridster({
					widget_margins: [10, 10],
					widget_base_dimensions: [120, 120],
					min_rows: 4,
					min_cols: 1,
					avoid_overlapped_widgets: true
				});
			}
		}

		function changeLayout(newLayoutName) {

			setCookie('Intranet_Layout', newLayoutName, 365);
			if (newLayoutName == "winkit") {
				setCookie('Intranet_Theme', "white", 365);
			} else {
				setCookie('Intranet_Theme', "blue", 365);
			}


			$(".DashContent").fadeOut("slow");

			$(".DashContent").promise().done(function () {

				//Are we currently in tiles mode? If so, just refresh the screen as there's too much loaded to reformat on the fly.
				var currentLayout = $("link[id=layoutLink]").attr("href");
				if (currentLayout.indexOf("tiles.css") > 0) {
					window.location = "LinksMain";
					return;
				}

				switch (newLayoutName) {
					case "tiles":
						//Hide all officebar tabs, except 'current' which WILL always be home tab...
						//$(".officebar > ul > li:not(.current)").hide();
						//$(".current > a").hide();

						window.changeTheme("theme-Blue");
						$("link[id=layoutLink]").attr({ href: "../Content/DashboardStyles/layouts/tiles.css" });
						setupTiles();
						break;
					case "wireframe":
						//$(".officebar > ul > li:not(.current)").hide();
						//$(".current > a").hide();

						window.changeTheme("theme-Blue");
						$("link[id=layoutLink]").attr({ href: "../Content/DashboardStyles/layouts/wireframe.css" });

						break;
					case "winkit":
						window.changeTheme("theme-White");
						$("link[id=layoutLink]").attr({ href: "../Content/DashboardStyles/layouts/winkit.css" });
						$(".officebar > ul > li:not(.current)").show();
						$(".current > a").show();

						//$(".hypertextlinks ul").addClass("menu");
					default:
						break;
				}

				$(".DashContent").fadeIn("slow");
			});
		}

	</script>


	<div class="DashContent" style="display: none;">
		<div class="tileContent">
		<%Dim fFirstSeparator = True%>
		<%Const iMaxRows As Integer = 4%>
		<%Dim iRowNum = 1%>
		<%Dim iColNum = 1%>
		<%Dim iSeparatorNum = 0%>
			<div class="hypertextlinks">
				<%For Each navlink In Model.NavigationLinks%>
				<%If navlink.LinkType = 0 Then	 ' hypertext link%>
				<%If navlink.Element_Type = 1 Then		' separator%>
				<%iRowNum = 1%>
				<%iColNum = 1%>
				<%If fFirstSeparator Then%>
				<%fFirstSeparator = False%>
				<%Else%>
				</ul>
			</div>
			</li> </ul>
			<%End If%>
			<%iSeparatorNum += 1%>
			<ul class="hypertextlinkseparatorframe" id="hypertextlinkseparatorframe_<%=iSeparatorNum %>">
				<li class="hypertextlink-displaytype"><a class="hypertextlinkseparator" href="#">
					<%: navlink.Text %></a>
					<div class="gridster">
						<ul>
							<%Else%>
							<%If iRowNum > iMaxRows Then%>
							<% iColNum += 1%>
							<%iRowNum = 1%>
							<script type="text/javascript">
								$("#hypertextlinkseparatorframe<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
								$("#hypertextlinkseparatorframe<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
							</script>
							<%End If%>
							<%
								Dim classIcon As String = ""
								Select Case navlink.Element_Type%>
							<%Case 0
									classIcon = "icon-external-link"

							 End Select%>

							<li class="hypertextlinktext greenTile flipTile" data-col="<%=iColNum %>" data-row="<%=iRowNum %>"
								data-sizex="1" data-sizey="1">
								<p class="hypertextlinktileIcon">
									<i class="<%=classIcon %>"></i>
								</p>
								<p>
									<a href="#">
										<%: navlink.Text %></a></p>
							</li>
							<%iRowNum += 1%>
							<%End If%>
							<%End If%>
							<%Next%>
							<%If Not fFirstSeparator Then		' close off the hypertext group%>
						</ul>
					</div>
				</li>
			</ul>
			<%End If%>
		</div>


		<%fFirstSeparator = True%>
		<div class="linkspagebutton">
			<div class="ButtonLinkColumn">
				<%For Each navlink In Model.NavigationLinks%>
				<%If navlink.LinkType = 1 Then	 ' main dashboard link%>
				<%If navlink.Element_Type = 1 Then		' separator%>
				<%iRowNum = 1%>
				<%iColNum = 1%>
				<%If fFirstSeparator Then%>
				<%fFirstSeparator = False%>
				<%Else%>
				</ul>
			</div>
			</li> </ul>
			<%End If%>
			<%If navlink.SeparatorOrientation = 1 Then	' Vertical break/new column %>
		</div>
		<div class="ButtonLinkColumn">
			<%End If%>
			<%iSeparatorNum += 1%>
			<ul class="linkspagebuttonseparatorframe" id="linkspagebuttonseparatorframe_<%=iSeparatorNum %>">
				<li class="linkspagebutton-displaytype">
					
					<a class="linkspagebuttonseparator" href="#"><%: navlink.Text %></a>

					<div class="gridster">
						<ul>
							<%Else%>
							<%If iRowNum > iMaxRows Then	 ' start a new column if required (affects tiles only)%>
							<% iColNum += 1%>
							<%iRowNum = 1%>
							<script type="text/javascript">
								$("#linkspagebuttonseparatorframe_<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
								$("#linkspagebuttonseparatorframe_<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
							</script>
							<%End If%>

							<%Select Case navlink.Element_Type%>

							<%Case 0		 ' Button Link	%>
								<%If navlink.UtilityType = -1 Then	' screen view%>
									<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"	class="linkspagebuttontext blueTile">
										<p class="linkspagebuttontileIcon"><i class="icon-table"></i></p>
										<p><a href="#"><%: navlink.Text %><img src="<%: Url.Content("~/Content/images/extlink2.png") %>" alt=""/></a></p>
									</li>								
								<%ElseIf navlink.UtilityType = 25 Then	' workflow launch%>
									<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"	class="linkspagebuttontext lightBlueTile">
										<p class="linkspagebuttontileIcon"><i class="icon-magic"></i></p>
										<p><a href="#"><%: navlink.Text %><img src="<%: Url.Content("~/Content/images/extlink2.png") %>" alt=""/></a></p>
									</li>								

								<%ElseIf navlink.UtilityType = 2 Then	 ' report/utility%>
									<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"	class="linkspagebuttontext orangeTile">
										<p class="linkspagebuttontileIcon"><i class="icon-file"></i></p>
										<p><a href="#"><%: navlink.Text %><img src="<%: Url.Content("~/Content/images/extlink2.png") %>" alt=""/></a></p>
									</li>								

								<%End If%>


								<%iRowNum += 1%>


							<%Case 2		' Chart %>
								<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"	class="linkspagebuttontext blueTile displayonly">
									<p class="linkspagebuttontileIcon">
										<i class="icon-bar-chart"></i>
									</p>
									<p>
										<a href="#"><%: navlink.Text %></a>
									</p>
									<div class="widgetplaceholder generaltheme">
										<div><i class="icon-bar-chart"></i></div>
										<a href="#">Chart</a>
									</div>
									
								</li>
								<%iRowNum += 1%>

							<%Case 3		 ' Pending Workflows	%>
								<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"	class="linkspagebuttontext blueTile displayonly">
									<p class="linkspagebuttontileIcon">
										<i class="icon-inbox"></i>
									</p>
									<p>
										<a href="#">Pending Workflows</a>
									</p>
									<div class="widgetplaceholder generaltheme">
										<div><i class="icon-inbox"></i></div>
										<a href="#">Pending Workflows</a>
									</div>
								</li>
								<%iRowNum += 1%>



							<%Case 4		' Database Value%>
								<li id="li_<%: navlink.id %>" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
									data-sizey="1" class="linkspagebuttontext redTile displayonly">
									<div class="DBValueScroller" id="marqueeDBV<%: navlink.id %>">
										<p class="DBValue" id="DBV<%: navlink.id %>">
											<img class="DBVSpinner" id="SpinnerDBV<%: navlink.id %>" src="<%: url.content("~/Content/images/spinner04.gif") %>"
												alt="..." />
										</p>
									</div>
									<a href="#">
										<p class="DBValueCaption">
											<%: navlink.Text %></p>
									</a>
								</li>
<%--								<script type="text/javascript">loadjscssfile('$.getScript("http://abs16090/dmi.net/scripts/widgetscripts/wdg_oHRDBV.js", function () { initialiseWidget(<%: navlink.id %>, "DBV<%: navlink.id %>", "DBV<%: navlink.Text %>", ""); });', 'ajax');</script>--%>
								<%iRowNum += 1%>

							<%Case 5		 ' Todays events	%>
								<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"	class="linkspagebuttontext blueTile displayonly">
									<p class="linkspagebuttontileIcon">
										<i class="icon-calendar"></i>
									</p>
									
									<div class="holidaycontainer" id="HolContainer<%: navlink.id %>"></div>
									
								</li>
								<script type="text/javascript">loadjscssfile('$.getScript("http://abs16090/dmi.net/scripts/widgetscripts/wdg_oHRHoliday.js", function () { initialiseWidget(<%: navlink.id %>, "HolContainer<%: navlink.id %>", 19, ""); });', 'ajax');</script>
								<%iRowNum += 1%>


							<%Case Else%>
								<li data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1" data-sizey="1"
									class="linkspagebuttontext blueTile displayonly"><a href="#">
										<%: navlink.Text %></a></li>
								<%iRowNum += 1%>

							<%End Select%>




							<%End If%>
							<%End If%>
							<%Next%>
								<%If Not fFirstSeparator Then%>
							</ul>
				</div>
						</li>
					</ul>
				<%End If%>
			</div>
		</div>

		<%If Model.NumberOfLinks > 0 Then%>
	<div class="dropdownlinks">
		<ul class="dropdownlinkseparatorframe" id="dropdownlinkseparatorframe_<%=iSeparatorNum %>">
		<li class="dropdownlink-displaytype">
			<a class="dropdownlinkseparator" href="#">Dropdown links:</a>
			<div class="gridster">
			<ul class="DropDownListMenu">
				<%iRowNum = 1%>
				<%iColNum = 1%>
				<%For Each navlink In Model.NavigationLinks%>
				<%If navlink.LinkType = 2 Then	 ' dropdown link%>
				<%If iRowNum > iMaxRows Then	 ' start a new column if required (affects tiles only)%>
				<% iColNum += 1%>
				<%iRowNum = 1%>
				<script type="text/javascript">
					$("#dropdownlinksseparatorframe<%=iSeparatorNum %>").removeClass("cols<%=iColNum-1 %>");
					$("#dropdownlinksseparatorframe<%=iSeparatorNum %>").addClass("cols<%=iColNum %>");
				</script>
				<%End If%>
				<li class="dropdownlinktext greenTile" data-col="<%=iColNum %>" data-row="<%=iRowNum %>" data-sizex="1"
					data-sizey="1">
					<p class="dropdownlinktileIcon">
						<i class="icon-external-link"></i>
					</p>
					<p>
						<a href="#">
							<%: navlink.Text %></a>
					</p>
				</li>
				<%iRowNum += 1%>
				<%End If%>
				<%Next%>
			</ul>
			</div>
			</li>

			</ul>
	</div>

		<%End If%>

		</div>
	</div>
</asp:Content>
