<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
	Dim objSessionContext = CType(Session("sessionContext"), SessionInfo)
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
	
	If objDataAccess Is Nothing Then Return
%>
<script type="text/javascript">
	$(document).ready(function () {
		
		////Org Chart Print button is split only for winkit mode.
		//if (window.currentLayout == "winkit") {
		//	//winkit has a split button
		//	$('.mnuBtnPrintOrgChart').parent().addClass("button split");
		//} else {
		//	$('.mnuBtnPrintOrgChart').parent().addClass("button");
		//}



		$(".officebar").officebar({});

		$("#officebar .button").addClass("ui-corner-all");

		if ((window.currentLayout == "winkit") || (window.currentLayout == "wireframe")) {
			$('#officebar .button').addClass('ui-state-default');

			$('#officebar .button').hover(
				function() { if (!$(this).hasClass("disabled")) $(this).addClass('ui-state-hover'); },
				function() { if (!$(this).hasClass("disabled")) $(this).removeClass('ui-state-hover'); }
			);
		}
		else {
		}

		setTimeout("wrapTileIcons();", 100);

		$("#fixedlinks").fadeIn("slow");

		//add a little function to jQuery which allows case insensitive searches..
		$.extend($.expr[":"], {
			"MyCaseInsensitiveContains": function (elem, i, match, array) {
				return (elem.textContent || elem.innerText || "").toLowerCase().indexOf((match[3] || "").toLowerCase()) >= 0;
			}
		});
		
		//Search dashboard functionality
		$("#searchDashboardString").keyup(function (event) {
			filterDashboard();
		});
		$("#searchDashboardString").mouseup(function (event) {
			setTimeout('filterDashboard();', 100);
		});		


		menu_setVisibleMenuItem("userDropdownmenu_Layout", menu_isSSIMode()); //Set visibility of Layout menu
		menu_setVisibleMenuItem("mnutoolFixedSelfService", '<%=objSessionContext.LoginInfo.IsSSIUser%>');
		menu_setVisibleMenuItem("mnutoolFixedOpenHR", '<%: (objSessionContext.LoginInfo.IsDMIUser) And Not Session("isMobileDevice")%>');

		$("#userDropdownmenu a").on("click", function () {
			$("#userDropdownmenu ul").css("visibility", "visible");
			$("#userDropdownmenu ul li").css("visibility", "visible");
			var userMenuHeight = Number($('#userDropdownmenu ul ul').height());
			var topPos = (userMenuHeight + 26) * -1;
			if (window.currentLayout == "tiles") $('#userDropdownmenu ul ul').css('top', topPos).css('height', userMenuHeight);
		});

		if (window.isMobileDevice == "True") {
			//This is a tablet or phone format - So make the Dashboard information smaller
			$('.ViewDescription p').css({
				'font-size': '1em',
				'margin-bottom': '10px'
			});
		}

		$('#userDropdownmenu').hover(function () {
			//On hover in, do nothing
		}, function () {
			//On hover out, hide the menu
			$("#userDropdownmenu ul").css("visibility", "hidden");
			$("#userDropdownmenu li").css("visibility", "hidden");
		});

		$("#userDropdownmenu_Items").menu();


	});

	function filterDashboard() {
		//Dashboard search functionality - tiles only.
		var searchString = $('#searchDashboardString').val();
		if (searchString.length == 0) {
			$('.pendingworkflowlinkcontent li').removeClass('dimmed');
			$('.dropdownlinkcontent li').removeClass('dimmed');
			$('.hypertextlinkcontent li').removeClass('dimmed');
			$('.buttonlinkcontent li').removeClass('dimmed');
		} else {
			$('.pendingworkflowlinkcontent li').addClass('dimmed');
			$('.dropdownlinkcontent li').addClass('dimmed');
			$('.hypertextlinkcontent li').addClass('dimmed');
			$('.buttonlinkcontent li').addClass('dimmed');
			$("span:MyCaseInsensitiveContains('" + searchString + "')").parent('li').removeClass('dimmed');
			$("a:MyCaseInsensitiveContains('" + searchString + "')").parent('li').removeClass('dimmed');
		}
	}

	function wrapTileIcons() {
		if (window.currentLayout == "tiles") {
			//Wrap the icons with circles or boxes or whatever...
			//$(".officetab i[class^='icon-']").css("padding-left", "7px");
			//$(".officetab i[class^='icon-']").css("padding-bottom", "4px");
			$(".officetab i[class^='icon-']").wrap("<span class='icon-stack' />");
			$(".officetab .icon-stack").prepend("<i class='icon-check-empty icon-stack-base'></i>");
			
			//$('#officebar .button').addClass('ui-state-default');
			$('#officebar .button').hover(
			//	function () { if (!$(this).hasClass("disabled")) $(this).addClass('ui-state-hover'); },
			//	function () { if (!$(this).hasClass("disabled")) $(this).removeClass('ui-state-hover'); }
			);
			
		}
	}

	function fixedlinks_mnutoolAboutHRPro() {
		if (OpenHR.currentWorkPage() == "FIND") {
			try {
				if (rowWasModified) {
					//Inform the user that they have unsaved changes on the Find window
					OpenHR.modalMessage("You have unsaved changes.<br/><br/>Please action them before navigating away.");
					return false;
				}
			} catch (e) { //continue with navigation 
			}
		} else {
			OpenHR.showAboutPopup();
		}
	}

	function showThemeEditor() {
		if (OpenHR.currentWorkPage() == "FIND") {
			try {
				if (rowWasModified) {
					//Inform the user that they have unsaved changes on the Find window
					OpenHR.modalMessage("You have unsaved changes.<br/><br/>Please action them before navigating away.");
					return false;
				}
			} catch (e) { //continue with navigation 
			}
		}
		else {
			$("#divthemeRoller").dialog("open");

			//load the themeeditor form now
			loadPartialView("themeEditor", "home", "divthemeRoller", null);
		}
	}

	//why was this here?...
	//$("#officebar").tabs();

</script>
<div id="fixedlinks">
	<div class="dashboardSearch" id="searchBox"><span>Search: <input type="text" id="searchDashboardString"/></span></div>
	<div class="ViewDescription">
		<p></p>
	</div>
	<div class="FixedLinksLeft">
		<div id="officebar" class="officebar">
			<ul>
						<%-- Home --%>
				<li class="current"><a class="ui-state-active ui-corner-top" id="toolbarHome" href="#" rel="home">Home</a>
					<ul>
						<li><span>&nbsp;   </span><%-- Fixed Links value removed By mayank to avoide duplicasy--%>
							<div id="mnutoolFixedSelfService" class="button">
								<a href="#" rel="table" title="Self Service">
<%--									<a href="<%: Url.Action("Main", "Home", New With {.SSIMode = vbTrue})%>" rel="table" title="Self-service">--%>
								<img src="<%: Url.Content("~/Scripts/officebar/winkit/abssmall.png")%>" alt="" />
									<i class="icon-user"></i>
										<h6>Self-service</h6>
									</a>
							</div>
							<div id="mnutoolFixedOpenHR" class="button">
								<a href="#" rel="table" title="OpenHR Web">
								<%--<a href="<%: Url.Action("Main", "Home", New With {.SSIMode = vbFalse})%>" rel="table" title="OpenHR Web">--%>
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/abssmall.png") %>" alt="" />
									<i class="icon-group"></i>
									<h6>OpenHR<br />Web</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record: Find Record--%>
				<li><a class="ui-state-default ui-corner-top" id="toolbarRecordFind" href="#" rel="Find">Find</a>
					<ul>
						<li id="mnuSectionRecordFindEdit"><span>Edit</span>											
							<div id="mnutoolNewRecordFind" class="button">
								<a href="#" rel="table" title="New Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt="" />
									<i class="icon-plus"></i>  
									<h6>New</h6>
								</a>
							</div>
							<div id="mnutoolCopyRecordFind" class="button">
								<a href="#" rel="table" title="Copy Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
									<i class="icon-copy"></i>
									<h6>Copy</h6>
								</a>
							</div>
							<div id="mnutoolEditRecordFind" class="button">
								<a href="#" rel="table" title="Edit Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
									<i class="icon-pencil"></i>
									<h6>Edit</h6>
								</a>
							</div>
							<div id="mnutoolDeleteRecordFind" class="button">
								<a href="#" rel="table" title="Delete Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cross64HOVER.png")%>" alt="" />
									<i class="icon-remove"></i>
									<h6>Delete</h6>
								</a>
							</div>
						</li>						
						<li id="mnuSectionRecordFindNavigate"><span>Navigate</span>
							<div id="mnutoolParentRecordFind" class="button">
								<a href="#" rel="table" title="Return to parent record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/upblack64HOVER.png") %>" alt="" />
									<i class="icon-arrow-up"></i>
									<h6>Parent</h6>
								</a>
							</div>
							<div id="mnutoolBackRecordFind" class="button">
								<a href="#" rel="table" title="Return to record editing">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BackRecord64HOVER.png") %>" alt="" />
									<i class="icon-arrow-left"></i>
									<h6>Back</h6>
								</a>
							</div>

							<div id="mnutoolAccessLinksFind" class="button hidden">
								<a href="#" rel="table" title="Access the links for the selected record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png")%>" alt="" />
									<i class="icon-pencil"></i>
									<h6>Access Links</h6>
								</a>
							</div>

							<div id="mnutoolCancelLinksFind" class="button hidden">
								<a href="#" rel="table" title="Return to links page">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/user64HOVER.png")%>" alt="" />
									<i class="icon-user"></i>
									<h6>Return to the<br/>links page</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionRecordFindOrder"><span>Order</span>
							<div id="mnutoolChangeOrderRecordFind" class="button" title="Change Order">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/sort64HOVER.png") %>" alt="" />
									<i class="icon-ChangeOrderRecordFind"></i>
									<h6>Change Order</h6>
								</a>
							</div>
							<div id="mnutoolFilterRecordFind" class="button" title="Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Filtering64HOVER.png") %>" alt="" />
									<i class="icon-FilterRecordFind"></i>
									<h6>Filter</h6>
								</a>
							</div>
							<div id="mnutoolClearFilterRecordFind" class="button" title="Clear Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/FilteringDelete64HOVER.png") %>"
										alt="" />
									<i class="icon-ClearFilterRecordFind"></i>
									<h6>Clear Filter</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionRecordFindCourseBooking"><span>Course Booking</span>
							<div id="mnutoolBookCourseFind" class="button">
								<a href="#" rel="table" title="Book Course">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BookCourse64HOVER.png") %>"
										alt="" />
									<i class="icon-BookCourseRecord"></i>
									<h6>Book<br />
										Course</h6>
								</a>
							</div>							
						</li>
						<li id="mnuSectionRecordFindTrainingBooking"><span>Training Booking</span>
							<div id="mnutoolBulkBookingRecordFind" class="button">
								<a href="#" rel="table" title="Bulk Booking">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BulkBooking64HOVER.png")%>"
										alt="" />
																		<i class="icon-BulkBookingRecordFind"></i>
									<h6>Bulk<br />
										Booking</h6>
								</a>
							</div>
							<div id="mnutoolAddFromWaitingListRecordFind" class="button">
								<a href="#" rel="table" title="Add from Waiting List">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/AddFromWaitingList64HOVER.png")%>"
										alt="" />
																		<i class="icon-AddFromWaitingListRecordFind"></i>
									<h6>Add from<br />
										Waiting List</h6>
								</a>
							</div>
							<div id="mnutoolTransferBookingRecordFind" class="button">
								<a href="#" rel="table" title="Transfer Booking">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/TransferBooking64HOVER.png") %>"
										alt="" />
																		<i class="icon-TransferBookingRecordFind"></i>
									<h6>Transfer<br />
										Booking</h6>
								</a>
							</div>
							<div id="mnutoolCancelBookingRecordFind" class="button">
								<a href="#" rel="table" title="Cancel Booking">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/CancelBooking64HOVER.png") %>"
										alt="" />
																		<i class="icon-CancelBookingRecordFind"></i>
									<h6>Cancel<br />
										Booking</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionPositionRecordFind"><span>Record Count</span>
							<div id="mnutoolPositionRecordFind" class="textboxlist">
								<ul>
									<li>
										<span>Record(s) : x [(filtered)]</span>
									</li>
								</ul>
							</div>
						</li>
				</ul>
			</li>

								<%-- Record: Record Edit --%>
				<li class="ui-corner-top"><a id="toolbarRecord" href="#" rel="Record">Record</a>
					<ul>
						<li id="mnuSectionRecordEdit"><span>Edit</span>											
							<div id="mnutoolNewRecord" class="button">
								<a href="#" rel="table" title="New Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt="" />
									<i class="icon-plus"></i>
									<h6>New</h6>
								</a>
							</div>
							<div id="mnutoolCopyRecord" class="button">
								<a href="#" rel="table" title="Copy Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
									<i class="icon-copy"></i>
									<h6>Copy</h6>
								</a>
							</div>
							<div id="mnutoolEditRecord" class="button">
								<a href="#" rel="table" title="Edit Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
									<i class="icon-pencil"></i>
									<h6>Edit</h6>
								</a>
							</div>
							<div id="mnutoolSaveRecord" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png") %>" alt="" />
									<i class="icon-save"></i>
									<h6>Save</h6>
								</a>
							</div>
							<div id="mnutoolDeleteRecord" class="button">
								<a href="#" rel="table" title="Delete Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cross64HOVER.png")%>" alt="" />
									<i class="icon-remove"></i>
									<h6>Delete</h6></a>
							</div>
						</li>
						<li id="mnuSectionRecordNavigate"><span>Navigate</span>
							<div id="mnutoolParentRecord" class="button">
								<a href="#" rel="table" title="Return to parent record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/upblack64HOVER.png") %>" alt="" />
									<i class="icon-arrow-up"></i>
									<h6>Parent</h6>
								</a>
							</div>
							<div id="mnutoolBackRecord" class="button">
								<a href="#" rel="table" title="Return to record editing">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BackRecord64HOVER.png") %>" alt="" />
									<i class="icon-arrow-left"></i>
									<h6>Back</h6>
								</a>
							</div>

							<div id="mnutoolFirstRecord" class="button">
								<a href="#" rel="table" title="First Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/first64HOVER.png") %>" alt="" />
									<i class="icon-step-backward"></i>
									<h6>First</h6>
								</a>
							</div>
							<div id="mnutoolPreviousRecord" class="button">
								<a href="#" rel="table" title="Previous Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/back64HOVER.png") %>" alt="" />
									<i class="icon-backward"></i>
									<h6>Previous</h6>
								</a>
							</div>
							<div id="mnutoolNextRecord" class="button">
								<a href="#" rel="table" title="Next Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/next64HOVER.png") %>" alt="" />
									<i class="icon-forward"></i>
									<h6>Next</h6>
								</a>
							</div>
							<div id="mnutoolLastRecord" class="button">
								<a href="#" rel="table" title="Last Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/last64HOVER.png") %>" alt="" />
									<i class="icon-step-forward"></i>
									<h6>Last</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionRecordFind"><span>Find</span>
							<div id="mnutoolFindRecord" class="button">
								<a href="#" rel="table" title="Find">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Find64HOVER.png") %>" alt="" />
									<i class="icon-search"></i>
									<h6>Find</h6>
								</a>
							</div>
							<div id="mnutoolQuickFindRecord" class="button">
								<a href="#" rel="table" title="Quick Find">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/QuickFind64HOVER.png") %>" alt="" />
									<i class="icon-QuickFindRecord"></i>
									<h6>Quick Find</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionRecordOrder"><span>Order</span>
							<div id="mnutoolChangeOrderRecord" class="button" title="Change Order">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/sort64HOVER.png") %>" alt="" />
									<i class="icon-ChangeOrderRecord"></i>
									<h6>Change Order</h6>
								</a>
							</div>
							<div id="mnutoolFilterRecord" class="button" title="Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Filtering64HOVER.png") %>" alt="" />
									<i class="icon-FilterRecord"></i>
									<h6>Filter</h6>
								</a>
							</div>
							<div id="mnutoolClearFilterRecord" class="button" title="Clear Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/FilteringDelete64HOVER.png") %>"
										alt="" />
									<i class="icon-ClearFilterRecord"></i>
									<h6>Clear Filter</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionRecordCourseBooking"><span>Course Booking</span>
<%--
							<div id="mnutoolBookCourseRecord" class="button">
								<a href="#" rel="table" title="Book Course">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BookCourse64HOVER.png") %>"
										alt="" />
																		<i class="icon-BookCourseRecord"></i>
									<h6>Book<br />
										Course</h6>
								</a>
							</div>
--%>
							<div id="mnutoolCancelCourseRecord" class="button">
								<a href="#" rel="table" title="Cancel Course">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/CancelCourse64HOVER.png") %>"
										alt="" />
																		<i class="icon-CancelCourseRecord"></i>
									<h6>Cancel<br />
										Course</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionRecordReports"><span>Reports</span>
							<div id="mnutoolCalendarReportsRecord" class="button">
								<a href="#" rel="table" title="Calendar Reports">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/CalendarReports64HOVER.png") %>"
										alt="" />
									<i class="icon-CalendarReportsRecord"></i>
									<h6>Calendar<br />
										Reports</h6>
								</a>
							</div>
							<div id="mnutoolAbsenceBreakdownRecord" class="button">
								<a href="#" rel="table" title="Absence Breakdown">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/AbsenceBreakdown64HOVER.png") %>"
										alt="" />
																		<i class="icon-AbsenceBreakdownRecord"></i>
									<h6>Absence<br />
										Breakdown</h6>
								</a>
							</div>
							<div id="mnutoolAbsenceCalendarRecord" class="button">
								<a href="#" rel="table" title="Absence Calendar">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/AbsenceCalendar64HOVER.png") %>"
										alt="" />
																		<i class="icon-AbsenceCalendarRecord"></i>
									<h6>Absence<br />
										Calendar</h6>
								</a>
							</div>
							<div id="mnutoolBradfordRecord" class="button">
								<a href="#" rel="table" title="Bradford Factor">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BradfordFactor64HOVER.png") %>"
										alt="" /><i class="icon-BradfordRecord"></i>
									<h6>Bradford<br />
										Factor</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionRecordMailmerge"><span>Mail Merge</span>
							<div id="mnutoolMailMergeRecord" class="button">
								<a href="#" rel="table" title="Mail Merge">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/MailMerge64HOVER.png") %>"
										alt="" />
																		<i class="icon-MailMergeRecord"></i>
									<h6>Mail<br />
										Merge</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionPositionRecord"><span>Record Position</span>
							<div id="mnutoolRecordPosition" class="textboxlist">
								<ul>
									<li>
										<span>Record n of m [(filtered)]</span>
									</li>
								</ul>
							</div>
						</li>
						<li id="mnuSectionRecordMF"><span>Record Tools</span>
							<div id="mnutoolMFRecord" class="button">
								<a href="#" rel="table" title="Toggle Mandatory Fields">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Registry64HOVER.png")%>"
										alt="" />
									<i class="icon-th-list"></i>
									<h6>Highlight Mandatory<br/>Columns</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>


				<%-- Record - Absence Calendar --%>
				<li class="ui-corner-top"><a id="toolbarRecordAbsence" href="#" rel="toolbarRecord_Absence">Absence Calendar</a>
					<ul>
						<li id="mnuSectionRecordAbsence"><span>Absence Calendar</span>
							<div id="mnutoolPrintRecordAbsence" class="button" title="Print">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png") %>" alt="" />
									<i class="icon-PrintRecordAbsence"></i>
									<h6>Output</h6>
								</a>
							</div>
							<div id="mnutoolCloseRecordAbsence" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-removeRecordAbsence"></i>
									<h6>Close</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record - Quick Find --%>
								<li class="ui-corner-top"><a id="toolbarRecordQuickFind" href="#" rel="toolbarRecord_QuickFind">Quick Find</a>
					<ul>
						<li id="mnuSectionRecordQuickFind"><span>Quick Find</span>
							<div id="mnutoolFindRecordQuickFind" class="button" title="Find">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Find64HOVER.png")%>" alt="" />
									<i class="icon-FindRecordQuickFind"></i>
									<h6>Find</h6>
								</a>
							</div>
							<div id="mnutoolCloseRecordQuickFind" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-removeRecordQuickFind"></i>
									<h6>Close</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record - Sort Order --%>
								<li class="ui-corner-top"><a id="toolbarRecordSortOrder" href="#" rel="Record_SortOrder">Sort Order</a>
					<ul>
						<li id="mnuSectionRecordSortOrder"><span>Sort Order</span>
							<div id="mnutoolCheckRecordSortOrder" class="button" title="Select">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/checkmark64HOVER.png")%>" alt="" />
									<i class="icon-CheckRecordSortOrder"></i>
									<h6>Select</h6>
								</a>
							</div>
							<div id="mnutoolCloseRecordSortOrder" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-removeRecordSortOrder"></i>
									<h6>Close</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record - Filter --%>
								<li class="ui-corner-top"><a id="toolbarRecordFilter" href="#" rel="Record_Filter">Filter</a>
					<ul>
						<li id="mnuSectionRecordFilter"><span>Filter</span>
							<div id="mnutoolApplyRecordFilter" class="button" title="Apply">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/checkmark64HOVER.png")%>" alt="" />
									<i class="icon-ApplyRecordFilter"></i>
									<h6>Apply</h6>
								</a>
							</div>
							<div id="mnutoolCloseRecordFilter" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-removeRecordFilter"></i>
									<h6>Close</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record - Mail Merge --%>
								<li class="ui-corner-top"><a id="toolbarRecordMailMerge" href="#" rel="Record_MailMerge">Filter</a>
					<ul>
						<li id="mnuSectionRecordMailMerge"><span>Filter</span>
							<div id="mnutoolRunRecordMailMerge" class="button" title="Run">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
									<i class="icon-RunRecordMailMerge"></i>
									<h6>Run</h6>
								</a>
							</div>
							<div id="mnutoolCloseRecordMailMerge" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-removeRecordMailMerge"></i>
									<h6>Close</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>

				<%-- Record - Booking - Transfer Booking / Add from waiting list --%>
				<li class="ui-corner-top"><a id="toolbarDelegateBookingTransfer" href="#" rel="Delegate Booking">Delegate Booking</a>
					<ul>
						<li id="mnuSectionSelectDelegateBookingTransfer"><span>Select</span>
							<div id="mnutoolSelectDelegateBookingTransfer" class="button" title="Select">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/checkmark64HOVER.png")%>" alt="" />
									<i class="icon-OK"></i>
									<h6>Select</h6>
								</a>
							</div>
							<div id="mnutoolCloseDelegateBookingTransfer" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-remove"></i>
								<h6>Close</h6></a>
							</div>
						</li>
						<li id="mnuSectionNavigateDelegateBookingTransfer"><span>Navigate</span>
							<div id="mnutoolFirstDelegateBookingTransfer" class="button">
								<a href="#" rel="table" title="First Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/first64HOVER.png") %>" alt="" />
									<i class="icon-step-backward"></i>
									<h6>First</h6>
								</a>
							</div>
							<div id="mnutoolPreviousDelegateBookingTransfer" class="button">
								<a href="#" rel="table" title="Previous Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/back64HOVER.png") %>" alt="" />
									<i class="icon-backward"></i>
									<h6>Previous</h6>
								</a>
							</div>
							<div id="mnutoolNextDelegateBookingTransfer" class="button">
								<a href="#" rel="table" title="Next Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/next64HOVER.png") %>" alt="" />
									<i class="icon-forward"></i>
									<h6>Next</h6>
								</a>
							</div>
							<div id="mnutoolLastDelegateBookingTransfer" class="button">
								<a href="#" rel="table" title="Last Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/last64HOVER.png") %>" alt="" />
									<i class="icon-step-forward"></i>
									<h6>Last</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>

				<%-- Record - Booking - Bulk Booking --%>

				<li class="ui-corner-top"><a id="toolbarDelegateBookingBulkBooking" href="#" rel="Report_NewEditCopy">Bulk Booking</a>
					<ul>
						<li id="mnuSectionDelegateBookingBulkBooking"><span>Report</span>
							<div id="mnutoolSaveDelegateBookingBulkBooking" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png")%>" alt="" />
									<i class="icon-save"></i>
									<h6>Save</h6>
								</a>
							</div>
							<div id="mnutoolCancelDelegateBookingBulkBooking" class="button">
								<a href="#" rel="table" title="Cancel">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cancel64HOVER.png")%>" alt="" />
								<i class="icon-remove"></i>
								<h6>Cancel</h6></a>
							</div>
						</li>
					</ul>
				</li>

				<%-- Report NewEditCopy --%>
								<%-- Report Find --%>
								<li class="ui-corner-top"><a id="toolbarReportFind" href="#" rel="Report_Find">Find</a>
					<ul>
						<li id="mnuSectionReportFind"><span>Find</span>
							<div id="mnutoolNewReportFind" class="button">
								<a href="#" rel="table" title="New">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt="" />
																<i class="icon-plus"></i>
									<h6>New</h6>
								</a>
							</div>
							<div id="mnutoolCopyReportFind" class="button">
								<a href="#" rel="table" title="Copy">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
								<i class="icon-copy"></i>
									<h6>Copy</h6>
								</a>
							</div>
							<div id="mnutoolEditReportFind" class="button">
								<a href="#" rel="table" title="Edit">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
								<i class="icon-pencil"></i>
									<h6>Edit</h6>
								</a>
							</div>
							<div id="mnutoolDeleteReportFind" class="button">
								<a href="#" rel="table" title="Delete">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cross64HOVER.png")%>" alt="" />
								<i class="icon-remove"></i>
									<h6>Delete</h6>
								</a>
							</div>
							<div id="mnutoolPropertiesReportFind" class="button">
								<a href="#" rel="table" title="Properties">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png")%>" alt="" />								
								<i class="icon-PropertiesReportFind"></i>
									<h6>Properties</h6>
								</a>
							</div>
							<div id="mnutoolRunReportFind" class="button">
								<a href="#" rel="table" title="Run">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
								<i class="icon-RunReportFind"></i>
									<h6>Run</h6>
								</a>
							</div>
							<div id="mnutoolCloseReportFind" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-removeReportFind"></i>
									<h6>Close</h6>
								</a>
							</div>
						</li>						
					</ul>
				</li>

								<%-- Report NewEditCopy --%>
								<li class="ui-corner-top"><a id="toolbarReportNewEditCopy" href="#" rel="Report_NewEditCopy">Definition</a>
					<ul>
						<li id="mnuSectionNewEditCopyReport"><span>Definition</span>
							<div id="mnutoolSaveReport" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png")%>" alt="" />
																<i class="icon-save"></i>
									<h6>Save</h6>
								</a>
							</div>
							<div id="mnutoolCancelReport" class="button">
								<a href="#" rel="table" title="Cancel">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cancel64HOVER.png")%>" alt="" />
								<i class="icon-CancelReport"></i>
									<h6>Cancel</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Report Run --%>
								<li class="ui-corner-top"><a id="toolbarReportRun" href="#" rel="Report_Run">Definition</a>
					<ul>
						<li id="mnuSectionRunReport"><span>Output</span>
							<div id="mnutoolOutputReport" class="button">
								<a href="#" rel="table" title="Output">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png")%>" alt="" />
																<i class="icon-OutputReportRun"></i>
									<h6>Output</h6>
								</a>
							</div>
							<div id="mnutoolCloseReport" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-mnutoolCloseReportRun"></i>
									<h6>Close</h6>
								</a>
							</div>
						</li>						
					</ul>
				</li>									

								<%-- Utilities Find --%>
								<li class="ui-corner-top"><a id="toolbarUtilitiesFind" href="#" rel="Utilities_Find">Find</a>
					<ul>
						<li id="mnuSectionUtilitiesFind"><span>Find</span>
							<div id="mnutoolNewUtilitiesFind" class="button">
								<a href="#" rel="table" title="New">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt="" />
																<i class="icon-plus"></i>
									<h6>New</h6>
								</a>
							</div>
							<div id="mnutoolCopyUtilitiesFind" class="button">
								<a href="#" rel="table" title="Copy">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
								<i class="icon-copy"></i>
									<h6>Copy</h6>
								</a>
							</div>
							<div id="mnutoolEditUtilitiesFind" class="button">
								<a href="#" rel="table" title="Edit">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
								<i class="icon-pencil"></i>
									<h6>Edit</h6>
								</a>
							</div>
							<div id="mnutoolDeleteUtilitiesFind" class="button">
								<a href="#" rel="table" title="Delete">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cross64HOVER.png")%>" alt="" />
								<i class="icon-remove"></i>
									<h6>Delete</h6>
								</a>
							</div>
							<div id="mnutoolPropertiesUtilitiesFind" class="button">
								<a href="#" rel="table" title="Properties">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png")%>" alt="" />								
								<i class="icon-PropertiesUtilitiesFind"></i>
									<h6>Properties</h6>
								</a>
							</div>
							<div id="mnutoolRunUtilitiesFind" class="button">
								<a href="#" rel="table" title="Run">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
								<i class="icon-RunUtilitiesFind"></i>
									<h6>Run</h6>
								</a>
							</div>
							<div id="mnutoolCloseUtilitiesFind" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-removeUtilitiesFind"></i>
									<h6>Close</h6>
								</a>
							</div>
						</li>						
					</ul>
				</li>

								<%-- Utilities NewEditCopy --%>
								<li class="ui-corner-top"><a id="toolbarUtilitiesNewEditCopy" href="#" rel="Utilities_NewEditCopy">Utilities</a>
					<ul>
						<li id="mnuSectionNewEditCopyUtilities"><span>Utilities</span>
							<div id="mnutoolSaveUtilities" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png")%>" alt="" />
																<i class="icon-save"></i>
									<h6>Save</h6>
								</a>
							</div>
							<div id="mnutoolCancelUtilities" class="button">
								<a href="#" rel="table" title="Cancel">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cancel64HOVER.png")%>" alt="" />
								<i class="icon-CancelUtilities"></i>
									<h6>Cancel</h6>
								</a>
							</div>
						</li>						
					</ul>
				</li>

								<%-- Tools Find --%>
				<li class="ui-corner-top"><a id="toolbarToolsFind" href="#" rel="Tools_Find">Find</a>
					<ul>
						<li id="mnuSectionToolsFind"><span>Find</span>
							<div id="mnutoolNewToolsFind" class="button">
								<a href="#" rel="table" title="New">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt="" />
																<i class="icon-plus"></i>
									<h6>New</h6>
								</a>
							</div>
							<div id="mnutoolCopyToolsFind" class="button">
								<a href="#" rel="table" title="Copy">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
								<i class="icon-copy"></i>
									<h6>Copy</h6>
								</a>
							</div>
							<div id="mnutoolEditToolsFind" class="button">
								<a href="#" rel="table" title="Edit">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
								<i class="icon-pencil"></i>
									<h6>Edit</h6>
								</a>
							</div>
							<div id="mnutoolDeleteToolsFind" class="button">
								<a href="#" rel="table" title="Delete">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cross64HOVER.png")%>" alt="" />
								<i class="icon-remove"></i>
									<h6>Delete</h6>
								</a>
							</div>
							<div id="mnutoolPropertiesToolsFind" class="button">
								<a href="#" rel="table" title="Properties">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png")%>" alt="" />								
								<i class="icon-PropertiesToolsFind"></i>
									<h6>Properties</h6>
								</a>
							</div>
							<div id="mnutoolRunToolsFind" class="button">
								<a href="#" rel="table" title="Run">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
								<i class="icon-RunToolsFind"></i>
									<h6>Run</h6>
								</a>
							</div>
						</li>						
					</ul>
				</li>
				

				<%-- Org Chart--%>
				<li class="ui-corner-top"><a id="toolbarOrgChart" href="#" rel="Find">Organisation Chart</a>
					<ul>
						<li><span>Output</span>
							<div class="button" id="divBtnPrintOrgChart">
								<a class="mnuBtnPrintOrgChart" href="#" rel="paste">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png")%>" alt="" /><span>
									<i class="icon-print"></i>
									<h6>Print</h6></span>
								</a>								
							</div>
							<div class="button">
								<a class="mnuBtnSelectOrgChart" href="#" rel="table" title="Select nodes to print">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Checkmark64HOVER.png")%>" alt="" /><span>			
									<i class="icon-check"></i>
									<h6>Select nodes<br/>to print</h6></span>
								</a>
							</div>
						</li>
						<li><span>Interact</span>
							<div id="mnutoolOrgChartExpand" class="button">
								<a href="#" rel="table" title="Expand all nodes">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt="" />
									<i class="icon-plus"></i>
									<h6>Expand All Nodes</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>
				
				<%-- EventLog Find--%>
				<li class="ui-corner-top"><a id="toolbarEventLogFind" href="#" rel="Find">Find</a>
					<ul>
						<li><span>Edit</span>											
							<div id="mnutoolViewEventLogFind" class="button">
								<a href="#" rel="table" title="View">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/preview64HOVER.png")%>" alt="" />
									<i class="icon-ViewEventLogFind"></i> 
									<h6>View</h6>
								</a>
							</div>
							<div id="mnutoolPurgeEventLogFind" class="button">
								<a href="#" rel="table" title="Purge">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Purge64HOVER.png")%>" alt="" />
									<i class="icon-PurgeEventLogFind"></i>
									<h6>Purge</h6>
								</a>
							</div>
							<div id="mnutoolEmailEventLogFind" class="button">
								<a href="#" rel="table" title="Email">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Email64HOVER.png")%>" alt="" />
									<i class="icon-EmailEventLogFind"></i>
									<h6>Email</h6>
								</a>
							</div>
							<div id="mnutoolDeleteEventLogFind" class="button">
								<a href="#" rel="table" title="Delete">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cross64HOVER.png")%>" alt="" />
									<i class="icon-remove"></i>
									<h6>Delete</h6>
								</a>
							</div>
						</li>
						<li id="mnuSectionNavigateRecords"><span>Navigate</span>
							<div id="mnutoolFirstEventLogFind" class="button">
								<a href="#" rel="table" title="First Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/first64HOVER.png") %>" alt="" />
									<i class="icon-step-backward"></i>
									<h6>First</h6>
								</a>
							</div>
							<div id="mnutoolPreviousEventLogFind" class="button">
								<a href="#" rel="table" title="Previous Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/back64HOVER.png") %>" alt="" />
									<i class="icon-backward"></i>
									<h6>Previous</h6>
								</a>
							</div>
							<div id="mnutoolNextEventLogFind" class="button">
								<a href="#" rel="table" title="Next Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/next64HOVER.png") %>" alt="" />
									<i class="icon-forward"></i>
									<h6>Next</h6>
								</a>
							</div>
							<div id="mnutoolLastEventLogFind" class="button">
								<a href="#" rel="table" title="Last Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/last64HOVER.png") %>" alt="" />
									<i class="icon-step-forward"></i>
									<h6>Last</h6>
								</a>
							</div>
						</li>						
						<li id="mnuSectionPositionEventLog"><span>Record Count</span>
							<div id="mnutoolRecordEventLog" class="textboxlist">
								<ul>
									<li>
										<span>Record n of m [(filtered)]</span>
									</li>
								</ul>
							</div>
						</li>						
					</ul>
				</li>

				<li class="ui-corner-top"><a id="toolbarEventLogView" href="#" rel="Find">View</a>
					<ul>
						<li><span>View</span>											
							<div id="mnutoolEmailEventLogView" class="button">
								<a href="#" rel="table" title="Email">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Email64HOVER.png")%>" alt="" />
									<i class="icon-EmailEventLogView"></i> 
									<h6>Email</h6>
								</a>
							</div>
							<div id="mnutoolOutputEventLogView" class="button">
								<a href="#" rel="table" title="Output">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png")%>" alt="" />
									<i class="icon-OutputEventLogView"></i> 
									<h6>Email</h6>
								</a>
							</div>
							<div id="mnutoolCloseEventLogView" class="button">
								<a href="#" rel="table" title="View">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
									<i class="icon-removeEventLogView"></i> 
									<h6>Email</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>


				<li class="ui-corner-top"><a id="toolbarWFPendingStepsFind" href="#" rel="Find">Find</a>
					<ul>
						<li><span>Find</span>											
							<div id="mnutoolRefreshWFPendingStepsFind" class="button">
								<a href="#" rel="table" title="Refresh">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Refresh64HOVER.png")%>" alt="" />
								<i class="icon-refresh"></i>
								<h6>Refresh</h6>
								</a>
							</div>
							<div id="mnutoolRunWFPendingStepsFind" class="button">
								<a href="#" rel="table" title="Run">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
								<i class="icon-play-circle"></i>
								<h6>Run</h6>
								</a>
							</div>
							<div id="mnutoolCloseWFPendingStepsFind" class="button">
								<%--<a href="#" rel="table" title="Close">--%>
									<a href="<%: Url.Action("Main", "Home", New With {.SSIMode = vbFalse})%>" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-remove"></i>
								<h6>Close</h6>
								</a>
							</div>
						</li>						
					</ul>
				</li>

				<li class="ui-corner-top"><a id="toolbarAdminConfig" href="#" rel="Find">Configure</a>
					<ul>
						<li><span>Configure</span>											
											<div id="mnutoolSaveAdminConfig" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png")%>" alt="" />
								<i class="icon-save"></i>
									<h6>Save</h6>
								</a>
							</div>
						</li>						
					</ul>
				</li>
				
				<li class="ui-corner-top"><a id="toolbarStandardReportConfig" href="#" rel="Report">Configure</a>
					<ul>
						<li><span>Configure</span>
							<div id="mnutoolSaveStandardReportConfig" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png")%>" alt="" />
									<i class="icon-save"></i>
									<h6>Save</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>
			</ul>
		</div>
	</div>
	<div class="FixedLinksRight" title="<%=Session("welcomemessage")%>">
		<div class="userpic">
			<img id="UserPicture" style="vertical-align: middle; height: 48px; width: 48px;" src="<%=Session("SelfServicePhotograph_src")%>" alt="Photo" />
		</div>
		<div class="userdetails">
		<div class="userid">
				&nbsp;
			</div>
		<div class="groupid">
			<%=Session("UserGroup")%>		
		</div>
		</div>
	</div>
	<!-- User dropdown menu -->
	<div id="userDropdownmenu">
		<a href="#"><%=Session("welcomeName")%> ▼</a>
		<ul>
			<li class="active has-sub last">
				<ul id="userDropdownmenu_Items" class="ui-widget-header">
					<li class="linkspagebuttontext">
						<a id="mnutoolFixedPasswordChange" href="#">
							<span>Change Password</span>
						</a>
					</li>
					<%If UCase(Session("WF_Enabled")) = "TRUE" Then%>
					<li class="linkspagebuttontext">
						<a id="mnutoolFixedPWFS" href="#" rel="table" title="">
							<span>Pending Workflow Steps</span>
						</a>
					</li>
					<%End If%>
					<li class="linkspagebuttontext">
						<a id="mnutoolFixedWorkflowOutOfOffice" href="#">
							<span>Out Of Office</span>
						</a>
					</li>
					<%If UCase(Session("ui-layout-selectable")) = "TRUE" Then%>
					<li class="linkspagebuttontext" id="userDropdownmenu_Layout">
						<a onclick="javascript: showThemeEditor();" href="#">
							<span>Layout</span>
						</a>
					</li>
					<%End If%>
					<li class="linkspagebuttontext">
						<a id="mnutoolFixedAbout" onclick="javascript: fixedlinks_mnutoolAboutHRPro();" href="#">
							<span>About</span>
						</a>
					</li>
					<li class="linkspagebuttontext">
						<a id="mnutoolLogoff" href="#">
							<span>Log Off</span>
						</a>
					</li>

				</ul>
			</li>
		</ul>
	</div>
</div>