<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%If Session("databaseConnection") Is Nothing Then Return%>
<script type="text/javascript">
	$(document).ready(function () {
		$(".officebar").officebar({});

		//$("#toolbarRecord").hide();
		//$("#mnuSectionUtilities").hide();

				// Commented out for now
		//menu_toolbarEnableItem("mnutoolNewUtil", false);
		//menu_toolbarEnableItem("mnutoolEditUtil", false);
		//menu_toolbarEnableItem("mnutoolCopyUtil", false);
		//menu_toolbarEnableItem("mnutoolDeleteUtil", false);
		//menu_toolbarEnableItem("mnutoolPrintUtil", false);
		//menu_toolbarEnableItem("mnutoolPropertiesUtil", false);
		//menu_toolbarEnableItem("mnutillRunUtil", false);

		//$("#toolbarHome").click();

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
		
		});




	function wrapTileIcons() {
		if (window.currentLayout == "tiles") {
			//Wrap the icons with circles or boxes or whatever...
			//$(".officetab i[class^='icon-']").css("padding-left", "7px");
			//$(".officetab i[class^='icon-']").css("padding-bottom", "4px");
			$(".officetab i[class^='icon-']").wrap("<span class='icon-stack' />");
			$(".officetab .icon-stack").prepend("<i class='icon-check-empty icon-stack-base'></i>");
		}
	}

	function fixedlinks_mnutoolAboutHRPro() {
		$("#About").dialog("open");		
	}

	function showThemeEditor() {
		$("#divthemeRoller").dialog("open");
		$("#themeeditoraccordion").accordion("resize");
		
		//load the themeeditor form now
		loadPartialView("themeEditor", "home", "divthemeRoller", null);

	}

	//why was this here?...
	//$("#officebar").tabs();

</script>
<div id="fixedlinks">
	<div class="RecordDescription">
		<p><a href="<%: Url.Action("Main", "Home", New With {.SSIMode = vbTrue})%>" title="Home"><%=Session("recdesc")%></a></p>
	</div>
	<div class="FixedLinksLeft">
		<div id="officebar" class="officebar ui-widget-header ui-widget-content">
			<ul>
						<%-- Home --%>
				<li class="current ui-state-default"><a id="toolbarHome" href="#" rel="home">Home</a>
					<ul>
						<li><span>Fixed Links</span>
							<div id="mnutoolFixedSelfService" class="button">
									<a href="<%: Url.Action("Main", "Home", New With {.SSIMode = vbTrue})%>" rel="table" title="Self-service">
								<img src="<%: Url.Content("~/Scripts/officebar/winkit/abssmall.png")%>" alt="" />
									<i class="icon-user"></i>
										<h6>Self-service</h6>
									</a>
							</div>
							<div id="mnutoolFixedOpenHR" class="button">
								<a href="<%: Url.Action("Main", "Home", New With {.SSIMode = vbFalse})%>" rel="table" title="OpenHR">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/abssmall.png") %>" alt="" />
									<i class="icon-group"></i>
									<h6>OpenHR</h6>
								</a>
							</div>
							<div id="mnutoolFixedLogoff" class="button">
								<a href="<%: Url.Action("LogOff", "Home") %>" rel="table" title="Log Off">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Logoff64HOVER.png")%>" alt="" />
									<i class="icon-off"></i>
									<h6>Log Off</h6>
								</a>
							</div>
							<div id="mnutoolFixedChangePassword" class="button">
								<a href="#" rel="table" title="Change Password">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/ChangePassword64HOVER.png") %>"
										alt="" />
									<i class="icon-lock"></i>
									<h6>Change<br/>Password</h6>
								</a>
							</div>
							<div id="mnutoolFixedLayout" class="button">
								<a href="javascript:showThemeEditor()" rel="layout" title="Change Layout">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png") %>" alt="" />
									<i class="icon-wrench"></i>
									<h6>Layout</h6>
								</a>
							</div>
<%--							
							<div class="button">
								<a href="#" rel="table" title="Pending Workflow Steps">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Workflow64HOVER.png") %>" alt="" />
									<i class="icon-inbox"></i>
									<h6>Workflow</h6>
								</a>
							</div>
--%>
							<div id="mnutoolFixedAbout" class="button">
								<a href="javascript:fixedlinks_mnutoolAboutHRPro()" rel="table" title="About OpenHR v8">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/info64HOVER.png")%>" alt="" />
									<i class="icon-question-sign"></i>
									<h6>Help</h6>
								</a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record: Find Record--%>
				<li class="ui-state-default ui-corner-top"><a id="toolbarRecordFind" href="#" rel="Find">Find</a>
					<ul>
						<li id="mnuSectionRecordFindEdit"><span>Edit</span>											
							<div id="mnutoolNewRecordFind" class="button">
								<a href="#" rel="table" title="New Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt="" />
									<i class="icon-plus"></i>  
									<h6>New</h6>
							</a></div>
							<div id="mnutoolCopyRecordFind" class="button">
								<a href="#" rel="table" title="Copy Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
									<i class="icon-copy"></i>
									<h6>Copy</h6></a>
							</div>
							<div id="mnutoolEditRecordFind" class="button">
								<a href="#" rel="table" title="Edit Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
									<i class="icon-pencil"></i>
									<h6>Edit</h6></a>
							</div>
							<div id="mnutoolDeleteRecordFind" class="button">
								<a href="#" rel="table" title="Delete Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/delete64HOVER.png") %>" alt="" />
									<i class="icon-close"></i>
									<h6>Delete</h6></a>
							</div>
						</li>
						<li id="mnuSectionRecordFindNavigate"><span>Navigate</span>
							<div id="mnutoolParentRecordFind" class="button">
								<a href="#" rel="table" title="Return to parent record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/upblack64HOVER.png") %>" alt="" />
									<i class="icon-arrow-up"></i>
									<h6>Parent</h6></a>
							</div>
							<div id="mnutoolBackRecordFind" class="button">
								<a href="#" rel="table" title="Return to record editing">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BackRecord64HOVER.png") %>" alt="" />
									<i class="icon-arrow-left"></i>
									<h6>Back</h6></a>
							</div>

							<div id="mnutoolFirstRecordFind" class="button">
								<a href="#" rel="table" title="First Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/first64HOVER.png") %>" alt="" />
									<i class="icon-previous"></i>
									<h6>First</h6></a>
							</div>
							<div id="mnutoolPreviousRecordFind" class="button">
								<a href="#" rel="table" title="Previous Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/back64HOVER.png") %>" alt="" />
									<i class="icon-backward"></i>
									<h6>Previous</h6></a>
							</div>
							<div id="mnutoolNextRecordFind" class="button">
								<a href="#" rel="table" title="Next Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/next64HOVER.png") %>" alt="" />
									<i class="icon-forward"></i>
									<h6>Next</h6></a>
							</div>
							<div id="mnutoolLastRecordFind" class="button">
								<a href="#" rel="table" title="Last Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/last64HOVER.png") %>" alt="" />
									<i class="icon-next"></i>
									<h6>Last</h6></a>
							</div>
						</li>
						<li id="mnuSectionLocateRecordFind"><span>Locate Record</span>
							<div id="mnutoolLocateRecordFind" class="textboxlist">
								<ul><li>Go To<input type="text" /></li></ul>
							</div>
						</li>
												<li id="mnuSectionRecordFindOrder"><span>Order</span>
							<div id="mnutoolChangeOrderRecordFind" class="button" title="Change Order">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/sort64HOVER.png") %>" alt="" />
									<i class="icon-ChangeOrderRecordFind"></i>
									<h6>Change Order</h6></a>
							</div>
							<div id="mnutoolFilterRecordFind" class="button" title="Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Filtering64HOVER.png") %>" alt="" />
									<i class="icon-FilterRecordFind"></i>
									<h6>Filter</h6></a>
							</div>
							<div id="mnutoolClearFilterRecordFind" class="button" title="Clear Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/FilteringDelete64HOVER.png") %>"
										alt="" />
									<i class="icon-ClearFilterRecordFind"></i>
									<h6>Clear Filter</h6></a>
							</div>
						</li>
						<li id="mnuSectionRecordFindTrainingBooking"><span>Training Booking</span>
							<div id="mnutoolBulkBookingRecordFind" class="button">
								<a href="#" rel="table" title="Bulk Booking">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BulkBooking64HOVER.png")%>"
										alt="" />
																		<i class="icon-BulkBookingRecordFind"></i>
																		<h6>Bulk<br />Booking</h6></a>
							</div>
							<div id="mnutoolAddFromWaitingListRecordFind" class="button">
								<a href="#" rel="table" title="Add from Waiting List">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/AddFromWaitingList64HOVER.png")%>"
										alt="" />
																		<i class="icon-AddFromWaitingListRecordFind"></i>
																		<h6>Add from<br />Waiting List</h6></a>
							</div>
							<div id="mnutoolTransferBookingRecordFind" class="button">
								<a href="#" rel="table" title="Transfer Booking">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/TransferBooking64HOVER.png") %>"
										alt="" />
																		<i class="icon-TransferBookingRecordFind"></i>
																		<h6>Transfer<br />Booking</h6></a>
							</div>
							<div id="mnutoolCancelBookingRecordFind" class="button">
								<a href="#" rel="table" title="Cancel Booking">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/CancelBooking64HOVER.png") %>"
										alt="" />
																		<i class="icon-CancelBookingRecordFind"></i>
																		<h6>Cancel<br />Booking</h6></a>
							</div>
						</li>
						<li id="mnuSectionPositionRecordFind"><span>Record Position</span>
							<div id="mnutoolPositionRecordFind" class="textboxlist">
								<ul>
									<li>
										<span>Record n of m [(filtered)]</span>
									</li>
								</ul>
							</div>
						</li>
				</ul>
			</li>

								<%-- Record: Record Edit --%>
				<li class="ui-state-default ui-corner-top"><a id="toolbarRecord" href="#" rel="Record">Record</a>
					<ul>
						<li id="mnuSectionRecordEdit"><span>Edit</span>											
							<div id="mnutoolNewRecord" class="button">
								<a href="#" rel="table" title="New Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt="" />
									<i class="icon-plus"></i>
									<h6>Add</h6>
							</a></div>
							<div id="mnutoolEditRecord" class="button">
								<a href="#" rel="table" title="Edit Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
									<i class="icon-pencil"></i>
									<h6>Edit</h6></a>
							</div>
							<div id="mnutoolSaveRecord" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png") %>" alt="" />
									<i class="icon-disk"></i>
									<h6>Save</h6></a>
							</div>
							<div id="mnutoolDeleteRecord" class="button">
								<a href="#" rel="table" title="Delete Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/delete64HOVER.png") %>" alt="" />
									<i class="icon-minus"></i>
									<h6>Delete</h6></a>
							</div>
						</li>
						<li id="mnuSectionRecordNavigate"><span>Navigate</span>
							<div id="mnutoolParentRecord" class="button">
								<a href="#" rel="table" title="Return to parent record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/upblack64HOVER.png") %>" alt="" />
									<i class="icon-arrow-up"></i>
									<h6>Parent</h6></a>
							</div>
							<div id="mnutoolBackRecord" class="button">
								<a href="#" rel="table" title="Return to record editing">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BackRecord64HOVER.png") %>" alt="" />
									<i class="icon-arrow-left"></i>
									<h6>Back</h6></a>
							</div>

							<div id="mnutoolFirstRecord" class="button">
								<a href="#" rel="table" title="First Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/first64HOVER.png") %>" alt="" />
									<i class="icon-previous"></i>
									<h6>First</h6></a>
							</div>
							<div id="mnutoolPreviousRecord" class="button">
								<a href="#" rel="table" title="Previous Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/back64HOVER.png") %>" alt="" />
									<i class="icon-backward"></i>
									<h6>Previous</h6></a>
							</div>
							<div id="mnutoolNextRecord" class="button">
								<a href="#" rel="table" title="Next Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/next64HOVER.png") %>" alt="" />
									<i class="icon-forward"></i>
									<h6>Next</h6></a>
							</div>
							<div id="mnutoolLastRecord" class="button">
								<a href="#" rel="table" title="Last Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/last64HOVER.png") %>" alt="" />
									<i class="icon-next"></i>
									<h6>Last</h6></a>
							</div>
						</li>
						<li id="mnuSectionRecordFind"><span>Find</span>
							<div id="mnutoolFindRecord" class="button">
								<a href="#" rel="table" title="Find">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Find64HOVER.png") %>" alt="" />
									<i class="icon-FindRecord"></i>
									<h6>Find</h6></a>
							</div>
							<div id="mnutoolQuickFindRecord" class="button">
								<a href="#" rel="table" title="Quick Find">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/QuickFind64HOVER.png") %>" alt="" />
									<i class="icon-QuickFindRecord"></i>
									<h6>Quick Find</h6></a>
							</div>
						</li>
						<li id="mnuSectionRecordOrder"><span>Order</span>
							<div id="mnutoolChangeOrderRecord" class="button" title="Change Order">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/sort64HOVER.png") %>" alt="" />
									<i class="icon-ChangeOrderRecord"></i>
									<h6>Change Order</h6></a>
							</div>
							<div id="mnutoolFilterRecord" class="button" title="Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Filtering64HOVER.png") %>" alt="" />
									<i class="icon-FilterRecord"></i>
									<h6>Filter</h6></a>
							</div>
							<div id="mnutoolClearFilterRecord" class="button" title="Clear Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/FilteringDelete64HOVER.png") %>"
										alt="" />
									<i class="icon-ClearFilterRecord"></i>
									<h6>Clear Filter</h6></a>
							</div>
							<div id="mnutoolPrintRecord" class="button" title="Output">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png") %>" alt="" />
									<i class="icon-print"></i>
									<h6>Output</h6></a>
							</div>
						</li>
						<li id="mnuSectionRecordCourseBooking"><span>Course Booking</span>
							<div id="mnutoolBookCourseRecord" class="button">
								<a href="#" rel="table" title="Book Course">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BookCourse64HOVER.png") %>"
										alt="" />
																		<i class="icon-BookCourseRecord"></i>
																		<h6>Book<br />Course</h6></a>
							</div>
							<div id="mnutoolCancelCourseRecord" class="button">
								<a href="#" rel="table" title="Cancel Course">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/CancelCourse64HOVER.png") %>"
										alt="" />
																		<i class="icon-CancelCourseRecord"></i>
																		<h6>Cancel<br />Course</h6></a>
							</div>
						</li>
						<li id="mnuSectionRecordReports"><span>Reports</span>
							<div id="mnutoolCalendarReportsRecord" class="button">
								<a href="#" rel="table" title="Calendar Reports">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/CalendarReports64HOVER.png") %>"
										alt="" />
									<i class="icon-CalendarReportsRecord"></i>
																		<h6>Calendar<br />Reports</h6></a>
							</div>
							<div id="mnutoolAbsenceBreakdownRecord" class="button">
								<a href="#" rel="table" title="Absence Breakdown">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/AbsenceBreakdown64HOVER.png") %>"
										alt="" />
																		<i class="icon-AbsenceBreakdownRecord"></i>
																		<h6>Absence<br />Breakdown</h6></a>
							</div>
							<div id="mnutoolAbsenceCalendarRecord" class="button">
								<a href="#" rel="table" title="Absence Calendar">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/AbsenceCalendar64HOVER.png") %>"
										alt="" />
																		<i class="icon-AbsenceCalendarRecord"></i>
																		<h6>Absence<br />Calendar</h6></a>
							</div>
							<div id="mnutoolBradfordRecord" class="button">
								<a href="#" rel="table" title="Bradford Factor">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BradfordFactor64HOVER.png") %>"
										alt="" /><i class="icon-BradfordRecord"></i>
																		<h6>Bradford<br />Factor</h6></a>
							</div>
						</li>
						<li id="mnuSectionRecordMailmerge"><span>Mail Merge</span>
							<div id="mnutoolMailMergeRecord" class="button">
								<a href="#" rel="table" title="Mail Merge">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/MailMerge64HOVER.png") %>"
										alt="" />
																		<i class="icon-MailMergeRecord"></i>
																		<h6>Mail<br/>Merge</h6></a>
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
					</ul>
				</li>


				<%-- Record - Absence Calendar --%>
				<li class="ui-state-default ui-corner-top"><a id="toolbarRecordAbsence" href="#" rel="toolbarRecord_Absence">Absence Calendar</a>
					<ul>
						<li id="mnuSectionRecordAbsence"><span>Absence Calendar</span>
							<div id="mnutoolPrintRecordAbsence" class="button" title="Print">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png") %>" alt="" />
									<i class="icon-PrintRecordAbsence"></i>
									<h6>Output</h6></a>
							</div>
							<div id="mnutoolCloseRecordAbsence" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-CloseRecordAbsence"></i>
								<h6>Close</h6></a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record - Quick Find --%>
								<li class="ui-state-default ui-corner-top"><a id="toolbarRecordQuickFind" href="#" rel="toolbarRecord_QuickFind">Quick Find</a>
					<ul>
						<li id="mnuSectionRecordQuickFind"><span>Quick Find</span>
							<div id="mnutoolFindRecordQuickFind" class="button" title="Find">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Find64HOVER.png")%>" alt="" />
									<i class="icon-FindRecordQuickFind"></i>
									<h6>Find</h6></a>
							</div>
							<div id="mnutoolCloseRecordQuickFind" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-CloseRecordQuickFind"></i>
								<h6>Close</h6></a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record - Sort Order --%>
								<li class="ui-state-default ui-corner-top"><a id="toolbarRecordSortOrder" href="#" rel="Record_SortOrder">Sort Order</a>
					<ul>
						<li id="mnuSectionRecordSortOrder"><span>Sort Order</span>
							<div id="mnutoolCheckRecordSortOrder" class="button" title="Select">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/checkmark64HOVER.png")%>" alt="" />
									<i class="icon-CheckRecordSortOrder"></i>
									<h6>Select</h6></a>
							</div>
							<div id="mnutoolCloseRecordSortOrder" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-CloseRecordSortOrder"></i>
								<h6>Close</h6></a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record - Filter --%>
								<li class="ui-state-default ui-corner-top"><a id="toolbarRecordFilter" href="#" rel="Record_Filter">Filter</a>
					<ul>
						<li id="mnuSectionRecordFilter"><span>Filter</span>
							<div id="mnutoolApplyRecordFilter" class="button" title="Apply">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/checkmark64HOVER.png")%>" alt="" />
									<i class="icon-ApplyRecordFilter"></i>
									<h6>Apply</h6></a>
							</div>
							<div id="mnutoolCloseRecordFilter" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-CloseRecordFilter"></i>
								<h6>Close</h6></a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Record - Mail Merge --%>
								<li class="ui-state-default ui-corner-top"><a id="toolbarRecordMailMerge" href="#" rel="Record_MailMerge">Filter</a>
					<ul>
						<li id="mnuSectionRecordMailMerge"><span>Filter</span>
							<div id="mnutoolRunRecordMailMerge" class="button" title="Run">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
									<i class="icon-RunRecordMailMerge"></i>
									<h6>Run</h6></a>
							</div>
							<div id="mnutoolCloseRecordMailMerge" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-CloseRecordMailMerge"></i>
								<h6>Close</h6></a>
							</div>
						</li>
					</ul>
				</li>

				<%-- Record - Booking - Transfer Booking / Add from waiting list --%>
				<li class="ui-state-default ui-corner-top"><a id="toolbarDelegateBookingTransfer" href="#" rel="Delegate Booking">Delegate Booking</a>
					<ul>
						<li id="mnuSectionSelectDelegateBookingTransfer"><span>Select</span>
							<div id="mnutoolSelectDelegateBookingTransfer" class="button" title="Select">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/checkmark64HOVER.png")%>" alt="" />
									<i class="icon-OK"></i>
									<h6>Select</h6></a>
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
									<i class="icon-previous"></i>
									<h6>First</h6></a>
							</div>
							<div id="mnutoolPreviousDelegateBookingTransfer" class="button">
								<a href="#" rel="table" title="Previous Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/back64HOVER.png") %>" alt="" />
									<i class="icon-backward"></i>
									<h6>Previous</h6></a>
							</div>
							<div id="mnutoolNextDelegateBookingTransfer" class="button">
								<a href="#" rel="table" title="Next Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/next64HOVER.png") %>" alt="" />
									<i class="icon-forward"></i>
									<h6>Next</h6></a>
							</div>
							<div id="mnutoolLastDelegateBookingTransfer" class="button">
								<a href="#" rel="table" title="Last Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/last64HOVER.png") %>" alt="" />
									<i class="icon-next"></i>
									<h6>Last</h6></a>
							</div>
						</li>
					</ul>
				</li>

				<%-- Record - Booking - Bulk Booking --%>

				<li class="ui-state-default ui-corner-top"><a id="toolbarDelegateBookingBulkBooking" href="#" rel="Report_NewEditCopy">Bulk Booking</a>
					<ul>
						<li id="mnuSectionDelegateBookingBulkBooking"><span>Report</span>
							<div id="mnutoolSaveDelegateBookingBulkBooking" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png")%>" alt ="" />
									<i class="icon-disk"></i>
									<h6>Save</h6></a>
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
								<li class="ui-state-default ui-corner-top"><a id="toolbarReportFind" href="#" rel="Report_Find">Find</a>
					<ul>
						<li id="mnuSectionReportFind"><span>Find</span>
							<div id="mnutoolNewReportFind" class="button">
								<a href="#" rel="table" title="New">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt ="" />
																<i class="icon-plus"></i>
																<h6>New</h6></a>
							</div>
							<div id="mnutoolCopyReportFind" class="button">
								<a href="#" rel="table" title="Copy">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
								<i class="icon-copy"></i>
								<h6>Copy</h6></a>
							</div>
							<div id="mnutoolEditReportFind" class="button">
								<a href="#" rel="table" title="Edit">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
								<i class="icon-pencil"></i>
								<h6>Edit</h6></a>
							</div>
							<div id="mnutoolDeleteReportFind" class="button">
								<a href="#" rel="table" title="Delete">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/delete64HOVER.png") %>" alt="" />
								<i class="icon-close"></i>
								<h6>Delete</h6></a>
							</div>
							<div id="mnutoolPropertiesReportFind" class="button">
								<a href="#" rel="table" title="Properties">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png")%>" alt="" />								
								<i class="icon-PropertiesReportFind"></i>
								<h6>Properties</h6></a>
							</div>
							<div id="mnutoolRunReportFind" class="button">
								<a href="#" rel="table" title="Run">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
								<i class="icon-RunReportFind"></i>
								<h6>Run</h6></a>
							</div>
							<div id="mnutoolCloseReportFind" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-CloseReportFind"></i>
								<h6>Close</h6></a>
							</div>
						</li>						
					</ul>
				</li>

								<%-- Report NewEditCopy --%>
								<li class="ui-state-default ui-corner-top"><a id="toolbarReportNewEditCopy" href="#" rel="Report_NewEditCopy">Report</a>
					<ul>
						<li id="mnuSectionNewEditCopyReport"><span>Report</span>
							<div id="mnutoolSaveReport" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png")%>" alt ="" />
																<i class="icon-disk"></i>
																<h6>Save</h6></a>
							</div>
							<div id="mnutoolCancelReport" class="button">
								<a href="#" rel="table" title="Cancel">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cancel64HOVER.png")%>" alt="" />
								<i class="icon-CancelReport"></i>
								<h6>Cancel</h6></a>
							</div>
						</li>
					</ul>
				</li>

								<%-- Report Run --%>
								<li class="ui-state-default ui-corner-top"><a id="toolbarReportRun" href="#" rel="Report_Run">Report</a>
					<ul>
						<li id="mnuSectionRunReport"><span>Output</span>
							<div id="mnutoolOutputReport" class="button">
								<a href="#" rel="table" title="Output">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png")%>" alt ="" />
																<i class="icon-OutputReportRun"></i>
																<h6>Output</h6></a>
							</div>
							<div id="mnutoolCloseReport" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-mnutoolCloseReportRun"></i>
								<h6>Close</h6></a>
							</div>
						</li>						
					</ul>
				</li>

								<%-- Utilities Find --%>
								<li class="ui-state-default ui-corner-top"><a id="toolbarUtilitiesFind" href="#" rel="Utilities_Find">Find</a>
					<ul>
						<li id="mnuSectionUtilitiesFind"><span>Find</span>
							<div id="mnutoolNewUtilitiesFind" class="button">
								<a href="#" rel="table" title="New">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt ="" />
																<i class="icon-plus"></i>
																<h6>New</h6></a>
							</div>
							<div id="mnutoolCopyUtilitiesFind" class="button">
								<a href="#" rel="table" title="Copy">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
								<i class="icon-copy"></i>
								<h6>Copy</h6></a>
							</div>
							<div id="mnutoolEditUtilitiesFind" class="button">
								<a href="#" rel="table" title="Edit">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
								<i class="icon-pencil"></i>
								<h6>Edit</h6></a>
							</div>
							<div id="mnutoolDeleteUtilitiesFind" class="button">
								<a href="#" rel="table" title="Delete">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/delete64HOVER.png") %>" alt="" />
								<i class="icon-close"></i>
								<h6>Delete</h6></a>
							</div>
							<div id="mnutoolPropertiesUtilitiesFind" class="button">
								<a href="#" rel="table" title="Properties">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png")%>" alt="" />								
								<i class="icon-PropertiesUtilitiesFind"></i>
								<h6>Properties</h6></a>
							</div>
							<div id="mnutoolRunUtilitiesFind" class="button">
								<a href="#" rel="table" title="Run">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
								<i class="icon-RunUtilitiesFind"></i>
								<h6>Run</h6></a>
							</div>
							<div id="mnutoolCloseUtilitiesFind" class="button">
								<a href="#" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-CloseUtilitiesFind"></i>
								<h6>Close</h6></a>
							</div>
						</li>						
					</ul>
				</li>

								<%-- Utilities NewEditCopy --%>
								<li class="ui-state-default ui-corner-top"><a id="toolbarUtilitiesNewEditCopy" href="#" rel="Utilities_NewEditCopy">Utilities</a>
					<ul>
						<li id="mnuSectionNewEditCopyUtilities"><span>Utilities</span>
							<div id="mnutoolSaveUtilities" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png")%>" alt ="" />
																<i class="icon-disk"></i>
																<h6>Save</h6></a>
							</div>
							<div id="mnutoolCancelUtilities" class="button">
								<a href="#" rel="table" title="Cancel">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/cancel64HOVER.png")%>" alt="" />
								<i class="icon-CancelUtilities"></i>
								<h6>Cancel</h6></a>
							</div>
						</li>						
					</ul>
				</li>

								<%-- Tools Find --%>
				<li class="ui-state-default ui-corner-top"><a id="toolbarToolsFind" href="#" rel="Tools_Find">Find</a>
					<ul>
						<li id="mnuSectionToolsFind"><span>Find</span>
							<div id="mnutoolNewToolsFind" class="button">
								<a href="#" rel="table" title="New">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt ="" />
																<i class="icon-plus"></i>
																<h6>New</h6></a>
							</div>
							<div id="mnutoolCopyToolsFind" class="button">
								<a href="#" rel="table" title="Copy">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
								<i class="icon-copy"></i>
								<h6>Copy</h6></a>
							</div>
							<div id="mnutoolEditToolsFind" class="button">
								<a href="#" rel="table" title="Edit">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
								<i class="icon-pencil"></i>
								<h6>Edit</h6></a>
							</div>
							<div id="mnutoolDeleteToolsFind" class="button">
								<a href="#" rel="table" title="Delete">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/delete64HOVER.png") %>" alt="" />
								<i class="icon-close"></i>
								<h6>Delete</h6></a>
							</div>
							<div id="mnutoolPropertiesToolsFind" class="button">
								<a href="#" rel="table" title="Properties">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png")%>" alt="" />								
								<i class="icon-PropertiesToolsFind"></i>
								<h6>Properties</h6></a>
							</div>
							<div id="mnutoolRunToolsFind" class="button">
								<a href="#" rel="table" title="Run">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
								<i class="icon-RunToolsFind"></i>
								<h6>Run</h6></a>
							</div>
						</li>						
					</ul>
				</li>

								<%-- EventLog Find--%>
				<li class="ui-state-default ui-corner-top"><a id="toolbarEventLogFind" href="#" rel="Find">Find</a>
					<ul>
						<li><span>Edit</span>											
							<div id="mnutoolViewEventLogFind" class="button">
								<a href="#" rel="table" title="View">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/preview64HOVER.png")%>" alt="" />
									<i class="icon-ViewEventLogFind"></i> 
									<h6>View</h6></a>
							</div>
							<div id="mnutoolPurgeEventLogFind" class="button">
								<a href="#" rel="table" title="Purge">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Purge64HOVER.png")%>"	alt="" />
									<i class="icon-PurgeEventLogFind"></i>
									<h6>Purge</h6></a>
							</div>
							<div id="mnutoolEmailEventLogFind" class="button">
								<a href="#" rel="table" title="Email">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Email64HOVER.png")%>" alt="" />
									<i class="icon-EmailEventLogFind"></i>
									<h6>Email</h6></a>
							</div>
							<div id="mnutoolDeleteEventLogFind" class="button">
								<a href="#" rel="table" title="Delete">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/delete64HOVER.png")%>" alt="" />
									<i class="icon-close"></i>
									<h6>Delete</h6></a>
							</div>
						</li>
						<li><span>Navigate</span>
							<div id="mnutoolFirstEventLogFind" class="button">
								<a href="#" rel="table" title="First Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/first64HOVER.png") %>" alt="" />
									<i class="icon-previous"></i>
									<h6>First</h6></a>
							</div>
							<div id="mnutoolPreviousEventLogFind" class="button">
								<a href="#" rel="table" title="Previous Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/back64HOVER.png") %>" alt="" />
									<i class="icon-backward"></i>
									<h6>Back</h6></a>
							</div>
							<div id="mnutoolNextEventLogFind" class="button">
								<a href="#" rel="table" title="Next Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/next64HOVER.png") %>" alt="" />
									<i class="icon-forward"></i>
									<h6>Next</h6></a>
							</div>
							<div id="mnutoolLastEventLogFind" class="button">
								<a href="#" rel="table" title="Last Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/last64HOVER.png") %>" alt="" />
									<i class="icon-next"></i>
									<h6>Last</h6></a>
							</div>
						</li>
					</ul>
				</li>

				<li class="ui-state-default ui-corner-top"><a id="toolbarEventLogView" href="#" rel="Find">View</a>
					<ul>
						<li><span>View</span>											
							<div id="mnutoolEmailEventLogView" class="button">
								<a href="#" rel="table" title="Email">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Email64HOVER.png")%>" alt="" />
									<i class="icon-EmailEventLogView"></i> 
									<h6>Email</h6></a>
							</div>
							<div id="mnutoolOutputEventLogView" class="button">
								<a href="#" rel="table" title="Output">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png")%>" alt="" />
									<i class="icon-OutputEventLogView"></i> 
									<h6>Email</h6></a>
							</div>
							<div id="mnutoolCloseEventLogView" class="button">
								<a href="#" rel="table" title="View">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
									<i class="icon-CloseEventLogView"></i> 
									<h6>Email</h6></a>
							</div>
						</li>
					</ul>
				</li>


				<li class="ui-state-default ui-corner-top"><a id="toolbarWFPendingStepsFind" href="#" rel="Find">Find</a>
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
								<i class="icon-refresh"></i>
								

								<h6>Run</h6></a>
							</div>
							<div id="mnutoolCloseWFPendingStepsFind" class="button">
								<%--<a href="#" rel="table" title="Close">--%>
									<a href="<%: Url.Action("Main", "Home", New With {.SSIMode = vbFalse})%>" rel="table" title="Close">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<i class="icon-refresh"></i>
								

								<h6>Close</h6>
								</a>
							</div>
						</li>						
					</ul>
				</li>

				<li class="ui-state-default ui-corner-top"><a id="toolbarAdminConfig" href="#" rel="Find">Configure</a>
					<ul>
						<li><span>Configure</span>											
											<div id="mnutoolSaveAdminConfig" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png")%>" alt="" />
								<i class="icon-disk"></i>
								<h6>Save</h6></a>
							</div>
						</li>						
					</ul>
				</li>
			</ul>
		</div>

	</div>
	<div class="FixedLinksRight">
		<ul>
			<li>
				</li>
			</ul>
	</div>

</div>


<%--<				<li class="ui-state-default ui-corner-top"><a id="toolbarRecord" href="#" rel="home">Record</a>
					<ul>
						<li id="mnuEdit"><span>Edit</span>

							<div id="mnutoolNewRecord" class="button">
								<a href="#" rel="table" title="New Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>" alt="" />
									<i class="icon-plus"></i>
									<h6>Add</h6>
							</a></div>
							<div id="mnutoolCopyRecord" class="button">
								<a href="#" rel="table" title="Copy Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>"
										alt="" />
									<i class="icon-copy"></i>
									<h6>Copy</h6></a>
							</div>
							<div id="mnutoolEditRecord" class="button">
								<a href="#" rel="table" title="Edit Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
									<i class="icon-pencil"></i>
									<h6>Edit</h6></a>
							</div>
							<div id="mnutoolSaveRecord" class="button">
								<a href="#" rel="table" title="Save">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png") %>" alt="" />
									<i class="icon-disk"></i>
									<h6>Save</h6></a>
							</div>
							<div id="mnutoolDeleteRecord" class="button">
								<a href="#" rel="table" title="Delete Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/delete64HOVER.png") %>" alt="" />
									<i class="icon-close"></i>
									<h6>Delete</h6></a>
							</div>
						</li>
						<li id="mnuNavigate"><span>Navigate</span>
							<div id="mnutoolParentRecord" class="button">
								<a href="#" rel="table" title="Return to parent record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/upblack64HOVER.png") %>" alt="" />
									<i class="icon-arrow-up"></i>
									<h6>Parent</h6></a>
							</div>
							<div id="mnutoolBack" class="button">
								<a href="#" rel="table" title="Return to record editing">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BackRecord64HOVER.png") %>" alt="" />
									<i class="icon-arrow-left"></i>
									<h6>Back</h6></a>
							</div>

							<div id="mnutoolFirstRecord" class="button">
								<a href="#" rel="table" title="First Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/first64HOVER.png") %>" alt="" />
									<i class="icon-previous"></i>
									<h6>First</h6></a>
							</div>
							<div id="mnutoolPreviousRecord" class="button">
								<a href="#" rel="table" title="Previous Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/back64HOVER.png") %>" alt="" />
									<i class="icon-backward"></i>
									<h6>Previous</h6></a>
							</div>
							<div id="mnutoolNextRecord" class="button">
								<a href="#" rel="table" title="Next Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/next64HOVER.png") %>" alt="" />
									<i class="icon-forward"></i>
									<h6>Next</h6></a>
							</div>
							<div id="mnutoolLastRecord" class="button">
								<a href="#" rel="table" title="Last Record">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/last64HOVER.png") %>" alt="" />
									<i class="icon-next"></i>
									<h6>Last</h6></a>
							</div>
						</li>

						<li id="mnuLocateRecord"><span>Locate Record</span>
							<div id="mnutoolLocateRecords" class="textboxlist">
								<ul><li>Go To<input type="text" /></li></ul>
							</div>
						</li>

						<li id="mnuFind"><span>Find</span>
							<div id="mnutoolFind" class="button">
								<a href="#" rel="table" title="Find">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Find64HOVER.png") %>" alt="" />
									<i class="icon-search"></i>
									<h6>Find</h6></a>
							</div>
							<div id="mnutoolQuickFind" class="button">
								<a href="#" rel="table" title="Quick Find">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/QuickFind64HOVER.png") %>"
										alt="" />
									<i class="icon-check-empty"></i>
									<h6>Quick Find</h6></a>
							</div>
						</li>
						<li id="mnuOrder"><span>Order</span>
							<div id="mnutoolOrder" class="button" title="Change Order">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/sort64HOVER.png") %>" alt="" />
									<i class="icon-sort"></i>
									<h6>Sort</h6></a>
							</div>
							<div id="mnutoolFilter" class="button" title="Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Filtering64HOVER.png") %>" alt="" />
									<i class="icon-filter"></i>
									<h6>Filter</h6></a>
							</div>
							<div id="mnutoolClearFilter" class="button" title="Clear Filter">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/FilteringDelete64HOVER.png") %>"
										alt="" />
									<i class="icon-filter" style="color: gray"></i>
									<h6>Clear Filter</h6></a>
							</div>
							<div id="mnutoolPrint" class="button" title="Print">
								<a href="#" rel="table">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png") %>" alt="" />
									<i class="icon-print"></i>
									<h6>Printer</h6></a>
							</div>
						</li>
						<li id="mnuReports"><span>Reports</span>
							<div id="mnutoolCalendarReportsRec" class="button">
								<a href="#" rel="table" title="Calendar Reports">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/CalendarReports64HOVER.png") %>"
										alt="" />
									<i class="icon-print"></i><h6>Calendar<br />Reports</h6></a>
							</div>
							<div id="mnutoolStdRpt_BreakdownREC" class="button">
								<a href="#" rel="table" title="Absence Breakdown">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/AbsenceBreakdown64HOVER.png") %>"
										alt="" /><i class="icon-print"></i><h6>Absence<br />Breakdown</h6></a>
							</div>
							<div id="mnutoolStdRpt_AbsenceCalendar" class="button">
								<a href="#" rel="table" title="Absence Calendar">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/AbsenceCalendar64HOVER.png") %>"
										alt="" /><i class="icon-print"></i><h6>Absence<br />Calendar</h6></a>
							</div>
							<div id="mnutoolStdRpt_BradfordREC" class="button">
								<a href="#" rel="table" title="Bradford Factor">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BradfordFactor64HOVER.png") %>"
										alt="" /><i class="icon-print"></i><h6>Bradford<br />Factor</h6></a>
							</div>
							<div id="mnutoolMailMergeRec" class="button">
								<a href="#" rel="table" title="Mail Merge">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/MailMerge64HOVER.png") %>"
										alt="" /><i class="icon-print"></i><h6>Mail<br/>Merge</h6></a>
							</div>
						</li>
						<li><span>Training Booking</span>
							<div id="mnutoolBulkBooking" class="button">
								<a href="#" rel="table" title="Bulk Booking">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BulkBooking64HOVER.png")%>"
										alt="" /><i class="icon-print"></i><h6>Bulk<br />Booking</h6></a>
							</div>
							<div id="mnutoolAddFromWaitingList" class="button">
								<a href="#" rel="table" title="Add from Waiting List">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/AddFromWaitingList64HOVER.png")%>"
										alt="" /><i class="icon-print"></i><h6>Add from<br />Waiting List</h6></a>
							</div>
							<div id="mnutoolTransferBooking" class="button">
								<a href="#" rel="table" title="Transfer Booking">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/TransferBooking64HOVER.png") %>"
										alt="" /><i class="icon-print"></i><h6>Transfer<br />Booking</h6></a>
							</div>
							<div id="mnutoolCancelBooking" class="button">
								<a href="#" rel="table" title="Cancel Booking">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/CancelBooking64HOVER.png") %>"
										alt="" /><i class="icon-print"></i><h6>Cancel<br />Booking</h6></a>
							</div>
						</li>
						<li><span>Course Booking</span>
							<div id="mnutoolBookCourse" class="button">
								<a href="#" rel="table" title="Book Course">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/BookCourse64HOVER.png") %>"
										alt="" /><i class="icon-print"></i><h6>Book<br />Course</h6></a>
							</div>
							<div id="mnutoolCancelCourse" class="button">
								<a href="#" rel="table" title="Cancel Course">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/CancelCourse64HOVER.png") %>"
										alt="" /><i class="icon-print"></i><h6>Cancel<br />Course</h6></a>
							</div>
						</li>
						<li style="display: none;"><span>Workflow</span>
							<div id="mnutoolWorkflow" class="button">
								<a href="#" rel="table" title="Workflow">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Workflow64HOVER.png") %>"
										alt="" /><i class="icon-print"></i><h6>Launch<br />Workflow</h6></a>
							</div>
							<div id="mnutoolWorkflowPendingSteps" class="button">
								<a href="#" rel="table" title="Pending Workflow Steps">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/PendingSteps64HOVER.png") %>" alt="" /><i class="icon-print"></i><h6>Pending<br />Workflow Steps</h6></a>
							</div>
						</li>
						<li><span>Record Position</span>
							<div id="mnutoolRecordPosition" class="textboxlist">
								<ul>
									<li>
										<span>Record n of m [(filtered)]</span>
									</li>
								</ul>
							</div>
						</li>
					</ul>
				</li>
				<li class="ui-state-default ui-corner-top"><a id="toolbarUtilities" href="#" rel="utilites">Utilities</a>
					<ul>
						<li id="mnuSectionUtilities"><span>Utilities</span>
							<div id="mnutoolNewUtil" class="button">
								<a href="javascript:setnew();" rel="table" title="Create new...">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png")%>"
										alt="" /><i class="icon-plus"></i><h6>New</h6></a>
							</div>
					
							<div id="mnutoolEditUtil" class="button">
								<a href="javascript:setedit();" rel="table" title="Edit...">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />
								<i class="icon-pencil"></i>
								<h6>Edit</h6></a>
							</div>
							<div id="mnutoolCopyUtil" class="button">
								<a href="javascript:setcopy();" rel="table" title="Copy...">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/copy64HOVER.png")%>" alt="" />
								<i class="icon-copy"></i>
								<h6>Copy</h6></a>
							</div>
							<div id="mnutoolDeleteUtil" class="button">
								<a href="javascript:setdelete();" rel="table" title="Delete...">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/delete64HOVER.png") %>" alt="" />
								<i class="icon-disk"></i>
								<h6>Delete</h6></a>
							</div>
							<div id="mnutoolPrintUtil" class="button">
								<a href="#" rel="table" title="Print...">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png") %>" alt="" />
								<i class="icon-print"></i>
								<h6>Print</h6></a>
							</div>
							<div id="mnutoolPropertiesUtil" class="button">
								<a href="javascript:showproperties();" rel="table" title="Properties...">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png")%>" alt="" />								
								<h6>Properties</h6></a>
							</div>
							<div id="mnutillRunUtil" class="button">
								<a href="javascript:setrun();" rel="table" title="Run...">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Run64HOVER.png")%>" alt="" />
								<h6>Run</h6></a>
							</div>
							<div id="mnuutilCancelUtil" class="button">
								<a href="javascript:setcancel();" rel="table" title="Close...">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/close64HOVER.png")%>" alt="" />
								<h6>Close</h6></a>
							</div>
						</li>						
					</ul>
				</li>
--%>