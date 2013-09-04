﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%If Session("databaseConnection") Is Nothing Then Return%>
<script type="text/javascript">
	$(document).ready(function () {
		$(".officebar").officebar({});

		//$("#toolbarRecord").hide();
		//$("#mnuSectionUtilities").hide();
		menu_toolbarEnableItem("mnutoolNewUtil", false);
		menu_toolbarEnableItem("mnutoolEditUtil", false);
		menu_toolbarEnableItem("mnutoolCopyUtil", false);
		menu_toolbarEnableItem("mnutoolDeleteUtil", false);
		menu_toolbarEnableItem("mnutoolPrintUtil", false);
		menu_toolbarEnableItem("mnutoolPropertiesUtil", false);
		menu_toolbarEnableItem("mnutillRunUtil", false);

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
			$(".officetab i[class^='icon-']").css("padding-left", "7px");
			$(".officetab i[class^='icon-']").css("padding-bottom", "4px");
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
				<li class="current ui-state-default"><a id="toolbarHome" href="#" rel="home">Home</a>
					<ul>
						<li><span>Fixed Links</span>
							<div class="button"><a href="<%: Url.Action("Main", "Home", New With {.SSIMode = vbTrue})%>" rel="table" title="Self-service Intranet">
								<img src="<%: Url.Content("~/Scripts/officebar/winkit/home64HOVER.png") %>" alt="" />
									<i class="icon-home"></i>
								<h6>SSI</h6></a></div>
							<div class="button">
								<a href="<%: Url.Action("LogOff", "Home") %>" rel="table" title="Log Off">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Logoff64HOVER.png")%>" alt="" />
									<i class="icon-off"></i>
									<h6>Log Off</h6>
								</a>
							</div>
							<div class="button">
								<a href="<%: Url.Action("Main", "Home", New With {.SSIMode = vbFalse})%>" rel="table" title="Data Manager Intranet">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/abssmall.png") %>" alt="" />
									<i class="icon-group"></i>
									<h6>DMI</h6>
								</a>
							</div>
							<div class="button">
								<a href="#" rel="table" title="Change Password">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/ChangePassword64HOVER.png") %>"
										alt="" />
									<i class="icon-lock"></i>
									<h6>Change<br/>Password</h6>
								</a>
							</div>
							<div class="button">
								<a href="javascript:showThemeEditor()" rel="layout" title="Change Layout">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png") %>" alt="" />
									<i class="icon-wrench"></i>
									<h6>Layout</h6>
								</a>
							</div>
							<div class="button">
								<a href="#" rel="table" title="Pending Workflow Steps">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/Workflow64HOVER.png") %>" alt="" />
									<i class="icon-inbox"></i>
									<h6>Workflow</h6>
								</a>
							</div>
							<div class="button">
								<a href="javascript:fixedlinks_mnutoolAboutHRPro()" rel="table" title="About OpenHR v8">
									<img src="<%: Url.Content("~/Scripts/officebar/winkit/info64HOVER.png")%>" alt="" />
									<i class="icon-question-sign"></i>
									<h6>Help</h6>
								</a>
							</div>
						</li>


					</ul>
				</li>
				<li class="ui-state-default ui-corner-top"><a id="toolbarRecord" href="#" rel="home">Record</a>
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
						<li><span>Navigate</span>
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
			</ul>
		</div>
	</div>
	<div class="FixedLinksRight">
		<ul>
			<li></li>
		</ul>
	</div>
</div>
