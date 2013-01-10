<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%If Session("databaseConnection") Is Nothing Or (Not Session("ErrorText") Is Nothing) Then Return%>
<script type="text/javascript">
	$(document).ready(function () {
		$(".officebar").officebar({});

		$("#toolbarRecord").hide();

		$("#fixedlinks").fadeIn("slow");

	});
</script>
<div class="RecDescPhoto">
	<%If Session("recdesc").ToString.ToUpper.Contains("AVERY") Then%>
	<img src="<%: Url.Action("ShowPhoto", "Home", new with { .ImageName="davery.jpg"}) %>"
		alt="" />
	<%--	<i class="icon-user"></i>--%>
	<%Else%>
	<img src="<%: Url.Action("ShowPhoto", "Home", new with { .ImageName="mworthing.jpg"}) %>"
		alt="" />
	<%End If%>
</div>
<div id="fixedlinks" style="display: none;" class="">
	<div class="RecordDescription">
		<p>
			<%=Session("recdesc")%>
		</p>
	</div>
	<div class="FixedLinksLeft">
	        
			  

            <div class="officebar">
                <ul>
                    <li class="current"><a id="toolbarHome" href="#" rel="home">Home</a>
                        <ul>
                            <li><span>Fixed Links</span>
                                <div class="button">
                                    <a href="<%: Url.Action("LinksMain", "Home") %>" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/home64HOVER.png") %>" alt="" />
													 <i class="icon-home"></i><h6>Home</h6>
												</a>
                                </div>
                                <div class="button">
                                    <a href="<%: Url.Action("LogOff", "Home") %>" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/cross64HOVER.png") %>" alt="" />
													 <i class="icon-off"></i><h6>Log Off</h6>
													 </a>
                                </div>
                                <div class="button">
                                    <a href="<%: Url.Action("Main", "Home") %>" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/abssmall.png") %>" alt="" />
													 <i class="icon-group"></i><h6>Main</h6>
													 </a>
                                </div>
                                <div class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/ChangePassword64HOVER.png") %>"
                                            alt="" />
														  <i class="icon-lock"></i><h6>Change Password</h6>
														  </a>
                                </div>
										 <div class="button">
											 <a href="#" rel="layout">
												 <img src="<%: Url.Content("~/Scripts/officebar/winkit/configuration64HOVER.png") %>" alt="" />
												 <i class="icon-wrench"></i><h6>Layout</h6>
												 </a>
										 </div>
										 <div class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/Workflow64HOVER.png") %>" alt="" />
														  <i class="icon-inbox"></i><h6>Workflow</h6>
														  </a>
                                </div>
                                <div class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/help64HOVER.png") %>" alt="" />
													 <i class="icon-question-sign"></i><h6>Help</h6>
													 </a>
                                </div>
                            </li>
                        </ul>
                    </li>
                    <li><a id="toolbarRecord" href="#" rel="home">Record</a>
                        <ul>
                            <li><span>Edit</span>
                                <div id="mnutoolNewRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/add64HOVER.png") %>" alt="" />Add</a>
                                </div>
                                <div id="mnutoolCopyRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/arrangeall64HOVER.png") %>"
                                            alt="" />Copy</a>
                                </div>
                                <div id="mnutoolEditRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/Edit64HOVER.png") %>" alt="" />Edit</a>
                                </div>
                                <div id="mnutoolSaveRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/diskette64HOVER.png") %>" alt="" />Save</a>
                                </div>
                                <div id="mnutoolDeleteRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/delete64HOVER.png") %>" alt="" />Delete</a>
                                </div>
                                <div id="mnutoolPrint" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/printer64HOVER.png") %>" alt="" />Printer</a>
                                </div>
                            </li>
                            <li><span>Navigate</span>
                                <div id="mnutoolParentRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/upblack64HOVER.png") %>" alt="" />Parent</a>
                                </div>                                
										  <div id="mnutoolBack" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/BackRecord64HOVER.png") %>" alt="" />Back</a>
                                </div>

                                <div id="mnutoolFirstRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/first64HOVER.png") %>" alt="" />First</a>
                                </div>
                                <div id="mnutoolPreviousRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/back64HOVER.png") %>" alt="" />Back</a>
                                </div>
                                <div id="mnutoolNextRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/next64HOVER.png") %>" alt="" />Next</a>
                                </div>
                                <div id="mnutoolLastRecord" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/last64HOVER.png") %>" alt="" />Last</a>
                                </div>
                            </li>

                            <li><span>Locate Record</span>
                                <div id="mnutoolLocateRecords" class="textboxlist">
                                    <ul>
                                        <li>
                                            <%--<img src="<%: Url.Content("~/Scripts/officebar/winkit/lxxxx") %>" alt="" />--%>
                                            Go To <input type="text" />
                                        </li>
                                    </ul>
                                </div>
                            </li>

                            <li><span>Find</span>
                                <div id="mnutoolFind" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/Find64HOVER.png") %>" alt="" />Find</a>
                                </div>
                                <div id="mnutoolQuickFind" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/QuickFind64HOVER.png") %>"
                                            alt="" />Quick Find</a>
                                </div>
                            </li>
                            <li><span>Order</span>
                                <div id="mnutoolOrder" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/sort64HOVER.png") %>" alt="" />Sort</a>
                                </div>
                                <div id="mnutoolFilter" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/Filtering64HOVER.png") %>" alt="" />Filter</a>
                                </div>
                                <div id="mnutoolClearFilter" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/FilteringDelete64HOVER.png") %>"
                                            alt="" />Clear Filter</a>
                                </div>
                            </li>
                            <li><span>Calendar Reports</span>
                                <div id="mnutoolCalendarReportsRec" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/CalendarReports64HOVER.png") %>"
                                            alt="" />Calendar Reports</a>
                                </div>
                                <div id="mnutoolStdRpt_BreakdownREC" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/AbsenceBreakdown64HOVER.png") %>"
                                            alt="" />Absence Breakdown</a>
                                </div>
                                <div id="mnutoolStdRpt_AbsenceCalendar" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/AbsenceCalendar64HOVER.png") %>"
                                            alt="" />Absence Calendar</a>
                                </div>
                                <div id="mnutoolStdRpt_BradfordREC" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/BradfordFactor64HOVER.png") %>"
                                            alt="" />Bradford Factor</a>
                                </div>
                                <div id="mnutoolMailMergeRec" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/MailMerge64HOVER.png") %>"
                                            alt="" />Mail Merge</a>
                                </div>
                            </li>
                            <li><span>Training Booking</span>
                                <div id="mnutoolBulkBooking" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/BulkBooking64HOVER.png") %>"
                                            alt="" />Bulk Booking</a>
                                </div>
                                <div id="mnutoolAddFromWaitingList" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/AddFormWaitingList64HOVER.png") %>"
                                            alt="" />Add From Waiting List</a>
                                </div>
                                <div id="mnutoolTransferBooking" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/TransferBooking64HOVER.png") %>"
                                            alt="" />Transfer Booking</a>
                                </div>
                                <div id="mnutoolCancelBooking" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/CancelBooking64HOVER.png") %>"
                                            alt="" />Cancel Booking</a>
                                </div>
                            </li>
                            <li><span>Course Booking</span>
                                <div id="mnutoolBookCourse" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/BookCourse64HOVER.png") %>"
                                            alt="" />Book</a>
                                </div>
                                <div id="mnutoolCancelCourse" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/CancelCourse64HOVER.png") %>"
                                            alt="" />Cancel</a>
                                </div>
                            </li>
                            <li><span>Workflow</span>
                                <div id="mnutoolWorkflow" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/Workflow64HOVER.png") %>"
                                            alt="" />Workflows</a>
                                </div>
                                <div id="mnutoolWorkflowPendingSteps" class="button">
                                    <a href="#" rel="table">
                                        <img src="<%: Url.Content("~/Scripts/officebar/winkit/PendingSteps64HOVER.png") %>" alt="" />Pending
                                        Steps</a>
                                </div>
                            </li>
                            <li><span>Record Position</span>
                                <div id="mnutoolRecordPosition" class="textboxlist">
                                    <ul>
                                        <li>
	                                        Record n of m [(filtered)]
                                        </li>
                                    </ul>
                                </div>
                            </li>
                        </ul>
                    </li>
                </ul>
            </div>
        </div>
        <div class="FixedLinksRight">
            <ul>
                <li>Layout:
                    <select onchange="try{changeLayout(this.value);}catch (e) {}">
								<option> </option>
                        <option>wireframe</option>
                        <option>winkit</option>
                        <option>tiles</option>
                    </select></li>
               
            </ul>
        </div>
</div>
