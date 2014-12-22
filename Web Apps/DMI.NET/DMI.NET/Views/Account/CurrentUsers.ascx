	<div class="pageTitleDiv" style="margin-bottom: 15px">
		<span class="pageTitle" id="PopupReportDefinition_PageTitle">Current Users</span>
	</div>
	
	<div class="stretchyfill" style="text-align:center">
		<table id="currentLoggedInUsers"></table>	
	</div>

	<br/>

	<div style="text-align:right">
			<input id="btnCancel" name="btnCancel" type="button" class="btn" value="OK"  onclick="currentUsers_cancelClick()" />
	</div>

<script type="text/javascript">

	function currentUsers_cancelClick() {
		$("#divCurrentUsers").dialog("close");
		return false;
	}

	$(document).ready(function () {

		var licence = $.connection['LicenceHub'];

		licence['client'].CurrentUserList = function (userList) {

			$("#currentLoggedInUsers").jqGrid('GridUnload');

			$("#currentLoggedInUsers").jqGrid({
				datatype: 'jsonstring',
				datastr: userList,
				mtype: 'GET',
				jsonReader: {
					root: "rows", //array containing actual data
					page: "page", //current page
					total: "total", //total pages for the query
					records: "records", //total number of records
					repeatitems: false,
					id: "UserName" //index of the column with the PK in it
				},
				colNames: ['User Name', 'Device', 'Area'],
				colModel: [
					{ name: 'UserName', index: 'UserName', width:160 },
					{ name: 'DeviceBrowser', index: 'Device', width: 230 },
					{ name: 'WebAreaName', index: 'WebAreaName', width: 120 }
				],
				autowidth: false,
				width: 560,
				height: 290,
				viewrecords: true,
				sortname: 'User',
				sortorder: "desc",
				rowNum: 10000
			});
		}
	});



</script>


