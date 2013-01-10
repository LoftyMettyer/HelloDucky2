function getLinks() {

	main();

	/******** Our main function ********/
	function main() {
		jQuery(document).ready(function ($) {

			//generate the HTML and bang it into the widget.

			//********** THIS IS AN OPENHR DATABASE VALUES WIDGET *****************
			//Get the database connection object
			var widgetName = "GetLinks";
			var isWidgetLogin = true;
			var widgetUser = "mworthing";
			var widgetPassword = "mworthing";
			var widgetDatabase = "openhr50_std";
			var widgetServer = ".\sql2012";

			$.ajax({
				url: "/dmi.net/account/getWidgetData",
				type: "POST",
				data: { widgetName: widgetName, isWidgetLogin: isWidgetLogin, widgetUser: widgetUser, widgetPassword: widgetPassword, widgetDatabase: widgetDatabase, widgetServer: widgetServer, widgetID: 0 },
				success: function (data) {


					var currentColumn = 1
					var currentRow = 1
					var maxColumn = 50
					var maxRow = 4
					var GroupID = "accordion"
					//First let's drop the LI links on to the page
					$.each(data, function (i, item) {

						var datax = 1
						var datay = 1

						switch (item.element_Type) {
							case 0:
								//Button
								break;
							case 1:
								//Separator
								datax = 2
								datay = 2
								break;
							case 2:
								//Chart
								break;
							case 3:
								//Pending Workflow Steps
								break;
							case 4:
								//Database Value
								break;
							case 5:
								//Today's events
								break;
							case 6:
								//uh oh.
								break;
						}

						if ((item.element_Type < 6)) {
							//is the next widget higher than the remaining room on the column?
							var remainingRows = maxRow - (currentRow - 1);
							if (datay > remainingRows) { currentRow = 1; currentColumn += 1; }

//							if (item.element_Type == 1) {
//								//start a new group.
//								GroupID = item.text;
//								//Call function on portal...
//								CreateGroup(GroupID);
//							}

							if (item.element_Type != 1) {
								addWidget(GroupID, "newWidget" + item.ID, currentColumn, currentRow, datax, datay);

							//calculate next tile position:
							currentRow += datay
							if (currentRow > maxRow) {
								currentRow = 1
								currentColumn += 1
							}
						}
					}



					});


					// calculate grid size.
					var x_cols = Math.ceil(data.length / 4) + 16;
					x_cols = 3;
					applyGridster(x_cols);


					//loop through data and call relevant .js file with parameters
					$.each(data, function (i, item) {

						switch (item.element_Type) {
							case 0:
								//Button
								addWidgetScript(item.ID, "wdg_oHRButton.js", item.text, "ben Tile lightBlueTile foundicon-people");
								break;
							case 1:
								//Separator
								addWidgetScript(item.ID, "wdg_oHRButton.js", item.text, "ben Tile greenTile foundicon-settings");
								break;
							case 2:
								//Chart
								addWidgetScript(item.ID, "wdg_oHRButton.js", item.text, "ben Tile BlueTile foundicon-graph");
								break;
							case 3:
								//Pending Workflow Steps
								addWidgetScript(item.ID, "wdg_oHRButton.js", item.text, "ben Tile lightBlueTile foundicon-edit");
								break;
							case 4:
								//Database Value
								//if (i < 6) {
								addWidgetScript(item.ID, "wdg_oHRDBV.js", item.text, "ben Tile redTile");
								//}
								break;
							case 5:
								//Today's events
								addWidgetScript(item.ID, "wdg_oHRButton.js", item.text, "ben Tile orangeTile foundicon-calendar");
								break;
							case 6:
								//uh oh.
								break;
						}

					});


				},
				error: function (req, status, errorObj) { alert("not ok"); }
			});
		});
	}
}

