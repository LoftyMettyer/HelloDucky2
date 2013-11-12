using System;
using DayPilot.Web.Ui;
using DayPilot.Web.Ui.Enums;
using DayPilot.Web.Ui.Events;

public partial class Events : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
    }

    protected void DayPilotCalendar1_TimeRangeSelected(object sender, TimeRangeSelectedEventArgs e)
    {
        Label1.Text = "Time cell starting at " + e.Start + " clicked.";
    }
    
    protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
    {
        switch (DropDownList1.SelectedValue)
        {
            case "client":
                DayPilotCalendar1.TimeRangeSelectedHandling = TimeRangeSelectedHandling.JavaScript;
                DayPilotCalendar1.EventClickHandling = EventClickHandlingEnum.JavaScript;
                break;
            case "server":
                DayPilotCalendar1.TimeRangeSelectedHandling = TimeRangeSelectedHandling.PostBack;
                DayPilotCalendar1.EventClickHandling = EventClickHandlingEnum.PostBack;
                break;
            case "disabled":
                DayPilotCalendar1.TimeRangeSelectedHandling = TimeRangeSelectedHandling.Disabled;
                DayPilotCalendar1.EventClickHandling = EventClickHandlingEnum.Disabled;
                break;
        }
    }
    protected void DayPilotCalendar1_EventClick(object sender, EventClickEventArgs e)
    {
        Label1.Text = "Event with ID " + e.Value + " clicked.";
    }
}
