/*
Copyright © 2005 - 2013 Annpoint, s.r.o.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

-------------------------------------------------------------------------

NOTE: Reuse requires the following acknowledgement (see also NOTICE):
This product includes DayPilot (http://www.daypilot.org) developed by Annpoint, s.r.o.
*/

using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing; 
using System.Web.UI;
using System.Web.UI.WebControls;
using DayPilot.Web.Ui.Design;
using DayPilot.Web.Ui.Enums;
using DayPilot.Web.Ui.Events;
using DayPilot.Web.Ui.Events.Calendar;

namespace DayPilot.Web.Ui
{
    /// <summary>
    /// DayPilot is a component for showing a day schedule.
    /// </summary>
    [Themeable(true)]
    [ToolboxBitmap(typeof(Calendar))]
    [Designer(typeof(DayPilotCalendarDesigner))]
    public partial class DayPilotCalendar : DataBoundControl, IPostBackEventHandler
    {
        private Day[] _days;

        private string _dataStartField;
        private string _dataEndField;
        private string _dataTextField;
        private string _dataValueField;
				private string _dataTypeField;

        // day header
        private bool _showHeader = true;
        private int _headerHeight = 21;
        private string _headerDateFormat = "d";

        private ArrayList _items;

        /// <summary>
        /// Event called when the user clicks an event in the calendar. It's only called when EventClickHandling is set to PostBack.
        /// </summary>
        [Category("User actions")]
        [Description("Event called when the user clicks an event in the calendar.")]
        public event EventClickEventHandler EventClick;

        /// <summary>
        /// Event called when the user clicks a free space in the calendar. It's only called when TimeRangeSelectedHandling is set to PostBack.
        /// </summary>
        [Category("User actions")]
        [Description("Event called when the user clicks a free space in the calendar.")]
        public event TimeRangeSelectedEventHandler TimeRangeSelected;


        /// <summary>
        /// Use this event to modify event properties before rendering.
        /// </summary>
        [Category("Preprocessing")]
        [Description("Use this event to modify event properties before rendering.")]
        public event BeforeEventRenderEventHandler BeforeEventRender;


        #region Viewstate

        /// <summary>
        /// Loads ViewState.
        /// </summary>
        /// <param name="savedState"></param>
        protected override void LoadViewState(object savedState)
        {
            if (savedState == null)
                return;

            object[] vs = (object[])savedState;

            if (vs.Length != 2)
            {
                throw new ArgumentException("Wrong savedState object.");
            }

            if (vs[0] != null)
            {
                base.LoadViewState(vs[0]);
            }

            if (vs[1] != null)
            {
                _items = (ArrayList)vs[1];
            }

        }

        /// <summary>
        /// Saves ViewState.
        /// </summary>
        /// <returns></returns>
        protected override object SaveViewState()
        {
            object[] vs = new object[2];
            vs[0] = base.SaveViewState();
            vs[1] = _items;

            return vs;
        }

        #endregion

        #region PostBack


        /// <summary>
        /// 
        /// </summary>
        /// <param name="eventArgument"></param>
        public void RaisePostBackEvent(string eventArgument)
        {

            if (eventArgument.StartsWith("PK:"))
            {
                string pk = eventArgument.Substring(3, eventArgument.Length - 3);
                DoEventClick(new EventClickEventArgs(pk));
            }
            else if (eventArgument.StartsWith("TIME:"))
            {
                DateTime time = Convert.ToDateTime(eventArgument.Substring(5, eventArgument.Length - 5));
                DoTimeRangeSelected(new TimeRangeSelectedEventArgs(time, time.AddHours(1), null));
            }
            else
            {
                throw new ArgumentException("Bad argument passed from postback event.");
            }
        }


        private void DoTimeRangeSelected(TimeRangeSelectedEventArgs e)
        {
            if (TimeRangeSelected != null)
            {
                TimeRangeSelected(this, e);
            }
        }

        private void DoEventClick(EventClickEventArgs e)
        {
            if (EventClick != null)
            {
                EventClick(this, e);
            }
        }

        #endregion

        #region Rendering

        /// <summary>
        /// Renders the component HTML code.
        /// </summary>
        /// <param name="output"></param>
        protected override void Render(HtmlTextWriter output)
        {
            LoadEventsToDays();

            // <div>
            if (CssOnly)
            {
                output.AddAttribute("class", PrefixCssClass("_main"));
            }
            output.AddAttribute("width", Width.ToString());
            output.AddStyleAttribute("position", "relative");
            output.RenderBeginTag("div");

            // <table>
            output.AddAttribute("id", ClientID);
            output.AddAttribute("cellpadding", "0");
            output.AddAttribute("cellspacing", "0");
            output.AddAttribute("border", "0");
            if (!CssOnly)
            {
                output.AddStyleAttribute("border-bottom", "1px solid " + ColorTranslator.ToHtml(BorderColor));
                output.AddStyleAttribute("text-align", "left");
            }
            output.RenderBeginTag("table");

            // <tr>
            output.RenderBeginTag("tr");

            // <td>
            output.AddAttribute("valign", "top");
            output.RenderBeginTag("td");

            if (ShowHours)
            {
                RenderHourNamesTable(output);
            }

            // </td>
            output.RenderEndTag();

            // <td>
            output.AddAttribute("width", "100%");
            output.AddAttribute("valign", "top");
            output.RenderBeginTag("td");

            output.AddStyleAttribute("position", "relative");
            output.RenderBeginTag("div");

            RenderEventsAndCells(output);

            output.RenderEndTag();

            // </td>
            output.RenderEndTag();
            // </tr>
            output.RenderEndTag();
            // </table>
            output.RenderEndTag();

            // </div>
            output.RenderEndTag();

        }


        private void RenderHourNamesTable(HtmlTextWriter output)
        {
            output.AddAttribute("cellpadding", "0");
            output.AddAttribute("cellspacing", "0");
            output.AddAttribute("border", "0");
            output.AddAttribute("width", HourWidth.ToString());

            if (!CssOnly)
            {
                output.AddStyleAttribute("border-left", "1px solid " + ColorTranslator.ToHtml(BorderColor));
            }
            output.RenderBeginTag("table");

            if (!CssOnly)
            {
                // <tr> first emtpy
                output.AddStyleAttribute("height", "1px");
                if (!CssOnly)
                {
                    output.AddStyleAttribute("background-color", ColorTranslator.ToHtml(BorderColor));
                }
                output.RenderBeginTag("tr");

                output.RenderBeginTag("td");
                output.RenderEndTag();

                // </tr> first empty
                output.RenderEndTag();
            }

            if (ShowHeader)
            {
                RenderHourHeader(output);
            }

            for (DateTime i = VisibleStart; i < VisibleEnd; i = i.AddHours(1))
            {
                RenderHourTr(output, i);
            }

            // </table>
            output.RenderEndTag();

        }

        private string PrefixCssClass(string name)
        {
            if (String.IsNullOrEmpty(CssClassPrefix))
            {
                return String.Empty;
            }
            return CssClassPrefix + name;
        }


        private void RenderHourTr(HtmlTextWriter output, DateTime i)
        {

            // <tr>
            output.AddStyleAttribute("height", (CellHeight*2) + "px");
            output.RenderBeginTag("tr");

            // <td>
            output.AddAttribute("valign", "bottom");
            if (!CssOnly)
            {
                output.AddStyleAttribute("background-color", ColorTranslator.ToHtml(HourNameBackColor));
                output.AddStyleAttribute("cursor", "default");
            }
            output.RenderBeginTag("td");

            // <div> block
            output.AddStyleAttribute("display", "block");
            output.AddStyleAttribute("height", ((CellHeight * 2) - 1) + "px");
            if (!CssOnly)
            {
                output.AddStyleAttribute("border-bottom", "1px solid " + ColorTranslator.ToHtml(HourNameBorderColor));
                output.AddStyleAttribute("text-align", "right");
            }
            else
            {
                output.AddStyleAttribute("position", "relative");
                output.AddStyleAttribute("height", CellHeight * 2 + "px");
                output.AddAttribute("class", PrefixCssClass("_rowheader"));
            }
            output.RenderBeginTag("div");

            // <div> text
            if (!CssOnly)
            {
                output.AddStyleAttribute("padding", "2px");
                output.AddStyleAttribute("font-family", HourFontFamily);
                output.AddStyleAttribute("font-size", HourFontSize);
            }
            else
            {
                output.AddAttribute("class", PrefixCssClass("_rowheader_inner"));
            }
            output.RenderBeginTag("div");

            int hour = i.Hour;
            bool am = (i.Hour / 12) == 0;
            if (TimeFormat == TimeFormat.Clock12Hours)
            {
                hour = i.Hour % 12;
                if (hour == 0)
                {
                    hour = 12;
                }
            }

            output.Write(hour);
            if (!CssOnly)
            {
                output.AddStyleAttribute("font-size", "10px");
                output.AddStyleAttribute("vertical-align", "super");
                output.RenderBeginTag("span");
                output.Write("&nbsp;");
                //output.Write("<span style='font-size:10px; vertical-align: super; '>&nbsp;");
            }
            else
            {
                output.AddAttribute("class", PrefixCssClass("_rowheader_minutes"));
                output.RenderBeginTag("span");
                //output.Write("<span>");
            }
            if (TimeFormat == TimeFormat.Clock24Hours)
            {
                output.Write("00");
            }
            else
            {
                output.Write(am ? "AM" : "PM");
            }
            output.RenderEndTag();
//            output.Write("</span>");

            output.RenderEndTag();
            output.RenderEndTag();
            output.RenderEndTag(); // </td>
            output.RenderEndTag(); // </tr>
        }


        private void RenderHourHeader(HtmlTextWriter output)
        {

            // <tr>
            output.AddStyleAttribute("height", (HeaderHeight) + "px");
            output.RenderBeginTag("tr");

            // <td>
            output.AddAttribute("valign", "bottom");
            if (!CssOnly)
            {
                output.AddStyleAttribute("background-color", ColorTranslator.ToHtml(HourNameBackColor));
                output.AddStyleAttribute("cursor", "default");
            }
            output.RenderBeginTag("td");

            // <div> block
            output.AddStyleAttribute("display", "block");
            if (!CssOnly)
            {
                output.AddStyleAttribute("border-bottom", "1px solid " + ColorTranslator.ToHtml(BorderColor));
                output.AddStyleAttribute("text-align", "right");
            }
            else
            {
                output.AddStyleAttribute("position", "relative");
                output.AddStyleAttribute("height", HeaderHeight + "px");
                output.AddAttribute("class", PrefixCssClass("_corner"));
            }
            output.RenderBeginTag("div");

            // <div> text
            if (!CssOnly)
            {
                output.AddStyleAttribute("padding", "2px");
                output.AddStyleAttribute("font-size", "6pt");
            }
            else
            {
                output.AddAttribute("class", PrefixCssClass("_corner_inner"));
            }
            output.RenderBeginTag("div");

            output.Write("&nbsp;");

            output.RenderEndTag();
            output.RenderEndTag();
            output.RenderEndTag(); // </td>
            output.RenderEndTag(); // </tr>
        }

        private void RenderEventsAndCells(HtmlTextWriter output)
        {
            RenderCellsTable(output);
            RenderEventsTable(output);
        }

        private void RenderCellsTable(HtmlTextWriter output)
        {
            // <table>
            output.AddAttribute("cellpadding", "0");
            output.AddAttribute("cellspacing", "0");
            output.AddAttribute("border", "0");
            output.AddAttribute("width", "100%");

            output.AddStyleAttribute("position", "absolute");
            output.AddStyleAttribute("left", "0px");
            output.AddStyleAttribute("top", "0px");
            if (!CssOnly)
            {
                output.AddStyleAttribute("border-left", "1px solid " + ColorTranslator.ToHtml(BorderColor));
            }
            output.RenderBeginTag("table");

            /*
            // <tr> first
            output.AddStyleAttribute("height", "1px");
            if (!CssOnly)
            {
                output.AddStyleAttribute("background-color", ColorTranslator.ToHtml(BorderColor));
            }
            output.RenderBeginTag("tr");

            RenderEventTds(output);

            // </tr> first
            output.RenderEndTag();
             * */

            // header
            if (ShowHeader)
            {
                RenderDayHeaders(output);
            }

            output.WriteLine("<!-- empty cells -->");

            // render all cells

            for (DateTime i = VisibleStart; i < VisibleEnd; i = i.AddHours(1))
            {

                // <tr> first half-hour
                output.RenderBeginTag("tr");

                AddHalfHourCells(output, i, true, false);

                // </tr>
                output.RenderEndTag();

                // <tr> second half-hour
                output.AddStyleAttribute("height", CellHeight + "px");
                output.RenderBeginTag("tr");

                bool isLastRow = (i == VisibleEnd.AddHours(-1));
                AddHalfHourCells(output, i, false, isLastRow);

                // </tr>
                output.RenderEndTag();
            }

            // </table>
            output.RenderEndTag();
            
        }

        private void RenderEventsTable(HtmlTextWriter output)
        {
            // <table>
            output.AddAttribute("cellpadding", "0");
            output.AddAttribute("cellspacing", "0");
            output.AddAttribute("border", "0");
            output.AddAttribute("width", "100%");

            output.AddStyleAttribute("position", "absolute");
            output.AddStyleAttribute("left", "0px");
            output.AddStyleAttribute("top", "0px");

            if (!CssOnly)
            {
                output.AddStyleAttribute("border-left", "1px solid " + ColorTranslator.ToHtml(BorderColor));
            }
            output.RenderBeginTag("table");

            // <tr> first
            output.AddStyleAttribute("height", "1px");
            if (!CssOnly)
            {
                output.AddStyleAttribute("background-color", ColorTranslator.ToHtml(BorderColor));
            }
            output.RenderBeginTag("tr");

            RenderEventTds(output);

            // </tr> first
            output.RenderEndTag();

            // </table>
            output.RenderEndTag();

        }

        private void RenderDayHeaders(HtmlTextWriter output)
        {
            if (!CssOnly)
            {
                output.AddStyleAttribute("background-color", ColorTranslator.ToHtml(HourNameBackColor));
            }
            output.AddStyleAttribute("height", HeaderHeight + "px");
            output.RenderBeginTag("tr");

            foreach (Day d in _days)
            {
                DateTime h = new DateTime(d.Start.Year, d.Start.Month, d.Start.Day, 0, 0, 0);

                // <td>
                output.AddAttribute("valign", "bottom");
                if (!CssOnly)
                {
                    output.AddStyleAttribute("background-color", ColorTranslator.ToHtml(HourNameBackColor));
                    output.AddStyleAttribute("cursor", "default");
                    output.AddStyleAttribute("border-right", "1px solid " + ColorTranslator.ToHtml(BorderColor));
                }
                output.RenderBeginTag("td");

                // <div> block
                //output.AddStyleAttribute("display", "block");
                output.AddStyleAttribute("height", HeaderHeight + "px");
                if (!CssOnly)
                {
                    output.AddStyleAttribute("border-bottom", "1px solid " + ColorTranslator.ToHtml(BorderColor));
                    output.AddStyleAttribute("text-align", "center");
                }
                else
                {
                    output.AddStyleAttribute("position", "relative");
                    output.AddAttribute("class", PrefixCssClass("_colheader"));
                }
                output.RenderBeginTag("div");

                // <div> text
                if (!CssOnly)
                {
                    output.AddStyleAttribute("padding", "2px");
                    output.AddStyleAttribute("font-family", DayFontFamily);
                    output.AddStyleAttribute("font-size", DayFontSize);
                }
                else
                {
                    output.AddAttribute("class", PrefixCssClass("_colheader_inner"));
                }
                output.RenderBeginTag("div");

                output.Write(h.ToString(_headerDateFormat));

                output.RenderEndTag();
                output.RenderEndTag();
                output.RenderEndTag(); // </td>


            }

            output.RenderEndTag();
        }



        private void RenderEventTds(HtmlTextWriter output)
        {

            int dayPctWidth = 100 / _days.Length;

            for (int i = 0; i < _days.Length; i++)
            {
                Day d = _days[i];


                // <td>
                output.AddStyleAttribute("height", "1px");
                if (!CssOnly)
                {
                    output.AddStyleAttribute("text-align", "left");
                }
                output.AddAttribute("width", dayPctWidth + "%");
                output.RenderBeginTag("td");

                // <div> position
                output.AddStyleAttribute("display", "block");
                output.AddStyleAttribute("margin-right", ColumnMarginRight + "px"); 
                output.AddStyleAttribute("position", "relative");
                output.AddStyleAttribute("height", "1px");
                if (!CssOnly)
                {
                    output.AddStyleAttribute("font-size", "1px");
                }
                output.AddStyleAttribute("margin-top", "-1px");
                output.RenderBeginTag("div");

                foreach (Event e in d.events)
                {
                    RenderEvent(output, e, d);
                }

                // </div> position
//                output.Write(divPosition.EndTag());
                output.RenderEndTag();

                // </td>
                output.RenderEndTag();
            }
        }

        private void DoBeforeEventRender(BeforeEventRenderEventArgs args)
        {
            if (BeforeEventRender != null)
            {
                BeforeEventRender(this, args);
            }
        }

        private void RenderEvent(HtmlTextWriter output, Event e, Day d)
        {

            string displayText = e.Name + " (" + e.Start.ToShortTimeString() + " - " + e.End.ToShortTimeString() + ")";

            BeforeEventRenderEventArgs ea = new BeforeEventRenderEventArgs(e);
            ea.InnerHTML = displayText;
            ea.ToolTip = displayText;
            ea.EventClickEnabled = EventClickHandling != EventClickHandlingEnum.Disabled;
            if (!CssOnly)
            {
                ea.DurationBarColor = ColorTranslator.ToHtml(DurationBarColor);
                ea.BackgroundColor = ColorTranslator.ToHtml(EventBackColor);
            }

            DoBeforeEventRender(ea);

            // real box dimensions and position
            DateTime dayVisibleStart = new DateTime(d.Start.Year, d.Start.Month, d.Start.Day, VisibleStart.Hour, 0, 0);
            DateTime realBoxStart = e.BoxStart < dayVisibleStart ? dayVisibleStart : e.BoxStart;

            DateTime dayVisibleEnd;
            if (VisibleEnd.Day == 1)
            {
                dayVisibleEnd = new DateTime(d.Start.Year, d.Start.Month, d.Start.Day, VisibleEnd.Hour, 0, 0);
            }
            else if (VisibleEnd.Day == 2)
            {
                dayVisibleEnd = new DateTime(d.Start.Year, d.Start.Month, d.Start.Day, VisibleEnd.Hour, 0, 0).AddDays(1);
            }
            else
            {
                throw new Exception("Unexpected time for dayVisibleEnd.");
            }

            DateTime realBoxEnd = e.BoxEnd > dayVisibleEnd ? dayVisibleEnd : e.BoxEnd;

            // top
            double top = (realBoxStart - dayVisibleStart).TotalHours * (CellHeight * 2) + 1;
            if (ShowHeader)
            {
                top += HeaderHeight;
            }

            // height
            double height = (realBoxEnd - realBoxStart).TotalHours * (CellHeight * 2) - 2;
            int startDelta = (int)Math.Floor((e.Start - realBoxStart).TotalHours * (CellHeight * 2));
            int endDelta = (int)Math.Floor((realBoxEnd - e.End).TotalHours * (CellHeight * 2));

            double barHeight = height - startDelta - endDelta;
            int barTop = startDelta;

            // It's outside of visible area (for NonBusinessHours set to Hide).
            // Don't draw it in that case.
            if (height <= 0)
            {
                return;
            }

            if (CssOnly)
            {
                output.AddStyleAttribute("-moz-user-select", "none"); // prevent text selection in FF
                output.AddStyleAttribute("-khtml-user-select", "none"); // prevent text selection
                output.AddStyleAttribute("-webkit-user-select", "none"); // prevent text selection
                output.AddStyleAttribute("user-select", "none"); // prevent text selection
                output.AddAttribute("unselectable", "on");
                output.AddStyleAttribute("position", "absolute");
                output.AddStyleAttribute("left", e.Column.StartsAtPct + "%");
                output.AddStyleAttribute("top", top + "px");
                output.AddStyleAttribute("width", e.Column.WidthPct + "%");
                output.AddStyleAttribute("height", (realBoxEnd - realBoxStart).TotalHours * (CellHeight * 2) + "px");
                output.AddAttribute("class", PrefixCssClass("_event"));


                if (ea.EventClickEnabled && EventClickHandling != EventClickHandlingEnum.Disabled)
                {
                    if (EventClickHandling == EventClickHandlingEnum.PostBack)
                    {
                        output.AddAttribute("onclick", "javascript:event.cancelBubble=true;" + Page.ClientScript.GetPostBackEventReference(this, "PK:" + e.PK));
                    }
                    else
                    {
                        output.AddAttribute("onclick", "javascript:event.cancelBubble=true;" + String.Format(EventClickJavaScript, e.PK));
                    }
                }

                output.AddAttribute("onmouseover", "this.className+=' " + PrefixCssClass("_event_hover") + "';event.cancelBubble=true;");
                output.AddAttribute("onmouseout", "if (this.className) { this.className = this.className.replace(' " + PrefixCssClass("_event_hover") + "', ''); } ;event.cancelBubble=true;");

                output.RenderBeginTag("div");

                // inner
                output.AddAttribute("class", PrefixCssClass("_event_inner"));
                output.AddAttribute("unselectable", "on");
                output.RenderBeginTag("div");
                output.Write(ea.InnerHTML);
                output.RenderEndTag();

                // bar
                output.AddAttribute("class", PrefixCssClass("_event_bar"));
                output.AddStyleAttribute("position", "absolute");
                output.RenderBeginTag("div");

                double barTopPct = (100.0*barTop/height);
                double barHeightPct = (100.0*barHeight/height);

                if (barTopPct + barHeightPct > 100)
                {
                    barHeightPct = 100 - barTopPct;
                }

                // bar_inner
                output.AddAttribute("class", PrefixCssClass("_event_bar_inner"));
                output.AddStyleAttribute("top", barTopPct + "%");
                output.AddStyleAttribute("height", barHeightPct + "%");
                output.RenderBeginTag("div");

                // bar_inner
                output.RenderEndTag();

                // bar                
                output.RenderEndTag();


                output.RenderEndTag();
            }
            else
            {
                // MAIN BOX
                output.AddAttribute("onselectstart", "return false;"); // prevent text selection in IE

                if (ea.EventClickEnabled && EventClickHandling != EventClickHandlingEnum.Disabled)
                {
                    if (EventClickHandling == EventClickHandlingEnum.PostBack)
                    {
                        output.AddAttribute("onclick", "javascript:event.cancelBubble=true;" + Page.ClientScript.GetPostBackEventReference(this, "PK:" + e.PK));
                    }
                    else
                    {
                        output.AddAttribute("onclick", "javascript:event.cancelBubble=true;" + String.Format(EventClickJavaScript, e.PK));
                    }

                    output.AddStyleAttribute("cursor", "pointer");

                }

                output.AddStyleAttribute("-moz-user-select", "none"); // prevent text selection in FF
                output.AddStyleAttribute("-khtml-user-select", "none"); // prevent text selection
                output.AddStyleAttribute("user-select", "none"); // prevent text selection
                output.AddStyleAttribute("position", "absolute");
                if (!CssOnly)
                {
                    output.AddStyleAttribute("font-family", EventFontFamily);
                    output.AddStyleAttribute("font-size", EventFontSize);
                    output.AddStyleAttribute("white-space", "no-wrap");
                    output.AddStyleAttribute("background-color", ColorTranslator.ToHtml(EventBorderColor));
                }
                output.AddStyleAttribute("left", e.Column.StartsAtPct + "%");
                output.AddStyleAttribute("top", top + "px");
                output.AddStyleAttribute("width", e.Column.WidthPct + "%");
                output.AddStyleAttribute("height", (realBoxEnd - realBoxStart).TotalHours * (CellHeight * 2) + "px");
                output.RenderBeginTag("div");

                // FIX BOX - to fix the outer/inner box differences in Mozilla/IE (to create border)

                if (ea.EventClickEnabled && EventClickHandling != EventClickHandlingEnum.Disabled)
                {
                    if (!CssOnly)
                    {
                        output.AddAttribute("onmouseover", "this.style.backgroundColor='" + ColorTranslator.ToHtml(EventHoverColor) + "';event.cancelBubble=true;");
                        output.AddAttribute("onmouseout", "this.style.backgroundColor='" + ea.BackgroundColor + "';event.cancelBubble=true;");
                    }
                }

                if (ShowToolTip)
                {
                    output.AddAttribute("title", ea.ToolTip);
                }

                output.AddStyleAttribute("margin-top", "1px");
                output.AddStyleAttribute("display", "block");
                output.AddStyleAttribute("height", height + "px");
                if (!CssOnly)
                {
                    output.AddStyleAttribute("background-color", ea.BackgroundColor);
                    output.AddStyleAttribute("border-left", "1px solid " + ColorTranslator.ToHtml(EventBorderColor));
                    output.AddStyleAttribute("border-right", "1px solid " + ColorTranslator.ToHtml(EventBorderColor));
                }
                output.AddStyleAttribute("overflow", "hidden");
                output.RenderBeginTag("div");

                // blue column
                if (e.Start > realBoxStart)
                {

                }

                output.AddStyleAttribute("float", "left");
                output.AddStyleAttribute("width", "5px");
                output.AddStyleAttribute("height", height - startDelta - endDelta + "px");
                output.AddStyleAttribute("margin-top", startDelta + "px");
                if (!CssOnly)
                {
                    output.AddStyleAttribute("background-color", ea.DurationBarColor);
                    output.AddStyleAttribute("font-size", "1px");
                }
                output.RenderBeginTag("div");
                output.RenderEndTag();

                // right border of blue column
                output.AddStyleAttribute("float", "left");
                output.AddStyleAttribute("width", "1px");
                if (!CssOnly)
                {
                    output.AddStyleAttribute("background-color", ColorTranslator.ToHtml(EventBorderColor));
                }
                output.AddStyleAttribute("height", "100%");
                output.RenderBeginTag("div");
                output.RenderEndTag();

                // space
                output.AddStyleAttribute("float", "left");
                output.AddStyleAttribute("width", "2px");
                output.AddStyleAttribute("height", "100%");
                output.RenderBeginTag("div");
                output.RenderEndTag();

                // PADDING BOX
                output.AddStyleAttribute("padding", "1px");
                output.RenderBeginTag("div");

                output.Write(ea.InnerHTML);

                // closing the PADDING BOX
                output.RenderEndTag();

                // closing the FIX BOX
                output.RenderEndTag();

                // closing the MAIN BOX
                output.RenderEndTag();
                
            }

        }

        private void AddHalfHourCells(HtmlTextWriter output, DateTime hour, bool hourStartsHere, bool isLast)
        {
            foreach (Day d in _days)
            {
                DateTime h = new DateTime(d.Start.Year, d.Start.Month, d.Start.Day, hour.Hour, 0, 0);
                AddHalfHourCell(output, h, hourStartsHere, isLast);
            }
        }

        private void AddHalfHourCell(HtmlTextWriter output, DateTime hour, bool hourStartsHere, bool isLast)
        {
            bool isBusiness = true;
            if (hour.Hour < BusinessBeginsHour || hour.Hour >= BusinessEndsHour || hour.DayOfWeek == DayOfWeek.Saturday || hour.DayOfWeek == DayOfWeek.Sunday)
            {
                isBusiness = false;
            }


            string cellBgColor = isBusiness ? ColorTranslator.ToHtml(BackColor) : ColorTranslator.ToHtml(NonBusinessBackColor);

            /*
            if (hour.Hour < BusinessBeginsHour || hour.Hour >= BusinessEndsHour || hour.DayOfWeek == DayOfWeek.Saturday || hour.DayOfWeek == DayOfWeek.Sunday)
            {
                cellBgColor = ColorTranslator.ToHtml(NonBusinessBackColor);
            }
            else
            {
                cellBgColor = ColorTranslator.ToHtml(BackColor);
            }*/

            string borderBottomColor = ColorTranslator.ToHtml(hourStartsHere ? HourHalfBorderColor : HourBorderColor);

            DateTime startingTime = hour;
            if (!hourStartsHere)
            {
                startingTime = hour.AddMinutes(30);
            }

            if (TimeRangeSelectedHandling != TimeRangeSelectedHandling.Disabled)
            {
                if (TimeRangeSelectedHandling == TimeRangeSelectedHandling.PostBack)
                {
                    output.AddAttribute("onclick", "javascript:" + Page.ClientScript.GetPostBackEventReference(this, "TIME:" + startingTime.ToString("s")));
                }
                else
                {
                    output.AddAttribute("onclick", "javascript:" + String.Format(TimeRangeSelectedJavaScript, startingTime.ToString("s")));
                }

                if (!CssOnly)
                {
                    output.AddAttribute("onmouseover", "this.style.backgroundColor='" + ColorTranslator.ToHtml(HoverColor) + "';");
                    output.AddAttribute("onmouseout", "this.style.backgroundColor='" + cellBgColor + "';");
                }

                output.AddStyleAttribute("cursor", "pointer");

            }
            else
            {
                output.AddStyleAttribute("cursor", "default");
            }

            output.AddAttribute("valign", "bottom");
            if (!CssOnly)
            {
                output.AddStyleAttribute("background-color", cellBgColor);
                output.AddStyleAttribute("border-right", "1px solid " + ColorTranslator.ToHtml(BorderColor));
            }
            else
            {
                output.AddStyleAttribute("position", "relative");
            }
            output.AddStyleAttribute("cursor", "hand");
            output.AddStyleAttribute("height", CellHeight + "px");
            output.RenderBeginTag("td");

            if (!CssOnly)
            {
                // FIX BOX - to fix the outer/inner box differences in Mozilla/IE (to create border)
                output.AddStyleAttribute("display", "block");
                output.AddStyleAttribute("height", "14px");
                if (!isLast)
                {
                    output.AddStyleAttribute("border-bottom", "1px solid " + borderBottomColor);
                }
            }
            else
            {
                output.AddStyleAttribute("position", "relative");
                output.AddStyleAttribute("height", CellHeight + "px");

                string className = PrefixCssClass("_cell");
                if (isBusiness)
                {
                    className += " " + PrefixCssClass("_cell_business");
                }
                output.AddAttribute("class", className);


            }
            output.RenderBeginTag("div");

            if (!CssOnly)
            {
                // required
                output.Write("<span style='font-size:1px'>&nbsp;</span>");
            }
            else
            {
                output.AddAttribute("class", PrefixCssClass("_cell_inner"));
                output.RenderBeginTag("div");
                output.RenderEndTag();
            }

            // closing the FIX BOX
            output.RenderEndTag();

            // </td>
            output.RenderEndTag();

        }


        #endregion

        #region Calculations


        /// <summary>
        /// This is only a relative time. The date part should be ignored.
        /// </summary>
        private DateTime VisibleStart
        {
            get
            {

                DateTime date = new DateTime(1900, 1, 1);

                //if (NonBusinessHours == NonBusinessHoursBehavior.Show)
                if (HeightSpec == HeightSpecEnum.Full && !HideFreeCells)
                {
                    return new DateTime(date.Year, date.Month, date.Day, 0, 0, 0);
                }

                DateTime start = new DateTime(date.Year, date.Month, date.Day, BusinessBeginsHour, 0, 0);

                if (HeightSpec == HeightSpecEnum.BusinessHoursNoScroll)
                {
                    return start;
                }

                if (_days == null)
                    return start;

                if (TotalEvents == 0)
                    return start;

                foreach (Day d in _days)
                {
                    DateTime boxStart = new DateTime(date.Year, date.Month, date.Day, d.BoxStart.Hour, d.BoxStart.Minute, d.BoxStart.Second);
                    if (boxStart < start)
                        start = boxStart;
                }

                return new DateTime(start.Year, start.Month, start.Day, start.Hour, 0, 0);


            }
        }

        /// <summary>
        /// This is only a relative time. The date part should be ignored.
        /// </summary>
        private DateTime VisibleEnd
        {
            get
            {
                DateTime date = new DateTime(1900, 1, 1);

                //if (NonBusinessHours == NonBusinessHoursBehavior.Show)
                if (HeightSpec == HeightSpecEnum.Full && !HideFreeCells)
                {
                    return new DateTime(date.Year, date.Month, date.Day, 0, 0, 0).AddDays(1);
                }

                DateTime end;
                if (BusinessEndsHour == 24)
                {
                    end = new DateTime(date.Year, date.Month, date.Day, 0, 0, 0).AddDays(1);
                }
                else
                {
                    end = new DateTime(date.Year, date.Month, date.Day, BusinessEndsHour, 0, 0);
                }

                if (HeightSpec == HeightSpecEnum.BusinessHoursNoScroll)
                    return end;

                if (_days == null)
                    return end;

                if (TotalEvents == 0)
                    return end;

                foreach (Day d in _days)
                {

                    bool addDay = false;
                    if (d.BoxEnd > DateTime.MinValue && d.BoxEnd.AddDays(-1) >= d.Start)
                        addDay = true;

                    DateTime boxEnd = new DateTime(date.Year, date.Month, date.Day, d.BoxEnd.Hour, d.BoxEnd.Minute, d.BoxEnd.Second);

                    if (addDay)
                        boxEnd = boxEnd.AddDays(1);

                    if (boxEnd > end)
                        end = boxEnd;
                }

                if (end.Minute != 0)
                    end = end.AddHours(1);

                return new DateTime(end.Year, end.Month, end.Day, end.Hour, 0, 0);
            }
        }


        private int TotalEvents
        {
            get
            {
                int ti = 0;
                foreach (Day d in _days)
                    ti += d.events.Count;

                return ti;
            }
        }
        #endregion

        #region Data binding


        protected override void PerformSelect()
        {
            // Call OnDataBinding here if bound to a data source using the
            // DataSource property (instead of a DataSourceID), because the
            // databinding statement is evaluated before the call to GetData.       
            if (!IsBoundUsingDataSourceID)
            {
                OnDataBinding(EventArgs.Empty);
            }

            // The GetData method retrieves the DataSourceView object from  
            // the IDataSource associated with the data-bound control.            
            GetData().Select(CreateDataSourceSelectArguments(), OnDataSourceViewSelectCallback);

            // The PerformDataBinding method has completed.
            RequiresDataBinding = false;
            MarkAsDataBound();

            // Raise the DataBound event.
            OnDataBound(EventArgs.Empty);
        }

        private void OnDataSourceViewSelectCallback(IEnumerable retrievedData)
        {
            // Call OnDataBinding only if it has not already been 
            // called in the PerformSelect method.
            if (IsBoundUsingDataSourceID)
            {
                OnDataBinding(EventArgs.Empty);
            }
            // The PerformDataBinding method binds the data in the  
            // retrievedData collection to elements of the data-bound control.
            PerformDataBinding(retrievedData);
        }

        protected override void PerformDataBinding(IEnumerable retrievedData)
        {
            // don't load events in design mode
            if (DesignMode)
            {
                return;
            }

            base.PerformDataBinding(retrievedData);

            if (String.IsNullOrEmpty(DataStartField))
                throw new NullReferenceException("DataStartField property must be specified.");

            if (String.IsNullOrEmpty(DataEndField))
                throw new NullReferenceException("DataEndField property must be specified.");

            if (String.IsNullOrEmpty(DataTextField))
                throw new NullReferenceException("DataTextField property must be specified.");

            if (String.IsNullOrEmpty(DataValueField))
                throw new NullReferenceException("DataValueField property must be specified.");


            // Verify data exists.
            // ReSharper disable ConditionIsAlwaysTrueOrFalse
            if (retrievedData != null)
            // ReSharper restore ConditionIsAlwaysTrueOrFalse
            {
                _items = new ArrayList();

                foreach (object dataItem in retrievedData)
                {

                    DateTime start = Convert.ToDateTime(DataBinder.GetPropertyValue(dataItem, DataStartField, null));
                    DateTime end = Convert.ToDateTime(DataBinder.GetPropertyValue(dataItem, DataEndField, null));
                    string name = Convert.ToString(DataBinder.GetPropertyValue(dataItem, DataTextField, null));
                    string pk = Convert.ToString(DataBinder.GetPropertyValue(dataItem, DataValueField, null));
										string eventType = Convert.ToString(DataBinder.GetPropertyValue(dataItem, DataTypeField, null));

										var ev = new Event(pk, start, end, name, eventType);
                    ev.Source = dataItem;
                    _items.Add(ev);

                }

                _items.Sort(new EventComparer());

            }
        }

        private void LoadEventsToDays()
        {

            if (EndDate < StartDate)
            {
                throw new ArgumentException("EndDate must be equal to or greater than StartDate.");
            }

            int dayCount = (int)(EndDate - StartDate).TotalDays + 1;
            _days = new Day[dayCount];

            for (int i = 0; i < _days.Length; i++)
            {
                _days[i] = new Day(StartDate.AddDays(i));

                if (_items != null)
                {
                    _days[i].Load(_items);
                }
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the start of the business day (in hours).
        /// </summary>
        [Description("Start of the business day (hour from 0 to 23).")]
        [Category("Appearance")]
        [DefaultValue(9)]
        public int BusinessBeginsHour
        {
            get
            {
                if (ViewState["BusinessBeginsHour"] == null)
                    return 9;
                return (int)ViewState["BusinessBeginsHour"];
            }
            set
            {
                if (value < 0)
                    ViewState["BusinessBeginsHour"] = 0;
                else if (value > 23)
                    ViewState["BusinessBeginsHour"] = 23;
                else
                    ViewState["BusinessBeginsHour"] = value;
            }
        }

        /// <summary>
        /// Gets or sets the end of the business day (hours).
        /// </summary>
        [Description("End of the business day (hour from 1 to 24).")]
        [Category("Appearance")]
        [DefaultValue(18)]
        public int BusinessEndsHour
        {
            get
            {
                if (ViewState["BusinessEndsHour"] == null)
                    return 18;
                return (int)ViewState["BusinessEndsHour"];
            }
            set
            {
                if (value < BusinessBeginsHour)
                    ViewState["BusinessEndsHour"] = BusinessBeginsHour + 1;
                else if (value > 24)
                    ViewState["BusinessEndsHour"] = 24;
                else
                    ViewState["BusinessEndsHour"] = value;
            }
        }

        /// <summary>
        /// Gets or sets the height of the time cells in pixels.
        /// </summary>
        [Description("Height of the time cells in pixels.")]
        [Category("Layout")]
        [DefaultValue(20)]
        public int CellHeight
        {
            get
            {
                if (ViewState["CellHeight"] == null)
                    return 20;
                return (int)ViewState["CellHeight"];
            }
            set
            {
                ViewState["CellHeight"] = value;
            }
        }

        /// <summary>
        /// Gets or sets the width of the hour cell in pixels.
        /// </summary>
        [Description("Width of the hour cell in pixels.")]
        [Category("Layout")]
        [DefaultValue(50)]
        public int HourWidth
        {
            get
            {
                if (ViewState["HourWidth"] == null)
                    return 50;
                return (int)ViewState["HourWidth"];
            }
            set
            {
                ViewState["HourWidth"] = value;
            }
        }

        /// <summary>
        /// Gets or sets the Javascript code that is executed when the users clicks on an event. '{0}' will be replaced by the primary key of the event.
        /// </summary>
        [Description("Javascript code that is executed when the users clicks on an event. '{0}' will be replaced by the primary key of the event.")]
        [Category("User actions")]
        [DefaultValue("alert('{0}');")]
        public string EventClickJavaScript
        {
            get
            {
                if (ViewState["EventClickJavaScript"] == null)
                    return "alert('{0}');";
                return (string)ViewState["EventClickJavaScript"];
            }
            set
            {
                ViewState["EventClickJavaScript"] = value;
            }
        }

        /// <summary>
        /// Gets or sets the Javascript code that is executed when the users clicks on a free time slot. '{0}' will be replaced by the starting time of that slot (i.e. '9:00'.
        /// </summary>
        [Description("Javascript code that is executed when the users clicks on a free time slot. '{0}' will be replaced by the starting time of that slot (i.e. '9:00'.")]
        [Category("User actions")]
        [DefaultValue("alert('{0}');")]
        public string TimeRangeSelectedJavaScript
        {
            get
            {
                if (ViewState["TimeRangeSelectedJavaScript"] == null)
                    return "alert('{0}');";
                return (string)ViewState["TimeRangeSelectedJavaScript"];
            }
            set
            {
                ViewState["TimeRangeSelectedJavaScript"] = value;
            }
        }

        /// <summary>
        /// Gets or sets the first day to be shown. Default is DateTime.Today.
        /// </summary>
        [Description("The first day to be shown. Default is DateTime.Today.")]
        public DateTime StartDate
        {
            get
            {
                if (ViewState["StartDate"] == null)
                {
                    return DateTime.Today;
                }

                return (DateTime)ViewState["StartDate"];

            }
            set
            {
                ViewState["StartDate"] = new DateTime(value.Year, value.Month, value.Day);
            }
        }

        /// <summary>
        /// Gets or sets the number of days to be displayed. Default is 1.
        /// </summary>
        [Description("The number of days to be displayed on the calendar. Default value is 1.")]
        [DefaultValue(1)]
        public int Days
        {
            get
            {
                if (ViewState["Days"] == null)
                    return 1;
                return (int)ViewState["Days"];
            }
            set
            {
                int daysCount = value;

                if (daysCount < 1)
                    daysCount = 1;

                ViewState["Days"] = daysCount;
            }
        }

        /// <summary>
        /// Gets the last day to be shown.
        /// </summary>
        public DateTime EndDate
        {
            get
            {
                return StartDate.AddDays(Days - 1);
            }
        }

        /// <summary>
        /// Gets or sets the name of the column that contains the event starting date and time (must be convertible to DateTime).
        /// </summary>
        [Description("The name of the column that contains the event starting date and time (must be convertible to DateTime).")]
        [Category("Data")]
        public string DataStartField
        {
            get
            {
                return _dataStartField;
            }
            set
            {
                _dataStartField = value;

                if (Initialized)
                {
                    OnDataPropertyChanged();
                }

            }
        }

        /// <summary>
        /// Gets or sets the name of the column that contains the event ending date and time (must be convertible to DateTime).
        /// </summary>
        [Description("The name of the column that contains the event ending date and time (must be convertible to DateTime).")]
        [Category("Data")]
        public string DataEndField
        {
            get
            {
                return _dataEndField;
            }
            set
            {
                _dataEndField = value;
                if (Initialized)
                {
                    OnDataPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Gets or sets the  name of the column that contains the name of an event.
        /// </summary>
        [Category("Data")]
        [Description("The name of the column that contains the name of an event.")]
        public string DataTextField
        {
            get
            {
                return _dataTextField;
            }
            set
            {
                _dataTextField = value;

                if (Initialized)
                {
                    OnDataPropertyChanged();
                }

            }
        }

        /// <summary>
        /// Gets or sets the name of the column that contains the primary key. The primary key will be used for rendering the custom JavaScript actions.
        /// </summary>
        [Category("Data")]
        [Description("The name of the column that contains the primary key. The primary key will be used for rendering the custom JavaScript actions.")]
        public string DataValueField
        {
            get
            {
                return _dataValueField;
            }
            set
            {
                _dataValueField = value;

                if (Initialized)
                {
                    OnDataPropertyChanged();
                }

            }
        }

				[Category("Data")]
				[Description("The name of the column that contains the event type. This will be used for rendering the custom JavaScript actions.")]
				public string DataTypeField
				{
					get
					{
						return _dataTypeField;
					}
					set
					{
						_dataTypeField = value;

						if (Initialized)
						{
							OnDataPropertyChanged();
						}

					}
				}

        /// <summary>
        /// Gets or sets whether the hour numbers should be visible.
        /// </summary>
        [Category("Appearance")]
        [Description("Should the hour numbers be visible?")]
        [DefaultValue(true)]
        public bool ShowHours
        {
            get
            {
                if (ViewState["ShowHours"] == null)
                    return true;
                return (bool)ViewState["ShowHours"];
            }
            set
            {
                ViewState["ShowHours"] = value;
            }
        }

        /// <summary>
        /// Gets or sets the time-format for hour numbers (on the left).
        /// </summary>
        [Category("Appearance")]
        [Description("The time-format that will be used for the hour numbers.")]
        [DefaultValue(TimeFormat.Clock12Hours)]
        public TimeFormat TimeFormat
        {
            get
            {
                if (ViewState["TimeFormat"] == null)
                    return TimeFormat.Clock12Hours;
                return (TimeFormat)ViewState["TimeFormat"];
            }
            set
            {
                ViewState["TimeFormat"] = value;
            }
        }

        /// <summary>
        /// Handling of user action (clicking an event).
        /// </summary>
        [Category("User actions")]
        [Description("Whether clicking an event should do a postback or run a javascript action. By default, it calls the javascript code specified in EventClickJavaScript property.")]
        [DefaultValue(EventClickHandlingEnum.JavaScript)]
        public EventClickHandlingEnum EventClickHandling
        {
            get
            {
                if (ViewState["EventClickHandling"] == null)
                    return EventClickHandlingEnum.Disabled;
                return (EventClickHandlingEnum)ViewState["EventClickHandling"];
            }
            set
            {
                ViewState["EventClickHandling"] = value;
            }
        }

        /// <summary>
        /// Handling of user action (clicking a free-time slot).
        /// </summary>
        [Category("User actions")]
        [Description("Whether clicking a free-time slot should do a postback or run a javascript action. By default, it calls the javascript code specified in TimeRangeSelectedJavaScript property.")]
        [DefaultValue(TimeRangeSelectedHandling.Disabled)]
        public TimeRangeSelectedHandling TimeRangeSelectedHandling
        {
            get
            {
                if (ViewState["TimeRangeSelectedHandling"] == null)
                    return TimeRangeSelectedHandling.Disabled;
                return (TimeRangeSelectedHandling)ViewState["TimeRangeSelectedHandling"];
            }
            set
            {
                ViewState["TimeRangeSelectedHandling"] = value;
            }
        }

        //headerDateFormat
        /// <summary>
        /// Gets or sets the format of the date display in the header columns.
        /// </summary>
        [Description("Format of the date display in the header columns.")]
        [Category("Appearance")]
        [DefaultValue("d")]
        public string HeaderDateFormat
        {
            get
            {
                return _headerDateFormat;
            }
            set
            {
                _headerDateFormat = value;
            }
        }

        /// <summary>
        /// Gets or sets whether the header should be visible.
        /// </summary>
        [Category("Appearance")]
        [Description("Should the header be visible?")]
        [DefaultValue(true)]
        public bool ShowHeader
        {
            get
            {
                return _showHeader;
            }
            set
            {
                _showHeader = value;
            }
        }


        /// <summary>
        /// Gets or sets whether the header should be visible.
        /// </summary>
        [Category("Layout")]
        [Description("Header height in pixels.")]
        [DefaultValue(21)]
        public int HeaderHeight
        {
            get
            {
                return _headerHeight;
            }
            set
            {
                _headerHeight = value;
            }
        }

        public override Color BackColor
        {
            get
            {
                if (ViewState["BackColor"] == null)
                    return ColorTranslator.FromHtml("#FFFFD5");
                return (Color)ViewState["BackColor"];
            }
            set
            {
                ViewState["BackColor"] = value;
            }
        }

        public override Color BorderColor
        {
            get
            {
                if (ViewState["BorderColor"] == null)
                    return ColorTranslator.FromHtml("#000000");
                return (Color)ViewState["BorderColor"];
            }
            set
            {
                ViewState["BorderColor"] = value;
            }
        }

        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color HoverColor
        {
            get
            {
                if (ViewState["HoverColor"] == null)
                    return ColorTranslator.FromHtml("#FFED95");
                return (Color)ViewState["HoverColor"];

            }
            set
            {
                ViewState["HoverColor"] = value;
            }
        }

        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color HourBorderColor
        {
            get
            {
                if (ViewState["HourBorderColor"] == null)
                    return ColorTranslator.FromHtml("#EAD098");
                return (Color)ViewState["HourBorderColor"];

            }
            set
            {
                ViewState["HourBorderColor"] = value;
            }
        }

        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color HourHalfBorderColor
        {
            get
            {
                if (ViewState["HourHalfBorderColor"] == null)
                    return ColorTranslator.FromHtml("#F3E4B1");
                return (Color)ViewState["HourHalfBorderColor"];

            }
            set
            {
                ViewState["HourHalfBorderColor"] = value;
            }
        }

        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color HourNameBorderColor
        {
            get
            {
                if (ViewState["HourNameBorderColor"] == null)
                    return ColorTranslator.FromHtml("#ACA899");
                return (Color)ViewState["HourNameBorderColor"];
            }
            set
            {
                ViewState["HourNameBorderColor"] = value;
            }
        }

        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color HourNameBackColor
        {
            get
            {
                if (ViewState["HourNameBackColor"] == null)
                    return ColorTranslator.FromHtml("#ECE9D8");
                return (Color)ViewState["HourNameBackColor"];
            }
            set
            {
                ViewState["HourNameBackColor"] = value;
            }
        }


        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color EventBackColor
        {
            get
            {
                if (ViewState["EventBackColor"] == null)
                    return ColorTranslator.FromHtml("#FFFFFF");
                return (Color)ViewState["EventBackColor"];
            }
            set
            {
                ViewState["EventBackColor"] = value;
            }
        }


        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color EventHoverColor
        {
            get
            {
                if (ViewState["EventHoverColor"] == null)
                    return ColorTranslator.FromHtml("#DCDCDC");
                return (Color)ViewState["EventHoverColor"];
            }
            set
            {
                ViewState["EventHoverColor"] = value;
            }
        }


        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color EventBorderColor
        {
            get
            {
                if (ViewState["EventBorderColor"] == null)
                    return ColorTranslator.FromHtml("#000000");
                return (Color)ViewState["EventBorderColor"];
            }
            set
            {
                ViewState["EventBorderColor"] = value;
            }
        }

        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color DurationBarColor
        {
            get
            {
                if (ViewState["DurationBarColor"] == null)
                    return ColorTranslator.FromHtml("blue");
                return (Color)ViewState["DurationBarColor"];
            }
            set
            {
                ViewState["DurationBarColor"] = value;
            }
        }

        [Category("Appearance")]
        [TypeConverter(typeof(WebColorConverter))]
        public Color NonBusinessBackColor
        {
            get
            {
                if (ViewState["NonBusinessBackColor"] == null)
                    return ColorTranslator.FromHtml("#FFF4BC");

                return (Color)ViewState["NonBusinessBackColor"];
            }
            set
            {
                ViewState["NonBusinessBackColor"] = value;
            }
        }

        [Category("Appearance")]
        public string EventFontFamily
        {
            get
            {
                if (ViewState["EventFontFamily"] == null)
                    return "Tahoma";

                return (string)ViewState["EventFontFamily"];
            }
            set
            {
                ViewState["EventFontFamily"] = value;
            }
        }


        [Category("Appearance")]
        public string HourFontFamily
        {
            get
            {
                if (ViewState["HourFontFamily"] == null)
                    return "Tahoma";

                return (string)ViewState["HourFontFamily"];
            }
            set
            {
                ViewState["HourFontFamily"] = value;
            }
        }

        [Category("Appearance")]
        public string DayFontFamily
        {
            get
            {
                if (ViewState["DayFontFamily"] == null)
                    return "Tahoma";

                return (string)ViewState["DayFontFamily"];
            }
            set
            {
                ViewState["DayFontFamily"] = value;
            }
        }

        [Category("Appearance")]
        public string EventFontSize
        {
            get
            {
                if (ViewState["EventFontSize"] == null)
                    return "8pt";

                return (string)ViewState["EventFontSize"];
            }
            set
            {
                ViewState["EventFontSize"] = value;
            }
        }

        [Category("Appearance")]
        public string HourFontSize
        {
            get
            {
                if (ViewState["HourFontSize"] == null)
                    return "16pt";

                return (string)ViewState["HourFontSize"];
            }
            set
            {
                ViewState["HourFontSize"] = value;
            }
        }

        [Category("Appearance")]
        public string DayFontSize
        {
            get
            {
                if (ViewState["DayFontSize"] == null)
                    return "10pt";

                return (string)ViewState["DayFontSize"];
            }
            set
            {
                ViewState["DayFontSize"] = value;
            }
        }

        /// <summary>
        /// Determines whether the event tooltip is active.
        /// </summary>
        [Description("Determines whether the event tooltip is active.")]
        [Category("Appearance")]
        [DefaultValue(true)]
        public bool ShowToolTip
        {
            get
            {
                if (ViewState["ShowToolTip"] == null)
                    return true;
                return (bool)ViewState["ShowToolTip"];
            }
            set
            {
                ViewState["ShowToolTip"] = value;
            }
        }

        /// <summary>
        /// Width of the right margin inside a column (in pixels).
        /// </summary>
        [Description("Width of the right margin inside a column (in pixels).")]
        [Category("Appearance")]
        [DefaultValue(5)]
        public int ColumnMarginRight
        {
            get
            {
                if (ViewState["ColumnMarginRight"] == null)
                    return 5;
                return (int)ViewState["ColumnMarginRight"];
            }
            set
            {
                ViewState["ColumnMarginRight"] = value;
            }
        }

        /// <summary>
        /// Hide non-business cells if there are no events. Works for HeightSpec="Full"
        /// </summary>
        [Category("Behavior")]
        [Description("Hide non-business cells if there are no events.")]
        [DefaultValue(false)]
        public bool HideFreeCells
        {
            get
            {
                if (ViewState["HideFreeCells"] == null)
                    return false;
                return (bool)ViewState["HideFreeCells"];
            }
            set
            {
                ViewState["HideFreeCells"] = value;
            }
        }

        /// <summary>
        /// Sets or get the way how the height of the scrolling area is determined. It can be either Full (the full height, prevents scrolling), or BusinessHoursNoScroll (it always shows business hours in full).
        /// </summary>
        [DefaultValue(HeightSpecEnum.BusinessHoursNoScroll)]
        [Category("Layout")]
        [Description("Sets or get the way how the height of the scrolling area is determined - Full (the full height, prevents scrolling), or BusinessHoursNoScroll (it always shows business hours in full).")]
        public HeightSpecEnum HeightSpec
        {
            get
            {
                if (ViewState["HeightSpec"] == null)
                    return HeightSpecEnum.BusinessHoursNoScroll;

                return (HeightSpecEnum)ViewState["HeightSpec"];
            }
            set
            {
                ViewState["HeightSpec"] = value;
            }
        }

        [Category("Appearance")]
        public bool CssOnly
        {
            get
            {
                if (ViewState["CssOnly"] == null)
                    return false;

                return (bool)ViewState["CssOnly"];
            }
            set
            {
                ViewState["CssOnly"] = value;
            }
        }

        [Category("Appearance")]
        public string CssClassPrefix
        {
            get
            {
                return (string)ViewState["CssClassPrefix"];
            }
            set
            {
                ViewState["CssClassPrefix"] = value;
            }
        }

        
        
        #endregion


    }
}
