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

namespace DayPilot.Web.Ui
{
	/// <summary>
	/// Summary description for Event.
	/// </summary>
	[Serializable]
	public class Event
	{

        [NonSerialized]
        internal object Source;

	    /// <summary>
	    /// Event start.
	    /// </summary>
	    public DateTime Start;

		/// <summary>
		/// Event end;
		/// </summary>
		public DateTime End;

		/// <summary>
		/// Event name;
		/// </summary>
		public string Name;

		/// <summary>
		/// Event primary key.
		/// </summary>
		public string PK;

		public string EventType = "";

	    public string Resource;

		/// <summary>
		/// Column to which this event belongs.
		/// </summary>
		[NonSerialized]
		public Column Column;

		/// <summary>
		/// Constructor.
		/// </summary>
		public Event()
		{
		}

		/// <summary>
		/// Constructor that prefills the fields.
		/// </summary>
		/// <param name="pk"></param>
		/// <param name="start"></param>
		/// <param name="end"></param>
		/// <param name="name"></param>
		/// <param name="eventType"></param>
		public Event(string pk, DateTime start, DateTime end, string name, string eventType) : this(pk, start, end, name, null, null, eventType)
		{

		}

		public Event(string pk, DateTime start, DateTime end, string name, string resource, object source, string eventType)
		{
			this.PK = pk;
			this.Start = start;
			this.End = end;
			this.Name = name;
			this.Resource = resource;
			this.EventType = eventType;
			Source = source;
		}

		/// <summary>
		/// Gets the starting time of an event box.
		/// </summary>
		public DateTime BoxStart
		{
			get
			{
				if (Start.Minute >= 30)
					return new DateTime(Start.Year, Start.Month, Start.Day, Start.Hour, 30, 0);
				else
					return new DateTime(Start.Year, Start.Month, Start.Day, Start.Hour, 0, 0);
			}
		}

		/// <summary>
		/// Gets the ending time of an event box.
		/// </summary>
		public DateTime BoxEnd
		{
			get
			{
				if (End.Minute > 30)
				{
					DateTime hourPlus = End.AddHours(1);
					return new DateTime(hourPlus.Year, hourPlus.Month, hourPlus.Day, hourPlus.Hour, 0, 0);
				}
				else if (End.Minute > 0)
				{
					return new DateTime(End.Year, End.Month, End.Day, End.Hour, 30, 0);
				}
				else
				{
					return new DateTime(End.Year, End.Month, End.Day, End.Hour, 0, 0);
				}
			}
		}

		/// <summary>
		/// Returns true if this box overlaps with e's box.
		/// </summary>
		/// <param name="e"></param>
		/// <returns></returns>
		public bool OverlapsWith(Event e)
		{
			return (this.BoxStart < e.BoxEnd && this.BoxEnd > e.Start);
		}

	}
}
