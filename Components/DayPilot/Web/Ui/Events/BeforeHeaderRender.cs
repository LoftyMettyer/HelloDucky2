/* Copyright © 2005 - ${year} Annpoint, s.r.o.
   Use of this software is subject to license terms. */

using System;
using System.Collections.Generic;
using DayPilot.Web.Ui.Data;

namespace DayPilot.Web.Ui.Events
{
    /// <summary>
    /// Delegate for BeforeResHeaderRender event.
    /// </summary>
    public delegate void BeforeHeaderRenderEventHandler(object sender, BeforeHeaderRenderEventArgs e);

    /// <summary>
    /// Class that holds event arguments for BeforeResHeaderRender event.
    /// </summary>
    public class BeforeHeaderRenderEventArgs : EventArgs
    {

        /// <summary>
        /// Get or set the column/row header HTML.
        /// </summary>
        public string InnerHTML { get; set; }

        /// <summary>
        /// Get the column/row value (see <see cref="Column.Value">Column.Value</see>, <see cref="Resource.Value">Resource.Value</see>).
        /// </summary>
        public string Value { get; internal set; }

        /// <summary>
        /// Get the column/row name (see <see cref="Column.Name">Column.Name</see>, <see cref="Resource.Name">Resource.Name</see>).
        /// </summary>
        public string Name { get; internal set; }

        /// <summary>
        /// Additional header columns collection.
        /// </summary>
        public List<ResourceColumn> Columns { get; internal set; }

        /// <summary>
        /// DataSource element containing the source data for this resource (Gantt view).
        /// </summary>
        public DataItemWrapper DataItem { get; internal set; }

    }

}
