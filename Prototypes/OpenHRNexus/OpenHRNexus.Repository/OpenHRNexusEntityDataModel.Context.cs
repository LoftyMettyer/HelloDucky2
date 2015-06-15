﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace OpenHRNexus.Repository
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    using System.Data.Entity.Core.Objects;
    using System.Linq;
    
    public partial class OpenHRNexusEntities : DbContext
    {
        public OpenHRNexusEntities()
            : base("name=OpenHRNexusEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<tbuser_Personnel_Records> tbuser_Personnel_Records { get; set; }
        public virtual DbSet<Personnel_Records> Personnel_Records { get; set; }
        public virtual DbSet<tbuser_Languages> tbuser_Languages { get; set; }
    
        public virtual int spASRIntGetLookupFindRecords2(Nullable<int> piTableID, Nullable<int> piViewID, Nullable<int> piOrderID, Nullable<int> piLookupColumnID, Nullable<int> piRecordsRequired, ObjectParameter pfFirstPage, ObjectParameter pfLastPage, string psLocateValue, ObjectParameter piColumnType, ObjectParameter piColumnSize, ObjectParameter piColumnDecimals, string psAction, ObjectParameter piTotalRecCount, ObjectParameter piFirstRecPos, Nullable<int> piCurrentRecCount, string psFilterValue, Nullable<int> piCallingColumnID, ObjectParameter piLookupColumnGridNumber, Nullable<bool> pfOverrideFilter)
        {
            var piTableIDParameter = piTableID.HasValue ?
                new ObjectParameter("piTableID", piTableID) :
                new ObjectParameter("piTableID", typeof(int));
    
            var piViewIDParameter = piViewID.HasValue ?
                new ObjectParameter("piViewID", piViewID) :
                new ObjectParameter("piViewID", typeof(int));
    
            var piOrderIDParameter = piOrderID.HasValue ?
                new ObjectParameter("piOrderID", piOrderID) :
                new ObjectParameter("piOrderID", typeof(int));
    
            var piLookupColumnIDParameter = piLookupColumnID.HasValue ?
                new ObjectParameter("piLookupColumnID", piLookupColumnID) :
                new ObjectParameter("piLookupColumnID", typeof(int));
    
            var piRecordsRequiredParameter = piRecordsRequired.HasValue ?
                new ObjectParameter("piRecordsRequired", piRecordsRequired) :
                new ObjectParameter("piRecordsRequired", typeof(int));
    
            var psLocateValueParameter = psLocateValue != null ?
                new ObjectParameter("psLocateValue", psLocateValue) :
                new ObjectParameter("psLocateValue", typeof(string));
    
            var psActionParameter = psAction != null ?
                new ObjectParameter("psAction", psAction) :
                new ObjectParameter("psAction", typeof(string));
    
            var piCurrentRecCountParameter = piCurrentRecCount.HasValue ?
                new ObjectParameter("piCurrentRecCount", piCurrentRecCount) :
                new ObjectParameter("piCurrentRecCount", typeof(int));
    
            var psFilterValueParameter = psFilterValue != null ?
                new ObjectParameter("psFilterValue", psFilterValue) :
                new ObjectParameter("psFilterValue", typeof(string));
    
            var piCallingColumnIDParameter = piCallingColumnID.HasValue ?
                new ObjectParameter("piCallingColumnID", piCallingColumnID) :
                new ObjectParameter("piCallingColumnID", typeof(int));
    
            var pfOverrideFilterParameter = pfOverrideFilter.HasValue ?
                new ObjectParameter("pfOverrideFilter", pfOverrideFilter) :
                new ObjectParameter("pfOverrideFilter", typeof(bool));
    
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction("spASRIntGetLookupFindRecords2", piTableIDParameter, piViewIDParameter, piOrderIDParameter, piLookupColumnIDParameter, piRecordsRequiredParameter, pfFirstPage, pfLastPage, psLocateValueParameter, piColumnType, piColumnSize, piColumnDecimals, psActionParameter, piTotalRecCount, piFirstRecPos, piCurrentRecCountParameter, psFilterValueParameter, piCallingColumnIDParameter, piLookupColumnGridNumber, pfOverrideFilterParameter);
        }
    }
}
