﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SystemManagerService.Entities
{
    public class PermissionItem
    {
        [Index]
        public int Id { get; set; }

        //[Column("PermissionCategoryId")]
        //public int CategoryId { get; set; }

        //[Column("PermissionFacetId")]
        //public int FacetId { get; set; }

        public string Category { get; set; }

        public string Facet { get; set; }

    }

}
