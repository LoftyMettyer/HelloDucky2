﻿using Nexus.Common.Classes;
using System;
using System.Collections.Generic;

namespace Nexus.Common.Models
{
    public class WebFormDataModel
    {
        public Guid stepid {get; set;}
        public Dictionary<string, object> data { get; set; }

        public void DataCleanse()
        {

        }
    }
}