using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenHRNexus.Common.Models {
	public class tbuser_Languages_Model {
		public int ID { get; set; }
		public Nullable<int> ID_1 { get; set; }
		public Nullable<decimal> Language_Level { get; set; }
		public string Language_Name { get; set; }
		public string Spoken_Fluency { get; set; }
		public string Written_Fluency { get; set; }
		public Nullable<int> updflag { get; set; }
		public string C_description { get; set; }
		public Nullable<bool> C_deleted { get; set; }
		public Nullable<System.DateTime> C_deleteddate { get; set; }
		public byte[] TimeStamp { get; set; }
	}
}
