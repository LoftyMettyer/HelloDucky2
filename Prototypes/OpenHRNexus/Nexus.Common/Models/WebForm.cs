using System.Collections.Generic;

namespace Nexus.Common.Models {
	public class WebForm //: IProcessRepository
	{
		public int id { get; set; }
		public string Name { get; set; }
		public List<WebFormField> Fields { get; set; }

		//public string GetBaseTableInForm()
		//{
		//    throw new NotImplementedException();
		//}

		//public string GetColumnsInForm()
		//{
		//    throw new NotImplementedException();
		//}

	}

}
