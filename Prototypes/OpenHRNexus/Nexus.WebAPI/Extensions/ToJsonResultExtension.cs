using System.Collections.Generic;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using Nexus.Common.Interfaces;

namespace Nexus.WebAPI.Extensions {
	public static class ToJsonResultExtension {

		public static MvcHtmlString ToJsonResult<T>(this IEnumerable<T> items) where T : IJsonSerialize {

			dynamic results = new {
				total = 1,
				page = 1,
				records = 1,
				rows = items
			};

			var jsonSerialiser = new JavaScriptSerializer();
			jsonSerialiser.MaxJsonLength = int.MaxValue;

			dynamic json = HttpUtility.JavaScriptStringEncode(jsonSerialiser.Serialize(results));
			return MvcHtmlString.Create(json);

		}

	}
}
