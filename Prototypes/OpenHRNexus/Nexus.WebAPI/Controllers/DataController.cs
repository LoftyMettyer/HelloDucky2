using System;
using System.Collections.Generic;
using System.Web.Http;
using Nexus.Common.Models;
using System.Security.Claims;
using Microsoft.AspNet.Identity;
using Nexus.Common.Classes.DataFilters;
using System.Web;
using Nexus.Common.Interfaces.Services;
using System.IO;
using System.Threading.Tasks;
using Nexus.WebAPI.Formatters;
using System.Reflection;

namespace Nexus.WebAPI.Controllers
{
    //	[Authorize(Roles = "OpenHRUser")]
    public class DataController : ApiController
    {
        private readonly IDataService _dataService;
        //private readonly IWorkflowService _workflowService;

        private ClaimsIdentity _identity;
        private string _language;

        public DataController()
        {
        }

        public DataController(IDataService dataService)
        {
            _dataService = dataService;
            _identity = User.Identity as ClaimsIdentity;
            _language = "en-gb";
            if (HttpContext.Current.Request.UserLanguages != null)
            {
                _language = HttpContext.Current.Request.UserLanguages[0].ToLowerInvariant().Trim();
            }
        }

        public DataController(IDataService dataService, ClaimsIdentity claims, string language)
        {
            _identity = claims;
            _dataService = dataService;
            _language = language;
        }

		/// <summary>
		/// Get information about the system, such as version number, build date, etc.
		/// </summary>
		/// <returns></returns>
		[HttpGet]
		public Dictionary<string,string> GetSystemInfo()
		{
			var apiVersion = Assembly.GetExecutingAssembly().GetName().Version;
			var systemInfo = new Dictionary<string, string>();

			systemInfo.Add("API Version", String.Format("{0}.{1:0}.{2}", apiVersion.Major, apiVersion.Minor, apiVersion.Build));
			systemInfo.Add("API build timestamp", RetrieveLinkerTimestamp().ToString());

			return systemInfo;
		}

        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public IEnumerable<CalendarEventModel> GetCalendarData(string calendarType, DateTime from, DateTime to)
        {

            var userId = new Guid(_identity.GetUserId());

            var filters = new List<DateRangeFilter>();
            filters.Add(new DateRangeFilter() { StartRange = from, EndRange = to });

            return _dataService.GetReportData(1, filters);

        }

        /// <summary>
        /// Maybe a temporary stub for the Angular donut control to work.
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public IEnumerable<SummaryDataModel> GetSummaryData()
        {
            var userId = new Guid(_identity.GetUserId());
            return _dataService.GetSummaryData(userId, 1, null);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataSourceId"></param>
        /// <param name="filters"></param>
        /// <returns></returns>
        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public async Task<GridRequestFormat> GetData(int dataSourceId)
        {

            var filters = new List<RangeFilter>()
            {
                new RangeFilter() {
                    RecordRange = 100
                    }
            };

            var userId = new Guid(_identity.GetUserId());

            var dataDescription = new List<ColumnDefinitionFormat>();

            var definitionType = _dataService.GetDataDefinition(dataSourceId);

            var typeAsColModel = ConvertTypeToColModel(definitionType);

            var data = await _dataService.GetData(dataSourceId, filters);

            return new GridRequestFormat()
            {
                total = 1, page = 1, records = 0, rows = data, colModel = typeAsColModel
            };

        }

        private List<ColumnDefinitionFormat> ConvertTypeToColModel(Type description)
        {
            var colModel = new List<ColumnDefinitionFormat>();
            var fieldCount = 0;

            foreach (var field in description.GetRuntimeFields())
            {
                fieldCount += 1;
                colModel.Add(new ColumnDefinitionFormat() { sortable = true, name = field.Name, index = field.Name});                   
            }

            return colModel;

        }

	#region Private methods
	//Get compilation datetime from assembly; as you can see from the code below, it's not a trivial thing to do!
	private DateTime RetrieveLinkerTimestamp()
	{
	  string filePath = Assembly.GetCallingAssembly().Location;
	  const int PEHeaderOffset = 60;
	  const int LinkerTimestampOffset = 8;
	  byte[] b = new byte[2047];
	  Stream s = null;

	  try
	  {
		s = new FileStream(filePath, FileMode.Open, FileAccess.Read);
		s.Read(b, 0, 2047);
	  }
	  finally
	  {
		  s?.Close();
	  }

		int i = BitConverter.ToInt32(b, PEHeaderOffset);
	  int secondsSince1970 = BitConverter.ToInt32(b, i + LinkerTimestampOffset);
	  var dt = new DateTime(1970, 1, 1, 1, 0, 0);
	  dt = dt.AddSeconds(secondsSince1970);

	  return dt;
	}
	#endregion
  }
}
