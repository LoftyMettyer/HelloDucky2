using System;
using System.Collections.Generic;
using System.Web.Http;
using Nexus.Common.Models;
using System.Security.Claims;
using Microsoft.AspNet.Identity;
using Nexus.Common.Classes;
using Nexus.Common.Classes.DataFilters;
using System.Web;
using Nexus.Common.Interfaces.Services;
using System.Collections;
using System.Threading.Tasks;
using Nexus.Common.Interfaces;
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

    }
}
