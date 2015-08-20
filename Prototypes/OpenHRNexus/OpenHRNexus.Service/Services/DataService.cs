using System;
using System.Collections.Generic;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class DataService : IDataService {
		private IDataRepository _dataRepository;

		public DataService(IDataRepository dataRepository) {
			_dataRepository = dataRepository;
		}

		public IEnumerable<DynamicDataModel> GetData(int id) {
			return _dataRepository.GetData(id);
		}

		public IEnumerable<DynamicDataModel> GetData() {
			return _dataRepository.GetData();
		}

		public IEnumerable<WebFormFields> GetWebFormFields(int id) {
			return _dataRepository.GetWebFormFields(id);
		}

	}
}
