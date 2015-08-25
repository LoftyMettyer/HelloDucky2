using System;
using System.Collections.Generic;
using Nexus.Common.Models;
using Nexus.Repository.Interfaces;
using Nexus.Service.Interfaces;

namespace Nexus.Service.Services {
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

        WebFormModel IDataService.GetWebForm(int id, Guid userId)
        {
            WebForm webForm = _dataRepository.GetWebForm(id);

            var result = _dataRepository.PopulateFormWithData(webForm, userId);




            // Implement translation as a design pattern (a template one? - I can't remember - need to review training notes)
            //result.translate(language)

            return result;
        }

    }
}
