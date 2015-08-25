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
