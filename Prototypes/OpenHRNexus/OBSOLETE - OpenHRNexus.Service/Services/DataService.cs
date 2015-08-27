using System;
using System.Collections.Generic;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class ProcessService : IBusinessProcessService {
		private IDataRepository _dataRepository;

		public ProcessService(IDataRepository dataRepository) {
			_dataRepository = dataRepository;
		}

		public IEnumerable<DynamicDataModel> GetData(int id) {
			return _dataRepository.GetData(id);
		}

		public IEnumerable<DynamicDataModel> GetData() {
			return _dataRepository.GetData();
		}

        WebFormModel IBusinessProcessService.GetWebFormForProcessAndUser(int id, Guid userId)
        {

            //BusinessProcess = _dataRepository

            WebForm webForm = _dataRepository.GetWebForm(id);

            var result = _dataRepository.PopulateFormWithData(webForm, userId);



            // Implement translation as a design pattern (a template one? - I can't remember - need to review training notes)
            //result.translate(language)

            return result;
        }

    }
}
