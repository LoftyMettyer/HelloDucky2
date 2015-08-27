﻿using System;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Service.Interfaces;

namespace Nexus.Service.Services {
	public class DataService : IDataService {
		private IDataRepository _dataRepository;

		public DataService(IDataRepository dataRepository) {
			_dataRepository = dataRepository;
		}

		//public IEnumerable<DynamicDataModel> GetData(int id) {
		//	return _dataRepository.GetData(id);
		//}

		//public IEnumerable<DynamicDataModel> GetData() {
		//	return _dataRepository.GetData();
		//}

        WebFormModel IDataService.GetWebForm(int businessProcessId, Guid userId)
        {

            // Move to a factory for flexibility and eaiser reading?
            var businessProcess = _dataRepository.GetBusinessProcess(businessProcessId);

            //var businessProcess = (BusinessProcessModel)_dataRepository.GetBusinessProcess(businessProcessId);

            //   BusinessProcessModel businessProcess2 = (BusinessProcessModel)_dataRepository.GetBusinessProcess(businessProcessId);
            //    var model = new BusinessProcessModel(businessProcess);
            //            var businessProcess = new BusinessProcessModel(_dataRepository);

            //, businessProcessId);
       //     var webForm = businessProcess.GetFirstStep();
           


         //   firstStep.Translate("en-GB");


            WebForm webForm = _dataRepository.GetWebForm(businessProcessId);
            webForm.Translate("en-GB");


            var result = _dataRepository.PopulateFormWithData(webForm, userId);
            //var result = new WebFormModel();



            // Implement translation as a design pattern (a template one? - I can't remember - need to review training notes)
            //result.translate(language)

            return result;
        }

    }
}
