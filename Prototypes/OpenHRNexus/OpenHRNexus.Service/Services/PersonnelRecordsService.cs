using System.Collections.Generic;
using AutoMapper;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class PersonnelRecordsService : IPersonnelRecordsService {
		private readonly IPersonnelRecordsRepository _personnelRecordsRepository;

		public PersonnelRecordsService(IPersonnelRecordsRepository personnelRecordsRepository) {
			_personnelRecordsRepository = personnelRecordsRepository;
		}

		public List<Personnel_Records_Model> List() {
			return Mapper.Map<List<Personnel_Records_Model>>(_personnelRecordsRepository.List());
		}
	}
}
