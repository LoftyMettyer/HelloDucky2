using System.Collections.Generic;
using AutoMapper;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class tbuser_LanguagesService : Itbuser_LanguagesService {
		private readonly Itbuser_LanguagesRepository _tbuser_LanguagesRepository;

		public tbuser_LanguagesService(Itbuser_LanguagesRepository tbuser_LanguagesRepository) {
			_tbuser_LanguagesRepository = tbuser_LanguagesRepository;
		}

		public List<tbuser_Languages_Model> List() {
			return Mapper.Map<List<tbuser_Languages_Model>>(_tbuser_LanguagesRepository.List());
		}
	}
}
