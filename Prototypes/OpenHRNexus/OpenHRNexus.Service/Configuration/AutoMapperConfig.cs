using AutoMapper;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository;

namespace OpenHRNexus.Service.Configuration {
	public class AutoMapperConfig {
		public static void Configure() {
			Mapper.CreateMap<Personnel_Records, Personnel_Records_Model>();
			Mapper.CreateMap<Personnel_Records_Model, Personnel_Records>();

			Mapper.CreateMap<tbuser_Languages, tbuser_Languages_Model>();
			Mapper.CreateMap<tbuser_Languages_Model, tbuser_Languages>();
		}
	}
}
