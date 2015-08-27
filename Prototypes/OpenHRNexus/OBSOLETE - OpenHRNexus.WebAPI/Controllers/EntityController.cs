using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.Common.Enums;

namespace OpenHRNexus.WebAPI.Controllers
{
    [Authorize(Roles = "OpenHRUser")]
    public class EntityController : ApiController
    {
        private readonly IEntityService _entityService;
        
        public EntityController(IEntityService entityService)
        {
            _entityService = entityService;
        }

        [HttpGet]
        public IEnumerable<EntityModel> GetEntities(EntityType? entityType)
        {
            return _entityService.GetEntities(entityType).ToList();
        }

    }
}
