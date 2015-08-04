using System;
using System.Data.Entity;
using System.Linq;
using SystemManagerService.Entities;
using SystemManagerService.Enums;
using SystemManagerService.Interfaces;
using SystemManagerService.Messages;

namespace SystemManagerService
{
    public class SecurityManager : SecurityManagerContext
    {

        public IModifyMessage AddPermissionGroup(string name, string description)
        {
            var result = new PermissionChangeMessage();

            try {
                PermissionGroup newGroup = new PermissionGroup() { Name = name, Description = description };
                PermissionGroups.Add(newGroup);
                SaveChanges();
                result.ModifiedId = newGroup.Id;
                result.status = SaveStatusEnum.Success;

            }
            catch (Exception e)
            {
                result.status = SaveStatusEnum.Failure;
            }

            return result;
        }

        public IModifyMessage AddPermissionToGroup(int groupId, string category, string facet)
        {
            var result = new PermissionChangeMessage();

            try
            {
                var group = PermissionGroups
                 .Where(g => g.Id == groupId)
                 .First();

                var newPermission = new PermissionItem() { Category = category, Facet = facet };
                group.PermissionItems.Add(newPermission);

                SaveChanges();
                result.status = SaveStatusEnum.Success;

            }
            catch (Exception e)
            {
                result.status = SaveStatusEnum.Failure;
            }

            return result;
        }




        public virtual DbSet<PermissionGroup> PermissionGroups { get; set; }

        public virtual DbSet<PermissionCategory> PermissionCategories { get; set; }

        public virtual DbSet<PermissionFacet> PermissionFacets { get; set; }

//        public virtual DbSet<PermissionItem> PermissionItems { get; set; }


    }




}
