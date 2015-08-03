using System;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Diagnostics;
using SystemManagerService.Entities;
using SystemManagerService.Enums;
using SystemManagerService.Interfaces;
using SystemManagerService.Messages;

namespace SystemManagerService
{
    public class SecurityManager : DbContext
    {
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
        }

        public IModifyMessage AddRole(string name, string description)
        {
            var result = new PermissionChangeMessage();

            try {
                Group newGroup = new Group() { Name = name, Description = description };
                Groups.Add(newGroup);
                SaveChanges();
                result.status = SaveStatusEnum.Success;

            }
            catch (Exception e)
            {
                Debug.Print(e.InnerException.ToString());
            }

            return result;
        }

        public virtual DbSet<Group> Groups { get; set; }

        public virtual DbSet<PermissionCategory> PermissionCategories { get; set; }

        public virtual DbSet<PermissionFacet> PermissionFacets { get; set; }

        public virtual DbSet<PermissionItem> PermissionItems { get; set; }


    }




}
