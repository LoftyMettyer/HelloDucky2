namespace SystemManagerService.Migrations
{
    using System;
    using System.Collections.Generic;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;
    using SystemManagerService.Entities;

    internal sealed class Configuration : DbMigrationsConfiguration<SystemManagerService.SecurityManager>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = true;
            //  AutomaticMigrationDataLossAllowed = true;
        }

        protected override void Seed(SystemManagerService.SecurityManager context)
        {
            //  This method will be called after migrating to the latest version.

            //  You can use the DbSet<T>.AddOrUpdate() helper extension method 
            //  to avoid creating duplicate seed data. E.g.
            //
            //    context.People.AddOrUpdate(
            //      p => p.FullName,
            //      new Person { FullName = "Andrew Peters" },
            //      new Person { FullName = "Brice Lambson" },
            //      new Person { FullName = "Rowan Miller" }
            //    );
            //

            context.PermissionGroups.AddOrUpdate(g => g.Name,
                new PermissionGroup { Name = "Admin" },
                new PermissionGroup { Name = "ProjectManager" },
                new PermissionGroup { Name = "Developer" }
                );

            context.PermissionCategories.AddOrUpdate(g => g.KeyName,
                new PermissionCategory { KeyName = "OpenHR", Description = "User" },
                new PermissionCategory { KeyName = "BatchJobs", Description = "Batch Jobs" },
                new PermissionCategory { KeyName = "CrossTabs", Description = "Cross Tabs" },
                new PermissionCategory { KeyName = "DataTransfer", Description = "Data Transfer" },
                new PermissionCategory { KeyName = "Diary", Description = "Diary" },
                new PermissionCategory { KeyName = "Export", Description = "Export" },
                new PermissionCategory { KeyName = "GlobalAdd", Description = "Global Add" },
                new PermissionCategory { KeyName = "GlobalUpdate", Description = "Global Update" },
                new PermissionCategory { KeyName = "GlobalDelete", Description = "Global Delete" },
                new PermissionCategory { KeyName = "Import", Description = "Import" },
                new PermissionCategory { KeyName = "MailMerge", Description = "Mail Merge" },
                new PermissionCategory { KeyName = "CustomReports", Description = "Custom Reports" },
                new PermissionCategory { KeyName = "StandardReports", Description = "Standard Reports" },
                new PermissionCategory { KeyName = "Filters", Description = "Filters" },
                new PermissionCategory { KeyName = "Picklists", Description = "Picklists" },
                new PermissionCategory { KeyName = "Orders", Description = "Orders" },
                new PermissionCategory { KeyName = "EventLog", Description = "Event Log" },
                new PermissionCategory { KeyName = "EmailQueue", Description = "Email Queue" },
                new PermissionCategory { KeyName = "DataManagerIntranet", Description = "Data Manager Intranet" },
                new PermissionCategory { KeyName = "CMG&Centrefile", Description = "CMG & Centrefile" },
                new PermissionCategory { KeyName = "Calculations", Description = "Calculations" },
                new PermissionCategory { KeyName = "Configuration", Description = "Configuration" },
                new PermissionCategory { KeyName = "MatchReports", Description = "Match Reports" },
                new PermissionCategory { KeyName = "CalendarReports", Description = "Calendar Reports" },
                new PermissionCategory { KeyName = "Envelopes&Labels", Description = "Envelopes & Labels" },
                new PermissionCategory { KeyName = "Envelope&LabelTemplates", Description = "Envelope & Label Templates" },
                new PermissionCategory { KeyName = "RecordProfile", Description = "Record Profile" },
                new PermissionCategory { KeyName = "EmailAddresses", Description = "Email Addresses" },
                new PermissionCategory { KeyName = "EmailGroups", Description = "Email Groups" },
                new PermissionCategory { KeyName = "Menu", Description = "Menu" },
                new PermissionCategory { KeyName = "SuccessionPlanning", Description = "Succession Planning" },
                new PermissionCategory { KeyName = "CareerProgression", Description = "Career Progression" },
                new PermissionCategory { KeyName = "OutlookCalendarQueue", Description = "Outlook Calendar Queue" },
                new PermissionCategory { KeyName = "PayrollTransfer", Description = "Payroll Transfer" },
                new PermissionCategory { KeyName = "Workflow", Description = "Workflow" },
                new PermissionCategory { KeyName = "DocumentTypes", Description = "Document Types" },
                new PermissionCategory { KeyName = "ReportPacks", Description = "Report Packs" },
                new PermissionCategory { KeyName = "9-BoxGridReports", Description = "9-Box Grid Reports" });

            context.PermissionFacets.AddOrUpdate(g => g.Name,
                new PermissionFacet { Name = "New" },
                new PermissionFacet { Name = "User" },
                new PermissionFacet { Name = "Edit" },
                new PermissionFacet { Name = "Run" },
                new PermissionFacet { Name = "Delete" },
                new PermissionFacet { Name = "User" },
                new PermissionFacet { Name = "Administrator" }
                );

        }
    }
}
