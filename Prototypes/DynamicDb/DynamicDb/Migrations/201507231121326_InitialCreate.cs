namespace DynamicDb.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialCreate : DbMigration
    {
        public override void Up()
        {
           // CreateTable(
           //"ddo.DynamicTemplate",
           //c => new
           //{
           //    CourseID = c.Int(nullable: false),
           //    Title = c.String(),
           //    Credits = c.Int(nullable: false),
           //})
           //.PrimaryKey(t => t.CourseID);


        }
        
        public override void Down()
        {
        }
    }
}
