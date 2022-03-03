using FluentMigrator;

namespace TesteBackendEnContact.Database
{
    [Migration(1)]
    public class DatabaseMigration : Migration
    {
        public override void Up()
        {
            Create.Table("ContactBook")
                .WithColumn("Id").AsInt32().PrimaryKey().Identity()
                .WithColumn("Name").AsString(50).NotNullable()
            ;

            Create.Table("Contact")
                .WithColumn("Id").AsInt32().PrimaryKey().Identity()
                .WithColumn("ContactBookId").AsInt32().NotNullable()
                .WithColumn("CompanyId").AsInt32().Nullable()
                .WithColumn("Name").AsString(50).NotNullable()
                .WithColumn("Phone").AsString(20).Nullable()
                .WithColumn("Email").AsString(50).Nullable()
                .WithColumn("Address").AsString(100).Nullable()
            ;

            Create.Table("Company")
                .WithColumn("Id").AsInt32().PrimaryKey().Identity()
                .WithColumn("ContactBookId").AsInt32().NotNullable()
                .WithColumn("Name").AsString(50).NotNullable()
            ;
        }

        public override void Down()
        {
            Delete.Table("Company");
            Delete.Table("Contact");
            Delete.Table("ContactBook");
        }
    }
}
