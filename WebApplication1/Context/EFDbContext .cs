using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Linq;
using System.Web;

namespace WebApplication1.Context
{
    public class EFDbContext :DbContext
    {
        public DbSet<Models.Blog> Blogs { get; set; }
        public DbSet<Models.Comment> Comments { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();  //去除“设置表名为复数”这条约定
        }
    }
}