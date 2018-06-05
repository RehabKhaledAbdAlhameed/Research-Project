namespace finalProjMVC.Models
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class DataModel : DbContext
    {
        public DataModel()
            : base("name=DataModel")
        {
        }

        public virtual DbSet<Company> Companies { get; set; }
        public virtual DbSet<Data_item> Data_item { get; set; }
        public virtual DbSet<Data_type> Data_type { get; set; }
        public virtual DbSet<Historical_Price> Historical_Price { get; set; }
        public virtual DbSet<Industry> Industries { get; set; }
        public virtual DbSet<NonPeriodic_Data_Draft> NonPeriodic_Data_Draft { get; set; }
        public virtual DbSet<NonPeriodic_Data_Live> NonPeriodic_Data_Live { get; set; }
        public virtual DbSet<Periodic_Data_Draft> Periodic_Data_Draft { get; set; }
        public virtual DbSet<Periodic_Date_Live> Periodic_Date_Live { get; set; }
        public virtual DbSet<Report> Reports { get; set; }
        public virtual DbSet<Sector> Sectors { get; set; }
        public virtual DbSet<sysdiagram> sysdiagrams { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Company>()
                .Property(e => e.Comp_Name)
                .IsFixedLength();

            modelBuilder.Entity<Company>()
                .Property(e => e.Reu_Code)
                .IsUnicode(false);

            modelBuilder.Entity<Company>()
                .Property(e => e.Trd_Curr)
                .IsFixedLength();

            modelBuilder.Entity<Company>()
                .Property(e => e.Curr_Out_Shares)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Company>()
                .Property(e => e.IMG_URL)
                .IsUnicode(false);

            modelBuilder.Entity<Company>()
                .Property(e => e.Com_Country)
                .IsUnicode(false);

            modelBuilder.Entity<Company>()
                .HasMany(e => e.Historical_Price)
                .WithRequired(e => e.Company)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Company>()
                .HasMany(e => e.NonPeriodic_Data_Draft)
                .WithRequired(e => e.Company)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Company>()
                .HasMany(e => e.Periodic_Data_Draft)
                .WithRequired(e => e.Company)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Company>()
                .HasMany(e => e.Periodic_Date_Live)
                .WithRequired(e => e.Company)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Company>()
                .HasMany(e => e.Reports)
                .WithRequired(e => e.Company)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Data_item>()
                .Property(e => e.isRecuired)
                .IsUnicode(false);

            modelBuilder.Entity<Data_item>()
                .HasMany(e => e.NonPeriodic_Data_Draft)
                .WithRequired(e => e.Data_item)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Data_item>()
                .HasMany(e => e.Periodic_Data_Draft)
                .WithRequired(e => e.Data_item)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Data_item>()
                .HasMany(e => e.Periodic_Date_Live)
                .WithRequired(e => e.Data_item)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Data_type>()
                .HasMany(e => e.Data_item)
                .WithRequired(e => e.Data_type)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Historical_Price>()
                .Property(e => e.Price_val)
                .HasPrecision(19, 4);

            modelBuilder.Entity<Industry>()
                .Property(e => e.Ind_Name)
                .IsFixedLength();

            modelBuilder.Entity<Industry>()
                .Property(e => e.IMG_URL)
                .IsUnicode(false);

            modelBuilder.Entity<Industry>()
                .HasMany(e => e.Companies)
                .WithRequired(e => e.Industry)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Sector>()
                .Property(e => e.Sec_Name)
                .IsFixedLength();

            modelBuilder.Entity<Sector>()
                .Property(e => e.IMG_URL)
                .IsUnicode(false);

            modelBuilder.Entity<Sector>()
                .Property(e => e.Sec_Desc)
                .IsUnicode(false);

            modelBuilder.Entity<Sector>()
                .HasMany(e => e.Companies)
                .WithRequired(e => e.Sector)
                .WillCascadeOnDelete(false);
        }
    }
}
