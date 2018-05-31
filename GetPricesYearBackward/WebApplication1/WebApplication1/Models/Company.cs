namespace WebApplication1.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Company")]
    public partial class Company
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Company()
        {
            Historical_Price = new HashSet<Historical_Price>();
            Periodic_Data_Draft = new HashSet<Periodic_Data_Draft>();
            Periodic_Date_Live = new HashSet<Periodic_Date_Live>();
        }

        [Key]
        public int Comp_id { get; set; }

        [Required]
        [StringLength(100)]
        public string Comp_Name { get; set; }

        public double Reu_Code { get; set; }

        [StringLength(10)]
        public string Trd_Curr { get; set; }

        [Column(TypeName = "money")]
        public decimal? Curr_Out_Shares { get; set; }

        [StringLength(100)]
        public string IMG_URL { get; set; }

        [StringLength(100)]
        public string Com_Country { get; set; }

        public int Sec_id { get; set; }

        public int Ind_id { get; set; }

        public virtual Industry Industry { get; set; }

        public virtual Sector Sector { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Historical_Price> Historical_Price { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Periodic_Data_Draft> Periodic_Data_Draft { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Periodic_Date_Live> Periodic_Date_Live { get; set; }
    }
}
