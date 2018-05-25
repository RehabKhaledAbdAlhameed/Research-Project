namespace ExcelProject
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Data_item
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Data_item()
        {
            NonPeriodic_Data_Draft = new HashSet<NonPeriodic_Data_Draft>();
            Periodic_Data_Draft = new HashSet<Periodic_Data_Draft>();
            Periodic_Date_Live = new HashSet<Periodic_Date_Live>();
        }

        [Key]
        public int item_code { get; set; }

        public string item_desc { get; set; }

        public int type_id { get; set; }

        [StringLength(50)]
        public string formate { get; set; }

        [Required]
        [StringLength(1)]
        public string isRecuired { get; set; }

        public virtual Data_type Data_type { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<NonPeriodic_Data_Draft> NonPeriodic_Data_Draft { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Periodic_Data_Draft> Periodic_Data_Draft { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Periodic_Date_Live> Periodic_Date_Live { get; set; }
    }
}
