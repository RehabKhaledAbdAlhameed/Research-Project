namespace WordAddIn1
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Sector
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Sector()
        {
            Companies = new HashSet<Company>();
        }

        [Key]
        public int Sec_id { get; set; }

        [Required]
        [StringLength(100)]
        public string Sec_Name { get; set; }

        [StringLength(100)]
        public string IMG_URL { get; set; }

        [StringLength(200)]
        public string Sec_Desc { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Company> Companies { get; set; }
    }
}
