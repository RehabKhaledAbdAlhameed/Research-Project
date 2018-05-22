namespace finalProj.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Industry")]
    public partial class Industry
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Industry()
        {
            Companies = new HashSet<Company>();
        }

        [Key]
        public int Ind_id { get; set; }

        [Required]
        [StringLength(100)]
        public string Ind_Name { get; set; }

        public int Sec_id { get; set; }

        [StringLength(100)]
        public string IMG_URL { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public ICollection<Company> Companies { get; set; }
    }
}
