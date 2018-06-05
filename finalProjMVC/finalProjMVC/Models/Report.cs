namespace finalProjMVC.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Report
    {
        public int id { get; set; }

        [Required]
        [StringLength(100)]
        public string Rep_Name { get; set; }

        [Required]
        public string Rep_Location { get; set; }

        [Required]
        [StringLength(100)]
        public string Tag_Sec { get; set; }

        [Required]
        [StringLength(100)]
        public string Tag_Ind { get; set; }

        [Required]
        [StringLength(100)]
        public string Tag_Comp { get; set; }

        [Column(TypeName = "date")]
        public DateTime? Date { get; set; }

        public int comp_id { get; set; }

        public virtual Company Company { get; set; }
    }
}
