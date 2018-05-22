namespace finalProj.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Periodic_Data_Draft
    {
        public int ID { get; set; }

        [Column(TypeName = "date")]
        public DateTime Date { get; set; }

        public int? Num_of_shares { get; set; }

        public int Comp_id { get; set; }

        public Company Company { get; set; }
    }
}
