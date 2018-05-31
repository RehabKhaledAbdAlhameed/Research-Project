namespace finalProj.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Periodic_Date_Live
    {
        public int ID { get; set; }

        [Column(TypeName = "date")]
        public DateTime Date { get; set; }

        public string op_value { get; set; }

        public int Comp_id { get; set; }

        public int item_code { get; set; }

        public  Company Company { get; set; }

        public  Data_item Data_item { get; set; }
    }
}
