namespace WordAddIn1
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class NonPeriodic_Data_Live
    {
        public int ID { get; set; }

        public int item_code { get; set; }

        public string op_value { get; set; }

        [Column(TypeName = "date")]
        public DateTime? op_date { get; set; }

        public int? comp_id { get; set; }
    }
}
