namespace WordAddIn1
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Historical_Price
    {
        public int ID { get; set; }

        public int Comp_id { get; set; }

        [Column(TypeName = "date")]
        public DateTime Date { get; set; }

        [Column(TypeName = "money")]
        public decimal Price_val { get; set; }

        public virtual Company Company { get; set; }
    }
}
