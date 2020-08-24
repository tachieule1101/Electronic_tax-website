namespace ReadExcel.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("TienDoCBQL")]
    public partial class TienDoCBQL
    {

        [Required]
        [StringLength(3)]
        public int TT { get; set; }

        [Required]
        [StringLength(50)]
        public string MA { get; set; }

        [Required]
        [StringLength(200)]
        public string CHITIEU { get; set; }

        public int THANG { get; set; }
        public int QUY { get; set; }
        public int KHQUY { get; set; }
        public int TLQUY { get; set; }
        public int NAM { get; set; }
       
        public int KHNAM { get; set; }
        public int TLNAM { get; set; }
    }
}