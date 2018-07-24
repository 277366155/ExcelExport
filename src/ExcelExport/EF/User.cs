using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace ExcelExport.EF
{
    public class ViewModel
    {
    }

    public class User: ViewModel
    {
        [DisplayName("ID")]
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        [DisplayName("姓名")]
        [MaxLength(50)]
        public string Name { get; set; }

        [DisplayName("性别")]
        public int Sex { get; set; }

        [DisplayName("年龄")]
        public int Age { get; set; }
    }
}