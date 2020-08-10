using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WebReport.Class;
using System.ComponentModel.DataAnnotations;
namespace WebReport.Models
{
    public class Importsql
    {
        [Required(ErrorMessage = "Please select file")]
        [FileExt(Allow = ".xls,.xlsx", ErrorMessage = "Only excel file")]
        public HttpPostedFileBase file { get; set; }
    }
}