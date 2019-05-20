using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace Allegato3API.Models
{
    [Table("Templates")]
    public class Templates
    {
        [Key]
        public int TemplateID { get; set; }

        public string JsonString { get; set; }

    }

}