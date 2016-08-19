using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Aspose_Assignment_App.Models
{
     public class Class1
     {
          [Display(Name = "Select an Excel file")]
          public string EFile { get; set; }

          public string FullName { get; set; }
          public string Email { get; set; }
          public string Address { get; set; }
          public string Salary { get; set; }
          public List<Class1> Info { get; set; }
     }
}