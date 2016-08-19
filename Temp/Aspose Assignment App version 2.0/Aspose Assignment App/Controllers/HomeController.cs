using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel1 = Microsoft.Office.Interop.Excel;
using Aspose_Assignment_App.Models;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Settings;
using System.IO;
using Excel;
using System.Data;
using Aspose.Cells;
using System.Web.UI.WebControls;

namespace Aspose_Assignment_App.Controllers
{
     public class HomeController : Controller
     {
          //Screen to let the HR Manager upload the excel file
          [HttpGet]
          public ActionResult Index()
          {
               return View();
          }

          [HttpPost]
          public ActionResult Index(HttpPostedFileBase file)
          {
               if (ModelState.IsValid)
               {
                    string path = Server.MapPath("~/Files/" + file.FileName);
                    if (System.IO.File.Exists(path))
                         System.IO.File.Delete(path);
                    file.SaveAs(path);

                    if (file != null && file.ContentLength > 0)
                    {
                         //Creating a file stream containing the Excel file to be opened
                         FileStream fstream = new FileStream(path,FileMode.Open);

                         //Instantiating a Workbook object
                         //Opening the Excel file through the file stream
                         Workbook workbook = new Workbook(fstream);

                         //Accessing the first worksheet in the Excel file
                         Worksheet worksheet = workbook.Worksheets[0];

                         //Exporting the contents of 6 rows and 4 columns starting from 1st cell to DataTable
                         DataTable dataTable = worksheet.Cells.ExportDataTable(0, 0, 6, 4, true);

                         //Closing the file stream to free all resources
                         fstream.Close();
                         @ViewBag.filepath = path;
                         return View(dataTable);
                         }
               else
               {
                   ModelState.AddModelError("File", "Please Upload Your file");
               }
               }
               return View();
          }
          
          public ActionResult Send(string path)
          {
               string physicalpath = Server.MapPath("~/Files/Employees Data Book1.xlsx");
              //Creating a file stream containing the Default Excel file to be opened
              FileStream fstream = new FileStream(physicalpath, FileMode.Open, FileAccess.ReadWrite);
              //Instantiating a Workbook object
              //Opening the Excel file through the file stream
              Workbook workbook = new Workbook(fstream);

              //Accessing the first worksheet in the Excel file
              Worksheet worksheet = workbook.Worksheets[0];

              //Creating a file stream containing the Default Excel file to be opened
              FileStream filestream = new FileStream(path, FileMode.Open);
              //Instantiating a Workbook object
              //Opening the Excel file through the file stream
              Workbook workbook1 = new Workbook(filestream);

              //Accessing the first worksheet in the Excel file
              Worksheet worksheet1 = workbook1.Worksheets[0];
              fstream.Close();
              filestream.Close();
               for(int rows = 1; rows < 6 ; rows++)
               {
                    int col = 0;
                    // Access the cell by row and col indices
                    Cell cell = worksheet.Cells[rows, col];
                    Cell cell1 = worksheet1.Cells[rows, col];
                    if(cell.Value.ToString() == cell1.Value.ToString())
                    {
                         worksheet.Cells[rows, 3].Value = worksheet1.Cells[rows, 3].Value;
                    }  
               }
               workbook.Save(physicalpath);


               for (int row = 1; row < 6; row++)
               {
                    string Documentpath = Server.MapPath("~/Files/Mail_Merge_Template.docx");
                    Document doc = new Document(Documentpath);

                    //Execute Mail Merge
                    doc.MailMerge.Execute(new string[] { "FullName", "Email", "Address", "Salary" }, new object[] { worksheet.Cells[row, 0].Value, worksheet.Cells[row, 1].Value, worksheet.Cells[row, 2].Value, worksheet.Cells[row, 3].Value });
                    
                    //Physical Path of Files Folder
                    string directory = Server.MapPath("~/Files/");

                    //Save document on path
                    doc.Save(directory + "Salary-Increment-Letter_" + worksheet.Cells[row, 0].Value + ".docx");

                    //Physical Path For Document to be mailed
                    string outputdocpath = Server.MapPath("~/Files/Salary-Increment-Letter_" + worksheet.Cells[row, 0].Value + ".docx");

                    //Send Increment Letters
                    // Create instance of Email class
                    SendEmail se = new SendEmail();

                    // Extract Email from worksheet
                    string email = worksheet.Cells[1, 1].Value.ToString();

                    //Pass email and path of doc to send email function
                    se.sendEmail(email, outputdocpath);
               }
              
               return View();
          }

     }
}