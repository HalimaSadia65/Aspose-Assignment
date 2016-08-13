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
                         // ExcelDataReader works with the binary Excel file, so it needs a FileStream
                         // to get started:
                         Stream stream = file.InputStream;
                         
                         IExcelDataReader reader = null;
                         if (file.FileName.EndsWith(".xls"))
                         {
                              reader = ExcelReaderFactory.CreateBinaryReader(stream);
                         }
                         else if (file.FileName.EndsWith(".xlsx"))
                         {
                              reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                         }
                         else
                         {
                              ModelState.AddModelError("File", "This file format is not supported");
                              return View();
                         }
                         reader.IsFirstRowAsColumnNames = true;
                         DataSet result = reader.AsDataSet();
                         reader.Close();
                         @ViewBag.filepath = path;
                         return View(result.Tables[0]);
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
              Excel1.Application app = new Excel1.Application();
              Excel1.Workbook workbook = app.Workbooks.Open("D:\\Aspose_Assignment\\Employees Data Book1.xlsx");
              Excel1.Worksheet worksheet = workbook.ActiveSheet;
              Excel1.Range range = worksheet.UsedRange;
              Excel1.Application app1 = new Excel1.Application();
              Excel1.Workbook workbook1 = app1.Workbooks.Open(path);
              Excel1.Worksheet worksheet1 = workbook1.ActiveSheet;
              Excel1.Range range1 = worksheet1.UsedRange;
              for (int row = 2; row <= range.Rows.Count; row++)
              {
                   for (int col = 1; col <= range.Columns.Count; col++)
                   {
                        worksheet.Cells[row, col].Value = worksheet1.Cells[row, col].Value;
                   }
              }
              workbook.Save();
              workbook.Close();

               for (int row = 2; row <= range.Rows.Count; row++)
               {
                    const string datadir = "D:\\Aspose_Assignment\\";
                    Document doc = new Document(datadir + "Mail_Merge_Template.docx");
                    doc.MailMerge.Execute(new string[] { "FullName", "Email", "Address", "Salary" }, new object[] { worksheet.Cells[row, 1].Value, worksheet.Cells[row, 2].Value, worksheet.Cells[row,3].Value, worksheet.Cells[row, 4].Value });
                    doc.Save(datadir + "Output.docx");
                    SendEmail se = new SendEmail();
                    string email = worksheet.Cells[row, 2].Value;
                    se.sendEmail(email);
               }
              
               return View();
          }

     }
}