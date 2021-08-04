using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelDataReader.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDataReader.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase ExcelFile)
        {
            if (ExcelFile==null || ExcelFile.ContentLength == 0)
            {
                ViewBag.Error = "Please select an Excel File<br>";
                return View("Index");
            }
            else
            {
                if(ExcelFile.FileName.EndsWith("xls")|| ExcelFile.FileName.EndsWith("xlsx")||ExcelFile.FileName.EndsWith("csv"))
                {
                    string path = Server.MapPath("~/Content/"+ExcelFile.FileName) ;
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    ExcelFile.SaveAs(path);
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<ExcelFile> list_ex = new List<ExcelFile>();
                    for (int row = 12; row < range.Rows.Count; row++)
                    {
                        ExcelFile ex = new ExcelFile();
                        ex.age = ((Excel.Range)range.Cells[row, 1]).Text;
                        ex.anaemia = ((Excel.Range)range.Cells[row,2]).Text;
                        ex.creatinine_phosphokinase = ((Excel.Range)range.Cells[row, 3]).Text;
                        ex.diabetes = ((Excel.Range)range.Cells[row, 4]).Text;
                        ex.ejection_fraction = ((Excel.Range)range.Cells[row,5]).Text;
                        ex.high_blood_pressure = ((Excel.Range)range.Cells[row,6]).Text;
                        ex.platelets = ((Excel.Range)range.Cells[row, 7]).Text;
                        ex.serum_creatinine = ((Excel.Range)range.Cells[row, 8]).Text;
                        ex.smoking = ((Excel.Range)range.Cells[row, 9]).Text;
                        ex.time = ((Excel.Range)range.Cells[row, 10]).Text;
                        ex.DEATH_EVENT = ((Excel.Range)range.Cells[row, 11]).Text;
                        list_ex.Add(ex);
                        ViewBag.ExcelFile = list_ex;
                    }
                            return View("Success");
                }
                else
                {
                    ViewBag.Error = "File type is incorrect<br>";
                    return View("Index");
                }
                
            }
        }

    }
}