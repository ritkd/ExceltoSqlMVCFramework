using ExceltoSqlMVCFramework.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExceltoSqlMVCFramework.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult ReadExcelUsingEpplus()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ReadExcel(HttpPostedFileBase upload)
        {
            var usersList = new List<Customer>();
            if (Request != null)
            {
                if (Path.GetExtension(upload.FileName) == ".xlsx" ||
                Path.GetExtension(upload.FileName) == ".xls")
                {
                    using (var package = new ExcelPackage(upload.InputStream))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var workSheet = currentSheet.First();
                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;
                        for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                        {
                            var user = new Customer();
                            user.CustomerID = Convert.ToInt32(workSheet.Cells[rowIterator, 1].Value);
                            user.CustomerFirstName = workSheet.Cells[rowIterator, 2].Value.ToString();
                            user.CustomerLastName = workSheet.Cells[rowIterator, 3].Value.ToString();
                            usersList.Add(user);
                        }
                    }
                }
            }
            using (ExcelImportDBEntities excelImportDBEntities = new ExcelImportDBEntities())
            {
                foreach (var item in usersList)
                {
                    excelImportDBEntities.Customers.Add(item);
                }
                excelImportDBEntities.SaveChanges();
            }
            return View("Index");
        }
    }
}