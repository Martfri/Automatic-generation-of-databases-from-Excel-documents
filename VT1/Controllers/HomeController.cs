using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using VT1.Models;
using VT1.Services;
using OfficeOpenXml;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using System.Net.Http.Headers;

namespace VT1.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;       
        private static string filePath = ".\\wwwroot\\Uploads\\Test.xlsx";
        //private ExcelPackage package;    
        //private ExcelService service;      


        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //package = new ExcelPackage(filePath);
            
        }

        [HttpGet]

        public IActionResult Index(IFormCollection form)
        {
            return View();
        }

        [HttpGet]
        public IActionResult Statistics()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ImportExcelFile(IFormFile FormFile)
        {
            //get file name
            var filename = ContentDispositionHeaderValue.Parse(FormFile.ContentDisposition).FileName.Trim('"');

            //get path
            var MainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploads");

            if (Directory.Exists(MainPath))
            {
                Directory.Delete(MainPath, true);
            }

            var filePath = Path.Combine(MainPath, FormFile.FileName);

            string extension = Path.GetExtension(filename);

            string conString = string.Empty;

            switch (extension)
            {
                case ".xls": //Excel 97-03.
                    conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                    break;
                case ".xlsx": //Excel 07 and above.
                    conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                    break;
            }

            //get extension
            if (extension != ".xlsx")
            {
                ViewBag.Message = "Uploaded file is not an xlsx document";
            }
            else
            {
                ViewBag.Message = "File uploaded";


            }

            //create directory "Uploads" if it doesn't exists
            if (!Directory.Exists(MainPath))
            {
                Directory.CreateDirectory(MainPath);
            }

            //get file path 
            filePath = Path.Combine(MainPath, "Test.xlsx");

            //service = new ExcelService(filePath);


            using (System.IO.Stream stream = new FileStream(filePath, FileMode.Create))
            {
                await FormFile.CopyToAsync(stream);
            }

            //get extension


            return View("Index");
        }

        [HttpGet]
        public IActionResult TableView(IFormCollection form)
        {
            var sss = new ExcelPackage(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var ss = new ExcelService(sss);
            var tables = ss.TableDetection();
            Table table = ss.MapToTable(tables.First());
            List<Table> Tables = new List<Table>();
            Tables.Add(table);
            return View(Tables);
        }

        [HttpPost]
        public IActionResult TableView(IFormFile FormFile)
        {
            var sss = new ExcelPackage(filePath);
            var ss = new ExcelService(sss);
            var tables = ss.TableDetection();
            var table = ss.MapToTable(tables.First());
            var transformator = new DbService(table);
            transformator.CreateDb("Test");
            transformator.CreateTable();
            transformator.TableInsert();
            return Ok();
        }
    }
}