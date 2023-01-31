using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Web;
using VT1.Models;
using VT1.Services;
using OfficeOpenXml;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using Microsoft.AspNetCore.StaticFiles;
using System.Xml.Linq;
using System.Collections.Generic;

namespace VT1.Controllers
{
    public class TableController : Controller
    {
        private IExcelService? service;

        public TableController(IExcelService excelService)
        {
            var MainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploads");
            service = excelService;
       
            if (Directory.Exists(MainPath))
            {
                
                Directory.Delete(MainPath, true);
            }

            //tablez.Clear();
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //package = new ExcelPackage(filePath);

        }


        [HttpGet]
        public IActionResult TableView(IFormCollection form)
        {

            //if (!tablez.Any())
            //{

            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //var sss = new ExcelPackage(filePath);
            //var ss = new ExcelService(sss);
            //var table = ExcelService.Instance.TableDetection();
            //List<Table> tables = ExcelService.Instance.MapToTable(table);
            //tables.Except(tables.Where(t => t.ContainsKey(t.name)));
            List<Table> tables = new List<Table>();
                tables = TempData.Get<List<Table>>("tables");
            TempData.Keep("tables");

            return View(tables);
            //}

            //return View(tables);
        }

        [HttpPost]
        public IActionResult TableView(IFormFile FormFile)
        {
            //var sss = new ExcelPackage(filePath);
            //var ss = new ExcelService(sss);
            //var tables = ss.TableDetection();
            //var tablez = ExcelService.Instance.FFinalTables;
            var tablez = service.FFinalTables;

            try {
                foreach (Table table in tablez)
                {
                    var transformator = new DbService(table);
                    transformator.CreateDb("Test");
                    transformator.CreateTable();
                    transformator.TableInsert();
                    TempData["Success"] = "Success";
                }
            }
            catch
            {
                TempData["Error"] = "Error";
            }

            return View(service.FFinalTables);
        }

        [HttpGet]
        public IActionResult Preview(string name)
        {


            //var sss = new ExcelPackage(filePath);
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //var ss = new ExcelService(sss);
            //var tables = ss.TableDetection();
            List<Table> tablez = service.FFinalTables.Where(t=> t.tableName == name).ToList();
            //List<Table> tablez = ExcelService.Instance.FFinalTables;

            return View(tablez);
        }

        [HttpPost]
        public bool Delete(string name)
        {
            try
            {
                //temp = new Dictionary<string, string>();
                ////temp.Add(name, name);
                //var sss = new ExcelPackage(filePath);
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //var ss = new ExcelService(sss);
                //var tables = ss.TableDetection();
                //List<Table> tablez = ExcelService.Instance.FFinalTables;
                service.FFinalTables.Remove(service.FFinalTables.Where(t => t.tableName == name).First());

                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }


        [HttpGet]
        public IActionResult Export()
        {
            var jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(service.FFinalTables, Formatting.Indented);

            //return Json(ExcelService.instance.models);

            string fileName = "jsonExport.json";
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(jsonData);

            var content = new System.IO.MemoryStream(bytes);
            return File(content, "application/json", fileName);
        }
    }
}
