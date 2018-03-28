using AspNetMVCTestTask.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.Helpers;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;

namespace AspNetMVCTestTask.Controllers
{
    public class HomeController : Controller
    {
        // создаем контекст данных
        StockExchangeDataContext db = new StockExchangeDataContext();

        public ActionResult Index()
        {
            // получаем из бд все объекты
            IEnumerable<StockExchangeData> data = db.StockExchangeData;
            // передаем все объекты в динамическое свойство в ViewBag
            ViewBag.Data = data;
            // возвращаем представление
            return View();
        }



        [HttpPost]
        public ActionResult Index(StockExchangeData data)
        {
            db.StockExchangeData.Add(data);
            // сохраняем в бд все изменения
            db.SaveChanges();
            return Index();
        }

        [HttpGet]
        public ActionResult MakeChart()
        {

            IEnumerable<StockExchangeData> data = db.StockExchangeData;

            var myChart = new Chart(width: 600, height: 400)
                .AddTitle("Chart Title")
                .AddSeries(
                    name: "Profit",
                    xValue: data.Select(d => d.BusinessDay).ToArray(),
                    yValues: data.Select(d => d.Profit).ToArray());

            ViewBag.Chart = myChart;


            // возвращаем представление
            return View();
        }


        [HttpGet]
        public FileResult Download()
        {
            string path = Server.MapPath("~/App_Data/Files/excelFile" + Guid.NewGuid() + ".xls");

            Excel.Application excelApp = new Excel.Application { Visible = false };
            var xlWorkBook = excelApp.Workbooks.Add();

            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];
            xlWorkSheet.Cells[1, 1] = "Показатели Заемщика";
            xlWorkSheet.Range["A1", "B1"].Merge(); 
            xlWorkSheet.Cells[1, 3] = "Валюта";
            xlWorkSheet.Cells[1, 4] = "Индексы";
            xlWorkSheet.Cells[2, 1] = "Дата";
            xlWorkSheet.Cells[2, 2] = "Выручка";
            xlWorkSheet.Cells[2, 3] = "серебро, руб. ";
            xlWorkSheet.Cells[2, 4] = "Индекс ММВБ Last";

            int row = 3;
            foreach (var data in db.StockExchangeData)
            {
                xlWorkSheet.Cells[row, 1] = data.BusinessDay;
                xlWorkSheet.Cells[row, 2] = data.Profit;
                xlWorkSheet.Cells[row, 3] = data.SilverPrice;
                xlWorkSheet.Cells[row, 4] = data.MoexPrice;
                row++;
            }

            xlWorkBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal);
            xlWorkBook.Close(true);
            excelApp.Quit();



            // Объект Stream
            FileStream fs = new FileStream(path, FileMode.Open);
            string file_type = "application/excel";
            string file_name = "excelFile.xls";
            return File(fs, file_type, file_name);

        }

      

    }
}
