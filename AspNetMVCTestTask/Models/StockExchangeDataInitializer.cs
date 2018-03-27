using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;

namespace AspNetMVCTestTask.Models
{
    public class StockExchangeDataInitializer : CreateDatabaseIfNotExists<StockExchangeDataContext>
    {
        private string inputFile = "Выручка.xlsx";
        protected override void Seed(StockExchangeDataContext db)
        {

            Excel.Application excelApp = new Excel.Application {Visible = false};

            var path = HttpContext.Current.Server.MapPath("~/App_Data") + "/" + inputFile;
            excelApp.Workbooks.Open(path);



            int row = 3;
            Excel.Worksheet currentSheet = (Excel.Worksheet)excelApp.Workbooks[1].Worksheets[2];
            while (currentSheet.Range["A" + row].Value2 != null)
            {
                var aoDate = Convert.ToDouble(currentSheet.Range["A" + row].Value2);
                var date = DateTime.FromOADate(aoDate);
                var profit = Convert.ToDecimal(currentSheet.Range["B" + row].Value2.ToString());
                var silver = Convert.ToDecimal(currentSheet.Range["C" + row].Value2.ToString());
                var moex = Convert.ToDecimal(currentSheet.Range["D" + row].Value2.ToString());

                db.StockExchangeData.Add(new StockExchangeData { BusinessDay = date, MoexPrice = moex, Profit = profit, SilverPrice = silver });


                row++;
            }

            excelApp.Quit();
            //db.StockExchangeData.Add(new StockExchangeData { BusinessDay = DateTime.Now, MoexPrice = 1, Profit = 220 , SilverPrice = 123});
            //db.StockExchangeData.Add(new StockExchangeData { BusinessDay = DateTime.Now.AddDays(-5), MoexPrice = 11, Profit = 2210, SilverPrice = 1223 });


            base.Seed(db);
        }
    }
}