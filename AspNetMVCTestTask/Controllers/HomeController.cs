using AspNetMVCTestTask.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

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
    }
}
