using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AspNetMVCTestTask.Models
{
    public class StockExchangeData
    {
        public int Id { get; set; }
        // дата
        public DateTime BusinessDay { get; set; }
        // выручка
        public decimal Profit { get; set; }
        // серебро
        public decimal SilverPrice { get; set; }
        // ММВБ
        public decimal MoexPrice { get; set; }
    }
}