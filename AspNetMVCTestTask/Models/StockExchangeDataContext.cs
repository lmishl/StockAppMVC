using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace AspNetMVCTestTask.Models
{
    public class StockExchangeDataContext : DbContext
    {
        public DbSet<StockExchangeData> StockExchangeData { get; set; }
    }
}