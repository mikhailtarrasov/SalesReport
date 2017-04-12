using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using AutoMapper;
using OfficeOpenXml;
using Тестовое_задание.Models;

namespace Тестовое_задание.Controllers
{
    public class HomeController : Controller
    {
        private NorthwindNETEntities Db = new NorthwindNETEntities();
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult ReportingPeriod()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ReportingPeriod(ReportingPeriodViewModel period)
        {
            if (period.StartDate.HasValue && period.EndDate.HasValue)
            {
                return RedirectToAction("Report", "Home", new ReportingPeriodViewModel(){ StartDate = period.StartDate.Value, EndDate = period.EndDate.Value });
            }
            return View();
        }

        [HttpGet]
        public ActionResult Report(ReportingPeriodViewModel reportPeriod)
        {
            if (reportPeriod.StartDate.HasValue && reportPeriod.EndDate.HasValue)
            {
                ViewBag.reportPeriod = reportPeriod;

                IQueryable<Order> orders = Db.Order.Where(x => x.OrderDate > reportPeriod.StartDate && x.OrderDate < reportPeriod.EndDate);
                if (orders.Any())
                {
                    var itemsList = CreateOrdersItemsList(orders);
                    CreateSpreadsheetReport(itemsList, reportPeriod);

                    return View(itemsList);
                }
            }
            return RedirectToAction("ReportingPeriod", "Home");
        }

        public List<OrderItem> CreateOrdersItemsList(IQueryable<Order> orders)
        {
            Mapper.Initialize(cfg => cfg.CreateMap<OrderDetail, OrderItem>()
                        .ForMember(x => x.OrderDate, x => x.MapFrom(od => od.Order.OrderDate.Value.ToShortDateString()))
                        .ForMember(x => x.ProductName, x => x.MapFrom(od => od.Product.Name))       // Нужно ли учитывать скидку здесь?
                        );

            List<OrderItem> orderItemsList = new List<OrderItem>();
            foreach (var order in orders)
            {
                foreach (var orderDetail in order.OrderDetail)
                {
                    orderItemsList.Add(Mapper.Map<OrderDetail, OrderItem>(orderDetail));
                }
            }
            return orderItemsList;
        }

        public void CreateSpreadsheetReport(List<OrderItem> orderItems, ReportingPeriodViewModel reportPeriod)
        {
            string filelocation = "D:\\Reports\\";
            string filename = "report_" + reportPeriod.StartDate.Value.ToShortDateString() + "-" +
                              reportPeriod.EndDate.Value.ToShortDateString() + ".xlsx";

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filelocation + filename)))
            {
                string worksheetName = "Report";
                if (excelPackage.Workbook.Worksheets.Count == 1)
                    excelPackage.Workbook.Worksheets.Delete(worksheetName);
                var worksheet = excelPackage.Workbook.Worksheets.Add(worksheetName);

                worksheet.SetValue(1, 1, "Номер заказа");
                worksheet.SetValue(1, 2, "Дата заказа");
                worksheet.SetValue(1, 3, "Артикул товара");
                worksheet.SetValue(1, 4, "Название товара");
                worksheet.SetValue(1, 5, "Кол-во реализ. ед. товара");
                worksheet.SetValue(1, 6, "Цена реализ. за ед. прод.");
                worksheet.SetValue(1, 7, "Сумма");

                for (int j = 0; j < orderItems.Count; j++)
                {
                    worksheet.SetValue(j + 2, 1, orderItems[j].OrderId);
                    worksheet.SetValue(j + 2, 2, orderItems[j].OrderDate);
                    worksheet.SetValue(j + 2, 3, orderItems[j].ProductId);
                    worksheet.SetValue(j + 2, 4, orderItems[j].ProductName);
                    worksheet.SetValue(j + 2, 5, orderItems[j].Quantity);
                    worksheet.SetValue(j + 2, 6, orderItems[j].UnitPrice);
                    var quanity = worksheet.Cells[j + 2, 5].Address;
                    var price = worksheet.Cells[j + 2, 6].Address;
                    worksheet.Cells[j + 2, 7].Formula = string.Format("({0}*{1})", quanity, price);
                }
                excelPackage.Save();
            }
        }
    }
}