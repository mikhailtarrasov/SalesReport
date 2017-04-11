using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using AutoMapper;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
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
            if (period.StartDate != null && period.EndDate != null)
            {
                return RedirectToAction("Report", "Home", new
                {
                    startDate = period.StartDate, 
                    endDate = period.EndDate
                });
            }
            return View();
        }

        [HttpGet]
        public ActionResult Report(DateTime? startDate, DateTime? endDate)
        {
            if (startDate.HasValue && endDate.HasValue)
            {
                ViewBag.startDate = startDate.Value.ToString("dd.MM.yyyy");
                ViewBag.endDate = endDate.Value.ToString("dd.MM.yyyy");
                var orders = Db.Order.Where(x => x.OrderDate > startDate && x.OrderDate < endDate);
                if (orders.Any())
                {
                    Mapper.Initialize(cfg => cfg.CreateMap<OrderDetail, OrderItem>()
                        .ForMember(x => x.OrderId, x => x.MapFrom(od => od.OrderID))
                        .ForMember(x => x.OrderDate, x => x.MapFrom(od => od.Order.OrderDate))
                        .ForMember(x => x.ProductId, x => x.MapFrom(od => od.ProductID))
                        .ForMember(x => x.ProductName, x => x.MapFrom(od => od.Product.Name))
                        .ForMember(x => x.Quantity, x => x.MapFrom(od => od.Quantity))
                        .ForMember(x => x.UnitPrice, x => x.MapFrom(od => od.UnitPrice))       // Нужно ли учитывать скидку здесь?
                        );

                    List<OrderItem> orderItemsList = new List<OrderItem>();
                    foreach (var order in orders)
                    {
                        foreach (var orderDetail in order.OrderDetail)
                        {
                            orderItemsList.Add(Mapper.Map<OrderDetail, OrderItem>(orderDetail));
                        }
                    }
                    CreateSpreadsheetReport(orderItemsList);

                    return View(orderItemsList);
                }
            }
            return RedirectToAction("ReportingPeriod", "Home");
        }

        public void CreateSpreadsheetReport(List<OrderItem> orderItems)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo("D:\\report.xlsx")))
            {
                excelPackage.Workbook.Worksheets.Add("Report");
                excelPackage.Workbook.Worksheets[1].SetValue(1, 1, "Номер заказа");
                excelPackage.Workbook.Worksheets[1].SetValue(1, 2, "Дата заказа");
                excelPackage.Workbook.Worksheets[1].SetValue(1, 3, "Артикул товара");
                excelPackage.Workbook.Worksheets[1].SetValue(1, 4, "Название товара");
                excelPackage.Workbook.Worksheets[1].SetValue(1, 5, "Кол-во реализ. ед. товара");
                excelPackage.Workbook.Worksheets[1].SetValue(1, 6, "Цена реализ. за ед. прод.");

                for (int j = 0; j < orderItems.Count; j++)
                {
                    excelPackage.Workbook.Worksheets[1].SetValue(j + 2, 1, orderItems[j].OrderId);
                    excelPackage.Workbook.Worksheets[1].SetValue(j + 2, 2, orderItems[j].OrderDate.ToShortDateString());
                    excelPackage.Workbook.Worksheets[1].SetValue(j + 2, 3, orderItems[j].ProductId);
                    excelPackage.Workbook.Worksheets[1].SetValue(j + 2, 4, orderItems[j].ProductName);
                    excelPackage.Workbook.Worksheets[1].SetValue(j + 2, 5, orderItems[j].Quantity);
                    excelPackage.Workbook.Worksheets[1].SetValue(j + 2, 6, orderItems[j].UnitPrice);
                }
                excelPackage.Save();
            }
        }
    }
}