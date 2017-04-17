using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web.Mvc;
using AutoMapper;
using OfficeOpenXml;
using Тестовое_задание.Models;

namespace Тестовое_задание.Controllers
{
    public class HomeController : Controller
    {
        private readonly NorthwindNETEntities _db = new NorthwindNETEntities();

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
                if (period.EndDate.Value > DateTime.Today)
                {
                    period.EndDate = DateTime.Today;
                }
                var totalItems = _db.OrderDetail.Count(x => x.Order.OrderDate >= period.StartDate.Value &&
                            x.Order.OrderDate <= period.EndDate.Value);
                if (totalItems > 0)
                {
                    var pageSize = 50;
                    var startDate = period.StartDate;
                    var endDate = period.EndDate;
                    var totalPages = (int)Math.Ceiling((decimal)totalItems / pageSize);
                    var page = 1;
                    var reportPageViewModel = new ReportPageViewModel(startDate, endDate, pageSize, totalItems, totalPages, page);
                    reportPageViewModel.OrderItems = GetOrderItems(startDate, endDate, pageSize, page);
                    ViewBag.ResultSending = default(KeyValuePair<string, string>);

                    return View("Report", reportPageViewModel);
                }
            }
            return View();
        }

        [HttpGet]
        public ActionResult Report(DateTime? startDate, DateTime? endDate, int pageSize, int totalItems, int totalPages, int page, KeyValuePair<string, string> resultSending = default(KeyValuePair<string, string>))
        {
            ReportPageViewModel reportPage = new ReportPageViewModel(startDate, endDate, pageSize, totalItems, totalPages, page);

            if (reportPage.ReportingPeriod != null && reportPage.ReportingPeriod.StartDate.HasValue && reportPage.ReportingPeriod.EndDate.HasValue)
            {
                reportPage.OrderItems = GetOrderItems(startDate, endDate, pageSize, page);
                ViewBag.ResultSending = resultSending;

                return View(reportPage);
            }
            return ReportingPeriod();
        }

        [HttpPost]
        public ActionResult Report(DateTime? startDate, DateTime? endDate, int pageSize, int pageNumber, int totalItems, int totalPages, string inputEmail)
        {
            var resultSending = new KeyValuePair<string, string>();
            try
            {
                string filelocation = "D:\\Reports\\";
                string filename = "report_" + startDate.Value.ToShortDateString() + "-" + endDate.Value.ToShortDateString() + ".xlsx";

                if (!System.IO.File.Exists(filelocation + filename))                    // Если файла за выбранный период не существует - то создаем его
                {
                    var ordersItemsList = GetOrderItems(startDate, endDate, totalItems, 1);
                    CreateSpreadsheetReport(ordersItemsList, filelocation, filename);
                }

                resultSending = SendReportToEmail(filelocation, filename, inputEmail);
            }
            catch (Exception e)
            {
                resultSending = new KeyValuePair<string, string>("error", e.Message);
            }
           
            return Report(startDate, endDate, pageSize, totalItems, totalPages, pageNumber, resultSending);
        }

        private List<OrderItem> GetOrderItems(DateTime? startDate, DateTime? endDate, int pageSize, int page)
        {
            var dbItems = _db.OrderDetail.OrderBy(x => x.Order.OrderDate).Where(x => x.Order.OrderDate >= startDate && x.Order.OrderDate <= endDate)
                .Skip((page - 1) * pageSize)
                .Take(pageSize);

            Mapper.Initialize(cfg => cfg.CreateMap<OrderDetail, OrderItem>()
                .ForMember(x => x.OrderDate, x => x.MapFrom(od => od.Order.OrderDate.Value.ToShortDateString()))
                .ForMember(x => x.ProductName, x => x.MapFrom(od => od.Product.Name)));       // Нужно ли учитывать скидку здесь?

            List<OrderItem> orderItems = new List<OrderItem>();
            foreach (var orderDetail in dbItems)
            {
                orderItems.Add(Mapper.Map<OrderDetail, OrderItem>(orderDetail));
            }
            return orderItems;
        }

        public void CreateSpreadsheetReport(List<OrderItem> ordersItemsList, string filelocation, string fileName)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filelocation + fileName)))
            {
                string worksheetName = "Report";
                var worksheet = excelPackage.Workbook.Worksheets.Add(worksheetName);

                worksheet.SetValue(1, 1, "Номер заказа");
                worksheet.SetValue(1, 2, "Дата заказа");
                worksheet.SetValue(1, 3, "Артикул товара");
                worksheet.SetValue(1, 4, "Название товара");
                worksheet.SetValue(1, 5, "Кол-во реализ. ед. товара");
                worksheet.SetValue(1, 6, "Цена реализ. за ед. прод.");
                worksheet.SetValue(1, 7, "Сумма");

                for (int j = 0; j < ordersItemsList.Count; j++)
                {
                    int i = 1;
                    foreach (var prop in ordersItemsList[j].GetType().GetProperties())      // Проходим по каждому
                    {                                                                       // из полей текущего orderItem
                        worksheet.SetValue(j + 2, i, prop.GetValue(ordersItemsList[j]));    // и записываем его значение 
                        ++i;                                                                // в соответствующую 
                    }                                                                       // ячейку
                    var quanity = worksheet.Cells[j + 2, 5].Address;
                    var price = worksheet.Cells[j + 2, 6].Address;
                    worksheet.Cells[j + 2, ordersItemsList[i].GetType().GetProperties().Length + 1].Formula = string.Format("({0}*{1})", quanity, price);
                }
                worksheet.Calculate();
                excelPackage.Save();
                excelPackage.Stream.Close();
            }
        }

        public KeyValuePair<string, string> SendReportToEmail(string filelocation, string fileName, string emailTo)
        {
            KeyValuePair<string, string> result;
            try
            {
                var emailFrom = "mzd742@yandex.ru";
                var pass = "Тут был мой пароль от почты";
                var period = fileName.Replace("report_", "").Replace(".xlsx", "");

                MailAddress from = new MailAddress(emailFrom, "Mikhail Tarasov");
                MailAddress to = new MailAddress(emailTo);

                MailMessage m = new MailMessage(from, to);
                m.Subject = "Отчёт по продажам";
                m.Body = "<h2>Отчётный период: " + period + "</h2>";
                m.Attachments.Add(new Attachment(filelocation + fileName));
                m.IsBodyHtml = true;

                SmtpClient smtp = new SmtpClient("smtp.yandex.ru", 25);
                smtp.Credentials = new NetworkCredential(emailFrom, pass);
                smtp.EnableSsl = true;
                smtp.Send(m);
                result = new KeyValuePair<string, string>("success", "Сообщение успешно доставлено на адрес " + emailTo + "!");
            }
            catch (Exception e)
            {
                result = new KeyValuePair<string, string>("error", "Упс! Что-то пошло не так! " + e.Message);
            }

            return result;
        }
    }
}