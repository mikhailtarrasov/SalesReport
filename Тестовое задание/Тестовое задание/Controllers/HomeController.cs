using System;
using System.Web.Mvc;
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
                return View();
            }
            return RedirectToAction("ReportingPeriod", "Home");
        }
    }
}