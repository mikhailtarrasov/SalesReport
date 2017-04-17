using System;
using System.Collections.Generic;

namespace Тестовое_задание.Models
{
    public class ReportPageViewModel
    {
        public List<OrderItem> OrderItems { get; set; }
        public ReportingPeriodViewModel ReportingPeriod { get; set; }
        public PageInfo PageInfo { get; set; }

        public ReportPageViewModel(DateTime? startDate, DateTime? endDate, int pageSize, int totalItems,
            int totalPages, int page)
        {
            ReportingPeriod = new ReportingPeriodViewModel()
            {
                StartDate = startDate,
                EndDate = endDate
            };
            PageInfo = new PageInfo()
            {
                PageSize = pageSize,
                PageNumber = page,
                TotalItems = totalItems
            };
        }

        public ReportPageViewModel()
        {
        }
    }

    public class PageInfo
    {
        public int PageNumber { get; set; }
        public int PageSize { get; set; }
        public int TotalItems { get; set; } // всего объектов
        public int TotalPages
        {
            get { return (int)Math.Ceiling((decimal)TotalItems / PageSize); }
        }
    }
}