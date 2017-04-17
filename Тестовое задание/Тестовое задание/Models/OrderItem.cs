using System.ComponentModel.DataAnnotations;

namespace Тестовое_задание.Models
{
    public class OrderItem
    {
        [Display(Name = "Номер заказа")]
        public int OrderId { get; set; }
        [Display(Name = "Дата заказа")]
        public string OrderDate { get; set; }         
        [Display(Name = "Артикул товара")]
        public int ProductId { get; set; }              
        [Display(Name = "Название товара")]
        public string ProductName { get; set; }         
        [Display(Name = "Кол-во реализ. ед. товара")]
        public short Quantity { get; set; }             
        [Display(Name = "Цена реализ. за ед. прод.")]
        public decimal UnitPrice { get; set; }          
    }
}