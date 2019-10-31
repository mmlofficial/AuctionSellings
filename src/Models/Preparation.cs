using System;
using System.ComponentModel.DataAnnotations.Schema;


namespace MyProj.Models
{
    public class Preparation
    {
        public int Id { get; set; }
        public int? OrderInfoId { get; set; }
        public string Name { get; set; }
        public int Amount { get; set; }
        public string ExpirationDate { get; set; }
        public double Total { get; set; }
        public double TotalVAT { get; set; }
        public string PaymentDate { get; set; }

        public OrderInfo OrderInfo { get; set; }

    }
}