using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace MyProj.Models
{
    public class OrderInfo
    {
        public int Id { get; set; }
        public string Date { get; set; }
        public string PreShipmentDate { get; set; }
        public string Status { get; set; } // 0 - in progress, 1 - completed, -1 - errors
        public string CustomerName { get; set; }
        public string CustomerLocationArea { get; set; }
        public string CustomerCity { get; set; }

        public ICollection<Preparation> Preparations { get; set; }

        public OrderInfo()
        {
            Preparations = new List<Preparation>();
        }

    }
}