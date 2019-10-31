using System;
using System.ComponentModel.DataAnnotations.Schema;


namespace MyProj.Models
{
    public class ShipmentInfo
    {
        public int Id { get; set; }
        public string Date { get; set; }
        public int? DistributorId { get; set; }
        public string Status { get; set; } // 0 - in progress, 1 - completed, -1 - errors

        public Distributor Distributor { get; set; }
    }
}