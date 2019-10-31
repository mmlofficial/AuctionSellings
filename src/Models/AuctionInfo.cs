using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace MyProj.Models
{
    public class AuctionInfo
    {
        public int Id { get; set; }
        public string AuctionNumber { get; set; }
        public string Date { get; set; }
        public string Status { get; set; } // 0 - in progress, 1 - completed true, -1 - completed false 
    }
}