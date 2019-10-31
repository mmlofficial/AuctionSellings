using System;
using System.Collections.Generic;



namespace MyProj.Models
{
    public class MainOrder
    {
        public int Id { get; set; }
        public int? OrderInfoId { get; set; }
        public int? AuctionId { get; set; }
        public int? ShipmentId { get; set; }

        public OrderInfo OrderInfo { get; set; }
        public Auction Auction { get; set; }
        public Shipment Shipment { get; set; }
    }
}