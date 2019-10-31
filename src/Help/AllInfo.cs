using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MyProj.Help
{
    public class AllInfo
    {
        public string AuctionNumber { get; set; }
        public string AuctionDate { get; set; }
        public string DistributorName { get; set; }
        public string OrderId { get; set; }
        public string OrderDate { get; set; }
        public string OrderPreShipmentDate { get; set; }
        public string CustomerName { get; set; }
        public string CustomerLocationArea { get; set; }
        public string CustomerCity { get; set; }
        public string OrderStatus { get; set; }
        public string ShipmentDate { get; set; }
        public string ShipmentStatus { get; set; }
        public List<MyProj.Models.Preparation> PreparationList { get; set; }

        public AllInfo(string a, string b, string c, string d, string e, string f, string g, string h, string i, string j, string k, string l, List<MyProj.Models.Preparation> p)
        {
            AuctionNumber = a;
            AuctionDate = b;
            DistributorName = c;
            OrderId = d;
            OrderDate = e;
            OrderPreShipmentDate = f;
            CustomerName = g;
            CustomerLocationArea = h;
            CustomerCity = i;
            OrderStatus = j;
            ShipmentDate = k;
            ShipmentStatus = l;
            PreparationList = new List<Models.Preparation>();
            foreach (var item in p)
                PreparationList.Add(item);
        }
    }
}