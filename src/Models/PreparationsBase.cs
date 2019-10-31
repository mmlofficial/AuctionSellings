using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MyProj.Models
{
    public class PreparationsBase
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Price { get; set; }
        public double Amount { get; set; }
        public string PackageType { get; set; }
        public string Unit { get; set; }
        public string ProductionCode { get; set; }
        public double Weight { get; set; }

    }
}