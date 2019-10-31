

namespace MyProj.Models
{
    public class Shipment
    {
        public int Id { get; set; }
        public int? ShipmentInfoId { get; set; }

        public ShipmentInfo ShipmentInfo { get; set; }
    }
}