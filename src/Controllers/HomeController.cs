using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MyProj.Models;
using MyProj.DAL;
using System.IO;
using MyProj.Help;

namespace MyProj.Controllers
{
    public class HomeController : Controller
    {
        MyProjContext db = new MyProjContext();

        public string pathx = "C:\\Maxim\\BMSTU\\6\\KursachDB\\MyProj\\MyProj\\Files\\AAA.xlsx"; //pepe

        public ActionResult Index() //TODO
        {
            return View();
        }


        /*[HttpGet]
        public ActionResult InvoicePreparation()
        {
            return View();
        }*/

        [HttpGet]
        public ActionResult InvoicePreparation()    //TODO
        {
            return View();
        }

        [HttpPost]
        public ActionResult InvoicePreparation(AllInfo allInfo)    //TODO
        {
            InvoiceGen IG = new InvoiceGen(allInfo);

            using (var stream = new MemoryStream())
            {
                IG.workbook.SaveAs(stream);
                stream.Flush();

                return new FileContentResult(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    FileDownloadName = $"Invoice_{allInfo.AuctionNumber}_{DateTime.UtcNow.ToShortDateString()}.xlsx"
                };
            }
        }


        public ActionResult Search()
        {
            return View();
        }

        [HttpPost]
        public ActionResult SearchInfo(string str)
        {
            var allinfo = (from order in db.MainOrders
                           join auction in db.Auctions on order.AuctionId equals auction.Id
                           join auctioninfo in db.AuctionInfos on auction.AuctionInfoId equals auctioninfo.Id
                           join orderinfo in db.OrderInfos on order.OrderInfoId equals orderinfo.Id
                           join shipment in db.Shipments on order.ShipmentId equals shipment.Id
                           join shipmentinfo in db.ShipmentInfos on shipment.ShipmentInfoId equals shipmentinfo.Id
                           join distributor in db.Distributors on shipmentinfo.DistributorId equals distributor.Id
                           where auctioninfo.AuctionNumber == str
                           select new
                           {
                               AuctionNumber = auctioninfo.AuctionNumber,
                               AuctionDate = auctioninfo.Date,
                               DistributorName = distributor.Name,
                               OrderId = orderinfo.Id.ToString(),
                               OrderDate = orderinfo.Date,
                               OrderPreShipmentDate = orderinfo.PreShipmentDate,
                               CustomerName = orderinfo.CustomerName,
                               CustomerLocationArea = orderinfo.CustomerLocationArea,
                               CustomerCity = orderinfo.CustomerCity,
                               OrderStatus = orderinfo.Status,
                               ShipmentDate = shipmentinfo.Date,
                               ShipmentStatus = shipmentinfo.Status,
                               OrderInfoId = orderinfo.Id
                           }).ToList();
            if (allinfo.Count <= 0)
            {
                return HttpNotFound();
            }

            int orderinfoid = allinfo[0].OrderInfoId;
            var allpreparations = db.Preparations.Where(a => a.OrderInfoId == orderinfoid).ToList();


            //дикий костыль для обхода анонимного типа IQueryble<'a>
            AllInfo info = new AllInfo(allinfo[0].AuctionNumber, allinfo[0].AuctionDate, allinfo[0].DistributorName, allinfo[0].OrderId,
                allinfo[0].OrderDate, allinfo[0].OrderPreShipmentDate, allinfo[0].CustomerName, allinfo[0].CustomerLocationArea,
                allinfo[0].CustomerCity, allinfo[0].OrderStatus, allinfo[0].ShipmentDate, allinfo[0].ShipmentStatus, allpreparations);
            var kek = (new[] { info }).ToList();

            return PartialView(kek);
        }

        //*****************************************
        [HttpGet]
        public ActionResult EditOrderInfo(int? id)
        {
            if (id == null)
                return HttpNotFound();
            OrderInfo orderInfo = db.OrderInfos.Find(id);
            if (orderInfo != null)
            {
                return View(orderInfo);
            }
            return HttpNotFound();
        }

        [HttpPost]
        public ActionResult EditOrderInfo(OrderInfo orderInfo)
        {
            db.Entry(orderInfo).State = System.Data.Entity.EntityState.Modified;
            db.SaveChanges();
            return RedirectToAction("Search/" + orderInfo.Id.ToString());
        }
        //***************************************************

        //*****************************************
        [HttpGet]
        public ActionResult EditShipmentInfo(int? id)
        {
            if (id == null)
                return HttpNotFound();
            ShipmentInfo shipmentInfo = db.ShipmentInfos.Find(id);
            if (shipmentInfo != null)
            {
                return View(shipmentInfo);
            }
            return HttpNotFound();
        }

        [HttpPost]
        public ActionResult EditShipmentInfo(ShipmentInfo shipmentInfo)
        {
            db.Entry(shipmentInfo).State = System.Data.Entity.EntityState.Modified;
            db.SaveChanges();
            //string status = shipmentInfo.Status; //удаление из склада
            //if (status == "Complete")
                
            return RedirectToAction("Search");
        }
        //***************************************************

        //*****************************************
        [HttpGet]
        public ActionResult EditPaymentDate(int? id)
        {
            if (id == null)
                return HttpNotFound();
            Preparation preparation = db.Preparations.Find(id);
            if (preparation != null)
            {
                return View(preparation);
            }
            return HttpNotFound();
        }

        [HttpPost]
        public ActionResult EditPaymentDate(Preparation preparation)
        {
            db.Entry(preparation).State = System.Data.Entity.EntityState.Modified;
            db.SaveChanges();
            return RedirectToAction("Search");
        }
        //***************************************************

        [HttpGet]
        public ActionResult AddPreparation()
        {
            return View();
        }

        [HttpPost]
        public ActionResult AddPreparation(PreparationsBase preparationsBase)
        {
            db.PreparationsBases.Add(preparationsBase);
            db.SaveChanges();
            return RedirectToAction("LookPreparationBase");
        }

        [HttpGet]
        public ActionResult DeletePreparation(int id)
        {
            PreparationsBase b = db.PreparationsBases.Find(id);
            if (b == null)
            {
                return HttpNotFound();
            }
            return View(b);
        }
        [HttpPost, ActionName("DeletePreparation")]
        public ActionResult DeleteConfirmed(int id)
        {
            PreparationsBase b = db.PreparationsBases.Find(id);
            if (b == null)
            {
                return HttpNotFound();
            }
            db.PreparationsBases.Remove(b);
            db.SaveChanges();
            return RedirectToAction("LookPreparationBase");
        }

        [HttpGet]
        public ActionResult EditPreparation(int? id)
        {
            if (id == null)
                return HttpNotFound();
            PreparationsBase preparationsBase = db.PreparationsBases.Find(id);
            if (preparationsBase != null)
            {
                return View(preparationsBase);
            }
            return HttpNotFound();
        }

        [HttpPost]
        public ActionResult EditPreparation(PreparationsBase preparationsBase)
        {
            db.Entry(preparationsBase).State = System.Data.Entity.EntityState.Modified;
            db.SaveChanges();
            return RedirectToAction("LookPreparationBase");
        }

        public ActionResult LookPreparationBase()
        {
            IEnumerable<PreparationsBase> preparations = db.PreparationsBases;
            ViewBag.PreparationsBases = preparations;
            return View();
        }

        public ActionResult LookPreparation()
        {
            IEnumerable<Preparation> preparations = db.Preparations;
            ViewBag.Preparations = preparations;
            return View();
        }

        public ActionResult LookShipmentInfo()
        {
            IEnumerable<ShipmentInfo> shipmentInfos = db.ShipmentInfos;
            ViewBag.ShipmentInfos = shipmentInfos;
            return View();
        }

        public ActionResult LookAuction()
        {
            IEnumerable<Auction> auctions = db.Auctions;
            ViewBag.Auctions = auctions;
            return View();
        }

        public ActionResult LookAuctionInfo()
        {
            IEnumerable<AuctionInfo> auctionInfos = db.AuctionInfos;
            ViewBag.AuctionInfos = auctionInfos;
            return View();
        }

        public ActionResult LookDistributor()
        {
            IEnumerable<Distributor> distributors = db.Distributors;
            ViewBag.Distributors = distributors;
            return View();
        }

        public ActionResult LookMainOrder()
        {
            IEnumerable<MainOrder> mainOrders = db.MainOrders;
            ViewBag.MainOrders = mainOrders;
            return View();
        }

        public ActionResult LookOrderInfo()
        {
            IEnumerable<OrderInfo> orderInfos = db.OrderInfos;
            ViewBag.OrderInfos = orderInfos;
            return View();
        }

        public ActionResult LookShipment()
        {
            IEnumerable<Shipment> shipments = db.Shipments;
            ViewBag.Shipments = shipments;
            return View();
        }

        public ActionResult LookAll()
        {
            return View();
        }

        public ActionResult AddFile()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadFiles(IEnumerable<HttpPostedFileBase> files)
        {
            foreach (var file in files)
            {
                FileWorks FW = new FileWorks(file, Server.MapPath("~/Files/"), db, this.pathx);
                ViewBag.Message = "Hello maxim";
                string filePath = Guid.NewGuid() + Path.GetExtension(file.FileName);
                file.SaveAs(Path.Combine(Server.MapPath("~/Files/"), filePath));
            }
            return Json("file uploaded successfully");
        }

        [HttpGet]
        public ActionResult AddTender()
        {
            return View();
        }
        [HttpPost]
        public ActionResult AddTender(HttpPostedFileBase file)
        {
            try
            {
                if (file.ContentLength > 0)
                {
                    string extension = Path.GetExtension(file.FileName);
                    if (extension == ".xlsx" || extension == ".xls")
                    {
                        file.SaveAs(Server.MapPath("~/Files/AAA.xlsx"));
                        ViewBag.Message = "Файл \"" + file.FileName.ToString() + "\" успешно загружен!";
                    }
                    else
                        ViewBag.Message = "Файл \"" + file.FileName.ToString() + "\" не загружен! Требуемые расширения : .xls .xlsx";
                }
                else
                    ViewBag.Message = "Файл \"" + file.FileName.ToString() + "\" пуст!";
                return View();
            }
            catch
            {
                ViewBag.Message = "Загрузка файла \"" + file.FileName.ToString() + "\" не произошла!";
                return View();
            }
        }
    }
}