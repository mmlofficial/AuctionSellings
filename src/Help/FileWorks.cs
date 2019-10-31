using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using MyProj.DAL;

namespace MyProj.Help
{
    public class FileWorks
    {
        private string path;
        private string pathx;
        public int error { get; set; }

        public FileWorks(HttpPostedFileBase file, string path, MyProjContext db, string tenderpath)
        {
            this.path = path;
            this.pathx = tenderpath;
            Int32 strLen = Convert.ToInt32(file.InputStream.Length);
            byte[] byteArr = new byte[strLen];     // Create a byte array.
            Int32 strRead = file.InputStream.Read(byteArr, 0, strLen);    // Read stream into byte array.
            string ext = Path.GetExtension(file.FileName);
            this.path += "Wow" + ext;
            File.WriteAllBytes(this.path, byteArr);
            if (ExcelWorks(db) == -1)
                this.error = -1;
        }

        private int ExcelWorks(MyProjContext db)
        {
            Excel.Application application = new Excel.Application();
            application.Visible = false;
            Excel.Workbook workbook = application.Workbooks.Open(this.path, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];


            int maxrows = worksheet.UsedRange.Rows.Count;
            int maxcols = worksheet.UsedRange.Columns.Count;
            int num = 0;
            if (maxcols != 11) //неправильное кол-во колонок в файле
            {
                application.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
                application = null;
                workbook = null;
                worksheet = null;
                System.GC.Collect();
                return -1;
            }
            Excel.Range cellRange;
            string[] array = new string[] { "null", "null", "null", "null", "null", "null", "null", "null", "null", "null", "null" };
            for (int i = 1; i < maxrows; i++)
            {
                
                for (int j = 0; j < maxcols; j++)
                {

                    cellRange = (Excel.Range)worksheet.Cells[i+1, j+1];
                    if (cellRange.Value != null)
                    {
                        array[j] = cellRange.Value.ToString();
                    }
                }


                int lastOrderinfo = db.OrderInfos.Max(u => u.Id) + 1;

                string name = array[2];

                //звериный костылина потому что LINQ не хочет работать
                List<string> namelst = (from preparationsBase in db.PreparationsBases
                                    select preparationsBase.Name).ToList();
                List<double> pricelst = (from preparationsBase in db.PreparationsBases
                                         select preparationsBase.Price).ToList();
                double price = -1;
                for (int j = 0; j < namelst.Count; j++)
                {
                    if (namelst[j] == name)
                    {
                        price = pricelst[j];
                        break;
                    }
                }
                if (price == -1)
                {
                    application.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
                    application = null;
                    workbook = null;
                    worksheet = null;
                    System.GC.Collect();
                    return -2; //нет такого препарата
                }


                int amount = Int32.Parse(array[3]);
                num++;
                db.Preparations.Add(new Models.Preparation // check for errors
                {
                    Name = array[2],
                    Amount = amount,
                    ExpirationDate = array[4].Split(' ')[0],
                    OrderInfoId = lastOrderinfo,
                    PaymentDate = "none",// status maybe?
                    Total = Helper.GetTotal(amount, price),
                    TotalVAT = Helper.GetTotalVAT(amount, price)
                });
            }

            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
            //application = null;
            //workbook = null;
            //worksheet = null;
            System.GC.Collect();


            application = new Excel.Application();
            application.Visible = false;
            //this.pathx = "C:\Maxim\BMSTU\6\KursachDB\MyProj\MyProj\Files\";
            workbook = application.Workbooks.Open(this.pathx, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            worksheet = (Excel.Worksheet)workbook.Sheets[1];

            string auctionNum = array[9];
            string distr = array[0];

            maxrows = worksheet.UsedRange.Rows.Count;
            int index = 0;
            string kek;
            for (int i = 2; i < maxrows + 1; i++)
            {
                kek = worksheet.Cells[i, 2].Text.ToString();
                if (kek == auctionNum)         //поле 1(2)
                {
                    index = i;
                    break;
                }
            }
            kek = worksheet.Cells[index, 1].Text.ToString();
            if (distr != kek)                              //поле 0(1)
            {
                int maxi = db.Preparations.Max(u => u.Id);
                for (int i = maxi; i > maxi - num; i--)
                {
                    Models.Preparation prep = db.Preparations.Find(i);
                    db.Preparations.Remove(prep);
                }


                application.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
                application = null;
                workbook = null;
                worksheet = null;
                System.GC.Collect();
                return -3;                                  //неверный дистрибьютор
            }

            
            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
            application = null;
            workbook = null;
            worksheet = null;
            System.GC.Collect();



            int distributorId = db.Distributors.Max(u => u.Id) + 1;
            int auctionInfoId = db.AuctionInfos.Max(u => u.Id) + 1;
            int shipmentInfoId = db.ShipmentInfos.Max(u => u.Id) + 1;
            int orderInfoId = db.OrderInfos.Max(u => u.Id) + 1;
            int shipmentId = db.Shipments.Max(u => u.Id) + 1;
            int auctionId = db.Auctions.Max(u => u.Id) + 1;
            int orderId = db.MainOrders.Max(u => u.Id) + 1;

            db.Distributors.Add(new Models.Distributor
            {
                Id = distributorId,
                Name = array[0]
            });

            db.OrderInfos.Add(new Models.OrderInfo  // check for errors
            {
                Id = orderInfoId,
                Date = array[1].Split(' ')[0],
                PreShipmentDate = array[5].Split(' ')[0],
                CustomerName = array[6],
                CustomerLocationArea = array[7],
                CustomerCity = array[8],
                Status = "InProgress"
            });

            db.AuctionInfos.Add(new Models.AuctionInfo
            {
                Id = auctionInfoId,
                AuctionNumber = array[9],
                Date = array[10].Split(' ')[0],
                Status = "OK"
            });

            
            db.ShipmentInfos.Add(new Models.ShipmentInfo
            {
                Id = shipmentInfoId,
                Date = array[5].Split(' ')[0],
                Status = "InProgress",
                DistributorId = distributorId
            });

            
            db.Shipments.Add(new Models.Shipment
            {
                Id = shipmentId,
                ShipmentInfoId = shipmentInfoId
            });

            
            db.Auctions.Add(new Models.Auction
            {
                Id = auctionId,
                AuctionInfoId = auctionInfoId
            });

            db.MainOrders.Add(new Models.MainOrder
            {
                Id = orderId,
                OrderInfoId = orderInfoId,
                ShipmentId = shipmentId,
                AuctionId = auctionId
            });

            db.SaveChanges();
            return 0;
        }
    }
}