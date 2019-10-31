using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MyProj.Help
{
    public class InvoiceGen
    {
        private string path = "C:/Maxim/BMSTU/6/KursachDB/MyProj/MyProj/Files/test.xlsx";
        public XLWorkbook workbook { get; set; }
        public InvoiceGen(AllInfo allInfo)
        {
            using (XLWorkbook workbook = new XLWorkbook(this.path, XLEventTracking.Disabled))
            {
                IXLWorksheet worksheet;
                workbook.Worksheets.TryGetWorksheet("1", out worksheet);

                worksheet.Cell("A1").Value = "Бренд";
                worksheet.Cell("B1").Value = "Модели";

                this.workbook = workbook;
            }
        }
    }
}