using System;
using System.Collections.Generic;
using System.IO;
using AfterServiceHelper.Models;
using OfficeOpenXml;

namespace AfterServiceHelper
{
    public class ShippedDetailExcelParser
    {
        private readonly string _path;

        public ShippedDetailExcelParser(string path)
        {
            _path = path;
        }

        public List<ShippedDetail> Parsing()
        {
            var details = new List<ShippedDetail>();

            var file = new FileInfo(_path);
            using (var p = new ExcelPackage(file))
            {
                var workSheet = p.Workbook.Worksheets["出貨日期"];
                for (var rowIndex = 2; rowIndex <= workSheet.Dimension.Rows; rowIndex++)
                {
                    details.Add(new ShippedDetail
                    {
                        ProductType = workSheet.Cells[rowIndex, 1].Text,
                        RobotSn = workSheet.Cells[rowIndex, 2].Text,
                        ControlBox = workSheet.Cells[rowIndex, 3].Text,
                        StickSn = workSheet.Cells[rowIndex, 4].Text,
                        CustomerCode = workSheet.Cells[rowIndex, 5].Text,
                        ShippedCustmer = workSheet.Cells[rowIndex, 6].Text,
                        ShippedWeek = workSheet.Cells[rowIndex, 7].Text,
                        ShippedDate = workSheet.Cells[rowIndex, 8].GetValue<DateTime?>(),
                        Destination = workSheet.Cells[rowIndex, 9].Text,
                        HmiVersion = workSheet.Cells[rowIndex, 10].Text,
                        PowerBoard = workSheet.Cells[rowIndex, 11].Text,
                        DriverFW = workSheet.Cells[rowIndex, 12].Text,
                        IoVersion = workSheet.Cells[rowIndex, 13].Text,
                        RtxSn = workSheet.Cells[rowIndex, 14].Text,
                        ShippedMonth = workSheet.Cells[rowIndex, 15].GetValue<int?>(),
                        ShippedYear = workSheet.Cells[rowIndex, 16].GetValue<int?>(),
                        Remark = workSheet.Cells[rowIndex, 17].Text,
                        Hw = workSheet.Cells[rowIndex, 18].Text,
                        HwService = workSheet.Cells[rowIndex, 19].Text,
                        Country = workSheet.Cells[rowIndex, 20].Text,
                        Region = workSheet.Cells[rowIndex, 21].Text,
                        CustomerType = workSheet.Cells[rowIndex, 22].Text
                    });
                }

                return details;
            }
        }
    }
}
