using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;
using Traceablility01.DataReading;
using ExcelPackage = OfficeOpenXml.ExcelPackage;
using ExcelWorksheet = OfficeOpenXml.ExcelWorksheet;

namespace SavingData
{
    class Program
    { 
        public static void SaveToXLSX(string filePath, Report data,string document, string startDate, string endDate)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var file = new FileInfo(@filePath);
            SaveExcelFile(data,file,document, startDate,endDate);
        }

        private static void SaveExcelFile( Report reportData, FileInfo file,string document,string startDate,string endDate)
        {
            DeleteIfExists(file);
            using var packge = new ExcelPackage(file);
            var ws = packge.Workbook.Worksheets.Add($"{document}");
            ExcelRangeBase range;
            range = ws.Cells["A1"].LoadFromCollection(reportData.ReportList, false);
            Sheet01(range);
            HeadSheet01(ws);
            FormatHead(ws, "A1:M1",range);
            OrderBatchs(packge, document, startDate, endDate);
            packge.Save();
        }
        private static void OrderBatchs(ExcelPackage packge, string document, string startDate, string endDate) 
        {
            var OrderList=new ProdOrderList( startDate, endDate, document);
            foreach(var order in OrderList.ProdOrderItems)
            {
                var ws = packge.Workbook.Worksheets.Add($"{order.ProdBatch}");
                MakeProdOrderWorksheet(ws, order.ProdLabel,startDate,endDate);
            }
            
        }
        private static void MakeProdOrderWorksheet(ExcelWorksheet ws, string document, string startDate, string endDate) 
        {
            var prodBatchList=new ProdBatchList(startDate, endDate, document);
            ExcelRangeBase range;
            range = ws.Cells["A1"].LoadFromCollection(prodBatchList.prodBatchList, true);
            Sheet01(range);
            HeadSheet02(ws);
            FormatHead(ws, "A1:N1",range);
        } 
        private static void FormatHead(ExcelWorksheet ws, string adress, ExcelRangeBase range)
        {
            if (!(range==null))
            {
                Color colFromHex = ColorTranslator.FromHtml("#B7DEE8");
                ws.Cells[adress].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[adress].Style.Fill.BackgroundColor.SetColor(colFromHex);
                ws.Cells[range.LocalAddress].AutoFilter = true;
                ws.Row(1).Style.Font.Bold = true;
            }
        }
        private static void HeadSheet01(ExcelWorksheet ws) 
        {
            if (!(ws == null))
            {
                ws.Cells["A1"].Value = "Data ważenia";
                ws.Cells["B1"].Value = "Wagowy";
                ws.Cells["C1"].Value = "Paria dostawy";
                ws.Cells["D1"].Value = "Magazyn";
                ws.Cells["E1"].Value = "Ilość";
                ws.Cells["F1"].Value = "Jm";
                ws.Cells["G1"].Value = "Indeks";
                ws.Cells["H1"].Value = "Nazwa";
                ws.Cells["I1"].Value = "Numer PZ";
                ws.Cells["J1"].Value = "Dostawca";
                ws.Cells["K1"].Value = "Zamównienie zakupu";
                ws.Cells["L1"].Value = "Partia dostawcy";
                ws.Cells["M1"].Value = "Wykorzystane do";
            }
        }
        private static void HeadSheet02(ExcelWorksheet ws)
        {
            if (!(ws == null))
            {
                ws.Cells["A1"].Value = "Typ";
                ws.Cells["B1"].Value = "Data ważenia";
                ws.Cells["C1"].Value = "Wagowy";
                ws.Cells["D1"].Value = "Partia surowca";
                ws.Cells["E1"].Value = "Partia wyrobu";
                ws.Cells["F1"].Value = "Ilość";
                ws.Cells["G1"].Value = "Jm";
                ws.Cells["H1"].Value = "Indeks";
                ws.Cells["I1"].Value = "Nazwa";
                ws.Cells["J1"].Value = "Proces";
                ws.Cells["K1"].Value = "Zlecenie";
                ws.Cells["L1"].Value = "Dokument";
                ws.Cells["M1"].Value = "Obiorca";
                ws.Cells["N1"].Value = "Data";
            }

        }

        private static void Sheet01(ExcelRangeBase range) 
        {
            if (!(range==null))
            {
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.AutoFitColumns();
            } 
        }
        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }
    }
}