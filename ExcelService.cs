using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using DuctOrderApp.Models;

namespace DuctOrderApp.Services
{
    /// <summary>
    /// Handles all Excel read/write operations via ClosedXML.
    /// Fix: replaced XLWorkbook.OpenOrCreate() (doesn't exist in 0.102.x)
    ///      with the correct pattern: new XLWorkbook(path) when file exists.
    /// </summary>
    public class ExcelService
    {
        private const string FILE_NAME     = "Orders.xlsx";
        private const string SHEET_NAME    = "Orders";

        private const int COL_CLIENT       = 1;
        private const int COL_ORDER_NAME   = 2;
        private const int COL_DUCT_TYPE    = 3;
        private const int COL_URGENT       = 4;
        private const int COL_DATE_CREATED = 5;
        private const int COL_DONE         = 6;

        private readonly string _filePath;

        public ExcelService()
        {
            _filePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                FILE_NAME);
        }

        // ── Public API ───────────────────────────────────────────────────────

        public void EnsureFileExists()
        {
            if (File.Exists(_filePath)) return;

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add(SHEET_NAME);
                WriteHeaders(ws);
                wb.SaveAs(_filePath);
            }
        }

        public void SaveOrder(OrderModel order)
        {
            EnsureFileExists();

            // FIX: use new XLWorkbook(path) to open existing file
            using (var wb = new XLWorkbook(_filePath))
            {
                var ws      = GetOrCreateSheet(wb);
                int nextRow = GetNextEmptyRow(ws);

                ws.Cell(nextRow, COL_CLIENT).Value       = order.Client;
                ws.Cell(nextRow, COL_ORDER_NAME).Value   = order.OrderName;
                ws.Cell(nextRow, COL_DUCT_TYPE).Value    = order.DuctType;
                ws.Cell(nextRow, COL_URGENT).Value       = order.Urgent  ? "YES" : "NO";
                ws.Cell(nextRow, COL_DATE_CREATED).Value = order.DateCreated;
                ws.Cell(nextRow, COL_DONE).Value         = order.Done    ? "YES" : "NO";

                wb.Save();
            }
        }

        public List<OrderModel> LoadOrders()
        {
            var orders = new List<OrderModel>();
            if (!File.Exists(_filePath)) return orders;

            // FIX: use new XLWorkbook(path) to open existing file
            using (var wb = new XLWorkbook(_filePath))
            {
                var ws      = GetOrCreateSheet(wb);
                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

                for (int row = 2; row <= lastRow; row++)
                {
                    string client = ws.Cell(row, COL_CLIENT).GetString();
                    if (string.IsNullOrWhiteSpace(client)) continue;

                    orders.Add(new OrderModel
                    {
                        Client      = client,
                        OrderName   = ws.Cell(row, COL_ORDER_NAME).GetString(),
                        DuctType    = ws.Cell(row, COL_DUCT_TYPE).GetString(),
                        Urgent      = ws.Cell(row, COL_URGENT).GetString()
                                        .Equals("YES", StringComparison.OrdinalIgnoreCase),
                        DateCreated = ws.Cell(row, COL_DATE_CREATED).GetString(),
                        Done        = ws.Cell(row, COL_DONE).GetString()
                                        .Equals("YES", StringComparison.OrdinalIgnoreCase)
                    });
                }
            }

            return orders;
        }

        public void OverwriteOrders(List<OrderModel> orders)
        {
            // FIX: use new XLWorkbook(path) to open existing file
            using (var wb = new XLWorkbook(_filePath))
            {
                if (wb.Worksheets.Contains(SHEET_NAME))
                    wb.Worksheets.Delete(SHEET_NAME);

                var ws = wb.Worksheets.Add(SHEET_NAME);
                WriteHeaders(ws);

                int row = 2;
                foreach (var o in orders)
                {
                    ws.Cell(row, COL_CLIENT).Value       = o.Client;
                    ws.Cell(row, COL_ORDER_NAME).Value   = o.OrderName;
                    ws.Cell(row, COL_DUCT_TYPE).Value    = o.DuctType;
                    ws.Cell(row, COL_URGENT).Value       = o.Urgent ? "YES" : "NO";
                    ws.Cell(row, COL_DATE_CREATED).Value = o.DateCreated;
                    ws.Cell(row, COL_DONE).Value         = o.Done   ? "YES" : "NO";
                    row++;
                }

                wb.Save();
            }
        }

        // ── Helpers ──────────────────────────────────────────────────────────

        private IXLWorksheet GetOrCreateSheet(XLWorkbook wb)
        {
            if (!wb.Worksheets.Contains(SHEET_NAME))
            {
                var ws = wb.Worksheets.Add(SHEET_NAME);
                WriteHeaders(ws);
                return ws;
            }
            return wb.Worksheet(SHEET_NAME);
        }

        private void WriteHeaders(IXLWorksheet ws)
        {
            ws.Cell(1, COL_CLIENT).Value       = "CLIENT";
            ws.Cell(1, COL_ORDER_NAME).Value   = "ORDER NAME";
            ws.Cell(1, COL_DUCT_TYPE).Value    = "DUCT TYPE";
            ws.Cell(1, COL_URGENT).Value       = "URGENT";
            ws.Cell(1, COL_DATE_CREATED).Value = "DATE CREATED";
            ws.Cell(1, COL_DONE).Value         = "DONE";

            var headerRow = ws.Row(1);
            headerRow.Style.Font.Bold                 = true;
            headerRow.Style.Fill.BackgroundColor      = XLColor.FromHtml("#1E88E5");
            headerRow.Style.Font.FontColor            = XLColor.White;
            headerRow.Style.Alignment.Horizontal      = XLAlignmentHorizontalValues.Center;

            ws.Columns().AdjustToContents();
        }

        private int GetNextEmptyRow(IXLWorksheet ws)
        {
            var lastRow = ws.LastRowUsed();
            return (lastRow == null) ? 2 : lastRow.RowNumber() + 1;
        }
    }
}
