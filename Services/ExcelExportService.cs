using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class ExcelExportService
    {
        public string ExportMachines(string filePath, IEnumerable<Machine> data)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Export");
                ws.Cell(1,1).Value = "Кабинет";
                ws.Cell(1,2).Value = "Hostname";
                ws.Cell(1,3).Value = "MAC";
                ws.Cell(1,4).Value = "Адаптер";
                ws.Cell(1,5).Value = "IP";
                ws.Cell(1,6).Value = "Прокси (on/off)";
                ws.Cell(1,7).Value = "Дата";

                int row = 2;
                foreach (var m in data)
                {
                    ws.Cell(row,1).Value = m.CabinetName;
                    ws.Cell(row,2).Value = m.Hostname;
                    ws.Cell(row,3).Value = m.Mac;
                    ws.Cell(row,4).Value = m.AdapterName;
                    ws.Cell(row,5).Value = m.Ip;
                    ws.Cell(row,6).Value = m.ProxyOn ? "on" : "off";
                    ws.Cell(row,7).Value = m.AssignedAt.ToString("yyyy-MM-dd HH:mm:ss");
                    row++;
                }
                ws.Columns().AdjustToContents();
                wb.SaveAs(filePath);
            }
            return filePath;
        }
    }
}
