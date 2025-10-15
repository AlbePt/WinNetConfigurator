using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class ExcelExportService
    {
        public void ExportDevices(IEnumerable<Device> devices, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("Устройства");
                sheet.Cell(1, 1).Value = "Тип устройства";
                sheet.Cell(1, 2).Value = "Кабинет";
                sheet.Cell(1, 3).Value = "Название";
                sheet.Cell(1, 4).Value = "IP адрес";
                sheet.Cell(1, 5).Value = "MAC адрес";
                sheet.Cell(1, 6).Value = "Описание";
                sheet.Cell(1, 7).Value = "Закреплено";

                int row = 2;
                foreach (var device in devices)
                {
                    sheet.Cell(row, 1).Value = TranslateType(device.Type);
                    sheet.Cell(row, 2).Value = device.CabinetName;
                    sheet.Cell(row, 3).Value = device.Name;
                    sheet.Cell(row, 4).Value = device.IpAddress;
                    sheet.Cell(row, 5).Value = device.MacAddress;
                    sheet.Cell(row, 6).Value = device.Description;
                    sheet.Cell(row, 7).Value = device.AssignedAt.ToString("dd.MM.yyyy HH:mm");
                    row++;
                }

                var range = sheet.Range(1, 1, row - 1, 7);
                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                sheet.Columns().AdjustToContents();

                workbook.SaveAs(filePath);
            }
        }

        string TranslateType(DeviceType type)
        {
            switch (type)
            {
                case DeviceType.Printer:
                    return "Принтер";
                case DeviceType.Other:
                    return "Другое устройство";
                default:
                    return "Рабочее место";
            }
        }
    }
}
