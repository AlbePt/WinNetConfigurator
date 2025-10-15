using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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

        public List<Device> ImportDevices(string filePath)
        {
            var devices = new List<Device>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var sheet = workbook.Worksheet(1);
                var usedRange = sheet.RangeUsed();
                if (usedRange == null)
                    return devices;

                foreach (var row in usedRange.RowsUsed().Skip(1))
                {
                    var device = new Device
                    {
                        Type = ParseType(row.Cell(1).GetString()),
                        CabinetName = row.Cell(2).GetString().Trim(),
                        Name = row.Cell(3).GetString().Trim(),
                        IpAddress = row.Cell(4).GetString().Trim(),
                        MacAddress = row.Cell(5).GetString().Trim(),
                        Description = row.Cell(6).GetString().Trim(),
                        AssignedAt = ParseAssignedAt(row.Cell(7).GetString())
                    };

                    if (string.IsNullOrWhiteSpace(device.CabinetName) || string.IsNullOrWhiteSpace(device.IpAddress))
                        continue;

                    devices.Add(device);
                }
            }

            return devices;
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

        DeviceType ParseType(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return DeviceType.Workstation;

            value = value.Trim();

            if (string.Equals(value, "Принтер", StringComparison.OrdinalIgnoreCase))
                return DeviceType.Printer;

            if (string.Equals(value, "Другое устройство", StringComparison.OrdinalIgnoreCase))
                return DeviceType.Other;

            if (string.Equals(value, "Рабочее место", StringComparison.OrdinalIgnoreCase))
                return DeviceType.Workstation;

            if (Enum.TryParse<DeviceType>(value, true, out var parsed))
                return parsed;

            return DeviceType.Workstation;
        }

        DateTime ParseAssignedAt(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return DateTime.Now;

            var trimmed = value.Trim();
            if (DateTime.TryParseExact(trimmed, "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var exact))
                return exact;

            if (DateTime.TryParse(trimmed, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out var parsed))
                return parsed;

            return DateTime.Now;
        }
    }
}
