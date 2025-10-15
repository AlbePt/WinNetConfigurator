using System;
using System.Collections.Generic;
using System.Linq;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class InventoryWorkflowService
    {
        readonly Dictionary<string, List<string>> _templates = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase)
        {
            { "Учебный кабинет", new List<string> { "Учительский ПК", "ПК-1", "ПК-2", "ПК-3", "Принтер" } },
            { "Компьютерный класс", new List<string> { "Сервер", "ПК-01", "ПК-02", "ПК-03", "ПК-04", "ПК-05", "Проектор" } },
            { "Лаборатория", new List<string> { "Лаб-терминал", "Контроллер", "ПК оператора" } }
        };

        public IReadOnlyList<string> Templates => _templates.Keys.ToList();

        public InventorySession CreateSession(string template, UserSession owner)
        {
            var session = new InventorySession
            {
                Cabinet = template,
                Owner = owner
            };

            if (!string.IsNullOrWhiteSpace(template) && _templates.TryGetValue(template, out var devices))
            {
                foreach (var device in devices)
                {
                    session.Entries.Add(new InventoryEntry
                    {
                        DeviceId = device.Replace(" ", "_"),
                        DisplayName = device
                    });
                }
            }

            return session;
        }

        public InventorySession CreateCustomSession(string cabinetName, IEnumerable<string> devices, UserSession owner)
        {
            var session = new InventorySession
            {
                Cabinet = cabinetName,
                Owner = owner
            };

            if (devices != null)
            {
                foreach (var device in devices)
                {
                    if (string.IsNullOrWhiteSpace(device)) continue;
                    session.Entries.Add(new InventoryEntry
                    {
                        DeviceId = device.Trim().Replace(" ", "_"),
                        DisplayName = device.Trim()
                    });
                }
            }

            return session;
        }
    }
}
