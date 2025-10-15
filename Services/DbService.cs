using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class DbService
    {
        readonly string _dbPath;
        readonly string _connStr;

        public DbService(string fileName = "school_net.db")
        {
            _dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
            _connStr = $"Data Source={_dbPath};Version=3;";
            EnsureDatabase();
        }

        SQLiteConnection Open()
        {
            var con = new SQLiteConnection(_connStr);
            con.Open();
            return con;
        }

        void EnsureDatabase()
        {
            if (!File.Exists(_dbPath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_dbPath) ?? AppDomain.CurrentDomain.BaseDirectory);
                SQLiteConnection.CreateFile(_dbPath);
            }

            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS app_settings (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    pool_start TEXT NOT NULL DEFAULT '',
    pool_end   TEXT NOT NULL DEFAULT '',
    netmask    TEXT NOT NULL DEFAULT '',
    gateway    TEXT NOT NULL DEFAULT '',
    dns1       TEXT NOT NULL DEFAULT '',
    dns2       TEXT NOT NULL DEFAULT ''
);

INSERT OR IGNORE INTO app_settings(id) VALUES(1);

CREATE TABLE IF NOT EXISTS cabinets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS devices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cabinet_id INTEGER NOT NULL REFERENCES cabinets(id) ON DELETE CASCADE,
    type TEXT NOT NULL,
    name TEXT NOT NULL,
    ip TEXT NOT NULL UNIQUE,
    mac TEXT,
    description TEXT,
    assigned_at TEXT NOT NULL
);
";
                cmd.ExecuteNonQuery();
            }
        }

        public AppSettings LoadSettings()
        {
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT pool_start, pool_end, netmask, gateway, dns1, dns2 FROM app_settings WHERE id=1";
                using (var r = cmd.ExecuteReader())
                {
                    if (r.Read())
                    {
                        return new AppSettings
                        {
                            PoolStart = r.GetString(0),
                            PoolEnd = r.GetString(1),
                            Netmask = r.GetString(2),
                            Gateway = r.GetString(3),
                            Dns1 = r.GetString(4),
                            Dns2 = r.IsDBNull(5) ? string.Empty : r.GetString(5)
                        };
                    }
                }
            }
            return new AppSettings();
        }

        public void SaveSettings(AppSettings settings)
        {
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"UPDATE app_settings SET
                    pool_start=@a, pool_end=@b, netmask=@c, gateway=@d, dns1=@e, dns2=@f
                    WHERE id=1";
                cmd.Parameters.AddWithValue("@a", settings.PoolStart ?? string.Empty);
                cmd.Parameters.AddWithValue("@b", settings.PoolEnd ?? string.Empty);
                cmd.Parameters.AddWithValue("@c", settings.Netmask ?? string.Empty);
                cmd.Parameters.AddWithValue("@d", settings.Gateway ?? string.Empty);
                cmd.Parameters.AddWithValue("@e", settings.Dns1 ?? string.Empty);
                cmd.Parameters.AddWithValue("@f", settings.Dns2 ?? string.Empty);
                cmd.ExecuteNonQuery();
            }
        }

        public List<Cabinet> GetCabinets()
        {
            var list = new List<Cabinet>();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT id, name FROM cabinets ORDER BY name COLLATE NOCASE";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        list.Add(new Cabinet
                        {
                            Id = r.GetInt32(0),
                            Name = r.GetString(1)
                        });
                    }
                }
            }
            return list;
        }

        public int EnsureCabinet(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Название кабинета не может быть пустым", nameof(name));
            name = name.Trim();

            using (var con = Open())
            using (var tx = con.BeginTransaction())
            using (var cmd = con.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "SELECT id FROM cabinets WHERE name=@n";
                cmd.Parameters.AddWithValue("@n", name);
                var existing = cmd.ExecuteScalar();
                if (existing != null && existing != DBNull.Value)
                {
                    tx.Commit();
                    return Convert.ToInt32(existing);
                }

                cmd.Parameters.Clear();
                cmd.CommandText = "INSERT INTO cabinets(name) VALUES(@n); SELECT last_insert_rowid();";
                cmd.Parameters.AddWithValue("@n", name);
                var id = Convert.ToInt32(cmd.ExecuteScalar());
                tx.Commit();
                return id;
            }
        }

        public List<Device> GetDevices()
        {
            var list = new List<Device>();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"SELECT d.id, d.cabinet_id, c.name, d.type, d.name, d.ip, d.mac, d.description, d.assigned_at
                                     FROM devices d
                                     JOIN cabinets c ON c.id = d.cabinet_id
                                     ORDER BY datetime(d.assigned_at) DESC, d.id DESC";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        list.Add(new Device
                        {
                            Id = r.GetInt32(0),
                            CabinetId = r.GetInt32(1),
                            CabinetName = r.GetString(2),
                            Type = Enum.TryParse<DeviceType>(r.GetString(3), out var type) ? type : DeviceType.Workstation,
                            Name = r.GetString(4),
                            IpAddress = r.GetString(5),
                            MacAddress = r.IsDBNull(6) ? string.Empty : r.GetString(6),
                            Description = r.IsDBNull(7) ? string.Empty : r.GetString(7),
                            AssignedAt = DateTime.Parse(r.GetString(8))
                        });
                    }
                }
            }
            return list;
        }

        public Device GetDeviceByIp(string ip)
        {
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"SELECT d.id, d.cabinet_id, c.name, d.type, d.name, d.ip, d.mac, d.description, d.assigned_at
                                     FROM devices d
                                     JOIN cabinets c ON c.id = d.cabinet_id
                                     WHERE d.ip=@ip";
                cmd.Parameters.AddWithValue("@ip", ip);
                using (var r = cmd.ExecuteReader())
                {
                    if (r.Read())
                    {
                        return new Device
                        {
                            Id = r.GetInt32(0),
                            CabinetId = r.GetInt32(1),
                            CabinetName = r.GetString(2),
                            Type = Enum.TryParse<DeviceType>(r.GetString(3), out var type) ? type : DeviceType.Workstation,
                            Name = r.GetString(4),
                            IpAddress = r.GetString(5),
                            MacAddress = r.IsDBNull(6) ? string.Empty : r.GetString(6),
                            Description = r.IsDBNull(7) ? string.Empty : r.GetString(7),
                            AssignedAt = DateTime.Parse(r.GetString(8))
                        };
                    }
                }
            }
            return null;
        }

        public void InsertDevice(Device device)
        {
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            using (var cmd = con.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = @"INSERT INTO devices(cabinet_id, type, name, ip, mac, description, assigned_at)
                                    VALUES(@cab, @type, @name, @ip, @mac, @desc, @ts);
                                    SELECT last_insert_rowid();";
                cmd.Parameters.AddWithValue("@cab", device.CabinetId);
                cmd.Parameters.AddWithValue("@type", device.Type.ToString());
                cmd.Parameters.AddWithValue("@name", device.Name);
                cmd.Parameters.AddWithValue("@ip", device.IpAddress);
                cmd.Parameters.AddWithValue("@mac", string.IsNullOrWhiteSpace(device.MacAddress) ? (object)DBNull.Value : device.MacAddress);
                cmd.Parameters.AddWithValue("@desc", string.IsNullOrWhiteSpace(device.Description) ? (object)DBNull.Value : device.Description);
                cmd.Parameters.AddWithValue("@ts", device.AssignedAt.ToString("o"));
                device.Id = Convert.ToInt32(cmd.ExecuteScalar());
                tx.Commit();
            }
        }

        public void UpdateDevice(Device device)
        {
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"UPDATE devices
                                    SET cabinet_id=@cab, type=@type, name=@name, ip=@ip, mac=@mac, description=@desc, assigned_at=@ts
                                    WHERE id=@id";
                cmd.Parameters.AddWithValue("@cab", device.CabinetId);
                cmd.Parameters.AddWithValue("@type", device.Type.ToString());
                cmd.Parameters.AddWithValue("@name", device.Name);
                cmd.Parameters.AddWithValue("@ip", device.IpAddress);
                cmd.Parameters.AddWithValue("@mac", string.IsNullOrWhiteSpace(device.MacAddress) ? (object)DBNull.Value : device.MacAddress);
                cmd.Parameters.AddWithValue("@desc", string.IsNullOrWhiteSpace(device.Description) ? (object)DBNull.Value : device.Description);
                cmd.Parameters.AddWithValue("@ts", device.AssignedAt.ToString("o"));
                cmd.Parameters.AddWithValue("@id", device.Id);
                cmd.ExecuteNonQuery();
            }
        }

        public bool DeleteDevice(int id)
        {
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM devices WHERE id=@id";
                cmd.Parameters.AddWithValue("@id", id);
                return cmd.ExecuteNonQuery() > 0;
            }
        }

        public void ClearDatabase()
        {
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            using (var cmd = con.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "DELETE FROM devices";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "DELETE FROM cabinets";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE app_settings SET pool_start='', pool_end='', netmask='', gateway='', dns1='', dns2='' WHERE id=1";
                cmd.ExecuteNonQuery();

                tx.Commit();
            }
        }

        public HashSet<string> GetUsedIps()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT ip FROM devices";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        set.Add(r.GetString(0));
                    }
                }
            }
            return set;
        }
    }
}
