using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using System.Data.SQLite;
using WinNetConfigurator.Models;
using Newtonsoft.Json;

namespace WinNetConfigurator.Services
{
    public class DbService
    {
        readonly string _dbPath;
        readonly string _connStr;

        public DbService(string dbFileName = "school_net.db")
        {
            _dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, dbFileName);
            _connStr = $"Data Source={_dbPath};Version=3;";
            EnsureDatabase();
            EnsureLeasesFromMachines(); // синхронизация аренды IP на старте
        }

        void EnsureDatabase()
        {
            if (!File.Exists(_dbPath))
                SQLiteConnection.CreateFile(_dbPath);

            using (var con = new SQLiteConnection(_connStr))
            {
                con.Open();
                using (var cmd = con.CreateCommand())
                {
                    cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS app_settings (
  id INTEGER PRIMARY KEY CHECK (id = 1),
  pool_start TEXT NOT NULL,
  pool_end   TEXT NOT NULL,
  netmask    TEXT NOT NULL,
  gateway    TEXT NOT NULL,
  dns1       TEXT NOT NULL,
  dns2       TEXT,
  proxy_hostport TEXT,
  proxy_bypass   TEXT,
  proxy_global_on INTEGER NOT NULL DEFAULT 0
);

INSERT OR IGNORE INTO app_settings
  (id, pool_start, pool_end, netmask, gateway, dns1, dns2, proxy_hostport, proxy_bypass, proxy_global_on)
VALUES
  (1,  '',         '',       '',      '',      '',   '',   '',            '',           0);

CREATE TABLE IF NOT EXISTS cabinets (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS machines (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  cabinet_id INTEGER NOT NULL REFERENCES cabinets(id) ON DELETE CASCADE,
  hostname TEXT NOT NULL,
  mac TEXT NOT NULL,
  adapter_name TEXT NOT NULL,
  ip TEXT NOT NULL UNIQUE,
  proxy_on INTEGER NOT NULL DEFAULT 0,
  assigned_at TEXT NOT NULL,
  source TEXT NOT NULL DEFAULT 'auto'
);

CREATE TABLE IF NOT EXISTS ip_leases (
  ip TEXT PRIMARY KEY,
  cabinet_id INTEGER REFERENCES cabinets(id) ON DELETE SET NULL,
  machine_id INTEGER REFERENCES machines(id) ON DELETE SET NULL,
  status TEXT NOT NULL,                  -- 'free' | 'assigned'
  updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS audit_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  event TEXT NOT NULL,
  details TEXT,
  created_at TEXT NOT NULL
);
";
                    cmd.ExecuteNonQuery();
                }
            }
        }

        SQLiteConnection Open()
        {
            var c = new SQLiteConnection(_connStr);
            c.Open();
            return c;
        }

        // ---------- Settings ----------

        public AppSettings LoadSettings()
        {
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"SELECT pool_start, pool_end, netmask, gateway, dns1, dns2, proxy_hostport, proxy_bypass, proxy_global_on
                                    FROM app_settings WHERE id=1";
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
                            Dns2 = r.IsDBNull(5) ? "" : r.GetString(5),
                            ProxyHostPort = r.IsDBNull(6) ? "" : r.GetString(6),
                            ProxyBypass = r.IsDBNull(7) ? "" : r.GetString(7),
                            ProxyGlobalOn = r.GetInt32(8) != 0
                        };
                    }
                }
            }
            return new AppSettings();
        }

        public void SaveSettings(AppSettings s)
        {
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"UPDATE app_settings SET 
                    pool_start=@a, pool_end=@b, netmask=@c, gateway=@d, dns1=@e, dns2=@f, proxy_hostport=@g, proxy_bypass=@h, proxy_global_on=@i
                    WHERE id=1";
                cmd.Parameters.AddWithValue("@a", s.PoolStart ?? "");
                cmd.Parameters.AddWithValue("@b", s.PoolEnd ?? "");
                cmd.Parameters.AddWithValue("@c", s.Netmask ?? "");
                cmd.Parameters.AddWithValue("@d", s.Gateway ?? "");
                cmd.Parameters.AddWithValue("@e", s.Dns1 ?? "");
                cmd.Parameters.AddWithValue("@f", s.Dns2 ?? "");
                cmd.Parameters.AddWithValue("@g", s.ProxyHostPort ?? "");
                cmd.Parameters.AddWithValue("@h", s.ProxyBypass ?? "");
                cmd.Parameters.AddWithValue("@i", s.ProxyGlobalOn ? 1 : 0);
                cmd.ExecuteNonQuery();
            }
        }

        // ---------- Cabinets ----------

        public int EnsureCabinet(string name)
        {
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            using (var cmd = con.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "SELECT id FROM cabinets WHERE name=@n";
                cmd.Parameters.AddWithValue("@n", name);
                var id = cmd.ExecuteScalar();
                if (id != null && id != DBNull.Value) { tx.Commit(); return Convert.ToInt32(id); }

                cmd.Parameters.Clear();
                cmd.CommandText = "INSERT INTO cabinets(name) VALUES(@n); SELECT last_insert_rowid();";
                cmd.Parameters.AddWithValue("@n", name);
                var newId = Convert.ToInt32(cmd.ExecuteScalar());
                tx.Commit();
                return newId;
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
                        list.Add(new Cabinet { Id = r.GetInt32(0), Name = r.GetString(1) });
                    }
                }
            }
            return list;
        }

        // ---------- Machines / Leases ----------

        public void UpsertIpLease(string ip, int? cabinetId, int? machineId, string status)
        {
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            using (var cmd = con.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = @"INSERT OR IGNORE INTO ip_leases(ip, cabinet_id, machine_id, status, updated_at)
                                    VALUES(@ip, @cab, @mach, @st, @ts)";
                cmd.Parameters.AddWithValue("@ip", ip);
                cmd.Parameters.AddWithValue("@cab", (object)cabinetId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@mach", (object)machineId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@st", status);
                cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                cmd.ExecuteNonQuery();

                cmd.Parameters.Clear();
                cmd.CommandText = @"UPDATE ip_leases
                                    SET cabinet_id=@cab, machine_id=@mach, status=@st, updated_at=@ts
                                    WHERE ip=@ip";
                cmd.Parameters.AddWithValue("@cab", (object)cabinetId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@mach", (object)machineId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@st", status);
                cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                cmd.Parameters.AddWithValue("@ip", ip);
                cmd.ExecuteNonQuery();

                tx.Commit();
            }
        }

        public void InsertMachine(Machine m)
        {
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"INSERT INTO machines(cabinet_id, hostname, mac, adapter_name, ip, proxy_on, assigned_at, source)
                                        VALUES(@cab, @h, @mac, @ad, @ip, @p, @ts, @src);
                                        SELECT last_insert_rowid();";
                    cmd.Parameters.AddWithValue("@cab", m.CabinetId);
                    cmd.Parameters.AddWithValue("@h", m.Hostname);
                    cmd.Parameters.AddWithValue("@mac", m.Mac);
                    cmd.Parameters.AddWithValue("@ad", m.AdapterName);
                    cmd.Parameters.AddWithValue("@ip", m.Ip);
                    cmd.Parameters.AddWithValue("@p", m.ProxyOn ? 1 : 0);
                    cmd.Parameters.AddWithValue("@ts", m.AssignedAt.ToString("o"));
                    cmd.Parameters.AddWithValue("@src", m.Source ?? "auto");
                    m.Id = Convert.ToInt32(cmd.ExecuteScalar());
                }

                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"INSERT OR IGNORE INTO ip_leases(ip, cabinet_id, machine_id, status, updated_at)
                                        VALUES(@ip, @cab, @mach, 'assigned', @ts)";
                    cmd.Parameters.AddWithValue("@ip", m.Ip);
                    cmd.Parameters.AddWithValue("@cab", m.CabinetId);
                    cmd.Parameters.AddWithValue("@mach", m.Id);
                    cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                    cmd.ExecuteNonQuery();

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"UPDATE ip_leases
                                        SET cabinet_id=@cab, machine_id=@mach, status='assigned', updated_at=@ts
                                        WHERE ip=@ip";
                    cmd.Parameters.AddWithValue("@cab", m.CabinetId);
                    cmd.Parameters.AddWithValue("@mach", m.Id);
                    cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@ip", m.Ip);
                    cmd.ExecuteNonQuery();
                }

                tx.Commit();
            }
        }

        /// <summary>
        /// Перегрузка, которую ждёт MachineEditForm: меняем кабинет и IP.
        /// Остальные поля оставляем без изменений.
        /// </summary>
        public void UpdateMachine(int id, int cabinetId, string newIp)
        {
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                string oldIp = null;

                // 1) читаем старый ip
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "SELECT ip FROM machines WHERE id=@id";
                    cmd.Parameters.AddWithValue("@id", id);
                    var o = cmd.ExecuteScalar();
                    oldIp = o == null || o == DBNull.Value ? null : Convert.ToString(o);
                }

                // 2) обновляем кабинет и IP
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"UPDATE machines
                                        SET cabinet_id=@cab, ip=@ip, assigned_at=@ts
                                        WHERE id=@id";
                    cmd.Parameters.AddWithValue("@cab", cabinetId);
                    cmd.Parameters.AddWithValue("@ip", newIp);
                    cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }

                // 3) отметить новый IP как assigned
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"INSERT OR IGNORE INTO ip_leases(ip, cabinet_id, machine_id, status, updated_at)
                                        VALUES(@ip, @cab, @id, 'assigned', @ts)";
                    cmd.Parameters.AddWithValue("@ip", newIp);
                    cmd.Parameters.AddWithValue("@cab", cabinetId);
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                    cmd.ExecuteNonQuery();

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"UPDATE ip_leases
                                        SET cabinet_id=@cab, machine_id=@id, status='assigned', updated_at=@ts
                                        WHERE ip=@ip";
                    cmd.Parameters.AddWithValue("@cab", cabinetId);
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@ip", newIp);
                    cmd.ExecuteNonQuery();
                }

                // 4) если ip изменился — старый освободить
                if (!string.IsNullOrWhiteSpace(oldIp) && !oldIp.Equals(newIp, StringComparison.OrdinalIgnoreCase))
                {
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"UPDATE ip_leases
                                            SET machine_id=NULL, status='free', updated_at=@ts
                                            WHERE ip=@old AND (machine_id IS NULL OR machine_id=@id)";
                        cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                        cmd.Parameters.AddWithValue("@old", oldIp);
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.ExecuteNonQuery();
                    }
                }

                tx.Commit();
            }
        }

        /// <summary>
        /// Полное обновление записи машины (если редактируется больше полей).
        /// </summary>
        public void UpdateMachine(Machine m)
        {
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                string oldIp = null;
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "SELECT ip FROM machines WHERE id=@id";
                    cmd.Parameters.AddWithValue("@id", m.Id);
                    var o = cmd.ExecuteScalar();
                    oldIp = o == null || o == DBNull.Value ? null : Convert.ToString(o);
                }

                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"UPDATE machines
                                        SET cabinet_id=@cab, hostname=@h, mac=@mac, adapter_name=@ad,
                                            ip=@ip, proxy_on=@p, assigned_at=@ts, source=@src
                                        WHERE id=@id";
                    cmd.Parameters.AddWithValue("@cab", m.CabinetId);
                    cmd.Parameters.AddWithValue("@h", m.Hostname);
                    cmd.Parameters.AddWithValue("@mac", m.Mac);
                    cmd.Parameters.AddWithValue("@ad", m.AdapterName);
                    cmd.Parameters.AddWithValue("@ip", m.Ip);
                    cmd.Parameters.AddWithValue("@p", m.ProxyOn ? 1 : 0);
                    cmd.Parameters.AddWithValue("@ts", m.AssignedAt.ToString("o"));
                    cmd.Parameters.AddWithValue("@src", m.Source ?? "auto");
                    cmd.Parameters.AddWithValue("@id", m.Id);
                    cmd.ExecuteNonQuery();
                }

                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"INSERT OR IGNORE INTO ip_leases(ip, cabinet_id, machine_id, status, updated_at)
                                        VALUES(@ip, @cab, @mach, 'assigned', @ts)";
                    cmd.Parameters.AddWithValue("@ip", m.Ip);
                    cmd.Parameters.AddWithValue("@cab", m.CabinetId);
                    cmd.Parameters.AddWithValue("@mach", m.Id);
                    cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                    cmd.ExecuteNonQuery();

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"UPDATE ip_leases
                                        SET cabinet_id=@cab, machine_id=@mach, status='assigned', updated_at=@ts
                                        WHERE ip=@ip";
                    cmd.Parameters.AddWithValue("@cab", m.CabinetId);
                    cmd.Parameters.AddWithValue("@mach", m.Id);
                    cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@ip", m.Ip);
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrWhiteSpace(oldIp) && !oldIp.Equals(m.Ip, StringComparison.OrdinalIgnoreCase))
                {
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"UPDATE ip_leases
                                            SET machine_id=NULL, status='free', updated_at=@ts
                                            WHERE ip=@old AND (machine_id IS NULL OR machine_id=@id)";
                        cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                        cmd.Parameters.AddWithValue("@old", oldIp);
                        cmd.Parameters.AddWithValue("@id", m.Id);
                        cmd.ExecuteNonQuery();
                    }
                }

                tx.Commit();
            }
        }

        public void DeleteMachine(int id)
        {
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                string ip = null;
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "SELECT ip FROM machines WHERE id=@id";
                    cmd.Parameters.AddWithValue("@id", id);
                    var o = cmd.ExecuteScalar();
                    ip = o == null || o == DBNull.Value ? null : Convert.ToString(o);
                }

                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "DELETE FROM machines WHERE id=@id";
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }

                if (!string.IsNullOrWhiteSpace(ip))
                {
                    using (var cmd = con.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"UPDATE ip_leases
                                            SET machine_id=NULL, status='free', updated_at=@ts
                                            WHERE ip=@ip";
                        cmd.Parameters.AddWithValue("@ip", ip);
                        cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                        cmd.ExecuteNonQuery();
                    }
                }

                tx.Commit();
            }
        }

        public Machine GetMachine(int id)
        {
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT m.id, m.cabinet_id, c.name, m.hostname, m.mac, m.adapter_name, m.ip, m.proxy_on, m.assigned_at, m.source
FROM machines m
JOIN cabinets c ON c.id = m.cabinet_id
WHERE m.id=@id";
                cmd.Parameters.AddWithValue("@id", id);

                using (var r = cmd.ExecuteReader())
                {
                    if (r.Read())
                    {
                        return new Machine
                        {
                            Id = r.GetInt32(0),
                            CabinetId = r.GetInt32(1),
                            CabinetName = r.GetString(2),
                            Hostname = r.GetString(3),
                            Mac = r.GetString(4),
                            AdapterName = r.GetString(5),
                            Ip = r.GetString(6),
                            ProxyOn = r.GetInt32(7) != 0,
                            AssignedAt = DateTime.Parse(r.GetString(8)),
                            Source = r.GetString(9)
                        };
                    }
                }
            }
            return null;
        }

        public int CountCabinetMachines(int cabinetId)
        {
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(*) FROM machines WHERE cabinet_id=@c";
                cmd.Parameters.AddWithValue("@c", cabinetId);
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        public HashSet<string> BusyIps()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"SELECT ip FROM machines
                                    UNION
                                    SELECT ip FROM ip_leases WHERE status='assigned'";
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) set.Add(r.GetString(0));
            }
            return set;
        }

        public List<Machine> ListMachines(string cabinetFilter = null)
        {
            var list = new List<Machine>();
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT m.id, m.cabinet_id, c.name, m.hostname, m.mac, m.adapter_name, m.ip, m.proxy_on, m.assigned_at, m.source
FROM machines m
JOIN cabinets c ON c.id = m.cabinet_id
" + (string.IsNullOrEmpty(cabinetFilter) ? "" : " WHERE c.name=@f") + @"
ORDER BY datetime(m.assigned_at) DESC, m.id DESC";
                if (!string.IsNullOrEmpty(cabinetFilter))
                    cmd.Parameters.AddWithValue("@f", cabinetFilter);

                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        list.Add(new Machine
                        {
                            Id = r.GetInt32(0),
                            CabinetId = r.GetInt32(1),
                            CabinetName = r.GetString(2),
                            Hostname = r.GetString(3),
                            Mac = r.GetString(4),
                            AdapterName = r.GetString(5),
                            Ip = r.GetString(6),
                            ProxyOn = r.GetInt32(7) != 0,
                            AssignedAt = DateTime.Parse(r.GetString(8)),
                            Source = r.GetString(9)
                        });
                    }
                }
            }
            return list;
        }

        public void AddAudit(string evt, object details)
        {
            var json = details == null ? null : JsonConvert.SerializeObject(details);
            using (var con = Open())
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "INSERT INTO audit_log(event, details, created_at) VALUES(@e, @d, @ts)";
                cmd.Parameters.AddWithValue("@e", evt);
                cmd.Parameters.AddWithValue("@d", (object)json ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"));
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Гарантируем, что все IP из machines отражены в ip_leases как 'assigned'.
        /// </summary>
        public void EnsureLeasesFromMachines()
        {
            using (var con = Open())
            using (var tx = con.BeginTransaction())
            {
                using (var select = con.CreateCommand())
                {
                    select.Transaction = tx;
                    select.CommandText = @"SELECT m.ip, m.cabinet_id, m.id FROM machines m";
                    using (var r = select.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            var ip = r.GetString(0);
                            var cab = r.GetInt32(1);
                            var mid = r.GetInt32(2);
                            var ts = DateTime.UtcNow.ToString("o");

                            using (var cmd = con.CreateCommand())
                            {
                                cmd.Transaction = tx;
                                cmd.CommandText = @"INSERT OR IGNORE INTO ip_leases(ip, cabinet_id, machine_id, status, updated_at)
                                                    VALUES(@ip, @cab, @mid, 'assigned', @ts)";
                                cmd.Parameters.AddWithValue("@ip", ip);
                                cmd.Parameters.AddWithValue("@cab", cab);
                                cmd.Parameters.AddWithValue("@mid", mid);
                                cmd.Parameters.AddWithValue("@ts", ts);
                                cmd.ExecuteNonQuery();

                                cmd.Parameters.Clear();
                                cmd.CommandText = @"UPDATE ip_leases
                                                    SET cabinet_id=@cab, machine_id=@mid, status='assigned', updated_at=@ts
                                                    WHERE ip=@ip";
                                cmd.Parameters.AddWithValue("@cab", cab);
                                cmd.Parameters.AddWithValue("@mid", mid);
                                cmd.Parameters.AddWithValue("@ts", ts);
                                cmd.Parameters.AddWithValue("@ip", ip);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
                tx.Commit();
            }
        }
    }
}
