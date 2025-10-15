using System;
using System.Linq;
using System.Windows.Forms;
using WinNetConfigurator.Models;
using WinNetConfigurator.Services;
using WinNetConfigurator.Utils;

namespace WinNetConfigurator.Forms
{
    public class MainForm : Form
    {
        readonly TabControl tabs = new TabControl();
        readonly TabPage tabSettings = new TabPage("Настройки") { AutoScroll = true };
        readonly TabPage tabAssign = new TabPage("Назначение IP") { AutoScroll = true };
        readonly TabPage tabInventory = new TabPage("Инвентаризация") { AutoScroll = true };
        readonly TabPage tabDb = new TabPage("База/Экспорт") { AutoScroll = true };

        // Services
        readonly DbService db = new DbService();
        readonly NetworkService netSvc = new NetworkService();
        readonly ProxyService proxySvc = new ProxyService();
        readonly InventoryService invSvc = new InventoryService();
        readonly IpPlanner ipPlanner = new IpPlanner();
        readonly ExcelExportService xlsx = new ExcelExportService();
        readonly AuditService audit;

        // Settings controls
        TextBox tbPoolStart = new TextBox() { Width = 120 };
        TextBox tbPoolEnd = new TextBox() { Width = 120 };
        TextBox tbMask = new TextBox() { Width = 120 };
        TextBox tbGw = new TextBox() { Width = 120 };
        TextBox tbDns1 = new TextBox() { Width = 120 };
        TextBox tbDns2 = new TextBox() { Width = 120 };
        TextBox tbProxyHost = new TextBox() { Width = 180 };
        TextBox tbProxyBypass = new TextBox() { Width = 260 };
        CheckBox cbProxyGlobal = new CheckBox() { Text = "Proxy on/off (глобально)" };
        Button btnSaveSettings = new Button() { Text = "Сохранить настройки", AutoSize = true };

        // Assign controls
        ComboBox cbCabinet = new ComboBox() { Width = 180, DropDownStyle = ComboBoxStyle.DropDown };
        CheckBox cbShowVirtualAssign = new CheckBox() { Text = "Показывать виртуальные/VPN" };
        ListBox lbAdaptersAssign = new ListBox() { Width = 520, Height = 140 };
        CheckBox cbManualIp = new CheckBox() { Text = "Выбрать IP вручную" };
        TextBox tbManualIp = new TextBox() { Width = 160, Enabled = false };
        // Новые для TCP-проб
        CheckBox cbEnableTcpProbe = new CheckBox() { Text = "TCP-проверка портов" };
        TextBox tbProbePorts = new TextBox() { Width = 160, Text = "135,139,445,3389" };

        Button btnSuggestIp = new Button() { Text = "Предложить IP", AutoSize = true };
        Label lblSuggested = new Label() { AutoSize = true, Text = "—" };
        CheckBox cbProxyOn = new CheckBox() { Text = "Прокси on/off (для этой операции)" };
        Button btnApply = new Button() { Text = "Применить", AutoSize = true };

        // Inventory controls
        CheckBox cbShowVirtualInv = new CheckBox() { Text = "Показывать виртуальные/VPN" };
        ListBox lbAdaptersInventory = new ListBox() { Width = 520, Height = 140 };
        Button btnReadCurrent = new Button() { Text = "Считать текущие настройки", AutoSize = true };
        Label lblCurrent = new Label() { AutoSize = true, Text = "—" };
        CheckBox cbInternetOk = new CheckBox() { Text = "Интернет работает (записать в БД без изменений)" };
        Button btnSaveInventory = new Button() { Text = "Записать в БД", AutoSize = true };

        // DB/Export
        Button btnExport = new Button() { Text = "Экспорт в XLSX", AutoSize = true };
        DataGridView grid = new DataGridView() { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = true };

        public MainForm()
        {
            Text = "WinNetConfigurator (offline)";
            Width = 900;
            Height = 600;
            StartPosition = FormStartPosition.CenterScreen;

            audit = new AuditService(db);
            BuildSettingsTab();
            BuildAssignTab();
            BuildInventoryTab();
            BuildDbTab();

            tabs.Dock = DockStyle.Fill;
            tabs.TabPages.AddRange(new[] { tabSettings, tabAssign, tabInventory, tabDb });
            Controls.Add(tabs);

            LoadSettingsToUI();
            RefreshAdaptersAssign();
            RefreshAdaptersInventory();
            RefreshGrid();
        }

        // ---------- SETTINGS ----------
        void BuildSettingsTab()
        {
            tabSettings.Controls.Clear();

            var tl = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Padding = new Padding(12),
            };
            tl.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            void AddRow(string label, Control input)
            {
                var row = tl.RowCount;
                tl.RowCount++;
                var lbl = new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
                input.Margin = new Padding(0, 3, 0, 3);
                input.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                if (input is TextBox tb) tb.Width = 260;
                tl.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tl.Controls.Add(lbl, 0, row);
                tl.Controls.Add(input, 1, row);
            }

            AddRow("Пул start:", tbPoolStart);
            AddRow("Пул end:", tbPoolEnd);
            AddRow("Маска:", tbMask);
            AddRow("Шлюз:", tbGw);
            AddRow("DNS1:", tbDns1);
            AddRow("DNS2:", tbDns2);
            AddRow("Proxy host:port:", tbProxyHost);
            AddRow("Bypass:", tbProxyBypass);

            var rowCb = tl.RowCount; tl.RowCount++;
            tl.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbProxyGlobal.Margin = new Padding(0, 8, 0, 8);
            cbProxyGlobal.Anchor = AnchorStyles.Left;
            tl.Controls.Add(cbProxyGlobal, 0, rowCb);
            tl.SetColumnSpan(cbProxyGlobal, 2);

            var btnCol = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                WrapContents = false,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 6, 0, 0)
            };
            btnCol.Controls.Add(btnSaveSettings);

            var rowBtns = tl.RowCount; tl.RowCount++;
            tl.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tl.Controls.Add(btnCol, 0, rowBtns);
            tl.SetColumnSpan(btnCol, 2);

            tabSettings.Controls.Add(tl);

            btnSaveSettings.Click += (s, e) => SaveSettingsFromUI();
        }

        // ---------- ASSIGN ----------
        void BuildAssignTab()
        {
            tabAssign.Controls.Clear();

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoScroll = true,
                ColumnCount = 2,
                Padding = new Padding(12)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var form = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2
            };
            form.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            void AddRow(string label, Control input)
            {
                var row = form.RowCount; form.RowCount++;
                var lbl = new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
                input.Margin = new Padding(0, 3, 0, 3);
                input.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                form.Controls.Add(lbl, 0, row);
                form.Controls.Add(input, 1, row);
            }

            AddRow("Кабинет:", cbCabinet);

            var rowV = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbShowVirtualAssign.Anchor = AnchorStyles.Left;
            cbShowVirtualAssign.Margin = new Padding(0, 6, 0, 6);
            form.Controls.Add(cbShowVirtualAssign, 0, rowV);
            form.SetColumnSpan(cbShowVirtualAssign, 2);

            var rowA = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var lblAdapters = new Label { Text = "Адаптеры:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
            form.Controls.Add(lblAdapters, 0, rowA);
            lbAdaptersAssign.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            form.Controls.Add(lbAdaptersAssign, 1, rowA);

            var rowM = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var manualPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false, Dock = DockStyle.Fill };
            manualPanel.Controls.Add(cbManualIp);
            manualPanel.Controls.Add(tbManualIp);
            form.Controls.Add(new Label() { Text = "IP:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, rowM);
            form.Controls.Add(manualPanel, 1, rowM);

            // --- Новые поля TCP-проб ---
            var rowTP = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            form.Controls.Add(new Label() { Text = "TCP-порты:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, rowTP);
            form.Controls.Add(tbProbePorts, 1, rowTP);

            var rowTCB = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbEnableTcpProbe.Anchor = AnchorStyles.Left;
            cbEnableTcpProbe.Margin = new Padding(0, 6, 0, 6);
            form.Controls.Add(cbEnableTcpProbe, 0, rowTCB);
            form.SetColumnSpan(cbEnableTcpProbe, 2);
            // --- /новые поля ---

            var rowP = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbProxyOn.Anchor = AnchorStyles.Left;
            cbProxyOn.Margin = new Padding(0, 6, 0, 6);
            form.Controls.Add(cbProxyOn, 0, rowP);
            form.SetColumnSpan(cbProxyOn, 2);

            var btnCol = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                WrapContents = false,
                Dock = DockStyle.Top,
                Margin = new Padding(12, 0, 0, 0)
            };
            btnCol.Controls.Add(btnSuggestIp);
            var suggPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            suggPanel.Controls.Add(new Label { Text = "Предложен:", AutoSize = true });
            suggPanel.Controls.Add(lblSuggested);
            btnCol.Controls.Add(suggPanel);
            btnCol.Controls.Add(btnApply);

            root.Controls.Add(form, 0, 0);
            root.Controls.Add(btnCol, 1, 0);

            tabAssign.Controls.Add(root);

            cbShowVirtualAssign.CheckedChanged += (s, e) => RefreshAdaptersAssign();
            cbManualIp.CheckedChanged += (s, e) => tbManualIp.Enabled = cbManualIp.Checked;
            btnSuggestIp.Click += (s, e) => SuggestIpClick();
            btnApply.Click += (s, e) => ApplyClick();
        }

        // ---------- INVENTORY ----------
        void BuildInventoryTab()
        {
            tabInventory.Controls.Clear();

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoScroll = true,
                ColumnCount = 2,
                Padding = new Padding(12)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var form = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2
            };
            form.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            var rowV = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbShowVirtualInv.Anchor = AnchorStyles.Left;
            cbShowVirtualInv.Margin = new Padding(0, 6, 0, 6);
            form.Controls.Add(cbShowVirtualInv, 0, rowV);
            form.SetColumnSpan(cbShowVirtualInv, 2);

            var rowA = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var lblAdapters = new Label { Text = "Адаптеры:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
            form.Controls.Add(lblAdapters, 0, rowA);
            lbAdaptersInventory.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            form.Controls.Add(lbAdaptersInventory, 1, rowA);

            var rowCur = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            form.Controls.Add(new Label { Text = "Текущие:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, rowCur);
            lblCurrent.Anchor = AnchorStyles.Left;
            form.Controls.Add(lblCurrent, 1, rowCur);

            var btnCol = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                WrapContents = false,
                Dock = DockStyle.Top,
                Margin = new Padding(12, 0, 0, 0)
            };
            btnCol.Controls.Add(btnReadCurrent);
            btnCol.Controls.Add(cbInternetOk);
            btnCol.Controls.Add(btnSaveInventory);

            root.Controls.Add(form, 0, 0);
            root.Controls.Add(btnCol, 1, 0);

            tabInventory.Controls.Add(root);

            cbShowVirtualInv.CheckedChanged += (s, e) => RefreshAdaptersInventory();
            btnReadCurrent.Click += (s, e) => ReadCurrentClick();
            btnSaveInventory.Click += (s, e) => SaveInventoryClick();
        }

        // ---------- DB / EXPORT ----------
        void BuildDbTab()
        {
            tabDb.Controls.Clear();

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoScroll = true,
                ColumnCount = 2,
                Padding = new Padding(12)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            grid.ReadOnly = true;
            layout.Controls.Add(grid, 0, 0);

            var btnCol = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                WrapContents = false,
                Dock = DockStyle.Top,
                Margin = new Padding(12, 0, 0, 0)
            };

            var btnRefresh = new Button { Text = "Обновить", AutoSize = true };
            var btnEdit = new Button { Text = "Редактировать...", AutoSize = true };
            var btnDelete = new Button { Text = "Удалить", AutoSize = true };

            btnCol.Controls.Add(btnRefresh);
            btnCol.Controls.Add(btnEdit);
            btnCol.Controls.Add(btnDelete);
            btnCol.Controls.Add(btnExport);

            layout.Controls.Add(btnCol, 1, 0);
            tabDb.Controls.Add(layout);

            btnRefresh.Click += (s, e) => RefreshGrid();
            btnEdit.Click += (s, e) => EditSelectedMachine();
            btnDelete.Click += (s, e) => DeleteSelectedMachine();
            btnExport.Click += (s, e) => ExportClick();
        }

        // ---------- SETTINGS LOGIC ----------
        void LoadSettingsToUI()
        {
            var s = db.LoadSettings();
            tbPoolStart.Text = s.PoolStart;
            tbPoolEnd.Text = s.PoolEnd;
            tbMask.Text = s.Netmask;
            tbGw.Text = s.Gateway;
            tbDns1.Text = s.Dns1;
            tbDns2.Text = s.Dns2;
            tbProxyHost.Text = s.ProxyHostPort;
            tbProxyBypass.Text = s.ProxyBypass;
            cbProxyGlobal.Checked = s.ProxyGlobalOn;
        }

        void SaveSettingsFromUI()
        {
            if (!Validation.IsValidIPv4(tbPoolStart.Text) || !Validation.IsValidIPv4(tbPoolEnd.Text))
            { MessageBox.Show("Некорректный пул."); return; }
            if (!Validation.IsValidIPv4(tbMask.Text) || !Validation.IsValidIPv4(tbGw.Text))
            { MessageBox.Show("Некорректная маска/шлюз."); return; }
            if (!Validation.IsValidIPv4(tbDns1.Text))
            { MessageBox.Show("Некорректный DNS1."); return; }

            var s = new AppSettings
            {
                PoolStart = tbPoolStart.Text.Trim(),
                PoolEnd = tbPoolEnd.Text.Trim(),
                Netmask = tbMask.Text.Trim(),
                Gateway = tbGw.Text.Trim(),
                Dns1 = tbDns1.Text.Trim(),
                Dns2 = tbDns2.Text.Trim(),
                ProxyHostPort = tbProxyHost.Text.Trim(),
                ProxyBypass = tbProxyBypass.Text.Trim(),
                ProxyGlobalOn = cbProxyGlobal.Checked
            };
            db.SaveSettings(s);
            if (s.ProxyGlobalOn)
            {
                try { proxySvc.SetGlobalProxy(true, s.ProxyHostPort, s.ProxyBypass); }
                catch (Exception ex) { MessageBox.Show("Ошибка применения прокси: " + ex.Message); }
            }
            audit.Log("SAVE_SETTINGS", s);
            MessageBox.Show("Сохранено.");
        }

        // ---------- ADAPTERS ----------
        void RefreshAdaptersAssign() => RefreshAdaptersInto(lbAdaptersAssign, cbShowVirtualAssign.Checked);
        void RefreshAdaptersInventory() => RefreshAdaptersInto(lbAdaptersInventory, cbShowVirtualInv.Checked);

        void RefreshAdaptersInto(ListBox target, bool includeVirtual)
        {
            target.Items.Clear();
            foreach (var a in netSvc.ListAdapters(includeVirtual))
                target.Items.Add($"{a.NetConnectionId} | {a.Name} | {a.MacAddress} | {a.IPv4Address} | {a.Type}");
            if (target.Items.Count > 0) target.SelectedIndex = 0;
        }

        (string adapterId, string adapterDisplay, string mac) GetSelectedAdapterFrom(ListBox listBox)
        {
            if (listBox.SelectedItem == null) { MessageBox.Show("Выберите адаптер."); return ("", "", ""); }
            var line = listBox.SelectedItem.ToString();
            var parts = line.Split('|').Select(x => x.Trim()).ToArray();
            string netId = parts.Length > 0 ? parts[0] : "";
            string name = parts.Length > 1 ? parts[1] : "";
            string mac = parts.Length > 2 ? parts[2] : "";
            return (netId, name, mac);
        }

        // ---------- HELPERS ----------
        int[] ParsePorts(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return Array.Empty<int>();
            return s.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(p => { int v; return int.TryParse(p, out v) ? v : -1; })
                    .Where(v => v > 0 && v <= 65535)
                    .Distinct()
                    .ToArray();
        }

        // ---------- ASSIGN FLOW ----------
        void SuggestIpClick()
        {
            var s = db.LoadSettings();
            if (string.IsNullOrWhiteSpace(s.PoolStart) || string.IsNullOrWhiteSpace(s.PoolEnd))
            { MessageBox.Show("Настройте пул в Настройках."); return; }

            var busy = db.BusyIps();
            var probe = new NetworkProbeService();
            var ports = cbEnableTcpProbe.Checked ? ParsePorts(tbProbePorts.Text) : Array.Empty<int>();

            string suggestion = null;
            foreach (var ip in Utils.IPRange.Enumerate(s.PoolStart, s.PoolEnd))
            {
                var ipStr = ip.ToString();
                if (busy.Contains(ipStr)) continue;            // занято по БД
                if (probe.IsIpInUse(ipStr, ports)) continue;   // занято в сети (ARP/ICMP/TCP)
                suggestion = ipStr; break;
            }
            if (suggestion == null) { MessageBox.Show("Свободных IP не найдено."); return; }
            lblSuggested.Text = suggestion;
        }

        void ApplyClick()
        {
            var s = db.LoadSettings();
            if (string.IsNullOrWhiteSpace(s.PoolStart) || string.IsNullOrWhiteSpace(s.PoolEnd))
            { MessageBox.Show("Настройте пул в Настройках."); return; }

            string cabinetName = string.IsNullOrWhiteSpace(cbCabinet.Text) ? "Не указан" : cbCabinet.Text.Trim();
            int cabId = db.EnsureCabinet(cabinetName);

            int count = db.CountCabinetMachines(cabId);
            if (count >= 3) { MessageBox.Show("В кабинете уже 3 IP. Назначение запрещено."); return; }

            var (adapterId, adapterName, mac) = GetSelectedAdapterFrom(lbAdaptersAssign);
            if (string.IsNullOrEmpty(adapterId)) return;

            string ip = cbManualIp.Checked ? tbManualIp.Text.Trim() : lblSuggested.Text;
            if (!Validation.IsValidIPv4(ip)) { MessageBox.Show("Некорректный IP/нет предложения."); return; }

            try
            {
                netSvc.ApplyIPv4(adapterId, ip, s.Netmask, s.Gateway,
                    string.IsNullOrWhiteSpace(s.Dns2) ? new[] { s.Dns1 } : new[] { s.Dns1, s.Dns2 });

                if (cbProxyOn.Checked) proxySvc.SetGlobalProxy(true, s.ProxyHostPort, s.ProxyBypass);

                var m = new Models.Machine
                {
                    CabinetId = cabId,
                    CabinetName = cabinetName,
                    Hostname = Environment.MachineName,
                    Mac = mac,
                    AdapterName = adapterName,
                    Ip = ip,
                    ProxyOn = cbProxyOn.Checked,
                    AssignedAt = DateTime.Now,
                    Source = "auto"
                };
                db.InsertMachine(m);
                audit.Log("ASSIGN_IP", new { cabinet = cabinetName, ip, adapterId, adapterName, proxy = cbProxyOn.Checked });

                // косметика после успешного применения
                lblSuggested.Text = "—";
                cbManualIp.Checked = false;
                tbManualIp.Enabled = false;

                MessageBox.Show("Настройки применены.");
                RefreshGrid();
            }
            catch (Exception ex)
            {
                db.UpsertIpLease(ip, null, null, "free"); // best-effort rollback
                MessageBox.Show("Ошибка применения: " + ex.Message);
            }
        }

        // ---------- INVENTORY FLOW ----------
        void ReadCurrentClick()
        {
            var (adapterId, adapterName, mac) = GetSelectedAdapterFrom(lbAdaptersInventory);
            if (string.IsNullOrEmpty(adapterId)) return;
            var cur = invSvc.ReadCurrentIPv4(adapterName);
            lblCurrent.Text = $"IP={cur.ip} MASK={cur.mask} GW={cur.gateway} DNS=[{string.Join(",", cur.dns)}]";
        }

        void SaveInventoryClick()
        {
            if (!cbInternetOk.Checked) { MessageBox.Show("Отметьте, что интернет работает, чтобы записать без изменений."); return; }

            var (adapterId, adapterName, mac) = GetSelectedAdapterFrom(lbAdaptersInventory);
            if (string.IsNullOrEmpty(adapterId)) return;
            var cur = invSvc.ReadCurrentIPv4(adapterName);
            if (string.IsNullOrEmpty(cur.ip)) { MessageBox.Show("Не удалось прочитать текущие настройки."); return; }

            string cabinetName = string.IsNullOrWhiteSpace(cbCabinet.Text) ? "Не указан" : cbCabinet.Text.Trim();
            int cabId = db.EnsureCabinet(cabinetName);

            var m = new Models.Machine
            {
                CabinetId = cabId,
                CabinetName = cabinetName,
                Hostname = Environment.MachineName,
                Mac = mac,
                AdapterName = adapterName,
                Ip = cur.ip,
                ProxyOn = false,
                AssignedAt = DateTime.Now,
                Source = "inventory"
            };
            db.InsertMachine(m);
            audit.Log("INVENTORY_SAVE", new { cabinet = cabinetName, ip = cur.ip });
            MessageBox.Show("Сохранено в БД.");
            RefreshGrid();
        }

        // ---------- GRID / EXPORT / EDIT ----------
        void RefreshGrid()
        {
            var items = db.ListMachines();
            grid.DataSource = items;
        }

        Machine GetSelectedMachine()
        {
            if (grid.CurrentRow == null || grid.CurrentRow.DataBoundItem == null)
            {
                MessageBox.Show("Выберите строку в таблице.");
                return null;
            }
            return grid.CurrentRow.DataBoundItem as Machine;
        }

        void EditSelectedMachine()
        {
            var sel = GetSelectedMachine();
            if (sel == null) return;

            var fresh = db.GetMachine(sel.Id) ?? sel;

            using (var dlg = new Forms.MachineEditForm(db, fresh))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    RefreshGrid();
                }
            }
        }

        void DeleteSelectedMachine()
        {
            var sel = GetSelectedMachine();
            if (sel == null) return;

            if (MessageBox.Show($"Удалить запись ID={sel.Id} (IP {sel.Ip})?",
                                "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    db.DeleteMachine(sel.Id);
                    RefreshGrid();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка удаления: " + ex.Message);
                }
            }
        }

        void ExportClick()
        {
            using (var sfd = new SaveFileDialog() { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "export.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    var data = db.ListMachines();
                    var path = xlsx.ExportMachines(sfd.FileName, data);
                    audit.Log("EXPORT_XLSX", new { file = path, count = data.Count });
                    MessageBox.Show("Экспортировано: " + path);
                }
            }
        }
    }
}


