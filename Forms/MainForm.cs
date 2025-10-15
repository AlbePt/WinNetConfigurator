
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
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
        readonly TabPage tabExceptions = new TabPage("Исключения") { AutoScroll = true };

        // Services
        readonly DbService db = new DbService();
        readonly NetworkService netSvc = new NetworkService();
        readonly ProxyService proxySvc = new ProxyService();
        readonly InventoryService invSvc = new InventoryService();
        readonly IpPlanner ipPlanner = new IpPlanner();
        readonly ExcelExportService xlsx = new ExcelExportService();
        readonly AuditService audit;
        readonly RoleAccessService access = new RoleAccessService();
        readonly DraftService draftSvc = new DraftService();
        readonly NotificationService notificationSvc = new NotificationService();
        readonly TaskQueueService taskQueue = new TaskQueueService();
        readonly InventoryWorkflowService inventoryWorkflow = new InventoryWorkflowService();
        readonly IpPolicyService ipPolicy;

        readonly ErrorProvider errorProvider = new ErrorProvider();

        // Settings controls
        readonly TextBox tbPoolStart = new TextBox() { Width = 120 };
        readonly TextBox tbPoolEnd = new TextBox() { Width = 120 };
        readonly TextBox tbMask = new TextBox() { Width = 120 };
        readonly TextBox tbGw = new TextBox() { Width = 120 };
        readonly TextBox tbDns1 = new TextBox() { Width = 120 };
        readonly TextBox tbDns2 = new TextBox() { Width = 120 };
        readonly TextBox tbProxyHost = new TextBox() { Width = 180 };
        readonly TextBox tbProxyBypass = new TextBox() { Width = 260 };
        readonly CheckBox cbProxyGlobal = new CheckBox() { Text = "Proxy on/off (глобально)" };
        readonly Button btnSaveSettings = new Button() { Text = "Сохранить настройки", AutoSize = true };

        // Assignment controls
        readonly ComboBox cbCabinet = new ComboBox() { Width = 180, DropDownStyle = ComboBoxStyle.DropDown };
        readonly CheckBox cbShowVirtualAssign = new CheckBox() { Text = "Показывать виртуальные/VPN" };
        readonly ListBox lbAdaptersAssign = new ListBox() { Width = 520, Height = 140 };
        readonly CheckBox cbManualIp = new CheckBox() { Text = "Выбрать IP вручную" };
        readonly TextBox tbManualIp = new TextBox() { Width = 160, Enabled = false };
        readonly CheckBox cbEnableTcpProbe = new CheckBox() { Text = "TCP-проверка портов" };
        readonly TextBox tbProbePorts = new TextBox() { Width = 160, Text = "135,139,445,3389" };
        readonly Button btnSuggestIp = new Button() { Text = "Предложить IP", AutoSize = true };
        readonly Label lblSuggested = new Label() { AutoSize = true, Text = "—" };
        readonly CheckBox cbProxyOn = new CheckBox() { Text = "Прокси on/off (для этой операции)" };
        readonly Button btnApply = new Button() { Text = "Применить", AutoSize = true };
        readonly TextBox tbAssignReason = new TextBox() { Width = 240 };
        readonly TextBox tbAssignAttributes = new TextBox() { Width = 240, Multiline = true, Height = 60 };
        readonly Label lblPolicyHint = new Label() { AutoSize = true, ForeColor = Color.DimGray };
        readonly ListView lvRiskSummary = new ListView() { Width = 420, Height = 150, View = View.Details, FullRowSelect = true, GridLines = true };
        readonly CheckBox cbAcknowledgeRisks = new CheckBox() { Text = "Я ознакомлен с рисками" };
        readonly Button btnSubmitApproval = new Button() { Text = "Отправить на утверждение", AutoSize = true };
        readonly ListBox lbDrafts = new ListBox() { Width = 220, Height = 160 };
        readonly Button btnNewDraft = new Button() { Text = "Новый черновик", AutoSize = true };
        readonly Label lblDraftMeta = new Label() { AutoSize = true };
        readonly Label lblAutosave = new Label() { AutoSize = true, ForeColor = Color.DarkGreen };

        // Inventory controls
        readonly CheckBox cbShowVirtualInv = new CheckBox() { Text = "Показывать виртуальные/VPN" };
        readonly ListBox lbAdaptersInventory = new ListBox() { Width = 520, Height = 140 };
        readonly Button btnReadCurrent = new Button() { Text = "Считать текущие настройки", AutoSize = true };
        readonly Label lblCurrent = new Label() { AutoSize = true, Text = "—" };
        readonly CheckBox cbInternetOk = new CheckBox() { Text = "Интернет работает (записать в БД без изменений)" };
        readonly Button btnSaveInventory = new Button() { Text = "Записать в БД", AutoSize = true };
        readonly ComboBox cbInventoryTemplate = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        readonly Button btnStartInventorySession = new Button() { Text = "Запустить обход", AutoSize = true };
        readonly ListBox lbInventoryChecklist = new ListBox() { Width = 520, Height = 160 };
        readonly Label lblInventoryStatus = new Label() { AutoSize = true };
        readonly Button btnInventoryMarkOk = new Button() { Text = "Проверено", AutoSize = true };
        readonly Button btnInventoryMarkMissing = new Button() { Text = "Отсутствует", AutoSize = true };
        readonly Button btnInventoryMarkReview = new Button() { Text = "Требует проверки", AutoSize = true };
        readonly Button btnInventoryNext = new Button() { Text = "Следующее", AutoSize = true };
        readonly Button btnInventoryPrev = new Button() { Text = "Предыдущее", AutoSize = true };

        // Export
        readonly Button btnExport = new Button() { Text = "Экспорт", AutoSize = true };
        readonly Button btnExportPreview = new Button() { Text = "Предпросмотр", AutoSize = true };
        readonly CheckedListBox clbExportLocations = new CheckedListBox() { Width = 200, Height = 100 };
        readonly DateTimePicker dtExportFrom = new DateTimePicker() { Format = DateTimePickerFormat.Short };
        readonly DateTimePicker dtExportTo = new DateTimePicker() { Format = DateTimePickerFormat.Short };
        readonly CheckedListBox clbExportStatuses = new CheckedListBox() { Width = 200, Height = 100 };
        readonly TextBox tbExportResponsible = new TextBox() { Width = 200 };
        readonly ComboBox cbExportFormat = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 140 };
        readonly DataGridView grid = new DataGridView() { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = true };

        // Side panels
        readonly ComboBox cbRoleSelector = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };
        readonly TextBox tbUserName = new TextBox() { Width = 180 };
        readonly Button btnSwitchUser = new Button() { Text = "Применить", AutoSize = true };
        readonly Label lblUserSummary = new Label() { AutoSize = true, Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold) };
        readonly ListView lvNotifications = new ListView() { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false };
        readonly ListView lvTaskQueue = new ListView() { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false };
        readonly Button btnCancelSelectedTask = new Button() { Text = "Отменить задачу", AutoSize = true };
        readonly ListView lvExceptions = new ListView() { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false };

        // State
        readonly BindingList<AssignmentDraft> draftBinding = new BindingList<AssignmentDraft>();
        readonly BindingList<Notification> notificationBinding = new BindingList<Notification>();
        readonly BindingList<BackgroundTask> taskBinding = new BindingList<BackgroundTask>();
        readonly BindingList<InventoryEntry> inventoryBinding = new BindingList<InventoryEntry>();

        UserSession currentUser;
        AssignmentDraft activeDraft;
        InventorySession currentInventory;
        bool suppressDraftEvents;
        bool hasUnsavedChanges;

        public MainForm()
        {
            Text = "WinNetConfigurator (offline)";
            Width = 1100;
            Height = 720;
            StartPosition = FormStartPosition.CenterScreen;

            audit = new AuditService(db);
            ipPolicy = new IpPolicyService(db, ipPlanner);
            errorProvider.ContainerControl = this;

            currentUser = new UserSession("operator", "Оператор", UserRole.Operator);

            BuildLayout();
            BuildSettingsTab();
            BuildAssignTab();
            BuildInventoryTab();
            BuildDbTab();
            BuildExceptionsTab();

            RegisterEvents();

            LoadSettingsToUI();
            PopulateCabinets();
            PopulateInventoryTemplates();
            PopulateExportFilters();
            RefreshRoleUi();
            RefreshDraftList();
            RefreshNotifications();
            RefreshTaskQueue();
            RefreshAdaptersAssign();
            RefreshAdaptersInventory();
            RefreshGrid();
        }

        void BuildLayout()
        {
            tabs.Dock = DockStyle.Fill;
            tabs.TabPages.Clear();
            tabs.TabPages.AddRange(new[] { tabSettings, tabAssign, tabInventory, tabDb, tabExceptions });

            var header = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(12, 8, 12, 8)
            };
            header.Controls.Add(new Label { Text = "Оператор:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
            header.Controls.Add(tbUserName);
            header.Controls.Add(new Label { Text = "Роль:", AutoSize = true, Margin = new Padding(12, 6, 6, 0) });
            foreach (var role in Enum.GetValues(typeof(UserRole)))
                cbRoleSelector.Items.Add(role);
            cbRoleSelector.SelectedItem = currentUser.Role;
            header.Controls.Add(cbRoleSelector);
            header.Controls.Add(btnSwitchUser);
            header.Controls.Add(lblUserSummary);
            header.Controls.Add(lblAutosave);

            var rightPanel = BuildSidePanel();

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 720,
                Panel2MinSize = 240
            };
            split.Panel1.Controls.Add(tabs);
            split.Panel2.Controls.Add(rightPanel);

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            mainLayout.Controls.Add(header, 0, 0);
            mainLayout.Controls.Add(split, 0, 1);

            Controls.Add(mainLayout);
        }

        Control BuildSidePanel()
        {
            lvNotifications.Columns.Clear();
            lvNotifications.Columns.Add("Время", 90);
            lvNotifications.Columns.Add("Сообщение", 200);
            lvNotifications.Columns.Add("Статус", 90);

            lvTaskQueue.Columns.Clear();
            lvTaskQueue.Columns.Add("Задача", 140);
            lvTaskQueue.Columns.Add("Состояние", 120);
            lvTaskQueue.Columns.Add("Прогресс", 80);

            lvExceptions.Columns.Clear();
            lvExceptions.Columns.Add("Тип", 120);
            lvExceptions.Columns.Add("Описание", 220);

            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                AutoScroll = true
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 30f));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 20f));

            var grpNotifications = new GroupBox { Text = "Уведомления", Dock = DockStyle.Fill };
            grpNotifications.Controls.Add(lvNotifications);
            panel.Controls.Add(grpNotifications, 0, 0);

            var grpTasks = new GroupBox { Text = "Очередь задач", Dock = DockStyle.Fill };
            var taskLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            taskLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            taskLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            taskLayout.Controls.Add(lvTaskQueue, 0, 0);
            taskLayout.Controls.Add(btnCancelSelectedTask, 0, 1);
            grpTasks.Controls.Add(taskLayout);
            panel.Controls.Add(grpTasks, 0, 1);

            var grpExceptions = new GroupBox { Text = "Исключения", Dock = DockStyle.Fill };
            grpExceptions.Controls.Add(lvExceptions);
            panel.Controls.Add(grpExceptions, 0, 2);

            return panel;
        }

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
        }

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
            root.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            var draftsPanel = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 1,
                Dock = DockStyle.Top
            };
            draftsPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            draftsPanel.Controls.Add(new Label { Text = "Черновики", AutoSize = true, Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold) }, 0, 0);
            lbDrafts.DataSource = draftBinding;
            lbDrafts.DisplayMember = nameof(AssignmentDraft.Cabinet);
            lbDrafts.ValueMember = nameof(AssignmentDraft.Id);
            draftsPanel.RowCount++;
            draftsPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            draftsPanel.Controls.Add(lbDrafts, 0, 1);
            draftsPanel.RowCount++;
            draftsPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            draftsPanel.Controls.Add(btnNewDraft, 0, 2);
            draftsPanel.RowCount++;
            draftsPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            draftsPanel.Controls.Add(lblDraftMeta, 0, 3);

            root.Controls.Add(draftsPanel, 0, 0);

            lvRiskSummary.Columns.Clear();
            lvRiskSummary.Columns.Add("Риск", 140);
            lvRiskSummary.Columns.Add("Описание", 220);
            lvRiskSummary.Columns.Add("Соглас.", 80);

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
                form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var lbl = new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) };
                input.Margin = new Padding(0, 3, 0, 3);
                input.Anchor = AnchorStyles.Left | AnchorStyles.Right;
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
            form.Controls.Add(new Label { Text = "Адаптеры:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, rowA);
            form.Controls.Add(lbAdaptersAssign, 1, rowA);

            var rowManual = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var manualPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false, Dock = DockStyle.Fill };
            manualPanel.Controls.Add(cbManualIp);
            manualPanel.Controls.Add(tbManualIp);
            form.Controls.Add(new Label { Text = "IP:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, rowManual);
            form.Controls.Add(manualPanel, 1, rowManual);

            AddRow("Причина:", tbAssignReason);
            AddRow("Атрибуты:", tbAssignAttributes);

            var rowTP = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            form.Controls.Add(new Label { Text = "TCP-порты:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, rowTP);
            form.Controls.Add(tbProbePorts, 1, rowTP);

            var rowTCB = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbEnableTcpProbe.Anchor = AnchorStyles.Left;
            cbEnableTcpProbe.Margin = new Padding(0, 6, 0, 6);
            form.Controls.Add(cbEnableTcpProbe, 0, rowTCB);
            form.SetColumnSpan(cbEnableTcpProbe, 2);

            var rowPolicy = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lblPolicyHint.Margin = new Padding(0, 6, 0, 6);
            lblPolicyHint.MaximumSize = new Size(420, 0);
            form.Controls.Add(lblPolicyHint, 0, rowPolicy);
            form.SetColumnSpan(lblPolicyHint, 2);

            var rowProxy = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbProxyOn.Anchor = AnchorStyles.Left;
            cbProxyOn.Margin = new Padding(0, 6, 0, 6);
            form.Controls.Add(cbProxyOn, 0, rowProxy);
            form.SetColumnSpan(cbProxyOn, 2);

            var rowRisk = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            form.Controls.Add(new Label { Text = "Риски:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, rowRisk);
            form.Controls.Add(lvRiskSummary, 1, rowRisk);

            var rowAck = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbAcknowledgeRisks.Margin = new Padding(0, 6, 0, 6);
            form.Controls.Add(cbAcknowledgeRisks, 0, rowAck);
            form.SetColumnSpan(cbAcknowledgeRisks, 2);

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
            btnCol.Controls.Add(btnSubmitApproval);

            var actionsRow = form.RowCount; form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            form.Controls.Add(btnCol, 1, actionsRow);

            root.Controls.Add(form, 1, 0);

            tabAssign.Controls.Add(root);
        }

        void BuildInventoryTab()
        {
            tabInventory.Controls.Clear();

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                ColumnCount = 2,
                Padding = new Padding(12)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

            var workflow = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 2,
                Dock = DockStyle.Top
            };
            workflow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            workflow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            void AddRow(string label, Control control)
            {
                var row = workflow.RowCount; workflow.RowCount++;
                workflow.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                workflow.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, row);
                control.Margin = new Padding(0, 3, 0, 3);
                control.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                workflow.Controls.Add(control, 1, row);
            }

            AddRow("Шаблон кабинета:", cbInventoryTemplate);
            AddRow("Статус:", lblInventoryStatus);

            var btnFlow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false };
            btnFlow.Controls.Add(btnStartInventorySession);
            btnFlow.Controls.Add(btnInventoryPrev);
            btnFlow.Controls.Add(btnInventoryNext);
            workflow.RowCount++;
            workflow.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            workflow.Controls.Add(btnFlow, 1, workflow.RowCount - 1);

            workflow.RowCount++;
            workflow.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var markFlow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false };
            markFlow.Controls.Add(btnInventoryMarkOk);
            markFlow.Controls.Add(btnInventoryMarkMissing);
            markFlow.Controls.Add(btnInventoryMarkReview);
            workflow.Controls.Add(markFlow, 1, workflow.RowCount - 1);

            workflow.RowCount++;
            workflow.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            workflow.Controls.Add(lbInventoryChecklist, 0, workflow.RowCount - 1);
            workflow.SetColumnSpan(lbInventoryChecklist, 2);

            root.Controls.Add(workflow, 0, 0);

            var adaptersPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2
            };
            adaptersPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            adaptersPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            var rowV = adaptersPanel.RowCount; adaptersPanel.RowCount++;
            adaptersPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            cbShowVirtualInv.Anchor = AnchorStyles.Left;
            cbShowVirtualInv.Margin = new Padding(0, 6, 0, 6);
            adaptersPanel.Controls.Add(cbShowVirtualInv, 0, rowV);
            adaptersPanel.SetColumnSpan(cbShowVirtualInv, 2);

            adaptersPanel.RowCount++;
            adaptersPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            adaptersPanel.Controls.Add(new Label { Text = "Адаптеры:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, adaptersPanel.RowCount - 1);
            adaptersPanel.Controls.Add(lbAdaptersInventory, 1, adaptersPanel.RowCount - 1);

            adaptersPanel.RowCount++;
            adaptersPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            adaptersPanel.Controls.Add(new Label { Text = "Текущие:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, adaptersPanel.RowCount - 1);
            adaptersPanel.Controls.Add(lblCurrent, 1, adaptersPanel.RowCount - 1);

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
            adaptersPanel.RowCount++;
            adaptersPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            adaptersPanel.Controls.Add(btnCol, 1, adaptersPanel.RowCount - 1);

            root.Controls.Add(adaptersPanel, 1, 0);

            tabInventory.Controls.Add(root);
        }

        void BuildDbTab()
        {
            tabDb.Controls.Clear();

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                ColumnCount = 2,
                Padding = new Padding(12)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            var filters = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 2,
                Dock = DockStyle.Top
            };
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            void AddRow(string label, Control control)
            {
                var row = filters.RowCount; filters.RowCount++;
                filters.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                filters.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, row);
                control.Margin = new Padding(0, 3, 0, 3);
                control.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                filters.Controls.Add(control, 1, row);
            }

            AddRow("Локации:", clbExportLocations);
            AddRow("Статусы:", clbExportStatuses);
            AddRow("Дата с:", dtExportFrom);
            AddRow("Дата по:", dtExportTo);
            AddRow("Ответственный:", tbExportResponsible);
            cbExportFormat.Items.AddRange(new object[] { "xlsx", "csv" });
            cbExportFormat.SelectedIndex = 0;
            AddRow("Формат:", cbExportFormat);

            var exportBtns = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false };
            exportBtns.Controls.Add(btnExportPreview);
            exportBtns.Controls.Add(btnExport);
            filters.RowCount++;
            filters.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            filters.Controls.Add(exportBtns, 1, filters.RowCount - 1);

            root.Controls.Add(filters, 0, 0);

            var gridPanel = new Panel { Dock = DockStyle.Fill };
            gridPanel.Controls.Add(grid);
            root.Controls.Add(gridPanel, 1, 0);

            tabDb.Controls.Add(root);
        }

        void BuildExceptionsTab()
        {
            tabExceptions.Controls.Clear();
            var panel = new Panel { Dock = DockStyle.Fill };
            panel.Controls.Add(lvExceptions);
            tabExceptions.Controls.Add(panel);
        }

        void RegisterEvents()
        {
            btnSaveSettings.Click += (s, e) => SaveSettingsFromUI();
            cbShowVirtualAssign.CheckedChanged += (s, e) => RefreshAdaptersAssign();
            cbShowVirtualInv.CheckedChanged += (s, e) => RefreshAdaptersInventory();
            cbManualIp.CheckedChanged += (s, e) => tbManualIp.Enabled = cbManualIp.Checked;
            btnSuggestIp.Click += (s, e) => SuggestIpClick();
            btnApply.Click += (s, e) => ApplyClick();
            btnSubmitApproval.Click += (s, e) => SubmitDraft();
            lbDrafts.SelectedIndexChanged += (s, e) => LoadSelectedDraft();
            btnNewDraft.Click += (s, e) => StartNewDraft();
            cbCabinet.TextChanged += (s, e) => OnAssignmentInputChanged();
            tbAssignReason.TextChanged += (s, e) => OnAssignmentInputChanged();
            tbAssignAttributes.TextChanged += (s, e) => OnAssignmentInputChanged();
            tbManualIp.TextChanged += (s, e) => OnAssignmentInputChanged();
            cbAcknowledgeRisks.CheckedChanged += (s, e) => UpdateSubmissionState();
            btnSwitchUser.Click += (s, e) => UpdateCurrentUser();
            cbRoleSelector.SelectedIndexChanged += (s, e) => UpdateCurrentUser();
            notificationSvc.NotificationsChanged += (s, e) => Invoke((Action)RefreshNotifications);
            draftSvc.DraftsChanged += (s, e) => Invoke((Action)RefreshDraftList);
            taskQueue.TasksChanged += (s, e) => Invoke((Action)RefreshTaskQueue);
            lvNotifications.DoubleClick += (s, e) => AcknowledgeSelectedNotification();
            btnCancelSelectedTask.Click += (s, e) => CancelSelectedTask();
            btnReadCurrent.Click += (s, e) => ReadCurrentClick();
            btnSaveInventory.Click += (s, e) => SaveInventoryClick();
            btnStartInventorySession.Click += (s, e) => StartInventorySession();
            btnInventoryMarkOk.Click += (s, e) => UpdateInventoryEntry(InventoryEntryStatus.Checked, "Проверено");
            btnInventoryMarkMissing.Click += (s, e) => UpdateInventoryEntry(InventoryEntryStatus.Missing, "Не найдено");
            btnInventoryMarkReview.Click += (s, e) => UpdateInventoryEntry(InventoryEntryStatus.NeedsReview, "Требует проверки");
            btnInventoryNext.Click += (s, e) => AdvanceInventory(1);
            btnInventoryPrev.Click += (s, e) => AdvanceInventory(-1);
            btnExportPreview.Click += (s, e) => PreviewExport();
            btnExport.Click += (s, e) => ExportClick();
            this.FormClosing += MainForm_FormClosing;
            lbDrafts.Format += (s, e) =>
            {
                if (e.ListItem is AssignmentDraft draft)
                {
                    var cabinet = string.IsNullOrWhiteSpace(draft.Cabinet) ? "(кабинет не указан)" : draft.Cabinet;
                    e.Value = $"{cabinet} — {draft.Status}";
                }
            };
            lbInventoryChecklist.DataSource = inventoryBinding;
            lbInventoryChecklist.DisplayMember = nameof(InventoryEntry.DisplayName);
        }

        void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (hasUnsavedChanges && activeDraft != null)
            {
                var res = MessageBox.Show("Есть несохранённые изменения в черновике. Закрыть без сохранения?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (res == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
            }
        }

        void UpdateCurrentUser()
        {
            if (cbRoleSelector.SelectedItem is UserRole role)
                currentUser.UpdateRole(role);
            if (!string.IsNullOrWhiteSpace(tbUserName.Text))
                currentUser.UpdateDisplayName(tbUserName.Text);
            RefreshRoleUi();
            RefreshDraftList();
            RefreshNotifications();
            RefreshTaskQueue();
        }

        void RefreshRoleUi()
        {
            tbUserName.Text = currentUser.DisplayName;
            cbRoleSelector.SelectedItem = currentUser.Role;
            lblUserSummary.Text = $"Текущий пользователь: {currentUser}";
        }

        void PopulateCabinets()
        {
            cbCabinet.Items.Clear();
            foreach (var cabinet in db.GetCabinets())
                cbCabinet.Items.Add(cabinet.Name);
        }

        void PopulateInventoryTemplates()
        {
            cbInventoryTemplate.Items.Clear();
            foreach (var template in inventoryWorkflow.Templates)
                cbInventoryTemplate.Items.Add(template);
            if (cbInventoryTemplate.Items.Count > 0)
                cbInventoryTemplate.SelectedIndex = 0;
        }

        void PopulateExportFilters()
        {
            clbExportLocations.Items.Clear();
            foreach (var cabinet in db.GetCabinets())
                clbExportLocations.Items.Add(cabinet.Name, false);

            clbExportStatuses.Items.Clear();
            clbExportStatuses.Items.Add("auto", true);
            clbExportStatuses.Items.Add("inventory", true);
            clbExportStatuses.Items.Add("draft", false);
        }

        void RefreshDraftList()
        {
            draftBinding.Clear();
            foreach (var draft in draftSvc.GetDrafts(currentUser))
            {
                if (access.CanView(currentUser, draft))
                    draftBinding.Add(draft);
            }
            if (draftBinding.Count == 0)
            {
                StartNewDraft();
            }
            else if (activeDraft != null)
            {
                var current = draftBinding.FirstOrDefault(d => d.Id == activeDraft.Id);
                if (current != null)
                    lbDrafts.SelectedItem = current;
            }
        }

        void StartNewDraft()
        {
            activeDraft = new AssignmentDraft();
            activeDraft.SetOwner(currentUser);
            draftSvc.SaveDraft(currentUser, activeDraft);
            RefreshDraftList();
            lbDrafts.SelectedItem = draftBinding.FirstOrDefault(d => d.Id == activeDraft.Id);
            LoadDraftIntoUi(activeDraft);
        }

        void LoadSelectedDraft()
        {
            if (lbDrafts.SelectedItem is AssignmentDraft draft)
            {
                var stored = draftSvc.GetDraft(draft.Id);
                if (stored != null)
                {
                    activeDraft = stored;
                    LoadDraftIntoUi(activeDraft);
                }
            }
        }

        void LoadDraftIntoUi(AssignmentDraft draft)
        {
            suppressDraftEvents = true;
            cbCabinet.Text = draft.Cabinet ?? string.Empty;
            tbAssignReason.Text = draft.Reason ?? string.Empty;
            tbAssignAttributes.Text = draft.ToAttributesText();
            cbManualIp.Checked = !string.IsNullOrWhiteSpace(draft.RequestedIp);
            tbManualIp.Text = draft.RequestedIp ?? string.Empty;
            lblSuggested.Text = string.IsNullOrWhiteSpace(draft.SuggestedIp) ? "—" : draft.SuggestedIp;
            UpdateDraftMeta();
            suppressDraftEvents = false;
            hasUnsavedChanges = false;
            lblAutosave.Text = string.Empty;
            UpdatePolicyHint();
            UpdateSubmissionState();
            var restoredDecision = new AssignmentDecision { SuggestedIp = draft.SuggestedIp };
            foreach (var warning in draft.Warnings) restoredDecision.Warnings.Add(warning);
            UpdateRiskSummary(restoredDecision);
        }

        void OnAssignmentInputChanged()
        {
            if (suppressDraftEvents) return;
            hasUnsavedChanges = true;
            AutoSaveDraft();
            UpdatePolicyHint();
        }

        void AutoSaveDraft()
        {
            if (activeDraft == null || suppressDraftEvents)
                return;

            CollectDraftFromUi();
            draftSvc.SaveDraft(currentUser, activeDraft);
            hasUnsavedChanges = false;
            lblAutosave.Text = $"Черновик сохранён {DateTime.Now:T}";
            UpdateDraftMeta();
        }

        void CollectDraftFromUi()
        {
            if (activeDraft == null) return;
            activeDraft.Cabinet = cbCabinet.Text?.Trim();
            activeDraft.Reason = tbAssignReason.Text?.Trim();
            activeDraft.SetAttributesFromText(tbAssignAttributes.Text);
            activeDraft.RequestedIp = cbManualIp.Checked ? tbManualIp.Text?.Trim() : null;
            activeDraft.SetOwner(currentUser);
            activeDraft.UpdatedAt = DateTime.UtcNow;
        }

        void UpdateDraftMeta()
        {
            if (activeDraft == null)
            {
                lblDraftMeta.Text = string.Empty;
                return;
            }
            lblDraftMeta.Text = $"Создан: {activeDraft.CreatedAt:g}, статус: {activeDraft.Status}";
        }

        void UpdatePolicyHint()
        {
            if (string.IsNullOrWhiteSpace(cbCabinet.Text))
            {
                lblPolicyHint.Text = "Укажите кабинет для применения лимитов.";
                return;
            }
            var policy = ipPolicy.State.GetPolicyForCabinet(cbCabinet.Text);
            if (policy != null)
            {
                lblPolicyHint.Text = $"Лимит кабинета: {policy.Limit}. Исключений: {policy.Exceptions.Count}.";
            }
            else
            {
                lblPolicyHint.Text = $"Применяется общий лимит {ipPolicy.State.DefaultCabinetLimit}.";
            }
        }

        void UpdateRiskSummary(AssignmentDecision decision)
        {
            lvRiskSummary.Items.Clear();
            if (decision == null) return;
            foreach (var warning in decision.Warnings)
            {
                var item = new ListViewItem(warning.Title ?? warning.Code ?? "Риск");
                item.SubItems.Add(warning.Message);
                item.SubItems.Add(warning.RequiresSeniorApproval ? "Да" : "Нет");
                lvRiskSummary.Items.Add(item);
            }
            if (!string.IsNullOrWhiteSpace(decision.SuggestedIp))
                lblSuggested.Text = decision.SuggestedIp;
            UpdateSubmissionState();
        }

        void UpdateSubmissionState()
        {
            bool canSubmit = activeDraft != null && (!activeDraft.RequiresSeniorApproval || cbAcknowledgeRisks.Checked || currentUser.Role == UserRole.SeniorOperator || currentUser.Role == UserRole.Administrator);
            btnSubmitApproval.Enabled = canSubmit;
        }

        void SuggestIpClick()
        {
            if (activeDraft == null)
                StartNewDraft();
            CollectDraftFromUi();
            try
            {
                var decision = ipPolicy.Evaluate(db.LoadSettings(), activeDraft);
                UpdateRiskSummary(decision);
                AutoSaveDraft();
                if (decision.RequiresSeniorApproval)
                {
                    notificationSvc.Publish(new Notification
                    {
                        Title = "Назначение требует подтверждения",
                        Message = $"Кабинет {activeDraft.Cabinet}: {decision.SuggestedIp}",
                        Severity = NotificationSeverity.Warning,
                        RouteToRole = UserRole.SeniorOperator,
                        RequiresAttention = true,
                        LinkedEntityId = activeDraft.Id.ToString()
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        void SubmitDraft()
        {
            if (activeDraft == null)
                return;

            CollectDraftFromUi();
            var decision = ipPolicy.Evaluate(db.LoadSettings(), activeDraft);
            UpdateRiskSummary(decision);

            if (decision.RequiresSeniorApproval && currentUser.Role == UserRole.Operator && !cbAcknowledgeRisks.Checked)
            {
                MessageBox.Show("Отметьте подтверждение рисков перед отправкой.");
                return;
            }

            if (decision.RequiresSeniorApproval && currentUser.Role == UserRole.Operator)
            {
                activeDraft.Status = AssignmentStatus.PendingApproval;
                draftSvc.SaveDraft(currentUser, activeDraft);
                notificationSvc.Publish(new Notification
                {
                    Title = "Черновик отправлен",
                    Message = $"Кабинет {activeDraft.Cabinet}, IP {decision.SuggestedIp}",
                    Severity = NotificationSeverity.Info,
                    RouteToRole = UserRole.SeniorOperator,
                    RequiresAttention = true,
                    LinkedEntityId = activeDraft.Id.ToString()
                });
                MessageBox.Show("Черновик направлен старшему оператору.");
            }
            else
            {
                activeDraft.Status = AssignmentStatus.Approved;
                draftSvc.SaveDraft(currentUser, activeDraft);
                notificationSvc.Publish(new Notification
                {
                    Title = "Черновик утверждён",
                    Message = $"IP {decision.SuggestedIp} доступен к применению",
                    Severity = NotificationSeverity.Info,
                    LinkedEntityId = activeDraft.Id.ToString()
                });
                MessageBox.Show("Черновик утверждён и готов к применению.");
            }

            RefreshDraftList();
            UpdateDraftMeta();
        }

        void ApplyClick()
        {
            var s = db.LoadSettings();
            if (string.IsNullOrWhiteSpace(s.PoolStart) || string.IsNullOrWhiteSpace(s.PoolEnd))
            {
                MessageBox.Show("Настройте пул в Настройках.");
                return;
            }

            string cabinetName = string.IsNullOrWhiteSpace(cbCabinet.Text) ? "Не указан" : cbCabinet.Text.Trim();
            int cabId = db.EnsureCabinet(cabinetName);

            var decision = ipPolicy.Evaluate(s, activeDraft ?? new AssignmentDraft { Cabinet = cabinetName, RequestedIp = cbManualIp.Checked ? tbManualIp.Text : null });
            if (decision.RequiresSeniorApproval && currentUser.Role == UserRole.Operator)
            {
                MessageBox.Show("Требуется подтверждение старшего оператора.");
                return;
            }

            var (adapterId, adapterName, mac) = GetSelectedAdapterFrom(lbAdaptersAssign);
            if (string.IsNullOrEmpty(adapterId)) return;

            string ip = cbManualIp.Checked ? tbManualIp.Text.Trim() : lblSuggested.Text;
            if (!Validation.IsValidIPv4(ip))
            {
                MessageBox.Show("Некорректный IP/нет предложения.");
                return;
            }

            try
            {
                netSvc.ApplyIPv4(adapterId, ip, s.Netmask, s.Gateway,
                    string.IsNullOrWhiteSpace(s.Dns2) ? new[] { s.Dns1 } : new[] { s.Dns1, s.Dns2 });

                if (cbProxyOn.Checked)
                    proxySvc.SetGlobalProxy(true, s.ProxyHostPort, s.ProxyBypass);

                var m = new Machine
                {
                    CabinetId = cabId,
                    CabinetName = cabinetName,
                    Hostname = Environment.MachineName,
                    Mac = mac,
                    AdapterName = adapterName,
                    Ip = ip,
                    ProxyOn = cbProxyOn.Checked,
                    AssignedAt = DateTime.Now,
                    Source = "auto",
                };
                db.InsertMachine(m);
                audit.Log("ASSIGN_IP", new { cabinet = cabinetName, ip, adapterId, adapterName, proxy = cbProxyOn.Checked });

                lblSuggested.Text = "—";
                cbManualIp.Checked = false;
                tbManualIp.Enabled = false;

                MessageBox.Show("Настройки применены.");
                RefreshGrid();

                notificationSvc.Publish(new Notification
                {
                    Title = "IP назначен",
                    Message = $"{cabinetName}: {ip}",
                    Severity = NotificationSeverity.Info
                });
            }
            catch (Exception ex)
            {
                db.UpsertIpLease(ip, null, null, "free");
                MessageBox.Show("Ошибка применения: " + ex.Message);
            }
        }

        void RefreshAdaptersAssign()
        {
            lbAdaptersAssign.Items.Clear();
            foreach (var a in netSvc.ListAdapters(cbShowVirtualAssign.Checked))
                lbAdaptersAssign.Items.Add(a);
            if (lbAdaptersAssign.Items.Count > 0)
                lbAdaptersAssign.SelectedIndex = 0;
        }

        void RefreshAdaptersInventory()
        {
            lbAdaptersInventory.Items.Clear();
            foreach (var a in netSvc.ListAdapters(cbShowVirtualInv.Checked))
                lbAdaptersInventory.Items.Add(a);
            if (lbAdaptersInventory.Items.Count > 0)
                lbAdaptersInventory.SelectedIndex = 0;
        }

        (string adapterId, string adapterName, string mac) GetSelectedAdapterFrom(ListBox list)
        {
            if (list.SelectedItem is AdapterInfo adapter)
                return (adapter.Id, adapter.Name, adapter.Mac);
            MessageBox.Show("Выберите сетевой адаптер.");
            return (null, null, null);
        }

        void ReadCurrentClick()
        {
            var (adapterId, adapterName, mac) = GetSelectedAdapterFrom(lbAdaptersInventory);
            if (string.IsNullOrEmpty(adapterId)) return;
            var cur = invSvc.ReadCurrentIPv4(adapterName);
            lblCurrent.Text = $"IP={cur.ip} MASK={cur.mask} GW={cur.gateway} DNS=[{string.Join(",", cur.dns)}]";
        }

        void SaveInventoryClick()
        {
            if (!cbInternetOk.Checked)
            {
                MessageBox.Show("Отметьте, что интернет работает, чтобы записать без изменений.");
                return;
            }

            var (adapterId, adapterName, mac) = GetSelectedAdapterFrom(lbAdaptersInventory);
            if (string.IsNullOrEmpty(adapterId)) return;
            var cur = invSvc.ReadCurrentIPv4(adapterName);
            if (string.IsNullOrEmpty(cur.ip))
            {
                MessageBox.Show("Не удалось прочитать текущие настройки.");
                return;
            }

            string cabinetName = string.IsNullOrWhiteSpace(cbCabinet.Text) ? "Не указан" : cbCabinet.Text.Trim();
            int cabId = db.EnsureCabinet(cabinetName);

            var m = new Machine
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

        void StartInventorySession()
        {
            string template = cbInventoryTemplate.SelectedItem?.ToString();
            currentInventory = inventoryWorkflow.CreateSession(template, currentUser);
            inventoryBinding.Clear();
            foreach (var entry in currentInventory.Entries)
                inventoryBinding.Add(entry);
            if (inventoryBinding.Count > 0)
                lbInventoryChecklist.SelectedIndex = 0;
            UpdateInventoryStatus();
        }

        void UpdateInventoryEntry(InventoryEntryStatus status, string note)
        {
            if (currentInventory == null || currentInventory.CurrentEntry == null)
                return;
            currentInventory.CurrentEntry.UpdateStatus(status, note);
            inventoryBinding.ResetBindings();
            UpdateInventoryStatus();
        }

        void AdvanceInventory(int delta)
        {
            if (currentInventory == null) return;
            currentInventory.CurrentIndex = Math.Max(0, Math.Min(currentInventory.Entries.Count - 1, currentInventory.CurrentIndex + delta));
            lbInventoryChecklist.SelectedIndex = currentInventory.CurrentIndex;
            UpdateInventoryStatus();
        }

        void UpdateInventoryStatus()
        {
            if (currentInventory == null || currentInventory.CurrentEntry == null)
            {
                lblInventoryStatus.Text = "Сессия не запущена";
                return;
            }
            lblInventoryStatus.Text = $"Устройство {currentInventory.CurrentIndex + 1}/{currentInventory.Entries.Count}: {currentInventory.CurrentEntry.DisplayName} ({currentInventory.CurrentEntry.Status})";
        }

        void RefreshGrid()
        {
            var items = db.ListMachines();
            grid.DataSource = items;
        }

        void PreviewExport()
        {
            var data = ApplyExportFilters();
            grid.DataSource = data;
            MessageBox.Show($"Отобрано записей: {data.Count}");
        }

        List<Machine> ApplyExportFilters()
        {
            var data = db.ListMachines();
            var selectedLocations = clbExportLocations.CheckedItems.Cast<object>().Select(o => o.ToString()).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (selectedLocations.Any())
            {
                data = data.Where(m => selectedLocations.Contains(m.CabinetName)).ToList();
            }

            var statuses = clbExportStatuses.CheckedItems.Cast<object>().Select(o => o.ToString()).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (statuses.Any())
            {
                data = data.Where(m => statuses.Contains(m.Source ?? string.Empty)).ToList();
            }

            if (dtExportFrom.Value.Date <= dtExportTo.Value.Date)
            {
                var from = dtExportFrom.Value.Date;
                var to = dtExportTo.Value.Date.AddDays(1);
                data = data.Where(m => m.AssignedAt >= from && m.AssignedAt < to).ToList();
            }

            return data;
        }

        ExportRequest BuildExportRequest()
        {
            if (string.IsNullOrWhiteSpace(tbExportResponsible.Text))
            {
                MessageBox.Show("Укажите ответственное лицо за отчёт.");
                return null;
            }

            var request = new ExportRequest
            {
                Responsible = tbExportResponsible.Text.Trim(),
                RequestedBy = currentUser,
                Format = cbExportFormat.SelectedItem?.ToString() ?? "xlsx",
                From = dtExportFrom.Value.Date,
                To = dtExportTo.Value.Date
            };
            foreach (var loc in clbExportLocations.CheckedItems.Cast<object>())
                request.Locations.Add(loc.ToString());
            foreach (var status in clbExportStatuses.CheckedItems.Cast<object>())
                request.Statuses.Add(status.ToString());
            return request;
        }

        void ExportClick()
        {
            var request = BuildExportRequest();
            if (request == null) return;

            using (var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "export.xlsx" })
            {
                if (sfd.ShowDialog(this) != DialogResult.OK)
                    return;

                string path = sfd.FileName;
                taskQueue.Enqueue("Экспорт данных", async (token, progress) =>
                {
                    progress.Report(new TaskProgressReport { Percent = 10, Message = "Подготовка данных" });
                    var data = await Task.Run(() => ApplyExportFilters(), token).ConfigureAwait(false);
                    progress.Report(new TaskProgressReport { Percent = 60, Message = "Формирование файла" });
                    var result = await Task.Run(() => xlsx.ExportMachines(path, data), token).ConfigureAwait(false);
                    audit.Log("EXPORT_XLSX", new { file = result, count = data.Count, responsible = request.Responsible });
                    notificationSvc.Publish(new Notification
                    {
                        Title = "Экспорт завершён",
                        Message = $"Файл: {result}",
                        Severity = NotificationSeverity.Info
                    });
                    progress.Report(new TaskProgressReport { Percent = 100, Message = "Готово" });
                    return result;
                });
            }
        }

        void RefreshNotifications()
        {
            notificationBinding.Clear();
            foreach (var n in notificationSvc.GetForUser(currentUser))
            {
                if (access.CanSeeNotification(currentUser, n))
                    notificationBinding.Add(n);
            }
            lvNotifications.Items.Clear();
            foreach (var n in notificationBinding.OrderByDescending(n => n.CreatedAt))
            {
                var item = new ListViewItem(n.CreatedAt.ToLocalTime().ToString("g"));
                item.SubItems.Add(n.Title);
                item.SubItems.Add(n.Severity.ToString());
                item.Tag = n;
                lvNotifications.Items.Add(item);
            }
            RefreshExceptionsTab();
        }

        void RefreshExceptionsTab()
        {
            lvExceptions.Items.Clear();
            foreach (var n in notificationBinding.Where(n => n.RequiresAttention || n.Severity >= NotificationSeverity.Warning))
            {
                var item = new ListViewItem(n.Category ?? n.Severity.ToString());
                item.SubItems.Add(n.Message);
                item.Tag = n;
                lvExceptions.Items.Add(item);
            }
        }

        void RefreshTaskQueue()
        {
            taskBinding.Clear();
            foreach (var task in taskQueue.ListTasks())
                taskBinding.Add(task);
            lvTaskQueue.Items.Clear();
            foreach (var task in taskBinding.OrderByDescending(t => t.CreatedAt))
            {
                var item = new ListViewItem(task.Name) { Tag = task };
                item.SubItems.Add(task.Status.ToString());
                item.SubItems.Add(task.Progress + "%");
                lvTaskQueue.Items.Add(item);
            }
        }

        void AcknowledgeSelectedNotification()
        {
            if (lvNotifications.SelectedItems.Count == 0) return;
            if (lvNotifications.SelectedItems[0].Tag is Notification notification)
            {
                notificationSvc.Acknowledge(notification.Id);
            }
        }

        void CancelSelectedTask()
        {
            if (lvTaskQueue.SelectedItems.Count == 0) return;
            if (lvTaskQueue.SelectedItems[0].Tag is BackgroundTask task)
            {
                taskQueue.Cancel(task.Id);
            }
        }

        void SaveSettingsFromUI()
        {
            var settings = new AppSettings
            {
                PoolStart = tbPoolStart.Text,
                PoolEnd = tbPoolEnd.Text,
                Netmask = tbMask.Text,
                Gateway = tbGw.Text,
                Dns1 = tbDns1.Text,
                Dns2 = tbDns2.Text,
                ProxyHostPort = tbProxyHost.Text,
                ProxyBypass = tbProxyBypass.Text,
                ProxyGlobalOn = cbProxyGlobal.Checked
            };
            db.SaveSettings(settings);
            if (cbProxyGlobal.Checked)
                proxySvc.SetGlobalProxy(true, settings.ProxyHostPort, settings.ProxyBypass);
            else
                proxySvc.SetGlobalProxy(false, settings.ProxyHostPort, settings.ProxyBypass);
            MessageBox.Show("Настройки сохранены.");
        }

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

            using (var dlg = new MachineEditForm(db, fresh))
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
    }
}
