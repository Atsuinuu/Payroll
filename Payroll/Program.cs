#nullable disable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Linq;
using ClosedXML.Excel;

// ══════════════════════════════════════════════════════════════
//  ENTRY POINT
// ══════════════════════════════════════════════════════════════
static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new LoginForm());
    }
}

// ══════════════════════════════════════════════════════════════
//  API CLIENT
// ══════════════════════════════════════════════════════════════
static class ApiClient
{
    public static readonly HttpClient Http = new HttpClient();

    public static void Initialize(string baseUrl, string token)
    {
        Http.BaseAddress = new Uri(baseUrl.TrimEnd('/') + "/");
        Http.DefaultRequestHeaders.Clear();
        Http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Token", token);
        Http.DefaultRequestHeaders.Accept.Clear();
        Http.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json"));
        Http.Timeout = TimeSpan.FromSeconds(30);
    }

    public static async Task<string> GetAsync(string url)
    {
        var response = await Http.GetAsync(url);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"HTTP {(int)response.StatusCode}\n\n{body}");
        return body;
    }
}

// ══════════════════════════════════════════════════════════════
//  SHARED MODELS
// ══════════════════════════════════════════════════════════════
class PagedResponse<T>
{
    [JsonProperty("count")] public int Count { get; set; }
    [JsonProperty("next")] public string Next { get; set; }
    [JsonProperty("data")] public List<T> Data { get; set; }
}

class Employee
{
    [JsonProperty("emp_code")] public string EmpCode { get; set; }
    [JsonProperty("first_name")] public string FirstName { get; set; }
    [JsonProperty("last_name")] public string LastName { get; set; }
    [JsonProperty("department")] public object Department { get; set; }
    [JsonProperty("hire_date")] public string HireDate { get; set; }

    public string DepartmentName =>
        Department is Newtonsoft.Json.Linq.JObject obj
            ? obj["dept_name"]?.ToString() ?? ""
            : Department?.ToString() ?? "";
}

class PunchRecord
{
    [JsonProperty("emp_code")] public string EmpCode { get; set; }
    [JsonProperty("first_name")] public string FirstName { get; set; }
    [JsonProperty("last_name")] public string LastName { get; set; }
    [JsonProperty("department")] public string Department { get; set; }
    [JsonProperty("punch_time")] public string PunchTime { get; set; }
    [JsonProperty("punch_state")] public string PunchState { get; set; }
    [JsonProperty("punch_state_display")] public string PunchStateDisplay { get; set; }
}

class DepartmentItem
{
    [JsonProperty("id")] public int Id { get; set; }
    [JsonProperty("dept_code")] public string DeptCode { get; set; }
    [JsonProperty("dept_name")] public string DeptName { get; set; }
}

// ══════════════════════════════════════════════════════════════
//  LOGIN FORM
// ══════════════════════════════════════════════════════════════
class LoginForm : Form
{
    TextBox txtServer, txtUsername, txtPassword;
    Button btnConnect;
    Label lblStatus;

    public LoginForm()
    {
        Text = "ZKBio Time Login";
        Size = new Size(450, 340);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(18, 18, 18);
        ForeColor = Color.White;
        BuildUI();
    }

    void BuildUI()
    {
        Controls.Add(new Label
        {
            Text = "PAYROLL SYSTEM",
            Font = new Font("Segoe UI", 16, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 200, 150),
            AutoSize = true,
            Location = new Point(20, 20)
        });

        Controls.Add(MakeLabel("Server URL", 70));
        txtServer = MakeTextBox(95, "http://");
        Controls.Add(txtServer);

        Controls.Add(MakeLabel("Username", 135));
        txtUsername = MakeTextBox(160, "admin");
        Controls.Add(txtUsername);

        Controls.Add(MakeLabel("Password", 200));
        txtPassword = MakeTextBox(225, "");
        txtPassword.UseSystemPasswordChar = true;
        Controls.Add(txtPassword);

        lblStatus = new Label
        {
            ForeColor = Color.OrangeRed,
            Location = new Point(20, 270),
            Size = new Size(280, 40)
        };
        Controls.Add(lblStatus);

        btnConnect = new Button
        {
            Text = "Connect",
            Size = new Size(120, 36),
            Location = new Point(290, 265),
            BackColor = Color.FromArgb(0, 200, 150),
            ForeColor = Color.Black,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        btnConnect.FlatAppearance.BorderSize = 0;
        btnConnect.Click += BtnConnect_Click;
        Controls.Add(btnConnect);
    }

    Label MakeLabel(string text, int y) => new Label
    {
        Text = text,
        Location = new Point(20, y),
        ForeColor = Color.Silver,
        AutoSize = true
    };

    TextBox MakeTextBox(int y, string text) => new TextBox
    {
        Text = text,
        Location = new Point(20, y),
        Size = new Size(390, 28),
        BackColor = Color.FromArgb(35, 35, 35),
        ForeColor = Color.White,
        BorderStyle = BorderStyle.FixedSingle
    };

    async void BtnConnect_Click(object sender, EventArgs e)
    {
        btnConnect.Enabled = false;
        btnConnect.Text = "Connecting...";
        lblStatus.Text = "";

        try
        {
            using var client = new HttpClient();
            client.BaseAddress = new Uri(txtServer.Text.Trim().TrimEnd('/') + "/");

            var payload = JsonConvert.SerializeObject(new
            {
                username = txtUsername.Text.Trim(),
                password = txtPassword.Text
            });
            var content = new StringContent(payload, Encoding.UTF8, "application/json");
            var response = await client.PostAsync("api-token-auth/", content);
            var json = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                lblStatus.Text = $"LOGIN FAILED\nHTTP {(int)response.StatusCode}\n{json}";
                return;
            }

            dynamic result = JsonConvert.DeserializeObject(json);
            string token = result.token;

            if (string.IsNullOrWhiteSpace(token))
            {
                lblStatus.Text = "No token returned by server.";
                return;
            }

            ApiClient.Initialize(txtServer.Text.Trim(), token);
            Hide();
            var dashboard = new DashboardForm();
            dashboard.FormClosed += (s2, e2) => Close();
            dashboard.Show();
        }
        catch (Exception ex) { lblStatus.Text = ex.Message; }
        finally
        {
            btnConnect.Enabled = true;
            btnConnect.Text = "Connect";
        }
    }
}

// ══════════════════════════════════════════════════════════════
//  DASHBOARD FORM
// ══════════════════════════════════════════════════════════════
class DashboardForm : Form
{
    DataGridView dgv;
    Button btnLoad;
    Label lblStatus;

    public DashboardForm()
    {
        Text = "Payroll System";
        Size = new Size(1000, 620);
        BackColor = Color.FromArgb(18, 18, 18);
        BuildUI();
    }

    void BuildUI()
    {
        var toolbar = new Panel
        {
            Dock = DockStyle.Top,
            Height = 50,
            BackColor = Color.FromArgb(28, 28, 28)
        };

        btnLoad = MakeToolbarButton("Load Employees", 10, Color.FromArgb(0, 200, 150), Color.Black);
        btnLoad.Click += async (s, e) => await LoadEmployeesAsync();

        var btnExport = MakeToolbarButton("Export Timecard", 170, Color.FromArgb(0, 150, 255), Color.Black);
        btnExport.Click += (s, e) => new ExportDialog().ShowDialog(this);

        var btnTimetables = MakeToolbarButton("Timetables", 330, Color.FromArgb(100, 80, 200), Color.White);
        btnTimetables.Click += (s, e) => new TimetableManagerForm().ShowDialog(this);

        lblStatus = new Label
        {
            ForeColor = Color.Silver,
            AutoSize = true,
            Location = new Point(490, 15)
        };

        toolbar.Controls.Add(btnLoad);
        toolbar.Controls.Add(btnExport);
        toolbar.Controls.Add(btnTimetables);
        toolbar.Controls.Add(lblStatus);

        dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            BackgroundColor = Color.FromArgb(22, 22, 22),
            BorderStyle = BorderStyle.None,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            AllowUserToAddRows = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            EnableHeadersVisualStyles = false
        };
        dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.FromArgb(0, 200, 150),
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        dgv.DefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(28, 28, 28),
            ForeColor = Color.White,
            SelectionBackColor = Color.FromArgb(0, 160, 120),
            SelectionForeColor = Color.White
        };

        Controls.Add(dgv);
        Controls.Add(toolbar);
    }

    Button MakeToolbarButton(string text, int x, Color back, Color fore)
    {
        var btn = new Button
        {
            Text = text,
            Location = new Point(x, 10),
            Size = new Size(150, 28),
            BackColor = back,
            ForeColor = fore,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    async Task LoadEmployeesAsync()
    {
        btnLoad.Enabled = false;
        lblStatus.ForeColor = Color.Silver;
        lblStatus.Text = "Loading employees...";
        dgv.DataSource = null;

        try
        {
            var all = new List<Employee>();
            var url = "personnel/api/employees/?page=1&page_size=100";

            while (!string.IsNullOrWhiteSpace(url))
            {
                var json = await ApiClient.GetAsync(url);
                var paged = JsonConvert.DeserializeObject<PagedResponse<Employee>>(json);
                if (paged?.Data != null) all.AddRange(paged.Data);
                url = paged?.Next?.Replace(ApiClient.Http.BaseAddress.ToString(), "");
            }

            dgv.DataSource = all.ConvertAll(x => new
            {
                Employee_Code = x.EmpCode,
                First_Name = x.FirstName,
                Last_Name = x.LastName,
                Department = x.DepartmentName,
                Hire_Date = x.HireDate,
                Rate = ""
            });

            lblStatus.ForeColor = Color.FromArgb(0, 200, 150);
            lblStatus.Text = $"{all.Count} employees loaded.";
        }
        catch (Exception ex)
        {
            lblStatus.ForeColor = Color.OrangeRed;
            lblStatus.Text = ex.Message;
        }
        finally { btnLoad.Enabled = true; }
    }
}

// ══════════════════════════════════════════════════════════════
//  EXPORT DIALOG
// ══════════════════════════════════════════════════════════════
class ExportDialog : Form
{
    DateTimePicker dtpFrom, dtpTo;
    CheckedListBox clbDepts;
    Button btnLoadDepts, btnExport, btnSelectAll, btnClearAll;
    Label lblStatus;
    List<DepartmentItem> _departments = new List<DepartmentItem>();
    TimetableStore _store;

    // ── Column definitions ────────────────────────────────────
    // Daily/Summary sheet: 20 cols (A–T)
    static readonly string[] TimeCardHeaders =
    {
        "Employee ID",          // A  1
        "First Name",           // B  2
        "Department",           // C  3
        "Date",                 // D  4
        "Clock In",             // E  5
        "Clock Out",            // F  6
        "Worked Hours",         // G  7
        "Break Out",            // H  8
        "Break In",             // I  9
        "Total OT",             // J  10
        "NORMAL OT",            // K  11
        "Regular(H)",           // L  12
        "OT1(H)",               // M  13
        "DCTC REG HRS(H)",      // N  14
        "DCTC OT SYS ADJ(H) (EDITABLE)",      // O  15
        "DCTC TOTAL OT(H)",     // P  16
        "DCTC MANUAL ADJ REG(H) (EDITABLE)",   // Q  17
        "DCTC MANUAL ADJ OT(H) (EDITABLE)",    // R  18
        "TOTAL REG HRS(H)",     // S  19
        "TOTAL OT HRS(H)"       // T  20
    };

    // Summary sheet: same + OVERALL TOTAL (col U = 21)
    static readonly string[] SummaryHeaders =
    {
        "Employee ID", "First Name", "Department", "Date",
        "Clock In", "Clock Out", "Worked Hours", "Break Out", "Break In",
        "Total OT", "NORMAL OT", "Regular(H)", "OT1(H)",
        "DCTC REG HRS(H)", "DCTC OT SYS ADJ(H) (EDITABLE)", "DCTC TOTAL OT(H)",
        "DCTC MANUAL ADJ REG(H) (EDITABLE)", "DCTC MANUAL ADJ OT(H) (EDITABLE)",
        "TOTAL REG HRS(H)", "TOTAL OT HRS(H)", "OVERALL TOTAL"
    };

    // Payroll sheet: 34 cols (A–AH)
    static readonly string[] PayrollHeaders =
    {
        "NO.", "NAME", "DEPARTMENT", "RATE", "DAYS", "REG", "OT", "TOTAL",
        "GROSS", "SAT VALE", "WED VALE", "PV", "HELMET", "GLOVES", "SHOES",
        "CHINSTRAP", "SAFETY BOOTS", "RAINCOAT", "SAFETY VEST", "GOGGLES",
        "METRO", "HEADGUARD", "UNIFORM", "EXCESS", "MEDICAL", "PV",
        "TRANSPO", "H2O", "101 BARRACKS", "CANTEEN1", "CANTEEN2",
        "NETPAY", "SIGNATURE", "NO."
    };

    public ExportDialog()
    {
        Text = "Export Timecard to Excel";
        Size = new Size(480, 580);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        BackColor = Color.FromArgb(18, 18, 18);
        ForeColor = Color.White;
        _store = TimetableStore.Load();
        BuildUI();
    }

    void BuildUI()
    {
        Controls.Add(MakeLabel("Date Range", 16));
        Controls.Add(MakeLabel("From:", 38));
        dtpFrom = new DateTimePicker
        {
            Location = new Point(55, 35),
            Size = new Size(150, 26),
            Format = DateTimePickerFormat.Short,
            Value = DateTime.Today.AddDays(-6)
        };
        Controls.Add(dtpFrom);

        Controls.Add(MakeLabel("To:", 38, 225));
        dtpTo = new DateTimePicker
        {
            Location = new Point(244, 35),
            Size = new Size(150, 26),
            Format = DateTimePickerFormat.Short,
            Value = DateTime.Today
        };
        Controls.Add(dtpTo);

        Controls.Add(MakeLabel("Departments", 76));

        btnLoadDepts = MakeButton("Load Departments", 16, 94, 160);
        btnLoadDepts.Click += async (s, e) => await LoadDepartmentsAsync();
        Controls.Add(btnLoadDepts);

        btnSelectAll = MakeButton("Select All", 190, 94, 100);
        btnSelectAll.Click += (s, e) =>
        {
            for (int i = 0; i < clbDepts.Items.Count; i++)
                clbDepts.SetItemChecked(i, true);
        };
        Controls.Add(btnSelectAll);

        btnClearAll = MakeButton("Clear All", 304, 94, 100);
        btnClearAll.BackColor = Color.FromArgb(60, 60, 60);
        btnClearAll.ForeColor = Color.White;
        btnClearAll.Click += (s, e) =>
        {
            for (int i = 0; i < clbDepts.Items.Count; i++)
                clbDepts.SetItemChecked(i, false);
        };
        Controls.Add(btnClearAll);

        clbDepts = new CheckedListBox
        {
            Location = new Point(16, 128),
            Size = new Size(432, 290),
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            CheckOnClick = true,
            Font = new Font("Segoe UI", 9)
        };
        Controls.Add(clbDepts);

        Controls.Add(new Label
        {
            Text = "ℹ Departments without a timetable will show raw punch times only.",
            ForeColor = Color.FromArgb(255, 200, 0),
            Font = new Font("Segoe UI", 8, FontStyle.Italic),
            Location = new Point(16, 424),
            Size = new Size(432, 32)
        });

        lblStatus = new Label
        {
            Location = new Point(16, 460),
            Size = new Size(300, 50),
            ForeColor = Color.Silver,
            Font = new Font("Segoe UI", 8)
        };
        Controls.Add(lblStatus);

        btnExport = MakeButton("Export to Excel", 322, 468, 130);
        btnExport.Click += async (s, e) => await ExportAsync();
        Controls.Add(btnExport);
    }

    async Task LoadDepartmentsAsync()
    {
        btnLoadDepts.Enabled = false;
        lblStatus.ForeColor = Color.Silver;
        lblStatus.Text = "Loading departments...";
        clbDepts.Items.Clear();
        _departments.Clear();

        try
        {
            var url = "personnel/api/departments/?page=1&page_size=200";
            while (url != null)
            {
                var json = await ApiClient.GetAsync(url);
                var paged = JsonConvert.DeserializeObject<PagedResponse<DepartmentItem>>(json);
                if (paged?.Data != null) _departments.AddRange(paged.Data);
                url = paged?.Next?.Replace(ApiClient.Http.BaseAddress.ToString(), "");
            }

            foreach (var d in _departments.OrderBy(d => d.DeptName))
            {
                var hasTimetable = _store.GetForDepartment(d.DeptName) != null;
                clbDepts.Items.Add(
                    hasTimetable ? d.DeptName : $"{d.DeptName}  ⚠ no timetable",
                    true);
            }

            lblStatus.ForeColor = Color.FromArgb(0, 200, 150);
            lblStatus.Text = $"{_departments.Count} departments loaded.";
        }
        catch (Exception ex)
        {
            lblStatus.ForeColor = Color.OrangeRed;
            lblStatus.Text = $"Error: {ex.Message}";
        }
        finally { btnLoadDepts.Enabled = true; }
    }

    async Task ExportAsync()
    {
        var selectedRaw = clbDepts.CheckedItems.Cast<string>().ToList();
        if (selectedRaw.Count == 0)
        {
            lblStatus.ForeColor = Color.OrangeRed;
            lblStatus.Text = "Please select at least one department.";
            return;
        }

        var selectedDeptNames = selectedRaw
            .Select(s => s.Contains("  ⚠") ? s.Substring(0, s.IndexOf("  ⚠")) : s)
            .ToList();

        var sfd = new SaveFileDialog
        {
            Title = "Save Timecard Export",
            Filter = "Excel Files (*.xlsx)|*.xlsx",
            FileName = $"TimeCard_{dtpFrom.Value:yyyyMMdd}_{dtpTo.Value:yyyyMMdd}.xlsx"
        };
        if (sfd.ShowDialog() != DialogResult.OK) return;

        btnExport.Enabled = false;
        lblStatus.ForeColor = Color.Silver;
        lblStatus.Text = "Fetching punch records...";

        try
        {
            var from = dtpFrom.Value.ToString("yyyy-MM-dd") + " 00:00:00";
            var to = dtpTo.Value.ToString("yyyy-MM-dd") + " 23:59:59";
            var punches = new List<PunchRecord>();
            var url = $"iclock/api/transactions/?start_time={Uri.EscapeDataString(from)}&end_time={Uri.EscapeDataString(to)}&page=1&page_size=500";

            int page = 1;
            while (url != null)
            {
                lblStatus.Text = $"Fetching page {page++}...";
                Application.DoEvents();
                var json = await ApiClient.GetAsync(url);
                var paged = JsonConvert.DeserializeObject<PagedResponse<PunchRecord>>(json);
                if (paged?.Data != null) punches.AddRange(paged.Data);
                url = paged?.Next?.Replace(ApiClient.Http.BaseAddress.ToString(), "");
            }

            lblStatus.Text = $"Calculating attendance for {punches.Count} records...";
            Application.DoEvents();

            var filtered = punches.Where(p => selectedDeptNames.Contains(p.Department ?? "")).ToList();
            var calculated = AttendanceCalculator.Calculate(filtered, _store);

            lblStatus.Text = "Building Excel file...";
            Application.DoEvents();

            BuildExcel(calculated, selectedDeptNames, sfd.FileName);

            lblStatus.ForeColor = Color.FromArgb(0, 200, 150);
            lblStatus.Text = $"Done! {calculated.Count} rows across {selectedDeptNames.Count} dept(s).";

            if (MessageBox.Show("Export complete! Open the file now?", "Done",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = sfd.FileName,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            lblStatus.ForeColor = Color.OrangeRed;
            lblStatus.Text = $"Error: {ex.Message}";
        }
        finally { btnExport.Enabled = true; }
    }

    // ══════════════════════════════════════════════════════════
    //  BUILD EXCEL — 3 sheets per department
    //  1. SUMMARY sheet   — one row per employee, formulas ref daily sheet
    //  2. Daily sheet     — every date in range, every employee, total rows
    //  3. PAYROLL sheet   — payroll form with deduction columns
    // ══════════════════════════════════════════════════════════
    void BuildExcel(List<CalculatedRecord> records, List<string> deptNames, string filePath)
    {
        using var wb = new XLWorkbook();

        // Build date list for the full range (every day from→to)
        var allDates = new List<DateTime>();
        for (var d = dtpFrom.Value.Date; d <= dtpTo.Value.Date; d = d.AddDays(1))
            allDates.Add(d);

        string dateLabel = $"{dtpFrom.Value:MMM d} - {dtpTo.Value:MMM d, yyyy}".ToUpper();
        string dailySheetName = dtpTo.Value.ToString("yyyyMMdd");

        var byDept = records
            .GroupBy(r => r.Department)
            .OrderBy(g => g.Key);

        foreach (var deptGroup in byDept)
        {
            string dept = deptGroup.Key;

            // Group by employee, sorted by emp code
            var byEmp = deptGroup
                .GroupBy(r => r.EmpCode)
                .OrderBy(g => g.Key)
                .ToList();

            // ── Sheet 2: Daily detail (named by end date) ────
            string safeDaily = SanitizeName(dailySheetName);
            if (wb.Worksheets.Any(s => s.Name == safeDaily))
                safeDaily = SanitizeName($"{dept}_{dailySheetName}");
            var wsDaily = wb.Worksheets.Add(safeDaily);

            WriteTimeCardHeader(wsDaily, TimeCardHeaders.Length);
            WriteTimeCardColumnHeaders(wsDaily, TimeCardHeaders, 3);

            // Track total row positions per employee for summary formulas
            var empTotalRows = new Dictionary<string, int>(); // empCode → total row number

            int row = 4;
            foreach (var empGroup in byEmp)
            {
                int empStartRow = row;
                var empRecs = empGroup.ToDictionary(r => r.Date);
                var firstRec = empGroup.First();

                // Write one row per date in range
                foreach (var date in allDates)
                {
                    string dateStr = date.ToString("yyyy-MM-dd");
                    empRecs.TryGetValue(dateStr, out var rec);
                    WriteDailyRow(wsDaily, row, dateStr, firstRec, rec, safeDaily);
                    row++;
                }

                // Total row — uses SUM formulas
                int empEndRow = row - 1;
                WriteDailyTotalRow(wsDaily, row, firstRec, empStartRow, empEndRow, TimeCardHeaders.Length);
                empTotalRows[empGroup.Key] = row;
                row++;
            }

            FinalizeSheet(wsDaily, TimeCardHeaders.Length);

            // ── Sheet 1: Summary (references daily sheet) ────
            string safeSummary = SanitizeName($"SUMMARY {dateLabel}");
            if (wb.Worksheets.Any(s => s.Name == safeSummary))
                safeSummary = SanitizeName($"SUMMARY {dept} {dateLabel}");
            var wsSummary = wb.Worksheets.Add(safeSummary);

            WriteTimeCardHeader(wsSummary, SummaryHeaders.Length);
            WriteTimeCardColumnHeaders(wsSummary, SummaryHeaders, 3);

            int sumRow = 4;
            foreach (var empGroup in byEmp)
            {
                int totalRowInDaily = empTotalRows[empGroup.Key];
                WriteSummaryRow(wsSummary, sumRow, empGroup.First(),
                                dept, safeDaily, totalRowInDaily,
                                SummaryHeaders.Length);
                sumRow++;
            }

            FinalizeSheet(wsSummary, SummaryHeaders.Length);
            wsSummary.Position = 1; // Summary first

            // ── Sheet 3: Payroll form ─────────────────────────
            string safePayroll = SanitizeName($"PAYROLL {dateLabel}");
            if (wb.Worksheets.Any(s => s.Name == safePayroll))
                safePayroll = SanitizeName($"PAYROLL {dept} {dateLabel}");
            var wsPayroll = wb.Worksheets.Add(safePayroll);

            WritePayrollSheet(wsPayroll, dept, dateLabel, byEmp,
                              wsSummary.Name, PayrollHeaders);
            wsPayroll.Position = wsSummary.Position + 1;
        }

        wb.SaveAs(filePath);
    }

    // ── TimeCard header (rows 1-2) ────────────────────────────
    void WriteTimeCardHeader(IXLWorksheet ws, int cols)
    {
        ws.Range(1, 1, 1, cols).Merge();
        ws.Cell(1, 1).Value = "DCTC";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 14;

        ws.Range(2, 1, 2, cols).Merge();
        ws.Cell(2, 1).Value = "Total Time Card";
        ws.Cell(2, 1).Style.Font.Bold = true;
        ws.Cell(2, 1).Style.Font.FontSize = 11;
    }

    // ── TimeCard column headers (row 3) ──────────────────────
    void WriteTimeCardColumnHeaders(IXLWorksheet ws, string[] headers, int headerRow)
    {
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(headerRow, c + 1);
            cell.Value = headers[c];
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;

            // A-I blue | J-M teal | N-T orange | U green (overall total)
            if (c < 9) cell.Style.Fill.BackgroundColor = XLColor.FromArgb(47, 117, 181);
            else if (c < 13) cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 128);
            else if (c < 20) cell.Style.Fill.BackgroundColor = XLColor.FromArgb(180, 95, 6);
            else cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 100, 0);
        }
        ws.Row(headerRow).Height = 42;
    }

    // ── Daily row (one per date per employee) ─────────────────
    // rec may be null if employee had no punch that day
    void WriteDailyRow(IXLWorksheet ws, int row, string dateStr,
                       CalculatedRecord identity, CalculatedRecord rec,
                       string dailySheetName)
    {
        bool hasPunch = rec != null && !rec.IsAbsent;
        bool hasClockOut = hasPunch && !string.IsNullOrEmpty(rec?.ClockOut);
        bool hasBreak = hasPunch &&
                           !string.IsNullOrEmpty(rec?.BreakOut) &&
                           !string.IsNullOrEmpty(rec?.BreakIn);

        // Identity columns
        ws.Cell(row, 1).Value = identity.EmpCode ?? "";
        ws.Cell(row, 2).Value = identity.FullName ?? "";
        ws.Cell(row, 3).Value = identity.Department ?? "";
        ws.Cell(row, 4).Value = dateStr;

        // Punch times
        ws.Cell(row, 5).Value = hasPunch ? rec.ClockIn ?? "" : "";
        ws.Cell(row, 6).Value = hasPunch ? rec.ClockOut ?? "" : "";

        // Worked Hours — rounded to 1 decimal (e.g. "8.0", "11.0")
        if (hasPunch && hasClockOut)
            ws.Cell(row, 7).Value = Math.Round(rec.NetWorkedH, 1).ToString("F1");
        else
            ws.Cell(row, 7).Value = "";

        ws.Cell(row, 8).Value = hasBreak ? rec.BreakOut ?? "" : "";
        ws.Cell(row, 9).Value = hasBreak ? rec.BreakIn ?? "" : "";

        // Total OT (J), NORMAL OT (K), Regular (L), OT1 (M)
        // Show 0 for no-punch days (matches sample)
        ws.Cell(row, 10).Value = (XLCellValue)(hasPunch && rec.TotalOTH > 0 ? Math.Round(rec.TotalOTH, 1) : 0);
        ws.Cell(row, 11).Value = (XLCellValue)(hasPunch && rec.LateOutH > 0 ? Math.Round(rec.LateOutH, 1) : 0);
        ws.Cell(row, 12).Value = (XLCellValue)(hasPunch && rec.RegularH > 0 ? Math.Round(rec.RegularH, 1) : 0);
        ws.Cell(row, 13).Value = (XLCellValue)(hasPunch && rec.OT1H > 0 ? Math.Round(rec.OT1H, 1) : 0);

        // DCTC REG HRS (N) — formula: only count if all 4 punch fields present
        ws.Cell(row, 14).FormulaA1 =
            $"IF(OR(E{row}=\"\",F{row}=\"\",H{row}=\"\",I{row}=\"\"),0," +
            $"IFERROR(IF(L{row}=\"\",0,VALUE(L{row})),0))";

        // DCTC OT SYS ADJ (O) — blank/yellow, manual
        ws.Cell(row, 15).Value = "";

        // DCTC TOTAL OT (P) — formula
        ws.Cell(row, 16).FormulaA1 =
            $"IFERROR(IF(K{row}=\"\",0,VALUE(K{row})),0)" +
            $"+IFERROR(IF(M{row}=\"\",0,VALUE(M{row})),0)" +
            $"+IF(O{row}=\"\",0,O{row})";

        // DCTC MANUAL ADJ REG (Q) / OT (R) — blank/yellow, manual
        ws.Cell(row, 17).Value = "";
        ws.Cell(row, 18).Value = "";

        // TOTAL REG HRS (S) — formula
        ws.Cell(row, 19).FormulaA1 = $"N{row}+IF(Q{row}=\"\",0,Q{row})";

        // TOTAL OT HRS (T) — formula
        ws.Cell(row, 20).FormulaA1 = $"P{row}+IF(R{row}=\"\",0,R{row})";

        // Styling
        StyleDataRow(ws, row, TimeCardHeaders.Length);
    }

    // ── Employee total row on daily sheet ─────────────────────
    void WriteDailyTotalRow(IXLWorksheet ws, int row,
                             CalculatedRecord first,
                             int startRow, int endRow, int cols)
    {
        ws.Cell(row, 1).Value = first.EmpCode ?? "";
        ws.Cell(row, 2).Value = $"{first.FullName} Total";
        // Cols 3-13 blank on total row
        ws.Cell(row, 14).FormulaA1 = $"SUM(N{startRow}:N{endRow})";
        ws.Cell(row, 15).Value = "";
        ws.Cell(row, 16).FormulaA1 = $"SUM(P{startRow}:P{endRow})";
        ws.Cell(row, 17).Value = "";
        ws.Cell(row, 18).Value = "";
        ws.Cell(row, 19).FormulaA1 = $"SUM(S{startRow}:S{endRow})";
        ws.Cell(row, 20).FormulaA1 = $"SUM(T{startRow}:T{endRow})";

        // Total row style — bold, light blue
        var range = ws.Range(row, 1, row, cols);
        range.Style.Fill.BackgroundColor = XLColor.FromArgb(221, 235, 247);
        range.Style.Font.Bold = true;
        range.Style.Font.FontColor = XLColor.Black;
        range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        range.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
        for (int c = 4; c <= cols; c++)
            ws.Cell(row, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }

    // ── Summary row (references daily sheet total row) ────────
    void WriteSummaryRow(IXLWorksheet ws, int row,
                         CalculatedRecord first,
                         string dept, string dailySheet,
                         int dailyTotalRow, int cols)
    {
        string ds = dailySheet.Contains("'") ? dailySheet : dailySheet;
        // ClosedXML sheet reference — wrap in single quotes if needed
        string sheetRef = $"'{dailySheet}'";

        ws.Cell(row, 1).Value = first.EmpCode ?? "";
        ws.Cell(row, 2).Value = $"{first.FullName} Total";
        ws.Cell(row, 3).Value = dept;
        // Cols 4-13 blank
        ws.Cell(row, 14).Value = "";  // DCTC REG HRS — filled via formula below
        ws.Cell(row, 15).Value = "";  // DCTC OT SYS ADJ — manual
        // DCTC TOTAL OT references daily sheet total row col P
        ws.Cell(row, 16).FormulaA1 =
            $"{sheetRef}!P{dailyTotalRow}+IF(O{row}=\"\",0,O{row})";
        ws.Cell(row, 17).Value = "";
        ws.Cell(row, 18).Value = "";
        // TOTAL REG HRS and TOTAL OT HRS
        ws.Cell(row, 19).FormulaA1 = $"N{row}+IF(Q{row}=\"\",0,Q{row})";
        ws.Cell(row, 20).FormulaA1 = $"P{row}+IF(R{row}=\"\",0,R{row})";
        // OVERALL TOTAL
        ws.Cell(row, 21).FormulaA1 = $"S{row}+T{row}";

        // DCTC REG HRS — reference daily total row col N
        ws.Cell(row, 14).FormulaA1 = $"{sheetRef}!N{dailyTotalRow}";

        StyleDataRow(ws, row, cols);
        // Yellow on manual columns
        ws.Cell(row, 15).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
        ws.Cell(row, 17).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
        ws.Cell(row, 18).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
    }

    // ── Payroll sheet ─────────────────────────────────────────
    void WritePayrollSheet(IXLWorksheet ws, string dept, string dateLabel,
                           List<IGrouping<string, CalculatedRecord>> byEmp,
                           string summarySheetName, string[] headers)
    {
        int cols = headers.Length; // 34

        // Row 1: PAYROLL FORM + covered date
        ws.Range(1, 1, 1, 8).Merge();
        ws.Cell(1, 1).Value = "PAYROLL FORM";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 14;

        ws.Range(1, 9, 1, cols).Merge();
        ws.Cell(1, 9).Value = $"COVERED DATE : {dateLabel}";
        ws.Cell(1, 9).Style.Font.Bold = true;

        // Row 3: project name
        ws.Range(3, 1, 3, cols).Merge();
        ws.Cell(3, 1).Value = $"PROJECT : {dept}";
        ws.Cell(3, 1).Style.Font.Bold = true;

        // Row 4: column headers
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(4, c + 1);
            cell.Value = headers[c];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(47, 117, 181);
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        }
        ws.Row(4).Height = 30;

        // Data rows — one per employee
        // REG and OT values come from summary sheet
        string sumRef = $"'{summarySheetName}'";
        int dataRow = 5;
        int empNum = 1;

        // Build a lookup: summary sheet row per emp
        // Summary data starts at row 4, one row per emp in same order
        int summaryDataStartRow = 4;

        foreach (var empGroup in byEmp)
        {
            var first = empGroup.First();
            int sRow = summaryDataStartRow + (empNum - 1); // summary row for this emp

            ws.Cell(dataRow, 1).Value = empNum;             // NO.
            ws.Cell(dataRow, 2).Value = first.FullName ?? ""; // NAME
            ws.Cell(dataRow, 3).Value = dept;               // DEPARTMENT
            ws.Cell(dataRow, 4).Value = "";                 // RATE — manual
            // DAYS = REG / 8
            ws.Cell(dataRow, 5).FormulaA1 = $"F{dataRow}/8";
            // REG — from summary TOTAL REG HRS (col S = 19)
            ws.Cell(dataRow, 6).FormulaA1 = $"{sumRef}!S{sRow}";
            // OT — from summary TOTAL OT HRS (col T = 20)
            ws.Cell(dataRow, 7).FormulaA1 = $"{sumRef}!T{sRow}";
            // TOTAL = REG + OT
            ws.Cell(dataRow, 8).FormulaA1 = $"F{dataRow}+G{dataRow}";
            // GROSS = RATE * TOTAL (blank if no rate)
            ws.Cell(dataRow, 9).FormulaA1 = $"IF(D{dataRow}=\"\",0,D{dataRow}*H{dataRow})";
            // Cols 10-31: deductions — blank, fill manually
            // NETPAY = GROSS - SUM(deductions)
            ws.Cell(dataRow, 32).FormulaA1 = $"I{dataRow}-SUM(J{dataRow}:AE{dataRow})";
            // SIGNATURE — blank
            ws.Cell(dataRow, 33).Value = "";
            // NO. repeated
            ws.Cell(dataRow, 34).Value = empNum;

            // Row border
            var range = ws.Range(dataRow, 1, dataRow, cols);
            range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            range.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            range.Style.Fill.BackgroundColor = dataRow % 2 == 0
                ? XLColor.FromArgb(242, 242, 242)
                : XLColor.White;
            range.Style.Font.FontColor = XLColor.Black;

            // Center numeric cols
            for (int c = 1; c <= cols; c++)
                ws.Cell(dataRow, c).Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;
            // Left-align name and dept
            ws.Cell(dataRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            ws.Cell(dataRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

            dataRow++;
            empNum++;
        }

        // Finalize payroll sheet
        ws.Columns().AdjustToContents();
        ws.Column(2).Width = Math.Max(ws.Column(2).Width, 24); // NAME
        ws.Column(3).Width = Math.Max(ws.Column(3).Width, 22); // DEPT
        ws.SheetView.FreezeRows(4);
        ws.SheetView.FreezeColumns(2);
    }

    // ── Row styling helper ────────────────────────────────────
    void StyleDataRow(IXLWorksheet ws, int row, int cols)
    {
        var range = ws.Range(row, 1, row, cols);
        range.Style.Fill.BackgroundColor = XLColor.White;
        range.Style.Font.FontColor = XLColor.Black;
        range.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
        range.Style.Border.BottomBorderColor = XLColor.LightGray;
        for (int c = 4; c <= cols; c++)
            ws.Cell(row, c).Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;
        // Yellow on manual adj columns
        ws.Cell(row, 15).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
        ws.Cell(row, 17).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
        ws.Cell(row, 18).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
    }

    void FinalizeSheet(IXLWorksheet ws, int cols)
    {
        ws.Columns().AdjustToContents();
        ws.Column(1).Width = Math.Max(ws.Column(1).Width, 13);
        ws.Column(2).Width = Math.Max(ws.Column(2).Width, 28);
        ws.Column(3).Width = Math.Max(ws.Column(3).Width, 22);
        for (int c = 9; c <= cols; c++)
            ws.Column(c).Width = Math.Max(ws.Column(c).Width, 14);
        ws.SheetView.FreezeRows(3);
        ws.SheetView.FreezeColumns(2);
    }

    string SanitizeName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "Sheet";
        name = name.Length > 31 ? name.Substring(0, 31) : name;
        foreach (var c in new[] { ':', '\\', '/', '?', '*', '[', ']' })
            name = name.Replace(c.ToString(), "");
        return string.IsNullOrWhiteSpace(name) ? "Sheet" : name.Trim();
    }

    Label MakeLabel(string text, int y, int x = 16) => new Label
    {
        Text = text,
        Location = new Point(x, y),
        ForeColor = Color.Silver,
        AutoSize = true,
        Font = new Font("Segoe UI", 9, FontStyle.Bold)
    };

    Button MakeButton(string text, int x, int y, int width = 130)
    {
        var btn = new Button
        {
            Text = text,
            Location = new Point(x, y),
            Size = new Size(width, 28),
            BackColor = Color.FromArgb(0, 200, 150),
            ForeColor = Color.Black,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }
}