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
        txtServer = MakeTextBox(95, "http://122.53.60.80:8081/");
        Controls.Add(txtServer);

        Controls.Add(MakeLabel("Username", 135));
        txtUsername = MakeTextBox(160, "IT");
        Controls.Add(txtUsername);

        Controls.Add(MakeLabel("Password", 200));
        txtPassword = MakeTextBox(225, "DCTC.L0c@L");
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

        // Load Employees
        btnLoad = MakeToolbarButton("Load Employees", 10, Color.FromArgb(0, 200, 150), Color.Black);
        btnLoad.Click += async (s, e) => await LoadEmployeesAsync();

        // Export Timecard
        var btnExport = MakeToolbarButton("Export Timecard", 170, Color.FromArgb(0, 150, 255), Color.Black);
        btnExport.Click += (s, e) => new ExportDialog().ShowDialog(this);

        // Timetable Manager
        var btnTimetables = MakeToolbarButton("Timetables", 330, Color.FromArgb(100, 80, 200), Color.White);
        btnTimetables.Click += (s, e) => new TimetableManagerForm().ShowDialog(this);

        lblStatus = new Label
        {
            ForeColor = Color.Silver,
            AutoSize = true,
            Location = new Point(470, 15)
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
                Rate = ""  // filled manually for payroll
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
    List<DepartmentItem> _departments = new();
    TimetableStore _store;

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
        // ── Date range ───────────────────────────────────────
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

        // ── Department selection ─────────────────────────────
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

        // Timetable warning
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

    // ── Load departments from API ─────────────────────────────
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

    // ── Export ────────────────────────────────────────────────
    async Task ExportAsync()
    {
        var selectedRaw = clbDepts.CheckedItems.Cast<string>().ToList();
        if (selectedRaw.Count == 0)
        {
            lblStatus.ForeColor = Color.OrangeRed;
            lblStatus.Text = "Please select at least one department.";
            return;
        }

        // Strip ⚠ suffix
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
            // Fetch punch records for date range
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

            // Filter to selected departments then run calculator
            var filtered = punches.Where(p => selectedDeptNames.Contains(p.Department ?? "")).ToList();
            var calculated = AttendanceCalculator.Calculate(filtered, _store);

            lblStatus.Text = "Building Excel file...";
            Application.DoEvents();

            BuildExcel(calculated, sfd.FileName);

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

    // ── Build Excel ───────────────────────────────────────────
    void BuildExcel(List<CalculatedRecord> records, string filePath)
    {
        using var wb = new XLWorkbook();

        // Column layout:
        // A-D:   Identity (Employee ID, Name, Department, Date)
        // E-H:   Punch times (Clock In, Clock Out, Break Out, Break In)
        // I-L:   Computed attendance (Total OT, Unscheduled, Regular(H), OT1(H))
        // M-S:   DCTC payroll columns (pre-filled from calc; adj columns are yellow)
        string[] headers =
        {
            "Employee ID", "Name", "Department", "Date",
            "Clock In", "Clock Out", "Break Out", "Break In",
            "Total OT(H)", "Unscheduled", "Regular(H)", "OT1(H)",
            "DCTC REG HRS(H)", "DCTC OT SYS ADJ(H)", "DCTC TOTAL OT(H)",
            "DCTC MANUAL ADJ REG(H)", "DCTC MANUAL ADJ OT(H)",
            "TOTAL REG HRS(H)", "TOTAL OT HRS(H)"
        };
        int cols = headers.Length;

        var byDept = records.GroupBy(r => r.Department).OrderBy(g => g.Key);

        foreach (var deptGroup in byDept)
        {
            var ws = AddSheet(wb, deptGroup.Key);
            WriteSheetHeader(ws, deptGroup.Key, headers, cols);
            int row = 7;
            foreach (var rec in deptGroup)
                WriteDataRow(ws, row++, rec, cols);
            FinalizeSheet(ws, cols);
        }

        if (byDept.Count() > 1)
        {
            var ws = wb.Worksheets.Add("ALL DEPARTMENTS");
            ws.TabColor = XLColor.DarkGreen;
            WriteSheetHeader(ws, "All Departments", headers, cols);
            int row = 7;
            foreach (var rec in records)
                WriteDataRow(ws, row++, rec, cols);
            FinalizeSheet(ws, cols);
            ws.Position = 1;
        }

        wb.SaveAs(filePath);
    }

    IXLWorksheet AddSheet(XLWorkbook wb, string deptName)
    {
        var name = (deptName.Length > 31 ? deptName.Substring(0, 31) : deptName);
        foreach (var c in new[] { ':', '\\', '/', '?', '*', '[', ']' })
            name = name.Replace(c.ToString(), "");
        if (string.IsNullOrWhiteSpace(name)) name = "Unknown";
        int suffix = 2;
        string baseName = name;
        while (wb.Worksheets.Any(s => s.Name == name))
            name = $"{baseName.Substring(0, Math.Min(baseName.Length, 28))}_{suffix++}";
        return wb.Worksheets.Add(name);
    }

    void WriteSheetHeader(IXLWorksheet ws, string deptName, string[] headers, int cols)
    {
        ws.Range(1, 1, 1, cols).Merge();
        ws.Cell(1, 1).Value = "DCTC";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 14;

        ws.Range(2, 1, 2, cols).Merge();
        ws.Cell(2, 1).Value = "Total Time Card";
        ws.Cell(2, 1).Style.Font.Bold = true;
        ws.Cell(2, 1).Style.Font.FontSize = 11;

        ws.Range(3, 1, 3, cols).Merge();
        ws.Cell(3, 1).Value = deptName;
        ws.Cell(3, 1).Style.Font.Bold = true;

        ws.Range(4, 1, 4, cols).Merge();
        ws.Cell(4, 1).Value =
            $"{dtpFrom.Value:yyyy-MM-dd}  to  {dtpTo.Value:yyyy-MM-dd}";
        ws.Cell(4, 1).Style.Font.Italic = true;
        ws.Cell(4, 1).Style.Font.FontColor = XLColor.Gray;

        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(6, c + 1);
            cell.Value = headers[c];
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            // Blue = punch data | Teal = computed | Orange = payroll
            cell.Style.Fill.BackgroundColor =
                c < 8 ? XLColor.FromArgb(47, 117, 181) :
                c < 12 ? XLColor.FromArgb(0, 128, 128) :
                         XLColor.FromArgb(180, 95, 6);
        }
        ws.Row(6).Height = 36;
    }

    void WriteDataRow(IXLWorksheet ws, int row, CalculatedRecord rec, int cols)
    {
        // Identity
        ws.Cell(row, 1).Value = rec.EmpCode ?? "";
        ws.Cell(row, 2).Value = rec.FullName ?? "";
        ws.Cell(row, 3).Value = rec.Department ?? "";
        ws.Cell(row, 4).Value = rec.Date ?? "";

        // Punch times
        ws.Cell(row, 5).Value = rec.ClockIn ?? "";
        ws.Cell(row, 6).Value = rec.ClockOut ?? "";
        ws.Cell(row, 7).Value = rec.BreakOut ?? "";
        ws.Cell(row, 8).Value = rec.BreakIn ?? "";

        ws.Cell(row, 9).Value = rec.IsAbsent ? (XLCellValue)"" : (XLCellValue)rec.TotalOTH;
        ws.Cell(row, 10).Value = rec.IsUnscheduled ? "Yes" : "";
        ws.Cell(row, 11).Value = rec.IsAbsent ? (XLCellValue)"" : (XLCellValue)rec.RegularH;
        ws.Cell(row, 12).Value = rec.IsAbsent ? (XLCellValue)"" : (XLCellValue)rec.OT1H;
        ws.Cell(row, 13).Value = rec.IsAbsent ? (XLCellValue)"" : (XLCellValue)rec.RegularH;
        ws.Cell(row, 15).Value = rec.IsAbsent ? (XLCellValue)"" : (XLCellValue)rec.TotalOTH;
        ws.Cell(row, 18).Value = rec.IsAbsent ? (XLCellValue)"" : (XLCellValue)rec.RegularH;
        ws.Cell(row, 19).Value = rec.IsAbsent ? (XLCellValue)"" : (XLCellValue)rec.TotalOTH;

        // Row styling — white background, light border
        var rowRange = ws.Range(row, 1, row, cols);
        rowRange.Style.Fill.BackgroundColor = XLColor.White;
        rowRange.Style.Font.FontColor = XLColor.Black;
        rowRange.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
        rowRange.Style.Border.BottomBorderColor = XLColor.LightGray;

        // Center numeric and time columns
        for (int c = 4; c <= cols; c++)
            ws.Cell(row, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        // Yellow tint on manual adjustment columns
        ws.Cell(row, 14).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
        ws.Cell(row, 16).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
        ws.Cell(row, 17).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);

        // Red text for absent rows
        if (rec.IsAbsent)
            ws.Range(row, 1, row, 4).Style.Font.FontColor = XLColor.OrangeRed;

        // Orange text for unscheduled
        if (rec.IsUnscheduled)
            ws.Cell(row, 10).Style.Font.FontColor = XLColor.OrangeRed;
    }

    void FinalizeSheet(IXLWorksheet ws, int cols)
    {
        ws.Columns().AdjustToContents();
        ws.Column(1).Width = Math.Max(ws.Column(1).Width, 13);
        ws.Column(2).Width = Math.Max(ws.Column(2).Width, 26);
        ws.Column(3).Width = Math.Max(ws.Column(3).Width, 22);
        for (int c = 9; c <= cols; c++)
            ws.Column(c).Width = Math.Max(ws.Column(c).Width, 14);
        ws.SheetView.FreezeRows(6);
        ws.SheetView.FreezeColumns(2);
    }

    // ── Helpers ──────────────────────────────────────────────
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