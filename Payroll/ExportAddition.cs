using ClosedXML.Excel;
using Newtonsoft.Json;
using Payroll;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

// ══════════════════════════════════════════════════════════════
//  PUNCH RECORD  (from iclock/api/transactions/)
// ══════════════════════════════════════════════════════════════
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

        // Timetable warning label
        var lblTimetable = new Label
        {
            Text = "ℹ Departments without a timetable assigned will show raw punch times only.",
            ForeColor = Color.FromArgb(255, 200, 0),
            Font = new Font("Segoe UI", 8, FontStyle.Italic),
            Location = new Point(16, 424),
            Size = new Size(432, 32)
        };
        Controls.Add(lblTimetable);

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

    // ── Load departments ──────────────────────────────────────
    async System.Threading.Tasks.Task LoadDepartmentsAsync()
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
    async System.Threading.Tasks.Task ExportAsync()
    {
        var selectedRaw = clbDepts.CheckedItems.Cast<string>().ToList();
        if (selectedRaw.Count == 0)
        {
            lblStatus.ForeColor = Color.OrangeRed;
            lblStatus.Text = "Please select at least one department.";
            return;
        }

        // Strip the ⚠ suffix to get real dept names
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
            // Fetch all punch records for the date range
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

            lblStatus.Text = $"Calculating attendance for {punches.Count} punch records...";
            Application.DoEvents();

            // Filter to selected departments
            var filtered = punches
                .Where(p => selectedDeptNames.Contains(p.Department ?? ""))
                .ToList();

            // Run calculation engine
            var calculated = AttendanceCalculator.Calculate(filtered, _store);

            lblStatus.Text = "Building Excel file...";
            Application.DoEvents();

            BuildExcel(calculated, sfd.FileName);

            lblStatus.ForeColor = Color.FromArgb(0, 200, 150);
            lblStatus.Text =
                $"Done! {calculated.Count} rows across {selectedDeptNames.Count} department(s).";

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

        var byDept = records
            .GroupBy(r => r.Department)
            .OrderBy(g => g.Key);

        // Column layout
        // A  Employee ID       B  Name             C  Department      D  Date
        // E  Clock In          F  Clock Out         G  Break Out       H  Break In
        // I  Total OT(H)       J  Unscheduled       K  Regular(H)      L  OT1(H)
        // M  DCTC REG HRS(H)   N  DCTC OT SYS ADJ  O  DCTC TOTAL OT
        // P  DCTC MAN ADJ REG  Q  DCTC MAN ADJ OT  R  TOTAL REG HRS   S  TOTAL OT HRS

        string[] headers =
        {
            "Employee ID", "Name", "Department", "Date",
            "Clock In", "Clock Out", "Break Out", "Break In",
            "Total OT(H)", "Unscheduled", "Regular(H)", "OT1(H)",
            "DCTC REG HRS(H)", "DCTC OT SYS ADJ(H)", "DCTC TOTAL OT(H)",
            "DCTC MANUAL ADJ REG(H)", "DCTC MANUAL ADJ OT(H)",
            "TOTAL REG HRS(H)", "TOTAL OT HRS(H)"
        };

        int totalCols = headers.Length; // 19

        foreach (var deptGroup in byDept)
        {
            var ws = AddSheet(wb, deptGroup.Key);
            WriteSheetHeader(ws, deptGroup.Key, headers, totalCols);

            int row = 7;
            foreach (var rec in deptGroup)
            {
                WriteDataRow(ws, row, rec, totalCols);
                row++;
            }

            FinalizeSheet(ws, totalCols);
        }

        // Summary sheet if multiple departments
        if (byDept.Count() > 1)
        {
            var ws = wb.Worksheets.Add("ALL DEPARTMENTS");
            ws.TabColor = XLColor.DarkGreen;
            WriteSheetHeader(ws, "All Departments", headers, totalCols);

            int row = 7;
            foreach (var rec in records)
            {
                WriteDataRow(ws, row, rec, totalCols);
                row++;
            }

            FinalizeSheet(ws, totalCols);
            ws.Position = 1;
        }

        wb.SaveAs(filePath);
    }

    IXLWorksheet AddSheet(XLWorkbook wb, string deptName)
    {
        var name = deptName.Length > 31 ? deptName.Substring(0, 31) : deptName;
        foreach (var c in new[] { ':', '\\', '/', '?', '*', '[', ']' })
            name = name.Replace(c.ToString(), "");
        if (string.IsNullOrWhiteSpace(name)) name = "Unknown";
        int suffix = 2;
        string baseName = name;
        while (wb.Worksheets.Any(s => s.Name == name))
            name = $"{baseName.Substring(0, Math.Min(baseName.Length, 28))}_{suffix++}";
        return wb.Worksheets.Add(name);
    }

    void WriteSheetHeader(IXLWorksheet ws, string deptName, string[] headers, int totalCols)
    {
        ws.Range(1, 1, 1, totalCols).Merge();
        ws.Cell(1, 1).Value = "DCTC";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 14;

        ws.Range(2, 1, 2, totalCols).Merge();
        ws.Cell(2, 1).Value = "Total Time Card";
        ws.Cell(2, 1).Style.Font.Bold = true;
        ws.Cell(2, 1).Style.Font.FontSize = 11;

        ws.Range(3, 1, 3, totalCols).Merge();
        ws.Cell(3, 1).Value = deptName;
        ws.Cell(3, 1).Style.Font.Bold = true;

        ws.Range(4, 1, 4, totalCols).Merge();
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

            // Blue = punch data, Teal = computed attendance, Orange = payroll
            if (c < 8)
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(47, 117, 181);
            else if (c < 12)
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 128);
            else
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(180, 95, 6);
        }

        ws.Row(6).Height = 36;
    }

    void WriteDataRow(IXLWorksheet ws, int row, CalculatedRecord rec, int totalCols)
    {
        // A-D: identity
        ws.Cell(row, 1).Value = rec.EmpCode ?? "";
        ws.Cell(row, 2).Value = rec.FullName ?? "";
        ws.Cell(row, 3).Value = rec.Department ?? "";
        ws.Cell(row, 4).Value = rec.Date ?? "";

        // E-H: punch times
        ws.Cell(row, 5).Value = rec.ClockIn ?? "";
        ws.Cell(row, 6).Value = rec.ClockOut ?? "";
        ws.Cell(row, 7).Value = rec.BreakOut ?? "";
        ws.Cell(row, 8).Value = rec.BreakIn ?? "";

        // I-L: computed attendance (from calculator)
        ws.Cell(row, 9).Value = rec.IsAbsent ? "" : (object)rec.TotalOTH;
        ws.Cell(row, 10).Value = rec.IsUnscheduled ? "Yes" : "";
        ws.Cell(row, 11).Value = rec.IsAbsent ? "" : (object)rec.RegularH;
        ws.Cell(row, 12).Value = rec.IsAbsent ? "" : (object)rec.OT1H;

        // M-S: DCTC payroll columns — yellow, fill manually
        // Pre-fill DCTC REG and DCTC TOTAL OT from computed values as a starting point
        ws.Cell(row, 13).Value = rec.IsAbsent ? "" : (object)rec.RegularH;   // DCTC REG HRS
        ws.Cell(row, 14).Value = "";                                           // DCTC OT SYS ADJ
        ws.Cell(row, 15).Value = rec.IsAbsent ? "" : (object)rec.TotalOTH;   // DCTC TOTAL OT
        ws.Cell(row, 16).Value = "";                                           // DCTC MANUAL ADJ REG
        ws.Cell(row, 17).Value = "";                                           // DCTC MANUAL ADJ OT
        ws.Cell(row, 18).Value = rec.IsAbsent ? "" : (object)rec.RegularH;   // TOTAL REG HRS
        ws.Cell(row, 19).Value = rec.IsAbsent ? "" : (object)rec.TotalOTH;   // TOTAL OT HRS

        // Row styling
        var rowRange = ws.Range(row, 1, row, totalCols);
        rowRange.Style.Fill.BackgroundColor = XLColor.White;
        rowRange.Style.Font.FontColor = XLColor.Black;
        rowRange.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
        rowRange.Style.Border.BottomBorderColor = XLColor.LightGray;

        // Center numeric and time columns
        for (int c = 4; c <= totalCols; c++)
            ws.Cell(row, c).Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

        // Yellow tint on manually-adjustable payroll columns (N, P, Q)
        ws.Cell(row, 14).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
        ws.Cell(row, 16).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);
        ws.Cell(row, 17).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 200);

        // Red text for absent rows
        if (rec.IsAbsent)
            ws.Range(row, 1, row, 4).Style.Font.FontColor = XLColor.OrangeRed;

        // Flag unscheduled in orange
        if (rec.IsUnscheduled)
            ws.Cell(row, 10).Style.Font.FontColor = XLColor.OrangeRed;
    }

    void FinalizeSheet(IXLWorksheet ws, int totalCols)
    {
        ws.Columns().AdjustToContents();
        ws.Column(1).Width = Math.Max(ws.Column(1).Width, 13);
        ws.Column(2).Width = Math.Max(ws.Column(2).Width, 26);
        ws.Column(3).Width = Math.Max(ws.Column(3).Width, 22);
        for (int c = 9; c <= totalCols; c++)
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
