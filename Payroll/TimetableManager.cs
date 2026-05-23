using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;

// ══════════════════════════════════════════════════════════════
//  DATA MODELS
// ══════════════════════════════════════════════════════════════

class BreakRule
{
    public string Name { get; set; } = "";
    public string StartTime { get; set; } = "12:00";
    public string EndTime { get; set; } = "13:00";
    public int Duration { get; set; } = 60;      // minutes
    public bool PunchRequired { get; set; } = true;    // "Punch Time Is Required"
}

class OvertimeRule
{
    public string Name { get; set; } = "OT1";
    public double HoursFrom { get; set; } = 8.0;
    public double HoursTo { get; set; } = 20.0;
}

class Timetable
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Name { get; set; } = "";

    // ── Basic Settings ───────────────────────────────────────
    public string CheckIn { get; set; } = "08:00";  // scheduled check-in
    public string CheckInStart { get; set; } = "07:00";  // earliest valid check-in
    public string CheckInEnd { get; set; } = "13:00";  // latest valid check-in
    public string CheckOut { get; set; } = "17:00";  // scheduled check-out
    public string CheckOutStart { get; set; } = "15:00";  // earliest valid check-out
    public string CheckOutEnd { get; set; } = "23:59";  // latest valid check-out
    public double WorkDay { get; set; } = 1.0;      // days (1.0 = full day)
    public double RegularHours { get; set; } = 8.0;      // regular hours per day

    // ── Break Settings ───────────────────────────────────────
    public List<BreakRule> Breaks { get; set; } = new();

    // ── Unscheduled Time Settings ────────────────────────────
    public bool EarlyInEnabled { get; set; } = true;
    public string EarlyInAssignTo { get; set; } = "Normal OT";
    public int EarlyInMinimum { get; set; } = 60;   // minutes
    public bool EarlyInCountMinimum { get; set; } = true;

    public bool LateOutEnabled { get; set; } = true;
    public string LateOutAssignTo { get; set; } = "Normal OT";
    public int LateOutMinimum { get; set; } = 60;   // minutes
    public bool LateOutCountMinimum { get; set; } = true;

    // ── Overtime Rules ───────────────────────────────────────
    public bool MaxOTEnabled { get; set; } = true;
    public int MaxOTMinutes { get; set; } = 1440;
    public bool OvertimeEnabled { get; set; } = true;
    public List<OvertimeRule> OvertimeRules { get; set; } = new()
    {
        new OvertimeRule { Name = "OT1", HoursFrom = 8.0,  HoursTo = 20.0 },
        new OvertimeRule { Name = "OT2", HoursFrom = 20.0, HoursTo = 22.0 },
        new OvertimeRule { Name = "OT3", HoursFrom = 22.0, HoursTo = 24.0 }
    };

    // ── Rule Settings ────────────────────────────────────────
    public bool ClockInRequired { get; set; } = true;
    public bool ClockOutRequired { get; set; } = true;
    public int AllowLateIn { get; set; } = 0;   // grace period minutes
    public int AllowEarlyOut { get; set; } = 0;   // grace period minutes
    public string DayChangeTime { get; set; } = "00:00";
    public bool AutoApproveOT { get; set; } = true;
}

class DepartmentAssignment
{
    public string DepartmentName { get; set; } = "";
    public string TimetableId { get; set; } = "";  // empty = unassigned
}

class TimetableStore
{
    public List<Timetable> Timetables { get; set; } = new();
    public List<DepartmentAssignment> Assignments { get; set; } = new();

    static readonly string FilePath = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "timetables.json");

    public static TimetableStore Load()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonConvert.DeserializeObject<TimetableStore>(
                    File.ReadAllText(FilePath)) ?? new TimetableStore();
        }
        catch { }
        return new TimetableStore();
    }

    public void Save()
    {
        File.WriteAllText(FilePath,
            JsonConvert.SerializeObject(this, Formatting.Indented));
    }

    public Timetable GetForDepartment(string deptName)
    {
        var assignment = Assignments.FirstOrDefault(
            a => a.DepartmentName == deptName);
        if (assignment == null || string.IsNullOrEmpty(assignment.TimetableId))
            return null;
        return Timetables.FirstOrDefault(t => t.Id == assignment.TimetableId);
    }
}

// ══════════════════════════════════════════════════════════════
//  TIMETABLE MANAGER FORM
// ══════════════════════════════════════════════════════════════
class TimetableManagerForm : Form
{
    TimetableStore _store;
    ListBox lstTimetables;
    Button btnNew, btnEdit, btnDelete;
    DataGridView dgvAssignments;
    Label lblInfo;

    public TimetableManagerForm()
    {
        Text = "Timetable Manager";
        Size = new Size(900, 600);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(18, 18, 18);
        ForeColor = Color.White;
        _store = TimetableStore.Load();
        BuildUI();
        RefreshAll();
    }

    void BuildUI()
    {
        // ── Left panel: timetable list ───────────────────────
        var leftPanel = new Panel
        {
            Location = new Point(0, 0),
            Size = new Size(280, 600),
            BackColor = Color.FromArgb(25, 25, 25)
        };

        var lblTimetables = new Label
        {
            Text = "Timetables",
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 200, 150),
            Location = new Point(12, 12),
            AutoSize = true
        };

        lstTimetables = new ListBox
        {
            Location = new Point(12, 40),
            Size = new Size(256, 460),
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Segoe UI", 10),
            SelectionMode = SelectionMode.One
        };
        lstTimetables.SelectedIndexChanged += (s, e) => UpdateButtonStates();

        btnNew = MakeButton("+ New", 12, 510, 78);
        btnNew.Click += (s, e) => NewTimetable();

        btnEdit = MakeButton("Edit", 98, 510, 78);
        btnEdit.Click += (s, e) => EditSelected();

        btnDelete = MakeButton("Delete", 184, 510, 78);
        btnDelete.BackColor = Color.FromArgb(180, 40, 40);
        btnDelete.Click += (s, e) => DeleteSelected();

        leftPanel.Controls.AddRange(new Control[]
        {
            lblTimetables, lstTimetables, btnNew, btnEdit, btnDelete
        });

        // ── Right panel: department assignments ──────────────
        var rightPanel = new Panel
        {
            Location = new Point(280, 0),
            Size = new Size(620, 600),
            BackColor = Color.FromArgb(18, 18, 18)
        };

        var lblAssign = new Label
        {
            Text = "Department → Timetable Assignments",
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 200, 150),
            Location = new Point(12, 12),
            AutoSize = true
        };

        lblInfo = new Label
        {
            Text = "Load departments from the dashboard first, then assign timetables here.",
            ForeColor = Color.Silver,
            Font = new Font("Segoe UI", 8, FontStyle.Italic),
            Location = new Point(12, 36),
            Size = new Size(590, 20)
        };

        var btnLoadDepts = MakeButton("Load Departments from API", 12, 60, 220);
        btnLoadDepts.Click += async (s, e) => await LoadDepartmentsAsync();

        var btnSaveAssign = MakeButton("Save Assignments", 246, 60, 160);
        btnSaveAssign.Click += (s, e) => SaveAssignments();

        dgvAssignments = new DataGridView
        {
            Location = new Point(12, 96),
            Size = new Size(590, 470),
            BackgroundColor = Color.FromArgb(22, 22, 22),
            ForeColor = Color.White,
            GridColor = Color.FromArgb(40, 40, 40),
            BorderStyle = BorderStyle.None,
            RowHeadersVisible = false,
            AllowUserToAddRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            Font = new Font("Segoe UI", 9),
            EnableHeadersVisualStyles = false
        };
        dgvAssignments.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.FromArgb(0, 200, 150),
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        dgvAssignments.DefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(28, 28, 28),
            ForeColor = Color.White,
            SelectionBackColor = Color.FromArgb(0, 160, 120),
            SelectionForeColor = Color.White
        };

        // Department column (read-only)
        dgvAssignments.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Department",
            HeaderText = "Department",
            ReadOnly = true,
            FillWeight = 50
        });

        // Timetable dropdown column
        var comboCol = new DataGridViewComboBoxColumn
        {
            Name = "Timetable",
            HeaderText = "Assigned Timetable",
            FillWeight = 50
        };
        dgvAssignments.Columns.Add(comboCol);

        rightPanel.Controls.AddRange(new Control[]
        {
            lblAssign, lblInfo, btnLoadDepts, btnSaveAssign, dgvAssignments
        });

        Controls.Add(leftPanel);
        Controls.Add(rightPanel);
    }

    // ── Refresh ───────────────────────────────────────────────
    void RefreshAll()
    {
        // Refresh timetable list
        lstTimetables.Items.Clear();
        foreach (var t in _store.Timetables)
            lstTimetables.Items.Add(t.Name);

        // Refresh combo options in assignment grid
        RefreshComboOptions();
        UpdateButtonStates();
    }

    void RefreshComboOptions()
    {
        if (dgvAssignments.Columns["Timetable"] is not DataGridViewComboBoxColumn combo)
            return;

        combo.Items.Clear();
        combo.Items.Add("(None)");
        foreach (var t in _store.Timetables)
            combo.Items.Add(t.Name);

        // Update existing rows
        foreach (DataGridViewRow row in dgvAssignments.Rows)
        {
            var current = row.Cells["Timetable"].Value?.ToString();
            if (string.IsNullOrEmpty(current) || !combo.Items.Contains(current))
                row.Cells["Timetable"].Value = "(None)";
        }
    }

    void UpdateButtonStates()
    {
        bool selected = lstTimetables.SelectedIndex >= 0;
        btnEdit.Enabled = selected;
        btnDelete.Enabled = selected;
    }

    // ── Load departments from API ─────────────────────────────
    async System.Threading.Tasks.Task LoadDepartmentsAsync()
    {
        try
        {
            var depts = new List<DepartmentItem>();
            var url = "personnel/api/departments/?page=1&page_size=200";

            while (url != null)
            {
                var json = await ApiClient.GetAsync(url);
                var paged = JsonConvert.DeserializeObject<PagedResponse<DepartmentItem>>(json);
                if (paged?.Data != null) depts.AddRange(paged.Data);
                url = paged?.Next?.Replace(ApiClient.Http.BaseAddress.ToString(), "");
            }

            // Merge with existing assignments
            foreach (var d in depts.OrderBy(d => d.DeptName))
            {
                if (!_store.Assignments.Any(a => a.DepartmentName == d.DeptName))
                    _store.Assignments.Add(new DepartmentAssignment
                    {
                        DepartmentName = d.DeptName,
                        TimetableId = ""
                    });
            }

            RefreshAssignmentGrid();
            lblInfo.Text = $"{depts.Count} departments loaded.";
            lblInfo.ForeColor = Color.FromArgb(0, 200, 150);
        }
        catch (Exception ex)
        {
            lblInfo.Text = $"Error: {ex.Message}";
            lblInfo.ForeColor = Color.OrangeRed;
        }
    }

    void RefreshAssignmentGrid()
    {
        dgvAssignments.Rows.Clear();
        RefreshComboOptions();

        foreach (var assignment in _store.Assignments.OrderBy(a => a.DepartmentName))
        {
            var timetable = _store.Timetables.FirstOrDefault(
                t => t.Id == assignment.TimetableId);
            var timetableName = timetable?.Name ?? "(None)";

            var rowIdx = dgvAssignments.Rows.Add(
                assignment.DepartmentName,
                timetableName);
        }
    }

    void SaveAssignments()
    {
        foreach (DataGridViewRow row in dgvAssignments.Rows)
        {
            var deptName = row.Cells["Department"].Value?.ToString() ?? "";
            var timetableName = row.Cells["Timetable"].Value?.ToString() ?? "(None)";

            var assignment = _store.Assignments.FirstOrDefault(
                a => a.DepartmentName == deptName);
            if (assignment == null) continue;

            var timetable = _store.Timetables.FirstOrDefault(
                t => t.Name == timetableName);
            assignment.TimetableId = timetable?.Id ?? "";
        }

        _store.Save();
        lblInfo.Text = "Assignments saved!";
        lblInfo.ForeColor = Color.FromArgb(0, 200, 150);
    }

    // ── CRUD ──────────────────────────────────────────────────
    void NewTimetable()
    {
        var form = new TimetableEditForm(new Timetable(), _store);
        if (form.ShowDialog() == DialogResult.OK)
        {
            _store.Timetables.Add(form.Result);
            _store.Save();
            RefreshAll();
            lstTimetables.SelectedIndex = lstTimetables.Items.Count - 1;
        }
    }

    void EditSelected()
    {
        if (lstTimetables.SelectedIndex < 0) return;
        var timetable = _store.Timetables[lstTimetables.SelectedIndex];
        var copy = JsonConvert.DeserializeObject<Timetable>(
            JsonConvert.SerializeObject(timetable)); // deep copy

        var form = new TimetableEditForm(copy, _store);
        if (form.ShowDialog() == DialogResult.OK)
        {
            _store.Timetables[lstTimetables.SelectedIndex] = form.Result;
            _store.Save();
            RefreshAll();
        }
    }

    void DeleteSelected()
    {
        if (lstTimetables.SelectedIndex < 0) return;
        var timetable = _store.Timetables[lstTimetables.SelectedIndex];

        // Check if assigned
        var inUse = _store.Assignments.Any(a => a.TimetableId == timetable.Id);
        if (inUse)
        {
            MessageBox.Show(
                "This timetable is assigned to one or more departments.\n" +
                "Please reassign those departments before deleting.",
                "Cannot Delete", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (MessageBox.Show($"Delete \"{timetable.Name}\"?", "Confirm Delete",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        _store.Timetables.RemoveAt(lstTimetables.SelectedIndex);
        _store.Save();
        RefreshAll();
    }

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

// ══════════════════════════════════════════════════════════════
//  TIMETABLE EDIT FORM
// ══════════════════════════════════════════════════════════════
class TimetableEditForm : Form
{
    public Timetable Result { get; private set; }
    Timetable _t;
    TabControl tabs;

    // Basic Settings controls
    TextBox txtName, txtCheckIn, txtCheckInStart, txtCheckInEnd;
    TextBox txtCheckOut, txtCheckOutStart, txtCheckOutEnd;
    NumericUpDown nudWorkDay, nudRegularHours;

    // Break controls
    DataGridView dgvBreaks;

    // Unscheduled controls
    CheckBox chkEarlyIn, chkEarlyInCountMin, chkLateOut, chkLateOutCountMin;
    ComboBox cmbEarlyInAssign, cmbLateOutAssign;
    NumericUpDown nudEarlyInMin, nudLateOutMin;

    // OT Rule controls
    CheckBox chkMaxOT, chkOTEnabled;
    NumericUpDown nudMaxOT;
    DataGridView dgvOTRules;

    // Rule Settings controls
    CheckBox chkClockInReq, chkClockOutReq, chkAutoApprove;
    NumericUpDown nudLateIn, nudEarlyOut;
    TextBox txtDayChange;

    public TimetableEditForm(Timetable timetable, TimetableStore store)
    {
        _t = timetable;
        Result = timetable;
        Text = string.IsNullOrEmpty(timetable.Name) ? "New Timetable" : $"Edit — {timetable.Name}";
        Size = new Size(700, 560);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(18, 18, 18);
        ForeColor = Color.White;
        BuildUI();
        LoadValues();
    }

    void BuildUI()
    {
        // Name row at top
        var lblName = MakeLabel("Name *", 14, 12);
        txtName = MakeTextBox(80, 14, 300);

        var lblAssignTo = MakeLabel("Work Hours Assign To *", 400, 14);
        var txtAssignTo = MakeTextBox(570, 14, 100);
        txtAssignTo.Text = "Regular";
        txtAssignTo.Enabled = false;

        tabs = new TabControl
        {
            Location = new Point(8, 44),
            Size = new Size(668, 440),
            Font = new Font("Segoe UI", 9)
        };

        tabs.TabPages.Add(BuildBasicTab());
        tabs.TabPages.Add(BuildBreakTab());
        tabs.TabPages.Add(BuildUnscheduledTab());
        tabs.TabPages.Add(BuildOTRuleTab());
        tabs.TabPages.Add(BuildRuleSettingsTab());

        // Bottom buttons
        var btnConfirm = new Button
        {
            Text = "Confirm",
            Location = new Point(490, 492),
            Size = new Size(90, 30),
            BackColor = Color.FromArgb(0, 200, 150),
            ForeColor = Color.Black,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        btnConfirm.FlatAppearance.BorderSize = 0;
        btnConfirm.Click += BtnConfirm_Click;

        var btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(592, 492),
            Size = new Size(90, 30),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

        Controls.AddRange(new Control[]
        {
            lblName, txtName, lblAssignTo, txtAssignTo,
            tabs, btnConfirm, btnCancel
        });
    }

    // ── Basic Settings Tab ────────────────────────────────────
    TabPage BuildBasicTab()
    {
        var page = MakePage("Basic Settings");

        // Check-In group
        AddLabel(page, "Check-In *", 12, 16);
        txtCheckIn = AddTextBox(page, 100, 16, 110);

        AddLabel(page, "Check-Out *", 350, 16);
        txtCheckOut = AddTextBox(page, 440, 16, 110);

        AddLabel(page, "Check-In Start *", 12, 52);
        txtCheckInStart = AddTextBox(page, 130, 52, 110);

        AddLabel(page, "Check-Out Start *", 350, 52);
        txtCheckOutStart = AddTextBox(page, 470, 52, 110);

        AddLabel(page, "Check-In End *", 12, 88);
        txtCheckInEnd = AddTextBox(page, 130, 88, 110);

        AddLabel(page, "Check-Out End *", 350, 88);
        txtCheckOutEnd = AddTextBox(page, 470, 88, 110);

        AddLabel(page, "WorkDay *", 12, 130);
        nudWorkDay = new NumericUpDown
        {
            Location = new Point(100, 127),
            Size = new Size(80, 24),
            DecimalPlaces = 1,
            Minimum = 0.5m,
            Maximum = 2.0m,
            Increment = 0.5m,
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White
        };
        page.Controls.Add(nudWorkDay);

        AddLabel(page, "Regular Hours/Day *", 12, 165);
        nudRegularHours = new NumericUpDown
        {
            Location = new Point(160, 162),
            Size = new Size(80, 24),
            DecimalPlaces = 1,
            Minimum = 1,
            Maximum = 24,
            Increment = 0.5m,
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White
        };
        page.Controls.Add(nudRegularHours);

        var notice = new Label
        {
            Text = "*Notice\n1. Times in HH:mm format (e.g. 08:00)\n2. Check-In/Out are the scheduled times.\n   Start/End define the valid punch window.",
            ForeColor = Color.Silver,
            Font = new Font("Segoe UI", 8, FontStyle.Italic),
            Location = new Point(12, 210),
            Size = new Size(600, 60)
        };
        page.Controls.Add(notice);

        return page;
    }

    // ── Break Time Tab ────────────────────────────────────────
    TabPage BuildBreakTab()
    {
        var page = MakePage("BreakTime Settings");

        dgvBreaks = new DataGridView
        {
            Location = new Point(8, 8),
            Size = new Size(638, 320),
            BackgroundColor = Color.FromArgb(28, 28, 28),
            ForeColor = Color.White,
            GridColor = Color.FromArgb(40, 40, 40),
            BorderStyle = BorderStyle.None,
            RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            Font = new Font("Segoe UI", 9),
            EnableHeadersVisualStyles = false
        };
        dgvBreaks.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.FromArgb(0, 200, 150),
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        dgvBreaks.DefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(28, 28, 28),
            ForeColor = Color.White,
            SelectionBackColor = Color.FromArgb(0, 160, 120),
            SelectionForeColor = Color.White
        };

        dgvBreaks.Columns.Add(new DataGridViewTextBoxColumn
        { Name = "BreakName", HeaderText = "Name", FillWeight = 25 });
        dgvBreaks.Columns.Add(new DataGridViewTextBoxColumn
        { Name = "StartTime", HeaderText = "Start Time", FillWeight = 20 });
        dgvBreaks.Columns.Add(new DataGridViewTextBoxColumn
        { Name = "EndTime", HeaderText = "End Time", FillWeight = 20 });
        dgvBreaks.Columns.Add(new DataGridViewTextBoxColumn
        { Name = "Duration", HeaderText = "Duration (min)", FillWeight = 15 });
        dgvBreaks.Columns.Add(new DataGridViewCheckBoxColumn
        { Name = "PunchReq", HeaderText = "Punch Required", FillWeight = 20 });

        var btnAddBreak = new Button
        {
            Text = "+ Add Break",
            Location = new Point(8, 336),
            Size = new Size(120, 28),
            BackColor = Color.FromArgb(0, 200, 150),
            ForeColor = Color.Black,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        btnAddBreak.FlatAppearance.BorderSize = 0;
        btnAddBreak.Click += (s, e) =>
            dgvBreaks.Rows.Add("Break", "12:00", "13:00", 60, true);

        var btnDelBreak = new Button
        {
            Text = "Delete Selected",
            Location = new Point(140, 336),
            Size = new Size(130, 28),
            BackColor = Color.FromArgb(180, 40, 40),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        btnDelBreak.FlatAppearance.BorderSize = 0;
        btnDelBreak.Click += (s, e) =>
        {
            if (dgvBreaks.CurrentRow != null)
                dgvBreaks.Rows.Remove(dgvBreaks.CurrentRow);
        };

        page.Controls.AddRange(new Control[] { dgvBreaks, btnAddBreak, btnDelBreak });
        return page;
    }

    // ── Unscheduled Time Tab ──────────────────────────────────
    TabPage BuildUnscheduledTab()
    {
        var page = MakePage("Unscheduled Time Settings");
        var assignOptions = new[] { "Normal OT", "OT1", "OT2", "OT3", "Regular" };

        // Early In
        chkEarlyIn = new CheckBox
        {
            Text = "Early In",
            Location = new Point(12, 16),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            AutoSize = true
        };
        AddLabel(page, "Work Hours Assign To", 12, 48);
        cmbEarlyInAssign = new ComboBox
        {
            Location = new Point(170, 45),
            Size = new Size(150, 24),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        cmbEarlyInAssign.Items.AddRange(assignOptions);

        AddLabel(page, "Minimum (min)", 12, 80);
        nudEarlyInMin = new NumericUpDown
        {
            Location = new Point(120, 77),
            Size = new Size(80, 24),
            Minimum = 1,
            Maximum = 480,
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White
        };
        chkEarlyInCountMin = new CheckBox
        {
            Text = "Count the Minimum",
            Location = new Point(220, 80),
            ForeColor = Color.White,
            AutoSize = true
        };

        // Separator
        var sep = new Label
        {
            Location = new Point(12, 115),
            Size = new Size(630, 1),
            BackColor = Color.FromArgb(50, 50, 50)
        };

        // Late Out
        chkLateOut = new CheckBox
        {
            Text = "Late Out",
            Location = new Point(12, 130),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            AutoSize = true
        };
        AddLabel(page, "Work Hours Assign To", 12, 162);
        cmbLateOutAssign = new ComboBox
        {
            Location = new Point(170, 159),
            Size = new Size(150, 24),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        cmbLateOutAssign.Items.AddRange(assignOptions);

        AddLabel(page, "Minimum (min)", 12, 195);
        nudLateOutMin = new NumericUpDown
        {
            Location = new Point(120, 192),
            Size = new Size(80, 24),
            Minimum = 1,
            Maximum = 480,
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White
        };
        chkLateOutCountMin = new CheckBox
        {
            Text = "Count the Minimum",
            Location = new Point(220, 195),
            ForeColor = Color.White,
            AutoSize = true
        };

        page.Controls.AddRange(new Control[]
        {
            chkEarlyIn, cmbEarlyInAssign, nudEarlyInMin, chkEarlyInCountMin,
            sep,
            chkLateOut, cmbLateOutAssign, nudLateOutMin, chkLateOutCountMin
        });
        return page;
    }

    // ── Overtime Rule Tab ─────────────────────────────────────
    TabPage BuildOTRuleTab()
    {
        var page = MakePage("Overtime Rule");

        chkMaxOT = new CheckBox
        {
            Text = "Max OT",
            Location = new Point(12, 16),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            AutoSize = true
        };
        AddLabel(page, "Max OT Duration (min)", 140, 16);
        nudMaxOT = new NumericUpDown
        {
            Location = new Point(290, 13),
            Size = new Size(100, 24),
            Minimum = 0,
            Maximum = 1440,
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White
        };

        chkOTEnabled = new CheckBox
        {
            Text = "Overtime Enabled",
            Location = new Point(12, 48),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            AutoSize = true
        };

        // OT rules grid
        dgvOTRules = new DataGridView
        {
            Location = new Point(8, 80),
            Size = new Size(638, 220),
            BackgroundColor = Color.FromArgb(28, 28, 28),
            ForeColor = Color.White,
            GridColor = Color.FromArgb(40, 40, 40),
            BorderStyle = BorderStyle.None,
            RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            Font = new Font("Segoe UI", 9),
            EnableHeadersVisualStyles = false
        };
        dgvOTRules.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.FromArgb(0, 200, 150),
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        dgvOTRules.DefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(28, 28, 28),
            ForeColor = Color.White,
            SelectionBackColor = Color.FromArgb(0, 160, 120),
            SelectionForeColor = Color.White
        };

        dgvOTRules.Columns.Add(new DataGridViewTextBoxColumn
        { Name = "OTName", HeaderText = "Work Hours Assign To", FillWeight = 30 });
        dgvOTRules.Columns.Add(new DataGridViewTextBoxColumn
        { Name = "HoursFrom", HeaderText = "Hours From", FillWeight = 35 });
        dgvOTRules.Columns.Add(new DataGridViewTextBoxColumn
        { Name = "HoursTo", HeaderText = "Hours To", FillWeight = 35 });

        var btnAddOT = new Button
        {
            Text = "+ Add OT Rule",
            Location = new Point(8, 308),
            Size = new Size(130, 28),
            BackColor = Color.FromArgb(0, 200, 150),
            ForeColor = Color.Black,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        btnAddOT.FlatAppearance.BorderSize = 0;
        btnAddOT.Click += (s, e) =>
            dgvOTRules.Rows.Add($"OT{dgvOTRules.Rows.Count + 1}", "8.0", "20.0");

        var btnDelOT = new Button
        {
            Text = "Delete Selected",
            Location = new Point(150, 308),
            Size = new Size(130, 28),
            BackColor = Color.FromArgb(180, 40, 40),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        btnDelOT.FlatAppearance.BorderSize = 0;
        btnDelOT.Click += (s, e) =>
        {
            if (dgvOTRules.CurrentRow != null)
                dgvOTRules.Rows.Remove(dgvOTRules.CurrentRow);
        };

        page.Controls.AddRange(new Control[]
        {
            chkMaxOT, nudMaxOT, chkOTEnabled, dgvOTRules, btnAddOT, btnDelOT
        });
        return page;
    }

    // ── Rule Settings Tab ─────────────────────────────────────
    TabPage BuildRuleSettingsTab()
    {
        var page = MakePage("Rule Settings");

        chkClockInReq = new CheckBox
        {
            Text = "Clock-In Is Required",
            Location = new Point(12, 16),
            ForeColor = Color.White,
            AutoSize = true
        };
        chkClockOutReq = new CheckBox
        {
            Text = "Clock-Out Is Required",
            Location = new Point(320, 16),
            ForeColor = Color.White,
            AutoSize = true
        };

        AddLabel(page, "Allow Late-In (min)", 12, 52);
        nudLateIn = new NumericUpDown
        {
            Location = new Point(160, 49),
            Size = new Size(80, 24),
            Minimum = 0,
            Maximum = 120,
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White
        };

        AddLabel(page, "Allow Early-Out (min)", 320, 52);
        nudEarlyOut = new NumericUpDown
        {
            Location = new Point(490, 49),
            Size = new Size(80, 24),
            Minimum = 0,
            Maximum = 120,
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White
        };

        AddLabel(page, "Day Change Time", 12, 88);
        txtDayChange = new TextBox
        {
            Location = new Point(160, 85),
            Size = new Size(100, 24),
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Text = "00:00"
        };

        chkAutoApprove = new CheckBox
        {
            Text = "Overtime Auto Approved",
            Location = new Point(12, 124),
            ForeColor = Color.White,
            AutoSize = true
        };

        page.Controls.AddRange(new Control[]
        {
            chkClockInReq, chkClockOutReq,
            nudLateIn, nudEarlyOut, txtDayChange, chkAutoApprove
        });
        return page;
    }

    // ── Load values into controls ─────────────────────────────
    void LoadValues()
    {
        txtName.Text = _t.Name;
        txtCheckIn.Text = _t.CheckIn;
        txtCheckInStart.Text = _t.CheckInStart;
        txtCheckInEnd.Text = _t.CheckInEnd;
        txtCheckOut.Text = _t.CheckOut;
        txtCheckOutStart.Text = _t.CheckOutStart;
        txtCheckOutEnd.Text = _t.CheckOutEnd;
        nudWorkDay.Value = (decimal)_t.WorkDay;
        nudRegularHours.Value = (decimal)_t.RegularHours;

        // Breaks
        foreach (var b in _t.Breaks)
            dgvBreaks.Rows.Add(b.Name, b.StartTime, b.EndTime, b.Duration, b.PunchRequired);

        // Unscheduled
        chkEarlyIn.Checked = _t.EarlyInEnabled;
        cmbEarlyInAssign.Text = _t.EarlyInAssignTo;
        nudEarlyInMin.Value = _t.EarlyInMinimum;
        chkEarlyInCountMin.Checked = _t.EarlyInCountMinimum;
        chkLateOut.Checked = _t.LateOutEnabled;
        cmbLateOutAssign.Text = _t.LateOutAssignTo;
        nudLateOutMin.Value = _t.LateOutMinimum;
        chkLateOutCountMin.Checked = _t.LateOutCountMinimum;

        // OT Rules
        chkMaxOT.Checked = _t.MaxOTEnabled;
        nudMaxOT.Value = _t.MaxOTMinutes;
        chkOTEnabled.Checked = _t.OvertimeEnabled;
        foreach (var r in _t.OvertimeRules)
            dgvOTRules.Rows.Add(r.Name, r.HoursFrom.ToString("F1"), r.HoursTo.ToString("F1"));

        // Rule Settings
        chkClockInReq.Checked = _t.ClockInRequired;
        chkClockOutReq.Checked = _t.ClockOutRequired;
        nudLateIn.Value = _t.AllowLateIn;
        nudEarlyOut.Value = _t.AllowEarlyOut;
        txtDayChange.Text = _t.DayChangeTime;
        chkAutoApprove.Checked = _t.AutoApproveOT;
    }

    // ── Save values from controls ─────────────────────────────
    void BtnConfirm_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtName.Text))
        {
            MessageBox.Show("Please enter a timetable name.", "Validation",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _t.Name = txtName.Text.Trim();
        _t.CheckIn = txtCheckIn.Text.Trim();
        _t.CheckInStart = txtCheckInStart.Text.Trim();
        _t.CheckInEnd = txtCheckInEnd.Text.Trim();
        _t.CheckOut = txtCheckOut.Text.Trim();
        _t.CheckOutStart = txtCheckOutStart.Text.Trim();
        _t.CheckOutEnd = txtCheckOutEnd.Text.Trim();
        _t.WorkDay = (double)nudWorkDay.Value;
        _t.RegularHours = (double)nudRegularHours.Value;

        // Breaks
        _t.Breaks.Clear();
        foreach (DataGridViewRow row in dgvBreaks.Rows)
        {
            if (row.IsNewRow) continue;
            _t.Breaks.Add(new BreakRule
            {
                Name = row.Cells["BreakName"].Value?.ToString() ?? "",
                StartTime = row.Cells["StartTime"].Value?.ToString() ?? "",
                EndTime = row.Cells["EndTime"].Value?.ToString() ?? "",
                Duration = int.TryParse(row.Cells["Duration"].Value?.ToString(), out int d) ? d : 60,
                PunchRequired = row.Cells["PunchReq"].Value is true
            });
        }

        // Unscheduled
        _t.EarlyInEnabled = chkEarlyIn.Checked;
        _t.EarlyInAssignTo = cmbEarlyInAssign.Text;
        _t.EarlyInMinimum = (int)nudEarlyInMin.Value;
        _t.EarlyInCountMinimum = chkEarlyInCountMin.Checked;
        _t.LateOutEnabled = chkLateOut.Checked;
        _t.LateOutAssignTo = cmbLateOutAssign.Text;
        _t.LateOutMinimum = (int)nudLateOutMin.Value;
        _t.LateOutCountMinimum = chkLateOutCountMin.Checked;

        // OT Rules
        _t.MaxOTEnabled = chkMaxOT.Checked;
        _t.MaxOTMinutes = (int)nudMaxOT.Value;
        _t.OvertimeEnabled = chkOTEnabled.Checked;
        _t.OvertimeRules.Clear();
        foreach (DataGridViewRow row in dgvOTRules.Rows)
        {
            if (row.IsNewRow) continue;
            _t.OvertimeRules.Add(new OvertimeRule
            {
                Name = row.Cells["OTName"].Value?.ToString() ?? "",
                HoursFrom = double.TryParse(row.Cells["HoursFrom"].Value?.ToString(), out double f) ? f : 8.0,
                HoursTo = double.TryParse(row.Cells["HoursTo"].Value?.ToString(), out double t) ? t : 20.0
            });
        }

        // Rule Settings
        _t.ClockInRequired = chkClockInReq.Checked;
        _t.ClockOutRequired = chkClockOutReq.Checked;
        _t.AllowLateIn = (int)nudLateIn.Value;
        _t.AllowEarlyOut = (int)nudEarlyOut.Value;
        _t.DayChangeTime = txtDayChange.Text.Trim();
        _t.AutoApproveOT = chkAutoApprove.Checked;

        Result = _t;
        DialogResult = DialogResult.OK;
        Close();
    }

    // ── Helpers ──────────────────────────────────────────────
    TabPage MakePage(string title)
    {
        var page = new TabPage(title)
        {
            BackColor = Color.FromArgb(22, 22, 22),
            ForeColor = Color.White
        };
        return page;
    }

    Label MakeLabel(string text, int x, int y)
    {
        return new Label
        {
            Text = text,
            Location = new Point(x, y),
            ForeColor = Color.Silver,
            AutoSize = true,
            Font = new Font("Segoe UI", 9)
        };
    }

    TextBox MakeTextBox(int x, int y, int width)
    {
        return new TextBox
        {
            Location = new Point(x, y),
            Size = new Size(width, 24),
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Segoe UI", 9)
        };
    }

    void AddLabel(TabPage page, string text, int x, int y)
    {
        page.Controls.Add(new Label
        {
            Text = text,
            Location = new Point(x, y),
            ForeColor = Color.Silver,
            AutoSize = true,
            Font = new Font("Segoe UI", 9)
        });
    }

    TextBox AddTextBox(TabPage page, int x, int y, int width)
    {
        var tb = new TextBox
        {
            Location = new Point(x, y),
            Size = new Size(width, 24),
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Segoe UI", 9)
        };
        page.Controls.Add(tb);
        return tb;
    }
}