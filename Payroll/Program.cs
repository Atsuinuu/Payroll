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
using System.IO;
using ClosedXML.Excel;

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

// ── Login Form ────────────────────────────────────────────────
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

// ── Dashboard Form ────────────────────────────────────────────
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

        btnLoad = new Button
        {
            Text = "Load Employees",
            Location = new Point(10, 10),
            Size = new Size(150, 28),
            BackColor = Color.FromArgb(0, 200, 150),
            ForeColor = Color.Black,
            FlatStyle = FlatStyle.Flat
        };
        btnLoad.FlatAppearance.BorderSize = 0;
        btnLoad.Click += async (s, e) => await LoadEmployeesAsync();

        var btnExport = new Button
        {
            Text = "Export Timecard",
            Location = new Point(170, 10),
            Size = new Size(150, 28),
            BackColor = Color.FromArgb(0, 150, 255),
            ForeColor = Color.Black,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        btnExport.FlatAppearance.BorderSize = 0;
        btnExport.Click += (s, e) => new ExportDialog().ShowDialog(this);

        lblStatus = new Label
        {
            ForeColor = Color.Silver,
            AutoSize = true,
            Location = new Point(340, 15)
        };

        toolbar.Controls.Add(btnLoad);
        toolbar.Controls.Add(btnExport);
        toolbar.Controls.Add(lblStatus);

        var btnTimetables = new Button
        {
            Text = "Timetables",
            Location = new Point(330, 10),
            Size = new Size(120, 28),
            BackColor = Color.FromArgb(100, 80, 200),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        btnTimetables.FlatAppearance.BorderSize = 0;
        btnTimetables.Click += (s, e) => new TimetableManagerForm().ShowDialog(this);
        toolbar.Controls.Add(btnTimetables);

        //⣿⣷⣀⡀⢠⠀⠈⠙⠻⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣇⠙⣶⠀⢀⠀⠤⣉⠛⠿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣷⣤⠣⢈⢎⡤⢈⠓⢦⣌⡙⠿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣏⠻⣿⣆⡸⢜⡡⢊⠀⠻⣿⣿⣾⣷⣽⣧⢎⡻⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣦⡍⢻⣷⣌⠖⠀⠀⠀⠈⢙⠛⣿⣿⣿⣿⣶⣧⣖⣯⡙⢮⡷⣯⢿⡿⣟⡟⢯⠻⡍⠶⡛⢿⠛⠿⣿⠟⡿⠿⠿⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠿⣿⢿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣷⠊⢿⣿⣄⠐⠀⢂⠈⠰⡘⢿⣏⠻⣿⣿⠟⣆⢏⢇⠳⢎⠷⡹⠬⡝⢪⠑⠌⠱⠌⠂⠉⠒⢠⠂⢈⠁⣂⡄⠈⡐⠀⢀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠁⠀⠉⠉⠁⠀⠀⠈⠛⠿⡿⠿⠻⠿⠿⠿⠿⠿⠿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⠋⣋⡤⣖⡻⢌⡠⠎⣠⣾⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⡿⣿⡿⢿⣟⣿⣦⡀⠘⡀⠀⠀⠘⠢⠘⢄⠉⠲⠌⠊⠠⠌⢪⠁⠘⡤⠐⡆⠈⢠⠂⢌⠣⠀⢩⡗⣾⣇⢻⠙⠄⡀⢁⠂⠠⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠄⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠉⠉⠉⠙⠛⠻⢿⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠿⢛⣋⣭⡖⣠⡾⢉⡰⡡⠙⢊⣴⣾⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣦⡒⢽⣮⣟⢷⣦⣄⡀⠀⠂⠄⠈⡀⠁⠈⠀⠐⠈⠀⠍⠆⢱⠀⠁⡐⣄⢹⡀⢷⡆⣅⢃⠈⠀⠀⠀⠀⠐⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠠⠀⠌⠂⠱⢠⢨⠌⡐⣱⠟⢩⣻⠋⢸⣿⣿⣿⣿⣿⠿⡿⠿⠿⠟⢛⣛⣿⣭⣴⣶⠞⢛⡛⢉⣼⠋⠔⠋⠁⣠⣶⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣄⣈⠻⣧⣙⠟⣿⢷⣦⣄⡀⠀⠀⠀⠀⢀⠀⠀⣠⡀⣀⡰⢷⡌⠺⣏⢳⠀⠙⠈⡜⢄⠘⢂⢠⠂⠰⠆⠀⠀⠂⠀⠀⠀⠀⠀⠲⡇⠘⡀⠠⡐⠀⡌⠂⠠⠀⠀⠀⠀⠀⠀⡀⠀⠀⠀⠀⠀⠐⠀⡄⠀⠙⠈⡌⠐⠰⢣⡾⢓⢣⡠⠃⢈⣄⠠⠡⠐⡠⡔⢲⠆⣿⣿⣿⣿⣿⣿⠍⣾⣿⠃⡩⠒⢀⠔⣠⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠛⠻⠦⣭⢛⠪⣷⡙⠿⣿⣿⡝⣶⣶⢦⡻⣶⣬⢿⡾⢏⠣⣅⠈⠯⠠⠁⢂⢣⡜⠎⢇⣄⣄⠃⠀⠀⠀⠀⠀⠀⡀⠐⡠⠁⢁⠀⠀⠁⡀⠄⠀⢱⡐⠀⠰⠀⡅⠁⠀⠈⡄⡀⠐⠀⠆⠀⠄⠀⠀⠻⠏⣴⣄⠈⠄⣸⢠⠃⡘⢡⠀⡞⡩⡀⢰⡫⠐⢠⠋⣠⢞⡿⣿⣿⠟⡡⠞⠃⣉⠤⠁⢈⣵⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡤⠐⣢⡉⠓⡻⠅⢳⣦⣌⡝⢿⣿⣷⣧⡘⠓⠮⢝⣦⠁⠌⢀⣷⡄⢁⡀⢃⠉⢻⠈⠹⠞⣗⠀⢦⠀⠀⠀⠀⣰⡀⠁⠀⠠⠀⠆⠀⠁⠀⠀⡌⠀⠀⠀⡃⠁⠀⠈⠄⡀⠁⠀⠈⠀⠂⢨⢀⡀⠁⢈⢛⠁⡜⠘⣡⢃⡔⠴⠞⠰⠋⣴⢃⡞⢁⢆⠁⢨⠃⣀⠜⠁⠀⠠⠐⠀⠔⣀⣤⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣶⢦⣍⠦⢉⡻⢶⣿⣾⣿⣾⢟⡳⡙⠿⣤⡀⣄⠙⢷⡸⣄⠈⠙⠀⢻⡄⢰⠀⠁⠀⢧⣱⢰⡈⢴⡘⣆⡏⣿⣧⡀⢆⠠⠀⠀⠀⠀⠀⠀⠂⠰⣄⠀⠱⠀⠀⠉⠀⠀⠀⠃⠀⢀⢀⣼⠠⡇⡄⡄⠀⣴⡡⢀⢧⡏⡎⡶⠋⠘⠘⡁⣾⢆⣿⠏⣀⣥⠞⡁⠄⡤⢊⣠⣠⣶⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡿⠋⠀⢀⡠⣤⣥⣝⣻⡾⢯⣍⠛⠿⣽⣮⡳⡆⠦⡇⢦⠘⢻⣼⠀⢠⢀⠀⠁⠘⡀⣦⣼⡄⢻⣇⢷⠞⠃⢻⠄⡟⣿⣳⡼⠀⠀⠀⠀⠀⠀⠀⠀⠄⠈⢡⠀⡀⠀⠀⠀⠀⠀⠂⠀⣌⡞⣿⣟⡄⡡⢼⡇⣿⢣⠉⡸⠀⠝⠁⢠⠀⣜⠰⢫⠊⠁⠜⡿⣿⣿⣿⣟⣵⡿⢿⣿⣿⢟⣽⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣤⠥⠆⠈⠛⣛⣿⣿⢿⣿⣎⠤⣄⢘⠻⢧⠁⠢⠁⡜⢠⡎⢙⣀⢫⣬⣧⠙⠳⡝⣿⣿⣷⠪⠛⢀⠂⠀⡀⡌⠳⢻⣧⣏⣷⣀⠀⠀⠐⠀⠀⠠⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣀⢠⣿⣟⣿⡏⣽⡇⡆⢱⡿⣋⣧⠃⣰⡆⢀⡼⠀⠋⣠⣿⠆⢊⡞⡚⣹⢩⣿⣺⣯⣾⣿⣋⣵⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣏⡀⣰⣾⣿⣿⣿⣿⣿⠷⠸⡿⣆⠈⠷⣷⠀⢸⣰⡇⠁⠈⠇⠁⣏⣷⡸⣎⣿⢶⣹⣎⣿⣿⣿⣇⠀⠇⢸⠇⠁⠸⣿⣸⢸⢹⣿⠀⠀⠀⠸⢀⠀⠀⠰⠀⠀⠀⠀⠀⠀⠀⠀⢹⣿⣹⣿⢾⡿⠀⡁⠁⠈⠹⢁⡏⣾⣿⣰⠏⠰⠀⠰⢸⠀⠀⢀⡎⡸⢱⢏⣹⢏⣿⣿⣱⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡥⠟⠿⠿⠿⢟⡻⠯⠵⣀⠑⣜⢧⠂⣿⣆⣈⡛⢁⣈⠐⢍⡲⠤⠬⣝⡲⢭⣷⣝⠿⣾⡿⣿⣿⣆⠈⠈⠀⢠⣿⡽⣿⡈⢿⡟⠀⠀⠀⠰⢰⡄⠀⠀⢀⢷⠀⠀⠀⡀⠀⠈⣏⡿⣿⣿⢾⣧⠐⡎⢀⠏⢠⢏⣾⣿⡯⣲⢃⢆⣡⣴⣶⠘⢠⢋⡔⢱⣿⢯⠇⢈⡜⣿⣿⡿⢛⣷⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡟⠁⠀⢀⣀⣻⣭⠳⡘⣤⠰⣤⡳⣾⡢⣓⣤⣤⣉⣦⣼⣭⣻⠶⠯⠷⣒⠤⠌⡱⢦⣌⣿⣮⣿⣻⣿⢿⡀⠀⠀⠀⣿⣷⡹⣷⡣⢻⣼⢣⢀⢰⠰⣶⡄⠀⠌⢎⣀⣴⡀⡀⠐⢠⡟⣞⣿⣿⢾⣏⠀⠑⠦⠀⣹⣾⣿⢻⣱⣳⠿⠫⠇⣠⠷⠊⣰⡼⣁⣬⠆⠰⢊⠂⣴⢟⣭⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣟⠛⠀⣴⢾⡿⢽⣽⣷⣜⢶⣝⢿⣏⠤⣻⣿⣿⣿⣿⣿⣿⣿⣿⣷⣿⣯⣖⣫⡛⡾⣝⣪⣟⠺⣷⣟⣾⣿⣷⡀⢀⢲⡼⣿⡰⣽⣧⠻⠹⠸⣏⢼⣾⢹⠃⠀⢦⡸⠁⢿⣇⢹⡇⢻⡇⣯⢻⣿⡞⡏⠀⠀⢀⣴⣿⣿⢳⡿⣟⣥⠞⣰⢗⡣⡐⣣⣙⣑⠛⣡⡵⠶⢫⠞⢃⠎⣾⠛⠉⠹⠿⠿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⣿⣿⣿⡿⠏⣀⠐⣿⠛⢷⣘⢦⣸⣷⣯⡳⣍⡳⣭⡛⣿⣿⣿⣿⣿⢿⣟⠿⠿⠷⠀⣉⠉⡛⢻⣷⠾⣽⣻⣷⣯⣛⣿⣿⣧⣷⣞⣮⣻⣿⡇⣝⡏⣆⠑⡄⠉⠁⢁⠈⣱⡀⠀⡳⡈⡄⡏⢲⡇⢸⡇⣿⢹⣽⣣⣤⣄⢃⢾⣿⣿⣟⣿⣿⣳⡶⠟⣡⠾⣋⠥⣚⣯⣿⣿⡷⢦⡴⣶⣼⡟⣡⢀⡶⢁⡀⠀⠀⠨⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⡿⢿⡉⣅⣰⢦⣁⣨⡁⠤⣛⠫⢿⣯⣟⠿⣦⣽⣳⣷⣮⣙⣫⡭⠭⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠉⠈⠂⠉⠓⠽⢻⢿⣷⣿⣿⣾⢱⣏⣿⣼⡄⠉⡔⠀⠸⡠⡟⠀⢯⠁⠀⠰⢠⠳⢐⣈⠃⣎⠟⣿⣷⢹⣇⣿⣟⣾⣿⣿⡿⣫⣷⣟⣵⢾⣯⣾⣿⣷⠿⠿⠿⣿⡿⢿⣿⣿⠿⠏⣶⣶⣾⣷⣿⣯⣔⠂⠀⣈⠙⢻⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⡿⠆⠀⠨⣉⠓⠶⠄⢏⡱⢬⡙⣳⣴⣯⢽⣚⣭⣻⡿⣿⣶⡢⠉⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠉⠛⣿⣿⣿⣎⣏⣿⣼⡿⣴⠃⠀⣇⢻⠀⠃⠘⠇⡔⣄⡌⠆⡄⡈⠀⢌⢀⢿⣿⣿⣿⣿⣿⣿⢏⣫⣾⠿⢫⠞⠁⠿⠞⠉⠈⠀⠂⠉⠑⠒⠉⣙⣛⣶⣿⣯⣥⡯⢿⣫⣿⣿⣿⣿⠀⢬⠓⣠⠘⠻⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣿⡦⣐⠈⡄⢉⢻⣴⣈⣧⠈⠛⢍⡉⠪⡙⠁⠙⠫⠑⠉⢋⡙⣓⢤⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣶⣿⡶⠀⠀⠀⠀⠀⠈⠻⣿⣿⣿⣷⣿⣏⣿⢃⣥⠎⠈⡀⠟⠁⠈⡀⡏⢣⡌⢳⠉⠄⢺⣾⡞⣿⣿⣿⣿⣿⡿⠋⠘⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣽⡿⡿⠯⠥⠴⣾⠟⣫⣾⣿⣣⡞⠈⠐⡡⢦⡄⣩⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⡕⠂⣁⠀⠘⠤⠶⢮⡉⠒⠄⢀⣀⠈⠀⠀⠒⠦⢄⣉⠻⣶⣿⣮⣿⣦⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠀⠀⠀⠀⠀⠀⠀⠀⠀⠙⣿⣿⣿⣿⣿⡽⢯⢆⣠⢶⡁⠣⠰⡲⡀⢈⣌⢦⠈⣿⢰⣿⣿⣿⣿⣿⣿⡿⠇⠀⠀⠀⠀⠀⠀⠀⣀⣤⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣰⣾⣿⣻⣟⣓⣲⣿⣿⡷⠗⣈⣠⠎⣐⠿⣃⠾⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⡿⠟⠋⣀⠈⢁⡀⠀⠈⠀⠤⣈⠀⠀⠐⠈⠂⠁⠂⢄⣩⣭⡳⢿⣿⣿⣿⣇⠀⠀⠈⢦⡁⠂⠀⠠⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢹⣿⣿⣿⣿⢣⡤⡿⡵⡛⢠⠀⣸⢇⠃⠚⠛⠸⡷⢉⢿⣿⣿⣿⣿⣿⠋⠀⠀⠀⠀⠀⠀⠀⠀⠈⠻⠛⠃⠀⠀⠀⠀⠀⠀⠀⠀⡀⠀⣯⣼⡿⠗⠿⠿⠿⢗⣺⣽⠿⠋⣡⠖⠁⠚⣡⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⡿⠇⠀⠘⠀⠀⢈⠒⠄⡀⢀⡁⢃⣀⠀⠀⠀⠀⠀⠠⢤⣊⡉⢟⣿⣿⣿⣿⣆⠀⠈⢆⡁⠄⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠘⣿⣿⣽⣿⡿⣣⣹⡵⠟⠟⠰⠂⡈⠆⠃⣠⡷⣷⣻⣼⣿⣿⣿⣿⡇⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠠⢀⠂⢌⠀⠀⠰⣯⣿⣷⣾⣉⣖⣋⣌⠹⠵⠯⠔⢀⡴⢒⡹⢿⣯⢿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣤⡀⠀⠂⠄⠀⠈⠨⢉⠂⠔⢢⠤⡀⠀⠄⠈⠑⠠⠐⠤⠬⡑⠲⣾⢷⣿⣿⣆⠀⠈⠉⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠨⣹⣿⣟⣿⣟⢏⡶⣻⡯⢣⣮⢹⣀⢩⢰⣽⠨⣳⣽⣿⣿⣿⠟⠀⠀⠀⠀⠀⠀⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⡰⠌⡿⣌⠀⠀⢘⣿⣿⣿⣿⣿⠿⠿⢗⠠⣴⠂⢀⠉⢰⠟⡐⣯⢀⣾⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⡒⠀⠀⠀⠀⠡⠀⠄⢂⠒⣄⢒⡠⢄⠐⠀⠀⠁⠀⠹⠤⠢⣀⠪⡅⠙⢿⣿⣧⣀⠀⠈⠂⢀⠀⠀⠀⠀⠀⠀⠀⠀⢀⠠⡒⠄⠀⠀⠀⠀⢸⣿⣿⣿⣟⣯⡿⢊⡿⣁⣘⠣⢎⡑⣨⣰⡝⣿⣽⣿⣿⣿⣯⡀⠀⠀⠀⠀⢤⡐⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⠏⡵⠃⠀⢀⣼⣿⣿⣿⣿⣿⣿⠿⢟⡋⠀⠀⢀⠲⣛⠛⠙⠋⣠⡾⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⠧⠁⠀⠀⡀⢂⠌⠢⣉⠟⣌⢧⡘⢾⣜⡒⠌⠐⠂⠠⢭⣐⡂⠃⠘⠕⣢⣀⠻⣟⣻⣶⣬⣅⣈⡑⠒⠐⠂⠐⠒⣒⣀⣡⣤⡄⠀⠀⠀⠀⣻⡿⣿⡿⡯⢙⣾⡛⡫⣾⠆⠠⠶⢽⣠⡛⣽⣙⢿⣿⣿⣻⡷⠀⠀⠀⠀⠈⢳⡘⠆⣀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣡⠎⠁⠀⣠⣾⣿⣿⢿⣛⣿⡽⢃⣌⠡⠐⠀⠀⠄⣠⠴⡎⡴⠌⠂⢡⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⢦⠀⠀⠀⢰⣁⣎⡱⠷⣾⣦⣄⣈⣡⢤⢠⡐⠀⠂⠡⠤⡠⢌⣉⠉⠀⠐⠉⣡⣩⡅⢎⠹⢿⡿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣀⢀⡴⢂⠐⠉⠼⠏⣿⡧⡏⡹⣵⣵⠋⠔⢐⣡⢐⠬⠋⢳⢓⡘⢾⣷⡿⡧⠀⠀⠒⠠⣀⣄⣀⡐⠠⠀⠀⠀⢀⣀⠀⠠⠔⠊⠄⣒⣠⣾⣿⡿⠟⠉⣉⣉⠒⢐⠆⠀⠀⠤⢐⢀⣊⠉⣉⣀⣄⠠⠘⢂⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣦⠀⠤⠐⠮⢿⣿⣿⣷⣶⣭⣭⣋⡁⡊⢥⣈⡵⠂⠀⠠⠀⠢⠒⠦⠈⠒⢂⣀⠀⡟⢣⠚⠍⢭⣭⣟⣻⣿⣿⣿⣿⣿⣛⡿⠟⠀⠀⠀⠀⠀⠀⢟⡟⢐⠀⠘⠷⠉⡰⠃⠋⣐⠁⢺⣛⡚⠲⡍⠸⡝⠡⠡⠀⠀⠀⢀⠸⢿⣿⣿⣿⣷⣶⣤⣤⣀⠤⣴⣶⣶⡾⠿⠿⠛⠉⠀⡤⠟⣉⠀⠈⣄⠂⣡⠖⣉⣥⠎⣡⣓⣡⠎⠄⠐⠈⢘⣻⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⡁⢉⠀⠀⠛⣿⣶⣿⣿⣿⣿⣿⣯⡵⣉⡒⠖⠋⠡⢖⡐⠀⠁⠲⢷⣄⣮⣉⠢⢀⠈⠈⠉⣰⠮⠍⣶⠭⢿⢯⣽⣿⣿⣿⡷⠌⠁⠀⠀⠀⠀⠃⡚⠅⢀⣀⢯⠞⢓⠃⠴⣃⠨⠖⠁⠩⣽⡛⣖⡈⠑⡄⠀⠀⠀⠀⠀⣻⣛⠿⣿⡿⣿⣷⣿⣯⣟⣿⡛⢋⠑⢂⡠⠂⣀⠔⠃⣀⠀⠀⠒⣡⣬⣥⣯⣿⣿⣻⣿⣿⣿⣇⣐⣊⡡⢀⣺⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣿⣤⠁⠎⡉⠓⢤⢻⣏⣿⣿⣿⡿⣿⣭⣑⡘⢂⠋⠳⠤⠌⠛⢤⣀⠂⡀⠁⠄⣒⠢⢙⣊⣀⠴⢍⣓⠐⢫⡿⣿⣭⣿⣯⣵⡰⠖⠀⠀⠀⢀⠴⠀⠀⠀⠴⠉⠉⠄⠁⠐⡸⢧⡠⠄⠆⠠⢁⣆⢢⡗⡘⠈⠄⠀⠀⠀⢘⡻⣿⣿⣷⣯⡳⣭⣃⡈⠄⣀⠀⢀⡄⠁⠀⠊⠥⢐⠥⢠⣤⣤⢘⡻⢋⡩⠵⢋⣥⠿⢿⣿⢿⣫⣷⣿⣯⠥⠤⠿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⡛⠂⠖⠤⡑⠈⢭⣍⣻⣯⣭⣿⣽⣮⣳⣏⣍⣂⠓⡲⢎⣓⣀⣈⡒⡉⠷⠖⡒⢀⠦⢶⠀⠉⢑⡲⠬⣉⢤⡤⣦⠬⠡⣄⠀⠁⡀⢀⠤⣗⠈⠀⠀⠀⠀⠘⠄⠀⢂⡈⠇⢠⠆⣁⠣⠢⡀⠀⠀⡐⠂⠀⠀⠙⡄⣐⠀⠀⠉⢙⠺⢷⣶⣛⣛⠓⣀⠀⠒⠅⢂⠄⠁⣠⡴⠚⢅⣚⣿⡫⣭⢯⣐⣻⣾⡟⠭⢶⣾⣿⣾⣿⡿⠿⢋⡡⢴⠶⠛⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣟⡛⠂⠐⠠⠐⠂⠄⣂⠤⡄⠚⢭⣿⣯⣍⣒⠎⠔⠣⠔⡳⢺⠼⠃⢉⡤⡴⠤⢌⡁⢂⡀⠺⠓⣲⣬⢭⣤⣦⡻⢺⣟⠳⢀⣢⣔⢾⣛⣴⠍⠐⠄⠀⠀⠀⠀⠀⠠⢀⡠⠸⠡⡀⠀⠠⠁⠀⠀⠈⠀⠄⠀⠄⠀⠀⠻⡶⠤⠄⡄⠛⠮⣥⡦⠤⠖⢈⠹⠃⠔⡠⣊⠐⢒⡨⡵⠻⠛⢷⣭⣷⣯⣭⣭⡷⠾⣲⣿⣿⣿⡿⢛⣵⠟⣁⣤⣠⠄⢺⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⣿⣿⣗⡢⠌⣀⡉⠔⠀⡠⢄⠩⠉⢐⣈⣩⢍⡒⠊⠭⢁⣋⠠⠀⡲⢺⠭⢶⣲⣙⠋⠀⠤⣠⣽⣧⣭⢿⡶⠮⠯⣤⡠⠝⣿⣷⣶⣉⡞⣒⣩⠿⣁⠬⡄⠀⠀⠀⠀⠘⠀⠀⠄⠀⠀⠀⠀⠠⠐⠂⠀⠀⠀⠀⠀⠔⢈⣧⣮⣗⣂⠘⠻⠁⣤⣒⢶⣛⣓⣀⣀⠉⢁⡨⠅⣲⢛⣥⠾⣻⠯⣛⣯⣭⢿⠿⢿⣿⡭⠟⣫⣵⣾⣿⠃⢉⣁⣾⣵⠾⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⡿⣟⠋⠁⠀⡀⠌⣉⢁⡐⢂⡤⢭⣝⣋⣉⡭⡙⠝⠎⢲⠠⢆⡡⣐⠢⠬⠌⣓⠓⠲⢶⣒⠈⣽⣻⠟⣋⣶⣛⣲⣨⡝⠓⢀⡈⣛⠫⣭⣙⡛⣿⣆⡀⠠⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠘⣰⡿⠛⣿⣿⣿⣯⡗⢭⡐⠚⣩⣝⣻⣾⢁⡀⠄⠠⠶⠖⠛⣡⠽⠶⣛⣽⣿⣿⣿⡿⣭⣅⣚⣛⣿⣿⠟⣑⣾⣿⣿⣭⣤⣯⢽⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⠻⡌⠂⠀⠐⠈⠘⡀⢣⡔⣫⣽⢿⣻⠶⠲⠀⠩⠉⣍⢂⡑⣈⠐⡡⢄⠀⠜⠔⠊⣭⠓⣢⠾⣟⣛⡿⠶⠬⠛⠞⠛⣿⣿⣖⠿⠿⢿⡿⣫⣿⣿⣯⣀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢐⣻⣿⣯⣿⣯⣔⣗⢚⠉⢐⣡⢶⣈⠶⣮⣍⢛⡩⠏⠵⣋⢉⡩⢽⣿⠿⢷⣬⣫⢥⣶⣿⣿⣿⣿⡿⢟⣻⣵⣿⣿⡿⣯⣴⠦⣻⣿⣿⣿⣿⣿⣿⣿
        //⣿⣿⡡⠐⡀⠠⠀⠩⢎⡝⡃⠉⠐⣀⠢⠁⠀⣀⡤⣭⡥⡔⢢⠒⡄⠫⠔⣊⠜⡰⢨⣉⠴⢞⣷⣯⣤⡄⠈⣉⣀⣀⠮⢤⢀⡭⣌⠀⠈⣲⣿⣿⣿⣿⣿⣿⡆⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣬⣿⣿⣯⣭⣽⣾⠦⣀⡈⢥⠥⠨⠞⣭⣩⣷⣭⠥⢒⣤⣲⣮⡿⠾⢙⣂⣷⣾⡿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⢯⡵⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⣯⣐⠃⠄⠁⠐⠠⢸⣜⣉⡩⢐⣐⣒⣫⣙⡳⣾⡷⢓⣭⣒⡭⢎⡱⡘⠤⢋⠴⠡⠒⢨⣽⣾⣿⣷⣶⣾⣿⣿⣿⣿⣿⣿⣿⣿⡶⣾⣿⣿⣿⣿⣿⣿⣿⣿⣤⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣼⣿⣿⣿⣷⢮⣧⠚⠷⣶⠭⢿⣒⡾⠶⠾⠿⣿⣟⣳⡲⠭⢖⠓⢶⡩⠭⣵⣲⠶⣻⢳⢫⢷⣻⡶⣶⣾⣿⣿⣿⢿⣿⣿⠿⠿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⠛⡄⠡⠈⠄⡁⠀⡀⠀⠒⡀⢀⠀⠒⠒⣡⣶⡏⠵⣛⣚⣉⡭⠭⠤⠶⢒⡒⡒⣓⡛⣽⣯⣯⡿⢿⣿⣿⣿⣿⡿⢛⣭⣿⣿⣛⢻⣿⣿⣿⢿⣿⣿⣿⣿⣿⣿⣿⣶⣀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣰⣿⣿⣿⣿⣿⣿⣿⣷⣤⣭⣥⣤⣤⣶⣄⣈⣉⣭⣭⣽⣿⣿⡭⡤⠴⣀⣉⣛⠒⠧⢻⡔⣦⢋⠲⢭⡷⣷⣾⣿⣿⣿⣿⣿⡵⡾⡽⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⡳⠌⠄⢃⠒⢀⠀⠠⠙⠦⠙⣀⣀⡤⠭⠍⣒⡺⠭⠝⣒⣲⡾⡽⢉⠃⠉⠀⢁⠂⠭⠔⠚⣉⣮⡷⢴⣚⣻⣭⣾⣷⢿⣾⣾⣿⢟⣿⣫⣷⣿⣽⣿⣷⣟⣿⡿⣻⣽⢟⣥⠀⠀⠀⠀⠀⠀⠀⠀⢀⣴⣿⡿⣶⣿⣿⣟⢿⣿⣶⣾⣾⢻⣿⡿⢿⣿⣿⣿⣿⣿⣿⣛⠿⣿⣿⣿⣿⣟⣖⣲⢢⢭⣭⣒⣒⣨⠭⢟⣛⠿⣿⣿⣿⣿⣿⣋⣳⣽⡶⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣿⡛⠌⡐⠀⠐⣀⡈⠥⠒⣚⡩⢅⣂⠰⠐⠉⠁⣠⠶⢮⣟⣻⡵⠊⠁⢀⡂⠡⠄⢒⣂⡬⣵⠻⢴⣾⡽⣿⣿⣿⢿⡿⣟⣿⣟⣿⢿⣻⣽⣿⣽⡿⠽⢿⢿⣱⠟⢋⡸⣩⣞⡿⡀⠀⠀⠀⠀⠀⣄⢶⡳⢷⣝⢿⣳⣝⣿⣿⣿⣿⣿⣯⣻⣭⣻⣛⡻⢯⣛⣻⠟⣿⠯⣭⣝⣒⡲⠬⠬⣁⣉⡊⠵⢪⣕⣏⣿⣿⣿⣶⣭⣵⣲⠮⢟⣋⡭⣼⣗⣾⣽⣿⣿⣿⣿⣿⣿⣿⣿
        //⠧⢉⡀⠐⠀⠉⠀⠀⠀⠁⠤⠝⠛⣈⡁⠌⡨⢱⢎⣻⡿⢯⡥⢒⠪⣉⡤⠴⠒⣊⡭⠖⠋⣀⠴⣛⠍⣼⠟⡹⠕⢏⣾⠿⢟⣿⠷⠿⢟⠿⣛⣴⡾⢋⠕⣢⣢⢴⢟⠽⢃⠌⣵⠃⠀⠀⠀⠀⢠⠛⠳⢏⠲⣍⣕⢗⢟⡾⣿⣿⡻⣿⣿⣻⢿⣷⣾⣶⣶⣾⣮⣽⣿⣿⣌⡛⢭⡹⢍⡳⠴⣮⢭⣝⣒⡒⠬⣍⡛⠿⣿⣿⣿⣿⣿⣷⣮⣍⣂⠽⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⠉⠀⠀⠀⠀⠀⠆⠁⡈⠰⠀⠎⠷⠾⢈⡰⣀⠏⣎⠱⡈⠇⣰⠾⣉⣷⠎⣱⠏⠱⢀⡰⢏⡰⠏⣁⡾⢱⣾⣷⢾⡏⣸⡱⢉⣰⢷⠏⣰⣹⣿⡿⢉⠶⠿⠹⣱⢏⠰⣰⢆⢾⠁⠀⠇⠀⠀⢀⠀⠆⢶⣏⢿⣆⢹⡎⡆⢈⢎⢿⢿⡏⢿⣿⣿⣏⣹⣿⣿⣏⣿⣿⣹⢿⣇⡉⢶⠿⣎⢷⡿⡶⢿⣾⣏⣿⣿⣿⣿⣹⢶⣉⡹⢿⣿⣆⣀⡶⣆⢷⣆⡹⣿⣿⣿⣿⣿⣿⣿⣿      
        //        ⠠⢀⠐⠈⣀⢎⡹⣐⠆⢚⣠⠵⠊⣡⣶⠿⡫⢔⡩⢒⠭⢒⣡⠶⡡⣫⠞⡡⢞⣱⠟⡵⠋⡴⠫⣾⠟⡵⣫⡾⠛⠙⢥⠐⠉⠂⠊⠘⠙⠉⠈⠘⠁⠀⠀⠀⠀⠀⠀⠀⠐⠀⠷⠘⣔⢍⡳⡊⢿⣎⠫⡻⡄⡻⢷⡙⠿⣷⣭⣝⡛⢦⣉⢽⡛⢥⡲⣍⡂⠝⡛⠾⢷⣿⡿⣿⣿⣿⣿⣿⣿⣿⣵⣦⣭⣗⡪⢝⡛⠾⣽⢿⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⠀⠄⡀⠄⡀⢃⠐⠈⠁⠂⠀⢀⠠⡄⢖⠥⣚⣭⣖⠤⠚⣫⠕⣭⡶⠟⣫⣵⣺⣭⡳⢏⡴⡡⢪⡵⡫⣡⠞⣠⠞⣡⢞⢄⡜⡼⢁⡴⢫⣃⠆⠀⠀⠀⠀⠀⠀⠂⠀⠀⠀⠀⠀⠀⢀⢀⠀⢀⠀⠀⠈⢀⠀⠁⠀⠀⠈⠻⡔⡷⡹⡹⣄⡹⣆⢹⣻⢿⣻⢦⡻⢷⣭⡲⣝⠿⣿⣷⣾⣷⣶⣦⣹⢿⠿⣦⣽⢿⣿⣿⣟⣿⣿⡿⡝⣳⢮⣕⡢⣍⡻⢿⣿⣿⣿⣿⣿⣿⣿⣿
        //⠉⠒⣁⠂⡀⠀⠀⠀⠁⠀⠁⢀⡤⠚⠱⣾⣿⡿⢁⡤⣚⣵⣾⠏⣒⣉⡼⢒⢟⣳⢏⡾⢊⡼⣣⠟⡵⢃⡼⢁⡼⣱⢋⢌⡜⣰⢯⣜⣧⢳⣞⡭⢃⢤⡆⢊⡅⣀⠀⢎⡄⣀⡄⠀⠉⠈⠣⠀⣠⠄⡀⣒⠀⣈⠀⠀⢂⡤⣜⡇⡇⢱⢸⢹⣮⢫⡳⡹⣮⡣⡌⣷⣼⣿⣶⣽⣮⡻⣿⣿⣿⣿⣿⣿⣿⠿⠿⢷⠶⣭⣿⡻⢾⠳⠧⣝⣫⠻⣝⡳⢬⣓⡮⣟⢿⣿⣿⣿⣿⣿
        //⢌⡱⢀⠞⠰⢂⠐⠀⠀⡠⠖⠁⠀⠀⡀⢤⠞⠵⣩⣾⠿⣉⣠⣶⠿⠁⠔⣡⠾⢉⡞⣠⢧⡳⢍⡞⢁⠞⣠⢏⣾⢃⠎⡜⣼⢯⣳⢞⣞⣳⣟⣋⣴⠇⢊⡭⣚⣥⣶⢟⠀⠜⡄⠀⢈⡀⢐⠦⢦⡙⣒⢬⡣⣄⢑⠶⣮⣙⢳⡇⠇⡿⠸⣸⣿⡷⣝⠵⣌⢷⣝⢮⡻⣿⣿⣿⣿⣿⣻⣿⣿⣿⣿⣯⣻⣿⣿⣝⣻⡮⠍⡉⠛⠳⣒⠤⢣⡻⣭⣟⣿⣿⣿⣾⣯⣻⣿⣿⣿⣿
        //⠠⡙⢎⠰⡡⢦⠘⡤⠊⠀⠀⠀⠀⠂⢀⠔⡡⠞⠟⢋⣼⣿⡟⢃⠤⠚⣉⠖⣱⡟⣜⣡⢧⢋⢎⢠⢏⣠⢧⣿⢃⡎⡜⣼⣯⢯⡳⢯⣞⣷⣻⡟⣧⣪⣯⣾⣿⠋⡑⢡⠔⣽⢰⡏⢤⣣⡁⢮⣞⢮⣪⢳⠝⣮⣫⠳⣵⡘⢿⢡⢰⠇⣇⣿⡷⣿⣯⣧⡘⣮⢻⣷⣽⣜⢿⣿⣿⣿⣿⣯⣿⣟⢿⣿⣿⣿⠿⣿⣮⣟⠛⠰⠤⣤⣌⣳⣷⣽⣶⣽⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣷⣹⣌⡱⢀⠔⠉⣀⢂⠡⠈⠐⠀⠡⠂⠀⢀⡠⠾⢋⠝⡟⠁⠀⣔⣲⠄⡸⢫⣼⣾⡵⢣⢏⢢⢏⣆⢇⣾⣷⡟⡸⣸⣿⣞⣷⢻⡗⣾⡜⣧⣛⡿⢿⡿⣽⣧⠇⡴⢣⣼⣆⣺⠰⠹⡓⣟⢮⣫⠚⢧⢳⣹⢼⣿⣷⣾⣿⠎⢆⣿⢸⢹⣯⣿⢿⣻⣿⣿⣮⢧⣻⣿⣾⡿⣻⣿⣟⣿⣿⣿⣿⡿⣝⠿⣿⣿⣦⣄⠉⠻⢦⣴⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⡽⡻⠟⡕⠁⠂⡘⠠⢌⠠⠁⡠⠊⠀⠀⠈⢀⠄⠀⠁⠀⠀⣠⠾⣛⡭⠚⢰⣿⣿⡿⣣⢫⣎⣏⡾⣸⢸⣿⢿⢡⣱⣿⣿⣟⣾⣻⢞⣵⣻⢃⠝⡙⣰⠏⡱⠃⡜⣴⣿⣿⠏⡇⡥⡇⢷⣿⣾⣷⣷⣽⣌⢮⠺⣟⡿⡽⠏⡜⣸⠇⡇⣿⣻⡿⣿⣿⣿⣿⡿⣧⣳⢻⣿⣿⣝⢮⢷⡻⢷⡽⢿⡿⣮⠲⡀⣀⠩⣘⠿⣶⣴⣿⣦⣽⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣡⡵⢊⡔⣩⠒⡤⢁⠂⢁⠖⠀⠠⢁⠐⠀⠀⠀⢀⠀⠀⢀⡴⠊⠁⠀⡌⣾⣿⡿⡳⠣⡟⡞⣼⡧⡇⣿⣟⢃⠇⣾⣿⣟⡾⣷⢯⣛⡶⣏⣡⠾⢡⢃⡼⣀⠟⣰⠯⣽⠋⠌⣼⣡⢣⣿⢻⣏⣿⣻⣝⣻⡌⢷⡉⢞⡽⡙⡰⠭⢰⢸⡳⣏⡿⣽⣯⣿⣿⣷⣬⡉⢧⢻⠟⠿⣎⢧⠉⠓⠈⡐⠂⠈⠁⠙⢦⠱⣭⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⠟⣴⢮⣔⣂⠊⠅⢋⠔⡡⢠⠉⡀⠄⠂⠈⠀⡐⠀⠀⡠⠋⠀⠀⠠⠂⡇⠔⠉⡰⢡⡞⡹⢸⠿⠱⠐⡷⡛⠸⢸⢟⠷⠮⢝⡏⠓⡭⢳⢯⠻⢡⠃⢮⠱⠎⡵⢫⠷⢉⠌⢘⠱⢂⡞⣜⠢⡝⢲⡹⠌⠳⠘⡄⢊⠈⡲⠴⣉⠃⠌⠲⡥⠉⠙⢧⠛⢾⡹⢯⢷⣛⡌⣆⠩⢀⠐⠂⢳⡀⢅⠀⠀⢀⣀⠈⢤⡳⡽⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣾⣿⣿⣿⣭⢟⠔⣁⠌⠡⡁⢆⠡⠘⡀⠃⠂⠄⠄⠐⠁⠄⡐⢀⠂⠰⠁⠀⡴⠁⠂⢠⠁⠈⠀⠘⠀⡴⠃⠃⠌⠀⠋⠀⠀⠀⠀⠀⠀⠀⠁⠀⠀⠀⠀⠈⠀⠋⠀⠈⠀⠀⠀⠀⠊⠀⠀⠀⠀⠠⠁⠀⠀⠀⠀⠁⠀⠀⠀⠀⠀⠀⠀⠁⠀⠀⠂⠀⠀⠀⠀⠉⠐⡘⡄⠀⢂⠐⢄⠳⡈⣟⣶⣾⣽⣿⣿⣿⣞⢮⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⠿⣏⠥⣶⡖⡡⣚⣭⠻⣴⡩⢄⠊⠥⣀⠣⠜⡠⠊⣀⢂⠐⡀⠂⠈⢸⠀⡴⠁⠀⠀⠆⠀⠀⠀⢀⠀⠀⠐⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⠀⠀⠀⠀⠀⠀⠀⢀⠀⠀⠀⢀⠀⠌⡐⠠⢱⠈⠀⢦⣘⣣⣱⡹⣿⣿⣿⣿⣿⣿⣿⣿⡳⡿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⡿⢿⠾⠿⠼⢍⣉⣉⣳⠷⣙⢆⠩⠐⠀⠅⢊⠰⡐⠤⢊⡰⠀⠆⠐⠈⡰⠁⢀⠂⠘⠀⢀⠀⢂⠈⠀⠀⠂⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⡀⠀⠐⢀⠂⡈⠄⡌⡐⠀⢂⡜⡰⢌⡳⠌⡇⣞⠺⣍⣛⣷⢻⢲⠿⣿⣿⣿⣿⣿⣿⣿⡿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⣶⣶⣿⣶⣶⣶⣞⠻⠋⠓⡽⣎⣛⠲⠥⠔⣢⢃⡌⠓⣌⠲⡑⢊⡅⡸⠁⠐⠀⢀⠂⢌⠠⡈⠀⠀⠀⠀⠃⠀⠀⠀⠀⠀⠀⢀⠂⠁⠀⠀⠀⠀⠀⠄⠠⠀⡀⠀⠁⠀⠀⠀⠀⢂⠀⠀⠀⠀⠀⠀⠀⠀⠀⢈⠐⠠⠐⡀⢀⠁⣀⠢⡀⠒⡈⢦⡈⠱⡃⢄⠘⢽⣾⣷⡾⢹⣬⣿⣾⣯⣉⣧⣯⢿⣮⣽⣿⣿⣿⣿⣿⣿⡾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //⡾⣽⣿⣿⣿⣷⣭⠎⠵⠠⠤⣀⠀⠀⢄⢢⡴⣌⡜⣩⠔⡣⠝⠢⡰⢱⠀⠃⠐⠘⠀⠄⠂⠁⠐⡀⠇⢨⠀⠀⠀⠀⠀⠀⠀⠀⠈⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⠀⠀⠀⠀⠀⡀⠀⠂⠀⠀⡁⠀⠀⢀⠠⠐⢀⠊⢡⠃⡔⠣⢎⠴⣡⣛⢶⣱⣦⠽⣧⡝⣶⣵⣦⣿⣿⣿⣌⣟⣿⣿⣿⣿⣷⡘⣾⣿⣿⣯⣝⣿⣿⣿⣿⣿⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
        //--------------------------------------------------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------------------------------------------------
        // Extra buttons for testing/debugging API responses — can be removed in production
   

        //---------------------------------------------------------------------------------------------------------------------------
        // Extra buttons for testing/debugging API responses — can be removed in production


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

    async Task LoadEmployeesAsync()
    {
        btnLoad.Enabled = false;
        lblStatus.ForeColor = Color.Silver;
        lblStatus.Text = "Loading employees...";
        dgv.DataSource = null;

        try
        {
            var employees = new List<Employee>();
            string url = "personnel/api/employees/?page=1&page_size=100";

            while (!string.IsNullOrWhiteSpace(url))
            {
                var json = await ApiClient.GetAsync(url);
                var paged = JsonConvert.DeserializeObject<PagedResponse<Employee>>(json);
                if (paged?.Data != null) employees.AddRange(paged.Data);
                url = paged?.Next?.Replace(ApiClient.Http.BaseAddress.ToString(), "");
            }

            // Email replaced with Rate (blank — to be filled manually)
            dgv.DataSource = employees.ConvertAll(x => new
            {
                Employee_Code = x.EmpCode,
                First_Name = x.FirstName,
                Last_Name = x.LastName,
                Department = x.DepartmentName,
                Hire_Date = x.HireDate,
                Rate = ""   // payroll rate — fill manually
            });

            lblStatus.ForeColor = Color.FromArgb(0, 200, 150);
            lblStatus.Text = $"{employees.Count} employees loaded.";
        }
        catch (Exception ex)
        {
            lblStatus.ForeColor = Color.OrangeRed;
            lblStatus.Text = ex.Message;
        }
        finally { btnLoad.Enabled = true; }
    }
}

// ── Punch record from iclock/api/transactions/ ────────────────
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

// ── One row per employee per day ──────────────────────────────
class DailyRecord
{
    public string EmpCode { get; set; }
    public string FullName { get; set; }
    public string Department { get; set; }
    public string Date { get; set; }
    public string ClockIn { get; set; }
    public string ClockOut { get; set; }
    public string BreakOut { get; set; }
    public string BreakIn { get; set; }
}

class DepartmentItem
{
    [JsonProperty("id")] public int Id { get; set; }
    [JsonProperty("dept_code")] public string DeptCode { get; set; }
    [JsonProperty("dept_name")] public string DeptName { get; set; }
}

// ── Export Dialog ─────────────────────────────────────────────
class ExportDialog : Form
{
    DateTimePicker dtpFrom, dtpTo;
    CheckedListBox clbDepts;
    Button btnLoadDepts, btnExport, btnSelectAll, btnClearAll;
    Label lblStatus;
    List<DepartmentItem> _departments = new List<DepartmentItem>();

    public ExportDialog()
    {
        Text = "Export Timecard to Excel";
        Size = new Size(480, 560);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        BackColor = Color.FromArgb(18, 18, 18);
        ForeColor = Color.White;
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
            Size = new Size(432, 300),
            BackColor = Color.FromArgb(35, 35, 35),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            CheckOnClick = true,
            Font = new Font("Segoe UI", 9)
        };
        Controls.Add(clbDepts);

        lblStatus = new Label
        {
            Location = new Point(16, 440),
            Size = new Size(300, 50),
            ForeColor = Color.Silver,
            Font = new Font("Segoe UI", 8)
        };
        Controls.Add(lblStatus);

        btnExport = MakeButton("Export to Excel", 322, 448, 130);
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
                clbDepts.Items.Add(d.DeptName, true);

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
        var selectedDeptNames = clbDepts.CheckedItems.Cast<string>().ToList();
        if (selectedDeptNames.Count == 0)
        {
            lblStatus.ForeColor = Color.OrangeRed;
            lblStatus.Text = "Please select at least one department.";
            return;
        }

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

            lblStatus.Text = $"Processing {punches.Count} punch records...";
            Application.DoEvents();

            var filtered = punches
                .Where(p => selectedDeptNames.Contains(p.Department ?? ""))
                .ToList();

            var daily = GroupToDailyRecords(filtered);

            lblStatus.Text = "Building Excel file...";
            Application.DoEvents();

            BuildExcel(daily, sfd.FileName);

            lblStatus.ForeColor = Color.FromArgb(0, 200, 150);
            lblStatus.Text = $"Done! {daily.Count} rows across {selectedDeptNames.Count} department(s).";

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

    List<DailyRecord> GroupToDailyRecords(List<PunchRecord> punches)
    {
        var result = new List<DailyRecord>();

        var groups = punches.GroupBy(p => new
        {
            p.EmpCode,
            Date = p.PunchTime?.Length >= 10 ? p.PunchTime.Substring(0, 10) : p.PunchTime
        });

        foreach (var g in groups)
        {
            var sorted = g
                .Select(p => p.PunchTime)
                .Where(t => !string.IsNullOrEmpty(t))
                .OrderBy(t => t)
                .ToList();

            var first = g.First();
            result.Add(new DailyRecord
            {
                EmpCode = g.Key.EmpCode,
                FullName = $"{first.LastName}, {first.FirstName}".Trim(',', ' '),
                Department = first.Department ?? "",
                Date = g.Key.Date ?? "",
                ClockIn = sorted.Count >= 1 ? TimeOnly(sorted[0]) : "",
                ClockOut = sorted.Count >= 2 ? TimeOnly(sorted[^1]) : "",
                BreakOut = sorted.Count >= 4 ? TimeOnly(sorted[1]) : "",
                BreakIn = sorted.Count >= 4 ? TimeOnly(sorted[^2]) : ""
            });
        }

        return result
            .OrderBy(r => r.Department)
            .ThenBy(r => r.EmpCode)
            .ThenBy(r => r.Date)
            .ToList();
    }

    string TimeOnly(string dt)
    {
        if (string.IsNullOrEmpty(dt) || dt.Length < 16) return dt ?? "";
        return dt.Substring(11, 5);
    }

    void BuildExcel(List<DailyRecord> records, string filePath)
    {
        using var wb = new XLWorkbook();

        var byDept = records
            .GroupBy(r => r.Department)
            .OrderBy(g => g.Key);

        // Column layout:
        // A  Employee ID
        // B  Name
        // C  Department
        // D  Date
        // E  Clock In
        // F  Clock Out
        // G  Break Out
        // H  Break In
        // --- from API (blank, ZKBio fills these) ---
        // I  Total OT
        // J  Unscheduled
        // K  Regular(H)
        // L  OT1(H)
        // --- payroll computed (blank, fill manually) ---
        // M  DCTC REG HRS(H)
        // N  DCTC OT SYS ADJ(H)
        // O  DCTC TOTAL OT(H)
        // P  DCTC MANUAL ADJ REG(H)
        // Q  DCTC MANUAL ADJ OT(H)
        // R  TOTAL REG HRS(H)
        // S  TOTAL OT HRS(H)

        string[] headers =
        {
            "Employee ID", "Name", "Department", "Date",
            "Clock In", "Clock Out", "Break Out", "Break In",
            "Total OT", "Unscheduled", "Regular(H)", "OT1(H)",
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

            // Auto-fit and minimum widths
            ws.Columns().AdjustToContents();
            ws.Column(1).Width = Math.Max(ws.Column(1).Width, 13);
            ws.Column(2).Width = Math.Max(ws.Column(2).Width, 24);
            ws.Column(3).Width = Math.Max(ws.Column(3).Width, 24);
            // Payroll columns — fixed width so they're easy to fill in
            for (int c = 9; c <= totalCols; c++)
                ws.Column(c).Width = 18;

            ws.SheetView.FreezeRows(6);
            ws.SheetView.FreezeColumns(2); // freeze Name so it stays visible while scrolling
        }

        // Summary sheet when more than one department
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

            ws.Columns().AdjustToContents();
            for (int c = 9; c <= totalCols; c++)
                ws.Column(c).Width = 18;
            ws.SheetView.FreezeRows(6);
            ws.SheetView.FreezeColumns(2);
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
        // Merge title cells across all columns
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

            if (c < 8)
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(47, 117, 181);
            else if (c < 12)
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 128);
            else
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(180, 95, 6);
        }

        ws.Row(6).Height = 36; // taller header row for wrapped text
    }

    void WriteDataRow(IXLWorksheet ws, int row, DailyRecord rec, int totalCols)
    {
        ws.Cell(row, 1).Value = rec.EmpCode ?? "";
        ws.Cell(row, 2).Value = rec.FullName ?? "";
        ws.Cell(row, 3).Value = rec.Department ?? "";
        ws.Cell(row, 4).Value = rec.Date ?? "";
        ws.Cell(row, 5).Value = rec.ClockIn ?? "";
        ws.Cell(row, 6).Value = rec.ClockOut ?? "";
        ws.Cell(row, 7).Value = rec.BreakOut ?? "";
        ws.Cell(row, 8).Value = rec.BreakIn ?? "";

       
        for (int c = 9; c <= totalCols; c++)
            ws.Cell(row, c).Value = "";

        
        var rowRange = ws.Range(row, 1, row, totalCols);
        rowRange.Style.Fill.BackgroundColor = XLColor.White;
        rowRange.Style.Font.FontColor = XLColor.Black;
        rowRange.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
        rowRange.Style.Border.BottomBorderColor = XLColor.LightGray;

       
        for (int c = 4; c <= totalCols; c++)
            ws.Cell(row, c).Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

  
        var payrollRange = ws.Range(row, 13, row, totalCols);
        payrollRange.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 220);
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