using ChoETL;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Cronos_Data
{
    public partial class Main_Form : Form
    {
        // Drag Header to Move
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        private bool isClose;
        Timer timer = new Timer();

        // FY ---
        List<FY_BetRecord> _fy_bet_records = new List<FY_BetRecord>();
        List<String> fy_datetime = new List<String>();
        List<String> fy_gettotal = new List<String>();
        List<String> fy_gettotal_test = new List<String>();
        private double _total_records_fy;
        private double _display_length_fy = 5000;
        private double _limit_fy = 250000;
        private int _total_page_fy;
        private int _result_count_json_fy;
        private JObject jo_fy;
        int _fy_displayinexel_i = 0;
        private int _fy_pages_count_display = 0;
        private int _fy_pages_count_last;
        private int _fy_pages_count = 0;
        private bool _detect_fy = false;
        private bool _fy_inserted_in_excel = true;
        private bool __turnover_detect = false;
        private int _fy_row = 1;
        private int _fy_row_count = 1;
        private int __send_email = 0;
        private bool _isDone_fy = false;
        private string _fy_filename = "";
        private string _fy_folder_path_result;
        private string _fy_folder_path_result_xlsx;
        private string _fy_folder_path_result_locate;
        int _fy_secho = 0;
        int _fy_i = 0;
        int _fy_ii = 0;
        private bool __isLimit = false;
        int _fy_get_ii = 1;
        int _fy_get_ii_display = 1;
        int _fy_current_index = 1;
        string _fy_path = Path.Combine(Path.GetTempPath(), "cd_fy.txt");
        StringBuilder _fy_csv = new StringBuilder();
        StringBuilder _fy_csv_memberrregister_custom = new StringBuilder();
        private string _fy_start_datetime;
        private string _fy_finish_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        private bool _fy_no_result;
        private string _fy_current_datetime = "";
        private int _test_fy_gettotal_count_record;
        // MEMBER LIST
        private string _fy_playerlist_cn = "";
        private string _fy_playerlist_ea = "";
        private string _fy_id_playerlist;

        private bool _fy_cn_ea;
        private bool isInsertMemberRegister = false;
        private bool _isSecondRequest_fy = false;
        private bool _isThirdRequest_fy = false;
        private bool isButtonStart_fy = false;
        private string _fy_ld;
        private bool _isSecondRequestFinish_fy = false;
        // asd added
        private bool isFYRegistrationDone = false;
        private int display_count_fy = 0;
        private int display_count_turnover_fy = 0;
        List<String> getmemberlist_fy = new List<String>();
        List<String> __getcontactemail = new List<String>();
        private bool isBetRecordInsert = false;
        private bool isStopClick_fy = false;
        private string get_value;
        private string __brand_code = "FY";
        private string __brand_color = "#DE1E70";

        // Border
        const int _ = 1;
        new Rectangle Top { get { return new Rectangle(0, 0, this.ClientSize.Width, _); } }
        new Rectangle Left { get { return new Rectangle(0, 0, _, this.ClientSize.Height); } }
        new Rectangle Bottom { get { return new Rectangle(0, this.ClientSize.Height - _, this.ClientSize.Width, _); } }
        new Rectangle Right { get { return new Rectangle(this.ClientSize.Width - _, 0, _, ClientSize.Height); } }

        // Minimize Click in Taskbar
        const int WS_MINIMIZEBOX = 0x20000;
        const int CS_DBLCLKS = 0x8;

        // Rounded Border Radius
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
        );

        public Main_Form()
        {
            InitializeComponent();
            Region = new Region(RoundedRectangle.Create(new Rectangle(0, 0, Size.Width, Size.Height), 8, RoundedRectangle.RectangleCorners.TopRight | RoundedRectangle.RectangleCorners.TopLeft | RoundedRectangle.RectangleCorners.BottomLeft | RoundedRectangle.RectangleCorners.BottomRight));

            Opacity = 0;
            timer.Interval = 20;
            timer.Tick += new EventHandler(FadeIn);
            timer.Start();
        }

        private void FadeIn(object sender, EventArgs e)
        {
            if (Opacity >= 1)
            {
                timer_landing.Start();
            }
            else
            {
                Opacity += 0.05;
            }
        }

        private void Main_Form_Load(object sender, EventArgs e)
        {
            // FY
            webBrowser_fy.Navigate("http://cs.ying168.bet/account/login");
            comboBox_fy.SelectedIndex = 0;
            comboBox_fy_list.SelectedIndex = 0;
            dateTimePicker_start_fy.Format = DateTimePickerFormat.Custom;
            dateTimePicker_start_fy.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker_end_fy.Format = DateTimePickerFormat.Custom;
            dateTimePicker_end_fy.CustomFormat = "yyyy-MM-dd HH:mm:ss";
        }

        private void Main_Form_Shown(object sender, EventArgs e)
        {
            FY_Cronos_Data.Properties.Settings.Default.filelocation = "\\\\192.168.10.22\\ssi-reporting";
            FY_Cronos_Data.Properties.Settings.Default.Save();
            label_filelocation.Text = FY_Cronos_Data.Properties.Settings.Default.filelocation;

            // asd comment
            // asdasd update
            GetMemberList_FY();
            GetBonusCode_FY();
            GetGamePlatform_FY();
            GetPaymentType_FY();
        }

        // ----------------
        // FY ----------------
        // ----------------

        // ComboBox Selected Index Changed
        private void comboBox_fy_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_fy.SelectedIndex == 0)
            {
                // Yesterday
                string start_fy = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_start_fy = DateTime.ParseExact(start_fy, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                dateTimePicker_start_fy.Value = datetime_start_fy;

                string end_fy = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 23:59:59");
                DateTime datetime_end_fy = DateTime.ParseExact(end_fy, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end_fy.Value = datetime_end_fy;
            }
            else if (comboBox_fy.SelectedIndex == 1)
            {
                // Last Week
                DayOfWeek weekStart = DayOfWeek.Sunday;
                DateTime startingDate = DateTime.Today;

                while (startingDate.DayOfWeek != weekStart)
                {
                    startingDate = startingDate.AddDays(-1);
                }

                DateTime datetime_start_fy = startingDate.AddDays(-7);
                dateTimePicker_start_fy.Value = datetime_start_fy;

                string last = startingDate.AddDays(-1).ToString("yyyy-MM-dd 23:59:59");
                DateTime datetime_end_fy = DateTime.ParseExact(last, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end_fy.Value = datetime_end_fy;
            }
            else if (comboBox_fy.SelectedIndex == 2)
            {
                // Last Month
                var today = DateTime.Today;
                var month = new DateTime(today.Year, today.Month, 1);
                var first = month.AddMonths(-1).ToString("yyyy-MM-dd 00:00:00");
                var last = month.AddDays(-1).ToString("yyyy-MM-dd 23:59:59");

                DateTime datetime_start_fy = DateTime.ParseExact(first, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_start_fy.Value = datetime_start_fy;

                DateTime datetime_end_fy = DateTime.ParseExact(last, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end_fy.Value = datetime_end_fy;
            }
        }
        
        static int LineNumber([System.Runtime.CompilerServices.CallerLineNumber] int lineNumber = 0)
        {
            return lineNumber;
        }

        // Loaded
        private async void webBrowser_fy_DocumentCompletedAsync(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser_fy.ReadyState == WebBrowserReadyState.Complete)
            {
                if (e.Url == webBrowser_fy.Url)
                {
                    try
                    {
                        if (webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/account/login"))
                        {
                            if (__isStart)
                            {
                                button_fy_stop.PerformClick();
                                button_fy_start.Visible = false;
                                timer_fy_start.Stop();

                                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>The application have been logout, please re-login again.</b></body></html>");
                                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>The application have been logout, please re-login again.</b></body></html>");
                                __send_email = 0;
                            }
                            
                            label_status.Text = "Logout";
                            __isLogin = false;
                            __isStart = true;
                            webBrowser_fy.Visible = true;
                            webBrowser_fy.BringToFront();
                            webBrowser_fy.Document.Window.ScrollTo(0, 180);
                            webBrowser_fy.Document.GetElementById("csname").SetAttribute("value", "fycronos");
                            webBrowser_fy.Document.GetElementById("cspwd").SetAttribute("value", "cr888@@@");
                            webBrowser_fy.Document.GetElementById("la").Enabled = false;
                            get_value = webBrowser_fy.Document.GetElementById("la").GetAttribute("value");

                            if (get_value == "")
                            {
                                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                                __send_email = 0;

                                isClose = false;
                                Environment.Exit(0);
                            }

                            // start detect
                            if (FY_Cronos_Data.Properties.Settings.Default.______start_detect == "")
                            {
                                // done
                                FY_Cronos_Data.Properties.Settings.Default.______start_detect = "0";
                                FY_Cronos_Data.Properties.Settings.Default.Save();
                            }
                        }

                        if (webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/player/list") || webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/site/index") || webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/player/online") || webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/message/platform"))
                        { 
                            if (!isButtonStart_fy)
                            {
                                button_fy_start.Visible = true;
                                webBrowser_fy.Visible = false;
                                panel_fy_status.Visible = true;
                                __isLogin = true;

                                string date = DateTime.Now.ToString("yyyy-MM-dd");
                                string get_path = label_filelocation.Text + "\\Cronos Data\\FY\\" + date + "\\FY Registration.xlsx";
                                if (File.Exists(get_path))
                                {
                                    // added auto
                                    // start detect
                                    // done
                                    if (FY_Cronos_Data.Properties.Settings.Default.______start_detect == "0")
                                    {
                                        label_status.Text = "Waiting";
                                        button_fy_start.Enabled = false;
                                        panel_fy_filter.Enabled = false;
                                    }
                                    // payment
                                    else if (FY_Cronos_Data.Properties.Settings.Default.______start_detect == "1")
                                    {
                                        label_status.Text = "Running";
                                        comboBox_fy_list.SelectedIndex = 0;
                                        button_fy_start.PerformClick();
                                    }
                                    // bonus
                                    else if (FY_Cronos_Data.Properties.Settings.Default.______start_detect == "2")
                                    {
                                        label_status.Text = "Running";
                                        comboBox_fy_list.SelectedIndex = 1;
                                        button_fy_start.PerformClick();
                                    }
                                    // bet record
                                    else if (FY_Cronos_Data.Properties.Settings.Default.______start_detect == "3")
                                    {
                                        label_status.Text = "Running";
                                        comboBox_fy_list.SelectedIndex = 2;
                                        button_fy_start.PerformClick();
                                    }
                                }
                                else
                                {
                                    // registration member list
                                    label_status.Text = "Running";
                                    panel_fy_status.Visible = true;
                                    label_fy_page_count_1.Visible = true;
                                    label_fy_total_records_1.Visible = true;
                                    label_fy_page_count.Visible = true;
                                    label_fy_currentrecord.Visible = true;
                                    panel_fy_datetime.Visible = true;
                                    panel_fy_status.BringToFront();
                                    button_fy_start.Visible = false;
                                    panel_fy_filter.Visible = false;
                                    pictureBox_fy_loader.Visible = true;
                                    _fy_current_datetime = "";
                                    label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                                    _fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                                    label_fy_status.Visible = true;
                                    timer_fy.Start();
                                    isButtonStart_fy = true;
                                    FY_GetPlayerListsAsync();
                                }
                            }
                            else
                            {
                                button_fy_start.Visible = true;
                                webBrowser_fy.Visible = false;
                                panel_fy_status.Visible = true;
                                __isLogin = true;

                                string date = DateTime.Now.ToString("yyyy-MM-dd");
                                string get_path = label_filelocation.Text + "\\Cronos Data\\FY\\" + date + "\\FY Registration.xlsx";
                                if (File.Exists(get_path))
                                {
                                    label_fy_insert.Visible = false;
                                    panel_fy_status.Visible = false;
                                    button_fy_start.Visible = true;
                                    panel_fy_filter.Enabled = true;
                                    //button_filelocation.Enabled = true;

                                    button_fy_proceed.Visible = false;
                                    label_fy_locatefolder.Visible = false;

                                    label_fy_page_count.Text = "-";
                                    label_fy_currentrecord.Text = "-";
                                    label_fy_finish_datetime.Text = "-";
                                    label_fy_elapsed.Text = "-";

                                    panel_fy_datetime.Location = new Point(66, 226);

                                    // set default variables
                                    _fy_bet_records.Clear();
                                    fy_datetime.Clear();
                                    fy_gettotal.Clear();
                                    _total_records_fy = 0;
                                    _total_page_fy = 0;
                                    _fy_displayinexel_i = 0;
                                    _fy_pages_count_display = 0;
                                    _fy_pages_count = 0;
                                    _detect_fy = false;
                                    _fy_inserted_in_excel = true;
                                    _fy_row = 1;
                                    _fy_row_count = 1;
                                    _isDone_fy = false;
                                    _fy_secho = 0;
                                    _fy_i = 0;
                                    _fy_ii = 0;
                                    _fy_get_ii = 1;
                                    _fy_get_ii_display = 1;
                                    _fy_folder_path_result = "";
                                    _fy_folder_path_result_locate = "";
                                    _fy_current_index = 1;
                                    _fy_csv.Clear();

                                    // MEMBER LIST
                                    _fy_playerlist_cn = "";
                                    _fy_playerlist_ea = "";
                                    _fy_id_playerlist = "";

                                    label_fy_status.Visible = true;
                                    label_fy_page_count_1.Visible = true;
                                    label_fy_total_records_1.Visible = true;
                                    label_fy_page_count.Visible = true;
                                    label_fy_currentrecord.Visible = true;
                                    panel_fy_datetime.Visible = true;

                                    timer_fy_start_button.Stop();
                                    label_fy_count.Visible = false;
                                    button_fy_stop.Visible = false;

                                    fy_datetime.Clear();
                                    fy_gettotal.Clear();
                                    fy_gettotal_test.Clear();

                                    button_fy_start.PerformClick();
                                }
                                else
                                {
                                    // registration member list
                                    _fy_bet_records.Clear();
                                    fy_datetime.Clear();
                                    fy_gettotal.Clear();
                                    _total_records_fy = 0;
                                    _total_page_fy = 0;
                                    _fy_displayinexel_i = 0;
                                    _fy_pages_count_display = 0;
                                    _fy_pages_count = 0;
                                    _detect_fy = false;
                                    _fy_inserted_in_excel = true;
                                    _fy_row = 1;
                                    _fy_row_count = 1;
                                    _isDone_fy = false;
                                    _fy_secho = 0;
                                    _fy_i = 0;
                                    _fy_ii = 0;
                                    _fy_get_ii = 1;
                                    _fy_get_ii_display = 1;
                                    _fy_folder_path_result = "";
                                    _fy_folder_path_result_locate = "";
                                    _fy_current_index = 1;
                                    _fy_csv.Clear();

                                    // MEMBER LIST
                                    _fy_playerlist_cn = "";
                                    _fy_playerlist_ea = "";
                                    _fy_id_playerlist = "";
                                    
                                    panel_fy_status.Visible = true;
                                    label_fy_page_count_1.Visible = true;
                                    label_fy_total_records_1.Visible = true;
                                    label_fy_page_count.Visible = true;
                                    label_fy_currentrecord.Visible = true;
                                    panel_fy_datetime.Visible = true;
                                    panel_fy_status.BringToFront();
                                    button_fy_start.Visible = false;
                                    panel_fy_filter.Visible = false;
                                    pictureBox_fy_loader.Visible = true;
                                    _fy_current_datetime = "";
                                    label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                                    _fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                                    label_fy_status.Visible = true;
                                    timer_fy.Start();
                                    isButtonStart_fy = true;
                                    FY_GetPlayerListsAsync();
                                }
                                
                                label_status.Text = "Running";
                                timer_fy_start.Start();
                            }
                        }
                    }
                    catch (Exception err)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                        __send_email = 0;

                        isClose = false;
                        Environment.Exit(0);
                    }
                }
            }
        }

        // MEMBER LIST ----------------
        private async void FY_GetPlayerListsAsync()
        {
            try
            {
                // status
                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                label_fy_status.Text = "status: doing calculation... --- MEMBER LIST";

                var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                var reqparm_gettotal = new NameValueCollection
                {
                    { "s_btype", ""},
                    { "skip", "0"},
                    { "groupid", "0"},
                    { "s_type", "1"},
                    { "s_status_search", ""},
                    { "s_keyword", ""},
                    { "s_playercurrency", "ALL"},
                    { "s_phone", "on"},
                    { "s_email", "on"},
                    { "data[0][name]", "sEcho"},
                    { "data[0][value]", _fy_secho++.ToString()},
                    { "data[1][name]", "iColumns"},
                    { "data[1][value]", "13"},
                    { "data[2][name]", "sColumns"},
                    { "data[2][value]", ""},
                    { "data[3][name]", "iDisplayStart"},
                    { "data[3][value]", "0"},
                    { "data[4][name]", "iDisplayLength"},
                    { "data[4][value]", "1"}
                };

                byte[] result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/player/listAjax1", "POST", reqparm_gettotal);
                string responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
                var deserializeObject_gettotal = JsonConvert.DeserializeObject(responsebody_gettotatal);
                JObject jo_gettotal = JObject.Parse(deserializeObject_gettotal.ToString());
                JToken jt_gettotal = jo_gettotal.SelectToken("$.iTotalRecords");
                _total_records_fy += double.Parse(jt_gettotal.ToString());
                double get_total_records_fy = 0;
                get_total_records_fy = double.Parse(jt_gettotal.ToString());

                fy_gettotal_test.Add(get_total_records_fy.ToString());
                double result_total_records = get_total_records_fy / _display_length_fy;

                if (result_total_records.ToString().Contains("."))
                {
                    _total_page_fy += Convert.ToInt32(Math.Floor(result_total_records)) + 1;
                }
                else
                {
                    _total_page_fy += Convert.ToInt32(Math.Floor(result_total_records));
                }

                label_fy_page_count.Text = "0 of " + _total_page_fy.ToString("N0");
                label_fy_currentrecord.Text = "0 of " + Convert.ToInt32(_total_records_fy).ToString("N0");

                var reqparm = new NameValueCollection
                {
                    { "s_btype", ""},
                    { "skip", "0"},
                    { "groupid", "0"},
                    { "s_type", "1"},
                    { "s_status_search", ""},
                    { "s_keyword", ""},
                    { "s_playercurrency", "ALL"},
                    { "data[0][name]", "sEcho"},
                    { "data[0][value]", _fy_secho++.ToString()},
                    { "data[1][name]", "iColumns"},
                    { "data[1][value]", "13"},
                    { "data[2][name]", "sColumns"},
                    { "data[2][value]", ""},
                    { "data[3][name]", "iDisplayStart"},
                    { "data[3][value]", "0"},
                    { "data[4][name]", "iDisplayLength"},
                    { "data[4][value]", _display_length_fy.ToString()}
                };

                label_fy_status.Text = "status: getting data... --- MEMBER LIST";

                byte[] result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/player/listAjax1", "POST", reqparm);
                string responsebody = Encoding.UTF8.GetString(result);
                var deserializeObject = JsonConvert.DeserializeObject(responsebody);

                jo_fy = JObject.Parse(deserializeObject.ToString());
                JToken count = jo_fy.SelectToken("$.aaData");
                _result_count_json_fy = count.Count();
                
                FY_PlayerListAsync();
            }
            catch (Exception err)
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() +"</b></body></html>");
                __send_email = 0;
            }
        }
        
        private async void FY_PlayerListAsync()
        {
            if (_fy_inserted_in_excel)
            {
                for (int i = _fy_i; i < _total_page_fy; i++)
                {
                    button_fy_start.Visible = false;

                    if (!_fy_inserted_in_excel)
                    {
                        break;
                    }
                    else
                    {
                        _fy_i = i;
                        _fy_pages_count_display++;
                    }

                    for (int ii = 0; ii < _result_count_json_fy; ii++)
                    {
                        Application.DoEvents();

                        if (!__isLogin)
                        {
                            _fy_inserted_in_excel = false;
                            
                            break;
                        }

                        _test_fy_gettotal_count_record++;

                        if (_fy_pages_count_display != 0 && _fy_pages_count_display <= _total_page_fy)
                        {
                            label_fy_page_count.Text = _fy_pages_count_display.ToString("N0") + " of " + _total_page_fy.ToString("N0");
                        }

                        _fy_ii = ii;
                        JToken username__id = jo_fy.SelectToken("$.aaData[" + ii + "][0]").ToString();
                        string username = Regex.Match(username__id.ToString(), "username=\\\"(.*?)\\\"").Groups[1].Value;
                        _fy_id_playerlist = Regex.Match(username__id.ToString(), "player=\\\"(.*?)\\\"").Groups[1].Value;
                        JToken name = jo_fy.SelectToken("$.aaData[" + ii + "][2]").ToString().Replace("\"", "");
                        JToken source = jo_fy.SelectToken("$.aaData[" + ii + "][3]").ToString().Replace("\"", "");
                        JToken vip = jo_fy.SelectToken("$.aaData[" + ii + "][4]").ToString().Replace("\"", "");
                        if (vip.ToString().Contains("label"))
                        {
                            string vip_get = Regex.Match(vip.ToString(), "<label(.*?)>(.*?)</label>").Groups[2].Value;
                            vip = vip_get;
                        }
                        JToken last_login_date___ip = jo_fy.SelectToken("$.aaData[" + ii + "][11]");
                        string last_login_date = last_login_date___ip.ToString().Substring(0, 19);
                        string ip = Regex.Match(last_login_date___ip.ToString(), "<label(.*?)>(.*?)</label>").Groups[2].Value;
                        JToken date_register__register_domain = jo_fy.SelectToken("$.aaData[" + ii + "][12]");
                        string date_register = date_register__register_domain.ToString().Substring(0, 10);
                        string date_time_register = date_register__register_domain.ToString().Substring(14, 8);
                        string register_domain = date_register__register_domain.ToString().Substring(27);
                        JToken status = jo_fy.SelectToken("$.aaData[" + ii + "][13]").ToString().Replace("\"", "");
                        DateTime month_registration = Convert.ToDateTime(date_register);

                        // uncomment
                        //getmemberlist_fy.Add(username);
                        //getmemberlist_fy.Add(vip.ToString());
                        
                        string cn = "";
                        string ea = "";

                        for (int i_ = 0; i_ < __getcontactemail.Count; i_ += 3)
                        {
                            if (__getcontactemail[i_] == username)
                            {
                                cn = __getcontactemail[i_ + 1];
                                ea = __getcontactemail[i_ + 2];
                                break;
                            }
                        }

                        if (cn == "" && ea == "")
                        {
                            await FY_PlayerListContactNumberEmailAsync(_fy_id_playerlist);
                        }
                        else
                        {
                            _fy_playerlist_cn = cn;
                            _fy_playerlist_ea = ea;
                        }

                        if (_fy_get_ii == 1)
                        {
                            var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}", "Brand", "Username", "Name", "Status", "Date Register", "Last Login Date", "Last Deposit Date", "Contact Number", "Email", "VIP Level", "Registration Time", "Month Register", "FD Date", "FD Month", "Source", "IP Address", "Register Domain");
                            _fy_csv.AppendLine(header);
                        }
                        
                        var newLine = "";
                        if (_fy_playerlist_cn != "" || _fy_playerlist_ea != "")
                        {
                            newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}", "FY", "\"" + username + "\"", "\"" + name + "\"", "\"" + status + "\"", "\"" + date_register + " " + date_time_register + "\"", "\"" + last_login_date + "\"", "\"" + "" + "\"", "\"" + _fy_playerlist_cn + "\"", "\"" + _fy_playerlist_ea + "\"", "\"" + vip + "\"", "\"" + date_register + "\"", "\"" + month_registration.ToString("MM/01/yyyy") + "\"", "\"" + "" + "\"", "\"" + "" + "\"", "\"" + source + "\"", "\"" + ip + "\"", "\"" + register_domain + "\"");
                        }
                        else
                        {
                            newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}", "FY", "\"" + username + "\"", "\"" + name + "\"", "\"" + status + "\"", "\"" + date_register + " " + date_time_register + "\"", "\"" + last_login_date + "\"", "\"" + "" + "\"", "\"" + "" + "\"", "\"" + "" + "\"", "\"" + vip + "\"", "\"" + date_register + "\"", "\"" + month_registration.ToString("MM/01/yyyy") + "\"", "\"" + "" + "\"", "\"" + "" + "\"", "\"" + source + "\"", "\"" + ip + "\"", "\"" + register_domain + "\"");
                        }

                        _fy_playerlist_cn = "";
                        _fy_playerlist_ea = "";

                        _fy_csv.AppendLine(newLine);

                        if ((_fy_get_ii) == _limit_fy)
                        {
                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: saving excel... --- MEMBER LIST";

                            __isLimit = true;
                            __pageCountPlayer++;

                            _fy_get_ii = 0;

                            if (_fy_current_datetime == "")
                            {
                                _fy_current_datetime = DateTime.Now.ToString("yyyy-MM-dd");
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data"))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data");
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY"))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY");
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime);
                            }

                            _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\FY Registration " + __pageCountPlayer + ".txt";
                            _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\FY Registration " + __pageCountPlayer + ".xlsxxt";
                            _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime;

                            if (File.Exists(_fy_folder_path_result))
                            {
                                File.Delete(_fy_folder_path_result);
                            }

                            if (File.Exists(_fy_folder_path_result_xlsx))
                            {
                                File.Delete(_fy_folder_path_result_xlsx);
                            }

                            _fy_csv.ToString().Reverse();
                            File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                            Excel.Application app = new Excel.Application();
                            Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            Excel.Worksheet worksheet = wb.ActiveSheet;
                            worksheet.Activate();
                            worksheet.Application.ActiveWindow.SplitRow = 1;
                            worksheet.Application.ActiveWindow.FreezePanes = true;
                            Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                            firstRow.AutoFilter(1,
                                                Type.Missing,
                                                Excel.XlAutoFilterOperator.xlAnd,
                                                Type.Missing,
                                                true);
                            worksheet.Columns[6].NumberFormat = "MM/dd/yyyy";
                            Excel.Range usedRange = worksheet.UsedRange;
                            Excel.Range rows = usedRange.Rows;
                            int count = 0;
                            foreach (Excel.Range row in rows)
                            {
                                if (count == 0)
                                {
                                    Excel.Range firstCell = row.Cells[1];

                                    string firstCellValue = firstCell.Value as String;

                                    if (!string.IsNullOrEmpty(firstCellValue))
                                    {
                                        row.Interior.Color = Color.FromArgb(222, 30, 112);
                                        row.Font.Color = Color.FromArgb(255, 255, 255);
                                    }

                                    break;
                                }

                                count++;
                            }
                            int i_;
                            for (i_ = 1; i_ <= 17; i_++)
                            {
                                worksheet.Columns[i_].ColumnWidth = 18;
                            }
                            wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            wb.Close();
                            app.Quit();
                            Marshal.ReleaseComObject(app);

                            if (File.Exists(_fy_folder_path_result))
                            {
                                File.Delete(_fy_folder_path_result);
                            }

                            _fy_csv.Clear();

                            label_fy_currentrecord.Text = (_fy_get_ii_display).ToString("N0") + " of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                            label_fy_currentrecord.Invalidate();
                            label_fy_currentrecord.Update();
                        }
                        else
                        {
                            label_fy_currentrecord.Text = (_fy_get_ii_display).ToString("N0") + " of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                            label_fy_currentrecord.Invalidate();
                            label_fy_currentrecord.Update();
                        }

                        _fy_get_ii++;
                        _fy_get_ii_display++;
                    }

                    _result_count_json_fy = 0;

                    // web client request
                    await GetDataFYPagesPlayerListAsync();
                }

                if (_fy_inserted_in_excel)
                {
                    FY_PlayerListInsertDone();
                    _isDone_fy = true;
                }
            }
        }
        
        private async Task GetDataFYPagesPlayerListAsync()
        {
            try
            {
                var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                int result_pages;

                if (_detect_fy)
                {
                    _detect_fy = false;
                    result_pages = (Convert.ToInt32(_display_length_fy) * _fy_pages_count);
                }
                else
                {
                    _fy_pages_count++;
                    result_pages = (Convert.ToInt32(_display_length_fy) * _fy_pages_count);
                }

                var reqparm = new NameValueCollection
                    {
                        { "s_btype", ""},
                        { "skip", "0"},
                        { "groupid", "0"},
                        { "s_type", "1"},
                        { "s_status_search", ""},
                        { "s_keyword", ""},
                        { "s_playercurrency", "ALL"},
                        { "data[0][name]", "sEcho"},
                        { "data[0][value]", _fy_secho++.ToString()},
                        { "data[1][name]", "iColumns"},
                        { "data[1][value]", "13"},
                        { "data[2][name]", "sColumns"},
                        { "data[2][value]", ""},
                        { "data[3][name]", "iDisplayStart"},
                        { "data[3][value]", result_pages.ToString()},
                        { "data[4][name]", "iDisplayLength"},
                        { "data[4][value]", _display_length_fy.ToString()}
                    };

                // status
                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                label_fy_status.Text = "status: getting data... --- MEMBER LIST";

                byte[] result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/player/listAjax1", "POST", reqparm);
                string responsebody = Encoding.UTF8.GetString(result);
                if (_fy_pages_count_display != _total_page_fy)
                {
                    var deserializeObject = JsonConvert.DeserializeObject(responsebody);

                    jo_fy = JObject.Parse(deserializeObject.ToString());
                    JToken count = jo_fy.SelectToken("$.aaData");
                    _result_count_json_fy = count.Count();
                }
                else
                {
                    _result_count_json_fy = 0;
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    await GetDataFYPagesPlayerListAsync();
                }
            }
        }

        private void FY_PlayerListInsertDone()
        {
            try
            {
                // status
                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                label_fy_status.Text = "status: saving excel... --- MEMBER LIST";

                if (_fy_current_datetime == "")
                {
                    _fy_current_datetime = DateTime.Now.ToString("yyyy-MM-dd");
                }

                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data"))
                {
                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data");
                }

                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY"))
                {
                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY");
                }

                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime))
                {
                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime);
                }

                if (__isLimit)
                {
                    __pageCountPlayer++;
                    _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\FY Registration " + __pageCountPlayer + ".txt";
                    _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\FY Registration " + __pageCountPlayer + ".xlsx";
                }
                else
                {
                    _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\FY Registration.txt";
                    _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\FY Registration.xlsx";
                }
                _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime;

                if (File.Exists(_fy_folder_path_result))
                {
                    File.Delete(_fy_folder_path_result);
                }

                if (File.Exists(_fy_folder_path_result_xlsx))
                {
                    File.Delete(_fy_folder_path_result_xlsx);
                }

                _fy_csv.ToString().Reverse();
                File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet worksheet = wb.ActiveSheet;
                worksheet.Activate();
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                firstRow.AutoFilter(1,
                                    Type.Missing,
                                    Excel.XlAutoFilterOperator.xlAnd,
                                    Type.Missing,
                                    true);
                worksheet.Columns[6].NumberFormat = "MM/dd/yyyy";
                Excel.Range usedRange = worksheet.UsedRange;
                Excel.Range rows = usedRange.Rows;
                int count = 0;
                foreach (Excel.Range row in rows)
                {
                    if (count == 0)
                    {
                        Excel.Range firstCell = row.Cells[1];

                        string firstCellValue = firstCell.Value as String;

                        if (!string.IsNullOrEmpty(firstCellValue))
                        {
                            row.Interior.Color = Color.FromArgb(222, 30, 112);
                            row.Font.Color = Color.FromArgb(255, 255, 255);
                        }

                        break;
                    }

                    count++;
                }
                int i;
                for (i = 1; i <= 17; i++)
                {
                    worksheet.Columns[i].ColumnWidth = 18;
                }
                wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();
                app.Quit();
                Marshal.ReleaseComObject(app);

                if (File.Exists(_fy_folder_path_result))
                {
                    File.Delete(_fy_folder_path_result);
                }

                _fy_csv.Clear();

                //FYHeader();

                Invoke(new Action(() =>
                {
                    label_fy_finish_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                    timer_fy.Stop();
                    pictureBox_fy_loader.Visible = false;
                    button_fy_proceed.Visible = true;
                    label_fy_locatefolder.Visible = true;
                    label_fy_status.ForeColor = Color.FromArgb(34, 139, 34);
                    label_fy_status.Text = "status: done --- MEMBER LIST";
                    panel_fy_datetime.Location = new Point(5, 226);

                    // added asdasd
                    button_fy_proceed.PerformClick();
                }));

                //var notification = new NotifyIcon()
                //{
                //    Visible = true,
                //    Icon = SystemIcons.Information,
                //    BalloonTipIcon = ToolTipIcon.Info,
                //    BalloonTipTitle = "FY MEMBER LIST",
                //    BalloonTipText = "Filter of...\nStart Time: " + dateTimePicker_start_fy.Text + "\nEnd Time: " + dateTimePicker_end_fy.Text + "\n\nStart-Finish...\nStart Time: " + label_start_fy.Text + "\nFinish Time: " + label_end_fy.Text,
                //};

                //notification.ShowBalloonTip(1000);

                //timer_fy_start.Start();
            }
            catch (Exception err)
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                __send_email = 0;
            }
        }

        private async Task FY_PlayerListContactNumberEmailAsync(string id)
        {
            try
            {
                var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                string result_gettotal_responsebody = await wc.DownloadStringTaskAsync("http://cs.ying168.bet/player/playerDetailBox?id=" + id);

                int i_label = 0;
                int cn = 0;
                int ea = 0;
                bool cn_detect = false;
                bool ea_detect = false;
                bool cn_ = false;
                bool ea_ = false;

                Regex ItemRegex_label = new Regex("<label class=\"control-label\">(.*?)</label>", RegexOptions.Compiled);
                foreach (Match ItemMatch in ItemRegex_label.Matches(result_gettotal_responsebody))
                {
                    string item = ItemMatch.Groups[1].Value;
                    i_label++;
                    //MessageBox.Show(item);
                    if (item.Contains("Cellphone No") || item.Contains("手机号"))
                    {
                        cn = i_label;
                        cn_detect = true;
                    }
                    else if (item.Contains("E-mail Address") || item.Contains("邮箱"))
                    {
                        ea = i_label;
                        ea_detect = true;
                    }
                    else if (item.Contains("Agent No") || item.Contains("代理编号"))
                    {
                        if (!cn_detect)
                        {
                            cn_ = true;
                        }

                        if (!ea_detect)
                        {
                            ea_ = true;
                        }
                    }
                }

                if (cn_)
                {
                    cn--;
                }

                if (ea_)
                {
                    ea--;
                }

                int i_span = 0;

                Regex ItemRegex_span = new Regex("<span class=\"text\">(.*?)</span>", RegexOptions.Compiled);
                foreach (Match ItemMatch in ItemRegex_span.Matches(result_gettotal_responsebody))
                {
                    i_span++;
                    string item = ItemMatch.Groups[1].Value;

                    if (i_span == cn)
                    {
                        if (ItemMatch.Groups[1].Value != "")
                        {
                            _fy_playerlist_cn = "86" + ItemMatch.Groups[1].Value;
                        }
                        else
                        {
                            _fy_playerlist_cn = ItemMatch.Groups[1].Value;
                        }
                    }
                    else if (i_span == ea)
                    {
                        _fy_playerlist_ea = ItemMatch.Groups[1].Value;
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    await FY_PlayerListContactNumberEmailAsync(id);
                }
            }
        }
        
        private async Task FY_GetTotal(string start_datetime, string end_datetime)
        {
            try
            {
                var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                byte[] result_gettotal = null;
                string responsebody_gettotatal = "";

                int selected_index = comboBox_fy_list.SelectedIndex;
                if (selected_index == 0)
                {
                    // payment
                    FY_Cronos_Data.Properties.Settings.Default.______start_detect = "1";
                    FY_Cronos_Data.Properties.Settings.Default.Save();
                    
                    if (!_isSecondRequestFinish_fy)
                    {
                        if (!_isSecondRequest_fy)
                        {
                            // Deposit Record
                            var reqparm = new NameValueCollection
                            {
                                {"s_btype", ""},
                                {"s_StartTime", start_datetime},
                                {"s_EndTime", end_datetime},
                                {"dno", ""},
                                {"s_dpttype", "0"},
                                {"s_type", "1"},
                                {"s_transtype", "0"},
                                {"s_ppid", "0"},
                                {"s_payoption", "0"},
                                {"groupid", "0"},
                                {"s_keyword", ""},
                                {"s_playercurrency", "ALL"},
                                {"skip", "0"},
                                {"data[0][name]", "sEcho"},
                                {"data[0][value]", _fy_secho++.ToString()},
                                {"data[1][name]", "iColumns"},
                                {"data[1][value]", "17"},
                                {"data[2][name]", "sColumns"},
                                {"data[2][value]", ""},
                                {"data[3][name]", "iDisplayStart"},
                                {"data[3][value]", "0"},
                                {"data[4][name]", "iDisplayLength"},
                                {"data[4][value]", "1"}
                            };

                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: doing calculation... DEPOSIT RECORD";

                            result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptHistoryAjax", "POST", reqparm);
                            if (get_value == "en")
                            {
                                responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal).Remove(0, 1);
                            }
                            else
                            {
                                responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
                            }
                        }
                        else
                        {
                            // Manual Deposit Record
                            var reqparm = new NameValueCollection
                            {
                                {"s_btype", ""},
                                {"ptype", "1212"},
                                {"fs_ptype", "1212"},
                                {"s_StartTime", start_datetime},
                                {"s_EndTime", end_datetime},
                                {"s_type", "1"},
                                {"s_keyword", ""},
                                {"s_playercurrency", "ALL"},
                                {"data[0][name]", "sEcho"},
                                {"data[0][value]", _fy_secho++.ToString()},
                                {"data[1][name]", "iColumns"},
                                {"data[1][value]", "18"},
                                {"data[2][name]", "sColumns"},
                                {"data[2][value]", ""},
                                {"data[3][name]", "iDisplayStart"},
                                {"data[3][value]", "0"},
                                {"data[4][name]", "iDisplayLength"},
                                {"data[4][value]", "1"}
                            };

                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: doing calculation... M-DEPOSIT RECORD";

                            result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptCorrectionAjax", "POST", reqparm);
                            responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
                        }
                    }
                    else
                    {
                        if (!_isThirdRequest_fy)
                        {
                            // Withdrawal Record
                            var reqparm = new NameValueCollection
                            {
                                {"s_btype", ""},
                                {"s_StartTime", start_datetime},
                                {"s_EndTime", end_datetime},
                                {"s_wtdAmtFr", ""},
                                {"s_wtdAmtTo", ""},
                                {"s_dpttype", "0"},
                                {"skip", "0"},
                                {"s_type", "1"},
                                {"s_keyword", "0"},
                                {"s_playercurrency", "ALL"},
                                {"wttype", "0"},
                                {"data[0][name]", "sEcho"},
                                {"data[0][value]", _fy_secho++.ToString()},
                                {"data[1][name]", "iColumns"},
                                {"data[1][value]", "18"},
                                {"data[2][name]", "sColumns"},
                                {"data[2][value]", ""},
                                {"data[3][name]", "iDisplayStart"},
                                {"data[3][value]", "0"},
                                {"data[4][name]", "iDisplayLength"},
                                {"data[4][value]", "1"}
                            };

                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: doing calculation... WITHDRAWAL RECORD";

                            result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/wtdHistoryAjax", "POST", reqparm);
                            responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
                        }
                        else
                        {
                            // Manual Withdrawal Record
                            var reqparm = new NameValueCollection
                            {
                                {"s_btype", ""},
                                {"ptype", "1313"},
                                {"fs_ptype", "1313"},
                                {"s_StartTime", start_datetime},
                                {"s_EndTime", end_datetime},
                                {"s_type", "1"},
                                {"s_keyword", ""},
                                {"s_playercurrency", "ALL"},
                                {"data[0][name]", "sEcho"},
                                {"data[0][value]", _fy_secho++.ToString()},
                                {"data[1][name]", "iColumns"},
                                {"data[1][value]", "18"},
                                {"data[2][name]", "sColumns"},
                                {"data[2][value]", ""},
                                {"data[3][name]", "iDisplayStart"},
                                {"data[3][value]", "0"},
                                {"data[4][name]", "iDisplayLength"},
                                {"data[4][value]", "1"}
                            };

                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: doing calculation... M-WITHDRAWAL RECORD";

                            result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptCorrectionAjax", "POST", reqparm);
                            responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
                        }
                    }
                }
                else if (selected_index == 1)
                {
                    if (!_isSecondRequest_fy)
                    {
                        // Manual Bonus Report
                        var reqparm = new NameValueCollection
                        {
                            {"s_btype", ""},
                            {"ptype", "1411"},
                            {"fs_ptype", "1411"},
                            {"s_StartTime", start_datetime},
                            {"s_EndTime", end_datetime},
                            {"s_type", "1"},
                            {"s_keyword", "0"},
                            {"s_playercurrency", "ALL"},
                            {"wttype", "0"},
                            {"data[0][name]", "sEcho"},
                            {"data[0][value]", _fy_secho++.ToString()},
                            {"data[1][name]", "iColumns"},
                            {"data[1][value]", "18"},
                            {"data[2][name]", "sColumns"},
                            {"data[2][value]", ""},
                            {"data[3][name]", "iDisplayStart"},
                            {"data[3][value]", "0"},
                            {"data[4][name]", "iDisplayLength"},
                            {"data[4][value]", "1"}
                        };

                        // status
                        label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                        label_fy_status.Text = "status: doing calculation... M-BONUS RECORD";

                        result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptCorrectionAjax", "POST", reqparm);
                        responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
                    }
                    else
                    {
                        // Generated Bonus Report
                        DateTime start_datetime_replace = DateTime.ParseExact(start_datetime, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        var reqparm = new NameValueCollection
                        {
                            {"s_btype", ""},
                            {"skip", "0"},
                            {"s_StartTime", start_datetime_replace.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss")},
                            {"s_EndTime", end_datetime},
                            {"s_type", "0"},
                            {"s_keyword", "0"},
                            {"data[0][name]", "sEcho"},
                            {"data[0][value]", _fy_secho++.ToString()},
                            {"data[1][name]", "iColumns"},
                            {"data[1][value]", "18"},
                            {"data[2][name]", "sColumns"},
                            {"data[2][value]", ""},
                            {"data[3][name]", "iDisplayStart"},
                            {"data[3][value]", "0"},
                            {"data[4][name]", "iDisplayLength"},
                            {"data[4][value]", "1"}
                        };

                        // status
                        label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                        label_fy_status.Text = "status: doing calculation... G-BONUS RECORD";

                        result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/getRakeBackHistory", "POST", reqparm);
                        responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
                    }
                }
                else if (selected_index == 2)
                {
                    // Bet Record
                    var reqparm = new NameValueCollection
                    {
                        {"s_btype", ""},
                        {"betNo", ""},
                        {"name", ""},
                        {"gpid", "0"},
                        {"wager_settle", "0"},
                        {"valid_inva", ""},
                        {"start",  start_datetime},
                        {"end", end_datetime},
                        {"skip", "0"},
                        {"ftime_188", "bettime"},
                        {"data[0][name]", "sEcho"},
                        {"data[0][value]", _fy_secho++.ToString()},
                        {"data[1][name]", "iColumns"},
                        {"data[1][value]", "12"},
                        {"data[2][name]", "sColumns"},
                        {"data[2][value]", ""},
                        {"data[3][name]", "iDisplayStart"},
                        {"data[3][value]", "0"},
                        {"data[4][name]", "iDisplayLength"},
                        {"data[4][value]", "1"}
                    };

                    // status
                    label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                    label_fy_status.Text = "status: doing calculation... BET RECORD";

                    result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm);
                    responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
                }

                var deserializeObject_gettotal = JsonConvert.DeserializeObject(responsebody_gettotatal);

                JObject jo_gettotal = JObject.Parse(deserializeObject_gettotal.ToString());
                JToken jt_gettotal = jo_gettotal.SelectToken("$.iTotalRecords");

                _total_records_fy += double.Parse(jt_gettotal.ToString());
                double get_total_records_fy = 0;
                get_total_records_fy = double.Parse(jt_gettotal.ToString());

                fy_gettotal_test.Add(get_total_records_fy.ToString());
                double result_total_records = get_total_records_fy / _display_length_fy;

                if (result_total_records.ToString().Contains("."))
                {
                    _total_page_fy += Convert.ToInt32(Math.Floor(result_total_records)) + 1;
                }
                else
                {
                    _total_page_fy += Convert.ToInt32(Math.Floor(result_total_records));
                }

                fy_gettotal.Add(_total_page_fy.ToString());

                label_fy_page_count.Text = "0 of " + _total_page_fy.ToString("N0");
                label_fy_currentrecord.Text = "0 of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                _fy_no_result = false;
            }
            catch (Exception err)
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                __send_email = 0;
            }
        }
        
        private async Task GetDataFYAsync()
        {
            try
            {
                string gettotal_start_datetime = "";
                string gettotal_end_datetime = "";

                fy_datetime.Reverse();

                int i = 0;
                foreach (var datetime in fy_datetime)
                {
                    i++;
                    string[] datetime_results = datetime.Split("*|*");
                    int ii = 0;
                    string get_start_datetime = "";
                    string get_end_datetime = "";

                    foreach (string datetime_result in datetime_results)
                    {
                        ii++;
                        if (i == 1)
                        {
                            if (ii == 1)
                            {
                                gettotal_start_datetime = datetime_result;
                            }
                            else if (ii == 2)
                            {
                                gettotal_end_datetime = datetime_result;
                            }
                        }

                        if (ii == 1)
                        {
                            get_start_datetime = datetime_result;
                        }
                        else if (ii == 2)
                        {
                            get_end_datetime = datetime_result;
                        }
                    }

                    await FY_GetTotal(get_start_datetime, get_end_datetime);
                }

                // asd label1.Text = gettotal_start_datetime + " ----- ghghg " + gettotal_end_datetime;
                var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                byte[] result = null;
                string responsebody = null;

                int selected_index = comboBox_fy_list.SelectedIndex;
                if (selected_index == 0)
                {
                    if (!_isSecondRequestFinish_fy)
                    {
                        if (!_isSecondRequest_fy)
                        {
                            // Deposit Record
                            var reqparm = new NameValueCollection
                            {
                                {"s_btype", ""},
                                {"s_StartTime", gettotal_start_datetime},
                                {"s_EndTime", gettotal_end_datetime},
                                {"dno", ""},
                                {"s_dpttype", "0"},
                                {"s_type", "1"},
                                {"s_transtype", "0"},
                                {"s_ppid", "0"},
                                {"s_payoption", "0"},
                                {"groupid", "0"},
                                {"s_keyword", ""},
                                {"s_playercurrency", "ALL"},
                                {"skip", "0"},
                                {"data[0][name]", "sEcho"},
                                {"data[0][value]", _fy_secho++.ToString()},
                                {"data[1][name]", "iColumns"},
                                {"data[1][value]", "17"},
                                {"data[2][name]", "sColumns"},
                                {"data[2][value]", ""},
                                {"data[3][name]", "iDisplayStart"},
                                {"data[3][value]", "0"},
                                {"data[4][name]", "iDisplayLength"},
                                {"data[4][value]", _display_length_fy.ToString()}
                            };

                            label_fy_status.Text = "status: getting data... DEPOSIT RECORD";

                            result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptHistoryAjax", "POST", reqparm);
                            if (get_value == "en")
                            {
                                responsebody = Encoding.UTF8.GetString(result).Remove(0, 1);
                            }
                            else
                            {
                                responsebody = Encoding.UTF8.GetString(result);
                            }
                        }
                        else
                        {
                            // Manual Deposit Record
                            var reqparm = new NameValueCollection
                            {
                                {"s_btype", ""},
                                {"ptype", "1212"},
                                {"fs_ptype", "1212"},
                                {"s_StartTime", gettotal_start_datetime},
                                {"s_EndTime", gettotal_end_datetime},
                                {"s_type", "1"},
                                {"s_keyword", ""},
                                {"data[0][name]", "sEcho"},
                                {"data[0][value]", _fy_secho++.ToString()},
                                {"data[1][name]", "iColumns"},
                                {"data[1][value]", "18"},
                                {"data[2][name]", "sColumns"},
                                {"data[2][value]", ""},
                                {"data[3][name]", "iDisplayStart"},
                                {"data[3][value]", "0"},
                                {"data[4][name]", "iDisplayLength"},
                                {"data[4][value]", _display_length_fy.ToString()}
                            };

                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: getting data... M-DEPOSIT RECORD";

                            result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptCorrectionAjax", "POST", reqparm);
                            responsebody = Encoding.UTF8.GetString(result);
                        }
                    }
                    else
                    {
                        if (!_isThirdRequest_fy)
                        {
                            // Withdrawal Record
                            var reqparm = new NameValueCollection
                        {
                            {"s_btype", ""},
                            {"s_StartTime", gettotal_start_datetime},
                            {"s_EndTime", gettotal_end_datetime},
                            {"s_wtdAmtFr", ""},
                            {"s_wtdAmtTo", ""},
                            {"s_dpttype", "0"},
                            {"skip", "0"},
                            {"s_type", "1"},
                            {"s_keyword", "0"},
                            {"s_playercurrency", "ALL"},
                            {"wttype", "0"},
                            {"data[0][name]", "sEcho"},
                            {"data[0][value]", _fy_secho++.ToString()},
                            {"data[1][name]", "iColumns"},
                            {"data[1][value]", "18"},
                            {"data[2][name]", "sColumns"},
                            {"data[2][value]", ""},
                            {"data[3][name]", "iDisplayStart"},
                            {"data[3][value]", "0"},
                            {"data[4][name]", "iDisplayLength"},
                            {"data[4][value]", _display_length_fy.ToString()}
                        };

                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: getting data... WITHDRAWAL RECORD";

                            result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/wtdHistoryAjax", "POST", reqparm);
                            responsebody = Encoding.UTF8.GetString(result);
                        }
                        else
                        {
                            // Manual Withdrawal Record
                            var reqparm = new NameValueCollection
                        {
                            {"s_btype", ""},
                            {"ptype", "1313"},
                            {"fs_ptype", "1313"},
                            {"s_StartTime", gettotal_start_datetime},
                            {"s_EndTime", gettotal_end_datetime},
                            {"s_type", "1"},
                            {"s_keyword", "0"},
                            {"s_playercurrency", "ALL"},
                            {"wttype", "0"},
                            {"data[0][name]", "sEcho"},
                            {"data[0][value]", _fy_secho++.ToString()},
                            {"data[1][name]", "iColumns"},
                            {"data[1][value]", "18"},
                            {"data[2][name]", "sColumns"},
                            {"data[2][value]", ""},
                            {"data[3][name]", "iDisplayStart"},
                            {"data[3][value]", "0"},
                            {"data[4][name]", "iDisplayLength"},
                            {"data[4][value]", _display_length_fy.ToString()}
                        };

                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: getting data... M-WITHDRAWAL RECORD";

                            result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptCorrectionAjax", "POST", reqparm);
                            responsebody = Encoding.UTF8.GetString(result);
                        }
                    }
                }
                else if (selected_index == 1)
                {
                    if (!_isSecondRequest_fy)
                    {
                        // Manual Bonus Report
                        var reqparm = new NameValueCollection
                        {
                            {"s_btype", ""},
                            {"ptype", "1411"},
                            {"fs_ptype", "1411"},
                            {"s_StartTime", gettotal_start_datetime},
                            {"s_EndTime", gettotal_end_datetime},
                            {"s_type", "1"},
                            {"s_keyword", "0"},
                            {"s_playercurrency", "ALL"},
                            {"wttype", "0"},
                            {"data[0][name]", "sEcho"},
                            {"data[0][value]", _fy_secho++.ToString()},
                            {"data[1][name]", "iColumns"},
                            {"data[1][value]", "18"},
                            {"data[2][name]", "sColumns"},
                            {"data[2][value]", ""},
                            {"data[3][name]", "iDisplayStart"},
                            {"data[3][value]", "0"},
                            {"data[4][name]", "iDisplayLength"},
                            {"data[4][value]", _display_length_fy.ToString()}
                        };

                        // status
                        label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                        label_fy_status.Text = "status: getting data... M-BONUS RECORD";

                        result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptCorrectionAjax", "POST", reqparm);
                        responsebody = Encoding.UTF8.GetString(result);
                    }
                    else
                    {
                        // Generated Bonus Report
                        DateTime start_datetime_replace = DateTime.ParseExact(gettotal_start_datetime, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        var reqparm = new NameValueCollection
                        {
                            {"s_btype", ""},
                            {"skip", "0"},
                            {"s_StartTime", start_datetime_replace.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss")},
                            {"s_EndTime", gettotal_end_datetime},
                            {"s_type", "0"},
                            {"s_keyword", "0"},
                            {"data[0][name]", "sEcho"},
                            {"data[0][value]", _fy_secho++.ToString()},
                            {"data[1][name]", "iColumns"},
                            {"data[1][value]", "18"},
                            {"data[2][name]", "sColumns"},
                            {"data[2][value]", ""},
                            {"data[3][name]", "iDisplayStart"},
                            {"data[3][value]", "0"},
                            {"data[4][name]", "iDisplayLength"},
                            {"data[4][value]", _display_length_fy.ToString()}
                        };

                        // status
                        label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                        label_fy_status.Text = "status: getting data... G-BONUS RECORD";

                        result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/getRakeBackHistory", "POST", reqparm);
                        responsebody = Encoding.UTF8.GetString(result);
                    }
                }
                else if (selected_index == 2)
                {
                    // Bet Record
                    var reqparm = new NameValueCollection
                    {
                        {"s_btype", ""},
                        {"betNo", ""},
                        {"name", ""},
                        {"gpid", "0"},
                        {"wager_settle", "0"},
                        {"valid_inva", ""},
                        {"start", gettotal_start_datetime},
                        {"end", gettotal_end_datetime},
                        {"skip", "0"},
                        {"ftime_188", "bettime"},
                        {"data[0][name]", "sEcho"},
                        {"data[0][value]", _fy_secho++.ToString()},
                        {"data[1][name]", "iColumns"},
                        {"data[1][value]", "12"},
                        {"data[2][name]", "sColumns"},
                        {"data[2][value]", ""},
                        {"data[3][name]", "iDisplayStart"},
                        {"data[3][value]", "0"},
                        {"data[4][name]", "iDisplayLength"},
                        {"data[4][value]", _display_length_fy.ToString()}
                    };

                    label_fy_status.Text = "status: getting data... BET RECORD";

                    result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm);
                    responsebody = Encoding.UTF8.GetString(result);
                }

                var deserializeObject = JsonConvert.DeserializeObject(responsebody);

                jo_fy = JObject.Parse(deserializeObject.ToString());
                JToken count = jo_fy.SelectToken("$.aaData");
                _result_count_json_fy = count.Count();
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    detect_fy++;
                    await GetDataFYAsync();
                }
            }
        }

        private async Task GetDataFYPagesAsync()
        {
            try
            {
                string gettotal_start_datetime = "";
                string gettotal_end_datetime = "";

                var last_item = fy_gettotal[fy_gettotal.Count - 1];

                if (last_item != _fy_pages_count_display.ToString())
                {

                    foreach (var gettotal in fy_gettotal)
                    {
                        if (gettotal == _fy_pages_count_display.ToString())
                        {
                            _fy_current_index++;
                            _fy_pages_count_last = _fy_pages_count;
                            _fy_pages_count = 0;
                            _detect_fy = true;
                            break;
                        }
                    }

                    int i = 0;
                    foreach (var datetime in fy_datetime)
                    {
                        i++;
                        string[] datetime_results = datetime.Split("*|*");
                        int ii = 0;

                        foreach (string datetime_result in datetime_results)
                        {
                            ii++;
                            if (i == _fy_current_index)
                            {
                                if (ii == 1)
                                {
                                    gettotal_start_datetime = datetime_result;
                                }
                                else if (ii == 2)
                                {
                                    gettotal_end_datetime = datetime_result;

                                    break;
                                }
                            }
                        }
                    }

                    // asd label1.Text = gettotal_start_datetime + " ----- dsadsadas " + gettotal_end_datetime;

                    var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
                    WebClient wc = new WebClient();

                    wc.Headers.Add("Cookie", cookie);
                    wc.Encoding = Encoding.UTF8;
                    wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                    int result_pages;

                    if (_detect_fy)
                    {
                        _detect_fy = false;
                        result_pages = (Convert.ToInt32(_display_length_fy) * _fy_pages_count);
                    }
                    else
                    {
                        _fy_pages_count++;
                        result_pages = (Convert.ToInt32(_display_length_fy) * _fy_pages_count);
                    }

                    byte[] result = null;
                    string responsebody = null;

                    int selected_index = comboBox_fy_list.SelectedIndex;
                    if (selected_index == 0)
                    {
                        if (!_isSecondRequestFinish_fy)
                        {
                            if (!_isSecondRequest_fy)
                            {
                                // Deposit Record
                                var reqparm = new NameValueCollection
                                {
                                    {"s_btype", ""},
                                    {"s_StartTime", gettotal_start_datetime},
                                    {"s_EndTime", gettotal_end_datetime},
                                    {"dno", ""},
                                    {"s_dpttype", "0"},
                                    {"s_type", "1"},
                                    {"s_transtype", "0"},
                                    {"s_ppid", "0"},
                                    {"s_payoption", "0"},
                                    {"groupid", "0"},
                                    {"s_keyword", ""},
                                    {"s_playercurrency", "ALL"},
                                    {"skip", "0"},
                                    {"data[0][name]", "sEcho"},
                                    {"data[0][value]", _fy_secho++.ToString()},
                                    {"data[1][name]", "iColumns"},
                                    {"data[1][value]", "17"},
                                    {"data[2][name]", "sColumns"},
                                    {"data[2][value]", ""},
                                    {"data[3][name]", "iDisplayStart"},
                                    {"data[3][value]", result_pages.ToString()},
                                    {"data[4][name]", "iDisplayLength"},
                                    {"data[4][value]", _display_length_fy.ToString()}
                                };

                                label_fy_status.Text = "status: getting data... DEPOSIT RECORD";

                                result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptHistoryAjax", "POST", reqparm);
                                if (get_value == "en")
                                {
                                    responsebody = Encoding.UTF8.GetString(result).Remove(0, 1);
                                }
                                else
                                {
                                    responsebody = Encoding.UTF8.GetString(result);
                                }
                            }
                            else
                            {
                                // Manual Deposit Record
                                var reqparm = new NameValueCollection
                                {
                                    {"s_btype", ""},
                                    {"ptype", "1212"},
                                    {"fs_ptype", "1212"},
                                    {"s_StartTime", gettotal_start_datetime},
                                    {"s_EndTime", gettotal_end_datetime},
                                    {"s_type", "1"},
                                    {"s_keyword", ""},
                                    {"data[0][name]", "sEcho"},
                                    {"data[0][value]", _fy_secho++.ToString()},
                                    {"data[1][name]", "iColumns"},
                                    {"data[1][value]", "18"},
                                    {"data[2][name]", "sColumns"},
                                    {"data[2][value]", ""},
                                    {"data[3][name]", "iDisplayStart"},
                                    {"data[3][value]", "0"},
                                    {"data[4][name]", "iDisplayLength"},
                                    {"data[4][value]", _display_length_fy.ToString()}
                                };

                                // status
                                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                                label_fy_status.Text = "status: getting data... M-DEPOSIT RECORD";

                                result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptCorrectionAjax", "POST", reqparm);
                                responsebody = Encoding.UTF8.GetString(result);
                            }
                        }
                        else
                        {
                            if (!_isThirdRequest_fy)
                            {
                                // Withdrawal Record
                                var reqparm = new NameValueCollection
                                {
                                    {"s_btype", ""},
                                    {"s_StartTime", gettotal_start_datetime},
                                    {"s_EndTime", gettotal_end_datetime},
                                    {"s_wtdAmtFr", ""},
                                    {"s_wtdAmtTo", ""},
                                    {"s_dpttype", "0"},
                                    {"skip", "0"},
                                    {"s_type", "1"},
                                    {"s_keyword", "0"},
                                    {"s_playercurrency", "ALL"},
                                    {"wttype", "0"},
                                    {"data[0][name]", "sEcho"},
                                    {"data[0][value]", _fy_secho++.ToString()},
                                    {"data[1][name]", "iColumns"},
                                    {"data[1][value]", "18"},
                                    {"data[2][name]", "sColumns"},
                                    {"data[2][value]", ""},
                                    {"data[3][name]", "iDisplayStart"},
                                    {"data[3][value]", result_pages.ToString()},
                                    {"data[4][name]", "iDisplayLength"},
                                    {"data[4][value]", _display_length_fy.ToString()}
                                };

                                // status
                                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                                label_fy_status.Text = "status: getting data... WITHDRAWAL RECORD";

                                result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/wtdHistoryAjax", "POST", reqparm);
                                responsebody = Encoding.UTF8.GetString(result);
                            }
                            else
                            {
                                // Manual Withdrawal Record
                                var reqparm = new NameValueCollection
                            {
                                {"s_btype", ""},
                                {"ptype", "1313"},
                                {"fs_ptype", "1313"},
                                {"s_StartTime", gettotal_start_datetime},
                                {"s_EndTime", gettotal_end_datetime},
                                {"s_type", "1"},
                                {"s_keyword", "0"},
                                {"s_playercurrency", "ALL"},
                                {"wttype", "0"},
                                {"data[0][name]", "sEcho"},
                                {"data[0][value]", _fy_secho++.ToString()},
                                {"data[1][name]", "iColumns"},
                                {"data[1][value]", "18"},
                                {"data[2][name]", "sColumns"},
                                {"data[2][value]", ""},
                                {"data[3][name]", "iDisplayStart"},
                                {"data[3][value]", "0"},
                                {"data[4][name]", "iDisplayLength"},
                                {"data[4][value]", _display_length_fy.ToString()}
                            };

                                // status
                                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                                label_fy_status.Text = "status: getting data... M-WITHDRAWAL RECORD";

                                result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptCorrectionAjax", "POST", reqparm);
                                responsebody = Encoding.UTF8.GetString(result);
                            }
                        }
                    }
                    else if (selected_index == 1)
                    {
                        if (!_isSecondRequest_fy)
                        {
                            // Manual Bonus Report
                            var reqparm = new NameValueCollection
                            {
                                {"s_btype", ""},
                                {"ptype", "1411"},
                                {"fs_ptype", "1411"},
                                {"s_StartTime", gettotal_start_datetime},
                                {"s_EndTime", gettotal_end_datetime},
                                {"s_type", "1"},
                                {"s_keyword", "0"},
                                {"s_playercurrency", "ALL"},
                                {"wttype", "0"},
                                {"data[0][name]", "sEcho"},
                                {"data[0][value]", _fy_secho++.ToString()},
                                {"data[1][name]", "iColumns"},
                                {"data[1][value]", "18"},
                                {"data[2][name]", "sColumns"},
                                {"data[2][value]", ""},
                                {"data[3][name]", "iDisplayStart"},
                                {"data[3][value]", result_pages.ToString()},
                                {"data[4][name]", "iDisplayLength"},
                                {"data[4][value]", _display_length_fy.ToString()}
                            };

                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: getting data... M-BONUS RECORD";

                            result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/playerFund/dptCorrectionAjax", "POST", reqparm);
                            responsebody = Encoding.UTF8.GetString(result);
                        }
                        else
                        {
                            // Generated Bonus Report
                            DateTime start_datetime_replace = DateTime.ParseExact(gettotal_start_datetime, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                            var reqparm = new NameValueCollection
                            {
                                {"s_btype", ""},
                                {"skip", "0"},
                                {"s_StartTime", start_datetime_replace.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss")},
                                {"s_EndTime", gettotal_end_datetime},
                                {"s_type", "0"},
                                {"s_keyword", "0"},
                                {"data[0][name]", "sEcho"},
                                {"data[0][value]", _fy_secho++.ToString()},
                                {"data[1][name]", "iColumns"},
                                {"data[1][value]", "18"},
                                {"data[2][name]", "sColumns"},
                                {"data[2][value]", ""},
                                {"data[3][name]", "iDisplayStart"},
                                {"data[3][value]", "0"},
                                {"data[4][name]", "iDisplayLength"},
                                {"data[4][value]", _display_length_fy.ToString()}
                            };

                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: getting data... G-BONUS RECORD";

                            result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/getRakeBackHistory", "POST", reqparm);
                            responsebody = Encoding.UTF8.GetString(result);
                        }
                    }
                    else if (selected_index == 2)
                    {
                        // Bet Record
                        var reqparm = new NameValueCollection
                        {
                            {"s_btype", ""},
                            {"betNo", ""},
                            {"name", ""},
                            {"gpid", "0"},
                            {"wager_settle", "0"},
                            {"valid_inva", ""},
                            {"start", gettotal_start_datetime},
                            {"end", gettotal_end_datetime},
                            {"skip", "0"},
                            {"ftime_188", "bettime"},
                            {"data[0][name]", "sEcho"},
                            {"data[0][value]", _fy_secho++.ToString()},
                            {"data[1][name]", "iColumns"},
                            {"data[1][value]", "12"},
                            {"data[2][name]", "sColumns"},
                            {"data[2][value]", ""},
                            {"data[3][name]", "iDisplayStart"},
                            {"data[3][value]", result_pages.ToString()},
                            {"data[4][name]", "iDisplayLength"},
                            {"data[4][value]", _display_length_fy.ToString()}
                        };

                        label_fy_status.Text = "status: getting data... BET RECORD";

                        result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm);
                        responsebody = Encoding.UTF8.GetString(result);
                    }

                    var deserializeObject = JsonConvert.DeserializeObject(responsebody);

                    jo_fy = JObject.Parse(deserializeObject.ToString());
                    JToken count = jo_fy.SelectToken("$.aaData");
                    _result_count_json_fy = count.Count();
                }
                else
                {
                    foreach (var gettotal in fy_gettotal)
                    {
                        if (gettotal == _fy_pages_count_display.ToString())
                        {
                            _fy_current_index++;
                            _fy_pages_count_last = _fy_pages_count;
                            _fy_pages_count = 0;
                            _detect_fy = true;

                            break;
                        }
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    if (!_detect_fy)
                    {
                        _fy_pages_count--;
                    }

                    detect_fy++;

                    await GetDataFYPagesAsync();
                }
            }
        }

        private async void FYAsync()
        {
            if (_fy_inserted_in_excel)
            {
                for (int i = _fy_i; i < _total_page_fy; i++)
                {
                    button_fy_start.Visible = false;

                    if (!_fy_inserted_in_excel)
                    {
                        break;
                    }
                    else
                    {
                        _fy_i = i;
                        _fy_pages_count_display++;
                    }

                    for (int ii = 0; ii < _result_count_json_fy; ii++)
                    {
                        Application.DoEvents();

                        if (!__isLogin)
                        {
                            _fy_inserted_in_excel = false;

                            break;
                        }

                        _test_fy_gettotal_count_record++;

                        if (_fy_pages_count_display != 0 && _fy_pages_count_display <= _total_page_fy)
                        {
                            label_fy_page_count.Text = _fy_pages_count_display.ToString("N0") + " of " + _total_page_fy.ToString("N0");
                        }

                        _fy_ii = ii;

                        int selected_index = comboBox_fy_list.SelectedIndex;
                        if (selected_index == 0)
                        {
                            if (!_isSecondRequestFinish_fy)
                            {
                                if (!_isSecondRequest_fy)
                                {
                                    // Deposit Record
                                    JToken submitted_date__transaction_id = jo_fy.SelectToken("$.aaData[" + ii + "][0]");
                                    string submitted_date = submitted_date__transaction_id.ToString().Substring(0, 19);
                                    string transaction_id = submitted_date__transaction_id.ToString().Substring(23);
                                    string date = submitted_date__transaction_id.ToString().Substring(0, 10);
                                    JToken member_get = jo_fy.SelectToken("$.aaData[" + ii + "][1]");
                                    string member = Regex.Match(member_get.ToString(), "<span(.*?)>(.*?)</span>").Groups[2].Value;
                                    JToken vip = jo_fy.SelectToken("$.aaData[" + ii + "][3]").ToString().Replace("\"", "");
                                    JToken amount = jo_fy.SelectToken("$.aaData[" + ii + "][5]").ToString().Replace("\"", "");
                                    JToken payment_type = jo_fy.SelectToken("$.aaData[" + ii + "][11]");
                                    string[] payment_types_replace = payment_type.ToString().Split("<br>");
                                    payment_type = "";
                                    int count_ = 0;
                                    foreach (string payment_type_replace in payment_types_replace)
                                    {
                                        count_++;
                                        if (count_ == 1)
                                        {
                                            payment_type += payment_type_replace + "-";
                                        }
                                        else if (count_ == 2)
                                        {
                                            payment_type += payment_type_replace;
                                            break;
                                        }
                                    }
                                    if (payment_type.ToString() != "")
                                    {
                                        if (payment_type.ToString().Substring(payment_type.ToString().Length - 1, 1) == "-")
                                        {
                                            payment_type = payment_type.ToString().Remove(payment_type.ToString().Length - 1);
                                        }
                                    }
                                    JToken status_get = jo_fy.SelectToken("$.aaData[" + ii + "][12]").ToString().Replace("\"", "");
                                    string status = Regex.Match(status_get.ToString(), "<font(.*?)>(.*?)</font>").Groups[2].Value;
                                    string replace_status = status.ToLower();
                                    if (replace_status == "success" || replace_status == "成功")
                                    {
                                        status = "Success";
                                    }
                                    else if (replace_status == "failure" || replace_status == "失败")
                                    {
                                        status = "Failure";
                                    }
                                    else if (replace_status == "取消")
                                    {
                                        status = "Cancelled";
                                    }
                                    replace_status = status.ToLower();
                                    JToken updated_date__updated_time = jo_fy.SelectToken("$.aaData[" + ii + "][13]").ToString().Replace("\"", "");
                                    string updated_date = updated_date__updated_time.ToString().Substring(0, 10);
                                    string updated_time = updated_date__updated_time.ToString().Substring(15);
                                    DateTime month = DateTime.ParseExact(submitted_date, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

                                    // Bank account
                                    string pg_company = "";
                                    string pg_type = "";
                                    string bank_account_fy_temp = Path.Combine(Path.GetTempPath(), "FY Payment Type Code.txt");
                                    
                                    if (payment_type.ToString().Trim() == "手工存款")
                                    {
                                        string replace_transaction_id = transaction_id.ToLower();
                                        if (replace_transaction_id.Contains("wechat"))
                                        {
                                            pg_company = "MANUAL WECHAT";
                                            pg_type = "MANUAL WECHAT";
                                            break;
                                        }
                                        else
                                        {
                                            pg_company = "LOCAL BANK";
                                            pg_type = "LOCAL BANK";
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        using (StreamReader sr = File.OpenText(bank_account_fy_temp))
                                        {
                                            string s = String.Empty;
                                            while ((s = sr.ReadLine()) != null)
                                            {
                                                Application.DoEvents();

                                                string[] results = s.Split("*|*");
                                                int bank_account_i = 0;
                                                bool isNext = false;
                                                foreach (string result in results)
                                                {
                                                    Application.DoEvents();

                                                    bank_account_i++;

                                                    if (bank_account_i == 1)
                                                    {
                                                        if (result == payment_type.ToString().Trim())
                                                        {
                                                            isNext = true;
                                                        }
                                                    }

                                                    if (isNext)
                                                    {
                                                        if (bank_account_i == 2)
                                                        {
                                                            pg_company = result;
                                                        }
                                                        else if (bank_account_i == 3)
                                                        {
                                                            pg_type = result;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (payment_type.ToString() == "")
                                    {
                                        pg_company = "";
                                        pg_type = "";
                                    }

                                    string duration_time = "";
                                    DateTime start = DateTime.ParseExact(submitted_date, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    DateTime end = DateTime.ParseExact(updated_date + " " + updated_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    TimeSpan span = end - start;
                                    double totalMinutes = Math.Floor(span.TotalMinutes);

                                    if (totalMinutes <= 5)
                                    {
                                        // 0-5
                                        duration_time = "0-5min";
                                    }
                                    else if (totalMinutes <= 10)
                                    {
                                        // 6-10
                                        duration_time = "6-10min";
                                    }
                                    else if (totalMinutes <= 15)
                                    {
                                        // 11-15
                                        duration_time = "11-15min";
                                    }
                                    else if (totalMinutes <= 20)
                                    {
                                        // 16-20
                                        duration_time = "16-20min";
                                    }
                                    else if (totalMinutes <= 25)
                                    {
                                        // 21-25
                                        duration_time = "21-25min";
                                    }
                                    else if (totalMinutes <= 30)
                                    {
                                        // 26-30
                                        duration_time = "26-30min";
                                    }
                                    else if (totalMinutes <= 60)
                                    {
                                        // 31-60
                                        duration_time = "31-60min";
                                    }
                                    else if (totalMinutes >= 61)
                                    {
                                        // >60
                                        duration_time = ">60min";
                                    }
                                    
                                    string retained = "";
                                    string fd_date = "";
                                    string new_fy = "";
                                    string reactivated = "";
                                    string last_deposit_get_replace = "";
                                    string first_deposit_get_replace = "";
                                    string first_deposit_get = "";
                                    
                                    if (replace_status == "success" && !member.ToString().ToLower().Contains("test") && !vip.ToString().ToLower().Contains("test"))
                                    {
                                        // get last deposit in temp file
                                        string memberlist_temp = Path.Combine(Path.GetTempPath(), "FY Registration Deposit.txt");
                                        if (File.Exists(memberlist_temp))
                                        {
                                            using (StreamReader sr = File.OpenText(memberlist_temp))
                                            {
                                                string s = String.Empty;
                                                while ((s = sr.ReadLine()) != null)
                                                {
                                                    int memberlist_i = 0;
                                                    string[] results = s.Split("*|*");
                                                    foreach (string result in results)
                                                    {
                                                        Application.DoEvents();

                                                        memberlist_i++;

                                                        if (memberlist_i == 1)
                                                        {
                                                            // Username
                                                            if (result.Trim().ToLower() == member.ToString().Trim().ToLower())
                                                            {
                                                                int memberlist_i_inner = 0;
                                                                string[] results_inner = s.Split("*|*");
                                                                foreach (string result_inner in results_inner)
                                                                {
                                                                    Application.DoEvents();

                                                                    memberlist_i_inner++;
                                                                    if (memberlist_i_inner == 2)
                                                                    {
                                                                        string[] first_deposit_get_results = result_inner.Split("/");
                                                                        int count = 0;
                                                                        foreach (string first_deposit_get_result in first_deposit_get_results)
                                                                        {
                                                                            Application.DoEvents();

                                                                            count++;

                                                                            if (count == 1)
                                                                            {
                                                                                // Month
                                                                                if (first_deposit_get_result.Length == 1)
                                                                                {
                                                                                    first_deposit_get += "0" + first_deposit_get_result + "/";
                                                                                }
                                                                                else
                                                                                {
                                                                                    first_deposit_get += first_deposit_get_result + "/";
                                                                                }
                                                                            }
                                                                            else if (count == 2)
                                                                            {
                                                                                // Day
                                                                                if (first_deposit_get_result.Length == 1)
                                                                                {
                                                                                    first_deposit_get += "0" + first_deposit_get_result + "/";
                                                                                }
                                                                                else
                                                                                {
                                                                                    first_deposit_get += first_deposit_get_result + "/";
                                                                                }
                                                                            }
                                                                            else if (count == 3)
                                                                            {
                                                                                // Year
                                                                                first_deposit_get += first_deposit_get_result.Substring(0, 4);
                                                                            }

                                                                            first_deposit_get_replace = first_deposit_get;
                                                                        }
                                                                    }
                                                                    
                                                                    if (memberlist_i_inner == 3)
                                                                    {
                                                                        string[] first_deposit_get_results = result_inner.Split("/");
                                                                        int count = 0;
                                                                        foreach (string first_deposit_get_result in first_deposit_get_results)
                                                                        {
                                                                            Application.DoEvents();

                                                                            count++;

                                                                            if (count == 1)
                                                                            {
                                                                                // Month
                                                                                if (first_deposit_get_result.Length == 1)
                                                                                {
                                                                                    last_deposit_get_replace += "0" + first_deposit_get_result + "/";
                                                                                }
                                                                                else
                                                                                {
                                                                                    last_deposit_get_replace += first_deposit_get_result + "/";
                                                                                }
                                                                            }
                                                                            else if (count == 2)
                                                                            {
                                                                                // Day
                                                                                if (first_deposit_get_result.Length == 1)
                                                                                {
                                                                                    last_deposit_get_replace += "0" + first_deposit_get_result + "/";
                                                                                }
                                                                                else
                                                                                {
                                                                                    last_deposit_get_replace += first_deposit_get_result + "/";
                                                                                }
                                                                            }
                                                                            else if (count == 3)
                                                                            {
                                                                                // Year
                                                                                last_deposit_get_replace += first_deposit_get_result.Substring(0, 4);
                                                                            }
                                                                        }
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (first_deposit_get_replace != "" && last_deposit_get_replace != "")
                                        {
                                            // get last deposit in temp file
                                            try
                                            {
                                                DateTime last_deposit = DateTime.ParseExact(last_deposit_get_replace.Trim().Substring(0, 10), "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                                DateTime first_deposit = DateTime.ParseExact(first_deposit_get_replace.Trim().Substring(0, 10), "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                                DateTime first_deposit_ = DateTime.ParseExact(first_deposit_get.Trim().Substring(0, 10), "MM/dd/yyyy", CultureInfo.InvariantCulture);

                                                // retained
                                                // 2 months current date
                                                double amount_get = Convert.ToDouble(amount);
                                                if (amount_get > 0)
                                                {
                                                    var last2month_get = DateTime.Today.AddMonths(-2);
                                                    DateTime last2month = DateTime.ParseExact(last2month_get.ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                    if (last_deposit >= last2month)
                                                    {
                                                        retained = "Retained";
                                                    }
                                                    else
                                                    {
                                                        retained = "Not Retained";
                                                    }
                                                }
                                                else
                                                {
                                                    retained = "Not Retained";
                                                }

                                                String month_get = DateTime.Now.Month.ToString();
                                                String year_get = DateTime.Now.Year.ToString();
                                                string year_month = year_get + "-" + month_get;

                                                // new
                                                if (first_deposit.ToString("yyyy-MM") == year_month)
                                                {
                                                    new_fy = "New";
                                                }
                                                else
                                                {
                                                    new_fy = "Not New";
                                                }

                                                // reactivated
                                                if (retained == "Not Retained" && new_fy == "Not New")
                                                {
                                                    reactivated = "Reactivated";
                                                }
                                                else
                                                {
                                                    reactivated = "Not Reactivated";
                                                }

                                                fd_date = first_deposit_.ToString("MM/dd/yyyy");
                                            }
                                            catch (Exception err)
                                            {
                                                MessageBox.Show(first_deposit_get);
                                                MessageBox.Show(first_deposit_get_replace);
                                                MessageBox.Show(last_deposit_get_replace);
                                                MessageBox.Show(err.ToString());
                                            }
                                        }
                                        else
                                        {
                                            retained = "";
                                            fd_date = "";
                                            new_fy = "";
                                            reactivated = "";
                                        }
                                    }
                                    else
                                    {
                                        retained = "";
                                        fd_date = "";
                                        new_fy = "";
                                        reactivated = "";
                                    }

                                    if (_fy_get_ii == 1)
                                    {
                                        var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19}", "Brand", "Month", "Date", "Submitted Date", "Updated Date", "Member", "Payment Type", "PG Company", "PG Type", "Transaction ID", "Amount", "Transaction Time", "Transaction Type", "Duration Time", "VIP", "Status", "Retained", "FD Date", "New", "Reactivated");
                                        _fy_csv.AppendLine(header);
                                    }

                                    String date_get = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                                    DateTime submitted_date_ = DateTime.ParseExact(submitted_date, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    DateTime updated_date_ = DateTime.ParseExact(updated_date + " " + updated_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    TimeSpan transaction_time = updated_date_ - submitted_date_;
                                    string transaction_time_get = "";
                                    if (transaction_time.Hours.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Hours + ":";
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Hours + ":";
                                    }
                                    if (transaction_time.Minutes.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Minutes + ":";
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Minutes + ":";
                                    }
                                    if (transaction_time.Seconds.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Seconds;
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Seconds;
                                    }

                                    var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19}", "FY", "\"" + month.ToString("MM/01/yyyy") + "\"", "\"" + date_get + "\"", "\"" + submitted_date + "\"", "\"" + updated_date + " " + updated_time + "\"", "\"" + member + "\"", "\"" + payment_type + "\"", "\"" + pg_company + "\"", "\"" + pg_type + "\"", "\"" + "'" + transaction_id.Replace(",", "") + "\"", "\"" + amount.ToString().Replace("-", "") + "\"", "\"" + transaction_time_get + "\"", "\"" + "Deposit" + "\"", "\"" + duration_time + "\"", "\"" + vip + "\"", "\"" + status + "\"", "\"" + retained + "\"", "\"" + fd_date + "\"", "\"" + new_fy + "\"", "\"" + reactivated + "\"");
                                    _fy_csv.AppendLine(newLine);
                                }
                                else
                                {
                                    // Manual Deposit Record
                                    JToken member_get = jo_fy.SelectToken("$.aaData[" + ii + "][1]");
                                    string member = Regex.Match(member_get.ToString(), "<span(.*?)>(.*?)</span>").Groups[2].Value;
                                    JToken vip = jo_fy.SelectToken("$.aaData[" + ii + "][3]").ToString().Replace("\"", "");
                                    JToken amount = jo_fy.SelectToken("$.aaData[" + ii + "][5]").ToString().Replace("(RMB) - ¥ ", "");
                                    JToken remark = jo_fy.SelectToken("$.aaData[" + ii + "][8]").ToString().Replace("\"", "");
                                    JToken submitted_date__submitted_time = jo_fy.SelectToken("$.aaData[" + ii + "][10]");
                                    string submitted_date = submitted_date__submitted_time.ToString().Substring(0, 10);
                                    string submitted_time = submitted_date__submitted_time.ToString().Substring(15);
                                    JToken payment_type = jo_fy.SelectToken("$.aaData[" + ii + "][7]").ToString().Replace("\"", "");
                                    if (payment_type.ToString() != "")
                                    {
                                        if (payment_type.ToString().Substring(payment_type.ToString().Length - 1, 1) == "-")
                                        {
                                            payment_type = payment_type.ToString().Remove(payment_type.ToString().Length - 1);
                                        }
                                    }
                                    DateTime month = DateTime.ParseExact(submitted_date, "yyyy-MM-dd", CultureInfo.InvariantCulture);

                                    // Bank account
                                    string pg_company = "";
                                    string pg_type = "";
                                    string bank_account_fy_temp = Path.Combine(Path.GetTempPath(), "FY Payment Type Code.txt");

                                    using (StreamReader sr = File.OpenText(bank_account_fy_temp))
                                    {
                                        string s = String.Empty;
                                        while ((s = sr.ReadLine()) != null)
                                        {
                                            Application.DoEvents();

                                            string[] results = s.Split("*|*");
                                            int bank_account_i = 0;
                                            bool isNext = false;
                                            foreach (string result in results)
                                            {
                                                Application.DoEvents();

                                                bank_account_i++;

                                                if (bank_account_i == 1)
                                                {
                                                    if (result == "手工存款")
                                                    {
                                                        string replace_transaction_id = remark.ToString().ToLower();
                                                        if (replace_transaction_id.Contains("wechat"))
                                                        {
                                                            pg_company = "MANUAL WECHAT";
                                                            pg_type = "MANUAL WECHAT";
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            pg_company = "LOCAL BANK";
                                                            pg_type = "LOCAL BANK";
                                                            break;
                                                        }
                                                    }
                                                    else if (result == payment_type.ToString().Trim())
                                                    {
                                                        isNext = true;
                                                    }
                                                }

                                                if (isNext)
                                                {
                                                    if (bank_account_i == 2)
                                                    {
                                                        pg_company = result;
                                                    }
                                                    else if (bank_account_i == 3)
                                                    {
                                                        pg_type = result;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (payment_type.ToString() == "")
                                    {
                                        pg_company = "";
                                        pg_type = "";
                                    }

                                    string duration_time = "";
                                    DateTime start = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    DateTime end = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    TimeSpan span = end - start;
                                    double totalMinutes = Math.Floor(span.TotalMinutes);

                                    if (totalMinutes <= 5)
                                    {
                                        // 0-5
                                        duration_time = "0-5min";
                                    }
                                    else if (totalMinutes <= 10)
                                    {
                                        // 6-10
                                        duration_time = "6-10min";
                                    }
                                    else if (totalMinutes <= 15)
                                    {
                                        // 11-15
                                        duration_time = "11-15min";
                                    }
                                    else if (totalMinutes <= 20)
                                    {
                                        // 16-20
                                        duration_time = "16-20min";
                                    }
                                    else if (totalMinutes <= 25)
                                    {
                                        // 21-25
                                        duration_time = "21-25min";
                                    }
                                    else if (totalMinutes <= 30)
                                    {
                                        // 26-30
                                        duration_time = "26-30min";
                                    }
                                    else if (totalMinutes <= 60)
                                    {
                                        // 31-60
                                        duration_time = "31-60min";
                                    }
                                    else if (totalMinutes >= 61)
                                    {
                                        // >60
                                        duration_time = ">60min";
                                    }
                                    
                                    string retained = "";
                                    string fd_date = "";
                                    string new_fy = "";
                                    string reactivated = "";
                                    string last_deposit_get_replace = "";
                                    string first_deposit_get_replace = "";
                                    string first_deposit_get = "";

                                    if (!member.ToString().ToLower().Contains("test") && !vip.ToString().ToLower().Contains("test"))
                                    {
                                        // get last deposit in temp file
                                        string memberlist_temp = Path.Combine(Path.GetTempPath(), "FY Registration Deposit.txt");
                                        if (File.Exists(memberlist_temp))
                                        {
                                            using (StreamReader sr = File.OpenText(memberlist_temp))
                                            {
                                                string s = String.Empty;
                                                while ((s = sr.ReadLine()) != null)
                                                {
                                                    int memberlist_i = 0;
                                                    string[] results = s.Split("*|*");
                                                    foreach (string result in results)
                                                    {
                                                        Application.DoEvents();

                                                        memberlist_i++;

                                                        if (memberlist_i == 1)
                                                        {
                                                            // Username
                                                            if (result.Trim().ToLower() == member.ToString().Trim().ToLower())
                                                            {
                                                                int memberlist_i_inner = 0;
                                                                string[] results_inner = s.Split("*|*");
                                                                foreach (string result_inner in results_inner)
                                                                {
                                                                    Application.DoEvents();

                                                                    memberlist_i_inner++;
                                                                    if (memberlist_i_inner == 2)
                                                                    {
                                                                        string[] first_deposit_get_results = result_inner.Split("/");
                                                                        int count = 0;
                                                                        foreach (string first_deposit_get_result in first_deposit_get_results)
                                                                        {
                                                                            Application.DoEvents();

                                                                            count++;

                                                                            if (count == 1)
                                                                            {
                                                                                // Month
                                                                                if (first_deposit_get_result.Length == 1)
                                                                                {
                                                                                    first_deposit_get += "0" + first_deposit_get_result + "/";
                                                                                }
                                                                                else
                                                                                {
                                                                                    first_deposit_get += first_deposit_get_result + "/";
                                                                                }
                                                                            }
                                                                            else if (count == 2)
                                                                            {
                                                                                // Day
                                                                                if (first_deposit_get_result.Length == 1)
                                                                                {
                                                                                    first_deposit_get += "0" + first_deposit_get_result + "/";
                                                                                }
                                                                                else
                                                                                {
                                                                                    first_deposit_get += first_deposit_get_result + "/";
                                                                                }
                                                                            }
                                                                            else if (count == 3)
                                                                            {
                                                                                // Year
                                                                                first_deposit_get += first_deposit_get_result.Substring(0, 4);
                                                                            }

                                                                            first_deposit_get_replace = first_deposit_get;
                                                                        }
                                                                    }
                                                                    
                                                                    if (memberlist_i_inner == 3)
                                                                    {
                                                                        string[] first_deposit_get_results = result_inner.Split("/");
                                                                        int count = 0;
                                                                        foreach (string first_deposit_get_result in first_deposit_get_results)
                                                                        {
                                                                            Application.DoEvents();

                                                                            count++;

                                                                            if (count == 1)
                                                                            {
                                                                                // Month
                                                                                if (first_deposit_get_result.Length == 1)
                                                                                {
                                                                                    last_deposit_get_replace += "0" + first_deposit_get_result + "/";
                                                                                }
                                                                                else
                                                                                {
                                                                                    last_deposit_get_replace += first_deposit_get_result + "/";
                                                                                }
                                                                            }
                                                                            else if (count == 2)
                                                                            {
                                                                                // Day
                                                                                if (first_deposit_get_result.Length == 1)
                                                                                {
                                                                                    last_deposit_get_replace += "0" + first_deposit_get_result + "/";
                                                                                }
                                                                                else
                                                                                {
                                                                                    last_deposit_get_replace += first_deposit_get_result + "/";
                                                                                }
                                                                            }
                                                                            else if (count == 3)
                                                                            {
                                                                                // Year
                                                                                last_deposit_get_replace += first_deposit_get_result.Substring(0, 4);
                                                                            }
                                                                        }
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (first_deposit_get_replace != "" && last_deposit_get_replace != "")
                                        {
                                            // get last deposit in temp file
                                            try
                                            {
                                                DateTime last_deposit = DateTime.ParseExact(last_deposit_get_replace.Trim().Substring(0, 10), "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                                DateTime first_deposit = DateTime.ParseExact(first_deposit_get_replace.Trim().Substring(0, 10), "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                                DateTime first_deposit_ = DateTime.ParseExact(first_deposit_get.Trim().Substring(0, 10), "MM/dd/yyyy", CultureInfo.InvariantCulture);

                                                // retained
                                                // 2 months current date
                                                var last2month_get = DateTime.Today.AddMonths(-2);
                                                DateTime last2month = DateTime.ParseExact(last2month_get.ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                if (last_deposit >= last2month)
                                                {
                                                    retained = "Retained";
                                                }
                                                else
                                                {
                                                    retained = "Not Retained";
                                                }

                                                String month_get = DateTime.Now.Month.ToString();
                                                String year_get = DateTime.Now.Year.ToString();
                                                string year_month = year_get + "-" + month_get;

                                                // new
                                                if (first_deposit.ToString("yyyy-MM") == year_month)
                                                {
                                                    new_fy = "New";
                                                }
                                                else
                                                {
                                                    new_fy = "Not New";
                                                }

                                                // reactivated
                                                if (retained == "Not Retained" && new_fy == "Not New")
                                                {
                                                    reactivated = "Reactivated";
                                                }
                                                else
                                                {
                                                    reactivated = "Not Reactivated";
                                                }

                                                fd_date = first_deposit_.ToString("MM/dd/yyyy");
                                            }
                                            catch (Exception err)
                                            {
                                                MessageBox.Show(first_deposit_get);
                                                MessageBox.Show(first_deposit_get_replace);
                                                MessageBox.Show(last_deposit_get_replace);
                                                MessageBox.Show(err.ToString());
                                            }
                                        }
                                        else
                                        {
                                            retained = "";
                                            fd_date = "";
                                            new_fy = "";
                                            reactivated = "";
                                        }
                                    }
                                    else
                                    {
                                        retained = "";
                                        fd_date = "";
                                        new_fy = "";
                                        reactivated = "";
                                    }

                                    String date_get = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                                    DateTime submitted_date_ = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    DateTime updated_date_ = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    TimeSpan transaction_time = updated_date_ - submitted_date_;
                                    string transaction_time_get = "";
                                    if (transaction_time.Hours.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Hours + ":";
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Hours + ":";
                                    }
                                    if (transaction_time.Minutes.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Minutes + ":";
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Minutes + ":";
                                    }
                                    if (transaction_time.Seconds.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Seconds;
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Seconds;
                                    }

                                    var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19}", "FY", "\"" + month.ToString("MM/01/yyyy") + "\"", "\"" + date_get + "\"", "\"" + submitted_date + " " + submitted_time + "\"", "\"" + submitted_date + " " + submitted_time + "\"", "\"" + member + "\"", "\"" + payment_type + "\"", "\"" + pg_company + "\"", "\"" + pg_type + "\"", "\"" + remark.ToString().Replace(",", "") + "\"", "\"" + amount.ToString().Replace("-", "") + "\"", "\"" + transaction_time_get + "\"", "\"" + "Deposit" + "\"", "\"" + duration_time + "\"", "\"" + vip + "\"", "\"" + "Success" + "\"", "\"" + retained + "\"", "\"" + fd_date + "\"", "\"" + new_fy + "\"", "\"" + reactivated + "\"");
                                    _fy_csv.AppendLine(newLine);
                                }
                            }
                            else
                            {
                                if (!_isThirdRequest_fy)
                                {
                                    // Withdrawal Record
                                    JToken transaction_id = jo_fy.SelectToken("$.aaData[" + ii + "][0]").ToString().Replace("\"", "");
                                    JToken member_get = jo_fy.SelectToken("$.aaData[" + ii + "][1]");
                                    string member = Regex.Match(member_get.ToString(), "<span(.*?)>(.*?)</span>").Groups[2].Value;
                                    JToken vip = jo_fy.SelectToken("$.aaData[" + ii + "][3]").ToString().Replace("\"", "");
                                    JToken amount = jo_fy.SelectToken("$.aaData[" + ii + "][6]").ToString().Replace("\"", "");
                                    JToken submitted_date__submitted_time = jo_fy.SelectToken("$.aaData[" + ii + "][8]").ToString().Replace("\"", "");
                                    string submitted_date = submitted_date__submitted_time.ToString().Substring(0, 10);
                                    string submitted_time = submitted_date__submitted_time.ToString().Substring(15);
                                    string date = submitted_date__submitted_time.ToString().Substring(0, 10);
                                    JToken status = jo_fy.SelectToken("$.aaData[" + ii + "][10]").ToString().Replace("</br>", "");
                                    if (status.ToString() == "出款成功")
                                    {
                                        status = "Success";
                                    }
                                    else
                                    {
                                        status = "Failure";
                                    }
                                    JToken payment_type = jo_fy.SelectToken("$.aaData[" + ii + "][12]").ToString().Replace("\"", "");
                                    string[] payment_type_get_array = payment_type.ToString().Split("<br />");
                                    int i_payment_type = 0;
                                    foreach (string obj in payment_type_get_array)
                                    {
                                        i_payment_type++;

                                        if (i_payment_type == 1)
                                        {
                                            payment_type = obj;
                                            break;
                                        }
                                    }
                                    if (payment_type.ToString() != "")
                                    {
                                        if (payment_type.ToString().Substring(payment_type.ToString().Length - 1, 1) == "-")
                                        {
                                            payment_type = payment_type.ToString().Remove(payment_type.ToString().Length - 1);
                                        }
                                    }
                                    JToken updated_date__updated_time = jo_fy.SelectToken("$.aaData[" + ii + "][13]").ToString().Replace("\"", "");
                                    string updated_date = updated_date__updated_time.ToString().Substring(0, 10);
                                    string updated_time = updated_date__updated_time.ToString().Substring(15);
                                    updated_date__updated_time = updated_date + " " + updated_time;
                                    DateTime month = DateTime.ParseExact(submitted_date, "yyyy-MM-dd", CultureInfo.InvariantCulture);

                                    // Bank account
                                    string pg_company = "";
                                    string pg_type = "";
                                    string bank_account_fy_temp = Path.Combine(Path.GetTempPath(), "FY Payment Type Code.txt");

                                    if (!transaction_id.ToString().ToLower().Contains("wrong"))
                                    {
                                        using (StreamReader sr = File.OpenText(bank_account_fy_temp))
                                        {
                                            string s = String.Empty;
                                            while ((s = sr.ReadLine()) != null)
                                            {
                                                Application.DoEvents();

                                                string[] results = s.Split("*|*");
                                                int bank_account_i = 0;
                                                bool isNext = false;
                                                foreach (string result in results)
                                                {
                                                    Application.DoEvents();

                                                    bank_account_i++;

                                                    if (bank_account_i == 1)
                                                    {
                                                        if (result == "手工存款")
                                                        {
                                                            string replace_transaction_id = transaction_id.ToString().ToLower();
                                                            if (replace_transaction_id.Contains("wechat"))
                                                            {
                                                                pg_company = "MANUAL WECHAT";
                                                                pg_type = "MANUAL WECHAT";
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                pg_company = "LOCAL BANK";
                                                                pg_type = "LOCAL BANK";
                                                                break;
                                                            }
                                                        }
                                                        else if (result == payment_type.ToString().Trim())
                                                        {
                                                            isNext = true;
                                                        }
                                                    }

                                                    if (isNext)
                                                    {
                                                        if (bank_account_i == 2)
                                                        {
                                                            pg_company = result;
                                                        }
                                                        else if (bank_account_i == 3)
                                                        {
                                                            pg_type = result;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (payment_type.ToString() == "")
                                        {
                                            pg_company = "";
                                            pg_type = "";
                                        }
                                    }
                                    else
                                    {
                                        pg_company = "LOCAL BANK ADJUSTMENT";
                                        pg_type = "LOCAL BANK ADJUSTMENT";
                                    }

                                    string duration_time = "";
                                    DateTime start = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    DateTime end = DateTime.ParseExact(updated_date__updated_time.ToString(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    TimeSpan span = end - start;
                                    double totalMinutes = Math.Floor(span.TotalMinutes);

                                    if (totalMinutes <= 5)
                                    {
                                        // 0-5
                                        duration_time = "0-5min";
                                    }
                                    else if (totalMinutes <= 10)
                                    {
                                        // 6-10
                                        duration_time = "6-10min";
                                    }
                                    else if (totalMinutes <= 15)
                                    {
                                        // 11-15
                                        duration_time = "11-15min";
                                    }
                                    else if (totalMinutes <= 20)
                                    {
                                        // 16-20
                                        duration_time = "16-20min";
                                    }
                                    else if (totalMinutes <= 25)
                                    {
                                        // 21-25
                                        duration_time = "21-25min";
                                    }
                                    else if (totalMinutes <= 30)
                                    {
                                        // 26-30
                                        duration_time = "26-30min";
                                    }
                                    else if (totalMinutes <= 60)
                                    {
                                        // 31-60
                                        duration_time = "31-60min";
                                    }
                                    else if (totalMinutes >= 61)
                                    {
                                        // >60
                                        duration_time = ">60min";
                                    }

                                    String date_get = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                                    DateTime submitted_date_ = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    DateTime updated_date_ = DateTime.ParseExact(updated_date__updated_time.ToString(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    TimeSpan transaction_time = updated_date_ - submitted_date_;
                                    string transaction_time_get = "";
                                    if (transaction_time.Hours.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Hours + ":";
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Hours + ":";
                                    }
                                    if (transaction_time.Minutes.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Minutes + ":";
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Minutes + ":";
                                    }
                                    if (transaction_time.Seconds.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Seconds;
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Seconds;
                                    }

                                    var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19}", "FY", "\"" + month.ToString("MM/01/yyyy") + "\"", "\"" + date_get + "\"", "\"" + submitted_date + " " + submitted_time + "\"", "\"" + updated_date__updated_time + "\"", "\"" + member + "\"", "\"" + payment_type + "\"", "\"" + pg_company + "\"", "\"" + pg_type + "\"", "\"" + "'" + transaction_id + "\"", "\"" + amount.ToString().Replace("-", "") + "\"", "\"" + transaction_time_get + "\"", "\"" + "Withdrawal" + "\"", "\"" + duration_time + "\"", "\"" + vip + "\"", "\"" + status + "\"", "\"" + "" + "\"", "\"" + "" + "\"", "\"" + "" + "\"", "\"" + "" + "\"");
                                    _fy_csv.AppendLine(newLine);
                                }
                                else
                                {
                                    // Manual Withdrawal Record
                                    JToken member_get = jo_fy.SelectToken("$.aaData[" + ii + "][1]");
                                    string member = Regex.Match(member_get.ToString(), "<span(.*?)>(.*?)</span>").Groups[2].Value;
                                    JToken vip = jo_fy.SelectToken("$.aaData[" + ii + "][3]").ToString().Replace("\"", "");
                                    JToken amount = jo_fy.SelectToken("$.aaData[" + ii + "][5]").ToString().Replace("(RMB) - ¥ ", "");
                                    JToken remark = jo_fy.SelectToken("$.aaData[" + ii + "][8]").ToString().Replace("\"", "");
                                    JToken submitted_date__submitted_time = jo_fy.SelectToken("$.aaData[" + ii + "][10]");
                                    string submitted_date = submitted_date__submitted_time.ToString().Substring(0, 10);
                                    string submitted_time = submitted_date__submitted_time.ToString().Substring(15);
                                    JToken payment_type = jo_fy.SelectToken("$.aaData[" + ii + "][7]").ToString().Replace("\"", "");
                                    if (payment_type.ToString() != "")
                                    {
                                        if (payment_type.ToString().Substring(payment_type.ToString().Length - 1, 1) == "-")
                                        {
                                            payment_type = payment_type.ToString().Remove(payment_type.ToString().Length - 1);
                                        }
                                    }
                                    DateTime month = DateTime.ParseExact(submitted_date, "yyyy-MM-dd", CultureInfo.InvariantCulture);

                                    // Bank account
                                    string pg_company = "";
                                    string pg_type = "";
                                    string bank_account_fy_temp = Path.Combine(Path.GetTempPath(), "FY Payment Type Code.txt");

                                    if (!remark.ToString().ToLower().Contains("wrong"))
                                    {
                                        using (StreamReader sr = File.OpenText(bank_account_fy_temp))
                                        {
                                            string s = String.Empty;
                                            while ((s = sr.ReadLine()) != null)
                                            {
                                                Application.DoEvents();

                                                string[] results = s.Split("*|*");
                                                int bank_account_i = 0;
                                                bool isNext = false;
                                                foreach (string result in results)
                                                {
                                                    Application.DoEvents();

                                                    bank_account_i++;

                                                    if (bank_account_i == 1)
                                                    {
                                                        if (result == "手工存款")
                                                        {
                                                            string replace_transaction_id = remark.ToString().ToLower();
                                                            if (replace_transaction_id.Contains("wechat"))
                                                            {
                                                                pg_company = "MANUAL WECHAT";
                                                                pg_type = "MANUAL WECHAT";
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                pg_company = "LOCAL BANK";
                                                                pg_type = "LOCAL BANK";
                                                                break;
                                                            }
                                                        }
                                                        else if (result == payment_type.ToString().Trim())
                                                        {
                                                            isNext = true;
                                                        }
                                                    }

                                                    if (isNext)
                                                    {
                                                        if (bank_account_i == 2)
                                                        {
                                                            pg_company = result;
                                                        }
                                                        else if (bank_account_i == 3)
                                                        {
                                                            pg_type = result;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (payment_type.ToString() == "")
                                        {
                                            pg_company = "";
                                            pg_type = "";
                                        }
                                    }
                                    else
                                    {
                                        pg_company = "LOCAL BANK ADJUSTMENT";
                                        pg_type = "LOCAL BANK ADJUSTMENT";
                                    }

                                    string duration_time = "";
                                    DateTime start = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    DateTime end = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    TimeSpan span = end - start;
                                    double totalMinutes = Math.Floor(span.TotalMinutes);

                                    if (totalMinutes <= 5)
                                    {
                                        // 0-5
                                        duration_time = "0-5min";
                                    }
                                    else if (totalMinutes <= 10)
                                    {
                                        // 6-10
                                        duration_time = "6-10min";
                                    }
                                    else if (totalMinutes <= 15)
                                    {
                                        // 11-15
                                        duration_time = "11-15min";
                                    }
                                    else if (totalMinutes <= 20)
                                    {
                                        // 16-20
                                        duration_time = "16-20min";
                                    }
                                    else if (totalMinutes <= 25)
                                    {
                                        // 21-25
                                        duration_time = "21-25min";
                                    }
                                    else if (totalMinutes <= 30)
                                    {
                                        // 26-30
                                        duration_time = "26-30min";
                                    }
                                    else if (totalMinutes <= 60)
                                    {
                                        // 31-60
                                        duration_time = "31-60min";
                                    }
                                    else if (totalMinutes >= 61)
                                    {
                                        // >60
                                        duration_time = ">60min";
                                    }

                                    String date_get = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                                    DateTime submitted_date_ = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    DateTime updated_date_ = DateTime.ParseExact(submitted_date + " " + submitted_time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    TimeSpan transaction_time = updated_date_ - submitted_date_;
                                    string transaction_time_get = "";
                                    if (transaction_time.Hours.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Hours + ":";
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Hours + ":";
                                    }
                                    if (transaction_time.Minutes.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Minutes + ":";
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Minutes + ":";
                                    }
                                    if (transaction_time.Seconds.ToString().Length == 1)
                                    {
                                        transaction_time_get += "0" + transaction_time.Seconds;
                                    }
                                    else
                                    {
                                        transaction_time_get += transaction_time.Seconds;
                                    }

                                    var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19}", "FY", "\"" + month.ToString("MM/01/yyyy") + "\"", "\"" + date_get + "\"", "\"" + submitted_date + " " + submitted_time + "\"", "\"" + submitted_date + " " + submitted_time + "\"", "\"" + member + "\"", "\"" + payment_type + "\"", "\"" + pg_company + "\"", "\"" + pg_type + "\"", "\"" + remark + "\"", "\"" + amount.ToString().Replace("-", "") + "\"", "\"" + transaction_time_get + "\"", "\"" + "Withdrawal" + "\"", "\"" + duration_time + "\"", "\"" + vip + "\"", "\"" + "Success" + "\"", "\"" + "" + "\"", "\"" + "" + "\"", "\"" + "" + "\"", "\"" + "" + "\"");
                                    _fy_csv.AppendLine(newLine);
                                }
                            }
                        }
                        else if (selected_index == 1)
                        {
                            if (!_isSecondRequest_fy)
                            {
                                // Manual Bonus Report
                                JToken member_get = jo_fy.SelectToken("$.aaData[" + ii + "][1]");
                                string member = Regex.Match(member_get.ToString(), "<span(.*?)>(.*?)</span>").Groups[2].Value;
                                JToken vip = jo_fy.SelectToken("$.aaData[" + ii + "][3]").ToString().Replace("\"", "");
                                JToken amount = jo_fy.SelectToken("$.aaData[" + ii + "][5]").ToString().Replace("(RMB) - ¥ ", "");
                                JToken remark = jo_fy.SelectToken("$.aaData[" + ii + "][8]").ToString().Replace("\"", "");
                                JToken submitted_date__submitted_time = jo_fy.SelectToken("$.aaData[" + ii + "][10]");
                                string submitted_date = submitted_date__submitted_time.ToString().Substring(0, 10);
                                string submitted_time = submitted_date__submitted_time.ToString().Substring(15);
                                DateTime month = DateTime.ParseExact(submitted_date, "yyyy-MM-dd", CultureInfo.InvariantCulture);

                                string replace_remark = "";
                                foreach (char c in remark.ToString())
                                {
                                    if (c == ';')
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        if (c != ' ')
                                        {
                                            replace_remark += c;
                                        }
                                    }
                                }

                                // Bonus code
                                string bonus_category = "";
                                string purpose = "";
                                string bank_account_fy_temp = Path.Combine(Path.GetTempPath(), "FY Bonus Code.txt");

                                using (StreamReader sr = File.OpenText(bank_account_fy_temp))
                                {
                                    string s = String.Empty;
                                    while ((s = sr.ReadLine()) != null)
                                    {
                                        string[] results = s.Split("*|*");
                                        int bonus_code_i = 0;
                                        bool isNext = false;
                                        foreach (string result in results)
                                        {
                                            bonus_code_i++;

                                            if (bonus_code_i == 1)
                                            {
                                                if (result == replace_remark)
                                                {
                                                    isNext = true;
                                                }
                                            }

                                            if (isNext)
                                            {
                                                if (bonus_code_i == 2)
                                                {
                                                    bonus_category = result;
                                                }
                                                else if (bonus_code_i == 3)
                                                {
                                                    purpose = result;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }

                                // added asdasd
                                if (!member.ToString().ToLower().Contains("test") && !vip.ToString().ToLower().Contains("test") && !remark.ToString().ToLower().Contains("test"))
                                {
                                    if (bonus_category == "" && purpose == "")
                                    {
                                        string get1 = replace_remark.Substring(6, 3);
                                        string get2 = get1.Substring(0, 2);
                                        string get3 = get1.Substring(2);
                                        string get4 = get1.Substring(0, 2);

                                        if (get2 == "FD" || get2 == "RA")
                                        {
                                            get1 = replace_remark.Substring(6, 4);
                                            get2 = get1.Substring(0, 3);
                                            get3 = get1.Substring(3);
                                        }

                                        ArrayList items_code = new ArrayList(new string[] { "AD", "FDB", "DP", "PZ", "RF", "RAF", "RB", "SU", "TO", "RR", "CB", "GW", "RW", "TE" });
                                        ArrayList items_bonus_category = new ArrayList(new string[] { "Adjustment", "FDB", "Deposit", "Prize", "Refer friend", "Refer friend", "Reload", "Signup Bonus", "Turnover", "Rebate", "Cashback", "Goodwill Bonus", "Reward", "Test" });
                                        int count_ = 0;
                                        foreach (var item in items_code)
                                        {
                                            if (get2 == item.ToString())
                                            {
                                                bonus_category = items_bonus_category[count_].ToString();
                                                break;
                                            }

                                            count_++;
                                        }

                                        if (get3 == "0")
                                        {
                                            if (get4 == "FD" || get4 == "RA")
                                            {
                                                get1 = replace_remark.Substring(6, 5);
                                                get2 = get1.Substring(0, 4);
                                                get3 = get1.Substring(4);
                                            }
                                            else
                                            {
                                                get1 = replace_remark.Substring(6, 4);
                                                get2 = get1.Substring(0, 3);
                                                get3 = get1.Substring(3);
                                            }
                                        }

                                        ArrayList items_code_ = new ArrayList(new string[] { "0", "1", "2", "3", "4" });
                                        ArrayList items_bonus_category_ = new ArrayList(new string[] { "Retention", "Acquisition", "Conversion", "Retention", "Reactivation" });
                                        int count__ = 0;
                                        foreach (var item in items_code_)
                                        {
                                            if (get3 == item.ToString())
                                            {
                                                purpose = items_bonus_category_[count__].ToString();
                                                break;
                                            }

                                            count__++;
                                        }

                                        if (bonus_category == "" && purpose == "")
                                        {
                                            bonus_category = "Rebate";
                                            purpose = "Retention";
                                        }
                                    }
                                }
                                else
                                {
                                    bonus_category = "Other";
                                    purpose = "Adjustment";
                                }

                                if (_fy_get_ii == 1)
                                {
                                    var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", "Brand", "Month", "Date", "Username", "Bonus Category", "Purpose", "Amount", "Remark", "VIP Level");
                                    _fy_csv.AppendLine(header);
                                }

                                var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", "FY", "\"" + month.ToString("MM/01/yyyy") + "\"", "\"" + submitted_date + "\"", "\"" + member + "\"", "\"" + bonus_category + "\"", "\"" + purpose + "\"", "\"" + amount + "\"", "\"" + remark + "\"", "\"" + vip + "\"");
                                _fy_csv.AppendLine(newLine);
                            }
                            else
                            {
                                // Generated Bonus Report
                                JToken submitted_year__submitted_month__submitted_day = jo_fy.SelectToken("$.aaData[" + ii + "][0]");
                                string submitted_year = submitted_year__submitted_month__submitted_day.ToString().Substring(0, 4);
                                string submitted_month = submitted_year__submitted_month__submitted_day.ToString().Substring(4, 2);
                                string submitted_day = submitted_year__submitted_month__submitted_day.ToString().Substring(6);
                                string submitted_date = submitted_month + "/" + submitted_day + "/" + submitted_year;
                                JToken member_get = jo_fy.SelectToken("$.aaData[" + ii + "][1]");
                                string member = Regex.Match(member_get.ToString(), "<span(.*?)>(.*?)</span>").Groups[2].Value;
                                JToken vip = jo_fy.SelectToken("$.aaData[" + ii + "][3]").ToString().Replace("\"", "");
                                JToken game_platform = jo_fy.SelectToken("$.aaData[" + ii + "][5]").ToString().Replace("\"", "");
                                JToken amount = jo_fy.SelectToken("$.aaData[" + ii + "][9]").ToString().Replace("(返0)", "");
                                DateTime month = DateTime.ParseExact(submitted_year + "-" + submitted_month + "-" + submitted_day, "yyyy-MM-dd", CultureInfo.InvariantCulture);

                                string replace_remark = "";
                                foreach (char c in game_platform.ToString())
                                {
                                    if (c == ';')
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        replace_remark += c;
                                    }
                                }

                                // Bonus code
                                string bonus_category = "";
                                string purpose = "";
                                string bank_account_fy_temp = Path.Combine(Path.GetTempPath(), "FY Bonus Code.txt");

                                using (StreamReader sr = File.OpenText(bank_account_fy_temp))
                                {
                                    string s = String.Empty;
                                    while ((s = sr.ReadLine()) != null)
                                    {
                                        string[] results = s.Split("*|*");
                                        int bonus_code_i = 0;
                                        bool isNext = false;
                                        foreach (string result in results)
                                        {
                                            bonus_code_i++;

                                            if (bonus_code_i == 1)
                                            {
                                                if (result == replace_remark)
                                                {
                                                    isNext = true;
                                                }
                                            }

                                            if (isNext)
                                            {
                                                if (bonus_code_i == 2)
                                                {
                                                    bonus_category = result;
                                                }
                                                else if (bonus_code_i == 3)
                                                {
                                                    purpose = result;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }

                                if (bonus_category == "" && purpose == "")
                                {
                                    bonus_category = "Rebate";
                                    purpose = "Retention";
                                }

                                var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", "FY", "\"" + month.ToString("MM/01/yyyy") + "\"", "\"" + submitted_date + "\"", "\"" + member + "\"", "\"" + bonus_category + "\"", "\"" + purpose + "\"", "\"" + amount + "\"", "\"" + game_platform + "\"", "\"" + vip + "\"");
                                _fy_csv.AppendLine(newLine);
                            }
                        }
                        else if (selected_index == 2)
                        {
                            // Bet Record
                            JToken game_platform = jo_fy.SelectToken("$.aaData[" + ii + "][0]");
                            JToken player_id = jo_fy.SelectToken("$.aaData[" + ii + "][1][0]");
                            JToken player_name = jo_fy.SelectToken("$.aaData[" + ii + "][1][1]");
                            JToken bet_no = jo_fy.SelectToken("$.aaData[" + ii + "][2]").ToString().Replace("BetTransaction:", "");
                            JToken bet_time = jo_fy.SelectToken("$.aaData[" + ii + "][3]");
                            JToken bet_type = jo_fy.SelectToken("$.aaData[" + ii + "][4]").ToString();
                            var match = Regex.Match(bet_type.ToString(), @"<div\b[^>]*>(.*?)<\/div>");
                            if (match.Success)
                            {
                                string bet_type_replace = match.Groups[1].Value;
                                bet_type_replace = bet_type_replace.Replace("<br/>", "");
                                bet_type_replace = bet_type_replace.Replace("<div>", "");
                                bet_type_replace = bet_type_replace.Replace("</div>", "");
                                bet_type_replace = bet_type_replace.Replace("<span>", "");
                                bet_type_replace = bet_type_replace.Replace("</span>", "");
                                bet_type_replace = bet_type_replace.Replace(",", "");
                                bet_type_replace = bet_type_replace.Replace("，", "");
                                bet_type = bet_type_replace.PadRight(225).Substring(0, 225).Trim();
                            }
                            else
                            {
                                string bet_type_replace = bet_type.ToString();
                                bet_type_replace = bet_type_replace.Replace("<br/>", "");
                                bet_type_replace = bet_type_replace.Replace("<div>", "");
                                bet_type_replace = bet_type_replace.Replace("</div>", "");
                                bet_type_replace = bet_type_replace.Replace("<span>", "");
                                bet_type_replace = bet_type_replace.Replace("</span>", "");
                                bet_type_replace = bet_type_replace.Replace(",", "");
                                bet_type_replace = bet_type_replace.Replace("，", "");
                                bet_type = bet_type_replace.PadRight(225).Substring(0, 225).Trim();
                            }
                            String result_bet_type = Regex.Replace(bet_type.ToString(), @"<[^>]*>", String.Empty);
                            JToken game_result = jo_fy.SelectToken("$.aaData[" + ii + "][5]").ToString().Replace("<br>", "");
                            JToken stake_amount_color = jo_fy.SelectToken("$.aaData[" + ii + "][6][0]");
                            JToken stake_amount = jo_fy.SelectToken("$.aaData[" + ii + "][6][1]");
                            JToken win_amount_color = jo_fy.SelectToken("$.aaData[" + ii + "][7][0]");
                            JToken win_amount = jo_fy.SelectToken("$.aaData[" + ii + "][7][1]");
                            JToken company_win_loss_color = jo_fy.SelectToken("$.aaData[" + ii + "][8][0]");
                            JToken company_win_loss = jo_fy.SelectToken("$.aaData[" + ii + "][8][1]");
                            JToken valid_bet_color = jo_fy.SelectToken("$.aaData[" + ii + "][9][0]");
                            JToken valid_bet = jo_fy.SelectToken("$.aaData[" + ii + "][9][1]");
                            JToken valid_invalid_id = jo_fy.SelectToken("$.aaData[" + ii + "][10][0]");
                            JToken valid_invalid = jo_fy.SelectToken("$.aaData[" + ii + "][10][1]");
                            string bet_time_date = bet_time.ToString().Substring(0, 10);
                            DateTime month = DateTime.ParseExact(bet_time_date, "yyyy-MM-dd", CultureInfo.InvariantCulture);

                            // get vip
                            // asd comment
                            string vip = "";
                            
                            for (int i_ = 0; i_ < getmemberlist_fy.Count; i_ += 2)
                            {
                                if (getmemberlist_fy[i_] == player_name.ToString())
                                {
                                    vip = getmemberlist_fy[i_ + 1];
                                    break;
                                }
                            }

                            // asd turnover
                            // provider
                            // category

                            if (!__turnover_detect)
                            {
                                int _fy_get_ii_display_result = _fy_get_ii_display;
                                int get = _fy_get_ii_display_result+=1;
                                FY_Cronos_Data.Properties.Settings.Default.______turnover_count = get.ToString();
                                FY_Cronos_Data.Properties.Settings.Default.Save();

                                Turnover_FY(player_name.ToString(), stake_amount.ToString().Replace(",", ""), win_amount.ToString().Replace(",", ""), company_win_loss.ToString().Replace(",", ""), valid_bet.ToString().Replace(",", ""), bet_time_date, month.ToString("MM/01/yyyy"), vip, game_platform.ToString());
                            }
                            else
                            {
                                if (FY_Cronos_Data.Properties.Settings.Default.______turnover_count == _fy_get_ii_display.ToString())
                                {
                                    Turnover_FY(player_name.ToString(), stake_amount.ToString().Replace(",", ""), win_amount.ToString().Replace(",", ""), company_win_loss.ToString().Replace(",", ""), valid_bet.ToString().Replace(",", ""), bet_time_date, month.ToString("MM/01/yyyy"), vip, game_platform.ToString());
                                    __turnover_detect = false;
                                }
                            }

                            if (_fy_get_ii == 1)
                            {
                                var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}", "Month", "Date", "VIP", "Game Platform", "Username", "Bet No", "Bet Time", "Bet Type", "Game Result", "Stake Amount", "Win Amount", "Company Win/Loss", "Valid Bet", "Valid/Invalid");
                                _fy_csv.AppendLine(header);
                            }

                            result_bet_type = result_bet_type.ToString().Replace(";", "");
                            string result_bet_type_replace = Regex.Replace(result_bet_type, @"\t|\n|\r", "");

                            var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}", "\"" + month.ToString("MM/01/yyyy") + "\"", "\"" + bet_time_date + "\"", "\"" + vip + "\"", "\"" + game_platform + "\"", "\"" + player_name + "\"", "\"" + "'" + bet_no + "\"", "\"" + bet_time + "\"", "\"" + result_bet_type_replace + "\"", "\"" + game_result + "\"", "\"" + stake_amount + "\"", "\"" + win_amount + "\"", "\"" + company_win_loss + "\"", "\"" + valid_bet + "\"", "\"" + valid_invalid + "\"");
                            _fy_csv.AppendLine(newLine);
                        }

                        if ((_fy_get_ii) == _limit_fy)
                        {
                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: saving excel... --- BET RECORD";

                            _fy_get_ii = 0;

                            _fy_displayinexel_i++;
                            StringBuilder replace_datetime_fy = new StringBuilder(dateTimePicker_start_fy.Text.Substring(0, 10) + "__" + dateTimePicker_end_fy.Text.Substring(0, 10));
                            replace_datetime_fy.Replace(" ", "_");

                            if (_fy_current_datetime == "")
                            {
                                _fy_current_datetime = DateTime.Now.ToString("yyyy-MM-dd");
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data"))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data");
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY"))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY");
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime);
                            }

                            string replace = _fy_displayinexel_i.ToString();

                            if (_fy_displayinexel_i.ToString().Length == 1)
                            {
                                replace = "0" + _fy_displayinexel_i;
                            }

                            if (selected_index == 0)
                            {
                                if (!_isSecondRequestFinish_fy)
                                {
                                    if (!_isSecondRequest_fy)
                                    {
                                        // Deposit Record
                                        if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report"))
                                        {
                                            Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report");
                                        }

                                        _fy_filename = "FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                        _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                                        _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                        _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\";

                                        if (File.Exists(_fy_folder_path_result))
                                        {
                                            File.Delete(_fy_folder_path_result);
                                        }

                                        if (File.Exists(_fy_folder_path_result_xlsx))
                                        {
                                            File.Delete(_fy_folder_path_result_xlsx);
                                        }

                                        //_fy_csv.ToString().Reverse();
                                        File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                                        Excel.Application app = new Excel.Application();
                                        Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                        Excel.Worksheet worksheet = wb.ActiveSheet;
                                        worksheet.Activate();
                                        worksheet.Application.ActiveWindow.SplitRow = 1;
                                        worksheet.Application.ActiveWindow.FreezePanes = true;
                                        Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                                        firstRow.AutoFilter(1,
                                                            Type.Missing,
                                                            Excel.XlAutoFilterOperator.xlAnd,
                                                            Type.Missing,
                                                            true);
                                        worksheet.Columns[2].NumberFormat = "MMM-yy";
                                        worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                                        worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                                        worksheet.Columns[12].NumberFormat = "HH:mm";
                                        //worksheet.Columns[8].Replace(" ", "");
                                        worksheet.Columns[8].NumberFormat = "@";
                                        //asd123
                                        Excel.Range usedRange = worksheet.UsedRange;
                                        Excel.Range rows = usedRange.Rows;
                                        int count = 0;
                                        foreach (Excel.Range row in rows)
                                        {
                                            if (count == 0)
                                            {
                                                Excel.Range firstCell = row.Cells[1];

                                                string firstCellValue = firstCell.Value as String;

                                                if (!string.IsNullOrEmpty(firstCellValue))
                                                {
                                                    row.Interior.Color = Color.FromArgb(222, 30, 112);
                                                    row.Font.Color = Color.FromArgb(255, 255, 255);
                                                    row.Font.Bold = true;
                                                    row.Font.Size = 12;
                                                }

                                                break;
                                            }

                                            count++;
                                        }
                                        int i_excel;
                                        for (i_excel = 1; i_excel <= 20; i_excel++)
                                        {
                                            worksheet.Columns[i_excel].ColumnWidth = 20;
                                        }
                                        wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                        wb.Close();
                                        app.Quit();
                                        Marshal.ReleaseComObject(app);

                                        _fy_csv.Clear();

                                        label_fy_currentrecord.Text = (_fy_get_ii_display).ToString("N0") + " of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                                        label_fy_currentrecord.Invalidate();
                                        label_fy_currentrecord.Update();

                                        //if (File.Exists(_fy_folder_path_result))
                                        //{
                                        //    File.Delete(_fy_folder_path_result);
                                        //}
                                    }
                                    else
                                    {
                                        // Manual Deposit Record
                                        if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report"))
                                        {
                                            Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report");
                                        }

                                        _fy_filename = "FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                        _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                                        _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                        _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\";

                                        //if (File.Exists(_fy_folder_path_result))
                                        //{
                                        //    File.Delete(_fy_folder_path_result);
                                        //}

                                        if (File.Exists(_fy_folder_path_result_xlsx))
                                        {
                                            File.Delete(_fy_folder_path_result_xlsx);
                                        }

                                        //var lines = File.ReadAllLines(_fy_folder_path_result).Where(arg => !string.IsNullOrWhiteSpace(arg));

                                        using (StreamWriter file = new StreamWriter(_fy_folder_path_result, true, Encoding.UTF8))
                                        {
                                            file.Write(_fy_csv.ToString());
                                        }

                                        //File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                                        Excel.Application app = new Excel.Application();
                                        Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                        Excel.Worksheet worksheet = wb.ActiveSheet;
                                        worksheet.Activate();
                                        worksheet.Application.ActiveWindow.SplitRow = 1;
                                        worksheet.Application.ActiveWindow.FreezePanes = true;
                                        Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                                        firstRow.AutoFilter(1,
                                                            Type.Missing,
                                                            Excel.XlAutoFilterOperator.xlAnd,
                                                            Type.Missing,
                                                            true);
                                        worksheet.Columns[2].NumberFormat = "MMM-yy";
                                        worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                                        worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                                        worksheet.Columns[12].NumberFormat = "HH:mm";
                                        //worksheet.Columns[8].Replace(" ", "");
                                        //worksheet.Columns[8].NumberFormat = "@";
                                        Excel.Range usedRange = worksheet.UsedRange;
                                        Excel.Range rows = usedRange.Rows;
                                        int count = 0;
                                        foreach (Excel.Range row in rows)
                                        {
                                            if (count == 0)
                                            {
                                                Excel.Range firstCell = row.Cells[1];

                                                string firstCellValue = firstCell.Value as String;

                                                if (!string.IsNullOrEmpty(firstCellValue))
                                                {
                                                    row.Interior.Color = Color.FromArgb(222, 30, 112);
                                                    row.Font.Color = Color.FromArgb(255, 255, 255);
                                                    row.Font.Bold = true;
                                                    row.Font.Size = 12;
                                                }

                                                break;
                                            }

                                            count++;
                                        }
                                        int i_excel;
                                        for (i_excel = 1; i_excel <= 20; i_excel++)
                                        {
                                            worksheet.Columns[i_excel].ColumnWidth = 20;
                                        }
                                        wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                        wb.Close();
                                        app.Quit();
                                        Marshal.ReleaseComObject(app);

                                        //if (File.Exists(_fy_folder_path_result))
                                        //{
                                        //    File.Delete(_fy_folder_path_result);
                                        //}
                                    }
                                }
                                else
                                {
                                    if (!_isThirdRequest_fy)
                                    {
                                        // Withdrawal Record
                                        if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report"))
                                        {
                                            Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report");
                                        }

                                        _fy_filename = "FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                        _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                                        _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                        _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\";

                                        //if (File.Exists(_fy_folder_path_result))
                                        //{
                                        //    File.Delete(_fy_folder_path_result);
                                        //}

                                        if (File.Exists(_fy_folder_path_result_xlsx))
                                        {
                                            File.Delete(_fy_folder_path_result_xlsx);
                                        }

                                        using (StreamWriter file = new StreamWriter(_fy_folder_path_result, true, Encoding.UTF8))
                                        {
                                            file.Write(_fy_csv.ToString());
                                        }

                                        //_fy_csv.ToString().Reverse();
                                        //File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                                        Excel.Application app = new Excel.Application();
                                        Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                        Excel.Worksheet worksheet = wb.ActiveSheet;
                                        worksheet.Activate();
                                        worksheet.Application.ActiveWindow.SplitRow = 1;
                                        worksheet.Application.ActiveWindow.FreezePanes = true;
                                        Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                                        firstRow.AutoFilter(1,
                                                            Type.Missing,
                                                            Excel.XlAutoFilterOperator.xlAnd,
                                                            Type.Missing,
                                                            true);
                                        worksheet.Columns[2].NumberFormat = "MMM-yy";
                                        worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                                        worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                                        worksheet.Columns[12].NumberFormat = "HH:mm";
                                        //worksheet.Columns[8].Replace(" ", "");
                                        worksheet.Columns[8].NumberFormat = "@";
                                        Excel.Range usedRange = worksheet.UsedRange;
                                        Excel.Range rows = usedRange.Rows;
                                        int count = 0;
                                        foreach (Excel.Range row in rows)
                                        {
                                            if (count == 0)
                                            {
                                                Excel.Range firstCell = row.Cells[1];

                                                string firstCellValue = firstCell.Value as String;

                                                if (!string.IsNullOrEmpty(firstCellValue))
                                                {
                                                    row.Interior.Color = Color.FromArgb(222, 30, 112);
                                                    row.Font.Color = Color.FromArgb(255, 255, 255);
                                                    row.Font.Bold = true;
                                                    row.Font.Size = 12;
                                                }

                                                break;
                                            }

                                            count++;
                                        }
                                        int i_excel;
                                        for (i_excel = 1; i_excel <= 20; i_excel++)
                                        {
                                            worksheet.Columns[i_excel].ColumnWidth = 20;
                                        }
                                        wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                        wb.Close();
                                        app.Quit();
                                        Marshal.ReleaseComObject(app);

                                        //if (File.Exists(_fy_folder_path_result))
                                        //{
                                        //    File.Delete(_fy_folder_path_result);
                                        //}
                                    }
                                    else
                                    {
                                        // Manual Withdrawal Record
                                        if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report"))
                                        {
                                            Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report");
                                        }

                                        _fy_filename = "FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                        _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                                        _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                        _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\";

                                        //if (File.Exists(_fy_folder_path_result))
                                        //{
                                        //    File.Delete(_fy_folder_path_result);
                                        //}

                                        if (File.Exists(_fy_folder_path_result_xlsx))
                                        {
                                            File.Delete(_fy_folder_path_result_xlsx);
                                        }

                                        using (StreamWriter file = new StreamWriter(_fy_folder_path_result, true, Encoding.UTF8))
                                        {
                                            file.Write(_fy_csv.ToString());
                                        }

                                        //File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                                        Excel.Application app = new Excel.Application();
                                        Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                        Excel.Worksheet worksheet = wb.ActiveSheet;
                                        worksheet.Activate();
                                        worksheet.Application.ActiveWindow.SplitRow = 1;
                                        worksheet.Application.ActiveWindow.FreezePanes = true;
                                        Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                                        firstRow.AutoFilter(1,
                                                            Type.Missing,
                                                            Excel.XlAutoFilterOperator.xlAnd,
                                                            Type.Missing,
                                                            true);
                                        worksheet.Columns[2].NumberFormat = "MMM-yy";
                                        worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                                        worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                                        worksheet.Columns[12].NumberFormat = "HH:mm";
                                        //worksheet.Columns[8].Replace(" ", "");
                                        //worksheet.Columns[8].NumberFormat = "@";
                                        Excel.Range usedRange = worksheet.UsedRange;
                                        Excel.Range rows = usedRange.Rows;
                                        int count = 0;
                                        foreach (Excel.Range row in rows)
                                        {
                                            if (count == 0)
                                            {
                                                Excel.Range firstCell = row.Cells[1];

                                                string firstCellValue = firstCell.Value as String;

                                                if (!string.IsNullOrEmpty(firstCellValue))
                                                {
                                                    row.Interior.Color = Color.FromArgb(222, 30, 112);
                                                    row.Font.Color = Color.FromArgb(255, 255, 255);
                                                    row.Font.Bold = true;
                                                    row.Font.Size = 12;
                                                }

                                                break;
                                            }

                                            count++;
                                        }
                                        int i_excel;
                                        for (i_excel = 1; i_excel <= 20; i_excel++)
                                        {
                                            worksheet.Columns[i_excel].ColumnWidth = 20;
                                        }
                                        wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                        wb.Close();
                                        app.Quit();
                                        Marshal.ReleaseComObject(app);

                                        // comment
                                        //if (File.Exists(_fy_folder_path_result))
                                        //{
                                        //    File.Delete(_fy_folder_path_result);
                                        //}
                                    }
                                }
                            }
                            else if (selected_index == 1)
                            {
                                if (!_isSecondRequest_fy)
                                {
                                    // Manual Bonus Report
                                    if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report"))
                                    {
                                        Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report");
                                    }

                                    _fy_filename = "FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                    _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                                    _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                    _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\";

                                    if (File.Exists(_fy_folder_path_result))
                                    {
                                        File.Delete(_fy_folder_path_result);
                                    }

                                    if (File.Exists(_fy_folder_path_result_xlsx))
                                    {
                                        File.Delete(_fy_folder_path_result_xlsx);
                                    }

                                    //_fy_csv.ToString().Reverse();
                                    File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                                    Excel.Application app = new Excel.Application();
                                    Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                    Excel.Worksheet worksheet = wb.ActiveSheet;
                                    worksheet.Activate();
                                    worksheet.Application.ActiveWindow.SplitRow = 1;
                                    worksheet.Application.ActiveWindow.FreezePanes = true;
                                    Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                                    firstRow.AutoFilter(1,
                                                        Type.Missing,
                                                        Excel.XlAutoFilterOperator.xlAnd,
                                                        Type.Missing,
                                                        true);
                                    //worksheet.Columns[2].NumberFormat = "MMM-yy";
                                    //worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                                    //worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                                    //worksheet.Columns[8].Replace(" ", "");
                                    //worksheet.Columns[8].NumberFormat = "@";
                                    Excel.Range usedRange = worksheet.UsedRange;
                                    Excel.Range rows = usedRange.Rows;
                                    int count = 0;
                                    foreach (Excel.Range row in rows)
                                    {
                                        if (count == 0)
                                        {
                                            Excel.Range firstCell = row.Cells[1];

                                            string firstCellValue = firstCell.Value as String;

                                            if (!string.IsNullOrEmpty(firstCellValue))
                                            {
                                                row.Interior.Color = Color.FromArgb(222, 30, 112);
                                                row.Font.Color = Color.FromArgb(255, 255, 255);
                                                row.Font.Bold = true;
                                                row.Font.Size = 12;
                                            }

                                            break;
                                        }

                                        count++;
                                    }
                                    int i_excel;
                                    for (i_excel = 1; i_excel <= 20; i_excel++)
                                    {
                                        worksheet.Columns[i_excel].ColumnWidth = 20;
                                    }
                                    wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                    wb.Close();
                                    app.Quit();
                                    Marshal.ReleaseComObject(app);

                                    _fy_csv.Clear();

                                    label_fy_currentrecord.Text = (_fy_get_ii_display).ToString("N0") + " of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                                    label_fy_currentrecord.Invalidate();
                                    label_fy_currentrecord.Update();

                                    //if (File.Exists(_fy_folder_path_result))
                                    //{
                                    //    File.Delete(_fy_folder_path_result);
                                    //}
                                }
                                else
                                {
                                    // Generated Bonus Report
                                    if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report"))
                                    {
                                        Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report");
                                    }

                                    _fy_filename = "FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                    _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                                    _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                    _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\";

                                    //if (File.Exists(_fy_folder_path_result))
                                    //{
                                    //    File.Delete(_fy_folder_path_result);
                                    //}

                                    if (File.Exists(_fy_folder_path_result_xlsx))
                                    {
                                        File.Delete(_fy_folder_path_result_xlsx);
                                    }

                                    using (StreamWriter file = new StreamWriter(_fy_folder_path_result, true, Encoding.UTF8))
                                    {
                                        file.WriteLine(_fy_csv.ToString());
                                    }

                                    //File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                                    Excel.Application app = new Excel.Application();
                                    Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                    Excel.Worksheet worksheet = wb.ActiveSheet;
                                    worksheet.Activate();
                                    worksheet.Application.ActiveWindow.SplitRow = 1;
                                    worksheet.Application.ActiveWindow.FreezePanes = true;
                                    Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                                    firstRow.AutoFilter(1,
                                                        Type.Missing,
                                                        Excel.XlAutoFilterOperator.xlAnd,
                                                        Type.Missing,
                                                        true);
                                    //worksheet.Columns[2].NumberFormat = "MMM-yy";
                                    //worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                                    //worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                                    //worksheet.Columns[8].Replace(" ", "");
                                    //worksheet.Columns[8].NumberFormat = "@";
                                    Excel.Range usedRange = worksheet.UsedRange;
                                    Excel.Range rows = usedRange.Rows;
                                    int count = 0;
                                    foreach (Excel.Range row in rows)
                                    {
                                        if (count == 0)
                                        {
                                            Excel.Range firstCell = row.Cells[1];

                                            string firstCellValue = firstCell.Value as String;

                                            if (!string.IsNullOrEmpty(firstCellValue))
                                            {
                                                row.Interior.Color = Color.FromArgb(222, 30, 112);
                                                row.Font.Color = Color.FromArgb(255, 255, 255);
                                                row.Font.Bold = true;
                                                row.Font.Size = 12;
                                            }

                                            break;
                                        }

                                        count++;
                                    }
                                    int i_excel;
                                    for (i_excel = 1; i_excel <= 20; i_excel++)
                                    {
                                        worksheet.Columns[i_excel].ColumnWidth = 20;
                                    }
                                    wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                    wb.Close();
                                    app.Quit();
                                    Marshal.ReleaseComObject(app);

                                    if (File.Exists(_fy_folder_path_result))
                                    {
                                        File.Delete(_fy_folder_path_result);
                                    }
                                }
                            }
                            else if (selected_index == 2)
                            {
                                // Bet Record
                                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record"))
                                {
                                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record");
                                }

                                _fy_filename = "FY_BetRecord_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record\\FY_BetRecord_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                                _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record\\FY_BetRecord_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                                _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record\\";

                                if (File.Exists(_fy_folder_path_result))
                                {
                                    File.Delete(_fy_folder_path_result);
                                }

                                if (File.Exists(_fy_folder_path_result_xlsx))
                                {
                                    File.Delete(_fy_folder_path_result_xlsx);
                                }

                                //_fy_csv.ToString().Reverse();
                                File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                                Excel.Application app = new Excel.Application();
                                Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                Excel.Worksheet worksheet = wb.ActiveSheet;
                                worksheet.Activate();
                                worksheet.Application.ActiveWindow.SplitRow = 1;
                                worksheet.Application.ActiveWindow.FreezePanes = true;
                                Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                                firstRow.AutoFilter(1,
                                                    Type.Missing,
                                                    Excel.XlAutoFilterOperator.xlAnd,
                                                    Type.Missing,
                                                    true);
                                //worksheet.Columns[3].Replace(" ", "");
                                //worksheet.Columns[3].NumberFormat = "@";
                                //worksheet.Columns[2].NumberFormat = "MMM-yy";
                                //worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                                //worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                                Excel.Range usedRange = worksheet.UsedRange;
                                Excel.Range rows = usedRange.Rows;
                                int count = 0;
                                foreach (Excel.Range row in rows)
                                {
                                    if (count == 0)
                                    {
                                        Excel.Range firstCell = row.Cells[1];

                                        string firstCellValue = firstCell.Value as String;

                                        if (!string.IsNullOrEmpty(firstCellValue))
                                        {
                                            row.Interior.Color = Color.FromArgb(222, 30, 112);
                                            row.Font.Color = Color.FromArgb(255, 255, 255);
                                            row.Font.Bold = true;
                                            row.Font.Size = 12;
                                        }

                                        break;
                                    }

                                    count++;
                                }
                                int i_excel;
                                for (i_excel = 1; i_excel <= 20; i_excel++)
                                {
                                    worksheet.Columns[i_excel].ColumnWidth = 20;
                                }
                                wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                wb.Close();
                                app.Quit();
                                Marshal.ReleaseComObject(app);

                                _fy_csv.Clear();

                                label_fy_currentrecord.Text = (_fy_get_ii_display).ToString("N0") + " of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                                label_fy_currentrecord.Invalidate();
                                label_fy_currentrecord.Update();
                                
                                //SaveAsTurnOver_FY(replace);
                                
                                // Database Bet Record FY
                                // asd comment
                                isBetRecordInsert = false;
                                InsertBetRecord_FY(_fy_folder_path_result);
                                label_fy_insert.Visible = false;
                            }
                        }
                        else
                        {
                            label_fy_currentrecord.Text = (_fy_get_ii_display).ToString("N0") + " of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                            label_fy_currentrecord.Invalidate();
                            label_fy_currentrecord.Update();
                        }

                        _fy_get_ii++;
                        _fy_get_ii_display++;
                    }

                    _result_count_json_fy = 0;

                    // web client request
                    await GetDataFYPagesAsync();
                }
                
                if (_fy_inserted_in_excel)
                {
                    FY_InsertDone();
                    _isDone_fy = true;
                }
            }
        }
        
        int detect_fy = 0;

        private void FY_InsertDone()
        {
            _fy_displayinexel_i++;
            StringBuilder replace_datetime_fy = new StringBuilder(dateTimePicker_start_fy.Text.Substring(0, 10) + "__" + dateTimePicker_end_fy.Text.Substring(0, 10));
            replace_datetime_fy.Replace(" ", "_");

            if (_fy_current_datetime == "")
            {
                _fy_current_datetime = DateTime.Now.ToString("yyyy-MM-dd");
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data");
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY");
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime);
            }

            string replace = _fy_displayinexel_i.ToString();

            if (_fy_displayinexel_i.ToString().Length == 1)
            {
                replace = "0" + _fy_displayinexel_i;
            }

            int selected_index = comboBox_fy_list.SelectedIndex;
            if (selected_index == 0)
            {
                if (!_isSecondRequestFinish_fy)
                {
                    if (!_isSecondRequest_fy)
                    {
                        // Deposit Record
                        if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report"))
                        {
                            Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report");
                        }

                        _fy_filename = "FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                        _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                        _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                        _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\";

                        if (File.Exists(_fy_folder_path_result))
                        {
                            File.Delete(_fy_folder_path_result);
                        }

                        if (File.Exists(_fy_folder_path_result_xlsx))
                        {
                            File.Delete(_fy_folder_path_result_xlsx);
                        }

                        //_fy_csv.ToString().Reverse();
                        File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                        Excel.Application app = new Excel.Application();
                        Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        Excel.Worksheet worksheet = wb.ActiveSheet;
                        worksheet.Activate();
                        worksheet.Application.ActiveWindow.SplitRow = 1;
                        worksheet.Application.ActiveWindow.FreezePanes = true;
                        Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                        firstRow.AutoFilter(1,
                                            Type.Missing,
                                            Excel.XlAutoFilterOperator.xlAnd,
                                            Type.Missing,
                                            true);
                        worksheet.Columns[2].NumberFormat = "MMM-yy";
                        worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                        worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                        worksheet.Columns[12].NumberFormat = "HH:mm";
                        //worksheet.Columns[8].Replace(" ", "");
                        worksheet.Columns[8].NumberFormat = "@";
                        //asd123
                        Excel.Range usedRange = worksheet.UsedRange;
                        Excel.Range rows = usedRange.Rows;
                        int count = 0;
                        foreach (Excel.Range row in rows)
                        {
                            if (count == 0)
                            {
                                Excel.Range firstCell = row.Cells[1];

                                string firstCellValue = firstCell.Value as String;

                                if (!string.IsNullOrEmpty(firstCellValue))
                                {
                                    row.Interior.Color = Color.FromArgb(222, 30, 112);
                                    row.Font.Color = Color.FromArgb(255, 255, 255);
                                    row.Font.Bold = true;
                                    row.Font.Size = 12;
                                }

                                break;
                            }

                            count++;
                        }
                        int i_excel;
                        for (i_excel = 1; i_excel <= 20; i_excel++)
                        {
                            worksheet.Columns[i_excel].ColumnWidth = 20;
                        }
                        wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        wb.Close();
                        app.Quit();
                        Marshal.ReleaseComObject(app);

                        //if (File.Exists(_fy_folder_path_result))
                        //{
                        //    File.Delete(_fy_folder_path_result);
                        //}
                    }
                    else
                    {
                        // Manual Deposit Record
                        if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report"))
                        {
                            Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report");
                        }

                        _fy_filename = "FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                        _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                        _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                        _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\";

                        //if (File.Exists(_fy_folder_path_result))
                        //{
                        //    File.Delete(_fy_folder_path_result);
                        //}

                        if (File.Exists(_fy_folder_path_result_xlsx))
                        {
                            File.Delete(_fy_folder_path_result_xlsx);
                        }

                        //var lines = File.ReadAllLines(_fy_folder_path_result).Where(arg => !string.IsNullOrWhiteSpace(arg));

                        using (StreamWriter file = new StreamWriter(_fy_folder_path_result, true, Encoding.UTF8))
                        {
                            file.Write(_fy_csv.ToString());
                        }

                        //File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                        Excel.Application app = new Excel.Application();
                        Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        Excel.Worksheet worksheet = wb.ActiveSheet;
                        worksheet.Activate();
                        worksheet.Application.ActiveWindow.SplitRow = 1;
                        worksheet.Application.ActiveWindow.FreezePanes = true;
                        Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                        firstRow.AutoFilter(1,
                                            Type.Missing,
                                            Excel.XlAutoFilterOperator.xlAnd,
                                            Type.Missing,
                                            true);
                        worksheet.Columns[2].NumberFormat = "MMM-yy";
                        worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                        worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                        worksheet.Columns[12].NumberFormat = "HH:mm";
                        //worksheet.Columns[8].Replace(" ", "");
                        worksheet.Columns[8].NumberFormat = "@";
                        Excel.Range usedRange = worksheet.UsedRange;
                        Excel.Range rows = usedRange.Rows;
                        int count = 0;
                        foreach (Excel.Range row in rows)
                        {
                            if (count == 0)
                            {
                                Excel.Range firstCell = row.Cells[1];

                                string firstCellValue = firstCell.Value as String;

                                if (!string.IsNullOrEmpty(firstCellValue))
                                {
                                    row.Interior.Color = Color.FromArgb(222, 30, 112);
                                    row.Font.Color = Color.FromArgb(255, 255, 255);
                                    row.Font.Bold = true;
                                    row.Font.Size = 12;
                                }

                                break;
                            }

                            count++;
                        }
                        int i_excel;
                        for (i_excel = 1; i_excel <= 20; i_excel++)
                        {
                            worksheet.Columns[i_excel].ColumnWidth = 20;
                        }
                        wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        wb.Close();
                        app.Quit();
                        Marshal.ReleaseComObject(app);

                        //if (File.Exists(_fy_folder_path_result))
                        //{
                        //    File.Delete(_fy_folder_path_result);
                        //}
                    }
                }
                else
                {
                    if (!_isThirdRequest_fy)
                    {
                        // Withdrawal Record
                        if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report"))
                        {
                            Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report");
                        }

                        _fy_filename = "FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                        _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                        _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                        _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\";

                        //if (File.Exists(_fy_folder_path_result))
                        //{
                        //    File.Delete(_fy_folder_path_result);
                        //}

                        if (File.Exists(_fy_folder_path_result_xlsx))
                        {
                            File.Delete(_fy_folder_path_result_xlsx);
                        }

                        using (StreamWriter file = new StreamWriter(_fy_folder_path_result, true, Encoding.UTF8))
                        {
                            file.Write(_fy_csv.ToString());
                        }

                        //_fy_csv.ToString().Reverse();
                        //File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                        Excel.Application app = new Excel.Application();
                        Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        Excel.Worksheet worksheet = wb.ActiveSheet;
                        worksheet.Activate();
                        worksheet.Application.ActiveWindow.SplitRow = 1;
                        worksheet.Application.ActiveWindow.FreezePanes = true;
                        Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                        firstRow.AutoFilter(1,
                                            Type.Missing,
                                            Excel.XlAutoFilterOperator.xlAnd,
                                            Type.Missing,
                                            true);
                        worksheet.Columns[2].NumberFormat = "MMM-yy";
                        worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                        worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                        worksheet.Columns[12].NumberFormat = "HH:mm";
                        //worksheet.Columns[8].Replace(" ", "");
                        worksheet.Columns[8].NumberFormat = "@";
                        Excel.Range usedRange = worksheet.UsedRange;
                        Excel.Range rows = usedRange.Rows;
                        int count = 0;
                        foreach (Excel.Range row in rows)
                        {
                            if (count == 0)
                            {
                                Excel.Range firstCell = row.Cells[1];

                                string firstCellValue = firstCell.Value as String;

                                if (!string.IsNullOrEmpty(firstCellValue))
                                {
                                    row.Interior.Color = Color.FromArgb(222, 30, 112);
                                    row.Font.Color = Color.FromArgb(255, 255, 255);
                                    row.Font.Bold = true;
                                    row.Font.Size = 12;
                                }

                                break;
                            }

                            count++;
                        }
                        int i_excel;
                        for (i_excel = 1; i_excel <= 20; i_excel++)
                        {
                            worksheet.Columns[i_excel].ColumnWidth = 20;
                        }
                        wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        wb.Close();
                        app.Quit();
                        Marshal.ReleaseComObject(app);

                        //if (File.Exists(_fy_folder_path_result))
                        //{
                        //    File.Delete(_fy_folder_path_result);
                        //}
                    }
                    else
                    {
                        // Manual Withdrawal Record
                        if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report"))
                        {
                            Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report");
                        }

                        _fy_filename = "FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                        _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                        _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\FY_PaymentReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                        _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Payment Report\\";

                        //if (File.Exists(_fy_folder_path_result))
                        //{
                        //    File.Delete(_fy_folder_path_result);
                        //}

                        if (File.Exists(_fy_folder_path_result_xlsx))
                        {
                            File.Delete(_fy_folder_path_result_xlsx);
                        }

                        using (StreamWriter file = new StreamWriter(_fy_folder_path_result, true, Encoding.UTF8))
                        {
                            file.Write(_fy_csv.ToString());
                        }

                        //File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                        Excel.Application app = new Excel.Application();
                        Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        Excel.Worksheet worksheet = wb.ActiveSheet;
                        worksheet.Activate();
                        worksheet.Application.ActiveWindow.SplitRow = 1;
                        worksheet.Application.ActiveWindow.FreezePanes = true;
                        Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                        firstRow.AutoFilter(1,
                                            Type.Missing,
                                            Excel.XlAutoFilterOperator.xlAnd,
                                            Type.Missing,
                                            true);
                        worksheet.Columns[2].NumberFormat = "MMM-yy";
                        worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                        worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                        worksheet.Columns[12].NumberFormat = "HH:mm";
                        //worksheet.Columns[8].Replace(" ", "");
                        worksheet.Columns[8].NumberFormat = "@";
                        Excel.Range usedRange = worksheet.UsedRange;
                        Excel.Range rows = usedRange.Rows;
                        int count = 0;
                        foreach (Excel.Range row in rows)
                        {
                            if (count == 0)
                            {
                                Excel.Range firstCell = row.Cells[1];

                                string firstCellValue = firstCell.Value as String;

                                if (!string.IsNullOrEmpty(firstCellValue))
                                {
                                    row.Interior.Color = Color.FromArgb(222, 30, 112);
                                    row.Font.Color = Color.FromArgb(255, 255, 255);
                                    row.Font.Bold = true;
                                    row.Font.Size = 12;
                                }

                                break;
                            }

                            count++;
                        }
                        int i_excel;
                        for (i_excel = 1; i_excel <= 20; i_excel++)
                        {
                            worksheet.Columns[i_excel].ColumnWidth = 20;
                        }
                        wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        wb.Close();
                        app.Quit();
                        Marshal.ReleaseComObject(app);

                        // comment
                        //if (File.Exists(_fy_folder_path_result))
                        //{
                        //    File.Delete(_fy_folder_path_result);
                        //}
                    }
                }
            }
            else if (selected_index == 1)
            {
                if (!_isSecondRequest_fy)
                {
                    // Manual Bonus Report
                    if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report"))
                    {
                        Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report");
                    }

                    _fy_filename = "FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                    _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                    _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                    _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\";

                    if (File.Exists(_fy_folder_path_result))
                    {
                        File.Delete(_fy_folder_path_result);
                    }

                    if (File.Exists(_fy_folder_path_result_xlsx))
                    {
                        File.Delete(_fy_folder_path_result_xlsx);
                    }

                    //_fy_csv.ToString().Reverse();
                    File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                    Excel.Application app = new Excel.Application();
                    Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet worksheet = wb.ActiveSheet;
                    worksheet.Activate();
                    worksheet.Application.ActiveWindow.SplitRow = 1;
                    worksheet.Application.ActiveWindow.FreezePanes = true;
                    Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                    firstRow.AutoFilter(1,
                                        Type.Missing,
                                        Excel.XlAutoFilterOperator.xlAnd,
                                        Type.Missing,
                                        true);
                    //worksheet.Columns[2].NumberFormat = "MMM-yy";
                    //worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                    //worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                    //worksheet.Columns[8].Replace(" ", "");
                    //worksheet.Columns[8].NumberFormat = "@";
                    Excel.Range usedRange = worksheet.UsedRange;
                    Excel.Range rows = usedRange.Rows;
                    int count = 0;
                    foreach (Excel.Range row in rows)
                    {
                        if (count == 0)
                        {
                            Excel.Range firstCell = row.Cells[1];

                            string firstCellValue = firstCell.Value as String;

                            if (!string.IsNullOrEmpty(firstCellValue))
                            {
                                row.Interior.Color = Color.FromArgb(222, 30, 112);
                                row.Font.Color = Color.FromArgb(255, 255, 255);
                                row.Font.Bold = true;
                                row.Font.Size = 12;
                            }

                            break;
                        }

                        count++;
                    }
                    int i_excel;
                    for (i_excel = 1; i_excel <= 20; i_excel++)
                    {
                        worksheet.Columns[i_excel].ColumnWidth = 20;
                    }
                    wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wb.Close();
                    app.Quit();
                    Marshal.ReleaseComObject(app);

                    //if (File.Exists(_fy_folder_path_result))
                    //{
                    //    File.Delete(_fy_folder_path_result);
                    //}
                }
                else
                {
                    // Generated Bonus Report
                    if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report"))
                    {
                        Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report");
                    }

                    _fy_filename = "FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                    _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                    _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\FY_BonusReport_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                    _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bonus Report\\";

                    //if (File.Exists(_fy_folder_path_result))
                    //{
                    //    File.Delete(_fy_folder_path_result);
                    //}

                    if (File.Exists(_fy_folder_path_result_xlsx))
                    {
                        File.Delete(_fy_folder_path_result_xlsx);
                    }

                    using (StreamWriter file = new StreamWriter(_fy_folder_path_result, true, Encoding.UTF8))
                    {
                        file.WriteLine(_fy_csv.ToString());
                    }

                    //File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                    Excel.Application app = new Excel.Application();
                    Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet worksheet = wb.ActiveSheet;
                    worksheet.Activate();
                    worksheet.Application.ActiveWindow.SplitRow = 1;
                    worksheet.Application.ActiveWindow.FreezePanes = true;
                    Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                    firstRow.AutoFilter(1,
                                        Type.Missing,
                                        Excel.XlAutoFilterOperator.xlAnd,
                                        Type.Missing,
                                        true);
                    //worksheet.Columns[2].NumberFormat = "MMM-yy";
                    //worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                    //worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                    //worksheet.Columns[8].Replace(" ", "");
                    //worksheet.Columns[8].NumberFormat = "@";
                    Excel.Range usedRange = worksheet.UsedRange;
                    Excel.Range rows = usedRange.Rows;
                    int count = 0;
                    foreach (Excel.Range row in rows)
                    {
                        if (count == 0)
                        {
                            Excel.Range firstCell = row.Cells[1];

                            string firstCellValue = firstCell.Value as String;

                            if (!string.IsNullOrEmpty(firstCellValue))
                            {
                                row.Interior.Color = Color.FromArgb(222, 30, 112);
                                row.Font.Color = Color.FromArgb(255, 255, 255);
                                row.Font.Bold = true;
                                row.Font.Size = 12;
                            }

                            break;
                        }

                        count++;
                    }
                    int i_excel;
                    for (i_excel = 1; i_excel <= 20; i_excel++)
                    {
                        worksheet.Columns[i_excel].ColumnWidth = 20;
                    }
                    wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wb.Close();
                    app.Quit();
                    Marshal.ReleaseComObject(app);

                    //if (File.Exists(_fy_folder_path_result))
                    //{
                    //    File.Delete(_fy_folder_path_result);
                    //}
                }
            }
            else if (selected_index == 2)
            {
                // Bet Record
                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record"))
                {
                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record");
                }

                _fy_filename = "FY_BetRecord_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record\\FY_BetRecord_" + _fy_current_datetime.ToString() + "_" + replace + ".txt";
                _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record\\FY_BetRecord_" + _fy_current_datetime.ToString() + "_" + replace + ".xlsx";
                _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Record\\";

                if (File.Exists(_fy_folder_path_result))
                {
                    File.Delete(_fy_folder_path_result);
                }

                if (File.Exists(_fy_folder_path_result_xlsx))
                {
                    File.Delete(_fy_folder_path_result_xlsx);
                }

                //_fy_csv.ToString().Reverse();
                File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet worksheet = wb.ActiveSheet;
                worksheet.Activate();
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                firstRow.AutoFilter(1,
                                    Type.Missing,
                                    Excel.XlAutoFilterOperator.xlAnd,
                                    Type.Missing,
                                    true);
                //worksheet.Columns[3].Replace(" ", "");
                //worksheet.Columns[3].NumberFormat = "@";
                //worksheet.Columns[2].NumberFormat = "MMM-yy";
                //worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                //worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                Excel.Range usedRange = worksheet.UsedRange;
                Excel.Range rows = usedRange.Rows;
                int count = 0;
                foreach (Excel.Range row in rows)
                {
                    if (count == 0)
                    {
                        Excel.Range firstCell = row.Cells[1];

                        string firstCellValue = firstCell.Value as String;

                        if (!string.IsNullOrEmpty(firstCellValue))
                        {
                            row.Interior.Color = Color.FromArgb(222, 30, 112);
                            row.Font.Color = Color.FromArgb(255, 255, 255);
                            row.Font.Bold = true;
                            row.Font.Size = 12;
                        }

                        break;
                    }

                    count++;
                }
                int i_excel;
                for (i_excel = 1; i_excel <= 20; i_excel++)
                {
                    worksheet.Columns[i_excel].ColumnWidth = 20;
                }
                wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();
                app.Quit();
                Marshal.ReleaseComObject(app);
                                
                SaveAsTurnOver_FY();
                display_count_turnover_fy = 0;
                //if (File.Exists(_fy_folder_path_result))
                //{
                //    File.Delete(_fy_folder_path_result);
                //}
            }

            //if (File.Exists(_fy_folder_path_result))
            //{
            //    File.Delete(_fy_folder_path_result);
            //}

            _fy_csv.Clear();

            //FYHeader();

            Invoke(new Action(async () =>
            {
                //label_fy_finish_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                //timer_fy.Stop();
                button_fy_proceed.Visible = true;
                label_fy_locatefolder.Visible = true;
                label_fy_status.ForeColor = Color.FromArgb(34, 139, 34);

                if (selected_index == 0)
                {
                    if (!_isSecondRequestFinish_fy)
                    {
                        if (!_isSecondRequest_fy)
                        {
                            // Deposit Record
                            label_fy_status.Text = "status: done --- DEPOSIT RECORD";
                            _isSecondRequest_fy = true;

                            button_fy_proceed.PerformClick();

                            fy_datetime.Clear();
                            fy_gettotal.Clear();
                            fy_gettotal_test.Clear();

                            string start_datetime = dateTimePicker_start_fy.Text;
                            DateTime start = DateTime.Parse(start_datetime);

                            string end_datetime = dateTimePicker_end_fy.Text;
                            DateTime end = DateTime.Parse(end_datetime);

                            string result_start = start.ToString("yyyy-MM-dd");
                            string result_end = end.ToString("yyyy-MM-dd");
                            string result_start_time = start.ToString("HH:mm:ss");
                            string result_end_time = end.ToString("HH:mm:ss");

                            if (result_start != result_end)
                            {
                                string end_get = "";
                                int i = 0;
                                while (result_start != result_end)
                                {
                                    end_get = end.AddDays(-i).ToString("yyyy-MM-dd");
                                    if (result_start == end_get)
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_start_time;
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);

                                        break;
                                    }
                                    else
                                    {
                                        if (i == 0)
                                        {
                                            string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                            string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_end_time;
                                            fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                        }
                                        else
                                        {
                                            string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                            string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                            fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                        }
                                    }

                                    i++;
                                }
                            }
                            else
                            {
                                fy_datetime.Add(start_datetime + "*|*" + end_datetime);
                            }

                            //_fy_current_datetime = "";
                            //label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                            //_fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                            //timer_fy.Start();
                            webBrowser_fy.Stop();
                            //timer_fy_start.Stop();
                            button_fy_start.Visible = false;
                            pictureBox_fy_loader.Visible = true;
                            panel_fy_filter.Enabled = false;
                            panel_fy_status.Visible = true;

                            await GetDataFYAsync();

                            if (!_fy_no_result)
                            {
                                FYAsync();
                            }
                        }
                        else
                        {
                            // Manual Deposit Record
                            label_fy_status.Text = "status: done --- M-DEPOSIT RECORD";
                            _isSecondRequest_fy = false;
                            _isSecondRequestFinish_fy = true;

                            button_fy_proceed.PerformClick();

                            fy_datetime.Clear();
                            fy_gettotal.Clear();
                            fy_gettotal_test.Clear();

                            string start_datetime = dateTimePicker_start_fy.Text;
                            DateTime start = DateTime.Parse(start_datetime);

                            string end_datetime = dateTimePicker_end_fy.Text;
                            DateTime end = DateTime.Parse(end_datetime);

                            string result_start = start.ToString("yyyy-MM-dd");
                            string result_end = end.ToString("yyyy-MM-dd");
                            string result_start_time = start.ToString("HH:mm:ss");
                            string result_end_time = end.ToString("HH:mm:ss");

                            if (result_start != result_end)
                            {
                                string end_get = "";
                                int i = 0;
                                while (result_start != result_end)
                                {
                                    end_get = end.AddDays(-i).ToString("yyyy-MM-dd");
                                    if (result_start == end_get)
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_start_time;
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);

                                        break;
                                    }
                                    else
                                    {
                                        if (i == 0)
                                        {
                                            string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                            string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_end_time;
                                            fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                        }
                                        else
                                        {
                                            string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                            string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                            fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                        }
                                    }

                                    i++;
                                }
                            }
                            else
                            {
                                fy_datetime.Add(start_datetime + "*|*" + end_datetime);
                            }

                            //_fy_current_datetime = "";
                            //label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                            //_fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                            //timer_fy.Start();
                            webBrowser_fy.Stop();
                            //timer_fy_start.Stop();
                            button_fy_start.Visible = false;
                            pictureBox_fy_loader.Visible = true;
                            panel_fy_filter.Enabled = false;
                            panel_fy_status.Visible = true;

                            await GetDataFYAsync();

                            if (!_fy_no_result)
                            {
                                FYAsync();
                            }
                        }
                    }
                    else
                    {
                        if (!_isThirdRequest_fy)
                        {
                            // Withdrawal Record
                            label_fy_status.Text = "status: done --- WITHDRAWAL RECORD";
                            _isThirdRequest_fy = true;

                            button_fy_proceed.PerformClick();

                            fy_datetime.Clear();
                            fy_gettotal.Clear();
                            fy_gettotal_test.Clear();

                            string start_datetime = dateTimePicker_start_fy.Text;
                            DateTime start = DateTime.Parse(start_datetime);

                            string end_datetime = dateTimePicker_end_fy.Text;
                            DateTime end = DateTime.Parse(end_datetime);

                            string result_start = start.ToString("yyyy-MM-dd");
                            string result_end = end.ToString("yyyy-MM-dd");
                            string result_start_time = start.ToString("HH:mm:ss");
                            string result_end_time = end.ToString("HH:mm:ss");

                            if (result_start != result_end)
                            {
                                string end_get = "";
                                int i = 0;
                                while (result_start != result_end)
                                {
                                    end_get = end.AddDays(-i).ToString("yyyy-MM-dd");
                                    if (result_start == end_get)
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_start_time;
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);

                                        break;
                                    }
                                    else
                                    {
                                        if (i == 0)
                                        {
                                            string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                            string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_end_time;
                                            fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                        }
                                        else
                                        {
                                            string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                            string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                            fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                        }
                                    }

                                    i++;
                                }
                            }
                            else
                            {
                                fy_datetime.Add(start_datetime + "*|*" + end_datetime);
                            }

                            //_fy_current_datetime = "";
                            //label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                            //_fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                            //timer_fy.Start();
                            webBrowser_fy.Stop();
                            //timer_fy_start.Stop();
                            button_fy_start.Visible = false;
                            pictureBox_fy_loader.Visible = true;
                            panel_fy_filter.Enabled = false;
                            panel_fy_status.Visible = true;

                            await GetDataFYAsync();

                            if (!_fy_no_result)
                            {
                                FYAsync();
                            }
                        }
                        else
                        {
                            // Manual Withdrawal Record
                            label_fy_status.Text = "status: done --- M-WITHDRAWAL RECORD";
                            _isThirdRequest_fy = false;
                            _isSecondRequestFinish_fy = false;

                            // asd textBox2.Text = _fy_folder_path_result;

                            // Database Member List FY
                            // asd comment
                            InsertPaymentReport_FY(_fy_folder_path_result);
                            display_count_fy = 0;
                        }
                    }
                }
                else if (selected_index == 1)
                {
                    if (!_isSecondRequest_fy)
                    {
                        // Manual Bonus Report
                        label_fy_status.Text = "status: done --- M-BONUS RECORD";
                        _isSecondRequest_fy = true;

                        button_fy_proceed.PerformClick();

                        fy_datetime.Clear();
                        fy_gettotal.Clear();
                        fy_gettotal_test.Clear();

                        string start_datetime = dateTimePicker_start_fy.Text;
                        DateTime start = DateTime.Parse(start_datetime);

                        string end_datetime = dateTimePicker_end_fy.Text;
                        DateTime end = DateTime.Parse(end_datetime);

                        string result_start = start.ToString("yyyy-MM-dd");
                        string result_end = end.ToString("yyyy-MM-dd");
                        string result_start_time = start.ToString("HH:mm:ss");
                        string result_end_time = end.ToString("HH:mm:ss");

                        if (result_start != result_end)
                        {
                            string end_get = "";
                            int i = 0;
                            while (result_start != result_end)
                            {
                                end_get = end.AddDays(-i).ToString("yyyy-MM-dd");
                                if (result_start == end_get)
                                {
                                    string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_start_time;
                                    string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                    fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);

                                    break;
                                }
                                else
                                {
                                    if (i == 0)
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_end_time;
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                    }
                                    else
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                    }
                                }

                                i++;
                            }
                        }
                        else
                        {
                            fy_datetime.Add(start_datetime + "*|*" + end_datetime);
                        }

                        //_fy_current_datetime = "";
                        //label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                        //_fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                        //timer_fy.Start();
                        webBrowser_fy.Stop();
                        //timer_fy_start.Stop();
                        button_fy_start.Visible = false;
                        pictureBox_fy_loader.Visible = true;
                        panel_fy_filter.Enabled = false;
                        panel_fy_status.Visible = true;

                        await GetDataFYAsync();

                        if (!_fy_no_result)
                        {
                            FYAsync();
                        }
                    }
                    else
                    {
                        // Generated Bonus Report
                        label_fy_status.Text = "status: done --- G-BONUS RECORD";
                        _isSecondRequest_fy = false;

                        // Database Bonus Report FY
                        // asd comment
                        InsertBonusReport_FY(_fy_folder_path_result);
                        display_count_fy = 0;
                    }
                }
                else if (selected_index == 2)
                {
                    // Bet Record
                    label_fy_status.Text = "status: done --- BET RECORD";
                    
                    // Database Bet Record FY
                    // asd comment
                    isBetRecordInsert = true;
                    InsertBetRecord_FY(_fy_folder_path_result);
                    display_count_fy = 0;
                }

            }));

            //var notification = new NotifyIcon()
            //{
            //    Visible = true,
            //    Icon = SystemIcons.Information,
            //    BalloonTipIcon = ToolTipIcon.Info,
            //    BalloonTipTitle = "FY BET RECORD DONE",
            //    BalloonTipText = "Filter of...\nStart Time: " + dateTimePicker_start_fy.Text + "\nEnd Time: " + dateTimePicker_end_fy.Text + "\n\nStart-Finish...\nStart Time: " + label_start_fy.Text + "\nFinish Time: " + label_end_fy.Text,
            //};

            //notification.ShowBalloonTip(1000);

            //timer_fy_start.Start();
        }

        private void timer_fy_Tick(object sender, EventArgs e)
        {
            string start_datetime = _fy_start_datetime;
            DateTime start = DateTime.ParseExact(start_datetime, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            string finish_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            DateTime finish = DateTime.ParseExact(finish_datetime, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            TimeSpan span = finish.Subtract(start);

            if (span.Hours == 0 && span.Minutes == 0)
            {
                label_fy_elapsed.Text = span.Seconds + " sec(s)";
            }
            else if (span.Hours != 0)
            {
                label_fy_elapsed.Text = span.Hours + " hr(s) " + span.Minutes + " min(s) " + span.Seconds + " sec(s)";
            }
            else if (span.Minutes != 0)
            {
                label_fy_elapsed.Text = span.Minutes + " min(s) " + span.Seconds + " sec(s)";
            }
            else
            {
                label_fy_elapsed.Text = span.Seconds + " sec(s)";
            }
        }

        private void timer_fy_detect_inserted_in_excel_Tick(object sender, EventArgs e)
        {
            // status
            label_fy_status.Text = "status: inserting data to excel...";

            if (_fy_inserted_in_excel)
            {
                FYAsync();
                timer_fy_detect_inserted_in_excel.Stop();
            }
        }

        public class FY_BetRecord
        {
            public string GAME_PLATFORM { get; set; }
            public string USERNAME { get; set; }
            public string BET_NO { get; set; }
            public string BET_TIME { get; set; }
            public string BET_TYPE { get; set; }
            public string GAME_RESULT { get; set; }
            public double STAKE_AMOUNT { get; set; }
            public double WIN_AMOUNT { get; set; }
            public double COMPANY_WIN_LOSS { get; set; }
            public double VALID_BET { get; set; }
            public string VALID_INVALID { get; set; }
        }

        private void timer_fy_start_Tick(object sender, EventArgs e)
        {
            webBrowser_fy.Navigate("http://cs.ying168.bet/player/listAjax1");
        }

        private async void button_fy_start_ClickAsync(object sender, EventArgs e)
        {
            isButtonStart_fy = true;
            panel_fy_filter.Enabled = false;
            label_fy_start_datetime.Text = "-";
            fy_datetime.Clear();
            fy_gettotal.Clear();
            fy_gettotal_test.Clear();

            string start_datetime = dateTimePicker_start_fy.Text;
            DateTime start = DateTime.Parse(start_datetime);

            string end_datetime = dateTimePicker_end_fy.Text;
            DateTime end = DateTime.Parse(end_datetime);

            string result_start = start.ToString("yyyy-MM-dd");
            string result_end = end.ToString("yyyy-MM-dd");
            string result_start_time = start.ToString("HH:mm:ss");
            string result_end_time = end.ToString("HH:mm:ss");

            if (start < end)
            {
                button_fy_stop.Visible = true;
                button_fy_start.Visible = false;
                timer_fy_count = 10;
                label_fy_count.Text = timer_fy_count.ToString();
                timer_fy_count = 9;
                label_fy_count.Visible = true;
                timer_fy_start_button.Start();
            }
            else
            {
                _fy_no_result = true;
                MessageBox.Show("No data found.", "FY Cronos Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                panel_fy_filter.Enabled = true;
            }
        }

        private void button_fy_stop_Click(object sender, EventArgs e)
        {
            panel_fy_filter.Enabled = true;
            button_fy_stop.Visible = false;
            button_fy_start.Visible = true;
            timer_fy_count = 10;
            label_fy_count.Visible = false;
            timer_fy_start_button.Stop();
            isStopClick_fy = true;
            label_status.Text = "Stop";
        }

        int timer_fy_count = 10;
        private async void timer_fy_start_button_TickAsync(object sender, EventArgs e)
        {
            if (__isLogin)
            {
                label_fy_count.Text = timer_fy_count--.ToString();
                if (label_fy_count.Text == "0")
                {
                    label_status.Text = "Running";
                    label_fy_status.Visible = true;
                    label_fy_page_count_1.Visible = true;
                    label_fy_total_records_1.Visible = true;
                    label_fy_page_count.Visible = true;
                    label_fy_currentrecord.Visible = true;
                    panel_fy_datetime.Visible = true;

                    timer_fy_start_button.Stop();
                    label_fy_count.Visible = false;
                    button_fy_stop.Visible = false;

                    fy_datetime.Clear();
                    fy_gettotal.Clear();
                    fy_gettotal_test.Clear();

                    string start_datetime = dateTimePicker_start_fy.Text;
                    DateTime start = DateTime.Parse(start_datetime);

                    string end_datetime = dateTimePicker_end_fy.Text;
                    DateTime end = DateTime.Parse(end_datetime);

                    string result_start = start.ToString("yyyy-MM-dd");
                    string result_end = end.ToString("yyyy-MM-dd");
                    string result_start_time = start.ToString("HH:mm:ss");
                    string result_end_time = end.ToString("HH:mm:ss");

                    int selected_index = comboBox_fy_list.SelectedIndex;
                    if (selected_index == 0)
                    {
                        // Deposit Record
                        if (result_start != result_end)
                        {
                            string end_get = "";
                            int i = 0;
                            while (result_start != result_end)
                            {
                                end_get = end.AddDays(-i).ToString("yyyy-MM-dd");
                                if (result_start == end_get)
                                {
                                    string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_start_time;
                                    string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                    fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);

                                    break;
                                }
                                else
                                {
                                    if (i == 0)
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_end_time;
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                    }
                                    else
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                    }
                                }

                                i++;
                            }
                        }
                        else
                        {
                            fy_datetime.Add(start_datetime + "*|*" + end_datetime);
                        }

                        _fy_current_datetime = "";
                        label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                        _fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                        timer_fy.Start();
                        webBrowser_fy.Stop();
                        //timer_fy_start.Stop();
                        button_fy_start.Visible = false;
                        pictureBox_fy_loader.Visible = true;
                        panel_fy_filter.Enabled = false;
                        panel_fy_status.Visible = true;

                        await GetDataFYAsync();

                        if (!_fy_no_result)
                        {
                            FYAsync();
                        }
                    }
                    else if (selected_index == 1)
                    {
                        // Bonus Report
                        if (result_start != result_end)
                        {
                            string end_get = "";
                            int i = 0;
                            while (result_start != result_end)
                            {
                                end_get = end.AddDays(-i).ToString("yyyy-MM-dd");
                                if (result_start == end_get)
                                {
                                    string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_start_time;
                                    string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                    fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);

                                    break;
                                }
                                else
                                {
                                    if (i == 0)
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_end_time;
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                    }
                                    else
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                    }
                                }

                                i++;
                            }
                        }
                        else
                        {
                            fy_datetime.Add(start_datetime + "*|*" + end_datetime);
                        }

                        _fy_current_datetime = "";
                        label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                        _fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                        timer_fy.Start();
                        webBrowser_fy.Stop();
                        //timer_fy_start.Stop();
                        button_fy_start.Visible = false;
                        pictureBox_fy_loader.Visible = true;
                        panel_fy_filter.Enabled = false;
                        panel_fy_status.Visible = true;

                        await GetDataFYAsync();

                        if (!_fy_no_result)
                        {
                            FYAsync();
                        }
                    }
                    else if (selected_index == 2)
                    {
                        string path_turnover = Path.Combine(Path.GetTempPath(), "FY Turnover.txt");

                        if (FY_Cronos_Data.Properties.Settings.Default.______turnover_count == "")
                        {
                            FY_Cronos_Data.Properties.Settings.Default.______turnover_count = "0";
                            FY_Cronos_Data.Properties.Settings.Default.Save();
                        }

                        if (FY_Cronos_Data.Properties.Settings.Default.______turnover_count == "0")
                        {
                            __turnover_detect = false;
                            if (File.Exists(path_turnover))
                            {
                                File.Delete(path_turnover);
                            }
                        }
                        else
                        {
                            __turnover_detect = true;
                        }

                        // Bet Record
                        if (result_start != result_end)
                        {
                            string end_get = "";
                            int i = 0;
                            while (result_start != result_end)
                            {
                                end_get = end.AddDays(-i).ToString("yyyy-MM-dd");
                                if (result_start == end_get)
                                {
                                    string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_start_time;
                                    string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                    fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);

                                    break;
                                }
                                else
                                {
                                    if (i == 0)
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_end_time;
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                    }
                                    else
                                    {
                                        string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                        string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                        fy_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                                    }
                                }

                                i++;
                            }
                        }
                        else
                        {
                            fy_datetime.Add(start_datetime + "*|*" + end_datetime);
                        }

                        _fy_current_datetime = "";
                        label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                        _fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                        timer_fy.Start();
                        webBrowser_fy.Stop();
                        //timer_fy_start.Stop();
                        button_fy_start.Visible = false;
                        pictureBox_fy_loader.Visible = true;
                        panel_fy_filter.Enabled = false;
                        panel_fy_status.Visible = true;

                        await GetDataFYAsync();

                        if (!_fy_no_result)
                        {
                            FYAsync();
                        }
                    }
                }
            }
            else
            {
                button_fy_stop.PerformClick();
            }
        }

        private void button_fy_proceed_Click(object sender, EventArgs e)
        {
            label_fy_insert.Visible = false;
            panel_fy_status.Visible = false;
            button_fy_start.Visible = true;
            panel_fy_filter.Enabled = true;
            //button_filelocation.Enabled = true;

            button_fy_proceed.Visible = false;
            label_fy_locatefolder.Visible = false;

            label_fy_page_count.Text = "-";
            label_fy_currentrecord.Text = "-";
            label_fy_finish_datetime.Text = "-";
            label_fy_elapsed.Text = "-";

            panel_fy_datetime.Location = new Point(66, 226);

            // set default variables
            _fy_bet_records.Clear();
            fy_datetime.Clear();
            fy_gettotal.Clear();
            _total_records_fy = 0;
            _total_page_fy = 0;
            _fy_displayinexel_i = 0;
            _fy_pages_count_display = 0;
            _fy_pages_count = 0;
            _detect_fy = false;
            _fy_inserted_in_excel = true;
            _fy_row = 1;
            _fy_row_count = 1;
            _isDone_fy = false;
            _fy_secho = 0;
            _fy_i = 0;
            _fy_ii = 0;
            _fy_get_ii = 1;
            _fy_get_ii_display = 1;
            _fy_folder_path_result = "";
            _fy_folder_path_result_locate = "";
            _fy_current_index = 1;
            _fy_csv.Clear();

            // MEMBER LIST
            _fy_playerlist_cn = "";
            _fy_playerlist_ea = "";
            _fy_id_playerlist = "";
            
            if (label_fy_status.Text.Contains("MEMBER LIST"))
            {
                label_fy_status.Text = "-";
                panel_fy_filter.Visible = true;
                button_fy_start.Enabled = true;
                button_fy_start.PerformClick();
            }
        }

        private void label_fy_locatefolder_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(_fy_folder_path_result_locate);
            }
            catch (Exception err)
            {
                MessageBox.Show("Can't locate folder.", "FY Cronos Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Get Insert
        
        private void InsertPaymentReport_FY(string path)
        {
            button_fy_proceed.Text = "SENDING...";
            button_fy_proceed.Enabled = false;
            WindowState = FormWindowState.Normal;
            label_fy_locatefolder.Enabled = false;
            label_fy_insert.Visible = true;

            try
            {
                var lines = File.ReadAllLines(path);
                if (lines.Count() == 0) return;
                var table = new DataTable();
                string header = "Month,Date,Submitted Date,Member,Amount,Payment Type,Status,Updated Date,Transaction Time,Duration Time,VIP,Transaction Type,Transaction ID,Time,Brand,PG Company,PG Type,Retained,FD Date,New,Reactivated,File Name";
                var columns = header.Split(',');
                foreach (var c in columns)
                {
                    table.Columns.Add(c);
                }

                for (int i = 1; i < lines.Count(); i++)
                {
                    if (lines[i].Length > 0)
                    {
                        Application.DoEvents();
                        display_count_fy++;
                        label_fy_insert.Text = display_count_fy.ToString("N0");

                        string[] get_column = lines[i].Split(",\"");

                        DataRow row = table.NewRow();
                        // Month
                        DateTime month = DateTime.ParseExact(get_column[1].Replace("\"", ""), "MM/dd/yyyy", CultureInfo.InvariantCulture);
                        row[0] = month.ToString("yyyy-MM-dd 00:00:00");
                        // Date
                        row[1] = get_column[2].Replace("\"", "") + " 00:00:00";
                        // Submitted Date
                        row[2] = get_column[3].Replace("\"", "");
                        // Member
                        row[3] = get_column[5].Replace("\"", "");
                        // Amount
                        row[4] = get_column[10].Replace("\"", "");
                        // Payment Type
                        row[5] = get_column[6].Replace("\"", "");
                        // Status
                        row[6] = get_column[15].Replace("\"", "");
                        // Updated Date
                        row[7] = get_column[4].Replace("\"", "");
                        // Transaction Time
                        row[8] = "1899-12-30 " + get_column[11].Replace("\"", "");
                        // Duration Time
                        row[9] = get_column[13].Replace("\"", "");
                        // VIP
                        row[10] = get_column[14].Replace("\"", "");
                        // Transaction Type
                        row[11] = get_column[12].Replace("\"", "");
                        // Transaction ID
                        get_column[9] = get_column[9].Replace("\"", "");
                        get_column[9] = get_column[9].Replace("'", "");
                        row[12] = get_column[9];
                        // Time
                        row[13] = DBNull.Value;
                        // Brand
                        row[14] = "FY";
                        // PG Company
                        row[15] = get_column[7].Replace("\"", "");
                        // PG Type
                        row[16] = get_column[8].Replace("\"", "");
                        // Retained
                        if (get_column[16].Replace("\"", "").Trim() != "")
                        {
                            row[17] = get_column[16].Replace("\"", "");
                        }
                        else
                        {
                            row[17] = DBNull.Value;
                        }
                        // FD Date
                        if (get_column[17].Replace("\"", "") != "" && get_column[17].Replace("\"", "") != "fd date")
                        {
                            DateTime fd_date_ = DateTime.ParseExact(get_column[17].Replace("\"", ""), "MM/dd/yyyy", CultureInfo.InvariantCulture);
                            row[18] = fd_date_.ToString("yyyy-MM-dd 00:00:00");
                        }
                        else
                        {
                            row[18] = DBNull.Value;
                        }
                        // New
                        if (get_column[18].Replace("\"", "").Trim() != "")
                        {
                            row[19] = get_column[18].Replace("\"", "");
                        }
                        else
                        {
                            row[19] = DBNull.Value;
                        }
                        // Reactivated
                        if (get_column[19].Replace("\"", "").Trim() != "")
                        {
                            row[20] = get_column[19].Replace("\"", "");
                        }
                        else
                        {
                            row[20] = DBNull.Value;
                        }
                        // File Name
                        row[21] = path.Replace(".txt", ".xlsx");
                        
                        table.Rows.Add(row);
                    }
                }

                var connection = @"Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";
                var sqlBulk = new SqlBulkCopy(connection);
                sqlBulk.DestinationTableName = "[testrain].[dbo].[FY.Payment Logs]";
                sqlBulk.WriteToServer(table);
                
                label_fy_finish_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                timer_fy.Stop();
                panel_fy_datetime.Location = new Point(5, 226);
                pictureBox_fy_loader.Visible = false;

                button_fy_proceed.Text = "PROCEED";
                button_fy_proceed.Enabled = true;
                label_fy_locatefolder.Enabled = true;

                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                // added auto
                if (!isStopClick_fy)
                {
                    // bonus
                    FY_Cronos_Data.Properties.Settings.Default.______start_detect = "2";
                    FY_Cronos_Data.Properties.Settings.Default.Save();
                    
                    button_fy_proceed.PerformClick();
                    comboBox_fy_list.SelectedIndex = 1;
                    button_fy_start.PerformClick();
                }
            }
            catch (Exception err)
            {
                button_fy_proceed.Text = "PROCEED";
                button_fy_proceed.Enabled = true;
                label_fy_locatefolder.Enabled = true;
                
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                __send_email = 0;

                isClose = false;
                Environment.Exit(0);
            }
        }

        private void InsertBonusReport_FY(string path)
        {
            button_fy_proceed.Text = "SENDING...";
            button_fy_proceed.Enabled = false;
            WindowState = FormWindowState.Normal;
            label_fy_locatefolder.Enabled = false;
            label_fy_insert.Visible = true;

            try
            {
                var lines = File.ReadAllLines(path);
                if (lines.Count() == 0) return;
                var table = new DataTable();
                string header = "Month,Date,Username,VIP,Amount,Bonus Category,Purpose,Transaction ID,Updated by,Bonus Code,Transaction Time,Product,Brand,File Name";
                var columns = header.Split(',');
                foreach (var c in columns)
                {
                    table.Columns.Add(c);
                }

                for (int i = 1; i < lines.Count(); i++)
                {
                    if (lines[i].Length > 0)
                    {
                        Application.DoEvents();
                        display_count_fy++;
                        label_fy_insert.Text = display_count_fy.ToString("N0");

                        string[] get_column = lines[i].Split(",\"");

                        DataRow row = table.NewRow();
                        // Month
                        DateTime month = DateTime.ParseExact(get_column[1].Replace("\"", ""), "MM/dd/yyyy", CultureInfo.InvariantCulture);
                        row[0] = month.ToString("yyyy-MM-dd 00:00:00");
                        // Date
                        row[1] = get_column[2].Replace("\"", "") + " 00:00:00";
                        // Username
                        row[2] = get_column[3].Replace("\"", "");
                        // VIP
                        row[3] = get_column[8].Replace("\"", "");
                        // Amount
                        row[4] = get_column[6].Replace("\"", "");
                        // Bonus Category
                        row[5] = get_column[4].Replace("\"", "");
                        // Purpose
                        row[6] = get_column[5].Replace("\"", "");
                        // Transaction ID
                        row[7] = DBNull.Value;
                        // Updated By
                        row[8] = "System";
                        // Bonus Code
                        get_column[7] = get_column[7].Replace("\"", "");
                        get_column[7] = get_column[7].Replace(";", "");
                        row[9] = get_column[7];
                        // Transaction Time
                        row[10] = DBNull.Value;
                        // Product
                        row[11] = "Main Wallet";
                        // Brand
                        row[12] = "FY";
                        // File Name
                        row[13] = path.Replace(".txt", ".xlsx");

                        table.Rows.Add(row);
                    }
                }

                var connection = @"Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";
                var sqlBulk = new SqlBulkCopy(connection);
                sqlBulk.DestinationTableName = "[testrain].[dbo].[FY.Bonus Report]";
                sqlBulk.WriteToServer(table);

                label_fy_finish_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                timer_fy.Stop();
                panel_fy_datetime.Location = new Point(5, 226);
                pictureBox_fy_loader.Visible = false;

                button_fy_proceed.Text = "PROCEED";
                button_fy_proceed.Enabled = true;
                label_fy_locatefolder.Enabled = true;

                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                // added auto
                if (!isStopClick_fy)
                {
                    // bet record
                    FY_Cronos_Data.Properties.Settings.Default.______start_detect = "3";
                    FY_Cronos_Data.Properties.Settings.Default.Save();
                    
                    button_fy_proceed.PerformClick();
                    comboBox_fy_list.SelectedIndex = 2;
                    button_fy_start.PerformClick();
                }
            }
            catch (Exception err)
            {
                button_fy_proceed.Text = "PROCEED";
                button_fy_proceed.Enabled = true;
                label_fy_locatefolder.Enabled = true;

                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                __send_email = 0;

                isClose = false;
                Environment.Exit(0);
            }
        }

        private void InsertBetRecord_FY(string path)
        {
            button_fy_proceed.Text = "SENDING...";
            button_fy_proceed.Enabled = false;
            WindowState = FormWindowState.Normal;
            label_fy_locatefolder.Enabled = false;
            label_fy_insert.Visible = true;
            string sdsdsdsds = "";
            try
            {
                var lines = File.ReadAllLines(path);
                if (lines.Count() == 0) return;
                var table = new DataTable();
                string header = "Date,Category,Platform,Username,Bet No,Bet Time,Game,Settlement,VIP,Bet Amount,Payout,Company WL,Turnover,Status,File Name";
                var columns = header.Split(',');
                foreach (var c in columns)
                {
                    table.Columns.Add(c);
                }

                for (int i = 1; i < lines.Count(); i++)
                {
                    if (lines[i].Length > 0)
                    {
                        Application.DoEvents();
                        display_count_fy++;
                        label_fy_insert.Text = display_count_fy.ToString("N0");
                        string[] get_column = lines[i].Split(",\"");

                        string category_get = "";
                        string gameplatform_temp = Path.Combine(Path.GetTempPath(), "FY Game Platform Code.txt");
                        if (File.Exists(gameplatform_temp))
                        {
                            using (StreamReader sr = File.OpenText(gameplatform_temp))
                            {
                                string s = String.Empty;
                                while ((s = sr.ReadLine()) != null)
                                {
                                    int gameplatform_i = 0;
                                    string[] results = s.Split("*|*");
                                    foreach (string result in results)
                                    {
                                        Application.DoEvents();
                                        gameplatform_i++;

                                        if (gameplatform_i == 1)
                                        {
                                            if (result.Trim() == get_column[3].Replace("\"", "").Trim())
                                            {
                                                int memberlist_i_inner = 0;
                                                string[] results_inner = s.Split("*|*");
                                                foreach (string result_inner in results_inner)
                                                {
                                                    Application.DoEvents();
                                                    memberlist_i_inner++;

                                                    if (memberlist_i_inner == 4)
                                                    {
                                                        category_get = result_inner;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        DataRow row = table.NewRow();
                        // Date
                        row[0] = get_column[1].Replace("\"", "") + " 00:00:00";
                        // Category
                        row[1] = category_get;
                        // Platform
                        row[2] = get_column[3].Replace("\"", "");
                        // Username
                        row[3] = get_column[4].Replace("\"", "");
                        // Bet Number
                        get_column[5] = get_column[5].Replace("\"", "");
                        get_column[5] = get_column[5].Replace("'", "");
                        if (IsDigitsOnly(get_column[5]))
                        {
                            row[4] = get_column[5].Trim();
                            sdsdsdsds = get_column[5];
                        }
                        else
                        {
                            row[4] = DBNull.Value;
                        }
                        // Bet Time
                        row[5] = get_column[6].Replace("\"", "");
                        // Game
                        row[6] = get_column[7].Replace("\"", "");
                        // Settlement
                        row[7] = get_column[8].Replace("\"", "");
                        // VIP
                        row[8] = get_column[2].Replace("\"", "");
                        // Bet Amount
                        row[9] = get_column[9].Replace("\"", "");
                        // Payout
                        row[10] = get_column[10].Replace("\"", "");
                        // Company WL
                        row[11] = get_column[11].Replace("\"", "");
                        // Turnover
                        row[12] = get_column[12].Replace("\"", "");
                        // Status
                        row[13] = get_column[13].Replace("\"", "");
                        // File Name
                        row[14] = path.Replace(".txt", ".xlsx");
                        
                        table.Rows.Add(row);
                    }
                }

                var connection = @"Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";
                var sqlBulk = new SqlBulkCopy(connection);
                sqlBulk.DestinationTableName = "[testrain].[dbo].[FY.Bet Record]";
                sqlBulk.WriteToServer(table);

                if (isBetRecordInsert)
                {
                    label_fy_finish_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                    timer_fy.Stop();
                    panel_fy_datetime.Location = new Point(5, 226);
                    pictureBox_fy_loader.Visible = false;

                    button_fy_proceed.Text = "PROCEED";
                    button_fy_proceed.Enabled = true;
                    label_fy_locatefolder.Enabled = true;

                    button_fy_proceed.PerformClick();
                    panel_fy_filter.Enabled = false;
                    button_fy_start.Enabled = false;
                    
                    // done
                    FY_Cronos_Data.Properties.Settings.Default.______start_detect = "0";
                    FY_Cronos_Data.Properties.Settings.Default.Save();
                    
                    comboBox_fy_list.SelectedIndex = 0;
                    label_status.Text = "Waiting";
                    _fy_current_datetime = "";

                    FY_Cronos_Data.Properties.Settings.Default.______turnover_count = "0";
                    FY_Cronos_Data.Properties.Settings.Default.Save();

                    SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>Message: <b>Reports has been completed.</b></body></html>");
                    SendEmail3("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>Message: <b>Reports has been completed.</b></body></html>");
                    __send_email = 0;
                }

                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception err)
            {
                button_fy_proceed.Text = "PROCEED";
                button_fy_proceed.Enabled = true;
                label_fy_locatefolder.Enabled = true;

                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                __send_email = 0;

                isClose = false;
                Environment.Exit(0);
            }
        }

        bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }

            return true;
        }

        private void GetMemberList_FY()
        {
            string path_deposit = Path.Combine(Path.GetTempPath(), "FY Registration Deposit.txt");
            if (File.Exists(path_deposit))
            {
                File.Delete(path_deposit);
            }

            try
            {
                string connection = "Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";

                using (SqlConnection conn = new SqlConnection(connection))
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM [testrain].[dbo].[FY.Registration Report]", conn);
                    SqlCommand command_count = new SqlCommand("SELECT COUNT(*) FROM [testrain].[dbo].[FY.Registration Report]", conn);
                    string columns_deposit = "";
                    string column_reg = "";

                    Int32 getcount = (Int32)command_count.ExecuteScalar();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            count++;
                            label_getdatacount_fy.Text = "Member List: " + count.ToString("N0") + " of " + getcount.ToString("N0");

                            Application.DoEvents();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Application.DoEvents();

                                if (i == 0)
                                {
                                    getmemberlist_fy.Add(reader[i].ToString());
                                    __getcontactemail.Add(reader[i].ToString());
                                    columns_deposit += reader[i].ToString() + "*|*";
                                }
                                else if (i == 2)
                                {
                                    // Date Registration
                                    string[] date_registration_get_results = reader[i].ToString().Split("/");
                                    int count_reg = 0;
                                    foreach (string first_deposit_get_result in date_registration_get_results)
                                    {
                                        Application.DoEvents();

                                        count_reg++;

                                        if (count_reg == 1)
                                        {
                                            // Month
                                            if (first_deposit_get_result.Length == 1)
                                            {
                                                column_reg += "0" + first_deposit_get_result + "/";
                                            }
                                            else
                                            {
                                                column_reg += first_deposit_get_result + "/";
                                            }
                                        }
                                        else if (count_reg == 2)
                                        {
                                            // Day
                                            if (first_deposit_get_result.Length == 1)
                                            {
                                                column_reg += "0" + first_deposit_get_result + "/";
                                            }
                                            else
                                            {
                                                column_reg += first_deposit_get_result + "/";
                                            }
                                        }
                                        else if (count_reg == 3)
                                        {
                                            // Year
                                            column_reg += first_deposit_get_result.Substring(0, 4);
                                        }
                                    }
                                }
                                else if (i == 4)
                                {
                                    getmemberlist_fy.Add(reader[i].ToString());
                                }
                                else if (i == 6)
                                {
                                    __getcontactemail.Add(reader[i].ToString());
                                }
                                else if (i == 7)
                                {
                                    __getcontactemail.Add(reader[i].ToString());
                                }
                                else if (i == 8)
                                {
                                    // FDD
                                    string first_deposit = "";
                                    string[] first_deposit_get_results = reader[i].ToString().Split("/");
                                    int count_reg = 0;
                                    foreach (string first_deposit_get_result in first_deposit_get_results)
                                    {
                                        Application.DoEvents();

                                        count_reg++;

                                        if (count_reg == 1)
                                        {
                                            // Month
                                            if (first_deposit_get_result.Length == 1)
                                            {
                                                first_deposit += "0" + first_deposit_get_result + "/";
                                            }
                                            else
                                            {
                                                first_deposit += first_deposit_get_result + "/";
                                            }
                                        }
                                        else if (count_reg == 2)
                                        {
                                            // Day
                                            if (first_deposit_get_result.Length == 1)
                                            {
                                                first_deposit += "0" + first_deposit_get_result + "/";
                                            }
                                            else
                                            {
                                                first_deposit += first_deposit_get_result + "/";
                                            }
                                        }
                                        else if (count_reg == 3)
                                        {
                                            // Year
                                            first_deposit += first_deposit_get_result.Substring(0, 4);
                                        }
                                    }

                                    columns_deposit += first_deposit + "*|*";
                                }
                                else if (i == 10)
                                {
                                    // LDD
                                    string last_deposit = "";
                                    string[] last_deposit_get_results = reader[i].ToString().Split("/");
                                    int count_reg = 0;
                                    foreach (string last_deposit_get_result in last_deposit_get_results)
                                    {
                                        Application.DoEvents();

                                        count_reg++;

                                        if (count_reg == 1)
                                        {
                                            // Month
                                            if (last_deposit_get_result.Length == 1)
                                            {
                                                last_deposit += "0" + last_deposit_get_result + "/";
                                            }
                                            else
                                            {
                                                last_deposit += last_deposit_get_result + "/";
                                            }
                                        }
                                        else if (count_reg == 2)
                                        {
                                            // Day
                                            if (last_deposit_get_result.Length == 1)
                                            {
                                                last_deposit += "0" + last_deposit_get_result + "/";
                                            }
                                            else
                                            {
                                                last_deposit += last_deposit_get_result + "/";
                                            }
                                        }
                                        else if (count_reg == 3)
                                        {
                                            // Year
                                            last_deposit += last_deposit_get_result.Substring(0, 4);
                                        }
                                    }

                                    columns_deposit += last_deposit + "*|*" + column_reg;
                                }
                            }

                            using (StreamWriter file = new StreamWriter(path_deposit, true, Encoding.UTF8))
                            {
                                file.WriteLine(columns_deposit);
                            }
                            columns_deposit = "";
                            column_reg = "";
                        }
                    }

                    conn.Close();
                }
            }
            catch (Exception err)
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                __send_email = 0;

                isClose = false;
                Environment.Exit(0);
            }
        }

        private void GetBonusCode_FY()
        {
            string path = Path.Combine(Path.GetTempPath(), "FY Bonus Code.txt");
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            try
            {
                string connection = "Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";

                using (SqlConnection conn = new SqlConnection(connection))
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM [testrain].[dbo].[FY.Bonus Code]", conn);
                    SqlCommand command_count = new SqlCommand("SELECT COUNT(*) FROM [testrain].[dbo].[FY.Bonus Code]", conn);
                    string columns = "";

                    Int32 getcount = (Int32)command_count.ExecuteScalar();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            count++;
                            label_getdatacount_fy.Text = "Bonus Code: " + count.ToString("N0") + " of " + getcount.ToString("N0");

                            Application.DoEvents();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Application.DoEvents();
                                columns += reader[i].ToString() + "*|*";
                            }

                            using (StreamWriter file = new StreamWriter(path, true, Encoding.UTF8))
                            {
                                file.WriteLine(columns);
                            }
                            columns = "";
                        }
                    }

                    conn.Close();
                }
            }
            catch (Exception err)
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                __send_email = 0;

                isClose = false;
                Environment.Exit(0);
            }
        }

        private void GetGamePlatform_FY()
        {
            string path = Path.Combine(Path.GetTempPath(), "FY Game Platform Code.txt");
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            try
            {
                string connection = "Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";

                using (SqlConnection conn = new SqlConnection(connection))
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM [testrain].[dbo].[FY.Game Platform Code]", conn);
                    SqlCommand command_count = new SqlCommand("SELECT COUNT(*) FROM [testrain].[dbo].[FY.Game Platform Code]", conn);
                    string columns = "";

                    Int32 getcount = (Int32)command_count.ExecuteScalar();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            count++;
                            label_getdatacount_fy.Text = "Game Platform Code: " + count.ToString("N0") + " of " + getcount.ToString("N0");

                            Application.DoEvents();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Application.DoEvents();
                                columns += reader[i].ToString() + "*|*";
                            }

                            using (StreamWriter file = new StreamWriter(path, true, Encoding.UTF8))
                            {
                                file.WriteLine(columns);
                            }
                            columns = "";
                        }
                    }

                    conn.Close();
                }
            }
            catch (Exception err)
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                __send_email = 0;

                isClose = false;
                Environment.Exit(0);
            }
        }

        private void GetPaymentType_FY()
        {
            string path = Path.Combine(Path.GetTempPath(), "FY Payment Type Code.txt");
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            try
            {
                string connection = "Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";

                using (SqlConnection conn = new SqlConnection(connection))
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM [testrain].[dbo].[FY.Payment Type Code]", conn);
                    SqlCommand command_count = new SqlCommand("SELECT COUNT(*) FROM [testrain].[dbo].[FY.Payment Type Code]", conn);
                    string columns = "";

                    Int32 getcount = (Int32)command_count.ExecuteScalar();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            count++;
                            label_getdatacount_fy.Text = "Payment Type Code: " + count.ToString("N0") + " of " + getcount.ToString("N0");

                            Application.DoEvents();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Application.DoEvents();
                                columns += reader[i].ToString() + "*|*";
                            }

                            using (StreamWriter file = new StreamWriter(path, true, Encoding.UTF8))
                            {
                                file.WriteLine(columns);
                            }
                            columns = "";
                        }
                    }

                    label_getdatacount_fy.Text = "";
                    panel_fy.Enabled = true;
                    conn.Close();
                }
            }
            catch (Exception err)
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                __send_email = 0;

                isClose = false;
                Environment.Exit(0);
            }
        }

        bool deposit_fy = false;
        String[] filecontent_deposit_fy = null;
        private int __pageCountPlayer;
        private bool __isLogin = false;
        private bool __isStart = false;

        private void Turnover_FY(string player_name, string stake_amount_get, string win_amount_get, string company_win_loss_get, string valid_bet_get, string date_get, string month_get, string vip_get, string gameplatform_get)
        {
            if (!deposit_fy)
            {
                filecontent_deposit_fy = File.ReadAllLines(Path.Combine(Path.GetTempPath(), "FY Registration Deposit.txt"));
                deposit_fy = true;
            }
            //// date reg
            //// fdd
            //// ldd
            string date_reg = "";
            string fdd = "";
            string ldd = "";
            string retained = "";
            string new_based_on_reg = "";
            string new_based_on_dep = "";
            string real_player = "";
            foreach (String dataLine_deposit in filecontent_deposit_fy)
            {
                String[] columns_deposit = dataLine_deposit.Split("*|*");

                if (columns_deposit[0] == player_name)
                {
                    fdd = columns_deposit[1];
                    ldd = columns_deposit[2];
                    date_reg = columns_deposit[3];

                    break;
                }
            }

            String month_get_ = DateTime.Now.Month.ToString();
            String year_get = DateTime.Now.Year.ToString();
            string year_month = year_get + "-" + month_get_;

            // New Based on Reg
            if (date_reg != "")
            {
                DateTime date_reg_get = DateTime.ParseExact(date_reg, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                if (date_reg_get.ToString("yyyy-MM") == year_month)
                {
                    new_based_on_reg = "Yes";
                }
                else
                {
                    new_based_on_reg = "No";
                }
            }
            else
            {
                new_based_on_reg = "No";
            }

            // New Based on Dep
            // Real Player
            if (fdd != "")
            {
                DateTime first_deposit = DateTime.ParseExact(fdd, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                if (first_deposit.ToString("yyyy-MM") == year_month)
                {
                    new_based_on_dep = "Yes";
                }
                else
                {
                    new_based_on_dep = "No";
                }
                real_player = "Yes";
            }
            else
            {
                new_based_on_dep = "No";
                real_player = "No";
            }

            // Retained
            if (fdd != "" && ldd != "")
            {
                DateTime last_deposit = DateTime.ParseExact(ldd, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                DateTime first_deposit = DateTime.ParseExact(fdd, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                // retained
                // 2 months current date
                var last2month_get = DateTime.Today.AddMonths(-2);
                DateTime last2month = DateTime.ParseExact(last2month_get.ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                if (last_deposit >= last2month)
                {
                    retained = "Yes";
                }
                else
                {
                    retained = "No";
                }
            }
            else
            {
                retained = "No";
            }

            string path_turnover = Path.Combine(Path.GetTempPath(), "FY Turnover.txt");
            if (!File.Exists(path_turnover))
            {
                //using (StreamWriter file = new StreamWriter(path_deposit, true, Encoding.UTF8))
                using (StreamWriter file = new StreamWriter(path_turnover, true, Encoding.UTF8))
                {
                    file.WriteLine("Brand,Provider,Category,Month,Date,Member,Currency,Stake,Stake Ex. Draw,Bet Count,Company Winloss,VIP,Retained,Reg Month,First Dep Month,New Based on Reg,New Based on Dep,Real Player");
                }
            }

            bool isFind = false;
            String[] fileContent_turnover_fy = File.ReadAllLines(Path.Combine(Path.GetTempPath(), "FY Turnover.txt"));
            foreach (String dataLine_turnover in fileContent_turnover_fy)
            {
                //Application.DoEvents();

                String[] columns_turnover = dataLine_turnover.Split(",");

                if (columns_turnover[5] == player_name)
                {
                    string text = File.ReadAllText(path_turnover);
                    int bet_count = Convert.ToInt32(columns_turnover[9]) + 1;
                    decimal stake_amount = Convert.ToDecimal(stake_amount_get) + Convert.ToDecimal(columns_turnover[7]);
                    decimal company_win_loss = Convert.ToDecimal(company_win_loss_get) + Convert.ToDecimal(columns_turnover[10]);
                    decimal valid_bet = Convert.ToDecimal(valid_bet_get) + Convert.ToDecimal(columns_turnover[8]);

                    string updated_text = columns_turnover[0] + "," + columns_turnover[1] + "," + columns_turnover[2] + "," + columns_turnover[3] + "," + columns_turnover[4] + "," + columns_turnover[5] + "," + columns_turnover[6] + "," + stake_amount + "," + valid_bet + "," + bet_count + "," + company_win_loss + "," + columns_turnover[11] + "," + columns_turnover[12] + "," + columns_turnover[13] + "," + columns_turnover[14] + "," + columns_turnover[15] + "," + columns_turnover[16] + "," + columns_turnover[17];
                    text = text.Replace(columns_turnover[0] + "," + columns_turnover[1] + "," + columns_turnover[2] + "," + columns_turnover[3] + "," + columns_turnover[4] + "," + columns_turnover[5] + "," + columns_turnover[6] + "," + columns_turnover[7] + "," + columns_turnover[8] + "," + columns_turnover[9] + "," + columns_turnover[10] + "," + columns_turnover[11] + "," + columns_turnover[12] + "," + columns_turnover[13] + "," + columns_turnover[14] + "," + columns_turnover[15] + "," + columns_turnover[16] + "," + columns_turnover[17], updated_text);
                    File.WriteAllText(path_turnover, text, Encoding.UTF8);
                    isFind = true;
                    break;
                }
                else
                {
                    isFind = false;
                }
            }

            if (!isFind)
            {
                // get category
                // get provider
                String[] fileContent_gameplatform_fy = File.ReadAllLines(Path.Combine(Path.GetTempPath(), "FY Game Platform Code.txt"));
                string category = "";
                string platform = "";
                foreach (String dataLine_gameplatform in fileContent_gameplatform_fy)
                {
                    //Application.DoEvents();

                    String[] columns_gameplatform = dataLine_gameplatform.Split("*|*");

                    if (columns_gameplatform[0] == gameplatform_get)
                    {
                        category = columns_gameplatform[1];
                        platform = columns_gameplatform[2];
                        break;
                    }
                }

                using (StreamWriter file = new StreamWriter(path_turnover, true, Encoding.UTF8))
                {
                    file.WriteLine("FY," + platform + "," + category + "," + month_get + "," + date_get + "," + player_name + ",CNY," + stake_amount_get + "," + valid_bet_get + ",1," + company_win_loss_get + "," + vip_get + "," + retained + "," + date_reg + "," + fdd + "," + new_based_on_reg + "," + new_based_on_dep + "," + real_player);
                }
            }
        }

        private void SaveAsTurnOver_FY()
        {
            try
            {
                // Turnover Record
                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Turnover Record"))
                {
                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Turnover Record");
                }

                string _fy_filename = "FY_TurnoverRecord_" + _fy_current_datetime.ToString() + "_1.xlsx";
                string _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Turnover Record\\FY_TurnoverRecord_" + _fy_current_datetime.ToString() + "_1.txt";
                string _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Turnover Record\\FY_TurnoverRecord_" + _fy_current_datetime.ToString() + "_1.xlsx";
                string _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Turnover Record\\";

                if (File.Exists(_fy_folder_path_result))
                {
                    File.Delete(_fy_folder_path_result);
                }

                if (File.Exists(_fy_folder_path_result_xlsx))
                {
                    File.Delete(_fy_folder_path_result_xlsx);
                }

                string path = Path.Combine(Path.GetTempPath(), "FY Turnover.txt");

                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet worksheet = wb.ActiveSheet;
                worksheet.Activate();
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                firstRow.AutoFilter(1,
                                    Type.Missing,
                                    Excel.XlAutoFilterOperator.xlAnd,
                                    Type.Missing,
                                    true);
                worksheet.Columns[4].NumberFormat = "dd-MMM";
                worksheet.Columns[5].NumberFormat = "dd-MMM";
                //worksheet.Columns[3].Replace(" ", "");
                //worksheet.Columns[3].NumberFormat = "@";
                //worksheet.Columns[2].NumberFormat = "MMM-yy";
                //worksheet.Columns[4].NumberFormat = "hh:mm:ss AM/PM";
                //worksheet.Columns[5].NumberFormat = "hh:mm:ss AM/PM";
                Excel.Range usedRange = worksheet.UsedRange;
                Excel.Range rows = usedRange.Rows;
                int count = 0;
                foreach (Excel.Range row in rows)
                {
                    if (count == 0)
                    {
                        Excel.Range firstCell = row.Cells[1];

                        string firstCellValue = firstCell.Value as String;

                        if (!string.IsNullOrEmpty(firstCellValue))
                        {
                            row.Interior.Color = Color.FromArgb(222, 30, 112);
                            row.Font.Color = Color.FromArgb(255, 255, 255);
                            row.Font.Bold = true;
                            row.Font.Size = 12;
                        }

                        break;
                    }

                    count++;
                }
                int i_excel;
                for (i_excel = 1; i_excel <= 20; i_excel++)
                {
                    worksheet.Columns[i_excel].ColumnWidth = 20;
                }
                wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();
                app.Quit();
                Marshal.ReleaseComObject(app);

                InsertTurnoverRecord_FY(path, _fy_folder_path_result_xlsx);
            }
            catch (Exception err)
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                __send_email = 0;
            }
        }

        private void InsertTurnoverRecord_FY(string path, string path_)
        {
            button_fy_proceed.Text = "SENDING...";
            button_fy_proceed.Enabled = false;
            WindowState = FormWindowState.Normal;
            label_fy_locatefolder.Enabled = false;
            label_fy_insert.Visible = true;

            try
            {
                var lines = File.ReadAllLines(path);
                if (lines.Count() == 0) return;
                var table = new DataTable();
                string header = "Brand,Provider,Category,Month,Date,Member,Currency,Stake,Stake Ex# Draw,Bet Count,Company Winloss,VIP,Retained,Reg Month,First Dep Month,New Based on Reg,New based on Dep,Real Player,File Name";
                var columns = header.Split(',');
                foreach (var c in columns)
                {
                    table.Columns.Add(c);
                }

                for (int i = 1; i < lines.Count(); i++)
                {
                    if (lines[i].Length > 0)
                    {
                        Application.DoEvents();
                        display_count_turnover_fy++;
                        label_fy_insert.Text = display_count_turnover_fy.ToString("N0");

                        string[] get_column = lines[i].Split(",");

                        DataRow row = table.NewRow();
                        // Brand
                        row[0] = "FY";
                        // Provider
                        row[1] = get_column[1].Replace("\"", "");
                        // Category
                        row[2] = get_column[2];
                        // Month
                        row[3] = get_column[3].Replace("\"", "");
                        // Date
                        row[4] = get_column[4].Replace("\"", "");
                        // Member
                        row[5] = get_column[5].Replace("\"", "");
                        // Currency
                        row[6] = get_column[6].Replace("\"", "");
                        // Stake
                        row[7] = get_column[7].Replace("\"", "");
                        // Stake Ex Draw
                        row[8] = get_column[8].Replace("\"", "");
                        // Bet Count
                        row[9] = get_column[9].Replace("\"", "");
                        // Company WL
                        row[10] = get_column[10].Replace("\"", "");
                        // VIP
                        row[11] = get_column[11].Replace("\"", "");
                        // Retained
                        row[12] = get_column[12].Replace("\"", "");
                        // Register Month
                        if (get_column[13].Replace("\"", "") != "")
                        {
                            row[13] = get_column[13].Replace("\"", "");
                        }
                        else
                        {
                            row[13] = DBNull.Value;
                        }
                        // First Deposit Date
                        if (get_column[14].Replace("\"", "") != "")
                        {
                            row[14] = get_column[14].Replace("\"", "");
                        }
                        else
                        {
                            row[14] = DBNull.Value;
                        }
                        // New Based on Registration
                        row[15] = get_column[15].Replace("\"", "");
                        // New Based on Department
                        row[16] = get_column[16].Replace("\"", "");
                        // Real Player
                        row[17] = get_column[17].Replace("\"", "");
                        // File Name
                        row[18] = path_;

                        table.Rows.Add(row);
                    }
                }

                var connection = @"Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";
                var sqlBulk = new SqlBulkCopy(connection);
                sqlBulk.DestinationTableName = "[testrain].[dbo].[FY.Turnover Report]";
                sqlBulk.WriteToServer(table);

                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception err)
            {
                button_fy_proceed.Text = "PROCEED";
                button_fy_proceed.Enabled = true;
                label_fy_locatefolder.Enabled = true;

                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                SendEmail2("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Message: <b>There's a problem to the server, please re-open the application.</b></body></html>");
                __send_email = 0;

                isClose = false;
                Environment.Exit(0);
            }
        }
        
        // Drag Header to Move
        private void panel_header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        // Drag Header to Move
        private void label_filelocation_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        // Drag Header to Move
        private void label_title_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        // Border
        protected override void OnPaint(PaintEventArgs e)
        {
            SolidBrush defaultColor = new SolidBrush(Color.FromArgb(78, 122, 159));
            e.Graphics.FillRectangle(defaultColor, Top);
            e.Graphics.FillRectangle(defaultColor, Left);
            e.Graphics.FillRectangle(defaultColor, Right);
            e.Graphics.FillRectangle(defaultColor, Bottom);
        }

        // Minimize Click in Taskbar
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.Style |= WS_MINIMIZEBOX;
                cp.ClassStyle |= CS_DBLCLKS;
                return cp;
            }
        }

        public object RoundedCorners { get; private set; }

        // Close
        private void pictureBox_close_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Exit the program?", "FY Cronos Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                isClose = false;
                Environment.Exit(0);
            }
        }

        private void Main_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!isClose)
            {
                DialogResult dr = MessageBox.Show("Exit the program?", "FY Cronos Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.No)
                {
                    e.Cancel = true;
                }
                else
                {
                    Environment.Exit(0);
                }
            }

            Environment.Exit(0);
        }

        // Minimize
        private void pictureBox_minimize_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        // FY Paint Panel
        private void panel_fy_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel_fy.ClientRectangle;
            rect.Width--;
            rect.Height--;
            e.Graphics.DrawRectangle(Pens.LightGray, rect);
        }

        // File Location
        private void button_filelocation_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Select File Location";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                if (FY_Cronos_Data.Properties.Settings.Default.filelocation == "")
                {
                    label_filelocation.Text = FY_Cronos_Data.Properties.Settings.Default.filelocation;
                }

                label_filelocation.Text = fbd.SelectedPath;
                FY_Cronos_Data.Properties.Settings.Default.filelocation = fbd.SelectedPath;
                FY_Cronos_Data.Properties.Settings.Default.Save();

                panel_fy.Enabled = true;
            }
        }

        private void label_updates_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"updater.exe");
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString());
            }
        }

        private void timer_landing_Tick(object sender, EventArgs e)
        {
            panel_landing.Visible = false;
            label_title.Visible = true;
            panel.Visible = true;
            pictureBox_minimize.Visible = true;
            pictureBox_close.Visible = true;
            //label_updates.Visible = true;
            label_version.Visible = true;
            label_status.Visible = true;
            label_status_1.Visible = true;
            label_cycle_in.Visible = true;
            label_cycle_in_1.Visible = true;
            timer_landing.Stop();
        }

        private void timer_midnight_Tick(object sender, EventArgs e)
        {
            if (FY_Cronos_Data.Properties.Settings.Default.______midnight_time == "")
            {
                DateTime today = DateTime.Now;
                DateTime date = today.AddDays(1);
                FY_Cronos_Data.Properties.Settings.Default.______midnight_time = date.ToString("yyyy-MM-dd 00:30");
                FY_Cronos_Data.Properties.Settings.Default.Save();
            }
            else
            {
                DateTime today = DateTime.Now;
                if (FY_Cronos_Data.Properties.Settings.Default.______midnight_time == today.ToString("yyyy-MM-dd HH:mm"))
                {
                    if (FY_Cronos_Data.Properties.Settings.Default.______start_detect == "0")
                    {
                        FY_Cronos_Data.Properties.Settings.Default.______midnight_time = "";
                        FY_Cronos_Data.Properties.Settings.Default.Save();

                        getmemberlist_fy.Clear();
                        __getcontactemail.Clear();
                        GetMemberList_FY();
                        GetBonusCode_FY();
                        GetGamePlatform_FY();
                        GetPaymentType_FY();

                        string start_datetime_ = dateTimePicker_start_fy.Text;
                        DateTime start_ = DateTime.Parse(start_datetime_).AddDays(1);

                        string end_datetime_ = dateTimePicker_end_fy.Text;
                        DateTime end_ = DateTime.Parse(end_datetime_).AddDays(1);

                        dateTimePicker_start_fy.Text = start_.ToString("yyyy-MM-dd 00:00:00");
                        dateTimePicker_end_fy.Text = end_.ToString("yyyy-MM-dd 23:59:59");

                        //button_fy_start.Enabled = true;
                        //button_fy_start.PerformClick();

                        // registration member list
                        panel_fy_status.Visible = true;
                        label_fy_page_count_1.Visible = true;
                        label_fy_total_records_1.Visible = true;
                        label_fy_page_count.Visible = true;
                        label_fy_currentrecord.Visible = true;
                        panel_fy_datetime.Visible = true;
                        panel_fy_status.BringToFront();
                        button_fy_start.Visible = false;
                        panel_fy_filter.Visible = false;
                        pictureBox_fy_loader.Visible = true;
                        _fy_current_datetime = "";
                        label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                        _fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                        label_fy_status.Visible = true;
                        timer_fy.Start();
                        isButtonStart_fy = true;
                        FY_GetPlayerListsAsync();
                    }
                }
                else
                {
                    string start_datetime = today.ToString("yyyy-MM-dd HH:mm");
                    DateTime start = DateTime.ParseExact(start_datetime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

                    string end_datetime = FY_Cronos_Data.Properties.Settings.Default.______midnight_time;
                    DateTime end = DateTime.ParseExact(end_datetime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

                    if (start > end)
                    {
                        if (FY_Cronos_Data.Properties.Settings.Default.______start_detect == "0")
                        {
                            FY_Cronos_Data.Properties.Settings.Default.______midnight_time = "";
                            FY_Cronos_Data.Properties.Settings.Default.Save();

                            getmemberlist_fy.Clear();
                            __getcontactemail.Clear();
                            GetMemberList_FY();
                            GetBonusCode_FY();
                            GetGamePlatform_FY();
                            GetPaymentType_FY();

                            string start_datetime_ = dateTimePicker_start_fy.Text;
                            DateTime start_ = DateTime.Parse(start_datetime_).AddDays(1);

                            string end_datetime_ = dateTimePicker_end_fy.Text;
                            DateTime end_ = DateTime.Parse(end_datetime_).AddDays(1);

                            dateTimePicker_start_fy.Text = start_.ToString("yyyy-MM-dd 00:00:00");
                            dateTimePicker_end_fy.Text = end_.ToString("yyyy-MM-dd 23:59:59");

                            //button_fy_start.Enabled = true;
                            //button_fy_start.PerformClick();

                            // registration member list
                            panel_fy_status.Visible = true;
                            label_fy_page_count_1.Visible = true;
                            label_fy_total_records_1.Visible = true;
                            label_fy_page_count.Visible = true;
                            label_fy_currentrecord.Visible = true;
                            panel_fy_datetime.Visible = true;
                            panel_fy_status.BringToFront();
                            button_fy_start.Visible = false;
                            panel_fy_filter.Visible = false;
                            pictureBox_fy_loader.Visible = true;
                            _fy_current_datetime = "";
                            label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                            _fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                            label_fy_status.Visible = true;
                            timer_fy.Start();
                            isButtonStart_fy = true;
                            FY_GetPlayerListsAsync();
                        }
                    }
                }
            }
        }

        string start_get = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        private void timer_cycle_in_Tick(object sender, EventArgs e)
        {
            if (FY_Cronos_Data.Properties.Settings.Default.______midnight_time != "")
            {
                string cyclein_parse = FY_Cronos_Data.Properties.Settings.Default.______midnight_time;
                DateTime cyclein = DateTime.ParseExact(cyclein_parse, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

                string start_parse = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                DateTime start = DateTime.ParseExact(start_parse, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

                TimeSpan difference = cyclein - start;
                int hrs = difference.Hours;
                int mins = difference.Minutes;
                int secs = difference.Seconds;

                TimeSpan spinTime = new TimeSpan(hrs, mins, secs);

                TimeSpan delta = DateTime.Now - start;
                TimeSpan timeRemaining = spinTime - delta;
                
                if (timeRemaining.Hours != 0 && timeRemaining.Minutes != 0)
                {
                    label_cycle_in.Text = timeRemaining.Hours + " hr(s) " + timeRemaining.Minutes + " min(s)";
                }
                else if (timeRemaining.Hours == 0 && timeRemaining.Minutes == 0)
                {
                    label_cycle_in.Text = timeRemaining.Seconds + " sec(s)";
                }
                else if (timeRemaining.Hours == 0)
                {
                    label_cycle_in.Text = timeRemaining.Minutes + " min(s) " + timeRemaining.Seconds + " sec(s)";
                }
            }
            else
            {
                label_cycle_in.Text = "-";
            }
        }

        private void SendEmail(string get_message)
        {
            try
            {
                int port = 587;
                string host = "smtp.gmail.com";
                string username = "drake@18tech.com";
                string password = "@ccess123418tech";
                string mailFrom = "noreply@mail.com";
                string mailTo = "drake@18tech.com";
                string mailTitle = "FY Cronos Data";
                string mailMessage = get_message;

                using (SmtpClient client = new SmtpClient())
                {
                    MailAddress from = new MailAddress(mailFrom);
                    MailMessage message = new MailMessage
                    {
                        From = from
                    };
                    message.To.Add(mailTo);
                    message.Subject = mailTitle;
                    message.Body = mailMessage;
                    message.IsBodyHtml = true;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = host;
                    client.Port = port;
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential
                    {
                        UserName = username,
                        Password = password
                    };
                    client.Send(message);
                }
            }
            catch (Exception err)
            {
                __send_email++;
                if (__send_email <= 5)
                {
                    SendEmail(get_message);
                }
                else
                {
                    MessageBox.Show(err.ToString());
                }
            }
        }
        
        private void SendEmail2(string get_message)
        {
            try
            {
                int port = 587;
                string host = "smtp.gmail.com";
                string username = "drake@18tech.com";
                string password = "@ccess123418tech";
                string mailFrom = "noreply@mail.com";
                string mailTo = "it@18tech.com";
                string mailTitle = "FY Cronos Data";
                string mailMessage = get_message;

                using (SmtpClient client = new SmtpClient())
                {
                    MailAddress from = new MailAddress(mailFrom);
                    MailMessage message = new MailMessage
                    {
                        From = from
                    };
                    message.To.Add(mailTo);
                    message.Subject = mailTitle;
                    message.Body = mailMessage;
                    message.IsBodyHtml = true;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = host;
                    client.Port = port;
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential
                    {
                        UserName = username,
                        Password = password
                    };
                    client.Send(message);
                }
            }
            catch (Exception err)
            {
                __send_email++;
                if (__send_email <= 5)
                {
                    SendEmail2(get_message);
                }
                else
                {
                    MessageBox.Show(err.ToString());
                }
            }
        }

        private void SendEmail3(string get_message)
        {
            try
            {
                int port = 587;
                string host = "smtp.gmail.com";
                string username = "drake@18tech.com";
                string password = "@ccess123418tech";
                string mailFrom = "noreply@mail.com";
                string mailTo = "christian@18tech.com";
                string mailTitle = "FY Cronos Data";
                string mailMessage = get_message;

                using (SmtpClient client = new SmtpClient())
                {
                    MailAddress from = new MailAddress(mailFrom);
                    MailMessage message = new MailMessage
                    {
                        From = from
                    };
                    message.To.Add(mailTo);
                    message.Subject = mailTitle;
                    message.Body = mailMessage;
                    message.IsBodyHtml = true;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = host;
                    client.Port = port;
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential
                    {
                        UserName = username,
                        Password = password
                    };
                    client.Send(message);
                }
            }
            catch (Exception err)
            {
                __send_email++;
                if (__send_email <= 5)
                {
                    SendEmail3(get_message);
                }
                else
                {
                    MessageBox.Show(err.ToString());
                }
            }
        }
        
        private void timer_flush_memory_Tick(object sender, EventArgs e)
        {
            FlushMemory();
        }

        public static void FlushMemory()
        {
            Process prs = Process.GetCurrentProcess();
            try
            {
                prs.MinWorkingSet = (IntPtr)(300000);
            }
            catch (Exception err)
            {
                // leave blank
            }
        }
    }
}