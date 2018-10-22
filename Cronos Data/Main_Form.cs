using ChoETL;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
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
        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

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
        private int _fy_row = 1;
        private int _fy_row_count = 1;
        private bool _isDone_fy = false;
        private string _fy_folder_path_result;
        private string _fy_folder_path_result_xlsx;
        private string _fy_folder_path_result_locate;
        int _fy_secho = 0;
        int _fy_i = 0;
        int _fy_ii = 0;
        int _fy_get_ii = 1;
        int _fy_get_ii_display = 1;
        int _fy_current_index = 1;
        string _fy_path = Path.Combine(Path.GetTempPath(), "cd_fy.txt");
        StringBuilder _fy_csv = new StringBuilder();
        private string _fy_start_datetime;
        private string _fy_finish_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        private bool _fy_no_result;
        private string _fy_current_datetime;
        private int _test_fy_gettotal_count_record;

        // TF ---
        List<TF_BetRecord> _tf_bet_records = new List<TF_BetRecord>();
        List<String> tf_datetime = new List<String>();
        List<String> tf_gettotal = new List<String>();
        List<String> tf_gettotal_test = new List<String>();
        private double _total_records_tf;
        private double _display_length_tf = 5000;
        private double _limit_tf = 250000;
        private int _total_page_tf;
        private int _result_count_json_tf;
        private JObject jo_tf;
        int _tf_displayinexel_i = 0;
        private int _tf_pages_count_display = 0;
        private int _tf_pages_count_last;
        private int _tf_pages_count = 0;
        private bool _detect_tf = false;
        private bool _tf_inserted_in_excel = true;
        private int _tf_row = 1;
        private int _tf_row_count = 1;
        private bool _isDone_tf = false;
        private string _tf_folder_path_result;
        private string _tf_folder_path_result_xlsx;
        private string _tf_folder_path_result_locate;
        int _tf_secho = 0;
        int _tf_i = 0;
        int _tf_ii = 0;
        int _tf_get_ii = 1;
        int _tf_get_ii_display = 1;
        int _tf_current_index = 1;
        string _tf_path = Path.Combine(Path.GetTempPath(), "cd_tf.txt");
        StringBuilder _tf_csv = new StringBuilder();
        private string _tf_start_datetime;
        private string _tf_finish_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        private bool _tf_no_result;
        private string _tf_current_datetime;
        private int _test_tf_gettotal_count_record;

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
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // height of ellipse
            int nHeightEllipse // width of ellipse
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
            dateTimePicker_start_fy.Format = DateTimePickerFormat.Custom;
            dateTimePicker_start_fy.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker_end_fy.Format = DateTimePickerFormat.Custom;
            dateTimePicker_end_fy.CustomFormat = "yyyy-MM-dd HH:mm:ss";

            // TF
            webBrowser_tf.Navigate("http://cs.tianfa86.org/account/login");
            comboBox_tf.SelectedIndex = 0;
            dateTimePicker_start_tf.Format = DateTimePickerFormat.Custom;
            dateTimePicker_start_tf.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker_end_tf.Format = DateTimePickerFormat.Custom;
            dateTimePicker_end_tf.CustomFormat = "yyyy-MM-dd HH:mm:ss";
        }

        private void Main_Form_Shown(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.filelocation == "")
            {
                panel_fy.Enabled = false;
                panel_tf.Enabled = false;
                MessageBox.Show("Select file location to start the process.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                label_filelocation.Text = Properties.Settings.Default.filelocation;
            }
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

        // Loaded
        private async void webBrowser_fy_DocumentCompletedAsync(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser_fy.ReadyState == WebBrowserReadyState.Complete)
            {
                if (e.Url == webBrowser_fy.Url)
                {
                    if (webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/account/login"))
                    {
                        webBrowser_fy.Document.Window.ScrollTo(0, 180);
                        webBrowser_fy.Document.GetElementById("csname").SetAttribute("value", "fyrain");
                        webBrowser_fy.Document.GetElementById("cspwd").SetAttribute("value", "djrain123@@@");
                    }

                    if (webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/player/list") || webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/site/index") || webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/player/online"))
                    {
                        if (panel_fy_status.Visible != true)
                        {
                            button_fy_start.Visible = true;
                        }

                        webBrowser_fy.Visible = false;
                        timer_fy_start.Start();
                    }
                }
            }
        }

        private async Task FY_GetTotal(string start_datetime, string end_datetime)
        {
            // status
            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
            label_fy_status.Text = "status: doing calculation...";

            var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
            WebClient wc = new WebClient();

            wc.Headers.Add("Cookie", cookie);
            wc.Encoding = Encoding.UTF8;
            wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            var reqparm_gettotal = new System.Collections.Specialized.NameValueCollection
            {
                { "s_btype", "" },
                { "betNo", "" },
                { "name", "" },
                { "gpid", "0" },
                { "wager_settle", "0" },
                { "valid_inva", "" },
                { "start",  start_datetime},
                { "end", end_datetime},
                { "skip", "0"},
                { "ftime_188", "bettime"},
                { "data[0][name]", "sEcho"},
                { "data[0][value]", _fy_secho++.ToString()},
                { "data[1][name]", "iColumns"},
                { "data[1][value]", "1"},
                { "data[2][name]", "sColumns"},
                { "data[2][value]", ""},
                { "data[3][name]", "iDisplayStart"},
                { "data[3][value]", "0"},
                { "data[4][name]", "iDisplayLength"},
                { "data[4][value]", "1"}
            };

            byte[] result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm_gettotal);
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
            
            fy_gettotal.Add(_total_page_fy.ToString());

            if (_total_records_fy > 0)
            {
                // status
                label_fy_page_count.Text = "0 of " + _total_page_fy.ToString("N0");
                label_fy_currentrecord.Text = "0 of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                _fy_no_result = false;
            }
            else
            {
                _fy_no_result = true;

                panel_fy_status.Visible = false;
                button_fy_start.Visible = true;
                panel_fy_filter.Enabled = true;
                button_filelocation.Enabled = true;

                button_fy_proceed.Visible = false;
                label_fy_locatefolder.Visible = false;

                label_fy_status.Text = "-";
                label_fy_page_count.Text = "-";
                label_fy_currentrecord.Text = "-";
                label_fy_inserting_count.Text = "-";
                label_fy_start_datetime.Text = "-";
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
                timer_fy_start.Start();

                MessageBox.Show("No data found.", "FY", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                        if (ii== 1)
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

                label1.Text = gettotal_start_datetime + " ----- ghghg " + gettotal_end_datetime;
                var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                
                var reqparm = new System.Collections.Specialized.NameValueCollection
                {
                    { "s_btype", "" },
                    { "betNo", "" },
                    { "name", "" },
                    { "gpid", "0" },
                    { "wager_settle", "0" },
                    { "valid_inva", "" },
                    { "start",  gettotal_start_datetime},
                    { "end", gettotal_end_datetime},
                    { "skip", "0"},
                    { "ftime_188", "bettime"},
                    { "data[0][name]", "sEcho"},
                    { "data[0][value]", _fy_secho++.ToString()},
                    { "data[1][name]", "iColumns"},
                    { "data[1][value]", "12"},
                    { "data[2][name]", "sColumns"},
                    { "data[2][value]", ""},
                    { "data[3][name]", "iDisplayStart"},
                    { "data[3][value]", "0"},
                    { "data[4][name]", "iDisplayLength"},
                    { "data[4][value]", _display_length_fy.ToString()}
                };
                                
                label_fy_status.Text = "status: getting data...";

                byte[] result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm);
                string responsebody = Encoding.UTF8.GetString(result);
                var deserializeObject = JsonConvert.DeserializeObject(responsebody);

                jo_fy = JObject.Parse(deserializeObject.ToString());
                JToken count = jo_fy.SelectToken("$.aaData");
                _result_count_json_fy = count.Count();
            }
            catch (Exception err)
            {
                detect_fy++;
                label2.Text = "detect ghghghghg" + detect_fy;
                await GetDataFYAsync();
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

                    label1.Text = gettotal_start_datetime + " ----- dsadsadas " + gettotal_end_datetime;

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

                    var reqparm = new System.Collections.Specialized.NameValueCollection
                    {
                        { "s_btype", "" },
                        { "betNo", "" },
                        { "name", "" },
                        { "gpid", "0" },
                        { "wager_settle", "0" },
                        { "valid_inva", "" },
                        { "start",  gettotal_start_datetime},
                        { "end", gettotal_end_datetime},
                        { "skip", "0"},
                        { "ftime_188", "bettime"},
                        { "data[0][name]", "sEcho"},
                        { "data[0][value]", _fy_secho++.ToString()},
                        { "data[1][name]", "iColumns"},
                        { "data[1][value]", "12"},
                        { "data[2][name]", "sColumns"},
                        { "data[2][value]", ""},
                        { "data[3][name]", "iDisplayStart"},
                        { "data[3][value]", result_pages.ToString()},
                        { "data[4][name]", "iDisplayLength"},
                        { "data[4][value]", _display_length_fy.ToString()}
                    };

                    // status
                    label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                    label_fy_status.Text = "status: getting data...";

                    byte[] result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm);
                    string responsebody = Encoding.UTF8.GetString(result);
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
                if (!_detect_fy)
                {
                    _fy_pages_count--;
                }

                detect_fy++;
                label4.Text = "detect ghghghghghghg " + detect_fy;

                await GetDataFYPagesAsync();
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

                        _test_fy_gettotal_count_record++;

                        if (_fy_pages_count_display != 0 && _fy_pages_count_display <= _total_page_fy)
                        {
                            label_fy_page_count.Text = _fy_pages_count_display.ToString("N0") + " of " + _total_page_fy.ToString("N0");
                        }

                        _fy_ii = ii;
                        JToken game_platform = jo_fy.SelectToken("$.aaData[" + ii + "][0]");
                        JToken player_id = jo_fy.SelectToken("$.aaData[" + ii + "][1][0]");
                        JToken player_name = jo_fy.SelectToken("$.aaData[" + ii + "][1][1]");
                        JToken bet_no = jo_fy.SelectToken("$.aaData[" + ii + "][2]").ToString().Replace("BetTransaction:", "");
                        JToken bet_time = jo_fy.SelectToken("$.aaData[" + ii + "][3]");
                        JToken bet_type = jo_fy.SelectToken("$.aaData[" + ii + "][4]").ToString().Replace("<br/>", "").PadRight(225).Substring(0, 225).Trim();
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

                        if (_fy_get_ii == 1)
                        {
                            var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}", "Game Platform", "Username", "Bet No.", "Bet Time", "Bet Type", "Game Result", "Stake Amount", "Win Amount", "Company Win/Loss", "Valid Bet", "Valid/Invalid");
                            _fy_csv.AppendLine(header);
                        }

                        var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}", game_platform, "\"" + player_name + "\"", "\"" + bet_no + "\"", "\"" + bet_time + "\"", "\"" + result_bet_type.ToString().Replace(";", "") + "\"", "\"" + game_result + "\"", "\"" + stake_amount + "\"", "\"" + win_amount + "\"", "\"" + company_win_loss + "\"", "\"" + valid_bet + "\"", "\"" + valid_invalid + "\"");
                        _fy_csv.AppendLine(newLine);
                        
                        if ((_fy_get_ii) == _limit_fy)
                        {
                            // status
                            label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_fy_status.Text = "status: saving excel...";
                            
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

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records"))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records");
                            }

                            string replace = _fy_displayinexel_i.ToString();

                            if (_fy_displayinexel_i.ToString().Length == 1)
                            {
                                replace = "0" + _fy_displayinexel_i;
                            }

                            _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records\\FY_BetRecords_" + replace_datetime_fy.ToString() + "_" + replace + ".txt";
                            _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records\\FY_BetRecords_" + replace_datetime_fy.ToString() + "_" + replace + ".xlsx";
                            _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records\\";

                            if (File.Exists(_fy_folder_path_result))
                            {
                                File.Delete(_fy_folder_path_result);
                            }

                            if (File.Exists(_fy_folder_path_result_xlsx))
                            {
                                File.Delete(_fy_folder_path_result_xlsx);
                            }
                            
                            File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

                            Application.DoEvents();

                            Excel.Application app = new Excel.Application();
                            Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            wb.SaveAs(_fy_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            wb.Close();
                            app.Quit();

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
                    await GetDataFYPagesAsync();
                }

                FY_InsertDone();

                if (_fy_inserted_in_excel)
                {
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

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records");
            }

            string replace = _fy_displayinexel_i.ToString();

            if (_fy_displayinexel_i.ToString().Length == 1)
            {
                replace = "0" + _fy_displayinexel_i;
            }

            _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records\\FY_BetRecords_" + replace_datetime_fy.ToString() + "_" + replace + ".txt";
            _fy_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records\\FY_BetRecords_" + replace_datetime_fy.ToString() + "_" + replace + ".xlsx";
            _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + _fy_current_datetime + "\\Bet Records\\";

            if (File.Exists(_fy_folder_path_result))
            {
                File.Delete(_fy_folder_path_result);
            }

            if (File.Exists(_fy_folder_path_result_xlsx))
            {
                File.Delete(_fy_folder_path_result_xlsx);
            }
            
            File.WriteAllText(_fy_folder_path_result, _fy_csv.ToString(), Encoding.UTF8);

            Application.DoEvents();

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(_fy_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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
                label_fy_status.Text = "status: done";
                panel_fy_datetime.Location = new Point(5, 226);
            }));

            var notification = new NotifyIcon()
            {
                Visible = true,
                Icon = SystemIcons.Information,
                BalloonTipIcon = ToolTipIcon.Info,
                BalloonTipTitle = "FY BET RECORD DONE",
                BalloonTipText = "Filter of...\nStart Time: " + dateTimePicker_start_fy.Text + "\nEnd Time: " + dateTimePicker_end_fy.Text + "\n\nStart-Finish...\nStart Time: " + label_start_fy.Text + "\nFinish Time: " + label_end_fy.Text,
            };

            notification.ShowBalloonTip(1000);

            timer_fy_start.Start();
        }

        private void FYHeader()
        {
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks.Open(_fy_folder_path_result_xlsx);
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            //int i;
            //for (i = 1; i <= 11; i++) // this will aply it form col 1 to 10
            //{
            //    worksheet.Columns[i].ColumnWidth = 15;
            //}   

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

            workbook.Save();
            workbook.Close();

            application.Quit();

            Marshal.ReleaseComObject(application);
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
            webBrowser_fy.Navigate("http://cs.ying168.bet/player/list");
        }

        private async void button_fy_start_ClickAsync(object sender, EventArgs e)
        {
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
                timer_fy_start.Stop();
                button_fy_start.Visible = false;
                pictureBox_fy_loader.Visible = true;
                panel_fy_filter.Enabled = false;
                button_filelocation.Enabled = false;
                panel_fy_status.Visible = true;

                await GetDataFYAsync();

                if (!_fy_no_result)
                {
                    FYAsync();
                }
            }
            else
            {
                _fy_no_result = true;
                MessageBox.Show("No data found.");
            }
        }
        
        private void button_fy_proceed_Click(object sender, EventArgs e)
        {
            panel_fy_status.Visible = false;
            button_fy_start.Visible = true;
            panel_fy_filter.Enabled = true;
            button_filelocation.Enabled = true;

            button_fy_proceed.Visible = false;
            label_fy_locatefolder.Visible = false;

            label_fy_status.Text = "-";
            label_fy_page_count.Text = "-";
            label_fy_currentrecord.Text = "-";
            label_fy_inserting_count.Text = "-";
            label_fy_start_datetime.Text = "-";
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
        }

        private void label_fy_locatefolder_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(_fy_folder_path_result_locate);
            }
            catch (Exception err)
            {
                MessageBox.Show("Can't locate folder.", "FY", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        



        
        // ----------------
        // TF ----------------
        // ----------------

        private void comboBox_tf_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_tf.SelectedIndex == 0)
            {
                // Yesterday
                string start_tf = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_start_tf = DateTime.ParseExact(start_tf, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                dateTimePicker_start_tf.Value = datetime_start_tf;

                string end_tf = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 23:59:59");
                DateTime datetime_end_tf = DateTime.ParseExact(end_tf, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end_tf.Value = datetime_end_tf;
            }
            else if (comboBox_tf.SelectedIndex == 1)
            {
                // Last Week
                DayOfWeek weekStart = DayOfWeek.Sunday;
                DateTime startingDate = DateTime.Today;

                while (startingDate.DayOfWeek != weekStart)
                {
                    startingDate = startingDate.AddDays(-1);
                }

                DateTime datetime_start_tf = startingDate.AddDays(-7);
                dateTimePicker_start_tf.Value = datetime_start_tf;

                string last = startingDate.AddDays(-1).ToString("yyyy-MM-dd 23:59:59");
                DateTime datetime_end_tf = DateTime.ParseExact(last, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end_tf.Value = datetime_end_tf;
            }
            else if (comboBox_tf.SelectedIndex == 2)
            {
                // Last Month
                var today = DateTime.Today;
                var month = new DateTime(today.Year, today.Month, 1);
                var first = month.AddMonths(-1).ToString("yyyy-MM-dd 00:00:00");
                var last = month.AddDays(-1).ToString("yyyy-MM-dd 23:59:59");

                DateTime datetime_start_tf = DateTime.ParseExact(first, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_start_tf.Value = datetime_start_tf;

                DateTime datetime_end_tf = DateTime.ParseExact(last, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end_tf.Value = datetime_end_tf;
            }
        }

        private void webBrowser_tf_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser_tf.ReadyState == WebBrowserReadyState.Complete)
            {
                if (e.Url == webBrowser_tf.Url)
                {
                    if (webBrowser_tf.Url.ToString().Equals("http://cs.tianfa86.org/account/login"))
                    {
                        webBrowser_tf.Document.Window.ScrollTo(0, 180);
                        //webBrowser_tf.Document.GetElementById("csname").SetAttribute("value", "tfrain");
                        //webBrowser_tf.Document.GetElementById("cspwd").SetAttribute("value", "djrain123$$");
                    }

                    if (webBrowser_tf.Url.ToString().Equals("http://cs.tianfa86.org/player/list") || webBrowser_tf.Url.ToString().Equals("http://cs.tianfa86.org/site/index") || webBrowser_tf.Url.ToString().Equals("http://cs.tianfa86.org/player/online"))
                    {
                        if (panel_tf_status.Visible != true)
                        {
                            button_tf_start.Visible = true;
                        }

                        webBrowser_tf.Visible = false;
                        timer_tf_start.Start();
                    }
                }
            }
        }

        private async Task TF_GetTotal(string start_datetime, string end_datetime)
        {
            // status
            label_tf_status.ForeColor = Color.FromArgb(78, 122, 159);
            label_tf_status.Text = "status: doing calculation...";

            var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_tf.Url, false);
            WebClient wc = new WebClient();

            wc.Headers.Add("Cookie", cookie);
            wc.Encoding = Encoding.UTF8;
            wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            var reqparm_gettotal = new System.Collections.Specialized.NameValueCollection
            {
                { "s_btype", "" },
                { "betNo", "" },
                { "name", "" },
                { "gpid", "0" },
                { "wager_settle", "0" },
                { "valid_inva", "" },
                { "start",  start_datetime},
                { "end", end_datetime},
                { "skip", "0"},
                { "ftime_188", "bettime"},
                { "data[0][name]", "sEcho"},
                { "data[0][value]", _tf_secho++.ToString()},
                { "data[1][name]", "iColumns"},
                { "data[1][value]", "1"},
                { "data[2][name]", "sColumns"},
                { "data[2][value]", ""},
                { "data[3][name]", "iDisplayStart"},
                { "data[3][value]", "0"},
                { "data[4][name]", "iDisplayLength"},
                { "data[4][value]", "1"}
            };

            byte[] result_gettotal = await wc.UploadValuesTaskAsync("http://cs.tianfa86.org/flow/wageredAjax2", "POST", reqparm_gettotal);
            string responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
            var deserializeObject_gettotal = JsonConvert.DeserializeObject(responsebody_gettotatal);

            JObject jo_gettotal = JObject.Parse(deserializeObject_gettotal.ToString());
            JToken jt_gettotal = jo_gettotal.SelectToken("$.iTotalRecords");

            _total_records_tf += double.Parse(jt_gettotal.ToString());
            double get_total_records_tf = 0;
            get_total_records_tf = double.Parse(jt_gettotal.ToString());

            tf_gettotal_test.Add(get_total_records_tf.ToString());
            double result_total_records = get_total_records_tf / _display_length_tf;

            if (result_total_records.ToString().Contains("."))
            {
                _total_page_tf += Convert.ToInt32(Math.Floor(result_total_records)) + 1;
            }
            else
            {
                _total_page_tf += Convert.ToInt32(Math.Floor(result_total_records));
            }

            tf_gettotal.Add(_total_page_tf.ToString());

            if (_total_records_tf > 0)
            {
                // status
                label_tf_page_count.Text = "0 of " + _total_page_tf.ToString("N0");
                label_tf_currentrecord.Text = "0 of " + Convert.ToInt32(_total_records_tf).ToString("N0");
                _tf_no_result = false;
            }
            else
            {
                _tf_no_result = true;

                panel_tf_status.Visible = false;
                button_tf_start.Visible = true;
                panel_tf_filter.Enabled = true;
                button_filelocation.Enabled = true;

                button_tf_proceed.Visible = false;
                label_tf_locatefolder.Visible = false;

                label_tf_status.Text = "-";
                label_tf_page_count.Text = "-";
                label_tf_currentrecord.Text = "-";
                label_tf_inserting_count.Text = "-";
                label_tf_start_datetime.Text = "-";
                label_tf_finish_datetime.Text = "-";
                label_tf_elapsed.Text = "-";

                panel_tf_datetime.Location = new Point(66, 226);

                // set default variables
                _tf_bet_records.Clear();
                tf_datetime.Clear();
                tf_gettotal.Clear();
                _total_records_tf = 0;
                _total_page_tf = 0;
                _tf_displayinexel_i = 0;
                _tf_pages_count_display = 0;
                _tf_pages_count = 0;
                _detect_tf = false;
                _tf_inserted_in_excel = true;
                _tf_row = 1;
                _tf_row_count = 1;
                _isDone_tf = false;
                _tf_secho = 0;
                _tf_i = 0;
                _tf_ii = 0;
                _tf_get_ii = 1;
                _tf_get_ii_display = 1;

                _tf_folder_path_result = "";
                _tf_folder_path_result_locate = "";
                _tf_current_index = 1;
                _tf_csv.Clear();
                timer_tf_start.Start();

                MessageBox.Show("No data found.", "TF", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async Task GetDataTFAsync()
        {
            try
            {
                string gettotal_start_datetime = "";
                string gettotal_end_datetime = "";

                tf_datetime.Reverse();

                int i = 0;
                foreach (var datetime in tf_datetime)
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
                    
                    await TF_GetTotal(get_start_datetime, get_end_datetime);
                }

                label1.Text = gettotal_start_datetime + " ----- ghghg " + gettotal_end_datetime;
                var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_tf.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                var reqparm = new System.Collections.Specialized.NameValueCollection
                {
                    { "s_btype", "" },
                    { "betNo", "" },
                    { "name", "" },
                    { "gpid", "0" },
                    { "wager_settle", "0" },
                    { "valid_inva", "" },
                    { "start",  gettotal_start_datetime},
                    { "end", gettotal_end_datetime},
                    { "skip", "0"},
                    { "ftime_188", "bettime"},
                    { "data[0][name]", "sEcho"},
                    { "data[0][value]", _tf_secho++.ToString()},
                    { "data[1][name]", "iColumns"},
                    { "data[1][value]", "12"},
                    { "data[2][name]", "sColumns"},
                    { "data[2][value]", ""},
                    { "data[3][name]", "iDisplayStart"},
                    { "data[3][value]", "0"},
                    { "data[4][name]", "iDisplayLength"},
                    { "data[4][value]", _display_length_tf.ToString()}
                };

                label_tf_status.Text = "status: getting data...";

                byte[] result = await wc.UploadValuesTaskAsync("http://cs.tianfa86.org/flow/wageredAjax2", "POST", reqparm);
                string responsebody = Encoding.UTF8.GetString(result);
                var deserializeObject = JsonConvert.DeserializeObject(responsebody);

                jo_tf = JObject.Parse(deserializeObject.ToString());
                JToken count = jo_tf.SelectToken("$.aaData");
                _result_count_json_tf = count.Count();
            }
            catch (Exception err)
            {
                detect_tf++;
                label2.Text = "detect ghghghghg" + detect_tf;
                await GetDataTFAsync();
            }
        }

        private async Task GetDataTFPagesAsync()
        {
            try
            {
                string gettotal_start_datetime = "";
                string gettotal_end_datetime = "";

                var last_item = tf_gettotal[tf_gettotal.Count - 1];

                if (last_item != _tf_pages_count_display.ToString())
                {

                    foreach (var gettotal in tf_gettotal)
                    {
                        if (gettotal == _tf_pages_count_display.ToString())
                        {
                            _tf_current_index++;
                            _tf_pages_count_last = _tf_pages_count;
                            _tf_pages_count = 0;
                            _detect_tf = true;
                            break;
                        }
                    }
                    
                    int i = 0;
                    foreach (var datetime in tf_datetime)
                    {
                        i++;
                        string[] datetime_results = datetime.Split("*|*");
                        int ii = 0;

                        foreach (string datetime_result in datetime_results)
                        {
                            ii++;
                            if (i == _tf_current_index)
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

                    label1.Text = gettotal_start_datetime + " ----- dsadsadas " + gettotal_end_datetime;

                    var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_tf.Url, false);
                    WebClient wc = new WebClient();

                    wc.Headers.Add("Cookie", cookie);
                    wc.Encoding = Encoding.UTF8;
                    wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                    int result_pages;

                    if (_detect_tf)
                    {
                        _detect_tf = false;
                        result_pages = (Convert.ToInt32(_display_length_tf) * _tf_pages_count);
                    }
                    else
                    {
                        _tf_pages_count++;
                        result_pages = (Convert.ToInt32(_display_length_tf) * _tf_pages_count);
                    }

                    var reqparm = new System.Collections.Specialized.NameValueCollection
                    {
                        { "s_btype", "" },
                        { "betNo", "" },
                        { "name", "" },
                        { "gpid", "0" },
                        { "wager_settle", "0" },
                        { "valid_inva", "" },
                        { "start",  gettotal_start_datetime},
                        { "end", gettotal_end_datetime},
                        { "skip", "0"},
                        { "ftime_188", "bettime"},
                        { "data[0][name]", "sEcho"},
                        { "data[0][value]", _tf_secho++.ToString()},
                        { "data[1][name]", "iColumns"},
                        { "data[1][value]", "12"},
                        { "data[2][name]", "sColumns"},
                        { "data[2][value]", ""},
                        { "data[3][name]", "iDisplayStart"},
                        { "data[3][value]", result_pages.ToString()},
                        { "data[4][name]", "iDisplayLength"},
                        { "data[4][value]", _display_length_tf.ToString()}
                    };

                    // status
                    label_tf_status.ForeColor = Color.FromArgb(78, 122, 159);
                    label_tf_status.Text = "status: getting data...";

                    byte[] result = await wc.UploadValuesTaskAsync("http://cs.tianfa86.org/flow/wageredAjax2", "POST", reqparm);
                    string responsebody = Encoding.UTF8.GetString(result);
                    var deserializeObject = JsonConvert.DeserializeObject(responsebody);

                    jo_tf = JObject.Parse(deserializeObject.ToString());
                    JToken count = jo_tf.SelectToken("$.aaData");
                    _result_count_json_tf = count.Count();
                }
                else
                {
                    foreach (var gettotal in tf_gettotal)
                    {
                        if (gettotal == _tf_pages_count_display.ToString())
                        {
                            _tf_current_index++;
                            _tf_pages_count_last = _tf_pages_count;
                            _tf_pages_count = 0;
                            _detect_tf = true;
                            
                            break;
                        }
                    }
                }
            }
            catch (Exception err)
            {
                if (!_detect_tf)
                {
                    _tf_pages_count--;
                }

                detect_tf++;
                label4.Text = "detect ghghghghghghg " + detect_tf;

                await GetDataTFPagesAsync();
            }
        }

        private async void TFAsync()
        {
            if (_tf_inserted_in_excel)
            {
                for (int i = _tf_i; i < _total_page_tf; i++)
                {
                    button_tf_start.Visible = false;
                    
                    if (!_tf_inserted_in_excel)
                    {
                        break;
                    }
                    else
                    {
                        _tf_i = i;
                        _tf_pages_count_display++;
                    }

                    for (int ii = 0; ii < _result_count_json_tf; ii++)
                    {
                        Application.DoEvents();

                        _test_tf_gettotal_count_record++;

                        if (_tf_pages_count_display != 0 && _tf_pages_count_display <= _total_page_tf)
                        {
                            label_tf_page_count.Text = _tf_pages_count_display.ToString("N0") + " of " + _total_page_tf.ToString("N0");
                        }

                        _tf_ii = ii;
                        JToken game_platform = jo_tf.SelectToken("$.aaData[" + ii + "][0]");
                        JToken player_id = jo_tf.SelectToken("$.aaData[" + ii + "][1][0]");
                        JToken player_name = jo_tf.SelectToken("$.aaData[" + ii + "][1][1]");
                        JToken bet_no = jo_tf.SelectToken("$.aaData[" + ii + "][2]").ToString().Replace("BetTransaction:", "");
                        JToken bet_time = jo_tf.SelectToken("$.aaData[" + ii + "][3]");
                        JToken bet_type = jo_tf.SelectToken("$.aaData[" + ii + "][4]").ToString().Replace("<br/>", "").PadRight(225).Substring(0, 225).Trim();
                        String result_bet_type = Regex.Replace(bet_type.ToString(), @"<[^>]*>", String.Empty);
                        JToken game_result = jo_tf.SelectToken("$.aaData[" + ii + "][5]").ToString().Replace("<br>", "");
                        JToken stake_amount_color = jo_tf.SelectToken("$.aaData[" + ii + "][6][0]");
                        JToken stake_amount = jo_tf.SelectToken("$.aaData[" + ii + "][6][1]");
                        JToken win_amount_color = jo_tf.SelectToken("$.aaData[" + ii + "][7][0]");
                        JToken win_amount = jo_tf.SelectToken("$.aaData[" + ii + "][7][1]");
                        JToken company_win_loss_color = jo_tf.SelectToken("$.aaData[" + ii + "][8][0]");
                        JToken company_win_loss = jo_tf.SelectToken("$.aaData[" + ii + "][8][1]");
                        JToken valid_bet_color = jo_tf.SelectToken("$.aaData[" + ii + "][9][0]");
                        JToken valid_bet = jo_tf.SelectToken("$.aaData[" + ii + "][9][1]");
                        JToken valid_invalid_id = jo_tf.SelectToken("$.aaData[" + ii + "][10][0]");
                        JToken valid_invalid = jo_tf.SelectToken("$.aaData[" + ii + "][10][1]");
                        
                        if (_tf_get_ii == 1)
                        {
                            var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}", "Game Platform", "Username", "Bet No.", "Bet Time", "Bet Type", "Game Result", "Stake Amount", "Win Amount", "Company Win/Loss", "Valid Bet", "Valid/Invalid");
                            _tf_csv.AppendLine(header);
                        }

                        var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}", game_platform, "\"" + player_name + "\"", "\"" + bet_no + "\"", "\"" + bet_time + "\"", "\"" + result_bet_type.ToString().Replace(";", "") + "\"", "\"" + game_result + "\"", "\"" + stake_amount + "\"", "\"" + win_amount + "\"", "\"" + company_win_loss + "\"", "\"" + valid_bet + "\"", "\"" + valid_invalid + "\"");
                        _tf_csv.AppendLine(newLine);
                        
                        if ((_tf_get_ii) == _limit_tf)
                        {
                            // status
                            label_tf_status.ForeColor = Color.FromArgb(78, 122, 159);
                            label_tf_status.Text = "status: saving excel...";

                            _tf_get_ii = 0;

                            _tf_displayinexel_i++;
                            StringBuilder replace_datetime_tf = new StringBuilder(dateTimePicker_start_tf.Text.Substring(0, 10) + "__" + dateTimePicker_end_tf.Text.Substring(0, 10));
                            replace_datetime_tf.Replace(" ", "_");

                            if (_tf_current_datetime == "")
                            {
                                _tf_current_datetime = DateTime.Now.ToString("yyyy-MM-dd");
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data"))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data");
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\TF"))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\TF");
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\tf\\" + _tf_current_datetime);
                            }

                            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records"))
                            {
                                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records");
                            }

                            string replace = _tf_displayinexel_i.ToString();

                            if (_tf_displayinexel_i.ToString().Length == 1)
                            {
                                replace = "0" + _tf_displayinexel_i;
                            }

                            _tf_folder_path_result = label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records\\TF_BetRecords_" + replace_datetime_tf.ToString() + "_" + replace + ".txt";
                            _tf_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records\\TF_BetRecords_" + replace_datetime_tf.ToString() + "_" + replace + ".xlsx";
                            _tf_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records\\";

                            if (File.Exists(_tf_folder_path_result))
                            {
                                File.Delete(_tf_folder_path_result);
                            }

                            if (File.Exists(_tf_folder_path_result_xlsx))
                            {
                                File.Delete(_tf_folder_path_result_xlsx);
                            }
                            
                            File.WriteAllText(_tf_folder_path_result, _tf_csv.ToString(), Encoding.UTF8);

                            Application.DoEvents();

                            Excel.Application app = new Excel.Application();
                            Excel.Workbook wb = app.Workbooks.Open(_tf_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            wb.SaveAs(_tf_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            wb.Close();
                            app.Quit();

                            if (File.Exists(_tf_folder_path_result))
                            {
                                File.Delete(_tf_folder_path_result);
                            }

                            _tf_csv.Clear();

                            label_tf_currentrecord.Text = (_tf_get_ii_display).ToString("N0") + " of " + Convert.ToInt32(_total_records_tf).ToString("N0");
                            label_tf_currentrecord.Invalidate();
                            label_tf_currentrecord.Update();
                        }
                        else
                        {
                            label_tf_currentrecord.Text = (_tf_get_ii_display).ToString("N0") + " of " + Convert.ToInt32(_total_records_tf).ToString("N0");
                            label_tf_currentrecord.Invalidate();
                            label_tf_currentrecord.Update();
                        }

                        _tf_get_ii++;
                        _tf_get_ii_display++;
                    }

                    _result_count_json_tf = 0;
                    
                    // web client request
                    await GetDataTFPagesAsync();
                }

                TF_InsertDone();

                if (_tf_inserted_in_excel)
                {
                    _isDone_tf = true;
                }

            }
        }

        public class TF_BetRecord
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

        int detect_tf = 0;

        private void TF_InsertDone()
        {
            _tf_displayinexel_i++;
            StringBuilder replace_datetime_tf = new StringBuilder(dateTimePicker_start_tf.Text.Substring(0, 10) + "__" + dateTimePicker_end_tf.Text.Substring(0, 10));
            replace_datetime_tf.Replace(" ", "_");

            if (_tf_current_datetime == "")
            {
                _tf_current_datetime = DateTime.Now.ToString("yyyy-MM-dd");
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data");
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\TF"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\TF");
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\tf\\" + _tf_current_datetime);
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records");
            }

            string replace = _tf_displayinexel_i.ToString();

            if (_tf_displayinexel_i.ToString().Length == 1)
            {
                replace = "0" + _tf_displayinexel_i;
            }

            _tf_folder_path_result = label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records\\TF_BetRecords_" + replace_datetime_tf.ToString() + "_" + replace + ".txt";
            _tf_folder_path_result_xlsx = label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records\\TF_BetRecords_" + replace_datetime_tf.ToString() + "_" + replace + ".xlsx";
            _tf_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\TF\\" + _tf_current_datetime + "\\Bet Records\\";

            if (File.Exists(_tf_folder_path_result))
            {
                File.Delete(_tf_folder_path_result);
            }

            if (File.Exists(_tf_folder_path_result_xlsx))
            {
                File.Delete(_tf_folder_path_result_xlsx);
            }
            
            File.WriteAllText(_tf_folder_path_result, _tf_csv.ToString(), Encoding.UTF8);

            Application.DoEvents();

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(_tf_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.SaveAs(_tf_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close();
            app.Quit();
            Marshal.ReleaseComObject(app);

            if (File.Exists(_tf_folder_path_result))
            {
                File.Delete(_tf_folder_path_result);
            }

            _tf_csv.Clear();

            //TFHeader();

            Invoke(new Action(() =>
            {
                label_tf_finish_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                timer_tf.Stop();
                pictureBox_tf_loader.Visible = false;
                button_tf_proceed.Visible = true;
                label_tf_locatefolder.Visible = true;
                label_tf_status.ForeColor = Color.FromArgb(34, 139, 34);
                label_tf_status.Text = "status: done";
                panel_tf_datetime.Location = new Point(5, 226);
            }));

            var notification = new NotifyIcon()
            {
                Visible = true,
                Icon = SystemIcons.Information,
                BalloonTipIcon = ToolTipIcon.Info,
                BalloonTipTitle = "TF BET RECORD DONE",
                BalloonTipText = "Filter of...\nStart Time: " + dateTimePicker_start_tf.Text + "\nEnd Time: " + dateTimePicker_end_tf.Text + "\n\nStart-Finish...\nStart Time: " + label_start_tf.Text + "\nFinish Time: " + label_end_tf.Text,
            };

            notification.ShowBalloonTip(1000);

            timer_tf_start.Start();
        }

        private void TFHeader()
        {
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks.Open(_tf_folder_path_result_xlsx);
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            //int i;
            //for (i = 1; i <= 11; i++) // this will aply it form col 1 to 10
            //{
            //    worksheet.Columns[i].ColumnWidth = 15;
            //}   

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

            workbook.Save();
            workbook.Close();

            application.Quit();

            Marshal.ReleaseComObject(application);
        }

        private void timer_tf_detect_inserted_in_excel_Tick(object sender, EventArgs e)
        {
            // status
            label_tf_status.Text = "status: inserting data to excel...";

            if (_tf_inserted_in_excel)
            {
                TFAsync();
                timer_tf_detect_inserted_in_excel.Stop();
            }
        }

        private void timer_tf_start_Tick(object sender, EventArgs e)
        {
            webBrowser_fy.Navigate("http://cs.tianfa86.org/player/list");
        }

        private void timer_tf_Tick(object sender, EventArgs e)
        {
            string start_datetime = _tf_start_datetime;
            DateTime start = DateTime.ParseExact(start_datetime, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            string finish_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            DateTime finish = DateTime.ParseExact(finish_datetime, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            TimeSpan span = finish.Subtract(start);

            if (span.Hours == 0 && span.Minutes == 0)
            {
                label_tf_elapsed.Text = span.Seconds + " sec(s)";
            }
            else if (span.Hours != 0)
            {
                label_tf_elapsed.Text = span.Hours + " hr(s) " + span.Minutes + " min(s) " + span.Seconds + " sec(s)";
            }
            else if (span.Minutes != 0)
            {
                label_tf_elapsed.Text = span.Minutes + " min(s) " + span.Seconds + " sec(s)";
            }
            else
            {
                label_tf_elapsed.Text = span.Seconds + " sec(s)";
            }
        }
        
        private async void button_tf_start_ClickAsync(object sender, EventArgs e)
        {
            tf_datetime.Clear();
            tf_gettotal.Clear();
            tf_gettotal_test.Clear();

            string start_datetime = dateTimePicker_start_tf.Text;
            DateTime start = DateTime.Parse(start_datetime);

            string end_datetime = dateTimePicker_end_tf.Text;
            DateTime end = DateTime.Parse(end_datetime);

            string result_start = start.ToString("yyyy-MM-dd");
            string result_end = end.ToString("yyyy-MM-dd");
            string result_start_time = start.ToString("HH:mm:ss");
            string result_end_time = end.ToString("HH:mm:ss");

            if (start < end)
            {
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
                            tf_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);

                            break;
                        }
                        else
                        {
                            if (i == 0)
                            {
                                string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd ") + result_end_time;
                                tf_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                            }
                            else
                            {
                                string start_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 00:00:00");
                                string end_get_to_list = end.AddDays(-i).ToString("yyyy-MM-dd 23:59:59");
                                tf_datetime.Add(start_get_to_list + "*|*" + end_get_to_list);
                            }
                        }

                        i++;
                    }
                }
                else
                {
                    tf_datetime.Add(start_datetime + "*|*" + end_datetime);
                }

                _tf_current_datetime = "";
                label_tf_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                _tf_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                timer_tf.Start();
                webBrowser_tf.Stop();
                timer_tf_start.Stop();
                button_tf_start.Visible = false;
                pictureBox_tf_loader.Visible = true;
                panel_tf_filter.Enabled = false;
                button_filelocation.Enabled = false;
                panel_tf_status.Visible = true;

                await GetDataTFAsync();

                if (!_tf_no_result)
                {
                    TFAsync();
                }
            }
            else
            {
                _tf_no_result = true;
                MessageBox.Show("No data found.");
            }
        }
        
        private void button_tf_proceed_Click(object sender, EventArgs e)
        {
            panel_tf_status.Visible = false;
            button_tf_start.Visible = true;
            panel_tf_filter.Enabled = true;
            button_filelocation.Enabled = true;

            button_tf_proceed.Visible = false;
            label_tf_locatefolder.Visible = false;

            label_tf_status.Text = "-";
            label_tf_page_count.Text = "-";
            label_tf_currentrecord.Text = "-";
            label_tf_inserting_count.Text = "-";
            label_tf_start_datetime.Text = "-";
            label_tf_finish_datetime.Text = "-";
            label_tf_elapsed.Text = "-";

            panel_tf_datetime.Location = new Point(66, 226);

            // set default variables
            _tf_bet_records.Clear();
            tf_datetime.Clear();
            tf_gettotal.Clear();
            _total_records_tf = 0;
            _total_page_tf = 0;
            _tf_displayinexel_i = 0;
            _tf_pages_count_display = 0;
            _tf_pages_count = 0;
            _detect_tf = false;
            _tf_inserted_in_excel = true;
            _tf_row = 1;
            _tf_row_count = 1;
            _isDone_tf = false;
            _tf_secho = 0;
            _tf_i = 0;
            _tf_ii = 0;
            _tf_get_ii = 1;
            _tf_get_ii_display = 1;

            _tf_folder_path_result = "";
            _tf_folder_path_result_locate = "";
            _tf_current_index = 1;
            _tf_csv.Clear();
        }

        private void label_tf_locatefolder_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(_tf_folder_path_result_locate);
            }
            catch (Exception err)
            {
                MessageBox.Show("Can't locate folder.", "TF", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            DialogResult dr = MessageBox.Show("Exit the program?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                isClose = true;
                Application.Exit();
            }
        }

        private void Main_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!isClose)
            {
                DialogResult dr = MessageBox.Show("Exit the program?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
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

        // TF Paint Panel
        private void panel_tf_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel_tf.ClientRectangle;
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
                label_filelocation.Text = fbd.SelectedPath;
                Properties.Settings.Default.filelocation = fbd.SelectedPath;
                Properties.Settings.Default.Save();

                panel_fy.Enabled = true;
                panel_tf.Enabled = true;
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

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var message = string.Join(Environment.NewLine, fy_gettotal.ToArray());
                MessageBox.Show(message);

                var messagesds = string.Join(Environment.NewLine, fy_gettotal_test.ToArray());
                MessageBox.Show(messagesds);

                var last_item = fy_gettotal[fy_gettotal.Count - 1];
                MessageBox.Show(last_item);
            }
            catch (Exception err)
            {
                // Leave blank
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            label_fy_page_count.Text = textBox1.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
            row.Cells[0].Value = "XYZ";
            dataGridView1.Rows.Add(row);
        }

        private void timer_landing_Tick(object sender, EventArgs e)
        {
            panel_landing.Visible = false;
            label_title.Visible = true;
            label_filelocation.Visible = true;
            pictureBox_minimize.Visible = true;
            pictureBox_close.Visible = true;
            //label_updates.Visible = true;
            label_version.Visible = true;
            timer_landing.Stop();
        }
    }
}