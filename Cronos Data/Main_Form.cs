using ChoETL;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
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
        List<FY_BetRecord> bet_records = new List<FY_BetRecord>();
        private string _total_records_fy;
        private double _display_length_fy = 5000;
        private double _limit_fy = 250000;
        private int _total_page_fy;
        private int _result_count_json_fy;
        private JObject jo_fy;
        int fy_displayinexel_i = 0;
        private int _fy_pages_count = 0;
        private bool fy_inserted_in_excel = true;
        private int _fy_row = 1;
        private int _fy_row_count = 1;
        private bool isDone_fy = false;
        private string _fy_folder_path_result;
        private string _fy_folder_path_result_locate;
        int _fy_secho = 0;
        int fy_i = 0;
        int fy_ii = 0;
        int _fy_get_ii = 1;
        int _fy_get_ii_display = 1;
        string _fy_path = Path.Combine(Path.GetTempPath(), "cd_fy.csv");
        private bool isCSV = true;

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
            this.FormBorderStyle = FormBorderStyle.None;
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
        }

        private void Main_Form_Load(object sender, EventArgs e)
        {
            // FY
            webBrowser_fy.Navigate("http://cs.ying168.bet/account/login");

            //SHDocVw.WebBrowser wb = (SHDocVw.WebBrowser)webBrowser_fy.ActiveXInstance;
            //wb.BeforeNavigate2 += new DWebBrowserEvents2_BeforeNavigate2EventHandler(
            //    (object pDisp,
            //     ref object URL,
            //     ref object Flags,
            //     ref object TargetFrameName,
            //     ref object PostData,
            //     ref object Headers,
            //     ref bool Cancel) =>
            //    {

            //        if (PostData == null)
            //        {
            //            MessageBox.Show("[GET] " + URL);
            //        }
            //        else
            //        {
            //            string PostDATAStr = Encoding.ASCII.GetString((Byte[])PostData);

            //            MessageBox.Show("[POST] " + URL);
            //            MessageBox.Show("[POST DATA] " + PostDATAStr);
            //            MessageBox.Show("[HEADERS] " + Headers);
            //            MessageBox.Show("[FLAGS] " + Flags);
            //        }
            //    });

            comboBox_fy.SelectedIndex = 0;
            dateTimePicker_start_fy.Format = DateTimePickerFormat.Custom;
            dateTimePicker_start_fy.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker_end_fy.Format = DateTimePickerFormat.Custom;
            dateTimePicker_end_fy.CustomFormat = "yyyy-MM-dd HH:mm:ss";

            // TF
            webBrowser_tf.Navigate("http://cs.tianfa86.org/account/login");
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
                // Lat Week
                DayOfWeek weekStart = DayOfWeek.Sunday;
                DateTime startingDate = DateTime.Today;

                while (startingDate.DayOfWeek != weekStart)
                    startingDate = startingDate.AddDays(-1);

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

                    if (webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/player/list"))
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
        
        private async Task GetDataFYAsync()
        {
            try
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
                { "start",  dateTimePicker_start_fy.Text},
                { "end", dateTimePicker_end_fy.Text},
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
                { "data[4][value]", "1"},
                { "data[5][name]", "mDataProp_0"},
                { "data[5][value]", "0"},
                { "data[6][name]", "mDataProp_1"},
                { "data[6][value]", "1"},
                { "data[7][name]", "mDataProp_2"},
                { "data[7][value]", "2"},
                { "data[8][name]", "mDataProp_3"},
                { "data[8][value]", "3"},
                { "data[9][name]", "mDataProp_4"},
                { "data[9][value]", "4"},
                { "data[10][name]", "mDataProp_5"},
                { "data[10][value]", "5"},
                { "data[11][name]", "mDataProp_6"},
                { "data[11][value]", "6"},
                { "data[12][name]", "mDataProp_7"},
                { "data[12][value]", "7"},
                { "data[13][name]", "mDataProp_8"},
                { "data[13][value]", "8"},
                { "data[14][name]", "mDataProp_9"},
                { "data[14][value]", "9"},
                { "data[15][name]", "mDataProp_10"},
                { "data[15][value]", "10"},
                { "data[16][name]", "mDataProp_11"},
                { "data[16][value]", "11"}
            };

                var reqparm = new System.Collections.Specialized.NameValueCollection
            {
                { "s_btype", "" },
                { "betNo", "" },
                { "name", "" },
                { "gpid", "0" },
                { "wager_settle", "0" },
                { "valid_inva", "" },
                { "start",  dateTimePicker_start_fy.Text},
                { "end", dateTimePicker_end_fy.Text},
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
                { "data[4][value]", "5000"},
                // edited 5000
                { "data[5][name]", "mDataProp_0"},
                { "data[5][value]", "0"},
                { "data[6][name]", "mDataProp_1"},
                { "data[6][value]", "1"},
                { "data[7][name]", "mDataProp_2"},
                { "data[7][value]", "2"},
                { "data[8][name]", "mDataProp_3"},
                { "data[8][value]", "3"},
                { "data[9][name]", "mDataProp_4"},
                { "data[9][value]", "4"},
                { "data[10][name]", "mDataProp_5"},
                { "data[10][value]", "5"},
                { "data[11][name]", "mDataProp_6"},
                { "data[11][value]", "6"},
                { "data[12][name]", "mDataProp_7"},
                { "data[12][value]", "7"},
                { "data[13][name]", "mDataProp_8"},
                { "data[13][value]", "8"},
                { "data[14][name]", "mDataProp_9"},
                { "data[14][value]", "9"},
                { "data[15][name]", "mDataProp_10"},
                { "data[15][value]", "10"},
                { "data[16][name]", "mDataProp_11"},
                { "data[16][value]", "11"}
            };

                byte[] result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm_gettotal);
                string responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);
                var deserializeObject_gettotal = JsonConvert.DeserializeObject(responsebody_gettotatal);

                JObject jo_gettotal = JObject.Parse(deserializeObject_gettotal.ToString());
                JToken jt_gettotal = jo_gettotal.SelectToken("$.iTotalRecords");
                _total_records_fy = jt_gettotal.ToString();

                double result_total_records = double.Parse(_total_records_fy) / _display_length_fy;

                if (result_total_records.ToString().Contains("."))
                {
                    _total_page_fy = Convert.ToInt32(Math.Floor(result_total_records)) + 1;
                }
                else
                {
                    _total_page_fy = Convert.ToInt32(Math.Floor(result_total_records));
                }

                // status
                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                label_fy_page_count.Text = "1 of " + _total_page_fy.ToString("N0");
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
                MessageBox.Show(err.ToString());
            }
        }

        private async Task GetDataFYPagesAsync()
        {
            var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
            WebClient wc = new WebClient();

            wc.Headers.Add("Cookie", cookie);
            wc.Encoding = Encoding.UTF8;
            wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            int result_pages = (5000 * _fy_pages_count);

            var reqparm = new System.Collections.Specialized.NameValueCollection
            {
                { "s_btype", "" },
                { "betNo", "" },
                { "name", "" },
                { "gpid", "0" },
                { "wager_settle", "0" },
                { "valid_inva", "" },
                { "start",  dateTimePicker_start_fy.Text},
                { "end", dateTimePicker_end_fy.Text},
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
                { "data[4][value]", "5000"},
                // edited 100
                { "data[5][name]", "mDataProp_0"},
                { "data[5][value]", "0"},
                { "data[6][name]", "mDataProp_1"},
                { "data[6][value]", "1"},
                { "data[7][name]", "mDataProp_2"},
                { "data[7][value]", "2"},
                { "data[8][name]", "mDataProp_3"},
                { "data[8][value]", "3"},
                { "data[9][name]", "mDataProp_4"},
                { "data[9][value]", "4"},
                { "data[10][name]", "mDataProp_5"},
                { "data[10][value]", "5"},
                { "data[11][name]", "mDataProp_6"},
                { "data[11][value]", "6"},
                { "data[12][name]", "mDataProp_7"},
                { "data[12][value]", "7"},
                { "data[13][name]", "mDataProp_8"},
                { "data[13][value]", "8"},
                { "data[14][name]", "mDataProp_9"},
                { "data[14][value]", "9"},
                { "data[15][name]", "mDataProp_10"},
                { "data[15][value]", "10"},
                { "data[16][name]", "mDataProp_11"},
                { "data[16][value]", "11"}
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


        StringBuilder csv = new StringBuilder();
        private string _fy_start_datetime;
        private string _fy_finish_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

        private async void FYAsync()
        {
            try
            {
                //var bet_records = new List<FY_BetRecord>();

                if (fy_inserted_in_excel)
                {
                    for (int i = fy_i; i < _total_page_fy; i++)
                    {
                        button_fy_start.Visible = false;

                        //if (File.Exists(_fy_path))
                        //{
                        //    File.Delete(_fy_path);
                        //}

                        if (!fy_inserted_in_excel)
                        {
                            break;
                        }
                        else
                        {
                            fy_i = i;
                            _fy_pages_count++;
                        }
                        
                        for (int ii = 0; ii < _result_count_json_fy; ii++)
                        {
                            Application.DoEvents();

                            if (_fy_pages_count != 0 && _fy_pages_count <= _total_page_fy)
                            {
                                label_fy_page_count.Text = _fy_pages_count.ToString("N0") + " of " + _total_page_fy.ToString("N0");
                            }

                            fy_ii = ii;
                            JToken game_platform = jo_fy.SelectToken("$.aaData[" + ii + "][0]");
                            JToken player_id = jo_fy.SelectToken("$.aaData[" + ii + "][1][0]");
                            JToken player_name = jo_fy.SelectToken("$.aaData[" + ii + "][1][1]");
                            JToken bet_no = jo_fy.SelectToken("$.aaData[" + ii + "][2]");
                            JToken bet_time = jo_fy.SelectToken("$.aaData[" + ii + "][3]");
                            JToken bet_type = jo_fy.SelectToken("$.aaData[" + ii + "][4]").ToString().Replace("<br/>", "").PadRight(225).Substring(0, 225).Trim();
                            String result_bet_type = Regex.Replace(bet_type.ToString(), @"<[^>]*>", String.Empty);
                            JToken game_result = jo_fy.SelectToken("$.aaData[" + ii + "][5]");
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

                            // insert to text file

                            if (_fy_get_ii == 1)
                            {
                                var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}", "Game Platform", "Username", "Bet No.", "Bet Time", "Bet Type", "Game Result", "Stake Amount", "Win Amount", "Company Win/Loss", "Valid Bet", "Valid/Invalid");
                                csv.AppendLine(header);
                            }

                            var newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}", game_platform, "\"" + player_name + "\"", "\"" + bet_no + "\"", "\"" + bet_time + "\"", "\"" + result_bet_type.ToString().Replace(";", "") + "\"", "\"" + game_result + "\"", "\"" + stake_amount + "\"", "\"" + win_amount + "\"", "\"" + company_win_loss + "\"", "\"" + valid_bet + "\"", "\"" + valid_invalid + "\"");
                            csv.AppendLine(newLine);

                            //StreamWriter sw = new StreamWriter(_fy_path, true, Encoding.UTF8);
                            //sw.WriteLine(game_platform + "*|*" + player_name + "*|*" + bet_no + "*|*" + bet_time + "*|*" + bet_type + "*|*" + game_result + "*|*" + stake_amount + "*|*" + win_amount + "*|*" + company_win_loss + "*|*" + valid_bet + "*|*" + valid_invalid);
                            //sw.Close();

                            //bet_records.Add(new FY_BetRecord
                            //{
                            //    GAME_PLATFORM = game_platform.ToString(),
                            //    USERNAME = player_name.ToString(),
                            //    BET_NO = bet_no.ToString(),
                            //    BET_TIME = bet_time.ToString(),
                            //    BET_TYPE = bet_type.ToString(),
                            //    GAME_RESULT = game_result.ToString(),
                            //    STAKE_AMOUNT = Convert.ToDouble(stake_amount),
                            //    WIN_AMOUNT = Convert.ToDouble(win_amount),
                            //    COMPANY_WIN_LOSS = Convert.ToDouble(company_win_loss),
                            //    VALID_BET = Convert.ToDouble(valid_bet),
                            //    VALID_INVALID = valid_invalid.ToString()
                            //});

                            // edited
                            if ((_fy_get_ii) == 250000)
                            {
                                // status
                                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                                label_fy_status.Text = "status: saving excel...";

                                // label show list to excel
                                // push to database

                                _fy_get_ii = 1;

                                // text file to excel

                                fy_displayinexel_i++;
                                StringBuilder replace_datetime_start_fy = new StringBuilder(dateTimePicker_start_fy.Text.Substring(0, 10));
                                replace_datetime_start_fy.Replace(" ", "_");

                                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data"))
                                {
                                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data");
                                }
                                
                                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY"))
                                {
                                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY");
                                }

                                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString()))
                                {
                                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString());
                                }

                                if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records"))
                                {
                                    Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records");
                                }

                                string replace = fy_displayinexel_i.ToString();

                                if (fy_displayinexel_i.ToString().Length == 1)
                                {
                                    replace = "0" + fy_displayinexel_i;
                                }

                                _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records\\FY_BetRecords_" + replace_datetime_start_fy.ToString() + "_" + replace + ".csv";
                                _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records\\";

                                //if (File.Exists(_fy_folder_path_result))
                                //{
                                //    File.Delete(_fy_folder_path_result);
                                //}

                                //after your loop
                                File.WriteAllText(_fy_folder_path_result, csv.ToString(), Encoding.UTF8);
                                
                                csv = new StringBuilder();

                                //Thread fy_displayinexcel_thread_limit = new Thread(delegate ()
                                //{
                                //    FY_DisplayInExcel(bet_records);
                                //});
                                //fy_displayinexcel_thread_limit.Start();

                                //fy_inserted_in_excel = false;

                                //timer_fy_detect_inserted_in_excel.Start();

                                label_fy_currentrecord.Text = (_fy_get_ii_display).ToString("N0") + " of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                                label_fy_currentrecord.Invalidate();
                                label_fy_currentrecord.Update();

                                _fy_get_ii_display++;

                                break;
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

                        // web client request
                        await GetDataFYPagesAsync();
                    }
                    
                    FY_InsertDone();

                    if (fy_inserted_in_excel)
                    {
                        isDone_fy = true;

                        // done

                        // text file to excel


                        //Thread fy_displayinexcel_thread = new Thread(delegate ()
                        //{
                        //    FY_DisplayInExcel(bet_records);
                        //});
                        //fy_displayinexcel_thread.Start();
                    }

                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString());
                button1.Visible = true;
                // web client request
                await GetDataFYPagesAsync();
            }
        }

        private void FY_InsertDone()
        {
            fy_displayinexel_i++;
            StringBuilder replace_datetime_start_fy = new StringBuilder(dateTimePicker_start_fy.Text.Substring(0, 10));
            replace_datetime_start_fy.Replace(" ", "_");

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data");
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY");
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString()))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString());
            }

            if (!Directory.Exists(label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records");
            }

            string replace = fy_displayinexel_i.ToString();

            if (fy_displayinexel_i.ToString().Length == 1)
            {
                replace = "0" + fy_displayinexel_i;
            }

            _fy_folder_path_result = label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records\\FY_BetRecords_" + replace_datetime_start_fy.ToString() + "_" + replace + ".csv";
            _fy_folder_path_result_locate = label_filelocation.Text + "\\Cronos Data\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records\\";

            //if (File.Exists(_fy_folder_path_result))
            //{
            //    File.Delete(_fy_folder_path_result);
            //}

            //after your loop
            File.WriteAllText(_fy_folder_path_result, csv.ToString(), Encoding.UTF8);

            csv = new StringBuilder();

            Invoke(new Action(() =>
            {
                label_fy_finish_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                timer_fy.Stop();
                //_fy_finish_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

                //string start_datetime = _fy_start_datetime;
                //DateTime start = DateTime.ParseExact(start_datetime, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                //string finish_datetime = _fy_finish_datetime;
                //DateTime finish = DateTime.ParseExact(finish_datetime, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                //TimeSpan span = finish.Subtract(start);

                //if (span.Hours == 0 && span.Minutes == 0)
                //{
                //    label_fy_elapsed.Text = span.Seconds + " sec(s)";
                //}
                //else if (span.Hours != 0)
                //{
                //    label_fy_elapsed.Text = span.Hours + " hr(s) " + span.Minutes + " min(s) " + span.Seconds + " sec(s)";
                //}
                //else if (span.Minutes != 0)
                //{
                //    label_fy_elapsed.Text = span.Minutes + " min(s) " + span.Seconds + " sec(s)";
                //}
                //else
                //{
                //    label_fy_elapsed.Text = span.Seconds + " sec(s)";
                //}

                pictureBox_fy_loader.Visible = false;
                button_fy_proceed.Visible = true;
                label_fy_locatefolder.Visible = true;
                label_fy_status.ForeColor = Color.FromArgb(34, 139, 34);
                label_fy_status.Text = "status: done";

                panel_datetime.Location = new Point(5, 226);
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
            // label show list to excel
            // push to database
            
            // status
            label_fy_status.Text = "status: inserting data to excel...";
            
            if (fy_inserted_in_excel)
            {
                FYAsync();
                timer_fy_detect_inserted_in_excel.Stop();
            }
        }









        











































        // ----------------
        // TF ----------------
        // ----------------






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
                Application.Exit();
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
        
        private void FY_DisplayInExcel(IEnumerable<FY_BetRecord> betRecords)
        {
            //if (Directory.Exists(label_filelocation.Text + "\\FY"))
            //{
            //    Directory.Delete(label_filelocation.Text + "\\FY", true);
            //}

            Invoke(new Action(() =>
            {
                // status
                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                label_fy_status.Text = "status: inserting data to excel...";
            }));

            fy_displayinexel_i++;

            var excelApp = new Excel.Application { };
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Cells[1, "A"] = "Game Platform";
            workSheet.Cells[1, "B"] = "Username";
            workSheet.Cells[1, "C"] = "Bet No";
            workSheet.Cells[1, "D"] = "Bet Time";
            workSheet.Cells[1, "E"] = "Bet Type";
            workSheet.Cells[1, "F"] = "Game Result";
            workSheet.Cells[1, "G"] = "Stake Amount";
            workSheet.Cells[1, "H"] = "Win Amount";
            workSheet.Cells[1, "I"] = "Company Win/Loss";
            workSheet.Cells[1, "J"] = "Valid bet";
            workSheet.Cells[1, "K"] = "Valid/Invalid";
            
            foreach (var betRecord in betRecords)
            {
                try
                {
                    Invoke(new Action(() =>
                    {
                        label_fy_inserting_count.Text = _fy_row_count.ToString("N0") + " of " + Convert.ToInt32(_total_records_fy).ToString("N0");
                    }));

                    _fy_row++;
                    _fy_row_count++;
                    workSheet.Cells[_fy_row, "A"] = betRecord.GAME_PLATFORM;
                    workSheet.Cells[_fy_row, "B"] = betRecord.USERNAME;
                    workSheet.Cells[_fy_row, "C"] = betRecord.BET_NO;
                    workSheet.Cells[_fy_row, "D"] = betRecord.BET_TIME;
                    workSheet.Cells[_fy_row, "E"] = betRecord.BET_TYPE;
                    workSheet.Cells[_fy_row, "F"] = betRecord.GAME_RESULT;
                    workSheet.Cells[_fy_row, "G"] = betRecord.STAKE_AMOUNT;
                    workSheet.Cells[_fy_row, "H"] = betRecord.WIN_AMOUNT;
                    workSheet.Cells[_fy_row, "I"] = betRecord.COMPANY_WIN_LOSS;
                    workSheet.Cells[_fy_row, "J"] = betRecord.VALID_BET;
                    workSheet.Cells[_fy_row, "K"] = betRecord.VALID_INVALID;
                }
                catch (Exception err)
                {
                    // Leave blank
                }
            }

            Invoke(new Action(() =>
            {
                // status
                label_fy_status.ForeColor = Color.FromArgb(78, 122, 159);
                label_fy_status.Text = "status: saving excel...";
            }));

            _fy_row = 1;
            
            StringBuilder replace_datetime_start_fy = new StringBuilder(dateTimePicker_start_fy.Text.Substring(0, 10));
            replace_datetime_start_fy.Replace(" ", "_");

            if (!Directory.Exists(label_filelocation.Text + "\\FY"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\FY");
            }

            if (!Directory.Exists(label_filelocation.Text + "\\FY\\" + replace_datetime_start_fy.ToString()))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\FY\\" + replace_datetime_start_fy.ToString());
            }

            if (!Directory.Exists(label_filelocation.Text + "\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records"))
            {
                Directory.CreateDirectory(label_filelocation.Text + "\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records");
            }

            string replace = fy_displayinexel_i.ToString();

            if (fy_displayinexel_i.ToString().Length == 1)
            {
                replace = "0" + fy_displayinexel_i;
            }
            
            _fy_folder_path_result = label_filelocation.Text + "\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records\\FY_BetRecords_" + replace_datetime_start_fy.ToString() + "_" + replace + ".xlsx";
            _fy_folder_path_result_locate = label_filelocation.Text + "\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records\\";

            if (File.Exists(_fy_folder_path_result))
            {
                File.Delete(_fy_folder_path_result);
            }

            workSheet.SaveAs(_fy_folder_path_result);
            Marshal.ReleaseComObject(workSheet);
            excelApp.Application.Quit();
            excelApp.Quit();

            bet_records.Clear();

            fy_inserted_in_excel = true;

            if (isDone_fy)
            {
                //if (WindowState == FormWindowState.Minimized)
                //{
                //    WindowState = FormWindowState.Normal;
                //}
                
                Invoke(new Action(() =>
                {
                    label_fy_finish_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                    pictureBox_fy_loader.Visible = false;
                    button_fy_proceed.Visible = true;
                    label_fy_locatefolder.Visible = true;
                    label_fy_status.ForeColor = Color.FromArgb(34, 139, 34);
                    label_fy_status.Text = "status: done";

                    panel_datetime.Location = new Point(5, 226);
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

        }

        private void timer_fy_start_Tick(object sender, EventArgs e)
        {
            webBrowser_fy.Navigate("http://cs.ying168.bet/player/list");
        }

        private async void button_fy_start_ClickAsync(object sender, EventArgs e)
        {
            label_fy_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
            _fy_start_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            timer_fy.Start();
            timer_fy_start.Stop();
            button_fy_start.Visible = false;
            pictureBox_fy_loader.Visible = true;
            panel.Enabled = false;
            panel_fy_status.Visible = true;
            await GetDataFYAsync();
            FYAsync();
        }

        private void label_title_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void button_fy_proceed_Click(object sender, EventArgs e)
        {
            panel_fy_status.Visible = false;
            button_fy_start.Visible = true;
            panel.Enabled = true;

            button_fy_proceed.Visible = false;
            label_fy_locatefolder.Visible = false;

            label_fy_status.Text = "-";
            label_fy_page_count.Text = "-";
            label_fy_currentrecord.Text = "-";
            label_fy_inserting_count.Text = "-";
            label_fy_start_datetime.Text = "-";
            label_fy_finish_datetime.Text = "-";
            label_fy_elapsed.Text = "-";

            panel_datetime.Location = new Point(66, 226);

            // set default variables
            bet_records.Clear();
            //_total_records_fy;
             _display_length_fy = 5000;
            _limit_fy = 250000;
            //_total_page_fy;
            //_result_count_json_fy;

            fy_displayinexel_i = 0;
            _fy_pages_count = 0;
            fy_inserted_in_excel = true;
            _fy_row = 1;
            _fy_row_count = 1;
            isDone_fy = false;
            //_fy_folder_path_result;
            //_fy_folder_path_result_locate;
            _fy_secho = 0;
            fy_i = 0;
            fy_ii = 0;
            _fy_get_ii = 1;
            _fy_get_ii_display = 1;
        }

        private void label_fy_locatefolder_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(_fy_folder_path_result_locate);
            }
            catch (Exception err)
            {
                MessageBox.Show("Can't locate folder.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async void button1_ClickAsync(object sender, EventArgs e)
        {
            await GetDataFYPagesAsync();
        }
    }
}