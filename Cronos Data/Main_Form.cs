using ChoETL;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Cronos_Data
{
    public partial class Main_Form : Form
    {// Drag Header to Move
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        private string _total_records_fy;
        private double _display_length_fy = 5000;
        private double _limit_fy = 250000;
        private int _total_page_fy;
        private int _result_count_json_fy;
        private JObject jo_fy;

        // Border
        const int _ = 1;
        new Rectangle Top { get { return new Rectangle(0, 0, this.ClientSize.Width, _); } }
        new Rectangle Left { get { return new Rectangle(0, 0, _, this.ClientSize.Height); } }
        new Rectangle Bottom { get { return new Rectangle(0, this.ClientSize.Height - _, this.ClientSize.Width, _); } }
        new Rectangle Right { get { return new Rectangle(this.ClientSize.Width - _, 0, _, ClientSize.Height); } }

        // Minimize Click in Taskbar
        const int WS_MINIMIZEBOX = 0x20000;
        const int CS_DBLCLKS = 0x8;

        public Main_Form()
        {
            InitializeComponent();
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
                MessageBox.Show("Select File Location to Start the Process.");
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
                        webBrowser_fy.Document.GetElementById("csname").SetAttribute("value", "fyrain");
                        webBrowser_fy.Document.GetElementById("cspwd").SetAttribute("value", "djrain123@@@");
                        webBrowser_fy.Document.Window.ScrollTo(0, webBrowser_fy.Document.Window.Size.Height);
                    }

                    if (webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/player/list"))
                    {
                        await GetDataFYAsync();
                        FYAsync();
                    }
                }
            }
        }
       
        private async Task GetDataFYAsync()
        {
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
                { "wager_settle", "" },
                { "valid_inva", "" },
                { "start",  dateTimePicker_start_fy.Text},
                { "end", dateTimePicker_end_fy.Text},
                { "skip", "0"},
                { "ftime_188", "bettime"},
                { "data[0][name]", "sEcho"},
                { "data[0][value]", "1"},
                { "data[1][name]", "iColumns"},
                { "data[1][value]", "12"},
                { "data[2][name]", "sColumns"},
                { "data[2][value]", ""},
                { "data[3][name]", "iDisplayStart"},
                { "data[3][value]", "0"},
                { "data[4][name]", "iDisplayLength"},
                { "data[4][value]", "1"}
            };

            var reqparm = new System.Collections.Specialized.NameValueCollection
            {
                { "s_btype", "" },
                { "betNo", "" },
                { "name", "" },
                { "gpid", "0" },
                { "wager_settle", "" },
                { "valid_inva", "" },
                { "start",  "2018-10-08 00:00:00"},
                { "end", "2018-10-08 23:59:59"},
                { "skip", "0"},
                { "ftime_188", "bettime"},
                { "data[0][name]", "sEcho"},
                { "data[0][value]", "1"},
                { "data[1][name]", "iColumns"},
                { "data[1][value]", "12"},
                { "data[2][name]", "sColumns"},
                { "data[2][value]", ""},
                { "data[3][name]", "iDisplayStart"},
                { "data[3][value]", "0"},
                { "data[4][name]", "iDisplayLength"},
                { "data[4][value]", "5000"}
            };

            byte[] result_gettotal = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm_gettotal);
            string responsebody_gettotatal = Encoding.UTF8.GetString(result_gettotal);

            JObject jo_gettotal = JObject.Parse(responsebody_gettotatal);
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


            byte[] result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm);
            string responsebody = Encoding.UTF8.GetString(result);

            jo_fy = JObject.Parse(responsebody);
            JToken count = jo_fy.SelectToken("$.aaData");
            _result_count_json_fy = count.Count();
        }

        private async Task GetDataFYPagesAsync()
        {
            var cookie = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
            WebClient wc = new WebClient();

            wc.Headers.Add("Cookie", cookie);
            wc.Encoding = Encoding.UTF8;
            wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            int result_pages = (5000 * _fy_pages_count) + 1;

            var reqparm = new System.Collections.Specialized.NameValueCollection
            {
                { "s_btype", "" },
                { "betNo", "" },
                { "name", "" },
                { "gpid", "0" },
                { "wager_settle", "" },
                { "valid_inva", "" },
                { "start",  "2018-10-08 00:00:00"},
                { "end", "2018-10-08 23:59:59"},
                { "skip", "0"},
                { "ftime_188", "bettime"},
                { "data[0][name]", "sEcho"},
                { "data[0][value]", "1"},
                { "data[1][name]", "iColumns"},
                { "data[1][value]", "12"},
                { "data[2][name]", "sColumns"},
                { "data[2][value]", ""},
                { "data[3][name]", "iDisplayStart"},
                { "data[3][value]", result_pages.ToString()},
                { "data[4][name]", "iDisplayLength"},
                { "data[4][value]", "5000"}
            };
            
            byte[] result = await wc.UploadValuesTaskAsync("http://cs.ying168.bet/flow/wageredAjax2", "POST", reqparm);
            string responsebody = Encoding.UTF8.GetString(result);

            jo_fy = JObject.Parse(responsebody);
            JToken count = jo_fy.SelectToken("$.aaData");
            _result_count_json_fy = count.Count();
        }

        int fy_i = 0;
        int fy_ii = 0;

        private async void FYAsync()
        {
            try
            {
                if (fy_inserted_in_excel)
                {
                    var bet_records = new List<FY_BetRecord>();
                    int get_ii = 1;

                    for (int i = 0; i < _total_page_fy; i++)
                    {
                        if (!fy_inserted_in_excel)
                        {
                            break;
                        }

                        //fy_i = i;

                        label_fy_page_count.Text = _fy_pages_count.ToString();

                        _fy_pages_count++;

                        for (int ii = 0; ii < _result_count_json_fy; ii++)
                        {
                            //fy_ii = ii;

                            label_fy_currentrecord.Text = (get_ii++).ToString();
                            label_fy_currentrecord.Invalidate();
                            label_fy_currentrecord.Update();

                            JToken game_platform = jo_fy.SelectToken("$.aaData[" + ii + "][0]");
                            JToken player_id = jo_fy.SelectToken("$.aaData[" + ii + "][1][0]");
                            JToken player_name = jo_fy.SelectToken("$.aaData[" + ii + "][1][1]");
                            JToken bet_no = jo_fy.SelectToken("$.aaData[" + ii + "][2]");
                            JToken bet_time = jo_fy.SelectToken("$.aaData[" + ii + "][3]");
                            JToken bet_type = jo_fy.SelectToken("$.aaData[" + ii + "][4]").ToString().Replace("<br />", Environment.NewLine).PadRight(125).Substring(0, 125);
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

                            bet_records.Add(new FY_BetRecord
                            {
                                GAME_PLATFORM = game_platform.ToString(),
                                USERNAME = player_name.ToString(),
                                BET_NO = bet_no.ToString(),
                                BET_TIME = bet_time.ToString(),
                                BET_TYPE = bet_type.ToString(),
                                GAME_RESULT = game_result.ToString(),
                                STAKE_AMOUNT = Convert.ToDouble(stake_amount),
                                WIN_AMOUNT = Convert.ToDouble(win_amount),
                                COMPANY_WIN_LOSS = Convert.ToDouble(company_win_loss),
                                VALID_BET = Convert.ToDouble(valid_bet),
                                VALID_INVALID = valid_invalid.ToString()
                            });

                            if ((get_ii++) == _limit_fy)
                            {
                                // label show list to excel
                                // push to database
                                label_fy_inserting_in_excel.Text = "inserting...";

                                get_ii = 1;

                                Thread fy_displayinexcel_thread = new Thread(delegate ()
                                {
                                    FY_DisplayInExcel(bet_records);
                                });
                                fy_displayinexcel_thread.Start();

                                fy_inserted_in_excel = false;
                                timer_fy_detect_inserted_in_excel.Start();

                                break;
                            }
                        }

                        // web client request
                        await GetDataFYPagesAsync();
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString());
            }
        }
        

        private void timer_fy_detect_inserted_in_excel_Tick(object sender, EventArgs e)
        {
            // label show list to excel
            // push to database

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

        // Close
        private void pictureBox_close_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Exit the program？", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                Close();
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


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string test = "2018-10-08 23:59:59";
                var bet_records = new List<FY_BetRecord> {
                    //new FY_BetRecord { GAME_PLATFORM = "Test game plaform", USERNAME = "John Doe", BET_NO = 123456789, BET_TIME = test, BET_TYPE = "ghgjh有效", GAME_RESULT = "game result", STAKE_AMOUNT = 0.00, WIN_AMOUNT = 0.00, COMPANY_WIN_LOSS = 0.60, VALID_BET = -2.43230, VALID_INVALID = "有效有效"},
                    //new FY_BetRecord { GAME_PLATFORM = "Test game plaform", USERNAME = "John Doe", BET_NO = 123456789, BET_TIME = test, BET_TYPE = "test bet type", GAME_RESULT = "game result", STAKE_AMOUNT = 0.00, WIN_AMOUNT = 3.23, COMPANY_WIN_LOSS = 0.00, VALID_BET = 0.00, VALID_INVALID = "test asdsa"},
                };

                FY_DisplayInExcel(bet_records);
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString());
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

        int fy_displayinexel_i = 0;
        private int _fy_pages_count = 1;
        private bool fy_inserted_in_excel = true;

        private void FY_DisplayInExcel(IEnumerable<FY_BetRecord> betRecords)
        {
            fy_displayinexel_i++;

            if (Directory.Exists(label_filelocation.Text + "\\FY"))
            {
                Directory.Delete(label_filelocation.Text + "\\FY", true);
            }

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

            var row = 1;
            foreach (var betRecord in betRecords)
            {
                row++;
                workSheet.Cells[row, "A"] = betRecord.GAME_PLATFORM;
                workSheet.Cells[row, "B"] = betRecord.USERNAME;
                workSheet.Cells[row, "C"] = betRecord.BET_NO;
                workSheet.Cells[row, "D"] = betRecord.BET_TIME;
                workSheet.Cells[row, "E"] = betRecord.BET_TYPE;
                workSheet.Cells[row, "F"] = betRecord.GAME_RESULT;
                workSheet.Cells[row, "G"] = betRecord.STAKE_AMOUNT;
                workSheet.Cells[row, "H"] = betRecord.WIN_AMOUNT;
                workSheet.Cells[row, "I"] = betRecord.COMPANY_WIN_LOSS;
                workSheet.Cells[row, "J"] = betRecord.VALID_BET;
                workSheet.Cells[row, "K"] = betRecord.VALID_INVALID;
            }
            
            //workSheet.Range["D2", "D" + row].NumberFormat = "DD/MM/YYYY hh:mm:ss";
            //workSheet.Columns[1].AutoFit();
            //workSheet.Columns[2].AutoFit();
            //workSheet.Columns[3].AutoFit();
            //workSheet.Columns[4].AutoFit();
            //workSheet.Columns[5].AutoFit();
            //workSheet.Columns[6].AutoFit();
            //workSheet.Columns[7].AutoFit();
            //workSheet.Columns[8].AutoFit();
            //workSheet.Columns[9].AutoFit();
            //workSheet.Columns[10].AutoFit();
            //workSheet.Columns[11].AutoFit();

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

            string result = label_filelocation.Text + "\\FY\\" + replace_datetime_start_fy.ToString() + "\\Bet Records\\FY_BetRecords_" + replace_datetime_start_fy.ToString() + "_" + fy_displayinexel_i + ".xlsx";
            workSheet.SaveAs(result);
            Marshal.ReleaseComObject(workSheet);
            excelApp.Application.Quit();
            excelApp.Quit();

            fy_inserted_in_excel = true;
        }
    }
}