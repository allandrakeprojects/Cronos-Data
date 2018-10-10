using ChoETL;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
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
                        FY();
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
                { "data[4][value]", "2"}
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
        
        private void FY()
        {
            int get_i = 1;
            for (int i = 0; i < _result_count_json_fy; i++)
            {
                get_i += i;
                MessageBox.Show(get_i.ToString());
                JToken game_provider = jo_fy.SelectToken("$.aaData[" + i + "][0]");
                JToken player_id = jo_fy.SelectToken("$.aaData[" + i + "][1][0]");
                JToken player_name = jo_fy.SelectToken("$.aaData[" + i + "][1][1]");
                JToken bet_no = jo_fy.SelectToken("$.aaData[" + i + "][2]");
                JToken bet_time = jo_fy.SelectToken("$.aaData[" + i + "][3]");
                JToken bet_type = jo_fy.SelectToken("$.aaData[" + i + "][4]");
                JToken bet_result = jo_fy.SelectToken("$.aaData[" + i + "][5]");
                JToken stake_amount_color = jo_fy.SelectToken("$.aaData[" + i + "][6][0]");
                JToken stake_amount = jo_fy.SelectToken("$.aaData[" + i + "][6][1]");
                JToken win_amount_color = jo_fy.SelectToken("$.aaData[" + i + "][7][0]");
                JToken win_amount = jo_fy.SelectToken("$.aaData[" + i + "][7][1]");
                JToken company_win_loss_color = jo_fy.SelectToken("$.aaData[" + i + "][8][0]");
                JToken company_win_loss = jo_fy.SelectToken("$.aaData[" + i + "][8][1]");
                JToken valid_bet_color = jo_fy.SelectToken("$.aaData[" + i + "][9][0]");
                JToken valid_bet = jo_fy.SelectToken("$.aaData[" + i + "][9][1]");
                JToken valid_invalid_id = jo_fy.SelectToken("$.aaData[" + i + "][10][0]");
                JToken valid_invalid = jo_fy.SelectToken("$.aaData[" + i + "][10][1]");


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
                var bankAccounts = new List<FY_BetRecord> {
                new FY_BetRecord { ID = 345678, Balance = 541.27},
                new FY_BetRecord {ID = 1230221,Balance = -1237.44},
                new FY_BetRecord {ID = 346777,Balance = 3532574},
                new FY_BetRecord {ID = 235788,Balance = 1500.033333}
                };

                DisplayInExcel(bankAccounts);














                //Excel.Application oXL;
                //Excel._Workbook oWB;
                //Excel._Worksheet oSheet;
                //Excel.Range oRng;


                ////Start Excel and get Application object.
                //oXL = new Excel.Application();

                ////Get a new workbook.
                //oWB = oXL.Workbooks.Add(Missing.Value);
                //oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                ////Add table headers going cell by cell.
                //oSheet.Cells[1, 1] = "First Name";
                //oSheet.Cells[1, 2] = "Last Name";
                //oSheet.Cells[1, 3] = "Full Name";
                //oSheet.Cells[1, 4] = "Salary";

                ////Format A1:D1 as bold, vertical alignment = center.
                //oSheet.get_Range("A1", "D1").Font.Bold = true;
                //oSheet.get_Range("A1", "D1").VerticalAlignment =
                //Excel.XlVAlign.xlVAlignCenter;

                //// Create an array to multiple values at once.
                //string[,] saNames = new string[5, 2];

                //saNames[0, 0] = "John";
                //saNames[0, 1] = "Smith";
                //saNames[1, 0] = "Tom";
                //saNames[1, 1] = "Brown";
                //saNames[2, 0] = "Sue";
                //saNames[2, 1] = "Thomas";
                //saNames[3, 0] = "Jane";
                //saNames[3, 1] = "Jones";
                //saNames[4, 0] = "Adam";
                //saNames[4, 1] = "Johnson";

                ////Fill A2:B6 with an array of values (First and Last Names).
                //oSheet.get_Range("A2", "B6").Value2 = saNames;

                ////Fill C2:C6 with a relative formula (=A2 & " " & B2).
                //oRng = oSheet.get_Range("C2", "C6");
                //oRng.Formula = "=A2 & \" \" & B2";

                ////Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                //oRng = oSheet.get_Range("D2", "D3");
                //oRng.Formula = "=RAND()*100000";
                //oRng.EntireColumn.NumberFormat = "$0.00";

                ////AutoFit columns A:D.
                //oRng = oSheet.get_Range("A1", "D1");
                //oRng.EntireColumn.AutoFit();

                ////Make sure Excel is visible and give the user control
                ////of Microsoft Excel's lifetime.
                //oXL.Visible = true;
                //oXL.UserControl = true;



























                //Excel.Application oXL;
                //Excel._Workbook oWB;
                //Excel._Worksheet oSheet;
                //Excel.Range oRng;

                ////Start Excel and get Application object.
                //oXL = new Excel.Application();

                ////Get a new workbook.
                //oWB = oXL.Workbooks.Add(Missing.Value);
                //oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                ////Add table headers going cell by cell.
                //oSheet.Cells[1, 1] = "Game Platform";
                //oSheet.Cells[1, 2] = "Username";
                //oSheet.Cells[1, 3] = "Bet No.";
                //oSheet.Cells[1, 4] = "Bet Time";
                //oSheet.Cells[1, 5] = "Bet Type";
                //oSheet.Cells[1, 6] = "Game Result";
                //oSheet.Cells[1, 7] = "Stake Amount";
                //oSheet.Cells[1, 8] = "Win Amount";
                //oSheet.Cells[1, 9] = "Company Win/Loss";
                //oSheet.Cells[1, 10] = "Valid Bet";
                //oSheet.Cells[1, 11] = "Valid/Invalid";

                ////Format A1:D1 as bold, vertical alignment = center.
                //oSheet.get_Range("A1", "K1").Font.Bold = true;
                //oSheet.get_Range("A1", "K1").VerticalAlignment =
                //Excel.XlVAlign.xlVAlignCenter;

                //// Create an array to multiple values at once.
                //string[,] saNames = new string[5, 11];

                //// R, C
                //saNames[0, 0] = "Game";
                //saNames[0, 1] = "John Doe";
                //saNames[0, 2] = "123456789";
                //saNames[0, 3] = "00-00-00 00:00:00";
                //saNames[0, 4] = "Type";
                //saNames[0, 5] = "Result";
                //saNames[0, 6] = "0.00";
                //saNames[0, 7] = "0.00";
                //saNames[0, 8] = "0.00";
                //saNames[0, 9] = "0.00";
                //saNames[0, 10] = "Valid";

                ////Fill A2:B6 with an array of values (First and Last Names).
                //oSheet.get_Range("A2", "K2").Value2 = saNames;
                ////oSheet.get_Range("C2").EntireColumn.NumberFormat = "@";

                ////AutoFit columns A:K.
                //oRng = oSheet.get_Range("A1", "K1");
                //oRng.EntireColumn.AutoFit();
                //oRng.EntireColumn.NumberFormat = "@";

                ////Excel.Range ThisRange = oSheet.get_Range("C2", Missing.Value);
                ////ThisRange.NumberFormat = "0.00";
                ////ThisRange.NumberFormat = "General";
                ////ThisRange.NumberFormat = "hh:mm:ss";
                ////ThisRange.NumberFormat = "DD/MM/YYYY";

                ////Marshal.FinalReleaseComObject(ThisRange);

                ////Make sure Excel is visible and give the user control
                ////of Microsoft Excel's lifetime.
                //oXL.Visible = true;
                //oXL.UserControl = true;

                ////StringBuilder replace_datetime_start_fy = new StringBuilder("testtest");
                ////replace_datetime_start_fy.Replace(":", "");
                ////replace_datetime_start_fy.Replace(" ", "_");

                ////string result = label_filelocation.Text + "\\FY_" + replace_datetime_start_fy.ToString() + "_1.xlsx";
                ////oWB.SaveAs(result, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                ////    false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                ////    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                ////oWB.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString());
            }         
        }


        public class FY_BetRecord
        {
            public int ID { get; set; }
            public double Balance { get; set; }
        }

        static void DisplayInExcel(IEnumerable<FY_BetRecord> accounts)
        {
            var excelApp = new Excel.Application { Visible = true };
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Cells[1, "A"] = "ID Number";
            workSheet.Cells[1, "B"] = "Current Balance";
            var row = 1;
            foreach (var acct in accounts)
            {
                row++;
                workSheet.Cells[row, "A"] = acct.ID;
                workSheet.Cells[row, "B"] = acct.Balance;

            }
            //workSheet.Range["B2", "B" + row].NumberFormat = "#,###.00 €";
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
        }
    }
}
