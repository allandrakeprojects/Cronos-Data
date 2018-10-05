using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
        private void webBrowser_fy_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
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
                        webBrowser_fy.Navigate("http://cs.ying168.bet/flow/wagered2");
                    }

                    if (webBrowser_fy.Url.ToString().Equals("http://cs.ying168.bet/flow/wagered2"))
                    {
                        int selected_index = comboBox_fy.SelectedIndex + 2;
                        webBrowser_fy.Document.GetElementById("TimeFli").Children[selected_index].SetAttribute("selected", "selected");
                        webBrowser_fy.Document.GetElementById("start").SetAttribute("value", dateTimePicker_start_fy.Text);
                        webBrowser_fy.Document.GetElementById("end").SetAttribute("value", dateTimePicker_end_fy.Text);
                        webBrowser_fy.Document.GetElementById("s_submit").InvokeMember("Click");
                        timer_stillsearhcing_fy.Start();
                    }
                }
            }
        }

        private async void timer_stillsearhcing_fy_TickAsync(object sender, EventArgs e)
        {
            //< div id = "data2_processing" class="dataTables_processing" style="visibility: hidden;">Processing...</div>

            HtmlElement htmlelement = webBrowser_fy.Document.GetElementById("data2_processing");
            if (htmlelement == null)
            {
                return;
            }
            else
            {
                string result = webBrowser_fy.Document.GetElementById("data2_processing").OuterHtml;
                if (!result.Contains("visibility: visible;"))
                {
                    timer_stillsearhcing_fy.Stop();

                    webBrowser_fy.Navigate("http://cs.ying168.bet/flow/wageredAjax2");
                    await GetTotalDataAsync();
                    await GetDataFYAsync();
                    MessageBox.Show("get data");
                }
            }
        }

        private async Task GetTotalDataAsync()
        {
            var cookies = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
            WebClient wc = new WebClient();
            wc.Headers.Add("Cookie: " + cookies);
            wc.Encoding = Encoding.UTF8;
            wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            string result = await wc.DownloadStringTaskAsync("http://cs.ying168.bet/flow/wageredTotalAjax2");
            MessageBox.Show(result);
        }

        private async Task GetDataFYAsync()
        {
            var cookies = FullWebBrowserCookie.GetCookieInternal(webBrowser_fy.Url, false);
            WebClient wc = new WebClient();
            wc.Headers.Add("Cookie: " + cookies);
            wc.Encoding = Encoding.UTF8;
            wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            string result = await wc.DownloadStringTaskAsync("http://cs.ying168.bet/flow/wageredAjax2");
            MessageBox.Show(result);
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
                Environment.Exit(0);
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
    }
}
