namespace Cronos_Data
{
    partial class Main_Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main_Form));
            this.panel_header = new System.Windows.Forms.Panel();
            this.label_title = new System.Windows.Forms.Label();
            this.pictureBox_minimize = new System.Windows.Forms.PictureBox();
            this.pictureBox_close = new System.Windows.Forms.PictureBox();
            this.label_filelocation = new System.Windows.Forms.Label();
            this.label_title_fy = new System.Windows.Forms.Label();
            this.panel_fy = new System.Windows.Forms.Panel();
            this.panel_fy_status = new System.Windows.Forms.Panel();
            this.label_fy_insert = new System.Windows.Forms.Label();
            this.button_fy_proceed = new System.Windows.Forms.Button();
            this.label_fy_locatefolder = new System.Windows.Forms.Label();
            this.panel_fy_datetime = new System.Windows.Forms.Panel();
            this.label_fy_elapsed = new System.Windows.Forms.Label();
            this.label_fy_elapsed_1 = new System.Windows.Forms.Label();
            this.label_fy_start_datetime_1 = new System.Windows.Forms.Label();
            this.label_fy_finish_datetime = new System.Windows.Forms.Label();
            this.label_fy_finish_datetime_1 = new System.Windows.Forms.Label();
            this.label_fy_start_datetime = new System.Windows.Forms.Label();
            this.pictureBox_fy_loader = new System.Windows.Forms.PictureBox();
            this.label_fy_currentrecord = new System.Windows.Forms.Label();
            this.label_fy_page_count = new System.Windows.Forms.Label();
            this.label_fy_page_count_1 = new System.Windows.Forms.Label();
            this.label_fy_total_records_1 = new System.Windows.Forms.Label();
            this.label_fy_status = new System.Windows.Forms.Label();
            this.webBrowser_fy = new System.Windows.Forms.WebBrowser();
            this.panel_fy_filter = new System.Windows.Forms.Panel();
            this.comboBox_fy_list = new System.Windows.Forms.ComboBox();
            this.comboBox_fy = new System.Windows.Forms.ComboBox();
            this.dateTimePicker_end_fy = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker_start_fy = new System.Windows.Forms.DateTimePicker();
            this.label_start_fy = new System.Windows.Forms.Label();
            this.label_end_fy = new System.Windows.Forms.Label();
            this.button_fy_stop = new System.Windows.Forms.Button();
            this.button_fy_start = new System.Windows.Forms.Button();
            this.timer_fy_detect_inserted_in_excel = new System.Windows.Forms.Timer(this.components);
            this.timer_fy_start = new System.Windows.Forms.Timer(this.components);
            this.button_filelocation = new System.Windows.Forms.Button();
            this.timer_fy = new System.Windows.Forms.Timer(this.components);
            this.panel_footer = new System.Windows.Forms.Panel();
            this.label_version = new System.Windows.Forms.Label();
            this.label_updates = new System.Windows.Forms.Label();
            this.panel_landing = new System.Windows.Forms.Panel();
            this.pictureBox_landing = new System.Windows.Forms.PictureBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.timer_tf_detect_inserted_in_excel = new System.Windows.Forms.Timer(this.components);
            this.timer_tf_start = new System.Windows.Forms.Timer(this.components);
            this.timer_tf = new System.Windows.Forms.Timer(this.components);
            this.timer_landing = new System.Windows.Forms.Timer(this.components);
            this.timer_fy_start_button = new System.Windows.Forms.Timer(this.components);
            this.label_fy_count = new System.Windows.Forms.Label();
            this.label_getdatacount_fy = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel = new System.Windows.Forms.Panel();
            this.panel_header.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_minimize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_close)).BeginInit();
            this.panel_fy.SuspendLayout();
            this.panel_fy_status.SuspendLayout();
            this.panel_fy_datetime.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_fy_loader)).BeginInit();
            this.panel_fy_filter.SuspendLayout();
            this.panel_footer.SuspendLayout();
            this.panel_landing.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_landing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.panel_header.Controls.Add(this.panel);
            this.panel_header.Controls.Add(this.label_title);
            this.panel_header.Controls.Add(this.pictureBox_minimize);
            this.panel_header.Controls.Add(this.pictureBox_close);
            this.panel_header.Controls.Add(this.label_filelocation);
            this.panel_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_header.Location = new System.Drawing.Point(0, 0);
            this.panel_header.Name = "panel_header";
            this.panel_header.Size = new System.Drawing.Size(569, 45);
            this.panel_header.TabIndex = 0;
            this.panel_header.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel_header_MouseDown);
            // 
            // label_title
            // 
            this.label_title.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_title.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.label_title.Location = new System.Drawing.Point(2, 0);
            this.label_title.Name = "label_title";
            this.label_title.Size = new System.Drawing.Size(166, 45);
            this.label_title.TabIndex = 2;
            this.label_title.Text = "Cronos Data";
            this.label_title.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_title.Visible = false;
            this.label_title.MouseDown += new System.Windows.Forms.MouseEventHandler(this.label_title_MouseDown);
            // 
            // pictureBox_minimize
            // 
            this.pictureBox_minimize.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.pictureBox_minimize.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox_minimize.Image = global::FY_Cronos_Data.Properties.Resources.minus;
            this.pictureBox_minimize.Location = new System.Drawing.Point(479, 10);
            this.pictureBox_minimize.Name = "pictureBox_minimize";
            this.pictureBox_minimize.Size = new System.Drawing.Size(24, 24);
            this.pictureBox_minimize.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox_minimize.TabIndex = 1;
            this.pictureBox_minimize.TabStop = false;
            this.pictureBox_minimize.Visible = false;
            this.pictureBox_minimize.Click += new System.EventHandler(this.pictureBox_minimize_Click);
            // 
            // pictureBox_close
            // 
            this.pictureBox_close.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.pictureBox_close.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox_close.Image = global::FY_Cronos_Data.Properties.Resources.close;
            this.pictureBox_close.Location = new System.Drawing.Point(518, 10);
            this.pictureBox_close.Name = "pictureBox_close";
            this.pictureBox_close.Size = new System.Drawing.Size(24, 24);
            this.pictureBox_close.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox_close.TabIndex = 0;
            this.pictureBox_close.TabStop = false;
            this.pictureBox_close.Visible = false;
            this.pictureBox_close.Click += new System.EventHandler(this.pictureBox_close_Click);
            // 
            // label_filelocation
            // 
            this.label_filelocation.ForeColor = System.Drawing.Color.White;
            this.label_filelocation.Location = new System.Drawing.Point(-7, 16);
            this.label_filelocation.Name = "label_filelocation";
            this.label_filelocation.Size = new System.Drawing.Size(580, 13);
            this.label_filelocation.TabIndex = 3;
            this.label_filelocation.Text = "-";
            this.label_filelocation.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_filelocation.Visible = false;
            this.label_filelocation.MouseDown += new System.Windows.Forms.MouseEventHandler(this.label_filelocation_MouseDown);
            // 
            // label_title_fy
            // 
            this.label_title_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_title_fy.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(30)))), ((int)(((byte)(112)))));
            this.label_title_fy.Location = new System.Drawing.Point(3, 3);
            this.label_title_fy.Name = "label_title_fy";
            this.label_title_fy.Size = new System.Drawing.Size(528, 30);
            this.label_title_fy.TabIndex = 2;
            this.label_title_fy.Text = "Feng Ying";
            this.label_title_fy.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel_fy
            // 
            this.panel_fy.Controls.Add(this.panel_fy_status);
            this.panel_fy.Controls.Add(this.label_title_fy);
            this.panel_fy.Controls.Add(this.webBrowser_fy);
            this.panel_fy.Controls.Add(this.panel_fy_filter);
            this.panel_fy.Enabled = false;
            this.panel_fy.Location = new System.Drawing.Point(17, 88);
            this.panel_fy.Name = "panel_fy";
            this.panel_fy.Size = new System.Drawing.Size(534, 408);
            this.panel_fy.TabIndex = 4;
            this.panel_fy.Paint += new System.Windows.Forms.PaintEventHandler(this.panel_fy_Paint);
            // 
            // panel_fy_status
            // 
            this.panel_fy_status.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.panel_fy_status.Controls.Add(this.label_fy_insert);
            this.panel_fy_status.Controls.Add(this.button_fy_proceed);
            this.panel_fy_status.Controls.Add(this.label_fy_locatefolder);
            this.panel_fy_status.Controls.Add(this.panel_fy_datetime);
            this.panel_fy_status.Controls.Add(this.pictureBox_fy_loader);
            this.panel_fy_status.Controls.Add(this.label_fy_currentrecord);
            this.panel_fy_status.Controls.Add(this.label_fy_page_count);
            this.panel_fy_status.Controls.Add(this.label_fy_page_count_1);
            this.panel_fy_status.Controls.Add(this.label_fy_total_records_1);
            this.panel_fy_status.Controls.Add(this.label_fy_status);
            this.panel_fy_status.Location = new System.Drawing.Point(7, 121);
            this.panel_fy_status.Name = "panel_fy_status";
            this.panel_fy_status.Size = new System.Drawing.Size(524, 284);
            this.panel_fy_status.TabIndex = 23;
            this.panel_fy_status.Visible = false;
            // 
            // label_fy_insert
            // 
            this.label_fy_insert.Location = new System.Drawing.Point(382, 207);
            this.label_fy_insert.Name = "label_fy_insert";
            this.label_fy_insert.Size = new System.Drawing.Size(125, 23);
            this.label_fy_insert.TabIndex = 34;
            this.label_fy_insert.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // button_fy_proceed
            // 
            this.button_fy_proceed.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button_fy_proceed.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button_fy_proceed.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_fy_proceed.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_fy_proceed.ForeColor = System.Drawing.Color.White;
            this.button_fy_proceed.Location = new System.Drawing.Point(382, 233);
            this.button_fy_proceed.Name = "button_fy_proceed";
            this.button_fy_proceed.Size = new System.Drawing.Size(126, 28);
            this.button_fy_proceed.TabIndex = 23;
            this.button_fy_proceed.Text = "PROCEED";
            this.button_fy_proceed.UseVisualStyleBackColor = false;
            this.button_fy_proceed.Visible = false;
            this.button_fy_proceed.Click += new System.EventHandler(this.button_fy_proceed_Click);
            // 
            // label_fy_locatefolder
            // 
            this.label_fy_locatefolder.AutoSize = true;
            this.label_fy_locatefolder.Cursor = System.Windows.Forms.Cursors.Hand;
            this.label_fy_locatefolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_locatefolder.Location = new System.Drawing.Point(436, 264);
            this.label_fy_locatefolder.Name = "label_fy_locatefolder";
            this.label_fy_locatefolder.Size = new System.Drawing.Size(72, 13);
            this.label_fy_locatefolder.TabIndex = 29;
            this.label_fy_locatefolder.Text = "Locate Folder";
            this.label_fy_locatefolder.Visible = false;
            this.label_fy_locatefolder.Click += new System.EventHandler(this.label_fy_locatefolder_Click);
            // 
            // panel_fy_datetime
            // 
            this.panel_fy_datetime.Controls.Add(this.label_fy_elapsed);
            this.panel_fy_datetime.Controls.Add(this.label_fy_elapsed_1);
            this.panel_fy_datetime.Controls.Add(this.label_fy_start_datetime_1);
            this.panel_fy_datetime.Controls.Add(this.label_fy_finish_datetime);
            this.panel_fy_datetime.Controls.Add(this.label_fy_finish_datetime_1);
            this.panel_fy_datetime.Controls.Add(this.label_fy_start_datetime);
            this.panel_fy_datetime.Location = new System.Drawing.Point(66, 226);
            this.panel_fy_datetime.Name = "panel_fy_datetime";
            this.panel_fy_datetime.Size = new System.Drawing.Size(287, 58);
            this.panel_fy_datetime.TabIndex = 28;
            this.panel_fy_datetime.Visible = false;
            // 
            // label_fy_elapsed
            // 
            this.label_fy_elapsed.AutoSize = true;
            this.label_fy_elapsed.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_elapsed.Location = new System.Drawing.Point(66, 38);
            this.label_fy_elapsed.Name = "label_fy_elapsed";
            this.label_fy_elapsed.Size = new System.Drawing.Size(11, 15);
            this.label_fy_elapsed.TabIndex = 29;
            this.label_fy_elapsed.Text = "-";
            // 
            // label_fy_elapsed_1
            // 
            this.label_fy_elapsed_1.AutoSize = true;
            this.label_fy_elapsed_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_elapsed_1.Location = new System.Drawing.Point(3, 36);
            this.label_fy_elapsed_1.Name = "label_fy_elapsed_1";
            this.label_fy_elapsed_1.Size = new System.Drawing.Size(55, 15);
            this.label_fy_elapsed_1.TabIndex = 28;
            this.label_fy_elapsed_1.Text = "Elapsed:";
            // 
            // label_fy_start_datetime_1
            // 
            this.label_fy_start_datetime_1.AutoSize = true;
            this.label_fy_start_datetime_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_start_datetime_1.Location = new System.Drawing.Point(3, 5);
            this.label_fy_start_datetime_1.Name = "label_fy_start_datetime_1";
            this.label_fy_start_datetime_1.Size = new System.Drawing.Size(35, 15);
            this.label_fy_start_datetime_1.TabIndex = 24;
            this.label_fy_start_datetime_1.Text = "Start:";
            // 
            // label_fy_finish_datetime
            // 
            this.label_fy_finish_datetime.AutoSize = true;
            this.label_fy_finish_datetime.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_finish_datetime.Location = new System.Drawing.Point(66, 21);
            this.label_fy_finish_datetime.Name = "label_fy_finish_datetime";
            this.label_fy_finish_datetime.Size = new System.Drawing.Size(11, 15);
            this.label_fy_finish_datetime.TabIndex = 27;
            this.label_fy_finish_datetime.Text = "-";
            // 
            // label_fy_finish_datetime_1
            // 
            this.label_fy_finish_datetime_1.AutoSize = true;
            this.label_fy_finish_datetime_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_finish_datetime_1.Location = new System.Drawing.Point(3, 21);
            this.label_fy_finish_datetime_1.Name = "label_fy_finish_datetime_1";
            this.label_fy_finish_datetime_1.Size = new System.Drawing.Size(43, 15);
            this.label_fy_finish_datetime_1.TabIndex = 25;
            this.label_fy_finish_datetime_1.Text = "Finish:";
            // 
            // label_fy_start_datetime
            // 
            this.label_fy_start_datetime.AutoSize = true;
            this.label_fy_start_datetime.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_start_datetime.Location = new System.Drawing.Point(66, 5);
            this.label_fy_start_datetime.Name = "label_fy_start_datetime";
            this.label_fy_start_datetime.Size = new System.Drawing.Size(11, 15);
            this.label_fy_start_datetime.TabIndex = 26;
            this.label_fy_start_datetime.Text = "-";
            // 
            // pictureBox_fy_loader
            // 
            this.pictureBox_fy_loader.Image = global::FY_Cronos_Data.Properties.Resources.loader;
            this.pictureBox_fy_loader.Location = new System.Drawing.Point(3, 180);
            this.pictureBox_fy_loader.Name = "pictureBox_fy_loader";
            this.pictureBox_fy_loader.Size = new System.Drawing.Size(60, 101);
            this.pictureBox_fy_loader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox_fy_loader.TabIndex = 23;
            this.pictureBox_fy_loader.TabStop = false;
            this.pictureBox_fy_loader.Visible = false;
            // 
            // label_fy_currentrecord
            // 
            this.label_fy_currentrecord.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_currentrecord.Location = new System.Drawing.Point(258, 116);
            this.label_fy_currentrecord.Name = "label_fy_currentrecord";
            this.label_fy_currentrecord.Size = new System.Drawing.Size(250, 18);
            this.label_fy_currentrecord.TabIndex = 12;
            this.label_fy_currentrecord.Text = "-";
            this.label_fy_currentrecord.Visible = false;
            // 
            // label_fy_page_count
            // 
            this.label_fy_page_count.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_page_count.Location = new System.Drawing.Point(259, 86);
            this.label_fy_page_count.Name = "label_fy_page_count";
            this.label_fy_page_count.Size = new System.Drawing.Size(249, 18);
            this.label_fy_page_count.TabIndex = 13;
            this.label_fy_page_count.Text = "-";
            this.label_fy_page_count.Visible = false;
            // 
            // label_fy_page_count_1
            // 
            this.label_fy_page_count_1.AutoSize = true;
            this.label_fy_page_count_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_page_count_1.Location = new System.Drawing.Point(134, 86);
            this.label_fy_page_count_1.Name = "label_fy_page_count_1";
            this.label_fy_page_count_1.Size = new System.Drawing.Size(46, 18);
            this.label_fy_page_count_1.TabIndex = 20;
            this.label_fy_page_count_1.Text = "Page:";
            this.label_fy_page_count_1.Visible = false;
            // 
            // label_fy_total_records_1
            // 
            this.label_fy_total_records_1.AutoSize = true;
            this.label_fy_total_records_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_total_records_1.Location = new System.Drawing.Point(132, 116);
            this.label_fy_total_records_1.Name = "label_fy_total_records_1";
            this.label_fy_total_records_1.Size = new System.Drawing.Size(98, 18);
            this.label_fy_total_records_1.TabIndex = 18;
            this.label_fy_total_records_1.Text = "Total Record:";
            this.label_fy_total_records_1.Visible = false;
            // 
            // label_fy_status
            // 
            this.label_fy_status.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_status.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.label_fy_status.Location = new System.Drawing.Point(3, 42);
            this.label_fy_status.Name = "label_fy_status";
            this.label_fy_status.Size = new System.Drawing.Size(518, 25);
            this.label_fy_status.TabIndex = 17;
            this.label_fy_status.Text = "-";
            this.label_fy_status.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_fy_status.Visible = false;
            // 
            // webBrowser_fy
            // 
            this.webBrowser_fy.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.webBrowser_fy.Location = new System.Drawing.Point(6, 35);
            this.webBrowser_fy.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser_fy.Name = "webBrowser_fy";
            this.webBrowser_fy.ScriptErrorsSuppressed = true;
            this.webBrowser_fy.Size = new System.Drawing.Size(522, 367);
            this.webBrowser_fy.TabIndex = 0;
            this.webBrowser_fy.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser_fy_DocumentCompletedAsync);
            // 
            // panel_fy_filter
            // 
            this.panel_fy_filter.Controls.Add(this.comboBox_fy_list);
            this.panel_fy_filter.Controls.Add(this.comboBox_fy);
            this.panel_fy_filter.Controls.Add(this.dateTimePicker_end_fy);
            this.panel_fy_filter.Controls.Add(this.dateTimePicker_start_fy);
            this.panel_fy_filter.Controls.Add(this.label_start_fy);
            this.panel_fy_filter.Controls.Add(this.label_end_fy);
            this.panel_fy_filter.Location = new System.Drawing.Point(3, 35);
            this.panel_fy_filter.Name = "panel_fy_filter";
            this.panel_fy_filter.Size = new System.Drawing.Size(528, 80);
            this.panel_fy_filter.TabIndex = 24;
            // 
            // comboBox_fy_list
            // 
            this.comboBox_fy_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_fy_list.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox_fy_list.FormattingEnabled = true;
            this.comboBox_fy_list.Items.AddRange(new object[] {
            "Payment Report",
            "Bonus Report",
            "Bet Record"});
            this.comboBox_fy_list.Location = new System.Drawing.Point(69, 47);
            this.comboBox_fy_list.Name = "comboBox_fy_list";
            this.comboBox_fy_list.Size = new System.Drawing.Size(133, 23);
            this.comboBox_fy_list.TabIndex = 12;
            // 
            // comboBox_fy
            // 
            this.comboBox_fy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox_fy.FormattingEnabled = true;
            this.comboBox_fy.Items.AddRange(new object[] {
            "Yesterday",
            "Last week",
            "Last month"});
            this.comboBox_fy.Location = new System.Drawing.Point(69, 15);
            this.comboBox_fy.Name = "comboBox_fy";
            this.comboBox_fy.Size = new System.Drawing.Size(133, 23);
            this.comboBox_fy.TabIndex = 7;
            this.comboBox_fy.SelectedIndexChanged += new System.EventHandler(this.comboBox_fy_SelectedIndexChanged);
            // 
            // dateTimePicker_end_fy
            // 
            this.dateTimePicker_end_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker_end_fy.Location = new System.Drawing.Point(296, 44);
            this.dateTimePicker_end_fy.Name = "dateTimePicker_end_fy";
            this.dateTimePicker_end_fy.Size = new System.Drawing.Size(169, 21);
            this.dateTimePicker_end_fy.TabIndex = 11;
            // 
            // dateTimePicker_start_fy
            // 
            this.dateTimePicker_start_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker_start_fy.Location = new System.Drawing.Point(296, 16);
            this.dateTimePicker_start_fy.Name = "dateTimePicker_start_fy";
            this.dateTimePicker_start_fy.Size = new System.Drawing.Size(169, 21);
            this.dateTimePicker_start_fy.TabIndex = 8;
            // 
            // label_start_fy
            // 
            this.label_start_fy.AutoSize = true;
            this.label_start_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_start_fy.Location = new System.Drawing.Point(225, 20);
            this.label_start_fy.Name = "label_start_fy";
            this.label_start_fy.Size = new System.Drawing.Size(66, 15);
            this.label_start_fy.TabIndex = 9;
            this.label_start_fy.Text = "Start Time:";
            // 
            // label_end_fy
            // 
            this.label_end_fy.AutoSize = true;
            this.label_end_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_end_fy.Location = new System.Drawing.Point(225, 49);
            this.label_end_fy.Name = "label_end_fy";
            this.label_end_fy.Size = new System.Drawing.Size(63, 15);
            this.label_end_fy.TabIndex = 10;
            this.label_end_fy.Text = "End Time:";
            // 
            // button_fy_stop
            // 
            this.button_fy_stop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button_fy_stop.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button_fy_stop.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_fy_stop.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_fy_stop.ForeColor = System.Drawing.Color.White;
            this.button_fy_stop.Image = ((System.Drawing.Image)(resources.GetObject("button_fy_stop.Image")));
            this.button_fy_stop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_fy_stop.Location = new System.Drawing.Point(204, 290);
            this.button_fy_stop.Name = "button_fy_stop";
            this.button_fy_stop.Padding = new System.Windows.Forms.Padding(15, 0, 22, 0);
            this.button_fy_stop.Size = new System.Drawing.Size(153, 59);
            this.button_fy_stop.TabIndex = 33;
            this.button_fy_stop.Text = "STOP";
            this.button_fy_stop.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_fy_stop.UseVisualStyleBackColor = false;
            this.button_fy_stop.Visible = false;
            this.button_fy_stop.Click += new System.EventHandler(this.button_fy_stop_Click);
            // 
            // button_fy_start
            // 
            this.button_fy_start.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button_fy_start.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button_fy_start.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_fy_start.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_fy_start.ForeColor = System.Drawing.Color.White;
            this.button_fy_start.Image = ((System.Drawing.Image)(resources.GetObject("button_fy_start.Image")));
            this.button_fy_start.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_fy_start.Location = new System.Drawing.Point(204, 290);
            this.button_fy_start.Name = "button_fy_start";
            this.button_fy_start.Padding = new System.Windows.Forms.Padding(15, 0, 15, 0);
            this.button_fy_start.Size = new System.Drawing.Size(153, 59);
            this.button_fy_start.TabIndex = 1;
            this.button_fy_start.Text = "START";
            this.button_fy_start.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_fy_start.UseVisualStyleBackColor = false;
            this.button_fy_start.Visible = false;
            this.button_fy_start.Click += new System.EventHandler(this.button_fy_start_ClickAsync);
            // 
            // timer_fy_detect_inserted_in_excel
            // 
            this.timer_fy_detect_inserted_in_excel.Interval = 10000;
            this.timer_fy_detect_inserted_in_excel.Tick += new System.EventHandler(this.timer_fy_detect_inserted_in_excel_Tick);
            // 
            // timer_fy_start
            // 
            this.timer_fy_start.Interval = 30000;
            this.timer_fy_start.Tick += new System.EventHandler(this.timer_fy_start_Tick);
            // 
            // button_filelocation
            // 
            this.button_filelocation.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button_filelocation.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_filelocation.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_filelocation.ForeColor = System.Drawing.Color.White;
            this.button_filelocation.Image = global::FY_Cronos_Data.Properties.Resources.folder;
            this.button_filelocation.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_filelocation.Location = new System.Drawing.Point(217, 52);
            this.button_filelocation.Name = "button_filelocation";
            this.button_filelocation.Padding = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.button_filelocation.Size = new System.Drawing.Size(134, 30);
            this.button_filelocation.TabIndex = 6;
            this.button_filelocation.Text = "File Location";
            this.button_filelocation.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_filelocation.UseVisualStyleBackColor = false;
            this.button_filelocation.Click += new System.EventHandler(this.button_filelocation_Click);
            // 
            // timer_fy
            // 
            this.timer_fy.Interval = 1000;
            this.timer_fy.Tick += new System.EventHandler(this.timer_fy_Tick);
            // 
            // panel_footer
            // 
            this.panel_footer.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.panel_footer.Controls.Add(this.label_version);
            this.panel_footer.Controls.Add(this.label_updates);
            this.panel_footer.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel_footer.Location = new System.Drawing.Point(0, 505);
            this.panel_footer.Name = "panel_footer";
            this.panel_footer.Size = new System.Drawing.Size(569, 20);
            this.panel_footer.TabIndex = 4;
            // 
            // label_version
            // 
            this.label_version.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label_version.AutoSize = true;
            this.label_version.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_version.ForeColor = System.Drawing.Color.White;
            this.label_version.Location = new System.Drawing.Point(513, 4);
            this.label_version.Name = "label_version";
            this.label_version.Size = new System.Drawing.Size(43, 13);
            this.label_version.TabIndex = 1;
            this.label_version.Text = "v1.0.1";
            this.label_version.Visible = false;
            // 
            // label_updates
            // 
            this.label_updates.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label_updates.AutoSize = true;
            this.label_updates.Cursor = System.Windows.Forms.Cursors.Hand;
            this.label_updates.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_updates.ForeColor = System.Drawing.Color.White;
            this.label_updates.Location = new System.Drawing.Point(406, 3);
            this.label_updates.Name = "label_updates";
            this.label_updates.Size = new System.Drawing.Size(99, 13);
            this.label_updates.TabIndex = 0;
            this.label_updates.Text = "Check for Updates.";
            this.label_updates.Visible = false;
            this.label_updates.Click += new System.EventHandler(this.label_updates_Click);
            // 
            // panel_landing
            // 
            this.panel_landing.Controls.Add(this.pictureBox_landing);
            this.panel_landing.Location = new System.Drawing.Point(1, 19);
            this.panel_landing.Name = "panel_landing";
            this.panel_landing.Size = new System.Drawing.Size(567, 485);
            this.panel_landing.TabIndex = 31;
            // 
            // pictureBox_landing
            // 
            this.pictureBox_landing.Image = global::FY_Cronos_Data.Properties.Resources.icon;
            this.pictureBox_landing.Location = new System.Drawing.Point(226, 188);
            this.pictureBox_landing.Name = "pictureBox_landing";
            this.pictureBox_landing.Size = new System.Drawing.Size(135, 134);
            this.pictureBox_landing.TabIndex = 0;
            this.pictureBox_landing.TabStop = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            this.dataGridView1.Location = new System.Drawing.Point(1167, 55);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(242, 170);
            this.dataGridView1.TabIndex = 30;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            // 
            // timer_landing
            // 
            this.timer_landing.Interval = 2000;
            this.timer_landing.Tick += new System.EventHandler(this.timer_landing_Tick);
            // 
            // timer_fy_start_button
            // 
            this.timer_fy_start_button.Interval = 1000;
            this.timer_fy_start_button.Tick += new System.EventHandler(this.timer_fy_start_button_TickAsync);
            // 
            // label_fy_count
            // 
            this.label_fy_count.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_count.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.label_fy_count.Location = new System.Drawing.Point(38, 350);
            this.label_fy_count.Name = "label_fy_count";
            this.label_fy_count.Size = new System.Drawing.Size(498, 38);
            this.label_fy_count.TabIndex = 0;
            this.label_fy_count.Text = "-";
            this.label_fy_count.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_fy_count.Visible = false;
            // 
            // label_getdatacount_fy
            // 
            this.label_getdatacount_fy.Location = new System.Drawing.Point(17, 57);
            this.label_getdatacount_fy.Name = "label_getdatacount_fy";
            this.label_getdatacount_fy.Size = new System.Drawing.Size(534, 29);
            this.label_getdatacount_fy.TabIndex = 34;
            this.label_getdatacount_fy.Text = "-";
            this.label_getdatacount_fy.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(476, 57);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 35;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(420, 57);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 36;
            this.label1.Text = "label1";
            this.label1.Visible = false;
            // 
            // panel
            // 
            this.panel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(30)))), ((int)(((byte)(112)))));
            this.panel.Location = new System.Drawing.Point(-12, -5);
            this.panel.Name = "panel";
            this.panel.Size = new System.Drawing.Size(170, 10);
            this.panel.TabIndex = 1;
            this.panel.Visible = false;
            // 
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(569, 525);
            this.Controls.Add(this.panel_landing);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button_filelocation);
            this.Controls.Add(this.label_getdatacount_fy);
            this.Controls.Add(this.button_fy_start);
            this.Controls.Add(this.label_fy_count);
            this.Controls.Add(this.button_fy_stop);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel_footer);
            this.Controls.Add(this.panel_fy);
            this.Controls.Add(this.panel_header);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Main_Form";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FY Cronos Data";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Main_Form_FormClosing);
            this.Load += new System.EventHandler(this.Main_Form_Load);
            this.Shown += new System.EventHandler(this.Main_Form_Shown);
            this.panel_header.ResumeLayout(false);
            this.panel_header.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_minimize)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_close)).EndInit();
            this.panel_fy.ResumeLayout(false);
            this.panel_fy_status.ResumeLayout(false);
            this.panel_fy_status.PerformLayout();
            this.panel_fy_datetime.ResumeLayout(false);
            this.panel_fy_datetime.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_fy_loader)).EndInit();
            this.panel_fy_filter.ResumeLayout(false);
            this.panel_fy_filter.PerformLayout();
            this.panel_footer.ResumeLayout(false);
            this.panel_footer.PerformLayout();
            this.panel_landing.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_landing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.Label label_title_fy;
        private System.Windows.Forms.Panel panel_fy;
        private System.Windows.Forms.WebBrowser webBrowser_fy;
        private System.Windows.Forms.Label label_title;
        private System.Windows.Forms.Button button_filelocation;
        private System.Windows.Forms.ComboBox comboBox_fy;
        private System.Windows.Forms.DateTimePicker dateTimePicker_start_fy;
        private System.Windows.Forms.Label label_start_fy;
        private System.Windows.Forms.Label label_end_fy;
        private System.Windows.Forms.DateTimePicker dateTimePicker_end_fy;
        private System.Windows.Forms.Timer timer_fy_detect_inserted_in_excel;
        private System.Windows.Forms.Button button_fy_start;
        private System.Windows.Forms.Timer timer_fy_start;
        private System.Windows.Forms.Panel panel_fy_filter;
        private System.Windows.Forms.Timer timer_fy;
        private System.Windows.Forms.PictureBox pictureBox_minimize;
        private System.Windows.Forms.PictureBox pictureBox_close;
        private System.Windows.Forms.Label label_filelocation;
        private System.Windows.Forms.Panel panel_footer;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.Timer timer_tf_detect_inserted_in_excel;
        private System.Windows.Forms.Timer timer_tf_start;
        private System.Windows.Forms.Timer timer_tf;
        private System.Windows.Forms.Label label_updates;
        private System.Windows.Forms.Label label_version;
        private System.Windows.Forms.Panel panel_landing;
        private System.Windows.Forms.PictureBox pictureBox_landing;
        private System.Windows.Forms.Timer timer_landing;
        private System.Windows.Forms.ComboBox comboBox_fy_list;
        private System.Windows.Forms.Timer timer_fy_start_button;
        private System.Windows.Forms.Button button_fy_stop;
        private System.Windows.Forms.Label label_fy_count;
        private System.Windows.Forms.Label label_getdatacount_fy;
        private System.Windows.Forms.Panel panel_fy_status;
        private System.Windows.Forms.Label label_fy_insert;
        private System.Windows.Forms.Button button_fy_proceed;
        private System.Windows.Forms.Label label_fy_locatefolder;
        private System.Windows.Forms.Panel panel_fy_datetime;
        private System.Windows.Forms.Label label_fy_elapsed;
        private System.Windows.Forms.Label label_fy_elapsed_1;
        private System.Windows.Forms.Label label_fy_start_datetime_1;
        private System.Windows.Forms.Label label_fy_finish_datetime;
        private System.Windows.Forms.Label label_fy_finish_datetime_1;
        private System.Windows.Forms.Label label_fy_start_datetime;
        private System.Windows.Forms.PictureBox pictureBox_fy_loader;
        private System.Windows.Forms.Label label_fy_currentrecord;
        private System.Windows.Forms.Label label_fy_page_count;
        private System.Windows.Forms.Label label_fy_page_count_1;
        private System.Windows.Forms.Label label_fy_total_records_1;
        private System.Windows.Forms.Label label_fy_status;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel;
    }
}