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
            this.label_title_tf = new System.Windows.Forms.Label();
            this.panel_fy = new System.Windows.Forms.Panel();
            this.panel_fy_status = new System.Windows.Forms.Panel();
            this.button_fy_proceed = new System.Windows.Forms.Button();
            this.label_fy_locatefolder = new System.Windows.Forms.Label();
            this.panel_datetime = new System.Windows.Forms.Panel();
            this.label_fy_start_datetime_1 = new System.Windows.Forms.Label();
            this.label_fy_finish_datetime = new System.Windows.Forms.Label();
            this.label_fy_finish_datetime_1 = new System.Windows.Forms.Label();
            this.label_fy_start_datetime = new System.Windows.Forms.Label();
            this.pictureBox_fy_loader = new System.Windows.Forms.PictureBox();
            this.label_fy_currentrecord = new System.Windows.Forms.Label();
            this.label_fy_inserting_count_1 = new System.Windows.Forms.Label();
            this.label_fy_page_count = new System.Windows.Forms.Label();
            this.label_fy_page_count_1 = new System.Windows.Forms.Label();
            this.label_fy_inserting_count = new System.Windows.Forms.Label();
            this.label_fy_total_records_1 = new System.Windows.Forms.Label();
            this.label_fy_status = new System.Windows.Forms.Label();
            this.webBrowser_fy = new System.Windows.Forms.WebBrowser();
            this.panel = new System.Windows.Forms.Panel();
            this.comboBox_fy = new System.Windows.Forms.ComboBox();
            this.dateTimePicker_end_fy = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker_start_fy = new System.Windows.Forms.DateTimePicker();
            this.label_start_fy = new System.Windows.Forms.Label();
            this.label_end_fy = new System.Windows.Forms.Label();
            this.button_fy_start = new System.Windows.Forms.Button();
            this.panel_tf = new System.Windows.Forms.Panel();
            this.webBrowser_tf = new System.Windows.Forms.WebBrowser();
            this.timer_fy_detect_inserted_in_excel = new System.Windows.Forms.Timer(this.components);
            this.timer_fy_start = new System.Windows.Forms.Timer(this.components);
            this.button_filelocation = new System.Windows.Forms.Button();
            this.label_fy_elapsed_1 = new System.Windows.Forms.Label();
            this.label_fy_elapsed = new System.Windows.Forms.Label();
            this.sPanel_separator = new Cronos_Data.SPanel();
            this.panel_header.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_minimize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_close)).BeginInit();
            this.panel_fy.SuspendLayout();
            this.panel_fy_status.SuspendLayout();
            this.panel_datetime.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_fy_loader)).BeginInit();
            this.panel.SuspendLayout();
            this.panel_tf.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.panel_header.Controls.Add(this.label_title);
            this.panel_header.Controls.Add(this.pictureBox_minimize);
            this.panel_header.Controls.Add(this.pictureBox_close);
            this.panel_header.Controls.Add(this.label_filelocation);
            this.panel_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_header.Location = new System.Drawing.Point(0, 0);
            this.panel_header.Name = "panel_header";
            this.panel_header.Size = new System.Drawing.Size(1140, 45);
            this.panel_header.TabIndex = 0;
            this.panel_header.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel_header_MouseDown);
            // 
            // label_title
            // 
            this.label_title.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_title.ForeColor = System.Drawing.Color.White;
            this.label_title.Location = new System.Drawing.Point(2, 0);
            this.label_title.Name = "label_title";
            this.label_title.Size = new System.Drawing.Size(166, 45);
            this.label_title.TabIndex = 2;
            this.label_title.Text = "Cronos Data";
            this.label_title.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_title.MouseDown += new System.Windows.Forms.MouseEventHandler(this.label_title_MouseDown);
            // 
            // pictureBox_minimize
            // 
            this.pictureBox_minimize.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.pictureBox_minimize.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox_minimize.Image = global::Cronos_Data.Properties.Resources.minus;
            this.pictureBox_minimize.Location = new System.Drawing.Point(1052, 10);
            this.pictureBox_minimize.Name = "pictureBox_minimize";
            this.pictureBox_minimize.Size = new System.Drawing.Size(24, 24);
            this.pictureBox_minimize.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox_minimize.TabIndex = 1;
            this.pictureBox_minimize.TabStop = false;
            this.pictureBox_minimize.Click += new System.EventHandler(this.pictureBox_minimize_Click);
            // 
            // pictureBox_close
            // 
            this.pictureBox_close.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.pictureBox_close.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox_close.Image = global::Cronos_Data.Properties.Resources.close;
            this.pictureBox_close.Location = new System.Drawing.Point(1091, 10);
            this.pictureBox_close.Name = "pictureBox_close";
            this.pictureBox_close.Size = new System.Drawing.Size(24, 24);
            this.pictureBox_close.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox_close.TabIndex = 0;
            this.pictureBox_close.TabStop = false;
            this.pictureBox_close.Click += new System.EventHandler(this.pictureBox_close_Click);
            // 
            // label_filelocation
            // 
            this.label_filelocation.ForeColor = System.Drawing.Color.White;
            this.label_filelocation.Location = new System.Drawing.Point(0, 16);
            this.label_filelocation.Name = "label_filelocation";
            this.label_filelocation.Size = new System.Drawing.Size(1140, 13);
            this.label_filelocation.TabIndex = 3;
            this.label_filelocation.Text = "-";
            this.label_filelocation.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
            // label_title_tf
            // 
            this.label_title_tf.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_title_tf.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(155)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.label_title_tf.Location = new System.Drawing.Point(572, 51);
            this.label_title_tf.Name = "label_title_tf";
            this.label_title_tf.Size = new System.Drawing.Size(556, 24);
            this.label_title_tf.TabIndex = 3;
            this.label_title_tf.Text = "Tian Fa";
            this.label_title_tf.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel_fy
            // 
            this.panel_fy.Controls.Add(this.panel_fy_status);
            this.panel_fy.Controls.Add(this.webBrowser_fy);
            this.panel_fy.Controls.Add(this.panel);
            this.panel_fy.Controls.Add(this.label_title_fy);
            this.panel_fy.Location = new System.Drawing.Point(16, 80);
            this.panel_fy.Name = "panel_fy";
            this.panel_fy.Size = new System.Drawing.Size(534, 408);
            this.panel_fy.TabIndex = 4;
            this.panel_fy.Paint += new System.Windows.Forms.PaintEventHandler(this.panel_fy_Paint);
            // 
            // panel_fy_status
            // 
            this.panel_fy_status.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.panel_fy_status.Controls.Add(this.button_fy_proceed);
            this.panel_fy_status.Controls.Add(this.label_fy_locatefolder);
            this.panel_fy_status.Controls.Add(this.panel_datetime);
            this.panel_fy_status.Controls.Add(this.pictureBox_fy_loader);
            this.panel_fy_status.Controls.Add(this.label_fy_currentrecord);
            this.panel_fy_status.Controls.Add(this.label_fy_inserting_count_1);
            this.panel_fy_status.Controls.Add(this.label_fy_page_count);
            this.panel_fy_status.Controls.Add(this.label_fy_page_count_1);
            this.panel_fy_status.Controls.Add(this.label_fy_inserting_count);
            this.panel_fy_status.Controls.Add(this.label_fy_total_records_1);
            this.panel_fy_status.Controls.Add(this.label_fy_status);
            this.panel_fy_status.Location = new System.Drawing.Point(7, 121);
            this.panel_fy_status.Name = "panel_fy_status";
            this.panel_fy_status.Size = new System.Drawing.Size(524, 284);
            this.panel_fy_status.TabIndex = 23;
            this.panel_fy_status.Visible = false;
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
            // panel_datetime
            // 
            this.panel_datetime.Controls.Add(this.label_fy_elapsed);
            this.panel_datetime.Controls.Add(this.label_fy_elapsed_1);
            this.panel_datetime.Controls.Add(this.label_fy_start_datetime_1);
            this.panel_datetime.Controls.Add(this.label_fy_finish_datetime);
            this.panel_datetime.Controls.Add(this.label_fy_finish_datetime_1);
            this.panel_datetime.Controls.Add(this.label_fy_start_datetime);
            this.panel_datetime.Location = new System.Drawing.Point(66, 226);
            this.panel_datetime.Name = "panel_datetime";
            this.panel_datetime.Size = new System.Drawing.Size(204, 58);
            this.panel_datetime.TabIndex = 28;
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
            this.pictureBox_fy_loader.Image = global::Cronos_Data.Properties.Resources.loader;
            this.pictureBox_fy_loader.Location = new System.Drawing.Point(3, 180);
            this.pictureBox_fy_loader.Name = "pictureBox_fy_loader";
            this.pictureBox_fy_loader.Size = new System.Drawing.Size(60, 101);
            this.pictureBox_fy_loader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox_fy_loader.TabIndex = 23;
            this.pictureBox_fy_loader.TabStop = false;
            // 
            // label_fy_currentrecord
            // 
            this.label_fy_currentrecord.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_currentrecord.Location = new System.Drawing.Point(258, 121);
            this.label_fy_currentrecord.Name = "label_fy_currentrecord";
            this.label_fy_currentrecord.Size = new System.Drawing.Size(151, 18);
            this.label_fy_currentrecord.TabIndex = 12;
            this.label_fy_currentrecord.Text = "-";
            // 
            // label_fy_inserting_count_1
            // 
            this.label_fy_inserting_count_1.AutoSize = true;
            this.label_fy_inserting_count_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_inserting_count_1.Location = new System.Drawing.Point(134, 150);
            this.label_fy_inserting_count_1.Name = "label_fy_inserting_count_1";
            this.label_fy_inserting_count_1.Size = new System.Drawing.Size(92, 18);
            this.label_fy_inserting_count_1.TabIndex = 21;
            this.label_fy_inserting_count_1.Text = "Insert Count:";
            this.label_fy_inserting_count_1.Visible = false;
            // 
            // label_fy_page_count
            // 
            this.label_fy_page_count.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_page_count.Location = new System.Drawing.Point(259, 91);
            this.label_fy_page_count.Name = "label_fy_page_count";
            this.label_fy_page_count.Size = new System.Drawing.Size(151, 18);
            this.label_fy_page_count.TabIndex = 13;
            this.label_fy_page_count.Text = "-";
            // 
            // label_fy_page_count_1
            // 
            this.label_fy_page_count_1.AutoSize = true;
            this.label_fy_page_count_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_page_count_1.Location = new System.Drawing.Point(134, 91);
            this.label_fy_page_count_1.Name = "label_fy_page_count_1";
            this.label_fy_page_count_1.Size = new System.Drawing.Size(46, 18);
            this.label_fy_page_count_1.TabIndex = 20;
            this.label_fy_page_count_1.Text = "Page:";
            // 
            // label_fy_inserting_count
            // 
            this.label_fy_inserting_count.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_inserting_count.Location = new System.Drawing.Point(258, 150);
            this.label_fy_inserting_count.Name = "label_fy_inserting_count";
            this.label_fy_inserting_count.Size = new System.Drawing.Size(151, 18);
            this.label_fy_inserting_count.TabIndex = 15;
            this.label_fy_inserting_count.Text = "-";
            this.label_fy_inserting_count.Visible = false;
            // 
            // label_fy_total_records_1
            // 
            this.label_fy_total_records_1.AutoSize = true;
            this.label_fy_total_records_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fy_total_records_1.Location = new System.Drawing.Point(132, 121);
            this.label_fy_total_records_1.Name = "label_fy_total_records_1";
            this.label_fy_total_records_1.Size = new System.Drawing.Size(98, 18);
            this.label_fy_total_records_1.TabIndex = 18;
            this.label_fy_total_records_1.Text = "Total Record:";
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
            // panel
            // 
            this.panel.Controls.Add(this.comboBox_fy);
            this.panel.Controls.Add(this.dateTimePicker_end_fy);
            this.panel.Controls.Add(this.dateTimePicker_start_fy);
            this.panel.Controls.Add(this.label_start_fy);
            this.panel.Controls.Add(this.label_end_fy);
            this.panel.Location = new System.Drawing.Point(3, 35);
            this.panel.Name = "panel";
            this.panel.Size = new System.Drawing.Size(528, 80);
            this.panel.TabIndex = 24;
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
            this.comboBox_fy.Location = new System.Drawing.Point(76, 18);
            this.comboBox_fy.Name = "comboBox_fy";
            this.comboBox_fy.Size = new System.Drawing.Size(108, 23);
            this.comboBox_fy.TabIndex = 7;
            this.comboBox_fy.SelectedIndexChanged += new System.EventHandler(this.comboBox_fy_SelectedIndexChanged);
            // 
            // dateTimePicker_end_fy
            // 
            this.dateTimePicker_end_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker_end_fy.Location = new System.Drawing.Point(278, 44);
            this.dateTimePicker_end_fy.Name = "dateTimePicker_end_fy";
            this.dateTimePicker_end_fy.Size = new System.Drawing.Size(169, 21);
            this.dateTimePicker_end_fy.TabIndex = 11;
            // 
            // dateTimePicker_start_fy
            // 
            this.dateTimePicker_start_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker_start_fy.Location = new System.Drawing.Point(278, 16);
            this.dateTimePicker_start_fy.Name = "dateTimePicker_start_fy";
            this.dateTimePicker_start_fy.Size = new System.Drawing.Size(169, 21);
            this.dateTimePicker_start_fy.TabIndex = 8;
            // 
            // label_start_fy
            // 
            this.label_start_fy.AutoSize = true;
            this.label_start_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_start_fy.Location = new System.Drawing.Point(207, 20);
            this.label_start_fy.Name = "label_start_fy";
            this.label_start_fy.Size = new System.Drawing.Size(66, 15);
            this.label_start_fy.TabIndex = 9;
            this.label_start_fy.Text = "Start Time:";
            // 
            // label_end_fy
            // 
            this.label_end_fy.AutoSize = true;
            this.label_end_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_end_fy.Location = new System.Drawing.Point(207, 49);
            this.label_end_fy.Name = "label_end_fy";
            this.label_end_fy.Size = new System.Drawing.Size(63, 15);
            this.label_end_fy.TabIndex = 10;
            this.label_end_fy.Text = "End Time:";
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
            this.button_fy_start.Location = new System.Drawing.Point(204, 275);
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
            // panel_tf
            // 
            this.panel_tf.Controls.Add(this.webBrowser_tf);
            this.panel_tf.Location = new System.Drawing.Point(588, 80);
            this.panel_tf.Name = "panel_tf";
            this.panel_tf.Size = new System.Drawing.Size(534, 408);
            this.panel_tf.TabIndex = 5;
            this.panel_tf.Paint += new System.Windows.Forms.PaintEventHandler(this.panel_tf_Paint);
            // 
            // webBrowser_tf
            // 
            this.webBrowser_tf.Location = new System.Drawing.Point(4, 3);
            this.webBrowser_tf.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser_tf.Name = "webBrowser_tf";
            this.webBrowser_tf.ScriptErrorsSuppressed = true;
            this.webBrowser_tf.Size = new System.Drawing.Size(528, 402);
            this.webBrowser_tf.TabIndex = 0;
            // 
            // timer_fy_detect_inserted_in_excel
            // 
            this.timer_fy_detect_inserted_in_excel.Interval = 10000;
            this.timer_fy_detect_inserted_in_excel.Tick += new System.EventHandler(this.timer_fy_detect_inserted_in_excel_Tick);
            // 
            // timer_fy_start
            // 
            this.timer_fy_start.Interval = 10000;
            this.timer_fy_start.Tick += new System.EventHandler(this.timer_fy_start_Tick);
            // 
            // button_filelocation
            // 
            this.button_filelocation.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button_filelocation.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_filelocation.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_filelocation.ForeColor = System.Drawing.Color.White;
            this.button_filelocation.Image = global::Cronos_Data.Properties.Resources.folder;
            this.button_filelocation.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_filelocation.Location = new System.Drawing.Point(498, 47);
            this.button_filelocation.Name = "button_filelocation";
            this.button_filelocation.Padding = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.button_filelocation.Size = new System.Drawing.Size(134, 30);
            this.button_filelocation.TabIndex = 6;
            this.button_filelocation.Text = "File Location";
            this.button_filelocation.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_filelocation.UseVisualStyleBackColor = false;
            this.button_filelocation.Click += new System.EventHandler(this.button_filelocation_Click);
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
            // sPanel_separator
            // 
            this.sPanel_separator.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.sPanel_separator.BackColor = System.Drawing.Color.Transparent;
            this.sPanel_separator.Cursor = System.Windows.Forms.Cursors.Default;
            this.sPanel_separator.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.sPanel_separator.Location = new System.Drawing.Point(557, 75);
            this.sPanel_separator.Name = "sPanel_separator";
            this.sPanel_separator.Size = new System.Drawing.Size(44, 459);
            this.sPanel_separator.TabIndex = 1;
            // 
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1140, 495);
            this.Controls.Add(this.button_fy_start);
            this.Controls.Add(this.button_filelocation);
            this.Controls.Add(this.panel_tf);
            this.Controls.Add(this.panel_fy);
            this.Controls.Add(this.sPanel_separator);
            this.Controls.Add(this.panel_header);
            this.Controls.Add(this.label_title_tf);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Main_Form";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Cronos Data";
            this.Load += new System.EventHandler(this.Main_Form_Load);
            this.Shown += new System.EventHandler(this.Main_Form_Shown);
            this.panel_header.ResumeLayout(false);
            this.panel_header.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_minimize)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_close)).EndInit();
            this.panel_fy.ResumeLayout(false);
            this.panel_fy_status.ResumeLayout(false);
            this.panel_fy_status.PerformLayout();
            this.panel_datetime.ResumeLayout(false);
            this.panel_datetime.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_fy_loader)).EndInit();
            this.panel.ResumeLayout(false);
            this.panel.PerformLayout();
            this.panel_tf.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.PictureBox pictureBox_close;
        private System.Windows.Forms.PictureBox pictureBox_minimize;
        private System.Windows.Forms.Label label_title_fy;
        private System.Windows.Forms.Label label_title_tf;
        private System.Windows.Forms.Panel panel_fy;
        private System.Windows.Forms.WebBrowser webBrowser_fy;
        private System.Windows.Forms.Label label_title;
        private System.Windows.Forms.Panel panel_tf;
        private System.Windows.Forms.WebBrowser webBrowser_tf;
        private System.Windows.Forms.Button button_filelocation;
        private System.Windows.Forms.Label label_filelocation;
        private System.Windows.Forms.ComboBox comboBox_fy;
        private System.Windows.Forms.DateTimePicker dateTimePicker_start_fy;
        private System.Windows.Forms.Label label_start_fy;
        private System.Windows.Forms.Label label_end_fy;
        private System.Windows.Forms.DateTimePicker dateTimePicker_end_fy;
        private System.Windows.Forms.Label label_fy_currentrecord;
        private System.Windows.Forms.Label label_fy_page_count;
        private System.Windows.Forms.Timer timer_fy_detect_inserted_in_excel;
        private System.Windows.Forms.Label label_fy_inserting_count;
        private System.Windows.Forms.Label label_fy_status;
        private System.Windows.Forms.Label label_fy_total_records_1;
        private System.Windows.Forms.Label label_fy_inserting_count_1;
        private System.Windows.Forms.Label label_fy_page_count_1;
        private System.Windows.Forms.Button button_fy_start;
        private System.Windows.Forms.Timer timer_fy_start;
        private System.Windows.Forms.Panel panel_fy_status;
        private System.Windows.Forms.Panel panel;
        private System.Windows.Forms.PictureBox pictureBox_fy_loader;
        private System.Windows.Forms.Label label_fy_finish_datetime_1;
        private System.Windows.Forms.Label label_fy_start_datetime_1;
        private System.Windows.Forms.Label label_fy_start_datetime;
        private System.Windows.Forms.Label label_fy_finish_datetime;
        private System.Windows.Forms.Panel panel_datetime;
        private System.Windows.Forms.Label label_fy_locatefolder;
        private System.Windows.Forms.Button button_fy_proceed;
        private SPanel sPanel_separator;
        private System.Windows.Forms.Label label_fy_elapsed_1;
        private System.Windows.Forms.Label label_fy_elapsed;
    }
}