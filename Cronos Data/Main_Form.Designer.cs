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
            this.panel_header = new System.Windows.Forms.Panel();
            this.label_title = new System.Windows.Forms.Label();
            this.pictureBox_minimize = new System.Windows.Forms.PictureBox();
            this.pictureBox_close = new System.Windows.Forms.PictureBox();
            this.label_filelocation = new System.Windows.Forms.Label();
            this.label_title_fy = new System.Windows.Forms.Label();
            this.label_title_tf = new System.Windows.Forms.Label();
            this.panel_fy = new System.Windows.Forms.Panel();
            this.dateTimePicker_end_fy = new System.Windows.Forms.DateTimePicker();
            this.label_end_fy = new System.Windows.Forms.Label();
            this.label_start_fy = new System.Windows.Forms.Label();
            this.dateTimePicker_start_fy = new System.Windows.Forms.DateTimePicker();
            this.comboBox_fy = new System.Windows.Forms.ComboBox();
            this.webBrowser_fy = new System.Windows.Forms.WebBrowser();
            this.panel_tf = new System.Windows.Forms.Panel();
            this.webBrowser_tf = new System.Windows.Forms.WebBrowser();
            this.button_filelocation = new System.Windows.Forms.Button();
            this.sPanel1 = new Cronos_Data.SPanel();
            this.button1 = new System.Windows.Forms.Button();
            this.panel_header.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_minimize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_close)).BeginInit();
            this.panel_fy.SuspendLayout();
            this.panel_tf.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.panel_header.Controls.Add(this.button1);
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
            // 
            // pictureBox_minimize
            // 
            this.pictureBox_minimize.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.pictureBox_minimize.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox_minimize.Image = global::Cronos_Data.Properties.Resources.minus;
            this.pictureBox_minimize.Location = new System.Drawing.Point(1065, 10);
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
            this.pictureBox_close.Location = new System.Drawing.Point(1104, 10);
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
            this.panel_fy.Controls.Add(this.dateTimePicker_end_fy);
            this.panel_fy.Controls.Add(this.label_end_fy);
            this.panel_fy.Controls.Add(this.label_start_fy);
            this.panel_fy.Controls.Add(this.dateTimePicker_start_fy);
            this.panel_fy.Controls.Add(this.comboBox_fy);
            this.panel_fy.Controls.Add(this.webBrowser_fy);
            this.panel_fy.Controls.Add(this.label_title_fy);
            this.panel_fy.Location = new System.Drawing.Point(16, 80);
            this.panel_fy.Name = "panel_fy";
            this.panel_fy.Size = new System.Drawing.Size(534, 570);
            this.panel_fy.TabIndex = 4;
            this.panel_fy.Paint += new System.Windows.Forms.PaintEventHandler(this.panel_fy_Paint);
            // 
            // dateTimePicker_end_fy
            // 
            this.dateTimePicker_end_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker_end_fy.Location = new System.Drawing.Point(282, 76);
            this.dateTimePicker_end_fy.Name = "dateTimePicker_end_fy";
            this.dateTimePicker_end_fy.Size = new System.Drawing.Size(169, 21);
            this.dateTimePicker_end_fy.TabIndex = 11;
            // 
            // label_end_fy
            // 
            this.label_end_fy.AutoSize = true;
            this.label_end_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_end_fy.Location = new System.Drawing.Point(211, 81);
            this.label_end_fy.Name = "label_end_fy";
            this.label_end_fy.Size = new System.Drawing.Size(63, 15);
            this.label_end_fy.TabIndex = 10;
            this.label_end_fy.Text = "End Time:";
            // 
            // label_start_fy
            // 
            this.label_start_fy.AutoSize = true;
            this.label_start_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_start_fy.Location = new System.Drawing.Point(211, 52);
            this.label_start_fy.Name = "label_start_fy";
            this.label_start_fy.Size = new System.Drawing.Size(66, 15);
            this.label_start_fy.TabIndex = 9;
            this.label_start_fy.Text = "Start Time:";
            // 
            // dateTimePicker_start_fy
            // 
            this.dateTimePicker_start_fy.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker_start_fy.Location = new System.Drawing.Point(282, 48);
            this.dateTimePicker_start_fy.Name = "dateTimePicker_start_fy";
            this.dateTimePicker_start_fy.Size = new System.Drawing.Size(169, 21);
            this.dateTimePicker_start_fy.TabIndex = 8;
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
            this.comboBox_fy.Location = new System.Drawing.Point(80, 50);
            this.comboBox_fy.Name = "comboBox_fy";
            this.comboBox_fy.Size = new System.Drawing.Size(108, 23);
            this.comboBox_fy.TabIndex = 7;
            this.comboBox_fy.SelectedIndexChanged += new System.EventHandler(this.comboBox_fy_SelectedIndexChanged);
            // 
            // webBrowser_fy
            // 
            this.webBrowser_fy.Location = new System.Drawing.Point(3, 109);
            this.webBrowser_fy.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser_fy.Name = "webBrowser_fy";
            this.webBrowser_fy.ScriptErrorsSuppressed = true;
            this.webBrowser_fy.Size = new System.Drawing.Size(528, 458);
            this.webBrowser_fy.TabIndex = 0;
            this.webBrowser_fy.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser_fy_DocumentCompletedAsync);
            // 
            // panel_tf
            // 
            this.panel_tf.Controls.Add(this.webBrowser_tf);
            this.panel_tf.Location = new System.Drawing.Point(588, 80);
            this.panel_tf.Name = "panel_tf";
            this.panel_tf.Size = new System.Drawing.Size(534, 570);
            this.panel_tf.TabIndex = 5;
            this.panel_tf.Paint += new System.Windows.Forms.PaintEventHandler(this.panel_tf_Paint);
            // 
            // webBrowser_tf
            // 
            this.webBrowser_tf.Location = new System.Drawing.Point(3, 3);
            this.webBrowser_tf.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser_tf.Name = "webBrowser_tf";
            this.webBrowser_tf.ScriptErrorsSuppressed = true;
            this.webBrowser_tf.Size = new System.Drawing.Size(528, 564);
            this.webBrowser_tf.TabIndex = 0;
            // 
            // button_filelocation
            // 
            this.button_filelocation.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button_filelocation.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_filelocation.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_filelocation.ForeColor = System.Drawing.Color.White;
            this.button_filelocation.Location = new System.Drawing.Point(498, 51);
            this.button_filelocation.Name = "button_filelocation";
            this.button_filelocation.Size = new System.Drawing.Size(143, 24);
            this.button_filelocation.TabIndex = 6;
            this.button_filelocation.Text = "File Location";
            this.button_filelocation.UseVisualStyleBackColor = false;
            this.button_filelocation.Click += new System.EventHandler(this.button_filelocation_Click);
            // 
            // sPanel1
            // 
            this.sPanel1.BackColor = System.Drawing.Color.Transparent;
            this.sPanel1.Cursor = System.Windows.Forms.Cursors.Default;
            this.sPanel1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.sPanel1.Location = new System.Drawing.Point(557, 71);
            this.sPanel1.Name = "sPanel1";
            this.sPanel1.Size = new System.Drawing.Size(44, 625);
            this.sPanel1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(349, 16);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1140, 664);
            this.Controls.Add(this.button_filelocation);
            this.Controls.Add(this.panel_tf);
            this.Controls.Add(this.panel_fy);
            this.Controls.Add(this.sPanel1);
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
            this.panel_fy.PerformLayout();
            this.panel_tf.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.PictureBox pictureBox_close;
        private System.Windows.Forms.PictureBox pictureBox_minimize;
        private SPanel sPanel1;
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
        private System.Windows.Forms.Button button1;
    }
}