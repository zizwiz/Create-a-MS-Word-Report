
namespace Create_a_MS_Word_Report
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tab_defaults = new System.Windows.Forms.TabPage();
            this.chkbx_page_footer = new System.Windows.Forms.CheckBox();
            this.chkbx_page_header = new System.Windows.Forms.CheckBox();
            this.txtbx_proj_name = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tab_header = new System.Windows.Forms.TabPage();
            this.chkbx_header_bold = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbobx_header_underline_style = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbobx_header_colour = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.cmbobx_header_style = new System.Windows.Forms.ComboBox();
            this.cmbobx_header_fontsize = new System.Windows.Forms.ComboBox();
            this.cmbobx_header_fontname = new System.Windows.Forms.ComboBox();
            this.tab_footer = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.btn_create = new System.Windows.Forms.Button();
            this.panel5 = new System.Windows.Forms.Panel();
            this.panel6 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.btn_close = new System.Windows.Forms.Button();
            this.chkbx_header_italic = new System.Windows.Forms.CheckBox();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tab_defaults.SuspendLayout();
            this.tab_header.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel7.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1086, 796);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1080, 730);
            this.panel1.TabIndex = 0;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tab_defaults);
            this.tabControl1.Controls.Add(this.tab_header);
            this.tabControl1.Controls.Add(this.tab_footer);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1080, 730);
            this.tabControl1.TabIndex = 0;
            // 
            // tab_defaults
            // 
            this.tab_defaults.Controls.Add(this.chkbx_page_footer);
            this.tab_defaults.Controls.Add(this.chkbx_page_header);
            this.tab_defaults.Controls.Add(this.txtbx_proj_name);
            this.tab_defaults.Controls.Add(this.label1);
            this.tab_defaults.Location = new System.Drawing.Point(4, 29);
            this.tab_defaults.Name = "tab_defaults";
            this.tab_defaults.Size = new System.Drawing.Size(1072, 697);
            this.tab_defaults.TabIndex = 2;
            this.tab_defaults.Text = "Defaults";
            this.tab_defaults.UseVisualStyleBackColor = true;
            // 
            // chkbx_page_footer
            // 
            this.chkbx_page_footer.AutoSize = true;
            this.chkbx_page_footer.Location = new System.Drawing.Point(56, 154);
            this.chkbx_page_footer.Name = "chkbx_page_footer";
            this.chkbx_page_footer.Size = new System.Drawing.Size(115, 24);
            this.chkbx_page_footer.TabIndex = 5;
            this.chkbx_page_footer.Text = "Use Footer";
            this.chkbx_page_footer.UseVisualStyleBackColor = true;
            // 
            // chkbx_page_header
            // 
            this.chkbx_page_header.AutoSize = true;
            this.chkbx_page_header.Location = new System.Drawing.Point(56, 124);
            this.chkbx_page_header.Name = "chkbx_page_header";
            this.chkbx_page_header.Size = new System.Drawing.Size(121, 24);
            this.chkbx_page_header.TabIndex = 4;
            this.chkbx_page_header.Text = "Use Header";
            this.chkbx_page_header.UseVisualStyleBackColor = true;
            this.chkbx_page_header.CheckedChanged += new System.EventHandler(this.chkbx_page_header_CheckedChanged);
            // 
            // txtbx_proj_name
            // 
            this.txtbx_proj_name.Location = new System.Drawing.Point(157, 49);
            this.txtbx_proj_name.Name = "txtbx_proj_name";
            this.txtbx_proj_name.Size = new System.Drawing.Size(579, 26);
            this.txtbx_proj_name.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(43, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "Project Name:";
            // 
            // tab_header
            // 
            this.tab_header.Controls.Add(this.chkbx_header_italic);
            this.tab_header.Controls.Add(this.chkbx_header_bold);
            this.tab_header.Controls.Add(this.label3);
            this.tab_header.Controls.Add(this.cmbobx_header_underline_style);
            this.tab_header.Controls.Add(this.label2);
            this.tab_header.Controls.Add(this.cmbobx_header_colour);
            this.tab_header.Controls.Add(this.label11);
            this.tab_header.Controls.Add(this.label10);
            this.tab_header.Controls.Add(this.label9);
            this.tab_header.Controls.Add(this.cmbobx_header_style);
            this.tab_header.Controls.Add(this.cmbobx_header_fontsize);
            this.tab_header.Controls.Add(this.cmbobx_header_fontname);
            this.tab_header.Location = new System.Drawing.Point(4, 29);
            this.tab_header.Name = "tab_header";
            this.tab_header.Padding = new System.Windows.Forms.Padding(3);
            this.tab_header.Size = new System.Drawing.Size(1072, 697);
            this.tab_header.TabIndex = 0;
            this.tab_header.Text = "Header";
            this.tab_header.UseVisualStyleBackColor = true;
            // 
            // chkbx_header_bold
            // 
            this.chkbx_header_bold.AutoSize = true;
            this.chkbx_header_bold.Location = new System.Drawing.Point(59, 376);
            this.chkbx_header_bold.Name = "chkbx_header_bold";
            this.chkbx_header_bold.Size = new System.Drawing.Size(67, 24);
            this.chkbx_header_bold.TabIndex = 78;
            this.chkbx_header_bold.Text = "Bold";
            this.chkbx_header_bold.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(31, 252);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(153, 20);
            this.label3.TabIndex = 77;
            this.label3.Text = "Font Underline Style";
            // 
            // cmbobx_header_underline_style
            // 
            this.cmbobx_header_underline_style.FormattingEnabled = true;
            this.cmbobx_header_underline_style.Items.AddRange(new object[] {
            "Dash",
            "DashHeavy",
            "DashLong",
            "DashLongHeavy",
            "DotDash",
            "DotDashHeavy",
            "DotDotDash",
            "DotDotDashHeavy",
            "Dotted",
            "DottedHeavy",
            "Double",
            "None",
            "Single",
            "Thick",
            "Wavy",
            "WavyDouble",
            "WavyHeavy",
            "Words"});
            this.cmbobx_header_underline_style.Location = new System.Drawing.Point(192, 249);
            this.cmbobx_header_underline_style.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbobx_header_underline_style.Name = "cmbobx_header_underline_style";
            this.cmbobx_header_underline_style.Size = new System.Drawing.Size(178, 28);
            this.cmbobx_header_underline_style.TabIndex = 76;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 290);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 20);
            this.label2.TabIndex = 75;
            this.label2.Text = "Font Colour";
            // 
            // cmbobx_header_colour
            // 
            this.cmbobx_header_colour.FormattingEnabled = true;
            this.cmbobx_header_colour.Items.AddRange(new object[] {
            "Black",
            "Blue",
            "Bright Green",
            "Dark Blue",
            "Dark Red",
            "Dark Yellow",
            "Grey 25",
            "Grey 50",
            "Green",
            "Pink",
            "Red",
            "Teal",
            "Turquoise",
            "Violet",
            "White",
            "Yellow"});
            this.cmbobx_header_colour.Location = new System.Drawing.Point(192, 287);
            this.cmbobx_header_colour.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbobx_header_colour.Name = "cmbobx_header_colour";
            this.cmbobx_header_colour.Size = new System.Drawing.Size(178, 28);
            this.cmbobx_header_colour.TabIndex = 74;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(31, 214);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(81, 20);
            this.label11.TabIndex = 73;
            this.label11.Text = "Font Style";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(31, 176);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(77, 20);
            this.label10.TabIndex = 72;
            this.label10.Text = "Font Size";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(31, 138);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(91, 20);
            this.label9.TabIndex = 71;
            this.label9.Text = "Font Family";
            // 
            // cmbobx_header_style
            // 
            this.cmbobx_header_style.FormattingEnabled = true;
            this.cmbobx_header_style.Items.AddRange(new object[] {
            "Italic",
            "Bold",
            "Regular",
            "Strikeout",
            "Underline"});
            this.cmbobx_header_style.Location = new System.Drawing.Point(192, 211);
            this.cmbobx_header_style.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbobx_header_style.Name = "cmbobx_header_style";
            this.cmbobx_header_style.Size = new System.Drawing.Size(178, 28);
            this.cmbobx_header_style.TabIndex = 70;
            // 
            // cmbobx_header_fontsize
            // 
            this.cmbobx_header_fontsize.FormattingEnabled = true;
            this.cmbobx_header_fontsize.Items.AddRange(new object[] {
            "8",
            "9",
            "11",
            "12",
            "14",
            "16",
            "18",
            "20",
            "22",
            "24",
            "26",
            "28",
            "32",
            "48",
            "72"});
            this.cmbobx_header_fontsize.Location = new System.Drawing.Point(192, 173);
            this.cmbobx_header_fontsize.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbobx_header_fontsize.Name = "cmbobx_header_fontsize";
            this.cmbobx_header_fontsize.Size = new System.Drawing.Size(178, 28);
            this.cmbobx_header_fontsize.TabIndex = 69;
            // 
            // cmbobx_header_fontname
            // 
            this.cmbobx_header_fontname.FormattingEnabled = true;
            this.cmbobx_header_fontname.Location = new System.Drawing.Point(192, 135);
            this.cmbobx_header_fontname.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbobx_header_fontname.Name = "cmbobx_header_fontname";
            this.cmbobx_header_fontname.Size = new System.Drawing.Size(370, 28);
            this.cmbobx_header_fontname.TabIndex = 68;
            // 
            // tab_footer
            // 
            this.tab_footer.Location = new System.Drawing.Point(4, 29);
            this.tab_footer.Name = "tab_footer";
            this.tab_footer.Padding = new System.Windows.Forms.Padding(3);
            this.tab_footer.Size = new System.Drawing.Size(1072, 697);
            this.tab_footer.TabIndex = 1;
            this.tab_footer.Text = "Footer";
            this.tab_footer.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tableLayoutPanel2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 739);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1080, 54);
            this.panel2.TabIndex = 1;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 5;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 180F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 180F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 180F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 180F));
            this.tableLayoutPanel2.Controls.Add(this.panel4, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.panel5, 2, 0);
            this.tableLayoutPanel2.Controls.Add(this.panel6, 3, 0);
            this.tableLayoutPanel2.Controls.Add(this.panel7, 4, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(1080, 54);
            this.tableLayoutPanel2.TabIndex = 0;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.btn_create);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(363, 3);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(174, 48);
            this.panel4.TabIndex = 1;
            // 
            // btn_create
            // 
            this.btn_create.Location = new System.Drawing.Point(20, 6);
            this.btn_create.Name = "btn_create";
            this.btn_create.Size = new System.Drawing.Size(134, 37);
            this.btn_create.TabIndex = 1;
            this.btn_create.Text = "Create";
            this.btn_create.UseVisualStyleBackColor = true;
            this.btn_create.Click += new System.EventHandler(this.btn_create_Click);
            // 
            // panel5
            // 
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(543, 3);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(174, 48);
            this.panel5.TabIndex = 2;
            // 
            // panel6
            // 
            this.panel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel6.Location = new System.Drawing.Point(723, 3);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(174, 48);
            this.panel6.TabIndex = 3;
            // 
            // panel7
            // 
            this.panel7.Controls.Add(this.btn_close);
            this.panel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel7.Location = new System.Drawing.Point(903, 3);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(174, 48);
            this.panel7.TabIndex = 4;
            // 
            // btn_close
            // 
            this.btn_close.Location = new System.Drawing.Point(20, 6);
            this.btn_close.Name = "btn_close";
            this.btn_close.Size = new System.Drawing.Size(134, 37);
            this.btn_close.TabIndex = 0;
            this.btn_close.Text = "Close";
            this.btn_close.UseVisualStyleBackColor = true;
            this.btn_close.Click += new System.EventHandler(this.btn_close_Click);
            // 
            // chkbx_header_italic
            // 
            this.chkbx_header_italic.AutoSize = true;
            this.chkbx_header_italic.Location = new System.Drawing.Point(132, 376);
            this.chkbx_header_italic.Name = "chkbx_header_italic";
            this.chkbx_header_italic.Size = new System.Drawing.Size(76, 24);
            this.chkbx_header_italic.TabIndex = 79;
            this.chkbx_header_italic.Text = "Italics";
            this.chkbx_header_italic.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1086, 796);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Create a MS Word Report";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tab_defaults.ResumeLayout(false);
            this.tab_defaults.PerformLayout();
            this.tab_header.ResumeLayout(false);
            this.tab_header.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tab_header;
        private System.Windows.Forms.TabPage tab_footer;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Button btn_close;
        private System.Windows.Forms.Button btn_create;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cmbobx_header_style;
        private System.Windows.Forms.ComboBox cmbobx_header_fontsize;
        private System.Windows.Forms.ComboBox cmbobx_header_fontname;
        private System.Windows.Forms.TabPage tab_defaults;
        private System.Windows.Forms.CheckBox chkbx_page_footer;
        private System.Windows.Forms.CheckBox chkbx_page_header;
        private System.Windows.Forms.TextBox txtbx_proj_name;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbobx_header_colour;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbobx_header_underline_style;
        private System.Windows.Forms.CheckBox chkbx_header_bold;
        private System.Windows.Forms.CheckBox chkbx_header_italic;
    }
}

