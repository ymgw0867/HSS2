namespace HSS2.Config
{
    partial class frmComSet
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmComSet));
            this.label1 = new System.Windows.Forms.Label();
            this.txtID = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtComName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtMaru = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.chkKyujitsuShinya = new System.Windows.Forms.CheckBox();
            this.chkLunch = new System.Windows.Forms.CheckBox();
            this.chkKyujitsuTime = new System.Windows.Forms.CheckBox();
            this.chkShinyaTime = new System.Windows.Forms.CheckBox();
            this.chkZanTime = new System.Windows.Forms.CheckBox();
            this.chkShitsumuTime = new System.Windows.Forms.CheckBox();
            this.chkShukkinTime = new System.Windows.Forms.CheckBox();
            this.chkKekkinNissu = new System.Windows.Forms.CheckBox();
            this.chkYukyuNissu = new System.Windows.Forms.CheckBox();
            this.chkTokkyuNissu = new System.Windows.Forms.CheckBox();
            this.chkKyushutsuNIssu = new System.Windows.Forms.CheckBox();
            this.chkShukinNissu = new System.Windows.Forms.CheckBox();
            this.cmbShokushu = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.txtHayade = new System.Windows.Forms.MaskedTextBox();
            this.txtZangyo = new System.Windows.Forms.MaskedTextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 27);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "会社ID：";
            // 
            // txtID
            // 
            this.txtID.BackColor = System.Drawing.SystemColors.Window;
            this.txtID.Location = new System.Drawing.Point(94, 24);
            this.txtID.Name = "txtID";
            this.txtID.ReadOnly = true;
            this.txtID.Size = new System.Drawing.Size(31, 27);
            this.txtID.TabIndex = 1;
            this.txtID.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(155, 27);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "会社名：";
            // 
            // txtComName
            // 
            this.txtComName.BackColor = System.Drawing.SystemColors.Window;
            this.txtComName.Location = new System.Drawing.Point(212, 24);
            this.txtComName.Name = "txtComName";
            this.txtComName.ReadOnly = true;
            this.txtComName.Size = new System.Drawing.Size(288, 27);
            this.txtComName.TabIndex = 2;
            this.txtComName.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 63);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 20);
            this.label3.TabIndex = 4;
            this.label3.Text = "丸め単位：";
            // 
            // txtMaru
            // 
            this.txtMaru.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtMaru.Location = new System.Drawing.Point(94, 60);
            this.txtMaru.MaxLength = 2;
            this.txtMaru.Name = "txtMaru";
            this.txtMaru.Size = new System.Drawing.Size(31, 27);
            this.txtMaru.TabIndex = 3;
            this.txtMaru.Enter += new System.EventHandler(this.txtMaru_Enter);
            this.txtMaru.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMaru_KeyPress);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(125, 63);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(22, 20);
            this.label4.TabIndex = 7;
            this.label4.Text = "分";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(155, 63);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 20);
            this.label5.TabIndex = 8;
            this.label5.Text = "早出開始時間：";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(313, 63);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(100, 20);
            this.label6.TabIndex = 9;
            this.label6.Text = "所定終了時刻：";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.chkKyujitsuShinya);
            this.panel1.Controls.Add(this.chkLunch);
            this.panel1.Controls.Add(this.chkKyujitsuTime);
            this.panel1.Controls.Add(this.chkShinyaTime);
            this.panel1.Controls.Add(this.chkZanTime);
            this.panel1.Controls.Add(this.chkShitsumuTime);
            this.panel1.Controls.Add(this.chkShukkinTime);
            this.panel1.Controls.Add(this.chkKekkinNissu);
            this.panel1.Controls.Add(this.chkYukyuNissu);
            this.panel1.Controls.Add(this.chkTokkyuNissu);
            this.panel1.Controls.Add(this.chkKyushutsuNIssu);
            this.panel1.Controls.Add(this.chkShukinNissu);
            this.panel1.Controls.Add(this.cmbShokushu);
            this.panel1.Location = new System.Drawing.Point(26, 118);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(474, 290);
            this.panel1.TabIndex = 11;
            // 
            // chkKyujitsuShinya
            // 
            this.chkKyujitsuShinya.AutoSize = true;
            this.chkKyujitsuShinya.Location = new System.Drawing.Point(151, 244);
            this.chkKyujitsuShinya.Name = "chkKyujitsuShinya";
            this.chkKyujitsuShinya.Size = new System.Drawing.Size(106, 24);
            this.chkKyujitsuShinya.TabIndex = 12;
            this.chkKyujitsuShinya.Text = "休日深夜時間";
            this.chkKyujitsuShinya.UseVisualStyleBackColor = true;
            // 
            // chkLunch
            // 
            this.chkLunch.AutoSize = true;
            this.chkLunch.Location = new System.Drawing.Point(151, 214);
            this.chkLunch.Name = "chkLunch";
            this.chkLunch.Size = new System.Drawing.Size(80, 24);
            this.chkLunch.TabIndex = 11;
            this.chkLunch.Text = "昼食回数";
            this.chkLunch.UseVisualStyleBackColor = true;
            // 
            // chkKyujitsuTime
            // 
            this.chkKyujitsuTime.AutoSize = true;
            this.chkKyujitsuTime.Location = new System.Drawing.Point(151, 184);
            this.chkKyujitsuTime.Name = "chkKyujitsuTime";
            this.chkKyujitsuTime.Size = new System.Drawing.Size(80, 24);
            this.chkKyujitsuTime.TabIndex = 10;
            this.chkKyujitsuTime.Text = "休日時間";
            this.chkKyujitsuTime.UseVisualStyleBackColor = true;
            // 
            // chkShinyaTime
            // 
            this.chkShinyaTime.AutoSize = true;
            this.chkShinyaTime.Location = new System.Drawing.Point(151, 154);
            this.chkShinyaTime.Name = "chkShinyaTime";
            this.chkShinyaTime.Size = new System.Drawing.Size(80, 24);
            this.chkShinyaTime.TabIndex = 9;
            this.chkShinyaTime.Text = "深夜時間";
            this.chkShinyaTime.UseVisualStyleBackColor = true;
            // 
            // chkZanTime
            // 
            this.chkZanTime.AutoSize = true;
            this.chkZanTime.Location = new System.Drawing.Point(151, 124);
            this.chkZanTime.Name = "chkZanTime";
            this.chkZanTime.Size = new System.Drawing.Size(106, 24);
            this.chkZanTime.TabIndex = 8;
            this.chkZanTime.Text = "早出残業時間";
            this.chkZanTime.UseVisualStyleBackColor = true;
            // 
            // chkShitsumuTime
            // 
            this.chkShitsumuTime.AutoSize = true;
            this.chkShitsumuTime.Location = new System.Drawing.Point(151, 94);
            this.chkShitsumuTime.Name = "chkShitsumuTime";
            this.chkShitsumuTime.Size = new System.Drawing.Size(80, 24);
            this.chkShitsumuTime.TabIndex = 7;
            this.chkShitsumuTime.Text = "執務時間";
            this.chkShitsumuTime.UseVisualStyleBackColor = true;
            // 
            // chkShukkinTime
            // 
            this.chkShukkinTime.AutoSize = true;
            this.chkShukkinTime.Location = new System.Drawing.Point(152, 64);
            this.chkShukkinTime.Name = "chkShukkinTime";
            this.chkShukkinTime.Size = new System.Drawing.Size(80, 24);
            this.chkShukkinTime.TabIndex = 6;
            this.chkShukkinTime.Text = "出勤時間";
            this.chkShukkinTime.UseVisualStyleBackColor = true;
            // 
            // chkKekkinNissu
            // 
            this.chkKekkinNissu.AutoSize = true;
            this.chkKekkinNissu.Location = new System.Drawing.Point(19, 184);
            this.chkKekkinNissu.Name = "chkKekkinNissu";
            this.chkKekkinNissu.Size = new System.Drawing.Size(80, 24);
            this.chkKekkinNissu.TabIndex = 5;
            this.chkKekkinNissu.Text = "欠勤日数";
            this.chkKekkinNissu.UseVisualStyleBackColor = true;
            // 
            // chkYukyuNissu
            // 
            this.chkYukyuNissu.AutoSize = true;
            this.chkYukyuNissu.Location = new System.Drawing.Point(19, 154);
            this.chkYukyuNissu.Name = "chkYukyuNissu";
            this.chkYukyuNissu.Size = new System.Drawing.Size(80, 24);
            this.chkYukyuNissu.TabIndex = 4;
            this.chkYukyuNissu.Text = "有休日数";
            this.chkYukyuNissu.UseVisualStyleBackColor = true;
            // 
            // chkTokkyuNissu
            // 
            this.chkTokkyuNissu.AutoSize = true;
            this.chkTokkyuNissu.Location = new System.Drawing.Point(19, 124);
            this.chkTokkyuNissu.Name = "chkTokkyuNissu";
            this.chkTokkyuNissu.Size = new System.Drawing.Size(80, 24);
            this.chkTokkyuNissu.TabIndex = 3;
            this.chkTokkyuNissu.Text = "特休日数";
            this.chkTokkyuNissu.UseVisualStyleBackColor = true;
            // 
            // chkKyushutsuNIssu
            // 
            this.chkKyushutsuNIssu.AutoSize = true;
            this.chkKyushutsuNIssu.Location = new System.Drawing.Point(19, 94);
            this.chkKyushutsuNIssu.Name = "chkKyushutsuNIssu";
            this.chkKyushutsuNIssu.Size = new System.Drawing.Size(80, 24);
            this.chkKyushutsuNIssu.TabIndex = 2;
            this.chkKyushutsuNIssu.Text = "休出日数";
            this.chkKyushutsuNIssu.UseVisualStyleBackColor = true;
            // 
            // chkShukinNissu
            // 
            this.chkShukinNissu.AutoSize = true;
            this.chkShukinNissu.Location = new System.Drawing.Point(19, 64);
            this.chkShukinNissu.Name = "chkShukinNissu";
            this.chkShukinNissu.Size = new System.Drawing.Size(80, 24);
            this.chkShukinNissu.TabIndex = 1;
            this.chkShukinNissu.Text = "出勤日数";
            this.chkShukinNissu.UseVisualStyleBackColor = true;
            // 
            // cmbShokushu
            // 
            this.cmbShokushu.FormattingEnabled = true;
            this.cmbShokushu.Location = new System.Drawing.Point(19, 18);
            this.cmbShokushu.Name = "cmbShokushu";
            this.cmbShokushu.Size = new System.Drawing.Size(212, 28);
            this.cmbShokushu.TabIndex = 0;
            this.cmbShokushu.SelectedIndexChanged += new System.EventHandler(this.cmbShokushu_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(42, 107);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(61, 20);
            this.label7.TabIndex = 12;
            this.label7.Text = "出力設定";
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(194, 424);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(138, 35);
            this.btnOk.TabIndex = 6;
            this.btnOk.Text = "OK(&O)";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.Location = new System.Drawing.Point(407, 424);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(93, 35);
            this.btnRtn.TabIndex = 7;
            this.btnRtn.Text = "終了(&Q)";
            this.btnRtn.UseVisualStyleBackColor = true;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // txtHayade
            // 
            this.txtHayade.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtHayade.Location = new System.Drawing.Point(252, 60);
            this.txtHayade.Mask = "90:00";
            this.txtHayade.Name = "txtHayade";
            this.txtHayade.Size = new System.Drawing.Size(54, 27);
            this.txtHayade.TabIndex = 4;
            this.txtHayade.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtHayade.ValidatingType = typeof(System.DateTime);
            this.txtHayade.Enter += new System.EventHandler(this.txtHayade_Enter);
            this.txtHayade.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMaru_KeyPress);
            // 
            // txtZangyo
            // 
            this.txtZangyo.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtZangyo.Location = new System.Drawing.Point(409, 60);
            this.txtZangyo.Mask = "90:00";
            this.txtZangyo.Name = "txtZangyo";
            this.txtZangyo.Size = new System.Drawing.Size(54, 27);
            this.txtZangyo.TabIndex = 5;
            this.txtZangyo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtZangyo.ValidatingType = typeof(System.DateTime);
            this.txtZangyo.Enter += new System.EventHandler(this.txtMaru_Enter);
            this.txtZangyo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMaru_KeyPress);
            // 
            // frmComSet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(527, 471);
            this.Controls.Add(this.txtZangyo);
            this.Controls.Add(this.txtHayade);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtMaru);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtComName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtID);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label5);
            this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.Name = "frmComSet";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "会社別設定";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmComSet_FormClosing);
            this.Load += new System.EventHandler(this.frmComSet_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtID;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtComName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtMaru;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox chkKyujitsuShinya;
        private System.Windows.Forms.CheckBox chkLunch;
        private System.Windows.Forms.CheckBox chkKyujitsuTime;
        private System.Windows.Forms.CheckBox chkShinyaTime;
        private System.Windows.Forms.CheckBox chkZanTime;
        private System.Windows.Forms.CheckBox chkShitsumuTime;
        private System.Windows.Forms.CheckBox chkShukkinTime;
        private System.Windows.Forms.CheckBox chkKekkinNissu;
        private System.Windows.Forms.CheckBox chkYukyuNissu;
        private System.Windows.Forms.CheckBox chkTokkyuNissu;
        private System.Windows.Forms.CheckBox chkKyushutsuNIssu;
        private System.Windows.Forms.CheckBox chkShukinNissu;
        private System.Windows.Forms.ComboBox cmbShokushu;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.MaskedTextBox txtHayade;
        private System.Windows.Forms.MaskedTextBox txtZangyo;
    }
}