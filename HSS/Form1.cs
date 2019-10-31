using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HSS2.Config;
using HSS2.Model;
using HSS2.ScanOcr;

namespace HSS2
{
    public partial class Form1 : Form
    {
        string sMstPath;
        string pMstPath;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 解放する
            this.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // フォーム最大値、最小値を設定する
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // ディレクトリ作成
            dirCreate(Properties.Settings.Default.instPath + Properties.Settings.Default.NG);
            dirCreate(Properties.Settings.Default.instPath + Properties.Settings.Default.OK);
            dirCreate(Properties.Settings.Default.instPath + Properties.Settings.Default.READ);

            //// 設定設定情報を取得します
            global.GetCommonYearMonth();
        }

        /// <summary>
        /// 該当パスが存在しなければ作成します
        /// </summary>
        /// <sPath>
        /// 該当パス
        /// </sPath>
        private void dirCreate(string sPath)
        {
            if (System.IO.Directory.Exists(sPath) == false)
                System.IO.Directory.CreateDirectory(sPath);
        }

        private void btnSetup_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmConfig frm = new frmConfig();
            frm.ShowDialog();
            this.Show();
        }

        private void btnPrePrint_Click(object sender, EventArgs e)
        {
            this.Hide();

            // 会社領域を選択します
            Model.frmComSelect frm = new frmComSelect();
            frm.ShowDialog();
            if (global.pblComNo == global.NON_SELECT)
            {
                this.Show();
                return;
            }

            // プレ印刷画面表示
            PrePrint.frmPrePrint frmPre = new PrePrint.frmPrePrint(global.pblDbName);
            frmPre.ShowDialog();
            this.Show();
        }

        private void btnOCR_Click(object sender, EventArgs e)
        {
            this.Hide();

            // 会社領域を選択します
            Model.frmComSelect frmCom = new frmComSelect();
            frmCom.ShowDialog();
            if (global.pblComNo == global.NON_SELECT)
            {
                this.Show();
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("スキャナで画像を読み取りOCR認識処理を行います。").Append(Environment.NewLine).Append(Environment.NewLine);
            sb.Append("よろしいですか？中止する場合は「いいえ」をクリックしてください。");
            if ((MessageBox.Show(sb.ToString(), this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No))
            {
                this.Show();
                return;
            }
            else
            {
                // 勤務票スキャン、ＯＣＲ処理
                frmOCR frm = new frmOCR();
                frm.ShowDialog();
                this.Show();
            }
        }

        private void btnDataEntry_Click(object sender, EventArgs e)
        {
            // 環境設定年月の確認
            string msg = "設定年月は " + global.sYear.ToString() + "年 " + global.sMonth.ToString() + "月です。よろしいですか？";
            if (MessageBox.Show(msg, "勤務データ登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;
            
            this.Hide();

            // 会社領域を選択します
            Model.frmComSelect frm = new frmComSelect();
            frm.ShowDialog();
            if (global.pblComNo == global.NON_SELECT)
            {
                this.Show();
                return;
            }

            // 勤怠データ登録画面表示
            ScanOcr.frmCorrect frmData = new frmCorrect(global.pblComNo, global.pblComName, global.pblDbName);
            frmData.ShowDialog();
            this.Show();
        }

        private void btnDataCombi_Click(object sender, EventArgs e)
        {
            this.Hide();

            // 会社領域を選択します
            Model.frmComSelect frm = new frmComSelect();
            frm.ShowDialog();
            if (global.pblComNo == global.NON_SELECT)
            {
                this.Show();
                return;
            }

            frmComSet frmc = new frmComSet(global.pblComNo, global.pblComName, global.pblDbName);
            frmc.ShowDialog();
            this.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();

            // 会社領域を選択します
            Model.frmComSelect frm = new frmComSelect();
            frm.ShowDialog();
            if (global.pblComNo == global.NON_SELECT)
            {
                this.Show();
                return;
            }

            frmCalendar frmc = new frmCalendar(global.pblComNo, global.pblComName);
            frmc.ShowDialog();
            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //string a = "123";
            //string b = a.Substring(4, 1);
            //MessageBox.Show(b);

            string n = ":";
            DateTime d = DateTime.Parse(n);
        }

     }
}
