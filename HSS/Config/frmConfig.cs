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
using Leadtools.Twain;

namespace HSS2.Config
{
    public partial class frmConfig : Form
    {
        const int ADDNEW = 0;
        const int EDIT = 1;
        const string SEIREKI = "20";
        int fMode = ADDNEW;

        // TWAINでの取得用にTwainSessionを宣言します。
        TwainSession _twainSession = new TwainSession();

        public frmConfig()
        {
            InitializeComponent();
        }

        private void frmConfig_Load(object sender, EventArgs e)
        {
            // フォームの最大サイズ、最小サイズの設定
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // バックアップ自動削除月数コンボボックス値設定
            Utility.comboBkDel.Load(this.cmbBkdels);

            // TwainSessionを初期化します
            _twainSession.Startup(this, "FKDL", "LEADTOOLS", "ver16.5J", "ocr", TwainStartupFlags.None);
            
            // データ表示
            DataShow();
        }

        private void DataShow()
        {
            // 表示項目クリア
            DispClear();

            // データベースの接続定義
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dr;

            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            OleDbConnection Cn = sDB.cnOpen();
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            sb.Append("select * from 環境設定");
            sCom.CommandText = sb.ToString();
            sCom.Connection = Cn;
            dr = sCom.ExecuteReader();

            try
            {
                while (dr.Read())
                {
                    fMode = EDIT;
                    if (dr["SYEAR"].ToString().Length == 4)
                        txtYear.Text = dr["SYEAR"].ToString().Substring(2, 2);
                    else txtYear.Text = dr["SYEAR"].ToString();
                    txtMonth.Text = dr["SMONTH"].ToString();
                    //txtScanner.Text = dr["SCAN"].ToString();
                    txtDat.Text = dr["DAT"].ToString();
                    txtTif.Text = dr["TIF"].ToString();

                    cmbBkdels.SelectedIndex = Utility.comboBkDel.selectedIndex(cmbBkdels, int.Parse(dr["BKDELS"].ToString()));

                    fMode = 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (dr.IsClosed == false) dr.Close();
                if (Cn.State == ConnectionState.Open) Cn.Close();
            }
        }

        private void DispClear()
        {
            txtYear.Text = string.Empty;
            txtMonth.Text = string.Empty;
            txtScanner.Text = string.Empty;
            txtTif.Text = string.Empty;
            txtDat.Text = string.Empty;
            cmbBkdels.Text = string.Empty;

            if (TwainSession.IsAvailable(this))
            {
                btnScanner.Enabled = true;
                txtScanner.Enabled = true;
                txtScanner.Text = _twainSession.SelectedSourceName();
            }
            else
            {
                btnScanner.Enabled = false;
                txtScanner.Enabled = false;
            }
        }

        /// <summary>
        /// フォルダダイアログ選択
        /// </summary>
        /// <returns>フォルダー名</returns>
        private string userFolderSelect()
        {
            string fName = string.Empty;

            //出力フォルダの選択ダイアログの表示
            // FolderBrowserDialog の新しいインスタンスを生成する (デザイナから追加している場合は必要ない)
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            // ダイアログの説明を設定する
            folderBrowserDialog1.Description = "フォルダを選択してください";

            // ルートになる特殊フォルダを設定する (初期値 SpecialFolder.Desktop)
            folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.Desktop;

            // 初期選択するパスを設定する
            folderBrowserDialog1.SelectedPath = @"C:\";

            // [新しいフォルダ] ボタンを表示する (初期値 true)
            folderBrowserDialog1.ShowNewFolderButton = true;

            // ダイアログを表示し、戻り値が [OK] の場合は、選択したディレクトリを表示する
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                fName = folderBrowserDialog1.SelectedPath + @"\";
            }
            else
            {
                // 不要になった時点で破棄する
                folderBrowserDialog1.Dispose();
                return fName;
            }

            // 不要になった時点で破棄する
            folderBrowserDialog1.Dispose();

            return fName;
        }

        private void btnDat_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            txtDat.Text = userFolderSelect();
        }

        private void btnTif_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            txtTif.Text = userFolderSelect();
        }

        private string getOpenFileName()
        {
            string f = string.Empty;

            // ダイアログボックスの表示
            openFileDialog1.Title = "ファイルの選択";
            openFileDialog1.Filter = "データファイル(*.csv *.txt *.dat)|*.csv;*.txt;*.dat";
            openFileDialog1.InitialDirectory = Properties.Settings.Default.instPath;
            openFileDialog1.FileName = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK) f = openFileDialog1.FileName;
            return f;
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();
            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            if (txtObj.Text == string.Empty) return;
            if (txtObj.Text.Length < 2) txtObj.Text = string.Format("{0:D2}", int.Parse(txtObj.Text)); 
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();
            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;
            if (sender == txtScanner) txtObj = txtScanner;
            if (sender == txtDat) txtObj = txtDat;
            if (sender == txtTif) txtObj = txtTif;

            txtObj.SelectAll();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (errCheck())
            {
                switch (fMode)
                {
                    case ADDNEW:
                        AddNewRecord();
                        break;

                    case EDIT:
                        EditRecord();
                        break;
                }
            }

            // 設定対象年月を取得します
            global.GetCommonYearMonth();

            // 終了します
            this.Close();
        }

        /// <summary>
        /// 環境設定データ新規登録処理
        /// </summary>
        private void AddNewRecord()
        {
            OleDbCommand sCom = new OleDbCommand();
            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            sCom.Connection = sDB.cnOpen();
            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Clear();
                sb.Append("insert into 環境設定 (");
                sb.Append("TIF,DAT,BKDELS,SCAN,SYEAR,SMONTH,更新年月日) values (");
                sb.Append("?,?,?,?,?,?,?)");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@tif", txtTif.Text);
                sCom.Parameters.AddWithValue("@dat", txtDat.Text);
                sCom.Parameters.AddWithValue("@bkdels", int.Parse(cmbBkdels.Text));
                sCom.Parameters.AddWithValue("@scan", txtScanner.Text);
                sCom.Parameters.AddWithValue("@year", SEIREKI + txtYear.Text);
                sCom.Parameters.AddWithValue("@Month", txtMonth.Text);
                sCom.Parameters.AddWithValue("@update", DateTime.Today.ToShortDateString());
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// 環境設定データ書き換え
        /// </summary>
        private void EditRecord()
        {
            OleDbCommand sCom = new OleDbCommand();
            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            sCom.Connection = sDB.cnOpen();
            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Clear();
                sb.Append("update 環境設定 set ");
                sb.Append("TIF=?,DAT=?,BKDELS=?,SCAN=?,SYEAR=?,SMONTH=?,更新年月日=?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@tif", txtTif.Text);
                sCom.Parameters.AddWithValue("@dat", txtDat.Text);
                sCom.Parameters.AddWithValue("@bkdels", int.Parse(cmbBkdels.Text));
                sCom.Parameters.AddWithValue("@scan", txtScanner.Text);
                sCom.Parameters.AddWithValue("@year", SEIREKI + txtYear.Text);
                sCom.Parameters.AddWithValue("@Month", txtMonth.Text);
                sCom.Parameters.AddWithValue("@update", DateTime.Today.ToShortDateString());
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private Boolean errCheck()
        {
            if (txtYear.Text == string.Empty)
            {
                MessageBox.Show("対象年が未入力です", "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (txtMonth.Text == string.Empty)
            {
                MessageBox.Show("対象月が未入力です", "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }

            int sMonth = int.Parse(txtMonth.Text);
            if (sMonth < 1 || sMonth > 12)
            {
                MessageBox.Show("対象月が正しくありません", "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }

            if (cmbBkdels.SelectedIndex == -1)
            {
                MessageBox.Show("スタッフのバックアップデータの自動削除設定を選択してください", "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbBkdels.Focus();
                return false;
            }
            return true;
        }

        private void frmConfig_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnScanner_Click(object sender, EventArgs e)
        {
            try
            {
                _twainSession.SelectSource(string.Empty);
                txtScanner.Text = _twainSession.SelectedSourceName();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "スキャナの選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
    }
}
