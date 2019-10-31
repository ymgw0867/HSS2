using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace HSS2.Model
{
    public partial class frmComSelect : Form
    {
        public frmComSelect()
        {
            InitializeComponent();
        }

        // 選択した会社ID
        Int64 _ComNum = 0;

        private void frmComSelect_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最大サイズ
            Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //DataGridViewの設定
            GridViewSetting(dg1);

            // 接続文字列取得 2016/10/12
            string sc = SqlControl.obcConnectSting.get(Properties.Settings.Default.sqlCurrentDB);     

            //データ表示
            GridViewShowData(sc,dg1);

            //終了時タグ初期化
            Tag = string.Empty;

        }
        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                tempDGV.Height = 187;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "会社ＩＤ");
                tempDGV.Columns.Add("col2", "会社名");
                tempDGV.Columns.Add("col3", "ＤＢ名称");

                tempDGV.Columns[0].Width = 100;
                tempDGV.Columns[1].Width = 200;
                tempDGV.Columns[2].Width = 100;

                tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ会社情報を表示する </summary>
        /// <param name="sConnect">
        ///     接続文字列</param>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        ///--------------------------------------------------------------------
        private void GridViewShowData(string sConnect, DataGridView tempDGV)
        {
            string sqlSTRING = string.Empty;

            SqlControl.DataControl sdcon = new SqlControl.DataControl(sConnect);

            //データリーダーを取得する
            SqlDataReader dR;

            sqlSTRING += "SELECT EntityCode,EntityName,DatabaseName from tbCorpDatabaseContext order by EntityCode";
            dR = sdcon.free_dsReader(sqlSTRING);

            try
            {
                //グリッドビューに表示する
                int iX = 0;
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    //データグリッドにデータを表示する
                    tempDGV.Rows.Add();
                    
                    // 会社ＩＤ
                    tempDGV[0, iX].Value = dR["EntityCode"].ToString().PadLeft(2,'0');

                    // 会社名
                    tempDGV[1, iX].Value = dR["EntityName"].ToString().Trim();

                    // データベース名
                    tempDGV[2, iX].Value = dR["DatabaseName"].ToString().Trim();

                    iX++;
                }
                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (sdcon.Cn.State == ConnectionState.Open) sdcon.Close();
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            //会社情報がないときはそのままクローズ
            if (dg1.RowCount == 0)
            {
                global.pblComNo = 0;                //会社№
                global.pblDbName = string.Empty;    //データベース名
                global.pblComName = string.Empty;
            }
            else
            {
                if (dg1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("会社を選択してください", "会社未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                //選択した会社情報を取得する
                int r = dg1.SelectedRows[0].Index;

                _ComNum = Int64.Parse(dg1[0, r].Value.ToString());    // 会社№
                global.pblComNo = _ComNum;                          // 会社№
                global.pblComName = dg1[1, r].Value.ToString();     // 会社名
                global.pblDbName = dg1[2, r].Value.ToString();      // データベース名
            }

            //フォームを閉じる
            Tag = "btn";
            this.Close();
        }

        private void frmComSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 会社選択のときディレクトリを作成する
            if (_ComNum != 0) CreateDir(_ComNum);

            this.Dispose();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            global.pblComNo = global.NON_SELECT;
            this.Close();
        }

        private void dg1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //選択した会社情報を取得する
            int r = dg1.SelectedRows[0].Index;

            _ComNum = Int64.Parse(dg1[0, r].Value.ToString());    // 会社№
            global.pblComNo = _ComNum;                          // 会社№
            global.pblComName = dg1[1, r].Value.ToString();     // 会社名
            global.pblDbName = dg1[2, r].Value.ToString();      // データベース名

            this.Close();
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     会社別のディレクトリを作成します </summary>
        /// <param name="comNum">
        ///     会社ID</param>
        ///-----------------------------------------------------------------
        private void CreateDir(Int64 comNum)
        {
            // 未選択のときは終了
            if (_ComNum == 0) return;

            // パスを定義します
            string pathData = Properties.Settings.Default.instPath + @"DATA\" + global.pblComNo.ToString().PadLeft(2, '0') + " " + global.pblComName;
            string pathTif = global.sTIF + global.pblComNo.ToString().PadLeft(2, '0') + " " + global.pblComName;
            string pathDat = global.sDAT + global.pblComNo.ToString().PadLeft(2, '0') + " " + global.pblComName;

            // DATAディレクトリが存在しなければ作成します
            if (!System.IO.Directory.Exists(pathData))
                System.IO.Directory.CreateDirectory(pathData);

            // 画像保存ディレクトリが存在しなければ作成します
            if (!System.IO.Directory.Exists(pathTif))
                System.IO.Directory.CreateDirectory(pathTif);

            // 受け渡しデータディレクトリが存在しなければ作成します
            if (!System.IO.Directory.Exists(pathDat))
                System.IO.Directory.CreateDirectory(pathDat);
        }
    }
}
