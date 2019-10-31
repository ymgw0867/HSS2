using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using HSS2.Model;

namespace HSS2.ScanOcr
{
    public partial class frmStaffSelect : Form
    {
        public frmStaffSelect(Entity.ShainMst [] m)
        {
            InitializeComponent();
            Mst = m;
            _msCode = string.Empty;
        }

        Entity.ShainMst[] Mst;

        private void frmStaffSelect_Load(object sender, EventArgs e)
        {
            Model.Utility.WindowsMaxSize(this, this.Width, this.Height);
            Model.Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビュー定義
            GridViewSetting(dataGridView1);         // スタッフ定義
            GridViewStaffShow(dataGridView1);       // スタッフマスター表示
        }

        // データグリッドビューカラム名
        private string cCode = "c1";
        private string cName = "c2";
        private string cKana = "c3";
        private string cShoID = "c4";
        private string cShoName = "c5";
        private string cTenban = "c6";
        private string cTenName = "c7";
        private string cKikan = "c8";
        private string cShoTime = "c9";
        private string cTokuTime = "c10";
        private string cTcode = "c11";

        /// <summary>
        /// グリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewSetting(DataGridView tempDGV)
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
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 380;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                tempDGV.Columns.Add(cCode, "個人番号");
                tempDGV.Columns.Add(cName, "氏名");
                tempDGV.Columns.Add(cKana, "カナ");
                tempDGV.Columns.Add(cShoID, "職種ID");
                tempDGV.Columns.Add(cShoName, "職種名称");
                tempDGV.Columns.Add(cTenban, "店番");
                tempDGV.Columns.Add(cTenName, "勤務場所名称");
                tempDGV.Columns.Add(cKikan, "契約期間");
                tempDGV.Columns.Add(cShoTime, "所定勤務時間");
                tempDGV.Columns.Add(cTokuTime, "特殊日勤務時間");
                tempDGV.Columns.Add(cTcode, "コード");
                tempDGV.Columns[cTcode].Visible = false;

                tempDGV.Columns[cCode].Width = 80;
                tempDGV.Columns[cName].Width = 120;
                tempDGV.Columns[cKana].Width = 120;
                tempDGV.Columns[cShoID].Width = 100;
                tempDGV.Columns[cShoName].Width = 160;
                tempDGV.Columns[cTenban].Width = 80;
                tempDGV.Columns[cTenName].Width = 160;
                tempDGV.Columns[cKikan].Width = 200;
                tempDGV.Columns[cShoTime].Width = 200;
                tempDGV.Columns[cTokuTime].Width = 200;

                tempDGV.Columns[cCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cShoID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cTenban].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKikan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cShoTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cTokuTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                
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

                // 列サイズ変更不可
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                tempDGV.StandardTab = false;

                // ソート禁止
                foreach (DataGridViewColumn c in tempDGV.Columns)
                {
                    c.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                //tempDGV.Columns[cDay].SortMode = DataGridViewColumnSortMode.NotSortable;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// スタッフマスタをグリッドビューへ表示します
        /// </summary>
        /// <param name="tempGrid">データグリッドビューオブジェクト</param>
        private void GridViewStaffShow(DataGridView tempGrid)
        {
            int iX = 0;
            tempGrid.RowCount = 0;

            for (int i = 0; i < Mst.Length; i++)
            {
                tempGrid.Rows.Add();

                tempGrid[cCode, iX].Value = Mst[i].sCode;
                tempGrid[cName, iX].Value = Mst[i].sName;
                tempGrid[cKana, iX].Value = Mst[i].sKana;
                tempGrid[cShoID, iX].Value = Mst[i].sSID;
                tempGrid[cShoName, iX].Value = Mst[i].sSName;
                tempGrid[cTenban, iX].Value = Mst[i].sTenpoCode;
                tempGrid[cTenName, iX].Value = Mst[i].sTenpoName;

                if (Mst[i].sKeiyakuS == string.Empty && Mst[i].sKeiyakuE == string.Empty)
                    tempGrid[cKikan, iX].Value = string.Empty;
                else tempGrid[cKikan, iX].Value = Mst[i].sKeiyakuS + " ～ " + Mst[i].sKeiyakuE;

                if (Mst[i].sWorkTimeS == string.Empty && Mst[i].sWorkTimeE == string.Empty)
                    tempGrid[cShoTime, iX].Value = string.Empty;
                else tempGrid[cShoTime, iX].Value = Mst[i].sWorkTimeS + " ～ " + Mst[i].sWorkTimeE;

                if (Mst[i].sWorkTimeS2 == string.Empty && Mst[i].sWorkTimeE2 == string.Empty)
                    tempGrid[cTokuTime, iX].Value = string.Empty;
                else tempGrid[cTokuTime, iX].Value = Mst[i].sWorkTimeS2 + " ～ " + Mst[i].sWorkTimeE2;

                iX++;
            }

            tempGrid.CurrentCell = null;  
        }

        private void GetGridViewData(DataGridView g)
        {
            if (g.SelectedRows.Count == 0) return;

            int r = g.SelectedRows[0].Index;

            _msCode = g[cCode, r].Value.ToString();

            //_msName = g[cName, r].Value.ToString();
            //_msTenpoCode = g[cKinmuCode, r].Value.ToString();
            //_msTenpoName = g[cKinmuName, r].Value.ToString();
            //if (_usrSel == global.STAFF_SELECT)
            //{
            //    _msBusho = g[cKinmuBusho, r].Value.ToString();
            //    _msTcode = g[cTcode, r].Value.ToString();
            //    _msOrderCode = g[cKOrderCode, r].Value.ToString();
            //}

            //_msStime = g[cKKikan, r].Value.ToString().Substring(0, 5);
            //_msEtime = g[cKKikan, r].Value.ToString().Substring(6, 5);
                   
            //TimeS = Replace(msStime, ":", "");
            //TimeE = Replace(msEtime, ":", "");
            //TimeJ = (((fncCMin(TimeE) - fncCMin(TimeS)) - 60) / 2);
    
            //_msHStime1 = msStime;
            //_msHEtime1 = Format(DateAdd("n", TimeJ, msStime), "Short Time");
            //_msHStime2 = Format(DateAdd("n", TimeJ + 60, msStime), "Short Time");
            //_msHEtime2 = msEtime;
        }

        public string _msCode { get; set; }
        public string _msName { get; set; }
        public string _msTenpoCode { get; set; }
        public string _msTenpoName { get; set; }
        public string _msBusho { get; set; }
        public string _msTcode { get; set; }
        public string _msOrderCode { get; set; }
        public string _msStime { get; set; }
        public string _msEtime { get; set; }
        public string _msHStime1 { get; set; }
        public string _msHEtime1 { get; set; }
        public string _msHStime2 { get; set; }
        public string _msHEtime2 { get; set; }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GetGridViewData(dataGridView1);
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            GetGridViewData(dataGridView1);
            this.Close();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmStaffSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            //this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int r = -1;
            if (txtsCode.Text != string.Empty) 
                r = GridViewFind(dataGridView1, cCode, txtsCode.Text);
            else if (txtsName.Text != string.Empty) 
                r = GridViewFind(dataGridView1, cName, txtsName.Text);

            if (r != -1)
            {
                dataGridView1.FirstDisplayedScrollingRowIndex = r;
                dataGridView1.CurrentCell = dataGridView1[cCode, r];
            }
        }

        /// <summary>
        /// コードまたは氏名で該当するデータの行を返します
        /// </summary>
        /// <param name="d">データグリッドビューオブジェクト</param>
        /// <param name="ColName">対象のカラム名</param>
        /// <param name="Val">検索する値</param>
        /// <returns>行のIndex</returns>
        private int GridViewFind(DataGridView d, string ColName, string Val)
        {
            int rtn = -1;

            for (int i = 0; i < d.Rows.Count; i++)
            {
                if (d[ColName, i].Value.ToString().Contains(Val))
                {
                    rtn = i;
                    break;
                }
            }

            return rtn;
        }

        private void txtsCode_TextChanged(object sender, EventArgs e)
        {
            if (txtsCode.Text.Length > 0) txtsName.Text = string.Empty;
        }

        private void txtsName_TextChanged(object sender, EventArgs e)
        {
            if (txtsName.Text.Length > 0) txtsCode.Text = string.Empty;
        }
    }
}
