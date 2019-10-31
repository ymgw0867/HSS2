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
using System.Globalization;
using HSS2.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace HSS2.PrePrint
{
    public partial class frmPrePrint : Form
    {
        string _dbName = string.Empty;

        Entity.ShainMst [] mst = new Entity.ShainMst[1];
        Entity.ShokushuMst [] sMst = new Entity.ShokushuMst[1];

        public frmPrePrint(string dbName)
        {
            InitializeComponent();
            _dbName = dbName;
        }

        private void frmPrePrint_Load(object sender, EventArgs e)
        {
            // フォーム最小値を設定します
            Model.Utility.WindowsMinSize(this, this.Width, this.Height);

            // マスターロード
            MstLoad(_dbName);

            // グリッド定義
            GridviewSet.Setting(dataGridView1);

            // グリッドビューデータ表示
            //GridviewSet.Show(dataGridView1, mst, string.Empty);

            // 職種コンボ値をロード
            cmbShokushuLoad();

            // 選択ボタン初期化
            btnSelAll.Tag = "OFF";

            // 対象年月
            if (global.sYear.ToString().Length == 4)
                txtYear.Text = global.sYear.ToString().Substring(2, 2);
            else txtYear.Text = DateTime.Today.Year.ToString().Substring(2, 2);

            txtMonth.Text = global.sMonth.ToString("00");
        }

        // データグリッドビュークラス
        private class GridviewSet
        {
            public const string col_ID = "col1";
            public const string col_Name = "col2";
            public const string col_Kana = "col3";
            public const string col_ShokuID = "col4";
            public const string col_ShokuName = "col5";
            public const string col_Tenban = "col6";
            public const string col_KinmuBashoName = "col7";
            public const string col_KeiyakuKikan = "col8";
            public const string col_Shoteikinmu = "col9";
            public const string col_TokushubiKinmu = "col10";

            /// <summary>
            /// データグリッドビューの定義を行います
            /// </summary>
            /// <param name="tempDGV">データグリッドビューオブジェクト</param>
            public static void Setting(DataGridView tempDGV)
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
                    tempDGV.ColumnHeadersHeight = 20;
                    tempDGV.RowTemplate.Height = 20;

                    // 全体の高さ
                    tempDGV.Height = 462;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                    DataGridViewCheckBoxColumn cl = new DataGridViewCheckBoxColumn();

                    //各列幅指定
                    tempDGV.Columns.Add(col_ID, "個人番号");
                    tempDGV.Columns.Add(col_Name, "氏名");
                    tempDGV.Columns.Add(col_Kana, "カナ");
                    tempDGV.Columns.Add(col_ShokuID, "職種ID");
                    tempDGV.Columns.Add(col_ShokuName, "職種名称");
                    tempDGV.Columns.Add(col_Tenban, "店番");
                    tempDGV.Columns.Add(col_KinmuBashoName, "勤務場所名称");
                    tempDGV.Columns.Add(col_KeiyakuKikan, "契約期間");
                    tempDGV.Columns.Add(col_Shoteikinmu, "所定勤務時間");
                    tempDGV.Columns.Add(col_TokushubiKinmu, "特殊日勤務時間");

                    tempDGV.Columns[col_ID].Width = 90;
                    tempDGV.Columns[col_Name].Width = 120;
                    tempDGV.Columns[col_Kana].Width = 100;
                    tempDGV.Columns[col_ShokuID].Width = 90;
                    tempDGV.Columns[col_ShokuName].Width = 120;
                    tempDGV.Columns[col_Tenban].Width = 80;
                    tempDGV.Columns[col_KinmuBashoName].Width = 220;
                    tempDGV.Columns[col_KeiyakuKikan].Width = 220;  // 2019/02/06
                    tempDGV.Columns[col_Shoteikinmu].Width = 120;
                    tempDGV.Columns[col_TokushubiKinmu].Width = 120;

                    tempDGV.Columns[col_ID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[col_ShokuID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[col_Tenban].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[col_Shoteikinmu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[col_TokushubiKinmu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

                    // 編集の可否設定
                    tempDGV.ReadOnly = true;

                    // 追加行表示しない
                    tempDGV.AllowUserToAddRows = false;

                    // データグリッドビューから行削除を禁止する
                    tempDGV.AllowUserToDeleteRows = false;

                    // 手動による列移動の禁止
                    tempDGV.AllowUserToOrderColumns = false;

                    // 列サイズ変更可
                    tempDGV.AllowUserToResizeColumns = true;

                    // 行サイズ変更禁止
                    tempDGV.AllowUserToResizeRows = false;

                    // 行ヘッダーの自動調節
                    //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                    //TAB動作
                    tempDGV.StandardTab = true;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            ///-----------------------------------------------------------------------
            /// <summary>
            ///     データグリッドビューにデータを表示します </summary>
            /// <param name="tempGrid">
            ///     データグリッドビューオブジェクト</param>
            ///-----------------------------------------------------------------------
            public static void Show(DataGridView tempGrid, Entity.ShainMst[] mst, string Sho)
            {
                tempGrid.RowCount = 0;
                int iX = 0;
                for (int i = 0; i < mst.Length; i++)
                {
                    // 「全ての職種」または指定した職種IDと一致するデータを印刷対象とします
                    if (Sho == string.Empty || Sho == mst[i].sSID)
                    {
                        tempGrid.Rows.Add();
                        tempGrid[col_ID, iX].Value = mst[i].sCode;
                        tempGrid[col_Name, iX].Value = mst[i].sName;
                        tempGrid[col_Kana, iX].Value = mst[i].sKana;
                        tempGrid[col_ShokuID, iX].Value = mst[i].sSID;
                        tempGrid[col_ShokuName, iX].Value = mst[i].sSName;
                        tempGrid[col_Tenban, iX].Value = mst[i].sTenpoCode;
                        tempGrid[col_KinmuBashoName, iX].Value = mst[i].sTenpoName;

                        if (mst[i].sKeiyakuS == string.Empty && mst[i].sKeiyakuE == string.Empty)
                            tempGrid[col_KeiyakuKikan, iX].Value = string.Empty;
                        else tempGrid[col_KeiyakuKikan, iX].Value = mst[i].sKeiyakuS + " ～ " + mst[i].sKeiyakuE;

                        if (mst[i].sWorkTimeS != string.Empty || mst[i].sWorkTimeE != string.Empty)
                            tempGrid[col_Shoteikinmu, iX].Value = mst[i].sWorkTimeS + " ～ " + mst[i].sWorkTimeE;
                        else tempGrid[col_Shoteikinmu, iX].Value = string.Empty;

                        if (mst[i].sWorkTimeS2 != string.Empty || mst[i].sWorkTimeE2 != string.Empty)
                            tempGrid[col_TokushubiKinmu, iX].Value = mst[i].sWorkTimeS2 + " ～ " + mst[i].sWorkTimeE2;
                        else tempGrid[col_TokushubiKinmu, iX].Value = string.Empty;
                        
                        iX++;
                    }
                }

                tempGrid.CurrentCell = null;
            }
        }

        private void btnSelAll_Click(object sender, EventArgs e)
        {
            switch (btnSelAll.Tag.ToString())
            {
                case "OFF":
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        dataGridView1.Rows[i].Selected = true;
                    }
                    btnSelAll.Tag = "ON";
                    btnSelAll.Text = "全て非選択とする(&N)";
                    break;

                case "ON":
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        dataGridView1.Rows[i].Selected = false;
                    }
                    btnSelAll.Tag = "OFF";
                    btnSelAll.Text = "全て選択する(&A)";
                    break;

                default:
                    break;
            }
        }

        private void frmPrePrint_Shown(object sender, EventArgs e)
        {
            dataGridView1.CurrentCell = null;
        }

        private void txtSCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\r')
            {
                e.Handled = true;
            }

            if (e.KeyChar == '\r')
            {
                if (txtSCode.Text == string.Empty) return;

                dataGridView1.CurrentCell = null;

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1[GridviewSet.col_ID, i].Value.ToString().Contains(txtSCode.Text))
                    {
                        dataGridView1.CurrentCell = dataGridView1[GridviewSet.col_ID, i];
                        break;
                    }
                }
            }
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            // 処理月を検証します
            if (!Utility.NumericCheck(txtMonth.Text))
            {
                MessageBox.Show("処理月が正しくありません","勤務票プレ印刷",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
            {
                MessageBox.Show("正しい処理月を入力してください","勤務票プレ印刷",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            // 印刷対象データ件数を調べます
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("印刷対象データがありません", "勤務票プレ印刷", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string xlsPath = Properties.Settings.Default.instPath + Properties.Settings.Default.XLS + Properties.Settings.Default.ExcelFile;
            if (!System.IO.File.Exists(xlsPath))
            {
                MessageBox.Show("Excelファイルが見つかりません", "勤務票プレ印刷", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("勤務記録用紙の印刷を開始します。よろしいですか？", "勤務票プレ印刷（指定件数：" + dataGridView1.SelectedRows.Count.ToString() + "）"
                , MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                return;

            // 勤務票印刷
            sReportStaff(Properties.Settings.Default.instPath +
             Properties.Settings.Default.XLS +
             Properties.Settings.Default.ExcelFile,
             dataGridView1);
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// スタッフ勤務票印刷
        /// </summary>
        /// <param name="xlsPath">勤務票エクセルシートパス</param>
        /// <param name="dg1">データグリッドビューオブジェクト</param>
        /// <param name="rBtn">オプション番号</param>
        private void sReportStaff(string xlsPath, DataGridView dg1)
        {
            string sID = string.Empty;
            int r = 0;

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;
                Excel.Application oXls = new Excel.Application();
                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(xlsPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                //Excel.Worksheet oxlsSheet = new Excel.Worksheet();

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                // データベース接続
                SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
                OleDbCommand sCom = new OleDbCommand();
                sCom.Connection = sDB.cnOpen();
                OleDbDataReader dr = null;

                try
                {
                    //グリッドを順番に読む
                    for (int i = 0; i < dg1.Rows.Count; i++)
                    {
                        if (dg1.Rows[i].Selected)
                        {
                            ////印刷2件目以降はシートを追加する
                            //pCnt++;

                            //if (pCnt > 1)
                            //{
                            //    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                            //    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];
                            //}

                            // シートを初期化します
                            oxlsSheet.Cells[3, 3] = string.Empty;
                            oxlsSheet.Cells[3, 4] = string.Empty;
                            oxlsSheet.Cells[3, 8] = string.Empty;
                            oxlsSheet.Cells[3, 9] = string.Empty;
                            oxlsSheet.Cells[3, 10] = string.Empty;
                            oxlsSheet.Cells[3, 11] = string.Empty;
                            oxlsSheet.Cells[3, 12] = string.Empty;
                            oxlsSheet.Cells[3, 13] = string.Empty;
                            oxlsSheet.Cells[3, 14] = string.Empty;
                            oxlsSheet.Cells[3, 19] = string.Empty;
                            oxlsSheet.Cells[3, 29] = string.Empty;
                            oxlsSheet.Cells[3, 30] = string.Empty;
                            oxlsSheet.Cells[3, 31] = string.Empty;
                            oxlsSheet.Cells[3, 32] = string.Empty;
                            oxlsSheet.Cells[3, 34] = string.Empty;
                            oxlsSheet.Cells[3, 35] = string.Empty;
                            oxlsSheet.Cells[4, 3] = string.Empty;
                            oxlsSheet.Cells[4, 4] = string.Empty;
                            oxlsSheet.Cells[4, 8] = string.Empty;
                            oxlsSheet.Cells[4, 29] = string.Empty;
                            oxlsSheet.Cells[5, 5] = string.Empty;
                            oxlsSheet.Cells[5, 6] = string.Empty;
                            oxlsSheet.Cells[5, 7] = string.Empty;
                            oxlsSheet.Cells[5, 13] = string.Empty;
                            oxlsSheet.Cells[5, 29] = string.Empty;

                            for (int ix = 8; ix <= 38; ix++)
                            {
                                oxlsSheet.Cells[ix, 1] = string.Empty;
                                oxlsSheet.Cells[ix, 2] = string.Empty;
                                rng[0] = (Excel.Range)oxlsSheet.Cells[ix, 1];
                                rng[1] = (Excel.Range)oxlsSheet.Cells[ix, 2];
                                oxlsSheet.get_Range(rng[0], rng[1]).Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                            }

                            // 会社ID
                            oxlsSheet.Cells[3, 3] = Utility.toWide(global.pblComNo.ToString().PadLeft(2, '0').Substring(0, 1));
                            oxlsSheet.Cells[3, 4] = Utility.toWide(global.pblComNo.ToString().PadLeft(2, '0').Substring(1, 1));

                            // 個人番号
                            oxlsSheet.Cells[3, 8] = Utility.toWide(dg1[GridviewSet.col_ID, i].Value.ToString().Substring(0, 1));
                            oxlsSheet.Cells[3, 9] = Utility.toWide(dg1[GridviewSet.col_ID, i].Value.ToString().Substring(1, 1));
                            oxlsSheet.Cells[3, 10] = Utility.toWide(dg1[GridviewSet.col_ID, i].Value.ToString().Substring(2, 1));
                            oxlsSheet.Cells[3, 11] = Utility.toWide(dg1[GridviewSet.col_ID, i].Value.ToString().Substring(3, 1));
                            oxlsSheet.Cells[3, 12] = Utility.toWide(dg1[GridviewSet.col_ID, i].Value.ToString().Substring(4, 1));
                            oxlsSheet.Cells[3, 13] = Utility.toWide(dg1[GridviewSet.col_ID, i].Value.ToString().Substring(5, 1));
                            oxlsSheet.Cells[3, 14] = Utility.toWide(dg1[GridviewSet.col_ID, i].Value.ToString().Substring(6, 1));

                            // 氏名
                            oxlsSheet.Cells[3, 19] = dg1[GridviewSet.col_Name, i].Value.ToString();

                            // 年
                            oxlsSheet.Cells[3, 29] = "２";
                            oxlsSheet.Cells[3, 30] = "０";
                            oxlsSheet.Cells[3, 31] = Utility.toWide(this.txtYear.Text.Substring(0, 1));
                            oxlsSheet.Cells[3, 32] = Utility.toWide(this.txtYear.Text.Substring(1, 1));

                            // 月
                            oxlsSheet.Cells[3, 34] = Utility.toWide(this.txtMonth.Text.Substring(0, 1));
                            oxlsSheet.Cells[3, 35] = Utility.toWide(this.txtMonth.Text.Substring(1, 1));

                            // 契約期間
                            oxlsSheet.Cells[4, 8] = dg1[GridviewSet.col_KeiyakuKikan, i].Value.ToString();

                            // 通常勤務時間
                            oxlsSheet.Cells[4, 29] = dg1[GridviewSet.col_Shoteikinmu, i].Value.ToString();

                            // 勤務場所コード
                            oxlsSheet.Cells[5, 5] = Utility.toWide(dg1[GridviewSet.col_Tenban, i].Value.ToString().Substring(0, 1));
                            oxlsSheet.Cells[5, 6] = Utility.toWide(dg1[GridviewSet.col_Tenban, i].Value.ToString().Substring(1, 1));
                            oxlsSheet.Cells[5, 7] = Utility.toWide(dg1[GridviewSet.col_Tenban, i].Value.ToString().Substring(2, 1));

                            // 勤務場所名称
                            oxlsSheet.Cells[5, 13] = dg1[GridviewSet.col_KinmuBashoName, i].Value.ToString();

                            // 特殊勤務時間
                            oxlsSheet.Cells[5, 29] = dg1[GridviewSet.col_TokushubiKinmu, i].Value.ToString();

                            // 日・曜日
                            DateTime dt;
                            string sDate = "20" + txtYear.Text + "/" + txtMonth.Text + "/";
                            for (int ix = 8; ix <= 38; ix++)
                            {
                                int days = ix - 7;
                                if (DateTime.TryParse(sDate + days.ToString(), out dt))
                                {
                                    string youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                                    oxlsSheet.Cells[ix, 1] = days.ToString();
                                    oxlsSheet.Cells[ix, 2] = youbi;
                                    rng[0] = (Excel.Range)oxlsSheet.Cells[ix, 1];

                                    // 土日の場合
                                    if (youbi == "日" || youbi == "土")
                                    {
                                        oxlsSheet.get_Range(rng[0], rng[0]).Interior.ColorIndex = 15;
                                    }
                                    else
                                    {
                                        // 休日テーブルを参照し休日に該当するか調べます
                                        sCom.CommandText = "select * from 休日 where 年=? and 月=? and 日=? and 会社ID=?";
                                        sCom.Parameters.Clear();
                                        sCom.Parameters.AddWithValue("@year", int.Parse("20" + txtYear.Text));
                                        sCom.Parameters.AddWithValue("@Month", int.Parse(txtMonth.Text));
                                        sCom.Parameters.AddWithValue("@day", days);
                                        sCom.Parameters.AddWithValue("@id", global.pblComNo.ToString("00"));
                                        dr = sCom.ExecuteReader();
                                        if (dr.HasRows)
                                        {
                                            oxlsSheet.get_Range(rng[0], rng[0]).Interior.ColorIndex = 15;
                                        }
                                        dr.Close();
                                    }
                                }
                            }

                            // 会社名称
                            oxlsSheet.Cells[42, 2] = global.pblComName;

                            // ウィンドウを非表示にする
                            //oXls.Visible = false;

                            // 印刷
                            //oxlsSheet.PrintPreview(false);
                            oxlsSheet.PrintOut(1, Type.Missing, 1, false, oXls.ActivePrinter, Type.Missing, Type.Missing, Type.Missing);
                            //oXlsBook.PrintOut();
                        }
                    }

                    // マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 終了メッセージ
                    MessageBox.Show("印刷が終了しました", "勤務票印刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                finally
                {
                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();

                    // COMオブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    // マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // データリーダーを閉じる
                    if (dr.IsClosed == false) dr.Close();

                    // データベースを切断
                    if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        /// <summary>
        /// パートタイマー勤務票印刷
        /// </summary>
        /// <param name="xlsPath">勤務票エクセルシートパス</param>
        /// <param name="dg1">データグリッドビューオブジェクト</param>
        //private void sReportPart(string xlsPath, DataGridView dg1)
        //{
        //    string sID = string.Empty;

        //    try
        //    {
        //        //マウスポインタを待機にする
        //        this.Cursor = Cursors.WaitCursor;
        //        Excel.Application oXls = new Excel.Application();
        //        Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(xlsPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing));

        //        Excel.Worksheet oxlsSheet = new Excel.Worksheet();

        //        // スタッフＩＤによるシートの判断
        //        oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets["ID1"];
        //        sID = "1";

        //        Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

        //        // データベース接続
        //        SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
        //        OleDbCommand sCom = new OleDbCommand();
        //        sCom.Connection = sDB.cnOpen();
        //        OleDbDataReader dr = null;

        //        try
        //        {
        //            //グリッドを順番に読む
        //            for (int i = 0; i < dg1.RowCount; i++)
        //            {
        //                //チェックがあるものを対象とする
        //                if (dg1[GridviewSet.col_Check, i].Value.ToString() == "True")
        //                {
        //                    ////印刷2件目以降はシートを追加する
        //                    //pCnt++;

        //                    //if (pCnt > 1)
        //                    //{
        //                    //    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
        //                    //    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];
        //                    //}

        //                    // シートを初期化します
        //                    oxlsSheet.Cells[3, 2] = string.Empty;
        //                    oxlsSheet.Cells[3, 5] = string.Empty;
        //                    oxlsSheet.Cells[3, 6] = string.Empty;
        //                    oxlsSheet.Cells[3, 7] = string.Empty;
        //                    oxlsSheet.Cells[3, 8] = string.Empty;
        //                    oxlsSheet.Cells[3, 9] = string.Empty;
        //                    oxlsSheet.Cells[3, 10] = string.Empty;
        //                    oxlsSheet.Cells[3, 11] = string.Empty;
        //                    oxlsSheet.Cells[3, 14] = string.Empty;
        //                    oxlsSheet.Cells[3, 26] = string.Empty;
        //                    oxlsSheet.Cells[3, 27] = string.Empty;
        //                    oxlsSheet.Cells[3, 28] = string.Empty;
        //                    oxlsSheet.Cells[3, 29] = string.Empty;
        //                    oxlsSheet.Cells[3, 31] = string.Empty;
        //                    oxlsSheet.Cells[3, 32] = string.Empty;
        //                    oxlsSheet.Cells[4, 4] = string.Empty;
        //                    oxlsSheet.Cells[4, 26] = string.Empty;
        //                    oxlsSheet.Cells[5, 4] = string.Empty;
        //                    oxlsSheet.Cells[5, 5] = string.Empty;
        //                    oxlsSheet.Cells[5, 6] = string.Empty;
        //                    oxlsSheet.Cells[5, 11] = string.Empty;

        //                    for (int ix = 8; ix <= 38; ix++)
        //                    {
        //                        oxlsSheet.Cells[ix, 1] = string.Empty;
        //                        oxlsSheet.Cells[ix, 2] = string.Empty;
        //                        rng[0] = (Excel.Range)oxlsSheet.Cells[ix, 1];
        //                        rng[1] = (Excel.Range)oxlsSheet.Cells[ix, 2];
        //                        oxlsSheet.get_Range(rng[0], rng[1]).Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
        //                    }

        //                    int sRow = i;

        //                    // ID
        //                    oxlsSheet.Cells[3, 2] = sID;

        //                    // 個人番号
        //                    oxlsSheet.Cells[3, 5] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(0, 1);
        //                    oxlsSheet.Cells[3, 6] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(1, 1);
        //                    oxlsSheet.Cells[3, 7] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(2, 1);
        //                    oxlsSheet.Cells[3, 8] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(3, 1);
        //                    oxlsSheet.Cells[3, 9] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(4, 1);
        //                    oxlsSheet.Cells[3, 10] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(5, 1);
        //                    oxlsSheet.Cells[3, 11] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(6, 1);

        //                    // 氏名
        //                    oxlsSheet.Cells[3, 14] = dg1[GridviewSet.col_Name, i].Value.ToString();

        //                    // 年
        //                    oxlsSheet.Cells[3, 26] = "２";
        //                    oxlsSheet.Cells[3, 27] = "０";
        //                    oxlsSheet.Cells[3, 28] = this.txtYear.Text.Substring(0, 1);
        //                    oxlsSheet.Cells[3, 29] = this.txtYear.Text.Substring(1, 1);

        //                    // 月
        //                    oxlsSheet.Cells[3, 31] = this.txtMonth.Text.Substring(0, 1);
        //                    oxlsSheet.Cells[3, 32] = this.txtMonth.Text.Substring(1, 1);

        //                    // 契約期間
        //                    oxlsSheet.Cells[4, 4] = dg1[GridviewSet.col_Keiyaku, i].Value.ToString();

        //                    // 就業時間
        //                    oxlsSheet.Cells[4, 26] = dg1[GridviewSet.col_WorkTime, i].Value.ToString();

        //                    // 勤務場所店番
        //                    oxlsSheet.Cells[5, 4] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(0, 1);
        //                    oxlsSheet.Cells[5, 5] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(1, 1);
        //                    oxlsSheet.Cells[5, 6] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(2, 1);

        //                    // 勤務場所名称
        //                    oxlsSheet.Cells[5, 11] = dg1[GridviewSet.col_sName1, i].Value.ToString();

        //                    // 日・曜日
        //                    DateTime dt;
        //                    string sDate = "20" + txtYear.Text + "/" + txtMonth.Text + "/";
        //                    for (int ix = 8; ix <= 38; ix++)
        //                    {
        //                        int days = ix - 7;
        //                        if (DateTime.TryParse(sDate + days.ToString(), out dt))
        //                        {
        //                            string youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
        //                            oxlsSheet.Cells[ix, 1] = days.ToString();
        //                            oxlsSheet.Cells[ix, 2] = youbi;
        //                            rng[0] = (Excel.Range)oxlsSheet.Cells[ix, 1];

        //                            // 土日の場合
        //                            if (youbi == "日" || youbi == "土")
        //                            {
        //                                oxlsSheet.get_Range(rng[0], rng[0]).Interior.ColorIndex = 15;
        //                            }
        //                            else
        //                            {
        //                                // 休日テーブルを参照し休日に該当するか調べます
        //                                sCom.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
        //                                sCom.Parameters.Clear();
        //                                sCom.Parameters.AddWithValue("@year", int.Parse("20" + txtYear.Text));
        //                                sCom.Parameters.AddWithValue("@Month", int.Parse(txtMonth.Text));
        //                                sCom.Parameters.AddWithValue("@day", days);
        //                                dr = sCom.ExecuteReader();
        //                                if (dr.HasRows)
        //                                {
        //                                    oxlsSheet.get_Range(rng[0], rng[0]).Interior.ColorIndex = 15;
        //                                }
        //                                dr.Close();
        //                            }
        //                        }
        //                    }
        //                    // ウィンドウを非表示にする
        //                    //oXls.Visible = false;
        //                    //印刷
        //                    //oxlsSheet.PrintPreview(false);
        //                    oxlsSheet.PrintOut(1, Type.Missing, 1, false, oXls.ActivePrinter, Type.Missing, Type.Missing, Type.Missing);
        //                    //oXlsBook.PrintOut();
        //                }
        //            }

        //            // マウスポインタを元に戻す
        //            this.Cursor = Cursors.Default;
        //            // 終了メッセージ
        //            MessageBox.Show("印刷が終了しました", "勤務票印刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        }

        //        catch (Exception e)
        //        {
        //            MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //        }

        //        finally
        //        {
        //            // ウィンドウを非表示にする
        //            oXls.Visible = false;

        //            // 保存処理
        //            oXls.DisplayAlerts = false;

        //            // Bookをクローズ
        //            oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

        //            // Excelを終了
        //            oXls.Quit();

        //            // COMオブジェクトの参照カウントを解放する 
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

        //            // マウスポインタを元に戻す
        //            this.Cursor = Cursors.Default;

        //            // データリーダーを閉じる
        //            if (dr.IsClosed == false) dr.Close();

        //            // データベースを切断
        //            if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //    }

        //    // マウスポインタを元に戻す
        //    this.Cursor = Cursors.Default;
        //}

        private void txtMonth_Leave(object sender, EventArgs e)
        {
            txtMonth.Text = txtMonth.Text.PadLeft(2, '0');
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            txtYear.Text = txtYear.Text.PadLeft(2, '0');
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmPrePrint_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void MstLoad(string dbName)
        {
            // 勘定奉行データベース接続文字列を取得する 2016/10/12
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //string sc = Utility.GetDBConnect(dbName);

            string sqlSTRING = string.Empty;

            SqlControl.DataControl sdcon = new SqlControl.DataControl(sc);

            //データリーダーを取得する
            SqlDataReader dr;

            // ------------------------------------------
            //   会社マスタからデータを取得
            //
            //       2010/10/08
            //       給与奉行iシリーズから取得
            // ------------------------------------------

            string mySQL = string.Empty;

            mySQL += "SELECT tbEmployeeBase.EmployeeID, tbSalarySystem.SalarySystemCode, tbHR_DivisionCategory.CategoryCode as zaisekikbn, ";
            mySQL += "tbSalarySystem.SalarySystemName, tbEmployeeBase.EmployeeNo, tbEmployeeBase.NameKana, tbEmployeeBase.Name, ";
            mySQL += "tbDepartment.DepartmentID, tbDepartment.DepartmentCode, tbDepartment.DepartmentName, ";
            mySQL += "tbEmployeeLaborAgreement.ContractStartDate, tbEmployeeLaborAgreement.ContractEndDate , ";
            mySQL += "tbEmployeeLaborAgreement.WeeklyContractTimeSlot1Start, tbEmployeeLaborAgreement.WeeklyContractTimeSlot1End, ";
            mySQL += "tbEmployeeLaborAgreement.WeeklyContractTimeSlot2Start, tbEmployeeLaborAgreement.WeeklyContractTimeSlot2End, ";
            mySQL += "tbEmployeeBase.RetireCorpScheduleDate ";
            mySQL += "from (((((tbEmployeeBase inner join "; 
            
            // 2012/05/15 給与体系データは基準日内最新データとする
            mySQL += "(select tbEmployeeUnitPrice.EmployeeID,tbEmployeeUnitPrice.SalarySystemID from tbEmployeeUnitPrice inner join ";
            mySQL += "(select EmployeeID,MAX(RevisionDate) as RevisionDate from tbEmployeeUnitPrice ";
            mySQL += "where RevisionDate <= '" + global.sYear + "/" + global.sMonth + "/01' ";
            mySQL += "group by EmployeeID) as e ";
            mySQL += "on tbEmployeeUnitPrice.EmployeeID = e.EmployeeID and tbEmployeeUnitPrice.RevisionDate = e.RevisionDate) as tbd ";
            
            mySQL += "on tbEmployeeBase.EmployeeID = tbd.EmployeeID) ";
            mySQL += "inner join tbSalarySystem on tbd.SalarySystemID = tbSalarySystem.SalarySystemID) ";
            mySQL += "inner join tbEmployeeLaborAgreement on tbEmployeeBase.EmployeeID = tbEmployeeLaborAgreement.EmployeeID) ";
            mySQL += "inner join (";

            mySQL += "select tbEmployeeMainDutyPersonnelChange.EmployeeID,tbEmployeeMainDutyPersonnelChange.AnnounceDate,";
            mySQL += "tbEmployeeMainDutyPersonnelChange.BelongID from tbEmployeeMainDutyPersonnelChange inner join (";
            mySQL += "select EmployeeID,max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ";
            mySQL += "where AnnounceDate <= '" + global.sYear + "/" + global.sMonth + "/01' ";
            mySQL += "group by EmployeeID) as a ";
            mySQL += "on (tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and ";
            mySQL += "(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate) ";
            mySQL += ") as d ";
            mySQL += "on tbEmployeeBase.EmployeeID = d.EmployeeID) ";

            mySQL += "inner join tbDepartment on d.BelongID = tbDepartment.DepartmentID) ";
            mySQL += "inner join tbHR_DivisionCategory on tbEmployeeBase.BeOnTheRegisterDivisionID = tbHR_DivisionCategory.CategoryID ";
            mySQL += "ORDER BY DepartmentCode,tbEmployeeBase.EmployeeNo,tbEmployeeLaborAgreement.ContractStartDate ";

            dr = sdcon.free_dsReader(mySQL);

            // マスタ読み込み
            int iX = 0;
            string yy;
            string mm;
            string dd;
            string sReBuf = string.Empty;
            string dt = string.Empty;
            string date1 = string.Empty;
            string date2 = string.Empty;

            while (dr.Read())
            {
                //--------------------------------------------------
                //  給与奉行iシリーズ対応　2010/10/11
                //  和暦を西暦表示に変える 2019/02/06
                //--------------------------------------------------
                if (int.Parse(dr["zaisekikbn"].ToString()) != 2)     // iシリーズでは"2"が退職
                {
                    // 退職予定日を取得する
                    yy = DateTime.Parse(dr["RetireCorpScheduleDate"].ToString()).Year.ToString();
                    mm = DateTime.Parse(dr["RetireCorpScheduleDate"].ToString()).Month.ToString().PadLeft(2, '0');
                    dd = DateTime.Parse(dr["RetireCorpScheduleDate"].ToString()).Day.ToString().PadLeft(2, '0');

                    //string rsDate = Utility.SeitoWa(yy).PadLeft(2, '0') + mm + dd;
                    string rsDate = yy + mm + dd;// 退職予定日・西暦 2019/02/06

                    // 退職予定日が環境設定年月の月初日より前のとき対象外とする
                    // 西暦対応 2019/02/06
                    if (int.Parse(rsDate) >= int.Parse(global.sYear.ToString() + global.sMonth.ToString().PadLeft(2, '0') + "01"))
                    {
                        if (sReBuf != dr["EmployeeID"].ToString().Trim())
                        {
                            // 契約開始日が環境設定年月の月初日以前を対象とする 2010/10/13
                            // 西暦年月日を返す 2019/02/06
                            dt = Utility.fcKikan2(dr["ContractStartDate"].ToString());

                            if (dt != string.Empty)
                            {
                                date1 = dt.Substring(0, 4) + dt.Substring(5, 2) + dt.Substring(8, 2);
                            }
                            else
                            {
                                date1 = "0";
                            }

                            date2 = global.sYear.ToString() + global.sMonth.ToString("00") + "01";

                            if (int.Parse(date1) <= int.Parse(date2))
                            {
                                //2件目以降なら要素数を追加
                                if (iX != 0) mst.CopyTo(mst = new Entity.ShainMst[iX + 1], 0);

                                mst[iX].sCode = Utility.subStringRight(dr["EmployeeNo"].ToString(), 7);
                                mst[iX].sName = dr["Name"].ToString().Trim();
                                mst[iX].sKana = dr["NameKana"].ToString().Trim();
                                mst[iX].sTenpoCode = Utility.subStringRight(dr["DepartmentCode"].ToString(), 3);
                                mst[iX].sTenpoName = dr["DepartmentName"].ToString().Trim();
                                mst[iX].sKeiyakuS = Utility.fcKikan2(dr["ContractStartDate"].ToString());   // 西暦対応　2019/02/06
                                mst[iX].sKeiyakuE = Utility.fcKikan2(dr["ContractEndDate"].ToString().Trim());  // 西暦対応　2019/02/06
                                mst[iX].sWorkTimeS = Utility.fcTime(dr["WeeklyContractTimeSlot1Start"].ToString());
                                mst[iX].sWorkTimeE = Utility.fcTime(dr["WeeklyContractTimeSlot1End"].ToString());
                                mst[iX].sWorkTimeS2 = Utility.fcTime(dr["WeeklyContractTimeSlot2Start"].ToString());
                                mst[iX].sWorkTimeE2 = Utility.fcTime(dr["WeeklyContractTimeSlot2End"].ToString());
                                mst[iX].sSID = Utility.subStringRight(dr["SalarySystemCode"].ToString(), 2).PadLeft(2, '0');
                                mst[iX].sSName = dr["SalarySystemName"].ToString().Trim();
                                iX++;

                                sReBuf = dr["EmployeeID"].ToString().Trim();
                            }
                        }
                        else
                        {
                            if (mst[iX - 1].sKeiyakuS != string.Empty)
                            {
                                date1 = mst[iX - 1].sKeiyakuS.Substring(0, 4) + mst[iX - 1].sKeiyakuS.Substring(5, 2) + mst[iX - 1].sKeiyakuS.Substring(8, 2);
                            }
                            else
                            {
                                date1 = "0";
                            }

                            // 西暦対応 2019/02/06
                            dt = Utility.fcKikan2(dr["ContractStartDate"].ToString());

                            // 2019/02/06
                            if (dt != string.Empty)
                            {
                                date2 = dt.Substring(0, 4) + dt.Substring(5, 2) + dt.Substring(8, 2);
                            }
                            else
                            {
                                date2 = "0";
                            }

                            if (int.Parse(date1) < int.Parse(date2))
                            {
                                // 2010/10/13  契約開始日が環境設定年月の月初日以前を対象とする
                                // 2019/02/06 西暦対応
                                dt = Utility.fcKikan2(dr["ContractStartDate"].ToString());

                                // 西暦対応　2019/02/06
                                if (dt != string.Empty)
                                {
                                    date1 = dt.Substring(0, 4) + dt.Substring(5, 2) + dt.Substring(8, 2);
                                }
                                else
                                {
                                    date1 = "0";
                                }

                                if (int.Parse(date1) <= int.Parse(global.sYear.ToString() + global.sMonth.ToString("00") + "01"))
                                {
                                    mst[iX-1].sCode = Utility.subStringRight(dr["EmployeeNo"].ToString(), 7);
                                    mst[iX-1].sName = dr["Name"].ToString().Trim();
                                    mst[iX-1].sKana = dr["NameKana"].ToString().Trim();
                                    mst[iX-1].sTenpoCode = Utility.subStringRight(dr["DepartmentCode"].ToString(), 3);
                                    mst[iX-1].sTenpoName = dr["DepartmentName"].ToString().Trim();
                                    mst[iX - 1].sKeiyakuS = Utility.fcKikan2(dr["ContractStartDate"].ToString());     // 西暦対応　2019/02/06
                                    mst[iX - 1].sKeiyakuE = Utility.fcKikan2(dr["ContractEndDate"].ToString().Trim());    // 西暦対応　2019/02/06
                                    mst[iX-1].sWorkTimeS = Utility.fcTime(dr["WeeklyContractTimeSlot1Start"].ToString());
                                    mst[iX-1].sWorkTimeE = Utility.fcTime(dr["WeeklyContractTimeSlot1End"].ToString());
                                    mst[iX-1].sWorkTimeS2 = Utility.fcTime(dr["WeeklyContractTimeSlot2Start"].ToString());
                                    mst[iX-1].sWorkTimeE2 = Utility.fcTime(dr["WeeklyContractTimeSlot2End"].ToString());
                                    mst[iX-1].sSID = Utility.subStringRight(dr["SalarySystemCode"].ToString(), 2).PadLeft(2, '0');
                                    mst[iX-1].sSName = dr["SalarySystemName"].ToString().Trim();
                                }                    
                            }
                        }
                    }
                }

                // 2019/02/06 和暦 → 西暦のため コメント化
                ////--------------------------------------------------
                ////   給与奉行iシリーズ対応　2010/10/11
                ////--------------------------------------------------
                //if (int.Parse(dr["zaisekikbn"].ToString()) != 2)     // iシリーズでは"2"が退職
                //{
                //    // 退職予定日を取得する
                //    yy = DateTime.Parse(dr["RetireCorpScheduleDate"].ToString()).Year.ToString();
                //    mm = DateTime.Parse(dr["RetireCorpScheduleDate"].ToString()).Month.ToString().PadLeft(2, '0');
                //    dd = DateTime.Parse(dr["RetireCorpScheduleDate"].ToString()).Day.ToString().PadLeft(2, '0');

                //    string rsDate = Utility.SeitoWa(yy).PadLeft(2, '0') + mm + dd;

                //    // 退職予定日が環境設定年月の月初日より前のとき対象外とする 2010/10/13
                //    if (int.Parse(rsDate) >= int.Parse(Utility.SeitoWa(global.sYear.ToString()) + global.sMonth.ToString().PadLeft(2, '0') + "01"))
                //    {
                //        if (sReBuf != dr["EmployeeID"].ToString().Trim())
                //        {
                //            // 契約開始日が環境設定年月の月初日以前を対象とする 2010/10/13
                //            dt = Utility.fcKikan(dr["ContractStartDate"].ToString());

                //            // 昭和に対応　2012/08/26
                //            if (dt != string.Empty)
                //            {
                //                if (dt.Substring(0, 1) == "H")
                //                    date1 = dt.Substring(1, 2) + dt.Substring(4, 2) + dt.Substring(7, 2);
                //                else date1 = "00" + dt.Substring(4, 2) + dt.Substring(7, 2); // 昭和のときは一時的に年は0にする
                //            }
                //            else date1 = "0";

                //            date2 = Utility.SeitoWa(global.sYear.ToString()) + global.sMonth.ToString("00") + "01";

                //            if (int.Parse(date1) <= int.Parse(date2))
                //            {
                //                //2件目以降なら要素数を追加
                //                if (iX != 0) mst.CopyTo(mst = new Entity.ShainMst[iX + 1], 0);

                //                mst[iX].sCode = Utility.subStringRight(dr["EmployeeNo"].ToString(), 7);
                //                mst[iX].sName = dr["Name"].ToString().Trim();
                //                mst[iX].sKana = dr["NameKana"].ToString().Trim();
                //                mst[iX].sTenpoCode = Utility.subStringRight(dr["DepartmentCode"].ToString(), 3);
                //                mst[iX].sTenpoName = dr["DepartmentName"].ToString().Trim();
                //                mst[iX].sKeiyakuS = Utility.fcKikan(dr["ContractStartDate"].ToString());
                //                mst[iX].sKeiyakuE = Utility.fcKikan(dr["ContractEndDate"].ToString().Trim());
                //                mst[iX].sWorkTimeS = Utility.fcTime(dr["WeeklyContractTimeSlot1Start"].ToString());
                //                mst[iX].sWorkTimeE = Utility.fcTime(dr["WeeklyContractTimeSlot1End"].ToString());
                //                mst[iX].sWorkTimeS2 = Utility.fcTime(dr["WeeklyContractTimeSlot2Start"].ToString());
                //                mst[iX].sWorkTimeE2 = Utility.fcTime(dr["WeeklyContractTimeSlot2End"].ToString());
                //                mst[iX].sSID = Utility.subStringRight(dr["SalarySystemCode"].ToString(), 2).PadLeft(2, '0');
                //                mst[iX].sSName = dr["SalarySystemName"].ToString().Trim();
                //                iX++;

                //                sReBuf = dr["EmployeeID"].ToString().Trim();
                //            }
                //        }
                //        else
                //        {
                //            // 昭和に対応　2012/08/26
                //            if (mst[iX - 1].sKeiyakuS != string.Empty)
                //            {
                //                if (mst[iX - 1].sKeiyakuS.Substring(0, 1) == "H")
                //                    date1 = mst[iX - 1].sKeiyakuS.Substring(1, 2) + mst[iX - 1].sKeiyakuS.Substring(4, 2) + mst[iX - 1].sKeiyakuS.Substring(7, 2);
                //                else date1 = "00" + mst[iX - 1].sKeiyakuS.Substring(4, 2) + mst[iX - 1].sKeiyakuS.Substring(7, 2); // 昭和のときは一時的に年は0にする
                //            }
                //            else date1 = "0";

                //            dt = Utility.fcKikan(dr["ContractStartDate"].ToString());

                //            // 昭和に対応　2012/08/26
                //            if (dt != string.Empty)
                //            {
                //                if (dt.Substring(0, 1) == "H")
                //                    date2 = dt.Substring(1, 2) + dt.Substring(4, 2) + dt.Substring(7, 2);
                //                else date2 = "00" + dt.Substring(4, 2) + dt.Substring(7, 2); // 昭和のときは一時的に年は0にする
                //            }
                //            else date2 = "0";

                //            if (int.Parse(date1) < int.Parse(date2))
                //            {
                //                // 2010/10/13  契約開始日が環境設定年月の月初日以前を対象とする
                //                dt = Utility.fcKikan(dr["ContractStartDate"].ToString());

                //                // 昭和に対応　2012/08/26
                //                if (dt != string.Empty)
                //                {
                //                    if (dt.Substring(0, 1) == "H")
                //                        date1 = dt.Substring(1, 2) + dt.Substring(4, 2) + dt.Substring(7, 2);
                //                    else date1 = "00" + dt.Substring(4, 2) + dt.Substring(7, 2); // 昭和のときは一時的に年は0にする
                //                }
                //                else date1 = "0";

                //                if (int.Parse(date1) <= int.Parse(Utility.SeitoWa(global.sYear.ToString()) + global.sMonth.ToString("00") + "01"))
                //                {
                //                    mst[iX - 1].sCode = Utility.subStringRight(dr["EmployeeNo"].ToString(), 7);
                //                    mst[iX - 1].sName = dr["Name"].ToString().Trim();
                //                    mst[iX - 1].sKana = dr["NameKana"].ToString().Trim();
                //                    mst[iX - 1].sTenpoCode = Utility.subStringRight(dr["DepartmentCode"].ToString(), 3);
                //                    mst[iX - 1].sTenpoName = dr["DepartmentName"].ToString().Trim();
                //                    mst[iX - 1].sKeiyakuS = Utility.fcKikan(dr["ContractStartDate"].ToString());
                //                    mst[iX - 1].sKeiyakuE = Utility.fcKikan(dr["ContractEndDate"].ToString().Trim());
                //                    mst[iX - 1].sWorkTimeS = Utility.fcTime(dr["WeeklyContractTimeSlot1Start"].ToString());
                //                    mst[iX - 1].sWorkTimeE = Utility.fcTime(dr["WeeklyContractTimeSlot1End"].ToString());
                //                    mst[iX - 1].sWorkTimeS2 = Utility.fcTime(dr["WeeklyContractTimeSlot2Start"].ToString());
                //                    mst[iX - 1].sWorkTimeE2 = Utility.fcTime(dr["WeeklyContractTimeSlot2End"].ToString());
                //                    mst[iX - 1].sSID = Utility.subStringRight(dr["SalarySystemCode"].ToString(), 2).PadLeft(2, '0');
                //                    mst[iX - 1].sSName = dr["SalarySystemName"].ToString().Trim();
                //                }
                //            }
                //        }
                //    }
                //}

            }

            dr.Close();

            //'----------------------------------------------------------------------------------
            //'   給与奉行iシリーズ対応　2010/10/11
            //'----------------------------------------------------------------------------------
            mySQL = "SELECT SalarySystemCode, SalarySystemName FROM tbSalarySystem ORDER BY SalarySystemCode";
            dr = sdcon.free_dsReader(mySQL);
            iX = 0;
        
            while (dr.Read())
	        {
                //2件目以降なら要素数を追加
                if (iX != 0) sMst.CopyTo(sMst = new Entity.ShokushuMst[iX + 1], 0);

                // 構造体にデータをセットする
                sMst[iX].sCode = Utility.subStringRight(dr["SalarySystemCode"].ToString(), 2).PadLeft(2, '0');
                sMst[iX].sName = dr["SalarySystemName"].ToString().Trim();
                iX++;
	        }
            dr.Close();

            sdcon.Close();
        }

        private void cmbShokushuLoad()
        {
            cmbShokushu.Items.Clear();
            cmbShokushu.Items.Add("全て表示");

            for (int i = 1; i < sMst.Length; i++)
            {
                cmbShokushu.Items.Add(sMst[i].sName);
            }

            cmbShokushu.SelectedIndex = 0;
        }

        private void cmbShokushu_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sID = string.Empty;
            if (cmbShokushu.SelectedIndex == 0) sID = string.Empty;
            else sID = sMst[cmbShokushu.SelectedIndex].sCode;

            GridviewSet.Show(dataGridView1, mst, sID);
        }
    }
}