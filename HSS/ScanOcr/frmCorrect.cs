using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient; // 2016/10/12
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using HSS2.Model;
using Leadtools;
using Leadtools.Codecs;
using Leadtools.ImageProcessing;
using Leadtools.WinForms;

namespace HSS2.ScanOcr
{
    public partial class frmCorrect : Form
    {
        public frmCorrect(Int64 sComID, string sComName, string sDbName)
        {
            InitializeComponent();

            // 会社情報を取得
            _sComID = sComID;
            _sComName = sComName;
            _sDBName = sDbName;

            // tagを初期化
            this.Tag = END_MAKEDATA;

            // 環境設定情報取得
            global.GetCommonYearMonth();

            // 受け渡しデータの出力先フォルダがあるか？なければ作成する
            global.sDATCOM = global.sDAT + global.pblComNo.ToString().PadLeft(2, '0') + " " + global.pblComName + @"\";
            if (!System.IO.Directory.Exists(global.sDATCOM)) System.IO.Directory.CreateDirectory(global.sDATCOM);

            // CSV・画像ファイルの退避先フォルダがあるか？なければ作成する
            global.sTIFCOM = global.sTIF + global.pblComNo.ToString().PadLeft(2, '0') + " " + global.pblComName + @"\";
            if (!System.IO.Directory.Exists(global.sTIFCOM)) System.IO.Directory.CreateDirectory(global.sTIFCOM);

            //// 登録モードのとき新規データを追加します
            //if (sMode == global.sADDMODE) AddNewData(_usrSel);

            // 対象会社が八十二リースか？  2017/12/08
            if (_sComName.Contains(Properties.Settings.Default.lease82))
            {
                global._82lease = int.Parse(global.FLGON);
            }
            else
            {
                global._82lease = int.Parse(global.FLGOFF);
            }
        }

        int _usrSel = 0;

        // 社員マスター配列
        Entity.ShainMst[] mst = new Entity.ShainMst[1];

        // 職種マスター配列
        Entity.ShokushuMst[] sMst = new Entity.ShokushuMst[1];

        //終了ステータス
        const string END_BUTTON = "btn";
        const string END_MAKEDATA = "data";
        const string END_CONTOROL = "close";

        // 会社ID
        private Int64 _sComID;

        // 会社名
        private string _sComName;

        // データベース名
        private string _sDBName;

        // 入力パス
        private string _InPath;

        // データグリッドビューカラム定義
        private string cDay = "col1";
        private string cWeek = "col2";
        private string cTokMK = "col3";
        private string cShoMK = "col4";
        private string cSH = "col5";
        private string cS = "col6";
        private string cSM = "col7";
        private string cEH = "col8";
        private string cE = "col9";
        private string cEM = "col10";
        private string cHiru1 = "col11";
        private string cKyuka = "col12";
        private string cKyukei = "col13";
        private string cRSH = "col14";
        private string cRS = "col15";
        private string cRSM = "col16";
        private string cREH = "col17";
        private string cRE = "col18";
        private string cREM = "col19";
        private string cTeisei = "col20";
        private string cID = "col21";

        //datagridview表示行数
        private const int _MULTIGYO = 31;

        // MDBデータキー配列
        Entity.kData[] inData;

        //カレントデータインデックス
        private int _cI;

        /// <summary>
        /// データグリッドビューの定義を行います
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
                tempDGV.DefaultCellStyle.Font = new Font("メイリオ", (float)9.5, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 22;
                tempDGV.RowTemplate.Height = 22;

                // 全体の高さ
                tempDGV.Height = 706;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                tempDGV.Columns.Add(cDay, "日");
                tempDGV.Columns.Add(cWeek, "曜");
                tempDGV.Columns.Add(cTokMK, "特");
                tempDGV.Columns.Add(cShoMK, "所");
                tempDGV.Columns.Add(cSH, "開");
                tempDGV.Columns.Add(cS, "");
                tempDGV.Columns.Add(cSM, "始");
                tempDGV.Columns.Add(cEH, "終");
                tempDGV.Columns.Add(cE, string.Empty);
                tempDGV.Columns.Add(cEM, "了");
                tempDGV.Columns.Add(cHiru1, "昼");
                tempDGV.Columns.Add(cKyuka, "暇");
                tempDGV.Columns.Add(cKyukei, "無");
                tempDGV.Columns.Add(cRSH, "開");
                tempDGV.Columns.Add(cRS, string.Empty);
                tempDGV.Columns.Add(cRSM, "始");
                tempDGV.Columns.Add(cREH, "終");
                tempDGV.Columns.Add(cRE, string.Empty);
                tempDGV.Columns.Add(cREM, "了");

                DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(column);
                tempDGV.Columns[19].Name = cTeisei;
                tempDGV.Columns[19].HeaderText = "訂正";
                tempDGV.Columns.Add(cID, "ID");
                tempDGV.Columns[cID].Visible = false; // IDカラムは非表示とする

                // 各列の定義を行う
                foreach (DataGridViewColumn c in tempDGV.Columns)
                {
                    // 幅
                    c.Width = 40;

                    // 表示位置、編集可否
                    if (c.Name == cDay)
                    {
                        c.Width = 30;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cWeek)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cTokMK)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cShoMK)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cSH)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cS)
                    {
                        c.Width = 15;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cSM)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cEH)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cE)
                    {
                        c.Width = 15;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cEM)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cKyuka)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cKyukei)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cHiru1)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cRSH)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cRS)
                    {
                        c.Width = 15;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cRSM)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cREH)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cRE)
                    {
                        c.Width = 15;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cREM)
                    {
                        c.Width = 26;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cTeisei)
                    {
                        c.Width = 40;
                        //c.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cID)
                    {
                        c.ReadOnly = true;
                    }

                    // 入力可能桁数
                    if (c.Name != cTeisei)
                    {
                        DataGridViewTextBoxColumn col = (DataGridViewTextBoxColumn)c;

                        if (c.Name == cTokMK) col.MaxInputLength = 1;
                        if (c.Name == cShoMK) col.MaxInputLength = 1;
                        if (c.Name == cSH) col.MaxInputLength = 2;
                        if (c.Name == cSM) col.MaxInputLength = 2;
                        if (c.Name == cEH) col.MaxInputLength = 2;
                        if (c.Name == cEM) col.MaxInputLength = 2;
                        if (c.Name == cKyuka) col.MaxInputLength = 1;
                        if (c.Name == cKyukei) col.MaxInputLength = 1;
                        if (c.Name == cHiru1) col.MaxInputLength = 1;
                        if (c.Name == cRSH) col.MaxInputLength = 2;
                        if (c.Name == cRSM) col.MaxInputLength = 2;
                        if (c.Name == cREH) col.MaxInputLength = 2;
                        if (c.Name == cREM) col.MaxInputLength = 2;
                    }
                }

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                tempDGV.MultiSelect = false;

                // 編集可とする
                //tempDGV.ReadOnly = false;

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

                // 編集モード
                tempDGV.EditMode = DataGridViewEditMode.EditOnEnter;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }
        
        private void frmCorrect_Load(object sender, EventArgs e)
        {
            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // グリッド定義
            GridViewSetting(dg1);
            
            // 奉行マスター情報を取得します
            LoadBugyoMst(global.pblDbName);

            // 入力ファイルフォルダ取得
            _InPath = Properties.Settings.Default.instPath + @"DATA\" + global.pblComNo.ToString().PadLeft(2, '0') + " " + global.pblComName + @"\";

            // CSVデータMDBへ登録
            GetCsvDataToMDB();

            //MDB件数カウント
            if (CountMDB() == 0)
            {
                MessageBox.Show("対象となる勤務票データがありません", "勤務票データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //終了処理
                this.Close();
                return;
            }

            // MDBへ所定勤務情報をセットします
            GetMstData();

            // MDBデータキー項目読み込み
            inData = LoadMdbID();

            //エラー情報初期化
            ErrInitial();

            // データ表示
            _cI = 0;
            DataShow(_cI, inData, this.dg1);
        }

        /// <summary>
        /// 勤務票ヘッダレコードに所定勤務時間,特殊勤務時間をセットします
        /// </summary>
        private void GetMstData()
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            OleDbDataReader dR;

            StringBuilder sb = new StringBuilder();

            // 勤務票ヘッダ情報から個人番号を取得します
            sCom.Parameters.Clear();
            sb.Clear();
            sb.Append("select * from 勤務票ヘッダ order by ID");
            sCom.CommandText = sb.ToString();
            dR = sCom.ExecuteReader();

            OleDbCommand sCom2 = new OleDbCommand();
            sCom2.Connection = Con.cnOpen();
            sb.Clear();
            sb.Append("update 勤務票ヘッダ set ");
            sb.Append("氏名=?,所定開始時間=?,所定終了時間=?,午前半休開始時間=?,午前半休終了時間=?, ");
            sb.Append("午後半休開始時間=?,午後半休終了時間=?,所定勤務時間=?, ");
            sb.Append("特殊開始時間=?,特殊終了時間=?,特殊午前半休開始時間=?,特殊午前半休終了時間=?, ");
            sb.Append("特殊午後半休開始時間=?,特殊午後半休終了時間=?,特殊勤務時間=? ");
            sb.Append("where 個人番号=?");
            sCom2.CommandText = sb.ToString();

            string TimeS = string.Empty;
            string TimeE = string.Empty;
            string TimeS2 = string.Empty;
            string TimeE2 = string.Empty;

            while (dR.Read())
            {
                sCom2.Parameters.Clear();

                for (int i = 0; i < mst.Length; i++)
                {
                    if (mst[i].sCode == dR["個人番号"].ToString())
                    {
                        string sSTime = mst[i].sWorkTimeS;
                        string sETime = mst[i].sWorkTimeE;
                        string sSTime2 = mst[i].sWorkTimeS2;
                        string sETime2 = mst[i].sWorkTimeE2;
                        int TimeJ = 0;
                        string amE = string.Empty;
                        string pmS = string.Empty;
                        double ShoteiTime = 0;
                        string amE2 = string.Empty;
                        string pmS2 = string.Empty;
                        double ShoteiTime2 = 0;
                        double ShoteiTime3 = 0; // 八十二リース特殊日午後所定勤務時間：2017/12/12
                        bool r = false;

                        // 所定勤務時間の開始時間、終了時間共に登録されている場合に所定勤務時間を取得します
                        if (sSTime != string.Empty && sETime != string.Empty)
                        {
                            TimeS = sSTime.PadLeft(5, '0');
                            TimeE = sETime.PadLeft(5, '0');
                            TimeJ = (Utility.TimeToMin(TimeE) - Utility.TimeToMin(TimeS) - 60) / 2;

                            DateTime dd = DateTime.Parse(TimeS);

                            // 午前半休終了時間取得 
                            amE = dd.AddMinutes(TimeJ).ToShortTimeString();

                            // 午後半休開始時間取得 
                            pmS = dd.AddMinutes(TimeJ + 60).ToShortTimeString();

                            // 所定勤務時間
                            ShoteiTime = double.Parse(TimeJ.ToString()) / 60;

                            // 書き換えステータス
                            r = true;
                        }

                        // 特殊勤務時間の開始時間、終了時間共に登録されている場合に特殊勤務時間を取得します
                        if (sSTime2 != string.Empty && sETime2 != string.Empty)
                        {
                            // 特殊勤務
                            TimeS2 = sSTime2.PadLeft(5, '0');
                            TimeE2 = sETime2.PadLeft(5, '0');

                            // 八十二リースの社員のとき：2017/12/21
                            if (global._82lease == int.Parse(global.FLGON) && Utility.StrToInt(mst[i].sSID) != 7)
                            {
                                // 八十二リース：2017/12/12
                                // 特殊勤務午前半休終了時間取得 
                                amE2 = "12:25";

                                // 特殊勤務午後半休開始時間取得 
                                pmS2 = "13:25";

                                // 特殊勤務所定勤務時間(午前勤務）
                                ShoteiTime2 = 225;

                                // 特殊勤務所定勤務時間(午後勤務）
                                ShoteiTime3 = 245;
                            }
                            else
                            {
                                // 八十二リースの社員以外のとき（リコーリースのパートタイマーも「以外」に含む）
                                TimeJ = (Utility.TimeToMin(TimeE2) - Utility.TimeToMin(TimeS2) - 60) / 2;

                                DateTime dd = DateTime.Parse(TimeS2);

                                // 特殊勤務午前半休終了時間取得 
                                amE2 = dd.AddMinutes(TimeJ).ToShortTimeString();

                                // 特殊勤務午後半休開始時間取得 
                                pmS2 = dd.AddMinutes(TimeJ + 60).ToShortTimeString();

                                // 特殊勤務所定勤務時間
                                ShoteiTime2 = double.Parse(TimeJ.ToString()) / 60;
                            }

                            // 八十二リースでかつ社員のみに限定のため下記コメント化 2017/12/21
                            //if (global._82lease == int.Parse(global.FLGOFF))
                            //{
                            //    TimeJ = (Utility.TimeToMin(TimeE2) - Utility.TimeToMin(TimeS2) - 60) / 2;

                            //    DateTime dd = DateTime.Parse(TimeS2);

                            //    // 特殊勤務午前半休終了時間取得 
                            //    amE2 = dd.AddMinutes(TimeJ).ToShortTimeString();

                            //    // 特殊勤務午後半休開始時間取得 
                            //    pmS2 = dd.AddMinutes(TimeJ + 60).ToShortTimeString();

                            //    // 特殊勤務所定勤務時間
                            //    ShoteiTime2 = double.Parse(TimeJ.ToString()) / 60;
                            //}
                            //else
                            //{
                            //    // 八十二リース：2017/12/12
                            //    // 特殊勤務午前半休終了時間取得 
                            //    amE2 = "12:25";

                            //    // 特殊勤務午後半休開始時間取得 
                            //    pmS2 = "13:25";

                            //    // 特殊勤務所定勤務時間(午前勤務）
                            //    ShoteiTime2 = 225;

                            //    // 特殊勤務所定勤務時間(午後勤務）
                            //    ShoteiTime3 = 245;
                            //}

                            // 書き換えステータス
                            r = true;
                        }

                        if (r)
                        {
                            sCom2.Parameters.AddWithValue("@Name", mst[i].sName);
                            sCom2.Parameters.AddWithValue("@ShoS", TimeS);
                            sCom2.Parameters.AddWithValue("@ShoE", TimeE);
                            sCom2.Parameters.AddWithValue("@amS", TimeS);
                            sCom2.Parameters.AddWithValue("@amE", amE);
                            sCom2.Parameters.AddWithValue("@pmS", pmS);
                            sCom2.Parameters.AddWithValue("@pmE", TimeE);
                            sCom2.Parameters.AddWithValue("@Sho", ShoteiTime);

                            sCom2.Parameters.AddWithValue("@ShoS2", TimeS2);
                            sCom2.Parameters.AddWithValue("@ShoE2", TimeE2);
                            sCom2.Parameters.AddWithValue("@amS2", TimeS2);
                            sCom2.Parameters.AddWithValue("@amE2", amE2);
                            sCom2.Parameters.AddWithValue("@pmS2", pmS2);
                            sCom2.Parameters.AddWithValue("@pmE2", TimeE2);
                            sCom2.Parameters.AddWithValue("@Sho2", ShoteiTime2);

                            sCom2.Parameters.AddWithValue("@sCode", dR["個人番号"].ToString());

                            sCom2.ExecuteNonQuery();

                        }
                        break;
                    }
                }
            }

            dR.Close();
            sCom.Connection.Close();
            sCom2.Connection.Close();
        }

        /// <summary>
        /// 個人別の所定勤務時間を取得します
        /// </summary>
        /// <param name="sCode">個人番号</param>
        private void GetShoteiData(string sCode)
        {
            string TimeS = string.Empty;
            string TimeE = string.Empty;
            string TimeS2 = string.Empty;
            string TimeE2 = string.Empty; 
            //string amE =

            // 該当項目を初期化します
            this.lblName.Text = string.Empty;
            this.txtShozokuCode.Text = string.Empty;
            this.txtShokuID.Text = string.Empty;
            this.lblShozoku.Text = string.Empty;
            this.lblShoteiS.Text = string.Empty;
            this.lblShoteiE.Text = string.Empty;
            this.lblHanAmS.Text = string.Empty;
            this.lblHanAmE.Text = string.Empty;
            this.lblHanPmS.Text = string.Empty;
            this.lblHanPmE.Text = string.Empty;
            this.lblTokushuS.Text = string.Empty;
            this.lblTokushuE.Text = string.Empty;
            this.lblTokAmS.Text = string.Empty;
            this.lblTokAmE.Text = string.Empty;
            this.lblTokPmS.Text = string.Empty;
            this.lblTokPmE.Text = string.Empty;

            // マスター情報を表示します
            for (int i = 0; i < mst.Length; i++)
            {
                if (mst[i].sCode == sCode)
                {
                    // マスター情報を画面表示します
                    this.lblName.Text = mst[i].sName;
                    this.txtShozokuCode.Text = mst[i].sTenpoCode;
                    this.lblShozoku.Text = mst[i].sTenpoName;
                    this.txtShokuID.Text = mst[i].sSID.PadLeft(2, '0');

                    string sSTime = mst[i].sWorkTimeS;
                    string sETime = mst[i].sWorkTimeE;
                    string sSTime2 = mst[i].sWorkTimeS2;
                    string sETime2 = mst[i].sWorkTimeE2;
                    
                    // 所定勤務時間の開始時間、終了時間共に登録されている場合に所定勤務時間を取得します
                    if (sSTime != string.Empty && sETime != string.Empty)
                    {
                        TimeS = sSTime.PadLeft(5, '0');
                        TimeE = sETime.PadLeft(5, '0');
                        int TimeJ = (Utility.TimeToMin(TimeE) - Utility.TimeToMin(TimeS) - 60) / 2;

                        DateTime dd = DateTime.Parse(TimeS);

                        // 午前半休終了時間取得 
                        string amE = dd.AddMinutes(TimeJ).ToShortTimeString();

                        // 午後半休開始時間取得 
                        string pmS = dd.AddMinutes(TimeJ + 60).ToShortTimeString();

                        // 所定勤務時間
                        double ShoteiTime = double.Parse(TimeJ.ToString()) / 60;
                        inData[_cI]._ShoTime = ShoteiTime;

                        // 所定勤務時間関連情報
                        lblShoteiS.Text = TimeS;    // 開始
                        lblShoteiE.Text = TimeE;    // 終了
                        lblHanAmS.Text = TimeS;     // 半休午前開始
                        lblHanAmE.Text = amE;       // 半休午前終了
                        lblHanPmS.Text = pmS;       // 半休午後開始
                        lblHanPmE.Text = TimeE;     // 半休午後終了
                    }
                    
                    // 特殊勤務時間の開始時間、終了時間共に登録されている場合に特殊勤務時間を取得します
                    if (sSTime2 != string.Empty && sETime2 != string.Empty)
                    {
                        // 特殊勤務
                        TimeS2 = sSTime2.PadLeft(5, '0');
                        TimeE2 = sETime2.PadLeft(5, '0');
                        int TimeJ = 0;
                        string amE2 = string.Empty;
                        string pmS2 = string.Empty;
                        double ShoteiTime2 = 0;
                        double ShoteiTime3 = 0;

                        // 八十二リースの社員のとき：2017/12/21
                        if (global._82lease == int.Parse(global.FLGON) && Utility.StrToInt(mst[i].sSID) != 7)
                        {
                            // 特殊勤務午前半休終了時間取得 
                            amE2 = "12:25";

                            // 特殊勤務午後半休開始時間取得 
                            pmS2 = "13:25";

                            // 特殊勤務所定勤務時間(午前勤務）
                            ShoteiTime2 = 225;
                            inData[_cI]._TokTime = ShoteiTime2;

                            // 特殊勤務所定勤務時間(午後勤務）
                            ShoteiTime3 = 245;
                            inData[_cI]._TokTime2 = ShoteiTime3;
                        }
                        else
                        {
                            TimeJ = (Utility.TimeToMin(TimeE2) - Utility.TimeToMin(TimeS2) - 60) / 2;

                            DateTime dd = DateTime.Parse(TimeS2);

                            // 特殊勤務午前半休終了時間取得 
                            amE2 = dd.AddMinutes(TimeJ).ToShortTimeString();

                            // 特殊勤務午後半休開始時間取得 
                            pmS2 = dd.AddMinutes(TimeJ + 60).ToShortTimeString();

                            // 特殊勤務所定勤務時間
                            ShoteiTime2 = double.Parse(TimeJ.ToString()) / 60;
                            inData[_cI]._TokTime = ShoteiTime2;
                        }


                        // 八十二リースの社員限定のため以下をコメント化 2017/12/21
                        //// 八十二リース以外：2017/12/12
                        //if (global._82lease == int.Parse(global.FLGOFF))
                        //{
                        //    TimeJ = (Utility.TimeToMin(TimeE2) - Utility.TimeToMin(TimeS2) - 60) / 2;

                        //    DateTime dd = DateTime.Parse(TimeS2);

                        //    // 特殊勤務午前半休終了時間取得 
                        //    amE2 = dd.AddMinutes(TimeJ).ToShortTimeString();

                        //    // 特殊勤務午後半休開始時間取得 
                        //    pmS2 = dd.AddMinutes(TimeJ + 60).ToShortTimeString();

                        //    // 特殊勤務所定勤務時間
                        //    ShoteiTime2 = double.Parse(TimeJ.ToString()) / 60;
                        //    inData[_cI]._TokTime = ShoteiTime2;
                        //}
                        //else
                        //{
                        //    // 八十二リース：2017/12/12
                        //    // 特殊勤務午前半休終了時間取得 
                        //    amE2 = "12:25";

                        //    // 特殊勤務午後半休開始時間取得 
                        //    pmS2 = "13:25";

                        //    // 特殊勤務所定勤務時間(午前勤務）
                        //    ShoteiTime2 = 225;
                        //    inData[_cI]._TokTime = ShoteiTime2;

                        //    // 特殊勤務所定勤務時間(午後勤務）
                        //    ShoteiTime3 = 245;
                        //    inData[_cI]._TokTime2 = ShoteiTime3;
                        //}

                        // 特殊勤務時間
                        lblTokushuS.Text = TimeS2;  // 開始
                        lblTokushuE.Text = TimeE2;  // 終了
                        lblTokAmS.Text = TimeS2;    // 半休午前開始
                        lblTokAmE.Text = amE2;      // 半休午前終了
                        lblTokPmS.Text = pmS2;      // 半休午後開始
                        lblTokPmE.Text = TimeE2;    // 半休午後終了
                    }

                    break;
                }
            }
        }

        private void dg1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex == 4 || e.ColumnIndex == 5 || e.ColumnIndex == 7 || e.ColumnIndex == 8)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;

                if (e.ColumnIndex == 5 || e.ColumnIndex == 6 || e.ColumnIndex == 8 || e.ColumnIndex == 9)
                {
                    e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None;
                }
                //else
                //    e.AdvancedBorderStyle.Left = dg1.AdvancedCellBorderStyle.Left;
            }

            if (e.ColumnIndex == 13 || e.ColumnIndex == 14 || e.ColumnIndex == 16 || e.ColumnIndex == 17)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;

                if (e.ColumnIndex == 14 || e.ColumnIndex == 15 || e.ColumnIndex == 17 || e.ColumnIndex == 18)
                {
                    e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None;
                }
                //else
                //    e.AdvancedBorderStyle.Left = dg1.AdvancedCellBorderStyle.Left;
            }
        }

        /// <summary>
        /// CSVデータをMDBへインサートする
        /// </summary>
        private void GetCsvDataToMDB()
        {
            //CSVファイル数をカウント
            string[] inCsv = System.IO.Directory.GetFiles(_InPath, "*.csv");

            //CSVファイルがなければ終了
            int cTotal = 0;
            if (inCsv.Length == 0) return;
            else cTotal = inCsv.Length;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // データベースへ接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                //CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = cCnt / cTotal * 100;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    // CSVファイルインポート
                    var s = File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // MDBへ登録する
                        // 勤務記録ヘッダテーブル
                        StringBuilder sb = new StringBuilder();
                        sb.Clear();
                        sb.Append("insert into 勤務票ヘッダ ");
                        sb.Append("(ID,会社ID,個人番号,年,月,職種ID,部門コード,");
                        sb.Append("出勤日数合計,昼食数合計,有休日数合計,特休日数合計,欠勤日数合計,");
                        sb.Append("画像名,更新年月日) ");
                        sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?,?)");

                        sCom.CommandText = sb.ToString();
                        sCom.Parameters.Clear();

                        // ID
                        sCom.Parameters.AddWithValue("@ID", Utility.GetStringSubMax(stCSV[0], 17));
                        // 会社ＩＤ
                        sCom.Parameters.AddWithValue("@ComID", Utility.GetStringSubMax(stCSV[2], 2).PadLeft(2,'0'));
                        // 個人番号
                        sCom.Parameters.AddWithValue("@kjn", Utility.GetStringSubMax(stCSV[3], 7));
                        // 年
                        sCom.Parameters.AddWithValue("@year", Utility.GetStringSubMax(stCSV[4], 4));
                        // 月
                        sCom.Parameters.AddWithValue("@month", Utility.GetStringSubMax(stCSV[5], 2));
                        // 職種ID
                        sCom.Parameters.AddWithValue("@ShoID", Utility.GetStringSubMax(stCSV[6], 2).PadLeft(2, '0'));
                        // 所属コード
                        sCom.Parameters.AddWithValue("@Szk", Utility.GetStringSubMax(stCSV[7], 3));
                        // 出勤日数合計
                        sCom.Parameters.AddWithValue("@ShukinTL", Utility.GetStringSubMax(stCSV[442], 2));
                        // 昼食回数合計
                        sCom.Parameters.AddWithValue("@ChushokuTL", Utility.GetStringSubMax(stCSV[443], 2));

                        // 有休日数合計
                        string yukyu = Utility.GetStringSubMax(stCSV[444], 2);
                        if (yukyu == string.Empty) yukyu = "0";
                        if (Utility.GetStringSubMax(stCSV[445], 1) != string.Empty)
                            yukyu += "." + Utility.GetStringSubMax(stCSV[445], 1);
                        //if (yukyu.Substring(0, 1) == ".")
                        //    yukyu = "0" + yukyu;
                        sCom.Parameters.AddWithValue("@YukyuTL", yukyu);

                        // 特休日数合計
                        sCom.Parameters.AddWithValue("@TokuTL", Utility.GetStringSubMax(stCSV[446], 2));
                        // 欠勤日数合計
                        sCom.Parameters.AddWithValue("@KekkinTL", Utility.GetStringSubMax(stCSV[447], 2));
                        // 画像名
                        sCom.Parameters.AddWithValue("@IMG", Utility.GetStringSubMax(stCSV[1], 21));
                        // 更新年月日
                        sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());

                        // テーブル書き込み
                        sCom.ExecuteNonQuery();

                        // 勤務票明細テーブル
                        sb.Clear();
                        sb.Append("insert into 勤務票明細 ");
                        sb.Append("(ヘッダID,日付,特殊日,所定勤務,開始時,開始分,終了時,終了分,昼食,休暇,");
                        sb.Append("休憩なし,離席開始時,離席開始分,離席終了時,離席終了分,訂正,休日,更新年月日) ");
                        sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
                        sCom.CommandText = sb.ToString();

                        int sDays = 0;
                        DateTime dt;

                        for (int i = 8; i <= 428; i += 14)
                        {
                            sCom.Parameters.Clear();
                            sDays++;

                            // 存在する日付のときにMDBへ登録する
                            string tempDt = global.sYear + "/" + global.sMonth + "/" + sDays.ToString();
                            if (DateTime.TryParse(tempDt, out dt))
                            {
                                // ヘッダID
                                sCom.Parameters.AddWithValue("@ID", Utility.GetStringSubMax(stCSV[0], 17));

                                // 日付
                                sCom.Parameters.AddWithValue("@Days", sDays);

                                // 特殊日
                                sCom.Parameters.AddWithValue("@TokMK", stCSV[i]);

                                // 所定勤務
                                sCom.Parameters.AddWithValue("@ShoMK", stCSV[i + 1]);

                                // 開始時間
                                string hh = Utility.GetStringSubMax(stCSV[i + 2], 2);
                                string mm = Utility.GetStringSubMax(stCSV[i + 3], 2);
                                sCom.Parameters.AddWithValue("@sh", hh);
                                sCom.Parameters.AddWithValue("@sm", mm);

                                // 終了時間
                                hh = Utility.GetStringSubMax(stCSV[i + 4], 2);
                                mm = Utility.GetStringSubMax(stCSV[i + 5], 2);
                                sCom.Parameters.AddWithValue("@eh", hh);
                                sCom.Parameters.AddWithValue("@em", mm);

                                // 昼食
                                sCom.Parameters.AddWithValue("@hiru1", stCSV[i + 6]);

                                // 休暇
                                sCom.Parameters.AddWithValue("@kyukei", stCSV[i + 7]);

                                // 休憩なし
                                sCom.Parameters.AddWithValue("@kyunashi", stCSV[i + 8]);

                                // 開始時間
                                hh = Utility.GetStringSubMax(stCSV[i + 9], 2);
                                mm = Utility.GetStringSubMax(stCSV[i + 10], 2);
                                sCom.Parameters.AddWithValue("@rsh", hh);
                                sCom.Parameters.AddWithValue("@rsm", mm);

                                // 終了時間
                                hh = Utility.GetStringSubMax(stCSV[i + 11], 2);
                                mm = Utility.GetStringSubMax(stCSV[i + 12], 2);
                                sCom.Parameters.AddWithValue("@reh", hh);
                                sCom.Parameters.AddWithValue("@rem", mm);

                                // 訂正
                                sCom.Parameters.AddWithValue("@teisei", stCSV[i + 13]);

                                // 休日区分の設定
                                // ①曜日で判断します
                                string Youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                                int sHol = 0;
                                if (Youbi == "土")
                                {
                                    sHol = global.hSATURDAY;
                                }
                                else if (Youbi == "日")
                                {
                                    sHol = global.hHOLIDAY;
                                }
                                else
                                {
                                    sHol = global.hWEEKDAY;
                                }

                                // ②休日テーブルを参照し休日に該当するか調べます
                                SysControl.SetDBConnect dc = new SysControl.SetDBConnect();
                                OleDbCommand sCom2 = new OleDbCommand();
                                sCom2.Connection = dc.cnOpen();
                                OleDbDataReader dr = null;
                                sCom2.CommandText = "select * from 休日 where 年=? and 月=? and 日=? and 会社ID=?";
                                sCom2.Parameters.Clear();
                                sCom2.Parameters.AddWithValue("@year", global.sYear);
                                sCom2.Parameters.AddWithValue("@Month", global.sMonth);
                                sCom2.Parameters.AddWithValue("@day", sDays);
                                sCom2.Parameters.AddWithValue("@id", global.pblComNo.ToString("00"));
                                dr = sCom2.ExecuteReader();
                                if (dr.Read())
                                {
                                    sHol = global.hHOLIDAY;
                                }
                                dr.Close();
                                sCom2.Connection.Close();

                                // ③休暇で判断
                                // 振替休暇のとき
                                if (stCSV[i + 5].Trim() == global.eFURIKYU)
                                {
                                    sHol = global.hHOLIDAY;
                                }
                                else if (stCSV[i + 5].Trim() == global.eFURIDE)　// 振替出勤のときは平日扱い
                                {
                                    sHol = global.hWEEKDAY;
                                }

                                sCom.Parameters.AddWithValue("@Hol", sHol);

                                // 更新年月日
                                sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());

                                // テーブル書き込み
                                sCom.ExecuteNonQuery();
                            }
                        }             
                    }
                }

                // トランザクションコミット
                sTran.Commit();

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                //ＣＳＶファイルを退避先フォルダへ移動する        
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    File.Move(files, global.sTIFCOM + System.IO.Path.GetFileName(files));
                }

                ////CSVファイルを削除する
                //foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                //{
                //    System.IO.File.Delete(files);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);

                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// MDBデータの件数をカウントする
        /// </summary>
        /// <returns></returns>
        private int CountMDB()
        {
            int rCnt = 0;

            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();
            OleDbDataReader dR;
            sCom.CommandText = "select ID from 勤務票ヘッダ order by ID";
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数加算
                rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            return rCnt;
        }

        /// <summary>
        /// MDBデータのキー項目を配列に読み込む
        /// </summary>
        /// <returns></returns>
        private Entity.kData[] LoadMdbID()
        {
            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            
            frmP.Show();

            //レコード件数取得
            int cTotal = CountMDB();

            Entity.kData[] DenID = new Entity.kData[1];

            int rCnt = 1;

            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();
            OleDbDataReader dR;
            sCom.CommandText = "select ID,所定勤務時間 from 勤務票ヘッダ order by ID";
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //プログレスバー表示
                frmP.Text = "勤務票データロード中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt / cTotal * 100;
                frmP.ProgressStep();

                //2件目以降は要素数を追加
                if (rCnt > 1)
                    DenID.CopyTo(DenID = new Entity.kData[rCnt], 0);

                DenID[rCnt - 1]._Meisai = new Entity.Gyou[_MULTIGYO];
                DenID[rCnt - 1]._sID = dR["ID"].ToString();
                DenID[rCnt - 1]._ShoTime = double.Parse(dR["所定勤務時間"].ToString());

                //データ件数加算
                rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            return DenID;
        }

        private void ErrInitial()
        {
            //エラー情報初期化
            lblErrMsg.Visible = false;
            global.errNumber = global.eNothing;     //エラー番号
            global.errMsg = string.Empty;           //エラーメッセージ
            lblErrMsg.Text = string.Empty;
        }

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     勤怠データを画面に表示します </summary>
        /// <param name="sIx">
        ///     勤怠データインデックス</param>
        /// <param name="rRec">
        ///     インデックス配列</param>
        /// <param name="dgv">
        ///     データグリッドビューオブジェクト</param>
        ///-----------------------------------------------------------------------------
        private void DataShow(int sIx, Entity.kData[] rRec, DataGridView dgv)
        {
            // 出勤日数合計
            int rDays = 0;

            // 昼食回数
            int Lunchs = 0;

            string SqlStr = string.Empty;

            // 画像ファイル名
            global.pblImageFile = string.Empty;

            // データグリッドビュー初期化
            dataGridInitial(this.dg1);

            //データ表示背景色初期化
            dsColorInitial(this.dg1);

            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;

            try
            {
                // 勤務票ヘッダ
                sCom.Connection = dCon.cnOpen();
                sCom.CommandText = "select * from 勤務票ヘッダ where ID = ?";
                sCom.Parameters.AddWithValue("@ID", rRec[sIx]._sID);

                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    txtID.Text = Utility.EmptytoZero(dR["会社ID"].ToString());
                    //txtYear.Text = Utility.EmptytoZero(dR["年"].ToString()).Substring(2, 2);    // 2019/07/22 コメント化
                    txtYear.Text = Utility.EmptytoZero(dR["年"].ToString()).PadLeft(4, '0').Substring(2, 2); // 4桁文字列を生成後に2文字切り出し 2019/07/22
                    txtMonth.Text = Utility.EmptytoZero(dR["月"].ToString());
                    txtNo.Text = Utility.EmptytoZero(dR["個人番号"].ToString());

                    if (Utility.EmptytoZero(dR["部門コード"].ToString()).Length > 8)
                    {
                        txtShozokuCode.Text = Utility.EmptytoZero(dR["部門コード"].ToString());
                    }
                    else
                    {
                        txtShozokuCode.Text = Utility.EmptytoZero(dR["部門コード"].ToString());
                    }

                    txtShokuID.Text = Utility.EmptytoZero(dR["職種ID"].ToString());

                    txtShuTl.Text = Utility.EmptytoZero(dR["出勤日数合計"].ToString());
                    txtChuTl.Text = Utility.EmptytoZero(dR["昼食数合計"].ToString());
                    txtYukyuTl.Text = Utility.EmptytoZero(dR["有休日数合計"].ToString());
                    txttokuTl.Text = Utility.EmptytoZero(dR["特休日数合計"].ToString());
                    txtKekkinTl.Text = Utility.EmptytoZero(dR["欠勤日数合計"].ToString());
                    global.pblImageFile = dR["画像名"].ToString();

                    lblShoteiS.Text = dR["所定開始時間"].ToString();
                    lblShoteiE.Text = dR["所定終了時間"].ToString();
                    lblHanAmS.Text = dR["午前半休開始時間"].ToString();
                    lblHanAmE.Text = dR["午前半休終了時間"].ToString();
                    lblHanPmS.Text = dR["午後半休開始時間"].ToString();
                    lblHanPmE.Text = dR["午後半休終了時間"].ToString();

                    lblTokushuS.Text = dR["特殊開始時間"].ToString();
                    lblTokushuE.Text = dR["特殊終了時間"].ToString();
                    lblTokAmS.Text = dR["特殊午前半休開始時間"].ToString();
                    lblTokAmE.Text = dR["特殊午前半休終了時間"].ToString();
                    lblTokPmS.Text = dR["特殊午後半休開始時間"].ToString();
                    lblTokPmE.Text = dR["特殊午後半休終了時間"].ToString();

                    //データ数表示
                    lblPage.Text = (_cI + 1).ToString() + "/" + inData.Length.ToString() + " 件目";
                }

                dR.Close();

                // マスター情報を取得して表示します
                GetShoteiData(txtNo.Text);

                // 勤務記録明細
                StringBuilder sb = new StringBuilder();
                sb.Append("select 勤務票明細.*,");
                sb.Append("勤務票ヘッダ.所定開始時間,勤務票ヘッダ.所定終了時間,");
                sb.Append("勤務票ヘッダ.午前半休開始時間,勤務票ヘッダ.午前半休終了時間,");
                sb.Append("勤務票ヘッダ.午後半休開始時間,勤務票ヘッダ.午後半休終了時間, ");
                sb.Append("勤務票ヘッダ.特殊開始時間,勤務票ヘッダ.特殊終了時間,");
                sb.Append("勤務票ヘッダ.特殊午前半休開始時間,勤務票ヘッダ.特殊午前半休終了時間,");
                sb.Append("勤務票ヘッダ.特殊午後半休開始時間,勤務票ヘッダ.特殊午後半休終了時間 ");
                sb.Append("from 勤務票明細 inner join 勤務票ヘッダ ");
                sb.Append("on 勤務票明細.ヘッダID = 勤務票ヘッダ.ID ");
                sb.Append("where ヘッダID = ? order by 勤務票明細.ID");
                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", rRec[sIx]._sID);
                dR = sCom.ExecuteReader();

                int r = 0;
                while (dR.Read())
                {
                    global.ShoS = dR["所定開始時間"].ToString();
                    global.ShoE = dR["所定終了時間"].ToString();
                    global.AmS = dR["午前半休開始時間"].ToString();
                    global.AmE = dR["午前半休終了時間"].ToString();
                    global.PmS = dR["午後半休開始時間"].ToString();
                    global.PmE = dR["午後半休終了時間"].ToString();

                    global.ShoS2 = dR["特殊開始時間"].ToString();
                    global.ShoE2 = dR["特殊終了時間"].ToString();
                    global.AmS2 = dR["特殊午前半休開始時間"].ToString();
                    global.AmE2 = dR["特殊午前半休終了時間"].ToString();
                    global.PmS2 = dR["特殊午後半休開始時間"].ToString();
                    global.PmE2 = dR["特殊午後半休終了時間"].ToString();

                    // ChangeValueイベント処理回避
                    global.dg1ChabgeValueStatus = false;

                    if (dR["特殊日"].ToString() == global.FLGON) dgv[cTokMK, r].Value = global.MARU;
                    else dgv[cTokMK, r].Value = string.Empty;

                    if (dR["所定勤務"].ToString() == global.FLGON) dgv[cShoMK, r].Value = global.MARU;
                    else dgv[cShoMK, r].Value = string.Empty;

                    // ChangeValueイベント処理実施
                    global.dg1ChabgeValueStatus = true;

                    if (dR["訂正"].ToString() == global.FLGON) dgv[cTeisei, r].Value = true;
                    else dgv[cTeisei, r].Value = false;

                    dgv[cKyuka, r].Value = dR["休暇"];
                    dgv[cSH, r].Value = dR["開始時"].ToString();
                    dgv[cSM, r].Value = dR["開始分"].ToString();
                    dgv[cEH, r].Value = dR["終了時"].ToString();
                    dgv[cEM, r].Value = dR["終了分"].ToString();

                    dgv[cRSH, r].Value = dR["離席開始時"].ToString();
                    dgv[cRSM, r].Value = dR["離席開始分"].ToString();
                    dgv[cREH, r].Value = dR["離席終了時"].ToString();
                    dgv[cREM, r].Value = dR["離席終了分"].ToString();

                    // ChangeValueイベント処理回避
                    global.dg1ChabgeValueStatus = false;

                    if (dR["休憩なし"].ToString() == global.FLGON) dgv[cKyukei, r].Value = global.MARU;
                    else dgv[cKyukei, r].Value = string.Empty;

                    // 昼食
                    if (dR["昼食"].ToString() == global.FLGON)
                    {
                        dgv[cHiru1, r].Value = global.MARU;
                        // 昼食回数加算
                        Lunchs++;
                    }
                    else dgv[cHiru1, r].Value = string.Empty;

                    // 出勤日数加算
                    if (dR["休日"].ToString() != global.hSATURDAY.ToString() &&
                        dR["休日"].ToString() != global.hHOLIDAY.ToString())
                    {
                        if (dR["特殊日"].ToString() == global.FLGON ||
                            dR["所定勤務"].ToString() == global.FLGON ||
                            dR["開始時"].ToString() != string.Empty ||
                            dR["休暇"].ToString() == global.eAMHANKYU ||
                            dR["休暇"].ToString() == global.ePMHANKYU)
                            rDays++;
                    }

                    dgv[cID, r].Value = dR["ID"].ToString();

                    // ChangeValueイベント処理実施
                    global.dg1ChabgeValueStatus = true;

                    r++;

                    // データグリッド最下行まで達したら終了する（月末日以降のデータは無視する）
                    if (dgv.Rows.Count == r) break;
                }

                dR.Close();
                sCom.Connection.Close();

                //画像イメージ表示
                ShowImage(_InPath + global.pblImageFile);

                // ヘッダ情報
                txtYear.ReadOnly = false;
                txtMonth.ReadOnly = false;
                txtShozokuCode.ReadOnly = false;
                txtNo.ReadOnly = false;

                // スクロールバー設定
                hScrollBar1.Enabled = true;
                hScrollBar1.Minimum = 0;
                hScrollBar1.Maximum = rRec.Length - 1;
                hScrollBar1.Value = sIx;
                hScrollBar1.LargeChange = 1;
                hScrollBar1.SmallChange = 1;

                //最初のレコード
                if (sIx == 0)
                {
                    btnBefore.Enabled = false;
                    btnFirst.Enabled = false;
                }

                //最終レコード
                if ((sIx + 1) == rRec.Length)
                {
                    btnNext.Enabled = false;
                    btnEnd.Enabled = false;
                }

                //カレントセル選択状態としない
                dgv.CurrentCell = null;

                // その他のボタンを有効とする
                btnErrCheck.Enabled = true;
                btnDataMake.Enabled = true;
                btnDel.Enabled = true;

                // データグリッドビュー編集
                dg1.ReadOnly = false;

                //エラー情報表示
                ErrShow();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (!dR.IsClosed) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                global.dg1ChabgeValueStatus = true;
            }
        }

        //表示初期化
        private void dataGridInitial(DataGridView dgv)
        {
            txtID.Text = string.Empty;
            txtYear.Text = string.Empty;
            txtMonth.Text = string.Empty;
            txtNo.Text = string.Empty;
            txtShozokuCode.Text = string.Empty;
            txtShokuID.Text = string.Empty;
            txtShuTl.Text = string.Empty;
            txtChuTl.Text = string.Empty;

            lblName.Text = string.Empty;
            lblShozoku.Text = string.Empty;
            lblShoteiS.Text = string.Empty;
            lblShoteiE.Text = string.Empty;
            lblHanAmS.Text = string.Empty;
            lblHanAmE.Text = string.Empty;
            lblHanPmS.Text = string.Empty;
            lblHanPmE.Text = string.Empty;
            lblTokushuS.Text = string.Empty;
            lblTokushuE.Text = string.Empty;
            lblTokAmS.Text = string.Empty;
            lblTokAmE.Text = string.Empty;
            lblTokPmS.Text = string.Empty;
            lblTokPmE.Text = string.Empty;

            txtID.BackColor = Color.Empty;
            txtYear.BackColor = Color.Empty;
            txtMonth.BackColor = Color.Empty;
            txtNo.BackColor = Color.Empty;
            txtShozokuCode.BackColor = Color.Empty;
            txtShokuID.BackColor = Color.Empty;
            txtShuTl.BackColor = Color.Empty;
            txtChuTl.BackColor = Color.Empty;
            txtYukyuTl.BackColor = Color.Empty;
            txttokuTl.BackColor = Color.Empty;
            txtKekkinTl.BackColor = Color.Empty;

            txtID.ForeColor = Color.Navy;
            txtYear.ForeColor = Color.Navy;
            txtMonth.ForeColor = Color.Navy;
            txtNo.ForeColor = Color.Navy;
            txtShozokuCode.ForeColor = Color.Navy;
            txtShokuID.ForeColor = Color.Navy;
            txtShuTl.ForeColor = Color.Navy;
            txtChuTl.ForeColor = Color.Navy;

            dgv.RowsDefaultCellStyle.ForeColor = Color.Navy;       //テキストカラーの設定
            dgv.DefaultCellStyle.SelectionBackColor = Color.Empty;
            dgv.DefaultCellStyle.SelectionForeColor = Color.Navy;

            //dgv.EditMode = EditMode.EditOnKeystrokeOrShortcutKey;

            //pictureBox1.Image = null;
            lblNoImage.Visible = false;
        }

        /// <summary>
        /// データ表示エリア背景色初期化
        /// </summary>
        /// <param name="dgv">データグリッドビューオブジェクト</param>
        private void dsColorInitial(DataGridView dgv)
        {
            // CellChangeValueイベント発生回避する
            global.dg1ChabgeValueStatus = false;

            txtID.BackColor = Color.White;
            txtYear.BackColor = Color.White;
            txtMonth.BackColor = Color.White;
            txtNo.BackColor = Color.White;
            txtShozokuCode.BackColor = Color.White;
            txtShokuID.BackColor = Color.White;
            txtShuTl.BackColor = Color.Empty;
            txtChuTl.BackColor = Color.Empty;
            txtYukyuTl.BackColor = Color.Empty;
            txttokuTl.BackColor = Color.Empty;
            txtKekkinTl.BackColor = Color.Empty;

            // 行数
            dgv.RowCount = 0;

            // 行追加、日付セット
            //SysControl.SetDBConnect db = new SysControl.SetDBConnect();
            //OleDbCommand sCom = new OleDbCommand();
            //sCom.Connection = db.cnOpen();
            //OleDbDataReader dr = null;

            for (int i = 0; i < _MULTIGYO; i++)
            {
                DateTime dt;
                if (DateTime.TryParse(global.sYear.ToString() + "/" + global.sMonth.ToString() + "/" + (i + 1).ToString(), out dt))
                {
                    // 行を追加
                    dgv.Rows.Add();
                    dgv.Rows[i].DefaultCellStyle.BackColor = Color.Empty;

                    // 日
                    dgv[cDay, i].Value = i + 1;
                    // 曜日
                    string Youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                    dgv[cWeek, i].Value = Youbi;

                    //// 土日の場合
                    //if (Youbi == "日" || Youbi == "土")
                    //{
                    //    dgv.Rows[i].DefaultCellStyle.BackColor = Color.MistyRose;
                    //}
                    //else
                    //{
                    //// 休日テーブルを参照し休日に該当するか調べます
                    //sCom.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                    //sCom.Parameters.Clear();
                    //sCom.Parameters.AddWithValue("@year", global.sYear);
                    //sCom.Parameters.AddWithValue("@Month", global.sMonth);
                    //sCom.Parameters.AddWithValue("@day", i + 1);
                    //dr = sCom.ExecuteReader();
                    //if (dr.HasRows)
                    //{
                    //    dgv.Rows[i].DefaultCellStyle.BackColor = Color.MistyRose;
                    //}
                    //dr.Close();
                    //}

                    // 時分区切り記号
                    dgv[cS, i].Value = ":";
                    dgv[cE, i].Value = ":";
                    dgv[cRS, i].Value = ":";
                    dgv[cRE, i].Value = ":";
                }
            }
            //sCom.Connection.Close();

            // CellChangeValueイベントステータスもどす
            global.dg1ChabgeValueStatus = true;
        }

        /// <summary>
        /// 伝票画像表示
        /// </summary>
        /// <param name="iX">現在の伝票</param>
        /// <param name="tempImgName">画像名</param>
        public void ShowImage(string tempImgName)
        {
            string wrkFileName;

            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                global.pblImageFile = string.Empty;
                return;
            }

            //画像ファイルがあるときのみ表示
            wrkFileName = tempImgName;
            if (System.IO.File.Exists(wrkFileName))
            {
                leadImg.Visible = true;

                //画像ロード
                RasterCodecs.Startup();
                RasterCodecs cs = new RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                RasterPaintProperties prop = new RasterPaintProperties();
                prop = RasterPaintProperties.Default;
                prop.PaintDisplayMode = RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(wrkFileName, 0, CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (global.miMdlZoomRate == 0f)
                {
                    leadImg.ScaleFactor *= global.ZOOM_RATE;
                }
                else
                {
                    leadImg.ScaleFactor *= global.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = RasterViewerInteractiveMode.Pan;

                ////右へ90°回転させる
                //RotateCommand rc = new RotateCommand();
                //rc.Angle = 90 * 100;
                //rc.FillColor = new RasterColor(255, 255, 255);
                ////rc.Flags = RotateCommandFlags.Bicubic;
                //rc.Flags = RotateCommandFlags.Resize;
                //rc.Run(leadImg.Image);

                // グレースケールに変換
                GrayscaleCommand grayScaleCommand = new GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                RasterCodecs.Shutdown();
                global.pblImageFile = wrkFileName;

                lblNoImage.Visible = false;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;
                btnFirst.Enabled = true;
                btnNext.Enabled = true;
                btnBefore.Enabled = true;
                btnEnd.Enabled = true;
            }
            else
            {
                //画像ファイルがないとき
                leadImg.Visible = false;
                global.pblImageFile = string.Empty;
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;
                btnFirst.Enabled = true;
                btnNext.Enabled = true;
                btnBefore.Enabled = true;
                btnEnd.Enabled = true;
            }
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = 0;
            DataShow(_cI, inData, dg1);
            txtDNum.Text = string.Empty;
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            if (_cI > 0)
            {
                _cI--;
                DataShow(_cI, inData, dg1);
                txtDNum.Text = string.Empty;
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            if (_cI + 1 < inData.Length)
            {
                _cI++;
                DataShow(_cI, inData, dg1);
                txtDNum.Text = string.Empty;
            }
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = inData.Length - 1;
            DataShow(_cI, inData, dg1);
            txtDNum.Text = string.Empty;
        }

        private void txtRiseki_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     カレントデータの更新 </summary>
        /// <param name="iX">
        ///     カレントレコードのインデックス</param>
        ///---------------------------------------------------------------------
        private void CurDataUpDate(int iX)
        {
            //カレントデータを更新する
            string mySql = string.Empty;

            // エラーメッセージ
            string errMsg = string.Empty;

            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();

            //勤務票ヘッダテーブル
            mySql += "update 勤務票ヘッダ set ";
            mySql += "会社ID=?,個人番号=?,氏名=?,年=?,月=?,職種ID=?,部門コード=?,";
            mySql += "出勤日数合計=?,昼食数合計=?,有休日数合計=?,特休日数合計=?,欠勤日数合計=?,";
            mySql += "所定開始時間=?,所定終了時間=?,午前半休開始時間=?,午前半休終了時間=?,";
            mySql += "午後半休開始時間=?,午後半休終了時間=?,所定勤務時間=?,";
            mySql += "特殊開始時間=?,特殊終了時間=?,特殊午前半休開始時間=?,特殊午前半休終了時間=?,";
            mySql += "特殊午後半休開始時間=?,特殊午後半休終了時間=?,特殊勤務時間=?,更新年月日=? ";
            mySql += "where ID = ?";

            errMsg = "勤務票ヘッダテーブル更新";

            sCom.CommandText = mySql;

            if (Utility.NumericCheck(txtID.Text))
                sCom.Parameters.AddWithValue("@SID", Utility.NulltoStr(txtID.Text).PadLeft(2,'0'));
            else sCom.Parameters.AddWithValue("@SID", Utility.NulltoStr(txtID.Text));
            
            sCom.Parameters.AddWithValue("@No", Utility.NulltoStr(txtNo.Text).PadLeft(7, '0'));
            sCom.Parameters.AddWithValue("@Name", Utility.NulltoStr(lblName.Text));
            sCom.Parameters.AddWithValue("@Year", "20" + Utility.NulltoStr(txtYear.Text));
            sCom.Parameters.AddWithValue("@Month", Utility.NulltoStr(txtMonth.Text));

            if (Utility.NumericCheck(txtShokuID.Text))
                sCom.Parameters.AddWithValue("@ShokuID", Utility.NulltoStr(txtShokuID.Text).PadLeft(2, '0'));
            else sCom.Parameters.AddWithValue("@ShokuID", Utility.NulltoStr(txtShokuID.Text));
                
            sCom.Parameters.AddWithValue("@Shozoku", Utility.NulltoStr(txtShozokuCode.Text));
            sCom.Parameters.AddWithValue("@ShuTl", Utility.NulltoStr(txtShuTl.Text));
            sCom.Parameters.AddWithValue("@ChuTl", Utility.NulltoStr(txtChuTl.Text));

            //sCom.Parameters.AddWithValue("@YuTl", Utility.NulltoStr(txtYukyuTl.Text));


            // 有給取得日数   2012/06/21
            string sYu = Utility.NulltoStr(txtYukyuTl.Text);
            string[] sRi;
            if (txtYukyuTl.Text.Contains('.'))
            {
                // 有給取得日数ピリオド区切りで分割して配列に格納する
                sRi = txtYukyuTl.Text.Split('.');
                string sSplit = "."; 
                if (sRi[0] == string.Empty) sRi[0] = "0";           // 整数部が値なしなら0をセットします
                if (sRi[1] == string.Empty) sSplit = string.Empty;  // 少数部が値なしなら小数点を消去します
                sCom.Parameters.AddWithValue("@YuTl", double.Parse(sRi[0] + sSplit + sRi[1]).ToString());
            }
            else
            {
                sCom.Parameters.AddWithValue("@YuTl", Utility.NulltoStr(txtYukyuTl.Text));
            }

            
            sCom.Parameters.AddWithValue("@TokTl", Utility.NulltoStr(txttokuTl.Text));
            sCom.Parameters.AddWithValue("@KekTl", Utility.NulltoStr(txtKekkinTl.Text));

            sCom.Parameters.AddWithValue("@ShoS", lblShoteiS.Text);
            sCom.Parameters.AddWithValue("@ShoE", lblShoteiE.Text);
            sCom.Parameters.AddWithValue("@ShoAmS", lblHanAmS.Text);
            sCom.Parameters.AddWithValue("@ShoAmE", lblHanAmE.Text);
            sCom.Parameters.AddWithValue("@ShoPmS", lblHanPmS.Text);
            sCom.Parameters.AddWithValue("@ShoPmE", lblHanPmE.Text);
            sCom.Parameters.AddWithValue("@ShoTime", inData[_cI]._ShoTime);
            sCom.Parameters.AddWithValue("@ShoS2", lblTokushuS.Text);
            sCom.Parameters.AddWithValue("@ShoE2", lblTokushuE.Text);
            sCom.Parameters.AddWithValue("@ShoAmS2", lblTokAmS.Text);
            sCom.Parameters.AddWithValue("@ShoAmE2", lblTokAmE.Text);
            sCom.Parameters.AddWithValue("@ShoPmS2", lblTokPmS.Text);
            sCom.Parameters.AddWithValue("@ShoPmE2", lblTokPmE.Text);
            sCom.Parameters.AddWithValue("@ShoTime2", inData[_cI]._TokTime);
            sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());

            sCom.Parameters.AddWithValue("@ID", inData[iX]._sID);

            sCom.Connection = dCon.cnOpen();

            // トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                sCom.ExecuteNonQuery();

                // 勤務票明細テーブル
                mySql = string.Empty;
                mySql += "update 勤務票明細 set ";
                mySql += "特殊日=?,所定勤務=?,開始時=?,開始分=?,終了時=?,終了分=?,昼食=?,休暇=?,休憩なし=?,";
                mySql += "離席開始時=?,離席開始分=?,離席終了時=?,離席終了分=?,訂正=?,休日=?,更新年月日=? ";
                mySql += "where ID = ?";
                errMsg = "勤務票明細テーブル更新";
                sCom.CommandText = mySql;

                for (int i = 0; i < dg1.Rows.Count; i++)
                {
                    sCom.Parameters.Clear();

                    // 特殊日マーク
                    if (dg1[cTokMK, i].Value.ToString() == global.MARU)
                        sCom.Parameters.AddWithValue("@mark", global.FLGON);
                    else sCom.Parameters.AddWithValue("@mark", global.FLGOFF);

                    // 所定勤務マーク
                    if (dg1[cShoMK, i].Value.ToString() == global.MARU)
                        sCom.Parameters.AddWithValue("@Smark", global.FLGON);
                    else sCom.Parameters.AddWithValue("@Smark", global.FLGOFF);

                    // 開始時
                    sCom.Parameters.AddWithValue("@SH", Utility.NulltoStr(dg1[cSH, i].Value).ToString().Trim());
                    // 開始分
                    sCom.Parameters.AddWithValue("@SM", Utility.NulltoStr(dg1[cSM, i].Value).ToString().Trim());
                    // 終了時
                    sCom.Parameters.AddWithValue("@EH", Utility.NulltoStr(dg1[cEH, i].Value).ToString().Trim());
                    // 終了分
                    sCom.Parameters.AddWithValue("@EM", Utility.NulltoStr(dg1[cEM, i].Value).ToString().Trim());

                    // 昼食
                    if (dg1[cHiru1, i].Value.ToString() == global.MARU)
                        sCom.Parameters.AddWithValue("@Hiru1", global.FLGON);
                    else sCom.Parameters.AddWithValue("@Hiru1", global.FLGOFF);

                    // 休暇
                    sCom.Parameters.AddWithValue("@kyuka", Utility.NulltoStr(dg1[cKyuka, i].Value).ToString().Trim());

                    // 休憩なし
                    if (dg1[cKyukei, i].Value.ToString() == global.MARU)
                        sCom.Parameters.AddWithValue("@Kyukei", global.FLGON);
                    else sCom.Parameters.AddWithValue("@Kyukei", global.FLGOFF);

                    // 離席開始時
                    sCom.Parameters.AddWithValue("@RSH", Utility.NulltoStr(dg1[cRSH, i].Value).ToString().Trim());
                    // 離席開始分
                    sCom.Parameters.AddWithValue("@RSM", Utility.NulltoStr(dg1[cRSM, i].Value).ToString().Trim());
                    // 離席終了時
                    sCom.Parameters.AddWithValue("@REH", Utility.NulltoStr(dg1[cREH, i].Value).ToString().Trim());
                    // 離席終了分
                    sCom.Parameters.AddWithValue("@REM", Utility.NulltoStr(dg1[cREM, i].Value).ToString().Trim());

                    // 訂正
                    if (dg1[cTeisei, i].Value.ToString() == "True")
                        sCom.Parameters.AddWithValue("@Teisei", global.FLGON);
                    else sCom.Parameters.AddWithValue("@Teisei", global.FLGOFF);

                    // 休日
                    sCom.Parameters.AddWithValue("@Hol", inData[iX]._Meisai[i]._clr);

                    // 更新年月日
                    sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());

                    // ID
                    sCom.Parameters.AddWithValue("@ID", Utility.NulltoStr(dg1[cID, i].Value).ToString());

                    // テーブル書き込み
                    sCom.ExecuteNonQuery();
                }

                //トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, errMsg, MessageBoxButtons.OK);

                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!Utility.NumericCheck(txtDNum.Text))
            {
                MessageBox.Show("何件目のデータへ移動するか指定してください", "データ移動", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                return;
            }

            if (int.Parse(txtDNum.Text) > inData.Length)
            {
                MessageBox.Show("データ件数（" + inData.Length.ToString() + "件）の範囲を超えています", "データ移動", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                return;
            }

            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = int.Parse((txtDNum.Text).ToString()) - 1;
            DataShow(_cI, inData, dg1);
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = hScrollBar1.Value;
            DataShow(_cI, inData, dg1);
            txtDNum.Text = string.Empty;
        }

        private void dg1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                string ColName = dg1.Columns[dg1.CurrentCell.ColumnIndex].Name;

                if (ColName == cSH || ColName == cSM || ColName == cEH || ColName == cEM || ColName == cKyuka ||
                    ColName == cRSH || ColName == cRSM || ColName == cREH || ColName == cREM)
                {
                    //イベントハンドラが複数回追加されてしまうので最初に削除する
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                    e.Control.KeyDown -= new KeyEventHandler(Control_KeyDown2);
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }

                if (ColName == cTokMK || ColName == cKyukei || ColName == cHiru1 || ColName == cShoMK)
                {
                    //イベントハンドラが複数回追加されてしまうので最初に削除する
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                    e.Control.KeyDown -= new KeyEventHandler(Control_KeyDown2);
                    //イベントハンドラを追加する
                    e.Control.KeyDown += new KeyEventHandler(Control_KeyDown2);
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress2);
                }
            }
        }

        void Control_KeyDown2(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Space && e.KeyCode != Keys.Delete && e.KeyCode != Keys.Tab && e.KeyCode != Keys.Enter)
            {
                e.Handled = true;
                return;
            }
        }

        void Control_KeyPress2(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != (char)Keys.Space && e.KeyChar != '\b' && e.KeyChar != (char)Keys.Delete && e.KeyChar != (char)Keys.Tab && e.KeyChar != (char)Keys.Enter)
            {
                e.Handled = true;
                return;
            }
        }

        void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        private void dg1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (global.dg1ChabgeValueStatus == false) return;
            global.dg1ChabgeValueStatus = false; // 自らChangeValueイベントを発生させない

            if (_cI == null) return;
            if (dg1.CurrentRow == null) return;

            // 該当カラム
            string Col = dg1.Columns[e.ColumnIndex].Name;

            // 該当行
            int Rn = e.RowIndex;

            // 以下のカラムは対象外
            if (Col == cWeek || Col == cTokMK || Col == cShoMK || Col == cS || Col == cE || 
                Col == cHiru1 || Col == cKyukei || Col == cRS || Col == cRE)
            {
                global.dg1ChabgeValueStatus = true;
                return;
            }

            global.lBackColorE = Color.FromArgb(251, 228, 183);
            //global.lBackColorN = Color.FromArgb(255, 255, 255);
            global.lBackColorK = Color.FromArgb(254, 252, 171);
            global.lBackColorR = Color.Blue;

            // 休暇
            if (Col == cKyuka)
            {
                if (dg1[Col, Rn].Value != null)
                {
                    switch (dg1[Col, Rn].Value.ToString())
                    {
                        case "4":
                            inData[_cI]._Meisai[Rn]._clr = 2;
                            break;

                        case "3":
                            if (inData[_cI]._Meisai[Rn]._clr == 1 || inData[_cI]._Meisai[Rn]._clr == 2)
                            {
                                inData[_cI]._Meisai[Rn]._clr = 0;
                                dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                                dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                                dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                            }
                            break;

                        default:
                            if (dg1[cWeek, Rn].Value.ToString() == "土")
                            {
                                inData[_cI]._Meisai[Rn]._clr = 1;
                            }
                            else if (dg1[cWeek, Rn].Value.ToString() == "日")
                            {
                                inData[_cI]._Meisai[Rn]._clr = 2;
                            }
                            else
                            {
                                inData[_cI]._Meisai[Rn]._clr = 0;
                            }

                            // 休日テーブルを参照し休日に該当するか調べます
                            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
                            OleDbCommand sCom = new OleDbCommand();
                            OleDbDataReader dr = null;
                            sCom.Connection = Con.cnOpen();
                            sCom.CommandText = "select * from 休日 where 年=? and 月=? and 日=? and 会社ID=?";
                            sCom.Parameters.Clear();
                            sCom.Parameters.AddWithValue("@year", global.sYear);
                            sCom.Parameters.AddWithValue("@Month", global.sMonth);
                            sCom.Parameters.AddWithValue("@day", Rn + 1);
                            sCom.Parameters.AddWithValue("@id", global.pblComNo.ToString("00"));
                            dr = sCom.ExecuteReader();
                            if (dr.Read())
                            {
                                inData[_cI]._Meisai[Rn]._clr = 2;
                            }
                            dr.Close();
                            sCom.Connection.Close();

                            break;
                    }

                    switch (inData[_cI]._Meisai[Rn]._clr)
                    {
                        case 1:
                            global.lBackColorN = Color.FromArgb(225, 244, 255);
                            break;
                        case 2:
                            global.lBackColorN = Color.FromArgb(255, 223, 227);
                            break;
                        default:
                            //global.lBackColorN = Color.FromArgb(255, 255, 255);
                            if (dg1[cTeisei, Rn].Value.ToString() == "False")
                                global.lBackColorN = Color.FromArgb(255, 255, 255);
                            else global.lBackColorN = global.lBackColorE;
                            break;
                    }

                    // 該当日付が休日のとき
                    if (inData[_cI]._Meisai[Rn]._clr == 1 || inData[_cI]._Meisai[Rn]._clr == 2)
                    {
                        //dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorN;
                        dg1[cDay, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cWeek, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cTokMK, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cShoMK, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cS, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cE, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cKyuka, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cKyukei, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cHiru1, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cRSH, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cRS, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cRSM, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cREH, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cRE, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cREM, Rn].Style.BackColor = global.lBackColorN;
                        dg1[cTeisei, Rn].Style.BackColor = global.lBackColorN;
                    }
                }
            }

            switch (inData[_cI]._Meisai[Rn]._clr)
            {
                case 1:
                    global.lBackColorN = Color.FromArgb(225, 244, 255);
                    break;
                case 2:
                    global.lBackColorN = Color.FromArgb(255, 223, 227);
                    break;
                default:
                    if (dg1[cTeisei, Rn].Value.ToString() == "False")
                        global.lBackColorN = Color.FromArgb(255, 255, 255);
                    else global.lBackColorN = global.lBackColorE;
                    break;
            }

            //// 訂正チェック
            //if (dg1[cTeisei, Rn].Value != null)
            //{
            //    if (dg1[cTeisei, Rn].Value.ToString() == "True")
            //    {
            //        global.lBackColorE = Color.FromArgb(251, 228, 183);
            //        global.lBackColorN = Color.FromArgb(251, 228, 183);
            //    }

            //    // 行表示色
            //    TeiseiColor(Rn);
            //}
            //else
            //{
            //    dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorN;
            //}


            // 開始時間
            // 訂正行でないとき表示色をリセットします
            //if ((Col == cSH || Col == cSM) && dg1[cTeisei, Rn].Value.ToString() == "False")
            if (Col == cSH || Col == cSM)
            {
                dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cS, Rn].Style.BackColor = global.lBackColorN;
                dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
            }

            if ((Col == cSH || Col == cSM) && (dg1[cSH, Rn].Value != null && dg1[cSM, Rn].Value != null))
            {
                // 開始時
                if (Utility.NumericCheck(dg1[cSH, Rn].Value.ToString()))
                    dg1[cSH, Rn].Value = dg1[cSH, Rn].Value.ToString().PadLeft(2, '0');

                if (Utility.StrToInt(dg1[cSH, Rn].Value.ToString()) < 8 && Utility.StrToInt(dg1[cSH, Rn].Value.ToString()) != 0)
                    dg1[cSH, Rn].Style.BackColor = global.lBackColorE;

                // 開始分
                if (Utility.NumericCheck(dg1[cSM, Rn].Value.ToString()))
                {
                    dg1[cSM, Rn].Value = dg1[cSM, Rn].Value.ToString().PadLeft(2, '0');

                    if (dg1[cSM, Rn].Value.ToString().Substring(1, 1) != "5" &&
                        dg1[cSM, Rn].Value.ToString().Substring(1, 1) != "0")
                    dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                }

                int hm = 0;

                // 時間が記入されているとき所定開始時間と比較する
                if (dg1[cSH, Rn].Value.ToString() != string.Empty || dg1[cSM, Rn].Value.ToString() != string.Empty)
                {
                    // 開始時間を取得
                    hm = Utility.StrToInt(dg1[cSH, Rn].Value.ToString()) * 100 +
                        Utility.StrToInt(dg1[cSM, Rn].Value.ToString());

                    // 開始が所定より早い場合
                    if (Utility.StrToInt(global.ShoS.Replace(":", string.Empty)) > hm)
                    {
                        dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                    }

                    // 午後半休開始が所定より早い場合
                    if (dg1[cKyuka, Rn].Value != null)  // 2015/08/18
                    {
                        if (dg1[cKyuka, Rn].Value.ToString() == "5" && (Utility.StrToInt(global.PmS.Replace(":", string.Empty)) > hm))
                        {
                            dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    // 開始が所定より30分以上遅い場合
                    DateTime st;
                    if (DateTime.TryParse(global.ShoS, out st))
                    {
                        if (Utility.StrToInt(st.AddMinutes(30).ToShortTimeString().Replace(":", string.Empty)) <= hm)
                        {
                            dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    // 休日出勤してる場合
                    if (inData[_cI]._Meisai[Rn]._clr == 1 || inData[_cI]._Meisai[Rn]._clr == 2)
                        //inData[_cI]._Meisai[Rn]._clr == 3)
                    {
                        if (dg1[cSH, Rn].Value.ToString() != string.Empty || dg1[cSM, Rn].Value.ToString() != string.Empty)
                        {
                            dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                        }
                        else
                        {
                            //dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                            //dg1[cS, Rn].Style.BackColor = global.lBackColorN;
                            //dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
                        }
                    }

                    //////// 半休で時間が掛かってる場合
                    //////if (dg1[cKyuka, Rn].Value.ToString() == "5")
                    //////{
                    //////    if (Utility.StrToInt(global.AmS.Replace(":", string.Empty)) <= hm &&
                    //////        Utility.StrToInt(global.AmE.Replace(":", string.Empty)) >= hm)
                    //////    {
                    //////        dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                    //////        dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                    //////        dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                    //////    }
                    //////}

                    //////if (dg1[cKyuka, Rn].Value.ToString() == "6")
                    //////{
                    //////    if (Utility.StrToInt(global.PmS.Replace(":", string.Empty)) <= hm &&
                    //////        Utility.StrToInt(global.PmE.Replace(":", string.Empty)) >= hm)
                    //////    {
                    //////        dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                    //////        dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                    //////        dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                    //////    }
                    //////}
                }
            }

            // 訂正行でないとき終了時間をリセットします
            //if ((Col == cEH || Col == cEM) && dg1[cTeisei, Rn].Value.ToString() == "False")
            if (Col == cEH || Col == cEM)
            {
                dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cE, Rn].Style.BackColor = global.lBackColorN;
                dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
            }

            if ((Col == cEH || Col == cEM) && (dg1[cEH, Rn].Value != null && dg1[cEM, Rn].Value != null))
            {
                if (Utility.NumericCheck(dg1[cEH, Rn].Value.ToString()))
                    dg1[cEH, Rn].Value = dg1[cEH, Rn].Value.ToString().PadLeft(2, '0');

                if (Utility.StrToInt(dg1[cEH, Rn].Value.ToString()) < 8 && Utility.StrToInt(dg1[cEH, Rn].Value.ToString()) >= 22)
                    dg1[cEH, Rn].Style.BackColor = global.lBackColorE;

                // 終了分
                if (Utility.NumericCheck(dg1[cEM, Rn].Value.ToString()))
                {
                    dg1[cEM, Rn].Value = dg1[cEM, Rn].Value.ToString().PadLeft(2, '0');

                    if (dg1[cEM, Rn].Value.ToString().Substring(1, 1) != "5" &&
                        dg1[cEM, Rn].Value.ToString().Substring(1, 1) != "0")
                    dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                }

                int hm = 0;
                
                // 時間が記入されているとき所定開始時間と比較する
                if (dg1[cEH, Rn].Value.ToString() != string.Empty || dg1[cEM, Rn].Value.ToString() != string.Empty)
                {
                    // 終了時間を取得
                    hm = Utility.StrToInt(dg1[cEH, Rn].Value.ToString()) * 100 +
                        Utility.StrToInt(dg1[cEM, Rn].Value.ToString());

                    // 終了が所定より早い場合
                    if (Utility.StrToInt(global.ShoE.Replace(":", string.Empty)) > hm)
                    {
                        dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                    }

                    // 午後半休終了が所定より早い場合
                    if (dg1[cKyuka, Rn].Value != null)  // 2015/08/18
                    {
                        if (dg1[cKyuka, Rn].Value.ToString() == "6" && (Utility.StrToInt(global.PmE.Replace(":", string.Empty)) > hm))
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    // 終了が所定より120分以上遅い場合
                    DateTime et;
                    if (DateTime.TryParse(global.ShoE, out et))
                    {
                        if (Utility.StrToInt(et.AddMinutes(120).ToShortTimeString().Replace(":", string.Empty)) <= hm)
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    // 所定が15時以前でかつ終了が16時以降の場合
                    if (Utility.StrToInt(global.ShoE.Replace(":", string.Empty)) < 1600)
                    {
                        if (hm >= 1600)
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    // 休日出勤してる場合
                    if (inData[_cI]._Meisai[Rn]._clr == 1 || inData[_cI]._Meisai[Rn]._clr == 2)
                    //inData[_cI]._Meisai[Rn]._clr == 3)
                    {
                        if (dg1[cEH, Rn].Value.ToString() != string.Empty || dg1[cEM, Rn].Value.ToString() != string.Empty)
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                        }
                        else
                        {
                            //dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                            //dg1[cE, Rn].Style.BackColor = global.lBackColorN;
                            //dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
                        }
                    }

                    //////// 半休で時間が掛かってる場合
                    //////if (dg1[cKyuka, Rn].Value.ToString() == "5")
                    //////{
                    //////    if (Utility.StrToInt(global.AmS.Replace(":", string.Empty)) <= hm &&
                    //////        Utility.StrToInt(global.AmE.Replace(":", string.Empty)) >= hm)
                    //////    {
                    //////        dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                    //////        dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                    //////        dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                    //////    }
                    //////}

                    //////if (dg1[cKyuka, Rn].Value.ToString() == "6")
                    //////{
                    //////    if (Utility.StrToInt(global.PmS.Replace(":", string.Empty)) <= hm &&
                    //////        Utility.StrToInt(global.PmE.Replace(":", string.Empty)) >= hm)
                    //////    {
                    //////        dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                    //////        dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                    //////        dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                    //////    }
                    //////}
                }

            }

            // 離席開始時
            if (Col == cRSH && dg1[cRSH, Rn].Value != null)
            {
                if (dg1[cRSH, Rn].Value.ToString() != string.Empty)
                {
                    if (Utility.NumericCheck(dg1[cRSH, Rn].Value.ToString()))
                        dg1[cRSH, Rn].Value = dg1[cRSH, Rn].Value.ToString().PadLeft(2, '0');
                    dg1[cDay, Rn].Style.BackColor = global.lBackColorR;
                }
                else dg1[cDay, Rn].Style.BackColor = global.lBackColorN;
            }

            // 離席開始分
            if (Col == cRSM && dg1[cRSM, Rn].Value != null)
            {
                if (dg1[cRSM, Rn].Value.ToString() != string.Empty)
                {
                    if (Utility.NumericCheck(dg1[cRSM, Rn].Value.ToString()))
                        dg1[cRSM, Rn].Value = dg1[cRSM, Rn].Value.ToString().PadLeft(2, '0');
                    dg1[cDay, Rn].Style.BackColor = global.lBackColorR;
                }
                else dg1[cDay, Rn].Style.BackColor = global.lBackColorN;
            }

            // 離席終了時
            if (Col == cREH && dg1[cREH, Rn].Value != null)
            {
                if (dg1[cREH, Rn].Value.ToString() != string.Empty)
                {
                    if (Utility.NumericCheck(dg1[cREH, Rn].Value.ToString()))
                        dg1[cREH, Rn].Value = dg1[cREH, Rn].Value.ToString().PadLeft(2, '0');
                    dg1[cDay, Rn].Style.BackColor = global.lBackColorR;
                }
                else dg1[cDay, Rn].Style.BackColor = global.lBackColorN;
            }

            // 離席終了分
            if (Col == cREM && dg1[cREM, Rn].Value != null)
            {
                if (dg1[cREM, Rn].Value.ToString() != string.Empty)
                {
                    if (Utility.NumericCheck(dg1[cREM, Rn].Value.ToString()))
                        dg1[cREM, Rn].Value = dg1[cREM, Rn].Value.ToString().PadLeft(2, '0');
                    dg1[cDay, Rn].Style.BackColor = global.lBackColorR;
                }
                else dg1[cDay, Rn].Style.BackColor = global.lBackColorN;
            }
            
            // 訂正チェック
            if (Col == cTeisei)
            {
                //if (inData[_cI]._Meisai[Rn]._clr == 1)
                //    global.lBackColorN = Color.FromArgb(225, 244, 255);
                //else if (inData[_cI]._Meisai[Rn]._clr == 2)
                //    global.lBackColorN = Color.FromArgb(255, 223, 227);
                //else
                //    global.lBackColorN = Color.FromArgb(255, 255, 255);

                // 行表示色
                TeiseiColor(Rn);
            }

            // ChangeValueイベントステータスを戻す
            global.dg1ChabgeValueStatus = true;
        }

        private void TeiseiColor(int Rn)
        {

            if (dg1[cTeisei, Rn].Value.ToString() == "True")
            {
                //dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorE;
                dg1[cDay, Rn].Style.BackColor = global.lBackColorE;
                dg1[cWeek, Rn].Style.BackColor = global.lBackColorE;
                dg1[cTokMK, Rn].Style.BackColor = global.lBackColorE;
                dg1[cShoMK, Rn].Style.BackColor = global.lBackColorE;
                dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                dg1[cKyuka, Rn].Style.BackColor = global.lBackColorK;
                dg1[cKyukei, Rn].Style.BackColor = global.lBackColorE;
                dg1[cHiru1, Rn].Style.BackColor = global.lBackColorE;
                dg1[cRSH, Rn].Style.BackColor = global.lBackColorE;
                dg1[cRS, Rn].Style.BackColor = global.lBackColorE;
                dg1[cRSM, Rn].Style.BackColor = global.lBackColorE;
                dg1[cREH, Rn].Style.BackColor = global.lBackColorE;
                dg1[cRE, Rn].Style.BackColor = global.lBackColorE;
                dg1[cREM, Rn].Style.BackColor = global.lBackColorE;
                dg1[cTeisei, Rn].Style.BackColor = global.lBackColorE;
            }
            else
            {
                //dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorN;
                dg1[cDay, Rn].Style.BackColor = global.lBackColorN;
                dg1[cWeek, Rn].Style.BackColor = global.lBackColorN;
                dg1[cTokMK, Rn].Style.BackColor = global.lBackColorN;
                dg1[cShoMK, Rn].Style.BackColor = global.lBackColorN;
                dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cS, Rn].Style.BackColor = global.lBackColorN;
                dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
                dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cE, Rn].Style.BackColor = global.lBackColorN;
                dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
                dg1[cKyuka, Rn].Style.BackColor = global.lBackColorK;
                dg1[cKyukei, Rn].Style.BackColor = global.lBackColorN;
                dg1[cHiru1, Rn].Style.BackColor = global.lBackColorN;
                dg1[cRSH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cRS, Rn].Style.BackColor = global.lBackColorN;
                dg1[cRSM, Rn].Style.BackColor = global.lBackColorN;
                dg1[cREH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cRE, Rn].Style.BackColor = global.lBackColorN;
                dg1[cREM, Rn].Style.BackColor = global.lBackColorN;
                dg1[cTeisei, Rn].Style.BackColor = global.lBackColorN;
            }
        }

        private void dg1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = dg1.Columns[dg1.CurrentCell.ColumnIndex].Name;
            if (colName == cTokMK || colName == cShoMK || colName == cKyukei || colName == cHiru1 || colName == cTeisei)
            {
                if (dg1.IsCurrentCellDirty)
                {
                    dg1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    dg1.RefreshEdit();
                }
            }
        }

        private void dg1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string colName = dg1.Columns[e.ColumnIndex].Name;
            if (colName == cTokMK || colName == cShoMK || colName == cHiru1 || colName == cKyukei)
            {
                if (dg1[colName, dg1.CurrentRow.Index].Value == null)
                {
                    dg1[colName, dg1.CurrentRow.Index].Value = global.MARU;
                }
                else if (dg1[colName, dg1.CurrentRow.Index].Value.ToString() == global.MARU)
                {
                    dg1[colName, dg1.CurrentRow.Index].Value = string.Empty;
                }
                else
                {
                    dg1[colName, dg1.CurrentRow.Index].Value = global.MARU;
                }

                dg1.RefreshEdit();

                // ChangeValueイベントステータスを戻す
                global.dg1ChabgeValueStatus = true; 
            }
        }

        private void dg1_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            // 入力値がSpaceのときに有効とします
            if (e.Value.ToString().Trim() != string.Empty) return;

            DataGridViewEx dgv = (DataGridViewEx)sender;
            string colName = dg1.Columns[e.ColumnIndex].Name;

            //セルの列を調べる

            if (colName == cTokMK || colName == cShoMK || colName == cHiru1 || colName == cKyukei)
            {
                if (dgv[colName, dgv.CurrentRow.Index].Value == null ||
                    dgv[colName, dgv.CurrentRow.Index].Value.ToString() != global.MARU)
                {
                    //○をセルの値とする
                    e.Value = global.MARU;
                }
                else if (dgv[colName, dgv.CurrentRow.Index].Value.ToString() == global.MARU)
                {
                    //セルをEmptyとする
                    e.Value = string.Empty;
                }

                //解析が不要であることを知らせる
                e.ParsingApplied = true;

                // ChangeValueイベントステータスを戻す
                global.dg1ChabgeValueStatus = true; 
            }
        }

        private void dg1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            string ColH = string.Empty;
            string ColM = dg1.Columns[dg1.CurrentCell.ColumnIndex].Name;

            // 開始時間または終了時間を判断
            if (ColM == cSM)
            {
                ColH = cSH;
            }
            else if (ColM == cEM)
            {
                ColH = cEH;
            }
            else if (ColM == cRSM)
            {
                ColH = cRSH;
            }
            else if (ColM == cREM)
            {
                ColH = cREH;
            }
            else
            {
                return;
            }

            // 開始時、終了時が入力済みで開始分、終了分が未入力のとき"00"を表示します
            if (dg1[ColH, dg1.CurrentRow.Index].Value != null)
            {
                if (dg1[ColH, dg1.CurrentRow.Index].Value.ToString().Trim() != string.Empty)
                {
                    if (dg1[ColM, dg1.CurrentRow.Index].Value == null)
                    {
                        dg1[ColM, dg1.CurrentRow.Index].Value = "00";
                    }
                    else if (dg1[ColM, dg1.CurrentRow.Index].Value.ToString().Trim() == string.Empty)
                    {
                        dg1[ColM, dg1.CurrentRow.Index].Value = "00";
                    }
                }
            }
        }

        private void btnErrCheck_Click(object sender, EventArgs e)
        {
            // カレントレコード更新
            CurDataUpDate(_cI);

            // エラーチェック実行①:カレントレコードから最終レコードまで
            if (ErrCheckMain(inData[_cI]._sID, inData[inData.Length - 1]._sID) == false) return;

            // エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (_cI > 0)
            {
                if (ErrCheckMain(inData[0]._sID, inData[_cI - 1]._sID) == false) return;
            }

            // エラーなし
            DataShow(_cI, inData, dg1);
            MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dg1.CurrentCell = null;
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     エラーチェックメイン処理 </summary>
        /// <param name="sID">
        ///     開始ID</param>
        /// <param name="eID">
        ///     終了ID</param>
        /// <returns>
        ///     True:エラーなし、false:エラーあり</returns>
        ///--------------------------------------------------------------------
        private Boolean ErrCheckMain(string sIx, string eIx)
        {
            int rCnt = 0;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            //レコード件数取得
            int cTotal = CountMDB();

            //エラー情報初期化
            ErrInitial();

            // 勤務記録表データ読み出し
            Boolean eCheck = true;
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            string mySql = string.Empty;

            mySql += "select * from 勤務票ヘッダ order by ID";

            sCom.CommandText = mySql;
            sCom.Connection = dCon.cnOpen();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                //指定範囲のIDならエラーチェックを実施する
                if (Int64.Parse(dR["ID"].ToString()) >= Int64.Parse(sIx) && Int64.Parse(dR["ID"].ToString()) <= Int64.Parse(eIx))
                {
                    eCheck = ErrCheckData(dR);
                    if (eCheck == false) break;　//エラーがあったとき
                }
            }

            dR.Close();
            sCom.Connection.Close();

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            //エラー有りの処理
            if (eCheck == false)
            {
                //エラーデータのインデックスを取得
                for (int i = 0; i < inData.Length; i++)
                {
                    if (inData[i]._sID == global.errID)
                    {
                        //エラーデータを画面表示
                        _cI = i;
                        DataShow(_cI, inData, dg1);
                        break;
                    }
                }
            }

            return eCheck;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     項目別エラーチェック </summary>
        /// <param name="cdR">
        ///     データリーダー</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///-------------------------------------------------------------------
        private Boolean ErrCheckData(OleDbDataReader cdR)
        {
            // 昼食回数
            int Lunchs = 0;

            // 特殊日回数
            int Toks = 0;
            
            // 所定勤務日回数
            int Shos = 0;
            
            // 出勤日数
            int rDays = 0;

            // 欠勤日数
            global.KekkinCnt = 0;

            // 有休日数
            double Yukyus = 0;

            // 特休日数
            int Tokkyus = 0;

            //対象年
            if (cdR["年"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "年を入力してください";

                return false;
            }

            if (Utility.NumericCheck(cdR["年"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "数字を入力してください";

                return false;
            }

            if (global.sYear != int.Parse(cdR["年"].ToString()))
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "処理年と異なっています。 処理年月：" + global.sYear.ToString() + "年 " + global.sMonth.ToString() + "月";

                return false;
            }

            //対象月
            if (cdR["月"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "月を入力してください";

                return false;
            }

            if (Utility.NumericCheck(cdR["月"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "数字を入力してください";

                return false;
            }

            if (int.Parse(cdR["月"].ToString()) < 1 || int.Parse(cdR["月"].ToString()) > 12)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "正しい月を入力してください";

                return false;
            }

            if (global.sMonth != int.Parse(cdR["月"].ToString()))
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "処理月と異なっています。 処理年月：" + global.sYear.ToString() + "年 " + global.sMonth.ToString() + "月";

                return false;
            }

            //個人番号
            //未入力のとき
            if (cdR["個人番号"] == null)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号を入力してください";

                return false;
            }

            if (cdR["個人番号"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号を入力してください";

                return false;
            }

            //数字以外のとき
            if (Utility.NumericCheck(cdR["個人番号"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号が不正です。";

                return false;
            }

            //個人番号マスター情報検査
            bool rHas = false;
            for (int i = 0; i < mst.Length; i++)
            {
                if (mst[i].sCode == cdR["個人番号"].ToString())
                {
                    rHas = true;
                    break;
                }
            }

            if (!rHas)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号がマスターに存在しません";
                return false;
            }
            
            // 会社ID
            // 未入力のとき
            if (cdR["会社ID"] == null)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eID;
                global.errRow = 0;
                global.errMsg = "会社IDを入力してください";

                return false;
            }

            if (cdR["会社ID"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eID;
                global.errRow = 0;
                global.errMsg = "会社IDを入力してください";

                return false;
            }

            //数字以外のとき
            if (Utility.NumericCheck(cdR["会社ID"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eID;
                global.errRow = 0;
                global.errMsg = "会社IDは数字を入力してください";

                return false;
            }

            //該当しないとき
            if (int.Parse(cdR["会社ID"].ToString()) != global.pblComNo)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eID;
                global.errRow = 0;
                global.errMsg = "会社IDが不正です";

                return false;
            }
            
            // 職種ID
            // 未入力のとき
            if (cdR["職種ID"] == null)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eSID;
                global.errRow = 0;
                global.errMsg = "職種IDを入力してください";

                return false;
            }

            if (cdR["職種ID"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eSID;
                global.errRow = 0;
                global.errMsg = "職種IDを入力してください";

                return false;
            }

            //数字以外のとき
            if (Utility.NumericCheck(cdR["職種ID"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eSID;
                global.errRow = 0;
                global.errMsg = "職種IDは数字を入力してください";

                return false;
            }

            // 所定勤務時間未取得のとき
            if (cdR["職種ID"].ToString() != "01" && cdR["職種ID"].ToString() != "02" && cdR["所定開始時間"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "給与奉行で所定開始、終了時間が登録されていません";

                return false;
            }

            ////該当しないとき
            //if (int.Parse(cdR["会社ID"].ToString()) != global.pblComNo)
            //{
            //    global.errID = cdR["ID"].ToString();
            //    global.errNumber = global.eID;
            //    global.errRow = 0;
            //    global.errMsg = "会社IDが不正です";

            //    return false;
            //}

            // 勤務票明細データ
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            sCom.Connection = dCon.cnOpen();
            sCom.CommandText = "select * from 勤務票明細 where ヘッダID = ? order by ID";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@HID", cdR["ID"].ToString());
            dR = sCom.ExecuteReader();

            // 日付別データ
            int iX = 1;
            int sID = 0;

            while (dR.Read())
            {
                // 職種IDを取得します
                sID = int.Parse(cdR["職種ID"].ToString());

                // 正社員、パートタイマーのとき明細行チェックを実施します
                if (sID == 5 || sID == 6 || sID == 7)
                {
                    if (!CheckMeisai(cdR, dR, iX, sID))
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errRow = iX - 1;
                        dR.Close();
                        sCom.Connection.Close();
                        return false;
                    }
                }

                // 昼食回数加算
                if (dR["昼食"].ToString() == global.FLGON) Lunchs++;

                // 特殊日回数加算
                if (dR["特殊日"].ToString() == global.FLGON) Toks++;

                // 所定勤務日加算
                if (dR["休日"].ToString() == global.hWEEKDAY.ToString()) Shos++;

                // 出勤日数加算①
                if ((dR["所定勤務"].ToString() == global.FLGON ||
                    dR["特殊日"].ToString() == global.FLGON ||
                    dR["開始時"].ToString() != string.Empty) &&
                    dR["休日"].ToString() == global.hWEEKDAY.ToString())
                {
                    rDays++;
                }
                
                // 出勤日数計算②
                // → 特殊日で有休、特休の場合は各々有休、特休扱いとする。
                // ①で一旦出勤加算されているのでここで減算する　　2010/10
                //if (dR["所定勤務"].ToString() == global.FLGON &&  ← dr["特殊日"]の誤り 2012/05/15
                if (dR["特殊日"].ToString() == global.FLGON &&
                   (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU) &&
                    dR["休日"].ToString() == global.hWEEKDAY.ToString())
                {
                    rDays--;
                }

                // 有給休暇日数加算
                if (dR["休暇"].ToString() == global.eYUKYU) Yukyus++;

                // 半休のとき　？？？？？あるの？？？？？
                // 2012/06/19 半休制度導入しました
                if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU) Yukyus += 0.5;

                // 特休日数加算
                if (dR["休暇"].ToString() == global.eKYUMU) Tokkyus++;

                iX++;
            }

            dR.Close();
            sCom.Connection.Close();

            // 役員のとき
            if (sID == 1 || sID == 2)
            {
                if (Utility.StrToInt(cdR["出勤日数合計"].ToString()) > 30)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eDays;
                    global.errRow = 0;
                    global.errMsg = "出勤回数合計が不正です";
                    return false;
                }

                // 出勤日数が所定日数と一致しているか
                if (Utility.StrToInt(cdR["出勤日数合計"].ToString()) != Shos)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eDays;
                    global.errRow = 0;
                    global.errMsg = "出勤回数合計が記入された回数合計と不一致です。 計:" + Shos.ToString() + "日";
                    return false;
                }
            }
            else // 正社員・パートタイマーのとき
            {
                // 出勤日数が記入回数と一致しているか
                if (Utility.StrToInt(cdR["出勤日数合計"].ToString()) != rDays)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eDays;
                    global.errRow = 0;
                    global.errMsg = "出勤回数合計が記入された回数合計と不一致です。 計:" + rDays.ToString() + "日";
                    return false;
                }
            }

            // 出勤日数が記入回数と一致しているか
            if ((sID == 5 || sID == 6) && Toks > 6)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eMark;
                global.errRow = 1;
                global.errMsg = "特殊日のマークは月６回までです";
                return false;
            }

            // 昼食回数チェック
            // 記入回数と一致しているか
            if (Utility.StrToInt(cdR["昼食数合計"].ToString()) != Lunchs)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eLunch;
                global.errRow = 0;
                global.errMsg = "昼食回数合計が記入された回数合計と不一致です。計:" + Lunchs.ToString() + "回";
                return false;
            }

            // 出勤日数内の値か
            if (Lunchs > rDays)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eLunch;
                global.errRow = 0;
                global.errMsg = "昼食回数合計が出勤回数合計を超えています";
                return false;
            }

            // 有休日数合計が記入回数と一致しているか
            if (Utility.StrToDbl(cdR["有休日数合計"].ToString()) != Yukyus)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYukyus;
                global.errRow = 0;
                global.errMsg = "有休日数合計が記入された回数合計と不一致です。計:" + Yukyus.ToString() + "日";
                return false;
            }

            // 特休日数合計が記入回数と一致しているか
            if (Utility.StrToInt(cdR["特休日数合計"].ToString()) != Tokkyus)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eTokkyus;
                global.errRow = 0;
                global.errMsg = "特休日数合計が記入された回数合計と不一致です。計:" + Tokkyus.ToString() + "日";
                return false;
            }

            // 欠勤日数
            if (sID == 7)
            {
                if ((cdR["欠勤日数合計"].ToString() != string.Empty) &&
                    (cdR["欠勤日数合計"].ToString() != global.FLGOFF))
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eKekkins;
                    global.errRow = 0;
                    global.errMsg = "欠勤日数合計は入力できません";
                    return false;
                }
            }
            else if (sID == 5 || sID == 6)
            {
                // 欠勤日数合計が記入回数と一致しているか
                if (Utility.StrToInt(cdR["欠勤日数合計"].ToString()) != global.KekkinCnt)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eKekkins;
                    global.errRow = 0;
                    global.errMsg = "欠勤日数合計が記入された日数合計と不一致です。計:" + global.KekkinCnt.ToString() + "日";
                    return false;
                }

                // 欠勤日数合計は23日以下
                if (Utility.StrToInt(cdR["欠勤日数合計"].ToString()) > 23)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eTokkyus;
                    global.errRow = 0;
                    global.errMsg = "欠勤日数合計が不正です。23日以下";
                    return false;
                }
            }

            return true;
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     明細行毎のエラーチェック </summary>
        /// <param name="cdR">
        ///     勤務票ヘッダデータリーダー</param>
        /// <param name="dR">
        ///     勤務票明細データリーダー</param>
        /// <param name="iX">
        ///     日付</param>
        /// <param name="sID">
        ///     職種ID</param>
        /// <returns>
        ///     </returns>
        ///----------------------------------------------------------------
        private bool CheckMeisai(OleDbDataReader cdR, OleDbDataReader dR, int iX, int sID)
        {
            bool Rtn = true;

            switch (sID)
            {
                // 正社員・準社員
                case 5:
                    Rtn = ChkMeisai05(cdR, dR, iX);
                    break;
                case 6:
                    Rtn = ChkMeisai05(cdR, dR, iX);
                    break;

                // パートタイマー
                case 7:
                    Rtn = ChkMeisai07(cdR, dR, iX);
                    break;

                default:
                    break;
            }

            return Rtn;
        }

        /// <summary>
        /// エラー情報表示
        /// </summary>
        private void ErrShow()
        {
            if (global.errNumber != global.eNothing)
            {
                lblErrMsg.Visible = true;
                lblErrMsg.Text = global.errMsg;

                //対象年
                if (global.errNumber == global.eYear)
                {
                    txtYear.BackColor = Color.Yellow;
                    txtYear.Focus();
                }

                //対象月
                if (global.errNumber == global.eMonth)
                {
                    txtMonth.BackColor = Color.Yellow;
                    txtMonth.Focus();
                }

                // 会社ID
                if (global.errNumber == global.eID)
                {
                    txtID.BackColor = Color.Yellow;
                    txtID.Focus();
                }

                //個人番号
                if (global.errNumber == global.eShainNo)
                {
                    txtNo.BackColor = Color.Yellow;
                    txtNo.Focus();
                }
                
                //特殊日マーク
                if (global.errNumber == global.eMark)
                {
                    dg1[cTokMK, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cTokMK, global.errRow];
                }

                //所定勤務マーク
                if (global.errNumber == global.eSMark)
                {
                    dg1[cShoMK, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cShoMK, global.errRow];
                }

                // 開始時
                if (global.errNumber == global.eSH)
                {
                    dg1[cSH, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cSH, global.errRow];
                }

                // 開始分
                if (global.errNumber == global.eSM)
                {
                    dg1[cSM, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cSM, global.errRow];
                }

                // 終了時
                if (global.errNumber == global.eEH)
                {
                    dg1[cEH, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cEH, global.errRow];
                }

                // 終了分
                if (global.errNumber == global.eEM)
                {
                    dg1[cEM, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cEM, global.errRow];
                }

                // 休暇
                if (global.errNumber == global.eKyuka)
                {
                    dg1[cKyuka, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cKyuka, global.errRow];
                }

                // 休憩なし
                if (global.errNumber == global.eKyukei)
                {
                    dg1[cKyukei, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cKyukei, global.errRow];
                }

                // 昼食
                if (global.errNumber == global.eHiru1)
                {
                    dg1[cHiru1, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cHiru1, global.errRow];
                }

                // 離席開始時
                if (global.errNumber == global.eRSH)
                {
                    dg1[cRSH, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cRSH, global.errRow];
                }

                // 離席開始分
                if (global.errNumber == global.eRSM)
                {
                    dg1[cRSM, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cRSM, global.errRow];
                }

                // 離席終了時
                if (global.errNumber == global.eREH)
                {
                    dg1[cREH, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cREH, global.errRow];
                }

                // 離席終了分
                if (global.errNumber == global.eREM)
                {
                    dg1[cREM, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cREM, global.errRow];
                }

                // 訂正欄 2015/08/18
                if (global.errNumber == global.eTeisei)
                {
                    dg1[cTeisei, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cTeisei, global.errRow];
                }

                // 出勤日数
                if (global.errNumber == global.eDays)
                {
                    txtShuTl.BackColor = Color.Yellow;
                    txtShuTl.Focus();
                }

                // 昼食回数
                if (global.errNumber == global.eLunch)
                {
                    txtChuTl.BackColor = Color.Yellow;
                    txtChuTl.Focus();
                }

                // 有休日数合計
                if (global.errNumber == global.eYukyus)
                {
                    txtYukyuTl.BackColor = Color.Yellow;
                    txtYukyuTl.Focus();
                }

                // 特休日数合計
                if (global.errNumber == global.eTokkyus)
                {
                    txttokuTl.BackColor = Color.Yellow;
                    txttokuTl.Focus();
                }

                // 欠勤日数合計
                if (global.errNumber == global.eKekkins)
                {
                    txtKekkinTl.BackColor = Color.Yellow;
                    txtKekkinTl.Focus();
                }
            }
        }

        private void dg1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnDataMake_Click(object sender, EventArgs e)
        {
            //カレントレコード更新
            CurDataUpDate(_cI);

            //エラーチェック実行①:カレントレコードから最終レコードまで
            if (ErrCheckMain(inData[_cI]._sID, inData[inData.Length - 1]._sID) == false) return;

            //エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (_cI > 0)
            {
                if (ErrCheckMain(inData[0]._sID, inData[_cI - 1]._sID) == false) return;
            }

            // 汎用データ作成
            if (MessageBox.Show("受け渡しデータを作成します。よろしいですか？", "勤怠データ登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;

            SaveData();
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     受け渡しデータ作成 </summary>
        ///-------------------------------------------------------------
        private void SaveData()
        {
            // 出力データ生成
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;

            string OutFileName = string.Empty;
            try
            {
                // オーナーフォームを無効にする
                this.Enabled = false;

                // プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = this;
                frmP.Show();

                // レコード件数取得
                int cTotal = CountMDBitem();
                int rCnt = 1;

                // データベース接続
                sCom.Connection = Con.cnOpen();
                StringBuilder sb = new StringBuilder();

                // ブレーク項目を定義
                string _DenID = string.Empty;   // ヘッダID
                string _kID = string.Empty;     // 個人番号
                string _ComID = string.Empty;   // 会社ID
                string _ShoID = string.Empty;   // 職種ID

                // 月間集計処理を開始
                sb.Clear();
                sb.Append("SELECT 勤務票ヘッダ.*,勤務票明細.* from ");
                sb.Append("勤務票ヘッダ inner join 勤務票明細 ");
                sb.Append("on 勤務票ヘッダ.ID = 勤務票明細.ヘッダID ");
                sb.Append("order by 勤務票ヘッダ.ID,勤務票明細.ID ");
                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                // 出力ファイル名
                if (Properties.Settings.Default.PC == global.PC1SELECT)
                    OutFileName = "PC1_" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒") + ".csv";
                else OutFileName = "PC2_" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒") + ".csv";

                // 出力ファイルインスタンス作成
                StreamWriter outFile = new StreamWriter(global.sDATCOM + OutFileName, false, System.Text.Encoding.GetEncoding(932));
                
                // 奉行シリーズ見出し行出力
                CsvHeaderWrite(outFile);

                // 集計値インスタンス生成
                Entity.saveData sd = new Entity.saveData();
                SaveDataInitial(sd);

                bool sRFlg = false;
                int dIdx = 1;

                // Debug : 個人別出力ファイル
                StreamWriter debFile = null;

                // 明細書き出し
                while (dR.Read())
                {
                    //プログレスバー表示
                    frmP.Text = "汎用データ作成中です・・・" + rCnt.ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = rCnt / cTotal * 100;
                    frmP.ProgressStep();

                    // ヘッダIDブレーク
                    if (_DenID != string.Empty && _DenID != dR["勤務票ヘッダ.ID"].ToString())
                    {
                        // 個人別CSVファイルに集計項目を書き込む
                        CsvWriteFooter(sd, debFile);

                        // 個人別出力CSVファイルを閉じる
                        debFile.Close();

                        // 受け渡しCSV書き込み
                        SaveDataToCsv(outFile, sd, _ComID, _ShoID);
                        dIdx++;
                    }

                    if (_DenID != dR["勤務票ヘッダ.ID"].ToString())
                    {
                        // ヘッダIDを格納する
                        _DenID = dR["勤務票ヘッダ.ID"].ToString();

                        // 個人番号を格納する
                        _kID = dR["個人番号"].ToString();

                        // 会社IDを格納する
                        _ComID = dR["会社ID"].ToString();

                        // 職種IDを格納する
                        _ShoID = dR["職種ID"].ToString();

                        // 集計クラス項目を初期化します
                        SaveDataInitial(sd);

                        // デバッグ区分から判断
                        if (Properties.Settings.Default.DEBUG == 1)
                        {
                            // 出力ファイルインスタンス作成
                            string DebugPath = Properties.Settings.Default.instPath + Properties.Settings.Default.OK +
                                dR["個人番号"].ToString() + " " + dR["氏名"].ToString() + ".csv";
                            debFile = new StreamWriter(DebugPath, false, System.Text.Encoding.GetEncoding(932));

                            // Debug 表題行出力
                            debFile.WriteLine("日付,特殊,所定,開始,終了,昼,休暇,休憩,離席開始,離席終了,所定,執務,早出残業,深夜");
                        }

                        // 会社別設定情報を取得します
                        SysControl.SetDBConnect Con2 = new SysControl.SetDBConnect();
                        OleDbCommand sCom2 = new OleDbCommand();
                        sCom2.Connection = Con2.cnOpen();
                        sCom2.CommandText = "select 丸め単位,早出時間,残業時間 from 会社別設定 where 会社ID=?";
                        sCom2.Parameters.AddWithValue("@Com", _ComID);
                        OleDbDataReader comDR = sCom2.ExecuteReader();

                        if (!comDR.HasRows)
                        {
                            sd.sHayadeSet = "08:40";
                            sd.sZanSTime = "17:00";     // 2010/12/14
                            sd.sMarume = 6;             // 2011/01/11 丸め単位
                        }
                        else
                        {
                            comDR.Read();
                            sd.sHayadeSet = comDR["早出時間"].ToString();      // 2010/12/05

                            // 会社別設定の所定終了時刻を取得 '2010/12/14
                            if (comDR["残業時間"].ToString().Trim() == string.Empty)
                                sd.sZanSTime = "17:00";
                            else sd.sZanSTime = comDR["残業時間"].ToString().Trim();

                            // 会社別設定の丸め単位を取得 '2011/01/11
                            if (comDR["丸め単位"].ToString().Trim() == string.Empty) sd.sMarume = 6;
                            else sd.sMarume = int.Parse(comDR["丸め単位"].ToString().Trim());
                        }
                        comDR.Close();
                        sCom2.Connection.Close();
                        
                        // 個人番号、欠勤日数を取得
                        sd.sCode = dR["個人番号"].ToString().PadLeft(7, '0');
                        sd.sKekkinNisu = Utility.StrToInt(dR["欠勤日数合計"].ToString()).ToString();
                    }

                    // 集計処理
                    switch (int.Parse(dR["職種ID"].ToString()))
                    {
                        // パートタイマー
                        case 7:
                            MakeSavePart(sd, dR, dIdx);
                            break;

                        // 正社員・準社員
                        default:
                            MakeSaveStaff(sd, dR, dIdx);
                            break;
                    }

                    // 個人別CSVファイル出力
                    CsvWriteSub(dR, sd, debFile);
                }

                // 個人別CSVファイルに集計項目を書き込む
                CsvWriteFooter(sd, debFile);

                // 個人別出力CSVファイルを閉じる
                debFile.Close();

                // 受け渡しCSVを書き込み
                SaveDataToCsv(outFile, sd, _ComID, _ShoID);

                // データリーダーをクローズします
                dR.Close();

                // 出力ファイルをクローズします
                outFile.Close();

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                // 画像ファイルを退避します
                tifFileMove(global.sTIFCOM);

                // 過去データ登録
                DeleteLastData();   // 登録済み過去データ削除
                SaveLastData();     // 過去データ登録

                // 勤務票レコード削除
                DelHeadRec();

                // 設定月数分経過した過去画像を削除する
                //delBackUpFiles(global.sBKDELS, global.sTIF);

                //終了
                MessageBox.Show("受け渡しデータが作成されました。", "勤怠データ作成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Tag = END_MAKEDATA;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (!dR.IsClosed) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();

                //MDBファイル最適化
                mdbCompact();
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     個人別のCSVファイルを出力 </summary>
        /// <param name="dR">
        ///     勤務票データリーダー</param>
        /// <param name="SD">
        ///     集計値配列</param>
        /// <param name="deb">
        ///     出力ストリーム</param>
        /// -------------------------------------------------------------------
        private void CsvWriteSub(OleDbDataReader dR, Entity.saveData SD, StreamWriter deb)
        {
            if (Properties.Settings.Default.DEBUG == 1)
            {
                StringBuilder sb = new StringBuilder();

                // 10進数表記可能な場合は小数点表記   2011/01/14
                sb.Append(dR["日付"].ToString().PadLeft(2, '0')).Append("日").Append(",");

                if (dR["特殊日"].ToString() == global.FLGON) sb.Append(global.MARU).Append(",");
                else sb.Append(string.Empty).Append(",");

                if (dR["所定勤務"].ToString() == global.FLGON) sb.Append(global.MARU).Append(",");
                else sb.Append(string.Empty).Append(",");

                if ((((dR["特殊日"].ToString() == global.FLGON || 
                       dR["所定勤務"].ToString() == global.FLGON) &&
                       dR["開始時"].ToString() == string.Empty)) ||
                      (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU))
                {
                    sb.Append(SD.ShoteiS.Hour.ToString("00") + ":" + SD.ShoteiS.Minute.ToString("00")).Append(",");
                    sb.Append(SD.ShoteiE.Hour.ToString("00") + ":" + SD.ShoteiE.Minute.ToString("00")).Append(",");
                }
                else 
                {
                    sb.Append(dR["開始時"].ToString()).Append(":").Append(dR["開始分"].ToString()).Append(",");
                    sb.Append(dR["終了時"].ToString()).Append(":").Append(dR["終了分"].ToString()).Append(",");
                }

                if (dR["昼食"].ToString() == global.FLGON) sb.Append(global.MARU).Append(",");
                else sb.Append(string.Empty).Append(",");

                sb.Append(dR["休暇"].ToString()).Append(",");
                
                if (dR["休憩なし"].ToString() == global.FLGON) sb.Append(global.MARU).Append(",");
                else sb.Append(string.Empty).Append(",");

                sb.Append(dR["離席開始時"].ToString()).Append(":").Append(dR["離席開始分"].ToString()).Append(",");
                sb.Append(dR["離席終了時"].ToString()).Append(":").Append(dR["離席終了分"].ToString()).Append(",");

                if ((SD.sMarume * 100) % 60 == 0)
                {
                    sb.Append((SD.TimeJ / 60).ToString()).Append(",");
                    sb.Append((Utility.fncTimeSet(SD.sSitumu2, SD.sMarume) / 60).ToString()).Append(",");
                    sb.Append((Utility.fncTimeSet(SD.Zangyo + SD.Hayade, SD.sMarume) / 60).ToString()).Append(",");
                    sb.Append((Utility.fncTimeSet(SD.sSinya + SD.kSinya, SD.sMarume) / 60).ToString());
                }
                else
                {
                    sb.Append(Utility.fncTimehhmm(SD.TimeJ)).Append(",");
                    sb.Append(Utility.fncTimehhmm(SD.sSitumu2)).Append(",");
                    sb.Append(Utility.fncTimehhmm(SD.Zangyo + SD.Hayade)).Append(",");
                    sb.Append(Utility.fncTimehhmm(SD.sSinya + SD.kSinya));
                }
                deb.WriteLine(sb.ToString());
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     個人別CSVファイルに集計値を出力する </summary>
        /// <param name="SD">
        ///     集計値配列</param>
        /// <param name="deb">
        ///     出力ストリーム</param>
        ///----------------------------------------------------------
        private void CsvWriteFooter(Entity.saveData SD, StreamWriter deb)
        {
            if (Properties.Settings.Default.DEBUG == 1)
            {
                deb.WriteLine("出勤日数," + SD.sJituKinmu1.ToString());
                deb.WriteLine("昼食回数," + SD._sHiru1.ToString());
                deb.WriteLine("有休日数," + SD.sJituKinmu2.ToString());
                deb.WriteLine("特休日数," + SD.sJituKinmu3.ToString());
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     パートタイマー受け渡しデータ集計 </summary>
        /// <param name="SD">
        ///     集計データ</param>
        /// <param name="dR">
        ///     勤務票データリーダー</param>
        /// <param name="iX">
        ///     データインデックス</param>
        ///----------------------------------------------------------
        private void MakeSavePart(Entity.saveData SD, OleDbDataReader dR, int iX)
        {
            string _sSH = string.Empty;
            string _sSM = string.Empty;
            string _sEH = string.Empty;
            string _sEM = string.Empty;

            double TimeSHO = 0;

            DateTime TimeS = DateTime.Now;
            DateTime TimeE = DateTime.Now;
            DateTime TimeRS = DateTime.Parse("1900/01/01 0:00");
            DateTime TimeRE = DateTime.Parse("1900/01/01 0:00");
            DateTime TimeNull = DateTime.Parse("1900/01/01 0:00");
            DateTime TimeEKyujitu;
            DateTime dt2200 = DateTime.Parse("22:00");

            double sShotei = 0;

            SD.TimeJ = 0;
            SD.TimeJ2 = 0;
            SD.kSinya = 0;

            // 執務・残業・早出・深夜
            SD.sSitumu1 = 0;
            SD.sSitumu2 = 0;
            SD.Zangyo = 0;
            SD.sSinya = 0;
            SD.Hayade = 0;

            DateTime dt;
            DateTime sHayade = DateTime.Parse(SD.sHayadeSet);
            DateTime amHankyuHayade = DateTime.Parse(SD.sHayadeSet); // 2015/05/13 午前半のときの午後開始時間の仮の初期値

            // 所定又は特殊開始終了時間を取得します
            if (dR["特殊日"].ToString() == global.FLGON)
            {
                // 午前半休のとき
                if (dR["休暇"].ToString() == global.eAMHANKYU)
                {
                    if (DateTime.TryParse(dR["特殊午後半休開始時間"].ToString(), out dt))
                    {
                        SD.ShoteiS = dt;

                        // 午前半休時の午後開始時間 2015/05/13
                        amHankyuHayade = dt;
                    }

                    if (DateTime.TryParse(dR["特殊午後半休終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
                // 午後半休のとき
                else if (dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    if (DateTime.TryParse(dR["特殊午前半休開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                    if (DateTime.TryParse(dR["特殊午前半休終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
                else
                {
                    // 特殊開始終了時間を取得します
                    if (DateTime.TryParse(dR["特殊開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                    if (DateTime.TryParse(dR["特殊終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
            }
            else
            {
                // 午前半休のとき
                if (dR["休暇"].ToString() == global.eAMHANKYU)
                {
                    if (DateTime.TryParse(dR["午後半休開始時間"].ToString(), out dt))
                    {
                        SD.ShoteiS = dt;

                        // 午前半休時の午後開始時間 2015/05/13
                        amHankyuHayade = dt;
                    }

                    if (DateTime.TryParse(dR["午後半休終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
                // 午後半休のとき
                else if (dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    if (DateTime.TryParse(dR["午前半休開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                    if (DateTime.TryParse(dR["午前半休終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
                else
                {
                    // 所定開始終了時間を取得します
                    if (DateTime.TryParse(dR["所定開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                    if (DateTime.TryParse(dR["所定終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
            }
            
            // マーク、有給、休務のときは所定勤務時間を適用します
            if ((((dR["特殊日"].ToString() == global.FLGON || 
                 dR["所定勤務"].ToString() == global.FLGON) &&
                 dR["開始時"].ToString() == string.Empty)) ||
                (dR["休暇"].ToString() == global.eYUKYU ||
                dR["休暇"].ToString() == global.eKYUMU))
            {
                _sSH = dR["所定開始時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                _sSM = dR["所定開始時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
                _sEH = dR["所定終了時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                _sEM = dR["所定終了時間"].ToString().Replace(":", string.Empty).Substring(2, 2);

                // 特殊マークのときは特殊勤務時間を適用します
                if (dR["特殊日"].ToString() == global.FLGON)
                {
                    _sSH = dR["特殊開始時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                    _sSM = dR["特殊開始時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
                    _sEH = dR["特殊終了時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                    _sEM = dR["特殊終了時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
                }
            }
                // 開始時間、終了時間が記入されているとき
            else if (dR["開始時"].ToString() != string.Empty)
            {
                _sSH = dR["開始時"].ToString();
                _sSM = dR["開始分"].ToString();
                _sEH = dR["終了時"].ToString();
                _sEM = dR["終了分"].ToString();
            }

            // 実勤務日数を加算
            if (dR["所定勤務"].ToString() == global.FLGON || dR["特殊日"].ToString() == global.FLGON || _sSH != string.Empty ||
                dR["休暇"].ToString() == global.eFURIDE || dR["休暇"].ToString() == global.eAMHANKYU ||
                dR["休暇"].ToString() == global.ePMHANKYU)
            {
                if ((dR["休日"].ToString() == global.hSATURDAY.ToString() ||
                     dR["休日"].ToString() == global.hHOLIDAY.ToString()) &&
                    (dR["休暇"].ToString() != global.eFURIDE))
                {
                    SD.sJituKinmu5++;
                }
                else if (dR["休暇"].ToString() != global.eYUKYU && dR["休暇"].ToString() != global.eKYUMU) 
                        SD.sJituKinmu1++;
            }

            // 実勤務日数 内有休日数
            if (dR["休暇"].ToString() == global.eYUKYU) SD.sJituKinmu2++;

            // 実勤務日数 内半休
            if (dR["休暇"].ToString() == global.eAMHANKYU ||
                dR["休暇"].ToString() == global.ePMHANKYU)
            {
                SD.sJituKinmu2 += 0.5;
                //SD.sKinmuJikan2 += double.Parse(dR["所定勤務時間"].ToString());
                SD.sKinmuJikan2 += Utility.fncTimeSet(Utility.GetTimeSpan(SD.ShoteiS, SD.ShoteiE).TotalMinutes, SD.sMarume);
            }

            // 実勤務日数 内特休日数
            if (dR["休暇"].ToString() == global.eKYUMU) SD.sJituKinmu3++;

            if (dR["所定勤務"].ToString() == global.FLGON && dR["特殊日"].ToString() == global.FLGON && 
                dR["開始時"].ToString() != string.Empty)
                SD.sJituKinmu4++;

            // 勤務時間 平常日所定勤務時間
            if (_sSH != string.Empty && _sSM != string.Empty && _sEH != string.Empty && _sEM != string.Empty)
            {
                // 開始時間を取得
                if (DateTime.TryParse(_sSH + ":" + _sSM, out dt)) TimeS = dt;

                // 終了時間を取得
                if (DateTime.TryParse(_sEH + ":" + _sEM, out dt)) TimeE = dt;

                // 開始時間のほうが大きいとき
                if (TimeS > TimeE)
                {
                    TimeSpan ts = new TimeSpan(24, 0, 0);
                    TimeE = TimeE + ts;
                }
                
                // 勤務終了時間が早出開始時間以前の場合、勤務終了時間を早出開始時間とする
                if (TimeE < sHayade) sHayade = TimeE;

                // 離席開始、離席終了時間
                if (dR["離席開始時"].ToString() != string.Empty && dR["離席開始分"].ToString() != string.Empty && 
                    dR["離席終了時"].ToString() != string.Empty && dR["離席終了分"].ToString() != string.Empty)
                {
                    if (DateTime.TryParse(dR["離席開始時"].ToString() + ":" + dR["離席開始分"].ToString(), out dt)) TimeRS = dt;
                    if (DateTime.TryParse(dR["離席終了時"].ToString() + ":" + dR["離席終了分"].ToString(), out dt)) TimeRE = dt;
                    SD.TimeJ2 = Utility.GetTimeSpan(TimeRS, TimeRE).TotalMinutes;
                }

                if (dR["休憩なし"].ToString() == global.FLGON)
                    SD.TimeJ = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes;
                else SD.TimeJ = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes - 60;
                
                // 休日だったら次の日へ
                if ((dR["休日"].ToString() == global.hSATURDAY.ToString() ||
                    dR["休日"].ToString() == global.hHOLIDAY.ToString()) &&
                    dR["休暇"].ToString() != global.eFURIDE)
                {                    
                    // 2011/01/18 勤務終了が22:00超か？
                    if (TimeE > dt2200) TimeEKyujitu = dt2200;
                    else TimeEKyujitu = TimeE;
                    
                    // 2011/01/18 所定勤務時間計算
                    if (dR["休憩なし"].ToString() == global.FLGON)
                        SD.TimeJ = Utility.GetTimeSpan(TimeS, TimeEKyujitu).TotalMinutes;
                    else SD.TimeJ = Utility.GetTimeSpan(TimeS, TimeEKyujitu).TotalMinutes - 60;

                    SD.TimeJ = Utility.fncTimeSet(SD.TimeJ, SD.sMarume);
                    SD.sKyujitu += SD.TimeJ;

                    // 休日深夜勤務時間
                    if (TimeE > dt2200) // 終了時間が22時以降である
                    {
                        SD.kSinya = Utility.GetTimeSpan(dt2200, TimeE).TotalMinutes;
                        
                        // 22時以降の離席時間を減算します
                        if (TimeRS != TimeNull)
                            SD.kSinya -= Utility.GetRisekiTimeSpan(dt2200, TimeE, TimeRS, TimeRE);

                        // '小数点10進ではなく分単位で丸め処理　2011/01/11
                        SD.KSinyaT += Utility.fncTimeSet(SD.kSinya, SD.sMarume);
                    }
                    return;
                }

                // 
                // 所定時間内の勤務時間を取得します
                if (TimeE <= SD.ShoteiS || TimeS > SD.ShoteiE)
                    sShotei = 0;
                else
                {
                    // 開始時間：開始時が所定開始時間以前のときは所定開始時間を適用します
                    DateTime st;
                    
                    // 2015/05/13 午前半休のとき開始時が午後開始時刻以前のときは午後開始時刻を適用します
                    if (dR["休暇"].ToString() == global.eAMHANKYU)
                    {
                        if (TimeS < amHankyuHayade)
                        {
                            st = amHankyuHayade;
                        }
                        else
                        {
                            st = TimeS;
                        }
                    }
                    else
                    {
                        if (TimeS < SD.ShoteiS)
                        {
                            st = SD.ShoteiS;
                        }
                        else
                        {
                            st = TimeS;
                        }
                    }

                    
                    // 終了時間：終了時が所定終了時間以降のときは所定終了時間を適用します
                    DateTime et;
                    if (TimeE > SD.ShoteiE) et = SD.ShoteiE;
                    else et = TimeE;

                    // 所定時間内の勤務時間を取得します
                    sShotei = Utility.GetTimeSpan(st, et).TotalMinutes;

                    // 離席時間を減算します
                    if (TimeRS != TimeNull)
                        sShotei -= Utility.GetRisekiTimeSpan(st, et, TimeRS, TimeRE);
                }

                // ②休憩の有無、半休のとき
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    // 分単位で丸め処理
                    SD.TimeJ = Utility.fncTimeSet(sShotei,SD.sMarume);
                    TimeSHO = sShotei;
                }
                // 半休のとき
                else if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    // 分単位で丸め処理
                    SD.TimeJ = Utility.fncTimeSet(sShotei, SD.sMarume);
                    TimeSHO = sShotei;
                }
                else
                {
                    // 休憩有り60分減らす
                    SD.TimeJ = Utility.fncTimeSet(sShotei - 60, SD.sMarume);
                    TimeSHO = sShotei - 60;
                }

                // ③平日のとき 実勤務時間に加算
                if (dR["休日"].ToString() == global.hWEEKDAY.ToString())
                    SD.sKinmuJikan1 += SD.TimeJ;
                else
                {
                    // ④休日に振替出勤のとき 実勤務時間に加算
                    if (dR["休暇"].ToString() == global.eFURIDE)
                        SD.sKinmuJikan1 += SD.TimeJ;
                }

                // ⑤有休または特休取得日のとき「実勤務時間2」に加算
                if (dR["休暇"].ToString() == global.eYUKYU ||
                    dR["休暇"].ToString() == global.eKYUMU)
                    SD.sKinmuJikan2 += SD.TimeJ;

                /* 
                 * → 半休はエラーとなるためこのロジックは通らない ⑥半休のとき半休所定勤務時間（所定勤務時間の1/2）を実勤務時間に加算
                 * → 2012/06/18 半休制度導入のため有効なロジックとなる。半休扱いの勤務時間を加算しています。　
                */
                if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    //SD.sKinmuJikan1 += double.Parse(dR["所定勤務時間"].ToString());
                    SD.sKinmuJikan1 += Utility.fncTimeSet(Utility.GetTimeSpan(SD.ShoteiS, SD.ShoteiE).TotalMinutes, SD.sMarume);
            }
            else
            {
                /*
                 * → 半休はエラーとなるためこのロジックは通らない ⑥半休のとき半休所定勤務時間（所定勤務時間の1/2）を実勤務時間に加算
                 * → 2012/06/18 半休制度導入 ただし開始、終了時間の記入は必須のためこのロジックは通らない
                */
                //if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                //    //SD.sKinmuJikan1 += double.Parse(dR["所定勤務時間"].ToString());
                //    SD.sKinmuJikan1 += Utility.fncTimeSet(Utility.GetTimeSpan(SD.ShoteiS, SD.ShoteiE).TotalMinutes, SD.sMarume);
            }

            // 昼食回数１
            if (dR["昼食"].ToString() == global.FLGON) SD._sHiru1++;


            // "○"または開始時間が記入されているとき
            if (dR["特殊日"].ToString() == global.FLGON || dR["所定勤務"].ToString() == global.FLGON || _sSH != string.Empty)
            {
                SD.TimeT = 0;
                SD.sSitumu2 = 0;

                //
                // 執務時間を設定します
                //


                // 2015/05/13 午前半休のとき
                if (dR["休暇"].ToString() == global.eAMHANKYU)
                {
                    if (TimeS > amHankyuHayade)
                    {
                        SD.SituS = TimeS;
                        SD.SituE = TimeS.AddMinutes(540);
                    }
                    else
                    {
                        // 午後開始時刻から9時間とします
                        SD.SituS = amHankyuHayade;
                        SD.SituE = amHankyuHayade.AddMinutes(540);
                    }
                }
                else
                {
                    // 開始時間が早出時間(開始時間）以降のとき開始時間から9時間とします
                    if (TimeS > sHayade)
                    {
                        SD.SituS = TimeS;
                        SD.SituE = TimeS.AddMinutes(540);
                    }
                    else
                    {
                        // 開始時間が早出時間(開始時間）以前のとき早出時間(開始時間）から9時間とします
                        SD.SituS = sHayade;
                        SD.SituE = sHayade.AddMinutes(540);
                    }
                }

                // 開始時間から実際の終了時間までの時間
                TimeSpan ts = Utility.GetTimeSpan(SD.SituS, TimeE);

                // 離席時間
                double r = 0;
                if (TimeRS != TimeNull)
                {
                    r = Utility.GetRisekiTimeSpan(SD.SituS, TimeE, TimeRS, TimeRE);
                }

                // 所定勤務時間以上のとき
                if (ts.TotalMinutes > sShotei)
                {
                    SD.sSitumu2 = ts.TotalMinutes - r - sShotei;
                }

                // 午前半休のとき午後開始時間より前に出勤した場合は執務とする 2015/05/13
                if (dR["休暇"].ToString() == global.eAMHANKYU)
                {
                    if (TimeS < amHankyuHayade)
                    {
                        SD.sSitumu2 += Utility.GetTimeSpan(TimeS, amHankyuHayade).TotalMinutes;
                    }

                    // 午後開始時間より前の出勤中の離席時間
                    if (TimeRS != TimeNull)
                    {
                        SD.sSitumu2 -= Utility.GetRisekiTimeSpan(TimeS, amHankyuHayade, TimeRS, TimeRE);
                    }
                }
                
                //
                // 平常日時間外 早出残業
                //
                // ①早出残業時間を取得します
                if (TimeS < sHayade)
                {
                    SD.Hayade = Utility.GetTimeSpan(TimeS, sHayade).TotalMinutes;

                    // 離席時間を減算します
                    if (TimeRS != TimeNull)
                    {
                        SD.Hayade -= Utility.GetRisekiTimeSpan(TimeS, sHayade, TimeRS, TimeRE);
                    }
                }

                // ②時間外残業
                if (TimeS < sHayade)
                {
                    if (dR["休憩なし"].ToString() == global.FLGON)
                    {
                        SD.ShoETime = sHayade.AddMinutes(480);
                    }
                    else
                    {
                        SD.ShoETime = sHayade.AddMinutes(540);
                    }
                }
                else
                {
                    if (dR["休憩なし"].ToString() == global.FLGON)
                    {
                        SD.ShoETime = TimeS.AddMinutes(480);
                    }
                    else
                    {
                        SD.ShoETime = TimeS.AddMinutes(540);
                    }
                }
                
                // 離席があるとき終了時刻に離席時間を加算する
                if (TimeRS != TimeNull)
                {
                    SD.ShoETime = SD.ShoETime.AddMinutes(SD.TimeJ2);
                }

                // 所定終了時間より終了時間が遅いとき 終了時間までの経過時間を取得します（但し22:00mまで）
                DateTime t2200 = DateTime.Parse("22:00");
                if (SD.ShoETime < TimeE)
                {
                    if (TimeE < t2200)
                    {
                        SD.Zangyo = Utility.GetTimeSpan(SD.ShoETime, TimeE).TotalMinutes;
                    }
                    else
                    {
                        SD.Zangyo = Utility.GetTimeSpan(SD.ShoETime, t2200).TotalMinutes;
                    }
                }

                SD.HayaZanT += Utility.fncTimeSet(SD.Zangyo + SD.Hayade, SD.sMarume);

                // ③平常日時間外 深夜（22:00以降）
                if (TimeE > t2200)
                {
                    SD.sSinya = Utility.GetTimeSpan(t2200, TimeE).TotalMinutes;

                    // 離席時間を減算します
                    if (TimeRS != TimeNull)
                    {
                        SD.sSinya -= Utility.GetRisekiTimeSpan(t2200, TimeE, TimeRS, TimeRE);
                    }

                    SD.SinyaT += Utility.fncTimeSet(SD.sSinya, SD.sMarume);
                }
            }

            // 執務時間計算
            SD.sSitumu2 = SD.sSitumu2 - SD.Zangyo - SD.sSinya;
            SD.sSitumuT = SD.sSitumuT + Utility.fncTimeSet(SD.sSitumu2, SD.sMarume);  // 小数点10進ではなく分単位で丸め処理　2011/01/11
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     スタッフ受け渡しデータ集計 </summary>
        /// <param name="SD">
        ///     集計データ</param>
        /// <param name="dR">
        ///     勤務票データリーダー</param>
        /// <param name="iX">
        ///     データインデックス</param>
        ///----------------------------------------------------------
        private void MakeSaveStaff(Entity.saveData SD, OleDbDataReader dR, int iX)
        {
            string _sSH = string.Empty;
            string _sSM = string.Empty;
            string _sEH = string.Empty;
            string _sEM = string.Empty;

            double TimeSHO = 0;

            DateTime TimeS = DateTime.Now;
            DateTime TimeE = DateTime.Now;
            DateTime TimeRS = DateTime.Parse("1900/01/01 0:00");
            DateTime TimeRE = DateTime.Parse("1900/01/01 0:00");
            DateTime TimeNull = DateTime.Parse("1900/01/01 0:00");
            DateTime TimeEKyujitu;
            DateTime dt2200 = DateTime.Parse("22:00");

            double sShotei = 0;
            DateTime dt;
            DateTime sHayade = DateTime.Parse(SD.sHayadeSet);  // 早出時刻
            DateTime amHankyuHayade = DateTime.Parse(SD.sHayadeSet); // 2015/05/13 午前半のときの午後開始時間の仮の初期値
            DateTime s;
            DateTime e;

            SD.TimeJ = 0;
            SD.TimeJ2 = 0;
            SD.kSinya = 0;

            // 執務・残業・早出・深夜
            SD.sSitumu1 = 0;
            SD.sSitumu2 = 0;
            SD.Zangyo = 0;
            SD.sSinya = 0;
            SD.Hayade = 0;

            // 所定又は特殊開始終了時間を取得します
            if (dR["特殊日"].ToString() == global.FLGON)
            {
                // 午前半休のとき
                if (dR["休暇"].ToString() == global.eAMHANKYU)
                {
                    if (DateTime.TryParse(dR["特殊午後半休開始時間"].ToString(), out dt))
                    {
                        SD.ShoteiS = dt;

                        // 午前半休時の午後開始時間 2015/05/13
                        amHankyuHayade = dt;
                    }

                    if (DateTime.TryParse(dR["特殊午後半休終了時間"].ToString(), out dt))
                    {
                        SD.ShoteiE = dt;
                    }

                }
                // 午後半休のとき
                else if (dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    if (DateTime.TryParse(dR["特殊午前半休開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                    if (DateTime.TryParse(dR["特殊午前半休終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
                else
                {
                    // 特殊開始終了時間を取得します
                    if (DateTime.TryParse(dR["特殊開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                    if (DateTime.TryParse(dR["特殊終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
            }
            else
            {
                // 午前半休のとき
                if (dR["休暇"].ToString() == global.eAMHANKYU)
                {
                    if (DateTime.TryParse(dR["午後半休開始時間"].ToString(), out dt))
                    {
                        SD.ShoteiS = dt;

                        // 午前半休時の午後開始時間 2015/05/13
                        amHankyuHayade = dt;
                    }

                    if (DateTime.TryParse(dR["午後半休終了時間"].ToString(), out dt))
                    {
                        SD.ShoteiE = dt;
                    }
                }
                // 午後半休のとき
                else if (dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    if (DateTime.TryParse(dR["午前半休開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                    if (DateTime.TryParse(dR["午前半休終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
                else
                {
                    // 所定開始終了時間を取得します
                    if (DateTime.TryParse(dR["所定開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                    if (DateTime.TryParse(dR["所定終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
                }
            }

            // マーク、有給、休務のときは所定勤務時間を適用します ※ 所定開始時間が取得されていること
            if (((((dR["特殊日"].ToString() == global.FLGON ||
                 dR["所定勤務"].ToString() == global.FLGON) &&
                 dR["開始時"].ToString() == string.Empty)) ||
                (dR["休暇"].ToString() == global.eYUKYU ||
                dR["休暇"].ToString() == global.eKYUMU)) &&
                dR["所定開始時間"].ToString() != string.Empty)
            {
                _sSH = dR["所定開始時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                _sSM = dR["所定開始時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
                _sEH = dR["所定終了時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                _sEM = dR["所定終了時間"].ToString().Replace(":", string.Empty).Substring(2, 2);

                // 特殊マークのときは特殊勤務時間を適用します ※ 特殊開始時間が取得されていること
                if (dR["特殊日"].ToString() == global.FLGON && dR["特殊開始時間"].ToString() != string.Empty)
                {
                    _sSH = dR["特殊開始時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                    _sSM = dR["特殊開始時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
                    _sEH = dR["特殊終了時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                    _sEM = dR["特殊終了時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
                }
            }
                // 開始時間、終了時間が記入されているとき
            else if (dR["開始時"].ToString() != string.Empty)
            {
                _sSH = dR["開始時"].ToString();
                _sSM = dR["開始分"].ToString();
                _sEH = dR["終了時"].ToString();
                _sEM = dR["終了分"].ToString();
            }

            // 実勤務日数を加算
            if (dR["所定勤務"].ToString() == global.FLGON || dR["特殊日"].ToString() == global.FLGON || _sSH != string.Empty ||
                dR["休暇"].ToString() == global.eFURIDE || dR["休暇"].ToString() == global.eAMHANKYU ||
                dR["休暇"].ToString() == global.ePMHANKYU)
            {
                if ((dR["休日"].ToString() == global.hSATURDAY.ToString() ||
                     dR["休日"].ToString() == global.hHOLIDAY.ToString()) &&
                    (dR["休暇"].ToString() != global.eFURIDE))
                {
                    SD.sJituKinmu5++;
                }
                else if (dR["休暇"].ToString() != global.eYUKYU && dR["休暇"].ToString() != global.eKYUMU)
                    SD.sJituKinmu1++;
            }

            // 役員の勤務日数
            if (int.Parse(dR["職種ID"].ToString()) == 1 || int.Parse(dR["職種ID"].ToString()) == 2)
                SD.sJituKinmu1 = Utility.StrToInt(dR["出勤日数合計"].ToString());

            // 実勤務日数 内有休日数
            if (dR["休暇"].ToString() == global.eYUKYU) SD.sJituKinmu2++;

            // 実勤務日数 内半休
            if (dR["休暇"].ToString() == global.eAMHANKYU ||
                dR["休暇"].ToString() == global.ePMHANKYU)
            {
                SD.sJituKinmu2 += 0.5;
                //SD.sKinmuJikan2 += double.Parse(dR["所定勤務時間"].ToString());
                SD.sKinmuJikan2 += Utility.fncTimeSet(Utility.GetTimeSpan(SD.ShoteiS, SD.ShoteiE).TotalMinutes, SD.sMarume);
            }

            // 実勤務日数 内特休日数
            if (dR["休暇"].ToString() == global.eKYUMU) SD.sJituKinmu3++;

            if (dR["所定勤務"].ToString() == global.FLGON && dR["特殊日"].ToString() == global.FLGON &&
                dR["開始時"].ToString() != string.Empty)
                SD.sJituKinmu4++;

            // 勤務時間 平常日所定勤務時間
            if (_sSH != string.Empty && _sSM != string.Empty && _sEH != string.Empty && _sEM != string.Empty)
            {
                // 開始時間を取得
                if (DateTime.TryParse(_sSH + ":" + _sSM, out dt)) TimeS = dt;

                // 終了時間を取得
                if (DateTime.TryParse(_sEH + ":" + _sEM, out dt)) TimeE = dt;

                // 開始時間のほうが大きいとき
                if (TimeS > TimeE)
                {
                    TimeSpan ts = new TimeSpan(24, 0, 0);
                    TimeE = TimeE + ts;
                }

                // 勤務終了時間が早出開始時間以前の場合、勤務終了時間を早出開始時間とする
                if (TimeE < sHayade) sHayade = TimeE;

                // 離席開始、離席終了時間
                if (dR["離席開始時"].ToString() != string.Empty && dR["離席開始分"].ToString() != string.Empty &&
                    dR["離席終了時"].ToString() != string.Empty && dR["離席終了分"].ToString() != string.Empty)
                {
                    if (DateTime.TryParse(dR["離席開始時"].ToString() + ":" + dR["離席開始分"].ToString(), out dt)) TimeRS = dt;
                    if (DateTime.TryParse(dR["離席終了時"].ToString() + ":" + dR["離席終了分"].ToString(), out dt)) TimeRE = dt;
                    SD.TimeJ2 = Utility.GetTimeSpan(TimeRS, TimeRE).TotalMinutes;
                }

                if (dR["休憩なし"].ToString() == global.FLGON)
                    SD.TimeJ = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes;
                else SD.TimeJ = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes - 60;

                // 休日だったら次の日へ
                if ((dR["休日"].ToString() == global.hSATURDAY.ToString() ||
                    dR["休日"].ToString() == global.hHOLIDAY.ToString()) &&
                    dR["休暇"].ToString() != global.eFURIDE)
                {
                    // 2011/01/18 勤務終了が22:00超か？
                    if (TimeE > dt2200) TimeEKyujitu = dt2200;
                    else TimeEKyujitu = TimeE;

                    // 2011/01/18 所定勤務時間計算
                    if (dR["休憩なし"].ToString() == global.FLGON)
                        SD.TimeJ = Utility.GetTimeSpan(TimeS, TimeEKyujitu).TotalMinutes;
                    else SD.TimeJ = Utility.GetTimeSpan(TimeS, TimeEKyujitu).TotalMinutes - 60;

                    SD.TimeJ = Utility.fncTimeSet(SD.TimeJ, SD.sMarume);
                    SD.sKyujitu += SD.TimeJ;

                    // 休日深夜勤務時間
                    if (TimeE > dt2200) // 終了時間が22時以降である
                    {
                        SD.kSinya = Utility.GetTimeSpan(dt2200, TimeE).TotalMinutes;

                        // 22時以降の離席時間を減算します
                        if (TimeRS != TimeNull)
                            SD.kSinya -= Utility.GetRisekiTimeSpan(dt2200, TimeE, TimeRS, TimeRE);

                        // '小数点10進ではなく分単位で丸め処理　2011/01/11
                        SD.KSinyaT += Utility.fncTimeSet(SD.kSinya, SD.sMarume);
                    }
                    return;
                }

                // 所定勤務時間を取得します
                if (TimeE <= sHayade || TimeS > DateTime.Parse(SD.sZanSTime))
                    sShotei = 0;
                else
                {
                    // 開始時間：開始時が早出時間以前のときは早出時間を適用します
                    DateTime st;

                    // 2015/05/13 午前半休のとき開始時が午後開始時刻以前のときは午後開始時刻を適用します
                    if (dR["休暇"].ToString() == global.eAMHANKYU)
                    {
                        if (TimeS < amHankyuHayade)
                        {
                            st = amHankyuHayade;
                        }
                        else
                        {
                            st = TimeS;
                        }
                    }
                    else
                    {
                        if (TimeS < sHayade)
                        {
                            st = sHayade;
                        }
                        else
                        {
                            st = TimeS;
                        }
                    }

                    // 終了時間：終了時が所定終了時間以降のときは所定終了時間を適用します
                    DateTime et;
                        
                    // 所定時間内の勤務時間を取得します
                    if (dR["特殊日"].ToString() == global.FLGON)
                    {
                        if (TimeE > SD.ShoteiE) et = SD.ShoteiE;
                        else et = TimeE;
                    }
                    // 半休のとき
                    else if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    {
                        // 終了時間：終了時が所定終了時間以降のときは所定終了時間を適用します
                        if (TimeE > SD.ShoteiE) et = SD.ShoteiE;
                        else et = TimeE;
                    }
                    else
                    // 終了時間：終了時が設定残業時間以降のときは設定残業時間を適用します
                    {
                        if (TimeE > DateTime.Parse(SD.sZanSTime)) et = DateTime.Parse(SD.sZanSTime);
                        else et = TimeE;
                    }
                    
                    // 所定時間内の勤務時間を取得します
                    sShotei = Utility.GetTimeSpan(st, et).TotalMinutes;
                        
                    // 離席時間を減算します
                    if (TimeRS != TimeNull)
                        sShotei -= Utility.GetRisekiTimeSpan(st, et, TimeRS, TimeRE);
                }

                // ②休憩の有無、半休のとき
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    // 分単位で丸め処理
                    SD.TimeJ = Utility.fncTimeSet(sShotei, SD.sMarume);
                    TimeSHO = sShotei;
                }
                // 半休のとき''''半休はエラーとなる → 半休制度導入 2012/06/19
                else if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    // 分単位で丸め処理
                    SD.TimeJ = Utility.fncTimeSet(sShotei, SD.sMarume);
                    TimeSHO = sShotei;
                }
                else
                {
                    // 休憩有り60分減らす
                    SD.TimeJ = Utility.fncTimeSet(sShotei - 60, SD.sMarume);
                    TimeSHO = sShotei - 60;
                }

                // ③平日のとき 実勤務時間に加算
                if (dR["休日"].ToString() == global.hWEEKDAY.ToString())
                    SD.sKinmuJikan1 += SD.TimeJ;
                else
                {
                    // ④休日に振替出勤のとき 実勤務時間に加算
                    if (dR["休暇"].ToString() == global.eFURIDE)
                        SD.sKinmuJikan1 += SD.TimeJ;
                }

                // ⑤有休または特休取得日のとき「実勤務時間2」に加算
                if (dR["休暇"].ToString() == global.eYUKYU ||
                    dR["休暇"].ToString() == global.eKYUMU)
                    SD.sKinmuJikan2 += SD.TimeJ;

                // →半休はエラーとなるためこのロジックは通らない ⑥半休のとき半休所定勤務時間（所定勤務時間の1/2）を実勤務時間に加算
                // → 半休制度導入 2012/06/19
                if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    // 八十二リース以外のとき：2017/12/12
                    if (global._82lease == int.Parse(global.FLGOFF))
                    {
                        //SD.sKinmuJikan1 += double.Parse(dR["所定勤務時間"].ToString());
                        SD.sKinmuJikan1 += Utility.fncTimeSet(Utility.GetTimeSpan(SD.ShoteiS, SD.ShoteiE).TotalMinutes, SD.sMarume);
                    }
                    else
                    {
                        // 八十二リースのとき：2017/12/12
                        if (dR["特殊日"].ToString() != global.FLGON)
                        {
                            SD.sKinmuJikan1 += Utility.fncTimeSet(Utility.GetTimeSpan(SD.ShoteiS, SD.ShoteiE).TotalMinutes, SD.sMarume);
                        }
                        else
                        {
                            // 八十二リースは特殊日の半休が午前と午後で勤務時間が違うため以下の処理が必要：2017/12/12
                            if (dR["休暇"].ToString() == global.eAMHANKYU)
                            {
                                // 午前半休のとき午前勤務時間を加算
                                SD.sKinmuJikan1 += Utility.fncTimeSet(225, SD.sMarume);
                            }
                            else if (dR["休暇"].ToString() == global.ePMHANKYU)
                            {
                                // 午後半休のとき午後勤務時間を加算
                                SD.sKinmuJikan1 += Utility.fncTimeSet(245, SD.sMarume);
                            }
                        }
                    }
                }
            }
            else
            {
                // →半休はエラーとなるためこのロジックは通らない 時間記入なしで半休のとき半休所定勤務時間（所定勤務時間の1/2）を実勤務時間に加算
                // → 2012/06/19 半休制度導入 ただし開始、終了時間の記入は必須のためこのロジックは通らない
                //if (dR["休暇"].ToString() == global.eAMHANKYU ||
                //    dR["休暇"].ToString() == global.ePMHANKYU)
                //    SD.sKinmuJikan1 += double.Parse(dR["所定勤務時間"].ToString());
            }

            // 昼食回数１
            if (dR["昼食"].ToString() == global.FLGON) SD._sHiru1++;


            // "○"または開始時間が記入されているとき
            if (dR["特殊日"].ToString() == global.FLGON || dR["所定勤務"].ToString() == global.FLGON || _sSH != string.Empty)
            {
                SD.TimeT = 0;
                SD.sSitumu2 = 0;

                //
                // 執務時間を設定します
                //
                // 八十二リース半休取得時以外の執務なし 2017/12/09
                //

                // 八十二リース以外の半休取得時の執務時間計算
                if (global._82lease == int.Parse(global.FLGOFF))
                {
                    // 2015/05/13 午前半休のとき
                    if (dR["休暇"].ToString() == global.eAMHANKYU)
                    {
                        if (TimeS > amHankyuHayade)
                        {
                            SD.SituS = TimeS;
                            SD.SituE = TimeS.AddMinutes(540);
                        }
                        else
                        {
                            // 午後開始時刻から9時間とします
                            SD.SituS = DateTime.Parse(SD.sZanSTime);
                            SD.SituE = amHankyuHayade.AddMinutes(540);
                        }
                    }
                    else
                    {
                        // 開始時間が早出設定時刻以降のとき開始時間から9時間とします
                        if (TimeS > sHayade)
                        {
                            SD.SituS = TimeS;
                            SD.SituE = TimeS.AddMinutes(540);
                        }
                        else
                        {
                            // 早出設定時刻から9時間とします
                            SD.SituS = DateTime.Parse(SD.sZanSTime);
                            SD.SituE = sHayade.AddMinutes(540);
                        }
                    }

                    // 計算対象とする終了時間
                    e = SD.SituE;
                    if (TimeE < SD.SituE)
                    {
                        e = TimeE;
                    }

                    // 執務時間を計算
                    if (SD.ShoteiE > e)
                    {
                        SD.sSitumu2 = 0;
                    }
                    else
                    {
                        SD.sSitumu2 = Utility.GetTimeSpan(SD.ShoteiE, e).TotalMinutes;

                        // 離席時間
                        if (TimeRS != TimeNull)
                        {
                            SD.sSitumu2 -= Utility.GetRisekiTimeSpan(SD.ShoteiE, e, TimeRS, TimeRE);
                        }

                        // 午前半休のとき午後開始時間より前に出勤した場合は執務とする 2015/05/13
                        if (dR["休暇"].ToString() == global.eAMHANKYU)
                        {
                            if (TimeS < amHankyuHayade)
                            {
                                SD.sSitumu2 += Utility.GetTimeSpan(TimeS, amHankyuHayade).TotalMinutes;
                            }

                            // 午後開始時間より前の出勤中の離席時間
                            if (TimeRS != TimeNull)
                            {
                                SD.sSitumu2 -= Utility.GetRisekiTimeSpan(TimeS, amHankyuHayade, TimeRS, TimeRE);
                            }
                        }
                    }
                }
                else
                {
                    // 2018/02/02
                    //// 八十二リースで半休取得時の執務時間計算：2017/12/09
                    //if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    //{
                    //    // 所定時間を超えた時間を執務とする
                    //    double stm = Utility.GetTimeSpan(SD.ShoteiS, SD.ShoteiE).TotalMinutes;
                    //    double wtm = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes;
                    //    if (wtm > stm)
                    //    {
                    //        SD.sSitumu2 = wtm - stm;
                    //    }
                    //}
                    

                    // 八十二リースで半休取得時の執務時間計算：2018/02/02
                    if (dR["休暇"].ToString() == global.eAMHANKYU)
                    {
                        // パターン１：午前半休のとき午後開始時間より前に出勤した場合は執務とする (2015/05/13), 2018/02/02
                        if (TimeS < amHankyuHayade)
                        {
                            SD.sSitumu2 += Utility.GetTimeSpan(TimeS, amHankyuHayade).TotalMinutes;

                            // 午後開始時間より前の出勤中の離席時間
                            if (TimeRS != TimeNull)
                            {
                                SD.sSitumu2 -= Utility.GetRisekiTimeSpan(TimeS, amHankyuHayade, TimeRS, TimeRE);
                            }
                        }

                        // パターン２：午前半休のとき所定終了時刻を超えて勤務した場合は執務とする 2018/02/02
                        if (SD.ShoteiE < TimeE)
                        {
                            // 終了時間の判断
                            if (TimeE < dt2200)
                            {
                                e = TimeE;
                            }
                            else
                            {
                                e = dt2200;
                            }

                            // 執務時間を取得します
                            SD.sSitumu2 += Utility.GetTimeSpan(SD.ShoteiE, e).TotalMinutes;

                            // 離席時間を減算します
                            if (TimeRS != TimeNull)
                            {
                                SD.sSitumu2 -= Utility.GetRisekiTimeSpan(SD.ShoteiE, e, TimeRS, TimeRE);
                            }
                        }
                    }
                    else if (dR["休暇"].ToString() == global.ePMHANKYU)
                    {
                        // パターン３：午後半休のとき午前開始時間より前に出勤した場合は執務とする 2018/02/02
                        if (TimeS < SD.ShoteiS)
                        {
                            SD.sSitumu2 += Utility.GetTimeSpan(TimeS, SD.ShoteiS).TotalMinutes;

                            // 午後開始時間より前の出勤中の離席時間
                            if (TimeRS != TimeNull)
                            {
                                SD.sSitumu2 -= Utility.GetRisekiTimeSpan(TimeS, SD.ShoteiS, TimeRS, TimeRE);
                            }
                        }

                        // パターン４：午後半休のとき午前終了時間以降に勤務した場合は執務とする 2018/02/02
                        if (SD.ShoteiE < TimeE)
                        {
                            // 終了時間の判断
                            if (TimeE < dt2200)
                            {
                                e = TimeE;
                            }
                            else
                            {
                                e = dt2200;
                            }

                            // 執務時間を取得します
                            SD.sSitumu2 += Utility.GetTimeSpan(SD.ShoteiE, e).TotalMinutes;

                            // 離席時間を減算します
                            if (TimeRS != TimeNull)
                            {
                                SD.sSitumu2 -= Utility.GetRisekiTimeSpan(SD.ShoteiE, e, TimeRS, TimeRE);
                            }
                        }
                    }
                }

                SD.sSitumuT = SD.sSitumuT + Utility.fncTimeSet(SD.sSitumu2, SD.sMarume);  // 小数点10進ではなく分単位で丸め処理　2011/01/11
                
                //
                // 平常日時間外 早出残業
                //
                // ①早出残業時間を取得します
                if (TimeS < sHayade)
                {
                    SD.Hayade = Utility.GetTimeSpan(TimeS, sHayade).TotalMinutes;

                    // 早出残業時間内に離席した場合
                    if (TimeRS != TimeNull)
                        SD.Hayade -= Utility.GetRisekiTimeSpan(TimeS, sHayade, TimeRS, TimeRE);
                }

                // ②時間外残業
                if (TimeS < sHayade)
                {
                    if (dR["休憩なし"].ToString() == global.FLGON)
                        SD.ShoETime = sHayade.AddMinutes(480);
                    else SD.ShoETime = sHayade.AddMinutes(540);
                }
                else
                {
                    if (dR["休憩なし"].ToString() == global.FLGON)
                        SD.ShoETime = TimeS.AddMinutes(480);
                    else SD.ShoETime = TimeS.AddMinutes(540);
                }

                // 離席があるとき終了時刻に離席時間を加算する
                //if (TimeRS != TimeNull) SD.ShoETime += TimeSpan.Parse(SD.TimeJ2.ToString());

                // 所定終了時間より終了時間が遅いとき 終了時間までの経過時間を取得します（但し22:00mまで）
                // 八十二リース以外：2017/12/08
                if (global._82lease == int.Parse(global.FLGOFF))
                {
                    if (SD.SituE < TimeE)
                    {
                        // 終了時間の判断
                        if (TimeE < dt2200)
                        {
                            e = TimeE;
                        }
                        else
                        {
                            e = dt2200;
                        }

                        // 残業時間を取得します
                        SD.Zangyo = Utility.GetTimeSpan(SD.SituE, e).TotalMinutes;

                        // 離席時間を減算します
                        if (TimeRS != TimeNull)
                        {
                            SD.Zangyo -= Utility.GetRisekiTimeSpan(SD.SituE, e, TimeRS, TimeRE);
                        }
                    }
                }
                else
                {
                    // 八十二リース：2017/12/08
                    if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    {
                        // 半休取得時は残業時間なし
                        SD.Zangyo = 0;
                        SD.Hayade = 0;
                    }
                    else
                    {
                        if (SD.ShoteiE < TimeE)
                        {
                            // 終了時間の判断
                            if (TimeE < dt2200)
                            {
                                e = TimeE;
                            }
                            else
                            {
                                e = dt2200;
                            }

                            // 残業時間を取得します
                            SD.Zangyo = Utility.GetTimeSpan(SD.ShoteiE, e).TotalMinutes;

                            // 離席時間を減算します
                            if (TimeRS != TimeNull)
                            {
                                SD.Zangyo -= Utility.GetRisekiTimeSpan(SD.ShoteiE, e, TimeRS, TimeRE);
                            }
                        }
                    }
                }

                SD.HayaZanT += Utility.fncTimeSet(SD.Zangyo + SD.Hayade, SD.sMarume);

                // ③平常日時間外 深夜（22:00以降）
                if (TimeE > dt2200)
                {
                    // 深夜残業時間を取得します
                    SD.sSinya = Utility.GetTimeSpan(dt2200, TimeE).TotalMinutes;

                    // 離席時間を減算します
                    if (TimeRS != TimeNull)
                        SD.sSinya -= Utility.GetRisekiTimeSpan(dt2200, TimeE, TimeRS, TimeRE);

                    SD.SinyaT += Utility.fncTimeSet(SD.sSinya, SD.sMarume);
                }
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     受け渡しデータ集計項目初期化 </summary>
        /// <param name="SD">
        ///     受け渡しデータクラス</param>
        ///----------------------------------------------------------
        private void SaveDataInitial(Entity.saveData SD)
        {
            SD.sJikan1 = " 0";
            SD.sJikan2 = " 0";
            SD.sKinmu1 = "0";
            SD.sKinmu2 = "0";
            SD.sKinmu3 = "0";
            SD.sHiru1 = "0";
            SD.sHiru2 = "0";
            SD.sJigai1 = "0";
            SD.sJigai2 = "0";
            SD.sJigai3 = "0";
            SD.sJigai4 = "0";

            // 計算項目
            SD.sJituKinmu1 = 0;
            SD.sJituKinmu2 = 0;
            SD.sJituKinmu3 = 0;
            SD.sKinmuJikan1 = 0;
            SD.sKinmuJikan2 = 0;
            SD.sJituKinmu4 = 0;
            SD.sJituKinmu5 = 0;

            SD._sHiru1 = 0;
            SD._sHiru2 = 0;
            SD.sSitumu1 = 0;
            SD.sSitumu2 = 0;
            SD.sSitumuT = 0;
            SD.sHayaZan = 0;
            SD.sSinya = 0;
            SD.kSinya = 0;
            SD.sKyujitu = 0;
            SD.Hayade = 0;
            SD.Zangyo = 0;
            SD.HayaZanT = 0;
            SD.KSinyaT = 0;
            SD.SinyaT = 0;
            SD.TimeT = 0;
            SD.Situ1 = 0;
            SD.Situ2 = 0;
            SD.sRiseki = 0;
            SD.TimeJ = 0;
            SD.TimeJ2 = 0;
            
            SD.sShukinNisu = string.Empty;
            SD.sYukyuNisu = string.Empty;
            SD.sTokyuNisu = string.Empty;
            SD.sKekkinNisu = string.Empty;
            SD.sKyujituNisu = string.Empty;
            SD.sKinmuJikan = string.Empty;
            SD.sSitumuJikan = string.Empty;
            SD.sHayaZanJikan = string.Empty;
            SD.sSinyaJikan = string.Empty;
            SD.sKyujituJikan = string.Empty;
            SD.sKyujituSinya = string.Empty;
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     スタッフ受け渡しデータ集計横目初期化 </summary>
        /// <param name="SD">
        ///     受け渡しデータクラス</param>
        ///----------------------------------------------------------
        private void StaffDataInitial(Entity.saveStaff SD)
        {
            // 計算項目
            SD.sJituKinmu1 = 0;
            SD.sJituKinmu2 = 0;
            SD.sJituKinmu3 = 0;
            SD.sJituKinmu4 = 0;
            SD.sJituKinmu5 = 0;
            SD.sKinmuJikan1 = 0;
            SD.sKinmuJikan2 = 0;

            SD._sHiru1 = 0;
            SD._sHiru2 = 0;
            SD.sSitumu1 = 0;
            SD.sSitumu2 = 0;
            SD.sSitumuT = 0;
            SD.sHayaZan = 0;
            SD.sSinya = 0;
            SD.kSinya = 0;
            SD.sKyujitu = 0;
            SD.Hayade = 0;
            SD.Zangyo = 0;
            SD.HayaZanT = 0;
            SD.SinyaT = 0;
            SD.TimeT = 0;
            SD.Situ1 = 0;
            SD.Situ2 = 0;

            SD.dKinmuNisu = 0;
            SD.sKyushutuNisu = 0;
            SD.Keiyakunai = 0;
            SD.YukyuNisu = 0;
            SD.YukyuJikan = 0;
            SD.TokukyuNisu = 0;
            SD.sShoteiKyujitu = 0;
            SD.sRiseki = 0;
            SD.dKinmuJikan = 0;
            SD.dShoteiNisu = 0;
            SD.dKekkinNisu = 0;
            SD.TimeJ = 0;
            SD.TimeJ2 = 0;
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     受け渡しデータを書き出す </summary>
        /// <param name="outFile">
        ///     出力先ストリーム</param>
        /// <param name="SD">
        ///     出力データクラス</param>
        ///----------------------------------------------------------
        private void SaveDataToCsv(StreamWriter outFile, Entity.saveData SD, string ComID, string ShoID)
        {
            // 会社別出力設定データ取得
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = "select * from 出力設定 where 会社ID = ? and 職種ID = ?";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@Com", ComID);
            sCom.Parameters.AddWithValue("@Sho", ShoID);
            OleDbDataReader dR = sCom.ExecuteReader();

            if (!dR.HasRows)
            {
                SD.sShukinNisu = SD.sJituKinmu1.ToString();
                SD.sKyujituNisu = SD.sJituKinmu5.ToString();
                SD.sTokyuNisu = SD.sJituKinmu3.ToString();
                SD.sYukyuNisu = SD.sJituKinmu2.ToString();
                //SD.sKekkinNisu = SD.sJituKinmu4.ToString();

                // 時間（分単位）より時間：分(hh:mm)文字列へ変換する 2011/01/11
                SD.sKinmuJikan = Utility.fncTimehhmm(SD.sKinmuJikan1);
                SD.sSitumuJikan = Utility.fncTimehhmm(SD.sSitumuT);
                SD.sHayaZanJikan = Utility.fncTimehhmm(SD.HayaZanT);
                SD.sSinyaJikan = Utility.fncTimehhmm(SD.SinyaT);
                SD.sKyujituJikan = Utility.fncTimehhmm(SD.sKyujitu);
                SD.sKyujituSinya = Utility.fncTimehhmm(SD.KSinyaT);

                SD.sHiru1 = SD._sHiru1.ToString();
            }
            else
            {
                dR.Read();
   
                if (dR["出勤日数"].ToString() == "0") SD.sShukinNisu = "0";
                else SD.sShukinNisu = SD.sJituKinmu1.ToString();

                if (dR["休日日数"].ToString() == "0") SD.sKyujituNisu = "0";
                else SD.sKyujituNisu = SD.sJituKinmu5.ToString();

                if (dR["特休日数"].ToString() == "0") SD.sTokyuNisu = "0";
                else SD.sTokyuNisu = SD.sJituKinmu3.ToString();

                if (dR["有休日数"].ToString() == "0") SD.sYukyuNisu = "0";
                else SD.sYukyuNisu = SD.sJituKinmu2.ToString();

                if (dR["欠勤日数"].ToString() == "0") SD.sKekkinNisu = "0";

                if (dR["出勤時間"].ToString() == "0") SD.sKinmuJikan = "0";
                else SD.sKinmuJikan = Utility.fncTimehhmm(SD.sKinmuJikan1);

                if (dR["執務時間"].ToString() == "0") SD.sSitumuJikan = "0";
                else SD.sSitumuJikan = Utility.fncTimehhmm(SD.sSitumuT);

                if (dR["早出残業時間"].ToString() == "0") SD.sHayaZanJikan = "0";
                else SD.sHayaZanJikan = Utility.fncTimehhmm(SD.HayaZanT);

                if (dR["深夜時間"].ToString() == "0") SD.sSinyaJikan = "0";
                else SD.sSinyaJikan = Utility.fncTimehhmm(SD.SinyaT);

                if (dR["休日時間"].ToString() == "0") SD.sKyujituJikan = "0";
                else SD.sKyujituJikan = Utility.fncTimehhmm(SD.sKyujitu);

                if (dR["休日深夜時間"].ToString() == "0") SD.sKyujituSinya = "0";
                else SD.sKyujituSinya = Utility.fncTimehhmm(SD.KSinyaT);

                if (dR["昼食回数"].ToString() == "0") SD.sHiru1 = "0";
                else SD.sHiru1 = SD._sHiru1.ToString();
            }
            dR.Close();
            sCom.Connection.Close();

            // CSVファイルを書き出す
            StringBuilder sb = new StringBuilder();
            sb.Append(SD.sCode).Append(",");
            sb.Append(SD.sShukinNisu).Append(",");
            sb.Append(SD.sKyujituNisu).Append(",");
            sb.Append(SD.sTokyuNisu).Append(",");
            sb.Append(SD.sYukyuNisu).Append(",");
            sb.Append(SD.sKekkinNisu).Append(",");
            sb.Append(SD.sKinmuJikan).Append(",");
            sb.Append(SD.sSitumuJikan).Append(",");
            sb.Append(SD.sHayaZanJikan).Append(",");
            sb.Append(SD.sSinyaJikan).Append(",");
            sb.Append(SD.sKyujituJikan).Append(",");
            sb.Append(SD.sHiru1).Append(",");
            sb.Append(SD.sKyujituSinya);

            ////明細ファイル出力
            outFile.WriteLine(sb.ToString());
        }

        /// <summary>
        /// 登録済み過去データ削除
        /// </summary>
        private void DeleteLastData()
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = "select 個人番号 from 勤務票ヘッダ";
            
            OleDbCommand sCom2 = new OleDbCommand();
            sCom2.Connection = Con.cnOpen();
            StringBuilder sb = new StringBuilder();
            sCom2.CommandText = "delete from 履歴 where 年=? and 月=? and 個人番号=?";

            OleDbDataReader dR = null;
            
            try 
	        {
                dR = sCom.ExecuteReader();

                while (dR.Read())
	            {
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@year", global.sYear.ToString());
                    sCom2.Parameters.AddWithValue("@month", global.sMonth.ToString());
                    sCom2.Parameters.AddWithValue("@Num", dR["個人番号"].ToString());
                    sCom2.ExecuteNonQuery();
	            }
	        }
	        catch (Exception ex)
	        {
                MessageBox.Show(ex.Message, "登録済み過去データ削除エラー", MessageBoxButtons.OK);
	        }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                if (sCom2.Connection.State == ConnectionState.Open) sCom2.Connection.Close();
            }
        }

        /// <summary>
        /// 過去データ作成
        /// </summary>
        private void SaveLastData()
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = "select 個人番号 from 勤務票ヘッダ";

            OleDbCommand sCom2 = new OleDbCommand();
            sCom2.Connection = Con.cnOpen();
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("insert into 履歴 (");
            sb.Append("年,月,個人番号,氏名) values ");
            sb.Append("(?,?,?,?)");
            sCom2.CommandText = sb.ToString();

            OleDbDataReader dR = null;

            try
            {
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@year", global.sYear.ToString());
                    sCom2.Parameters.AddWithValue("@month", global.sMonth.ToString());
                    sCom2.Parameters.AddWithValue("@Num", dR["個人番号"].ToString());
                    sCom2.Parameters.AddWithValue("@Name", string.Empty);
                    sCom2.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "過去データ作成エラー", MessageBoxButtons.OK);
            }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                if (sCom2.Connection.State == ConnectionState.Open) sCom2.Connection.Close();
            }
        }

        /// <summary>
        /// 勤務票レコード削除
        /// </summary>
        private void DelHeadRec()
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;            

            try
            {
                // 勤務票ヘッダ削除
                sCom.CommandText = "delete from 勤務票ヘッダ";
                sCom.ExecuteNonQuery();

                // 勤務票明細削除
                sCom.CommandText = "delete from 勤務票明細";
                sCom.ExecuteNonQuery();

                // トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票データ削除エラー", MessageBoxButtons.OK);
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     画像・CSVファイル退避処理 </summary>
        /// <param name="tifPath">
        ///     退避先フォルダパス</param>
        ///----------------------------------------------------------
        private void tifFileMove(string tifPath)
        {
            //移動先フォルダがあるか？なければ作成する（TIFフォルダ）
            if (!System.IO.Directory.Exists(tifPath)) System.IO.Directory.CreateDirectory(tifPath);

            //画像を退避先フォルダへ移動する            
            foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.tif"))
            {
                File.Move(files, tifPath + @"\" + System.IO.Path.GetFileName(files));
            }

            //CSVファイルを退避先フォルダへ移動する　→　実際この時点でCSVファイルはフォルダ内に存在しない        
            foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
            {
                File.Move(files, tifPath + @"\" + System.IO.Path.GetFileName(files));
            }

            //Tifフォルダ内の全てのファイルを削除します            
            foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.*"))
            {
                File.Delete(files);
            }
        }

        /// <summary>
        /// 設定月数分経過した過去画像を削除する    
        /// </summary>
        private void imageDelete(int sdel)
        {
            //削除月設定が0のとき、「過去画像削除しない」とみなし終了する
            if (sdel == 0) return;

            try
            {
                //削除年月の取得
                DateTime delDate = DateTime.Today.AddMonths(sdel * (-1));
                int _dYY = delDate.Year;            //基準年
                int _dMM = delDate.Month;           //基準月
                int _dYYMM = _dYY * 100 + _dMM;     //基準年月
                int _DataYYMM;
                string fileYYMM;

                // 設定月数分経過した過去画像を削除する
                // ①スタッフ
                foreach (string files in System.IO.Directory.GetFiles(global.sTIF, "*.tif"))
                {
                    //ファイル名より年月を取得する
                    fileYYMM = System.IO.Path.GetFileName(files).Substring(0, 6);

                    if (Utility.NumericCheck(fileYYMM))
                    {
                        _DataYYMM = int.Parse(fileYYMM);

                        //基準年月以前なら削除する
                        if (_DataYYMM <= _dYYMM) File.Delete(files);
                    }
                }

                // ②パートタイマー
                foreach (string files in System.IO.Directory.GetFiles(global.sTIF, "*.tif"))
                {
                    //ファイル名より年月を取得する
                    fileYYMM = System.IO.Path.GetFileName(files).Substring(0, 6);

                    if (Utility.NumericCheck(fileYYMM))
                    {
                        _DataYYMM = int.Parse(fileYYMM);

                        //基準年月以前なら削除する
                        if (_DataYYMM <= _dYYMM) File.Delete(files);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("過去画像削除中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
                return;
            }
        }

        /// <summary>
        /// 設定月数を経過したバックアップファイルを削除します
        /// </summary>
        /// <param name="sSpan">経過月数</param>
        /// <param name="sDir">ディレクトリ</param>
        private void delBackUpFiles(int sSpan, string sDir)
        {
            if (System.IO.Directory.Exists(sDir))
            {
                if (sSpan > 0)
                {
                    foreach (string fName in System.IO.Directory.GetFiles(sDir, "*.tif"))
                    {
                        string f = System.IO.Path.GetFileName(fName);

                        // ファイル名の長さを検証（日付情報あり？）
                        if (f.Length > 12)
                        {
                            // ファイル名から日付部分を取得します
                            DateTime dt;
                            string stDt = f.Substring(0, 4) + "/" + f.Substring(4, 2) + "/" + f.Substring(6, 2);
                            if (DateTime.TryParse(stDt, out dt))
                            {
                                // 設定月数を加算した日付を取得します
                                DateTime Fdt = dt.AddMonths(sSpan);

                                // 今日の日付と比較して設定月数を加算したファイル日付が既に経過している場合、ファイルを削除します
                                if (DateTime.Today.CompareTo(Fdt) == 1)
                                {
                                    System.IO.File.Delete(fName);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// MDB明細データの件数をカウントする
        /// </summary>
        /// <returns>レコード件数</returns>
        private int CountMDBitem()
        {
            int rCnt = 0;

            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();
            OleDbDataReader dR;

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("SELECT 勤務票ヘッダ.*,勤務票明細.* from ");
            sb.Append("勤務票ヘッダ inner join 勤務票明細 ");
            sb.Append("on 勤務票ヘッダ.ID = 勤務票明細.ヘッダID ");
            sb.Append("order by 勤務票ヘッダ.ID,勤務票明細.ヘッダID ");
            sCom.CommandText = sb.ToString();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数加算
                rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            return rCnt;
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Tag.ToString() != END_MAKEDATA)
            {
                if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                // カレントデータ更新
                CurDataUpDate(_cI);
            }

            // 解放する
            this.Dispose();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            //フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        /// <summary>
        /// MDBファイルを最適化する
        /// </summary>
        private void mdbCompact()
        {
            try
            {
                JRO.JetEngine jro = new JRO.JetEngine();

                string OldDb = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBFILENAME;

                string NewDb = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBTEMPFILE;

                // 最適化した一時MDBを作成する
                jro.CompactDatabase(OldDb, NewDb);

                //今までのバックアップファイルを削除する
                System.IO.File.Delete(Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBBACKUP);

                //今までのファイルをバックアップとする
                System.IO.File.Move(Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBFILENAME,
                                    Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBBACKUP);

                //一時ファイルをMDBファイルとする
                System.IO.File.Move(Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBTEMPFILE,
                                    Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBFILENAME);
            }
            catch (Exception e)
            {
                MessageBox.Show("MDB最適化中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        private void txtNo_Leave(object sender, EventArgs e)
        {
            txtNo.Text = txtNo.Text.PadLeft(7, '0');
            GetShoteiData(txtNo.Text);
        }

        private void txtNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space)
            {
                frmStaffSelect frmS = new frmStaffSelect(mst);
                frmS.ShowDialog();
                if (frmS._msCode != string.Empty)
                {
                    txtNo.Text = frmS._msCode;
                    GetShoteiData(txtNo.Text);
                }
                frmS.Dispose();
            }
        }

        /// <summary>
        /// 勤務票ヘッダデータと勤務表明細データに新規レコードを追加する
        /// </summary>
        private void AddNewData(int usr)
        {
            // IDを取得します
            string _ID = string.Format("{0:0000}", DateTime.Today.Year) + 
                         string.Format("{0:00}", DateTime.Today.Month) + 
                         string.Format("{0:00}", DateTime.Today.Day) + 
                         string.Format("{0:00}", DateTime.Now.Hour) + 
                         string.Format("{0:00}", DateTime.Now.Minute) + 
                         string.Format("{0:00}", DateTime.Now.Second) + "001";

            // データベースへ接続します
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                // 勤務票ヘッダテーブル追加登録
                StringBuilder sb = new StringBuilder();
                sb.Clear();
                sb.Append("insert into 勤務票ヘッダ ");
                sb.Append("(ID,年,月,データ区分) values (?,?,?,?)");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@id", _ID);
                sCom.Parameters.AddWithValue("@year", global.sYear);
                sCom.Parameters.AddWithValue("@month", global.sMonth);
                sCom.Parameters.AddWithValue("@kbn", usr);
                sCom.ExecuteNonQuery();

                // 勤務票明細テーブル追加登録
                sb.Clear();
                sb.Append("insert into 勤務票明細 ");
                sb.Append("(ヘッダID, 日付, 休日) values (?,?,?)");
                sCom.CommandText = sb.ToString();

                int sDays = 1;
                DateTime dt;

                string tempDt = global.sYear + "/" + global.sMonth + "/" + sDays.ToString();

                // 存在する日付のときにMDBへ登録する
                while (DateTime.TryParse(tempDt, out dt))
                {
                    sCom.Parameters.Clear();

                    // ヘッダID
                    sCom.Parameters.AddWithValue("@ID", _ID);

                    // 日付
                    sCom.Parameters.AddWithValue("@Days", sDays);

                    // 休日区分の設定
                    // ①曜日で判断します
                    string Youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                    int sHol = 0;
                    if (Youbi == "土")
                    {
                        sHol = global.hSATURDAY;
                    }
                    else if (Youbi == "日")
                    {
                        sHol = global.hHOLIDAY;
                    }
                    else
                    {
                        sHol = global.hWEEKDAY;
                    }

                    // ②休日テーブルを参照し休日に該当するか調べます
                    SysControl.SetDBConnect dc = new SysControl.SetDBConnect();
                    OleDbCommand sCom2 = new OleDbCommand();
                    sCom2.Connection = dc.cnOpen();
                    OleDbDataReader dr = null;
                    sCom2.CommandText = "select * from 休日 where 年=? and 月=? and 日=? and 会社ID=?";
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@year", global.sYear);
                    sCom2.Parameters.AddWithValue("@Month", global.sMonth);
                    sCom2.Parameters.AddWithValue("@day", sDays);
                    sCom2.Parameters.AddWithValue("@id", global.pblComNo.ToString("00"));
                    dr = sCom2.ExecuteReader();
                    if (dr.Read())
                    {
                        sHol = global.hHOLIDAY;
                    }
                    dr.Close();
                    sCom2.Connection.Close();

                    sCom.Parameters.AddWithValue("@Hol", sHol);

                    // テーブル書き込みます
                    sCom.ExecuteNonQuery();

                    // 日付をインクリメントします
                    sDays++;
                    tempDt = global.sYear + "/" + global.sMonth + "/" + sDays.ToString();
                }

                // トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票データ新規登録処理", MessageBoxButtons.OK);
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の勤務票データを削除します。よろしいですか", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // レコード・画像を削除する
            DataDelete(_cI);

            //テーブル件数カウント：ゼロならばプログラム終了
            if (CountMDB() == 0)
            {
                MessageBox.Show("全ての勤務票データが削除されました。処理を終了します。", "勤務票削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //終了処理
                Environment.Exit(0);
            }

            //テーブルデータキー項目読み込み
            inData = LoadMdbID();

            //エラー情報初期化
            ErrInitial();

            //レコードを表示
            if (inData.Length - 1 < _cI) _cI = inData.Length - 1;

            DataShow(_cI, inData, this.dg1);
        }

        /// <summary>
        /// カレント勤務票データを削除します
        /// </summary>
        /// <param name="iX">勤務票データインデックス</param>
        private void DataDelete(int iX)
        {
            //カレントデータを削除します
            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();

            // 画像ファイル名を取得します
            string sImgNm = string.Empty;
            OleDbDataReader dR;
            sCom.CommandText = "select 画像名 from 勤務票ヘッダ where ID = ?";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@ID", inData[iX]._sID);
            dR = sCom.ExecuteReader();
            while (dR.Read())
            {
                sImgNm = dR["画像名"].ToString();
            }
            dR.Close();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                //勤務票ヘッダデータを削除します
                sCom.CommandText = "delete from 勤務票ヘッダ where ID = ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", inData[iX]._sID);
                sCom.ExecuteNonQuery();

                //勤務票明細データを削除します
                sCom.CommandText = "delete from 勤務票明細 where ヘッダID = ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", inData[iX]._sID);
                sCom.ExecuteNonQuery();

                //画像ファイルを削除します
                if (System.IO.File.Exists(_InPath + sImgNm)) System.IO.File.Delete(_InPath + sImgNm);

                //ＣＳＶファイル退避先パスを取得します
                string csvFnm = global.sTIFCOM + System.IO.Path.GetFileNameWithoutExtension(sImgNm) + ".csv";
                
                //ＣＳＶファイルを削除します
                if (System.IO.File.Exists(csvFnm)) System.IO.File.Delete(csvFnm);

                // トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("勤務表の削除に失敗しました" + Environment.NewLine + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     社員マスターと職種を配列に取得します </summary>
        /// <param name="dbName">
        ///     指定会社のデータベース名</param>
        ///----------------------------------------------------------------------
        private void LoadBugyoMst(string dbName)
        {
            // 勘定奉行データベース接続文字列を取得する 2016/10/12
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //string sc = Utility.GetDBConnect(dbName);

            SqlControl.DataControl sdcon = new SqlControl.DataControl(sc);

            //データリーダーを取得する
            SqlDataReader dr;
            string mySQL = string.Empty;
            
            //'------------------------------------------
            //'   会社マスタからデータを取得
            //'
            //'       2010/10/08
            //'       給与奉行iシリーズから取得
            //'------------------------------------------

            // テーブルを開く
            mySQL = string.Empty;

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
            mySQL += "where AnnounceDate <= '" + global.sYear.ToString() + "/" + global.sMonth.ToString() + "/01' ";
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
                //   給与奉行iシリーズ対応　2010/10/11
                //--------------------------------------------------
                if (int.Parse(dr["zaisekikbn"].ToString()) != 2)     // iシリーズでは"2"が退職
                {
                    // 退職予定日が環境設定年月の月初日より前のとき対象外とする 2010/10/13
                    yy = DateTime.Parse(dr["RetireCorpScheduleDate"].ToString()).Year.ToString();
                    mm = DateTime.Parse(dr["RetireCorpScheduleDate"].ToString()).Month.ToString().PadLeft(2, '0');
                    dd = DateTime.Parse(dr["RetireCorpScheduleDate"].ToString()).Day.ToString().PadLeft(2, '0');

                    string rsDate = Utility.SeitoWa(yy) + mm + dd;

                    if (int.Parse(rsDate) >= int.Parse(Utility.SeitoWa(global.sYear.ToString()) + global.sMonth.ToString().PadLeft(2, '0') + "01"))
                    {
                        if (sReBuf != dr["EmployeeID"].ToString().Trim())
                        {
                            // 契約開始日が環境設定年月の月初日以前を対象とする 2010/10/13
                            dt = Utility.fcKikan(dr["ContractStartDate"].ToString());

                            // 昭和に対応　2012/08/26
                            if (dt != string.Empty)
                            {
                                if (dt.Substring(0, 1) == "H")
                                    date1 = dt.Substring(1, 2) + dt.Substring(4, 2) + dt.Substring(7, 2);
                                else date1 = "00" + dt.Substring(4, 2) + dt.Substring(7, 2); // 昭和のときは一時的に年は0にする
                            }
                            else date1 = "0";

                            //if (dt != string.Empty) date1 = dt.Substring(1, 2) + dt.Substring(4, 2) + dt.Substring(7, 2);
                            //else date1 = "0";

                            date2 = Utility.SeitoWa(global.sYear.ToString()) + global.sMonth.ToString("00") + "01";

                            if (int.Parse(date1) <= int.Parse(date2))
                            {
                                //2件目以降なら要素数を追加
                                if (iX != 0) mst.CopyTo(mst = new Entity.ShainMst[iX + 1], 0);

                                mst[iX].sCode = Utility.subStringRight(dr["EmployeeNo"].ToString(), 7);
                                mst[iX].sName = dr["Name"].ToString().Trim();
                                mst[iX].sKana = dr["NameKana"].ToString().Trim();
                                mst[iX].sTenpoCode = Utility.subStringRight(dr["DepartmentCode"].ToString(), 3);
                                mst[iX].sTenpoName = dr["DepartmentName"].ToString().Trim();
                                mst[iX].sKeiyakuS = Utility.fcKikan(dr["ContractStartDate"].ToString());
                                mst[iX].sKeiyakuE = Utility.fcKikan(dr["ContractEndDate"].ToString().Trim());
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
                            // 昭和に対応　2012/08/26
                            if (mst[iX - 1].sKeiyakuS != string.Empty)
                            {
                                if (mst[iX - 1].sKeiyakuS.Substring(0, 1) == "H")
                                    date1 = mst[iX - 1].sKeiyakuS.Substring(1, 2) + mst[iX - 1].sKeiyakuS.Substring(4, 2) + mst[iX - 1].sKeiyakuS.Substring(7, 2);
                                else date1 = "00" + mst[iX - 1].sKeiyakuS.Substring(4, 2) + mst[iX - 1].sKeiyakuS.Substring(7, 2); // 昭和のときは一時的に年は0にする

                                //date1 = mst[iX - 1].sKeiyakuS.Substring(1, 2) + mst[iX - 1].sKeiyakuS.Substring(4, 2) + mst[iX - 1].sKeiyakuS.Substring(7, 2);
                            }
                            else date1 = "0";

                            dt = Utility.fcKikan(dr["ContractStartDate"].ToString());

                            // 昭和に対応　2012/08/26
                            if (dt != string.Empty)
                            {
                                if (dt.Substring(0, 1) == "H")
                                    date2 = dt.Substring(1, 2) + dt.Substring(4, 2) + dt.Substring(7, 2);
                                else date2 = "00" + dt.Substring(4, 2) + dt.Substring(7, 2); // 昭和のときは一時的に年は0にする
                            }
                            else date2 = "0";

                            //if (dt != string.Empty) 
                            //    date2 = dt.Substring(1, 2) + dt.Substring(4, 2) + dt.Substring(7, 2);
                            //else date2 = "0";

                            if (int.Parse(date1) < int.Parse(date2))
                            {
                                // 2010/10/13  契約開始日が環境設定年月の月初日以前を対象とする
                                dt = Utility.fcKikan(dr["ContractStartDate"].ToString());

                                // 昭和に対応　2012/08/26
                                if (dt != string.Empty)
                                {
                                    if (dt.Substring(0, 1) == "H")
                                        date1 = dt.Substring(1, 2) + dt.Substring(4, 2) + dt.Substring(7, 2);
                                    else date1 = "00" + dt.Substring(4, 2) + dt.Substring(7, 2); // 昭和のときは一時的に年は0にする
                                }
                                else date1 = "0";

                                //if (dt != string.Empty) date1 = dt.Substring(1, 2) + dt.Substring(4, 2) + dt.Substring(7, 2);
                                //else date1 = "0";

                                if (int.Parse(date1) <= int.Parse(Utility.SeitoWa(global.sYear.ToString()) + global.sMonth.ToString("00") + "01"))
                                {
                                    mst[iX - 1].sCode = Utility.subStringRight(dr["EmployeeNo"].ToString(), 7);
                                    mst[iX - 1].sName = dr["Name"].ToString().Trim();
                                    mst[iX - 1].sKana = dr["NameKana"].ToString().Trim();
                                    mst[iX - 1].sTenpoCode = Utility.subStringRight(dr["DepartmentCode"].ToString(), 3);
                                    mst[iX - 1].sTenpoName = dr["DepartmentName"].ToString().Trim();
                                    mst[iX - 1].sKeiyakuS = Utility.fcKikan(dr["ContractStartDate"].ToString());
                                    mst[iX - 1].sKeiyakuE = Utility.fcKikan(dr["ContractEndDate"].ToString().Trim());
                                    mst[iX - 1].sWorkTimeS = Utility.fcTime(dr["WeeklyContractTimeSlot1Start"].ToString());
                                    mst[iX - 1].sWorkTimeE = Utility.fcTime(dr["WeeklyContractTimeSlot1End"].ToString());
                                    mst[iX - 1].sWorkTimeS2 = Utility.fcTime(dr["WeeklyContractTimeSlot2Start"].ToString());
                                    mst[iX - 1].sWorkTimeE2 = Utility.fcTime(dr["WeeklyContractTimeSlot2End"].ToString());
                                    mst[iX - 1].sSID = Utility.subStringRight(dr["SalarySystemCode"].ToString(), 2).PadLeft(2, '0');
                                    mst[iX - 1].sSName = dr["SalarySystemName"].ToString().Trim();
                                }
                            }
                        }
                    }
                }
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

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     社員のエラーチェック（職種ID=05,06）</summary>
        /// <param name="cdR">
        ///     勤務票ヘッダデータリーダー</param>
        /// <param name="dR">
        ///     勤務票明細データリーダー</param>
        /// <param name="iX">
        ///     勤務票明細日付</param>
        /// <returns>
        ///     true:エラーなし、false:エラー有り</returns>
        ///----------------------------------------------------------------------------
        private bool ChkMeisai05(OleDbDataReader cdR, OleDbDataReader dR, int iX)
        {
            DateTime ssS;   // 所定開始または特殊勤務開始時間
            DateTime ssE;   // 所定終了または特殊勤務終了時間
    
            // 所定時間マーク有り、特殊日マークなし：週契約時間帯１を摘要
            if (dR["特殊日"].ToString() == global.FLGOFF && dR["所定勤務"].ToString() == global.FLGON)
            {
                ssS = DateTime.Parse(cdR["所定開始時間"].ToString());
                ssE = DateTime.Parse(cdR["所定終了時間"].ToString());
            }        
            // 所定時間マーク有り、特殊日マーク有り：週契約時間帯２を摘要
            else if (dR["特殊日"].ToString() == global.FLGON && dR["所定勤務"].ToString() == global.FLGON)
            {
                ssS = DateTime.Parse(cdR["特殊開始時間"].ToString());
                ssE = DateTime.Parse(cdR["特殊終了時間"].ToString());
            }

            // その他：週契約時間帯１を摘要
            else
            {
                ssS = DateTime.Parse(cdR["所定開始時間"].ToString());
                ssE = DateTime.Parse(cdR["所定終了時間"].ToString());
            }
            
            // 土日以外で全て無記入のときは欠勤扱い
            if (dR["特殊日"].ToString() == global.FLGOFF && dR["所定勤務"].ToString() == global.FLGOFF && 
                dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() == string.Empty && 
                dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() == string.Empty && 
                dR["休暇"].ToString() == string.Empty && dR["休憩なし"].ToString() == global.FLGOFF && 
                dR["昼食"].ToString() == global.FLGOFF && 
                dR["離席開始時"].ToString() == string.Empty && dR["離席開始分"].ToString() == string.Empty && 
                dR["離席終了時"].ToString() == string.Empty && dR["離席終了分"].ToString() == string.Empty &&
                dR["休日"].ToString() == global.hWEEKDAY.ToString())
            {
                // 欠勤日数を加算
                global.KekkinCnt++;
                return true;
            }
            
            // 休日に半休が記入されている 2015/05/13
            if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
            {
                // 休日のとき
                if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休日に半休は記入できません";
                    return false;
                }
            }

            // 休日に有休または特休が記入されている 2015/05/13
            if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU)
            {
                // 休日のとき 2015/05/13
                if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休日に有休・特休は記入できません";
                    return false;
                }
            }

            // 開始時間、終了時間いずれかに記入があるとき
            if (dR["開始時"].ToString() != string.Empty || dR["開始分"].ToString() != string.Empty ||
                dR["終了時"].ToString() != string.Empty || dR["終了分"].ToString() != string.Empty)
            {
                if (!chkMeisai05_Sub01(cdR, dR, iX)) return false;
            }
            else
            {
                // 半休のとき開始時間、終了時間は必須とする 2012/06/19
                if (dR["休暇"].ToString() == global.eAMHANKYU ||
                    dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "半休のときは開始時間、終了時間は必須項目です";
                    return false;
                }
                else if (!chkMeisai05_Sub02(cdR, dR, iX, ssS, ssE)) return false;
            }

            // 休暇 2012/06/19
            if (dR["休暇"].ToString() != string.Empty)
            {
                if (int.Parse(dR["休暇"].ToString()) < int.Parse(global.eYUKYU) ||
                    int.Parse(dR["休暇"].ToString()) > int.Parse(global.ePMHANKYU))
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休暇欄が不正です";
                    return false;
                }
            }

            // 昼食
            if (dR["昼食"].ToString() == global.FLGON)
            {
                if (dR["所定勤務"].ToString() == global.FLGOFF && 
                    dR["開始時"].ToString() == string.Empty)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "時刻を入力してください";
                    return false;
                }

                // 半休のときはエラーとする 2012/06/19
                if (dR["休暇"].ToString() == global.eAMHANKYU ||
                    dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "半休で昼食はできません";
                    return false;
                }

                // 休憩なし記入 2015/08/17
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = "昼食と休憩マークは両方付けられません";
                    return false;
                }
            }
    
            // 半休はエラーとする 2010/12/03
            // 2012/06/19 半休制度導入しました
            //if (dR["休暇"].ToString() == global.eAMHANKYU || 
            //    dR["休暇"].ToString() == global.ePMHANKYU)
            //{
            //    global.errNumber = global.eKyuka;
            //    global.errMsg = "半休の入力はできません";
            //    return false;
            //}

            // 半休のとき 2012/06/19
            if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
            {
                // 休日のとき 2015/05/13
                if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休日に半休は記入できません";
                    return false;
                }

                if (dR["所定勤務"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eSMark;
                    global.errMsg = "半休なのでマークは付けられません";
                    return false;
                }
            }

            // 有休または特休のとき
            if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU)
            {
                // 休日のとき 2015/05/13
                if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休日に有休・特休は記入できません";
                    return false;
                }
        
                if (dR["所定勤務"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eSMark;
                    global.errMsg = "有休・特休なのでマークは付けられません";
                    return false;
                }

                if (dR["開始時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }
            
                if (dR["開始分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }
            
                if (dR["終了時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }
            
                if (dR["終了分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }
            
                if (dR["昼食"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eHiru1;
                    global.errMsg = "有休・特休なので昼は付けられません";
                    return false;
                }
            
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = "有休・特休なので休憩マークは付けられません";
                    return false;
                }
            
                if (dR["離席開始時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }
            
                if (dR["離席開始分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eRSM;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }
            
                if (dR["離席終了時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }
            
                if (dR["離席終了分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eREM;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }
            }

            // 休日だった場合
            if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
            {
                if (dR["所定勤務"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eSMark;
                    global.errMsg = "休日出勤なので「○」は付けられません";
                    return false;
                }
                
                if (dR["特殊日"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "休日出勤なので「○」は付けられません";
                    return false;
                }

                if (dR["昼食"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eHiru1;
                    global.errMsg = "休日出勤なので昼食は無効です";
                    return false;
                }
            }

            // 振休あり出勤だった場合
            if (dR["休暇"].ToString() == global.eFURIDE)
            {
                // 特殊日マークあり
                if (dR["特殊日"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "振休あり出勤なので「○」は付けられません";
                    return false;
                }

                // 所定勤務マークなし
                if (dR["所定勤務"].ToString() == global.FLGOFF)
                {
                    if (dR["開始時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eSH;
                        global.errMsg = "開始時刻を入力してください";
                        return false;
                    }
                    if (dR["開始分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eSM;
                        global.errMsg = "開始時刻を入力してください";
                        return false;
                    }
                    if (dR["終了時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eEH;
                        global.errMsg = "終了時刻を入力してください";
                        return false;
                    }

                    if (dR["終了時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eEM;
                        global.errMsg = "終了時刻を入力してください";
                        return false;
                    }
                }

                // 離席時間記入あり
                if (dR["離席開始時"].ToString() != string.Empty || dR["離席開始分"].ToString() != string.Empty ||
                    dR["離席終了時"].ToString() != string.Empty || dR["離席終了分"].ToString() != string.Empty)
                {
                    if (dR["離席開始時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eRSH;
                        global.errMsg = "離席開始時刻が不正です";
                        return false;
                    }

                    if (dR["離席開始分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eRSM;
                        global.errMsg = "離席開始時刻が不正です";
                        return false;
                    }

                    if (dR["離席終了時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eREH;
                        global.errMsg = "離席終了時刻が不正です";
                        return false;
                    }

                    if (dR["離席終了分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eREM;
                        global.errMsg = "離席終了時刻が不正です";
                        return false;
                    }
                }
            
                // 所定勤務マークなし
                if (dR["所定勤務"].ToString() == global.FLGOFF)
                {

                    // 開始時間と終了時間
                    int stm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());
                    int etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

                    if (stm >= etm)
                    {
                        global.errNumber = global.eEH;
                        global.errMsg = "終了時間が開始時間以前となっています";
                        return false;
                    }

                    // 離席時間に記入あり
                    if (dR["離席開始時"].ToString() != string.Empty && dR["離席終了時"].ToString() != string.Empty)
                    {
                        // 離席開始時間と離席終了時間
                        stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                        etm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());

                        if (stm >= etm)
                        {
                            global.errNumber = global.eREH;
                            global.errMsg = "離席終了時間が離席開始時間以前となっています";
                            return false;
                        }

                        // 開始時間と離席開始時間
                        stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                        etm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());

                        if (stm <= etm)
                        {
                            global.errNumber = global.eRSH;
                            global.errMsg = "離席開始時間が開始時間以前となっています";
                            return false;
                        }

                        // 終了時間と離席終了時間と離席開始時間
                        stm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());
                        etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

                        if (stm >= etm)
                        {
                            global.errNumber = global.eREH;
                            global.errMsg = "離席終了時間が終了時間を超えています";
                            return false;
                        }
                    }
                }
            }

            // 振替休日のとき
            if (dR["休暇"].ToString() == global.eFURIKYU)
            {
                if (dR["所定勤務"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eSMark;
                    global.errMsg = "振替休日なので「○」は付けられません";
                    return false;
                }
                
                if (dR["特殊日"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "振替休日なので「○」は付けられません";
                    return false;
                }
                
                if (dR["開始時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }
                
                if (dR["開始分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["終了時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["終了分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }
                
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = "振替休日なので休憩は付けられません";
                    return false;
                }

                if (dR["昼食"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eHiru1;
                    global.errMsg = "振替休日なので昼は付けられません";
                    return false;
                }
            }

            // 訂正欄チェック状態のときエラーとする 2015/08/17
            if (dR["訂正"].ToString() == global.FLGON)
            {
                global.errNumber = global.eTeisei;
                global.errMsg = "訂正欄がチェックされています";
                return false;
            }

            return true;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     社員の開始時間、終了時間いずれかに記入があるときのチェック </summary>
        /// <param name="cdR">
        ///     勤務票ヘッダデータリーダー</param>
        /// <param name="dR">
        ///     勤務票明細データリーダー</param>
        /// <param name="iX">
        ///     日付</param>
        /// <returns>
        ///     </returns>
        ///---------------------------------------------------------------------------------
        private bool chkMeisai05_Sub01(OleDbDataReader cdR, OleDbDataReader dR, int iX)
        {
            TimeSpan ts;

            // 開始時間 数字以外または範囲外
            if (!Utility.NumericCheck(dR["開始時"].ToString()) ||
                int.Parse(dR["開始時"].ToString()) < 0 || int.Parse(dR["開始時"].ToString()) > 23)
            {
                global.errNumber = global.eSH;
                global.errMsg = "開始時刻が不正です";
                return false;
            }

            // 開始分 数字以外または範囲外
            if (!Utility.NumericCheck(dR["開始分"].ToString()) ||
                int.Parse(dR["開始分"].ToString()) < 0 || int.Parse(dR["開始分"].ToString()) > 59)
            {
                global.errNumber = global.eSM;
                global.errMsg = "開始時刻が不正です";
                return false;
            }

            // 終了時間 数字以外または範囲外
            if (!Utility.NumericCheck(dR["終了時"].ToString()) ||
                int.Parse(dR["終了時"].ToString()) < 0 || int.Parse(dR["終了時"].ToString()) > 23)
            {
                global.errNumber = global.eEH;
                global.errMsg = "終了時刻が不正です";
                return false;
            }

            // 終了分 数字以外または範囲外
            if (!Utility.NumericCheck(dR["終了分"].ToString()) ||
                int.Parse(dR["終了分"].ToString()) < 0 || int.Parse(dR["終了分"].ToString()) > 59)
            {
                global.errNumber = global.eEM;
                global.errMsg = "終了時刻が不正です";
                return false;
            }
            
            // 離席開始時間 数字以外
            if (dR["離席開始時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席開始時"].ToString()) ||
                    int.Parse(dR["離席開始時"].ToString()) < 0 || int.Parse(dR["離席開始時"].ToString()) > 23)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }
            }

            // 離席開始分 数字以外
            if (dR["離席開始分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席開始分"].ToString()) ||
                    int.Parse(dR["離席開始分"].ToString()) < 0 || int.Parse(dR["離席開始分"].ToString()) > 59)
                {
                    global.errNumber = global.eRSM;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }
            }

            // 離席終了時間 数字以外
            if (dR["離席終了時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席終了時"].ToString()) ||
                    int.Parse(dR["離席終了時"].ToString()) < 0 || int.Parse(dR["離席終了時"].ToString()) > 23)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時が不正です";
                    return false;
                }
            }

            // 離席終了分 数字以外
            if (dR["離席終了分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席終了分"].ToString()) ||
                    int.Parse(dR["離席終了分"].ToString()) < 0 || int.Parse(dR["離席終了分"].ToString()) > 59)
                {
                    global.errNumber = global.eREM;
                    global.errMsg = "離席終了時刻が不正です";
                    return false;
                }
            }
            
            // 離席時間記入
            if (dR["離席開始時"].ToString() != string.Empty || dR["離席開始分"].ToString() != string.Empty || 
                dR["離席終了時"].ToString() != string.Empty || dR["離席終了分"].ToString() != string.Empty)
            {
                if (dR["離席開始時"].ToString() == string.Empty)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }

                if (dR["離席開始分"].ToString() == string.Empty)
                {
                    global.errNumber = global.eRSM;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }

                if (dR["離席終了時"].ToString() == string.Empty)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時刻が不正です";
                    return false;
                }

                if (dR["離席終了分"].ToString() == string.Empty)
                {
                    global.errNumber = global.eREM;
                    global.errMsg = "離席終了時刻が不正です";
                    return false;
                }
            }

            // 開始時間と終了時間
            int stm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());
            int etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

            if (stm < 600)
            {
                global.errNumber = global.eSH;
                global.errMsg = "開始時刻が6時以前になっています";
                return false;
            }

            if (stm >= etm)
            {
                global.errNumber = global.eEH;
                global.errMsg = "終了時間が開始時間以前となっています";
                return false;
            }
            
            // 離席時間に記入あり
            if (dR["離席開始時"].ToString() != string.Empty && dR["離席終了時"].ToString() != string.Empty)
            {
                // 離席開始時間と離席終了時間
                stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                etm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());

                if (stm >= etm)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時間が離席開始時間以前となっています";
                    return false;
                }

                // 開始時間と離席開始時間
                stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                etm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());

                if (stm <= etm)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時間が開始時間以前となっています";
                    return false;
                }

                // 終了時間と離席終了時間と離席開始時間
                stm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());
                etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

                if (stm >= etm)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時間が終了時間を超えています";
                    return false;
                }
            }

            // 所定勤務
            if (dR["所定勤務"].ToString() == global.FLGON)
            {
                global.errNumber = global.eSMark;
                global.errMsg = "時間とマークが両方入力されています";
                return false;
            }
                            
            // 勤務時間数を計算する　2010/12/04
            DateTime sHM = DateTime.Parse(dR["開始時"].ToString() + ":" + dR["開始分"].ToString());
            DateTime eHM = DateTime.Parse(dR["終了時"].ToString() + ":" + dR["終了分"].ToString());
            ts = Utility.GetTimeSpan(sHM,eHM);

            // 半休のとき 2012/06/19
            if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
            {
                // 勤務時間が6H超のときエラーとする
                if (ts.TotalMinutes > 360)
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "半休で勤務時間が6hを超えています";
                    return false;
                }

                // 休憩なしマーク
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = "半休で休憩なしは付けられません";
                    return false;
                }

                // 午前半休終了時間 2015/08/17
                DateTime amEtm = DateTime.Parse(cdR["午前半休終了時間"].ToString());
                DateTime pmStm = DateTime.Parse(cdR["午後半休開始時間"].ToString());

                // 同　特殊日のとき 2015/08/17
                if (dR["特殊日"].ToString() == global.FLGON)
                {
                    amEtm = DateTime.Parse(cdR["特殊午前半休終了時間"].ToString());
                    pmStm = DateTime.Parse(cdR["特殊午後半休開始時間"].ToString());
                }

                // 午前半休で午前中の勤務が記入されていたらエラー 2015/08/17
                if (dR["休暇"].ToString() == global.eAMHANKYU)
                {
                    // 開始時刻が午前半休終了時刻前のとき
                    if (sHM.CompareTo(amEtm) == -1)
                    {
                        global.errNumber = global.eKyuka;
                        global.errMsg = "午前半休で午前勤務が記入されています";
                        return false;
                    }
                }

                // 午後半休で午後勤務が記入されていたらエラー 2015/08/17
                if (dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    // 開始時刻が午前半休終了時刻超、または終了時刻が午後半休開始時間超のとき
                    if (sHM.CompareTo(amEtm) == 1 || eHM.CompareTo(pmStm) == 1)
                    {
                        global.errNumber = global.eKyuka;
                        global.errMsg = "午後半休で午後勤務が記入されています";
                        return false;
                    }
                }
            }

            // 勤務時間がゼロ以下になる場合はエラーとする  2010/12/04
            if ((ts.TotalMinutes <= 60) && (dR["休憩なし"].ToString() == global.FLGOFF))
            {
                global.errNumber = global.eKyukei;
                global.errMsg = "勤務時間がゼロまたはマイナスになります";
                return false;
            }
            
            // 勤務時間と休憩なしマーク 2015/08/17
            if ((ts.TotalMinutes > 360) && (dR["休憩なし"].ToString() == global.FLGON))
            {
                global.errNumber = global.eKyukei;
                global.errMsg = "勤務時間が6hを超えているので休憩なしは付けられません";
                return false;
            }

            return true;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     社員の開始時間、終了時間いずれにも記入がないときのチェック </summary>
        /// <param name="cdR">
        ///     勤務票ヘッダデータリーダー</param>
        /// <param name="dR">
        ///     勤務票明細データリーダー</param>
        /// <param name="iX">
        ///     日付</param>
        /// <param name="sTime">
        ///     所定勤務開始時間</param>
        /// <param name="eTime">
        ///     所定勤務終了時間</param>
        /// <returns>
        ///     </returns>
        ///---------------------------------------------------------------------------------
        private bool chkMeisai05_Sub02(OleDbDataReader cdR, OleDbDataReader dR, int iX, DateTime sTime, DateTime eTime)
        {
            DateTime dt;

            // 離席開始時間 数字以外
            if (dR["離席開始時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席開始時"].ToString()) ||
                    int.Parse(dR["離席開始時"].ToString()) < 0 || int.Parse(dR["離席開始時"].ToString()) > 23)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }
            }

            // 離席開始分 数字以外
            if (dR["離席開始分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席開始分"].ToString()) ||
                    int.Parse(dR["離席開始分"].ToString()) < 0 || int.Parse(dR["離席開始分"].ToString()) > 59)
                {
                    global.errNumber = global.eRSM;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }
            }

            // 離席終了時間 数字以外
            if (dR["離席終了時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席終了時"].ToString()) ||
                    int.Parse(dR["離席終了時"].ToString()) < 0 || int.Parse(dR["離席終了時"].ToString()) > 23)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時が不正です";
                    return false;
                }
            }

            // 離席終了分 数字以外
            if (dR["離席終了分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席終了分"].ToString()) ||
                    int.Parse(dR["離席終了分"].ToString()) < 0 || int.Parse(dR["離席終了分"].ToString()) > 59)
                {
                    global.errNumber = global.eREM;
                    global.errMsg = "離席終了時刻が不正です";
                    return false;
                }
            }

            // 離席時間に記入あり
            if (dR["離席開始時"].ToString() != string.Empty && dR["離席終了時"].ToString() != string.Empty)
            {
                // 離席開始時間と離席終了時間
                int stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                int etm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());

                if (stm >= etm)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時間が離席開始時間以前となっています";
                    return false;
                }

                // 開始時間と離席開始時間
                dt = DateTime.Parse(dR["離席開始時"].ToString() + ":" + dR["離席開始分"].ToString());
                if (dt <= sTime)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時間が開始時間以前となっています";
                    return false;
                }

                // 終了時間と離席終了時間
                dt = DateTime.Parse(dR["離席終了時"].ToString() + ":" + dR["離席終了分"].ToString());
                if (dt >= eTime)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時間が終了時間を超えています";
                    return false;
                }
            }

            //「特殊日マーク有り」：「所定時間マークなし」：「有休または特休でない」とき
            if (dR["特殊日"].ToString() == global.FLGON &&
                dR["所定勤務"].ToString() == global.FLGOFF &&
                dR["休暇"].ToString() != global.eYUKYU &&
                dR["休暇"].ToString() != global.eKYUMU)  // 特殊日で有休または特休はエラーとしない　'2010/10/15
            {
                // 土、日のとき
                if (dR["休日"].ToString() == global.hHOLIDAY.ToString() ||
                    dR["休日"].ToString() == global.hSATURDAY.ToString())
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "休日に特殊日マークは付けられません";
                    return false;
                }
                else
                {
                    global.errNumber = global.eSMark;
                    global.errMsg = "所定マークを付けるか時刻を入力してください";
                    return false;
                }
            }
            
            // 休暇に記入あり：特殊日マーク有り
            // 有休と特休はエラーとしない　'2010/10/19
            if (dR["休暇"].ToString() != string.Empty && 
                dR["休暇"].ToString() != global.eYUKYU && 
                dR["休暇"].ToString() != global.eKYUMU && 
                dR["特殊日"].ToString() == global.FLGON)
            {
                global.errNumber = global.eKyuka;
                global.errMsg = "休暇欄が不正です";
                return false;
            }

            // 勤務時間数を計算する　2010/12/04
            TimeSpan ts = Utility.GetTimeSpan(sTime, eTime);

            // 勤務時間がゼロ以下になる場合はエラーとする  2010/12/04
            if ((ts.TotalMinutes <= 60) && (dR["休憩なし"].ToString() == global.FLGOFF))
            {
                global.errNumber = global.eKyukei;
                global.errMsg = "勤務時間がゼロまたはマイナスになります";
                return false;
            }

            return true;
        }


        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     パートタイマーのエラーチェック（職種ID=07）</summary>
        /// <param name="cdR">
        ///     勤務票ヘッダデータリーダー</param>
        /// <param name="dR">
        ///     勤務票明細データリーダー</param>
        /// <param name="iX">
        ///     勤務票明細日付</param>
        /// <returns>
        ///     true:エラーなし、false:エラー有り</returns>
        ///--------------------------------------------------------------------------------------
        private bool ChkMeisai07(OleDbDataReader cdR, OleDbDataReader dR, int iX)
        {
            DateTime ssS;   // 所定開始または特殊勤務開始時間
            DateTime ssE;   // 所定終了または特殊勤務終了時間

            // 所定時間マーク有り、特殊日マークなし：週契約時間帯１を摘要
            if (dR["特殊日"].ToString() == global.FLGOFF && dR["所定勤務"].ToString() == global.FLGON)
            {
                ssS = DateTime.Parse(cdR["所定開始時間"].ToString());
                ssE = DateTime.Parse(cdR["所定終了時間"].ToString());
            }
            // 所定時間マーク有り、特殊日マーク有り：週契約時間帯２を摘要
            else if (dR["特殊日"].ToString() == global.FLGON && dR["所定勤務"].ToString() == global.FLGON)
            {
                ssS = DateTime.Parse(cdR["特殊開始時間"].ToString());
                ssE = DateTime.Parse(cdR["特殊終了時間"].ToString());
            }

            // その他：週契約時間帯１を摘要
            else
            {
                ssS = DateTime.Parse(cdR["所定開始時間"].ToString());
                ssE = DateTime.Parse(cdR["所定終了時間"].ToString());
            }

            // 土日以外で全て無記入のときは欠勤扱い
            if (dR["特殊日"].ToString() == global.FLGOFF && dR["所定勤務"].ToString() == global.FLGOFF &&
                dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() == string.Empty &&
                dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() == string.Empty &&
                dR["休暇"].ToString() == string.Empty && dR["休憩なし"].ToString() == global.FLGOFF &&
                dR["昼食"].ToString() == global.FLGOFF &&
                dR["離席開始時"].ToString() == string.Empty && dR["離席開始分"].ToString() == string.Empty &&
                dR["離席終了時"].ToString() == string.Empty && dR["離席終了分"].ToString() == string.Empty &&
                dR["休日"].ToString() == global.hWEEKDAY.ToString())
            {
                // 欠勤日数を加算
                global.KekkinCnt++;
                return true;
            }

            // 休日に半休が記入されている 2015/05/13
            if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
            {
                // 休日のとき
                if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休日に半休は記入できません";
                    return false;
                }
            }

            // 休日に有休または特休が記入されている 2015/05/13
            if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU)
            {
                // 休日のとき 2015/05/13
                if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休日に有休・特休は記入できません";
                    return false;
                }
            }

            // 開始時間、終了時間いずれかに記入があるとき
            if (dR["開始時"].ToString() != string.Empty || dR["開始分"].ToString() != string.Empty ||
                dR["終了時"].ToString() != string.Empty || dR["終了分"].ToString() != string.Empty)
            {
                if (!chkMeisai07_Sub01(cdR, dR, iX)) return false;
            }
            else
            {
                // 半休のとき開始時間、終了時間は必須とする 2012/06/19
                if (dR["休暇"].ToString() == global.eAMHANKYU ||
                    dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "半休のときは開始時間、終了時間は必須項目です";
                    return false;
                }
                else if (!chkMeisai07_Sub02(cdR, dR, iX, ssS, ssE)) return false;
            }

            //// 半休はエラーとする 2010/12/03
            //if (dR["休暇"].ToString() == global.eAMHANKYU ||
            //    dR["休暇"].ToString() == global.ePMHANKYU)
            //{
            //    global.errNumber = global.eKyuka;
            //    global.errMsg = "半休の入力はできません";
            //    return false;
            //}
            
            // 休暇 2012/06/19
            if (dR["休暇"].ToString() != string.Empty)
            {
                if (int.Parse(dR["休暇"].ToString()) < int.Parse(global.eYUKYU) ||
                    int.Parse(dR["休暇"].ToString()) > int.Parse(global.ePMHANKYU))
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休暇欄が不正です";
                    return false;
                }
            }

            // 昼食
            if (dR["昼食"].ToString() == global.FLGON)
            {
                if (dR["所定勤務"].ToString() == global.FLGOFF &&
                    dR["開始時"].ToString() == string.Empty)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "時刻を入力してください";
                    return false;
                }

                // 半休のときはエラーとする 2012/06/19
                if (dR["休暇"].ToString() == global.eAMHANKYU ||
                    dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "半休で昼食はできません";
                    return false;
                }

                // 休憩なし記入 2015/08/17
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = "昼食と休憩なしマークは両方付けられません";
                    return false;
                }
            }

            // 半休のとき 2012/06/19
            if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
            {
                // 休日のとき 2015/05/13
                if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休日に半休は記入できません";
                    return false;
                }

                if (dR["所定勤務"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eSMark;
                    global.errMsg = "半休なのでマークは付けられません";
                    return false;
                }
            }

            // 有休または特休のとき
            if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU)
            {
                // 休日のとき 2015/05/13
                if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休日に有休・特休は記入できません";
                    return false;
                }
        
                if (dR["所定勤務"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eSMark;
                    global.errMsg = "有休・特休なのでマークは付けられません";
                    return false;
                }

                if (dR["特殊日"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "有休・特休なので特殊日マークは付けられません";
                    return false;
                }

                if (dR["開始時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }

                if (dR["開始分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }

                if (dR["終了時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }

                if (dR["終了分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }

                if (dR["昼食"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eHiru1;
                    global.errMsg = "有休・特休なので昼は付けられません";
                    return false;
                }

                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = "有休・特休なので休憩マークは付けられません";
                    return false;
                }

                if (dR["離席開始時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }

                if (dR["離席開始分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eRSM;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }

                if (dR["離席終了時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }

                if (dR["離席終了分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eREM;
                    global.errMsg = "有休・特休なので時刻の入力はできません";
                    return false;
                }
            }

            // 休日だった場合
            if (dR["休日"].ToString() != global.hWEEKDAY.ToString())
            {
                if (dR["所定勤務"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eSMark;
                    global.errMsg = "休日出勤なので「○」は付けられません";
                    return false;
                }

                if (dR["特殊日"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "休日出勤なので「○」は付けられません";
                    return false;
                }

                if (dR["昼食"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eHiru1;
                    global.errMsg = "休日出勤なので昼食は無効です";
                    return false;
                }
            }

            // 振休あり出勤だった場合
            if (dR["休暇"].ToString() == global.eFURIDE)
            {
                // 特殊日マークあり
                if (dR["特殊日"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "振休あり出勤なので「○」は付けられません";
                    return false;
                }

                // 所定勤務マークなし
                if (dR["所定勤務"].ToString() == global.FLGOFF)
                {
                    if (dR["開始時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eSH;
                        global.errMsg = "開始時刻を入力してください";
                        return false;
                    }
                    if (dR["開始分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eSM;
                        global.errMsg = "開始時刻を入力してください";
                        return false;
                    }
                    if (dR["終了時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eEH;
                        global.errMsg = "終了時刻を入力してください";
                        return false;
                    }

                    if (dR["終了時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eEM;
                        global.errMsg = "終了時刻を入力してください";
                        return false;
                    }
                }

                // 離席時間記入あり
                if (dR["離席開始時"].ToString() != string.Empty || dR["離席開始分"].ToString() != string.Empty ||
                    dR["離席終了時"].ToString() != string.Empty || dR["離席終了分"].ToString() != string.Empty)
                {
                    if (dR["離席開始時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eRSH;
                        global.errMsg = "離席開始時刻を入力してください";
                        return false;
                    }

                    if (dR["離席開始分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eRSM;
                        global.errMsg = "離席開始時刻を入力してください";
                        return false;
                    }

                    if (dR["離席終了時"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eREH;
                        global.errMsg = "離席終了時刻を入力してください";
                        return false;
                    }

                    if (dR["離席終了分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eREM;
                        global.errMsg = "離席終了時刻を入力してください";
                        return false;
                    }
                }

                // 所定勤務マークなし
                if (dR["所定勤務"].ToString() == global.FLGOFF)
                {

                    // 開始時間と終了時間
                    int stm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());
                    int etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

                    if (stm >= etm)
                    {
                        global.errNumber = global.eEH;
                        global.errMsg = "終了時間が開始時間以前となっています";
                        return false;
                    }

                    // 離席時間に記入あり
                    if (dR["離席開始時"].ToString() != string.Empty && dR["離席終了時"].ToString() != string.Empty)
                    {
                        // 離席開始時間と離席終了時間
                        stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                        etm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());

                        if (stm >= etm)
                        {
                            global.errNumber = global.eREH;
                            global.errMsg = "離席終了時間が離席開始時間以前となっています";
                            return false;
                        }

                        // 開始時間と離席開始時間
                        stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                        etm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());

                        if (stm <= etm)
                        {
                            global.errNumber = global.eRSH;
                            global.errMsg = "離席開始時間が開始時間以前となっています";
                            return false;
                        }

                        // 終了時間と離席終了時間と離席開始時間
                        stm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());
                        etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

                        if (stm >= etm)
                        {
                            global.errNumber = global.eREH;
                            global.errMsg = "離席終了時間が終了時間を超えています";
                            return false;
                        }
                    }
                }
            }

            // 振替休日のとき
            if (dR["休暇"].ToString() == global.eFURIKYU)
            {
                if (dR["所定勤務"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eSMark;
                    global.errMsg = "振替休日なので「○」は付けられません";
                    return false;
                }

                if (dR["特殊日"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "振替休日なので「○」は付けられません";
                    return false;
                }

                if (dR["開始時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["開始分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["終了時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["終了分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = "振替休日なので休憩は付けられません";
                    return false;
                }

                if (dR["昼食"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eHiru1;
                    global.errMsg = "振替休日なので昼は付けられません";
                    return false;
                }
            }

            if (dR["特殊日"].ToString() == global.FLGON)
            {
                global.errNumber = global.eMark;
                global.errMsg = "特殊日マークは付けられません";
                return false;
            }
            
            // 訂正欄チェック状態のときエラーとする 2015/08/17
            if (dR["訂正"].ToString() == global.FLGON)
            {
                global.errNumber = global.eTeisei;
                global.errMsg = "訂正欄がチェックされています";
                return false;
            }

            return true;
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     パートタイマーの開始時間、終了時間いずれかに記入があるときのチェック </summary>
        /// <param name="cdR">
        ///     勤務票ヘッダデータリーダー</param>
        /// <param name="dR">
        ///     勤務票明細データリーダー</param>
        /// <param name="iX">
        ///     日付</param>
        /// <returns>
        ///     true:エラーなし, false:エラーあり</returns>
        ///--------------------------------------------------------------------------------------
        private bool chkMeisai07_Sub01(OleDbDataReader cdR, OleDbDataReader dR, int iX)
        {
            TimeSpan ts;

            // 開始時間 数字以外または範囲外
            if (!Utility.NumericCheck(dR["開始時"].ToString()) ||
                int.Parse(dR["開始時"].ToString()) < 0 || int.Parse(dR["開始時"].ToString()) > 23)
            {
                global.errNumber = global.eSH;
                global.errMsg = "開始時刻が不正です";
                return false;
            }

            // 開始分 数字以外または範囲外
            if (!Utility.NumericCheck(dR["開始分"].ToString()) ||
                int.Parse(dR["開始分"].ToString()) < 0 || int.Parse(dR["開始分"].ToString()) > 59)
            {
                global.errNumber = global.eSM;
                global.errMsg = "開始時刻が不正です";
                return false;
            }

            // 終了時間 数字以外または範囲外
            if (!Utility.NumericCheck(dR["終了時"].ToString()) ||
                int.Parse(dR["終了時"].ToString()) < 0 || int.Parse(dR["終了時"].ToString()) > 23)
            {
                global.errNumber = global.eEH;
                global.errMsg = "終了時刻が不正です";
                return false;
            }

            // 終了分 数字以外または範囲外
            if (!Utility.NumericCheck(dR["終了分"].ToString()) ||
                int.Parse(dR["終了分"].ToString()) < 0 || int.Parse(dR["終了分"].ToString()) > 59)
            {
                global.errNumber = global.eEM;
                global.errMsg = "終了時刻が不正です";
                return false;
            }

            // 離席開始時間 数字以外
            if (dR["離席開始時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席開始時"].ToString()) ||
                    int.Parse(dR["離席開始時"].ToString()) < 0 || int.Parse(dR["離席開始時"].ToString()) > 23)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }
            }

            // 離席開始分 数字以外
            if (dR["離席開始分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席開始分"].ToString()) ||
                    int.Parse(dR["離席開始分"].ToString()) < 0 || int.Parse(dR["離席開始分"].ToString()) > 59)
                {
                    global.errNumber = global.eRSM;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }
            }

            // 離席終了時間 数字以外
            if (dR["離席終了時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席終了時"].ToString()) ||
                    int.Parse(dR["離席終了時"].ToString()) < 0 || int.Parse(dR["離席終了時"].ToString()) > 23)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時が不正です";
                    return false;
                }
            }

            // 離席終了分 数字以外
            if (dR["離席終了分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席終了分"].ToString()) ||
                    int.Parse(dR["離席終了分"].ToString()) < 0 || int.Parse(dR["離席終了分"].ToString()) > 59)
                {
                    global.errNumber = global.eREM;
                    global.errMsg = "離席終了時刻が不正です";
                    return false;
                }
            }

            // 離席時間記入
            if (dR["離席開始時"].ToString() != string.Empty || dR["離席開始分"].ToString() != string.Empty ||
                dR["離席終了時"].ToString() != string.Empty || dR["離席終了分"].ToString() != string.Empty)
            {
                if (dR["離席開始時"].ToString() == string.Empty)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }

                if (dR["離席開始分"].ToString() == string.Empty)
                {
                    global.errNumber = global.eRSM;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }

                if (dR["離席終了時"].ToString() == string.Empty)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時刻が不正です";
                    return false;
                }

                if (dR["離席終了分"].ToString() == string.Empty)
                {
                    global.errNumber = global.eREM;
                    global.errMsg = "離席終了時刻が不正です";
                    return false;
                }
            }

            // 開始時間と終了時間
            int stm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());
            int etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

            if (stm < 600)
            {
                global.errNumber = global.eSH;
                global.errMsg = "開始時刻が6時以前になっています";
                return false;
            }

            if (stm >= etm)
            {
                global.errNumber = global.eEH;
                global.errMsg = "終了時間が開始時間以前となっています";
                return false;
            }

            // 離席時間に記入あり
            if (dR["離席開始時"].ToString() != string.Empty && dR["離席終了時"].ToString() != string.Empty)
            {
                // 離席開始時間と離席終了時間
                stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                etm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());

                if (stm >= etm)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時間が離席開始時間以前となっています";
                    return false;
                }

                // 開始時間と離席開始時間
                stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                etm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());

                if (stm <= etm)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時間が開始時間以前となっています";
                    return false;
                }

                // 終了時間と離席終了時間と離席開始時間
                stm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());
                etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

                if (stm >= etm)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時間が終了時間を超えています";
                    return false;
                }
            }

            // 所定勤務
            if (dR["所定勤務"].ToString() == global.FLGON)
            {
                global.errNumber = global.eSMark;
                global.errMsg = "時間とマークが両方入力されています";
                return false;
            }

            // 勤務時間数を計算する　2010/12/04
            DateTime sHM = DateTime.Parse(dR["開始時"].ToString() + ":" + dR["開始分"].ToString());
            DateTime eHM = DateTime.Parse(dR["終了時"].ToString() + ":" + dR["終了分"].ToString());
            ts = Utility.GetTimeSpan(sHM, eHM);

            // 半休のとき 2012/06/19
            if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
            {
                // 勤務時間が6H超のときエラーとする
                if (ts.TotalMinutes > 360)
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "半休で勤務時間が6hを超えています";
                    return false;
                }

                // 休憩なしマーク
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = "半休で休憩なしは付けられません";
                    return false;
                }

                // 午前半休終了時間 2015/08/17
                DateTime amEtm = DateTime.Parse(cdR["午前半休終了時間"].ToString());
                DateTime pmStm = DateTime.Parse(cdR["午後半休開始時間"].ToString());

                /// パートタイマーに特殊日なし
                //// 同　特殊日のとき 2015/08/17
                //if (dR["特殊日"].ToString() == global.FLGON)
                //{
                //    amEtm = DateTime.Parse(cdR["特殊午前半休終了時間"].ToString());
                //    pmStm = DateTime.Parse(cdR["特殊午後半休開始時間"].ToString());
                //}

                // 午前半休で午前中の勤務が記入されていたらエラー 2015/08/17
                if (dR["休暇"].ToString() == global.eAMHANKYU)
                {
                    // 開始時刻が午前半休終了時刻前のとき
                    if (sHM.CompareTo(amEtm) == -1)
                    {
                        global.errNumber = global.eKyuka;
                        global.errMsg = "午前半休で午前勤務が記入されています";
                        return false;
                    }
                }

                // 午後半休で午後勤務が記入されていたらエラー 2015/08/17
                if (dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    // 開始時刻が午前半休終了時刻超、または終了時刻が午後半休開始時間超のとき
                    if (sHM.CompareTo(amEtm) == 1 || eHM.CompareTo(pmStm) == 1)
                    {
                        global.errNumber = global.eKyuka;
                        global.errMsg = "午後半休で午後勤務が記入されています";
                        return false;
                    }
                }
            }

            // 勤務時間がゼロ以下になる場合はエラーとする  2010/12/04
            if ((ts.TotalMinutes <= 60) && (dR["休憩なし"].ToString() == global.FLGOFF))
            {
                global.errNumber = global.eKyukei;
                global.errMsg = "勤務時間がゼロまたはマイナスになります";
                return false;
            }
            
            // 勤務時間と休憩なしマーク
            if ((ts.TotalMinutes > 360) && (dR["休憩なし"].ToString() == global.FLGON))
            {
                global.errNumber = global.eKyukei;
                global.errMsg = "勤務時間が6hを超えているので休憩なしは付けられません";
                return false;
            }

            return true;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     パートタイマーの開始時間、終了時間いずれにも記入がないときのチェック </summary>
        /// <param name="cdR">
        ///     勤務票ヘッダデータリーダー</param>
        /// <param name="dR">
        ///     勤務票明細データリーダー</param>
        /// <param name="iX">
        ///     日付</param>
        /// <param name="sTime">
        ///     所定勤務開始時間</param>
        /// <param name="eTime">
        ///     所定勤務終了時間</param>
        /// <returns>
        ///     </returns>
        ///---------------------------------------------------------------------------------
        private bool chkMeisai07_Sub02(OleDbDataReader cdR, OleDbDataReader dR, int iX, DateTime sTime, DateTime eTime)
        {
            DateTime dt;
 
            // 離席開始時間 数字以外
            if (dR["離席開始時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席開始時"].ToString()) ||
                    int.Parse(dR["離席開始時"].ToString()) < 0 || int.Parse(dR["離席開始時"].ToString()) > 23)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }
            }

            // 離席開始分 数字以外
            if (dR["離席開始分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席開始分"].ToString()) ||
                    int.Parse(dR["離席開始分"].ToString()) < 0 || int.Parse(dR["離席開始分"].ToString()) > 59)
                {
                    global.errNumber = global.eRSM;
                    global.errMsg = "離席開始時刻が不正です";
                    return false;
                }
            }

            // 離席終了時間 数字以外
            if (dR["離席終了時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席終了時"].ToString()) ||
                    int.Parse(dR["離席終了時"].ToString()) < 0 || int.Parse(dR["離席終了時"].ToString()) > 23)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時が不正です";
                    return false;
                }
            }

            // 離席終了分 数字以外
            if (dR["離席終了分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["離席終了分"].ToString()) ||
                    int.Parse(dR["離席終了分"].ToString()) < 0 || int.Parse(dR["離席終了分"].ToString()) > 59)
                {
                    global.errNumber = global.eREM;
                    global.errMsg = "離席終了時刻が不正です";
                    return false;
                }
            }

            // 離席時間に記入あり
            if (dR["離席開始時"].ToString() != string.Empty && dR["離席終了時"].ToString() != string.Empty)
            {
                // 離席開始時間と離席終了時間
                int stm = int.Parse(dR["離席開始時"].ToString()) * 100 + int.Parse(dR["離席開始分"].ToString());
                int etm = int.Parse(dR["離席終了時"].ToString()) * 100 + int.Parse(dR["離席終了分"].ToString());

                if (stm >= etm)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時間が離席開始時間以前となっています";
                    return false;
                }

                // 所定開始時間と離席開始時間
                dt = DateTime.Parse(dR["離席開始時"].ToString() + ":" + dR["離席開始分"].ToString());
                if (dt <= sTime)
                {
                    global.errNumber = global.eRSH;
                    global.errMsg = "離席開始時間が開始時間以前となっています";
                    return false;
                }

                // 終了時間と離席終了時間と離席開始時間
                dt = DateTime.Parse(dR["離席終了時"].ToString() + ":" + dR["離席終了分"].ToString());
                if (dt >= eTime)
                {
                    global.errNumber = global.eREH;
                    global.errMsg = "離席終了時間が終了時間を超えています";
                    return false;
                }
            }

            // 特殊日マーク有り
            if (dR["特殊日"].ToString() == global.FLGON)
            {
                global.errNumber = global.eMark;
                global.errMsg = "特殊日マークは付けられません";
                return false;
            }

            // 休憩なし
            if (dR["休憩なし"].ToString() == global.FLGON)
            {
                global.errNumber = global.eKyukei;
                global.errMsg = "休憩なしマークは付けられません";
                return false;
            }

            // 特殊日あり、所定勤務マークなし、休暇有休、特休以外のとき
            if (dR["特殊日"].ToString() == global.FLGON &&
                dR["所定勤務"].ToString() == global.FLGOFF &&
                dR["休暇"].ToString() != global.eYUKYU &&
                dR["休暇"].ToString() != global.eKYUMU)
            {
                global.errNumber = global.eSMark;
                global.errMsg = "所定マークを付けるか時刻を入力してください";
                return false;
            }

            // 特殊日に休暇あり
            if (dR["休暇"].ToString() != string.Empty &&
                dR["特殊日"].ToString() == global.FLGON)
            {
                global.errNumber = global.eKyuka;
                global.errMsg = "休暇欄が不正です";
                return false;
            }

            // 勤務時間数を計算する　2010/12/04
            TimeSpan ts = Utility.GetTimeSpan(sTime, eTime);

            // 勤務時間がゼロ以下になる場合はエラーとする  2010/12/04
            if ((ts.TotalMinutes <= 60) && (dR["休憩なし"].ToString() == global.FLGOFF))
            {
                global.errNumber = global.eKyukei;
                global.errMsg = "勤務時間がゼロまたはマイナスになります";
                return false;
            }

            // 勤務時間と休憩なしマーク
            if ((ts.TotalMinutes > 360) && (dR["休憩なし"].ToString() == global.FLGON))
            {
                global.errNumber = global.eKyukei;
                global.errMsg = "勤務時間が6hを超えてるので休憩なしは付けられません";
                return false;
            }

            return true;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     受け渡しデータ見出し行出力 </summary>
        /// <param name="f">
        ///     出力ストリーム</param>
        ///------------------------------------------------------------
        private void CsvHeaderWrite(StreamWriter f)
        {
            // 奉行iシリーズ版 2010/10
            StringBuilder sb = new StringBuilder();
            sb.Append(@"""EBAS001""").Append(",");
            sb.Append(@"""SWDF010""").Append(",");
            sb.Append(@"""SWDF020""").Append(",");
            sb.Append(@"""SWDF030""").Append(",");
            sb.Append(@"""SWDF040""").Append(",");
            sb.Append(@"""SWDF050""").Append(",");
            sb.Append(@"""SWTF010""").Append(",");
            sb.Append(@"""SWTF030""").Append(",");
            sb.Append(@"""SWTF040""").Append(",");
            sb.Append(@"""SWTF050""").Append(",");
            sb.Append(@"""SWTF060""").Append(",");
            sb.Append(@"""SPPM170""").Append(",");
            sb.Append(@"""SWTF080""");

            f.WriteLine(sb.ToString());
        }
    }
}
