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
using HSS2.Model;

namespace HSS2.Config
{
    public partial class frmComSet : Form
    {
        public frmComSet(Int64 sComID, string sComName, string sDbName)
        {
            InitializeComponent();

            _sComID = sComID.ToString().PadLeft(2, '0');
            _sComName = sComName;
            _sDbName = sDbName;
        }

        // 会社ID
        string _sComID;

        // 会社名
        string _sComName;

        // データベース名
        string _sDbName;

        // 職種ID配列
        Entity.ShokushuMst[] sMst = new Entity.ShokushuMst[1];

        // 選択中の職種ID,職種名
        string _SID = string.Empty;
        string _SName = string.Empty;

        // 会社別設定モード
        string _ComSet = string.Empty;

        // 出力設定モード
        string _OutSet = string.Empty;

        private void frmComSet_Load(object sender, EventArgs e)
        {
            this.txtID.Text = _sComID.PadLeft(2, '0');
            this.txtComName.Text = _sComName;

            // 職種ID取得
            MstLoad(_sDbName);

            // データ表示
            DataShow();
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     奉行マスター参照 </summary>
        /// <param name="dbName">
        ///     データベース名</param>
        ///-----------------------------------------------------------------
        private void MstLoad(string dbName)
        {
            // 勘定奉行データベース接続文字列を取得する 2016/10/12
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            SqlControl.DataControl sdcon = new SqlControl.DataControl(sc);

            //データリーダーを取得する
            SqlDataReader dr;
            string mySQL = string.Empty;

            // マスタ読み込み
            int iX = 0;
            string sReBuf = string.Empty;

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

            // コンボボックス値ロード
            cmbShokushuLoad();
        }

        private void cmbShokushuLoad()
        {
            cmbShokushu.Items.Clear();
            for (int i = 1; i < sMst.Length; i++)
            {
                cmbShokushu.Items.Add(sMst[i].sCode + "：" + sMst[i].sName);
            }

            cmbShokushu.SelectedIndex = 0;
        }

        private void cmbShokushu_SelectedIndexChanged(object sender, EventArgs e)
        {
            _SID = sMst[cmbShokushu.SelectedIndex + 1].sCode;
            _SName = sMst[cmbShokushu.SelectedIndex + 1].sName;
            getShokushu(_SID);
        }

        private void DataShow()
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT 会社別設定.会社ID as SID,会社別設定.会社名,会社別設定.丸め単位,");
            sb.Append("会社別設定.早出時間,会社別設定.残業時間, 出力設定.* FROM ");
            sb.Append("会社別設定 INNER JOIN 出力設定 ");
            sb.Append("ON 会社別設定.会社ID = 出力設定.会社ID ");
            sb.Append("where 会社別設定.会社ID = ? ");
            sb.Append("order by 会社別設定.会社ID,職種ID");
            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@id", _sComID);
            
            OleDbDataReader dr;
            dr = sCom.ExecuteReader();
            if (dr.HasRows)
            {
                dr.Read();
                txtMaru.Text = dr["丸め単位"].ToString();
                txtHayade.Text = dr["早出時間"].ToString().PadLeft(5, '0');
                txtZangyo.Text = dr["残業時間"].ToString().PadLeft(5, '0');

                // 職種ID
                for (int i = 0; i < sMst.Length; i++)
                {
                    if (sMst[i].sCode == dr["職種ID"].ToString())
                    {
                        cmbShokushu.SelectedIndex = i - 1;
                        break;
                    }
                }
            }
            dr.Close();
            sCom.Connection.Close();
        }

        private void getShokushu(string sID)
        {
            // チェック欄初期化
            CheckClear();

            // ローカルデータベースへ接続します
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM 出力設定 ");
            sb.Append("where 会社ID=? and 職種ID=?");
            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@id", _sComID);
            sCom.Parameters.AddWithValue("@sid", sID);

            OleDbDataReader dr;
            dr = sCom.ExecuteReader();
            while (dr.Read())
            {
                if (dr["出勤日数"].ToString() == "1") this.chkShukinNissu.Checked = true;
                if (dr["休日日数"].ToString() == "1") this.chkKyushutsuNIssu.Checked = true;
                if (dr["特休日数"].ToString() == "1") this.chkTokkyuNissu.Checked = true;
                if (dr["有休日数"].ToString() == "1") this.chkYukyuNissu.Checked = true;
                if (dr["欠勤日数"].ToString() == "1") this.chkKekkinNissu.Checked = true;
                if (dr["出勤時間"].ToString() == "1") this.chkShukkinTime.Checked = true;
                if (dr["執務時間"].ToString() == "1") this.chkShitsumuTime.Checked = true;
                if (dr["早出残業時間"].ToString() == "1") this.chkZanTime.Checked = true;
                if (dr["深夜時間"].ToString() == "1") this.chkShinyaTime.Checked = true;
                if (dr["休日時間"].ToString() == "1") this.chkKyujitsuTime.Checked = true;
                if (dr["昼食回数"].ToString() == "1") this.chkLunch.Checked = true;
                if (dr["休日深夜時間"].ToString() == "1") this.chkKyujitsuShinya.Checked = true;
            }
            dr.Close();
            sCom.Connection.Close();
        }

        private void CheckClear()
        {
            this.chkShukinNissu.Checked = false;
            this.chkKyushutsuNIssu.Checked = false;
            this.chkTokkyuNissu.Checked = false;
            this.chkYukyuNissu.Checked = false;
            this.chkKekkinNissu.Checked = false;
            this.chkShukkinTime.Checked = false;
            this.chkShitsumuTime.Checked = false;
            this.chkZanTime.Checked = false;
            this.chkShinyaTime.Checked = false;
            this.chkKyujitsuTime.Checked = false;
            this.chkLunch.Checked = false;
            this.chkKyujitsuShinya.Checked = false;
        }

        private void cmbShokushu_SelectedIndexChanged_1(object sender, EventArgs e)
        {
        }

        private void txtMaru_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }

        }

        private void txtMaru_Enter(object sender, EventArgs e)
        {
            if (sender == txtMaru) txtMaru.SelectAll();
            if (sender == txtZangyo) txtZangyo.SelectAll();
        }

        private void txtHayade_Enter(object sender, EventArgs e)
        {
            txtHayade.SelectAll();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!Utility.NumericCheck(txtMaru.Text))
            {
                MessageBox.Show("丸め単位は数字で登録してください",this.Text,MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                return;
            }

            if (int.Parse(txtMaru.Text) < 1 || int.Parse(txtMaru.Text) > 60)
            {
                MessageBox.Show("丸め単位は1～60の範囲で登録してください", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 更新確認
            if (MessageBox.Show("設定を更新します。よろしいですか", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) 
                return;

            // 会社別設定 登録or更新
            if (ComHasRow(_sComID)) _ComSet = global.FLGON;
            else _ComSet = global.FLGOFF;

            // 出力設定 登録or更新
            if (outHasRow(_sComID, _SID)) _OutSet = global.FLGON;
            else _OutSet = global.FLGOFF;

            // 設定更新
            comUpdate(_ComSet, _OutSet);

            // 終了
            this.Close();
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     会社別設定テーブルに登録済みか調べる </summary>
        /// <param name="ComID">
        ///     会社ID</param>
        /// <returns>
        ///     true:登録済み、false:未登録</returns>
        ///-----------------------------------------------------------
        private bool ComHasRow(string ComID)
        {
            // ローカルデータベースへ接続します
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Append("select * from 会社別設定 ");
            sb.Append("where 会社ID=?");
            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@id", ComID);

            sCom.CommandText = sb.ToString();
            OleDbDataReader dr = sCom.ExecuteReader();
            bool rtn;
            if (dr.HasRows) rtn = true;
            else rtn = false;
            sCom.Connection.Close();

            return rtn;
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     出力設定テーブルに登録済みか調べる </summary>
        /// <param name="ComID">
        ///     会社ID</param>
        /// <param name="ShokuID">
        ///     職種ID</param>
        /// <returns>
        ///     true:登録済み、false:未登録</returns>
        ///-----------------------------------------------------------
        private bool outHasRow(string ComID, string ShokuID)
        {
            // ローカルデータベースへ接続します
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Append("select * from 出力設定 ");
            sb.Append("where 会社ID=? and 職種ID=?");
            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@id", ComID);
            sCom.Parameters.AddWithValue("@sid", ShokuID);

            sCom.CommandText = sb.ToString();
            OleDbDataReader dr = sCom.ExecuteReader();
            bool rtn;
            if (dr.HasRows) rtn = true;
            else rtn = false;
            sCom.Connection.Close();

            return rtn;
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     設定データ更新 </summary>
        /// <param name="comMode">
        ///     会社別設定更新モード（0:新規登録、1:更新）</param>
        /// <param name="OutMode">
        ///     出力設定更新モード（0:新規登録、1:更新）</param>
        ///-----------------------------------------------------------
        private void comUpdate(string comMode, string OutMode)
        {
            // ローカルデータベースへ接続します
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();

            // 会社別設定テーブル
            switch (comMode)
            {
                case global.FLGOFF: // 新規登録
                    sb.Clear();
                    sb.Append("insert into 会社別設定 (");
                    sb.Append("会社ID,会社名,丸め単位,早出時間,残業時間) values (");
                    sb.Append("?,?,?,?,?) ");
                    sCom.CommandText = sb.ToString();
                    sCom.Parameters.Clear();
                    sCom.Parameters.AddWithValue("@ID", txtID.Text);
                    sCom.Parameters.AddWithValue("@Name", txtComName.Text);
                    sCom.Parameters.AddWithValue("@maru", txtMaru.Text);

                    if (txtHayade.Text.Trim() == ":") sCom.Parameters.AddWithValue("@hayade", string.Empty);
                    else if (txtHayade.Text.Replace("0", string.Empty) == ":") sCom.Parameters.AddWithValue("@hayade", string.Empty);
                    else sCom.Parameters.AddWithValue("@hayade", txtHayade.Text);

                    if (txtZangyo.Text.Trim() == ":") sCom.Parameters.AddWithValue("@zan", string.Empty);
                    else if (txtZangyo.Text.Replace("0", string.Empty) == ":") sCom.Parameters.AddWithValue("@zan", string.Empty);
                    else sCom.Parameters.AddWithValue("@zan", txtZangyo.Text);
                    
                    sCom.Parameters.AddWithValue("@id", _sComID);
                    sCom.Parameters.AddWithValue("@sid", _SID);
                    break;

                case global.FLGON:  // 更新
                    sb.Clear();
                    sb.Append("update 会社別設定 set ");
                    sb.Append("会社名=?,丸め単位=?, 早出時間=?, 残業時間=? ");
                    sb.Append("where 会社ID=?");
                    sCom.CommandText = sb.ToString();
                    sCom.Parameters.Clear();
                    sCom.Parameters.AddWithValue("@Com", txtComName.Text);
                    sCom.Parameters.AddWithValue("@maru", txtMaru.Text);
                    
                    if (txtHayade.Text.Trim() == ":") sCom.Parameters.AddWithValue("@hayade", string.Empty);
                    else if (txtHayade.Text.Replace("0", string.Empty) == ":") sCom.Parameters.AddWithValue("@hayade", string.Empty);
                    else sCom.Parameters.AddWithValue("@hayade", txtHayade.Text);

                    if (txtZangyo.Text.Trim() == ":") sCom.Parameters.AddWithValue("@zan", string.Empty);
                    else if (txtZangyo.Text.Replace("0", string.Empty) == ":") sCom.Parameters.AddWithValue("@zan", string.Empty);
                    else sCom.Parameters.AddWithValue("@zan", txtZangyo.Text);
                    
                    sCom.Parameters.AddWithValue("@id", _sComID);
                    sCom.Parameters.AddWithValue("@sid", _SID);
                    break;

                default:
                    break;
            }

            // トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                sCom.ExecuteNonQuery();

                switch (OutMode)
                {
                    case global.FLGOFF: // 新規登録
                        sb.Clear();
                        sb.Append("insert into 出力設定 (");
                        sb.Append("会社ID,会社名,職種ID,職種名,出勤日数,休日日数,特休日数,有休日数,欠勤日数,");
                        sb.Append("出勤時間,執務時間,早出残業時間,深夜時間,休日時間,昼食回数,休日深夜時間) values (");
                        sb.Append("?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ");
                        sCom.CommandText = sb.ToString();
                        sCom.Parameters.Clear();

                        sCom.Parameters.AddWithValue("@ComID", txtID.Text);
                        sCom.Parameters.AddWithValue("@ComName", txtComName.Text);
                        sCom.Parameters.AddWithValue("@ShokuID", _SID);
                        sCom.Parameters.AddWithValue("@ShokuName", _SName);

                        if (chkShukinNissu.Checked) sCom.Parameters.AddWithValue("@n1", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n1", global.FLGOFF);

                        if (chkKyushutsuNIssu.Checked) sCom.Parameters.AddWithValue("@n2", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n2", global.FLGOFF);

                        if (chkTokkyuNissu.Checked) sCom.Parameters.AddWithValue("@n3", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n3", global.FLGOFF);

                        if (chkYukyuNissu.Checked) sCom.Parameters.AddWithValue("@n4", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n4", global.FLGOFF);

                        if (chkKekkinNissu.Checked) sCom.Parameters.AddWithValue("@n5", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n5", global.FLGOFF);

                        if (chkShukkinTime.Checked) sCom.Parameters.AddWithValue("@t1", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t1", global.FLGOFF);

                        if (chkShitsumuTime.Checked) sCom.Parameters.AddWithValue("@t2", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t2", global.FLGOFF);

                        if (chkZanTime.Checked) sCom.Parameters.AddWithValue("@t3", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t3", global.FLGOFF);

                        if (chkShinyaTime.Checked) sCom.Parameters.AddWithValue("@t4", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t4", global.FLGOFF);

                        if (chkKyujitsuTime.Checked) sCom.Parameters.AddWithValue("@t5", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t5", global.FLGOFF);

                        if (chkLunch.Checked) sCom.Parameters.AddWithValue("@t6", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t6", global.FLGOFF);

                        if (chkKyujitsuShinya.Checked) sCom.Parameters.AddWithValue("@t7", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t7", global.FLGOFF);

                        break;

                    case global.FLGON:  // 更新
                        sb.Clear();
                        sb.Append("update 出力設定 set ");
                        sb.Append("職種名=?,出勤日数=?,休日日数=?,特休日数=?,有休日数=?,欠勤日数=?,");
                        sb.Append("出勤時間=?,執務時間=?,早出残業時間=?,深夜時間=?,休日時間=?,昼食回数=?,休日深夜時間=? ");
                        sb.Append("where 会社ID=? and 職種ID=?");
                        sCom.CommandText = sb.ToString();
                        sCom.Parameters.Clear();

                        sCom.Parameters.AddWithValue("@ShokuName", _SName);

                        if (chkShukinNissu.Checked) sCom.Parameters.AddWithValue("@n1", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n1", global.FLGOFF);

                        if (chkKyushutsuNIssu.Checked) sCom.Parameters.AddWithValue("@n2", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n2", global.FLGOFF);

                        if (chkTokkyuNissu.Checked) sCom.Parameters.AddWithValue("@n3", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n3", global.FLGOFF);

                        if (chkYukyuNissu.Checked) sCom.Parameters.AddWithValue("@n4", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n4", global.FLGOFF);

                        if (chkKekkinNissu.Checked) sCom.Parameters.AddWithValue("@n5", global.FLGON);
                        else sCom.Parameters.AddWithValue("@n5", global.FLGOFF);

                        if (chkShukkinTime.Checked) sCom.Parameters.AddWithValue("@t1", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t1", global.FLGOFF);

                        if (chkShitsumuTime.Checked) sCom.Parameters.AddWithValue("@t2", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t2", global.FLGOFF);

                        if (chkZanTime.Checked) sCom.Parameters.AddWithValue("@t3", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t3", global.FLGOFF);

                        if (chkShinyaTime.Checked) sCom.Parameters.AddWithValue("@t4", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t4", global.FLGOFF);

                        if (chkKyujitsuTime.Checked) sCom.Parameters.AddWithValue("@t5", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t5", global.FLGOFF);

                        if (chkLunch.Checked) sCom.Parameters.AddWithValue("@t6", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t6", global.FLGOFF);

                        if (chkKyujitsuShinya.Checked) sCom.Parameters.AddWithValue("@t7", global.FLGON);
                        else sCom.Parameters.AddWithValue("@t7", global.FLGOFF);
                        
                        sCom.Parameters.AddWithValue("@ComID", txtID.Text);
                        sCom.Parameters.AddWithValue("@ShokuID", _SID);

                        break;

                    default:
                        break;
                }

                sCom.ExecuteNonQuery();

                //トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,this.Text,MessageBoxButtons.OK,MessageBoxIcon.Exclamation);

                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void frmComSet_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
