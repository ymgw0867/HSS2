using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Drawing;

namespace HSS2.Model
{
    class global
    {
        // 表示画像名
        public static string pblImageFile;
        // MDBファイル名
        public const string MDBFILENAME = "HSS2.mdb";
        // MDBバックアップファイル名
        public const string MDBBACKUP = "HSS2_Back.mdb";
        // MDB一時ファイル名
        public const string MDBTEMPFILE = "HSS2_temp.mdb";

        // スタッフ、パート処理指定
        public const int STAFF_SELECT = 0;
        public const int PART_SELECT = 1;
        public const int END_SELECT = 9;
        public const int NON_SELECT = 0;

        // ＰＣ指定
        public const int PC1SELECT = 1;
        public const int PC2SELECT = 2;

        // 環境設定データ
        public static int sYear = 0;
        public static int sMonth = 0;
        public static string sTIF = string.Empty;       // 画像バックアップ作成先パス
        public static string sTIFCOM = string.Empty;    // 会社別画像バックアップ作成先パス
        public static string sDAT = string.Empty;       // 受け渡しデータ作成先パス
        public static string sDATCOM = string.Empty;    // 会社別受け渡しデータ作成先パス
        public static int sBKDELS = 0;                  // バックアップデータ保存月数

        // 会社領域情報
        public static long pblComNo;                   // 会社№
        public static string pblDbName;                 // 選択された会社のデータベース名
        public static string pblComName;                // 会社名

        // フラグオン・オフ
        public const string FLGON = "1";
        public const string FLGOFF = "0";

        // ○印
        public const string MARU = "○";

        //エラーチェック関連
        public static string errID;         //エラーデータID
        public static int errNumber;        //エラー項目番号
        public static int errRow;           //エラー行
        public static string errMsg;        //エラーメッセージ

        //エラー項目番号
        public const int eNothing = 0;      // エラーなし
        public const int eYear = 1;         // 対象年
        public const int eMonth = 2;        // 対象月
        public const int eShainNo = 3;      // 個人番号
        public const int eID = 4;           // 会社ID
        public const int eSID = 5;         // 職種ID
        public const int eRiseki = 6;       // 離席時間
        public const int eKyukei = 7;       // 休憩
        public const int eKyuka = 8;        // 休暇
        public const int eHiru1 = 9;        // 昼１
        public const int eHiru2 = 10;        // 昼２
        public const int eSH = 11;          // 開始時
        public const int eSM = 12;          // 開始分
        public const int eEH = 13;          // 終了時
        public const int eEM = 14;          // 終了分
        public const int eMark = 15;        // 特殊日マーク
        public const int eSMark = 16;       // 所定勤務マーク
        public const int eDays = 17;        // 出勤日数
        public const int eLunch = 18;       // 昼食回数
        public const int eRSH = 19;         // 離席開始時
        public const int eRSM = 20;         // 離席開始分
        public const int eREH = 21;         // 離席終了時
        public const int eREM = 22;         // 離席終了分
        public const int eYukyus = 23;      // 有休日数合計
        public const int eTokkyus = 24;     // 特休日数合計
        public const int eKekkins = 25;     // 欠勤日数合計
        public const int eTeisei = 26;      // 訂正欄  2015/08/18

        //表示関係
        public static float miMdlZoomRate = 0;      //現在の表示倍率

        //表示倍率（%）
        public static float ZOOM_RATE = 0.20f;      //標準倍率
        public static float ZOOM_MAX = 2.00f;       //最大倍率
        public static float ZOOM_MIN = 0.05f;       //最小倍率
        public static float ZOOM_STEP = 0.02f;      //ステップ倍率

        // 表示色関連
        public static Color lBackColorE;
        public static Color lBackColorN;
        public static Color lBackColorK;
        public static Color lBackColorR;

        // DataGridChangeValueイベント発生制御
        public static bool dg1ChabgeValueStatus;

        // 所定時間情報
        public static string ShoS = string.Empty;
        public static string ShoE = string.Empty;
        public static string AmS = string.Empty;
        public static string AmE = string.Empty;
        public static string PmS = string.Empty;
        public static string PmE = string.Empty;

        // 特殊勤務時間情報
        public static string ShoS2 = string.Empty;
        public static string ShoE2 = string.Empty;
        public static string AmS2 = string.Empty;
        public static string AmE2 = string.Empty;
        public static string PmS2 = string.Empty;
        public static string PmE2 = string.Empty;

        // 休暇区分
        public static string eYUKYU = "1";
        public static string eKYUMU = "2";
        public static string eFURIDE = "3";
        public static string eFURIKYU = "4";
        public static string eAMHANKYU = "5";
        public static string ePMHANKYU = "6";

        // 休日区分
        public static int hWEEKDAY = 0;
        public static int hSATURDAY = 1;
        public static int hHOLIDAY = 2;

        // 勤務データ登録区分
        public const int sADDMODE = 1;
        public const int sEDITMODE = 0;

        // 人事奉行iシリーズ
        // 週契約時間帯 引数
        public const string KEIYAKUTIME_HIKISU = "6000000000000000";

        // 欠勤数カウント
        public static int KekkinCnt;

        // 八十二リースステータス : 2017/12/08
        public static int _82lease = 0;

        /// <summary>
        /// 環境設定情報を取得します
        /// </summary>
        public static void GetCommonYearMonth()
        {
            Model.SysControl.SetDBConnect sDB = new Model.SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = sDB.cnOpen();

            string mySql = "select * from 環境設定";
            sCom.CommandText = mySql;
            OleDbDataReader dr = sCom.ExecuteReader();

            try
            {
                while (dr.Read())
                {
                    sYear = int.Parse(dr["SYEAR"].ToString());
                    sMonth = int.Parse(dr["SMONTH"].ToString());
                    sTIF = dr["TIF"].ToString();
                    sDAT = dr["DAT"].ToString();
                    sBKDELS = int.Parse(dr["BKDELS"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定年月取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (dr.IsClosed == false) dr.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }
    }
}
