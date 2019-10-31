using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace HSS2.Model
{
    public class Entity
    {
        public struct Gyou
        {
            public int _clr;
        }

        // MDBデータキー配列
        public struct kData
        {
            public string _sID;
            public double _ShoTime;
            public double _TokTime;
            public double _TokTime2;
            public Gyou[] _Meisai;
        }

        // 受け渡しデータ
        public class saveData
        { 
            public string   sRecID;
            public string   sCode;
            public string   sJinji;
            public string   sNengetu;
            public string   sKeiyakuE;
            public string   sKinmu1;
            public string   sKinmu2;
            public string   sKinmu3;
            public string   sJikan1;
            public string   sJikan2;
            public string   sHiru1;
            public string   sHiru2;
            public string   sJigai1;
            public string   sJigai2;
            public string   sJigai3;
            public string   sJigai4;
            public string   sSeqNo;
            public string   sRidatu;
    
            // 以下固定項目
            public string   sShukinNisu;
            public string   sYukyuNisu;
            public string   sTokyuNisu;
            public string   sKekkinNisu;
            public string   sKyujituNisu;
            public string   sKinmuJikan;
            public string   sSitumuJikan;
            public string   sHayaZanJikan;
            public string   sSinyaJikan;
            public string   sKyujituJikan;
            public string   sKyujituSinya;

            // 計算項目
            public int sJituKinmu1;
            public double sJituKinmu2;
            public double sJituKinmu3;
            public int sJituKinmu4;
            public int sJituKinmu5;
            public double sKinmuJikan1;
            public double sKinmuJikan2;
            
            public int      _sHiru1;
            public int      _sHiru2;
            public double   sSitumu1;
            public double   sSitumu2;
            public double   sSitumuT;
            public double   sHayaZan;
            public double   sSinya;
            public double   kSinya;
            public double   sKyujitu;
            public double   Hayade;
            public double   Zangyo;
            public double   HayaZanT;
            public double   KSinyaT;
            public double   SinyaT;
            public double   TimeT;
            public DateTime SituE;
            public DateTime SituS;
            public double   Situ1;
            public int      Situ2;
            public DateTime sETime;
            public DateTime ShoETime;
            public DateTime ShoteiS;
            public DateTime ShoteiE;
            public double sRiseki;
            public string sHayadeSet;
            public string sZanSTime;    // 残業開始時間：会社別設定の所定終了時刻を取得
            public int sMarume;         // 丸め単位：会社別設定の丸め単位を取得 2011/01/11

            public double TimeJ;
            public double TimeJ2;
        }
        
        // スタッフ受け渡しデータ
        public class saveStaff
        { 
            public string   sCode;
            public string   sOrderCode;
            public string   sKikanS;
            public string   sKikanE;           
            public string   sKinmuNisu;         
            public string   sKinmuJikan;        
            public string   sYukyuNisu;         
            public string   sYukyuJikan;        
            public string   sTokyu;             
            public string   sKyushutu;          
            public string   sKekkin;            
            public string   sKeiyakunai;        
            public string   sZangyo1;           
            public string   sZangyo2;           
            public string   sZangyo3;           
            public string   sZangyo4;           
            public string   sShoteiNisu;
            public string sRidatu;             
    
            public string   S09; 
            public string   S10; 
            public string   S11; 
            public string   S19;   
            public string   S20;   
            public string   S21;   
            public string   S22;   
            public string   S27;   
            public string   S28;   
            public string   S29;   
            public string   S37;   
            public string   S38;   
            public string   S39;   
            public string   S40;   
            public string   S41;   
            public string   S42;   
            public string   S43;   
            public string   S44;   
            public string   S45;   
            public string   S46;   
            public string   S47;   
            public string   S48;   
            public string   S49;   
            public string   S50;   
            public string   S51;   
            public string   S52;   
            public string   S53;   
            public string   S54;   
            public string   S55;   
            public string   S56;   
            public string   S57;   
            public string   S58;   
            public string   S59;   
            public string   S60;   
            public string   S61;   
            public string   S62;   
            public string   S63;   
            public string   S64;   
            public string   S65;   
            public string   S66;   
            public string   S67;   
            public string   S68;   
            public string   S69;   
            public string   S70;   
            public string   S71;   
            public string   S72;   
            public string   S73;   
            public string   S74;   
            public string   S75;   
            public string   S76;   
            public string   S77;   
            public string   S78;   
            public string   S79;   
            public string   S80;   
            public string   S81;   
            public string   S82;
            public string   S83;

            // 計算項目
            public int sJituKinmu1;
            public double sJituKinmu2;
            public double sJituKinmu3;
            public double sJituKinmu4;
            public double sJituKinmu5;
            public double sKinmuJikan1;
            public double sKinmuJikan2;

            public int _sHiru1;
            public int _sHiru2;
            public double sSitumu1;
            public double sSitumu2;
            public double sSitumuT;
            public double sHayaZan;
            public double sSinya;
            public double kSinya;
            public double sKyujitu;
            public double Hayade;
            public double Zangyo;
            public double HayaZanT;
            public double SinyaT;
            public double TimeT;
            public DateTime SituE;
            public DateTime SituS;
            public double Situ1;
            public double Situ2;
            public DateTime sETime;
            public DateTime ShoETime;
            public DateTime ShoteiS;
            public DateTime ShoteiE;

            public double dKinmuNisu;
            public double sKyushutuNisu;
            public double Keiyakunai;
            public double YukyuNisu;
            public double YukyuJikan;
            public double TokukyuNisu;
            public double sShoteiKyujitu;
            public double sRiseki;
            public double dKinmuJikan;
            public double dShoteiNisu;
            public double dKekkinNisu;
            public string sHayadeSet;
            public string sZanSTime;    // 残業開始時間：会社別設定の所定終了時刻を取得
            public int sMarume;         // 丸め単位：会社別設定の丸め単位を取得 2011/01/11

            public double TimeJ;
            public double TimeJ2;
        }

        /// <summary>
        /// 給与奉行社員マスター読み込み構造体
        /// </summary>
        public struct ShainMst
        {
            public string sCode;
            public string sName;
            public string sKana;
            public string sOrderCode;
            public string sKeiyakuS;
            public string sKeiyakuE;
            public string sWorkTimeS;
            public string sWorkTimeE;
            public string sWorkTimeS2;
            public string sWorkTimeE2;
            public string sTenpoCode;
            public string sTenpoName;
            public string sBushoName;
            public string sJikansu;
            public string sKyuyoKbn;
            public string sSID;
            public string sSName;
        }
        
        /// <summary>
        /// 給与奉行職種マスター読み込み構造体
        /// </summary>
        public struct ShokushuMst
        {
            public string sCode;
            public string sName;
        }
    }
}
