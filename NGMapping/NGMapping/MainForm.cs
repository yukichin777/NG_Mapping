using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Media.Media3D;

namespace NGMapping
{
    public partial class MainForm : Form
    {
        private readonly List<NgCounter> ngCounters = []; // クリックデータのリスト
        private ContextMenuStrip contextMenu; // コンテキストメニュー
        private Rectangle cachedImageRect;
        private Point lastRightClickPoint = Point.Empty;
        private int pBoxInd = 0; // PictureBoxのインデックス

        private readonly PictureBox[] picBoxes;
        private bool isRotated = false; // 画像が回転しているかどうかのフラグ    


        private float[] xy_A = [];
        private float[] xy_B = [];
        private readonly RotateImageGenerator[] ImgGen = [new(), new()];

        readonly ExLabel[] LCountA;
        readonly ExLabel[] LCountB;
        readonly ExLabel[] LCount;


        readonly RadioButton[] Ra_NgItems;
        readonly Label[] La_NgItems;
        readonly List<string> NgTexts_jp = ["はんだボール", "はんだ屑", "炭化物", "ヒュミシール付着", "ヒュミシール未塗布", "浮き", "破損", "リードカット異常"];
        readonly List<string> NgTexts_en = ["Solder ball", "Solder residue", "Carbide", "Humiseal adhesion", "No Humiseal adhesion", "Lift or Floating", "Damage", "Lead cutting abnormality"];
        readonly List<string> NgTexts_pt = ["Bola de solda", "Resíduo de solda", "Carboneto", "Adesão de Humiseal", "Sem adesão de Humiseal", "Levantamento ou Flutuação", "Danos", "Anormalidade no corte de terminal"];
        readonly List<string> NgTexts_rm = ["handa booru", "handa kotsubu", "tankabutsu", "hyumishiiru fuchaku", "hyumishiiru mitofu", "uki", "hason", "riido katto ijou"];
        readonly List<string> etcTexts_jp = ["異物", "汚れ", "フラックス"];
        readonly List<string> etcTexts_en = [ "Foreign material", "Dirt/Contamination", "Flux"];
        readonly List<string> etcTexts_pt = [ "Material estranho", "Sujeira/Contaminação", "Fluxo"];

        readonly List<UndoInfo> undoList = []; // クリックデータの履歴


        private GlobalKeyboardHook hook = new(); // グローバルキーボードフック

        

        bool isQRDisp = false; // QRコードの表示フラグ
        bool isQRead = false; // QRコード読み取りフラグ

        SQLiteCon db;

        readonly f_Login loginForm;

        //DB用データ(T_Daicho用)
        private static DateTime InitDt = new(1900, 1, 1); // 初期化用の検査日

        private int mode = 0; // mode(0:紙から入力、1:検査場で入力) 
        private string QRText = ""; // シリアル番号
        private string SN = ""; // シリアル番号
        private string operatorName = "";
        private DateTime TestDate = InitDt;//  DateTime.Now; // 検査日時(Time情報あり）
        private DateTime SaveDate = InitDt; // 保存日時(Time情報あり）
        private DateTime BoardDate = InitDt; // 基板製造年月日(Time情報なし）
        private bool isDay = false;
        private string LineName = "";
        private string hinban = "";
        private string AB = "";
        #region constructor
        public MainForm(string opeName, f_Login frm)
        {
            InitializeComponent();
            loginForm = frm;

            LCountA = [L_CountA_0, L_CountA_1, L_CountA_2, L_CountA_3, L_CountA_4, L_CountA_5, L_CountA_6, L_CountA_7, L_CountA];
            LCountB = [L_CountB_0, L_CountB_1, L_CountB_2, L_CountB_3, L_CountB_4, L_CountB_5, L_CountB_6, L_CountB_7, L_CountB];
            LCount = [L_Count_0, L_Count_1, L_Count_2, L_Count_3, L_Count_4, L_Count_5, L_Count_6, L_Count_7, L_Count];

            picBoxes = [picBox_A, picBox_B];
            ImgGen[0].pBox = picBox_A;
            ImgGen[1].pBox = picBox_B;

            Ra_NgItems = [radioButton1, radioButton2, radioButton3, radioButton4, radioButton5, radioButton6, radioButton7, radioButton8, radioButton9];
            La_NgItems = [L_Color1, L_Color2, L_Color3, L_Color4, L_Color5, L_Color6, L_Color7, L_Color8];

            t_operator.Text = opeName;
            operatorName = opeName;
            this.loginForm = frm;
        }
        #endregion
        #region event------Formロード
        private void MainForm_Load(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CSet.DbPath) || !File.Exists(CSet.DbPath))
            {
                string myDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string dbfolder = Path.Combine(myDocumentsPath, "NgMapping");
                Directory.CreateDirectory(dbfolder);
                string dbPath = Path.Combine(dbfolder, "NgMapping.sqlite");
                db = new(dbPath, CSet.DbTableCFG(), admin: true);
                CSet.DbPath = dbPath;
            }
            else
            {
                SQLiteCon db = new(CSet.DbPath, CSet.DbTableCFG(), true);
            }

            // PictureBoxのイベント設定
            picBox_A.MouseClick += PictureBox_MouseClick;
            picBox_B.MouseClick += PictureBox_MouseClick;
            dtPicker_TestDate.Value = DateTime.Now.Date; // デフォルトで今日の日付を設定

            CSet.SetFormLocState(this);

            Init();

            tsCombo_Language.ComboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;

        }

        #endregion
        #region method-----初期化
        private void Init()
        {

            db = new(CSet.DbPath, CSet.DbTableCFG(), false);
            ColorInit(); // 色の初期化とコンテキストメニューの作成
            tsCombo_Language.ComboBox.Text = CSet.Language;
            SetLanguage(CSet.Language); // 言語設定を適用

            isQRead = operatorName != "##"; 
            isQRDisp = isQRead && CSet.FLG_DispDummyQR;

            menu_Save.Enabled = !isQRead; // QRコード読み取りモードでは保存ボタンを無効化
            dtPicker_TestDate.Enabled = !isQRead; // QRコード読み取りモードでは検査日ピッカーを無効化
            cb_Hinban.Enabled = !isQRead; // QRコード読み取りモードでは品番コンボボックスを無効化

            if(isQRead)
            {
                //hook=new GlobalKeyboardHook(); // QRコード読み取りモードではグローバルキーボードフックを使用
                hook.SubmitKeyMode= SubmitKey.CR; // QRコード読み取りモードではEnterキーで送信
                hook.InputSubmitted += Hook_InputSubmitted; 
                hook.Start();
            }

        }
        #endregion

        private bool InDayRange(DateTime dt)
        {
            var dtRange = CSet.DayTimeRange; // 日付範囲設定
            return dt.TimeOfDay >= dtRange[0].TimeOfDay && dt.TimeOfDay < dtRange[1].TimeOfDay;
        }
        #region event-----QRコード読み取りイベント


        private void Hook_InputSubmitted(object sender, KeyInputEventArgs e)
        {
            //例：250528H659000B10213IF
            //分離
            //0-6（250528）：製造年月日
            //6-6（H）：ライン名
            //7-12（659000）：品番　　前１桁+"365"+2桁目から3桁　// 6+365+590で、6365590
            //13-13（B）：基板面（A or B）
            //14-20（10213IF）：シリアル番号（7桁）
            mode = 1; // QRコード読み取りモードに設定
            QRText = "";
            SN = "";
            hinban = ""; // 品番を初期化
            TestDate = InitDt; // 検査日を初期化
            SaveDate = InitDt; // 保存日時を初期化
            BoardDate = InitDt; // 基板製造年月日を初期化
            isDay = InDayRange(DateTime.Now);//現在時間で、昼勤か夜勤かを判定

            string QRCode = e.InputText.Trim();
            if (QRCode.Length != 21) return;

            string dtTxt = "20" + QRCode.Substring(0, 2) + "/" + QRCode.Substring(2, 2) + "/" + QRCode.Substring(4, 2);

            if (!DateTime.TryParse(dtTxt, out DateTime bDate))
            {
                return;
            }


            switch (QRCode.Substring(7, 6))
            {
                case "659000":
                    hinban = "6365590";
                    break;
                case "663000":
                    hinban = "6365630";
                    break;
                default:
                    return; // 無効なS/N形式の場合は何もしない
            }

            BoardDate = bDate; // 基板製造年月日を設定
            LineName = QRCode.Substring(6, 1); // ライン名を取得
            AB = QRCode.Substring(13, 1); // ライン名を取得
            SN = QRCode.Substring(14, 7); // シリアル番号を取得

        }
        #endregion
        #region method-----色初期化/コンテキストメニュー作成
        private void ColorInit()
        {
            Color[] colors = CSet.NgColors;
            ImgGen[0].PointColors = colors;
            ImgGen[1].PointColors = colors;
            for (int i = 0; i < colors.Length; i++)
            {
                La_NgItems[i].BackColor = colors[i];
                La_NgItems[i].Tag = i;
            }


            // コンテキストメニューの設定
            contextMenu = new ContextMenuStrip();
            for (int i = 0; i < 8; i++)
            {
                int ngType = i; // ローカル変数でNG種類を保存
                ToolStripMenuItem menuItem = new(NgTexts_jp[i]) { Tag = i };
                menuItem.Click += MenuItem_Click;
                contextMenu.Items.Add(menuItem);
            }


            if (CSet.NgItem != null)
            {
                // CSet.NgItemをDataGridViewに復元
                foreach (var entry in CSet.NgItem)
                {
                    // NGワードとNG種類（ComboBoxのインデックス）を取得
                    string ngWord = entry.Key;
                    int selInd = entry.Value;

                    ToolStripMenuItem menuItem = new(ngWord) { Tag = selInd };
                    menuItem.Click += MenuItem_Click;
                    contextMenu.Items.Add(menuItem);
                }
            }
        }
        #endregion
        #region event------tsCombo_Language_IndexChanged
        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string lang = ((ComboBox)sender).SelectedItem.ToString();
            SetLanguage(lang); // 言語設定を更新
            CSet.Language = lang; // 言語設定を保存
        }
        #endregion
        #region method-----言語設定
        private void SetLanguage(string lang)
        {           
            List<string> ngtxt;
            int wd0 = 130;
            int wd1 = 130;

            switch (lang)
            {
                case "RomanAlphabet":
                    ngtxt = NgTexts_rm;
                    ngtxt.Add("Keshigomu"); // ポルトガル語のNGテキストを追加
                    wd0 = 160;
                    wd1 = 125;
                    label3.Text = "Kensabi";
                    label4.Text = "Kiban";
                    ra_Day.Text = "Chukin";
                    ra_Night.Text = "Yakin";
                    L_Operator.Text = "Kensasya";
                    label10.Text = "Kei";
                    label44.Text = "Kei";
                    L_A.Text = "A_Men";
                    L_B.Text = "B_Men";
                    button3.Text = "Hanten";
                    ra_Normal.Text = "Tsujo"; // 通常
                    ra_TofuArea.Text = "Tofu"; // 塗布
                    ra_KinshiArea.Text = "Kinshi"; // 禁止
                    break;
                case "English":
                    ngtxt = NgTexts_en;
                    ngtxt.Add("Eraser"); // 英語のNGテキストを追加
                    wd0 = 175;
                    wd1 = 190;
                    label3.Text = "Test Date";    // 検査日
                    label4.Text = "Board";      // 対象基盤
                    ra_Day.Text = "Day Shift";         // 昼勤
                    ra_Night.Text = "Night Shift";     // 夜勤
                    L_Operator.Text = "Operator";     // 検査者
                    label10.Text = "Sum";            // 計
                    label44.Text = "Sum";            // 計
                    L_A.Text = "Side A";               // A面
                    L_B.Text = "Side B";               // B面
                    button3.Text = "Toggle";           // 反転

                    ra_Normal.Text = "regular"; // 通常
                    ra_TofuArea.Text = "apply"; // 塗布
                    ra_KinshiArea.Text = "no apply"; // 禁止

                    break;
                case "Português":
                    ngtxt = NgTexts_pt;
                    ngtxt.Add("Borracha"); // ポルトガル語のNGテキストを追加
                    wd0 = 195;
                    wd1 = 210;
                    label3.Text = "Data de Teste";       // 検査日 (Test Date)
                    label4.Text = "Placa";          // 対象基盤
                    ra_Day.Text = "Turno Diurno";        // 昼勤
                    ra_Night.Text = "Turno Noturno";     // 夜勤
                    L_Operator.Text = "Operador";        // 検査者
                    label10.Text = "Soma";              // 計
                    label44.Text = "Soma";              // 計
                    L_A.Text = "Lado A";                 // A面
                    L_B.Text = "Lado B";                 // B面
                    button3.Text = "Inverter";

                    ra_Normal.Text = "regular"; // 通常
                    ra_TofuArea.Text = "aplicar"; // 塗布
                    ra_KinshiArea.Text = "proibido"; // 禁止

                    break;
                case "Japanease":
                default:
                    label3.Text = "検査日";
                    label4.Text = "対象基盤";
                    ra_Day.Text = "昼勤";
                    ra_Night.Text = "夜勤";
                    L_Operator.Text = "検査者";
                    label10.Text = "計";
                    label44.Text = "計";
                    L_A.Text = "A面";
                    L_B.Text = "B面";
                    button3.Text = "反転";

                    ra_Normal.Text = "通常"; // 通常
                    ra_TofuArea.Text = "塗布エリア"; // 塗布
                    ra_KinshiArea.Text = "禁止エリア"; // 禁止


                    ngtxt = NgTexts_jp;
                    ngtxt.Add("消しゴム"); // 日本語のNGテキストを追加
                    ngtxt = NgTexts_jp;
                    wd0 = 130;
                    wd1 = 115;
                    break;
            }

            for (int i = 0; i < Ra_NgItems.Length; i++)
            {
                Ra_NgItems[i].Text = ngtxt[i];

            }
            //tableLayoutPanel2.ColumnStyles[5].SizeType = SizeType.Percent; // サイズを％で設定可能にする
            tableLayoutPanel2.ColumnStyles[5].Width = wd0;
            tableLayoutPanel2.ColumnStyles[11].Width = wd1;
        }
        #endregion
        #region event------コンテキストメニューのクリックイベント
        private void MenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem menuItem = sender as ToolStripMenuItem;

            int ngType = (int)menuItem.Tag; // NG種類を取得
            string ngText = menuItem.Text; // NGテキストを取得

            if (lastRightClickPoint.IsEmpty || !cachedImageRect.Contains(lastRightClickPoint)) return;
            ImgGen[pBoxInd].addPoint(lastRightClickPoint, ngType, ngText);
            undoList.Add(new UndoInfo(pBoxInd, -1, null));
            CountUpdate(); // カウントを更新
        }
        #endregion
        #region method-----NGカウント更新
        private void CountUpdate()
        {
            int[] NgCountA = ImgGen[0].NgCount;
            int[] NgCountB = ImgGen[1].NgCount;

            for (int i = 0; i < 8; i++)
            {
                LCountA[i].Value = NgCountA[i];
                LCountB[i].Value = NgCountB[i];
                LCount[i].Value = NgCountA[i] + NgCountB[i];
            }

            LCountA[8].Value = NgCountA.Sum();
            LCountB[8].Value = NgCountB.Sum();
            LCount[8].Value = NgCountA.Sum() + NgCountB.Sum();

        }
        #endregion
        #region event-----PictureBoxのクリックイベント
        private void PictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            if (sender is not PictureBox pBox) return; // nullチェック

            cachedImageRect = GetImageDisplayRectangle(pBox);

            pBoxInd = int.Parse(pBox.Tag.ToString());
            bool CTRL = (Control.ModifierKeys & Keys.Control) == Keys.Control;

            if (e.Button == MouseButtons.Left)
            {
                int ngType = GetSelectedNGType();
                if (CTRL || ngType==8) {
                    int ListInd = ImgGen[pBoxInd].DeletePoint(e.Location, out NgCounter ngcnt);
                    if (ListInd != -1)
                    {
                        undoList.Add(new UndoInfo(pBoxInd,ListInd,ngcnt));
                    }
                }
                else
                {
                    if (ngType == -1)
                    {
                        MessageBox.Show("NG種類を選択してください。");
                        return;
                    }
                    ImgGen[pBoxInd].addPoint(e.Location, ngType, NgTexts_jp[ngType]);
                    undoList.Add(new UndoInfo(pBoxInd, -1, null));
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                if (cachedImageRect.Contains(e.Location)) // 画像内で右クリックした場合
                {
                    lastRightClickPoint = e.Location; // 右クリック位置を記録
                    contextMenu.Show(pBox, e.Location);
                }
            }
            CountUpdate();
        }
        #endregion
        #region method-----PictureBox内のImage部分のRectangleを返す関数
        private Rectangle GetImageDisplayRectangle(PictureBox pb)
        {
            if (pb.Image == null) return Rectangle.Empty;

            float imageRatio = (float)pb.Image.Width / pb.Image.Height;
            float controlRatio = (float)pb.Width / pb.Height;

            int width, height;
            if (imageRatio > controlRatio)
            {
                width = pb.Width;
                height = (int)(pb.Width / imageRatio);
            }
            else
            {
                height = pb.Height;
                width = (int)(pb.Height * imageRatio);
            }

            int x = (pb.Width - width) / 2;
            int y = (pb.Height - height) / 2;

            return new Rectangle(x, y, width, height);
        }
        #endregion
        #region method-----選択されているNG種類を取得する関数
        private int GetSelectedNGType()
        {
            for (int i = 0;i<Ra_NgItems.Length;i++)
            {
                if (Ra_NgItems[i].Checked) return i;
            }
            return -1;
        }
        #endregion
        #region event-----品番コンボボックス選択
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            setSN(true); // S/Nを設定
        }
        #endregion
        private void DtPicker_TestDate_ValueChanged(object sender, EventArgs e)
        {
            setSN(true);
        }

        private void ra_Day_Click(object sender, EventArgs e)
        {
            setSN(true);
        }
        private void setSN(bool FLG_ReadSN_FromDB) 
        {

            SN = "";

            if (cb_Hinban.SelectedIndex < 0) return;
            if(!ra_Day.Checked && !ra_Night.Checked) return;
            DateTime dt = dtPicker_TestDate.Value.Date;
            DateTime dtNow = DateTime.Now;
            if (dt.Year < 2025 || dt > dtNow.Date) return; // 2025年以前や未来は処理しない

            string hinban = cb_Hinban.Text;
            int isDay = ra_Day.Checked ? 1 : 0; // 昼勤か夜勤かを判定
            string DummySN = $"{cb_Hinban.Text}_{dt:yyyyMMdd}_{dtNow:HHmmss}";


            if (!LoadImageToPictureBox(out string mes))
            {
                MessageBox.Show(mes);
            }



            if (isQRead)//QRコード読み取りモードの場合
            {                
                if (isQRDisp) // QRコード表示
                {
                    DispQRCode(DummySN); // QRコードの表示
                }
            }
            else //紙から入力の場合
            {
                if (FLG_ReadSN_FromDB)
                {
                    //まず、検査日がdtで、品番がcb_Hinban.TextのS/NﾘｽﾄをDBから取得する
                    //string sql = $"SELECT SN FROM T_Daicho WHERE TestDate = '{dt:yyyy-MM-dd}' AND Hinban = '{hinban.Replace("'", "''")}' AND isDay = {isDay} ORDER BY SaveDateTime DESC;";

                    string sql = 
                        $"SELECT DISTINCT SN " +
                        $"FROM T_Daicho " +
                        $"WHERE TestDate = '{dt:yyyy-MM-dd}' " +
                        $"AND Hinban = '{hinban.Replace("'", "''")}' " +
                        $"AND isDay = {isDay} " +
                        $"ORDER BY SaveDateTime;";


                    db.GetData(sql, out DataTable dtSNList);

                    listBox1.Items.Clear(); // ListBoxの内容をクリア
                    listBox1.Items.AddRange([.. dtSNList.Rows.Cast<DataRow>().Select(row => row.Field<string>("SN"))]);
                }
                SN = DummySN;

            }

            L_SN.Text = SN; // ラベルに現在のシリアル番号を表示


        }



        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex < 0) return; // 選択されていない場合は何もしない

            string SN= listBox1.Text;
            ReadSN(SN); // 選択されたS/Nを読み込む

        }

        private void ReadSN(string sn)
        {
            if (string.IsNullOrWhiteSpace(sn)) return; // 空文字列の場合は何もしない
            List<string> sql = [];
            string[] AB = ["A", "B"];

            ImgGen[0].ClearPoints(); // 画像のポイントをクリア
            ImgGen[1].ClearPoints(); // 画像のポイントをクリア


            for (int i = 0; i < 2; i++)
            {
                sql.Add(
                $"SELECT " +
                $"  T1.TestDate AS TestDate, " +
                $"  T1.Hinban AS Hinban, " +
                $"  T2.Board AS Board, " +
                $"  T2.NgType AS NgType, " +
                $"  T2.NgText AS NgText, " +
                $"  T2.X AS X, " +
                $"  T2.Y AS Y " +
                $"FROM " +
                $"            T_Daicho T1 " +
                $"  LEFT JOIN T_Data   T2 ON T2.SN = T1.SN " +
                $"WHERE " +
                $"  T1.SN = '{sn}' " +
                $"  AND T2.Board='{AB[i]}';");
            }

            db.GetData(sql, out DataTable[] dts);


            for (int i = 0; i < 2; i++)
            {
                ImgGen[i].ClearPoints(); // 画像のポイントをクリア

                if (dts[i].Rows.Count == 0) continue; // データがない場合はスキップ
                foreach (DataRow ro in dts[i].Rows)
                {

                    float x = (float)ro.Field<double>("X");
                    float y = (float)ro.Field<double>("Y");
                    int typ= (int)ro.Field<long>("NgType");
                    string txt = ro.Field<string>("NgText");

                    NgCounter ngCounter = new NgCounter(x, y, typ, txt);
                    ImgGen[i].addPoint(ngCounter); // 画像にポイントを追加
                }
            }

            // QRコードのテキストを設定
            SN = sn;
            L_SN.Text = sn;
            // QRコードを表示
            DispQRCode(sn);
        }

        private void DispQRCode(string txt) 
        {
            if (!isQRead || !isQRDisp) return;

            CM.MakeQRImage(txt, out Image QrImage);
            pictureBox2.Image = QrImage;
            L_SN.Text = txt;

        }
        string[] QrCodeText = ["", ""];
        private void DispQRCode(int ind)
        {
            //no　1：次のワーク、0：今のワーク

            if (cb_Hinban.SelectedIndex < 0) return; // 品番が選択されていない場合は処理を中止
            string hinban = cb_Hinban.Text;
            DateTime dt0 = dtPicker_TestDate.Value;
            DateTime dt1 = DateTime.Now;

            int year = dt0.Year;
            int month = dt0.Month;
            int day = dt0.Day;
            int hour = dt1.Hour;
            int minute = dt1.Minute;
            int second = dt1.Second;
            DateTime dateTime = new(year, month, day, hour, minute, second);

            switch (ind)
            {
                case 1:
                    if (QrCodeText[1] == "")
                    {
                        QrCodeText[1] = $"{hinban}_{dateTime:yyyyMMdd_HHmmss}";
                    }
                    break;
                default:
                    break;
            }

            if (QrCodeText[ind] == "")
            {
                pictureBox2.Image = null;
                L_SN.Text = "";
            }
            else
            {
                CM.MakeQRImage(QrCodeText[ind], out Image QrImage);
                pictureBox2.Image = QrImage;
                L_SN.Text = QrCodeText[ind];

            }
        }
        #region method----画像読み込み
        private bool LoadImageToPictureBox(out string mes)
        {
            

            string Hinban = cb_Hinban.Text;
            string appFolderPath = Application.StartupPath;// アプリケーション自身のフォルダのパスを取得
            string picFolderPath = Path.Combine(appFolderPath, "pic");
            string areaName=ra_TofuArea.Checked?"塗布":ra_KinshiArea.Checked?"禁止":""; // 勤務区分の取得

            try
            {
                if (!Directory.Exists(picFolderPath))
                {
                    mes = $"画像ファイル保存フォルダ {picFolderPath} が存在しません。";
                    return false;
                }

                // 指定品番の "_A.png" と "_B.png" のファイルパスを組み立てる
                string[] pngFiles = 
                    [
                    Path.Combine(picFolderPath, $"{Hinban}_A{areaName}.png"), 
                    Path.Combine(picFolderPath, $"{Hinban}_B{areaName}.png")
                    ];

                // すべてのファイルが存在するか確認
                if (!pngFiles.All(File.Exists))
                {
                    mes = $"画像フォルダ内に {Hinban} のPNGファイルが存在しません。\n必要なファイル: {string.Join(", ", pngFiles)}";
                    return false;
                }

                
                switch (Hinban)
                {
                    case "6365590":
                        xy_A = [.. CSet.XY_6365590_A];
                        xy_B = [.. CSet.XY_6365590_B];
                        break;
                    case "6365630":
                        xy_A = [.. CSet.XY_6365630_A];
                        xy_B = [.. CSet.XY_6365630_B];
                        break;
                    default:

                        break;
                }


                //---------------------------------

                ImgGen[0].AddLine("X1", xy_A[0], 0F, xy_A[1], 100F);
                ImgGen[0].AddLine("X2", xy_A[2], 0F, xy_A[3], 100F);
                ImgGen[0].AddLine("Y1", 0F, xy_A[4], 100F, xy_A[5]);
                ImgGen[0].AddLine("Y2", 0F, xy_A[6], 100F, xy_A[7]);
                ImgGen[0].ImageFile = pngFiles[0];
                //picBox_A.Image = ImgGen[0].RotatedImage();

                ImgGen[1].AddLine("X1", xy_B[0], 0F, xy_B[1], 100F);
                ImgGen[1].AddLine("X2", xy_B[2], 0F, xy_B[3], 100F);
                ImgGen[1].AddLine("Y1", 0F, xy_B[4], 100F, xy_B[5]);
                ImgGen[1].AddLine("Y2", 0F, xy_B[6], 100F, xy_B[7]);
                ImgGen[1].ImageFile = pngFiles[1];
                //picBox_B.Image = ImgGen[1].RotatedImage();


                //RotateImage();
                mes = "画像の読み込み成功";
                return true;
            }
            catch (Exception ex)
            {
                // エラーの場合
                mes = $"画像の読み込み中にエラーが発生しました: {ex.Message}";
                return false;
            }
        }
        #endregion
        #region method----画像回転
        private void RotateImage()
        {
            if (!isRotated)
            {
                ImgGen[0].RoteAngle = RotateAngle.None;
                ImgGen[1].RoteAngle = RotateAngle.None;
            }
            else
            {
                ImgGen[0].RoteAngle = RotateAngle.Rotate180;
                ImgGen[1].RoteAngle = RotateAngle.Rotate180;
            }
        }
        #endregion
        #region event-----PicBoxの描画
        //private void PictureBox_Paint(object sender, PaintEventArgs e)
        //{

        //    PictureBox picBox = sender as PictureBox;
        //    if (picBox == null) return; // nullチェック
        //    int ind = int.Parse(picBox.Tag.ToString());
        //    if (picBox.Image == null) return;

        //    Size imageSize = picBox.Image.Size;
        //    Rectangle displayRect = GetImageDisplayRectangle(picBox);


        //    float[] perX0;
        //    float[] perX1;
        //    float[] perY0;
        //    float[] perY1;

        //    // 線の描画
        //    if (!isRotated)
        //    {
        //        perX0 = [xyA[0], xyA[2], 0, 0];
        //        perX1 = [xyA[1], xyA[3], 1, 1];
        //        perY0 = [0, 0, xyA[4], xyA[6]];
        //        perY1 = [1, 1, xyA[5], xyA[7]];
        //    }
        //    else
        //    {
        //        perX0 = [1 - xyA[3], 1 - xyA[1], 0, 0];
        //        perX1 = [1 - xyA[2], 1 - xyA[0], 1, 1];
        //        perY0 = [0, 0, 1 - xyA[7], 1 - xyA[5]];
        //        perY1 = [1, 1, 1 - xyA[6], 1 - xyA[4]];
        //    }

        //    using Pen pen = new(Color.Yellow, 2); // 線の色と太さ
        //    for (int i = 0; i < 4; i++)
        //    {
        //        float x0 = displayRect.Left + perX0[i] * displayRect.Width;
        //        float y0 = displayRect.Top + perY0[i] * displayRect.Height;
        //        float x1 = displayRect.Left + perX1[i] * displayRect.Width;
        //        float y1 = displayRect.Top + perY1[i] * displayRect.Height;
        //        e.Graphics.DrawLine(pen, x0, y0, x1, y1);
        //    }

        //    // NgCounter情報に基づいて塗りつぶし円を描画
        //    foreach (var counter in ngCounters)
        //    {
        //        Brush brush = GetNgTypeBrush(counter.NgType); // NG種類に応じた色を取得
        //        float x = cachedImageRect.X + counter.XPercent / 100 * cachedImageRect.Width;
        //        float y = cachedImageRect.Y + counter.YPercent / 100 * cachedImageRect.Height;

        //        e.Graphics.FillEllipse(brush, x - 1.5f, y - 1.5f, 3, 3);
        //    }
        //}
        #endregion
        #region event-----AとBの入れ替え
        private void Button2_Click(object sender, EventArgs e)
        {
            int[] Cols = [.. picBoxes.Select(x => tableLayoutPanel3.GetColumn(x))];
            int[] Rows = [.. picBoxes.Select(x => tableLayoutPanel3.GetRow(x))];


            int Col1 = tableLayoutPanel3.GetColumn(picBox_A);
            int Row1 = tableLayoutPanel3.GetRow(picBox_A);
            int Col2 = tableLayoutPanel3.GetColumn(picBox_B);
            int Row2 = tableLayoutPanel3.GetRow(picBox_B);

            // 一時的に削除
            tableLayoutPanel3.Controls.Remove(picBox_A);
            tableLayoutPanel3.Controls.Remove(picBox_B);
            tableLayoutPanel3.Controls.Remove(L_A);
            tableLayoutPanel3.Controls.Remove(L_B);

            // 入れ替えて再配置
            tableLayoutPanel3.Controls.Add(picBoxes[0], Cols[1], Rows[1]);
            tableLayoutPanel3.Controls.Add(picBoxes[1], Cols[0], Rows[0]);
            tableLayoutPanel3.Controls.Add(L_A, Cols[1], 0);
            tableLayoutPanel3.Controls.Add(L_B, Cols[0], 0);
        }
        #endregion
        #region event-----画像回転  
        private void B_Rotate_Click(object sender, EventArgs e)
        {
            isRotated = !isRotated;
            RotateImage();

            if (isRotated)
            {
                ImgGen[0].RoteAngle = RotateAngle.Rotate180;
                ImgGen[1].RoteAngle = RotateAngle.Rotate180;
            }
            else
            {
                ImgGen[0].RoteAngle = RotateAngle.None;
                ImgGen[1].RoteAngle = RotateAngle.None;
            }
        }
        #endregion
        

        private void B_Setting_Click(object sender, EventArgs e)
        {
            F_Setting frm = new(NgTexts_jp);
            frm.ShowDialog();
            ColorInit();
        }

        List<float[]> x;
        List<float[]> y;
        long DaichoID = -1;
        private void b_Save_Click(object sender, EventArgs e)
        {
            x = [
                [xy_A[0],xy_A[1] ,xy_A[2], xy_A[3],0,100,0,100],
                [xy_B[0],xy_B[1] ,xy_B[2], xy_B[3],0,100,0,100],
            ];
            y = [
                [0,100,0,100,xy_A[4],xy_A[5] ,xy_A[6], xy_A[7]],
                [0, 100, 0, 100, xy_B[4],xy_B[5] ,xy_B[6], xy_B[7]],
            ];

            //SQLiteCon db = new (CSet.DbPath, CSet.DbTableCFG(), false);
            if (!db.IsConnected)
            {
                MessageBox.Show("DB接続失敗");
                return;
            }


            if(!ra_Day.Checked && !ra_Night.Checked)
            {

                MessageBox.Show("昼勤、夜勤を選択してください");
                return;
            }

            DateTime testDate= dtPicker_TestDate.Value.Date;
            string mkDate = dtPicker_TestDate.Value.ToString("yyyy-MM-dd");
            string svDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string hinban = cb_Hinban.Text;
            int isDay = ra_Day.Checked ? 1 : 0;
            string ope = t_operator.Text;
            if (string.IsNullOrEmpty(ope)) ope = "unknown";


            //現在のデータを削除する
            //この保存メソッドは紙から入力modeなので、operator別の処理は行わない
            List<string> SqlList0 =
                [
                $"DELETE FROM T_Daicho WHERE SN='{SN}';",
                $"DELETE FROM T_Data WHERE dID IN (SELECT ID FROM T_Daicho WHERE SN = '{SN}');"
                ];


            db.Execute(SqlList0);

            string sql = 
                dbCM.MakeInsertSQL(
                    "T_Daicho",
                    new Dictionary<string, object>
                    {
                        { "mode", 0 },
                        { "SN", SN },
                        { "Operator", ope },
                        { "TestDate",testDate },
                        { "SaveDate", DateTime.Now },
                        { "BoardDate",testDate},
                        { "Hinban", hinban },
                        { "isDay", ra_Day.Checked }
                    });
            db.InsertDataAndGetId(sql, out int dID);

            List<string> SqlList = []; 

            string[] board =["A", "B"];

            for (int i = 0; i < 2; i++)
            {

                foreach (var item in ImgGen[i].NgPoints)
                {
                    int areaNo = GetArea(item.XYPercent, i);
                    SqlList.Add(dbCM.MakeInsertSQL(
                        "T_Data",
                        new Dictionary<string, object>
                        {
                            { "dID", dID },
                            { "Board", board[i] },
                            { "NgType", item.NgType },
                            { "NgText", item.NgText },
                            { "X", item.XPercent }, // X座標を計算
                            { "Y", item.YPercent }, // Y座標を計算
                            { "Area", areaNo }
                        }));
                }
            }

            db.Execute(SqlList);

            
            if(!listBox1.Items.Contains(SN))listBox1.Items.Add(SN); // ListBoxの内容をクリア
            setSN(false); // S/Nを更新

            undoList.Clear(); // アンドゥリストをクリア
            ImgGen[0].ClearPoints();
            ImgGen[1].ClearPoints();
            CountUpdate();
        }


        private void b_Mae_Click(object sender, EventArgs e)
        {
            DateTime dTime = dtPicker_TestDate.Value;
            string dTimSt = dTime.ToString("yyyy-MM-dd");

            string sql =
                "SELECT " +
                "  T1.ID AS DaichoID, " +
                "  T1.MakeDate AS MakeDate, " +
                "  T1.SaveDateTime AS SaveDateTime, " +
                "  T1.Hinban AS Hinban, " +
                "  T1.Board AS Board, " +
                "  T1.isDay AS isDay, " +
                "  T2.NgType AS NgType, " +
                "  T2.NgText AS NgText, " +
                "  T2.X AS X, " +
                "  T2.Y AS Y, " +
                "  T2.Area AS Area " +
                "FROM" +
                "            T_Daicho T1" +
                "  LEFT JOIN T_Data   T2 ON T1.ID = T2.DaichoID " +
                "WHERE " +
                $"  T1.ID = (SELECT MAX(ID) FROM T_Daicho WHERE MakeDate='{dTimSt}'";
            if (DaichoID == -1) 
            {
                sql += ");";
            }
            else
            {
                sql += $" AND ID<{DaichoID});";
            }

            SQLiteCon db = new(CSet.DbPath, CSet.DbTableCFG(), false);

            if (!db.GetData(sql, out DataTable dt)) 
            {
                MessageBox.Show("DB読み込みエラー");
                return;
            }

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("もうないよ");
                return;
            }

            ImgGen[0].ClearPoints();
            ImgGen[1].ClearPoints();

            DataRow dr = dt.Rows[0];
            string hinban = dr.Field<string>("Hinban");
            long isDay = dr.Field<long>("isDay");
            cb_Hinban.Text = hinban;
            ra_Day.Checked = (isDay != 0);
            ra_Night.Checked = (isDay == 0);

            int ind;
            foreach (DataRow item in dt.Rows)
            {
                string brd = item.Field<string>("Board");
                ind = brd == "B" ? 1 : 0;
                // 各列のデータを取得
                int ngType = (int)item.Field<long>("NgType");           //NgTypeの値を取得しint型に変換
                string ngText = item.Field<string>("NgText"); //NgTypeCountの値を取得しint型に変換
                float x = (float)item.Field<double>("X");
                float y = (float)item.Field<double>("Y");

                ImgGen[ind].addPoint(new NgCounter(x, y, ngType, ngText));
            }



        }

        private int GetArea(PointF loc, int ind)
        {
            int row = 2;
            int col = 2;

            float[] x0 = [x[ind][0], x[ind][2]];
            float[] x1 = [x[ind][1], x[ind][3]];
            float[] y0 = [y[ind][4], y[ind][6]];
            float[] y1 = [y[ind][5], y[ind][7]];

            for (int i = 0; i < 2; i++)
            {
                float xx = loc.Y / 100 * (x1[i] - x0[i]) + x0[i];
                if (loc.X < xx)
                {
                    col = i;
                    break;
                }
            }

            for (int i = 0; i < 2; i++)
            {
                float yy = loc.X / 100 * (y1[i] - y0[i]) + y0[i];
                if (loc.Y < yy)
                {
                    row = i;
                    break;
                }
            }

            int re = row * 3 + col + 1;

            return re ;
        }

        private void B_Excel_Click(object sender, EventArgs e)
        {
            DateTime mkDt = dtPicker_TestDate.Value;

            string appFolderPath = Application.StartupPath;// アプリケーション自身のフォルダのパスを取得
            string myDocumentsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "NgMapping");
            Directory.CreateDirectory(myDocumentsPath);

            string openExcelFile = Path.Combine(appFolderPath, "【最終検査】マッピング.xlsx");
            string saveExcelFile = Path.Combine(myDocumentsPath, $"【最終検査】マッピング({mkDt:yyMMdd}).xlsx");

            int no = 0;
            while (File.Exists(saveExcelFile))
            {
                saveExcelFile = Path.Combine(myDocumentsPath, $"【最終検査】マッピング({mkDt:yyMMdd})_{no:00}.xlsx");
                no++;
            }

            SQLiteCon db = new(CSet.DbPath, CSet.DbTableCFG(), false);

            string TestDate = dtPicker_TestDate.Value.ToString("yyyy-MM-dd");

            string[] hinban = ["6365590", "6365590", "6365590", "6365590", "6365630", "6365630", "6365630", "6365630"];
            string[] board = ["A", "B", "A", "B", "A", "B", "A", "B"];
            int[] isDay = [1, 1, 0, 0, 1, 1, 0, 0];
            //DataTable[] dt = new DataTable[8];

            List<string> sqlList =[];

            for (int i = 0; i < 8; i++)
            {
                string sql =
                     "SELECT " +
                     "   T2.Area, " +
                     "   T2.NgType, " +
                     "   COUNT(*) AS NgTypeCount " +
                     "FROM " +
                     "              T_Daicho T1 " +
                     "   INNER JOIN T_Data   T2 ON T2.dID = T1.ID " +
                     "WHERE " +
                    $"       T1.TestDate LIKE '{TestDate}%' " +
                    $"   AND T1.Hinban = '{hinban[i]}' " +
                    $"   AND T2.Board = '{board[i]}' " +
                    $"   AND T1.isDay = {isDay[i]} " +
                     "GROUP BY " +
                     "   T2.Area, " +
                     "   T2.NgType; ";
                sqlList.Add(sql);

            }
            if (!db.GetData(sqlList, out DataTable[] dt))
            {
                MessageBox.Show("データ取得失敗");
                return;
            }
            string[] sName1 = ["6365590D(ALL)", "6365630D(ALL)", "6365590D", "6365590N", "6365630D", "6365630N"];
            string[] sName2 = ["6365590D", "6365590N", "6365630D", "6365630N"];

            //ファイルを開く
            using var workbook = new XLWorkbook(openExcelFile);

            List<string> worksheetNames = [];

            //foreach (var worksheet in workbook.Worksheets)
            //{
            //    worksheetNames.Add(worksheet.Name);
            //}


            // 指定されたシートを選択
            foreach (string sName in sName1)
            {
                var wsheet = workbook.Worksheet(sName);
                wsheet.Cell("H2").Value = mkDt; // A1セルに値を書き込む

            }

            for (int i = 0; i < 4; i++)
            {
                var worksheet = workbook.Worksheet(sName2[i]);
                worksheet.Range("I5:Q13").Clear(XLClearOptions.Contents);
                worksheet.Range("I19:Q27").Clear(XLClearOptions.Contents);

                foreach (DataRow item in dt[i * 2].Rows)
                {
                    // 各列のデータを取得
                    int areaNo = (int)item.Field<long>("Area");             //Areaの値を取得しint型に変換
                    int ngType = (int)item.Field<long>("NgType");           //NgTypeの値を取得しint型に変換
                    int ngTypeCount = (int)item.Field<long>("NgTypeCount"); //NgTypeCountの値を取得しint型に変換
                    worksheet.Cell(ngType + 5, areaNo + 8).Value = ngTypeCount;  //セルに値を設定
                }
                foreach (DataRow item in dt[i * 2 + 1].Rows)
                {
                    // 各列のデータを取得
                    int areaNo = (int)item.Field<long>("Area");                 // Areaの値を取得しint型に変換
                    int ngType = (int)item.Field<long>("NgType");       // NgTypeの値を取得しint型に変換
                    int ngTypeCount = (int)item.Field<long>("NgTypeCount"); // NgTypeCountの値を取得しint型に変換
                    worksheet.Cell(ngType + 19, areaNo + 8).Value = ngTypeCount; // セルに値を設定
                }
            }
            // ファイルを保存
            workbook.SaveAs(saveExcelFile);
            //MessageBox.Show($"保存しました。\r\n{saveExcelFile}");

            DialogResult result = MessageBox.Show($"Excelファイルを保存しました。\n\nファイルを開きますか？\n\nファイル場所: {saveExcelFile}",
                                                          "Excelファイル生成",
                                                          MessageBoxButtons.YesNo,
                                                          MessageBoxIcon.Information);

            // Yesが選択された場合、ファイルを開く
            if (result == DialogResult.Yes)
            {
                try
                {
                    // Excelアプリケーションでファイルを開く
                    Process.Start(new ProcessStartInfo
                    {
                        FileName =  saveExcelFile,
                        UseShellExecute = true // 必須：既定のアプリケーションで開く
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ファイルを開けませんでした。\n\nエラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }



        }

        private void B_Clear_Click(object sender, EventArgs e)
        {
            ImgGen[0].ClearPoints();
            ImgGen[1].ClearPoints();
            CountUpdate();
        }

        private void B_Analyze_Click(object sender, EventArgs e)
        {

            MessageBox.Show("まだ使えません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;

            F_Ana frm = new();

            // メインフォームが表示されているスクリーンを取得
            Screen mainScreen = Screen.FromControl(this);

            // ダイアログをメインフォームと同じスクリーンに配置


            frm.StartPosition = FormStartPosition.Manual;
            frm.Location = new System.Drawing.Point(
                mainScreen.Bounds.X + (mainScreen.Bounds.Width - frm.Width) / 2,
                mainScreen.Bounds.Y + (mainScreen.Bounds.Height - frm.Height) / 2);

            frm.ShowDialog();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            CSet.SaveFormLocState(this);
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            loginForm.Close();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {

            int LastIndex = undoList.Count - 1;
            if (LastIndex < 0) return; // アンドゥリストが空の場合は何もしない

            UndoInfo uinfo= undoList[LastIndex]; // NgCounterを解放
            if(uinfo.ListIndex==-1) // addしたとき
            {
                ImgGen[uinfo.PanelIndex].delLastPoint(); // 最後のポイントを削除
            }
            else // DeletePointしたとき
            {
                ImgGen[uinfo.PanelIndex].insertPoint(uinfo.ngCounter,uinfo.ListIndex); // NgCounterを復元
            }
            undoList.RemoveAt(LastIndex); // アンドゥリストから削除
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void ra_Normal_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb = sender as RadioButton;
            if (rb == null || !rb.Checked) return; // nullチェック
            if (!LoadImageToPictureBox(out string mes))
            {
                MessageBox.Show(mes);
            }
        }
    }
    public class UndoInfo
    {
        public int PanelIndex { get; set; } // クリックしたPictureBoxのインデックス（0: A面, 1: B面）
        public int ListIndex { get; set; } // DeletePointしたときのListIndex。addしたときは-1。
        public NgCounter ngCounter { get; set; }//addしたときは、nullのままでOK
        public UndoInfo(int PanelInd, int ListInd, NgCounter ngcnt)
        {
            this.PanelIndex = PanelInd;
            this.ListIndex = ListInd;
            this.ngCounter = ngcnt;
        }
    }
}
