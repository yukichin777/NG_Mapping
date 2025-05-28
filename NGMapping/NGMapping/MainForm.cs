using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
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
        readonly List<string> NgTexts_jp = ["はんだボール",  "はんだ屑",         "炭化物",    "ヒュミシール付着",  "ヒュミシール未塗布",      "浮き",                      "破損",   "リードカット異常"];
        readonly List<string> NgTexts_en = ["Solder ball",   "Solder residue",   "Carbide",   "Humiseal adhesion",  "No Humiseal adhesion",   "Lift or Floating",          "Damage", "Lead cutting abnormality"];
        readonly List<string> NgTexts_pt = ["Bola de solda", "Resíduo de solda", "Carboneto", "Adesão de Humiseal", "Sem adesão de Humiseal", "Levantamento ou Flutuação", "Danos",  "Anormalidade no corte de terminal"];
        readonly List<string> etcTexts_jp = ["異物","汚れ","フラックス"];
        readonly List<string> etcTexts_en = [ "Foreign material", "Dirt/Contamination", "Flux"];
        readonly List<string> etcTexts_pt = [ "Material estranho", "Sujeira/Contaminação", "Fluxo"];


        bool isQRDisp = false; // QRコードの表示フラグ
        bool isQRead = false; // QRコード読み取りフラグ
        string NowSerial = ""; // シリアル番号

        SQLiteCon db;




        #region constructor
        public MainForm()
        {
            InitializeComponent();
            LCountA = [L_CountA_0, L_CountA_1, L_CountA_2, L_CountA_3, L_CountA_4, L_CountA_5, L_CountA_6, L_CountA_7,L_CountA];
            LCountB = [L_CountB_0, L_CountB_1, L_CountB_2, L_CountB_3, L_CountB_4, L_CountB_5, L_CountB_6, L_CountB_7, L_CountB];
            LCount = [L_Count_0, L_Count_1, L_Count_2, L_Count_3, L_Count_4, L_Count_5, L_Count_6, L_Count_7, L_Count];

            picBoxes = [picBox_A, picBox_B];
            ImgGen[0].pBox = picBox_A;
            ImgGen[1].pBox = picBox_B;

            Ra_NgItems = [radioButton1, radioButton2, radioButton3, radioButton4, radioButton5, radioButton6, radioButton7, radioButton8, radioButton9];
            La_NgItems = [L_Color1, L_Color2, L_Color3, L_Color4, L_Color5, L_Color6, L_Color7, L_Color8];
        }
        #endregion
        #region event-----Formロード
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


            toolStripComboBox1.ComboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;

            Init();
        }

        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string lang = ((ComboBox)sender).SelectedItem.ToString();
            List<string> ngtxt;
            int wd=130;

            switch (lang)
            {
                case "English":
                    ngtxt = NgTexts_en;
                    ngtxt.Add("Eraser"); // 英語のNGテキストを追加
                    wd = 160;
                    break;
                case "Português":
                    ngtxt = NgTexts_pt;
                    ngtxt.Add("Borracha"); // ポルトガル語のNGテキストを追加
                    wd = 160;
                    break;
                case "Japanease":
                default:
                    ngtxt = NgTexts_jp;
                    ngtxt.Add("消しゴム"); // 日本語のNGテキストを追加
                    ngtxt = NgTexts_jp; break;
                    wd = 130;
            }

            for (int i = 0; i < Ra_NgItems.Length; i++)
            {
                Ra_NgItems[i].Text = ngtxt[i];
                
            }
            //tableLayoutPanel2.ColumnStyles[5].SizeType = SizeType.Percent; // サイズを％で設定可能にする
            tableLayoutPanel2.ColumnStyles[11].Width = wd;
            tableLayoutPanel2.ColumnStyles[11].Width = wd;

        }


        #endregion
        #region method-----初期化
        private void Init()
        {

            db = new(CSet.DbPath, CSet.DbTableCFG(), false);
            ColorInit(); // 色の初期化とコンテキストメニューの作成

            isQRead = CSet.isQRCodeReadMode;
            isQRDisp = isQRead && CSet.FLG_DispDummyQR;
            menu_Save.Enabled = !isQRead; // QRコード読み取りモードでは保存ボタンを無効化

        }
        #endregion
        #region method-----色初期化
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

        #region event-----コンテキストメニューのクリックイベント
        private void MenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem menuItem = sender as ToolStripMenuItem;

            int ngType = (int)menuItem.Tag; // NG種類を取得
            string ngText = menuItem.Text; // NGテキストを取得

            if (lastRightClickPoint.IsEmpty || !cachedImageRect.Contains(lastRightClickPoint)) return;
            ImgGen[pBoxInd].addPoint(lastRightClickPoint, ngType, ngText);
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
                    ImgGen[pBoxInd].DeletePoint(e.Location);
                }
                else
                {
                    if (ngType == -1)
                    {
                        MessageBox.Show("NG種類を選択してください。");
                        return;
                    }
                    ImgGen[pBoxInd].addPoint(e.Location, ngType, NgTexts_jp[ngType]);
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

            NowSerial = "";

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
                        $"ORDER BY SaveDateTime DESC;";


                    db.GetData(sql, out DataTable dtSNList);

                    listBox1.Items.Clear(); // ListBoxの内容をクリア
                    listBox1.Items.AddRange([.. dtSNList.Rows.Cast<DataRow>().Select(row => row.Field<string>("SN"))]);
                }
                NowSerial = DummySN;

            }

            L_SN.Text = NowSerial; // ラベルに現在のシリアル番号を表示


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
            NowSerial = sn;
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
                    Path.Combine(picFolderPath, $"{Hinban}_A.png"), 
                    Path.Combine(picFolderPath, $"{Hinban}_B.png")
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

            string mkDate = dtPicker_TestDate.Value.ToString("yyyy-MM-dd");
            string svDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string hinban = cb_Hinban.Text;
            int isDay = ra_Day.Checked ? 1 : 0;


            //現在のデータを削除する
            List<string> SqlList0 =
                [
                $"DELETE FROM T_Daicho WHERE SN='{NowSerial}';",
                $"DELETE FROM T_Data WHERE SN='{NowSerial}';",
                ];
            db.Execute(SqlList0);

            List<string> SqlList = 
                [
                $"INSERT INTO T_Daicho (SN,TestDate, SaveDateTime,Hinban,isDay) VALUES ('{NowSerial}','{mkDate}', '{svDate}','{hinban}',{isDay});"
                ];

            string[] board =["A", "B"];

            for (int i = 0; i < 2; i++)
            {

                foreach (var item in ImgGen[i].NgPoints)
                {
                    int areaNo = GetArea(item.XYPercent, i);
                    SqlList.Add($"INSERT INTO T_Data (SN,Board, NgType,NgText,X,Y,Area) VALUES ('{NowSerial}','{board[i]}', {item.NgType},'{item.NgText}',{item.XPercent},{item.YPercent},{areaNo});");
                }
            }

            db.Execute(SqlList);

            
            if(!listBox1.Items.Contains(NowSerial))listBox1.Items.Add(NowSerial); // ListBoxの内容をクリア
            setSN(false); // S/Nを更新

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

            string mkDate = dtPicker_TestDate.Value.ToString("yyyy-MM-dd");

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
                     "   INNER JOIN T_Data   T2 ON T2.SN = T1.SN " +
                     "WHERE " +
                    $"       T1.TestDate = '{mkDate}' " +
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

       
    }
}
