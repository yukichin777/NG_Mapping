using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NGMapping
{
    public partial class F_Ana : Form
    {
        private readonly RotateImageGenerator[] ImgGen = [new(), new()];
        private float[] xy_A = [];
        private float[] xy_B = [];
        private PictureBox[] picBoxes;
        private DataTable[] dt = new DataTable[2];

        public F_Ana()
        {
            InitializeComponent();
            picBoxes = [picBox_A, picBox_B];
            ImgGen[0].pBox = picBoxes[0];
            ImgGen[1].pBox = picBoxes[1];
        }

        private void Cb_Hinban_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 画像を読み込む
            if (!LoadImageToPictureBox(out string mes))
            {
                MessageBox.Show(mes);
            }
        }
        #region method----画像読み込み
        private bool LoadImageToPictureBox(out string mes)
        {
            if (cb_Hinban.SelectedIndex < 0) { mes = "品番が選択されていません"; return false; }

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

        private void Button1_Click(object sender, EventArgs e)
        {
            if(cb_Hinban.SelectedIndex < 0)
            {
                MessageBox.Show("品番が選択されていません");
                return;
            }
            string hinban = cb_Hinban.Text;

            DateTime dt1 = dateTimePicker1.Value;
            DateTime dt2 = dateTimePicker2.Value;
            if (dt1 > dt2)
            {
                CSet.Swap(ref dt1, ref dt2);
                dateTimePicker1.Value = dt1;
                dateTimePicker2.Value = dt2;
            }

            string dt1st = dt1.ToString("yyyy-MM-dd");
            string dt2st = dt2.ToString("yyyy-MM-dd");

            int isDay= ra_Day.Checked ?0:ra_Night.Checked ? 1 : 2;
            string[] st0 = [" AND T1.Board = 'A' ", " AND T1.Board = 'B' "];
            string[] st1 = [" AND T1.isDay = 1;", " AND T1.isDay = 0;", ";"];

            string sql0 =
                "SELECT " +
                "  T1.ID AS DaichoID, " +
                "  T1.MakeDate, " +
                "  T1.SaveDateTime,  " +
                "  T1.Hinban, " +
                "  T1.Board, " +
                "  T1.isDay, " +
                "  T2.ID AS DataID, " +
                "  T2.NgType, " +
                "  T2.NgText, " +
                "  T2.X, " +
                "  T2.Y," +
                "  T2.Area " +
                "FROM" +
                "             T_Daicho T1 " +
                "  INNER JOIN T_Data   T2 ON T1.ID = T2.DaichoID " +
                "WHERE " +
                $"  T1.MakeDate BETWEEN '{dt1st}' AND '{dt2st}'" +
                $"    AND T1.Hinban = '{hinban}' " +
                "    AND T1.Board = '基板名' " +
                "    AND T1.isDay = '昼勤or夜勤区分';";



            SQLiteCon db = new(CSet.DbPath, CSet.DbTableCFG(), false);            
            for (int i = 0; i < dt.Length; i++)
            {
                string sql1 = sql0 + st0[i]+st1[isDay];

                if (!db.GetData(sql1, out dt[i]))
                {
                    MessageBox.Show("データ取得失敗");
                    return;
                }
            }
            SetSplitLine();//分割線の設定(9分割固定)
            DrawBubleImage();



        }

        private void DrawBubleImage()
        {
            int NgType = GetNgType();
            List<int[]>[] ngCount = NgCount(dt, NgType);//NG数のカウント

            for (int i = 0; i < 2; i++)
            {
                for (int j = 0; j < ngCount[i].Count; j++)
                {
                    for (int k = 0; k < ngCount[i][j].Length; k++)
                    {
                        if (ngCount[i][j][k] > 0)
                        {
                            ImgGen[i].AddBubbleData(xy[i][j][k],  ngCount[i][j][k]);
                        }
                    }
                }
            }
        }

        private int GetNgType()
        {
            if (radioButton1.Checked) return 0;
            if (radioButton2.Checked) return 1;
            if (radioButton3.Checked) return 2;
            if (radioButton4.Checked) return 3;
            if (radioButton5.Checked) return 4;
            if (radioButton6.Checked) return 5;
            if (radioButton7.Checked) return 6;
            if (radioButton8.Checked) return 7;
            return -1;
        }


        List<float[]> SpX0;
        List<float[]> SpX1;
        List<float[]> SpY0;
        List<float[]> SpY1;
        List<PointF[]>[] xy;

        private void SetSplitLine()
        {
            SpX0 = [Split9(xy_A[0], xy_A[2]), Split9(xy_B[0], xy_B[2])];
            SpX1 = [Split9(xy_A[1], xy_A[3]), Split9(xy_B[1], xy_B[3])];
            SpY0 = [Split9(xy_A[4], xy_A[6]), Split9(xy_B[4], xy_B[6])];
            SpY1 = [Split9(xy_A[5], xy_A[7]), Split9(xy_B[5], xy_B[7])];


            float[] x0 = [0f, .. SpX0[0], 100f];
            float[] x1 = [0f, .. SpX1[0], 100f];
            float[] y0 = [0f, .. SpY0[0], 100f];
            float[] y1 = [0f, .. SpY1[0], 100f];


            xy = new List<PointF[]>[2];
            for (int ind = 0; ind < 2; ind++)
            {
                xy[ind] = [];
                for (int i = 0; i < y0.Length - 1; i++)
                {
                    PointF[] xy0 = new PointF[x0.Length - 1];
                    for (int j = 0; j < x0.Length - 1; j++)
                    {
                        PointF px1 = new (x0[j], 0);
                        PointF px2 = new (x1[j], 100);
                        PointF px3 = new (x0[j + 1], 0);
                        PointF px4 = new (x1[j + 1], 100);

                        PointF py1 = new (0, y0[i]);
                        PointF py2 = new (100, y1[i]);
                        PointF py3 = new (0, y0[i + 1]);
                        PointF py4 = new (100, y1[i + 1]);


                        PointF? pp0 = GetIntersection(px1, px2, py1, py2);
                        PointF? pp1 = GetIntersection(px1, px2, py3, py4);
                        PointF? pp2 = GetIntersection(px3, px4, py1, py2);
                        PointF? pp3 = GetIntersection(px3, px4, py3, py4);
                        if (pp0 != null && pp1 != null && pp2 != null && pp3 != null)
                        {
                            // 交点がすべて存在する場合の処理
                            PointF p0 = (PointF)pp0;
                            PointF p1 = (PointF)pp1;
                            PointF p2 = (PointF)pp2;
                            PointF p3 = (PointF)pp3;
                            PointF? pp = GetIntersection(p0, p3, p1, p2);

                            if (pp != null)
                            {
                                xy0[j] = (PointF)pp;
                            }
                        }
                    }
                    xy[ind].Add(xy0);
                }
            }



        }
        private float[]Split9(float a,float b)
        {
            float[] re = [
                a / 3,
                a * 2 / 3,
                a,
                a + (b - a) / 3,
                a + (b - a)*2 / 3,
                b,
                100-  (100-b)*2 / 3,
                100-  (100-b) / 3
                ];
            return re;
        }

        public static PointF? GetIntersection(PointF loc1, PointF loc2, PointF loc3, PointF loc4)
        {
            // 線1の方程式: A1x + B1y = C1
            float A1 = loc2.Y - loc1.Y;
            float B1 = loc1.X - loc2.X;
            float C1 = A1 * loc1.X + B1 * loc1.Y;

            // 線2の方程式: A2x + B2y = C2
            float A2 = loc4.Y - loc3.Y;
            float B2 = loc3.X - loc4.X;
            float C2 = A2 * loc3.X + B2 * loc3.Y;

            // 行列を使って交点を計算するためのデルタを求める
            float determinant = A1 * B2 - A2 * B1;

            // determinantが0の場合、線は平行または同一
            if (Math.Abs(determinant) < 1e-6)
            {
                return null; // 交点なし
            }

            // 交点の座標を計算
            float x = (B2 * C1 - B1 * C2) / determinant;
            float y = (A1 * C2 - A2 * C1) / determinant;

            return new PointF(x, y);
        }



        private List<int[]>[] NgCount(DataTable[] dts,int NgType)
        {
            List<int[]>[] re = new List<int[]>[2];

            for (int i = 0; i < dts.Length; i++)
            {
                re[i] = [];
                for (int j = 0; j < 9; j++) { re[i].Add(new int[9]); }
                foreach (DataRow item in dts[i].Rows)
                {
                    if(item.Field<int>("NgType") != NgType && NgType!=-1) continue;

                    PointF loc = new ((float)item.Field<double>("X"), (float)item.Field<double>("Y"));
                    int[] xyInd = GetArea(loc, i);
                    re[i][xyInd[1]][xyInd[0]]++;
                }

            }
            return re;

        }

        private int[] GetArea(PointF loc, int ind)
        {
            int[] XY = [8,8];

            for (int i = 0; i < SpX0[ind].Length; i++)
            {
                float xx = loc.Y / 100 * (SpX1[ind][i] - SpX0[ind][i]) + SpX0[ind][i];
                if (loc.X < xx)
                {
                    XY[0] = i;
                    break;
                }
            }

            for (int i = 0; i < SpY0[ind].Length; i++)
            {
                float yy = loc.X / 100 * (SpY1[ind][i] - SpY0[ind][i]) + SpY0[ind][i];
                if (loc.Y < yy)
                {
                    XY[1] = i;
                    break;
                }
            }

            return XY;
        }

        private void F_Ana_Load(object sender, EventArgs e)
        {

        }
    }
}
