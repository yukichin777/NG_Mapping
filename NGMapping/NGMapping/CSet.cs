using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Drawing;
using System.Windows.Forms; // JSONシリアライズライブラリ
using QRCoder;
using System;         // QRCoderライブラリを使用


namespace NGMapping
{
    public static class CSet
    {
        public static void Swap<T>(ref T a, ref T b)
        {
            T temp = a; // 一時変数を使用して値を保存
            a = b;      // bの値をaに代入
            b = temp;   // 保存したaの値をbに代入
        }


        public static void SetFormLocState(Form frm)
        {
            Point loc= Point.Empty;
            Size siz=Size.Empty;
            FormWindowState state = FormWindowState.Normal; // デフォルトのウィンドウ状態


            switch (frm.Name.ToLower())
            {
                case "mainform":
                    loc = Properties.Settings.Default.WindowLocation;
                    siz = Properties.Settings.Default.WindowSize;
                    string dum = Properties.Settings.Default.WindowState;
                    if (dum == FormWindowState.Maximized.ToString()) state= FormWindowState.Maximized;else state= FormWindowState.Normal;
                    if (loc != Point.Empty && siz != Size.Empty && IsOnScreen(loc, siz))
                    {
                        frm.Location = loc;
                        frm.Size = siz;
                    }
                    else
                    {
                        frm.StartPosition = FormStartPosition.CenterScreen;
                    }
                    break;
                case "f_login":
                    loc = Properties.Settings.Default.WindowLocation;

                    // フォームを対象スクリーンの中央に配置
                    Screen targetScreen = Screen.AllScreens.FirstOrDefault(screen =>
                        screen.WorkingArea.Contains(loc)) ?? Screen.PrimaryScreen;

                    frm.StartPosition = FormStartPosition.Manual; // 手動で位置を指定
                    frm.Location = new Point(
                        targetScreen.WorkingArea.X + (targetScreen.WorkingArea.Width - frm.Width) / 2,
                        targetScreen.WorkingArea.Y + (targetScreen.WorkingArea.Height - frm.Height) / 2
                    );
                    break;
                default:
                    break;
            }

            // ウィンドウ状態を復元
            frm.WindowState = state;
        }
        private static bool IsOnScreen(Point location, Size size)
        {
            Rectangle rect = new(location, size);

            foreach (var screen in Screen.AllScreens)
            {
                if (screen.WorkingArea.IntersectsWith(rect))
                {
                    return true;
                }
            }

            return false;
        }



        public static void SaveFormLocState(Form frm)
        {
            switch (frm.Name.ToLower())
            {
                case "mainform":
                    // フォームの位置を保存
                    Properties.Settings.Default.WindowLocation = frm.WindowState == FormWindowState.Normal ? frm.Location : frm.RestoreBounds.Location;

                    // フォームのサイズを保存
                    Properties.Settings.Default.WindowSize = frm.WindowState == FormWindowState.Normal ? frm.Size : frm.RestoreBounds.Size;

                    // フォームのウィンドウ状態を保存
                    Properties.Settings.Default.WindowState = frm.WindowState.ToString();

                    // 変更を適用
                    Properties.Settings.Default.Save();
                    break;
                default:
                    break;
            }
        }

        public static string Language
        {
            get { return Properties.Settings.Default.Language;  }
            set
            {
                Properties.Settings.Default.Language = value;
                Properties.Settings.Default.Save();
            }
        }
        public static DateTime[] DayTimeRange
        {
            get 
            {
                string[] dum = Properties.Settings.Default.DayTimeRange.Split(',');
                if(dum.Length == 2 && DateTime.TryParse(dum[0], out DateTime start) && DateTime.TryParse(dum[1], out DateTime end))
                {
                    return [start, end]; // 正しい形式ならばDateTime配列を返す
                }
                else
                {
                    // デフォルトの時間範囲を返す
                    return [new DateTime(2025, 1, 1, 8, 30, 0), new DateTime(2025, 1, 1, 20,30, 0)];
                }            
            }
            set
            {
                if (value.Length == 2)
                {
                    Properties.Settings.Default.DayTimeRange = value[0].ToString("HH:mm") + "," + value[1].ToString("HH:mm");
                    Properties.Settings.Default.Save();
                }
            }
        }
 





        public static float[] XY_6365590_A
        {
            get
            {
                string[] sp = Properties.Settings.Default.XY_6365590_A.Split(',');
                if (sp.Length == 8 && sp.All(x => float.TryParse(x, out _)))
                {
                    float[] xy = [.. sp.Select(x => float.Parse(x))];
                    return xy;
                }
                else
                {
                    return [32.23F, 32.45F, 65.09F, 64.40F, 34.22F, 33.66F, 68.55F, 68.20F];
                }
            }
            set
            {
                if (value.Length == 8)
                {
                    Properties.Settings.Default.XY_6365590_A = string.Join(",", value);
                    Properties.Settings.Default.Save();
                }
            }
        }
        public static float[] XY_6365590_B
        {
            get
            {
                string[] sp = Properties.Settings.Default.XY_6365590_B.Split(',');
                if (sp.Length == 8 && sp.All(x => float.TryParse(x, out _)))
                {
                    float[] xy = [.. sp.Select(x => float.Parse(x))];
                    return xy;
                }
                else
                {
                    return [32.24F, 31.06F, 66.16F, 65.49F, 39.21F, 41.23F, 74.04F, 75.93F];
                }
            }
            set
            {
                if (value.Length == 8)
                {
                    Properties.Settings.Default.XY_6365590_B = string.Join(",", value);
                    Properties.Settings.Default.Save();
                }
            }
        }
        public static float[] XY_6365630_A
        {
            get
            {
                string[] sp = Properties.Settings.Default.XY_6365630_A.Split(',');
                if (sp.Length == 8 && sp.All(x => float.TryParse(x, out _)))
                {
                    float[] xy = [.. sp.Select(x => float.Parse(x))];
                    return xy;
                }
                else
                {
                    return [28F, 28.74F, 56.38F, 56.88F, 34.19F, 34.25F, 70.96F, 70.74F];
                }
            }
            set
            {
                if (value.Length == 8)
                {
                    Properties.Settings.Default.XY_6365630_A = string.Join(",", value);
                    Properties.Settings.Default.Save();
                }
            }
        }
        public static float[] XY_6365630_B
        {
            get
            {
                string[] sp = Properties.Settings.Default.XY_6365630_B.Split(',');
                if (sp.Length == 8 && sp.All(x => float.TryParse(x, out _)))
                {
                    float[] xy = [.. sp.Select(x => float.Parse(x))];
                    return xy;
                }
                else
                {
                    return [35.41F, 35.49F, 65.65F, 65.99F, 31.67F, 33.13F, 69.26F, 71.37F];
                }
            }
            set
            {
                if (value.Length == 8)
                {
                    Properties.Settings.Default.XY_6365630_B = string.Join(",", value);
                    Properties.Settings.Default.Save();
                }
            }
        }
        public static Dictionary<string, int> NgItem
        {
            get
            {
                string json = Properties.Settings.Default.NgDictionary; // JSON形式の文字列を取得
                if (string.IsNullOrEmpty(json))
                {
                    Dictionary<string, int> re = [];

                    //NgTexts = ["はんだボール", "はんだ屑", "炭化物", "ヒュミシール付着", "ヒュミシール未塗布", "浮き", "破損", "リードカット異常"];
                    re.Add("異物", 2); // 異物は炭化物とする
                    re.Add("フラックス", 3); // フラックスはヒュミシール付着とする
                    return re; // 空の辞書を返す
                }
                if (json == "nothing")
                {
                    return []; // 空の辞書を返す
                }
                // JSON文字列からDictionaryにデシリアライズ
                return JsonSerializer.Deserialize<Dictionary<string, int>>(json) ?? [];
            }
            set
            {
                // DictionaryオブジェクトをJSON形式に変換して保存
                if (value.Count == 0)
                {
                    Properties.Settings.Default.NgDictionary = "nothing"; // 空の辞書を保存
                }
                else
                {
                    Properties.Settings.Default.NgDictionary = JsonSerializer.Serialize(value);
                }
                Properties.Settings.Default.Save(); // 保存を確定
            }
        }

        public static string DbPath
        {
            get { return Properties.Settings.Default.DbPath; }
            set
            {
                Properties.Settings.Default.DbPath = value;
                Properties.Settings.Default.Save();
            }
        }

        public static Color[] NgColors
        {
            get
            {
                // Settingsから保存されている色情報（カンマ区切りで保存されたint値）を取得し分割
                string[] sp = Properties.Settings.Default.NgColors.Split(',');

                // 分割した文字列をint型に変換し、それをColor型に戻す
                if (sp.Length == 8 && sp.All(x => int.TryParse(x, out _))) // 全て正しいint型かチェック
                {
                    Color[] colors = [..sp.Select(x => Color.FromArgb(int.Parse(x)))];
                    return colors;
                }
                else
                {
                    // デフォルトの色配列を返す
                    return [Color.Blue, Color.Green, Color.Purple, Color.Red, Color.Magenta, Color.Cyan, Color.Yellow, Color.Orange];
                }
            }
            set
            {
                // 渡された値が正しい場合（8個のColor型がある場合）、Settingsに保存
                Color[] defColor= [Color.Blue, Color.Green, Color.Purple, Color.Red, Color.Magenta, Color.Cyan, Color.Yellow, Color.Orange];

                if (value.Length == 8)
                {
                    for(int i = 0; i < value.Length; i++)
                    {
                        if (value[i]!=Color.Transparent)
                        {
                            defColor[i] = value[i];
                        }
                    }
                    // 各ColorをARGB整数値に変換し、カンマ区切りの文字列として保存
                    Properties.Settings.Default.NgColors = string.Join(",", defColor.Select(c => c.ToArgb()));
                    // 設定を保存
                    Properties.Settings.Default.Save();
                }
            }
        }


        public static List<TableInfo>  DbTableCFG()
        {
            // テーブル情報を定義
            List<TableInfo> tables =
            [
                // T_Daicho テーブル定義
                // mode  0:紙から入力, 1:QRコードから入力
                // SN  シリアル番号(mode=0のときはDummySN、mode=1の時はQRからの読み取りSN)
                // Operator  オペレータ名
                // TestDate  検査実施日時
                // SaveDateTime  データ保存日時
                // BoardDate  ボード製造日
                // Hinban  品番
                // isDay  昼勤かどうか(true:昼勤, false:夜勤)
                new TableInfo(
                    "T_Daicho",
                    [
                        new("ID", DataType.Integer, isNullable: false, isPrimaryKey: true, isAutoIncrement: true),
                        new("mode", DataType.Integer, isNullable:false),
                        new("QRText", DataType.Text, isNullable: false,maxLength:50),
                        new("SN", DataType.Text, isNullable: false,maxLength:50),
                        new("Operator", DataType.Text, isNullable: false,maxLength:50),
                        new("TestDate", DataType.DateTime, isNullable: false),
                        new("SaveDate", DataType.DateTime, isNullable: false),
                        new("BoardDate", DataType.DateTime, isNullable: false),
                        new("LineName", DataType.Text, isNullable: true,maxLength:2),
                        new("Hinban", DataType.Text, isNullable: true,maxLength:20),
                        new("isDay", DataType.Boolean, isNullable: false)
                    ]
                ),
                // T_Data テーブル定義
                // dID  台帳ID（T_DaichoのID）
                // Board  ボード名（例：6365590A, 6365630Bなど）
                // NgType  NGの種類（0:異物, 1:フラックス, 2:はんだボール, 3:はんだ屑, 4:炭化物, 5:ヒュミシール付着, 6:ヒュミシール未塗布, 7:浮き, 8:破損, 9:リードカット異常）
                // NgText  NGの詳細テキスト（例：はんだボール, フラックスなど）
                // X  X座標
                // Y  Y座標
                // Area  エリア番号（1～9）
                new TableInfo(
                    "T_Data",
                    [
                        new("ID", DataType.Integer, isNullable : false, isPrimaryKey : true, isAutoIncrement : true),
                        new("dID", DataType.Integer, isNullable: false),
                        new("Board", DataType.Text, isNullable: true, maxLength : 2),
                        new("NgType", DataType.Integer, isNullable: false),
                        new("NgText", DataType.Text, isNullable: true, maxLength : 20),
                        new("X", DataType.Double, isNullable: false),
                        new("Y", DataType.Double, isNullable: false),
                        new("Area", DataType.Integer, isNullable: false),
                    ]
                )
            ];

            return tables;
        }       

        public static bool FLG_DispDummyQR
        {
            get { return Properties.Settings.Default.FLG_DispQRCode; }
            set
            {
                Properties.Settings.Default.FLG_DispQRCode = value;
                Properties.Settings.Default.Save();
            }
        }
    }

    public static class CM
    {

        public static bool MakeQRImage(string txt, out Image QrImg)
        {
            QrImg = null; // 初期化
            try
            {
                // QRコード生成ロジック
                using QRCodeGenerator qrGenerator = new();
                // 入力文字列からQRコードデータを生成
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(txt, QRCodeGenerator.ECCLevel.Q);

                // QRコードから画像を生成
                using QRCode qrCode = new(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(20); // 20はピクセルサイズ（調整可能）
                QrImg = (Image)qrCodeImage.Clone(); // BitmapをImageに変換

                return true; // 成功
            }
            catch
            {
                return false; // 何らかのエラーが発生
            }
        }

    }

}
