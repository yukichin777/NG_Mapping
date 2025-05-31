using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;

namespace NGMapping
{
    public enum RotateAngle
    {
        None = 0,
        Rotate90 = 90,
        Rotate180 = 180,
        Rotate270 = 270
    }
    public class RotateImageGenerator
    {
        private Image _OriginalImage;
        private string _ImageFile = "";
        private RotateAngle _RoteAngle = RotateAngle.None;
        private readonly List<LineOnPicture> _Lines = [];
        private readonly List<NgCounter> _NgInfo = [];

        private readonly List<BubbleData>_BubbleDatas =[];
        private readonly int splitCount = 9; // バブルの最大直径を決定するための分割数

        private static Color[] _PointColors = [Color.Blue, Color.Green, Color.Purple, Color.Red, Color.Magenta, Color.Cyan, Color.Yellow, Color.Orange];
        public PictureBox pBox = null;

        public RotateImageGenerator()
        {
        }
        public List<NgCounter>NgPoints
        {
            get { return _NgInfo; }
        }
        public Color[] PointColors
        {
            get { return _PointColors; }
            set
            {
                if (value.Length == 8)
                {
                    _PointColors = value;
                }
            }
        }
        public string ImageFile
        {
            get { return _ImageFile; }
            set
            {
                _ImageFile = value;
                if (string.IsNullOrWhiteSpace(_ImageFile))
                {
                    _OriginalImage = null;
                    return;
                }
                try
                {
                    _OriginalImage = Image.FromFile(_ImageFile);
                }
                catch (Exception)
                {
                    _OriginalImage = null;
                    return;
                }
                refresh();
            }
        }
        public RotateAngle RoteAngle
        {
            get { return _RoteAngle; }
            set { _RoteAngle = value; refresh(); }
        }
        public void AddLine(string tag, float? perX0 = null, float? perY0 = null, float? perX1 = null, float? perY1 = null, Color? lineColor = null, float? lineWidth = null)
        {
            var existingLine = _Lines.FirstOrDefault(line => line.Tag == tag);

            if (existingLine != null)
            {
                // 更新処理: 既存の線を修正する
                if (perX0.HasValue) existingLine.perX0 = perX0.Value;
                if (perY0.HasValue) existingLine.perY0 = perY0.Value;
                if (perX1.HasValue) existingLine.perX1 = perX1.Value;
                if (perY1.HasValue) existingLine.perY1 = perY1.Value;
                if (lineColor.HasValue) existingLine.LineColor = lineColor.Value;
                if (lineWidth.HasValue) existingLine.LineWidth = lineWidth.Value;
            }
            else
            {
                // 新規追加処理: 必須項目が指定されていない場合は例外をスロー
                if (!perX0.HasValue || !perY0.HasValue || !perX1.HasValue || !perY1.HasValue)
                {
                    throw new ArgumentException("New line requires all coordinates (perX0, perY0, perX1, perY1) to be specified.");
                }
                LineOnPicture newLine = new()
                {
                    Tag = tag,
                    perX0 = perX0.Value,
                    perY0 = perY0.Value,
                    perX1 = perX1.Value,
                    perY1 = perY1.Value,
                    LineColor = lineColor ?? Color.Yellow,
                    LineWidth = lineWidth ?? 2.0F
                };
                _Lines.Add(newLine);
            }
            refresh();
        }
        public void DeleteLine(string tag)
        {
            var lineToRemove = _Lines.FirstOrDefault(line => line.Tag == tag);
            if (lineToRemove != null)
            {
                _Lines.Remove(lineToRemove);
                refresh();
            }
        }
        public void UpdateLine(string tag, float? perX0 = null, float? perY0 = null,
                           float? perX1 = null, float? perY1 = null,
                           Color? lineColor = null, float? lineWidth = null)
        {
            var lineToUpdate = _Lines.FirstOrDefault(line => line.Tag == tag);
            if (lineToUpdate != null)
            {
                if (perX0.HasValue) lineToUpdate.perX0 = perX0.Value;
                if (perY0.HasValue) lineToUpdate.perY0 = perY0.Value;
                if (perX1.HasValue) lineToUpdate.perX1 = perX1.Value;
                if (perY1.HasValue) lineToUpdate.perY1 = perY1.Value;
                if (lineColor.HasValue) lineToUpdate.LineColor = lineColor.Value;
                if (lineWidth.HasValue) lineToUpdate.LineWidth = lineWidth.Value;
                refresh();
            }
        }
        public void ClearLines()
        {
            _Lines.Clear();
            refresh();
        }        
        public void AddBubbleData(PointF loc, float value)
        {
            _BubbleDatas.Add(new BubbleData(loc, value));
            refresh();
        }
        public void AddBubbleData(BubbleData bData)
        {
            _BubbleDatas.Add(bData);
            refresh();
        }
        public void AddBubbleData(List<BubbleData>bDatas)
        {
            _BubbleDatas.Clear();
            _BubbleDatas.AddRange(bDatas);
            refresh();
        }
        public void ClearBubbleData()
        {
            _BubbleDatas.Clear();
            refresh();
        }
        public void ClearPoints()
        {
            _NgInfo.Clear();
            refresh();
        }
        public int[] NgCount
        {
            get
            {
                // PointColors.Length に合わせた配列を生成
                int[] counts = new int[_PointColors.Length];

                foreach (var  item in _NgInfo)
                {
                    if(item.NgType >= 0 && item.NgType < counts.Length)
                    {
                        counts[item.NgType]++;
                    }
                }

                // 各Indexのカウントを計算
                //foreach (var (_, index) in _PointsWithIndex)
                //{
                //    if (index >= 0 && index < counts.Length)
                //    {
                //        counts[index]++;
                //    }
                //}
                return counts;
            }
        }

        public void insertPoint(NgCounter ngc,int ListIndex)
        {
            _NgInfo.Insert(ListIndex,ngc);
            refresh();
        }

        public void addPoint(NgCounter ngc) 
        {
            _NgInfo.Add(ngc);
            refresh();
        }
        public void addPoint(Point Loc, int NgType,string NgText)
        {
            if (_OriginalImage == null) return;
            if (NgType < 0 || NgType > 7) NgType = 0;

            // 表示領域を取得
            Rectangle displayRect = GetImageDisplayRectangle();

            // 入力座標を画像表示エリア内の座標に変換
            float normalizedX = (Loc.X - displayRect.Left) / (float)displayRect.Width;
            float normalizedY = (Loc.Y - displayRect.Top) / (float)displayRect.Height;

            //正規化された座標が表示領域外なら無視
            if (normalizedX < 0 || normalizedX > 1 || normalizedY < 0 || normalizedY > 1) return;

            //回転を考慮して元の画像の座標に逆変換
            PointF originalPoint = TransformPointBack(new PointF(normalizedX, normalizedY), _RoteAngle);
            // 比率(%)で保持
            float percentX = originalPoint.X * 100f;
            float percentY = originalPoint.Y * 100f;
            
            _NgInfo.Add(new NgCounter(percentX, percentY, NgType,NgText) );

            //_PointsWithIndex.Add((new PointF(percentX, percentY), NgType));
            refresh();
        }

        #region method----Pointを削除する(Locに一番近い点を削除する。削除したListIndexを返す)
        public int DeletePoint(Point Loc,out NgCounter ngcnt)
        {
            if (_NgInfo.Count == 0) { ngcnt = null; return -1; }

            // 表示領域を取得
            Rectangle displayRect = GetImageDisplayRectangle();

            // 入力座標を画像表示エリア内の座標に変換
            float normalizedX = (Loc.X - displayRect.Left) / (float)displayRect.Width;
            float normalizedY = (Loc.Y - displayRect.Top) / (float)displayRect.Height;

            // 正規化された座標が表示領域外なら何もしない
            if (normalizedX < 0 || normalizedX > 1 || normalizedY < 0 || normalizedY > 1) { ngcnt = null; return -1; }

            // 回転を考慮して元の画像の座標に逆変換
            PointF originalPoint = TransformPointBack(new PointF(normalizedX, normalizedY), _RoteAngle);

            // 比率(%)で表現する
            float percentX = originalPoint.X * 100f;
            float percentY = originalPoint.Y * 100f;

            // 最も近いポイントとそのインデックスを探す
            int closestIndex = -1;
            float minDistanceSquared = float.MaxValue;


            for (int i = 0; i < _NgInfo.Count; i++)
            {
                float dx = _NgInfo[i].XPercent - percentX;
                float dy = _NgInfo[i].YPercent - percentY;
                float distanceSquared = dx * dx + dy * dy; // 距離の二乗を計算

                if (distanceSquared < minDistanceSquared)
                {
                    minDistanceSquared = distanceSquared;
                    closestIndex = i;
                }
            }
            // 最も近いインデックスが見つかった場合、その点を削除
            if (closestIndex != -1)
            {
                ngcnt=_NgInfo[closestIndex];

                _NgInfo.RemoveAt(closestIndex);
                refresh();
                return closestIndex; // 削除した点のNgTypeを返す
            }
            else 
            {                 // 近い点が見つからなかった場合は何もしない
                { ngcnt = null; return -1; }
            }
        }
        #endregion
        public void delLastPoint()
        {
            if (_NgInfo.Count > 0)
            {
                _NgInfo.RemoveAt(_NgInfo.Count - 1);
                refresh();
            }           
        }
        private Rectangle GetImageDisplayRectangle()
        {
            if (pBox == null) return Rectangle.Empty;
            if (_OriginalImage == null) return Rectangle.Empty;
            float imageRatio = (float)pBox.Image.Width / pBox.Image.Height;
            float controlRatio = (float)pBox.Width / pBox.Height;

            int width, height;
            if (imageRatio > controlRatio)
            {
                width = pBox.Width;
                height = (int)(pBox.Width / imageRatio);
            }
            else
            {
                height = pBox.Height;
                width = (int)(pBox.Height * imageRatio);
            }
            int x = (pBox.Width - width) / 2;
            int y = (pBox.Height - height) / 2;
            return new Rectangle(x, y, width, height);
        }
        private PointF TransformPointBack(PointF point, RotateAngle angle)
        {
            float x = point.X - 0.5f; //正規化座標を中心基準に変更
            float y = point.Y - 0.5f;

            double angleRad = -(double)angle * Math.PI / 180.0; //ラジアン変換
            float originalX = (float)(x * Math.Cos(angleRad) - y * Math.Sin(angleRad)) + 0.5f;
            float originalY = (float)(x * Math.Sin(angleRad) + y * Math.Cos(angleRad)) + 0.5f;
            return new PointF(originalX, originalY);
        }
        public Bitmap RotatedImage()
        {
            if (_OriginalImage == null) return null;
            float angle = (float)_RoteAngle;
            int width = _OriginalImage.Width;
            int height = _OriginalImage.Height;
            Bitmap rotatedImage = new(width, height);

            using (var graphics = Graphics.FromImage(rotatedImage))
            {
                //画像を回転
                graphics.TranslateTransform(width / 2.0f, height / 2.0f);
                graphics.RotateTransform(angle);
                graphics.TranslateTransform(-width / 2.0f, -height / 2.0f);
                graphics.DrawImage(_OriginalImage, new Point(0, 0));
                //線描画
                foreach (var line in _Lines)
                {
                    float x0 = line.perX0 / 100 * width;
                    float y0 = line.perY0 / 100 * height;
                    float x1 = line.perX1 / 100 * width;
                    float y1 = line.perY1 / 100 * height;
                    graphics.DrawLine(line.LinePen, x0, y0, x1, y1);
                }
                //塗潰円描画
                foreach (var item in _NgInfo)
                {
                    Color pointColor = _PointColors[item.NgType];
                    float x = item.XPercent/ 100f * width;
                    float y = item.YPercent / 100f * height;
                    int radius = 8;
                    using Brush brush = new SolidBrush(pointColor);
                    graphics.FillEllipse(brush, x - radius, y - radius, radius * 2, radius * 2);
                }




                //----
                int maxDiameter = Math.Min(width, height) / splitCount;

                Brush bubbleBrush = new SolidBrush(Color.FromArgb(128, Color.Blue)); // 半透明の青色
                Pen borderPen = new Pen(Color.Black, 2);                             // 境界線

                // バブルデータを描画
                foreach (var bubble in _BubbleDatas)
                {
                    // 画像サイズを基準に座標を計算（loc: 0～100 の割合）
                    float x = width * (bubble.Loc.X / 100f);
                    float y = height * (bubble.Loc.Y / 100f);

                    // バブルのサイズを計算 (値に基づき最大直径をスケール)
                    float diameter = maxDiameter * (bubble.Value / 100f);

                    // バブルの描画位置 (中心を指定座標にセット)
                    float drawX = x - diameter / 2;
                    float drawY = y - diameter / 2;

                    // 塗りつぶしの円を描画
                    graphics.FillEllipse(bubbleBrush, drawX, drawY, diameter, diameter);

                    // 境界線を描画
                    graphics.DrawEllipse(borderPen, drawX, drawY, diameter, diameter);
                }



























                //------



                //foreach (var (point, index) in _PointsWithIndex)
                //{
                //    Color pointColor = _PointColors[index];
                //    float x = point.X / 100f * width;
                //    float y = point.Y / 100f * height;
                //    int radius = 8;
                //    using Brush brush = new SolidBrush(pointColor);
                //    graphics.FillEllipse(brush, x - radius, y - radius, radius * 2, radius * 2);
                //}
            }
            return rotatedImage;
        }
        public void refresh()
        {
            if (pBox == null) return;
            if (_OriginalImage == null) return;
            Bitmap rotatedImage = RotatedImage();
            pBox.Image?.Dispose();
            pBox.Image = rotatedImage;
        }
    }
    public class LineOnPicture
    {
        public float X0 { get; set; }
        public float Y0 { get; set; }
        public float X1 { get; set; }
        public float Y1 { get; set; }
        public Color LineColor { get; set; } = Color.Yellow;
        public float LineWidth { get; set; } = 2.0F;
        public string Tag { get; set; }
        public float perX0 { get { return X0 * 100; } set { X0 = value / 100; } }
        public float perY0 { get { return Y0 * 100; } set { Y0 = value / 100; } }
        public float perX1 { get { return X1 * 100; } set { X1 = value / 100; } }
        public float perY1 { get { return Y1 * 100; } set { Y1 = value / 100; } }
        public Pen LinePen => new(LineColor, LineWidth);
    }
    public class BubbleData
    {
        public PointF Loc { get; set; }  // 座標 (0～100%)
        public float Value { get; set; } // 値 (バブルサイズに影響)

        public BubbleData(PointF loc, float value)
        {
            Loc = loc;
            Value = value;
        }
    }
}