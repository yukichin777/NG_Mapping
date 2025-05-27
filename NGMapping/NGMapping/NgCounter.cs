using System.Drawing;

namespace NGMapping
{
    public class NgCounter
    {
        public float XPercent { get; private set; } // X座標の画像に対する％値
        public float YPercent { get; private set; } // Y座標の画像に対する％値
        public PointF XYPercent
        {
            get { return new PointF(XPercent, YPercent); }
        }
        public int NgType { get; private set; }     // NG種類
        public string NgText { get; set; }         // NGテキスト

        public NgCounter(float xPercent, float yPercent, int ngType,string ngText)
        {
            XPercent = xPercent;
            YPercent = yPercent;
            NgType = ngType;
            NgText = ngText;
        }
    }
}
