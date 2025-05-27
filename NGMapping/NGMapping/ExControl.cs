using System.Drawing;
using System.Windows.Forms;

namespace NGMapping
{
    internal class ExControl
    {
    }
    public class ExNumUpDown : NumericUpDown
    {
        int TagInt { get; set; } = 0;
        string TagStr { get; set; } = "";

        public float F_Value
        {
            get { return (float)this.Value; }
            set { this.Value = (decimal)value; }
        }
        public double Dbl_Value
        {
            get { return (double)this.Value; }
            set { this.Value = (decimal)value; }
        }
        public int Int_Value
        {
            get { return (int)this.Value; }
            set { this.Value = (decimal)value; }
        }
    }

    public class ExLabel : Label
    {
        private int _value = 0;
        int TagInt { get; set; } = 0;
        string TagStr { get; set; } = "";

        public ExLabel()
        {
            this.Text = "0";
        }

        public void Increment()
        {
            Value++;
        }
        public int Value
        {
            get { return _value; }
            set { _value = value; this.Text = value.ToString(); }
        }
    }

    public class ExColorButton : Button
    {
        private int tagInt = 0;
        private string tagStr = "";
        private Color? buttonColor = null;

        public int TagInt
        {
            get { return tagInt; }
            set { tagInt = value; }
        }

        public string TagStr
        {
            get { return tagStr; }
            set { tagStr = value; }
        }

        public Color? ButtonColor
        {
            get
            {
                if (buttonColor.HasValue && IsSystemColor(buttonColor.Value))
                {
                    return null;
                }
                return buttonColor;
            }
            set
            {
                // SystemColor の場合、BackColorをデフォルトに設定しプロパティに保存しない。
                if (value.HasValue && IsSystemColor(value.Value))
                {
                    base.BackColor = SystemColors.Control; // デフォルトの色を設定
                    buttonColor = null;
                }
                else
                {
                    buttonColor = value;
                    base.BackColor = value ?? SystemColors.Control; // ボタンのBackColorを設定
                }
            }
        }

        /// <summary>
        /// 指定された色が SystemColors に該当するかを判定します。
        /// </summary>
        /// <param name="color">判定する色</param>
        /// <returns>SystemColors に該当する場合は true、そうでない場合は false</returns>
        private bool IsSystemColor(Color color)
        {
            return color.ToKnownColor() >= KnownColor.AliceBlue && color.ToKnownColor() <= KnownColor.WindowText;
        }
    }


    public class ExToolStripMenu: ToolStripMenuItem
    {
        private int tagInt = 0;
        private string tagStr = "";
        public int TagInt
        {
            get { return tagInt; }
            set { tagInt = value; }
        }
        public string TagStr
        {
            get { return tagStr; }
            set { tagStr = value; }
        }
    }
}
