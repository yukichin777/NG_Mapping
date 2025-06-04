using DocumentFormat.OpenXml.ExtendedProperties;
using KeyHook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NGMapping
{
    public partial class f_Login : Form
    {
        private GlobalKeyboardHook hook = new();

        private MainForm mainfrm;
        public f_Login()
        {
            InitializeComponent();
            hook.SubmitKeyMode = SubmitKey.CR;
        }

        private void f_Login_Load(object sender, EventArgs e)
        {
            CSet.SetFormLocState(this);

            //hook.CharReceived += _hook_CharReceived;
            hook.InputSubmitted += hook_InputSubmitted;
            hook.Start();
        }

        private void hook_InputSubmitted(object sender, KeyInputEventArgs e)
        {
            
            string st=e.InputText.Trim();
            string Operator= "";

            if (st.Length > 5 && string.Equals(st.Substring(0, 5), "Name-", StringComparison.OrdinalIgnoreCase))  
            {   
                hook.Stop();
                Operator = st.Substring(5);
            }
            if (st.ToLower() is "admin" or "ad" or "paper" or "pa" or "qc")
            {
                hook.Stop();                
                Operator = "##";
            }

            if (Operator != "") 
            {
                mainfrm = new MainForm(Operator,this);
                mainfrm.Show();
                this.ShowInTaskbar = false; // タスクバーに表示しない
                this.Hide();
                return; // ここで処理を終了
            }
            //LabelText = "";
        }

        private void _hook_CharReceived(object sender, char e)
        {
            string dummy = e.ToString();
            //LabelText = (LabelText + e.ToString()).Trim();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Dispose();
        }

        private void F_Login_Shown(object sender, EventArgs e)
        {
            label1.Focus();
        }
    }
}
