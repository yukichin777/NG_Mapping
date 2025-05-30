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


        private GlobalKeyboardHook _hook = new();

        private MainForm mainfrm;
        public f_Login()
        {
            InitializeComponent();
            _hook.SubmitKeyMode = SubmitKey.CR;

            //_hook.CharReceived += (s, c) => Console.WriteLine($"Char: {c}");
            //_hook.InputSubmitted += (s, e) => Console.WriteLine($"Submitted: [{e.InputText}]");

        }

        private void f_Login_Load(object sender, EventArgs e)
        {

            //_hook.CharReceived += _hook_CharReceived;
            _hook.InputSubmitted += _hook_InputSubmitted;
            _hook.Start();
        }

        private void _hook_InputSubmitted(object sender, KeyInputEventArgs e)
        {
            
            string st=e.InputText.Trim();
            string Operator= "";

            if (st.Length>5 && st.Substring(0, 5) == "Name-") 
            {   
                _hook.Stop();
                Operator = st.Substring(5);
            }
            if(st=="admin" || st == "ad"||st=="paper" || st=="pa" || st == "qc")
            {
                _hook.Stop();                
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
