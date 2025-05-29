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


        private GlobalKeyboardHook _hook;

        bool Flg_ready = false;
        string buf = "";
        int Flg_KeyDispose = 1;//1を指定すると、ﾌｯｸしたｷｰを破棄する。


        public f_Login()
        {
            InitializeComponent();
            _hook = new GlobalKeyboardHook
            {
                SubmitKeyMode = SubmitKey.CRLF
            };

            _hook.CharReceived += (s, c) => Console.WriteLine($"Char: {c}");
            _hook.InputSubmitted += (s, e) => Console.WriteLine($"Submitted: [{e.InputText}]");

        }

        private void f_Login_Load(object sender, EventArgs e)
        {

            _hook.CharReceived += _hook_CharReceived;
            _hook.InputSubmitted += _hook_InputSubmitted;

            //_hook.OnKeyUp += _hook_OnKeyUp;

            _hook.Hook();
        }

        private void _hook_InputSubmitted(object sender, KeyInputEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void _hook_CharReceived(object sender, char e)
        {
            throw new NotImplementedException();
        }

        private void _hook_OnKeyUp(object sender, KeyboardHookEventArgs e)
        {
            //throw new NotImplementedException();
        }
        #region event----hool_OnKeyDown
        private void _hook_OnKeyDown(object sender, KeyboardHookEventArgs e)
        {
            int maxL = 21;
            bool FlgEnter = false;
            bool FlgSpace = false;
            string st = "";

            string st1 = e.Character.ToString();

            switch (e.Key)
            {
                #region 数字,ｱﾙﾌｧﾍﾞｯﾄ
                case Keys.D0: st = "0"; break;
                case Keys.D1: st = "1"; break;
                case Keys.D2: st = "2"; break;
                case Keys.D3: st = "3"; break;
                case Keys.D4: st = "4"; break;
                case Keys.D5: st = "5"; break;
                case Keys.D6: st = "6"; break;
                case Keys.D7: st = "7"; break;
                case Keys.D8: st = "8"; break;
                case Keys.D9: st = "9"; break;
                case Keys.A: st = "A"; break;
                case Keys.B: st = "B"; break;
                case Keys.C: st = "C"; break;
                case Keys.D: st = "D"; break;
                case Keys.E: st = "E"; break;
                case Keys.F: st = "F"; break;
                case Keys.G: st = "G"; break;
                case Keys.H: st = "H"; break;
                case Keys.I: st = "I"; break;
                case Keys.J: st = "J"; break;
                case Keys.K: st = "K"; break;
                case Keys.L: st = "L"; break;
                case Keys.M: st = "M"; break;
                case Keys.N: st = "N"; break;
                case Keys.O: st = "O"; break;
                case Keys.P: st = "P"; break;
                case Keys.Q: st = "Q"; break;
                case Keys.R: st = "R"; break;
                case Keys.S: st = "S"; break;
                case Keys.T: st = "T"; break;
                case Keys.U: st = "U"; break;
                case Keys.V: st = "V"; break;
                case Keys.W: st = "W"; break;
                case Keys.X: st = "X"; break;
                case Keys.Y: st = "Y"; break;
                case Keys.Z: st = "Z"; break;
                
                #endregion
                #region ﾃﾝｷｰ
                case Keys.NumPad0: st = "0"; break;
                case Keys.NumPad1: st = "1"; break;
                case Keys.NumPad2: st = "2"; break;
                case Keys.NumPad3: st = "3"; break;
                case Keys.NumPad4: st = "4"; break;
                case Keys.NumPad5: st = "5"; break;
                case Keys.NumPad6: st = "6"; break;
                case Keys.NumPad7: st = "7"; break;
                case Keys.NumPad8: st = "8"; break;
                case Keys.NumPad9: st = "9"; break;
                case Keys.Divide: st = "/"; break;
                case Keys.Multiply: st = "*"; break;
                case Keys.Subtract: st = "-"; break;
                case Keys.Add: st = "+"; break;
                case Keys.Decimal: st = "."; break;
                #endregion
                #region その他いろいろ
                case Keys.Oem1: st = ":"; break;//Keys.OemSemicolon同じ                    
                case Keys.Oemplus: st = ";"; break;//「+」とならない事に注意                    
                case Keys.Oem3: st = "@"; break;//Keys.Oemtilde:
                case Keys.Oem4: st = "["; break;//Keys.OemOpenBrackets
                case Keys.Oem5: st = "\\"; break;//Keys.OemPipe
                case Keys.Oem6: st = "]"; break;//Keys.OemCloseBrackets                    
                case Keys.Oem7: st = "^"; break;//Keys.OemQuotes                    
                case Keys.Oemcomma: st = ","; break;
                case Keys.OemMinus: st = "-"; break;
                case Keys.OemPeriod: st = "."; break;
                case Keys.Oem2: st = "/"; break;//Keys.OemQuestion:
                case Keys.Oem102: st = "\\"; break;//Keys.OemBackslash(ﾊﾞｯｸｽﾗｯｼｭ)
                #endregion
                #region Enter,Space
                case Keys.Enter://Keys.Return:も同じ
                    FlgEnter = true;
                    st = "\r\n";
                    break;
                case Keys.Space:
                    st = " ";
                    FlgSpace = true;
                    break;
                #endregion
                #region F12
                case Keys.F12: Flg_KeyDispose = 1 - Flg_KeyDispose; break;//ﾌｯｸｼﾀｰｷｰｦ破棄するかのﾄｸﾞﾙ
                #endregion
                #region これらは無視する
                case Keys.KeyCode:
                case Keys.Modifiers:
                case Keys.None:
                case Keys.LButton:
                case Keys.RButton:
                case Keys.Cancel:
                case Keys.MButton:
                case Keys.XButton1:
                case Keys.XButton2:
                case Keys.Back:
                case Keys.Tab:
                case Keys.LineFeed:
                case Keys.Clear:
                case Keys.ShiftKey:
                case Keys.ControlKey:
                case Keys.Menu://menuｷｰ(ｺﾝﾃｷｽﾄﾒﾆｭｰを開くｷｰ)
                case Keys.Pause://PauseBreakｷｰ
                case Keys.CapsLock://Keys.Capitalも同じ
                case Keys.KanaMode://Keys.HanguelMode,Keys.HangulModeも同じ
                case Keys.JunjaMode:
                case Keys.FinalMode:
                case Keys.KanjiMode://Keys.HanjaModeも同じ
                case Keys.Escape:
                case Keys.IMEConvert://変換ｷｰ
                case Keys.IMENonconvert://無変換ｷｰ
                case Keys.IMEAccept://Keys.IMEAceeptも同じ
                case Keys.IMEModeChange:
                case Keys.PageUp://Keys.Priorも同じ
                case Keys.PageDown://Keys.Nextも同じ
                case Keys.End:
                case Keys.Home:
                case Keys.Left:
                case Keys.Up:
                case Keys.Right:
                case Keys.Down:
                case Keys.Select:
                case Keys.Print:
                case Keys.Execute:
                case Keys.PrintScreen://Keys.Snapshot同じ
                case Keys.Insert:
                case Keys.Delete:
                case Keys.Help:
                case Keys.LWin:
                case Keys.RWin:
                case Keys.Apps:
                case Keys.Sleep:
                case Keys.Separator:
                case Keys.F1:
                case Keys.F2:
                case Keys.F3:
                case Keys.F4:
                case Keys.F5:
                case Keys.F6:
                case Keys.F7:
                case Keys.F8:
                case Keys.F9:
                case Keys.F10:
                case Keys.F11:

                case Keys.F13:
                case Keys.F14:
                case Keys.F15:
                case Keys.F16:
                case Keys.F17:
                case Keys.F18:
                case Keys.F19:
                case Keys.F20:
                case Keys.F21:
                case Keys.F22:
                case Keys.F23:
                case Keys.F24:
                case Keys.NumLock:
                case Keys.Scroll:
                case Keys.LShiftKey:
                case Keys.RShiftKey:
                case Keys.LControlKey:
                case Keys.RControlKey:
                case Keys.LMenu:
                case Keys.RMenu:
                case Keys.BrowserBack:
                case Keys.BrowserForward:
                case Keys.BrowserRefresh:
                case Keys.BrowserStop:
                case Keys.BrowserSearch:
                case Keys.BrowserFavorites:
                case Keys.BrowserHome:
                case Keys.VolumeMute:
                case Keys.VolumeDown:
                case Keys.VolumeUp:
                case Keys.MediaNextTrack:
                case Keys.MediaPreviousTrack:
                case Keys.MediaStop:
                case Keys.MediaPlayPause:
                case Keys.LaunchMail:
                case Keys.SelectMedia:
                case Keys.LaunchApplication1:
                case Keys.LaunchApplication2:
                case Keys.Oem8:
                case Keys.ProcessKey:
                case Keys.Packet:
                case Keys.Attn:
                case Keys.Crsel:
                case Keys.Exsel:
                case Keys.EraseEof:
                case Keys.Play:
                case Keys.Zoom:
                case Keys.NoName:
                case Keys.Pa1:
                case Keys.OemClear:
                case Keys.Shift:
                case Keys.Control:
                case Keys.Alt:
                default:
                    break;
                    #endregion
            }

            if (FlgEnter || FlgSpace)
            {
                if(CodeCheck(buf,out string name))
                {
                }
                buf = "";
            }
            else
            {
                Flg_ready = false;
                buf += st1;
            }
            label5.Text = buf;

            e.RetCode = 1;//他のプログラムにキーを転送したくない場合は1を代入;
        }
        #endregion
        #region ﾒｿｯﾄﾞ-----正しいQRｺｰﾄﾞか確認する
        private bool CodeCheck(string txt,out string name)
        {
            txt=txt.Trim();
            if(txt.Length>5 && txt.Substring(0,5)=="Name-")
            {
                name = txt.Substring(5);
                return true;
            }
            else {
                name = "";
                return false;
            }
            //for (int i = 0; i < txt.Length; i++) { L_Result[i].Text = txt.Substring(i, 1); }

            //qr = txt;
            //key = txt.Substring(0, 8);
            //dtSt = txt.Substring(9, 6);
            //sn = txt.Substring(15, 4);


            //if (txt.Substring(8, 1) != "-") { L_Result[8].BackColor = Color.Red; flg = true; }
            //if (txt.Substring(19, 2) != "PP") { L_Result[19].BackColor = Color.Red; L_Result[20].BackColor = Color.Red; flg = true; }

            //dt = stToDate(dtSt);

            //if (dt.Year == 1900)
            //{
            //    L_Result[9].BackColor = Color.Red;
            //    L_Result[10].BackColor = Color.Red;
            //    L_Result[11].BackColor = Color.Red;
            //    L_Result[12].BackColor = Color.Red;
            //    L_Result[13].BackColor = Color.Red;
            //    L_Result[14].BackColor = Color.Red;
            //    flg = true;
            //}

            //if (flg)
            //{
            //    L_Target.Text = "-";
            //    L_HexFile.Text = "-";
            //    L_Mes.Text = "QRコードが間違っています。\r\n再度QRコードを読み込んでください。";
            //    return false;
            //}

            //flg = false;
            //for (int i = 0; i < RPJList.Count; i++)
            //{
            //    if (key == RPJList[i].CodeKey)
            //    {
            //        L_Target.Text = RPJList[i].Name;
            //        L_HexFile.Text = RPJList[i].RPJFileName;
            //        L_Status.Text = "QRコード読込完了";
            //        L_Mes.Text = "ワークを治具にセットし、\r\n外付けSw/Enter/Space/書込ボタンのいずれかを押してください。\r\n対象ワークを変更するときは、再度QRコードを読み込んでください。";
            //        rpj = RPJList[i].RPJFileName;
            //        return true;
            //    }
            //}
            //if (!flg)
            //{
            //    L_Result[0].BackColor = Color.Red;
            //    L_Result[1].BackColor = Color.Red;
            //    L_Result[2].BackColor = Color.Red;
            //    L_Result[3].BackColor = Color.Red;
            //    L_Result[4].BackColor = Color.Red;
            //    L_Result[5].BackColor = Color.Red;
            //    L_Result[6].BackColor = Color.Red;
            //    L_Result[7].BackColor = Color.Red;
            //    L_Target.Text = "Not Found";
            //    L_HexFile.Text = "-";
            //    L_Mes.Text = "QRコードが間違っています。\r\n再度QRコードを読み込んでください。";
            //    return false;
            //}
            //else
            //{
            //    return true;
            //}

        }
        #endregion

    }
}
