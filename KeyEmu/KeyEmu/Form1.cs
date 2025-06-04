using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KeyEmu
{
    public partial class Form1 : Form
    {
        //[StructLayout(LayoutKind.Sequential)]
        //struct INPUT
        //{
        //    public uint type;
        //    public InputUnion u;
        //}

        //[StructLayout(LayoutKind.Sequential)]
        //struct InputUnion
        //{
        //    public KEYBDINPUT ki;
        //}

        //[StructLayout(LayoutKind.Sequential)]
        //struct KEYBDINPUT
        //{
        //    public ushort wVk;
        //    public ushort wScan;
        //    public uint dwFlags;
        //    public uint time;
        //    public IntPtr dwExtraInfo;
        //}

        //private const uint INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_KEYUP = 0x0002;

        //[DllImport("user32.dll")]
        //private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);


        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);

        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string textToSend =textBox1.Text;

            textToSend = comboBox1.Text+"\r";

            if (!string.IsNullOrWhiteSpace(textToSend))
            {
                EmulateKeyboardInput(textToSend); // 入力内容をエミュレート
            }
            else
            {
                MessageBox.Show("Please enter some text to send.");
            }
        }

        static void EmulateKeyboardInput(string text)
        {
            foreach (char c in text)
            {
                byte vk = GetVirtualKey(c);

                // キー押下
                keybd_event(vk, 0, 0, IntPtr.Zero);

                // キー解放
                keybd_event(vk, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
            }
        }
        static byte GetVirtualKey(char c)
        {
            if (char.IsLetterOrDigit(c))
            {
                return (byte)c;
            }
            else
            {
                switch (c)
                {
                    case '\r': return 0x0D; // Enterキー
                    case '\n': return 0x0D; // 改行キー
                    case '\t': return 0x09; // Tabキー
                    default: return (byte)c;
                }
            }
        }
        //static void EmulateKeyboardInput_Old(string text)
        //{
        //    foreach (char c in text)
        //    {
        //        INPUT[] inputs = new INPUT[2];

        //        // キー押下
        //        inputs[0].type = INPUT_KEYBOARD;
        //        inputs[0].u.ki.wVk = (ushort)c;
        //        inputs[0].u.ki.dwFlags = 0;

        //        // キー解放
        //        inputs[1].type = INPUT_KEYBOARD;
        //        inputs[1].u.ki.wVk = (ushort)c;
        //        inputs[1].u.ki.dwFlags = KEYEVENTF_KEYUP;

        //        SendInput(2, inputs, Marshal.SizeOf(typeof(INPUT)));
        //    }
        //}
    }
}
