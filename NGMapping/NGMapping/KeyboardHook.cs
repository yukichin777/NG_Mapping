using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KeyHook
{
    internal class KeyboardHook
    {
        private const int WH_KEYBOARD_LL = 0x0D;
        private const int WM_KEYBOARD_DOWN = 0x100;
        private const int WM_KEYBOARD_UP = 0x101;
        private const int WM_SYSKEY_DOWN = 0x104;
        private const int WM_SYSKEY_UP = 0x105;

        //イベントハンドラの定義
        public event EventHandler<KeyboardHookEventArgs> OnKeyDown = delegate { };
        public event EventHandler<KeyboardHookEventArgs> OnKeyUp = delegate { };

        //コールバック関数のdelegate 定義
        private delegate IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam);

        //キーボードフックに必要なDLLのインポート
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookCallback lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        //フックプロシージャのハンドル
        private static IntPtr _hookHandle = IntPtr.Zero;

        //フック時のコールバック関数
        private static HookCallback _callback;

        /// <summary>
        /// キーボードHook の開始
        /// </summary>
        /// <param name="callback"></param>
        public void Hook()
        {
            _callback = CallbackProc;
            using (Process process = Process.GetCurrentProcess())
            {
                using (ProcessModule module = process.MainModule)
                {
                    _hookHandle = SetWindowsHookEx(
                       WH_KEYBOARD_LL,                                          //フックするイベントの種類
                       _callback, //フック時のコールバック関数
                       GetModuleHandle(module.ModuleName),                      //インスタンスハンドル
                       0                                                        //スレッドID（0：全てのスレッドでフック）
                   );
                }
            }
        }
        /// <summary>
        /// コールバック関数
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        //private IntPtr CallbackProc(int nCode, IntPtr wParam, IntPtr lParam)
        //{
        //    var args = new KeyboardHookEventArgs();
        //    Keys key = (Keys)(short)Marshal.ReadInt32(lParam);
        //    args.Key = key;

        //    if ((int)wParam == WM_KEYBOARD_DOWN || (int)wParam == WM_SYSKEY_DOWN) OnKeyDown(this, args);
        //    if ((int)wParam == WM_KEYBOARD_UP || (int)wParam == WM_SYSKEY_UP) OnKeyUp(this, args);

        //    return (args.RetCode == 0) ? CallNextHookEx(_hookHandle, nCode, wParam, lParam) : (IntPtr)1;
        //}
        private IntPtr CallbackProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }

            var args = new KeyboardHookEventArgs();
            Keys key = (Keys)(short)Marshal.ReadInt32(lParam);
            args.Key = key;

            // キーイベントの種類を判定
            int eventType = (int)wParam;
            if (eventType == WM_KEYBOARD_DOWN || eventType == WM_SYSKEY_DOWN)
            {
                // WM_CHAR から文字を取得
                int virtualKeyCode = Marshal.ReadInt32(lParam);
                char character = ConvertVirtualKeyToChar(virtualKeyCode);
                args.Character = character;

                OnKeyDown(this, args);
            }

            if (eventType == WM_KEYBOARD_UP || eventType == WM_SYSKEY_UP)
            {
                OnKeyUp(this, args);
            }

            return (args.RetCode == 0)
                ? CallNextHookEx(_hookHandle, nCode, wParam, lParam)
                : (IntPtr)1;
        }
        private char ConvertVirtualKeyToChar(int virtualKeyCode)
        {
            byte[] kbdState = new byte[256];
            StringBuilder buffer = new StringBuilder(2);

            // 現在のキーボード状態を取得
            if (!GetKeyboardState(kbdState))
            {
                return '\0'; // 失敗した場合は空文字を返す
            }

            // 仮想キーを文字に変換
            ToUnicode((uint)virtualKeyCode, 0, kbdState, buffer, buffer.Capacity, 0);

            return buffer.Length > 0 ? buffer[0] : '\0';
        }

        [DllImport("user32.dll")]
        private static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        private static extern int ToUnicode(
            uint wVirtKey,
            uint wScanCode,
            byte[] lpKeyState,
            [Out] StringBuilder pwszBuff,
            int cchBuff,
            uint wFlags
        );
        /// <summary>
        /// キーボードHockの終了
        /// </summary>
        public void UnHook()
        {
            UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
        }
    }
    /// <summary>
    /// キーボードフックのイベント引数
    /// </summary>
    public class KeyboardHookEventArgs
    {
        public Keys Key { get; set; }
        public char Character { get; set; } // 押されたキーに対応する文字
        public int RetCode { get; set; } = 0;
    }
}
