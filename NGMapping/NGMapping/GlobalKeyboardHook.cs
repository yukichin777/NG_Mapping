using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;


namespace NGMapping
{
    public enum SubmitKey
    {
        CR,
        LF,
        CRLF,
        TAB
    }

    public class KeyInputEventArgs : EventArgs
    {
        public string InputText { get; }

        public KeyInputEventArgs(string inputText)
        {
            InputText = inputText;
        }
    }

    public class GlobalKeyboardHook
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;

        private StringBuilder _buffer = new();
        private LowLevelKeyboardProc _proc;
        private IntPtr _hookID = IntPtr.Zero;

        public SubmitKey SubmitKeyMode { get; set; } = SubmitKey.CR;

        public event EventHandler<KeyInputEventArgs>? InputSubmitted;
        public event EventHandler<char>? CharReceived;

        public GlobalKeyboardHook()
        {
            _proc = HookCallback;
            // 初期状態ではフックしない
        }
        public void Start()
        {
            if (_hookID == IntPtr.Zero)
            {
                _hookID = SetHook(_proc);
            }
        }

        public void Stop()
        {
            if (_hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
            }
        }
        public string GetCharsFromKey(Keys key)
        {
            byte[] keyboardState = new byte[256];
            if (!GetKeyboardState(keyboardState))
                return "";

            uint scanCode = MapVirtualKey((uint)key, 0);
            StringBuilder sb = new(10);
            int result = ToUnicode((uint)key, scanCode, keyboardState, sb, sb.Capacity, 0);

            return result > 0 ? sb.ToString() : "";
        }

        private string GetSubmitKeyString()
        {
            return SubmitKeyMode switch
            {
                SubmitKey.CR => "\r",
                SubmitKey.LF => "\n",
                SubmitKey.CRLF => "\r\n",
                SubmitKey.TAB => "\t",
                _ => "\r"
            };
        }

        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using var curProcess = Process.GetCurrentProcess();
            using var curModule = curProcess.MainModule!;
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys key = (Keys)vkCode;

                if (key == Keys.Back)
                {
                    if (_buffer.Length > 0)
                        _buffer.Length--;
                }
                else
                {
                    string ch = GetCharsFromKey(key);
                    if (!string.IsNullOrEmpty(ch))
                    {
                        _buffer.Append(ch);
                        CharReceived?.Invoke(this, ch[0]);

                        string submitPattern = GetSubmitKeyString();
                        if (_buffer.Length >= submitPattern.Length &&
                            _buffer.ToString(_buffer.Length - submitPattern.Length, submitPattern.Length) == submitPattern)
                        {
                            InputSubmitted?.Invoke(this, new KeyInputEventArgs(_buffer.ToString()));
                            _buffer.Clear();
                        }
                    }
                }
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        #region WinAPI

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook,
            LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk,
            int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        private static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpKeyState,
            [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff, int cchBuff, uint wFlags);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        #endregion
    }


}
