using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace decksterity
{
    /// <summary>
    /// Manages global keyboard shortcuts for Decksterity PowerPoint add-in
    /// </summary>
    public class KeyboardHookManager
    {
        // Windows API imports
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        // Hook constants
        private const int WH_KEYBOARD = 2;
        private const int HC_ACTION = 0;

        // Delegate and hook variables
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private readonly LowLevelKeyboardProc keyboardProc = HookCallback;
        private static IntPtr keyboardHook = IntPtr.Zero;
        private static KeyboardHookManager instance;

        /// <summary>
        /// Installs the keyboard hook
        /// </summary>
        public void InstallHook()
        {
            try
            {
                instance = this;
                keyboardHook = SetWindowsHookEx(WH_KEYBOARD, keyboardProc, 
                    IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());
                
                if (keyboardHook == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine("Failed to install keyboard hook");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error installing keyboard hook: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes the keyboard hook
        /// </summary>
        public void RemoveHook()
        {
            try
            {
                if (keyboardHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(keyboardHook);
                    keyboardHook = IntPtr.Zero;
                }
                instance = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error removing keyboard hook: {ex.Message}");
            }
        }

        /// <summary>
        /// Main hook callback that processes keyboard messages
        /// </summary>
        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            int PreviousStateBit = 31;
            bool KeyWasAlreadyPressed = false;
            Int64 bitmask = (Int64)Math.Pow(2, (PreviousStateBit - 1));

            if (nCode < 0)
            {
                return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
            }

            if (nCode == HC_ACTION)
            {
                Keys keyData = (Keys)wParam;
                KeyWasAlreadyPressed = ((Int64)lParam & bitmask) > 0;

                // Only process key press (not key release)
                if (!KeyWasAlreadyPressed && IsKeyDown(Keys.ControlKey) && IsKeyDown(Keys.ShiftKey))
                {
                    HandleShortcut(keyData);
                }
            }

            return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
        }

        /// <summary>
        /// Handles the actual shortcut execution
        /// </summary>
        private static void HandleShortcut(Keys keyData)
        {
            try
            {
                switch (keyData)
                {
                    case Keys.D1: // Ctrl+Shift+1 - Align Left
                        AlignmentHelper.AlignLeft();
                        break;
                    case Keys.D2: // Ctrl+Shift+2 - Align Center
                        AlignmentHelper.AlignCenter();
                        break;
                    case Keys.D3: // Ctrl+Shift+3 - Align Right
                        AlignmentHelper.AlignRight();
                        break;
                    case Keys.D4: // Ctrl+Shift+4 - Align Top
                        AlignmentHelper.AlignTop();
                        break;
                    case Keys.D5: // Ctrl+Shift+5 - Align Middle
                        AlignmentHelper.AlignMiddle();
                        break;
                    case Keys.D6: // Ctrl+Shift+6 - Align Bottom
                        AlignmentHelper.AlignBottom();
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error handling keyboard shortcut: {ex.Message}");
            }
        }

        /// <summary>
        /// Helper function to check if a key is currently pressed
        /// </summary>
        private static bool IsKeyDown(Keys keys)
        {
            return (GetKeyState((int)keys) & 0x8000) == 0x8000;
        }
    }
}