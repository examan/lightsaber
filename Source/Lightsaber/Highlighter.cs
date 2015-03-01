using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace System.Runtime.CompilerServices
{
    public class ExtensionAttribute : Attribute { }
}

namespace Lightsaber
{
    /* Reference: http://stackoverflow.com/a/1563724 */
    static class RichTextExtensions
    {
        public static void ClearSelectionBackColor(this RichTextBox richTextBox)
        {
            NativeMethods.CHARFORMAT2 charFormat = new NativeMethods.CHARFORMAT2();

            charFormat.cbSize = Marshal.SizeOf(charFormat);
            charFormat.dwMask = NativeMethods.CFM_BACKCOLOR;
            charFormat.dwEffects = NativeMethods.CFM_BACKCOLOR;

            charFormat.crBackColor = 0;

            NativeMethods.SendMessage(richTextBox.Handle, NativeMethods.EM_SETCHARFORMAT, NativeMethods.SCF_SELECTION, ref charFormat);
        }
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        internal static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, ref CHARFORMAT2 lParam);

        internal const UInt32 WM_USER = 0x0400;
        internal const UInt32 EM_SETCHARFORMAT = (WM_USER + 68);
        internal const UInt32 CFM_BACKCOLOR = 0x04000000;
        internal static IntPtr SCF_SELECTION = (IntPtr)0x0001;

        [StructLayout(LayoutKind.Sequential, Pack = 4, CharSet = CharSet.Auto)]
        internal struct CHARFORMAT2
        {
            public int cbSize;
            public uint dwMask;
            public uint dwEffects;
            public int yHeight;
            public int yOffset;
            public int crTextColor;
            public byte bCharSet;
            public byte bPitchAndFamily;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szFaceName;
            public short wWeight;
            public short sSpacing;
            public int crBackColor;
            public int lcid;
            public int dwReserved;
            public short sStyle;
            public short wKerning;
            public byte bUnderlineType;
            public byte bAnimation;
            public byte bRevAuthor;
            public byte bReserved1;
        }
    }

    public partial class Addin
    {
        private bool HighlightEnabled
        {
            get
            {
                try
                {
                    using (var activeWindow = this.Application.ActiveWindow)
                    {
                        return 0 < activeWindow.Selection.TextRange.Count && activeWindow.Selection.TextRange.Text.Trim() != String.Empty;
                    }
                }
                catch
                {
                    return false;
                }
            }
        }
        
        private void Highlight(Color color)
        {
            if (!HighlightEnabled)
            {
                return;
            }

            var data = Clipboard.GetDataObject();
            if (data != null && 0 < data.GetFormats().Length)
            {
                if (MessageBox.Show(Resources.Str.ClipboardWarningMessage, "Lightsaber", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK)
                {
                    return;
                }

                Clipboard.Clear();
            }

            using (var activeWindow = this.Application.ActiveWindow)
            using (var richTextBox = new RichTextBox())
            {
                activeWindow.Selection.TextRange.Copy();
                richTextBox.Paste();
                richTextBox.SelectAll();

                if (color == Color.Transparent)
                {
                    richTextBox.ClearSelectionBackColor();
                }
                else
                {
                    richTextBox.SelectionBackColor = color;
                }

                richTextBox.Copy();

                activeWindow.Selection.TextRange.Paste();
            }

            Clipboard.Clear();

            try
            {
                this.Application.StartNewUndoEntry();
            }
            catch { }
        }

        private void ClearHighlight()
        {
            Highlight(Color.Transparent);
        }
    }
}
