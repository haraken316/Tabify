using System;

namespace Tabify
{
    public class WindowItem
    {
        public IntPtr HWnd { get; set; }
        public string Title { get; set; } = string.Empty;

        public override string ToString() => Title;
    }
}

