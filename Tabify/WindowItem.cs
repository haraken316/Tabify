using System;

namespace Tabify
{
    public class WindowItem
    {
        public IntPtr HWnd { get; set; }
        public string Title { get; set; }

        public override string ToString() => Title;
    }
}

