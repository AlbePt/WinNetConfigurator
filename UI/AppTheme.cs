using System.Drawing;
using System.Windows.Forms;

namespace WinNetConfigurator.UI
{
    public static class AppTheme
    {
        public static readonly Color PrimaryBackground = Color.FromArgb(247, 249, 252);
        public static readonly Color MainBackground = PrimaryBackground;
        public static readonly Color SecondaryBackground = Color.White;
        public static readonly Color Accent = Color.FromArgb(30, 99, 255);
        public static readonly Color AccentSoft = Color.FromArgb(226, 236, 255);
        public static readonly Color TextPrimary = Color.Black;
        public static readonly Color TextMuted = Color.FromArgb(90, 99, 110);
        public static readonly Color Border = Color.FromArgb(220, 225, 232);

        public static readonly Padding PaddingSmall = new Padding(4);
        public static readonly Padding PaddingNormal = new Padding(8);
        public static readonly Padding PaddingLarge = new Padding(12);
    }
}
