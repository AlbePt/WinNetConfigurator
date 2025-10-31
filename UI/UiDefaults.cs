using System.Drawing;
using System.Windows.Forms;

namespace WinNetConfigurator.UI
{
    // Очень простая фабрика UI-элементов
    public static class UiDefaults
    {
        // Единый ToolTip для формы можно создавать снаружи, а можно и здесь
        public static Button CreateTopButton(string text, string tooltipText, ToolTip tt)
        {
            var btn = new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 0, 8, 0),
                Height = 28
            };

            if (tt != null && !string.IsNullOrEmpty(tooltipText))
                tt.SetToolTip(btn, tooltipText);

            return btn;
        }

        public static StatusStrip CreateStatusStrip()
        {
            var strip = new StatusStrip
            {
                SizingGrip = false
            };
            var label = new ToolStripStatusLabel("Готово");
            strip.Items.Add(label);
            return strip;
        }
    }
}
