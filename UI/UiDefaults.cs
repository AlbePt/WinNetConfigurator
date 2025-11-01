using System.Drawing;
using System.Windows.Forms;
using WinNetConfigurator.UI;

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
                BackColor = AppTheme.SecondaryBackground,
                ForeColor = AppTheme.TextPrimary,
                Padding = AppTheme.PaddingNormal,
                Margin = AppTheme.PaddingNormal,
                Height = 28
            };

            if (tt != null && !string.IsNullOrEmpty(tooltipText))
                tt.SetToolTip(btn, tooltipText);

            return btn;
        }

        public static void StyleTopCommandButton(Button button)
        {
            button.Height = 32;
            button.MinimumSize = new Size(120, 32);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = AppTheme.SecondaryBackground;
            button.ForeColor = AppTheme.TextPrimary;
            button.Margin = new Padding(0, 0, 8, 0);
        }

        public static StatusStrip CreateStatusStrip()
        {
            var strip = new StatusStrip
            {
                SizingGrip = false,
                BackColor = AppTheme.SecondaryBackground
            };
            var label = new ToolStripStatusLabel("Готово")
            {
                ForeColor = AppTheme.TextPrimary
            };
            strip.Items.Add(label);
            return strip;
        }

        public static void ApplyFormBaseStyle(Form form)
        {
            if (form == null)
            {
                return;
            }

            form.BackColor = AppTheme.MainBackground;
            form.Padding = AppTheme.PaddingLarge;
        }
    }
}
