using System.Windows.Forms;

namespace SheetMultiplication
{
    internal static class SheetCopyCountForm
    {
        public static int Show()
        {
            int result = 0;

            System.Windows.Forms.Form form = new System.Windows.Forms.Form { Text = "Ile kopii?", Width = 300, Height = 150 };
            NumericUpDown numeric = new NumericUpDown { Minimum = 1, Maximum = 100, Value = 3, Dock = DockStyle.Top };
            Button ok = new Button { Text = "OK", Dock = DockStyle.Bottom };
            ok.Click += (s, e) => form.DialogResult = DialogResult.OK;

            form.Controls.Add(numeric);
            form.Controls.Add(ok);

            if (form.ShowDialog() == DialogResult.OK)
                result = (int)numeric.Value;

            return result;
        }
    }
}
