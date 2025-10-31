using Autodesk.Revit.DB;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SheetMultiplication
{
    internal enum InputFormatType
    {
        Form,
        XLS,
        TEST
    }

    internal class SheetSelectionFormResult
    {
        public SheetSelectionFormResult(ViewSheet selectedSheet, InputFormatType type)
        {
            SelectedSheet = selectedSheet;
            Type = type;
        }

        public ViewSheet SelectedSheet { get; set; }
        public InputFormatType Type { get; set; }
    }

    internal class SheetSelectionForm
    {
        public SheetSelectionFormResult Show(List<ViewSheet> sheets)
        {
            ViewSheet result = null;
            InputFormatType type = InputFormatType.Form;

            System.Windows.Forms.Form form = new System.Windows.Forms.Form { Text = "Wybierz arkusz", Width = 400, Height = 500 };
            ListBox listBox = new ListBox { Dock = DockStyle.Fill };
            var map = new Dictionary<string, ViewSheet>();
            foreach (var sheet in sheets)
            {
                string label = $"{sheet.SheetNumber} - {sheet.Name}";
                map[label] = sheet;
                listBox.Items.Add(label);
            }

            Button importFromForm = new Button { Text = "Import z Okna", Dock = DockStyle.Bottom, Enabled = false };
            Button importFromXLS = new Button { Text = "Import z XLS", Dock = DockStyle.Bottom, Enabled = false };
            Button test = new Button { Text = "Test", Dock = DockStyle.Bottom, Enabled = false };
            importFromForm.Click += (s, e) => { form.DialogResult = DialogResult.OK; type = InputFormatType.Form; };
            importFromXLS.Click += (s, e) => { form.DialogResult = DialogResult.OK; type = InputFormatType.XLS; };
            test.Click += (s, e) => { form.DialogResult = DialogResult.OK; type = InputFormatType.TEST; };
            listBox.SelectedValueChanged += (s, e) => { importFromForm.Enabled = true; importFromXLS.Enabled = true; test.Enabled = true; };


            form.Controls.Add(listBox);
            form.Controls.Add(importFromForm);
            form.Controls.Add(importFromXLS);
            form.Controls.Add(test);

            if (form.ShowDialog() == DialogResult.OK && listBox.SelectedItem != null)
                result = map[listBox.SelectedItem.ToString()];
            return new SheetSelectionFormResult(selectedSheet: result, type: type);
        }
    }
}
