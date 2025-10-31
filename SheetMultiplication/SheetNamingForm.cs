using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SheetMultiplication
{
    internal static class SheetNamingForm
    {
        public static (List<string>, List<string>) Show(int count, HashSet<string> existingNumbers)
        {
            List<string> names = new List<string>();
            List<string> numbers = new List<string>();

            // Use case-insensitive check if desired (caller can provide this)
            // existingNumbers = new HashSet<string>(existingNumbers, StringComparer.OrdinalIgnoreCase);

            Form form = new Form
            {
                Width = 760,
                Height = Math.Min(800, Math.Max(260, 200 + count * 30)),
                StartPosition = FormStartPosition.CenterParent,
                AutoScroll = true
            };

            Label numHeader = new Label { Text = "Numer arkusza", Left = 10, Top = 10, Width = 200 };
            Label nameHeader = new Label { Text = "Nazwa arkusza", Left = 220, Top = 10, Width = 450 };
            Label actionHeader = new Label { Text = "Akcja", Left = 680, Top = 10, Width = 70 };

            form.Controls.Add(numHeader);
            form.Controls.Add(nameHeader);
            form.Controls.Add(actionHeader);

            // Flow panel to hold rows
            FlowLayoutPanel flow = new FlowLayoutPanel
            {
                Left = 8,
                Top = 40,
                Width = form.ClientSize.Width - 20,
                Height = form.ClientSize.Height - 120,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };
            form.Controls.Add(flow);

            // Helper to create a single row panel (numBox, nameBox, deleteBtn)
            Panel CreateRow(string initialNumber = "", string initialName = "")
            {
                Panel row = new Panel
                {
                    Width = flow.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 4,
                    Height = 30,
                    Margin = new Padding(0, 0, 0, 3)
                };

                TextBox numBox = new TextBox
                {
                    Left = 0,
                    Top = 3,
                    Width = 200,
                    Text = initialNumber
                };
                TextBox nameBox = new TextBox
                {
                    Left = 210,
                    Top = 3,
                    Width = 450,
                    Text = initialName
                };
                Button deleteBtn = new Button
                {
                    Text = "Usuń",
                    Left = 670,
                    Top = 1,
                    Width = 70,
                    Height = 24
                };

                deleteBtn.Click += (s, e) =>
                {
                    // remove the entire row panel
                    flow.Controls.Remove(row);
                };

                // Make controls resize with panel
                numBox.Anchor = AnchorStyles.Left | AnchorStyles.Top;
                nameBox.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
                deleteBtn.Anchor = AnchorStyles.Top | AnchorStyles.Right;

                row.Controls.Add(numBox);
                row.Controls.Add(nameBox);
                row.Controls.Add(deleteBtn);

                // Store controls in Tag for easy retrieval later
                row.Tag = new Tuple<TextBox, TextBox>(numBox, nameBox);

                return row;
            }

            // add initial rows
            for (int i = 0; i < Math.Max(1, count); i++)
            {
                flow.Controls.Add(CreateRow());
            }

            // Buttons
            Button ok = new Button { Text = "OK", Width = 100, Height = 28, Left = 250 };
            Button addRowButton = new Button { Text = "Dodaj wiersz", Width = 120, Height = 28, Left = 360 };
            Button cancel = new Button { Text = "Anuluj", Width = 100, Height = 28, Left = 490 };

            // Position buttons panel
            Panel buttons = new Panel
            {
                Left = 8,
                Top = form.ClientSize.Height - 60,
                Width = form.ClientSize.Width - 20,
                Height = 40,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            ok.Left = (buttons.Width / 2) - ok.Width - 10;
            addRowButton.Left = (buttons.Width / 2) + 10;
            cancel.Left = addRowButton.Left + addRowButton.Width + 10;

            ok.Anchor = AnchorStyles.None;
            addRowButton.Anchor = AnchorStyles.None;
            cancel.Anchor = AnchorStyles.None;

            buttons.Controls.Add(ok);
            buttons.Controls.Add(addRowButton);
            buttons.Controls.Add(cancel);
            form.Controls.Add(buttons);

            form.AcceptButton = ok;
            form.CancelButton = cancel;

            addRowButton.Click += (s, e) =>
            {
                flow.Controls.Add(CreateRow());
            };

            cancel.Click += (s, e) =>
            {
                form.DialogResult = DialogResult.Cancel;
            };

            ok.Click += (s, e) =>
            {
                var tempNumbers = new List<string>();
                var tempNames = new List<string>();
                // Collect rows from flow panel
                foreach (Control ctrl in flow.Controls)
                {
                    if (!(ctrl is Panel row)) continue;
                    var tuple = row.Tag as Tuple<TextBox, TextBox>;
                    if (tuple == null) continue;
                    string numVal = tuple.Item1.Text.Trim();
                    string nameVal = tuple.Item2.Text.Trim();

                    if (string.IsNullOrEmpty(numVal))
                    {
                        MessageBox.Show($"Nieprawidłowy numer arkusza: \"{numVal}\"", "Błąd");
                        return;
                    }

                    // Case-insensitive duplicate check against existingNumbers and tempNumbers
                    if (existingNumbers != null && existingNumbers.Contains(numVal))
                    {
                        MessageBox.Show($"Numer arkusza już istnieje: \"{numVal}\"", "Błąd");
                        return;
                    }
                    if (tempNumbers.Contains(numVal))
                    {
                        MessageBox.Show($"Powtórzony numer arkusza w formularzu: \"{numVal}\"", "Błąd");
                        return;
                    }

                    if (string.IsNullOrEmpty(nameVal))
                    {
                        MessageBox.Show($"Podaj nazwę arkusza dla wiersza.", "Błąd");
                        return;
                    }

                    tempNumbers.Add(numVal);
                    tempNames.Add(nameVal);
                }

                if (tempNumbers.Count == 0)
                {
                    MessageBox.Show("Musisz dodać co najmniej jeden wiersz.", "Błąd");
                    return;
                }

                numbers = tempNumbers;
                names = tempNames;
                form.DialogResult = DialogResult.OK;
            };

            var result = form.ShowDialog();
            return result == DialogResult.OK ? (names, numbers) : (null, null);
        }
    }
}