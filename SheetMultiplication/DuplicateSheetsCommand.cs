
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

namespace SheetMultiplication
{
    [Transaction(TransactionMode.Manual)]
    public class DuplicateSheetsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 1. Lista arkuszy
            var allSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsTemplate)
                .ToList();

            ViewSheet selectedSheet = SheetSelectionForm.Show(allSheets);
            if (selectedSheet == null) return Result.Cancelled;

            // 2. Liczba kopii
            int copyCount = SheetCopyCountForm.Show();
            if (copyCount <= 0) return Result.Cancelled;

            // 3. Nazwa i numery
            var existingSheetNumbers = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Select(s => s.SheetNumber)
                .ToHashSet();

            var (sheetName, sheetNumbers) = SheetNamingForm.Show(copyCount, existingSheetNumbers);
            if (sheetName == null || sheetNumbers == null) return Result.Cancelled;

            // 4. Pobierz title block
            var tbCollector = new FilteredElementCollector(doc, selectedSheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType();
            var titleBlock = tbCollector.FirstOrDefault();
            if (titleBlock == null)
            {
                TaskDialog.Show("Błąd", "Nie znaleziono bloku tytułowego.");
                return Result.Failed;
            }
            ElementId titleBlockTypeId = titleBlock.GetTypeId();

            // 5. Tworzenie kopii
            using (Transaction t = new Transaction(doc, "Powielanie arkuszy"))
            {
                t.Start();
                // Get all viewports on the selected sheet
                var viewports = new FilteredElementCollector(doc, selectedSheet.Id)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .ToList();

                // Filter only legend viewports
                var legendViewports = viewports
                    .Where(vp =>
                    {
                        var view = doc.GetElement(vp.ViewId) as Autodesk.Revit.DB.View;
                        return view != null && view.ViewType == ViewType.Legend;
                    })
                    .ToList();

                for (int i = 0; i < copyCount; i++)
                {
                    ViewSheet sheet = ViewSheet.Create(doc, titleBlockTypeId);
                    sheet.Name = sheetName;
                    sheet.SheetNumber = sheetNumbers[i];

                    // Get bounding box of source and target sheets
                    BoundingBoxUV sourceOutline = selectedSheet.Outline;
                    BoundingBoxUV targetOutline = sheet.Outline;

                    // Lower-left corner of the source and target sheets
                    XYZ sourceOrigin = new XYZ(sourceOutline.Min.U, sourceOutline.Min.V, 0);
                    XYZ targetOrigin = new XYZ(targetOutline.Min.U, targetOutline.Min.V, 0);

                    foreach (var legendVp in legendViewports)
                    {
                        var legendView = doc.GetElement(legendVp.ViewId) as Autodesk.Revit.DB.View;
                        if (legendView != null)
                        {
                            // Offset from source sheet origin to legend center
                            XYZ sourceCenter = legendVp.GetBoxCenter();
                            XYZ offset = sourceCenter - sourceOrigin;

                            // Place at the same offset on the new sheet
                            XYZ targetCenter = targetOrigin + offset;

                            Viewport newVp = Viewport.Create(doc, sheet.Id, legendView.Id, targetCenter);

                            if (newVp != null && legendVp.GetTypeId() != newVp.GetTypeId())
                            {
                                newVp.ChangeTypeId(legendVp.GetTypeId());
                            }
                        }
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("Sukces", $"Utworzono {copyCount} arkuszy.");
            return Result.Succeeded;
        }
    }

    public static class SheetSelectionForm
    {
        public static ViewSheet Show(List<ViewSheet> sheets)
        {
            ViewSheet result = null;

            System.Windows.Forms.Form form = new System.Windows.Forms.Form { Text = "Wybierz arkusz", Width = 400, Height = 500 };
            ListBox listBox = new ListBox { Dock = DockStyle.Fill };
            var map = new Dictionary<string, ViewSheet>();
            foreach (var sheet in sheets)
            {
                string label = $"{sheet.SheetNumber} - {sheet.Name}";
                map[label] = sheet;
                listBox.Items.Add(label);
            }

            Button ok = new Button { Text = "OK", Dock = DockStyle.Bottom };
            ok.Click += (s, e) => form.DialogResult = DialogResult.OK;

            form.Controls.Add(listBox);
            form.Controls.Add(ok);

            if (form.ShowDialog() == DialogResult.OK && listBox.SelectedItem != null)
                result = map[listBox.SelectedItem.ToString()];
            return result;
        }
    }

    public static class SheetCopyCountForm
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

    public static class SheetNamingForm
    {
        public static (string, List<string>) Show(int count, HashSet<string> existingNumbers)
        {
            string nameResult = null;
            List<string> numbers = new List<string>();

            System.Windows.Forms.Form form = new System.Windows.Forms.Form { Width = 420, Height = 200 + count * 30 };
            form.AutoScroll = true;

            Label nameLabel = new Label { Left = 10, Top = 10, Width = 200 };
            System.Windows.Forms.TextBox nameBox = new System.Windows.Forms.TextBox { Left = 10, Top = 30, Width = 370 };

            Label numLabel = new Label { Text = "Numery arkuszy:", Left = 10, Top = 60, Width = 200 };

            var boxes = new List<System.Windows.Forms.TextBox>();
            for (int i = 0; i < count; i++)
            {
                var box = new System.Windows.Forms.TextBox
                {
                    Left = 10,
                    Top = 90 + i * 30,
                    Width = 370
                };
                form.Controls.Add(box);
                boxes.Add(box);
            }

            Button ok = new Button { Left = 150, Top = 100 + count * 30, Width = 100 };
            ok.Click += (s, e) =>
            {
                var tempList = new List<string>();
                foreach (var b in boxes)
                {
                    string val = b.Text.Trim();
                    if (string.IsNullOrEmpty(val) || existingNumbers.Contains(val) || tempList.Contains(val))
                    {
                        MessageBox.Show($"Nieprawidłowy lub powtórzony numer arkusza: \"{val}\"", "Błąd");
                        return;
                    }
                    tempList.Add(val);
                }

                nameResult = nameBox.Text.Trim();
                if (string.IsNullOrEmpty(nameResult))
                {
                    MessageBox.Show("Podaj nazwę arkusza.", "Błąd");
                    return;
                }

                numbers = tempList;
                form.DialogResult = DialogResult.OK;
            };

            form.Controls.Add(nameLabel);
            form.Controls.Add(nameBox);
            form.Controls.Add(numLabel);
            form.Controls.Add(ok);

            var result = form.ShowDialog();
            return result == DialogResult.OK ? (nameResult, numbers) : (null, null);
        }



    }
}
