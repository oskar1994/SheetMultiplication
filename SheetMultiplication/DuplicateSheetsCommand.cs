
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

            var (sheetNames, sheetNumbers) = SheetNamingForm.Show(copyCount, existingSheetNumbers);
            if (sheetNames == null || sheetNumbers == null) return Result.Cancelled;

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
                    sheet.Name = sheetNames[i];
                    sheet.SheetNumber = sheetNumbers[i];


                    // Find the new Title Block instance on the new sheet
                    var newTbCollector = new FilteredElementCollector(doc, sheet.Id)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsNotElementType();
                    var newTitleBlock = newTbCollector.FirstOrDefault();

                    if (newTitleBlock != null && titleBlock != null)
                    {
                        // Copy instance parameters
                        foreach (Parameter param in titleBlock.Parameters)
                        {
                            // Skip copying the "Sheet Name" parameter
                            if (param.Definition is InternalDefinition idef &&
                               (idef.BuiltInParameter == BuiltInParameter.SHEET_NAME ||
                                idef.BuiltInParameter == BuiltInParameter.SHEET_NUMBER))
                                continue;

                            if (!param.IsReadOnly && param.StorageType != StorageType.None)
                            {
                                Parameter newParam = newTitleBlock.get_Parameter(param.Definition);
                                if (newParam != null && !newParam.IsReadOnly)
                                {
                                    switch (param.StorageType)
                                    {
                                        case StorageType.String:
                                            newParam.Set(param.AsString());
                                            break;
                                        case StorageType.Double:
                                            newParam.Set(param.AsDouble());
                                            break;
                                        case StorageType.Integer:
                                            newParam.Set(param.AsInteger());
                                            break;
                                        case StorageType.ElementId:
                                            newParam.Set(param.AsElementId());
                                            break;
                                    }
                                }
                            }
                        }
                    }

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
        public static (List<string>, List<string>) Show(int count, HashSet<string> existingNumbers)
        {
            List<string> names = new List<string>();
            List<string> numbers = new List<string>();

            System.Windows.Forms.Form form = new System.Windows.Forms.Form { Width = 600, Height = 200 + count * 30 };
            form.AutoScroll = true;

            // Headery
            Label numHeader = new Label { Text = "Numer arkusza", Left = 10, Top = 10, Width = 200 };
            Label nameHeader = new Label { Text = "Nazwa arkusza", Left = 220, Top = 10, Width = 350 };
            form.Controls.Add(numHeader);
            form.Controls.Add(nameHeader);

            // Pola tekstowe
            var numBoxes = new List<System.Windows.Forms.TextBox>();
            var nameBoxes = new List<System.Windows.Forms.TextBox>();
            for (int i = 0; i < count; i++)
            {
                var numBox = new System.Windows.Forms.TextBox
                {
                    Left = 10,
                    Top = 40 + i * 30,
                    Width = 200
                };
                var nameBox = new System.Windows.Forms.TextBox
                {
                    Left = 220,
                    Top = 40 + i * 30,
                    Width = 350
                };
                form.Controls.Add(numBox);
                form.Controls.Add(nameBox);
                numBoxes.Add(numBox);
                nameBoxes.Add(nameBox);
            }

            Button ok = new Button { Text = "OK", Left = 250, Top = 50 + count * 30, Width = 100 };
            ok.Click += (s, e) =>
            {
                var tempNumbers = new List<string>();
                var tempNames = new List<string>();
                for (int i = 0; i < count; i++)
                {
                    string numVal = numBoxes[i].Text.Trim();
                    string nameVal = nameBoxes[i].Text.Trim();
                    if (string.IsNullOrEmpty(numVal) || existingNumbers.Contains(numVal) || tempNumbers.Contains(numVal))
                    {
                        MessageBox.Show($"Nieprawidłowy lub powtórzony numer arkusza: \"{numVal}\"", "Błąd");
                        return;
                    }
                    if (string.IsNullOrEmpty(nameVal))
                    {
                        MessageBox.Show($"Podaj nazwę arkusza dla wiersza {i + 1}.", "Błąd");
                        return;
                    }
                    tempNumbers.Add(numVal);
                    tempNames.Add(nameVal);
                }
                numbers = tempNumbers;
                names = tempNames;
                form.DialogResult = DialogResult.OK;
            };

            form.Controls.Add(ok);

            var result = form.ShowDialog();
            return result == DialogResult.OK ? (names, numbers) : (null, null);
        }



    }
}
