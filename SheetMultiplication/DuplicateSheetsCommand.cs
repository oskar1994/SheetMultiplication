
using System;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;

namespace SheetMultiplication
{
    [Transaction(TransactionMode.Manual)]
    public class DuplicateSheetsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {

                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // 1. Lista arkuszy
                var allSheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Where(s => !s.IsTemplate)
                    .ToList();

                SheetSelectionForm form = new SheetSelectionForm();
                var sheetSelectionFormResult = form.Show(allSheets);
                if (sheetSelectionFormResult.SelectedSheet == null) return Result.Cancelled;

                var selectedSheet = sheetSelectionFormResult.SelectedSheet;

                var existingSheetNumbers = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .Select(s => s.SheetNumber)
                        .ToHashSet();

                int copyCount = 0;
                List<string> sheetNames = new List<string>();
                List<string> sheetNumbers = new List<string>();

                if (sheetSelectionFormResult.Type == InputFormatType.Form)
                {
                    copyCount = SheetCopyCountForm.Show();
                    if (copyCount <= 0) return Result.Cancelled;
                    var sheetNamingFormResult = SheetNamingForm.Show(copyCount, existingSheetNumbers);
                    sheetNames = sheetNamingFormResult.Item1;
                    sheetNumbers = sheetNamingFormResult.Item2;
                }
                else if (sheetSelectionFormResult.Type == InputFormatType.XLS)
                {
                    var xlsImporter = new XLSSheetDataImporter();
                    var sheets = xlsImporter.ImportSheetData(existingSheetNumbers);
                    if (sheets == null) return Result.Cancelled;
                    sheetNames = sheets.Select(x => x.Name).ToList();
                    sheetNumbers = sheets.Select(x => x.Number).ToList();
                    copyCount = sheets.Count;
                }
                else if (sheetSelectionFormResult.Type == InputFormatType.TEST)
                {
                    var testForm = new TestForm();
                    var dialog = testForm.ShowDialog();
                    return Result.Cancelled;
                }

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


                        // Skopiuj parametry z sekcji "Dane Identyfikacyjne" (i inne niestandardowe) z oryginalnego arkusza
                        foreach (Parameter param in selectedSheet.Parameters)
                        {
                            // Pomijamy wbudowane parametry, które już ustawiliśmy lub których nie chcemy nadpisywać
                            if (param.IsReadOnly ||
                                param.StorageType == StorageType.None ||
                                (param.Definition is InternalDefinition idef &&
                                 (idef.BuiltInParameter == BuiltInParameter.SHEET_NAME ||
                                  idef.BuiltInParameter == BuiltInParameter.SHEET_NUMBER)))
                                continue;

                            Parameter newParam = sheet.get_Parameter(param.Definition);
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

            }
            catch (Exception ex)
            {
                TaskDialog.Show("Błąd", ex.Message + " " + ex.InnerException + " " + ex.StackTrace);
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}
