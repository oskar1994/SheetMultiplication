using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace SheetMultiplication
{

    internal class XLSSheetDataImporter
    {
        internal List<Sheet> ImportSheetData(HashSet<string> existingSheetNumbers)
        {
            List<Sheet> sheets = new List<Sheet>();

            using (OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Wybierz plik Excel"
            })
            {
                if(ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook(ofd.FileName))
                        {
                            var worksheet = workbook.Worksheet(1); // First sheet
                            
                            bool firstRow = true;
                            foreach (var row in worksheet.RowsUsed())
                            {                             
                                if (firstRow)
                                {
                                    firstRow = false;
                                    continue;  // skip header row
                                }

                                sheets.Add(new Sheet
                                {
                                    Number = row.Cell(1).GetValue<string>(),
                                    Name = row.Cell(2).GetValue<string>(),
                                });                
                            }                                           
                        }

                        if(!ValidateImportedSheets(sheets, existingSheetNumbers))
                        {
                            return null;
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Błąd podczas importu arkuszy z pliku Excel: {ex.Message}", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }
                }
            }


            return sheets;
        }


        private bool ValidateImportedSheets(List<Sheet> sheets, HashSet<string> existingSheetNumbers)
        {
            // Check for empty or null sheet numbers
            var emptyNumbers = sheets
                .Select((s, index) => new { Sheet = s, Row = index + 2 }) // +2 if row 1 is header
                .Where(x => string.IsNullOrWhiteSpace(x.Sheet.Number))
                .ToList();

            if (emptyNumbers.Any())
            {
                string rows = string.Join(", ", emptyNumbers.Select(x => x.Row));
                MessageBox.Show(
                    $"Błąd podczas importu arkuszy z pliku Excel:\nNiektóre wiersze nie mają przypisanego numeru arkusza (wiersze: {rows}).",
                    "Błąd",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }

            // Check for empty or null sheet names
            var emptyNames = sheets
                .Select((s, index) => new { Sheet = s, Row = index + 2 })
                .Where(x => string.IsNullOrWhiteSpace(x.Sheet.Name))
                .ToList();

            if (emptyNames.Any())
            {
                string rows = string.Join(", ", emptyNames.Select(x => x.Row));
                MessageBox.Show(
                    $"Błąd podczas importu arkuszy z pliku Excel:\nNiektóre wiersze nie mają przypisanej nazwy arkusza (wiersze: {rows}).",
                    "Błąd",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }

            // Check for duplicates in the imported file itself
            var duplicatedInFile = sheets
                .GroupBy(s => s.Number.Trim(), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .ToList();

            if (duplicatedInFile.Any())
            {
                string duplicates = string.Join(", ", duplicatedInFile);
                MessageBox.Show(
                    $"Błąd podczas importu arkuszy z pliku Excel:\nW pliku występują zduplikowane numery arkuszy: {duplicates}",
                    "Błąd",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }

            // Check for duplicates against existing system data
            var duplicatedInSystem = sheets
                .Where(s => existingSheetNumbers.Contains(s.Number.Trim()))
                .Select(s => s.Number)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (duplicatedInSystem.Any())
            {
                string duplicates = string.Join(", ", duplicatedInSystem);
                MessageBox.Show(
                    $"Błąd podczas importu arkuszy z pliku Excel:\nW systemie istnieją już arkusze o numerach: {duplicates}",
                    "Błąd",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

    }
}