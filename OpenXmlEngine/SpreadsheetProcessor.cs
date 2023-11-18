using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEngine;

public class SpreadsheetProcessor
{
    private readonly string _filePath;

    /// <summary>
    /// Creates a SpreadsheetDocument using supported file name
    /// </summary>
    /// <param name="filePath">Path to file</param>
    public SpreadsheetProcessor(string filePath)
    {
        _filePath = filePath;
    }

    /// <summary>
    /// Creates a SpreadsheetDocument with given sheet names
    /// </summary>
    /// <param name="sheetNames">Sheet names</param>
    public void CreateSpreadsheet(IEnumerable<string> sheetNames)
    {
        using var spreadsheet =
            SpreadsheetDocument.Create(_filePath, SpreadsheetDocumentType.Workbook);

        // Add new workbook part to the document
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        // Add new workbook to the workbook part
        var wb = workbookPart.Workbook;

        foreach (var name in sheetNames)
        {
            // Add worksheet part to the workbook
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add worksheet to the worksheet part
            var ws = worksheetPart.Worksheet;

            // Add sheets to the workbook
            var sheets = wb.GetFirstChild<Sheets>() ?? wb.AppendChild<Sheets>(new Sheets());

            sheets.Append(new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheets?.Elements<Sheet>().Max(s => s.SheetId?.Value) + 1 ?? 1,
                Name = name
            });
        }

        //Add WorkbookStylesPart to the document
        var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        workbookStylesPart.Stylesheet = new Stylesheet();
    }

    /// <summary>
    /// Adds new tables to SpreadsheetDocument
    /// </summary>
    /// <param name="sheetNames">Sheet names</param>
    public void AddSheetsToSpreadsheet(IEnumerable<string> sheetNames)
    {
        using var spreadsheet =
            SpreadsheetDocument.Open(_filePath, true);

        // Set wbPart
        var wbPart = spreadsheet.WorkbookPart;
        var wb = wbPart?.Workbook;

        foreach (var name in sheetNames)
        {
            // Add worksheet part to the workbook
            var worksheetPart = wbPart?.AddNewPart<WorksheetPart>();
            worksheetPart!.Worksheet = new Worksheet(new SheetData());

            // Add worksheet to the worksheet part
            var ws = worksheetPart.Worksheet;

            // Add sheets to the workbook
            var sheets = wb?.GetFirstChild<Sheets>() ?? wb?.AppendChild<Sheets>(new Sheets());

            sheets?.Append(new Sheet()
            {
                Id = wbPart?.GetIdOfPart(worksheetPart),
                SheetId = sheets?.Elements<Sheet>().Max(s => s.SheetId?.Value) + 1 ?? 1,
                Name = name
            });
        }
    }
}