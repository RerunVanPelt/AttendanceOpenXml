using AttendanceLibrary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEngine;
using OpenXmlEngine.HelperClasses;
using System.Globalization;

const string filePath = @"D:\SampleData\OpenXmlSamples\attendace.xlsx";

var year = DateTime.Now.Year;

SpreadsheetProcessor excel = new(filePath);
var firstSheetName = $"{year}";
var secondSheetName = $"{year}_FT";

var headlineList = new List<string>()
{
    "Mitarbeiter",
    "U", "Urlaub",
    "GZ", "Gleitzeit",
    "HO", "Home-office",
    "D", "Dienstreise",
    "AU", "A-unfähig",
    "SU", "S-Urlaub",
    "Urlaubstage",
    "Resturlaub"
};

var employeeList = new List<string>()
{
    "Employee 1",
    "Employee 2",
    "Employee 3",
    "Employee 4",
    "Employee 5",
    "Employee 6",
    "Employee 7"
};

excel.CreateSpreadsheet(new[]
{
    firstSheetName,
    secondSheetName
});

using var spreadsheet = SpreadsheetDocument.Open(filePath, true);

#region Spreadsheet Parts

var wbPart = spreadsheet.WorkbookPart;
var wb = wbPart?.Workbook;

var wbStylesPart = wbPart?.WorkbookStylesPart ?? wbPart?.AddNewPart<WorkbookStylesPart>();
var styleSheet = wbStylesPart?.Stylesheet ?? new Stylesheet();

var firstSheetId = wb?.Descendants<Sheet>()
    .First(s => s.Name == firstSheetName).Id?.Value;

var wsPart = (WorksheetPart)wbPart?.GetPartById(firstSheetId!)!;
var ws = wsPart.Worksheet;

var firstSheetData = ws.GetFirstChild<SheetData>();

var mergeCells = ws.Elements<MergeCells>().Any()
    ? ws.Elements<MergeCells>().First()
    : new MergeCells();

#endregion

#region Style

// Numbering Formats
styleSheet.AppendChild(StyleSheetBuilder.NumberingFormats());

// Fonts
styleSheet.AppendChild(StyleSheetBuilder.SetFonts());

// Fills
styleSheet.AppendChild(StyleSheetBuilder.SetFills());

// Borders
styleSheet.AppendChild(StyleSheetBuilder.SetBorders());

// CellFormats
styleSheet.AppendChild(StyleSheetBuilder.SetCellFormats());

//TODO: CellStyleFormats
//styleSheet.AppendChild<CellStyleFormats>(StyleSheetBuilder.SetCellStyleFormats());

//TODO: CellStyles
//styleSheet.AppendChild<CellStyles>(StyleSheetBuilder.SetCellStyles());

// DifferntialFormats
styleSheet.AppendChild(StyleSheetBuilder.SetDifferentialFormats());

//TODO: TableStyles

//TODO: StylesheetExtensionList

#endregion

#region Columns

var columns = LayoutBuilder.SetColumns();
wsPart.Worksheet.InsertBefore(columns, firstSheetData);

#endregion

#region Lists

var rowList = new List<Row>();
var mergeList = new List<MergeCell>();
var conditionList = new List<ConditionalFormatting>();


var daysCellReferences = new ListValue<StringValue>();
var weekendCellReferences = new ListValue<StringValue>();
var specialDaysCellReferences = new ListValue<StringValue>();


// TODO: List with ListValue<StringValue>
//List<ListValue<StringValue>> list = new();

Dictionary<string, uint> specialDaysFormats = new()
{
    { "U", 2U },
    { "GZ", 3U },
    { "HO", 4U },
    { "D", 5U },
    { "AU", 6U },
    { "SU", 7U }
};

Dictionary<string, string> specialDaysReferences = new();

var rowIndex = 1U;

#endregion

#region Title

var titleRow = new Row().AddIndexToRow(rowIndex).AddCellsToRow(34);
titleRow.Descendants<Cell>()
    .First()
    .AddCellText($"Personalplannung {firstSheetName}")
    .AddCellStyle(1);
rowList.Add(titleRow);

mergeList.Add(new MergeCell().AddMergeCell(
    titleRow.Descendants<Cell>().First().CellReference,
    titleRow.Descendants<Cell>().Last().CellReference));
rowIndex++;

#endregion

#region Headline Row (Table)

var headlineRow = new Row().AddIndexToRow(++rowIndex).AddCellsToRow(34);
var cells = headlineRow.Descendants<Cell>().ToList();
var listIndex = 0;
var styleIndex = 6U;
for (var i = 0; i < cells.Count; i++)
{
    switch (i)
    {
        case 0:
            cells[i].AddCellStyle(2);
            break;
        case 1:
            cells[i].AddCellText(headlineList[listIndex++]).AddCellStyle(3);
            break;
        case >= 2 and <= 25:
            cells[i].AddCellText(headlineList[listIndex++]).AddCellStyle(styleIndex++);
            cells[i + 1].AddCellText(headlineList[listIndex++]).AddCellStyle(4);
            cells[i + 2].AddCellText(headlineList[listIndex]).AddCellStyle(4);
            cells[i + 3].AddCellText(headlineList[listIndex]).AddCellStyle(4);
            mergeList.Add(new MergeCell().AddMergeCell(
                cells[i + 1].CellReference,
                cells[i + 3].CellReference));


            var firstValue = cells[i].CellValue?.Text;
            var lastValue = cells[i].CellReference?.Value;
            if (firstValue != null &&
                lastValue != null)
            {
                specialDaysReferences.Add(firstValue, lastValue);
            }

            i += 3;
            break;
        case > 25 and < 30:
            cells[i].AddCellText(headlineList[13]).AddCellStyle(4);
            cells[i + 1].AddCellText(headlineList[13]).AddCellStyle(4);
            cells[i + 2].AddCellText(headlineList[13]).AddCellStyle(4);
            cells[i + 3].AddCellText(headlineList[13]).AddCellStyle(4);
            mergeList.Add(new MergeCell().AddMergeCell(
                cells[i].CellReference,
                cells[i + 3].CellReference));
            i += 3;
            break;
        case 30:
            cells[i].AddCellText(headlineList[14]).AddCellStyle(5);
            cells[i + 1].AddCellText(headlineList[14]).AddCellStyle(5);
            cells[i + 2].AddCellText(headlineList[14]).AddCellStyle(5);
            cells[i + 3].AddCellText(headlineList[14]).AddCellStyle(5);
            mergeList.Add(new MergeCell().AddMergeCell(
                cells[i].CellReference,
                cells[i + 3].CellReference));
            i += 3;
            break;
    }
}

rowList.Add(headlineRow);

#endregion

#region Employee Rows

// Set Employee rows in table

foreach (var employee in employeeList)
{
    var index = employeeList.IndexOf(employee);

    var employeeRow = new Row().AddIndexToRow(++rowIndex).AddCellsToRow(34);

    cells = employeeRow.Descendants<Cell>().ToList();
    for (var i = 0; i < cells.Count; i++)
    {
        switch (i)
        {
            case 0:
                cells[i].AddCellStyle(2);
                break;

            case 1:
                styleIndex = index != employeeList.Count - 1 ? 13 : (uint)16;
                cells[i].AddCellText(employee).AddCellStyle(styleIndex);
                break;

            case >= 2 and < 30:
                styleIndex = index != employeeList.Count - 1 ? 15 : (uint)18;
                cells[i].AddCellStyle(styleIndex);
                cells[i + 1].AddCellStyle(styleIndex);
                cells[i + 2].AddCellStyle(styleIndex);
                cells[i + 3].AddCellStyle(styleIndex);
                mergeList.Add(new MergeCell().AddMergeCell(
                    cells[i].CellReference,
                    cells[i + 3].CellReference));
                i += 3;
                break;

            case >= 30:
                styleIndex = index != employeeList.Count - 1 ? 14 : (uint)17;
                cells[i].AddCellStyle(styleIndex);
                cells[i + 1].AddCellStyle(styleIndex);
                cells[i + 2].AddCellStyle(styleIndex);
                cells[i + 3].AddCellStyle(styleIndex);
                mergeList.Add(new MergeCell().AddMergeCell(
                    cells[i].CellReference,
                    cells[i + 3].CellReference));
                i += 3;
                break;
        }
    }

    rowList.Add(employeeRow);
}

++rowIndex;
++rowIndex;

#endregion


#region Calender

// Set Calender

for (var month = 1; month <= 12; month++)
{
    ++rowIndex;
    var daysInMonth = DateTime.DaysInMonth(year, month);
    var numberOfCells = (uint)(daysInMonth + 2);

    #region Week Numbers

    var weekRow = new Row().AddIndexToRow(rowIndex).AddCellsToRow(numberOfCells);

    cells = weekRow.Descendants<Cell>().ToList();

    var firstCellToMerge = cells[2];
    var lastCellToMerge = new Cell();
    var lastCell = cells.LastOrDefault();

    for (var i = 0; i < cells.Count; i++)
    {
        switch (i)
        {
            case 0:
                cells[i].AddCellStyle(19);
                break;

            case 1:
                cells[i].AddCellStyle(20).AddCellText("KW");
                break;

            default:
                var currWeekNumber = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(
                    new DateTime(year, month, i - 1), CalendarWeekRule.FirstFullWeek,
                    DayOfWeek.Monday);

                var currCell = cells[i].AddCellStyle(21).AddCellNumber(currWeekNumber);

                if (firstCellToMerge.InnerText == currCell.InnerText)
                {
                    lastCellToMerge = currCell;
                }
                else
                {
                    if (firstCellToMerge.CellReference != lastCellToMerge.CellReference)
                    {
                        mergeList.Add(new MergeCell().AddMergeCell(
                            firstCellToMerge.CellReference,
                            lastCellToMerge.CellReference));
                    }

                    firstCellToMerge = currCell;
                }

                if (currCell == lastCell &&
                    firstCellToMerge.CellReference != lastCell.CellReference)

                {
                    mergeList.Add(new MergeCell().AddMergeCell(
                        firstCellToMerge.CellReference,
                        lastCell.CellReference));
                }

                break;
        }
    }

    rowList.Add(weekRow);

    #endregion

    #region Days in Month

    ++rowIndex;
    var daysRow = new Row().AddIndexToRow(rowIndex).AddCellsToRow(numberOfCells);
    cells = daysRow.Descendants<Cell>().ToList();

    for (var i = 0; i < cells.Count; i++)
    {
        DateTime currDay;
        switch (i)
        {
            case 0:
                cells[i].AddCellStyle(19);
                break;
            case 1:
                currDay = new DateTime(year, month, i);
                cells[i].AddCellStyle(22).AddCellDate(currDay);
                break;
            default:
                currDay = new DateTime(year, month, i - 1);
                cells[i].AddCellStyle(23).AddCellDate(currDay);
                break;
        }
    }

    // Cell References for Conditional Formatting 
    daysCellReferences.Items.Add($"{cells[2].CellReference}:{cells.Last().CellReference}");
    weekendCellReferences.Items.Add($"{cells[2].CellReference}:{cells.Last().CellReference}");

    // Add Row to List
    rowList.Add(daysRow);

    #endregion

    #region Employees table

    foreach (var employee in employeeList)
    {
        ++rowIndex;
        var employeeRow = new Row().AddIndexToRow(rowIndex).AddCellsToRow(numberOfCells);
        cells = employeeRow.Descendants<Cell>().ToList();

        for (var i = 0; i < cells.Count; i++)
        {
            switch (i)
            {
                case 0:
                    cells[i].AddCellStyle(19);
                    break;
                case 1:
                    cells[i].AddCellStyle(24).AddCellText(employee);
                    break;
                default:
                    cells[i].AddCellStyle(15);
                    break;
            }
        }


        // Cell References for Conditional Formatting
        weekendCellReferences.Items.Add($"{cells[2].CellReference}:{cells.Last().CellReference}");
        specialDaysCellReferences.Items.Add(
            $"{cells[2].CellReference}:{cells.Last().CellReference}");

        // Add Row to List
        rowList.Add(employeeRow);
    }

    #endregion

    // Weekend Formating 
    // TODO: Try List<ListValue<StringValue>()
    conditionList.Add(ConditionalsBuilder.WeekendFormatting(weekendCellReferences));
    weekendCellReferences = new ListValue<StringValue>();

    // Space between month table
    ++rowIndex;
}

#endregion


// Append rows
firstSheetData?.Append(rowList);

// Append merge cells
mergeCells.Append(mergeList);
ws.Append(mergeCells);

// Append conditional formatting
conditionList.Add(ConditionalsBuilder.TodayFormatting(daysCellReferences));

foreach (var reference in specialDaysReferences)
{
    var format = specialDaysFormats.First(c => c.Key == reference.Key).Value;

    conditionList.Add(ConditionalsBuilder.SpecialDayFormatting(
        specialDaysCellReferences,
        reference.Value,
        format));
}


ws.Append(conditionList);


spreadsheet.Dispose();

//Console.WriteLine("Press any key to exit application ...");
//Console.ReadKey();