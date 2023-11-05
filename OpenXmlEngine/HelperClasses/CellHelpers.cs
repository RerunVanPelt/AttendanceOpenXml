using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEngine.HelperClasses;

public static class CellHelpers
{
    public static Cell AddCellReference(
        this Cell cell,
        uint column,
        uint row)
    {
        cell.CellReference = WorksheetProcessor.SetCellReference(column, row);

        return cell;
    }

    public static Cell AddCellStyle(
        this Cell cell,
        uint styleIndex)
    {
        cell.StyleIndex = styleIndex;

        return cell;
    }

    public static Cell AddCellText(
        this Cell cell,
        string text)
    {
        cell.DataType = CellValues.String;
        cell.CellValue = new CellValue() { Text = text };

        return cell;
    }

    public static Cell AddCellNumber(
        this Cell cell,
        int number)
    {
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(number);

        return cell;
    }

    public static Cell AddCellDate(
        this Cell cell,
        DateTime date)
    {
        cell.DataType = CellValues.Date;
        cell.CellValue = new CellValue(date);

        return cell;
    }

    public static MergeCell AddMergeCell(
        this MergeCell mergeCell,
        StringValue? firstCell,
        StringValue? lastCell)
    {
        mergeCell.Reference = $"{firstCell}:{lastCell}";
        return mergeCell;
    }
}