using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AttendanceLibrary;

public static class EquationBuilder
{
    public static string SumCountIf(ListValue<StringValue> range, string specialDayPosition)
    {
        var text = "SUM(";

        foreach (var pos in range)
        {
            text += $"COUNTIF({pos}, {specialDayPosition}),";
        }

        text = text.Substring(0, text.Length - 1);
        text += ")";

        return text;
    }

    public static string Subtraktion(Cell cell1, Cell cell2)
    {
        var text = $"({cell1.CellReference?.Value}-{cell2.CellReference?.Value})";

        return text;
    }
}