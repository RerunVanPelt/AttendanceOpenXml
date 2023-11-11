using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlEngine;

namespace AttendanceLibrary;

public static class ConditionalsBuilder
{
    public static ConditionalFormatting TodayFormatting(ListValue<StringValue> cellReferences)
    {
        var firstValue = cellReferences.Items.First().Value;
        var firstCell = firstValue?[..firstValue.IndexOf(":", StringComparison.Ordinal)];

        ConditionalFormatting condition = new()
        {
            SequenceOfReferences = cellReferences
        };

        ConditionalFormattingRule todayRule = new()
        {
            Type = ConditionalFormatValues.TimePeriod,
            FormatId = 0U,
            Priority = 1,
            TimePeriod = TimePeriodValues.Today,
            StopIfTrue = true
        };


        Formula todayFormula = new()
        {
            Text = $"FLOOR({firstCell},1)=TODAY()"
        };

        todayRule.Append(todayFormula);
        condition.Append(todayRule);
        return condition;
    }

    public static ConditionalFormatting WeekendFormatting(ListValue<StringValue> cellReferences)
    {
        var firstValue = cellReferences.First().Value;
        var firstCell = firstValue?[..firstValue.IndexOf(":", StringComparison.Ordinal)];

        var (column, row) = WorksheetProcessor.SplitCellReference(firstCell);

        ConditionalFormatting condition = new()
        {
            SequenceOfReferences = cellReferences
        };

        ConditionalFormattingRule weekendRule = new()
        {
            Type = ConditionalFormatValues.Expression,
            FormatId = 1U,
            Priority = 2,
            StopIfTrue = true
        };

        Formula weekendFormula = new()
        {
            Text = $"WEEKDAY({column}${row},2)>5"
        };

        weekendRule.Append(weekendFormula);
        condition.Append(weekendRule);

        return condition;
    }

    public static ConditionalFormatting SpecialDayFormatting(
        ListValue<StringValue> cellReferences,
        string dayReference,
        uint formatId)
    {
        var (column, row) = WorksheetProcessor.SplitCellReference(dayReference);

        ConditionalFormatting condition = new()
        {
            SequenceOfReferences = cellReferences
        };

        ConditionalFormattingRule uDayRule = new()
        {
            Type = ConditionalFormatValues.CellIs,
            FormatId = formatId,
            Priority = 3,
            Operator = ConditionalFormattingOperatorValues.Equal,
            StopIfTrue = true
        };

        Formula uDayFormula = new()
        {
            Text = $"${column}${row}"
        };

        uDayRule.Append(uDayFormula);
        condition.Append(uDayRule);
        return condition;
    }

    public static ConditionalFormatting CrossHolidaysFormatting(
        ListValue<StringValue> cellReferences)
    {
        var firstValue = cellReferences.First().Value;
        var firstCell = firstValue?[..firstValue.IndexOf(":", StringComparison.Ordinal)];

        var (column, row) = WorksheetProcessor.SplitCellReference(firstCell);

        ConditionalFormatting condition = new()
        {
            SequenceOfReferences = cellReferences
        };

        ConditionalFormattingRule holidayRule = new()
        {
            Type = ConditionalFormatValues.Expression,
            FormatId = 10U,
            Priority = 4,
            StopIfTrue = true
        };

        Formula holidayFormula = new()
        {
            Text =
                $"AND(MATCH({column}${row},\'2023_FT\'!$B:$B, 0), MATCH({column}${row},\'2023_FT\'!$D:$D, 0))"
        };

        holidayRule.Append(holidayFormula);
        condition.Append(holidayRule);

        return condition;

        // formula1925.Text = "AND(MATCH(C$13, \'2023_FT\'!$B:$B, 0), MATCH(C$13, \'2023_FT\'!$D:$D, 0))";
    }


    public static ConditionalFormatting HolidaysFormatting(ListValue<StringValue> cellReferences)
    {
        var firstValue = cellReferences.First().Value;
        var firstCell = firstValue?[..firstValue.IndexOf(":", StringComparison.Ordinal)];

        var (column, row) = WorksheetProcessor.SplitCellReference(firstCell);

        ConditionalFormatting condition = new()
        {
            SequenceOfReferences = cellReferences
        };

        ConditionalFormattingRule holidayRule = new()
        {
            Type = ConditionalFormatValues.Expression,
            FormatId = 8U,
            Priority = 5,
            StopIfTrue = true
        };

        Formula holidayFormula = new()
        {
            Text = $"MATCH({column}${row},\'2023_FT\'!$B:$B, 0)"
        };

        holidayRule.Append(holidayFormula);
        condition.Append(holidayRule);

        return condition;

        // formula1925.Text = "AND(MATCH(C$13, \'2023_FT\'!$B:$B, 0), MATCH(C$13, \'2023_FT\'!$D:$D, 0))";
    }

    public static ConditionalFormatting SchoolHolidaysFormatting(
        ListValue<StringValue> cellReferences)
    {
        var firstValue = cellReferences.First().Value;
        var firstCell = firstValue?[..firstValue.IndexOf(":", StringComparison.Ordinal)];

        var (column, row) = WorksheetProcessor.SplitCellReference(firstCell);

        ConditionalFormatting condition = new()
        {
            SequenceOfReferences = cellReferences
        };

        ConditionalFormattingRule holidayRule = new()
        {
            Type = ConditionalFormatValues.Expression,
            FormatId = 9U,
            Priority = 5,
            StopIfTrue = true
        };

        Formula holidayFormula = new()
        {
            Text = $"MATCH({column}${row},\'2023_FT\'!$D:$D, 0)"
        };

        holidayRule.Append(holidayFormula);
        condition.Append(holidayRule);

        return condition;

        // formula1925.Text = "AND(MATCH(C$13, \'2023_FT\'!$B:$B, 0), MATCH(C$13, \'2023_FT\'!$D:$D, 0))";
    }
}