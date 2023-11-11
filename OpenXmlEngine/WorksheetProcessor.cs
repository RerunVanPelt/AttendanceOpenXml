using DocumentFormat.OpenXml;

namespace OpenXmlEngine;

public static class WorksheetProcessor
{
    /// <summary>
    /// Get cell reference from column index
    /// </summary>
    /// <param name="columnIndex">Cell Column index</param>
    /// <returns>Cell reference eg. A</returns>
    public static string SetCellReference(uint columnIndex)
    {
        var result = string.Empty;

        const char firstRef = 'A';
        const uint firstIndex = (uint)firstRef;

        while (columnIndex > 0)
        {
            var mod = (columnIndex - 1) % 26;
            result = result.Insert(0, ((char)(firstIndex + mod)).ToString());
            columnIndex = (columnIndex - mod) / 26;
        }

        return result;
    }

    /// <summary>
    /// Get cell reference from column index and row index
    /// </summary>
    /// <param name="columnIndex">Column index</param>
    /// <param name="rowIndex">Row index</param>
    /// <returns>Cell reference eg. A1</returns>
    public static string SetCellReference(
        uint columnIndex,
        uint rowIndex)
    {
        var result = string.Empty;

        const char firstRef = 'A';
        const uint firstIndex = (uint)firstRef;

        while (columnIndex > 0)
        {
            var mod = (columnIndex - 1) % 26;
            result = result.Insert(0, ((char)(firstIndex + mod)).ToString());
            columnIndex = (columnIndex - mod) / 26;
        }

        result += rowIndex;
        return result;
    }

    public static (string column, uint row) SplitCellReference(string? cellReference)
    {
        var column = string.Empty;
        var rowString = string.Empty;

        if (cellReference != null)
        {
            foreach (var ch in cellReference)
                if (int.TryParse(ch.ToString(), out var r))
                {
                    rowString += r;
                }
                else
                {
                    column += ch;
                }
        }


        var row = uint.Parse(rowString);

        return (column, row);
    }


    public static StringValue GetRange(List<StringValue> cellReferences)
    {
        var firstReference = cellReferences.First().Value?.Split(':')[0];
        var lastReference = cellReferences.Last().Value?.Split(':')[1];

        StringValue output = $"{firstReference}:{lastReference}";

        return output;
    }

    public static List<ListValue<StringValue>> JoinLists(
        List<StringValue> cellReferences1,
        List<StringValue> cellReferences2)
    {
        var output = new List<ListValue<StringValue>>();

        for (var i = 0; i < cellReferences1.Count; i++)
        {
            var firstReference = cellReferences1[i].Value?.Split(":")[0];
            var lastReference = cellReferences2[i].Value?.Split(":")[1];

            output.Add(new ListValue<StringValue>()
            {
                Items = { $"{firstReference}:{lastReference}" }
            });
        }

        return output;
    }
}