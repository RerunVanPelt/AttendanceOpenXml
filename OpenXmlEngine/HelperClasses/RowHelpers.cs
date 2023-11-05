using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlEngine.HelperClasses
{
    public static class RowHelpers
    {
        public static Row AddIndexToRow(this Row row, uint rowIndex)
        {
            row.RowIndex = rowIndex;

            return row;
        }

        public static Row AddCellsToRow(this Row row, uint numberOfCells)
        {
            for (uint i = 1; i <= numberOfCells; i++)
            {
                row.Append(new Cell().AddCellReference(i, row.RowIndex));
            }

            return row;
        }
    }
}
