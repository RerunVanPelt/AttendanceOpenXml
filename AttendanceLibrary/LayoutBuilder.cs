using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AttendanceLibrary;

public static class LayoutBuilder
{
    public static Columns SetColumns()
    {
        var columns = new Columns();

        Column col1 = new()
        {
            Min = 1,
            Max = 2,
            Width = 3,
            CustomWidth = true
        };
        columns.Append(col1);

        Column col2 = new()
        {
            Min = 2,
            Max = 2,
            Width = 20,
            CustomWidth = true
        };
        columns.Append(col2);

        Column col3To34 = new()
        {
            Min = 3,
            Max = 34,
            Width = 4,
            CustomWidth = true
        };
        columns.Append(col3To34);

        return columns;
    }
}