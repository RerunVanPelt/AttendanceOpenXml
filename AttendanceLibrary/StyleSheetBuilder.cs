using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AttendanceLibrary;

public static class StyleSheetBuilder
{
    /// <summary>
    /// Set Fonts for spreadsheet
    /// </summary>
    /// <returns></returns>
    public static Fonts SetFonts()
    {
        Fonts fonts = new();

        // FontId = 0 - Default Font
        Font defaultFont = new()
        {
            FontName = new FontName() { Val = "Calibri" },
            FontSize = new FontSize() { Val = 11 }
        };
        fonts.Append(defaultFont);

        // FontId = 1 - Default Bold Font
        Font boldFont = new()
        {
            FontName = new FontName() { Val = "Calibri" },
            FontSize = new FontSize() { Val = 11 },
            Bold = new Bold()
        };
        fonts.Append(boldFont);

        // FontId = 2 - Title Font
        Font titleFont = new()
        {
            FontSize = new FontSize { Val = 20 },
            Color = new Color { Rgb = HexBinaryValue.FromString("800080") },
            Bold = new Bold()
        };
        fonts.Append(titleFont);

        // Get number of fonts
        fonts.Count = (uint)fonts.ChildElements.Count;
        return fonts;
    }

    public static Fills SetFills()
    {
        Fills fills = new();

        // FillId = 0 - Default Fill (reserved by Excel)
        Fill defaultFill1 = new()
        {
            PatternFill = new PatternFill() { PatternType = PatternValues.None }
        };
        fills.Append(defaultFill1);

        // FillId = 1 - Default Fill (reserved by Excel)
        Fill defaultFill2 = new()
        {
            PatternFill = new PatternFill { PatternType = PatternValues.Gray125 }
        };
        fills.Append(defaultFill2);

        // FillId = 2 - U
        Fill uFill = new()
        {
            PatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFCC00") }
            }
        };
        fills.Append(uFill);

        // FillId = 3 - GZ
        Fill gzFill = new()
        {
            PatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FCE4D6") }
            }
        };
        fills.Append(gzFill);

        // FillId = 4 - HO
        Fill hoFill = new()
        {
            PatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("00CCFF") }
            }
        };
        fills.Append(hoFill);

        // FillId = 5 - D
        Fill dFill = new()
        {
            PatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("99CC00") }
            }
        };
        fills.Append(dFill);

        // FillId = 6 - AU
        Fill auFill = new()
        {
            PatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FF0000") }
            }
        };
        fills.Append(auFill);

        // FillId = 7 - SU
        Fill suFill = new()
        {
            PatternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FF00FF") }
            }
        };
        fills.Append(suFill);


        //FillId = 8 - GradientFill
        Fill myFill = new();
        GradientFill gradient = new()
        {
            Type = GradientValues.Linear
            //Type = GradientValues.Path,
            //Left = 0.5,
            //Bottom = 0.5,
            //Right = 0.5,
            //Top = 0.5
        };
        gradient.Append(new GradientStop()
        {
            Position = 0.0D,
            Color = new Color() { Rgb = HexBinaryValue.FromString("e50000") }
        });
        gradient.Append(new GradientStop()
        {
            Position = 0.2D,
            Color = new Color() { Rgb = HexBinaryValue.FromString("ff8d00") }
        });
        gradient.Append(new GradientStop()
        {
            Position = 0.4D,
            Color = new Color() { Rgb = HexBinaryValue.FromString("ffee00") }
        });
        gradient.Append(new GradientStop()
        {
            Position = 0.6D,
            Color = new Color() { Rgb = HexBinaryValue.FromString("028121") }
        });
        gradient.Append(new GradientStop()
        {
            Position = 0.8D,
            Color = new Color() { Rgb = HexBinaryValue.FromString("004cff") }
        });
        gradient.Append(new GradientStop()
        {
            Position = 1.0D,
            Color = new Color() { Rgb = HexBinaryValue.FromString("770088") }
        });
        myFill.GradientFill = gradient;
        fills.Append(myFill);

        // Get number of fills
        fills.Count = (uint)fills.ChildElements.Count;
        return fills;
    }

    public static Borders SetBorders()
    {
        Borders borders = new();

        // BorderId = 0 - 0 0 0 0
        Border borderId0 = new();
        borders.Append(borderId0);

        // BorderId = 1 - 0 0 M 0
        Border borderId1 = new()
        {
            LeftBorder = new LeftBorder { Style = BorderStyleValues.None },
            TopBorder = new TopBorder { Style = BorderStyleValues.None },
            RightBorder = new RightBorder { Style = BorderStyleValues.Medium },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.None }
        };
        borders.Append(borderId1);

        // BorderId = 2 - 0 M M M
        var borderId2 = new Border()
        {
            LeftBorder = new LeftBorder() { Style = BorderStyleValues.None },
            TopBorder = new TopBorder { Style = BorderStyleValues.Medium },
            RightBorder = new RightBorder { Style = BorderStyleValues.Medium },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.Medium }
        };
        borders.Append(borderId2);

        // BorderId = 3 - 0 M T M
        var borderId3 = new Border()
        {
            LeftBorder = new LeftBorder() { Style = BorderStyleValues.None },
            TopBorder = new TopBorder { Style = BorderStyleValues.Medium },
            RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.Medium }
        };
        borders.Append(borderId3);

        // BorderId = 4 - 0 0 M T
        var borderId4 = new Border()
        {
            LeftBorder = new LeftBorder { Style = BorderStyleValues.None },
            TopBorder = new TopBorder { Style = BorderStyleValues.None },
            RightBorder = new RightBorder { Style = BorderStyleValues.Medium },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin }
        };
        borders.Append(borderId4);

        // BorderId = 5 - 0 0 T T
        var borderId5 = new Border()
        {
            LeftBorder = new LeftBorder { Style = BorderStyleValues.None },
            TopBorder = new TopBorder { Style = BorderStyleValues.None },
            RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin }
        };
        borders.Append(borderId5);

        // BorderId = 6 - 0 0 M M
        var borderId6 = new Border()
        {
            LeftBorder = new LeftBorder { Style = BorderStyleValues.None },
            TopBorder = new TopBorder { Style = BorderStyleValues.None },
            RightBorder = new RightBorder { Style = BorderStyleValues.Medium },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.Medium }
        };
        borders.Append(borderId6);

        // BorderId = 7 - 0 0 T M
        var borderId7 = new Border()
        {
            LeftBorder = new LeftBorder { Style = BorderStyleValues.None },
            TopBorder = new TopBorder { Style = BorderStyleValues.None },
            RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.Medium }
        };
        borders.Append(borderId7);

        // BorderId = 8 - 0 0 T 0
        var borderId8 = new Border()
        {
            LeftBorder = new LeftBorder { Style = BorderStyleValues.None },
            TopBorder = new TopBorder { Style = BorderStyleValues.None },
            RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.None }
        };
        borders.Append(borderId8);

        // BorderId = 9 - 0 T T T
        var borderId9 = new Border()
        {
            LeftBorder = new LeftBorder { Style = BorderStyleValues.None },
            TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
            RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin }
        };
        borders.Append(borderId9);

        // Get count of Elements
        borders.Count = (uint)borders.ChildElements.Count;

        return borders;
    }

    public static CellFormats SetCellFormats()
    {
        CellFormats cellFormats = new();

        // CellStyle = 0 - Default style
        CellFormat style0 = new();
        cellFormats.Append(style0);

        // CellStyle = 1 - Title
        CellFormat style1 = new()
        {
            FontId = 2,
            FillId = 8, // Default 0 - LGBT 8
            BorderId = 0
        };
        cellFormats.Append(style1);

        // CellStyle = 2 - Default - Border 0 0 M 0
        CellFormat style2 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 1
        };
        cellFormats.Append(style2);

        // CellStyle = 3 - Bold - Alignment Left/Center - Border 0 M M M
        CellFormat style3 = new()
        {
            FontId = 1,
            FillId = 0,
            BorderId = 2,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style3);

        // CellStyle = 4 - Default - Alignment Center/Center - Border 0 M T M
        CellFormat style4 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 3,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style4);

        // CellStyle = 5 - Default - Alignment Center/Center - Border 0 M M M
        CellFormat style5 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 2,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style5);

        // CellStyle = 6 - U - Alignment Center/Center - Border 0 M T M
        CellFormat style6 = new()
        {
            FontId = 0,
            FillId = 2,
            BorderId = 3,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style6);

        // CellStyle = 7 - GZ - Alignment Center/Center - Border 0 M T M
        CellFormat style7 = new()
        {
            FontId = 0,
            FillId = 3,
            BorderId = 3,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style7);

        // CellStyle = 8 - HO - Alignment Center/Center - Border 0 M T M
        CellFormat style8 = new()
        {
            FontId = 0,
            FillId = 4,
            BorderId = 3,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style8);

        // CellStyle = 9 - D - Alignment Center/Center - Border 0 M T M
        CellFormat style9 = new()
        {
            FontId = 0,
            FillId = 5,
            BorderId = 3,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style9);

        // CellStyle = 10 - AU - Alignment Center/Center - Border 0 M T M
        CellFormat style10 = new()
        {
            FontId = 0,
            FillId = 6,
            BorderId = 3,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style10);

        // CellStyle = 11 - SU - Alignment Center/Center - Border 0 M T M
        CellFormat style11 = new()
        {
            FontId = 0,
            FillId = 7,
            BorderId = 3,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style11);

        // CellStyle = 12 - SU - Alignment Center/Center - Border 0 M T M
        CellFormat style12 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 3,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style12);

        // CellStyle = 13 - Default - Alignment Left/Center - Border 0 0 M T
        CellFormat style13 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 4,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style13);

        // CellStyle = 14 - Default - Alignment Center/Center - Border 0 0 M T
        CellFormat style14 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 4,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style14);

        // CellStyle = 15 - Default - Alignment Center/Center - Border 0 0 T T
        CellFormat style15 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 5,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style15);

        // CellStyle = 16 - Default - Alignment Left/Center - Border 0 0 M M
        CellFormat style16 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 6,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style16);

        // CellStyle = 17 - Default - Alignment Center/Center - Border 0 0 M M
        CellFormat style17 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 6,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style17);

        // CellStyle = 18 - Default - Alignment Center/Center - Border 0 0 T M
        CellFormat style18 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 7,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style18);

        // CellStyle = 19 - Default - Alignment Left/Center - Border 0 0 T 0
        CellFormat style19 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 8,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style19);

        // CellStyle = 20 - Bold - Alignment Left/Center - Border 0 T T T
        CellFormat style20 = new()
        {
            FontId = 1,
            FillId = 0,
            BorderId = 9,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style20);

        // CellStyle = 21 - Bold - Alignment Center/Center - Border 0 T T T
        CellFormat style21 = new()
        {
            FontId = 1,
            FillId = 0,
            BorderId = 9,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style21);

        // CellStyle = 22 - Bold (Month) - Alignment Left/Center - Border 0 T T T
        CellFormat style22 = new()
        {
            FontId = 1,
            FillId = 0,
            BorderId = 9,
            NumberFormatId = 164,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style22);

        // CellStyle = 23 - Bold - Alignment Center/Center - Border 0 T T T
        CellFormat style23 = new()
        {
            FontId = 1,
            FillId = 0,
            BorderId = 9,
            NumberFormatId = 165,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style23);

        // CellStyle = 24 - Default - Alignment Left/Center - Border 0 0 T T
        CellFormat style24 = new()
        {
            FontId = 0,
            FillId = 0,
            BorderId = 5,
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            }
        };
        cellFormats.Append(style24);


        // Get number of CellFormats
        cellFormats.Count = (uint)cellFormats.ChildElements.Count;

        return cellFormats;
    }

    public static CellStyleFormats SetCellStyleFormats()
    {
        throw new NotImplementedException();
    }

    public static CellStyles SetCellStyles()
    {
        throw new NotImplementedException();
    }

    public static DifferentialFormats SetDifferentialFormats()
    {
        DifferentialFormats differentialFormats = new();

        // FormatID = 0 - Current Day
        DifferentialFormat currentDayFormat = new()
        {
            Fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    BackgroundColor = new BackgroundColor()
                        { Rgb = HexBinaryValue.FromString("FFFF00") }
                }
            }
        };
        differentialFormats.Append(currentDayFormat);

        // FormatID = 1 - Weekday
        DifferentialFormat weekDayFormat = new()
        {
            Fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    BackgroundColor = new BackgroundColor()
                    {
                        Rgb = HexBinaryValue.FromString("A5A5A5")
                    }
                }
            }
        };
        differentialFormats.Append(weekDayFormat);

        // FormatId = 2 - U
        DifferentialFormat uFormat = new()
        {
            Fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    BackgroundColor = new BackgroundColor
                        { Rgb = HexBinaryValue.FromString("FFCC00") }
                }
            }
        };
        differentialFormats.Append(uFormat);

        // FormatId = 3 - GZ
        DifferentialFormat gzFormat = new()
        {
            Fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    BackgroundColor = new BackgroundColor
                        { Rgb = HexBinaryValue.FromString("FCE4D6") }
                }
            }
        };
        differentialFormats.Append(gzFormat);

        // FormatId = 4 - HO
        DifferentialFormat hoFormat = new()
        {
            Fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    BackgroundColor = new BackgroundColor
                        { Rgb = HexBinaryValue.FromString("00CCFF") }
                }
            }
        };
        differentialFormats.Append(hoFormat);

        // FormatId = 5 - D
        DifferentialFormat dFormat = new()
        {
            Fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    BackgroundColor = new BackgroundColor
                        { Rgb = HexBinaryValue.FromString("99CC00") }
                }
            }
        };
        differentialFormats.Append(dFormat);

        // FormatId = 6 - AU
        DifferentialFormat auFormat = new()
        {
            Fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    BackgroundColor = new BackgroundColor
                        { Rgb = HexBinaryValue.FromString("FF0000") }
                }
            }
        };
        differentialFormats.Append(auFormat);

        // FormatId = 7 - SU
        DifferentialFormat suFormat = new()
        {
            Fill = new Fill()
            {
                PatternFill = new PatternFill()
                {
                    BackgroundColor = new BackgroundColor
                        { Rgb = HexBinaryValue.FromString("FF00FF") }
                }
            }
        };
        differentialFormats.Append(suFormat);

        // FormatId = 8 - Holidays
        DifferentialFormat holidaysFormat = new()
        {
            Fill = new Fill()
            {
                PatternFill = new PatternFill
                {
                    BackgroundColor = new BackgroundColor
                        { Rgb = HexBinaryValue.FromString("CCFFCC") }
                }
            }
        };
        differentialFormats.Append(holidaysFormat);

        // FormatId = 9 - School holidays
        DifferentialFormat schoolHolidaysFormat = new()
        {
            Fill = new Fill
            {
                PatternFill = new PatternFill
                {
                    BackgroundColor = new BackgroundColor
                        { Rgb = HexBinaryValue.FromString("CCCCFF") }
                }
            }
        };
        differentialFormats.Append(schoolHolidaysFormat);

        // FormatId = 10 - Cross Holidays
        DifferentialFormat crossHolidaysFormat = new()
        {
            Fill = new Fill()
            {
                GradientFill = new GradientFill
                (
                    new GradientStop()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("CCFFCC") },
                        Position = 0.0
                    },
                    new GradientStop()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("CCCCFF") },
                        Position = 1.0
                    }
                )
            }
        };
        differentialFormats.Append(crossHolidaysFormat);

        // Get number of DifferntialFormats
        differentialFormats.Count = (uint)differentialFormats.ChildElements.Count;

        return differentialFormats;
    }

    public static NumberingFormats NumberingFormats()
    {
        var numberingFormats = new NumberingFormats();

        // NumberingFormatId 164 - Month
        NumberingFormat monthFormat = new()
        {
            NumberFormatId = 164,
            FormatCode = StringValue.FromString("mmmm")
        };
        numberingFormats.Append(monthFormat);

        // NumberingFormatId 165 - Day
        NumberingFormat dayFormat = new()
        {
            NumberFormatId = 165,
            FormatCode = StringValue.FromString("dd")
        };
        numberingFormats.Append(dayFormat);


        return numberingFormats;
    }
}