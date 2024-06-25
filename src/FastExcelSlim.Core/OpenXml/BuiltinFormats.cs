namespace FastExcelSlim.OpenXml;

public static class BuiltinFormats
{
    public const int FirstUserDefinedFormatIndex = 164;

    public const string DefaultDateTimeFormat = "m/d/yy h:mm";

    private static readonly string[] Formats =
    [
        "General",
        "0",
        "0.00",
        "#,##0",
        "#,##0.00",
        "\"$\"#,##0_);(\"$\"#,##0)",
        "\"$\"#,##0_);[Red](\"$\"#,##0)",
        "\"$\"#,##0.00_);(\"$\"#,##0.00)",
        "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        "0%",
        "0.00%",
        "0.00E+00",
        "# ?/?",
        "# ??/??",
        "m/d/yy",
        "d-mmm-yy",
        "d-mmm",
        "mmm-yy",
        "h:mm AM/PM",
        "h:mm:ss AM/PM",
        "h:mm",
        "h:mm:ss",
        DefaultDateTimeFormat,

        // 0x17 - 0x24 reserved for international and undocumented
        // TODO - one junit relies on these values which seems incorrect
        "reserved-0x17",
        "reserved-0x18",
        "reserved-0x19",
        "reserved-0x1A",
        "reserved-0x1B",
        "reserved-0x1C",
        "reserved-0x1D",
        "reserved-0x1E",
        "reserved-0x1F",
        "reserved-0x20",
        "reserved-0x21",
        "reserved-0x22",
        "reserved-0x23",
        "reserved-0x24",

        "#,##0_);(#,##0)",
        "#,##0_);[Red](#,##0)",
        "#,##0.00_);(#,##0.00)",
        "#,##0.00_);[Red](#,##0.00)",
        "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
        "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
        "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
        "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
        "mm:ss",
        "[h]:mm:ss",
        "mm:ss.0",
        "##0.0E+0",
        "@"
    ];

    public static string? GetBuiltinFormat(int index)
    {
        if (index < 0 || index >= Formats.Length)
        {
            return null;
        }
        return Formats[index];
    }

    public static int GetBuiltinFormat(string pFmt)
    {
        var fmt = "TEXT".Equals(pFmt, StringComparison.OrdinalIgnoreCase) ? "@" : pFmt;

        int i = -1;
        foreach (var f in Formats)
        {
            i++;
            if (f.Equals(fmt))
            {
                return i;
            }
        }

        return -1;
    }
}