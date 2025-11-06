using System.Globalization;

namespace PDFQFZ.CLI;

/// <summary>
/// Location of the seam (骑缝章) relative to the page edges.
/// </summary>
internal enum SeamSide
{
    Left,
    Right,
    Top,
    Bottom,
}

/// <summary>
/// Defines which pages participate in seam stamping.
/// </summary>
internal enum SeamScope
{
    None,
    All,
    Odd,
    Even,
    Custom,
}

/// <summary>
/// Defines which pages receive the standard page stamp.
/// </summary>
internal enum PageStampScope
{
    None,
    SkipFirst,
    SkipLast,
    All,
    Custom,
}

/// <summary>
/// Defines how the tool should provision cryptographic credentials.
/// </summary>
internal enum SignatureMode
{
    None,
    SelfSigned,
    CustomCertificate,
}

internal sealed record SeamStampOptions
{
    public bool Enabled => Scope != SeamScope.None;
    public SeamScope Scope { get; init; } = SeamScope.None;
    public List<int> CustomPages { get; } = new();
    public SeamSide Side { get; init; } = SeamSide.Right;
    /// <summary>
    /// Offset percentage (0-100) along the edge where the seam stamp should be placed.
    /// </summary>
    public float EdgeOffsetPercent { get; init; } = 50f;
    /// <summary>
    /// Maximum number of slices per batch when splitting the seam image.
    /// </summary>
    public int MaxSlicesPerBatch { get; init; } = 20;
}

internal sealed record PageStampOptions
{
    public bool Enabled => Scope != PageStampScope.None;
    public PageStampScope Scope { get; init; } = PageStampScope.None;
    public List<int> CustomPages { get; } = new();
    /// <summary>
    /// Default relative X position (0-1, measured from left).
    /// </summary>
    public float PositionX { get; init; } = 0.5f;
    /// <summary>
    /// Default relative Y position (0-1, measured from bottom).
    /// </summary>
    public float PositionY { get; init; } = 0.5f;
    public Dictionary<int, (float X, float Y)> PerPagePositions { get; } = new();
    public bool RandomizePerPage { get; init; } = false;
}

internal sealed record SignatureOptions
{
    public SignatureMode Mode { get; init; } = SignatureMode.None;
    public string? CertificatePath { get; init; }
    public string? Password { get; init; }
    public string SelfSignedSubject { get; init; } = "CN=PDFQFZ CLI";
}

internal sealed record StampingOptions
{
    public List<string> Inputs { get; } = new();
    public string? OutputDirectory { get; init; }
    public bool ProcessInPlace => string.IsNullOrWhiteSpace(OutputDirectory);
    public bool IncludeSubdirectories { get; init; } = false;
    public bool Overwrite { get; init; } = false;
    public string OutputSuffix { get; init; } = "_stamped";
    public string StampImagePath { get; init; } = "";
    public bool ConvertWhiteToTransparent { get; init; } = true;
    public float StampWidthMillimetres { get; init; } = 40f;
    public int StampRotationDegrees { get; init; } = 0;
    public bool KeepOriginalBoundsWhenRotating { get; init; } = true;
    public int StampOpacity { get; init; } = 100;
    public bool FlattenToImages { get; init; } = false;
    public string? InputPassword { get; init; }
    public string? OutputPassword { get; init; }
    public SeamStampOptions Seam { get; init; } = new();
    public PageStampOptions PageStamp { get; init; } = new();
    public SignatureOptions Signature { get; init; } = new();
}

internal static class OptionParsers
{
    public static IEnumerable<int> ParsePageList(string value)
    {
        foreach (var segment in value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!int.TryParse(segment, NumberStyles.Integer, CultureInfo.InvariantCulture, out var page) || page < 1)
            {
                throw new ArgumentException($"Invalid page number '{segment}'. Page numbers must be >= 1.");
            }

            yield return page;
        }
    }

    /// <summary>
    /// Parses strings in the format "page@x,y;page@x,y" to a dictionary.
    /// </summary>
    public static Dictionary<int, (float X, float Y)> ParsePositionMap(string value)
    {
        var map = new Dictionary<int, (float X, float Y)>();
        var entries = value.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var entry in entries)
        {
            var parts = entry.Split('@');
            if (parts.Length != 2)
            {
                throw new ArgumentException($"Invalid custom position entry '{entry}', expected format 'page@x,y'.");
            }

            if (!int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var page) || page < 1)
            {
                throw new ArgumentException($"Invalid page number '{parts[0]}' in custom position entry '{entry}'.");
            }

            var coords = parts[1].Split(',');
            if (coords.Length != 2 ||
                !float.TryParse(coords[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var x) ||
                !float.TryParse(coords[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var y))
            {
                throw new ArgumentException($"Invalid coordinates '{parts[1]}' in custom position entry '{entry}'.");
            }

            ValidateRatio(x, nameof(x));
            ValidateRatio(y, nameof(y));

            map[page] = (x, y);
        }

        return map;
    }

    public static (float X, float Y) ParsePosition(string value)
    {
        var parts = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length != 2 ||
            !float.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var x) ||
            !float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var y))
        {
            throw new ArgumentException($"Invalid position '{value}'. Expected 'x,y' with values between 0 and 1.");
        }

        ValidateRatio(x, nameof(x));
        ValidateRatio(y, nameof(y));
        return (x, y);
    }

    public static float ParsePercentage(string value)
    {
        if (!float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
        {
            throw new ArgumentException($"Invalid percentage value '{value}'.");
        }

        if (result < 0 || result > 100)
        {
            throw new ArgumentOutOfRangeException(nameof(value), "Percentage must be between 0 and 100.");
        }

        return result;
    }

    public static float ParseRatio(string value)
    {
        if (!float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
        {
            throw new ArgumentException($"Invalid ratio '{value}'.");
        }

        ValidateRatio(result, nameof(value));
        return result;
    }

    private static void ValidateRatio(float value, string name)
    {
        if (value < 0f || value > 1f)
        {
            throw new ArgumentOutOfRangeException(name, "Value must be between 0 and 1.");
        }
    }
}
