using System.Globalization;

namespace PDFQFZ.CLI;

internal static class CommandLineParser
{
    public static (StampingOptions Options, bool ShowHelp) Parse(string[] args)
    {
        if (args.Length == 0)
        {
            return (CreateDefaultOptions(), true);
        }

        var inputs = new List<string>();
        string? outputDir = null;
        string? stampImage = null;
        string suffix = "_stamped";
        bool recursive = false;
        bool overwrite = false;
        bool convertWhite = true;
        float sizeMm = 40f;
        int rotation = 0;
        bool keepBounds = true;
        int opacity = 100;
        bool randomOffset = false;
        string? inputPassword = null;
        string? outputPassword = null;

        var seam = new SeamStampOptions();
        var seamScope = seam.Scope;
        var seamCustomPages = new List<int>();
        var seamSide = seam.Side;
        float seamOffset = seam.EdgeOffsetPercent;
        int seamMaxSlices = seam.MaxSlicesPerBatch;

        var pageStampScope = PageStampScope.None;
        var pageCustomPages = new List<int>();
        (float X, float Y) position = (0.5f, 0.5f);
        Dictionary<int, (float X, float Y)> perPagePositions = new();

        var signatureMode = SignatureMode.None;
        string? pfxPath = null;
        string? signPassword = null;
        string signSubject = "CN=PDFQFZ CLI";

        var showHelp = false;

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            switch (arg)
            {
                case "-h":
                case "--help":
                    showHelp = true;
                    break;
                case "-i":
                case "--input":
                    inputs.Add(RequireValue(args, ref i, arg));
                    break;
                case "-o":
                case "--output":
                    outputDir = RequireValue(args, ref i, arg);
                    break;
                case "-s":
                case "--stamp-image":
                    stampImage = RequireValue(args, ref i, arg);
                    break;
                case "--recursive":
                    recursive = true;
                    break;
                case "--overwrite":
                    overwrite = true;
                    break;
                case "--suffix":
                    suffix = RequireValue(args, ref i, arg);
                    break;
                case "--no-white-transparent":
                    convertWhite = false;
                    break;
                case "--size-mm":
                    sizeMm = ParseFloat(RequireValue(args, ref i, arg), arg);
                    break;
                case "--rotation":
                    rotation = ParseInt(RequireValue(args, ref i, arg), arg);
                    break;
                case "--rotate-crop":
                    keepBounds = false;
                    break;
                case "--opacity":
                    opacity = Math.Clamp(ParseInt(RequireValue(args, ref i, arg), arg), 0, 100);
                    break;
                case "--seam-scope":
                    seamScope = ParseSeamScope(RequireValue(args, ref i, arg));
                    break;
                case "--seam-pages":
                    seamCustomPages.AddRange(OptionParsers.ParsePageList(RequireValue(args, ref i, arg)));
                    break;
                case "--seam-side":
                    seamSide = ParseSeamSide(RequireValue(args, ref i, arg));
                    break;
                case "--seam-offset":
                    seamOffset = OptionParsers.ParsePercentage(RequireValue(args, ref i, arg));
                    break;
                case "--seam-max-slices":
                    seamMaxSlices = Math.Max(1, ParseInt(RequireValue(args, ref i, arg), arg));
                    break;
                case "--page-scope":
                    pageStampScope = ParsePageScope(RequireValue(args, ref i, arg));
                    break;
                case "--page-list":
                    pageCustomPages.AddRange(OptionParsers.ParsePageList(RequireValue(args, ref i, arg)));
                    break;
                case "--position":
                    position = OptionParsers.ParsePosition(RequireValue(args, ref i, arg));
                    break;
                case "--position-map":
                    perPagePositions = OptionParsers.ParsePositionMap(RequireValue(args, ref i, arg));
                    break;
                case "--random-offset":
                    randomOffset = true;
                    break;
                case "--input-password":
                    inputPassword = RequireValue(args, ref i, arg);
                    break;
                case "--encrypt-password":
                    outputPassword = RequireValue(args, ref i, arg);
                    break;
                case "--sign-mode":
                    signatureMode = ParseSignatureMode(RequireValue(args, ref i, arg));
                    break;
                case "--sign-pfx":
                    pfxPath = RequireValue(args, ref i, arg);
                    break;
                case "--sign-pass":
                    signPassword = RequireValue(args, ref i, arg);
                    break;
                case "--sign-subject":
                    signSubject = RequireValue(args, ref i, arg);
                    break;
                default:
                    if (arg.StartsWith("-"))
                    {
                        throw new ArgumentException($"Unknown option '{arg}'. Use --help to see available options.");
                    }

                    inputs.Add(arg);
                    break;
            }
        }

        if (showHelp)
        {
            return (CreateDefaultOptions(), true);
        }

        if (inputs.Count == 0)
        {
            throw new ArgumentException("At least one --input argument or positional input path is required.");
        }

        if (string.IsNullOrWhiteSpace(outputDir))
        {
            throw new ArgumentException("--output must be specified.");
        }

        if (string.IsNullOrWhiteSpace(stampImage))
        {
            throw new ArgumentException("--stamp-image must be specified.");
        }

        var seamOptions = new SeamStampOptions
        {
            Scope = seamScope,
            Side = seamSide,
            EdgeOffsetPercent = seamOffset,
            MaxSlicesPerBatch = seamMaxSlices
        };

        if (seamScope == SeamScope.Custom && seamCustomPages.Count == 0)
        {
            throw new ArgumentException("--seam-pages must be provided when seam scope is custom.");
        }
        seamOptions.CustomPages.AddRange(seamCustomPages);

        var pageOptions = new PageStampOptions
        {
            Scope = pageStampScope,
            PositionX = position.X,
            PositionY = position.Y,
            RandomizePerPage = randomOffset
        };
        pageOptions.CustomPages.AddRange(pageCustomPages);
        foreach (var kvp in perPagePositions)
        {
            pageOptions.PerPagePositions[kvp.Key] = kvp.Value;
        }

        var signatureOptions = new SignatureOptions
        {
            Mode = signatureMode,
            CertificatePath = pfxPath,
            Password = signPassword,
            SelfSignedSubject = signSubject
        };

        var stampingOptions = new StampingOptions
        {
            OutputDirectory = string.IsNullOrWhiteSpace(outputDir) ? null : outputDir,
            IncludeSubdirectories = recursive,
            Overwrite = overwrite,
            OutputSuffix = suffix,
            StampImagePath = stampImage!,
            ConvertWhiteToTransparent = convertWhite,
            StampWidthMillimetres = sizeMm,
            StampRotationDegrees = rotation,
            KeepOriginalBoundsWhenRotating = keepBounds,
            StampOpacity = opacity,
            InputPassword = inputPassword,
            OutputPassword = outputPassword,
            Seam = seamOptions,
            PageStamp = pageOptions,
            Signature = signatureOptions
        };

        stampingOptions.Inputs.AddRange(inputs);

        return (stampingOptions, false);
    }

    private static float ParseFloat(string value, string optionName)
    {
        if (!float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
        {
            throw new ArgumentException($"Option {optionName} expects a floating-point value.");
        }

        return result;
    }

    private static int ParseInt(string value, string optionName)
    {
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result))
        {
            throw new ArgumentException($"Option {optionName} expects an integer value.");
        }

        return result;
    }

    private static SeamScope ParseSeamScope(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "none" => SeamScope.None,
            "all" => SeamScope.All,
            "odd" => SeamScope.Odd,
            "even" => SeamScope.Even,
            "custom" => SeamScope.Custom,
            _ => throw new ArgumentException($"Unknown seam scope '{value}'. Valid values: none, all, odd, even, custom.")
        };
    }

    private static SeamSide ParseSeamSide(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "left" => SeamSide.Left,
            "right" => SeamSide.Right,
            "top" => SeamSide.Top,
            "bottom" => SeamSide.Bottom,
            _ => throw new ArgumentException($"Unknown seam side '{value}'. Valid values: left, right, top, bottom.")
        };
    }

    private static PageStampScope ParsePageScope(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "none" => PageStampScope.None,
            "all" => PageStampScope.All,
            "skip-first" => PageStampScope.SkipFirst,
            "skip-last" => PageStampScope.SkipLast,
            "custom" => PageStampScope.Custom,
            _ => throw new ArgumentException($"Unknown page scope '{value}'. Valid values: none, all, skip-first, skip-last, custom.")
        };
    }

    private static SignatureMode ParseSignatureMode(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "none" => SignatureMode.None,
            "self" => SignatureMode.SelfSigned,
            "selfsigned" => SignatureMode.SelfSigned,
            "pfx" => SignatureMode.CustomCertificate,
            "custom" => SignatureMode.CustomCertificate,
            _ => throw new ArgumentException($"Unknown sign mode '{value}'. Valid values: none, self, pfx.")
        };
    }

    private static string RequireValue(string[] args, ref int index, string optionName)
    {
        if (index + 1 >= args.Length)
        {
            throw new ArgumentException($"Option {optionName} expects a value.");
        }

        index++;
        return args[index];
    }

    private static StampingOptions CreateDefaultOptions() => new();
}
