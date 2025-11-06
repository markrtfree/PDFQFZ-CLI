using System.Text;

var (options, showHelp) = global::PDFQFZ.CLI.CommandLineParser.Parse(args);
if (showHelp)
{
    PrintUsage();
    return;
}

try
{
    using var cts = new CancellationTokenSource();
    Console.CancelKeyPress += (_, eventArgs) =>
    {
        Console.WriteLine("Cancellation requested...");
        eventArgs.Cancel = true;
        cts.Cancel();
    };

    using var processor = new global::PDFQFZ.CLI.StampProcessor(options);
    var processed = processor.Execute(cts.Token);
    Console.WriteLine($"Completed stamping for {processed} file(s).");
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    Environment.ExitCode = 1;
}

static void PrintUsage()
{
    var builder = new StringBuilder();
    builder.AppendLine("PDFQFZ.CLI - Apply seam and page stamps to PDF files using command-line options.");
    builder.AppendLine();
    builder.AppendLine("Usage:");
    builder.AppendLine("  pdfqfz-cli --input <path>... --stamp-image <image.png> [options]");
    builder.AppendLine();
    builder.AppendLine("Required arguments:");
    builder.AppendLine("  -i, --input <path>           PDF file or directory to process (can be repeated).");
    builder.AppendLine("  -s, --stamp-image <path>     Image used as the stamp (PNG recommended).");
    builder.AppendLine();
    builder.AppendLine("Output options:");
    builder.AppendLine("      --output <dir>           Write stamped PDFs to this directory instead of replacing inputs.");
    builder.AppendLine("      --overwrite              Overwrite existing files in the output directory.");
    builder.AppendLine("      --suffix <text>          Suffix appended to output file names when using --output (default: _stamped).");
    builder.AppendLine();
    builder.AppendLine("General options:");
    builder.AppendLine("      --recursive              Process PDFs in subdirectories when input is a folder.");
    builder.AppendLine("      --size-mm <value>        Target stamp width in millimetres (default: 40).");
    builder.AppendLine("      --rotation <degrees>     Rotate stamp image before stamping (default: 0).");
    builder.AppendLine("      --rotate-crop            Crop rotated stamp instead of keeping bounds.");
    builder.AppendLine("      --opacity <0-100>        Overall stamp opacity (default: 100).");
    builder.AppendLine("      --no-white-transparent   Do not convert white pixels to transparent.");
    builder.AppendLine("      --input-password <pwd>   Password for opening encrypted PDFs.");
    builder.AppendLine("      --encrypt-password <pwd> Apply password protection to stamped PDFs.");
    builder.AppendLine();
    builder.AppendLine("Seam stamp options:");
    builder.AppendLine("      --seam-scope <mode>      Seam mode: none, all, odd, even, custom.");
    builder.AppendLine("      --seam-pages <list>      Comma-separated pages for custom seam scope.");
    builder.AppendLine("      --seam-side <edge>       Edge for seam: left, right, top, bottom (default: right).");
    builder.AppendLine("      --seam-offset <percent>  Offset along the edge (default: 50).");
    builder.AppendLine("      --seam-max-slices <n>    Maximum seam slices per processing batch (default: 20).");
    builder.AppendLine();
    builder.AppendLine("Page stamp options:");
    builder.AppendLine("      --page-scope <mode>      Page mode: none, all, skip-first, skip-last, custom.");
    builder.AppendLine("      --page-list <list>       Pages for custom page scope.");
    builder.AppendLine("      --position <x,y>         Relative position (0-1) from left/bottom (default: 0.5,0.5).");
    builder.AppendLine("      --position-map <spec>    Per-page positions: e.g. 1@0.5,0.2;5@0.6,0.3.");
    builder.AppendLine("      --random-offset          Apply small random offsets and rotation per page.");
    builder.AppendLine();
    builder.AppendLine("Digital signature options:");
    builder.AppendLine("      --sign-mode <mode>       none, self, pfx.");
    builder.AppendLine("      --sign-pfx <path>        Path to PFX for sign-mode pfx.");
    builder.AppendLine("      --sign-pass <pwd>        Password for PFX or generated certificate.");
    builder.AppendLine("      --sign-subject <text>    Subject for self-signed certificate (default: CN=PDFQFZ CLI).");
    builder.AppendLine();
    builder.AppendLine("Examples:");
    builder.AppendLine("  pdfqfz-cli -i input.pdf -o output --stamp-image seam.png --seam-scope all --page-scope all");
    builder.AppendLine("  pdfqfz-cli -i C:/docs -o stamped --stamp-image seal.png --recursive --page-scope skip-first");
    builder.AppendLine();
    Console.WriteLine(builder.ToString());
}
