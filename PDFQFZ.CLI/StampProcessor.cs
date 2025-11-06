using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.security;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using DrawingRectangle = System.Drawing.Rectangle;
using PdfImage = iTextSharp.text.Image;
using BcX509Certificate = Org.BouncyCastle.X509.X509Certificate;

namespace PDFQFZ.CLI;

internal sealed class StampProcessor : IDisposable
{
    private readonly StampingOptions _options;
    private readonly List<string> _inputFiles;
    private readonly Bitmap _stampBitmap;
    private readonly int _originalStampWidthPixels;
    private readonly float _scalePercent;
    private readonly SignatureContext? _signatureContext;
    private readonly Random _random = new();
    private bool _disposed;

    public StampProcessor(StampingOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        ValidateOptions(_options);

        _inputFiles = ResolveInputFiles(_options);

        (_stampBitmap, _originalStampWidthPixels) = PrepareStampBitmap(_options);
        _scalePercent = CalculateScalePercent(_options.StampWidthMillimetres, _originalStampWidthPixels);
        _signatureContext = PrepareSignatureContext(_options.Signature);
    }

    public int Execute(CancellationToken cancellationToken)
    {
        ThrowIfDisposed();
        if (_inputFiles.Count == 0)
        {
            Console.WriteLine("No PDF files matched the provided input paths.");
            return 0;
        }

        if (!_options.ProcessInPlace)
        {
            Directory.CreateDirectory(_options.OutputDirectory!);
        }

        var processed = 0;
        foreach (var inputPath in _inputFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (_options.ProcessInPlace)
            {
                Console.WriteLine($"Stamping '{inputPath}' (in-place).");
                StampInPlace(inputPath, cancellationToken);
            }
            else
            {
                var outputPath = ComposeOutputPath(inputPath);
                EnsureOutputPathValid(inputPath, outputPath, _options.Overwrite);

                Console.WriteLine($"Stamping '{inputPath}' -> '{outputPath}'");
                ProcessSingleFile(inputPath, outputPath, cancellationToken);
            }
            processed++;
        }

        return processed;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _stampBitmap.Dispose();
        _signatureContext?.Dispose();
        _disposed = true;
    }

    private void ProcessSingleFile(string inputPath, string outputPath, CancellationToken cancellationToken)
    {
        using var outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        var passwordBytes = string.IsNullOrEmpty(_options.InputPassword)
            ? null
            : Encoding.UTF8.GetBytes(_options.InputPassword);

        if (_options.FlattenToImages)
        {
            Console.WriteLine("Flatten-to-images mode is not yet implemented; continuing without flattening.");
        }

        var signatureRequired = _signatureContext is not null;

        PdfReader? reader = null;
        PdfStamper? stamper = null;
        var signatureCompleted = false;
        try
        {
            reader = passwordBytes is null
                ? new PdfReader(inputPath)
                : new PdfReader(inputPath, passwordBytes);

            stamper = signatureRequired
                ? PdfStamper.CreateSignature(reader, outputStream, '\0', null, true)
                : new PdfStamper(reader, outputStream);
            stamper.Writer.CloseStream = true;

            if (!string.IsNullOrEmpty(_options.OutputPassword))
            {
                if (signatureRequired)
                {
                    throw new NotSupportedException("Output password protection cannot be combined with digital signatures in the CLI version.");
                }

                var pwd = Encoding.UTF8.GetBytes(_options.OutputPassword);
                stamper.SetEncryption(pwd, pwd, PdfWriter.ALLOW_PRINTING | PdfWriter.ALLOW_COPY, PdfWriter.ENCRYPTION_AES_128);
            }

            var totalPages = reader.NumberOfPages;
            if (_options.Seam.Enabled)
            {
                ApplySeamStamp(reader, stamper, totalPages, cancellationToken);
            }

            if (_options.PageStamp.Enabled)
            {
                ApplyPageStamp(reader, stamper, totalPages, cancellationToken);
            }

            if (signatureRequired)
            {
                ApplyDigitalSignature(stamper, reader);
                signatureCompleted = true;
            }
        }
        finally
        {
            if (!signatureCompleted)
            {
                stamper?.Close();
            }

            reader?.Close();
        }
    }

    private void StampInPlace(string inputPath, CancellationToken cancellationToken)
    {
        var tempPath = CreateTemporarySiblingPath(inputPath);
        try
        {
            ProcessSingleFile(inputPath, tempPath, cancellationToken);
            CommitInPlaceReplacement(inputPath, tempPath);
        }
        catch
        {
            SafeDelete(tempPath);
            throw;
        }
    }

    private void ApplySeamStamp(PdfReader reader, PdfStamper stamper, int totalPages, CancellationToken cancellationToken)
    {
        var seamPages = ResolveSeamPages(_options.Seam, totalPages);
        if (seamPages.Count < 2)
        {
            if (seamPages.Count == 1)
            {
                Console.WriteLine("Skipping seam stamping: requires at least 2 pages.");
            }
            return;
        }

        var slicePlan = BuildSeamSlicePlan(seamPages.Count);
        var maxBatch = Math.Max(1, _options.Seam.MaxSlicesPerBatch);
        for (var batchStart = 0; batchStart < seamPages.Count; batchStart += maxBatch)
        {
            var batchEnd = Math.Min(batchStart + maxBatch, seamPages.Count);
            for (var index = batchStart; index < batchEnd; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var pageNumber = seamPages[index];
                if (pageNumber < 1 || pageNumber > totalPages)
                {
                    continue;
                }

                var (startX, width) = slicePlan[index];
                if (width <= 0)
                {
                    continue;
                }

                using var sliceBitmap = ExtractSlice(_stampBitmap, startX, width);
                Bitmap? rotatedSlice = null;
                try
                {
                    var workingSlice = sliceBitmap;
                    if (_options.Seam.Side is SeamSide.Top or SeamSide.Bottom)
                    {
                        rotatedSlice = ImageUtilities.Rotate(sliceBitmap, 90, keepBounds: false);
                        workingSlice = rotatedSlice;
                    }

                    var image = PdfImage.GetInstance(workingSlice, ImageFormat.Png);
                    image.ScalePercent(_scalePercent);
                    var scaledWidth = image.ScaledWidth;
                    var scaledHeight = image.ScaledHeight;

                    var canvas = stamper.GetOverContent(pageNumber);
                    var (pageWidth, pageHeight, _) = GetPageMetrics(reader, pageNumber);
                    var position = ComputeSeamPosition(_options.Seam.Side, _options.Seam.EdgeOffsetPercent, pageWidth, pageHeight, scaledWidth, scaledHeight);

                    image.SetAbsolutePosition(position.X, position.Y);
                    canvas.AddImage(image);
                }
                finally
                {
                    rotatedSlice?.Dispose();
                }
            }
        }
    }

    private void ApplyPageStamp(PdfReader reader, PdfStamper stamper, int totalPages, CancellationToken cancellationToken)
    {
        var targetPages = ResolvePageStampPages(_options.PageStamp, totalPages);
        if (targetPages.Count == 0)
        {
            return;
        }

        foreach (var pageNumber in targetPages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (pageNumber < 1 || pageNumber > totalPages)
            {
                continue;
            }

            var canvas = stamper.GetOverContent(pageNumber);
            var (pageWidth, pageHeight, _) = GetPageMetrics(reader, pageNumber);

            var (posX, posY) = ResolvePagePosition(pageNumber, _options.PageStamp);
            var (adjustedX, adjustedY, rotationAdjustment) = ApplyRandomOffsetIfNeeded(posX, posY, _options.PageStamp.RandomizePerPage);

            var image = PdfImage.GetInstance(_stampBitmap, ImageFormat.Png);
            image.ScalePercent(_scalePercent);
            image.RotationDegrees = rotationAdjustment;

            var scaledWidth = image.ScaledWidth;
            var scaledHeight = image.ScaledHeight;

            var absolutePosition = ComputePagePosition(adjustedX, adjustedY, pageWidth, pageHeight, scaledWidth, scaledHeight);
            image.SetAbsolutePosition(absolutePosition.X, absolutePosition.Y);
            canvas.AddImage(image);
        }
    }

    private void ApplyDigitalSignature(PdfStamper stamper, PdfReader reader)
    {
        if (_signatureContext is null)
        {
            return;
        }

        var appearance = stamper.SignatureAppearance;
        appearance.SignDate = DateTime.Now;
        appearance.CertificationLevel = PdfSignatureAppearance.NOT_CERTIFIED;
        appearance.SignatureCreator = "PDFQFZ CLI";
        appearance.SetVisibleSignature(new iTextSharp.text.Rectangle(0, 0, 0, 0), reader.NumberOfPages, null);

        var signature = new PrivateKeySignature(_signatureContext.PrivateKey, DigestAlgorithms.SHA256);
        MakeSignature.SignDetached(appearance, signature, _signatureContext.Chain, null, null, null, 0, CryptoStandard.CMS);
    }

    private static (float X, float Y) ComputePagePosition(float ratioX, float ratioY, float pageWidth, float pageHeight, float imageWidth, float imageHeight)
    {
        ratioX = Math.Clamp(ratioX, 0f, 1f);
        ratioY = Math.Clamp(ratioY, 0f, 1f);

        var x = (pageWidth - imageWidth) * ratioX;
        var y = (pageHeight - imageHeight) * ratioY;
        return (x, y);
    }

    private static (float X, float Y) ComputeSeamPosition(SeamSide side, float edgeOffsetPercent, float pageWidth, float pageHeight, float imageWidth, float imageHeight)
    {
        var ratio = Math.Clamp(edgeOffsetPercent / 100f, 0f, 1f);

        return side switch
        {
            SeamSide.Left => (0f, (pageHeight - imageHeight) * (1f - ratio)),
            SeamSide.Right => (pageWidth - imageWidth, (pageHeight - imageHeight) * (1f - ratio)),
            SeamSide.Top => ((pageWidth - imageWidth) * ratio, pageHeight - imageHeight),
            SeamSide.Bottom => ((pageWidth - imageWidth) * ratio, 0f),
            _ => throw new ArgumentOutOfRangeException(nameof(side), side, "Unsupported seam side.")
        };
    }

    private static (float Width, float Height, int Rotation) GetPageMetrics(PdfReader reader, int pageNumber)
    {
        var rotation = reader.GetPageRotation(pageNumber);
        var size = reader.GetPageSize(pageNumber);
        var width = rotation is 90 or 270 ? size.Height : size.Width;
        var height = rotation is 90 or 270 ? size.Width : size.Height;
        return (width, height, rotation);
    }

    private (float X, float Y) ResolvePagePosition(int pageNumber, PageStampOptions options)
    {
        if (options.PerPagePositions.TryGetValue(pageNumber, out var custom))
        {
            return (custom.X, custom.Y);
        }

        return (options.PositionX, options.PositionY);
    }

    private (float X, float Y, float Rotation) ApplyRandomOffsetIfNeeded(float x, float y, bool enabled)
    {
        if (!enabled)
        {
            return (x, y, 0f);
        }

        var deltaX = 0.01f * _random.Next(-2, 3);
        var deltaY = 0.01f * _random.Next(-2, 3);
        var rotation = _random.Next(-2, 3);

        var adjustedX = x;
        var adjustedY = y;
        if (adjustedX + deltaX is > 0f and < 1f)
        {
            adjustedX += deltaX;
        }
        if (adjustedY + deltaY is > 0f and < 1f)
        {
            adjustedY += deltaY;
        }

        return (Math.Clamp(adjustedX, 0f, 1f), Math.Clamp(adjustedY, 0f, 1f), rotation);
    }

    private static List<int> ResolveSeamPages(SeamStampOptions options, int totalPages)
    {
        var pages = options.Scope switch
        {
            SeamScope.None => new List<int>(),
            SeamScope.All => Enumerable.Range(1, totalPages).ToList(),
            SeamScope.Odd => Enumerable.Range(1, totalPages).Where(p => p % 2 == 1).ToList(),
            SeamScope.Even => Enumerable.Range(1, totalPages).Where(p => p % 2 == 0).ToList(),
            SeamScope.Custom => options.CustomPages
                .Where(p => p >= 1 && p <= totalPages)
                .Distinct()
                .OrderBy(p => p)
                .ToList(),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Scope), options.Scope, "Unknown seam scope.")
        };

        return pages;
    }

    private static List<int> ResolvePageStampPages(PageStampOptions options, int totalPages)
    {
        return options.Scope switch
        {
            PageStampScope.None => new(),
            PageStampScope.All => Enumerable.Range(1, totalPages).ToList(),
            PageStampScope.SkipFirst => totalPages <= 1 ? new List<int>() : Enumerable.Range(2, totalPages - 1).ToList(),
            PageStampScope.SkipLast => totalPages <= 1 ? new List<int>() : Enumerable.Range(1, totalPages - 1).ToList(),
            PageStampScope.Custom => options.CustomPages
                .Where(p => p >= 1 && p <= totalPages)
                .Distinct()
                .OrderBy(p => p)
                .ToList(),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Scope), options.Scope, "Unknown page stamp scope.")
        };
    }

    private static (Bitmap Bitmap, int OriginalWidth) PrepareStampBitmap(StampingOptions options)
    {
        if (!File.Exists(options.StampImagePath))
        {
            throw new FileNotFoundException("Stamp image file was not found.", options.StampImagePath);
        }

        var loaded = ImageUtilities.LoadStampBitmap(options.StampImagePath, options.ConvertWhiteToTransparent, options.StampOpacity);
        var originalWidth = loaded.Width;

        if (options.StampRotationDegrees % 360 != 0)
        {
            var rotated = ImageUtilities.Rotate(loaded, options.StampRotationDegrees, options.KeepOriginalBoundsWhenRotating);
            loaded.Dispose();
            loaded = rotated;
        }

        return (loaded, originalWidth);
    }

    private SignatureContext? PrepareSignatureContext(SignatureOptions options)
    {
        return options.Mode switch
        {
            SignatureMode.None => null,
            SignatureMode.SelfSigned => throw new NotSupportedException("Self-signed certificate generation is not supported in the CLI version."),
            SignatureMode.CustomCertificate => LoadSignatureFromPfx(options),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Mode), options.Mode, "Unknown signature mode.")
        };
    }

    private static SignatureContext LoadSignatureFromPfx(SignatureOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.CertificatePath))
        {
            throw new ArgumentException("Certificate path must be supplied when using sign-mode 'pfx'.");
        }

        if (!File.Exists(options.CertificatePath))
        {
            throw new FileNotFoundException("Signature certificate file was not found.", options.CertificatePath);
        }

        var password = options.Password ?? string.Empty;
        var storageFlags = X509KeyStorageFlags.Exportable | X509KeyStorageFlags.EphemeralKeySet;
        var collection = new X509Certificate2Collection();
        collection.Import(options.CertificatePath, password, storageFlags);

        if (collection.Count == 0)
        {
            throw new InvalidOperationException("The provided PFX file does not contain any certificates.");
        }

        var certificates = new List<X509Certificate2>(collection.Count);
        var chain = new List<BcX509Certificate>(collection.Count);
        var parser = new X509CertificateParser();

        X509Certificate2? signingCertificate = null;
        foreach (X509Certificate2 cert in collection)
        {
            certificates.Add(cert);
            chain.Add(parser.ReadCertificate(cert.RawData));
            if (signingCertificate is null && cert.HasPrivateKey)
            {
                signingCertificate = cert;
            }
        }

        signingCertificate ??= certificates[0];
        using RSA? rsa = signingCertificate.GetRSAPrivateKey();
        if (rsa is null)
        {
            throw new InvalidOperationException("The signing certificate does not contain an RSA private key.");
        }

        var keyPair = DotNetUtilities.GetKeyPair(rsa);
        return new SignatureContext(certificates, chain.ToArray(), keyPair.Private);
    }

    private sealed class SignatureContext : IDisposable
    {
        public SignatureContext(List<X509Certificate2> certificates, BcX509Certificate[] chain, ICipherParameters privateKey)
        {
            Certificates = certificates;
            Chain = chain;
            PrivateKey = privateKey;
        }

        public IReadOnlyList<X509Certificate2> Certificates { get; }
        public BcX509Certificate[] Chain { get; }
        public ICipherParameters PrivateKey { get; }

        public void Dispose()
        {
            foreach (var certificate in Certificates)
            {
                certificate.Dispose();
            }
        }
    }

    private static List<string> ResolveInputFiles(StampingOptions options)
    {
        if (options.Inputs.Count == 0)
        {
            throw new ArgumentException("At least one input path must be specified.");
        }

        var searchOption = options.IncludeSubdirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var results = new List<string>();

        foreach (var input in options.Inputs)
        {
            var fullPath = Path.GetFullPath(input);
            if (File.Exists(fullPath))
            {
                if (!IsPdf(fullPath))
                {
                    throw new InvalidOperationException($"Input file '{fullPath}' is not a PDF.");
                }
                if (set.Add(fullPath))
                {
                    results.Add(fullPath);
                }
                continue;
            }

            if (!Directory.Exists(fullPath))
            {
                throw new DirectoryNotFoundException($"Input path '{fullPath}' was not found.");
            }

            var files = Directory.EnumerateFiles(fullPath, "*.pdf", searchOption)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);

            foreach (var file in files)
            {
                var normalized = Path.GetFullPath(file);
                if (set.Add(normalized))
                {
                    results.Add(normalized);
                }
            }
        }

        return results;
    }

    private static void ValidateOptions(StampingOptions options)
    {
        if (!options.ProcessInPlace && string.IsNullOrWhiteSpace(options.OutputDirectory))
        {
            throw new ArgumentException("Output directory must be specified when not stamping in-place.");
        }

        if (string.IsNullOrWhiteSpace(options.StampImagePath))
        {
            throw new ArgumentException("Stamp image path must be specified.");
        }

        if (options.StampWidthMillimetres <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options.StampWidthMillimetres), "Stamp width must be greater than zero.");
        }

        if (options.Seam.Scope == SeamScope.Custom && options.Seam.CustomPages.Count == 0)
        {
            throw new ArgumentException("Seam scope 'custom' requires at least one page number.");
        }

        if (options.PageStamp.Scope == PageStampScope.Custom && options.PageStamp.CustomPages.Count == 0)
        {
            throw new ArgumentException("Page stamp scope 'custom' requires at least one page number.");
        }

        if (options.Signature.Mode != SignatureMode.None && !string.IsNullOrEmpty(options.OutputPassword))
        {
            throw new ArgumentException("Output encryption cannot be combined with digital signatures in the CLI version.");
        }
    }

    private static string ComposeOutputPath(string inputPath, string outputDirectory, string suffix)
    {
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(inputPath);
        var outputFileName = string.Concat(fileNameWithoutExtension, suffix, ".pdf");
        return Path.Combine(outputDirectory, outputFileName);
    }

    private string ComposeOutputPath(string inputPath)
    {
        if (_options.OutputDirectory is null)
        {
            throw new InvalidOperationException("Output directory was not provided.");
        }

        return ComposeOutputPath(inputPath, _options.OutputDirectory, _options.OutputSuffix);
    }

    private static void EnsureOutputPathValid(string inputPath, string outputPath, bool overwrite)
    {
        if (string.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Output path resolves to the input file. Please choose a different output directory or suffix.");
        }

        if (!overwrite && File.Exists(outputPath))
        {
            throw new IOException($"Output file '{outputPath}' already exists. Use --overwrite to replace it.");
        }
    }

    private List<(int StartX, int Width)> BuildSeamSlicePlan(int totalSlices)
    {
        var plan = new List<(int, int)>(totalSlices);
        if (totalSlices <= 0)
        {
            return plan;
        }

        var totalWidth = _stampBitmap.Width;
        if (totalWidth <= 0)
        {
            return plan;
        }

        var firstSliceWidth = totalSlices == 1 ? totalWidth : totalWidth / 3.0;
        var remainingSlices = Math.Max(1, totalSlices - 1);
        var otherSliceWidth = totalSlices <= 1 ? totalWidth : (totalWidth - firstSliceWidth) / remainingSlices;

        double cursor = 0;
        for (var index = 0; index < totalSlices; index++)
        {
            double rawWidth;
            if (totalSlices == 1)
            {
                rawWidth = totalWidth;
            }
            else if (index == 0)
            {
                rawWidth = firstSliceWidth;
            }
            else if (index == totalSlices - 1)
            {
                rawWidth = totalWidth - cursor;
            }
            else
            {
                rawWidth = otherSliceWidth;
            }

            var start = (int)Math.Round(cursor, MidpointRounding.AwayFromZero);
            if (start >= totalWidth)
            {
                start = Math.Max(0, totalWidth - 1);
            }

            var width = (int)Math.Round(rawWidth, MidpointRounding.AwayFromZero);
            if (width <= 0)
            {
                width = 1;
            }

            if (start + width > totalWidth || index == totalSlices - 1)
            {
                width = Math.Max(1, totalWidth - start);
            }

            plan.Add((start, width));
            cursor = start + width;
        }

        return plan;
    }

    private static Bitmap ExtractSlice(Bitmap source, int startX, int width)
    {
        if (source.Width <= 0)
        {
            throw new InvalidOperationException("Stamp image width must be greater than zero.");
        }

        startX = Math.Clamp(startX, 0, Math.Max(0, source.Width - 1));
        width = Math.Clamp(width, 1, Math.Max(1, source.Width - startX));
        var slice = new Bitmap(width, source.Height);
        slice.SetResolution(source.HorizontalResolution, source.VerticalResolution);
        using var graphics = Graphics.FromImage(slice);
        graphics.DrawImage(source, new DrawingRectangle(0, 0, width, source.Height), new DrawingRectangle(startX, 0, width, source.Height), GraphicsUnit.Pixel);
        return slice;
    }

    private static bool IsPdf(string path) => string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase);

    private static float CalculateScalePercent(float widthMillimetres, int originalWidthPixels)
    {
        if (originalWidthPixels <= 0)
        {
            return 100f;
        }

        var targetWidthPoints = widthMillimetres * 72f / 25.4f;
        return targetWidthPoints / originalWidthPixels * 100f;
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(StampProcessor));
        }
    }

    private static string CreateTemporarySiblingPath(string inputPath)
    {
        var directory = Path.GetDirectoryName(inputPath);
        if (string.IsNullOrEmpty(directory))
        {
            directory = Directory.GetCurrentDirectory();
        }

        var name = Path.GetFileNameWithoutExtension(inputPath);
        var tempFile = $"{name}.{Guid.NewGuid():N}.pdf";
        return Path.Combine(directory, tempFile);
    }

    private static void CommitInPlaceReplacement(string inputPath, string tempPath)
    {
        try
        {
            File.Replace(tempPath, inputPath, null);
        }
        catch (PlatformNotSupportedException)
        {
            FallbackCopyReplace(inputPath, tempPath);
        }
        catch (IOException)
        {
            FallbackCopyReplace(inputPath, tempPath);
        }
        catch (UnauthorizedAccessException)
        {
            FallbackCopyReplace(inputPath, tempPath);
        }
    }

    private static void FallbackCopyReplace(string inputPath, string tempPath)
    {
        try
        {
            EnsureWritable(inputPath);
            File.Copy(tempPath, inputPath, true);
            File.Delete(tempPath);
        }
        catch
        {
            SafeDelete(tempPath);
            throw;
        }
    }

    private static void SafeDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Swallow cleanup exceptions to avoid masking the original failure.
        }
    }

    private static void EnsureWritable(string inputPath)
    {
        try
        {
            if (!File.Exists(inputPath))
            {
                return;
            }

            var attributes = File.GetAttributes(inputPath);
            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                File.SetAttributes(inputPath, attributes & ~FileAttributes.ReadOnly);
            }
        }
        catch
        {
            // Ignore attribute failures; copy attempt will surface any real issue.
        }
    }
}
