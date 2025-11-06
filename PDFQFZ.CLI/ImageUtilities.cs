using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace PDFQFZ.CLI;

internal static class ImageUtilities
{
    public static Bitmap LoadStampBitmap(string path, bool convertWhiteToTransparent, int opacityPercent)
    {
        var bitmap = new Bitmap(path);
        try
        {
            if (convertWhiteToTransparent && !Equals(bitmap.RawFormat, ImageFormat.Png))
            {
                var converted = SetWhiteToTransparent(bitmap);
                bitmap.Dispose();
                bitmap = converted;
            }

            if (opacityPercent < 100)
            {
                var adjusted = SetImageOpacity(bitmap, opacityPercent);
                bitmap.Dispose();
                bitmap = adjusted;
            }

            return bitmap;
        }
        catch
        {
            bitmap.Dispose();
            throw;
        }
    }

    public static Bitmap Rotate(Bitmap bitmap, int degrees, bool keepBounds)
    {
        if (degrees % 360 == 0)
        {
            return new Bitmap(bitmap);
        }

        var angle = degrees % 360;
        var radian = angle * Math.PI / 180.0;
        var cos = Math.Cos(radian);
        var sin = Math.Sin(radian);

        var w = bitmap.Width;
        var h = bitmap.Height;
        var newWidth = (int)Math.Round(Math.Max(Math.Abs(w * cos - h * sin), Math.Abs(w * cos + h * sin)));
        var newHeight = (int)Math.Round(Math.Max(Math.Abs(w * sin - h * cos), Math.Abs(w * sin + h * cos)));

        if (keepBounds)
        {
            newHeight = newHeight * w / newWidth;
            newWidth = w;
        }

        var rotatedImage = new Bitmap(newWidth, newHeight);
        rotatedImage.SetResolution(bitmap.HorizontalResolution, bitmap.VerticalResolution);

        using (var g = Graphics.FromImage(rotatedImage))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            var offsetX = (float)(newWidth - w) / 2;
            var offsetY = (float)(newHeight - h) / 2;

            using (var matrix = new Matrix())
            {
                matrix.RotateAt(angle, new PointF(newWidth / 2f, newHeight / 2f));
                g.Transform = matrix;
                g.DrawImage(bitmap, new RectangleF(offsetX, offsetY, w, h));
            }
        }

        return rotatedImage;
    }

    public static Bitmap[] SplitForSeam(Bitmap source, int slices)
    {
        if (slices <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(slices));
        }

        var result = new Bitmap[slices];
        var totalHeight = source.Height;
        var totalWidth = source.Width;

        if (slices == 1)
        {
            result[0] = new Bitmap(source);
            return result;
        }

        var firstSliceWidth = totalWidth / 3;
        var defaultSliceWidth = (totalWidth - firstSliceWidth) / slices;
        var indexLimit = slices - 1;
        var remainingWidth = totalWidth;

        for (var index = 0; index <= indexLimit; index++)
        {
            int sliceWidth;
            if (index == indexLimit)
            {
                sliceWidth = remainingWidth;
            }
            else if (index == 0)
            {
                sliceWidth = firstSliceWidth;
            }
            else
            {
                sliceWidth = defaultSliceWidth;
            }

            var segment = new Bitmap(sliceWidth, totalHeight);
            segment.SetResolution(source.HorizontalResolution, source.VerticalResolution);
            using (var g = Graphics.FromImage(segment))
            {
                var sourceRect = new Rectangle(totalWidth - remainingWidth, 0, sliceWidth, totalHeight);
                g.DrawImage(source, new Rectangle(0, 0, sliceWidth, totalHeight), sourceRect, GraphicsUnit.Pixel);
            }

            result[index] = segment;
            remainingWidth -= sliceWidth;
        }

        return result;
    }

    private static Bitmap SetWhiteToTransparent(Bitmap source)
    {
        var target = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
        target.SetResolution(source.HorizontalResolution, source.VerticalResolution);
        for (var x = 0; x < source.Width; x++)
        {
            for (var y = 0; y < source.Height; y++)
            {
                var pixel = source.GetPixel(x, y);
                if (pixel.R > 230 && pixel.G > 230 && pixel.B > 230)
                {
                    target.SetPixel(x, y, Color.FromArgb(0, pixel));
                }
                else
                {
                    target.SetPixel(x, y, pixel);
                }
            }
        }

        return target;
    }

    private static Bitmap SetImageOpacity(Bitmap source, int opacityPercentage)
    {
        var alpha = Math.Clamp(opacityPercentage, 0, 100) * 255 / 100;
        var target = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
        target.SetResolution(source.HorizontalResolution, source.VerticalResolution);
        for (var x = 0; x < source.Width; x++)
        {
            for (var y = 0; y < source.Height; y++)
            {
                var pixel = source.GetPixel(x, y);
                if (pixel.A == 0)
                {
                    target.SetPixel(x, y, pixel);
                }
                else
                {
                    var withAlpha = Color.FromArgb(alpha, pixel.R, pixel.G, pixel.B);
                    target.SetPixel(x, y, withAlpha);
                }
            }
        }

        return target;
    }
}
