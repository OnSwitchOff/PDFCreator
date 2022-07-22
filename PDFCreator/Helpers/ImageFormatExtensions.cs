using System.Drawing.Imaging;

namespace PDFCreator.Helpers
{
    public static class ImageFormatExtensions
    {
        /// <summary>
        /// Gets format of image.
        /// </summary>
        /// <param name="img">Image to get format of image.</param>
        /// <returns>ImageFormat/returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>24.06.2022.</date>
        public static ImageFormat GetImageFormat(this System.Drawing.Image img)
        {
            ImageCodecInfo[] decoders = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo decoder in decoders)
            {
                if (img.RawFormat.Guid == decoder.FormatID)
                {
                    switch (decoder.FormatDescription.ToLower())
                    {
                        case "bmp":
                            return ImageFormat.Bmp;
                        case "jpeg":
                            return ImageFormat.Jpeg;
                        case "png":
                            return ImageFormat.Png;
                        case "gif":
                            return ImageFormat.Gif;
                        case "tiff":
                            return ImageFormat.Tiff;
                        case "ico":
                            return ImageFormat.Icon;
                        case "emf":
                            return ImageFormat.Emf;
                        case "wmf":
                            return ImageFormat.Wmf;
                    }
                }
            }

            return null;
        }
    }
}
