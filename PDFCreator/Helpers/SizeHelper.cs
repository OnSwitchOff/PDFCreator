using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace PDFCreator.Helpers
{
    public class SizeHelper
    {
        #region "Declarations"
        // pixel count in 1 cm
        private const float PixelToCentimeterRatio = 2.54f;
        // point count in 1 mm
        private const float PointToMillimeterRatio = 0.35277777777778f;
        // error image format message
        private const string ErrorMessage = "Could not recognise image format.";
        // dictionary with images header 
        private static Dictionary<byte[], Func<BinaryReader, SizeHelper>> imageFormatDecoders;
        #endregion

        #region "Constructors"
        /// <summary>
        /// Initialize images header dictionary
        /// </summary>
        static SizeHelper()
        {
            imageFormatDecoders = new Dictionary<byte[], Func<BinaryReader, SizeHelper>>()
            {
                { new byte[] { 0x42, 0x4D }, DecodeBitmap },
                { new byte[] { 0x47, 0x49, 0x46, 0x38, 0x37, 0x61 }, DecodeGif },
                { new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }, DecodeGif },
                { new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, DecodePng },
                { new byte[] { 0xff, 0xd8 }, DecodeJpeg },
            };
        }

        /// <summary>
        /// Create new SizeHelper object
        /// </summary>
        /// <param name="w">Object width</param>
        /// <param name="h">Object height</param>
        public SizeHelper(ushort w, ushort h)
        {
            Width = w;
            Height = h;
        }
        #endregion

        #region "Properties"
        public ushort Width { get; set; }
        public ushort Height { get; set; }
        #endregion

        #region "Public methods"
        /// <summary>        
        /// Gets the dimensions of an image.        
        /// </summary>        
        /// <param name="path">The path of the image to get the dimensions of.</param>        
        /// <returns>The dimensions of the specified image.</returns>        
        /// <exception cref="ArgumentException">The image was of an unrecognised format.</exception>        
        public static SizeHelper GetImageSize(string path)
        {
            try
            {
                using (BinaryReader binaryReader = new BinaryReader(File.OpenRead(path)))
                {
                    try
                    {
                        return GetImageSize(binaryReader);
                    }
                    catch (ArgumentException e)
                    {
                        string newMessage = string.Format("{0} file: '{1}' ", ErrorMessage, path);

                        throw new ArgumentException(newMessage, "path", e);
                    }
                }
            }
            catch (ArgumentException)
            {
                //do it the old fashioned way
                using (Bitmap b = new Bitmap(path))
                {
                    return new SizeHelper((ushort)b.Width, (ushort)b.Height);
                }
            }
        }

        /// <summary>        
        /// Gets the dimensions of an image.        
        /// </summary>        
        /// <param name="path">The path of the image to get the dimensions of.</param>        
        /// <returns>The dimensions of the specified image.</returns>        
        /// <exception cref="ArgumentException">The image was of an unrecognised format.</exception>            
        public static SizeHelper GetImageSize(BinaryReader binaryReader)
        {
            int maxMagicBytesLength = imageFormatDecoders.Keys.OrderByDescending(x => x.Length).First().Length;
            byte[] magicBytes = new byte[maxMagicBytesLength];
            for (int i = 0; i < maxMagicBytesLength; i += 1)
            {
                magicBytes[i] = binaryReader.ReadByte();
                foreach (var kvPair in imageFormatDecoders)
                {
                    if (StartsWith(magicBytes, kvPair.Key))
                    {
                        return kvPair.Value(binaryReader);
                    }
                }
            }

            throw new ArgumentException(ErrorMessage, "binaryReader");
        }

        /// <summary>
        /// Convert size in pixels to size in centimeters 
        /// </summary>
        /// <param name="pixelsSize">Size in pixels</param>
        /// <param name="dpi">Dots per inch</param>
        /// <returns>Size in centimeters</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>16.03.2021</date>
        public static float ConvertPixelToCentimeter(float pixelsSize, int dpi = 300)
        {
            return pixelsSize * PixelToCentimeterRatio / dpi;
        }

        /// <summary>
        /// Convert size in centimeters to size in pixels 
        /// </summary>
        /// <param name="centimetersSize">Size in centimeters</param>
        /// <param name="dpi">Dots per inch</param>
        /// <returns>Size in pixels</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>16.03.2021</date>
        public static float ConvertCentimeterToPixel(float centimetersSize, int dpi = 300)
        {
            return dpi * centimetersSize / PixelToCentimeterRatio;
        }

        /// <summary>
        /// Convert size in points (for example font size) to size in millimeters 
        /// </summary>
        /// <param name="centimetersSize">Size in centimeters</param>
        /// <returns>Size in pixels</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>16.03.2021</date>
        public static float ConvertPointToMillimeter(float pointSize)
        {
            return pointSize * PointToMillimeterRatio;
        }
        #endregion

        #region "Private methods"
        /// <summary>
        /// Compare two byte arrays
        /// </summary>
        /// <param name="thisBytes">First byte array</param>
        /// <param name="thatBytes">Second byte array</param>
        /// <returns>true - if the first byte array equels to the second byte array else return false</returns>
        private static bool StartsWith(byte[] thisBytes, byte[] thatBytes)
        {
            for (int i = 0; i < thatBytes.Length; i += 1)
            {
                if (thisBytes[i] != thatBytes[i])
                {
                    return false;
                }
            }

            return true;
        }

        private static short ReadLittleEndianInt16(BinaryReader binaryReader)
        {
            byte[] bytes = new byte[sizeof(short)];

            for (int i = 0; i < sizeof(short); i += 1)
            {
                bytes[sizeof(short) - 1 - i] = binaryReader.ReadByte();
            }
            return BitConverter.ToInt16(bytes, 0);
        }

        private static ushort ReadLittleEndianUInt16(BinaryReader binaryReader)
        {
            byte[] bytes = new byte[sizeof(ushort)];

            for (int i = 0; i < sizeof(ushort); i += 1)
            {
                bytes[sizeof(ushort) - 1 - i] = binaryReader.ReadByte();
            }
            return BitConverter.ToUInt16(bytes, 0);
        }

        private static int ReadLittleEndianInt32(BinaryReader binaryReader)
        {
            byte[] bytes = new byte[sizeof(int)];
            for (int i = 0; i < sizeof(int); i += 1)
            {
                bytes[sizeof(int) - 1 - i] = binaryReader.ReadByte();
            }
            return BitConverter.ToInt32(bytes, 0);
        }

        /// <summary>
        /// Get image size from Bitmap image header data
        /// </summary>
        /// <param name="binaryReader">Image file header data</param>
        /// <returns>Width and height of image</returns>
        private static SizeHelper DecodeBitmap(BinaryReader binaryReader)
        {
            binaryReader.ReadBytes(16);
            int width = binaryReader.ReadInt32();
            int height = binaryReader.ReadInt32();
            return new SizeHelper((ushort)width, (ushort)height);
        }

        /// <summary>
        /// Get image size from Gif image header data
        /// </summary>
        /// <param name="binaryReader">Image file header data</param>
        /// <returns>Width and height of image</returns>
        private static SizeHelper DecodeGif(BinaryReader binaryReader)
        {
            int width = binaryReader.ReadInt16();
            int height = binaryReader.ReadInt16();
            return new SizeHelper((ushort)width, (ushort)height);
        }

        /// <summary>
        /// Get image size from Png image header data
        /// </summary>
        /// <param name="binaryReader">Image file header data</param>
        /// <returns>Width and height of image</returns>
        private static SizeHelper DecodePng(BinaryReader binaryReader)
        {
            binaryReader.ReadBytes(8);
            int width = ReadLittleEndianInt32(binaryReader);
            int height = ReadLittleEndianInt32(binaryReader);
            return new SizeHelper((ushort)width, (ushort)height);
        }

        /// <summary>
        /// Get image size from Jpeg image header data
        /// </summary>
        /// <param name="binaryReader">Image file header data</param>
        /// <returns>Width and height of image</returns>
        private static SizeHelper DecodeJpeg(BinaryReader binaryReader)
        {
            while (binaryReader.ReadByte() == 0xff)
            {
                byte marker = binaryReader.ReadByte();
                short chunkLength = ReadLittleEndianInt16(binaryReader);
                if (marker == 0xc0)
                {
                    binaryReader.ReadByte();
                    int height = ReadLittleEndianInt16(binaryReader);
                    int width = ReadLittleEndianInt16(binaryReader);
                    return new SizeHelper((ushort)width, (ushort)height);
                }

                if (chunkLength < 0)
                {
                    ushort uchunkLength = (ushort)chunkLength;
                    binaryReader.ReadBytes(uchunkLength - 2);
                }
                else
                {
                    binaryReader.ReadBytes(chunkLength - 2);
                }
            }

            throw new ArgumentException(ErrorMessage);
        }
        #endregion
    }
}
