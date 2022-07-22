using MigraDoc.DocumentObjectModel;

namespace PDFCreator.Models
{
    public class ColorModel
    {
        #region "Constructors"
        /// <summary>
        /// Create ColorModel object with default properties (black color): Color_R, Color_G, Color_B - 0. 
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public ColorModel() : this(0, 0, 0)
        {

        }

        /// <summary>
        /// Create ColorModel object
        /// </summary>
        /// <param name="r">Red in an additive color model</param>
        /// <param name="g">Green in an additive color model</param>
        /// <param name="b">Blue in an additive color model</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public ColorModel(byte r, byte g, byte b)
        {
            Color_R = r;
            Color_G = g;
            Color_B = b;
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Red in an additive color model
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public byte Color_R { get; set; }

        /// <summary>
        /// Green in an additive color model
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public byte Color_G { get; set; }

        /// <summary>
        /// Blue in an additive color model
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public byte Color_B { get; set; }
        #endregion

        #region "Methods"
        /// <summary>
        /// Transform ColorModel object to Color
        /// </summary>
        /// <param name="colorModel">ColorModel object</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public static explicit operator Color(ColorModel colorModel)
        {
            return Color.FromRgb(colorModel.Color_R, colorModel.Color_G, colorModel.Color_B);
        }
        #endregion
    }
}
