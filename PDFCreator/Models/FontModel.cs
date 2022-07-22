using System;
using MigraDoc.DocumentObjectModel;
using Underline = PDFCreator.Enums.Underline;

namespace PDFCreator.Models
{
    public class FontModel
    {
        #region "constructor"
        /// <summary>
        /// Create an object "FontModel" with the following page setupes: 
        /// - FontName - Arial;
        /// - FontSize - 12;
        /// - IsBold, IsItalic, Subscript, Superscript - false;
        /// - Underline - None;
        /// - FontColor - black.
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public FontModel()
        {
            FontName = "Arial";
            FontSize = 12;
            IsBold = false;
            IsItalic = false;
            Underline = Underline.None;
            Subscript = false;
            Superscript = false;
            FontColor = new ColorModel();
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Font family name
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public string FontName { get; set; }

        /// <summary>
        /// Font size
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public double FontSize { get; set; }

        /// <summary>
        /// Is the font italic
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public bool IsItalic { get; set; }

        /// <summary>
        /// Is the font bold
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public bool IsBold { get; set; }

        /// <summary>
        /// Type of the text underline
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public Underline Underline { get; set; }

        /// <summary>
        /// Subscript font
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public bool Subscript { get; set; }

        /// <summary>
        /// Superscript font
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public bool Superscript { get; set; }

        /// <summary>
        /// Font color
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public ColorModel FontColor { get; set; }
        #endregion

        #region "Methods"
        /// <summary>
        /// Transform object FontModel to Font
        /// </summary>
        /// <param name="_font"></param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public static explicit operator Font(FontModel _font)
        {
            Font font = new Font();
            font.Name = _font.FontName;
            font.Size = _font.FontSize;
            font.Bold = _font.IsBold;
            font.Italic = _font.IsItalic;
            font.Underline = (MigraDoc.DocumentObjectModel.Underline)Enum.Parse(typeof(MigraDoc.DocumentObjectModel.Underline), _font.Underline.ToString());
            font.Subscript = _font.Subscript;
            font.Superscript = _font.Superscript;
            font.Color = (Color)_font.FontColor;

            return font;
        }
        #endregion
    }
}
