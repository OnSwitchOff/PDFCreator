using PDFCreator.Enums;

namespace PDFCreator.Models
{
    public class TableVisualizationModel
    {
        #region "Constructor"
        /// <summary>
        /// Create TableVisualizationModel object with default properties:
        /// - HeaderFont, ContentFont: - FontName - Helvetica;
        ///                            - FontSize - 12;
        ///                            - IsBold, IsItalic, Subscript, Superscript - false;
        ///                            - Underline - None;
        ///                            - FontColor - black.
        /// - HeaderBackground, ContentBackground - white;
        /// - DifferentEvenRowsBackground - false;
        /// - EvenRowsBackground - grey;
        /// - BorderWidth - 0.5;
        /// - BorderColor - black;
        /// - BorderStyle - Single;
        /// - BorderPosition - AllBorders;
        /// - TopCellPadding, BottomCellPadding - 0.1.
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public TableVisualizationModel()
        {
            HeaderFont = new FontModel();
            ContentFont = new FontModel();
            HeaderBackground = new ColorModel(255, 255, 255);
            ContentBackground = new ColorModel(255, 255, 255);
            DifferentEvenRowsBackground = false;
            EvenRowsBackground = new ColorModel(220, 220, 220);

            BorderWidth = 0.5;
            BorderColor = new ColorModel();
            BorderStyle = BorderStyle.Single;
            BorderLocation = BordersOption.AllBorders;

            TopCellPadding = 0.1;
            BottomCellPadding = 0.1;
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Table header font
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public FontModel HeaderFont { get; set; }

        /// <summary>
        /// Table content font
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public FontModel ContentFont { get; set; }

        /// <summary>
        /// Table header background
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public ColorModel HeaderBackground { get; set; }

        /// <summary>
        /// Table content background
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public ColorModel ContentBackground { get; set; }

        /// <summary>
        /// Is there background different between even and not even rows
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public bool DifferentEvenRowsBackground { get; set; }

        /// <summary>
        /// Even rows background
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public ColorModel EvenRowsBackground { get; set; }

        /// <summary>
        /// Border width
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public double BorderWidth { get; set; }

        /// <summary>
        /// Border color
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public ColorModel BorderColor { get; set; }

        /// <summary>
        /// Style of the line of the border
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public BorderStyle BorderStyle { get; set; }

        /// <summary>
        /// Lines border location
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public BordersOption BorderLocation { get; set; }

        /// <summary>
        /// Top cell table padding
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public double TopCellPadding { get; set; }

        /// <summary>
        /// Bottom cell table padding
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public double BottomCellPadding { get; set; }
        #endregion
    }
}
