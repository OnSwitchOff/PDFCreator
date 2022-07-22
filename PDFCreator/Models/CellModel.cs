using PDFCreator.Enums;

namespace PDFCreator.Models
{
    public class CellModel
    {
        #region "Constructor"
        /// <summary>
        /// Create a CellModel object with the following properties: 
        /// - Value - Empty;
        /// - Font: 
        ///     - FontName - Helvetica;
        ///     - FontSize - 12;
        ///     - IsBold, IsItalic, Subscript, Superscript - false;
        ///     - Underline - None;
        ///     - FontColor - black.
        /// - MergeDirection - NoMerge;
        /// - MergeCellCount - 0;
        /// - HorizontalAlignment - Left;
        /// - VerticalAlignment - Center;
        /// - TopPadding, BottomPadding - 0.
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public CellModel()
        {
            Value = string.Empty;
            Font = new FontModel();
            MergeDirection = MergeDirection.NoMerge;
            MergeCellCount = 0;
            HorizontalAlignment = HorizontalAlignment.Left;
            VerticalAlignment = VerticalAlignment.Center;

            TopPadding = 0;
            BottomPadding = 0;
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Cell text
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string Value { get; set; }

        /// <summary>
        /// Cell text font
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public FontModel Font { get; set; }

        /// <summary>
        /// Cells merge direction
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public MergeDirection MergeDirection { get; set; }

        /// <summary>
        /// Cells count for merge 
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public int MergeCellCount { get; set; }

        /// <summary>
        /// Cells text horizontal alignment 
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public HorizontalAlignment HorizontalAlignment { get; set; }

        /// <summary>
        /// Cells text vertical alignment 
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public VerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// Top cell padding
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public double TopPadding { get; set; }

        /// <summary>
        /// Bottom cell padding
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public double BottomPadding { get; set; }
        #endregion
    }
}
