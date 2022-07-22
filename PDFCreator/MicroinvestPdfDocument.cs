using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using PDFCreator.Enums;
using PDFCreator.Helpers;
using PDFCreator.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Resources;
using PageFormat = PDFCreator.Enums.PageFormat;

namespace PDFCreator
{
    public class MicroinvestPdfDocument : PdfDocument
    {

        #region "Declarations"
        // header's table  used as Grid
        private Table headerTable;
        // header's table  used as Grid (in case DifferentFirstPageHeaderFooter is "true") 
        private Table firstPageHeaderTable;
        // footer's table  used as Grid
        private Table footerTable;
        // footer's table  used as Grid (in case DifferentFirstPageHeaderFooter is "true")
        private Table firstPageFooterTable;
        // section with client and document data
        private Table captionSection;
        // section with additional data like payment type, our bank name etc.
        private Table additionalDataSection;
        // table goods involved in the current document
        private DataTable documentData;
        #endregion

        #region "Constructors"
        /// <summary>
        /// Create a blank pdf document with the following page setupes: 
        /// - PageFormat - A4;
        /// - Orientation - portrait;
        /// - TopIndentation, BottomIndentation, RightIndentation - 1;
        /// - LeftIndentation - 2;
        /// - DifferentFirstPageHeaderFooter - false;
        /// - TemplatePageFormat - A4;
        /// - HeaderHeight - 0;
        /// - FirstPageHeaderHeight - 4;
        /// - FooterHeight - 1.3;
        /// - IsUseSectionNumbers - true;
        /// - PriceFormat, QuantityFormat - 2.
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public MicroinvestPdfDocument() : this(PageFormat.A4, PageOrientation.Portrait, 1, 1, 2, 1, false)
        {

        }

        /// <summary>
        /// Create a blank pdf document with the following page setupes: 
        /// - PageFormat - A4;
        /// - Orientation - portrait;
        /// - TopIndentation, BottomIndentation, RightIndentation - 1;
        /// - LeftIndentation - 2;
        /// - TemplatePageFormat - A4;
        /// - HeaderHeight - 0;
        /// - FirstPageHeaderHeight - 4;
        /// - FooterHeight - 1.3;
        /// - IsUseSectionNumbers - true;
        /// - PriceFormat, QuantityFormat - 2.
        /// </summary>
        /// <param name="differentFirstPageHeaderFooter">The first page header and/or footer is different from others</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public MicroinvestPdfDocument(bool differentFirstPageHeaderFooter) : this(PageFormat.A4, PageOrientation.Portrait, 1, 1, 2, 1, differentFirstPageHeaderFooter)
        {

        }

        /// <summary>
        /// Create a blank pdf document with the following page setupes:;
        /// - TemplatePageFormat - A4
        /// - HeaderHeight - 0;
        /// - FirstPageHeaderHeight - 4;
        /// - FooterHeight - 1.3;
        /// - IsUseSectionNumbers - true;
        /// - PriceFormat, QuantityFormat - 2.
        /// </summary>
        /// <param name="pageFormat">Standard page sizes</param>
        /// <param name="orientation">Specifies the page orientation</param>
        /// <param name="topMargin">Indent from the top of the page</param>
        /// <param name="bottomMargin">Indent from the bottom of the page</param>
        /// <param name="leftMargin">Indent from the left side of the page</param>
        /// <param name="rightMargin">Indent from the right side of the page</param>
        /// <param name="differentFirstPageHeaderFooter">The first page header and/or footer is different from others</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public MicroinvestPdfDocument(PageFormat pageFormat, PageOrientation orientation, double topMargin, double bottomMargin, double leftMargin, double rightMargin, bool differentFirstPageHeaderFooter) : base(pageFormat, orientation, topMargin, bottomMargin, leftMargin, rightMargin, differentFirstPageHeaderFooter)
        {
            Saler = new CompanyModel();
            Client = new CompanyModel();
            Document = new DocumentModel();
            DocumentData = new DataTable();
            VATs = new List<VATModel>();

            HeaderHeight = 0;
            FirstPageHeaderHeight = 4;
            FooterHeight = 1.3;
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// First page header grid
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public Table FirstPageHeaderGrid
        {
            get => firstPageHeaderTable == null ? throw new Exception("First page Header grid doesn't exist!!!") : firstPageHeaderTable;
        }

        /// <summary>
        /// First page footer grid
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public Table FirstPageFooterGrid
        {
            get => firstPageFooterTable == null ? throw new Exception("First page Footer grid doesn't exist!!!") : firstPageFooterTable;
        }

        /// <summary>
        /// Header grid
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public Table HeaderGrid
        {
            get => headerTable == null ? throw new Exception("Header grid doesn't exist!!!") : headerTable;

        }

        /// <summary>
        /// Footer grid
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public Table FooterGrid
        {
            get => footerTable == null ? throw new Exception("Footer grid doesn't exist!!!") : footerTable;
        }

        /// <summary>
        /// Our company data
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public CompanyModel Saler { get; set; }

        /// <summary>
        /// Client company data
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public CompanyModel Client { get; set; }

        /// <summary>
        /// Document data
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public DocumentModel Document { get; set; }

        /// <summary>
        /// VAT information of all goods involved in the current document
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public List<VATModel> VATs { get; set; }

        /// <summary>
        /// Table goods involved in the current document
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public DataTable DocumentData
        {
            get => documentData == null ? documentData = new DataTable() : documentData;
            set => documentData = value;
        }

        /// <summary>
        /// Section with client and document data
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public Table CaptionSection
        {
            get
            {
                if (captionSection == null) CreateCaptionSection();

                return captionSection;
            }
            set => captionSection = value;
        }

        /// <summary>
        /// Section with additional data like payment type, our bank name etc.
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>18.03.2021</date>
        public Table AdditionalDataSection
        {
            get
            {
                if (additionalDataSection == null) CreateAdditionalDataSection();

                return additionalDataSection;
            }
            set => additionalDataSection = value;
        }

        /// <summary>
        /// Default header image (company logo)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>30.03.2021</date>
        public System.Drawing.Bitmap DefaultHeaderImage
        {
            get
            {
                using(System.IO.MemoryStream stream = new System.IO.MemoryStream(Resources.Resources.DefaultHeader))
                {
                    return (System.Drawing.Bitmap)System.Drawing.Image.FromStream(stream);
                }
            }
        }

        /// <summary>
        /// Default footer image (company signature)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>30.03.2021</date>
        public System.Drawing.Bitmap DefaultFooterImage
        {
            get
            {
                using (System.IO.MemoryStream stream = new System.IO.MemoryStream(Resources.Resources.DefaultFooter))
                {
                    return (System.Drawing.Bitmap)System.Drawing.Image.FromStream(stream);
                }
            }
        }
        #endregion

        #region "Methods for creating document regions"
        /// <summary>
        /// Create and insert Table (without borders) into Header/Footer on the first page
        /// </summary>
        /// <param name="documentArea">The area (Header or Footer) of the first page ​​document in which the grid will be created</param>
        /// <param name="columnsWidth">Columns width</param>
        /// <param name="rowNumber">Number of rows</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public void CreateFirstPageHeaderFooterGrid(DocumentArea documentArea, double[] columnsWidth, int rowNumber)
        {
            switch (documentArea)
            {
                case DocumentArea.Header:
                    // добавляем таблицу в верхний колонтитул первой страницы
                    firstPageHeaderTable = FirstPageHeader.AddTable();
                    // добавляем колонки заданной пользователем ширины
                    foreach (double width in columnsWidth)
                    {
                        firstPageHeaderTable.AddColumn(string.Format("{0}cm", width));
                    }
                    // вставляем соответствующее кол-во строк                    
                    for (int i = 0; i < rowNumber; i++)
                    {
                        firstPageHeaderTable.AddRow();
                    }

                    break;
                case DocumentArea.Footer:
                    // добавляем таблицу в нижний колонтитул первой страницы
                    firstPageFooterTable = FirstPageFooter.AddTable();
                    // добавляем колонки заданной пользователем ширины
                    foreach (double width in columnsWidth)
                    {
                        firstPageFooterTable.AddColumn(string.Format("{0}cm", width));
                    }
                    // вставляем соответствующее кол-во строк  
                    for (int i = 0; i < rowNumber; i++)
                    {
                        firstPageFooterTable.AddRow();
                    }
                    break;
            }
        }

        /// <summary>
        /// Create and insert Table (without borders) into Header/Footer
        /// </summary>
        /// <param name="documentArea">The area (Header or Footer) of the ​​document in which the grid will be created</param>
        /// <param name="columnsWidth">Columns width</param>
        /// <param name="rowNumber">Number of rows</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public void CreateHeaderFooterGrid(DocumentArea documentArea, double[] columnsWidth, int rowNumber)
        {
            switch (documentArea)
            {
                case DocumentArea.Header:
                    // добавляем таблицу в верхний колонтитул 
                    headerTable = Header.AddTable();
                    // добавляем колонки заданной пользователем ширины
                    foreach (double width in columnsWidth)
                    {
                        HeaderGrid.AddColumn(string.Format("{0}cm", width));
                    }
                    // вставляем соответствующее кол-во строк 
                    for (int i = 0; i < rowNumber; i++)
                    {
                        HeaderGrid.AddRow();
                    }
                    break;
                case DocumentArea.Footer:
                    // добавляем таблицу в нижний колонтитул
                    footerTable = Footer.AddTable();
                    // добавляем колонки заданной пользователем ширины
                    foreach (double width in columnsWidth)
                    {
                        footerTable.AddColumn(string.Format("{0}cm", width));
                    }
                    // вставляем соответствующее кол-во строк 
                    for (int i = 0; i < rowNumber; i++)
                    {
                        footerTable.AddRow();
                    }
                    break;
            }
        }

        /// <summary>
        /// Create caption grid (grid for client and document data visualization)
        /// </summary>
        /// <param name="columnsWidth">Columns width</param>
        /// <param name="rowCount">Number of rows</param>
        /// <remarks>
        /// If columnsWidth is "null", a grid having two columns 2/3 and 1/3 of the page wide respectively will be created.
        /// </remarks>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public void CreateCaptionSection(double[] columnsWidth = null, int rowCount = 1)
        {
            // если пользователь не передал массив с указанием ширины столбцов - создаем его сами
            if (columnsWidth == null)
            {
                // создаем две колонки
                columnsWidth = new double[2];
                // первая колонка с шириной 2/3 ширины страницы без учета правого и левого отступов
                columnsWidth[0] = (PageWidth - RightIndentation - LeftIndentation) / 3 * 2;
                // вторая - 1/3 ширины страницы без учета правого и левого отступов
                columnsWidth[1] = (PageWidth - RightIndentation - LeftIndentation) / 3;
            }

            // создаем таблицу
            CaptionSection = new Table();
            // добавляем в таблицу колонки с заданной пользователем шириной
            for (int i = 0; i < columnsWidth.Length; i++)
            {
                CaptionSection.AddColumn(string.Format("{0}cm", columnsWidth[i]));
            }
            // вставляем заданное пользователем кол-во строк и устанавливаем 
            for (int i = 0; i < rowCount; i++)
            {
                CaptionSection.AddRow();
            }
        }

        /// <summary>
        /// Create additional data grid (grid for payment data and description  of document data)
        /// </summary>
        /// <param name="columnsWidth">Columns width</param>
        /// <param name="rowCount">Number of rows</param>
        /// <remarks>
        /// If columnsWidth is "null", a grid having two columns 1/2 and 1/2 of the page wide respectively will be created.
        /// </remarks>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public void CreateAdditionalDataSection(double[] columnsWidth = null, int rowCount = 2)
        {
            // если пользователь не передал массив с указанием ширины столбцов - создаем его сами
            if (columnsWidth == null)
            {
                // создаем две колонки
                columnsWidth = new double[2];
                // первая колонка с шириной 1/2 ширины страницы без учета правого и левого отступов
                columnsWidth[0] = (PageWidth - RightIndentation - LeftIndentation) / 2;
                // первая колонка с шириной 1/2 ширины страницы без учета правого и левого отступов
                columnsWidth[1] = (PageWidth - RightIndentation - LeftIndentation) / 2;
            }
            // создаем таблицу
            AdditionalDataSection = new Table();

            for (int i = 0; i < columnsWidth.Length; i++)
            {
                AdditionalDataSection.AddColumn(string.Format("{0}cm", columnsWidth[i]));
            }

            Row row;
            // вставляем заданное пользователем кол-во строк и устанавливаем 
            for (int i = 1; i <= rowCount; i++)
            {
                row = AdditionalDataSection.AddRow();
                // "связываем" текущую строку с предыдущей, чтобы все строки оставались на одной странице и не переносились на следующую 
                row.KeepWith = rowCount - i;
            }
        }

        /// <summary>
        /// Create and insert grid with two columns into Header and Footer
        /// </summary>
        /// <param name="headerFirstColumnWidth">The ratio of the width of the first column (in the Header) to the width of the page</param>
        /// <param name="footerFirstColumnWidth">The ratio of the width of the first column (in the Footer) to the width of the page</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public void CreateDefaultHeaderFooterGrid(double headerFirstColumnWidth = (2.0 / 3.0), double footerFirstColumnWidth = (1.0 / 4.0))
        {
            // создаем массив ширины столбцов верхнего колонтитула
            double[] headerColumnsWidth = new double[] { (PageWidth - RightIndentation - LeftIndentation) * headerFirstColumnWidth, (PageWidth - RightIndentation - LeftIndentation) * (1 - headerFirstColumnWidth) };
            // создаем массив ширины столбцов нижнего колонтитула
            double[] footerColumnsWidth = new double[] { (PageWidth - RightIndentation - LeftIndentation) * footerFirstColumnWidth, (PageWidth - RightIndentation - LeftIndentation) * (1 - footerFirstColumnWidth) };
            // формируем сетку верхнего колонтитула
            CreateHeaderFooterGrid(DocumentArea.Header, headerColumnsWidth, 1);
            // формируем сетку нижнего колонтитула
            CreateHeaderFooterGrid(DocumentArea.Footer, footerColumnsWidth, 1);

            // если данный документ имеет другой верхний и/или нижний колонтитулы - формируем также разметку и для колонтитулов первой страницы
            if (DifferentFirstPageHeaderFooter)
            {
                CreateFirstPageHeaderFooterGrid(DocumentArea.Header, headerColumnsWidth, 1);
                CreateFirstPageHeaderFooterGrid(DocumentArea.Footer, footerColumnsWidth, 1);
            }

            // добавляем нумерацию страниц в нижний колонтитул
            AddPageCounterToFooter("{0} / {1}", HorizontalAlignment.Left, Enums.VerticalAlignment.Center);
        }

        #endregion

        #region "Methods for adding data to document"
        /// <summary>
        /// Add page counter into Footer grid (cell 0x0)
        /// </summary>
        /// <param name="visualizerTemplate">Template of the page counter data visualization</param>
        /// <param name="horizontalAlignment">Horizontal alignment content</param>
        /// <param name="verticalAlignment">Vertical alignment content</param>
        /// <param name="font">Content font</param>
        /// <param name="skipFirstPage">Skip page counter on the first page</param>
        /// <remarks>
        /// Examples of the visualizerTemplate: Page {0} of {1}; {0} / {1}; {0} etc., - where:
        /// - {0} - current page;
        /// - {1} - all pages
        /// </remarks>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public void AddPageCounterToFooter(string visualizerTemplate, HorizontalAlignment horizontalAlignment, Enums.VerticalAlignment verticalAlignment, FontModel font = null, bool skipFirstPage = true)
        {
            // если нижний колонтитул размечен и имеет строки
            if (footerTable != null && footerTable.Rows.Count > 0)
            {
                // получаем первую строку
                Row row = FooterGrid.Rows[0];
                // устанавливаем выравнивание контента по горизонтали 
                row.Format.Alignment = ConvertParagraphHorizontalAlignment(horizontalAlignment);
                // устанавливаем выравнивание контента по вертикали 
                row.VerticalAlignment = ConvertVerticalAlignment(verticalAlignment);
                // устаналиваем в ячейку 0х0 текст с нумерацией страниц
                row.Cells[0].Add(CreatePageCounterParagraph(visualizerTemplate, font));

                // если нумерация строк должна быть включая с первой страницы, нижний колонтитул первой страницы размечен и имеет строки 
                if (!skipFirstPage && firstPageFooterTable != null && firstPageFooterTable.Rows.Count > 0)
                {
                    // получаем первую строку
                    row = FirstPageFooterGrid.Rows[0];
                    // устанавливаем выравнивание контента по горизонтали 
                    row.Format.Alignment = ConvertParagraphHorizontalAlignment(horizontalAlignment);
                    // устанавливаем выравнивание контента по вертикали 
                    row.VerticalAlignment = ConvertVerticalAlignment(verticalAlignment);
                    // устаналиваем в ячейку 0х0 текст с нумерацией страниц
                    row.Cells[0].Add(CreatePageCounterParagraph(visualizerTemplate));
                }
            }
        }

        /// <summary>
        /// Add new pdf document item to "Caption section" (grid with client and document data)
        /// </summary>
        /// <param name="newObject">New pdf document item</param>
        /// <param name="rowIndex">Row index which new item will be added in</param>
        /// <param name="columnIndex">Column index which new item will be added in</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>18.03.2021</date>
        public void AddNewItemToCaptionSection(DocumentObject newObject, int rowIndex, int columnIndex)
        {
            AddNewItemToGridCell(CaptionSection, rowIndex, columnIndex, newObject);
        }

        /// <summary>
        /// Add new pdf document item to "Additional data section" (grid with payment data and description of document data)
        /// </summary>
        /// <param name="newObject">New pdf document item</param>
        /// <param name="rowIndex">Row index which new item will be added in</param>
        /// <param name="columnIndex">Column index which new item will be added in</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>18.03.2021</date>
        public void AddNewItemToAdditionalDataSection(DocumentObject newObject, int rowIndex, int columnIndex)
        {
            AddNewItemToGridCell(AdditionalDataSection, rowIndex, columnIndex, newObject);
        }

        /// <summary>
        /// Add new document element to pdf content area
        /// </summary>
        /// <param name="newElement">Pdf document element</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public void AddNewItemToContent(DocumentObject newElement)
        {
            // если это первый добавленный элемент и
            // высота верхнего колонтитула первой страницы больше высоты верхнего колонтитула остальных страниц 
            if (CurrentContent.Elements.Count == 0 && FirstPageHeaderHeight > HeaderHeight)
            {
                // вставляем сначала пустой фрейм
                TextFrame textFrame = CurrentContent.AddTextFrame();
                // устанавливаем ему высоту, равную разнице верхних колонтитулов первой и остальных страниц
                textFrame.Height = string.Format("{0}cm", FirstPageHeaderHeight - HeaderHeight);

                // данная манипуляция нужна, чтобы сместить основной контент ниже и, тем самым,
                // предотвратить перекрытие верхнего колонтитула основным контентом
            }

            // добавляем новый элемент в поле документа
            CurrentContent.Elements.Add(newElement);
        }

        /// <summary>
        /// Add new document object to  the firs page Header or Footer grid
        /// </summary>
        /// <param name="documentArea">Header or Footer</param>
        /// <param name="newObject">Document object for adding</param>
        /// <param name="rowIndex">Row index which new document object will be added in</param>
        /// <param name="columnIndex">Column index which new document object will be added in</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public void AddNewItemToFirsPageHeaderFooterGrid(DocumentArea documentArea, DocumentObject newObject, int rowIndex, int columnIndex)
        {
            switch (documentArea)
            {
                case DocumentArea.Footer:
                    AddNewItemToGridCell(FirstPageFooterGrid, rowIndex, columnIndex, newObject);
                    break;
                case DocumentArea.Header:
                    AddNewItemToGridCell(FirstPageHeaderGrid, rowIndex, columnIndex, newObject);
                    break;
            }
        }

        /// <summary>
        /// Add new document object to Header or Footer grid
        /// </summary>
        /// <param name="documentArea">Header or Footer</param>
        /// <param name="newObject">Document object for adding</param>
        /// <param name="rowIndex">Row index which new document object will be added in</param>
        /// <param name="columnIndex">Column index which new document object will be added in</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public void AddNewItemToHeaderFooterGrid(DocumentArea documentArea, DocumentObject newObject, int rowIndex, int columnIndex)
        {
            switch (documentArea)
            {
                case DocumentArea.Footer:
                    AddNewItemToGridCell(FooterGrid, rowIndex, columnIndex, newObject);
                    break;
                case DocumentArea.Header:
                    AddNewItemToGridCell(HeaderGrid, rowIndex, columnIndex, newObject);
                    break;
            }
        }

        /// <summary>
        /// Add new document object to grid
        /// </summary>
        /// <param name="grid">Grid which new document object will be added in</param>
        /// <param name="rowIndex">Row index which new document object will be added in</param>
        /// <param name="columnIndex">Column index which new document object will be added in</param>
        /// <param name="newItem">Document object for adding</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        private void AddNewItemToGridCell(Table grid, int rowIndex, int columnIndex, DocumentObject newItem)
        {
            // получем ячейку, в которую мы будем добавлять новый элемент
            Cell cell = GetCell(grid, rowIndex, columnIndex);

            // если ячейка получена - добавляем новый элемент
            if (cell != null)
            {
                cell.Elements.Add(newItem);
            }
        }
        #endregion

        #region "Methods for preparing data"       
        /// <summary>
        /// Prepare grid data (with two columns) in according to our company data
        /// </summary>
        /// <param name="firstColumnWidth">The first column width</param>
        /// <param name="secondColumnWidth">The second column width</param>
        /// <param name="propertyFont">The first column font style</param>
        /// <param name="valueFont">The second column font style</param>
        /// <returns>Grid (table without borders) filled with our company data</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public Table PrepareOurCompanyData(double firstColumnWidth, double secondColumnWidth, FontModel propertyFont = null, FontModel valueFont = null)
        {
            // если шрифт для первого столбца не передан - устанавливаем шрифт по умолчанию для первого столбца
            if (propertyFont == null)
            {
                propertyFont = new FontModel() { FontSize = 8, IsItalic = true };
            }
            // если шрифт для второго столбца не передан - устанавливаем шрифт по умолчанию для второго столбца
            if (valueFont == null)
            {
                valueFont = new FontModel() { FontSize = 8 };
            }

            // заполняем список данными о нашей компании
            List<TwoColumnsDataModel> companyData = new List<TwoColumnsDataModel>();

            TwoColumnsDataModel twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("Name") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Saler.Name, Font = valueFont };
            companyData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("Address") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Saler.Address, Font = valueFont };
            companyData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("Principal") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Saler.Principal, Font = valueFont };
            companyData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("TaxNumber") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Saler.TaxNumber, Font = valueFont };
            companyData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("VATNumber") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Saler.VATNumber, Font = valueFont };
            companyData.Add(twoColumns);

            // возвращаем таблицу (без границ) с заполненными данными
            return CreateAndFillGrid(companyData, firstColumnWidth, secondColumnWidth);
        }

        /// <summary>
        /// Prepare grid data (with two columns) in according to client data
        /// </summary>
        /// <param name="firstColumnWidth">The first column width</param>
        /// <param name="secondColumnWidth">The second column width</param>
        /// <param name="rightPadding">Padding on the right table side (cm)</param>
        /// <param name="propertyFont">The first column font style</param>
        /// <param name="valueFont">The second column font style</param>
        /// <returns>Grid (table without borders) filled with client data</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date> 
        public Table PrepareClientData(double firstColumnWidth, double secondColumnWidth, double rightPadding = 1, FontModel propertyFont = null, FontModel valueFont = null)
        {
            // если шрифт для первого столбца не передан - устанавливаем шрифт по умолчанию для первого столбца
            if (propertyFont == null)
            {
                propertyFont = new FontModel() { FontSize = 11, IsItalic = true };
            }
            // если шрифт для второго столбца не передан - устанавливаем шрифт по умолчанию для второго столбца
            if (valueFont == null)
            {
                valueFont = new FontModel() { FontSize = 11 };
            }

            // заполняем список данными о клиенте
            List<TwoColumnsDataModel> companyData = new List<TwoColumnsDataModel>();

            TwoColumnsDataModel twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("Client") + ":",
                Font = new FontModel() { FontSize = 11, FontColor = new ColorModel(57, 186, 236), IsBold = true }
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Client.Name,
                Font = new FontModel() { FontSize = 11, FontColor = new ColorModel(), IsBold = true }
            };
            companyData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("Address") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Client.Address, Font = valueFont };
            companyData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("Principal") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Client.Principal, Font = valueFont };
            companyData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("TaxNumber") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Client.TaxNumber, Font = valueFont };
            companyData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("VATNumber") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Client.VATNumber, Font = valueFont };
            companyData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel() { Value = TranslationHelper.GetLocalizedString("Phone") + ":", Font = propertyFont };
            twoColumns.SecondColumnValue = new CellModel() { Value = Client.Phone, Font = valueFont };
            companyData.Add(twoColumns);

            // создаём таблицу (без границ) с заполненными данными
            Table table = CreateAndFillGrid(companyData, firstColumnWidth, secondColumnWidth);
            // устанавливаем отступ справа
            table.RightPadding = string.Format("{0}cm", rightPadding);

            return table;
        }

        /// <summary>
        /// Prepare grid data (with two columns) in according to document data
        /// </summary>
        /// <param name="firstColumnWidth">The first column width</param>
        /// <param name="secondColumnWidth">The second column width</param>
        /// <param name="propertyFont">The first column font style</param>
        /// <param name="valueFont">The second column font style</param>
        /// <returns>Grid (table without borders) filled with document data</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public Table PrepareDocumentData(double firstColumnWidth, double secondColumnWidth, FontModel propertyFont = null, FontModel valueFont = null)
        {
            // если шрифт для первого столбца не передан - устанавливаем шрифт по умолчанию для первого столбца
            if (propertyFont == null)
            {
                propertyFont = new FontModel() { FontSize = 10 };
            }
            // если шрифт для второго столбца не передан - устанавливаем шрифт по умолчанию для второго столбца
            if (valueFont == null)
            {
                valueFont = new FontModel() { FontSize = 10, IsBold = true };
            }

            // заполняем список данными о документе
            List<TwoColumnsDataModel> documentData = new List<TwoColumnsDataModel>();

            TwoColumnsDataModel twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = Document.DocumentName.ToUpper(),
                Font = new FontModel() { FontSize = 30, FontColor = new ColorModel(), IsBold = true },
                MergeDirection = MergeDirection.Right,
                MergeCellCount = 2,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = Enums.VerticalAlignment.Top,
            };
            twoColumns.SecondColumnValue = new CellModel() { Value = "", Font = valueFont };
            documentData.Add(twoColumns);

            if (!string.IsNullOrEmpty(Document.DocumentDescription) && string.IsNullOrEmpty(Document.SourceDocumentNumber))
            {
                twoColumns = new TwoColumnsDataModel();
                twoColumns.FirstColumnValue = new CellModel()
                {
                    Value = Document.DocumentDescription,
                    Font = valueFont,
                    MergeDirection = MergeDirection.Right,
                    MergeCellCount = 2,
                    HorizontalAlignment = HorizontalAlignment.Right
                };
                twoColumns.SecondColumnValue = new CellModel() { Value = "", Font = valueFont };
                documentData.Add(twoColumns);
            }

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = Document.DocumentAuthenticity == DocumentAuthenticity.Unknown ? "" : TranslationHelper.GetLocalizedString(Document.DocumentAuthenticity.ToString()).ToUpper(),
                Font = new FontModel() { FontSize = 16, FontColor = new ColorModel(57, 186, 236) },
                MergeDirection = MergeDirection.Right,
                MergeCellCount = 2,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            twoColumns.SecondColumnValue = new CellModel() { Value = "", Font = valueFont };
            documentData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("Number") + ":",
                Font = propertyFont
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Document.DocumentNumber,
                Font = valueFont,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            documentData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("Date") + ":",
                Font = propertyFont
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Document.DocumentDate.ToShortDateString(),
                Font = valueFont,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            documentData.Add(twoColumns);

            if (!string.IsNullOrEmpty(Document.SourceDocumentNumber))
            {
                twoColumns = new TwoColumnsDataModel();
                twoColumns.FirstColumnValue = new CellModel()
                {
                    Value = Document.DocumentDescription
                    + " № "
                    + Document.SourceDocumentNumber
                    + " "
                    + TranslationHelper.GetLocalizedString("From")
                    + " "
                    + Document.SourceDocumentDate.ToShortDateString(),
                    Font = valueFont,
                    MergeDirection = MergeDirection.Right,
                    MergeCellCount = 2
                };
                twoColumns.SecondColumnValue = new CellModel() { Value = "", Font = valueFont };
                documentData.Add(twoColumns);
            }

            // возвращаем таблицу (без границ) с заполненными данными
            return CreateAndFillGrid(documentData, firstColumnWidth, secondColumnWidth);
        }

        /// <summary>
        /// Prepare table in according to goods data table
        /// </summary>
        /// <param name="columnsWidth">Table columns width array</param>
        /// <param name="tableVisualization">Table visualization data</param>
        /// <returns>Table filled with document data</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public Table PrepareOperationData(double[] columnsWidth, TableVisualizationModel tableVisualization = null)
        {
            // если модель, описывающая параметры визуализации таблицы, пуста - создаём модель с параметрами по умолчанию
            if (tableVisualization == null)
            {
                tableVisualization = new TableVisualizationModel();
                tableVisualization.BorderLocation = BordersOption.NoBorder;

                tableVisualization.HeaderBackground = new ColorModel(255, 255, 255);
                tableVisualization.HeaderFont.FontColor = new ColorModel(147, 149, 152);
                tableVisualization.HeaderFont.IsItalic = true;
                tableVisualization.HeaderFont.FontSize = 9;

                tableVisualization.DifferentEvenRowsBackground = true;
                tableVisualization.ContentBackground = new ColorModel(137, 207, 240);
                tableVisualization.ContentFont.FontSize = 9;
                tableVisualization.EvenRowsBackground = new ColorModel(255, 255, 255);
            }

            // возвращаем таблицу с заполненными данными
            return CreateAndFillTable(DocumentData, columnsWidth, tableVisualization);
        }

        /// <summary>
        /// Prepare grid data (with two columns) in according to VAT data
        /// </summary>
        /// <param name="firstColumnWidth">The first column width</param>
        /// <param name="secondColumnWidth">The second column width</param>
        /// <param name="horizontalAlignment">Horizontal alignment content</param>
        /// <param name="borderStyle">Border style</param>
        /// <param name="bordersOption">Borders option</param>
        /// <param name="propertyFont">The first column font style</param>
        /// <param name="valueFont">The second column font style</param>
        /// <returns>Grid (table without borders) filled with VAT data</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>18.03.2021</date>
        public Table PrepareVATData(double firstColumnWidth, double secondColumnWidth, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Right, Border borderStyle = null, BordersOption bordersOption = BordersOption.InsideHorizontalBorder, FontModel propertyFont = null, FontModel valueFont = null)
        {
            // если шрифт для первого столбца не передан - устанавливаем шрифт по умолчанию для первого столбца
            if (propertyFont == null)
            {
                propertyFont = new FontModel() { FontSize = 9 };
            }
            // если шрифт для второго столбца не передан - устанавливаем шрифт по умолчанию для второго столбца
            if (valueFont == null)
            {
                valueFont = new FontModel() { FontSize = 9, IsBold = true };
            }
            // если объект класса, описывающего границы таблицы, отсутствует -
            // создаем объект данного класса с параметрами по умолчанию шрифт 
            if (borderStyle == null)
            {
                borderStyle = new Border()
                {
                    Width = 0.5,
                    Visible = true,
                    Color = MigraDoc.DocumentObjectModel.Color.FromRgb(137, 207, 240)
                };
            }

            // заполняем список данными об НДС
            List<TwoColumnsDataModel> vATData = new List<TwoColumnsDataModel>();
            TwoColumnsDataModel twoColumns;
            // переменная для хранения общей суммы по всем товарам
            double totalSum = 0;
            foreach (VATModel model in VATs)
            {
                // формируем запись с данными о базовой сумме налогообложения
                twoColumns = new TwoColumnsDataModel();
                twoColumns.FirstColumnValue = new CellModel()
                {
                    Value = TranslationHelper.GetLocalizedString("VATBase"),
                    Font = propertyFont,
                    TopPadding = 0.1,
                    BottomPadding = 0.1
                };
                twoColumns.SecondColumnValue = new CellModel()
                {
                    Value = model.VATBase.ToString(),
                    Font = valueFont,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    TopPadding = 0.1,
                    BottomPadding = 0.1
                };
                vATData.Add(twoColumns);

                // формируем данные с размером налога
                twoColumns = new TwoColumnsDataModel();
                twoColumns.FirstColumnValue = new CellModel()
                {
                    Value = TranslationHelper.GetLocalizedString("VAT") + "\t" + model.VATRate.ToString() + "%",
                    Font = propertyFont,
                    TopPadding = 0.1,
                    BottomPadding = 0.1
                };
                twoColumns.SecondColumnValue = new CellModel()
                {
                    Value = model.VATSum.ToString(),
                    Font = valueFont,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    TopPadding = 0.1,
                    BottomPadding = 0.1
                };
                vATData.Add(twoColumns);

                totalSum += model.VATBase;
            }

            // формируем таблицу с данными по НДС
            Table vATDataTable = CreateAndFillGrid(vATData, firstColumnWidth, secondColumnWidth, borderStyle, bordersOption);
            // добавляем в данную таблицу итоговую строку
            Row totalSumRow = vATDataTable.AddRow();
            // определяем смещение таблицы от левого края, чтобы установить выравнивание таблицы по горизонтали
            double gap = 0;
            switch (horizontalAlignment)
            {
                case HorizontalAlignment.Left:
                    gap = -0.15;
                    break;
                case HorizontalAlignment.Center:
                    gap = (PageWidth - LeftIndentation - RightIndentation - firstColumnWidth - secondColumnWidth) / 2;
                    break;
                case HorizontalAlignment.Right:
                    gap = PageWidth - LeftIndentation - RightIndentation - firstColumnWidth - secondColumnWidth - 0.15;
                    break;
            }
            // устанавливаем выравнивание таблицы по вертикали
            vATDataTable.Rows.LeftIndent = string.Format("{0}cm", gap);
            // устанавливаем белый шрифт для текста итоговой строки
            propertyFont.FontColor = new ColorModel(255, 255, 255);
            valueFont.FontColor = new ColorModel(255, 255, 255);
            valueFont.FontSize = 11;
            // добавляем данные по итоговой строке
            totalSumRow.Cells[0].Elements.Add(CreateParagraph(PrepareParagraphText(TranslationHelper.GetLocalizedString("PaymentSum")), propertyFont));
            totalSumRow.Cells[1].Elements.Add(CreateParagraph(PrepareParagraphText(totalSum), valueFont, HorizontalAlignment.Right));
            totalSumRow.Shading.Color = borderStyle.Color;
            return vATDataTable;

        }

        /// <summary>
        /// Prepare grid data (with two columns) in according to payment data
        /// </summary>
        /// <param name="firstColumnWidth">The first column width</param>
        /// <param name="secondColumnWidth">The second column width</param>
        /// <param name="propertyFont">The first column font style</param>
        /// <param name="valueFont">The second column font style</param>
        /// <returns>Grid (table without borders) filled with payment data</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>18.03.2021</date>
        public Table PreparePaymentData(double firstColumnWidth, double secondColumnWidth, FontModel propertyFont = null, FontModel valueFont = null)
        {
            // если шрифт для первого столбца не передан - устанавливаем шрифт по умолчанию для первого столбца
            if (propertyFont == null)
            {
                propertyFont = new FontModel() { FontSize = 9, FontColor = new ColorModel(147, 149, 152), IsItalic = true };
            }
            // если шрифт для второго столбца не передан - устанавливаем шрифт по умолчанию для второго столбца
            if (valueFont == null)
            {
                valueFont = new FontModel() { FontSize = 9 };
            }

            // заполняем список данными по оплате
            List<TwoColumnsDataModel> paymentData = new List<TwoColumnsDataModel>();

            TwoColumnsDataModel twoColumns;
            if (!string.IsNullOrEmpty(Saler.StoreName))
            {
                twoColumns = new TwoColumnsDataModel();
                twoColumns.FirstColumnValue = new CellModel()
                {
                    Value = TranslationHelper.GetLocalizedString("Object") + ":",
                    Font = propertyFont,
                    TopPadding = 0.1,
                    BottomPadding = 0.1
                };
                twoColumns.SecondColumnValue = new CellModel()
                {
                    Value = Saler.StoreName,
                    Font = valueFont,
                    TopPadding = 0.1,
                    BottomPadding = 0.1
                };
                paymentData.Add(twoColumns);
            }

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("PaymentType") + ":",
                Font = propertyFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Document.PaymentType,
                Font = valueFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            paymentData.Add(twoColumns);

            if (!string.IsNullOrEmpty(Document.DealReason))
            {
                twoColumns = new TwoColumnsDataModel();
                twoColumns.FirstColumnValue = new CellModel()
                {
                    Value = TranslationHelper.GetLocalizedString("DealReason") + ":",
                    Font = propertyFont,
                    TopPadding = 0.1,
                    BottomPadding = 0.1
                };
                twoColumns.SecondColumnValue = new CellModel()
                {
                    Value = Document.DealReason,
                    Font = valueFont,
                    TopPadding = 0.1,
                    BottomPadding = 0.1
                };
                paymentData.Add(twoColumns);
            }

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("InWords") + ":",
                Font = propertyFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Document.DocumentSum,
                Font = valueFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            paymentData.Add(twoColumns);

            // возвращаем таблицу (без границ) с заполненными данными
            return CreateAndFillGrid(paymentData, firstColumnWidth, secondColumnWidth);
        }

        /// <summary>
        /// Prepare grid data (with two columns) in according to bank data
        /// </summary>
        /// <param name="firstColumnWidth">The first column width</param>
        /// <param name="secondColumnWidth">The second column width</param>
        /// <param name="propertyFont">The first column font style</param>
        /// <param name="valueFont">The second column font style</param>
        /// <returns>Grid (table without borders) filled with bank data</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>18.03.2021</date>
        public Table PrepareBankData(double firstColumnWidth, double secondColumnWidth, FontModel propertyFont = null, FontModel valueFont = null)
        {
            // если шрифт для первого столбца не передан - устанавливаем шрифт по умолчанию для первого столбца
            if (propertyFont == null)
            {
                propertyFont = new FontModel() { FontSize = 9, FontColor = new ColorModel(147, 149, 152), IsItalic = true };
            }
            // если шрифт для второго столбца не передан - устанавливаем шрифт по умолчанию для второго столбца
            if (valueFont == null)
            {
                valueFont = new FontModel() { FontSize = 9 };
            }

            // заполняем список данными по банку
            List<TwoColumnsDataModel> bankData = new List<TwoColumnsDataModel>();

            TwoColumnsDataModel twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("Bank") + ":",
                Font = propertyFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Saler.BankName,
                Font = valueFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            bankData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = "BIC:",
                Font = propertyFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Saler.BIC,
                Font = valueFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            bankData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = "IBAN:",
                Font = propertyFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Saler.IBAN,
                Font = valueFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            bankData.Add(twoColumns);

            // возвращаем таблицу (без границ) с заполненными данными
            return CreateAndFillGrid(bankData, firstColumnWidth, secondColumnWidth);
        }

        /// <summary>
        /// Prepare grid data (with two columns) in according to deal description data
        /// </summary>
        /// <param name="firstColumnWidth">The first column width</param>
        /// <param name="secondColumnWidth">The second column width</param>
        /// <param name="propertyFont">The first column font style</param>
        /// <param name="valueFont">The second column font style</param>
        /// <returns>Grid (table without borders) filled with deal description data</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>18.03.2021</date>
        public Table PrepareDealDescriptionData(double firstColumnWidth, double secondColumnWidth, FontModel propertyFont = null, FontModel valueFont = null)
        {
            // если шрифт для первого столбца не передан - устанавливаем шрифт по умолчанию для первого столбца
            if (propertyFont == null)
            {
                propertyFont = new FontModel() { FontSize = 9, FontColor = new ColorModel(147, 149, 152), IsItalic = true };
            }
            // если шрифт для второго столбца не передан - устанавливаем шрифт по умолчанию для второго столбца
            if (valueFont == null)
            {
                valueFont = new FontModel() { FontSize = 9 };
            }

            // заполняем список данными по описанию сделки
            List<TwoColumnsDataModel> dealDescriptionData = new List<TwoColumnsDataModel>();

            TwoColumnsDataModel twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("TaxDate") + ":",
                Font = propertyFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Document.TaxDate.ToShortDateString(),
                Font = valueFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            dealDescriptionData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("DealPlace") + ":",
                Font = propertyFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Document.DealPlace,
                Font = valueFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            dealDescriptionData.Add(twoColumns);

            twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("DealDescription") + ":",
                Font = propertyFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = Document.DealDescription,
                Font = valueFont,
                TopPadding = 0.1,
                BottomPadding = 0.1
            };
            dealDescriptionData.Add(twoColumns);

            // возвращаем таблицу (без границ) с заполненными данными
            return CreateAndFillGrid(dealDescriptionData, firstColumnWidth, secondColumnWidth);
        }

        /// <summary>
        /// Prepare info about the creator and recipient of the document
        /// </summary>
        /// <param name="font">Font style</param>
        /// <returns>Frame filled with deal description data</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>18.03.2021</date>
        public TextFrame PrepareSignatureData(FontModel font = null)
        {
            // если шрифт не задан - устанавливаем шрифт по умолчанию
            if (font == null)
            {
                font = new FontModel() { FontSize = 10 };
            }

            // заполняем список данными о том, кто подготовил документ и кто его получит
            List<TwoColumnsDataModel> signatureData = new List<TwoColumnsDataModel>();

            TwoColumnsDataModel twoColumns = new TwoColumnsDataModel();
            twoColumns.FirstColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("ReceivedBy") + ": " + Document.ReceivedBy,
                Font = font
            };
            twoColumns.SecondColumnValue = new CellModel()
            {
                Value = TranslationHelper.GetLocalizedString("CreatedBy") + ": " + Document.CreatedBy,
                Font = font
            };
            signatureData.Add(twoColumns);

            // определяем ширину колонки как половина ширины страницы (без учёта левой и правой границ)
            double columnWidth = (PageWidth - LeftIndentation - RightIndentation) / 2;
            // определяем высоту букв в миллиметрах
            float fontSizeInMm = SizeHelper.ConvertPointToMillimeter((float)font.FontSize);
            // находим максимальную длину записи о создателе/получатале 
            int maxNoteLenght = Math.Max(twoColumns.FirstColumnValue.Value.Length, twoColumns.SecondColumnValue.Value.Length);
            // определяем кол-во строк, необходимое для отображения контента 
            int necessaryRowsCount = (int)Math.Ceiling(fontSizeInMm * maxNoteLenght / (columnWidth * 10));

            // создаём новый фрейм
            TextFrame signatureFrame = new TextFrame();
            // устанавливаем ему расположение внизу страницы
            signatureFrame.RelativeVertical = RelativeVertical.Margin;
            signatureFrame.RelativeHorizontal = RelativeHorizontal.Margin;
            signatureFrame.Top = ConvertVerticalAlignmentToShapePosition(Enums.VerticalAlignment.Bottom);
            // устанавливаем ему высоту, необходимую для отображения всех строк контенкта
            signatureFrame.Height = string.Format(
                "{0}mm",
                fontSizeInMm * necessaryRowsCount
                + (FooterHeight > BottomIndentation ? ((FooterHeight - BottomIndentation) * 10) : 0));
            // вставляем в данный фрейм таблицу (без границ) с данными о создателе и получателе документа
            signatureFrame.Add(CreateAndFillGrid(signatureData, columnWidth, columnWidth));

            // вставляем в конец документа пустой фрейм высотой, равной высоте фрейма с данными о создателе и получателе документа.
            // Это нужно для того, чтобы зарезервировать место на случай, если фрейм с данными займет больше одной строки и
            // поднимется вверх (так как нижнюю границу мы зафиксировали), тем самым накладываясь на последниt строки нашего документа
            AddNewItemToContent(new TextFrame() { Height = signatureFrame.Height });

            return signatureFrame;
        }
        #endregion
    }
}
