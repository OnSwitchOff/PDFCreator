using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PDFCreator.Enums;
using PDFCreator.Helpers;
using PDFCreator.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using Docnet.Core;
using Docnet.Core.Converters;
using Docnet.Core.Models;
using HorizontalAlignment = PDFCreator.Enums.HorizontalAlignment;
using VerticalAlignment = PDFCreator.Enums.VerticalAlignment;
using PageFormat = PDFCreator.Enums.PageFormat;


namespace PDFCreator
{
    public class PdfDocument
    {
        #region "Declarations"
        // current document
        private Document document;
        //// current section of document
        //private List<Section> sections;
        // convert document into Pdf
        private PdfDocumentRenderer pdfRenderer;
        // the page sizes used to design the page template 
        private PageFormat templatePageFormat;
        // the width used to design the page template 
        private double templateWidth;
        // the height used to design the page template 
        private double templateHeight;
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
        /// - HeaderHeight, FirstPageHeaderHeight, FooterHeight - 0;
        /// - IsUseSectionNumbers - true;
        /// - PriceFormat, QuantityFormat - 2.
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        [Obsolete]
        public PdfDocument() : this(PageFormat.A4, PageOrientation.Portrait, 1, 1, 2, 1, false)
        {

        }

        /// <summary>
        /// Create a blank pdf document with the following page setupes: 
        /// - PageFormat - A4;
        /// - Orientation - portrait;
        /// - TopIndentation, BottomIndentation, RightIndentation - 1;
        /// - LeftIndentation - 2;
        /// - TemplatePageFormat - A4;
        /// - HeaderHeight, FirstPageHeaderHeight, FooterHeight - 0;
        /// - IsUseSectionNumbers - true;
        /// - PriceFormat, QuantityFormat - 2.
        /// </summary>
        /// <param name="differentFirstPageHeaderFooter">The first page header and/or footer is different from others</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        [Obsolete]
        public PdfDocument(bool differentFirstPageHeaderFooter) : this(PageFormat.A4, PageOrientation.Portrait, 1, 1, 2, 1, differentFirstPageHeaderFooter)
        {

        }

        /// <summary>
        /// Create a blank pdf document with the following property:
        /// - TemplatePageFormat - A4;
        /// - HeaderHeight, FirstPageHeaderHeight, FooterHeight - 0;
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
        [Obsolete]
        public PdfDocument(PageFormat pageFormat, PageOrientation orientation, double topMargin, double bottomMargin, double leftMargin, double rightMargin, bool differentFirstPageHeaderFooter)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            document = new Document();
            document.AddSection();

            TemplatePageFormat = PageFormat.A4;
            PageFormat = pageFormat;
            Orientation = orientation;
            if (Orientation == PageOrientation.Portrait)
            {
                TopIndentation = topMargin;
                BottomIndentation = bottomMargin;
                LeftIndentation = leftMargin;
                RightIndentation = rightMargin;
            }
            else
            {
                TopIndentation = leftMargin;
                BottomIndentation = rightMargin;
                LeftIndentation = bottomMargin;
                RightIndentation = topMargin;
            }

            DifferentFirstPageHeaderFooter = differentFirstPageHeaderFooter;

            FirstPageHeaderHeight = 0;
            HeaderHeight = 0;
            FooterHeight = 0;
            IsUseSectionNumbers = true;
            PriceFormat = 2;
            QuantityFormat = 2;
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Standard page sizes
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public PageFormat PageFormat
        {
            get => (PageFormat)Enum.Parse(typeof(PageFormat), CurrentContent.PageSetup.PageFormat.ToString());
            set => CurrentContent.PageSetup.PageFormat = (MigraDoc.DocumentObjectModel.PageFormat)Enum.Parse(typeof(MigraDoc.DocumentObjectModel.PageFormat), value.ToString());
        }

        /// <summary>
        /// Specifies the page orientation
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public PageOrientation Orientation
        {
            get => (PageOrientation)Enum.Parse(typeof(PageOrientation), CurrentContent.PageSetup.Orientation.ToString());
            set => CurrentContent.PageSetup.Orientation = (Orientation)Enum.Parse(typeof(Orientation), value.ToString());
        }

        /// <summary>
        /// Indent from the top of the page (cm)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public double TopIndentation
        {
            get => document.DefaultPageSetup.HeaderDistance.Centimeter;
            set
            {
                document.DefaultPageSetup.HeaderDistance = string.Format("{0}cm", value);
                CurrentContent.PageSetup.TopMargin = string.Format("{0}cm", value);
            }
        }

        /// <summary>
        /// Indent from the bottom of the page (cm)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public double BottomIndentation
        {
            get => document.DefaultPageSetup.FooterDistance.Centimeter;
            set
            {
                document.DefaultPageSetup.FooterDistance = string.Format("{0}cm", value);
                CurrentContent.PageSetup.BottomMargin = string.Format("{0}cm", value);
            }
        }

        /// <summary>
        /// Indent from the left side of the page (cm)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public double LeftIndentation
        {
            get => CurrentContent.PageSetup.LeftMargin.Centimeter;
            set => CurrentContent.PageSetup.LeftMargin = string.Format("{0}cm", value);
        }

        /// <summary>
        /// Indent from the right side of the page (cm)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public double RightIndentation
        {
            get => CurrentContent.PageSetup.RightMargin.Centimeter;
            set => CurrentContent.PageSetup.RightMargin = string.Format("{0}cm", value);
        }

        /// <summary>
        /// The first page header and/or footer is different from others
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public bool DifferentFirstPageHeaderFooter
        {
            get => CurrentContent.PageSetup.DifferentFirstPageHeaderFooter;
            set => CurrentContent.PageSetup.DifferentFirstPageHeaderFooter = value;
        }

        /// <summary>
        /// Width of the page (cm)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public double PageWidth
        {
            get
            {
                Unit width, height;

                PageSetup.GetPageSize((MigraDoc.DocumentObjectModel.PageFormat)Enum.Parse(typeof(MigraDoc.DocumentObjectModel.PageFormat), PageFormat.ToString()), out width, out height);

                if (Orientation == PageOrientation.Portrait)
                {
                    return width.Centimeter;
                }
                else
                {
                    return height.Centimeter;
                }
            }

        }

        /// <summary>
        /// Height of the page (cm)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public double PageHeight
        {
            get
            {
                Unit width, height;

                PageSetup.GetPageSize((MigraDoc.DocumentObjectModel.PageFormat)Enum.Parse(typeof(MigraDoc.DocumentObjectModel.PageFormat), PageFormat.ToString()), out width, out height);

                if (Orientation == PageOrientation.Portrait)
                {
                    return height.Centimeter;
                }
                else
                {
                    return width.Centimeter;
                }
            }
        }

        /// <summary>
        /// First page header
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public HeaderFooter FirstPageHeader
        {
            get => CurrentContent.Headers.FirstPage;
        }

        /// <summary>
        /// First page footer
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public HeaderFooter FirstPageFooter
        {
            get => CurrentContent.Footers.FirstPage;
        }

        /// <summary>
        /// Header
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public HeaderFooter Header
        {
            get => CurrentContent.Headers.Primary;
        }

        /// <summary>
        /// Footer
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public HeaderFooter Footer
        {
            get => CurrentContent.Footers.Primary;
        }

        /// <summary>
        /// Current page content area
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public Section CurrentContent
        {
            get => document.LastSection;
        }

        /// <summary>
        /// The page sizes used to design the page template 
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public PageFormat TemplatePageFormat
        {
            get => templatePageFormat;
            set
            {
                templatePageFormat = value;

                Unit width, height;

                PageSetup.GetPageSize((MigraDoc.DocumentObjectModel.PageFormat)Enum.Parse(typeof(MigraDoc.DocumentObjectModel.PageFormat), value.ToString()), out width, out height);

                templateWidth = width.Centimeter;
                templateHeight = height.Centimeter;
            }
        }

        /// <summary>
        /// Width correction ratio to match the page width selected by the user
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public double WidthRatio
        {
            get => PageWidth / templateWidth;
        }

        /// <summary>
        /// Height correction ratio to match the page height selected by the user
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public double HeightRatio
        {
            get => PageHeight / templateHeight;
        }

        /// <summary>
        /// First page Header area height
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>18.03.2021</date>
        public double FirstPageHeaderHeight { get; set; }

        /// <summary>
        /// Header area height
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public double HeaderHeight
        {
            get => CurrentContent.PageSetup.TopMargin.Centimeter - document.DefaultPageSetup.HeaderDistance.Centimeter;
            set
            {
                // устанавливаем свойству "TopMargin" (расстояние от верха страницы до начала содержимого страницы (не до верхнего колонтитула!!!)
                // значение, равное сумме "HeaderDistance" (расстояние от верха страницы до верхнего колонтитула) и
                // установленной пользователем высоте верхнего колонтитула
                CurrentContent.PageSetup.TopMargin = string.Format("{0}cm", document.DefaultPageSetup.HeaderDistance.Centimeter + value);
                // 
                if (FirstPageHeaderHeight == 0)
                {
                    FirstPageHeaderHeight = HeaderHeight;
                }
            }
        }

        /// <summary>
        /// Footer area height
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public double FooterHeight
        {
            get => CurrentContent.PageSetup.BottomMargin.Centimeter - document.DefaultPageSetup.FooterDistance.Centimeter;
            set => CurrentContent.PageSetup.BottomMargin = string.Format("{0}cm", document.DefaultPageSetup.FooterDistance.Centimeter + value);
        }

        /// <summary>
        /// Document language
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public Language Language
        {
            get => TranslationHelper.Language;
            set => TranslationHelper.Language = value;
        }

        /// <summary>
        /// Each section has its own total page count in the page counter paragraph, not the page count of the entire document
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.04.2021</date>
        public bool IsUseSectionNumbers { get; set; }

        /// <summary>
        /// Number of digits after the decimal point in a field of type decimal 
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>20.04.2021</date>
        public int PriceFormat { get; set; }

        /// <summary>
        /// Number of digits after the decimal point in a field of type double 
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>20.04.2021</date>
        public int QuantityFormat { get; set; }

        /// <summary>
        /// Number of digits after the decimal point in a field of type double 
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>20.04.2021</date>
        public bool IsRenderedDocument
        {
            get => pdfRenderer.PdfDocument != null;
        }
        #endregion

        #region "Methods for add data"
        /// <summary>
        /// Create a new section (a block of data that can be located on the one or more pages)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public void AddSection()
        {
            Section section = CurrentContent.Clone();
            section.Elements.Clear();
            
            document.Add(section);
            if(IsUseSectionNumbers)
            {
                CurrentContent.PageSetup.StartingNumber = 1;
            }
        }

        /// <summary>
        /// Insert Image into table's cell
        /// </summary>
        /// <param name="cell">Cell which will be inserted image in</param>
        /// <param name="imagePath">Image path</param>
        /// <param name="horizontalAlignment">Image horizontal alignment</param>
        /// <param name="maxHeight">Max image height</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public void AddImageToTableCell(Cell cell, string imagePath, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Left, double maxHeight = 0)
        {
            if (cell != null)
            {
                cell.Add(CreateImageObject(imagePath, cell.Column.Width.Centimeter, horizontalAlignment, maxHeight));
            }
        }

        /// <summary>
        /// Add image to the first page Header or Footer
        /// </summary>
        /// <param name="documentArea">Header or Footer</param>
        /// <param name="imagePath">Image path</param>
        /// <param name="imageWidth">Image width</param>
        /// <param name="horizontalAlignment">Image horizontal alignment</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public void AddImageToFirstPageHeaderFooter(DocumentArea documentArea, string imagePath, double imageWidth, HorizontalAlignment horizontalAlignment)
        {
            switch (documentArea)
            {
                case DocumentArea.Header:
                    FirstPageHeader.Add(CreateImageObject(imagePath, imageWidth, horizontalAlignment, FirstPageHeaderHeight));
                    break;
                case DocumentArea.Footer:
                    FirstPageFooter.Add(CreateImageObject(imagePath, imageWidth, horizontalAlignment, FooterHeight));
                    break;
            }
        }

        /// <summary>
        /// Add image to a page Header or Footer
        /// </summary>
        /// <param name="documentArea">Header or Footer</param>
        /// <param name="imagePath">Image path</param>
        /// <param name="imageWidth">Image width</param>
        /// <param name="horizontalAlignment">Image horizontal alignment</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public void AddImageToHeaderFooter(DocumentArea documentArea, string imagePath, double imageWidth, HorizontalAlignment horizontalAlignment)
        {
            switch (documentArea)
            {
                case DocumentArea.Header:
                    Header.Add(CreateImageObject(imagePath, imageWidth, horizontalAlignment, HeaderHeight));
                    break;
                case DocumentArea.Footer:
                    Footer.Add(CreateImageObject(imagePath, imageWidth, horizontalAlignment, FooterHeight));
                    break;
            }
        }
        #endregion

        #region "Methods for create data"
        /// <summary>
        /// Create Image object
        /// </summary>
        /// <param name="imagePath">Image path</param>
        /// <param name="_imageWidth">Image width</param>
        /// <param name="horizontalAlignment">Image horizontal alignment</param>
        /// <param name="maxHeight">Max image height</param>
        /// <remarks>It returns a Paragraph object because it is horizontally aligned</remarks>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public Paragraph CreateImageObject(string imagePath, double _imageWidth, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Left, double maxHeight = 0)
        {
            // создаём объект "Paragraph", который позволит нам устанавливать выравнивание контента по горизонтали
            Paragraph paragraph = new Paragraph();
            // добавляем в параграф изображение
            Image image = paragraph.AddImage(imagePath);

            // если пользователь передал максимальную высоту изображения
            if (maxHeight > 0)
            {
                // получаем реальный размер изображения (в пикселях)
                SizeHelper sizeImage = SizeHelper.GetImageSize(imagePath);

                // преобразуем ширину изображения пикселях в ширину в см
                float imageWidth = SizeHelper.ConvertPixelToCentimeter(sizeImage.Width);
                // преобразуем высоту изображения пикселях в высоту в см
                float imageHeight = SizeHelper.ConvertPixelToCentimeter(sizeImage.Height);

                // определяем соотношение желаемой ширины и реальной ширины изображения
                double widthRatio = _imageWidth / imageWidth;

                // если скорректированная высота изображения получится больше, чем максимально возможная высота
                if (imageHeight * widthRatio > maxHeight)
                {
                    // устанавливаем высоту изображения равной максимально возможной ширине
                    image.Height = string.Format("{0}cm", maxHeight);
                    // устанавливаем горизонтальное выравнивание "по центру"
                    paragraph.Format.Alignment = ParagraphAlignment.Center;
                }
                else // иначе
                {
                    // устанавливаем длину изображения
                    image.Width = string.Format("{0}cm", _imageWidth);
                    // устанавливаем горизонтальное выравнивание согласно полученному от пользователя значению
                    paragraph.Format.Alignment = ConvertParagraphHorizontalAlignment(horizontalAlignment);
                }
            }
            else // иначе если максимальная высота пользователем не передавалась
            {
                // устанавливаем длину изображения
                image.Width = string.Format("{0}cm", _imageWidth);
                // устанавливаем горизонтальное выравнивание согласно полученному от пользователя значению
                paragraph.Format.Alignment = ConvertParagraphHorizontalAlignment(horizontalAlignment);
            }

            // устанавливаем сохранение пропорций изображения
            image.LockAspectRatio = true;
            // устанавливаем вертикальное выравнивание исходя из установленных пользователет отступов
            image.RelativeVertical = RelativeVertical.Margin;
            // устанавливаем горизонтальное выравнивание исходя из параметров колонки
            image.RelativeHorizontal = RelativeHorizontal.Column;

            return paragraph;
        }

        /// <summary>
        /// Create Image object
        /// </summary>
        /// <param name="_image">Image object</param>
        /// <param name="_imageWidth">Image width</param>
        /// <param name="horizontalAlignment">Image horizontal alignment</param>
        /// <param name="maxHeight">Max image height</param>
        /// <remarks>It returns a Paragraph object because it is horizontally aligned</remarks>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>30.03.2021</date>
        public Paragraph CreateImageObject(System.Drawing.Image _image, double _imageWidth, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Left, double maxHeight = 0)
        {
            // создаём объект "Paragraph", который позволит нам устанавливать выравнивание контента по горизонтали
            Paragraph paragraph = new Paragraph();
            Image image;
            // добавляем в параграф изображение
            using (System.Drawing.Image tmpImage = new System.Drawing.Bitmap(_image))
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    tmpImage.Save(stream, _image.GetImageFormat());
                    image = paragraph.AddImage(CreateImageStringFromByteArray(stream.ToArray()));
                }
            }

            // если пользователь передал максимальную высоту изображения
            if (maxHeight > 0)
            {
                // преобразуем ширину изображения пикселях в ширину в см
                float imageWidth = SizeHelper.ConvertPixelToCentimeter(_image.Width);
                // преобразуем высоту изображения пикселях в высоту в см
                float imageHeight = SizeHelper.ConvertPixelToCentimeter(_image.Height);

                // определяем соотношение желаемой ширины и реальной ширины изображения
                double widthRatio = _imageWidth / imageWidth;

                // если скорректированная высота изображения получится больше, чем максимально возможная высота
                if (imageHeight * widthRatio > maxHeight)
                {
                    // устанавливаем высоту изображения равной максимально возможной ширине
                    image.Height = string.Format("{0}cm", maxHeight);
                    // устанавливаем горизонтальное выравнивание "по центру"
                    paragraph.Format.Alignment = ParagraphAlignment.Center;
                }
                else // иначе
                {
                    // устанавливаем длину изображения
                    image.Width = string.Format("{0}cm", _imageWidth);
                    // устанавливаем горизонтальное выравнивание согласно полученному от пользователя значению
                    paragraph.Format.Alignment = ConvertParagraphHorizontalAlignment(horizontalAlignment);
                }
            }
            else // иначе если максимальная высота пользователем не передавалась
            {
                // устанавливаем длину изображения
                image.Width = string.Format("{0}cm", _imageWidth);
                // устанавливаем горизонтальное выравнивание согласно полученному от пользователя значению
                paragraph.Format.Alignment = ConvertParagraphHorizontalAlignment(horizontalAlignment);
            }

            // устанавливаем сохранение пропорций изображения
            image.LockAspectRatio = true;
            // устанавливаем вертикальное выравнивание исходя из установленных пользователет отступов
            image.RelativeVertical = RelativeVertical.Margin;
            // устанавливаем горизонтальное выравнивание исходя из параметров колонки
            image.RelativeHorizontal = RelativeHorizontal.Column;

            return paragraph;
        }


        /// <summary>
        /// Create paragraph with page counter
        /// </summary>
        /// <param name="visualizationTemplate">Вata visualization template</param>
        /// <param name="font">Font style</param>
        /// <remarks>
        /// Examples of the visualizerTemplate: Page {0} of {1}; {0} / {1}; {0} etc., - where:
        /// - {0} - current page;
        /// - {1} - all pages
        /// </remarks>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public Paragraph CreatePageCounterParagraph(string visualizationTemplate, FontModel font = null)
        {
            // создаём объект класса "Paragraph"
            Paragraph paragraph = new Paragraph();
            // если фон пользователем не установлен - устанавливаем фон по умолчанию
            if (font != null) paragraph.Format.Font = (Font)font;

            // "разбиваем" шаблон на фрагменты
            string[] templateComponents = visualizationTemplate.Split(new string[] { "{0}", "{1}" }, StringSplitOptions.None);
            // формируем наполнение параграфа с нумерацией строк в зависимости от полученного шаблона
            switch (templateComponents.Length)
            {
                case 0: // только текущая страница
                    paragraph.AddPageField();
                    break;
                case 1:
                    // если в шаблоне присутствует и текущая страница, и общее кол-во страниц - выводим оба параметра
                    if (visualizationTemplate.Contains("{0}") && visualizationTemplate.Contains("{1}"))
                    {
                        paragraph.AddPageField();
                        paragraph.AddText(templateComponents[0]);
                        if (IsUseSectionNumbers)
                        {
                            paragraph.AddSectionPagesField();
                        }
                        else
                        {
                            paragraph.AddNumPagesField();
                        }
                    }
                    else // иначе
                    {
                        // вставляем сопроводительное слово(-а)
                        paragraph.AddText(templateComponents[0]);
                        // если в шаблоне присутствует текущая страница - добавляем её в объект класса "Paragraph"
                        if (visualizationTemplate.Contains("{0}"))
                        {
                            paragraph.AddPageField();
                        }
                    }
                    break;
                case 2:
                    paragraph.AddText(templateComponents[0]);
                    if (visualizationTemplate.Contains("{0}"))
                    {
                        paragraph.AddPageField();
                    }
                    paragraph.AddText(templateComponents[1]);
                    if (visualizationTemplate.Contains("{1}"))
                    {
                        if (IsUseSectionNumbers)
                        {
                            paragraph.AddSectionPagesField();
                        }
                        else
                        {
                            paragraph.AddNumPagesField();
                        }
                    }
                    break;
                case 3:
                    paragraph.AddText(templateComponents[0]);
                    if (visualizationTemplate.Contains("{0}"))
                    {
                        paragraph.AddPageField();
                    }
                    paragraph.AddText(templateComponents[1]);
                    if (visualizationTemplate.Contains("{1}"))
                    {
                        if (IsUseSectionNumbers)
                        {
                            paragraph.AddSectionPagesField();
                        }
                        else
                        {
                            paragraph.AddNumPagesField();
                        }
                    }
                    paragraph.AddText(templateComponents[2]);
                    break;
            }

            // устанавливаем горизонтальное выравнивание "по центру"
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            return paragraph;
        }

        /// <summary>
        /// Create object with displayed text
        /// </summary>
        /// <param name="text">Displayed text</param>
        /// <param name="font">Font style</param>
        /// <returns>Paragraph with some text</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>09.03.2021</date>
        public Paragraph CreateParagraph(string text, FontModel font = null, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Left)
        {
            // если фон пользователем не установлен - устанавливаем фон по умолчанию
            if (font == null) font = new FontModel();

            // создаём объект класса "Paragraph"
            Paragraph paragraph = new Paragraph();
            // записываем соответствующий текст
            paragraph.Add(new Text(text));
            // устанавливаем шрифт
            paragraph.Format.Font = (Font)font;
            // устанавливаем горизонтальное выравнивание
            paragraph.Format.Alignment = ConvertParagraphHorizontalAlignment(horizontalAlignment);

            return paragraph;
        }

        /// <summary>
        /// Creates an empty table with the appropriate number of columns
        /// </summary>
        /// <param name="columnsWidth">Columns width array</param>
        /// <param name="bordersOption">Borders option</param>
        /// <param name="bordersStyle">Borders style</param>
        /// <param name="headerBackground">Table header background</param>
        /// <param name="headerFont">Table header font style</param>
        /// <param name="columnsNames">Columns name array</param>
        /// <returns>Table</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public Table CreateGridTemplate(double[] columnsWidth, BordersOption bordersOption, Border bordersStyle, ColorModel headerBackground, FontModel headerFont, DataColumnCollection columnsNames = null)
        {
            // если длина массива имен не соответствует длина массива ширина столбцов - уведомляем об этом 
            if (columnsNames != null && columnsWidth.Length != columnsNames.Count)
            {
                throw new Exception("Array length of columns widths does not equal list count of columns names!!!");
            }

            // создаем таблицу
            Table table = new Table();
            table.Style = "Table";
            // если у таблицы нет границ либо установлены все границы - применяем настройки для всей таблицы
            switch (bordersOption)
            {
                case BordersOption.NoBorder:
                    table.Borders.ClearAll();
                    break;
                case BordersOption.AllBorders:
                    table.Borders = new Borders()
                    {
                        Width = bordersStyle.Width,
                        Color = bordersStyle.Color,
                        Style = bordersStyle.Style
                    };
                    break;
            }

            Column column;
            // создаём колонки
            for (int i = 0; i < columnsWidth.Length; i++)
            {
                // добавляем новую колонку
                column = table.AddColumn(string.Format("{0}cm", columnsWidth[i]));
                // если у таблицы есть частично прорисованные вертикальные линии - применяем настройки к колонкам
                switch (bordersOption)
                {
                    case BordersOption.InsideBorders:
                    case BordersOption.InsideVerticalBorder:
                        if (i > 0)
                        {
                            column.Borders.Left = bordersStyle.Clone();
                        }
                        break;
                    case BordersOption.LeftBorder:
                        if (i == 0)
                        {
                            column.Borders.Left = bordersStyle.Clone();
                        }
                        break;
                    case BordersOption.RightBorder:
                        if (i == columnsWidth.Length - 1)
                        {
                            column.Borders.Right = bordersStyle.Clone();
                        }
                        break;
                    case BordersOption.OutsideBorders:
                        if (i == 0)
                        {
                            column.Borders.Left = bordersStyle.Clone();
                        }
                        if (i == columnsWidth.Length - 1)
                        {
                            column.Borders.Right = bordersStyle.Clone();
                        }
                        break;
                }
            }

            // если есть названия колонок таблицы - устанавливаем их
            if (columnsNames != null)
            {
                Row row = table.AddRow();
                // устанавливаем повторение "шапки" на каждой странице
                row.HeadingFormat = true;
                for (int i = 0; i < columnsNames.Count; i++)
                {
                    // устаналиваем цвет заливки "шапки" таблицы;
                    row.Shading.Color = (MigraDoc.DocumentObjectModel.Color)headerBackground;
                    // устанавливаем наименование колонки
                    row.Cells[i].Add(CreateParagraph(columnsNames[i].ColumnName, headerFont, HorizontalAlignment.Center));
                }

                // устанавливаем вертикальное выравнивание "по центру"
                row.VerticalAlignment = ConvertVerticalAlignment(VerticalAlignment.Center);
            }

            return table;
        }

        /// <summary>
        /// Create and fill grid (table with two columns)
        /// </summary>
        /// <param name="data">Data for filling</param>
        /// <param name="firstColumnWidth">The first column width</param>
        /// <param name="secondColumnWidth">The second column width</param>
        /// <param name="borderStyle">Table border style</param>
        /// <param name="bordersOption">Table borders option</param>
        /// <returns>Table</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public Table CreateAndFillGrid(List<TwoColumnsDataModel> data, double firstColumnWidth, double secondColumnWidth, Border borderStyle = null, BordersOption bordersOption = BordersOption.NoBorder)
        {
            // если стиль границ таблицы не передан - создаем таблицу без границ
            if (borderStyle == null)
            {
                borderStyle = new Border() { Width = 0 };
            }

            // создаём пустую таблицу без "шапки"
            Table table = CreateGridTemplate(new double[] { firstColumnWidth, secondColumnWidth }, bordersOption, borderStyle, new ColorModel(), new FontModel());

            Row row;
            for (int i = 0; i < data.Count; i++)
            {
                // если значение хотя бы одной из колонок не задано - извещаем об этом
                if (data[i].FirstColumnValue == null || data[i].SecondColumnValue == null)
                {
                    throw new Exception("One or more values of TwoColumnsDataModel object isn't initialized!!!");
                }

                // создаём строки и заполняем их
                row = table.AddRow();
                // устанавливаем отступы содержимого ячейки от нижнего и верхнего края 
                row.BottomPadding = string.Format("{0}cm", Math.Max(data[i].FirstColumnValue.BottomPadding, data[i].SecondColumnValue.BottomPadding));
                row.TopPadding = string.Format("{0}cm", Math.Max(data[i].FirstColumnValue.TopPadding, data[i].SecondColumnValue.TopPadding));

                // если у таблицы есть частично прорисованные горизонтальные линии - применяем настройки к соответствующим строкам
                SetHorizontalBordersOption(row, i, data.Count - 1, bordersOption, borderStyle);

                // устанавливаем вертикальное выравнивание "по верхнему краю"
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Top;
                // добавляем значение в первую колонку
                row.Cells[0].Add(CreateParagraph(PrepareParagraphText(data[i].FirstColumnValue.Value), data[i].FirstColumnValue.Font, data[i].FirstColumnValue.HorizontalAlignment));
                // объединяем ячейки (если необходимо)
                MergeCell(row.Cells[0], data[i].FirstColumnValue.MergeDirection, data[i].FirstColumnValue.MergeCellCount, 2, data.Count);
                // добавляем значение во вторую колонку
                row.Cells[1].Add(CreateParagraph(PrepareParagraphText(data[i].SecondColumnValue.Value), data[i].SecondColumnValue.Font, data[i].SecondColumnValue.HorizontalAlignment));
                // объединяем ячейки (если необходимо)
                MergeCell(row.Cells[1], data[i].SecondColumnValue.MergeDirection, data[i].SecondColumnValue.MergeCellCount, 2, data.Count);
            }

            return table;
        }

        /// <summary>
        /// Create and fill Table
        /// </summary>
        /// <param name="dataTable">Data table</param>
        /// <param name="columnsWidth">Columns width array</param>
        /// <param name="tableVisualization">Table visualization data</param>
        /// <returns>Table</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public Table CreateAndFillTable(DataTable dataTable, double[] columnsWidth, TableVisualizationModel tableVisualization)
        {
            // если кол-во колонок в таблице и длина списка с длиной колонок не соответствуют друг другу - выбрасываем исключение
            if (dataTable.Columns.Count != columnsWidth.Length)
            {
                throw new Exception("Table columns count doesn't equel columns width array length!!!");
            }

            // создаем объект с описанием параметров отображения линий таблицы
            Border borderStyle = new Border()
            {
                Width = tableVisualization.BorderWidth,
                Color = (MigraDoc.DocumentObjectModel.Color)tableVisualization.BorderColor,
                Style = (MigraDoc.DocumentObjectModel.BorderStyle)Enum.Parse(typeof(MigraDoc.DocumentObjectModel.BorderStyle), tableVisualization.BorderStyle.ToString())
            };


            // создаём пустую таблицу
            Table table = CreateGridTemplate(columnsWidth, tableVisualization.BorderLocation, borderStyle, tableVisualization.HeaderBackground, tableVisualization.HeaderFont, dataTable.Columns);

            Row row;
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                // добавляем новую строку
                row = table.AddRow();
                // устанавливаем отступы содержимого ячейки от нижнего и верхнего края 
                row.BottomPadding = string.Format("{0}cm", tableVisualization.BottomCellPadding);
                row.TopPadding = string.Format("{0}cm", tableVisualization.TopCellPadding);

                // если у таблицы есть частично прорисованные горизонтальные линии - применяем настройки к соответствующим строкам
                SetHorizontalBordersOption(row, i, dataTable.Rows.Count - 1, tableVisualization.BorderLocation, borderStyle);
                // заполняем строку данными
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    // получаем значение текущей ячейки из таблицы-источника
                    var cellValue = dataTable.Rows[i][j];
                    // 
                    row.Cells[j].Add(CreateParagraph(PrepareParagraphText(cellValue), tableVisualization.ContentFont, GetHorizontalAlignmentDependOnObjectType(cellValue)));
                    //RowContentHorizontalAlignment(row, HorizontalAlignmentInAccordingToType(cellValue));
                    table.Rows[i].VerticalAlignment = ConvertVerticalAlignment(VerticalAlignment.Center);

                    if (tableVisualization.DifferentEvenRowsBackground && (i - 1) % 2 == 0)
                    {
                        row.Shading.Color = (Color)tableVisualization.EvenRowsBackground;
                    }
                    else
                    {
                        row.Shading.Color = (Color)tableVisualization.ContentBackground;
                    }
                }
            }

            return table;
        }

        /// <summary>
        /// Create empty TextFrame (to rezerve some space in a document for example)
        /// </summary>
        /// <param name="height">TextFrame height</param>
        /// <returns>TextFrame</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        public TextFrame CreateSpace(double height)
        {
            return new TextFrame() { Height = string.Format("{0}cm", height) };
        }

        /// <summary>
        /// Create image name from image byte array for create image object of PdfSharp library 
        /// </summary>
        /// <param name="imageArray">Image byte array</param>
        /// <returns>Image name</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        protected string CreateImageStringFromByteArray(byte[] imageArray)
        {
            return "base64:" + Convert.ToBase64String(imageArray);
        }
        #endregion

        #region "Method for get data"
        /// <summary>
        /// Get cell from a table
        /// </summary>
        /// <param name="table">Table which get cell from</param>
        /// <param name="rowIndex">Row index for cell search</param>
        /// <param name="columnIndex">Column index for cell search</param>
        /// <returns>Found cell</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public Cell GetCell(Table table, int rowIndex, int columnIndex)
        {
            // если таблица не создана или индекс колонки/строки находится вне диапазона колонок/строк таблицы -
            // уведомляем об этом пользователя
            if (table == null)
            {
                throw new Exception("Table isn't initialized!!!");
            }
            else if (table.Columns.Count - 1 < columnIndex || columnIndex < 0)
            {
                throw new Exception("Column index is outside of existing columns!!!");
            }
            else if (table.Rows.Count - 1 < rowIndex || rowIndex < 0)
            {
                throw new Exception("Row index is outside of existing rows!!!");
            }
            else // иначе возвращаем соответствующую ячейку
            {
                return table.Rows[rowIndex].Cells[columnIndex];
            }
        }

        /// <summary>
        /// Get the width of a specific column
        /// </summary>
        /// <param name="table">The table whose column width we are looking for</param>
        /// <param name="columnIndex">The index of the column you are looking for</param>
        /// <returns>Column width (cm)</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public double GetColumnWidth(Table table, int columnIndex)
        {
            if (table.Columns.Count > columnIndex)
            {
                return table.Columns[columnIndex].Width.Centimeter;
            }

            throw new Exception("Column index is outside table columns range!!!");
        }

        /// <summary>
        /// Get HorizontalAlignment depending on object type
        /// </summary>
        /// <param name="obj">Current object</param>
        /// <returns>ShapePosition</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        protected HorizontalAlignment GetHorizontalAlignmentDependOnObjectType(object obj)
        {
            if (obj == null)
            {
                throw new Exception("Object doesn't initialize!!!");
            }

            string type = obj.GetType().Name;

            switch (type.ToLower())
            {
                case "datetime":
                    return HorizontalAlignment.Center;
                case "int16":
                case "int32":
                case "int64":
                case "single":
                case "double":
                case "decimal":
                    return HorizontalAlignment.Right;
                default:
                    return HorizontalAlignment.Left;
            }
        }
        #endregion

        #region "Methods for converte data"
        /// <summary>
        /// Convert custom HorizontalAlignment to HorizontalAlignment of the PDFSharp library 
        /// </summary>
        /// <param name="horizontalAlignment">Custom horizontal alignment</param>
        /// <returns>ParagraphAlignment</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        protected ParagraphAlignment ConvertParagraphHorizontalAlignment(HorizontalAlignment horizontalAlignment)
        {
            switch (horizontalAlignment)
            {
                case HorizontalAlignment.Left:
                    return ParagraphAlignment.Left;
                case HorizontalAlignment.Center:
                    return ParagraphAlignment.Center;
                case HorizontalAlignment.Right:
                    return ParagraphAlignment.Right;
            }

            return ParagraphAlignment.Left;
        }

        /// <summary>
        /// Convert custom VerticalAlignment to VerticalAlignment of the PDFSharp library 
        /// </summary>
        /// <param name="verticalAlignment">Custom vertical alignment</param>
        /// <returns>VerticalAlignment of the PDFSharp library</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        protected MigraDoc.DocumentObjectModel.Tables.VerticalAlignment ConvertVerticalAlignment(VerticalAlignment verticalAlignment)
        {
            switch (verticalAlignment)
            {
                case VerticalAlignment.Top:
                    return MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Top;
                case VerticalAlignment.Center:
                    return MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                case VerticalAlignment.Bottom:
                    return MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            }

            return MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
        }

        /// <summary>
        /// Convert custom HorizontalAlignment to ShapePosition of the PDFSharp library (horizontal)
        /// </summary>
        /// <param name="horizontalAlignment">Custom horizontal alignment</param>
        /// <returns>ShapePosition</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        protected ShapePosition ConvertHorizontalAlignmentToShapePosition(HorizontalAlignment horizontalAlignment)
        {
            switch (horizontalAlignment)
            {
                case HorizontalAlignment.Left:
                    return ShapePosition.Left;
                case HorizontalAlignment.Center:
                    return ShapePosition.Center;
                case HorizontalAlignment.Right:
                    return ShapePosition.Right;
            }

            return ShapePosition.Undefined;
        }

        /// <summary>
        /// Convert custom HorizontalAlignment to ShapePosition of the PDFSharp library (vertical)
        /// </summary>
        /// <param name="verticalAlignment">Custom vertical alignment</param>
        /// <returns>ShapePosition</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        protected ShapePosition ConvertVerticalAlignmentToShapePosition(VerticalAlignment verticalAlignment)
        {
            switch (verticalAlignment)
            {
                case VerticalAlignment.Top:
                    return ShapePosition.Top;
                case VerticalAlignment.Center:
                    return ShapePosition.Center;
                case VerticalAlignment.Bottom:
                    return ShapePosition.Bottom;
            }

            return ShapePosition.Undefined;
        }
        #endregion

        #region "Other methods"
        /// <summary>
        /// Merge cell
        /// </summary>
        /// <param name="cell">Current cell</param>
        /// <param name="direction">Merge direction</param>
        /// <param name="count">Cell count for merge</param>
        /// <param name="columnsCount">Columns count</param>
        /// <param name="rowsCount">Rows count</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        protected void MergeCell(Cell cell, MergeDirection direction, int count, int columnsCount, int rowsCount)
        {
            switch (direction)
            {
                case MergeDirection.Right:
                    // если объединение происходит в пределах таблицы - производим объединение
                    if (cell.Column.Index + count <= columnsCount)
                    {
                        cell.MergeRight = count - 1;
                    }
                    else // иначе выбрасываем исключение
                    {
                        throw new Exception("The number of cells to be merged is out of the number of columns!!!");
                    }
                    break;
                case MergeDirection.Down:
                    // если объединение происходит в пределах таблицы - производим объединение
                    if (cell.Row.Index + count <= rowsCount)
                    {
                        cell.MergeDown = count - 1;
                    }
                    else // иначе выбрасываем исключение
                    {
                        throw new Exception("The number of cells to be merged is out of the number of rows!!!");
                    }
                    break;
            }
        }

        /// <summary>
        /// Set horizontal row borders
        /// </summary>
        /// <param name="row">Current row</param>
        /// <param name="curRowIndex">Current row index</param>
        /// <param name="lastRowIndex">Last row index</param>
        /// <param name="bordersOption">Borders option</param>
        /// <param name="borderStyle">Border style</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>15.03.2021</date>
        protected void SetHorizontalBordersOption(Row row, int curRowIndex, int lastRowIndex, BordersOption bordersOption, Border borderStyle)
        {
            switch (bordersOption)
            {
                case BordersOption.InsideBorders:
                case BordersOption.InsideHorizontalBorder:
                    if (curRowIndex > 0)
                    {
                        row.Borders.Top = borderStyle.Clone();
                    }
                    break;
                case BordersOption.TopBorder:
                    if (curRowIndex == 0)
                    {
                        row.Borders.Top = borderStyle.Clone();
                    }
                    break;
                case BordersOption.BottomBorder:
                    if (curRowIndex == lastRowIndex)
                    {
                        row.Borders.Bottom = borderStyle.Clone();
                    }
                    break;
                case BordersOption.OutsideBorders:
                    if (curRowIndex == 0)
                    {
                        row.Borders.Top = borderStyle.Clone();
                    }
                    if (curRowIndex == lastRowIndex)
                    {
                        row.Borders.Bottom = borderStyle.Clone();
                    }
                    break;
            }
        }

        /// <summary>
        /// Format object to string in according to object type
        /// </summary>
        /// <param name="obj">Object to format</param>
        /// <returns>Formating string</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>20.04.2021</date>
        protected string PrepareParagraphText(object obj)
        {
            if (obj == null)
            {
                throw new Exception("Object doesn't initialize!!!");
            }

            string type = obj.GetType().Name;

            switch (type.ToLower())
            {
                case "datetime":
                    return ((DateTime)obj).ToShortDateString();
                case "int16":
                case "int32":
                case "int64":
                case "single":
                    return obj.ToString();
                case "double":
                    return double.Parse(obj.ToString()).ToString("F" + QuantityFormat);
                case "decimal":
                    return decimal.Parse(obj.ToString()).ToString("F" + PriceFormat);
                default:
                    if (obj != null)
                    {
                        return obj.ToString();
                    }

                    return string.Empty;
            }
        }

        /// <summary>
        /// Create pdf document in according to added pdf objects
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        [Obsolete]
        public void RenderDocument()
        {
            try
            {
                pdfRenderer.Document = document.Clone();
                pdfRenderer.RenderDocument();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Clear all pdf document
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>30.03.2021</date>
        [Obsolete]
        public void Clear()
        {
            PageSetup currentSetup = CurrentContent.PageSetup;
            document.Sections = new Sections();
            document.Sections.AddSection().PageSetup = currentSetup.Clone();

            pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
        }

        /// <summary>
        /// Clear content area of pdf document
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>30.03.2021</date>
        [Obsolete]
        public void ClearContentArea()
        {
            Section currentSection = CurrentContent.Clone();
            currentSection.Elements.Clear();

            document.Sections = new Sections();
            document.Sections.Add(currentSection);

            pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
        }

        /// <summary>
        /// Save pdf file
        /// </summary>
        /// <param name="path">Path to save file</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        [Obsolete]
        public void Save(string path)
        {
            //// строим документ
            //RenderDocument();
            // если документ занят - бросаем исключение
            if (FileIsLocked(path))
            {
                throw new Exception("File is locked!!!");
            }
            else // иначе сохраняем его
            {
                pdfRenderer.PdfDocument.Save(path);
            }
        }

        /// <summary>
        /// Check if file is locked
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>True if file is locked else false</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>10.03.2021</date>
        public static bool FileIsLocked(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.OpenOrCreate, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Extract images from current pdf document
        /// </summary>
        /// <returns>List of System.Drawing.Image objects</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.03.2021</date>
        [Obsolete]
        public List<System.Drawing.Image> ExtractImagesFromPdfDocument()
        {
            // создаём список изображений
            List<System.Drawing.Image> images = new List<System.Drawing.Image>();

            // создаём объект DocumentRenderer библиотеки PdfSharp на базе нашего документа, построенного с помощью MicraDoc
            DocumentRenderer documentRenderer = new DocumentRenderer(document);
            documentRenderer.PrepareDocument();
            // определяем кол-во страниц
            int docPages = documentRenderer.FormattedDocument.PageCount;

            // создаём PdfDocument библиотеки PdfSharp 
            PdfSharp.Pdf.PdfDocument pdfDocument = new PdfSharp.Pdf.PdfDocument();
            PdfPage page;
            PageInfo pageInfo;
            // запоняем PdfDocument библиотеки PdfSharp  
            for (int i = 1; i <= docPages; i++)
            {
                page = pdfDocument.AddPage();
                pageInfo = documentRenderer.FormattedDocument.GetPageInfo(i);
                page.Width = pageInfo.Width;
                page.Height = pageInfo.Height;
                page.Orientation = pageInfo.Orientation;

                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                    gfx.MUH = PdfFontEncoding.Unicode;

                    documentRenderer.RenderPage(gfx, i);
                }
            }

            // анализируем содержимое страниц документа
            foreach (PdfPage _page in pdfDocument.Pages)
            {
                // Get resources dictionary
                PdfDictionary resources = _page.Elements.GetDictionary("/Resources");
                if (resources != null)
                {
                    // Get external objects dictionary
                    PdfDictionary xObjects = resources.Elements.GetDictionary("/XObject");
                    if (xObjects != null)
                    {
                        ICollection<PdfItem> items = xObjects.Elements.Values;
                        foreach (PdfItem item in items)
                        {
                            PdfReference reference = item as PdfReference;
                            if (reference != null)
                            {
                                PdfDictionary xObject = reference.Value as PdfDictionary;
                                if (xObject != null && xObject.Elements.GetString("/Subtype") == "/Image")
                                {
                                    System.Drawing.Image extractedImage = ExtractImageFromPdf(xObject);
                                    if (extractedImage != null)
                                    {
                                        images.Add(extractedImage);
                                    }
                                }
                            }
                        }

                    }
                }
            }

            return images;
        }

        /// <summary>
        /// Extract images from pdf document
        /// </summary>
        /// <param name="filePath">Pdf document path</param>
        /// <returns>List of System.Drawing.Image objects</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.03.2021</date>
        public static List<System.Drawing.Image> ExtractImagesFromPdfDocument(string filePath)
        {
            // создаём список изображений
            List<System.Drawing.Image> images = new List<System.Drawing.Image>();

            PdfSharp.Pdf.PdfDocument document = PdfReader.Open(filePath);
            foreach (PdfPage page in document.Pages)
            {
                // Get resources dictionary
                PdfDictionary resources = page.Elements.GetDictionary("/Resources");
                if (resources != null)
                {
                    // Get external objects dictionary
                    PdfDictionary xObjects = resources.Elements.GetDictionary("/XObject");
                    if (xObjects != null)
                    {
                        ICollection<PdfItem> items = xObjects.Elements.Values;
                        foreach (PdfItem item in items)
                        {
                            PdfReference reference = item as PdfReference;
                            if (reference != null)
                            {
                                PdfDictionary xObject = reference.Value as PdfDictionary;
                                if (xObject != null && xObject.Elements.GetString("/Subtype") == "/Image")
                                {
                                    System.Drawing.Image extractedImage = ExtractImageFromPdf(xObject);
                                    if (extractedImage != null)
                                    {
                                        images.Add(extractedImage);
                                    }
                                }
                            }
                        }

                    }
                }
            }

            return images;
        }

        /// <summary>
        /// Extract image from pdf image object
        /// </summary>
        /// <param name="image">Pdf image object</param>
        /// <returns>System.Drawing.Image</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.03.2021</date>
        private static System.Drawing.Image ExtractImageFromPdf(PdfDictionary image)
        {
            string filter = image.Elements.GetName("/Filter");
            switch (filter)
            {
                case "/DCTDecode":
                    return ExtractJpegImage(image);
                case "/FlateDecode":
                    return ExtractPngImage(image);
                default:
                    return null;
            }
        }

        /// <summary>
        /// Extract jpeg image from pdf image object
        /// </summary>
        /// <param name="image">Pdf image object</param>
        /// <returns>System.Drawing.Image</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.03.2021</date>
        private static System.Drawing.Image ExtractJpegImage(PdfDictionary image)
        {
            using (MemoryStream stream = new MemoryStream(image.Stream.Value))
            {
                return System.Drawing.Image.FromStream(stream);
            }

            //// Fortunately JPEG has native support in PDF and exporting an image is just writing the stream to a file.
            //byte[] stream = image.Stream.Value;//FileStream fs = new FileStream(
            //System.IO.Path.Combine(@"C:\Users\serhii.rozniuk\Desktop\New folder", String.Format("Image{0}.jpeg", count++)), 
            //    FileMode.Create, 
            //    FileAccess.Write);
            //BinaryWriter bw = new BinaryWriter(fs);
            //bw.Write(stream);
            //bw.Close();
        }

        /// <summary>
        /// Extract png image from pdf image object
        /// </summary>
        /// <param name="image">Pdf image object</param>
        /// <returns>System.Drawing.Image</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.03.2021</date>
        private static System.Drawing.Image ExtractPngImage(PdfDictionary image)
        {
            int width = image.Elements.GetInteger(PdfImage.Keys.Width);
            int height = image.Elements.GetInteger(PdfImage.Keys.Height);

            var canUnfilter = image.Stream.TryUnfilter();
            byte[] decodedBytes;

            if (canUnfilter)
            {
                decodedBytes = image.Stream.Value;
            }
            else
            {
                PdfSharp.Pdf.Filters.FlateDecode flate = new PdfSharp.Pdf.Filters.FlateDecode();
                decodedBytes = flate.Decode(image.Stream.Value);
            }

            int bitsPerComponent = 0;
            while (decodedBytes.Length - ((width * height) * bitsPerComponent / 8) != 0)
            {
                bitsPerComponent++;
            }

            PixelFormat pixelFormat;
            switch (bitsPerComponent)
            {
                case 1:
                    pixelFormat = PixelFormat.Format1bppIndexed;
                    break;
                case 8:
                    pixelFormat = PixelFormat.Format8bppIndexed;
                    break;
                case 16:
                    pixelFormat = PixelFormat.Format16bppArgb1555;
                    break;
                case 24:
                    pixelFormat = PixelFormat.Format24bppRgb;
                    break;
                case 32:
                    pixelFormat = PixelFormat.Format32bppArgb;
                    break;
                case 64:
                    pixelFormat = PixelFormat.Format64bppArgb;
                    break;
                default:
                    throw new Exception("Unknown pixel format " + bitsPerComponent);
            }

            decodedBytes = decodedBytes.Reverse().ToArray();

            System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(width, height, pixelFormat);
            BitmapData bmpData = bmp.LockBits(new System.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height), System.Drawing.Imaging.ImageLockMode.WriteOnly, bmp.PixelFormat);
            int length = (int)Math.Ceiling(width * (bitsPerComponent / 8.0));
            for (int i = 0; i < height; i++)
            {
                int offset = i * length;
                int scanOffset = i * bmpData.Stride;
                Marshal.Copy(decodedBytes, offset, new IntPtr(bmpData.Scan0.ToInt32() + scanOffset), length);
            }
            bmp.UnlockBits(bmpData);
            bmp.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone);


            return bmp;
        }

        /// <summary>
        /// Convert current pdf document to list of images
        /// </summary>
        /// <returns>List of System.Drawing.Image objects</returns>
        /// <remarks>If ImageFormat is absent it will be ImageFormat.Png</remarks>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.03.2021</date>
        [Obsolete]
        public List<System.Drawing.Image> ConvertPdfToImage(ImageFormat imageFormat = null, SizeHelper pageDimensions = null, int dpi = 300)
        {
            // если формат изображения пользователем не передан - изображение будет формата "Png"
            if (imageFormat == null)
            {
                imageFormat = ImageFormat.Png;
            }

            if (pageDimensions == null)
            {
                pageDimensions = new SizeHelper(
                    (ushort)SizeHelper.ConvertCentimeterToPixel((float)PageWidth, dpi),
                    (ushort)SizeHelper.ConvertCentimeterToPixel((float)PageHeight, dpi));
            }

            // создаём список изображений
            List<System.Drawing.Image> images = new List<System.Drawing.Image>();
            using (MemoryStream pdfDocStream = new MemoryStream())
            {
                //// создаём текущий документ
                //RenderDocument();

                // записываем текущий документ в поток
                pdfRenderer.Save(pdfDocStream, false);

                // создаём объект DocLib библиотеки Docnet.Core для работы с документом 
                using (IDocLib docNet = DocLib.Instance)
                {
                    // получаем текущий документ из потока
                    using (var docReader = docNet.GetDocReader(pdfDocStream.ToArray(), new PageDimensions(Math.Min(pageDimensions.Width, pageDimensions.Height), Math.Max(pageDimensions.Width, pageDimensions.Height))))
                    {
                        // определяем кол-во страниц документа
                        int pageCount = docReader.GetPageCount();
                        // проходим по каждой странице
                        for (int i = 0; i < pageCount; i++)
                        {
                            // получаем текущую страницу
                            using (var pageReader = docReader.GetPageReader(i))
                            {
                                // получаем изображения текущей страницы
                                var rawBytes = pageReader.GetImage(new NaiveTransparencyRemover(255, 255, 255), RenderFlags.RenderForPrinting);
                                // запоминаем ширину страницы
                                var width = pageReader.GetPageWidth();
                                // запоминаем высоту страницы
                                var height = pageReader.GetPageHeight();
                                // получаем все символы текущей страницы
                                var characters = pageReader.GetCharacters();

                                // создаём пустое изображение размерами нашей страницы
                                using (var bmp = new System.Drawing.Bitmap(width, height, PixelFormat.Format32bppArgb))
                                {
                                    // добавляем изображения
                                    AddBytes(bmp, rawBytes);
                                    // добавляем текст
                                    DrawRectangles(bmp, characters);
                                    // добавляем изображение в коллекцию
                                    images.Add((System.Drawing.Image)bmp.Clone());
                                }
                            }
                        }
                    }
                }
            }

            return images;
        }

        /// <summary>
        /// Convert pdf document to list of images
        /// </summary>
        /// <returns>List of System.Drawing.Image objects</returns>
        /// <remarks>If ImageFormat is absent it will be ImageFormat.Png</remarks>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.03.2021</date>
        public static List<System.Drawing.Image> ConvertPdfToImage(string filePath, ImageFormat imageFormat = null)
        {
            // если формат изображения пользователем не передан - изображение будет формата "Png"
            if (imageFormat == null)
            {
                imageFormat = ImageFormat.Png;
            }

            // создаём список изображений
            List<System.Drawing.Image> images = new List<System.Drawing.Image>();

            // создаём объект DocLib библиотеки Docnet.Core для работы с документом  
            using (IDocLib docNet = DocLib.Instance)
            {
                // загружаем документ из файла
                using (var docReader = docNet.GetDocReader(filePath, new PageDimensions(2480, 3508)))
                {
                    // определяем кол-во страниц документа
                    int pageCount = docReader.GetPageCount();
                    // проходим по каждой странице
                    for (int i = 0; i < pageCount; i++)
                    {
                        // получаем текущую страницу
                        using (var pageReader = docReader.GetPageReader(i))
                        {
                            // получаем изображения текущей страницы
                            var rawBytes = pageReader.GetImage(new NaiveTransparencyRemover(255, 255, 255), RenderFlags.RenderForPrinting);
                            // запоминаем ширину страницы
                            var width = pageReader.GetPageWidth();
                            // запоминаем высоту страницы
                            var height = pageReader.GetPageHeight();
                            // получаем все символы текущей страницы
                            var characters = pageReader.GetCharacters();

                            // создаём пустое изображение размерами нашей страницы
                            using (var bmp = new System.Drawing.Bitmap(width, height, PixelFormat.Format32bppArgb))
                            {
                                // добавляем изображения
                                AddBytes(bmp, rawBytes);
                                // добавляем текст
                                DrawRectangles(bmp, characters);
                                // добавляем изображение в коллекцию
                                images.Add((System.Drawing.Image)bmp.Clone());
                            }
                        }
                    }
                }
            }

            return images;
        }

        /// <summary>
        /// Add pictures to image object 
        /// </summary>
        /// <param name="bmp">Image which need to add pictures in</param>
        /// <param name="rawBytes">Image bytes</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.03.2021</date>
        private static void AddBytes(System.Drawing.Bitmap bmp, byte[] rawBytes)
        {
            var rect = new System.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height);

            var bmpData = bmp.LockBits(rect, ImageLockMode.WriteOnly, bmp.PixelFormat);
            var pNative = bmpData.Scan0;

            Marshal.Copy(rawBytes, 0, pNative, rawBytes.Length);
            bmp.UnlockBits(bmpData);
        }

        /// <summary>
        /// Draw rectangles in an image and add characters into rectangles
        /// </summary>
        /// <param name="bmp">Image which need to add characters in</param>
        /// <param name="characters">Characters for adding</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>19.03.2021</date>
        private static void DrawRectangles(System.Drawing.Bitmap bmp, IEnumerable<Docnet.Core.Models.Character> characters)
        {
            var pen = new System.Drawing.Pen(System.Drawing.Color.Transparent);

            using (var graphics = System.Drawing.Graphics.FromImage(bmp))
            {
                foreach (var c in characters)
                {
                    var rect = new System.Drawing.Rectangle(c.Box.Left, c.Box.Top, c.Box.Right - c.Box.Left, c.Box.Bottom - c.Box.Top);
                    graphics.DrawRectangle(pen, rect);
                }
            }
        }

        /// <summary>
        /// Convert jpeg image to pdf document and save pdf 
        /// </summary>
        /// <param name="imagePath">Image path</param>
        /// <param name="saveFilePath">Pdf document save path</param>
        /// <param name="pdfDocumentSize">Prefered pdf document size</param>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>23.03.2021</date>
        public static void ConvertJpegImageToPdf(string imagePath, string saveFilePath, SizeHelper pdfDocumentSize = null)
        {
            // если какой-либо из входных параметров имеет неверный формат - выбрасываем исключение
            string pdfFileFormat = saveFilePath.Substring(saveFilePath.Length - 4).ToLower();
            if (!pdfFileFormat.Equals(".pdf"))
            {
                throw new Exception(string.Format("Incorrect pdf file format: \"{0}\"!!!", pdfFileFormat));
            }

            string imageShortFileFormat = imagePath.Substring(imagePath.Length - 4).ToLower();
            if (!pdfFileFormat.Equals(".jpg"))
            {
                throw new Exception(string.Format("Incorrect image file format: \"{0}\"!!!", imageShortFileFormat));
            }

            string imageLongFileFormat = imagePath.Substring(imagePath.Length - 5).ToLower();
            if (!pdfFileFormat.Equals(".jpeg"))
            {
                throw new Exception(string.Format("Incorrect image file format: \"{0}\"!!!", imageLongFileFormat));
            }

            // создаём объект "JpegImage" и записываем в него наше изображение
            Docnet.Core.Editors.JpegImage file = new Docnet.Core.Editors.JpegImage
            {
                Bytes = File.ReadAllBytes(imagePath)
            };

            // получаем реальный размер изображения
            SizeHelper size = SizeHelper.GetImageSize(imagePath);

            // устанавливаем размер нашего pdf документа
            if (pdfDocumentSize != null)
            {
                double widthRatio = pdfDocumentSize.Width / SizeHelper.ConvertPixelToCentimeter(size.Width);
                double heightRatio = pdfDocumentSize.Height / SizeHelper.ConvertPixelToCentimeter(size.Height);
                if (widthRatio > heightRatio)
                {
                    file.Height = pdfDocumentSize.Height;
                    file.Width = (int)(SizeHelper.ConvertPixelToCentimeter(size.Width) * heightRatio);
                }
                else
                {
                    file.Width = pdfDocumentSize.Width;
                    file.Height = (int)(SizeHelper.ConvertPixelToCentimeter(size.Height) * widthRatio);
                }
            }
            else
            {
                file.Width = size.Width;
                file.Height = size.Height;
            }

            // сохраняем pdf документ
            using (IDocLib docLib = DocLib.Instance)
            {
                var bytes = docLib.JpegToPdf(new[] { file });

                File.WriteAllBytes(saveFilePath, bytes);
            }
        }
        #endregion
    }
}
