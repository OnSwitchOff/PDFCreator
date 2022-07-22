using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using ConsoleApp1.Enums;
using ConsoleApp1.Models;
using PDFCreator;
using PDFCreator.Enums;
using PDFCreator.Models;
using HorizontalAlignment = PDFCreator.Enums.HorizontalAlignment;


namespace ConsoleApp1
{
    public class DocumentFactory
    {
        #region "Declarations"
        // класс, генерирующий pdf документ
        private MicroinvestPdfDocument pdfDocument;
        // параметры страницы
        private PrintParametersModel pageParameters;
        // высота верхнего колонтитула первой страницы
        private double firstPageHeaderHeight;
        // путь к изображению с логотипом компании
        private string logoPath;
        // путь к изображению с реквизитами компании
        private string signaturePath;
        // изображение с логотипом компании
        private System.Drawing.Image logo;
        // изображение с реквизитами компании
        private System.Drawing.Image signature;
        // таблица с данными отчёта
        private DataTable reportTable;
        // ширина колонок таблицы отчётов
        private double[] reportTableColumnsWidth;
        // ширина колонок таблицы товаров в печатных формах
        private double[] goodsTableColumnsWidth;
        #endregion

        #region "Constructor"
        /// <summary>
        /// Создаёт экземпляр класса MicroinvestPdfDocument и 
        /// устанавливает ему флаг использования на первой странице отличного от остальных верхнего и нижнего колонтитулов.
        /// Устанавливает логотип и реквизиты компании по умолчанию
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public DocumentFactory()
        {
            // создаём экземпляр класса MicroinvestPdfDocument, который будет генерировать Pdf документ
            pdfDocument = new MicroinvestPdfDocument();
            pdfDocument.Orientation = PageOrientation.Portrait;
            pdfDocument.PageFormat = PageFormat.A3;
            // устанавливаем флаг, что верхний и нижний колонтитулы первой страницы будут отлисными от остальных страниц
            pdfDocument.DifferentFirstPageHeaderFooter = true;
            // устанавливаем кол-во знаков после запятой при указании цены !!!взять из CurrentCultureHelper
            pdfDocument.PriceFormat = 2;
            // устанавливаем кол-во знаков после запятой при указании количества !!!взять из CurrentCultureHelper
            pdfDocument.QuantityFormat = 3;

            // запоминаем высоту верхнего колонтитула первой страницы, чтобы иметь возможность восстановить её в дальнейшем
            firstPageHeaderHeight = pdfDocument.FirstPageHeaderHeight;


            // устанавливаем путь к изображению с логотипом фирмы !!!!!! получить путь к изображению из базы данных !!!!!!!!!!!!!!!
            logoPath = string.Empty;
            // устанавливаем путь к изображению с реквизитами фирмы !!!!!! получить путь к изображению из базы данных !!!!!!!!!!!!!!!
            signaturePath = string.Empty;
            // устанавливаем логотип и изображение с реквизитами компании по умолчанию
            logo = pdfDocument.DefaultHeaderImage;
            signature = pdfDocument.DefaultFooterImage;

            // устанавливаем данные по нашей компании !!!! грузим данные из базы !!!!!!!!!!!!!!!!
            pdfDocument.Saler.Name = "Микроинвест ООД";
            pdfDocument.Saler.Address = "София бул. Цар Борис III 215";
            pdfDocument.Saler.Principal = "Елена Ширкова";
            pdfDocument.Saler.TaxNumber = "831826092";
            pdfDocument.Saler.VATNumber = "BG831826092";
            pdfDocument.Saler.BankName = "ПроКредит Банк България ЕАД";
            pdfDocument.Saler.BIC = "PRCBBGSF";
            pdfDocument.Saler.IBAN = "BG12PRCB92301015430322";
            pdfDocument.Saler.StoreName = "Служебен обект";
            // устанавливаем данные по группам НДС !!!! грузим данные из базы !!!!!!!!!!!!!!!
            pdfDocument.VATs.Add(new VATModel() { VATRate = 20, VATBase = 86.19, VATSum = 17.24 });
            pdfDocument.VATs.Add(new VATModel() { VATRate = 9, VATBase = 86.19, VATSum = 17.24 });
            pdfDocument.VATs.Add(new VATModel() { VATRate = 10, VATBase = 86.19, VATSum = 17.24 });
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Таблица с данными, необходимыми для построения отчёта
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public DataTable ReportTable
        {
            get => pdfDocument.DocumentData == null ? pdfDocument.DocumentData = new DataTable() : pdfDocument.DocumentData;
            set => pdfDocument.DocumentData = value;
            //get => reportTable == null ? reportTable = new DataTable() : reportTable;
            //set => reportTable = value;
        }

        /// <summary>
        /// Ширина колонок таблицы с данными, необходимыми для построения отчёта
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public double[] ReportTableColumnsWidth
        {
            get
            {
                // если массив с шириной колонок таблицы не сформирован - формируем его с данными по умолчанию
                if (reportTableColumnsWidth == null)
                {
                    reportTableColumnsWidth = new double[ReportTable.Columns.Count > 0 ? ReportTable.Columns.Count : 1];
                    double defaultWidth = (pdfDocument.PageWidth - pdfDocument.LeftIndentation - pdfDocument.RightIndentation) / reportTableColumnsWidth.Length;
                    for (int i = 0; i < reportTableColumnsWidth.Length; i++)
                    {
                        reportTableColumnsWidth[i] = defaultWidth;
                    }
                }

                return reportTableColumnsWidth;
            }
            set => reportTableColumnsWidth = value;
        }


        /// <summary>
        /// Таблицы с данными о товарах, участвующих в документе
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public DataTable GoodsTable
        {
            get => pdfDocument.DocumentData == null ? pdfDocument.DocumentData = new DataTable() : pdfDocument.DocumentData;
            set => pdfDocument.DocumentData = value;
        }

        /// <summary>
        /// Ширина колонок таблицы с данными о товарах
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public double[] GoodsTableColumnsWidth 
        { 
            get
            {
                // если массив с шириной колонок таблицы не сформирован - формируем его с данными по умолчанию
                if (goodsTableColumnsWidth == null)
                {
                    goodsTableColumnsWidth = new double[GoodsTable.Columns.Count > 0 ? GoodsTable.Columns.Count : 1];
                    double defaultWidth = (pdfDocument.PageWidth - pdfDocument.LeftIndentation - pdfDocument.RightIndentation) / goodsTableColumnsWidth.Length;
                    for(int i = 0; i < goodsTableColumnsWidth.Length; i++)
                    {
                        goodsTableColumnsWidth[i] = defaultWidth;
                    }
                    //GoodsTableColumnsWidth = new double[] { 1.5, 1.5, (pdfDocument.PageWidth - pdfDocument.LeftIndentation - pdfDocument.RightIndentation - 7), 2, 2 };
                }

                return goodsTableColumnsWidth;
            }
            set => goodsTableColumnsWidth = value; 
        }        

        /// <summary>
        /// Путь к изображению с логотипом фирмы
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public string LogoPath
        {
            get => logoPath;
            set => SetImageValues(ref logoPath, value, ref logo);
        }

        /// <summary>
        /// Путь к изображению с реквизитами фирмы
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public string SignaturePath
        {
            get => signaturePath;
            set => SetImageValues(ref signaturePath, value, ref signature);
        }

        /// <summary>
        /// Информация о покупателе
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public PartnerModel CustomerData
        {
            set => pdfDocument.Client = (PDFCreator.Models.CompanyModel)value;
        }

        /// <summary>
        /// Информация о формируемом документе
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public DocumentModel DocumentDescription
        {
            get => pdfDocument.Document;
            set => pdfDocument.Document = value;
        }

        /// <summary>
        /// Информация о параметрах страницы документа
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        public PrintParametersModel PageParameters
        {
            get => pageParameters == null ? pageParameters = new PrintParametersModel() : pageParameters;
            set
            {
                pageParameters = value;

                pdfDocument.LeftIndentation = pageParameters.LeftMargin;
                pdfDocument.TopIndentation = pageParameters.TopMargin;
                pdfDocument.RightIndentation = pageParameters.RightMargin;
                pdfDocument.BottomIndentation = pageParameters.BottomMargin;
                pdfDocument.Orientation = (PageOrientation)Enum.Parse(typeof(PageOrientation), pageParameters.SelectedPageOrientation.Value.ToString());
                pdfDocument.PageFormat = (PageFormat)Enum.Parse(typeof(PageFormat), pageParameters.SelectedPageFormat.Value.ToString());
            }
        }
        #endregion

        #region "Public methods"
        /// <summary>
        /// Сгенерировать разметку документа
        /// </summary>
        /// <param name="documentType">Тип документа</param>
        /// <param name="versionPrinting">Версия документа для печати (оригинал, копия и т.п.)</param>
        /// <param name="paymentTypes">Тип оплаты по текущему документу</param>
        /// <param name="operationType">Тип операции</param>
        /// <returns>true - если разметка удачно сгенерирована, иначе false</returns>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        [Obsolete]
        public bool GenerateDocument(DocumentType documentType, DocumentVersionPrinting versionPrinting = DocumentVersionPrinting.Original, PaymentTypes paymentTypes = PaymentTypes.Cash, OperationType operationType = OperationType.Sale)
        {
            // очищаем предыдущую разметку (если таковая осталась с предыдущего раза) 
            pdfDocument.Clear();
            try
            {
                // если необходимо разметить отчёт
                if (documentType == DocumentType.Report)
                {
                    // обнуляем "задел", выделенный для верхнего колонтитула первой страницы
                    pdfDocument.FirstPageHeaderHeight = 0;
                    // устанавливаем стиль для визуализации таблицы
                    TableVisualizationModel tableVisualization = new TableVisualizationModel();
                    tableVisualization.HeaderBackground.Color_R = 30;
                    tableVisualization.HeaderBackground.Color_G = 144;
                    tableVisualization.HeaderBackground.Color_B = 255;
                    tableVisualization.HeaderFont.IsBold = true;
                    tableVisualization.DifferentEvenRowsBackground = true;
                    // формируем разметку таблицы
                    pdfDocument.AddNewItemToContent(pdfDocument.CreateAndFillTable(ReportTable, ReportTableColumnsWidth, tableVisualization));
                }
                else // иначе
                {
                    // восстанавливаем высоту "задела", выделенного для верхнего колонтитула первой страницы
                    pdfDocument.FirstPageHeaderHeight = firstPageHeaderHeight;
                    // размечаем верхний и нижний колонтитулы
                    pdfDocument.CreateDefaultHeaderFooterGrid();
                    // добавляем в ячейку 0Х0 верхнего колонтитула первой страницы изображение с логотипом нашей фирмы
                    pdfDocument.AddNewItemToFirsPageHeaderFooterGrid(
                        DocumentArea.Header,
                        pdfDocument.CreateImageObject(
                            logo, pdfDocument.GetColumnWidth(pdfDocument.FirstPageHeaderGrid, 0),
                            HorizontalAlignment.Left,
                            pdfDocument.FirstPageHeaderHeight),
                        0,
                        0);
                    // добавляем в ячейку 0х1 верхнего колонтитула первой страницы реквизиты нашей фирмы
                    pdfDocument.AddNewItemToFirsPageHeaderFooterGrid(
                        DocumentArea.Header,
                        pdfDocument.PrepareOurCompanyData(
                            pdfDocument.GetColumnWidth(pdfDocument.FirstPageHeaderGrid, 1) / 3,
                            pdfDocument.GetColumnWidth(pdfDocument.FirstPageHeaderGrid, 1) / 3 * 2),
                        0,
                        1);
                    // добавляем в ячейку 0х1 нижнего колонтитула первой страницы изображение 
                    pdfDocument.AddNewItemToFirsPageHeaderFooterGrid(
                        DocumentArea.Footer,
                        pdfDocument.CreateImageObject(
                            signature, pdfDocument.GetColumnWidth(pdfDocument.FirstPageFooterGrid, 1),
                            HorizontalAlignment.Right,
                            pdfDocument.FooterHeight),
                        0,
                        1);
                    // в ячейке 0х0 нижнего колонтитула будет находиться нумерация строк (вставлена при выполнении функции "CreateDefaultHeaderFooterGrid")
                    // добавляем в ячейку 0х1 нижнего колонтитула изображение 
                    pdfDocument.AddNewItemToHeaderFooterGrid(
                        DocumentArea.Footer,
                        pdfDocument.CreateImageObject(
                            signature, pdfDocument.GetColumnWidth(pdfDocument.FirstPageFooterGrid, 1),
                            HorizontalAlignment.Right,
                            pdfDocument.FooterHeight),
                        0,
                        1);
                    // формируем тело документа в зависимости от выбранных пользователем параметров
                    switch (versionPrinting)
                    {
                        case DocumentVersionPrinting.Original:
                            GenerateDocumentBody(DocumentAuthenticity.Original, documentType, operationType, paymentTypes);
                            break;
                        case DocumentVersionPrinting.Copy:
                            GenerateDocumentBody(DocumentAuthenticity.Copy, documentType, operationType, paymentTypes);
                            break;
                        case DocumentVersionPrinting.OriginalAndCopy:
                            // формируем оригинал документа
                            GenerateDocumentBody(DocumentAuthenticity.Original, documentType, operationType, paymentTypes);
                            // добавляем секцию для следующего экземпляра документа
                            pdfDocument.AddSection();
                            // формируем копию документа
                            GenerateDocumentBody(DocumentAuthenticity.Copy, documentType, operationType, paymentTypes);
                            break;
                        case DocumentVersionPrinting.OriginalAndTwoCopies:
                            GenerateDocumentBody(DocumentAuthenticity.Original, documentType, operationType, paymentTypes);
                            pdfDocument.AddSection();
                            GenerateDocumentBody(DocumentAuthenticity.Copy, documentType, operationType, paymentTypes);
                            pdfDocument.AddSection();
                            GenerateDocumentBody(DocumentAuthenticity.Copy, documentType, operationType, paymentTypes);
                            break;
                        default:
                            break;
                    }
                }

                // генерируем разметку
                pdfDocument.RenderDocument();

                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Сохранить документ по указанному пути
        /// </summary>
        /// <param name="path">Путь, по которому будет сохранён документ</param>
        /// <returns>true - если документ сохранён, иначе false</returns>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        [Obsolete]
        public bool SaveDocument(string path)
        {
            // если документ размечен - сохраняем его и возвращаем true
            if (pdfDocument.IsRenderedDocument)
            {
                pdfDocument.Save(path);

                return true;
            }

            // иначе возвращаем false
            return false;
        }

        /// <summary>
        /// Преобразовать страницы документа в изображения
        /// </summary>
        /// <returns>Коллекция страниц документа</returns>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        [Obsolete]
        public List<System.Drawing.Image> ConvertDocumentToImage()
        {
            // если документ размечен - возвращаем коллекцию с изображениями страниц
            if(pdfDocument.IsRenderedDocument)
            {
                return pdfDocument.ConvertPdfToImage();
            }
            
            // иначе возвращаем пустую коллекцию
            return new List<System.Drawing.Image>();
        }
        #endregion

        #region "Private methods"
        /// <summary>
        /// Сгенерировать тело документа
        /// </summary>
        /// <param name="documentType">Тип документа</param>
        /// <param name="documentAuthenticity">Версия документа (оригинал или копия)</param>
        /// <param name="paymentTypes">Тип оплаты по текущему документу</param>
        /// <param name="operationType">Тип операции</param>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        [Obsolete]
        private void GenerateDocumentBody(DocumentAuthenticity documentAuthenticity, DocumentType documentType, OperationType operationType, PaymentTypes paymentTypes)
        {
            // создаём и размечаем блок с данными о клиенте и документе
            pdfDocument.CreateCaptionSection();
            // добаляем данный блок в тело документа
            pdfDocument.AddNewItemToContent(pdfDocument.CaptionSection);
            // добавляем отступ
            pdfDocument.AddNewItemToContent(pdfDocument.CreateSpace(1.3));
            // добавляем в тело документа таблицу с товарами
            pdfDocument.AddNewItemToContent(pdfDocument.PrepareOperationData(GoodsTableColumnsWidth));
            // добавляем в тело документа таблицу с разбивкой по таварам в зависимости от групп НДС
            pdfDocument.AddNewItemToContent(pdfDocument.PrepareVATData(4, 2));
            // создаём и размечаем блок с информацией об оплате и описании документа
            pdfDocument.CreateAdditionalDataSection();
            // добавляем в тело данный блок 
            pdfDocument.AddNewItemToContent(pdfDocument.AdditionalDataSection);
            // добавляем в тело документа информацию о том, кто составил данный документ и кто его получит 
            pdfDocument.AddNewItemToContent(pdfDocument.PrepareSignatureData());


            pdfDocument.Document.DocumentName = documentType.ToString(); // взять со словаря!!!
            pdfDocument.Document.DocumentDate = DateTime.Now;
            pdfDocument.Document.DocumentAuthenticity = documentAuthenticity;
            pdfDocument.Document.PaymentType = paymentTypes.ToString(); // взять со словаря!!!
            pdfDocument.Document.DealReason = operationType.ToString(); // взять со словаря!!!

            switch (documentType)
            {
                case DocumentType.Invoice:
                    pdfDocument.Document.DocumentDescription = "";
                    pdfDocument.Document.SourceDocumentNumber = "";
                    pdfDocument.Document.SourceDocumentDate = DateTime.Now;
                    break;
                case DocumentType.DebitNote:
                    break;
                case DocumentType.CreditNote:
                    break;
                case DocumentType.ProformInvoice:
                    pdfDocument.Document.DocumentDescription = "";
                    pdfDocument.Document.SourceDocumentNumber = "";
                    pdfDocument.Document.SourceDocumentDate = DateTime.Now;
                    break;
                case DocumentType.Receipt:
                    pdfDocument.Document.DocumentDescription = "за продажба на стоки"; // взять со словаря в зависимости от типа операции!!!
                    pdfDocument.Document.SourceDocumentNumber = "";
                    pdfDocument.Document.SourceDocumentDate = DateTime.Now;

                    pdfDocument.Document.DocumentAuthenticity = DocumentAuthenticity.Unknown;
                    break;
                default:
                    break;
            }


            // вставляем в ячейку 0Х0 CaptionSection информацию о клиенте
            pdfDocument.AddNewItemToCaptionSection(
                pdfDocument.PrepareClientData(
                    pdfDocument.GetColumnWidth(
                        pdfDocument.CaptionSection, 0) / 4,
                    pdfDocument.GetColumnWidth(
                        pdfDocument.CaptionSection, 0) / 4 * 3
                    ),
                0,
                0);
            // вставляем в ячейку 0Х1 CaptionSection информацию о документе
            pdfDocument.AddNewItemToCaptionSection(
                pdfDocument.PrepareDocumentData(
                    pdfDocument.GetColumnWidth(
                        pdfDocument.CaptionSection, 1) / 2,
                    pdfDocument.GetColumnWidth(
                        pdfDocument.CaptionSection, 1) / 2
                    ),
                0,
                1);

            // вставляем в ячейку 0Х0 AdditionalDataSection информацию об оплате
            pdfDocument.AddNewItemToAdditionalDataSection(
                pdfDocument.PreparePaymentData(
                    pdfDocument.GetColumnWidth(pdfDocument.AdditionalDataSection, 0) / 2,
                    pdfDocument.GetColumnWidth(pdfDocument.AdditionalDataSection, 0) / 2),
                0,
                0
                );
            // если данный документ не является квитанцией
            if (documentType != DocumentType.Receipt)
            {
                // вставляем в ячейку 1Х0 AdditionalDataSection информацию о банке
                pdfDocument.AddNewItemToAdditionalDataSection(
                    pdfDocument.PrepareBankData(
                        pdfDocument.GetColumnWidth(pdfDocument.AdditionalDataSection, 0) / 4,
                        pdfDocument.GetColumnWidth(pdfDocument.AdditionalDataSection, 0) / 4 * 3),
                    1,
                    0
                    );
                // вставляем в ячейку 1Х1 AdditionalDataSection описание по текущей сделке 
                pdfDocument.AddNewItemToAdditionalDataSection(
                    pdfDocument.PrepareDealDescriptionData(
                        pdfDocument.GetColumnWidth(pdfDocument.AdditionalDataSection, 1) / 2,
                        pdfDocument.GetColumnWidth(pdfDocument.AdditionalDataSection, 1) / 2),
                    1,
                    1
                    );
            }
        }

        /// <summary>
        /// Установить значения в поля, отвечающие за изображения (напр., логотип, реквизиты и т.п.)
        /// </summary>
        /// <param name="property">Поле, в которое необходимо записать путь к изображению</param>
        /// <param name="imagePath">Путь к изображению</param>
        /// <param name="image">Объект с изображение</param>
        /// <developer>Сергей Рознюк</developer>
        /// <date>19.04.2021</date>
        private void SetImageValues(ref string property, string imagePath, ref System.Drawing.Image image)
        {
            // если файл существует
            if (System.IO.File.Exists(imagePath))
            {
                // получаем разрешение файла
                int lastPointIndex = imagePath.LastIndexOf('.');
                string fileFormat = string.Empty;
                if (lastPointIndex > 0)
                {
                    fileFormat = imagePath.Substring(lastPointIndex + 1, imagePath.Length - (lastPointIndex + 1)).ToLower();
                }

                // если данный файл является изображением - сохраняем параметры в соответствующих полях
                if (!string.IsNullOrEmpty(fileFormat) && (fileFormat.Equals("png") || fileFormat.Equals("jpg") || fileFormat.Equals("jpeg") || fileFormat.Equals("bmp")))
                {
                    property = imagePath;
                    var imageBytes = System.IO.File.ReadAllBytes(imagePath);
                    System.IO.MemoryStream ms = new System.IO.MemoryStream(imageBytes);
                    image = System.Drawing.Image.FromStream(ms);
                }
                else // иначе "бросаем исключение" для информирования о некорректном формате файла
                {
                    throw new Exception("Incorrect image format!!!");
                }
            }
            else // иначе "бросаем исключение" для информирования об отсутствии файла по указанному пути
            {
                throw new Exception("File not exist!!!");
            }
        }
        #endregion
    }
}
