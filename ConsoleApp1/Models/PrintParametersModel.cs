using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Models
{
    public class PrintParametersModel
    {
        #region "Declarations"
        // перечень подключенных к данному ПК принтеров
        private List<string> installedPrinters;
        // выбранный пользователем принтер
        private string selectedPrinter;
        // перечень форматов бумаги
        private List<ComboBoxObject> pageFormats;
        // выбранный пользователем формат бумаги
        private ComboBoxObject selectedPageFormat;
        // перечень вариантов ориентации бумаги
        private List<ComboBoxObject> pageOrientations;
        // выбранный пользователем вариант ориентации бумаги
        private ComboBoxObject selectedPageOrientation;
        // оступ слева от края листа 
        private double leftMargin;
        // оступ сверху от края листа
        private double topMargin;
        // оступ справа от края листа
        private double rightMargin;
        // оступ снизу от края листа
        private double bottomMargin;
        // кол-во копий документа
        private int countCopies;
        #endregion

        #region "Constructor"
        /// <summary>
        /// Создаёт экземпляр класса и заполняет свойства данными, полученными от PrintService
        /// </summary>
        /// <param name="printService">Объект класса, обеспечивающего печать документов</param>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public PrintParametersModel()
        {
            installedPrinters = new List<string>();
            installedPrinters.Add("Printer A");
            installedPrinters.Add("Printer B");
            installedPrinters.Add("Printer C");
            SelectedPrinter = "Printer A";
            //foreach (string printerName in printService.InstalledPrinters)
            //{
            //    InstalledPrinters.Add(printerName);
            //}

            //if (InstalledPrinters.Count > 0)
            //{
            //    SelectedPrinter = installedPrinters[0];
            //}

            ComboBoxObject pageFormat;
            pageFormat = new ComboBoxObject("A3", "A3");
            PageFormats.Add(pageFormat);
            pageFormat = new ComboBoxObject("A4", "A4");
            PageFormats.Add(pageFormat);
            pageFormat = new ComboBoxObject("A5", "A5");
            PageFormats.Add(pageFormat);
            pageFormat = new ComboBoxObject("B5", "B5");
            PageFormats.Add(pageFormat);
            //foreach (PageFormat format in Enum.GetValues(typeof(PageFormat)))
            //{
            //    pageFormat = new ComboBoxObject(format.ToString(), format);
            //    PageFormats.Add(pageFormat);

            //    if (format == printService.PageFormat)
            //    {
            //        SelectedPageFormat = pageFormat;
            //    }
            //}

            ComboBoxObject pageOrientation;
            pageOrientation = new ComboBoxObject("Portrait", "Portrait");
            PageOrientations.Add(pageOrientation);
            pageOrientation = new ComboBoxObject("Landscape", "Landscape");
            PageOrientations.Add(pageOrientation);
            //foreach (PageOrientation orientation in Enum.GetValues(typeof(PageOrientation)))
            //{
            //    pageOrientation = new ComboBoxObject("string" + orientation.ToString(), orientation);
            //    PageOrientations.Add(pageOrientation);

            //    if(orientation == printService.PageOrientation)
            //    {
            //        SelectedPageOrientation = pageOrientation;
            //    }
            //}

            LeftMargin = 0.5;
            TopMargin = 0;
            RightMargin = 0;
            BottomMargin = 0;

            CountCopies = 1; // printService.CountCopies;
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Перечень подключенных к данному ПК принтеров
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public List<string> InstalledPrinters
        {
            get => installedPrinters == null ? installedPrinters = new List<string>() : installedPrinters;
            set => installedPrinters = value;
        }

        /// <summary>
        /// Выбранный пользователем принтер
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public string SelectedPrinter
        {
            get => selectedPrinter;
            set => selectedPrinter = value;
        }

        /// <summary>
        /// Перечень доступных форматов бумаги
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public List<ComboBoxObject> PageFormats
        {
            get => pageFormats == null ? pageFormats = new List<ComboBoxObject>() : pageFormats;
            set => pageFormats = value;
        }

        /// <summary>
        /// Выбранный пользователем формат бумаги
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public ComboBoxObject SelectedPageFormat
        {
            get => selectedPageFormat == null ? selectedPageFormat = new ComboBoxObject() : selectedPageFormat;
            set => selectedPageFormat = value;
        }

        /// <summary>
        /// Перечень вариантов ориентации бумаги
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public List<ComboBoxObject> PageOrientations
        {
            get => pageOrientations == null ? pageOrientations = new List<ComboBoxObject>() : pageOrientations;
            set => pageOrientations = value;
        }

        /// <summary>
        /// Выбранный пользователем вариант ориентации бумаги
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public ComboBoxObject SelectedPageOrientation
        {
            get => selectedPageOrientation == null ? selectedPageOrientation = new ComboBoxObject() : selectedPageOrientation;
            set => selectedPageOrientation = value;
        }

        /// <summary>
        /// Оступ слева от края листа 
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public double LeftMargin
        {
            get => leftMargin;
            set => leftMargin = value;
        }

        /// <summary>
        /// Оступ сверху от края листа 
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public double TopMargin
        {
            get => topMargin;
            set => topMargin = value;
        }

        /// <summary>
        /// Оступ справа от края листа 
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public double RightMargin
        {
            get => rightMargin;
            set => rightMargin = value;
        }

        /// <summary>
        /// Оступ снизу от края листа 
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public double BottomMargin
        {
            get => bottomMargin;
            set => bottomMargin = value;
        }

        /// <summary>
        /// Кол-во печатаемых копий документа
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>02.04.2021</date>
        public int CountCopies
        {
            get => countCopies;
            set => countCopies = value;
        }
        #endregion
    }
}
