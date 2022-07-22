using PDFCreator.Models;
using PDFCreator.Enums;
using System;
using System.Data;
using ConsoleApp1.Models;

namespace ConsoleApp1
{
    class Program
    {
        [Obsolete]
        static void Main(string[] args)
        {
            PartnerModel partner = new PartnerModel();
            partner.Company = "ММБ 7 ЕООД";
            partner.Address = "София - ж.к. Дружба бл. 221 вх. Б эт. 4 ап. 37"; 
            partner.Principal = "Мария Михайлова Борисова";
            partner.TaxNumber = "203456185";
            partner.Phone = "0896712690";
            partner.VATNumber = "";

            DocumentFactory documentFactory = new DocumentFactory();
            documentFactory.CustomerData = partner;
            string imagePath = @"C:\Users\serhii.rozniuk\Desktop\unlock.png";
            documentFactory.LogoPath = imagePath;

            documentFactory.DocumentDescription.DocumentName = "Фактура";
            documentFactory.DocumentDescription.DocumentDescription = "към фактура";
            documentFactory.DocumentDescription.DocumentNumber = "0000120053";
            documentFactory.DocumentDescription.SourceDocumentNumber = "0000000008";
            documentFactory.DocumentDescription.SourceDocumentDate = new DateTime(2021, 3, 29);
            documentFactory.DocumentDescription.DocumentDate = new DateTime(2021, 2, 5);
            documentFactory.DocumentDescription.DocumentAuthenticity = DocumentAuthenticity.Original;
            documentFactory.DocumentDescription.PaymentType = "Превод по сметка";
            documentFactory.DocumentDescription.DocumentSum = "Сто и три лв. и 43 ст.";
            documentFactory.DocumentDescription.DealReason = "Продажба";
            documentFactory.DocumentDescription.DealPlace = "Центр. офис София";
            documentFactory.DocumentDescription.DealDescription = "";
            documentFactory.DocumentDescription.TaxDate = new DateTime(2021, 2, 5);
            documentFactory.DocumentDescription.CreatedBy = "Стефка Стойчева";
            documentFactory.DocumentDescription.ReceivedBy = "Мария Михайлова Борисова";

            //documentFactory.GoodsTable.Columns.Add("Кол.", typeof(double));
            //documentFactory.GoodsTable.Columns.Add("Мярка", typeof(string));
            //documentFactory.GoodsTable.Columns.Add("Описание на стоката или услугата", typeof(string));
            //documentFactory.GoodsTable.Columns.Add("Цена", typeof(double));
            //documentFactory.GoodsTable.Columns.Add("Стойност", typeof(double));
            //for (int i = 1; i < 39; i++)
            //{
            //    documentFactory.GoodsTable.Rows.Add(i, "бр.", "Актуализация на програмен продукт - 1 бр. актуализационен код", 86.19, 86.19);
            //}

            documentFactory.ReportTable.Columns.Add("Col1", typeof(int));
            documentFactory.ReportTable.Columns.Add("Col2", typeof(string));
            documentFactory.ReportTable.Columns.Add("Col3", typeof(double));
            documentFactory.ReportTable.Columns.Add("Col4", typeof(DateTime));
            Random random = new Random();

            for(int i = 0; i < 20; i++)
            {
                documentFactory.ReportTable.Rows.Add(i + 1, string.Format("Item{0}", i + 1), random.NextDouble() * 1000, DateTime.Now);
            }

            documentFactory.GenerateDocument(Enums.DocumentType.Receipt, Enums.DocumentVersionPrinting.Copy, Enums.PaymentTypes.Cash);
            
            documentFactory.SaveDocument(System.IO.Path.Combine(@"C:\Users\serhii.rozniuk\Desktop", "NewPdf.pdf"));

            //try
            //{
            //    System.IO.File.Delete(imagePath);
            //}
            //catch (Exception ex)
            //{

            //}

            Console.WriteLine("Hello World!");
        }
    }
}
