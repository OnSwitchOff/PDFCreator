using System;
using PDFCreator.Enums;

namespace PDFCreator.Models
{
    public class DocumentModel
    {
        #region "Constructor"
        public DocumentModel()
        {
            DocumentName = "";
            DocumentDescription = "";
            DocumentNumber = "1";
            DocumentDate = DateTime.Now;
            DocumentAuthenticity = DocumentAuthenticity.Unknown;
            SourceDocumentNumber = "";
            SourceDocumentDate = DateTime.Now;
            PaymentType = "";
            DealReason = "";
            TaxDate = DateTime.Now;
            DealPlace = "";
            DealDescription = "";
            ReceivedBy = "Unknown";
            CreatedBy = "Unknown";
            DocumentSum = "0";
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Document name
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string DocumentName { get; set; }

        /// <summary>
        /// Additional document info (for example "for sale", "to invoice" etc)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string DocumentDescription { get; set; }

        /// <summary>
        /// Document number
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string DocumentNumber { get; set; }

        /// <summary>
        /// Document date
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public DateTime DocumentDate { get; set; }

        /// <summary>
        /// Document authenticity (original or copy)
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public DocumentAuthenticity DocumentAuthenticity { get; set; }

        /// <summary>
        /// Number of source document for whose will be created current document
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>30.03.2021</date>
        public string SourceDocumentNumber { get; set; }

        /// <summary>
        /// Date of source document for whose will be created current document
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>30.03.2021</date>
        public DateTime SourceDocumentDate { get; set; }

        /// <summary>
        /// Payment type
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string PaymentType { get; set; }

        /// <summary>
        /// Deal reason
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string DealReason { get; set; }

        /// <summary>
        /// Tax date
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public DateTime TaxDate { get; set; }

        /// <summary>
        /// Deal place
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string DealPlace { get; set; }

        /// <summary>
        /// Deal description
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string DealDescription { get; set; }

        /// <summary>
        /// Document recipient
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string ReceivedBy { get; set; }

        /// <summary>
        /// Document creator
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Document sum
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string DocumentSum { get; set; }
        #endregion
    }
}
