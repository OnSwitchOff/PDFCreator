using System;
using System.Collections.Generic;
using System.Text;

namespace PDFCreator.Models
{
    public class CompanyModel
    {
        #region "Constructors"
        public CompanyModel()
        {

        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Company name
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string Name { get; set; }

        /// <summary>
        /// Company address
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string Address { get; set; }

        /// <summary>
        /// Company principal
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string Principal { get; set; }

        /// <summary>
        /// Company phone
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string Phone { get; set; }

        /// <summary>
        /// Company tax number
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string TaxNumber { get; set; }

        /// <summary>
        /// Company VAT number
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string VATNumber { get; set; }

        /// <summary>
        /// Company bank name
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string BankName { get; set; }

        /// <summary>
        /// Company BIC
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string BIC { get; set; }

        /// <summary>
        /// Company IBAN
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>12.03.2021</date>
        public string IBAN { get; set; }

        /// <summary>
        /// Store name in which goods were sold
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>30.03.2021</date>
        public string StoreName { get; set; }
        #endregion
    }
}
