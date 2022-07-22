using System;
using System.Collections.Generic;
using System.Text;

namespace PDFCreator.Models
{
    public class VATModel
    {
        #region "Constructor"
        /// <summary>
        /// Create VATModel object with the properties equel "0".
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public VATModel()
        {
            VATRate = 0f;
            VATBase = 0.0;
            VATSum = 0.0;
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// VAT rate
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public float VATRate { get; set; }

        /// <summary>
        /// VAT base for VAT sum calculation
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public double VATBase { get; set; }

        /// <summary>
        /// VAT sum
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>17.03.2021</date>
        public double VATSum { get; set; }
        #endregion
    }
}
