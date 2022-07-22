using System;
using System.Collections.Generic;
using System.Text;

namespace PDFCreator.Models
{
    public class TwoColumnsDataModel
    {
        #region "Constructor"
        public TwoColumnsDataModel()
        {

        }
        #endregion

        #region "Properties"
        public CellModel FirstColumnValue { get; set; }
        public CellModel SecondColumnValue { get; set; }
        #endregion
    }
}
