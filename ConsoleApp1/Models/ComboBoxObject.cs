using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Models
{
    public class ComboBoxObject
    {
        #region "Constructors"
        public ComboBoxObject()
        {
            Text = string.Empty;
            Value = null;
        }

        public ComboBoxObject(string text, object value)
        {
            Text = text;
            Value = value;
        }

        #endregion

        #region "Properties"
        /// <summary>
        /// Текст, который будет отображаться в ComboBox-е
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Дополнительный параметр (параметры), записываемые в элемент ComboBox-а
        /// </summary>
        public object Value { get; set; }
        #endregion

        #region "Overrides methods"
        /// <summary>
        /// Значение, отображаемое пользователю на экране
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Text;
        }
        #endregion
    }
}
