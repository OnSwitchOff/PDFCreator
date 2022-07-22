using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Models
{
    public class CompanyModel 
    {
        #region "Declarations"
        // название фирмы
        private string company;
        // материально-ответственное лицо
        private string principal;
        // город размещения/регистрации фирмы
        private string city;
        // адрес, по которому размещена/зарегистрирована фирма
        private string address;
        // телефон
        private string phone;
        // налоговый номер
        private string taxNumber;
        // номер НДС
        private string vatNumber;
        // номер банковского счета
        private string iBAN;
        #endregion

        #region "Constructor"
        /// <summary>
        /// Заполняет экземпляр пустыми данными
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public CompanyModel()
        {
            Company = string.Empty;
            Principal = string.Empty;
            City = string.Empty;
            Address = string.Empty;
            Phone = string.Empty;
            TaxNumber = string.Empty;
            VATNumber = string.Empty;
            IBAN = string.Empty;
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Название фирмы
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public string Company
        {
            get => company;
            set => company = value;
        }

        /// <summary>
        /// Материально-ответственное лицо
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public string Principal
        {
            get => principal;
            set => principal = value;
        }

        /// <summary>
        /// Город размещения/регистрации фирмы
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public string City
        {
            get => city;
            set => city = value;
        }

        /// <summary>
        /// Адрес, по которому размещена/зарегистрирована фирма
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public string Address
        {
            get => address;
            set => address = value;
        }

        /// <summary>
        /// Телефон
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public string Phone
        {
            get => phone;
            set => phone = value;
        }

        /// <summary>
        /// Налоговый номер
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public string TaxNumber
        {
            get => taxNumber;
            set => taxNumber = value;
        }

        /// <summary>
        /// Номер НДС
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public string VATNumber
        {
            get => vatNumber;
            set => vatNumber = value;
        }

        /// <summary>
        /// Номер банковского счета
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>12.02.2021</date>
        public string IBAN
        {
            get => iBAN;
            set => iBAN = value;
        }
        #endregion
    }
}
