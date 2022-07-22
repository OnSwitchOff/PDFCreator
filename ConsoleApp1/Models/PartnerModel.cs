using ConsoleApp1.Enums;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Models
{
    public class PartnerModel : CompanyModel
    {
        #region "Declarations"
        // ID партнёра
        private int id;
        // номер дисконтной карты
        private string discountCard;
        // электронная почта
        private string e_mail;
        // ID группы партнёра
        private int groupId;
        // скидка по группе партнёра
        private int partnerGroupDiscount;
        // статус партнёра (напр., удалён, часто используется и т.п.)
        private int status;
        #endregion

        #region "Constructor"
        /// <summary>
        /// Заполняет экземпляр пустыми данными
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public PartnerModel() : base()
        {
            ID = 0;
            DiscountCard = string.Empty;
            E_mail = string.Empty;
            GroupID = 0;
            Status = (int)Statuses.Available;
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// ID партнёра
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public int ID
        {
            get => id;
            set => id = value;
        }

        /// <summary>
        /// Номер дисконтной карты
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public string DiscountCard
        {
            get => discountCard;
            set => discountCard = value;
        }

        /// <summary>
        /// Электронная почта
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public string E_mail
        {
            get => e_mail;
            set => e_mail = value;
        }

        /// <summary>
        /// ID группы партнёра
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public int GroupID
        {
            get => groupId;
            set => groupId = value;
        }

        /// <summary>
        /// Скидка по группе партнёра
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>01.02.2021</date>
        public int PartnerGroupDiscount
        {
            get => partnerGroupDiscount;
            set => partnerGroupDiscount = value;
        }

        /// <summary>
        /// Статус партнёра (напр., удалён, часто используется и т.п.)
        /// </summary>
        /// <developer>Сергей Рознюк</developer>
        /// <date>15.01.2021</date>
        public int Status
        {
            get => status;
            set => status = value;
        }
        #endregion

        public static explicit operator PDFCreator.Models.CompanyModel(PartnerModel partnerModel)
        {
            PDFCreator.Models.CompanyModel companyModel = new PDFCreator.Models.CompanyModel();
            companyModel.Name = partnerModel.Company;
            companyModel.Address = partnerModel.Address;
            companyModel.Principal = partnerModel.Principal;
            companyModel.TaxNumber = partnerModel.TaxNumber;
            companyModel.Phone = partnerModel.Phone;
            companyModel.VATNumber = partnerModel.VATNumber;

            return companyModel;
        }
    }
}
