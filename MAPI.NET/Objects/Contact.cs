using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MAPI.NET
{
    /// <summary>
    ///  The Contact class is used to maintain and provide access to the properties of contacts.
    /// </summary>
    public class Contact : MAPIMessage
    {
        enum OutlookTag
        {
            OUTLOOK_DATA1 = 0x00062004,
            OUTLOOK_EMAIL1 = 0x8083,
            OUTLOOK_EMAIL2 = 0x8093,
            OUTLOOK_EMAIL3 = 0x80A3,
            OUTLOOK_IM_ADDRESS = 0x8062,
            OUTLOOK_FILE_AS = 0x8005,
            OUTLOOK_POSTAL_ADDRESS = 0x8022,
            OUTLOOK_DISPLAY_ADDRESS_HOME = 0x801A,
            OUTLOOK_PICTURE_FLAG = 0x8015,
            OUTLOOK_CATEGORIES = 0xF101E
        }

        /// <exclude/>
        public Contact(IMessage contact)
            : base(contact, MAPIObject.ContactItem)
        {
            SetPropertyValue(PropTags.PR_MESSAGE_CLASS, MAPIObject.ContactItem);
        }

        #region Public Properties

        /// <summary>
        /// Gets/Sets FullName
        /// </summary>
        public string FullName
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_DISPLAY_NAME);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_DISPLAY_NAME, value);
            }
        }
        /// <summary>
        /// Gets/sets given name
        /// </summary>
        public string GivenName
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_GIVEN_NAME);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_GIVEN_NAME, value);
            }
        }

        /// <summary>
        /// Gets/sets Surname
        /// </summary>
        public string Surname
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_SURNAME);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_SURNAME, value);
            }
        }

        /// <summary>
        /// Gets/sets Middle Name
        /// </summary>
        public string MiddleName
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_MIDDLE_NAME);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_MIDDLE_NAME, value);
            }
        }

        /// <summary>
        /// Gets/Sets PrimaryTelephoneNumber
        /// </summary>
        public string PrimaryTelephoneNumber
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_PRIMARY_TELEPHONE_NUMBER);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_PRIMARY_TELEPHONE_NUMBER, value);
            }

        }

        /// <summary>
        /// Gets/sets FileAs
        /// </summary>
        public string FileAs
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_FILE_AS);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_FILE_AS, true);
                SetPropertyValue(tag, value);
            }
        }
        /// <summary>
        /// Gets/Sets PrimaryFaxNumber
        /// </summary>
        public string PrimaryFaxNumber
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_PRIMARY_FAX_NUMBER);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_PRIMARY_FAX_NUMBER, value);
            }
        }

        /// <summary>
        /// Gets/Sets BusinessTelephoneNumber
        /// </summary>
        public string BusinessTelephoneNumber
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_BUSINESS_TELEPHONE_NUMBER);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_BUSINESS_TELEPHONE_NUMBER, value);
            }
        }

        /// <summary>
        /// Gets/Sets Business2TelephoneNumber
        /// </summary>
        public string Business2TelephoneNumber
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_BUSINESS2_TELEPHONE_NUMBER);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_BUSINESS2_TELEPHONE_NUMBER, value);
            }
        }

        /// <summary>
        /// Gets/Sets Bussiness Fax Number
        /// </summary>
        public string BusinessFaxNumber
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_BUSINESS_FAX_NUMBER);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_PRIMARY_FAX_NUMBER, value);
            }
        }

        /// <summary>
        /// Gets/Sets HomeTelephoneNumber
        /// </summary>
        public string HomeTelephoneNumber
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_HOME_TELEPHONE_NUMBER);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_HOME_TELEPHONE_NUMBER, value);
            }
        }

        /// <summary>
        /// Gets/Sets Home2TelephoneNumber
        /// </summary>
        public string Home2TelephoneNumber
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_HOME2_TELEPHONE_NUMBER);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_HOME2_TELEPHONE_NUMBER, value);
            }
        }

        /// <summary>
        /// Gets/Sets HomeFaxNumber
        /// </summary>
        public string HomeFaxNumber
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_HOME_FAX_NUMBER);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_HOME_FAX_NUMBER, value);
            }
        }

        /// <summary>
        /// Gets/Sets MobileNumber
        /// </summary>
        public string MobileTelephoneNumber
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_MOBILE_TELEPHONE_NUMBER);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_MOBILE_TELEPHONE_NUMBER, value);
            }
        }

        /// <summary>
        /// Gets/Sets the email1 address of the contact
        /// </summary>
        public string EmailAddress1
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL1);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL1, true);
                SetPropertyValue(tag, value);
            }
        }

        /// <summary>
        /// Gets the email1 display as text of the contact
        /// </summary>
        public string Email1DisplayAs
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL1 - 3);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL1 - 3, true);
                SetPropertyValue(tag, value);
            }
        }

        /// <summary>
        /// Gets/Sets the email2 address of the contact
        /// </summary>
        public string EmailAddress2
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL2);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL2, true);
                SetPropertyValue(tag, value);
            }
        }

        /// <summary>
        /// Gets the email2 display as text of the contact
        /// </summary>
        public string Email2DisplayAs
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL2 - 3);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL2 - 3, true);
                SetPropertyValue(tag, value);
            }
        }

        /// <summary>
        /// Gets/Sets the email3 address of the contact
        /// </summary>
        public string EmailAddress3
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL3);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL3, true);
                SetPropertyValue(tag, value);
            }
        }

        /// <summary>
        /// Gets the email3 display as text of the contact
        /// </summary>
        public string Email3DisplayAs
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL3 - 3);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_EMAIL3 - 3, true);
                SetPropertyValue(tag, value);
            }
        }

        /// <summary>
        /// Gets/Sets the IM address of the contact
        /// </summary>
        public string IMAddress
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_IM_ADDRESS);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_IM_ADDRESS, true);
                SetPropertyValue(tag, value);
            }
        }

        /// <summary>
        /// Gets/Sets the profession of the contact
        /// </summary>
        public string Profession
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_PROFESSION);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_PROFESSION, value);
            }
        }

        /// <summary>
        /// Gets/Sets the business homepage of the contact
        /// </summary>
        public string BusinessHomePage
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_BUSINESS_HOME_PAGE);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_BUSINESS_HOME_PAGE, value);
            }
        }

        /// <summary>
        /// Gets/Sets the personal homepage of the contact
        /// </summary>
        public string PersonalHomePage
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_PERSONAL_HOME_PAGE);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_PERSONAL_HOME_PAGE, value);
            }
        }

        /// <summary>
        /// Gets the postal address, for quick access to a single string
        /// </summary>
        public string PostalAddress
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_POSTAL_ADDRESS);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_POSTAL_ADDRESS, value);
            }
        }

        /// <summary>
        ///  Gets/Sets the job title field
        /// </summary>
        public string JobTitle
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_TITLE);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_TITLE, value);
            }
        }

        /// <summary>
        /// Gets/Sets the Company field
        /// </summary>
        public string CompanyName
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_COMPANY_NAME);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_COMPANY_NAME, value);
            }
        }

        /// <summary>
        /// Gets/Sets the prefix (Mr., Dr. etc)
        /// </summary>
        public string DisplayNamePrefix
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_DISPLAY_NAME_PREFIX);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_DISPLAY_NAME_PREFIX, value);
            }
        }

        /// <summary>
        /// Gets/Sets the generation
        /// </summary>
        public string Generation
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_GENERATION);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_GENERATION, value);
            }
        }

        /// <summary>
        ///  Gets/sets the department field
        /// </summary>
        public string Department
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_DEPARTMENT_NAME);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_DEPARTMENT_NAME, value);
            }
        }

        /// <summary>
        /// Gets/sets the office field
        /// </summary>
        public string Office
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_OFFICE_LOCATION);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_OFFICE_LOCATION, value);
            }
        }

        /// <summary>
        /// Gets the name of the contact's manager
        /// </summary>
        public string Manager
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_MANAGER_NAME);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_MANAGER_NAME, value);
            }
        }

        /// <summary>
        /// Gets the name of the contact's assistant
        /// </summary>
        public string Assistant
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_ASSISTANT);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_ASSISTANT, value);
            }
        }

        /// <summary>
        /// Gets/sets the contact's nickname
        /// </summary>
        public string NickName
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_NICKNAME);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_NICKNAME, value);
            }
        }

        /// <summary>
        /// Gets/sets the name of the contact's spouse
        /// </summary>
        public string Spouse
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_SPOUSE_NAME);
                if (value != null)
                    return value.AsString;
                return string.Empty;
            }
            set
            {
                SetPropertyValue(PropTags.PR_SPOUSE_NAME, value);
            }
        }

        /// <summary>
        /// Gets/sets the contact's birthday
        /// </summary>
        public DateTime? Birthday
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_BIRTHDAY);
                if (value != null)
                    return value.AsDateTime;
                return null;
            }
            set
            {
                SetPropertyValue(PropTags.PR_BIRTHDAY, value);
            }
        }

        /// <summary>
        /// Gets the contact's wedding anniversary
        /// </summary>
        public DateTime? Anniversary
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_WEDDING_ANNIVERSARY);
                if (value != null)
                    return value.AsDateTime;
                return null;
            }
            set
            {
                SetPropertyValue(PropTags.PR_WEDDING_ANNIVERSARY, value);
            }
        }

        /// <summary>
        /// Gets the categories (stored under MV property "Keywords"
        /// </summary>
        public string Categories
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_CATEGORIES);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA1, (uint)OutlookTag.OUTLOOK_CATEGORIES, true);
                SetPropertyValue(tag, value);
            }
        }

        #endregion

    }
}
