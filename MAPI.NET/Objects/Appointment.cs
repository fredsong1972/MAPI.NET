using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MAPI.NET
{
    /// <summary>
    /// The Appointment class is used to maintain and provide access to the properties of appointments.
    /// </summary>
    public class Appointment : MAPIMessage
    {
        enum OutlookTag
        {
            OUTLOOK_DATA2 = 0x00062002,
            OUTLOOK_APPOINTMENT_START = 0x820D,
            OUTLOOK_APPOINTMENT_END = 0x820E,
            OUTLOOK_APPOINTMENT_LOCATION = 0x8208,
            OUTLOOK_APPOINTMENT_ISRECURRING = 0x8223,
            OUTLOOK_APPOINTMENT_RESPONSE = 0x8218
        }

        /// <exclude/>
        public Appointment(IMessage appointment)
            : base(appointment, MAPIObject.AppointmentItem)
        {
            SetPropertyValue(PropTags.PR_MESSAGE_CLASS, MAPIObject.AppointmentItem);
        }

        #region Public Properties

        /// <exclude/>
        public string Location
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA2, (uint)OutlookTag.OUTLOOK_APPOINTMENT_LOCATION);
                return p.AsString;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA2, (uint)OutlookTag.OUTLOOK_APPOINTMENT_LOCATION, true);
                SetPropertyValue(tag, value);
            }
        }

        /// <exclude/>
        public DateTime? StartDateTime
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA2, (uint)OutlookTag.OUTLOOK_APPOINTMENT_START);
                return p.AsDateTime;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA2, (uint)OutlookTag.OUTLOOK_APPOINTMENT_START, true);
                SetPropertyValue(tag, value);
            }
        }

        /// <exclude/>
        public DateTime? EndtDateTime
        {
            get
            {
                IPropValue p = GetOutlookProperty((int)OutlookTag.OUTLOOK_DATA2, (uint)OutlookTag.OUTLOOK_APPOINTMENT_END);
                return p.AsDateTime;
            }
            set
            {
                uint tag = GetOutlookPropTag((int)OutlookTag.OUTLOOK_DATA2, (uint)OutlookTag.OUTLOOK_APPOINTMENT_END, true);
                SetPropertyValue(tag, value);
            }
        }


        #endregion

    }
}
