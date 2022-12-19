using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MAPI.NET
{
    /// <summary>
    /// Recipient address type enumeration
    /// </summary>
    public enum AddressType
    {
        /// <summary>
        /// Exchange
        /// </summary>
        EX,
        /// <summary>
        /// SMTP
        /// </summary>
        SMTP,
    }

    /// <summary>
    /// Defines a Recipient for MAPI messages. 
    /// </summary>
    public class Recipient
    {
        /// <summary>
        /// Initializes a new instance of the recipient class.
        /// </summary>
        /// <param name="id">Entry identification of recipient</param>
        /// <param name="name">Name of recipient</param>
        /// <param name="email">Email of recipient</param>
        /// <param name="type">Address type of recipient</param>
        public Recipient(EntryID id, string name, string email, AddressType type)
        {
            EntryId = id;
            Name = name;
            Email = email;
            Type = type;
        }
        /// <summary>
        /// Gets entry identification
        /// </summary>
        public EntryID EntryId {get; private set;}
        /// <summary>
        /// Gets name
        /// </summary>
        public string Name { get; private set; }
        /// <summary>
        /// Gets email
        /// </summary>
        public string Email { get; private set; }
        /// <summary>
        /// Gets address type
        /// </summary>
        public AddressType Type { get; private set; }
        /// <summary>
        /// Get exchange address
        /// </summary>
        /// <param name="addressBook">MAPI address book object</param>
        public void GetExchangeAddress(MAPIAddressBook addressBook)
        {
            if (Type != AddressType.EX || EntryId == null)
                return;
            Email = addressBook.GetEmailAddress(EntryId);
        }
    }
}
