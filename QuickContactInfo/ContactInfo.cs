using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickContactInfo
{
    public class ContactInfo
    {

        public static ContactInfo Empty(string name = null)
        {
            var empty = new ContactInfo();
            empty.DisplayName = name ?? string.Empty;
            return empty;
        }

        public static ContactInfo UnResolved(string id, bool addressEntry)
        {
            var unresolved = new ContactInfo();
            if (addressEntry) unresolved.AddressEntryId = id;
            else unresolved.NameIdentifier = id;
            return unresolved;
        }

        public ContactInfo()
        {
            Initialize();
        }

        public ContactInfo(Outlook.AddressEntry address)
        {
            Initialize();
            Process(address);
        }

        public ContactInfo(Outlook.ContactItem contact)
        {
            Initialize();
            ProcessContact(contact);
        }

        public ContactInfo(Outlook.ExchangeUser exchangeUser)
        {
            Initialize();
            ProcessExchangeUser(exchangeUser);
            DisplayName = exchangeUser.Name ?? string.Empty;
        }
        

        private void Initialize()
        {
            Resolved = false;
            Email = new List<string>();
            Phone = new List<string>();
            Mobile = new List<string>();
        }
        
        public void Process(Outlook.AddressEntry address)
        {
            if (address == null)
                return;

            if (address.AddressEntryUserType == Microsoft.Office.Interop.Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                address.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                ProcessExchangeUser(address);
            else
                ProcessContact(address);

            DisplayName = address.Name ?? string.Empty;
        }

        private void Process(Outlook.ContactItem contact)
        {
            ProcessContact(contact);
            DisplayName = contact.FirstName ?? string.Empty + " " + contact.LastName ?? string.Empty;
        }

        private void ProcessExchangeUser(Outlook.AddressEntry address)
        {

            if (Globals.ThisAddIn.Offline)
            {
                var contact = address.GetContact();
                if(contact != null)
                {
                    ProcessContact(contact);
                    return;
                }
            }

            ProcessExchangeUser(address.GetExchangeUser());
        }

        private void ProcessExchangeUser(Outlook.ExchangeUser exchangeUser)
        {
            try
            {
                Resolved = true;
                Email.Add(exchangeUser.PrimarySmtpAddress);
                Phone.Add(exchangeUser.BusinessTelephoneNumber);
                Mobile.Add(exchangeUser.MobileTelephoneNumber);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Resolved = false;
            }
        }
        
        private void ProcessContact(Outlook.AddressEntry address, bool useAddress = true)
        {
            var contact = address.GetContact();
            if(contact == null && useAddress)
            {
                Email.Add(address.Address);
                return;
            }

            ProcessContact(contact);
        }

        private void ProcessContact(Outlook.ContactItem contact)
        {
            Resolved = true;
            Phone.Add(contact.PrimaryTelephoneNumber);
            Mobile.Add(contact.MobileTelephoneNumber);
            Email.Add(contact.Email1Address);
            Email.Add(contact.Email2Address);
            Email.Add(contact.Email3Address);
        }

        public string DisplayName { get; set; }
        public List<string> Email { get; set; }
        public List<string> Phone { get; set; }
        public List<string> Mobile { get; set; }

        public bool Resolved { get; set; }
        public string NameIdentifier { get; set; }
        public string AddressEntryId { get; set; }
    }
}
