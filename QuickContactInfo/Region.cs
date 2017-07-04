using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace QuickContactInfo.Core
{
    internal class Region
    {
        public List<ContactInfo> Contacts { get; set; }
        private string activeEmailAddress = string.Empty;

        public Region()
        {
            Contacts = new List<ContactInfo>();
            activeEmailAddress = Globals.ThisAddIn.CurrentUserEmail ?? string.Empty; 
        }

        public static bool CanInitialize(Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
        {
            if(Properties.Settings.Default.MeetingsOn == false && Properties.Settings.Default.MessagesOn == false)
            {
                return true;
            }

            if (e.OutlookItem is Outlook.MailItem)
            {
                if (!Properties.Settings.Default.MessagesOn)
                    return true;

                Outlook.MailItem mailItem = (Outlook.MailItem)e.OutlookItem;
                if (mailItem != null && mailItem.Sent)
                {
                    return false;
                }
            }
            else if (e.OutlookItem is Outlook.AppointmentItem)
            {
                if (!Properties.Settings.Default.MeetingsOn)
                    return true;

                Outlook.AppointmentItem appItem = (Outlook.AppointmentItem)e.OutlookItem;
                if (appItem != null && appItem.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting)
                {
                    return false;
                }
            }

            return true;
        }
        

        public void ProcessContacts(object outlookItem, bool sync = true)
        {
            if (outlookItem == null)
                return;

            if (outlookItem is Outlook.MailItem)
            {
                if (sync)
                    ProcessMailItem(outlookItem);
                else ProcessMailItemAsync(outlookItem);
            }
            else if (outlookItem is Outlook.AppointmentItem)
            {
                if (sync)
                    ProcessAppointmentItem(outlookItem);
                else ProcessAppointmentItemAsync(outlookItem);
            }
        }

        private void ProcessMailItem(object outlookItem)
        {
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookItem;
            if (mailItem != null)
            {
                Contacts.Add(new ContactInfo(mailItem.Sender));

                foreach (Outlook.Recipient recipient in mailItem.Recipients)
                {
                    if (!recipient.Resolved)
                    {
                        if (!recipient.Resolve())
                            continue;
                    }

                    Contacts.Add(new ContactInfo(recipient.AddressEntry));
                }
            }
        }

        private void ProcessMailItemAsync(object outlookItem)
        {
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookItem;
            if (mailItem != null)
            {
                Contacts.Add(ContactInfo.UnResolved(mailItem.Sender.ID, true));
                
                foreach (Outlook.Recipient recipient in mailItem.Recipients)
                {
                    Contacts.Add(ContactInfo.UnResolved(recipient.AddressEntry.ID, true));
                }
            }
        }

        private void ProcessAppointmentItem(object outlookItem)
        {
            Outlook.AppointmentItem appItem = (Outlook.AppointmentItem)outlookItem;
            if (appItem != null)
            {
                AddMeetingLocation(appItem);

                foreach (Outlook.Recipient recipient in appItem.Recipients)
                {
                    if (!recipient.Resolved)
                    {
                        if (!recipient.Resolve())
                            continue;
                    }

                    Contacts.Add(new ContactInfo(recipient.AddressEntry));
                }
            }
        }

        private void ProcessAppointmentItemAsync(object outlookItem)
        {
            Outlook.AppointmentItem appItem = (Outlook.AppointmentItem)outlookItem;
            if (appItem != null)
            {
                Contacts.Add(ContactInfo.UnResolved(appItem.Location, false));

                foreach (Outlook.Recipient recipient in appItem.Recipients)
                {
                    Contacts.Add(ContactInfo.UnResolved(recipient.AddressEntry.ID, true));
                }
            }
        }

        private void AddMeetingLocation(Outlook.AppointmentItem appItem)
        {
            //search gal
            var exchange = SearchExchange(appItem.Location);
            if (exchange != null)
            {
                Contacts.Add(new ContactInfo(exchange));
                return;
            }

            //search local contacts
            var contact = SearchOutlookContacts(appItem.Location);
            if (contact != null)
            {
                Contacts.Add(new ContactInfo(contact));
                return;
            }            

            //add empty
            Contacts.Add(ContactInfo.Empty(appItem.Location));
        }


        private Outlook.ExchangeUser SearchExchange(string name, Outlook.NameSpace _ns = null)
        {
            if (Globals.ThisAddIn.Offline)
                return null;

            Outlook.NameSpace ns = _ns;
            if(ns == null)
            {
                ns = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            }

            if (ns != null)
            {
                var addresslist = ns.GetGlobalAddressList();

                if (addresslist != null)
                {
                    try
                    {
                        var addressItem = addresslist.AddressEntries[name];
                        if (addressItem != null && (
                            addressItem.AddressEntryUserType == Microsoft.Office.Interop.Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                            addressItem.AddressEntryUserType == Microsoft.Office.Interop.Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry))
                        {
                            return addressItem.GetExchangeUser();
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        return null;
                    }
                }
            }
            
            return null;
        }

        private Outlook.ContactItem SearchOutlookContacts(string name, Outlook.NameSpace _ns = null)
        {
            Outlook.NameSpace ns = _ns;
            if(ns == null)
            {
                ns = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            }

            if (ns != null)
            {
                Outlook.MAPIFolder folder = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
                Outlook.Items contacts = folder.Items;

                try
                {
                    var contact = (Outlook.ContactItem)contacts.Find(string.Format("[FullName]='{0}'", name));
                    return contact;
                }
                catch (Exception)
                {
                    return null;
                }
            }

            return null;
        }

        public IEnumerable<bool> ResolveContacts()
        {
            var ns = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            if (ns != null)
            {
                foreach (var contact in Contacts.Where(x => x.Resolved == false))
                {
                    if (!string.IsNullOrEmpty(contact.AddressEntryId))
                    {
                        var addressEntry = ns.GetAddressEntryFromID(contact.AddressEntryId);
                        if(addressEntry != null)
                        {
                            contact.Process(addressEntry);
                        }
                    }
                    else if (!string.IsNullOrEmpty(contact.NameIdentifier))
                    {
                        var resolvedExcahnge = SearchExchange(contact.NameIdentifier, ns);
                        if(resolvedExcahnge != null)
                        {
                            contact.Resolved = true;
                            contact.DisplayName = resolvedExcahnge.Name;
                            contact.Email.Add(resolvedExcahnge.PrimarySmtpAddress);
                            contact.Phone.Add(resolvedExcahnge.BusinessTelephoneNumber);
                            contact.Mobile.Add(resolvedExcahnge.MobileTelephoneNumber);
                        }
                        else
                        {
                            var resolvedContact = SearchOutlookContacts(contact.NameIdentifier, ns);
                            if(resolvedContact != null)
                            {
                                contact.Resolved = true;
                                contact.DisplayName = resolvedContact.FullName;
                                contact.Email.Add(resolvedContact.Email1Address);
                                contact.Email.Add(resolvedContact.Email2Address);
                                contact.Email.Add(resolvedContact.Email3Address);
                                contact.Phone.Add(resolvedContact.BusinessTelephoneNumber);
                                contact.Phone.Add(resolvedContact.Business2TelephoneNumber);
                                contact.Mobile.Add(resolvedContact.MobileTelephoneNumber);
                            }
                        }
                    }

                    yield return true;                  
                }
            }
        }
    }
}
