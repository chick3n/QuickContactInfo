using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace QuickContactInfo
{
    public partial class ThisAddIn
    {
        public string CurrentUserEmail { get; set; }
        public Outlook.Recipient CurrentUser
        {
            get
            {
                try
                {
                    return Application.Session.CurrentUser;
                }
                catch(System.Runtime.InteropServices.COMException)
                {
                    return null;
                }
            }
        }

        public Boolean Offline
        {
            get
            {
                return Globals.ThisAddIn.Application.GetNamespace("MAPI").Offline;
            }
        }
        

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            setCurrentUserEmail();
        }

        private void setCurrentUserEmail()
        {
            CurrentUserEmail = string.Empty;
            var currentUser = CurrentUser;
            if (currentUser != null)
            {
                var contact = new ContactInfo(currentUser.AddressEntry);
                if(contact.Resolved && contact.Email.Count > 0)
                {
                    CurrentUserEmail = contact.Email[0] ?? string.Empty;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }
        

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
