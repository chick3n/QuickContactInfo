using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuickContactInfo.Core
{
    internal class CustomUserRegions
    {
        private TableLayoutPanel tblContacts;
        private Core.Region regionManager;
        private CancellationTokenSource tokenSource;

        public CustomUserRegions()
        {
            
            regionManager = new Region();
        }

        public void LinkControls(TableLayoutPanel tblContacts)
        {
            this.tblContacts = tblContacts;
        }

        public void LinkContacts(Core.Region regionManager)
        {
            this.regionManager = regionManager;
        }

        public void Cancel()
        {
            if(tokenSource != null)
            {
                tokenSource.Cancel();
            }
        }

        public async void StartAsync(object outlookItem)
        {
            if(outlookItem == null || tblContacts == null)
            {
                return;
            }
            
            UpdateRegionWaiting();
            regionManager.ProcessContacts(outlookItem, false);
            tokenSource = new CancellationTokenSource(10000);

            await Task.Run(() => ResolveContacts(tokenSource.Token));
            if (tokenSource.IsCancellationRequested)
                UpdateRegionWaiting("Timed out.");
            else UpdateRegion();
        }

        internal void UpdateRegion()
        {
            if (tblContacts == null)
                return;

            var contacts = regionManager.Contacts ?? new List<ContactInfo>();

            tblContacts.RowStyles.Clear();
            tblContacts.Controls.Clear();
            
            foreach (var contact in contacts.Where(x => x.Resolved && !string.IsNullOrEmpty(x.DisplayName)))
            {
                if (contact.Email.Contains(Globals.ThisAddIn.CurrentUserEmail))
                {
                    continue;
                }

                var index = tblContacts.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                Label displayName = new Label();
                displayName.AutoSize = true;
                displayName.Text = contact.DisplayName;

                tblContacts.Controls.Add(displayName, 0, index);
                tblContacts.Controls.Add(LabelMultipleItems(contact.Email), 1, index);
                tblContacts.Controls.Add(LabelMultipleItems(contact.Phone), 2, index);
                tblContacts.Controls.Add(LabelMultipleItems(contact.Mobile, prefix:"(M) "), 3, index);
            }
        }

        internal void UpdateRegionWaiting(string message = "Gathering...")
        {
            tblContacts.RowStyles.Clear();
            tblContacts.Controls.Clear();

            var index = tblContacts.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Label displayName = new Label();
            displayName.AutoSize = true;
            displayName.Text = message;

            tblContacts.Controls.Add(displayName, 0, index);
        }

        private Label LabelMultipleItems(List<string> items, string sep = "\n", string prefix = "")
        {
            Label item = new Label();
            item.AutoSize = true;
            item.Text = string.Empty;
            if (items != null && items.Count > 0)
            {
                item.Text = string.Join(sep, items.Select(x => string.IsNullOrEmpty(x) ? string.Empty : prefix + x));
            }
            return item;
        }

        private void ResolveContacts(CancellationToken token)
        {
            if (regionManager == null || tblContacts == null)
                return;

            foreach (var result in regionManager.ResolveContacts())
            {
                if (token.IsCancellationRequested)
                    break;
            }
        }
        
        
    }
}
