using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace QuickContactInfo
{
    partial class MeetingContactInfoRegion
    {
        Core.CustomUserRegions uiManager = new Core.CustomUserRegions();

        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("QuickContactInfo.MeetingContactInfoRegion")]
        public partial class MeetingContactInfoRegionFactory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void MeetingContactInfoRegionFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
                e.Cancel = Core.Region.CanInitialize(e);
            }
        }

        #endregion

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void MeetingContactInfoRegion_FormRegionShowing(object sender, System.EventArgs e)
        {
            uiManager.LinkControls(tblContacts);
            uiManager.StartAsync(OutlookItem);
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void MeetingContactInfoRegion_FormRegionClosed(object sender, System.EventArgs e)
        {
            uiManager.Cancel();
        }
    }
}
