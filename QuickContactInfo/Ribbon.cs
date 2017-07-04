using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace QuickContactInfo
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            mnuToggle.Label = "Quick Pane";
            mnuToggle.SuperTip = "Enable/Disable the quick contact info pane.";

            UpdateButton();
        }

        private void UpdateButton()
        {
            var meetingsOn = Properties.Settings.Default.MeetingsOn;
            var messagesOn = Properties.Settings.Default.MessagesOn;

            if(meetingsOn == false && messagesOn == false)
            {
                btnOff.Checked = true;
                btnEnableMeeting.Checked = false;
                btnEnableMessage.Checked = false;
            }
            else
            {
                btnOff.Checked = false;
                btnEnableMessage.Checked = messagesOn;
                btnEnableMeeting.Checked = meetingsOn;
            }
        }

        private void btnEnableMessage_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.MessagesOn = btnEnableMessage.Checked;
            Properties.Settings.Default.Save();
            UpdateButton();
        }

        private void btnEnableMeeting_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.MeetingsOn = btnEnableMeeting.Checked;
            Properties.Settings.Default.Save();
            UpdateButton();
        }

        private void btnOff_Click(object sender, RibbonControlEventArgs e)
        {
            if (btnOff.Checked)
            {
                Properties.Settings.Default.MessagesOn = false;
                Properties.Settings.Default.MeetingsOn = false;
            }
            else
            {
                Properties.Settings.Default.MessagesOn = true;
                Properties.Settings.Default.MeetingsOn = true;
            }
            Properties.Settings.Default.Save();
            UpdateButton();
        }
        
    }
}
