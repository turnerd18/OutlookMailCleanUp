using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace MailCleanUp
{
    public partial class MailCleanupRibbon
    {
        private void MailCleanupRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void MaxItemsBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var inte = int.Parse(((RibbonEditBox) sender).Text);
        }

        private void StartCleanupButton_Click(object sender, RibbonControlEventArgs e)
        {
            int maxItemsToDelete;
            if (int.TryParse(MaxItemsBox.Text, out maxItemsToDelete))
            {
                Globals.ThisAddIn.DeleteOldMessages(maxItemsToDelete);
            }
        }
    }
}
