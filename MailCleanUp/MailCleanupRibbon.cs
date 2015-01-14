using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MailCleanUp.Properties;
using Microsoft.Office.Tools.Ribbon;

namespace MailCleanUp
{
    public partial class MailCleanupRibbon
    {
        private int _maxItemsToDelete;

        private string MaxItemsToDeleteText
        {
            get { return _maxItemsToDelete.ToString("G0"); }
        }

        private void MailCleanupRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _maxItemsToDelete = 1000;
            MaxItemsBox.Text = MaxItemsToDeleteText;
        }

        private void MaxItemsBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            int maxItemsToDelete;
            if (int.TryParse(MaxItemsBox.Text, out maxItemsToDelete))
            {
                _maxItemsToDelete = maxItemsToDelete;
            }
            else
            {
                MaxItemsBox.Text = MaxItemsToDeleteText;
                MessageBox.Show(Resources.Max_Items_Not_Valid_Integer);
            }
        }

        private void StartCleanupButton_Click(object sender, RibbonControlEventArgs e)
        {
            var error = Globals.ThisAddIn.RunDeleteOldMessagesTask(_maxItemsToDelete);
            if (error != null)
            {
                MessageBox.Show(error);
            }
        }
    }
}
