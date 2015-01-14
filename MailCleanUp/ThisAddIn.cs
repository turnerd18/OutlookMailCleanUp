using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Timer = System.Timers.Timer;

namespace MailCleanUp
{
    public partial class ThisAddIn
    {
        private Outlook.NameSpace _appNameSpace;
        private Timer _timer;
        private static bool _running;
        private const int MaxItemsToDeleteInOneRun = 1000;
        private const string FoldersToCleanUpKey = "FoldersToCleanUpPath";
        
        private static Dictionary<string, int> _deleteDictionary;
        
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            InitializeDeleteDictionary();

            _appNameSpace = Application.GetNamespace("MAPI");
            if (_appNameSpace == null)
            {
                return;
            }

            var oneHourMillis = new TimeSpan(1, 0, 0).TotalMilliseconds;
            _timer = new Timer(oneHourMillis) { AutoReset = true };
            _timer.Elapsed += (timerSender, eventArgs) => RunDeleteOldMessagesTask();
            _timer.Start();
        }

        private static void InitializeDeleteDictionary()
        {
            _deleteDictionary = new Dictionary<string, int>();

            var folderListFileName = ConfigurationManager.AppSettings[FoldersToCleanUpKey] ?? string.Empty;
            try
            {
                if (!File.Exists(folderListFileName))
                {
                    throw new Exception(string.Concat("Folders to delete file not found at ", folderListFileName));
                }

                using (var reader = new StreamReader(folderListFileName))
                {
                    var body = reader.ReadToEnd();
                    var linesSplit =
                        body.Split(new[] {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                            .Select(line => line.Split('|'))
                            .Where(split => split.Length >= 2);
                    foreach (var lineSplit in linesSplit)
                    {
                        var folderName = lineSplit[0];
                        int daysOldBeforeDelete;
                        if (!int.TryParse(lineSplit[1], out daysOldBeforeDelete) 
                            || _deleteDictionary.ContainsKey(folderName))
                        {
                            continue;
                        }

                        _deleteDictionary.Add(folderName, daysOldBeforeDelete);
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _timer.Stop();
        }

        private void RunDeleteOldMessagesTask()
        {
            if (_running)
            {
                return;
            }

            Task.Run((Action) DeleteOldMessages);
        }

        private void DeleteOldMessages()
        {
            DeleteOldMessages(MaxItemsToDeleteInOneRun);
        }

        public void DeleteOldMessages(int maxItemsToDelete)
        {
            _running = true;
            var inbox = GetInbox();
            if (inbox == null)
            {
                MessageBox.Show(@"ERROR: Could not find inbox folder.");
                return;
            }

            var itemsDeleted = 0;
            foreach (var kvp in _deleteDictionary)
            {
                var folder =
                    inbox.Folders.Cast<Outlook.Folder>()
                        .FirstOrDefault(
                            f => string.Equals(f.Name, kvp.Key, StringComparison.InvariantCultureIgnoreCase));
                if (folder == null)
                {
                    MessageBox.Show(string.Format("Could not find folder {0}.", kvp.Key));
                    continue;
                }

                try
                {
                    var oldestAllowedDate = DateTime.Today.AddDays(-kvp.Value);
                    DeleteOldItems(folder, oldestAllowedDate, maxItemsToDelete, ref itemsDeleted);
                }
                finally
                {
                    Marshal.ReleaseComObject(folder);
                }
            }

            //MessageBox.Show(string.Format("Deleted {0} emails.", itemsDeleted));
            _running = false; 
        }

        private Outlook.Folder GetInbox()
        {
            var folders = _appNameSpace.Folders.Cast<Outlook.Folder>().ToArray();
            if (!folders.Any())
            {
                return null;
            }

            var root =
                folders.FirstOrDefault(
                    f => string.Equals(f.Name, "dturner@smg.com", StringComparison.InvariantCultureIgnoreCase));
            if (root == null)
            {
                return null;
            }

            Outlook.Folder inbox;
            try
            {
                inbox =
                    root.Folders.Cast<Outlook.Folder>()
                        .FirstOrDefault(
                            f => string.Equals(f.Name, "Inbox", StringComparison.InvariantCultureIgnoreCase));
            }
            finally
            {
                Marshal.ReleaseComObject(root);
            }

            return inbox;
        }

        private static void DeleteOldItems(Outlook.Folder folder, DateTime oldestAllowedDate, int maxItemsToDelete, ref int itemsDeleted)
        {
            foreach (var mailItem in folder.Items.Cast<Outlook.MailItem>())
            {
                try
                {
                    if (mailItem.ReceivedTime >= oldestAllowedDate)
                    {
                        continue;
                    }

                    if (itemsDeleted >= maxItemsToDelete)
                    {
                        break;
                    }

                    mailItem.Delete();
                    itemsDeleted++;
                }
                finally
                {
                    Marshal.ReleaseComObject(mailItem);
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
