using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using OfficeCore = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SentimentAnalysisAddIn
{
    public partial class ThisAddIn
    {
        private Outlook.Folder _inbox;
        private Outlook.Folder _sentimentFolder;
        internal Outlook.Folder _highPriorityFolder, _mediumPriorityFolder, _lowPriorityFolder;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Build your Sentiment folder
            EnsureSentimentFolder();

            // Add the root to Favorites
            var mailModule = Application.ActiveExplorer()
                .NavigationPane.Modules
                .GetNavigationModule(Outlook.OlNavigationModuleType.olModuleMail)
                as Outlook.MailModule;

            var favGroup = mailModule.NavigationGroups
                .GetDefaultNavigationGroup(Outlook.OlGroupType.olFavoriteFoldersGroup);

            if (!favGroup.NavigationFolders
                .OfType<Outlook.NavigationFolder>()
                .Any(nf => nf.Folder.EntryID == _sentimentFolder.EntryID))
            {
                favGroup.NavigationFolders.Add(_sentimentFolder);
            }
        }
        
        private void EnsureSentimentFolder()
        {
            _inbox = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
            var root = _inbox.Parent as Outlook.Folder;

            // Find or create the Sentiment Analysis folder
            try
            {
                _sentimentFolder = root.Folders["Sentiment Analysis"] as Outlook.Folder;
            }
            catch
            {
                _sentimentFolder = root.Folders.Add(
                    "Sentiment Analysis",
                    Outlook.OlDefaultFolders.olFolderInbox
                ) as Outlook.Folder;
            }

            _highPriorityFolder = CreateSubFolder(_sentimentFolder, "High Priority");
            _mediumPriorityFolder = CreateSubFolder(_sentimentFolder, "Medium Priority");
            _lowPriorityFolder = CreateSubFolder(_sentimentFolder, "Low Priority");
        }

        private Outlook.Folder CreateSubFolder(Outlook.Folder parent, string name)
        {
            Outlook.Folder folder;
            // Find or create the sub folders "Low Priority, Medium Priority, High Priority" folder
            try
            {
                folder = parent.Folders[name] as Outlook.Folder;
            }
            catch
            {
                folder = parent.Folders.Add(
                    name,
                    Outlook.OlDefaultFolders.olFolderInbox
                ) as Outlook.Folder;
            }
            return folder;
        }

        // 4) (Optional) display that folder immediately
        // sentimentFolder.Display(false);

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}