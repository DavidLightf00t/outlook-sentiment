using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace SentimentAnalysisAddIn
{
    public partial class ExplorerRibbon : RibbonBase
    {
        public ExplorerRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void ExplorerRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            dropdownSubfolders.Items.Clear();

            string[] labels = {"------------", "High Priority", "Medium Priority", "Low Priority" };
            string[] tags = {"null", "high", "medium", "low" };

            for (int i = 0; i < labels.Length; i++)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = labels[i];
                item.Tag = tags[i];

                dropdownSubfolders.Items.Add(item);
            }
        }

        private void dropdownSubfolders_SelectionChanged(
            object sender, RibbonControlEventArgs e)
        {
            var tag = dropdownSubfolders.SelectedItem.Tag as string;
            Outlook.Folder dest = null;
            switch (tag)
            {
                case "low":
                    dest = Globals.ThisAddIn._lowPriorityFolder;
                    break;
                case "medium":
                    dest = Globals.ThisAddIn._mediumPriorityFolder;
                    break;
                case "high":
                    dest = Globals.ThisAddIn._highPriorityFolder;
                    break;
            }

            if (dest != null)
            {
                // Select that folder in the SAME Explorer window:
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                explorer.SelectFolder(dest);
            }
        }
    }
}