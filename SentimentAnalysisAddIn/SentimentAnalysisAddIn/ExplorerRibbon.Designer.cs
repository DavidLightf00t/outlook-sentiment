using Microsoft.Office.Tools.Ribbon;

namespace SentimentAnalysisAddIn
{
    partial class ExplorerRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        internal RibbonTab tabSentiment;
        internal RibbonGroup groupSentiment;
        internal RibbonDropDown dropdownSubfolders;

        private void InitializeComponent()
        {
            var factory = this.Factory;

            // Create a brand-new tab:
            this.tabSentiment = factory.CreateRibbonTab();
            this.tabSentiment.Label = "Sentiment Analysis";
            this.tabSentiment.Name = "tabSentiment";

            // Create a group on that tab:
            this.groupSentiment = factory.CreateRibbonGroup();
            this.groupSentiment.Label = "Priority";
            this.groupSentiment.Name = "groupSentiment";
            this.tabSentiment.Groups.Add(this.groupSentiment);

            // Add your dropdown there:
            this.dropdownSubfolders = factory.CreateRibbonDropDown();
            this.dropdownSubfolders.Label = "Go To…";
            this.dropdownSubfolders.Name = "dropdownSubfolders";
            this.dropdownSubfolders.SelectionChanged +=
                new RibbonControlEventHandler(this.dropdownSubfolders_SelectionChanged);
            this.groupSentiment.Items.Add(this.dropdownSubfolders);

            // Finish
            this.Name = "ExplorerRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabSentiment);
            this.Load += new RibbonUIEventHandler(this.ExplorerRibbon_Load);
        }
    }
}