using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace SentimentAnalysisAddIn
{
    public enum PriorityLevel { Low, Medium, High };
    public partial class SelectPriorityForm : Form
    {
        private readonly ComboBox combo;
        private readonly Button buttonOkay;
        private readonly Button buttonCancel;

        public PriorityLevel Priority { get; private set; }

        public SelectPriorityForm()
        {
            this.Text = "Select Priority";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Width = 250;
            this.Height = 150;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            combo = new ComboBox()
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 25
            };
            combo.Items.AddRange(new[] { "Low", "Medium", "High" });
            combo.SelectedIndex = 1; // Default = Medium
            this.Controls.Add(combo);

            buttonOkay = new Button()
            {
                Text = "Okay",
                DialogResult = DialogResult.OK,
                Dock = DockStyle.Bottom,
                Height = 30
            };
            buttonOkay.Click += ButtonOkay_Click;
            this.Controls.Add(buttonOkay);

            buttonCancel = new Button()
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Dock = DockStyle.Bottom,
                Height = 30
            };
            this.Controls.Add(buttonCancel);

            this.AcceptButton = buttonOkay;
            this.CancelButton = buttonCancel;
        }

        private void ButtonOkay_Click(object sender, EventArgs e)
        {
            // Map selected button to enum
            Priority = (PriorityLevel)combo.SelectedIndex;
            this.Close();
        }
    }
}
